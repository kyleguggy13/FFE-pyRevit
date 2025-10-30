# -*- coding: utf-8 -*-
__title__     = "PDF Export \nSettings"
__version__   = 'Version = v0.1'
__doc__       = """Version = v0.1
Date    = 10.30.2025
________________________________________________________________
Tested Revit Versions: 
______________________________________________________________
Description:

______________________________________________________________
How-to:
 -> Click the button

______________________________________________________________
Last update:
 - [10.30.2025] - v0.1 Beta Release
 - [MM.DD.2025] - v1.0 First Release
______________________________________________________________
Author: Kyle Guggenheim"""


#____________________________________________________________________ IMPORTS (SYSTEM)
import os
import datetime

#____________________________________________________________________ IMPORTS (AUTODESK)
import clr
# Revit API
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import (
    FilteredElementCollector, ViewSheet, BuiltInCategory, ElementId, Transaction,
    # PDF API (Revit 2022+)
    PDFExportOptions, PDFExportQualityType, TableCellCombinedParameterData
)

clr.AddReference("RevitServices")
from RevitServices.Persistence import DocumentManager

#____________________________________________________________________ IMPORTS (PYREVIT)
from pyrevit import revit, DB
from pyrevit.script import output
from pyrevit import forms


output_window = output.get_output()



# # Optional: pyRevit output console
# try:
#     from pyrevit import script
#     out = script.get_output()
#     print_md = out.print_md
# except Exception:
#     def print_md(x): print(x)

# doc = DocumentManager.Instance.CurrentDBDocument
doc = revit.doc


# USER SETTINGS — EDIT THESE
#____________________________________________________________________ VARIABLES

# 1) Which sheets to export?
#    - If you prefer a picker UI later, swap this to a list you build from a dialog.
EXPORT_SELECTION_ONLY = False    # False => all sheets in the model
ALLOW_PLACEHOLDER_SHEETS = False # True to include placeholders; usually keep False

# 2) Output folder (defaults to user's Downloads\{ProjectName}\YYYY-MM-DD_HHMM)
USE_DEFAULT_DOWNLOADS = True
CUSTOM_OUTPUT_FOLDER = r""        # e.g., r"C:\_PDFs"; leave blank to auto-build

# 3) PDF export quality (DPI)
PDF_QUALITY = PDFExportQualityType.DPI300  # 144/300/600/... see enum

# 4) Graphics / page placement
ALWAYS_USE_RASTER = False         # True forces raster; False keeps vectors where possible
HIDE_CROP_BOUNDARIES = True
HIDE_UNREFERENCED_TAGS = True
MASK_COINCIDENT_LINES = False
REPLACE_HALFTONE_WITH_THIN_LINES = False
VIEW_LINKS_IN_BLUE = False
STOP_ON_ERROR = True              # True = stop if any view fails to export

# 5) Fit / Zoom
#    ZoomType is controlled by FitToPage vs. ZoomPercentage.
#    When ZoomPercentage is set (>0), Revit uses specific percentage.
FIT_TO_PAGE = False
ZOOM_PERCENT = 100                # only used if FIT_TO_PAGE=False

# 6) Naming Rule — order of fields in the filename (no extension).
#    You can mix:
#       - built-ins via the tokens: "SHEET_NUMBER", "SHEET_NAME"
#       - names of Sheet-bound shared/project parameters (exact display names)
#    Each entry can also define a custom separator/prefix/suffix if desired (see builder below).
NAMING_ORDER = [
    "SHEET_NUMBER",
    "SHEET_NAME",
    # Example shared/project parameters bound to Sheets:
    # "Discipline",
    # "Package",
    # "Submission",
]

# Separator used *between* each parameter block (Revit’s UI defaults to " - ")
GLOBAL_SEPARATOR = " - "

# Optionally prepend/append fixed text around the whole filename (rarely needed)
GLOBAL_PREFIX = ""     # e.g., "FFE_"
GLOBAL_SUFFIX = ""     # e.g., "_ISSUED"


#____________________________________________________________________ FUNCTIONS

def get_output_folder():
    """Build the output folder path."""
    if CUSTOM_OUTPUT_FOLDER.strip():
        base = CUSTOM_OUTPUT_FOLDER
    elif USE_DEFAULT_DOWNLOADS:
        try:
            # Get user's Downloads folder
            from System import Environment
            from System.Environment import SpecialFolder
            downloads = Environment.GetFolderPath(SpecialFolder.UserProfile)
            base = os.path.join(downloads, "Downloads")
        except:
            base = os.path.expanduser("~/Downloads")
    else:
        base = os.path.expanduser("~/Desktop")

    proj_name = doc.Title.replace(".rvt", "")
    ts = datetime.datetime.now().strftime("%Y-%m-%d_%H%M")
    folder = os.path.join(base, proj_name, ts)
    if not os.path.exists(folder):
        os.makedirs(folder)
    return folder

def collect_sheets():
    """Return a list[ViewSheet] based on settings."""
    coll = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
    sheets = []
    for s in coll:
        # Skip placeholders unless allowed
        if not ALLOW_PLACEHOLDER_SHEETS and s.IsPlaceholder:
            continue
        sheets.append(s)
    # If exporting only selection, filter down
    if EXPORT_SELECTION_ONLY:
        sel_ids = set([el.Id.IntegerValue for el in doc.Selection.GetElementIds()]) if hasattr(doc, "Selection") else set()
        sheets = [s for s in sheets if s.Id.IntegerValue in sel_ids]
    # Keep sorted by Sheet Number for predictable ordering
    try:
        sheets.sort(key=lambda x: x.SheetNumber)
    except Exception:
        pass
    return sheets

def build_table_cell_param_data_for_builtin(bip_enum, category_id, prefix="", suffix="", separator=""):
    """Create TableCellCombinedParameterData for a built-in parameter."""
    data = TableCellCombinedParameterData.Create()
    # ParamId expects ElementId. For built-ins, wrap the BuiltInParameter in an ElementId.
    data.ParamId = ElementId(bip_enum)
    data.CategoryId = category_id
    data.Prefix = prefix or ""
    data.Suffix = suffix or ""
    data.Separator = separator or ""
    return data

def find_parameter_element_id_by_name(param_display_name, category_id):
    """
    Resolve a *Sheet-bound* Project/Shared parameter by its display name.
    We walk the document's ParameterBindings and ensure it is bound to OST_Sheets.
    Returns ElementId for the ParameterElement (or InvalidElementId).
    """
    # ParameterBindings only exposes user-defined (project/shared) parameters.
    bm = doc.ParameterBindings
    it = bm.ForwardIterator()
    it.Reset()
    from Autodesk.Revit.DB import Definition, ElementId, CategorySet
    from Autodesk.Revit.DB import Binding, InstanceBinding, TypeBinding

    while it.MoveNext():
        definition = it.Key  # Internal Definition of the parameter
        binding = it.Current
        if definition is None or definition.Name != param_display_name:
            continue
        cats = None
        if isinstance(binding, InstanceBinding):
            cats = binding.Categories
        elif isinstance(binding, TypeBinding):
            cats = binding.Categories
        # Check it is bound to Sheets
        if cats:
            for c in cats:
                if c.Id == category_id:
                    # Get the ParameterElement by matching Definition
                    # (ParameterElement.GetDefinition().Name == param_display_name)
                    from Autodesk.Revit.DB import FilteredElementCollector, ParameterElement
                    for pe in FilteredElementCollector(doc).OfClass(ParameterElement):
                        try:
                            if pe.GetDefinition() and pe.GetDefinition().Name == param_display_name:
                                return pe.Id
                        except:
                            pass
    return ElementId.InvalidElementId

def build_naming_rule_for_sheets(naming_items, global_sep):
    """
    Build IList<TableCellCombinedParameterData> matching Revit’s Naming Rules UI:
    - Supports built-ins SHEET_NUMBER, SHEET_NAME
    - Supports Sheet-bound project/shared parameters by display name
    The 'Separator' applied to each block controls the delimiter after that block.
    """
    items = []
    sheet_cat_id = ElementId(BuiltInCategory.OST_Sheets)

    # Iterate each token in desired order and append a block
    for i, token in enumerate(naming_items):
        # Set per-block separator: use GLOBAL except after the last block
        sep = global_sep if i < len(naming_items) - 1 else ""

        if token == "SHEET_NUMBER":
            from Autodesk.Revit.DB import BuiltInParameter
            items.append(
                build_table_cell_param_data_for_builtin(BuiltInParameter.SHEET_NUMBER, sheet_cat_id, separator=sep)
            )
        elif token == "SHEET_NAME":
            from Autodesk.Revit.DB import BuiltInParameter
            items.append(
                build_table_cell_param_data_for_builtin(BuiltInParameter.SHEET_NAME, sheet_cat_id, separator=sep)
            )
        else:
            # Treat as a Sheet-bound project/shared parameter by its display name
            peid = find_parameter_element_id_by_name(token, sheet_cat_id)
            if peid and peid.IntegerValue != -1:
                data = TableCellCombinedParameterData.Create()
                data.ParamId = peid
                data.CategoryId = sheet_cat_id
                data.Separator = sep
                items.append(data)
            else:
                # If not found, silently skip (or you could raise/log)
                output_window.print_md(u"- ⚠️ Parameter not found on Sheets: `{}` (skipped)".format(token))

    return items

def build_pdf_options():
    """
    Configure PDFExportOptions to:
      - export per-sheet files (Combine=False)
      - set quality / graphics flags
      - set naming rule (the magic!)
    """
    opt = PDFExportOptions()

    # --- Per-file export ---
    opt.Combine = False  # one file per view/sheet

    # --- Quality (DPI/tessellation) ---
    opt.ExportQuality = PDF_QUALITY  # e.g., DPI300

    # --- Graphics / page placement & visibility ---
    opt.AlwaysUseRaster = ALWAYS_USE_RASTER
    opt.HideCropBoundaries = HIDE_CROP_BOUNDARIES
    opt.HideUnreferencedViewTags = HIDE_UNREFERENCED_TAGS
    opt.MaskCoincidentLines = MASK_COINCIDENT_LINES
    opt.ReplaceHalftoneWithThinLines = REPLACE_HALFTONE_WITH_THIN_LINES
    opt.ViewLinksInBlue = VIEW_LINKS_IN_BLUE
    opt.StopOnError = STOP_ON_ERROR

    # --- Fit to page / Zoom ---
    from Autodesk.Revit.DB import ZoomType
    if FIT_TO_PAGE:
        opt.ZoomType = ZoomType.FitToPage
    else:
        opt.ZoomType = ZoomType.Zoom
        opt.ZoomPercentage = ZOOM_PERCENT

    # --- Naming rule ---
    rule = build_naming_rule_for_sheets(NAMING_ORDER, GLOBAL_SEPARATOR)

    # Validate and apply naming rule (mirrors Revit UI behavior)
    # IsValidNamingRule returns False for empty rule or illegal characters like \ / : * ? " < > |
    from System.Collections.Generic import List as CsList
    cs_rule = CsList[TableCellCombinedParameterData]()
    for r in rule:
        cs_rule.Add(r)

    # Sanity check (optional)
    if not PDFExportOptions.IsValidNamingRule(cs_rule):
        raise Exception("Invalid PDF naming rule. Check your tokens and separators.")

    # You can also prepend/append global prefix/suffix by adding tiny 0-width blocks, but
    # simpler is to post-process filenames. Here we use Prefix/Suffix on the first/last item.
    if GLOBAL_PREFIX and cs_rule.Count > 0:
        first = cs_rule[0]
        first.Prefix = (GLOBAL_PREFIX or "") + (first.Prefix or "")
    if GLOBAL_SUFFIX and cs_rule.Count > 0:
        last = cs_rule[cs_rule.Count - 1]
        last.Suffix = (last.Suffix or "") + (GLOBAL_SUFFIX or "")

    opt.SetNamingRule(cs_rule)
    return opt


#____________________________________________________________________ MAIN

def main():
    out_folder = get_output_folder()
    sheets = collect_sheets()
    if not sheets:
        output_window.print_md("No sheets to export with current filters.")
        return

    # Build options
    opt = build_pdf_options()

    # Gather ElementIds for export
    id_list = [s.Id for s in sheets]

    # Perform export (native PDF). Revit will name files using our naming rule.
    # API signature: Document.Export(string folder, IList<ElementId> views, PDFExportOptions options)
    # (FileName is ignored because Combine=False; naming rule drives per-sheet filenames.)
    t = Transaction(doc, "Export Sheets to PDF (Native)")
    t.Start()
    try:
        doc.Export(out_folder, id_list, opt)
    finally:
        t.Commit()

    output_window.print_md(u"**Export complete** → `{}`  \nExported {} sheet PDF(s).".format(out_folder, len(id_list)))


#_____________________________________________________________________ 🏃‍➡️ RUN 
if __name__ == "__main__":
    main()
