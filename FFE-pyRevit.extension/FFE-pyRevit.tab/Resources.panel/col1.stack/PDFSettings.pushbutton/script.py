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
import clr

#____________________________________________________________________ IMPORTS (AUTODESK)
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import (
    Transaction, ElementId, ViewSheet, BuiltInCategory,
    FilteredElementCollector,  # not used for sheets; only for ParameterElement lookup
    ParameterElement,
    # Print / sheet set
    PrintManager, ViewSheetSetting,
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

doc = revit.doc


# USER SETTINGS — EDIT THESE
#____________________________________________________________________ VARIABLES

# 1) Output folder: defaults to Downloads\<ProjectName>\YYYY-MM-DD_HHMM
USE_DEFAULT_DOWNLOADS = True
CUSTOM_OUTPUT_FOLDER = r""        # e.g., r"C:\_PDFs"; leave blank to auto-build

# 2) PDF export quality (DPI)
PDF_QUALITY = PDFExportQualityType.DPI300  # options include DPI144, DPI300, DPI600, etc.

# 3) Graphics / visibility flags
ALWAYS_USE_RASTER = False                 # True forces raster; False prefers vector
HIDE_CROP_BOUNDARIES = False
HIDE_UNREFERENCED_TAGS = False
MASK_COINCIDENT_LINES = False
REPLACE_HALFTONE_WITH_THIN_LINES = False
VIEW_LINKS_IN_BLUE = False
STOP_ON_ERROR = True                      # Stop export if any sheet fails

# 4) Fit / Zoom behavior
FIT_TO_PAGE = True                        # True = Fit to page; False = use ZOOM_PERCENT
ZOOM_PERCENT = 100                        # Only used if FIT_TO_PAGE is False (valid: 10–500)

# 5) Filename Naming Rule (order of fields; no extension)
#    Built-in tokens: "SHEET_NUMBER", "SHEET_NAME"
#    Add any *Sheet-bound* shared/project parameter display names (exactly as seen in Revit).
NAMING_ORDER = [
    "SHEET_NUMBER",
    "SHEET_NAME",
    # Example shared parameters bound to Sheets:
    # "Discipline",
    # "Package",
    # "Submission",
]

# 6) Global filename separators and optional fixed affixes
GLOBAL_SEPARATOR = " - "                  # Placed *between* blocks in NAMING_ORDER
GLOBAL_PREFIX = ""                        # Optional prefix for entire filename
GLOBAL_SUFFIX = ""                        # Optional suffix for entire filename


#____________________________________________________________________ FUNCTIONS

# ---------------- Helpers: filesystem ----------------
def _downloads_folder():
    """Return user's Downloads folder path (Windows/macOS)."""
    try:
        from System import Environment
        from System.Environment import SpecialFolder
        profile = Environment.GetFolderPath(SpecialFolder.UserProfile)
        return os.path.join(profile, "Downloads")
    except:
        return os.path.expanduser("~/Downloads")

def _output_folder():
    """Build the output folder path per settings."""
    if CUSTOM_OUTPUT_FOLDER.strip():
        base = CUSTOM_OUTPUT_FOLDER
    elif USE_DEFAULT_DOWNLOADS:
        base = _downloads_folder()
    else:
        base = os.path.expanduser("~/Desktop")
    proj_name = doc.Title.replace(".rvt", "")
    ts = datetime.datetime.now().strftime("%Y-%m-%d_%H%M")
    folder = os.path.join(base, proj_name, ts)
    if not os.path.exists(folder):
        os.makedirs(folder)
    return folder


# ---------------- Helpers: naming rule ----------------
def _build_table_cell_param_data_for_builtin(bip_enum, category_id, prefix="", suffix="", separator=""):
    """Create a block for a built-in parameter (e.g., SHEET_NUMBER / SHEET_NAME)."""
    data = TableCellCombinedParameterData.Create()
    data.ParamId = ElementId(bip_enum)   # wrap BuiltInParameter as ElementId
    data.CategoryId = category_id
    data.Prefix = prefix or ""
    data.Suffix = suffix or ""
    data.Separator = separator or ""
    return data

def _find_parameter_element_id_by_name(param_display_name, category_id):
    """
    Resolve a *Sheet-bound* Project/Shared parameter ElementId by display name.
    We scan ParameterElement instances (faster & reliable across versions)
    and verify it's bound to OST_Sheets.
    """
    # Gather all ParameterElement and pick the one whose Definition.Name matches.
    # (We avoid document-wide sheet collection; this does not touch sheets.)
    for pe in FilteredElementCollector(doc).OfClass(ParameterElement):
        try:
            defi = pe.GetDefinition()
            if not defi or defi.Name != param_display_name:
                continue
            # We still need to verify it's bound to Sheets:
            # ParameterElement.GetCategories() is available in newer versions; else skip verification.
            cats = getattr(pe, "GetCategories", None)
            if callable(cats):
                cs = cats()
                if cs:
                    for c in cs:
                        if c.Id == category_id:
                            return pe.Id
                    continue
            # If we can't verify via categories (older API), assume OK and return.
            return pe.Id
        except:
            pass
    return ElementId.InvalidElementId

def _build_naming_rule_for_sheets(naming_items, global_sep):
    """
    Build IList<TableCellCombinedParameterData> to mirror Revit's PDF Naming Rules UI.
    Supports:
      - "SHEET_NUMBER" and "SHEET_NAME" built-ins
      - Any Sheet-bound project/shared parameter display names
    """
    from Autodesk.Revit.DB import BuiltInParameter
    from System.Collections.Generic import List as CsList

    items = CsList[TableCellCombinedParameterData]()
    sheet_cat_id = ElementId(BuiltInCategory.OST_Sheets)

    for i, token in enumerate(naming_items):
        # per-block separator (nothing after the final block)
        sep = global_sep if i < len(naming_items) - 1 else ""

        if token == "SHEET_NUMBER":
            items.Add(_build_table_cell_param_data_for_builtin(BuiltInParameter.SHEET_NUMBER, sheet_cat_id, separator=sep))
        elif token == "SHEET_NAME":
            items.Add(_build_table_cell_param_data_for_builtin(BuiltInParameter.SHEET_NAME, sheet_cat_id, separator=sep))
        else:
            peid = _find_parameter_element_id_by_name(token, sheet_cat_id)
            if peid and peid.IntegerValue != -1:
                d = TableCellCombinedParameterData.Create()
                d.ParamId = peid
                d.CategoryId = sheet_cat_id
                d.Separator = sep
                items.Add(d)
            else:
                # Minimal console output; do not reveal sheet names or contents.
                output_window.print_md(u"⚠️ Skipped missing Sheet parameter: `{}`".format(token))

    # Optional global prefix/suffix via first/last block
    if GLOBAL_PREFIX and items.Count > 0:
        f = items[0]
        f.Prefix = (GLOBAL_PREFIX or "") + (f.Prefix or "")
    if GLOBAL_SUFFIX and items.Count > 0:
        l = items[items.Count - 1]
        l.Suffix = (l.Suffix or "") + (GLOBAL_SUFFIX or "")

    return items

def _pdf_options():
    """Configure PDFExportOptions: per-file, quality, graphics, zoom, naming rule."""
    from Autodesk.Revit.DB import ZoomType, PDFExportOptions

    opt = PDFExportOptions()

    # --- One file per sheet ---
    opt.Combine = False

    # --- Quality ---
    opt.ExportQuality = PDF_QUALITY

    # --- Graphics / visibility flags ---
    opt.AlwaysUseRaster = ALWAYS_USE_RASTER
    opt.HideCropBoundaries = HIDE_CROP_BOUNDARIES
    opt.HideUnreferencedViewTags = HIDE_UNREFERENCED_TAGS
    opt.MaskCoincidentLines = MASK_COINCIDENT_LINES
    opt.ReplaceHalftoneWithThinLines = REPLACE_HALFTONE_WITH_THIN_LINES
    opt.ViewLinksInBlue = VIEW_LINKS_IN_BLUE
    opt.StopOnError = STOP_ON_ERROR

    # --- Fit/Zoom ---
    if FIT_TO_PAGE:
        opt.ZoomType = ZoomType.FitToPage
    else:
        opt.ZoomType = ZoomType.Zoom
        opt.ZoomPercentage = ZOOM_PERCENT

    # --- Naming Rule (required for per-file naming when Combine=False) ---
    rule = _build_naming_rule_for_sheets(NAMING_ORDER, GLOBAL_SEPARATOR)

    # Validate & apply (e.g., catches illegal characters \ / : * ? " < > |)
    if not PDFExportOptions.IsValidNamingRule(rule):
        raise Exception("Invalid PDF naming rule. Check NAMING_ORDER, separators, and illegal characters.")
    opt.SetNamingRule(rule)

    return opt

#____________________________________________________________________ MAIN

def _get_current_sheet_set_ids_only():
    """
    Return a list of ElementId for *sheets only* contained in the *current* View/Sheet Set
    from Revit's Print dialog. This does not run a document-wide sheet collector.
    """
    pm = doc.PrintManager
    vss = pm.ViewSheetSetting
    curr = vss.CurrentViewSheetSet  # "In-Session" or currently active saved set

    if curr is None:
        return []

    # CurrentViewSheetSet.Views is a ViewSet iterable (could include views). Filter to sheets only.
    ids = []
    it = curr.Views.ForwardIterator()
    it.Reset()
    while it.MoveNext():
        v = it.Current
        if isinstance(v, ViewSheet):
            ids.append(v.Id)
    return ids

def main():
    out_folder = _output_folder()
    ids = _get_current_sheet_set_ids_only()
    if not ids:
        output_window.print_md(
            "No sheets found in the **current View/Sheet Set**.\n"
            "- Open **Print… → Select…**, choose or create a *View/Sheet Set* (sheets only), then run this command."
        )
        return

    opt = _pdf_options()

    # Native Export > PDF (file names are produced by the naming rule; we do not log sheet names)
    t = Transaction(doc, "Export Sheets to PDF (Native via Current Set)")
    t.Start()
    try:
        doc.Export(out_folder, ids, opt)
    finally:
        t.Commit()

    output_window.print_md(u"**Export complete** → `{}`  \nExported {} PDF file(s).".format(out_folder, len(ids)))


#_____________________________________________________________________ 🏃‍➡️ RUN 
if __name__ == "__main__":
    main()