# -*- coding: utf-8 -*-
__title__     = "Collect Family \nParameters"
__version__   = 'Version = v0.1'
__doc__       = """Version = v0.1
Date    = 02.19.2026
______________________________________________________________
Description:
-> Export all loadable families and their Family Parameters to JSON.
______________________________________________________________
How-to:
-> 
______________________________________________________________
Last update:
- [02.09.2026] - v0.1 BETA RELEASE
______________________________________________________________
Author: Kyle Guggenheim"""

#____________________________________________________________________ IMPORTS (SYSTEM)
import os
import json
import clr

#____________________________________________________________________ IMPORTS (AUTODESK)
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Family,
    InternalDefinition,
    BuiltInParameter,
    IFamilyLoadOptions
)

from Autodesk.Revit.DB import Transaction

#____________________________________________________________________ IMPORTS (PYREVIT)
from pyrevit import revit, DB, UI, script
from pyrevit.script import output
from pyrevit import forms


#____________________________________________________________________ VARIABLES
doc = revit.doc


log_status = ""
action = "Collect Family Parameters"

output_window = output.get_output()
"""Output window for displaying results."""




# ----------------------------
# Utilities
# ----------------------------

def safe_str(x):
    try:
        return None if x is None else str(x)
    except:
        return None


def get_param_value_type_string(fp):
    d = fp.Definition
    try:
        dt = d.GetDataType()
        if dt:
            return safe_str(dt.TypeId) if hasattr(dt, "TypeId") else safe_str(dt)
    except:
        pass
    try:
        return safe_str(d.ParameterType)
    except:
        return ""


def get_formula(fm, fp):
    try:
        return fm.GetFormula(fp) or ""
    except:
        return ""


def is_builtin(defn):
    try:
        if isinstance(defn, InternalDefinition):
            bip = defn.BuiltInParameter
            return bip != BuiltInParameter.INVALID
    except:
        pass
    return False


def param_type_label(fp):
    # Shared vs Built-in vs Family
    try:
        if fp.IsShared:
            return "Shared Parameter"
    except:
        pass

    try:
        if is_builtin(fp.Definition):
            return "Built In Parameter"
    except:
        pass

    return "Family Parameter"


# ----------------------------
# Data Models (WPF binding needs attributes)
# ----------------------------

# from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs

# class NotifyBase(INotifyPropertyChanged):
#     def __init__(self):
#         self.PropertyChanged = None

#     def _raise(self, prop_name):
#         if self.PropertyChanged:
#             self.PropertyChanged(self, PropertyChangedEventArgs(prop_name))


class FamilyRow(object):
    def __init__(self, fam_name, cat_name):
        self.IsCheckedFamily = False
        # self.IsChecked = False
        self.FamilyName = fam_name
        self.CategoryName = cat_name

    # @property
    # def IsChecked(self):
    #     return self.IsCheckedFamily

    # @IsChecked.setter
    # def IsChecked(self, val):
    #     self.IsCheckedFamily = bool(val)
    #     self._raise("IsChecked")


class ParamRow(object):
    def __init__(self, dct):
        self.IsCheckedParameter = False
        # self.IsChecked = False
        self.Name = dct.get("name") or ""
        self.ParameterTypeLabel = dct.get("parameter_type_label") or ""
        self.ParamValueType = dct.get("param_value_type") or ""
        self.Group = dct.get("group") or ""
        self.InstanceTypeLabel = "Instance" if dct.get("is_instance") else "Type"
        self.Formula = dct.get("formula") or ""

        # Keep a key used for deletion lookup
        self._key_name = self.Name
        self._key_group = self.Group

    # @property
    # def IsChecked(self):
    #     return self.IsCheckedParameter

    # @IsChecked.setter
    # def IsChecked(self, val):
    #     self.IsCheckedParameter = bool(val)
    #     self._raise("IsChecked")


# ----------------------------
# Collect Families + Parameters (definitions + type label)
# ----------------------------

def read_family_parameters(fam_doc):
    fm = fam_doc.FamilyManager
    out = []

    for fp in fm.Parameters:
        try:
            d = fp.Definition
            out.append({
                "name": safe_str(d.Name) or "",
                "group": safe_str(d.ParameterGroup) or "",
                "param_value_type": get_param_value_type_string(fp),
                "is_instance": bool(fp.IsInstance),
                "formula": get_formula(fm, fp),
                "parameter_type_label": param_type_label(fp)
            })
        except:
            continue

    out.sort(key=lambda x: ((x.get("group") or ""), (x.get("name") or "")))
    return out


def collect_family_list(document):
    """Fast: only list families + category (no EditFamily)."""
    rows = []
    fam_elems = FilteredElementCollector(document).OfClass(Family).ToElements()
    for fam in fam_elems:
        try:
            if fam.IsInPlace:
                continue
        except:
            continue

        fam_name = safe_str(fam.Name) or "<Unnamed>"
        cat_name = ""
        try:
            if fam.FamilyCategory:
                cat_name = safe_str(fam.FamilyCategory.Name) or ""
        except:
            pass

        rows.append(FamilyRow(fam_name, cat_name))

    rows.sort(key=lambda r: (r.FamilyName or "").lower())
    return rows


# ----------------------------
# Family load options (overwrite)
# ----------------------------

class OverwriteLoadOptions(IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = True
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        overwriteParameterValues.Value = True
        return True


# ----------------------------
# Window
# ----------------------------

class FamiliesWindow(forms.WPFWindow):
    def __init__(self, xaml_path):
        forms.WPFWindow.__init__(self, xaml_path)

        # data
        self._families = collect_family_list(doc)
        self._current_family_name = None
        self._current_param_dicts = []   # raw dicts
        self._current_param_rows = []    # bound rows

        # bind
        self.DgFamilies.ItemsSource = self._families
        self.DgParams.ItemsSource = []
        self.LblFamiliesCount.Text = str(len(self._families))
        self.LblSelectedFamily.Text = "(no family selected)"
        # self.LblStatus.Text = "Ready"

        # events
        self.DgFamilies.SelectionChanged += self.on_family_selection_changed
        self.TxtFamilySearch.TextChanged += self.on_family_search_changed
        self.TxtParamFilter.TextChanged += self.on_param_filter_changed
        self.CmbParamTypeFilter.SelectionChanged += self.on_param_filter_changed

        self.BtnRefresh.Click += self.on_refresh
        # self.BtnClear.Click += self.on_clear

        # (3) modify dropdown -> delete selected
        self.CmbModify.SelectionChanged += self.on_modify_action_changed

        # optional: keep Apply as “Export JSON” if you want
        # self.BtnApply.Click += self.on_export_json

    # ------------------------
    # UI helpers
    # ------------------------

    def _selected_param_type_filter(self):
        try:
            item = self.CmbParamTypeFilter.SelectedItem
            if item is None:
                return "All"
            # ComboBoxItem.Content
            return safe_str(item.Content) or "All"
        except:
            return "All"

    def _apply_param_filters(self):
        name_needle = (self.TxtParamFilter.Text or "").strip().lower()
        type_filter = self._selected_param_type_filter()

        rows = []
        for d in self._current_param_dicts:
            # type filter
            if type_filter != "All":
                if (d.get("parameter_type_label") or "") != type_filter:
                    continue

            # name filter
            if name_needle:
                if name_needle not in (d.get("name") or "").lower():
                    continue

            rows.append(ParamRow(d))

        self._current_param_rows = rows
        self.DgParams.ItemsSource = rows
        # self.LblStatus.Text = "Parameters: {}".format(len(rows))

    # ------------------------
    # Data loaders
    # ------------------------

    def _load_params_for_family(self, fam_name):
        self._current_family_name = fam_name
        self.LblSelectedFamily.Text = fam_name
        # self.LblStatus.Text = "Loading parameters..."

        fam_elem = None
        for f in FilteredElementCollector(doc).OfClass(Family).ToElements():
            if safe_str(f.Name) == fam_name:
                fam_elem = f
                break

        if fam_elem is None:
            self._current_param_dicts = []
            self._apply_param_filters()
            # self.LblStatus.Text = "Family not found in document."
            return

        try:
            fam_doc = doc.EditFamily(fam_elem)
            try:
                self._current_param_dicts = read_family_parameters(fam_doc)
            finally:
                fam_doc.Close(False)
        except:
            self._current_param_dicts = []

        self._apply_param_filters()

    # ------------------------
    # Event handlers
    # ------------------------

    def on_family_selection_changed(self, sender, args):
        sel = self.DgFamilies.SelectedItem
        if not sel:
            return
        self._load_params_for_family(sel.FamilyName)

    def on_family_search_changed(self, sender, args):
        needle = (self.TxtFamilySearch.Text or "").strip().lower()
        if not needle:
            filtered = self._families
        else:
            filtered = [r for r in self._families
                        if needle in (r.FamilyName or "").lower()
                        or needle in (r.CategoryName or "").lower()]
        self.DgFamilies.ItemsSource = filtered
        self.LblFamiliesCount.Text = str(len(filtered))

    def on_param_filter_changed(self, sender, args):
        if self._current_family_name:
            self._apply_param_filters()

    def on_refresh(self, sender, args):
        # self.LblStatus.Text = "Refreshing..."
        self._families = collect_family_list(doc)
        self.DgFamilies.ItemsSource = self._families
        self.LblFamiliesCount.Text = str(len(self._families))

        self._current_family_name = None
        self._current_param_dicts = []
        self.DgParams.ItemsSource = []
        self.LblSelectedFamily.Text = "(no family selected)"
        # self.LblStatus.Text = "Ready"

    # def on_clear(self, sender, args):
    #     self.TxtFamilySearch.Text = ""
    #     self.TxtParamFilter.Text = ""
    #     self.DgFamilies.SelectedItem = None
    #     self.DgParams.ItemsSource = []
    #     self._current_family_name = None
    #     self._current_param_dicts = []
    #     self.LblSelectedFamily.Text = "(no family selected)"
    #     self.LblStatus.Text = "Ready"

    # ------------------------
    # (3) Delete selected parameters
    # ------------------------

    def on_modify_action_changed(self, sender, args):
        # Identify selected action
        try:
            item = self.CmbModify.SelectedItem
            action = safe_str(item.Content) if item else ""
        except:
            return

        if action != "Delete Selected Parameters":
            return

        # Reset dropdown back to label item
        try:
            self.CmbModify.SelectedIndex = 0
        except:
            pass

        # Preconditions
        if not self._current_family_name:
            forms.alert("Select a family first.", title="Delete Parameters")
            return

        selected = list(self.DgParams.SelectedItems) if self.DgParams.SelectedItems else []
        if not selected:
            forms.alert("Select one or more parameters in the right grid.", title="Delete Parameters")
            return

        names_to_delete = sorted(set([s.Name for s in selected if getattr(s, "Name", None)]))
        if not names_to_delete:
            forms.alert("No valid parameters selected.", title="Delete Parameters")
            return

        msg = "Delete {} parameter(s) from family:\n\n{}\n\nBuilt-in parameters will be skipped.".format(
            len(names_to_delete), self._current_family_name
        )
        if not forms.alert(msg, title="Confirm Delete", ok=False, yes=True, no=True):
            return

        # Find the Family element in the project
        fam_elem = None
        for f in FilteredElementCollector(doc).OfClass(Family).ToElements():
            if safe_str(f.Name) == self._current_family_name:
                fam_elem = f
                break

        if fam_elem is None:
            forms.alert("Family not found in the document.", title="Delete Parameters")
            return

        deleted = 0
        skipped_builtin = 0
        not_found = 0
        failed = 0
        fail_messages = []

        fam_doc = None
        try:
            fam_doc = doc.EditFamily(fam_elem)
            fm = fam_doc.FamilyManager

            # Build name->FamilyParameter list
            by_name = {}
            for fp in fm.Parameters:
                try:
                    nm = safe_str(fp.Definition.Name) or ""
                    by_name.setdefault(nm, []).append(fp)
                except:
                    continue

            # IMPORTANT: real DB.Transaction on the FAMILY DOCUMENT
            t = Transaction(fam_doc, "Delete Family Parameters")
            t.Start()
            try:
                for nm in names_to_delete:
                    fps = by_name.get(nm)
                    if not fps:
                        not_found += 1
                        continue

                    for fp in fps:
                        try:
                            # Skip built-ins (cannot be removed)
                            if is_builtin(fp.Definition):
                                skipped_builtin += 1
                                continue

                            fm.RemoveParameter(fp)
                            deleted += 1
                        except Exception as ex:
                            failed += 1
                            fail_messages.append("{}: {}".format(nm, ex))
                t.Commit()
            except Exception as ex:
                try:
                    t.RollBack()
                except:
                    pass
                raise

            # Reload into project and VERIFY it worked
            try:
                ok = fam_doc.LoadFamily(doc, OverwriteLoadOptions())
                if not ok:
                    forms.alert(
                        "Family reload failed (LoadFamily returned False).\n"
                        "The family may be in use, read-only, or blocked by shared family rules.",
                        title="Delete Parameters"
                    )
            except Exception as ex:
                forms.alert("Family reload threw an exception:\n{}".format(ex), title="Delete Parameters")

        finally:
            if fam_doc:
                try:
                    fam_doc.Close(False)
                except:
                    pass

        # Refresh UI from project state
        self._load_params_for_family(self._current_family_name)

        # Report
        report = "Deleted: {}\nSkipped built-in: {}\nNot found: {}\nFailed: {}".format(
            deleted, skipped_builtin, not_found, failed
        )
        if deleted == 0:
            # If nothing changed, show details to diagnose quickly
            if fail_messages:
                report += "\n\nFailures (first 10):\n" + "\n".join(fail_messages[:10])
            forms.alert(report, title="Delete Parameters Result")
        else:
            # self.LblStatus.Text = report.replace("\n", " | ")
            pass

    # ------------------------
    # Optional: export JSON (Apply button)
    # ------------------------

    def on_export_json(self, sender, args):
        # exports current family’s filtered list (or all if none)
        downloads = os.path.join(os.path.expanduser("~"), "Downloads")
        fp = os.path.join(downloads, "Revit_Family_Parameters.json")
        try:
            import json
            data = {}
            if self._current_family_name:
                data[self._current_family_name] = self._current_param_dicts
            else:
                data["_note"] = "Select a family to export its parameters."
            with open(fp, "w") as f:
                json.dump(data, f, indent=2)
            # self.LblStatus.Text = "Exported JSON to Downloads"
        except:
            # self.LblStatus.Text = "Export failed."
            pass


# Run
xaml_path = os.path.join(os.path.dirname(__file__), "FamiliesWindow.xaml")
FamiliesWindow(xaml_path).ShowDialog()



log_status = "Success"
#______________________________________________________ LOG ACTION
def log_action(action, log_status):
    """Log action to user JSON log file."""
    import os, json, time
    from pyrevit import revit

    doc = revit.doc
    # doc_path = doc.PathName or "<Untitled>"
    doc_path = doc.PathName if doc.PathName else "<Untitled>"

    doc_title = doc.Title
    version_build = doc.Application.VersionBuild
    version_number = doc.Application.VersionNumber
    username = doc.Application.Username
    action = action

    # json log location
    # \FFE Inc\FFE Revit Users - Documents\00-General\Revit_Add-Ins\FFE-pyRevit\Logs
    log_dir = os.path.join(os.path.expanduser("~"), "FFE Inc", "FFE Revit Users - Documents", "00-General", "Revit_Add-Ins", "FFE-pyRevit", "Logs")
    log_file = os.path.join(log_dir, username + "_revit_log.json")

    dataEntry = {
        "datetime": time.strftime("%Y-%m-%d %H:%M:%S"),
        "username": username,
        "doc_title": doc_title,
        "doc_path": doc_path,
        "revit_version_number": version_number,
        "revit_build": version_build,
        "action": action,
        "status": log_status
    }

    # Function to write JSON data
    def write_json(dataEntry, filename=log_file):
        with open(filename,'r+') as file:
            file_data = json.load(file)                 # First we load existing data into a dict.   
            file_data['action'].append(dataEntry)       # Join new_data with file_data inside emp_details
            file.seek(0)                                # Sets file's current position at offset.
            json.dump(file_data, file, indent = 4)      # convert back to json.


    # Check if log file exists, if not create it
    logcheck = False
    if not os.path.exists(log_file):
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        with open(log_file, 'w') as file:    
            file.write('{"action": []}')                # create json structure
        
        # output_window.print_md("### **Created log file:** `{}`".format(log_file))

    # If it does exist, write to it
    # Check if "action" key exists, if not create it
    with open(log_file,'r+') as file:
        file_data = json.load(file)
        if 'action' not in file_data:
            file_data['action'] = []
            file.seek(0)
            json.dump(file_data, file, indent = 4)

    try:
        write_json(dataEntry)
        logcheck = True
        # output_window.print_md("### **Logged sync to JSON:** `{}`".format(log_file))
    except Exception as e:
        logcheck = False

    return dataEntry

log_action(action, log_status)