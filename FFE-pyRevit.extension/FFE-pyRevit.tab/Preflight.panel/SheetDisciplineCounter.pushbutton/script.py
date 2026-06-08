# -*- coding: utf-8 -*-
__title__ = "Sheet Discipline \nOrder"
__version__ = "Version = v0.2"
__persistentengine__ = True
__doc__ = """Version = v0.2
Date    = 06.08.2026
______________________________________________________________
Description:
-> Opens a WebView2 sheet order manager.
-> Lists sheets grouped by FFE_Sheet_Discipline.
-> Writes FFE_Sheet_Order from the saved drag order.
______________________________________________________________
How-to:
-> Press Button
-> Drag sheets inside each discipline group
-> Click Save Order
______________________________________________________________
Last update:
- [05.13.2026] - v0.1 BETA RELEASE
- [06.08.2026] - v0.2 Added WebView2 drag ordering UI
______________________________________________________________
Author: Kyle Guggenheim"""


# ____________________________________________________________________ IMPORTS (SYSTEM)
import json
import os
import re
import traceback

import clr
clr.AddReference("System")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System import Uri
from System.Windows import ResizeMode, Thickness, Visibility, Window, WindowStartupLocation
from System.Windows.Controls import Grid, TextBlock
from System.Windows.Interop import WindowInteropHelper


# ____________________________________________________________________ IMPORTS (AUTODESK)
from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
    StorageType,
    Transaction,
    ViewSheet,
)
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler


# ____________________________________________________________________ IMPORTS (PYREVIT)
from pyrevit import forms, script


# ____________________________________________________________________ VARIABLES
revit_app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

PATH_SCRIPT = os.path.dirname(__file__)
PATH_SUPPORT = os.path.join(PATH_SCRIPT, "support")
PATH_INDEX = os.path.join(PATH_SUPPORT, "index.html")

DISCIPLINE_PARAM = "FFE_Sheet_Discipline"
ORDER_PARAM = "FFE_Sheet_Order"

OUTPUT = script.get_output()
LOGGER = script.get_logger()
WINDOW_REFS = []


# ____________________________________________________________________ BASIC HELPERS
def safe_str(value):
    if value is None:
        return ""
    try:
        return str(value)
    except:
        try:
            return value.ToString()
        except:
            return ""


def json_dumps(value):
    return json.dumps(value, ensure_ascii=True)


def make_file_uri(path):
    absolute_path = os.path.abspath(path).replace("\\", "/")
    return Uri("file:///" + absolute_path.replace(" ", "%20"))


def element_id_value(element_id):
    try:
        return int(element_id.IntegerValue)
    except:
        try:
            return int(element_id.Value)
        except:
            return int(str(element_id))


def natural_key(text):
    parts = re.split(r"(\d+)", safe_str(text))
    key = []
    for part in parts:
        key.append(int(part) if part.isdigit() else part.lower())
    return key


def parse_int_maybe(value):
    if value is None:
        return None

    if isinstance(value, int):
        return value

    text = safe_str(value).strip()
    match = re.match(r"^-?\d+", text)
    if not match:
        return None

    try:
        return int(match.group(0))
    except:
        return None


def get_revit_install_dir():
    try:
        app_path = revit_app.ApplicationPath
        if app_path:
            return os.path.dirname(app_path)
    except:
        pass

    version = safe_str(getattr(revit_app, "VersionNumber", "2026")) or "2026"
    return os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"), "Autodesk", "Revit {0}".format(version))


def load_webview2_types():
    revit_dir = get_revit_install_dir()
    dll_paths = [
        os.path.join(revit_dir, "Microsoft.Web.WebView2.Core.dll"),
        os.path.join(revit_dir, "Microsoft.Web.WebView2.Wpf.dll"),
    ]

    for dll_path in dll_paths:
        if not os.path.exists(dll_path):
            raise Exception("WebView2 assembly was not found: {0}".format(dll_path))
        try:
            clr.AddReferenceToFileAndPath(dll_path)
        except AttributeError:
            clr.AddReference(dll_path)

    path_parts = os.environ.get("PATH", "").split(os.pathsep)
    if revit_dir not in path_parts:
        os.environ["PATH"] = revit_dir + os.pathsep + os.environ.get("PATH", "")

    from Microsoft.Web.WebView2.Wpf import CoreWebView2CreationProperties, WebView2
    return WebView2, CoreWebView2CreationProperties


def get_webview_user_data_folder():
    base_folder = os.environ.get("LOCALAPPDATA")
    if not base_folder:
        base_folder = os.path.join(os.path.expanduser("~"), "AppData", "Local")

    user_data_folder = os.path.join(base_folder, "FFE-pyRevit", "SheetDisciplineOrderWebView2")
    if not os.path.exists(user_data_folder):
        os.makedirs(user_data_folder)
    return user_data_folder


# ____________________________________________________________________ REVIT DATA HELPERS
def parameter_exists_on_category(current_doc, param_name, built_in_category):
    bindings = current_doc.ParameterBindings
    iterator = bindings.ForwardIterator()
    iterator.Reset()

    category_id = None
    try:
        category_id = current_doc.Settings.Categories.get_Item(built_in_category).Id
    except:
        pass

    while iterator.MoveNext():
        definition = iterator.Key
        if safe_str(definition.Name) != param_name:
            continue

        if category_id is None:
            return True

        try:
            binding = iterator.Current
            for category in binding.Categories:
                if element_id_value(category.Id) == element_id_value(category_id):
                    return True
        except:
            return True

    return False


def get_param(element, param_name):
    try:
        return element.LookupParameter(param_name)
    except:
        return None


def param_has_value(param):
    if param is None:
        return False
    try:
        return bool(param.HasValue)
    except:
        return True


def get_param_text(element, param_name):
    param = get_param(element, param_name)
    if param is None or not param_has_value(param):
        return ""

    try:
        storage_type = param.StorageType
        if storage_type == StorageType.String:
            return safe_str(param.AsString())
        if storage_type == StorageType.Integer:
            return safe_str(param.AsInteger())
        if storage_type == StorageType.Double:
            return safe_str(param.AsDouble())
        if storage_type == StorageType.ElementId:
            return safe_str(element_id_value(param.AsElementId()))
    except:
        pass

    return ""


def get_discipline_value(sheet):
    return get_param_text(sheet, DISCIPLINE_PARAM).strip()


def get_order_number(sheet):
    return parse_int_maybe(get_param_text(sheet, ORDER_PARAM))


def get_order_display(sheet):
    return get_param_text(sheet, ORDER_PARAM).strip()


def can_write_order(sheet):
    param = get_param(sheet, ORDER_PARAM)
    if param is None:
        return False

    try:
        if param.IsReadOnly:
            return False
    except:
        pass

    return True


def discipline_label(discipline_value):
    value = safe_str(discipline_value).strip()
    return value if value else "Unassigned"


def sheet_sort_key(sheet):
    order_number = get_order_number(sheet)
    if order_number is None:
        order_number = 999999

    return (
        order_number,
        natural_key(sheet.SheetNumber),
        safe_str(sheet.Name).lower(),
        element_id_value(sheet.Id),
    )


def collect_sheets(current_doc):
    collector = (
        FilteredElementCollector(current_doc)
        .OfClass(ViewSheet)
        .WhereElementIsNotElementType()
        .ToElements()
    )

    sheets = []
    for sheet in collector:
        try:
            if getattr(sheet, "IsPlaceholder", False):
                continue
        except:
            pass

        try:
            scheduled_param = sheet.get_Parameter(BuiltInParameter.SHEET_SCHEDULED)
            if scheduled_param is None or scheduled_param.AsInteger() != 1:
                continue
        except:
            continue

        sheets.append(sheet)

    return sheets


def validate_required_parameters(current_doc):
    sheets = collect_sheets(current_doc)
    missing = []

    for param_name in [DISCIPLINE_PARAM, ORDER_PARAM]:
        exists_on_binding = parameter_exists_on_category(current_doc, param_name, BuiltInCategory.OST_Sheets)
        exists_on_sheet = False

        for sheet in sheets:
            if get_param(sheet, param_name) is not None:
                exists_on_sheet = True
                break

        if not exists_on_binding and not exists_on_sheet:
            missing.append(param_name)

    if missing:
        forms.alert(
            "Required sheet parameters are missing:\n\n{0}\n\nAdd them from Parameter Service and run the tool again.".format(
                "\n".join(["- {0}".format(item) for item in missing])
            ),
            title="Sheet Discipline Order",
            warn_icon=True
        )
        return False

    return True


def build_sheet_payload(current_doc):
    sheets = collect_sheets(current_doc)
    grouped = {}

    for sheet in sheets:
        discipline = get_discipline_value(sheet)
        if discipline not in grouped:
            grouped[discipline] = []
        grouped[discipline].append(sheet)

    for discipline in grouped:
        grouped[discipline] = sorted(grouped[discipline], key=sheet_sort_key)

    discipline_values = sorted(grouped.keys(), key=lambda value: natural_key(discipline_label(value)))
    sheet_items = []
    discipline_items = []

    for discipline in discipline_values:
        group = grouped[discipline]
        discipline_items.append({
            "value": discipline,
            "label": discipline_label(discipline),
            "sheetCount": len(group),
        })

        for index, sheet in enumerate(group, start=1):
            sheet_items.append({
                "id": element_id_value(sheet.Id),
                "uniqueId": safe_str(sheet.UniqueId),
                "discipline": discipline,
                "disciplineLabel": discipline_label(discipline),
                "sheetNumber": safe_str(sheet.SheetNumber),
                "name": safe_str(sheet.Name),
                "displayName": "{0} - {1}".format(safe_str(sheet.SheetNumber), safe_str(sheet.Name)),
                "orderValue": get_order_display(sheet),
                "orderNumber": get_order_number(sheet),
                "position": index,
                "canWriteOrder": can_write_order(sheet),
            })

    return {
        "documentTitle": safe_str(current_doc.Title),
        "disciplineParam": DISCIPLINE_PARAM,
        "orderParam": ORDER_PARAM,
        "sheetCount": len(sheet_items),
        "disciplineCount": len(discipline_items),
        "disciplines": discipline_items,
        "sheets": sheet_items,
    }


def group_current_sheets_by_discipline(current_doc):
    grouped = {}
    by_id = {}

    for sheet in collect_sheets(current_doc):
        discipline = get_discipline_value(sheet)
        if discipline not in grouped:
            grouped[discipline] = []
        grouped[discipline].append(sheet)
        by_id[element_id_value(sheet.Id)] = sheet

    for discipline in grouped:
        grouped[discipline] = sorted(grouped[discipline], key=sheet_sort_key)

    return grouped, by_id


def set_sheet_order(sheet, position, width):
    param = get_param(sheet, ORDER_PARAM)
    if param is None:
        return False, "Sheet {0} is missing {1}.".format(safe_str(sheet.SheetNumber), ORDER_PARAM)

    try:
        if param.IsReadOnly:
            return False, "Sheet {0} has a read-only {1} parameter.".format(safe_str(sheet.SheetNumber), ORDER_PARAM)
    except:
        pass

    try:
        storage_type = param.StorageType
        if storage_type == StorageType.Integer:
            param.Set(int(position))
            return True, None
        if storage_type == StorageType.String:
            param.Set(str(position).zfill(width))
            return True, None

        return False, "Sheet {0} has an unsupported {1} parameter type.".format(
            safe_str(sheet.SheetNumber),
            ORDER_PARAM
        )
    except Exception as exc:
        return False, "Sheet {0}: {1}".format(safe_str(sheet.SheetNumber), safe_str(exc))


def apply_order_payload(current_doc, payload):
    payload = payload or {}
    discipline_payloads = payload.get("disciplines") or []

    if not discipline_payloads:
        return {
            "status": "warning",
            "message": "No edited disciplines were submitted.",
            "issues": [],
            "writtenCount": 0,
            "editedCount": 0,
        }

    grouped, sheet_by_id = group_current_sheets_by_discipline(current_doc)
    issues = []
    written_count = 0
    edited_count = len(discipline_payloads)

    transaction = Transaction(current_doc, "FFE Sheet Discipline Order")
    transaction.Start()

    try:
        for discipline_payload in discipline_payloads:
            discipline = safe_str(discipline_payload.get("discipline")).strip()
            submitted_ids = discipline_payload.get("sheetIds") or []
            current_group = grouped.get(discipline, [])
            current_ids = set([element_id_value(sheet.Id) for sheet in current_group])
            ordered_sheets = []
            seen_ids = set()

            for raw_id in submitted_ids:
                try:
                    sheet_id = int(raw_id)
                except:
                    issues.append("A submitted sheet id could not be read: {0}.".format(safe_str(raw_id)))
                    continue

                if sheet_id in seen_ids:
                    continue

                sheet = sheet_by_id.get(sheet_id)
                if sheet is None:
                    issues.append("Sheet id {0} no longer exists or is no longer in the sheet list.".format(sheet_id))
                    continue

                if sheet_id not in current_ids:
                    issues.append("Sheet {0} is no longer in discipline {1} and was skipped.".format(
                        safe_str(sheet.SheetNumber),
                        discipline_label(discipline)
                    ))
                    continue

                ordered_sheets.append(sheet)
                seen_ids.add(sheet_id)

            for sheet in current_group:
                sheet_id = element_id_value(sheet.Id)
                if sheet_id not in seen_ids:
                    ordered_sheets.append(sheet)
                    seen_ids.add(sheet_id)

            if not ordered_sheets:
                issues.append("Discipline {0} has no current sheets to save.".format(discipline_label(discipline)))
                continue

            width = max(2, len(str(len(ordered_sheets))))
            for position, sheet in enumerate(ordered_sheets, start=1):
                ok, issue = set_sheet_order(sheet, position, width)
                if ok:
                    written_count += 1
                elif issue:
                    issues.append(issue)

        transaction.Commit()
    except Exception:
        try:
            transaction.RollBack()
        except:
            pass
        raise

    status = "success" if not issues else "warning"
    message = "Saved {0} sheet order value{1}.".format(
        written_count,
        "" if written_count == 1 else "s"
    )
    if issues:
        message = "{0} {1} issue{2} reported.".format(
            message,
            len(issues),
            "" if len(issues) == 1 else "s"
        )

    return {
        "status": status,
        "message": message,
        "issues": issues,
        "writtenCount": written_count,
        "editedCount": edited_count,
    }


# ____________________________________________________________________ EXTERNAL EVENT
class SheetOrderEventHandler(IExternalEventHandler):
    def __init__(self):
        self.window = None
        self.pending_action = None
        self.pending_payload = None

    def GetName(self):
        return "FFE Sheet Discipline Order Bridge"

    def queue_refresh(self):
        self.pending_action = "refresh"
        self.pending_payload = None

    def queue_save(self, payload):
        self.pending_action = "save"
        self.pending_payload = payload

    def clear_pending(self):
        self.pending_action = None
        self.pending_payload = None

    def Execute(self, uiapp):
        action = self.pending_action
        payload = self.pending_payload
        self.pending_action = None
        self.pending_payload = None

        result = {
            "status": "error",
            "message": "No Revit action was queued.",
            "issues": [],
        }
        sheet_payload = None

        try:
            current_doc = doc
            try:
                if hasattr(current_doc, "IsValidObject") and not current_doc.IsValidObject:
                    raise Exception("The document that opened this window is no longer available.")
            except Exception as doc_error:
                raise doc_error

            if action == "refresh":
                sheet_payload = build_sheet_payload(current_doc)
                result = {
                    "status": "ready",
                    "message": "Sheet data refreshed.",
                    "issues": [],
                }
            elif action == "save":
                result = apply_order_payload(current_doc, payload)
                sheet_payload = build_sheet_payload(current_doc)
            else:
                sheet_payload = build_sheet_payload(current_doc)
        except Exception as exc:
            result = {
                "status": "error",
                "message": safe_str(exc) or "Revit could not complete the sheet order action.",
                "issues": [],
            }
            try:
                LOGGER.debug(traceback.format_exc())
            except:
                pass
            try:
                sheet_payload = build_sheet_payload(current_doc)
            except:
                sheet_payload = None

        if self.window is not None:
            if sheet_payload is not None:
                self.window.send_sheet_data(sheet_payload)

            if action == "save":
                self.window.send_save_result(result)
            else:
                self.window.send_refresh_result(result)


# ____________________________________________________________________ WEBVIEW WINDOW
class SheetDisciplineOrderWindow(Window):
    def __init__(self, webview_type, creation_properties_type, initial_payload, event_handler, external_event):
        Window.__init__(self)

        # Keep window above Revit main window and show in taskbar
        try:
            WindowInteropHelper(self).Owner = __revit__.MainWindowHandle
            self.ShowInTaskbar = True
        except:
            pass

        self.initial_payload = initial_payload
        self.event_handler = event_handler
        self.external_event = external_event
        self.has_sent_initial_data = False
        self.index_uri = make_file_uri(PATH_INDEX)

        self.Title = "Sheet Discipline Order"
        self.Width = 980
        self.Height = 760
        self.MinWidth = 760
        self.MinHeight = 560
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.CanResize

        self.browser = webview_type()
        self.status_text = TextBlock()
        self.status_text.Margin = Thickness(18)
        self.status_text.Text = "Initializing Sheet Discipline Order WebView2..."

        # Configure WebView2 user data folder to avoid conflicts with other WebView2 instances in Revit
        try:
            creation_properties = creation_properties_type()
            creation_properties.UserDataFolder = get_webview_user_data_folder()
            self.browser.CreationProperties = creation_properties
        except Exception as exc:
            self.status_text.Text = "Could not configure WebView2 user data folder:\n{0}".format(exc)

        content_grid = Grid()
        content_grid.Children.Add(self.browser)
        content_grid.Children.Add(self.status_text)
        self.Content = content_grid

        self.Loaded += self.on_loaded
        self.Closed += self.on_closed
        self.browser.CoreWebView2InitializationCompleted += self.on_core_webview2_initialized
        self.browser.NavigationCompleted += self.on_navigation_completed

    def on_loaded(self, sender, args):
        """
        WebView2 can be slow to initialize, so show a status message while it's loading.
        """
        self.status_text.Visibility = Visibility.Visible
        self.status_text.Text = "Loading Sheet Discipline Order from:\n{0}".format(self.index_uri.AbsoluteUri)
        try:
            self.browser.EnsureCoreWebView2Async()
        except Exception as exc:
            self.status_text.Text = "Could not initialize WebView2:\n{0}".format(exc)

    def on_core_webview2_initialized(self, sender, args):
        """
        Called when the WebView2 core is initialized.
        """
        try:
            if args.IsSuccess:
                self.browser.CoreWebView2.WebMessageReceived += self.on_web_message_received
                self.browser.CoreWebView2.Navigate(self.index_uri.AbsoluteUri)
            else:
                message = "WebView2 initialization failed."
                try:
                    message = "{0}\n{1}".format(message, args.InitializationException.Message)
                except:
                    pass
                self.status_text.Text = message
        except:
            self.status_text.Text = "WebView2 initialized but navigation failed:\n{0}".format(traceback.format_exc())

    def on_navigation_completed(self, sender, args):
        """
        Called when the WebView2 finishes navigating to the index.html page.
        """
        try:
            if not args.IsSuccess:
                self.status_text.Visibility = Visibility.Visible
                self.status_text.Text = "Sheet Discipline Order navigation failed.\nURI: {0}\nWeb error: {1}".format(
                    self.index_uri.AbsoluteUri,
                    safe_str(args.WebErrorStatus)
                )
                return
        except:
            pass

        self.status_text.Visibility = Visibility.Collapsed
        self.send_initial_data()

    def on_closed(self, sender, args):
        try:
            self.external_event.Dispose()
        except:
            pass
        try:
            self.browser.Dispose()
        except:
            pass
        try:
            WINDOW_REFS.remove(self)
        except:
            pass

    def execute_script(self, script_text):
        try:
            self.browser.ExecuteScriptAsync(script_text)
        except:
            try:
                LOGGER.debug(traceback.format_exc())
            except:
                pass

    def call_sheet_api(self, method_name, payload):
        script_text = "window.ffeSheets && window.ffeSheets.{0}({1});".format(
            method_name,
            json_dumps(payload)
        )
        self.execute_script(script_text)

    def send_initial_data(self):
        if self.has_sent_initial_data:
            return
        self.has_sent_initial_data = True
        self.send_sheet_data(self.initial_payload)

    def send_sheet_data(self, payload):
        self.call_sheet_api("loadData", payload)

    def send_save_result(self, result):
        self.call_sheet_api("handleSaveResult", result)

    def send_refresh_result(self, result):
        self.call_sheet_api("handleRefreshResult", result)

    def send_status(self, status, message):
        self.call_sheet_api("setStatus", {
            "status": status,
            "message": message,
        })

    def raise_revit_event(self, result_method_name, error_prefix):
        try:
            self.external_event.Raise()
        except Exception as exc:
            self.event_handler.clear_pending()
            self.call_sheet_api(result_method_name, {
                "status": "error",
                "message": "{0}: {1}".format(error_prefix, safe_str(exc)),
                "issues": [],
            })

    def on_web_message_received(self, sender, args):
        raw_message = ""
        try:
            raw_message = args.TryGetWebMessageAsString()
        except:
            try:
                raw_message = args.WebMessageAsJson
            except:
                raw_message = ""

        if not raw_message:
            return

        try:
            message = json.loads(raw_message)
        except:
            return

        message_type = message.get("type")
        if message_type == "appReady":
            self.send_initial_data()
            return

        if message_type == "refreshSheets":
            self.event_handler.queue_refresh()
            self.send_status("busy", "Refreshing sheets from Revit...")
            self.raise_revit_event("handleRefreshResult", "Could not raise the Revit refresh event")
            return

        if message_type == "saveOrder":
            self.event_handler.queue_save(message.get("payload") or {})
            self.send_status("busy", "Saving sheet order in Revit...")
            self.raise_revit_event("handleSaveResult", "Could not raise the Revit save event")
            return


# ____________________________________________________________________ MAIN
def main():
    if not os.path.exists(PATH_INDEX):
        forms.alert(
            "The Sheet Discipline Order web app was not found:\n{0}".format(PATH_INDEX),
            title="Sheet Discipline Order",
            exitscript=True
        )

    if not validate_required_parameters(doc):
        return

    try:
        initial_payload = build_sheet_payload(doc)
    except Exception as data_error:
        forms.alert(
            "Could not read sheets from the current project.\n\n{0}".format(safe_str(data_error)),
            title="Sheet Discipline Order",
            warn_icon=True,
            exitscript=True
        )

    try:
        WebView2, CoreWebView2CreationProperties = load_webview2_types()
    except Exception as webview_error:
        forms.alert(
            "Could not load WebView2 from the Revit installation.\n\n{0}".format(webview_error),
            title="Sheet Discipline Order",
            exitscript=True
        )

    event_handler = SheetOrderEventHandler()
    external_event = ExternalEvent.Create(event_handler)
    window = SheetDisciplineOrderWindow(
        WebView2,
        CoreWebView2CreationProperties,
        initial_payload,
        event_handler,
        external_event
    )
    event_handler.window = window
    WINDOW_REFS.append(window)
    window.Show()


if __name__ == "__main__":
    main()
