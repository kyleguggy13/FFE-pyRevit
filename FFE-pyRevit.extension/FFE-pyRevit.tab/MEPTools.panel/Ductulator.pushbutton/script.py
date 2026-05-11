# -*- coding: utf-8 -*-
__title__     = "Ductulator"
__version__   = "Version = v1.0"
__persistentengine__ = True
__doc__       = """Version = v1.0
Date    = 05.11.2026
______________________________________________________________
Description:
-> Opens the local ductulator web app inside Revit, prefilled from
   a selected duct, and applies same-shape size changes back to Revit.
______________________________________________________________
How-to:
-> Pick one rigid duct, generate/select a size, then Apply to Revit.
______________________________________________________________
Last update:
- [07.08.2025] - v0.1 BETA RELEASE
- [05.07.2026] - v0.2 WEBVIEW2 REVIT MVP
- [05.11.2026] - v0.3 ENABLED OPENING WITHOUT SELECTION, BETTER ERROR HANDLING, LOGGING, AND CODE REFACTORING
- [05.11.2026] - v1.0 FINALIZED MVP RELEASE
______________________________________________________________
Author: Kyle Guggenheim"""


# ____________________________________________________________________ IMPORTS (SYSTEM)
import json
import os
import traceback

import clr
clr.AddReference("System")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System import Uri
from System.Windows import ResizeMode, Thickness, Visibility, Window, WindowStartupLocation
from System.Windows.Controls import Grid, TextBlock


# ____________________________________________________________________ IMPORTS (AUTODESK)
from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    Transaction,
    UnitTypeId,
    UnitUtils,
)
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType


# ____________________________________________________________________ IMPORTS (PYREVIT)
from pyrevit import forms, script


# ____________________________________________________________________ VARIABLES
revit_app   = __revit__.Application
uidoc       = __revit__.ActiveUIDocument
doc         = uidoc.Document
selection   = uidoc.Selection

PATH_SCRIPT = os.path.dirname(__file__)
PATH_SUPPORT = os.path.join(PATH_SCRIPT, "support")
PATH_INDEX = os.path.join(PATH_SUPPORT, "index.html")

OUTPUT = script.get_output()
LOGGER = script.get_logger()
WINDOW_REFS = []

EPSILON = 1e-9

log_status = ""
action = "Ductulator"

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


def element_id_value(element_id):
    try:
        return int(element_id.IntegerValue)
    except:
        try:
            return int(element_id.Value)
        except:
            return 0


def get_element_id_value(element):
    try:
        return element_id_value(element.Id)
    except:
        return 0


def get_revit_install_dir():
    try:
        app_path = revit_app.ApplicationPath
        if app_path and os.path.exists(app_path):
            return os.path.dirname(app_path)
    except:
        pass

    version = safe_str(getattr(revit_app, "VersionNumber", "2026")) or "2026"
    return os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"), "Autodesk", "Revit {0}".format(version))


def load_webview2_types():
    """Load Revit's bundled WebView2 assemblies and return the WPF control types."""
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


def json_dumps(value):
    return json.dumps(value, ensure_ascii=True)


def make_file_uri(path):
    absolute_path = os.path.abspath(path).replace("\\", "/")
    return Uri("file:///" + absolute_path.replace(" ", "%20"))


def get_webview_user_data_folder():
    base_folder = os.environ.get("LOCALAPPDATA")
    if not base_folder:
        base_folder = os.path.join(os.path.expanduser("~"), "AppData", "Local")

    user_data_folder = os.path.join(base_folder, "FFE-pyRevit", "DuctulatorWebView2")
    if not os.path.exists(user_data_folder):
        os.makedirs(user_data_folder)
    return user_data_folder


# ____________________________________________________________________ REVIT DATA HELPERS
def get_param(element, built_in_parameter):
    try:
        return element.get_Parameter(built_in_parameter)
    except:
        return None


def param_double(element, built_in_parameter):
    param = get_param(element, built_in_parameter)
    if param is None:
        return None

    try:
        if hasattr(param, "HasValue") and not param.HasValue:
            return None
    except:
        pass

    try:
        return param.AsDouble()
    except:
        return None


def param_string(element, built_in_parameter):
    param = get_param(element, built_in_parameter)
    if param is None:
        return ""
    try:
        value = param.AsString()
        return value if value else ""
    except:
        return ""


def lookup_parameter_string(element, parameter_name):
    try:
        param = element.LookupParameter(parameter_name)
        if param:
            value = param.AsString()
            return value if value else ""
    except:
        pass
    return ""


def internal_feet_to_inches(value):
    if value is None:
        return None
    return value * 12.0


def inches_to_internal_feet(value):
    return float(value) / 12.0


def convert_from_internal(value, unit_type_id):
    if value is None:
        return None
    try:
        return UnitUtils.ConvertFromInternalUnits(value, unit_type_id)
    except:
        return None


def is_positive(value):
    try:
        return value is not None and float(value) > EPSILON
    except:
        return False


def is_duct_curve_element(element):
    if element is None:
        return False

    try:
        if element.Category is None:
            return False
    except:
        return False

    try:
        duct_category_id = element_id_value(ElementId(BuiltInCategory.OST_DuctCurves))
        return element_id_value(element.Category.Id) == duct_category_id
    except:
        try:
            return element.Category.Name == "Ducts"
        except:
            return False


def get_display_name(element):
    mark = param_string(element, BuiltInParameter.ALL_MODEL_MARK)
    if mark:
        return mark

    try:
        name = element.Name
        if name:
            return name
    except:
        pass

    return "Element {0}".format(get_element_id_value(element))


def get_system_context(element):
    system_name = ""
    system_type = ""

    try:
        mep_system = element.MEPSystem
    except:
        mep_system = None

    if mep_system is not None:
        try:
            system_name = safe_str(mep_system.Name)
        except:
            system_name = ""

        try:
            system_type = safe_str(mep_system.SystemType)
        except:
            system_type = ""

    return system_name, system_type


def get_parameter_write_state(element, built_in_parameter):
    param = get_param(element, built_in_parameter)
    if param is None:
        return "missing"
    try:
        if param.IsReadOnly:
            return "readonly"
    except:
        pass
    return "writable"


def read_duct_data(element):
    if not is_duct_curve_element(element):
        raise Exception("Select a rigid duct. Flex ducts and fittings are not supported in this MVP.")

    diameter_internal = param_double(element, BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
    width_internal = param_double(element, BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
    height_internal = param_double(element, BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)

    shape = None
    diameter_in = internal_feet_to_inches(diameter_internal)
    width_in = internal_feet_to_inches(width_internal)
    height_in = internal_feet_to_inches(height_internal)

    if is_positive(diameter_in):
        shape = "round"
    elif is_positive(width_in) and is_positive(height_in):
        shape = "rectangular"

    if shape is None:
        raise Exception("Could not resolve the selected duct shape or size parameters.")

    flow_internal = param_double(element, BuiltInParameter.RBS_DUCT_FLOW_PARAM)
    velocity_internal = param_double(element, BuiltInParameter.RBS_VELOCITY)
    friction_internal = param_double(element, BuiltInParameter.RBS_FRICTION)

    flow_cfm = convert_from_internal(flow_internal, UnitTypeId.CubicFeetPerMinute)
    velocity_fpm = convert_from_internal(velocity_internal, UnitTypeId.FeetPerMinute)
    friction_inwg_per_100ft = convert_from_internal(
        friction_internal,
        UnitTypeId.InchesOfWater60DegreesFahrenheitPer100Feet
    )

    system_name, system_type = get_system_context(element)

    data = {
        "elementId": get_element_id_value(element),
        "displayName": get_display_name(element),
        "category": "Ducts",
        "systemName": system_name,
        "systemType": system_type,
        "shape": shape,
        "sizeText": lookup_parameter_string(element, "Size"),
        "flowCfm": flow_cfm,
        "velocityFpm": velocity_fpm,
        "frictionInWgPer100Ft": friction_inwg_per_100ft,
        "writeState": {},
    }

    if shape == "round":
        data["diameterIn"] = diameter_in
        data["writeState"]["diameter"] = get_parameter_write_state(element, BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
    else:
        data["widthIn"] = width_in
        data["heightIn"] = height_in
        data["writeState"]["width"] = get_parameter_write_state(element, BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
        data["writeState"]["height"] = get_parameter_write_state(element, BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)

    return data


class DuctSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return is_duct_curve_element(element)

    def AllowReference(self, reference, position):
        return False


def pick_duct():
    picked_ref = selection.PickObject(
        ObjectType.Element,
        DuctSelectionFilter(),
        "Select one rigid duct to size."
    )
    element = doc.GetElement(picked_ref.ElementId)
    if not is_duct_curve_element(element):
        raise Exception("Select a rigid duct. Flex ducts, fittings, accessories, and equipment are not supported yet.")
    return element


def get_preselected_duct():
    try:
        selected_ids = list(selection.GetElementIds())
    except:
        return None

    if len(selected_ids) != 1:
        return None

    try:
        element = doc.GetElement(selected_ids[0])
    except:
        return None

    if is_duct_curve_element(element):
        return element

    return None


def ensure_writable_parameter(element, built_in_parameter, label):
    param = get_param(element, built_in_parameter)
    if param is None:
        raise Exception("The selected duct does not expose a {0} parameter.".format(label))
    try:
        if param.IsReadOnly:
            raise Exception("The selected duct's {0} parameter is read-only.".format(label))
    except Exception as exc:
        if "read-only" in safe_str(exc):
            raise
    return param


def request_number(payload, key, label):
    try:
        value = float(payload.get(key))
    except:
        raise Exception("{0} was not provided by the web app.".format(label))

    if value <= 0:
        raise Exception("{0} must be greater than zero.".format(label))

    return value


def apply_duct_size(element, payload):
    if payload is None:
        raise Exception("No duct size was provided by the web app.")

    current_data = read_duct_data(element)
    requested_element_id = payload.get("elementId")
    if requested_element_id is not None and int(requested_element_id) != current_data["elementId"]:
        raise Exception("The selected Revit duct no longer matches the web app selection.")

    requested_shape = safe_str(payload.get("shape"))
    if requested_shape != current_data["shape"]:
        raise Exception("Shape changes are not supported in this MVP. Select a {0} size.".format(current_data["shape"]))

    if requested_shape == "round":
        diameter_in = request_number(payload, "diameterIn", "Diameter")
        diameter_param = ensure_writable_parameter(
            element,
            BuiltInParameter.RBS_CURVE_DIAMETER_PARAM,
            "diameter"
        )

        transaction = Transaction(doc, "Resize Ductulator Duct")
        try:
            transaction.Start()
            diameter_param.Set(inches_to_internal_feet(diameter_in))
            doc.Regenerate()
            transaction.Commit()
        except:
            try:
                transaction.RollBack()
            except:
                pass
            raise

        return "Updated duct {0} to {1:g} in diameter.".format(current_data["elementId"], diameter_in)

    width_in = request_number(payload, "widthIn", "Width")
    height_in = request_number(payload, "heightIn", "Height")
    width_param = ensure_writable_parameter(element, BuiltInParameter.RBS_CURVE_WIDTH_PARAM, "width")
    height_param = ensure_writable_parameter(element, BuiltInParameter.RBS_CURVE_HEIGHT_PARAM, "height")

    transaction = Transaction(doc, "Resize Ductulator Duct")
    try:
        transaction.Start()
        width_param.Set(inches_to_internal_feet(width_in))
        height_param.Set(inches_to_internal_feet(height_in))
        doc.Regenerate()
        transaction.Commit()
    except:
        try:
            transaction.RollBack()
        except:
            pass
        raise

    return "Updated duct {0} to {1:g} x {2:g} in.".format(current_data["elementId"], width_in, height_in)


# ____________________________________________________________________ EXTERNAL EVENT
class DuctResizeHandler(IExternalEventHandler):
    def __init__(self, duct):
        self.duct = duct
        self.window = None
        self.pending_action = None
        self.pending_payload = None

    def GetName(self):
        return "FFE Ductulator Revit Bridge"

    def queue_select(self):
        self.pending_action = "select"
        self.pending_payload = None

    def queue_resize(self, payload):
        self.pending_action = "resize"
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
        }

        if action == "select":
            try:
                duct = pick_duct()
                duct_data = read_duct_data(duct)
                self.duct = duct
                result = {
                    "status": "ready",
                    "message": "Loaded Revit duct: {0}".format(duct_data.get("displayName") or duct_data.get("elementId")),
                    "duct": duct_data,
                }
            except OperationCanceledException:
                result = {
                    "status": "warning",
                    "message": "Duct selection canceled.",
                }
            except Exception as exc:
                result = {
                    "status": "error",
                    "message": safe_str(exc) or "Revit could not load the selected duct.",
                }
                try:
                    LOGGER.debug(traceback.format_exc())
                except:
                    pass

        elif action == "resize":
            try:
                if self.duct is None:
                    raise Exception("No Revit duct is loaded. Click Select Duct first.")

                message = apply_duct_size(self.duct, payload)
                duct_data = read_duct_data(self.duct)
                result = {
                    "status": "ready",
                    "message": message,
                    "duct": duct_data,
                }
            except Exception as exc:
                result = {
                    "status": "error",
                    "message": safe_str(exc) or "Revit could not apply the selected duct size.",
                }
                try:
                    LOGGER.debug(traceback.format_exc())
                except:
                    pass

        try:
            if self.window is not None and result.get("duct"):
                self.window.set_duct_data(result.get("duct"))
        except:
            pass

        if self.window is not None:
            self.window.send_resize_result(result)


# ____________________________________________________________________ WEBVIEW WINDOW
class DuctulatorWindow(Window):
    def __init__(self, webview_type, creation_properties_type, duct_data, initial_status, resize_handler, resize_event):
        Window.__init__(self)

        self.duct_data = duct_data
        self.initial_status = initial_status
        self.resize_handler = resize_handler
        self.resize_event = resize_event
        self.has_sent_initial_data = False
        self.index_uri = make_file_uri(PATH_INDEX)

        self.Title = self.get_window_title(duct_data)
        self.Width = 1320
        self.Height = 860
        self.MinWidth = 980
        self.MinHeight = 620
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.CanResize

        self.browser = webview_type()
        self.status_text = TextBlock()
        self.status_text.Margin = Thickness(18)
        self.status_text.Text = "Initializing Ductulator WebView2..."

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
        self.status_text.Visibility = Visibility.Visible
        self.status_text.Text = "Loading Ductulator from:\n{0}".format(self.index_uri.AbsoluteUri)
        try:
            self.browser.EnsureCoreWebView2Async()
        except Exception as exc:
            self.status_text.Text = "Could not initialize WebView2:\n{0}".format(exc)

    def on_core_webview2_initialized(self, sender, args):
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
        try:
            if not args.IsSuccess:
                self.status_text.Visibility = Visibility.Visible
                self.status_text.Text = "Ductulator navigation failed.\nURI: {0}\nWeb error: {1}".format(
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
            self.resize_event.Dispose()
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

    def call_revit_api(self, method_name, payload):
        script_text = "window.ffeRevit && window.ffeRevit.{0}({1});".format(
            method_name,
            json_dumps(payload)
        )
        self.execute_script(script_text)

    def get_window_title(self, duct_data):
        if duct_data:
            label = duct_data.get("displayName") or duct_data.get("elementId")
            if label:
                return "Ductulator - {0}".format(label)
        return "Ductulator"

    def set_duct_data(self, duct_data):
        self.duct_data = duct_data
        self.Title = self.get_window_title(duct_data)

    def send_initial_data(self):
        if self.has_sent_initial_data:
            return
        self.has_sent_initial_data = True
        self.call_revit_api("loadDuct", self.duct_data)
        if self.initial_status:
            self.call_revit_api("setStatus", self.initial_status)

    def send_resize_result(self, result):
        self.call_revit_api("handleResizeResult", result)

    def send_status(self, status, message):
        self.call_revit_api("setStatus", {
            "status": status,
            "message": message,
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

        if message_type == "selectDuct":
            self.resize_handler.queue_select()
            self.send_status("warning", "Select one rigid duct in Revit...")

            try:
                self.resize_event.Raise()
            except Exception as exc:
                self.resize_handler.clear_pending()
                self.send_resize_result({
                    "status": "error",
                    "message": "Could not raise the Revit selection event: {0}".format(exc),
                })
            return

        if message_type != "applyDuctSize":
            return

        payload = message.get("payload")
        self.resize_handler.queue_resize(payload)
        self.send_status("warning", "Applying selected size in Revit...")

        try:
            self.resize_event.Raise()
        except Exception as exc:
            self.resize_handler.clear_pending()
            self.send_resize_result({
                "status": "error",
                "message": "Could not raise the Revit resize event: {0}".format(exc),
            })



#______________________________________________________ LOG ACTION
action = "Ductulator"
log_status = "Success"
def log_action(action, log_status):
    """Log action to user JSON log file."""
    import os, json, time
    from pyrevit import revit

    doc = revit.doc
    doc_path = doc.PathName or "<Untitled>"

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
# output_window.print_md("Logging action: {}".format(log_action(action, log_status)))



# ____________________________________________________________________ MAIN
if not os.path.exists(PATH_INDEX):
    forms.alert(
        "The Ductulator web app was not found:\n{0}".format(PATH_INDEX),
        title="Ductulator",
        exitscript=True
    )

duct = get_preselected_duct()
initial_duct_data = None
initial_status = None

if duct is not None:
    try:
        initial_duct_data = read_duct_data(duct)
    except Exception as data_error:
        duct = None
        initial_status = {
            "status": "warning",
            "message": safe_str(data_error) or "The selected duct could not be loaded.",
        }

try:
    WebView2, CoreWebView2CreationProperties = load_webview2_types()
except Exception as webview_error:
    forms.alert(
        "Could not load WebView2 from the Revit installation.\n\n{0}".format(webview_error),
        title="Ductulator",
        exitscript=True
    )

resize_handler = DuctResizeHandler(duct)
resize_event = ExternalEvent.Create(resize_handler)
window = DuctulatorWindow(
    WebView2,
    CoreWebView2CreationProperties,
    initial_duct_data,
    initial_status,
    resize_handler,
    resize_event
)
resize_handler.window = window
WINDOW_REFS.append(window)
window.Show()
