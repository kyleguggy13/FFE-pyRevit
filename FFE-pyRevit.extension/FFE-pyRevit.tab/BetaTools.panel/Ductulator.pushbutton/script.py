# -*- coding: utf-8 -*-
__title__     = "Ductulator"
__version__   = "Version = v0.2"
__persistentengine__ = True
__doc__       = """Version = v0.2
Date    = 05.07.2026
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
from System.Windows import ResizeMode, Window, WindowStartupLocation
from System.Windows.Controls import Grid


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


def load_webview2_control():
    """Load Revit's bundled WebView2 assemblies and return the WPF control type."""
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

    from Microsoft.Web.WebView2.Wpf import WebView2
    return WebView2


def json_dumps(value):
    return json.dumps(value, ensure_ascii=True)


def make_file_uri(path):
    return Uri(os.path.abspath(path))


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
    try:
        picked_ref = selection.PickObject(
            ObjectType.Element,
            DuctSelectionFilter(),
            "Select one rigid duct to size."
        )
    except OperationCanceledException:
        script.exit()
    except:
        script.exit()

    element = doc.GetElement(picked_ref.ElementId)
    if not is_duct_curve_element(element):
        forms.alert(
            "Select a rigid duct. Flex ducts, fittings, accessories, and equipment are not supported yet.",
            title="Ductulator",
            exitscript=True
        )
    return element


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
        self.pending_payload = None

    def GetName(self):
        return "FFE Ductulator Resize"

    def queue_resize(self, payload):
        self.pending_payload = payload

    def Execute(self, uiapp):
        payload = self.pending_payload
        self.pending_payload = None

        result = {
            "status": "error",
            "message": "No resize request was queued.",
        }

        try:
            message = apply_duct_size(self.duct, payload)
            result = {
                "status": "ready",
                "message": message,
                "duct": read_duct_data(self.duct),
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

        if self.window is not None:
            self.window.send_resize_result(result)


# ____________________________________________________________________ WEBVIEW WINDOW
class DuctulatorWindow(Window):
    def __init__(self, webview_type, duct_data, resize_handler, resize_event):
        Window.__init__(self)

        self.duct_data = duct_data
        self.resize_handler = resize_handler
        self.resize_event = resize_event
        self.has_sent_initial_data = False

        self.Title = "Ductulator - {0}".format(duct_data.get("displayName") or duct_data.get("elementId"))
        self.Width = 1320
        self.Height = 860
        self.MinWidth = 980
        self.MinHeight = 620
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.CanResize

        self.browser = webview_type()

        content_grid = Grid()
        content_grid.Children.Add(self.browser)
        self.Content = content_grid

        self.Closed += self.on_closed
        self.browser.CoreWebView2InitializationCompleted += self.on_core_webview2_initialized
        self.browser.NavigationCompleted += self.on_navigation_completed

        try:
            self.browser.Source = make_file_uri(PATH_INDEX)
        except Exception as exc:
            forms.alert(
                "Could not load the Ductulator web app:\n{0}".format(exc),
                title="Ductulator",
                exitscript=True
            )

    def on_core_webview2_initialized(self, sender, args):
        try:
            if args.IsSuccess:
                self.browser.CoreWebView2.WebMessageReceived += self.on_web_message_received
        except:
            pass

    def on_navigation_completed(self, sender, args):
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

    def send_initial_data(self):
        if self.has_sent_initial_data:
            return
        self.has_sent_initial_data = True
        self.call_revit_api("loadDuct", self.duct_data)

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

        if message_type != "applyDuctSize":
            return

        payload = message.get("payload")
        self.resize_handler.queue_resize(payload)
        self.send_status("warning", "Applying selected size in Revit...")

        try:
            self.resize_event.Raise()
        except Exception as exc:
            self.send_status("error", "Could not raise the Revit resize event: {0}".format(exc))


# ____________________________________________________________________ MAIN
if not os.path.exists(PATH_INDEX):
    forms.alert(
        "The Ductulator web app was not found:\n{0}".format(PATH_INDEX),
        title="Ductulator",
        exitscript=True
    )

duct = pick_duct()

try:
    initial_duct_data = read_duct_data(duct)
except Exception as data_error:
    forms.alert(safe_str(data_error), title="Ductulator", exitscript=True)

try:
    WebView2 = load_webview2_control()
except Exception as webview_error:
    forms.alert(
        "Could not load WebView2 from the Revit installation.\n\n{0}".format(webview_error),
        title="Ductulator",
        exitscript=True
    )

resize_handler = DuctResizeHandler(duct)
resize_event = ExternalEvent.Create(resize_handler)
window = DuctulatorWindow(WebView2, initial_duct_data, resize_handler, resize_event)
resize_handler.window = window
WINDOW_REFS.append(window)
window.Show()
