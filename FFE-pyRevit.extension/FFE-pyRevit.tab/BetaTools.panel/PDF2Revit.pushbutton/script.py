# -*- coding: utf-8 -*-
__title__ = "PDF2\nRevit"
__version__ = "Version = v0.1"
__persistentengine__ = True
__min_revit_ver__ = 2024
__doc__ = """Version = v0.1
Date    = 05.15.2026
__________________________________________________________________
Description:
Creates approximate Revit floor plan geometry from one vector PDF page.

V1 focuses on Floors, Walls, Doors, and Windows.
__________________________________________________________________
How-To:
- Select a vector PDF.
- Pick one page.
- Calibrate scale in the preview window.
- Select Revit targets and analyze.
- Review detected elements before creating the model geometry.
__________________________________________________________________
Last update:
- [05.15.2026] - v0.1 Beta Release
__________________________________________________________________
Author: Kyle Guggenheim"""


# ____________________________________________________________________ IMPORTS
import json
import os
import subprocess
import time
import traceback

import clr
clr.AddReference("System")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System import Uri
from System.Collections.Generic import List
from System.Windows import ResizeMode, Thickness, Visibility, Window, WindowStartupLocation
from System.Windows.Controls import Grid, TextBlock

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    CurveLoop,
    Element,
    ElementId,
    FilteredElementCollector,
    Floor,
    FloorType,
    FamilySymbol,
    Level,
    Line,
    Transaction,
    Wall,
    WallType,
    XYZ,
)
from Autodesk.Revit.DB.Structure import StructuralType

from pyrevit import forms, revit


# ____________________________________________________________________ VARIABLES
doc = revit.doc
revit_app = __revit__.Application

PATH_SCRIPT = os.path.dirname(__file__)
PATH_SUPPORT = os.path.join(PATH_SCRIPT, "support")
PATH_INDEX = os.path.join(PATH_SUPPORT, "index.html")
PATH_HELPER = os.path.join(PATH_SUPPORT, "pdf2revit_helper.py")

APP_NAME = "FFE PDF2Revit"
APP_VERSION = "v0.1"
LOCAL_APP_NAME = "PDF2Revit"
DEFAULT_WALL_HEIGHT_FT = 10.0
MIN_CURVE_LENGTH_FT = 0.10

try:
    WINDOW_REFS
except NameError:
    WINDOW_REFS = []


# ____________________________________________________________________ BASICS
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


def json_loads(value):
    return json.loads(value)


def element_id_value(element_id):
    if element_id is None:
        return None
    try:
        return int(element_id.Value)
    except:
        try:
            return int(element_id.IntegerValue)
        except:
            return None


def make_element_id(value):
    try:
        return ElementId(int(value))
    except:
        return ElementId.InvalidElementId


def get_element_name(element):
    if element is None:
        return ""
    try:
        name = Element.Name.GetValue(element)
        if name:
            return safe_str(name)
    except:
        pass
    try:
        return safe_str(element.Name)
    except:
        return safe_str(element)


def get_type_name(element):
    if element is None:
        return ""
    try:
        family_name = safe_str(element.FamilyName)
        type_name = get_element_name(element)
        if family_name and type_name:
            return "{0} : {1}".format(family_name, type_name)
    except:
        pass
    return get_element_name(element)


def make_file_uri(path):
    absolute_path = os.path.abspath(path).replace("\\", "/")
    return Uri("file:///" + absolute_path.replace(" ", "%20"))


def ensure_dir(path):
    if not os.path.exists(path):
        os.makedirs(path)
    return path


def get_local_app_dir():
    base_folder = os.environ.get("LOCALAPPDATA")
    if not base_folder:
        base_folder = os.path.join(os.path.expanduser("~"), "AppData", "Local")
    return ensure_dir(os.path.join(base_folder, "FFE-pyRevit", LOCAL_APP_NAME))


def get_run_dir():
    run_stamp = time.strftime("%Y%m%d-%H%M%S")
    return ensure_dir(os.path.join(get_local_app_dir(), "runs", run_stamp))


def write_json(path, payload):
    with open(path, "w") as file_obj:
        file_obj.write(json.dumps(payload, ensure_ascii=True, indent=2))


def read_json(path):
    with open(path, "r") as file_obj:
        return json.loads(file_obj.read())


# ____________________________________________________________________ PROCESS / DEPENDENCY BOOTSTRAP
def run_process(command, cwd=None):
    try:
        proc = subprocess.Popen(
            command,
            cwd=cwd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )
        stdout_data, stderr_data = proc.communicate()
        return proc.returncode, safe_str(stdout_data), safe_str(stderr_data)
    except Exception as exc:
        return -1, "", safe_str(exc)


def get_venv_python():
    venv_dir = os.path.join(get_local_app_dir(), ".venv")
    return os.path.join(venv_dir, "Scripts", "python.exe")


def get_venv_dir():
    return os.path.dirname(os.path.dirname(get_venv_python()))


def python_can_import(python_exe, module_name):
    if not os.path.exists(python_exe):
        return False
    code = "import {0}".format(module_name)
    return_code, stdout_data, stderr_data = run_process([python_exe, "-c", code])
    return return_code == 0


def find_base_python_command():
    candidates = [
        ["py", "-3"],
        ["python"],
        ["python3"],
    ]

    for candidate in candidates:
        command = list(candidate) + ["-c", "import sys; print(sys.executable)"]
        return_code, stdout_data, stderr_data = run_process(command)
        if return_code == 0:
            return candidate

    return None


def ensure_helper_python():
    venv_python = get_venv_python()

    if python_can_import(venv_python, "fitz"):
        return venv_python

    venv_dir = get_venv_dir()
    base_python = find_base_python_command()
    if not base_python:
        raise Exception(
            "Could not find CPython. Install Python 3, or make sure the Windows "
            "'py' launcher or 'python' command is available."
        )

    if not os.path.exists(venv_python):
        ensure_dir(os.path.dirname(venv_dir))
        return_code, stdout_data, stderr_data = run_process(
            list(base_python) + ["-m", "venv", venv_dir]
        )
        if return_code != 0:
            raise Exception(
                "Could not create the PDF2Revit Python environment.\n\n{0}".format(
                    stderr_data or stdout_data
                )
            )

    if not os.path.exists(venv_python):
        raise Exception("The PDF2Revit Python environment was created, but python.exe was not found.")

    if not python_can_import(venv_python, "fitz"):
        return_code, stdout_data, stderr_data = run_process(
            [venv_python, "-m", "pip", "install", "PyMuPDF>=1.24,<2"]
        )
        if return_code != 0:
            raise Exception(
                "Could not install PyMuPDF into the PDF2Revit environment.\n\n{0}".format(
                    stderr_data or stdout_data
                )
            )

    if not python_can_import(venv_python, "fitz"):
        raise Exception("PyMuPDF installed, but the helper still could not import it.")

    return venv_python


def run_helper(payload):
    if not os.path.exists(PATH_HELPER):
        raise Exception("The PDF2Revit helper was not found:\n{0}".format(PATH_HELPER))

    helper_python = ensure_helper_python()
    run_dir = payload.get("output_dir") or get_run_dir()
    ensure_dir(run_dir)

    request_path = os.path.join(run_dir, "request-{0}.json".format(int(time.time() * 1000)))
    response_path = os.path.join(run_dir, "response-{0}.json".format(int(time.time() * 1000)))

    payload["output_dir"] = run_dir
    write_json(request_path, payload)

    return_code, stdout_data, stderr_data = run_process(
        [helper_python, PATH_HELPER, request_path, response_path],
        cwd=PATH_SUPPORT
    )

    if not os.path.exists(response_path):
        raise Exception(
            "The PDF2Revit helper did not return a response.\n\n{0}".format(
                stderr_data or stdout_data
            )
        )

    result = read_json(response_path)
    if return_code != 0 or not result.get("ok"):
        message = result.get("error") or stderr_data or stdout_data or "Unknown helper error."
        raise Exception(message)

    return result


# ____________________________________________________________________ REVIT OPTIONS
def collect_levels():
    levels = list(FilteredElementCollector(doc).OfClass(Level).ToElements())
    levels.sort(key=lambda level: level.Elevation)
    return levels


def collect_wall_types():
    return sorted(
        list(FilteredElementCollector(doc).OfClass(WallType).ToElements()),
        key=lambda item: get_element_name(item).lower()
    )


def collect_floor_types():
    return sorted(
        list(FilteredElementCollector(doc).OfClass(FloorType).ToElements()),
        key=lambda item: get_element_name(item).lower()
    )


def collect_family_symbols(category):
    return sorted(
        list(
            FilteredElementCollector(doc)
            .OfCategory(category)
            .WhereElementIsElementType()
            .ToElements()
        ),
        key=lambda item: get_type_name(item).lower()
    )


def option_payload(element, selected=False, extra=None):
    payload = {
        "id": element_id_value(element.Id),
        "label": get_type_name(element),
        "selected": bool(selected),
    }

    if extra:
        payload.update(extra)

    return payload


def get_active_level():
    try:
        active_level = doc.ActiveView.GenLevel
        if active_level:
            return active_level
    except:
        pass
    levels = collect_levels()
    if levels:
        return levels[0]
    return None


def find_next_level(level):
    if not level:
        return None
    levels = collect_levels()
    for candidate in levels:
        if candidate.Elevation > level.Elevation + 0.01:
            return candidate
    return None


def wall_type_width(wall_type):
    try:
        width = float(wall_type.Width)
        if width > 0:
            return width
    except:
        pass
    return 0.5


def make_option_payload():
    active_level = get_active_level()
    active_level_id = element_id_value(active_level.Id) if active_level else None

    wall_types = collect_wall_types()
    floor_types = collect_floor_types()
    door_types = collect_family_symbols(BuiltInCategory.OST_Doors)
    window_types = collect_family_symbols(BuiltInCategory.OST_Windows)
    levels = collect_levels()

    if not levels:
        forms.alert("This model does not contain any Levels.", title=APP_NAME, exitscript=True)
    if not wall_types:
        forms.alert("This model does not contain any Wall Types.", title=APP_NAME, exitscript=True)
    if not floor_types:
        forms.alert("This model does not contain any Floor Types.", title=APP_NAME, exitscript=True)
    if not door_types:
        forms.alert("This model does not contain any Door family types.", title=APP_NAME, exitscript=True)
    if not window_types:
        forms.alert("This model does not contain any Window family types.", title=APP_NAME, exitscript=True)

    return {
        "levels": [
            option_payload(level, element_id_value(level.Id) == active_level_id, {
                "elevation": level.Elevation,
            })
            for level in levels
        ],
        "wallTypes": [
            option_payload(wall_type, index == 0, {
                "widthFeet": wall_type_width(wall_type),
            })
            for index, wall_type in enumerate(wall_types)
        ],
        "floorTypes": [
            option_payload(floor_type, index == 0)
            for index, floor_type in enumerate(floor_types)
        ],
        "doorTypes": [
            option_payload(door_type, index == 0)
            for index, door_type in enumerate(door_types)
        ],
        "windowTypes": [
            option_payload(window_type, index == 0)
            for index, window_type in enumerate(window_types)
        ],
    }


# ____________________________________________________________________ PDF SELECTION
def select_pdf_path():
    pdf_path = forms.pick_file()
    if not pdf_path:
        forms.alert("No PDF was selected.", title=APP_NAME, exitscript=True)

    if not pdf_path.lower().endswith(".pdf"):
        forms.alert("Please select a PDF file.", title=APP_NAME, exitscript=True)

    return pdf_path


def select_page(info_payload):
    page_count = int(info_payload.get("page_count") or 0)
    if page_count <= 0:
        raise Exception("The selected PDF does not contain any readable pages.")

    if page_count == 1:
        return 0

    options = ["Page {0}".format(index + 1) for index in range(page_count)]
    selected = forms.SelectFromList.show(
        options,
        title="Select PDF Page",
        button_name="Open Preview"
    )

    if not selected:
        forms.alert("No PDF page was selected.", title=APP_NAME, exitscript=True)

    return options.index(selected)


# ____________________________________________________________________ REVIT CREATION
def get_selected_ids(accepted):
    if not accepted:
        return {}
    return {
        "walls": set(accepted.get("walls") or []),
        "floors": set(accepted.get("floors") or []),
        "doors": set(accepted.get("doors") or []),
        "windows": set(accepted.get("windows") or []),
    }


def is_accepted(accepted_ids, category, element_id):
    category_ids = accepted_ids.get(category)
    if not category_ids:
        return False
    return element_id in category_ids


def get_setting_element(settings, key):
    value = settings.get(key)
    if value is None:
        return None
    try:
        return doc.GetElement(make_element_id(value))
    except:
        return None


def point_from_xy(point, elevation):
    return XYZ(float(point[0]), float(point[1]), elevation)


def curve_length(line):
    try:
        return line.Length
    except:
        return 0.0


def create_floor_from_loop(floor_item, floor_type, level):
    loop_points = floor_item.get("loop") or []
    if len(loop_points) < 3:
        return None

    curve_loop = CurveLoop()
    elevation = level.Elevation
    for index, point in enumerate(loop_points):
        next_point = loop_points[(index + 1) % len(loop_points)]
        start = point_from_xy(point, elevation)
        end = point_from_xy(next_point, elevation)
        line = Line.CreateBound(start, end)
        if curve_length(line) >= MIN_CURVE_LENGTH_FT:
            curve_loop.Append(line)

    curve_loops = List[CurveLoop]()
    curve_loops.Add(curve_loop)
    return Floor.Create(doc, curve_loops, floor_type.Id, level.Id)


def activate_symbol(symbol):
    if symbol and hasattr(symbol, "IsActive") and not symbol.IsActive:
        symbol.Activate()
        return True
    return False


def create_revit_elements(analysis_result, accepted, settings):
    accepted_ids = get_selected_ids(accepted)
    level = get_setting_element(settings, "levelId")
    wall_type = get_setting_element(settings, "wallTypeId")
    floor_type = get_setting_element(settings, "floorTypeId")
    door_type = get_setting_element(settings, "doorTypeId")
    window_type = get_setting_element(settings, "windowTypeId")

    if not level:
        raise Exception("Selected Level was not found in the active model.")
    if not isinstance(wall_type, WallType):
        raise Exception("Selected Wall Type was not found in the active model.")
    if not isinstance(floor_type, FloorType):
        raise Exception("Selected Floor Type was not found in the active model.")
    if not isinstance(door_type, FamilySymbol):
        raise Exception("Selected Door Type was not found in the active model.")
    if not isinstance(window_type, FamilySymbol):
        raise Exception("Selected Window Type was not found in the active model.")

    elements = analysis_result.get("elements") or {}
    walls = elements.get("walls") or []
    floors = elements.get("floors") or []
    doors = elements.get("doors") or []
    windows = elements.get("windows") or []
    next_level = find_next_level(level)

    created = {
        "walls": 0,
        "floors": 0,
        "doors": 0,
        "windows": 0,
    }
    warnings = []
    wall_map = {}

    transaction = Transaction(doc, "PDF2Revit - Create Floor Plan")
    transaction.Start()
    try:
        symbols_activated = False
        if any(is_accepted(accepted_ids, "doors", item.get("id")) for item in doors):
            symbols_activated = activate_symbol(door_type) or symbols_activated
        if any(is_accepted(accepted_ids, "windows", item.get("id")) for item in windows):
            symbols_activated = activate_symbol(window_type) or symbols_activated
        if symbols_activated:
            doc.Regenerate()

        for wall_item in walls:
            wall_id = wall_item.get("id")
            if not is_accepted(accepted_ids, "walls", wall_id):
                continue

            points = wall_item.get("points") or []
            if len(points) != 2:
                warnings.append("Skipped wall {0}: invalid point data.".format(wall_id))
                continue

            start = point_from_xy(points[0], level.Elevation)
            end = point_from_xy(points[1], level.Elevation)
            curve = Line.CreateBound(start, end)
            if curve_length(curve) < MIN_CURVE_LENGTH_FT:
                warnings.append("Skipped wall {0}: too short.".format(wall_id))
                continue

            wall = Wall.Create(
                doc,
                curve,
                wall_type.Id,
                level.Id,
                DEFAULT_WALL_HEIGHT_FT,
                0.0,
                False,
                False
            )

            if next_level:
                try:
                    wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(next_level.Id)
                except:
                    pass

            wall_map[wall_id] = wall
            created["walls"] += 1

        if wall_map:
            doc.Regenerate()

        for floor_item in floors:
            floor_id = floor_item.get("id")
            if not is_accepted(accepted_ids, "floors", floor_id):
                continue

            try:
                create_floor_from_loop(floor_item, floor_type, level)
                created["floors"] += 1
            except Exception as floor_error:
                warnings.append("Skipped floor {0}: {1}".format(floor_id, floor_error))

        for door_item in doors:
            door_id = door_item.get("id")
            if not is_accepted(accepted_ids, "doors", door_id):
                continue

            host_wall = wall_map.get(door_item.get("host_wall_id"))
            if not host_wall:
                warnings.append("Skipped door {0}: host wall was not created.".format(door_id))
                continue

            try:
                point = point_from_xy(door_item.get("point"), level.Elevation)
                doc.Create.NewFamilyInstance(point, door_type, host_wall, level, StructuralType.NonStructural)
                created["doors"] += 1
            except Exception as door_error:
                warnings.append("Skipped door {0}: {1}".format(door_id, door_error))

        for window_item in windows:
            window_id = window_item.get("id")
            if not is_accepted(accepted_ids, "windows", window_id):
                continue

            host_wall = wall_map.get(window_item.get("host_wall_id"))
            if not host_wall:
                warnings.append("Skipped window {0}: host wall was not created.".format(window_id))
                continue

            try:
                point = point_from_xy(window_item.get("point"), level.Elevation)
                doc.Create.NewFamilyInstance(point, window_type, host_wall, level, StructuralType.NonStructural)
                created["windows"] += 1
            except Exception as window_error:
                warnings.append("Skipped window {0}: {1}".format(window_id, window_error))

        transaction.Commit()
    except:
        transaction.RollBack()
        raise

    return {
        "ok": True,
        "created": created,
        "warnings": warnings,
    }


# ____________________________________________________________________ WEBVIEW
def get_revit_install_dir():
    try:
        app_path = revit_app.ApplicationPath
        if app_path and os.path.exists(app_path):
            return os.path.dirname(app_path)
    except:
        pass

    try:
        version = safe_str(revit_app.VersionNumber) or "2026"
    except:
        version = "2026"

    return os.path.join(
        os.environ.get("ProgramFiles", r"C:\Program Files"),
        "Autodesk",
        "Revit {0}".format(version)
    )


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
    return ensure_dir(os.path.join(get_local_app_dir(), "WebView2"))


class PDF2RevitWindow(Window):
    def __init__(self, webview_type, creation_properties_type, payload):
        Window.__init__(self)

        self.payload = payload
        self.has_sent_payload = False
        self.analysis_result = None
        self.index_uri = make_file_uri(PATH_INDEX)

        self.Title = "{0} - {1}".format(APP_NAME, os.path.basename(payload.get("pdfPath") or "PDF"))
        self.Width = 1420
        self.Height = 900
        self.MinWidth = 1060
        self.MinHeight = 700
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.CanResize

        self.browser = webview_type()
        self.status_text = TextBlock()
        self.status_text.Margin = Thickness(18)
        self.status_text.Text = "Initializing PDF2Revit preview..."

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
        self.status_text.Text = "Loading PDF2Revit preview..."
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
                self.status_text.Text = "PDF2Revit navigation failed.\nURI: {0}\nWeb error: {1}".format(
                    self.index_uri.AbsoluteUri,
                    safe_str(args.WebErrorStatus)
                )
                return
        except:
            pass

        self.status_text.Visibility = Visibility.Collapsed
        self.send_payload()

    def on_closed(self, sender, args):
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
            self.status_text.Visibility = Visibility.Visible
            self.status_text.Text = "Could not send data to the PDF2Revit web app:\n{0}".format(
                traceback.format_exc()
            )

    def send_payload(self, force=False):
        if self.has_sent_payload and not force:
            return
        self.has_sent_payload = True
        self.execute_script(
            "window.pdf2revit && window.pdf2revit.loadData({0});".format(
                json_dumps(self.payload)
            )
        )

    def send_error(self, message):
        self.execute_script(
            "window.pdf2revit && window.pdf2revit.showError({0});".format(
                json_dumps(safe_str(message))
            )
        )

    def send_analysis(self, result):
        self.execute_script(
            "window.pdf2revit && window.pdf2revit.loadAnalysis({0});".format(
                json_dumps(result)
            )
        )

    def send_create_result(self, result):
        self.execute_script(
            "window.pdf2revit && window.pdf2revit.loadCreateResult({0});".format(
                json_dumps(result)
            )
        )

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
            message = json_loads(raw_message)
        except:
            return

        message_type = message.get("type")

        if message_type == "appReady":
            self.send_payload()
            return

        if message_type == "closeWindow":
            self.Close()
            return

        if message_type == "analyze":
            self.handle_analyze(message)
            return

        if message_type == "create":
            self.handle_create(message)
            return

    def handle_analyze(self, message):
        try:
            request = {
                "operation": "analyze",
                "pdf_path": self.payload.get("pdfPath"),
                "page_index": self.payload.get("pageIndex"),
                "output_dir": self.payload.get("runDir"),
                "calibration": message.get("calibration") or {},
                "settings": message.get("settings") or {},
            }
            result = run_helper(request)
            self.analysis_result = result
            self.send_analysis(result)
        except Exception as exc:
            self.send_error(exc)

    def handle_create(self, message):
        if not self.analysis_result:
            self.send_error("Analyze the PDF before creating Revit elements.")
            return

        try:
            result = create_revit_elements(
                self.analysis_result,
                message.get("accepted") or {},
                message.get("settings") or {}
            )
            self.send_create_result(result)
        except Exception as exc:
            self.send_error(exc)


def focus_existing_window():
    for window in list(WINDOW_REFS):
        try:
            if window.IsVisible:
                window.Activate()
                window.send_payload(force=True)
                return True
        except:
            try:
                WINDOW_REFS.remove(window)
            except:
                pass
    return False


# ____________________________________________________________________ MAIN
if not os.path.exists(PATH_INDEX):
    forms.alert(
        "The PDF2Revit web app was not found:\n{0}".format(PATH_INDEX),
        title=APP_NAME,
        exitscript=True
    )

if not os.path.exists(PATH_HELPER):
    forms.alert(
        "The PDF2Revit helper was not found:\n{0}".format(PATH_HELPER),
        title=APP_NAME,
        exitscript=True
    )

if not focus_existing_window():
    try:
        pdf_path = select_pdf_path()
        run_dir = get_run_dir()
        options_payload = make_option_payload()

        info_result = run_helper({
            "operation": "info",
            "pdf_path": pdf_path,
            "output_dir": run_dir,
        })
        page_index = select_page(info_result)
        preview_result = run_helper({
            "operation": "preview",
            "pdf_path": pdf_path,
            "page_index": page_index,
            "output_dir": run_dir,
        })

        payload = {
            "name": APP_NAME,
            "version": APP_VERSION,
            "pdfPath": pdf_path,
            "pdfName": os.path.basename(pdf_path),
            "pageIndex": page_index,
            "pageNumber": page_index + 1,
            "pageCount": info_result.get("page_count"),
            "runDir": run_dir,
            "previewUri": make_file_uri(preview_result.get("preview_image_path")).AbsoluteUri,
            "page": preview_result.get("page") or {},
            "options": options_payload,
        }

        WebView2, CoreWebView2CreationProperties = load_webview2_types()
    except Exception as startup_error:
        forms.alert(
            "Could not start PDF2Revit.\n\n{0}".format(startup_error),
            title=APP_NAME,
            exitscript=True
        )

    window = PDF2RevitWindow(WebView2, CoreWebView2CreationProperties, payload)
    WINDOW_REFS.append(window)
    window.Show()
