# -*- coding: utf-8 -*-
__title__ = "Pipe One-Line\nEditor"
__version__ = "Version = v0.1"
__persistentengine__ = True
__min_revit_ver__ = 2026
__doc__ = """Version = v0.1
Date    = 06.05.2026
__________________________________________________________________
Description:
Persistent WebView2 window for generating and editing piping one-line diagrams.
The editor reads the selected piping system, renders a schematic web canvas,
and saves the edited one-line back to a dedicated Revit drafting view.
__________________________________________________________________
How-To:
Select a pipe, fitting, accessory, or equipment element in a piping system,
then run the command. Use Select System in the editor to pick a different
system.
__________________________________________________________________
Last update:
- [06.05.2026] - v0.1 BETA
__________________________________________________________________
Author: Kyle Guggenheim"""

import json
import math
import os
import re
import traceback
from collections import deque

import clr
clr.AddReference("System")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System import Guid, String, Uri
from System.Collections.Generic import List
from System.Windows import Thickness, Visibility, Window
from System.Windows.Controls import Grid, TextBlock
from System.Windows.Media import Brushes

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    ConnectorType,
    CurveElement,
    ElementId,
    ElementTypeGroup,
    FamilyInstance,
    FilteredElementCollector,
    Line,
    TextNote,
    TextNoteType,
    Transaction,
    ViewDrafting,
    ViewFamily,
    ViewFamilyType,
    ViewType,
    XYZ,
)
from Autodesk.Revit.DB.ExtensibleStorage import AccessLevel, DataStorage, Entity, Schema, SchemaBuilder
from Autodesk.Revit.DB.Plumbing import PipingSystem
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

from pyrevit import forms, revit, script


# ____________________________________________________________________ CONSTANTS
APP_NAME = "FFE Pipe One-Line Editor"
APP_VERSION = "v0.1"

PATH_SCRIPT = os.path.dirname(__file__)
PATH_SUPPORT = os.path.join(PATH_SCRIPT, "support")
PATH_INDEX = os.path.join(PATH_SUPPORT, "index.html")

SCHEMA_GUID = Guid("79f6634a-37bb-4c54-a843-2e7f92379e55")
SCHEMA_NAME = "FFEPipeOneLineDiagram"
SCHEMA_FIELD_TOOL = "Tool"
SCHEMA_FIELD_SYSTEM_UNIQUE_ID = "SystemUniqueId"
SCHEMA_FIELD_SYSTEM_ID = "SystemId"
SCHEMA_FIELD_DOCUMENT_TITLE = "DocumentTitle"
SCHEMA_FIELD_PAYLOAD = "Payload"

SCHEMA_VERSION = 1
GENERATION_MODE = "compact-significant-nodes"
DRAWING_VIEW_PREFIX = "FFE Pipe One-Line"
SVG_TO_FEET = 1.0 / 24.0
MIN_DETAIL_LINE_LENGTH = 0.0005
VISIBLE_NODE_KINDS = ["branch", "equipment", "accessory", "valve", "pump", "strainer", "meter"]
SIZE_CHANGE_KIND = "sizeChange"

try:
    WINDOW_REFS
except NameError:
    WINDOW_REFS = []

uidoc = revit.uidoc
doc = revit.doc
LOGGER = script.get_logger()


# ____________________________________________________________________ BASIC HELPERS
def safe_str(value):
    if value is None:
        return ""
    try:
        return unicode(value)  # noqa: F821 - IronPython
    except:
        try:
            return str(value)
        except:
            return ""


def json_dumps(value):
    return json.dumps(value, ensure_ascii=True, separators=(",", ":"))


def element_id_value(element_id):
    if element_id is None:
        return None
    try:
        return int(element_id.IntegerValue)
    except:
        pass
    try:
        return int(element_id.Value)
    except:
        pass
    try:
        return int(str(element_id))
    except:
        return None


def element_key(element):
    if element is None:
        return ""
    return safe_str(element_id_value(element.Id))


def get_element_name(element):
    if element is None:
        return ""
    try:
        name = element.Name
        if name:
            return safe_str(name)
    except:
        pass
    try:
        return safe_str(element.GetType().Name)
    except:
        return safe_str(element)


def get_type_name(element):
    if element is None:
        return ""
    try:
        type_id = element.GetTypeId()
        if type_id and type_id != ElementId.InvalidElementId:
            element_type = element.Document.GetElement(type_id)
            if element_type is not None:
                return get_element_name(element_type)
    except:
        pass
    return ""


def get_category_name(element):
    try:
        return safe_str(element.Category.Name)
    except:
        return ""


def normalize_for_label(value):
    text = safe_str(value).strip()
    text = re.sub(r"\s+", " ", text)
    return text


def number_to_label(value, precision):
    try:
        number = float(value)
    except:
        return ""
    if math.isnan(number) or math.isinf(number):
        return ""
    formatted = ("{0:." + str(precision) + "f}").format(number)
    formatted = formatted.rstrip("0").rstrip(".")
    return formatted


def safe_get_parameter_text(element, parameter_name):
    if element is None or not parameter_name:
        return ""
    try:
        parameter = element.LookupParameter(parameter_name)
        if parameter and parameter.HasValue:
            text = parameter.AsString()
            if text:
                return safe_str(text)
            text = parameter.AsValueString()
            if text:
                return safe_str(text)
    except:
        pass
    return ""


def get_builtin_parameter_double(element, builtin_name):
    if element is None:
        return None
    try:
        bip = getattr(BuiltInParameter, builtin_name)
    except:
        return None
    try:
        parameter = element.get_Parameter(bip)
        if parameter and parameter.HasValue:
            return float(parameter.AsDouble())
    except:
        return None
    return None


def get_pipe_diameter_label(element):
    diameter_ft = get_builtin_parameter_double(element, "RBS_PIPE_DIAMETER_PARAM")
    if diameter_ft is None:
        diameter_ft = get_builtin_parameter_double(element, "RBS_CURVE_DIAMETER_PARAM")
    if diameter_ft is None:
        return ""
    diameter_in = diameter_ft * 12.0
    diameter_text = number_to_label(diameter_in, 2)
    if not diameter_text:
        return ""
    return "Dia {0} in".format(diameter_text)


def get_flow_label(element, mep_system):
    flow_value = get_builtin_parameter_double(element, "RBS_PIPE_FLOW_PARAM")
    if flow_value is None and mep_system is not None:
        try:
            flow_value = float(mep_system.GetFlow())
        except:
            flow_value = None
    if flow_value is None:
        return ""
    flow_text = number_to_label(flow_value, 0)
    if not flow_text:
        return ""
    return flow_text


def get_revit_install_dir():
    try:
        revit_app = __revit__.Application  # noqa: F821 - pyRevit host
    except:
        revit_app = None
    try:
        version = safe_str(revit_app.VersionNumber) or "2026"
    except:
        version = "2026"
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
        clr.AddReferenceToFileAndPath(dll_path)
    from Microsoft.Web.WebView2.Wpf import CoreWebView2CreationProperties, WebView2
    return WebView2, CoreWebView2CreationProperties


def ensure_dir(path):
    if not os.path.exists(path):
        os.makedirs(path)
    return path


def get_local_app_dir():
    base_folder = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    return ensure_dir(os.path.join(base_folder, "FFE-pyRevit"))


def get_webview2_user_data_folder():
    return ensure_dir(os.path.join(get_local_app_dir(), "PipeOneLineWebView2"))


def make_file_uri(path):
    return Uri(os.path.abspath(path))


# ____________________________________________________________________ SYSTEM + CONNECTOR GRAPH
def is_piping_system(value):
    try:
        return isinstance(value, PipingSystem)
    except:
        return False


def get_connectors_from_element(element):
    connectors = []
    seen = set()
    if element is None:
        return connectors

    managers = []
    try:
        mep_model = element.MEPModel
        if mep_model is not None and mep_model.ConnectorManager is not None:
            managers.append(mep_model.ConnectorManager)
    except:
        pass
    try:
        if element.ConnectorManager is not None:
            managers.append(element.ConnectorManager)
    except:
        pass

    for manager in managers:
        try:
            for connector in manager.Connectors:
                key = connector_key(connector)
                if key in seen:
                    continue
                seen.add(key)
                connectors.append(connector)
        except:
            pass
    return connectors


def connector_key(connector):
    if connector is None:
        return ""
    owner_key = ""
    try:
        owner_key = element_key(connector.Owner)
    except:
        pass
    try:
        origin = connector.Origin
        return "{0}:{1:.5f}:{2:.5f}:{3:.5f}".format(owner_key, origin.X, origin.Y, origin.Z)
    except:
        return "{0}:{1}".format(owner_key, id(connector))


def connector_is_supported_physical(connector):
    if connector is None:
        return False
    try:
        connector_type = connector.ConnectorType
        unsupported = [
            ConnectorType.Logical,
            ConnectorType.Reference,
            ConnectorType.NodeReference,
            ConnectorType.Invalid,
        ]
        if connector_type in unsupported:
            return False
    except:
        pass
    try:
        if connector.Owner is None:
            return False
    except:
        return False
    return True


def connector_belongs_to_system(connector, mep_system):
    if connector is None or mep_system is None:
        return False
    try:
        connector_system = connector.MEPSystem
        if connector_system is not None and connector_system.Id == mep_system.Id:
            return True
    except:
        pass
    try:
        owner_system = resolve_piping_system_from_element(connector.Owner)
        if owner_system is not None and owner_system.Id == mep_system.Id:
            return True
    except:
        pass
    return False


def iter_physical_refs(connector, mep_system):
    if not connector_is_supported_physical(connector):
        return []
    if mep_system is not None and not connector_belongs_to_system(connector, mep_system):
        return []

    references = []
    try:
        all_refs = connector.AllRefs
    except:
        all_refs = []

    for ref_connector in all_refs:
        if not connector_is_supported_physical(ref_connector):
            continue
        try:
            if ref_connector.Owner.Id == connector.Owner.Id:
                continue
        except:
            continue
        if mep_system is not None and not connector_belongs_to_system(ref_connector, mep_system):
            continue
        references.append(ref_connector)
    return references


def resolve_piping_system_from_element(element):
    if element is None:
        return None
    if is_piping_system(element):
        return element

    try:
        direct_system = element.MEPSystem
        if is_piping_system(direct_system):
            return direct_system
    except:
        pass

    systems = {}
    for connector in get_connectors_from_element(element):
        try:
            connector_system = connector.MEPSystem
        except:
            connector_system = None
        if is_piping_system(connector_system):
            systems[element_key(connector_system)] = connector_system

    if len(systems) == 1:
        return list(systems.values())[0]
    if systems:
        best_system = None
        best_size = -1
        for system in systems.values():
            try:
                size = int(system.Elements.Size)
            except:
                size = 0
            if size > best_size:
                best_system = system
                best_size = size
        return best_system
    return None


def get_selected_element(active_uidoc):
    try:
        selected_ids = list(active_uidoc.Selection.GetElementIds())
    except:
        selected_ids = []
    if not selected_ids:
        return None
    try:
        return active_uidoc.Document.GetElement(selected_ids[0])
    except:
        return None


class PipingSystemSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return resolve_piping_system_from_element(element) is not None

    def AllowReference(self, reference, point):
        return False


def pick_piping_system_element(active_uidoc):
    picked_ref = active_uidoc.Selection.PickObject(
        ObjectType.Element,
        PipingSystemSelectionFilter(),
        "Select a pipe, fitting, accessory, or equipment element in a piping system."
    )
    return active_uidoc.Document.GetElement(picked_ref.ElementId)


def should_show_node(node_data):
    return safe_str(node_data.get("kind")) in VISIBLE_NODE_KINDS


def collect_system_graph(seed_element, mep_system):
    raw_nodes = {}
    raw_adjacency = {}
    queue = deque()
    visited_elements = set()
    warnings = []

    seeds = []
    try:
        if mep_system.BaseEquipment is not None:
            seeds.append(mep_system.BaseEquipment)
    except:
        pass
    if seed_element is not None:
        seeds.append(seed_element)
    try:
        for element in mep_system.Elements:
            seeds.append(element)
    except:
        pass

    for seed in seeds:
        if seed is None:
            continue
        key = element_key(seed)
        if key:
            queue.append(seed)

    while queue:
        element = queue.popleft()
        key = element_key(element)
        if not key or key in visited_elements:
            continue
        visited_elements.add(key)
        raw_nodes[key] = build_node_data(element, mep_system)
        raw_adjacency.setdefault(key, {})

        for connector in get_connectors_from_element(element):
            for ref_connector in iter_physical_refs(connector, mep_system):
                try:
                    neighbor = ref_connector.Owner
                except:
                    neighbor = None
                if neighbor is None:
                    continue
                neighbor_key = element_key(neighbor)
                if not neighbor_key:
                    continue
                if neighbor_key not in raw_nodes:
                    raw_nodes[neighbor_key] = build_node_data(neighbor, mep_system)
                raw_adjacency.setdefault(neighbor_key, {})

                edge_key_parts = sorted([key, neighbor_key])
                edge_key = "{0}:{1}".format(edge_key_parts[0], edge_key_parts[1])
                raw_edge = build_edge_data(edge_key, element, neighbor, mep_system)
                raw_adjacency[key][neighbor_key] = raw_edge
                raw_adjacency[neighbor_key][key] = raw_edge

                if neighbor_key not in visited_elements:
                    queue.append(neighbor)

    if not raw_nodes and seed_element is not None:
        raw_nodes[element_key(seed_element)] = build_node_data(seed_element, mep_system)
        warnings.append("Only the selected element could be read from this piping system.")

    nodes, edges, compact_warnings = compact_system_graph(raw_nodes, raw_adjacency, seed_element)
    warnings.extend(compact_warnings)
    return nodes, edges, warnings


def compact_system_graph(raw_nodes, raw_adjacency, seed_element):
    warnings = []
    if not raw_nodes:
        return {}, {}, warnings

    visible_keys = [key for key, node in raw_nodes.items() if should_show_node(node)]
    if not visible_keys:
        fallback_key = element_key(seed_element)
        if fallback_key not in raw_nodes:
            fallback_key = sorted(raw_nodes.keys())[0]
        warnings.append("No tees, equipment, or pipe accessories were found; showing the selected element only.")
        return {fallback_key: raw_nodes[fallback_key]}, {}, warnings

    visible_set = set(visible_keys)
    nodes = dict((key, raw_nodes[key]) for key in visible_keys)
    edges = {}
    processed_visible_pairs = set()

    for start_key in visible_keys:
        for neighbor_key, raw_edge in raw_adjacency.get(start_key, {}).items():
            queue = deque()
            queue.append((neighbor_key, [start_key, neighbor_key], [raw_edge]))
            visited_hidden = set([start_key])

            while queue:
                current_key, path_keys, path_edges = queue.popleft()
                if current_key in visible_set:
                    if current_key != start_key:
                        visible_pair = tuple(sorted([start_key, current_key]))
                        if visible_pair not in processed_visible_pairs:
                            processed_visible_pairs.add(visible_pair)
                            add_compact_edges(edges, nodes, start_key, current_key, path_keys, path_edges, raw_nodes)
                    continue

                if current_key in visited_hidden:
                    continue
                visited_hidden.add(current_key)

                for next_key, next_edge in raw_adjacency.get(current_key, {}).items():
                    if next_key in path_keys:
                        continue
                    queue.append((next_key, path_keys + [next_key], path_edges + [next_edge]))

    if len(nodes) == 1:
        warnings.append("Only one tee, equipment item, or pipe accessory was found in this piping system.")
    return nodes, edges, warnings


def add_compact_edges(edges, nodes, from_key, to_key, path_keys, path_edges, raw_nodes):
    stops = [{"id": from_key, "rawIndex": 0}]
    for stop in find_size_change_stops(path_keys, raw_nodes):
        if stop["id"] not in nodes:
            nodes[stop["id"]] = stop["node"]
        stops.append({"id": stop["id"], "rawIndex": stop["rawIndex"]})
    stops.append({"id": to_key, "rawIndex": len(path_keys) - 1})

    for index in range(len(stops) - 1):
        start = stops[index]
        end = stops[index + 1]
        segment_start = start["rawIndex"]
        segment_end = end["rawIndex"]
        segment_keys = path_keys[segment_start:segment_end + 1]
        segment_edges = path_edges[segment_start:segment_end]
        add_compact_edge(edges, start["id"], end["id"], segment_keys, segment_edges, raw_nodes)


def find_size_change_stops(path_keys, raw_nodes):
    stops = []
    last_diameter = ""
    for index, key in enumerate(path_keys):
        node = raw_nodes.get(key) or {}
        diameter = normalize_for_label(node.get("diameter"))
        if not diameter:
            continue
        if last_diameter and diameter != last_diameter and index not in [0, len(path_keys) - 1]:
            stop_id = "sizechange-{0}".format(key)
            stops.append({
                "id": stop_id,
                "rawIndex": index,
                "node": build_size_change_node(stop_id, key, node, last_diameter, diameter),
            })
        last_diameter = diameter
    return stops


def build_size_change_node(stop_id, raw_key, raw_node, from_diameter, to_diameter):
    label = "{0} to {1}".format(from_diameter, to_diameter)
    return {
        "id": stop_id,
        "kind": SIZE_CHANGE_KIND,
        "elementId": raw_node.get("elementId"),
        "uniqueId": raw_node.get("uniqueId"),
        "sourceElementId": raw_node.get("elementId"),
        "sourceNodeId": raw_key,
        "label": label,
        "diameter": to_diameter,
        "fromDiameter": from_diameter,
        "toDiameter": to_diameter,
        "flow": raw_node.get("flow") or "",
        "x": 0,
        "y": 0,
    }


def add_compact_edge(edges, from_key, to_key, path_keys, path_edges, raw_nodes):
    edge_key_parts = sorted([from_key, to_key])
    edge_key = "{0}:{1}".format(edge_key_parts[0], edge_key_parts[1])
    if edge_key in edges:
        return
    edge = build_compact_edge_data(edge_key, from_key, to_key, path_keys, path_edges, raw_nodes)
    edges[edge_key] = edge


def first_compact_value(path_keys, path_edges, raw_nodes, field_name):
    for edge in path_edges:
        value = normalize_for_label(edge.get(field_name))
        if value:
            return value
    for key in path_keys:
        value = normalize_for_label((raw_nodes.get(key) or {}).get(field_name))
        if value:
            return value
    return ""


def build_compact_edge_data(edge_id, from_key, to_key, path_keys, path_edges, raw_nodes):
    diameter = first_compact_value(path_keys, path_edges, raw_nodes, "diameter")
    flow = first_compact_value(path_keys, path_edges, raw_nodes, "flow")
    label_parts = []
    if flow:
        label_parts.append(flow)
    if diameter:
        label_parts.append(diameter)

    element_ids = []
    for key in path_keys:
        element_id = (raw_nodes.get(key) or {}).get("elementId")
        if element_id is not None and element_id not in element_ids:
            element_ids.append(element_id)

    return {
        "id": edge_id,
        "from": from_key,
        "to": to_key,
        "elementIds": element_ids,
        "label": " | ".join(label_parts),
        "diameter": diameter,
        "flow": flow,
        "points": [],
        "collapsedElementCount": max(0, len(path_keys) - 2),
    }


def element_kind(element):
    category = get_category_name(element).lower()
    name_bundle = " ".join([
        get_element_name(element),
        get_type_name(element),
        safe_get_parameter_text(element, "Type Mark"),
        safe_get_parameter_text(element, "Mark"),
    ]).lower()
    type_name = ""
    try:
        type_name = safe_str(element.GetType().Name).lower()
    except:
        pass

    if "pipe" in type_name or "pipes" in category or "pipe curves" in category:
        return "pipe"
    if "pump" in name_bundle:
        return "pump"
    if "strainer" in name_bundle:
        return "strainer"
    if "meter" in name_bundle:
        return "meter"
    if "valve" in name_bundle:
        return "valve"
    if "accessor" in category:
        return "accessory"
    if "fitting" in category:
        if "tee" in name_bundle or "wye" in name_bundle:
            return "branch"
        return "fitting"
    if "equipment" in category:
        return "equipment"
    return "junction"


def build_node_label(element):
    candidates = [
        safe_get_parameter_text(element, "Mark"),
        safe_get_parameter_text(element, "Type Mark"),
        safe_get_parameter_text(element, "Abbreviation"),
        get_element_name(element),
        get_type_name(element),
    ]
    for candidate in candidates:
        label = normalize_for_label(candidate)
        if label and label.lower() not in ["<unnamed>", "none"]:
            return label
    return "Element {0}".format(element_key(element))


def build_node_data(element, mep_system):
    node_id = element_key(element)
    diameter = get_pipe_diameter_label(element)
    flow = get_flow_label(element, mep_system)
    try:
        unique_id = safe_str(element.UniqueId)
    except:
        unique_id = ""
    return {
        "id": node_id,
        "kind": element_kind(element),
        "elementId": element_id_value(element.Id),
        "uniqueId": unique_id,
        "label": build_node_label(element),
        "diameter": diameter,
        "flow": flow,
        "x": 0,
        "y": 0,
    }


def build_edge_data(edge_id, element_a, element_b, mep_system):
    diameter = get_pipe_diameter_label(element_a) or get_pipe_diameter_label(element_b)
    flow = get_flow_label(element_a, mep_system) or get_flow_label(element_b, mep_system)
    label_parts = []
    if flow:
        label_parts.append(flow)
    if diameter:
        label_parts.append(diameter)
    return {
        "id": edge_id,
        "from": element_key(element_a),
        "to": element_key(element_b),
        "elementIds": [element_id_value(element_a.Id), element_id_value(element_b.Id)],
        "label": " | ".join(label_parts),
        "diameter": diameter,
        "flow": flow,
        "points": [],
    }


def apply_schematic_layout(nodes, edges, root_key):
    if not nodes:
        return 900, 600

    adjacency = {}
    for key in nodes:
        adjacency[key] = []
    for edge in edges.values():
        from_key = safe_str(edge.get("from"))
        to_key = safe_str(edge.get("to"))
        if from_key in adjacency and to_key in adjacency:
            adjacency[from_key].append(to_key)
            adjacency[to_key].append(from_key)

    if root_key not in nodes:
        root_key = sorted(nodes.keys())[0]

    levels = {}
    parent = {}
    order = []
    queue = deque([root_key])
    levels[root_key] = 0
    parent[root_key] = None

    while queue:
        current = queue.popleft()
        order.append(current)
        neighbors = sorted(adjacency.get(current, []), key=lambda item: nodes[item].get("label", item))
        for neighbor in neighbors:
            if neighbor in levels:
                continue
            levels[neighbor] = levels[current] + 1
            parent[neighbor] = current
            queue.append(neighbor)

    for key in sorted(nodes.keys()):
        if key not in levels:
            levels[key] = max(levels.values() or [0]) + 1
            parent[key] = None
            order.append(key)

    level_groups = {}
    for key, level in levels.items():
        level_groups.setdefault(level, []).append(key)
    for level in level_groups:
        level_groups[level].sort(key=lambda item: order.index(item) if item in order else 9999)

    row_index = {}
    next_row = 0
    for level in sorted(level_groups.keys()):
        for key in level_groups[level]:
            row_index[key] = next_row
            next_row += 1

    max_x = 0
    max_y = 0
    for key, node in nodes.items():
        level = levels.get(key, 0)
        row = row_index.get(key, 0)
        node["x"] = 70 + level * 120
        node["y"] = 88 + row * 56
        max_x = max(max_x, node["x"])
        max_y = max(max_y, node["y"])

    for edge in edges.values():
        from_key = safe_str(edge.get("from"))
        to_key = safe_str(edge.get("to"))
        if levels.get(to_key, 0) < levels.get(from_key, 0):
            edge["from"], edge["to"] = edge.get("to"), edge.get("from")
        edge["flowDirection"] = "fromTo"

    for edge in edges.values():
        from_node = nodes.get(safe_str(edge.get("from")))
        to_node = nodes.get(safe_str(edge.get("to")))
        if not from_node or not to_node:
            continue
        x1 = float(from_node.get("x") or 0)
        y1 = float(from_node.get("y") or 0)
        x2 = float(to_node.get("x") or 0)
        y2 = float(to_node.get("y") or 0)
        mid_x = (x1 + x2) / 2.0
        edge["points"] = [
            {"x": x1, "y": y1},
            {"x": mid_x, "y": y1},
            {"x": mid_x, "y": y2},
            {"x": x2, "y": y2},
        ]

    return max(760, int(max_x + 140)), max(460, int(max_y + 110))


def build_empty_payload(active_doc, warnings):
    return {
        "schemaVersion": SCHEMA_VERSION,
        "generationMode": GENERATION_MODE,
        "documentTitle": safe_str(getattr(active_doc, "Title", "")),
        "systemId": None,
        "systemUniqueId": "",
        "systemName": "",
        "viewId": None,
        "warnings": warnings or [],
        "nodes": [],
        "edges": [],
        "symbols": [],
        "labels": [{
            "id": "title",
            "kind": "title",
            "text": "Pipe One-Line",
            "x": 96,
            "y": 54,
        }],
        "canvas": {"width": 760, "height": 460},
    }


def build_diagram_payload(active_doc, mep_system, seed_element, extra_warnings):
    warnings = list(extra_warnings or [])
    nodes, edges, graph_warnings = collect_system_graph(seed_element, mep_system)
    warnings.extend(graph_warnings)

    try:
        if not mep_system.IsWellConnected:
            warnings.append("The selected piping system is not well connected.")
    except:
        pass
    try:
        networks = mep_system.GetPhysicalNetworksNumber()
        if networks and int(networks) > 1:
            warnings.append("The selected piping system has {0} physical networks.".format(networks))
    except:
        pass

    root_key = ""
    try:
        if mep_system.BaseEquipment is not None:
            root_key = element_key(mep_system.BaseEquipment)
    except:
        pass
    if not root_key and seed_element is not None:
        root_key = element_key(seed_element)

    canvas_width, canvas_height = apply_schematic_layout(nodes, edges, root_key)

    system_id = element_id_value(mep_system.Id)
    system_name = normalize_for_label(get_element_name(mep_system)) or "Piping System {0}".format(system_id)
    try:
        system_unique_id = safe_str(mep_system.UniqueId)
    except:
        system_unique_id = safe_str(system_id)

    title_text = system_name
    subtitle_text = safe_str(getattr(active_doc, "Title", ""))

    return {
        "schemaVersion": SCHEMA_VERSION,
        "generationMode": GENERATION_MODE,
        "documentTitle": safe_str(getattr(active_doc, "Title", "")),
        "systemId": system_id,
        "systemUniqueId": system_unique_id,
        "systemName": system_name,
        "viewId": None,
        "warnings": warnings,
        "nodes": [nodes[key] for key in sorted(nodes.keys(), key=lambda item: (nodes[item].get("x", 0), nodes[item].get("y", 0)))],
        "edges": [edges[key] for key in sorted(edges.keys())],
        "symbols": [],
        "labels": [
            {"id": "title", "kind": "title", "text": title_text, "x": 90, "y": 46},
            {"id": "subtitle", "kind": "note", "text": subtitle_text, "x": 90, "y": 68},
        ],
        "canvas": {"width": canvas_width, "height": canvas_height},
    }


def update_saved_payload_identity(saved_payload, generated_payload):
    if not isinstance(saved_payload, dict):
        return generated_payload
    if saved_payload.get("generationMode") != generated_payload.get("generationMode"):
        generated_warnings = list(generated_payload.get("warnings") or [])
        generated_warnings.append("Regenerated with the compact one-line layout.")
        generated_payload["warnings"] = generated_warnings
        return generated_payload
    for key in ["schemaVersion", "documentTitle", "systemId", "systemUniqueId", "systemName"]:
        saved_payload[key] = generated_payload.get(key)
    saved_payload["generationMode"] = generated_payload.get("generationMode")
    saved_warnings = list(generated_payload.get("warnings") or [])
    saved_warnings.append("Loaded the saved editable one-line for this piping system.")
    saved_payload["warnings"] = saved_warnings
    if "canvas" not in saved_payload:
        saved_payload["canvas"] = generated_payload.get("canvas") or {"width": 900, "height": 600}
    return saved_payload


def build_payload_from_selection(active_uidoc, force_regenerate=False):
    active_doc = active_uidoc.Document
    selected_element = get_selected_element(active_uidoc)
    if selected_element is None:
        return build_empty_payload(active_doc, ["Select an element in a piping system to generate a one-line."])

    mep_system = resolve_piping_system_from_element(selected_element)
    if mep_system is None:
        return build_empty_payload(active_doc, ["The selected element is not part of a piping system."])

    generated_payload = build_diagram_payload(active_doc, mep_system, selected_element, [])
    if not force_regenerate:
        saved_payload = read_saved_diagram(active_doc, mep_system)
        if saved_payload:
            return update_saved_payload_identity(saved_payload, generated_payload)
    return generated_payload


def build_payload_from_picked_element(active_uidoc):
    active_doc = active_uidoc.Document
    picked_element = pick_piping_system_element(active_uidoc)
    mep_system = resolve_piping_system_from_element(picked_element)
    if mep_system is None:
        return build_empty_payload(active_doc, ["The selected element is not part of a piping system."])
    generated_payload = build_diagram_payload(active_doc, mep_system, picked_element, [])
    saved_payload = read_saved_diagram(active_doc, mep_system)
    if saved_payload:
        return update_saved_payload_identity(saved_payload, generated_payload)
    return generated_payload


# ____________________________________________________________________ EXTENSIBLE STORAGE
def get_storage_schema():
    schema = Schema.Lookup(SCHEMA_GUID)
    if schema is not None:
        return schema
    builder = SchemaBuilder(SCHEMA_GUID)
    builder.SetSchemaName(SCHEMA_NAME)
    builder.SetReadAccessLevel(AccessLevel.Public)
    builder.SetWriteAccessLevel(AccessLevel.Public)
    builder.AddSimpleField(SCHEMA_FIELD_TOOL, String)
    builder.AddSimpleField(SCHEMA_FIELD_SYSTEM_UNIQUE_ID, String)
    builder.AddSimpleField(SCHEMA_FIELD_SYSTEM_ID, String)
    builder.AddSimpleField(SCHEMA_FIELD_DOCUMENT_TITLE, String)
    builder.AddSimpleField(SCHEMA_FIELD_PAYLOAD, String)
    return builder.Finish()


def entity_get_string(entity, schema, field_name):
    if entity is None or schema is None:
        return ""
    field = schema.GetField(field_name)
    if field is None:
        return ""
    try:
        return safe_str(entity.Get[String](field))
    except:
        try:
            return safe_str(entity.Get(field))
        except:
            return ""


def entity_set_string(entity, schema, field_name, value):
    field = schema.GetField(field_name)
    if field is None:
        return
    try:
        entity.Set[String](field, safe_str(value))
    except:
        entity.Set(field, safe_str(value))


def find_diagram_storage(active_doc, system_unique_id):
    schema = get_storage_schema()
    for data_storage in FilteredElementCollector(active_doc).OfClass(DataStorage).ToElements():
        try:
            entity = data_storage.GetEntity(schema)
        except:
            entity = None
        try:
            if entity is None or not entity.IsValid():
                continue
        except:
            continue
        stored_unique_id = entity_get_string(entity, schema, SCHEMA_FIELD_SYSTEM_UNIQUE_ID)
        if stored_unique_id == safe_str(system_unique_id):
            return data_storage, entity, schema
    return None, None, schema


def read_saved_diagram(active_doc, mep_system):
    try:
        system_unique_id = safe_str(mep_system.UniqueId)
    except:
        system_unique_id = safe_str(element_id_value(mep_system.Id))
    try:
        data_storage, entity, schema = find_diagram_storage(active_doc, system_unique_id)
        if data_storage is None or entity is None:
            return None
        payload_text = entity_get_string(entity, schema, SCHEMA_FIELD_PAYLOAD)
        if not payload_text:
            return None
        payload = json.loads(payload_text)
        if isinstance(payload, dict):
            return payload
    except:
        try:
            LOGGER.debug(traceback.format_exc())
        except:
            pass
    return None


def save_diagram_storage(active_doc, payload):
    schema = get_storage_schema()
    system_unique_id = safe_str(payload.get("systemUniqueId"))
    if not system_unique_id:
        raise Exception("No piping system is loaded. Use Select System before saving.")

    data_storage, entity, schema = find_diagram_storage(active_doc, system_unique_id)
    if data_storage is None:
        data_storage = DataStorage.Create(active_doc)
        entity = Entity(schema)
    else:
        entity = Entity(schema)

    entity_set_string(entity, schema, SCHEMA_FIELD_TOOL, APP_NAME)
    entity_set_string(entity, schema, SCHEMA_FIELD_SYSTEM_UNIQUE_ID, system_unique_id)
    entity_set_string(entity, schema, SCHEMA_FIELD_SYSTEM_ID, safe_str(payload.get("systemId")))
    entity_set_string(entity, schema, SCHEMA_FIELD_DOCUMENT_TITLE, safe_str(payload.get("documentTitle")))
    entity_set_string(entity, schema, SCHEMA_FIELD_PAYLOAD, json_dumps(payload))
    data_storage.SetEntity(entity)


# ____________________________________________________________________ REVIT DRAFTING OUTPUT
def sanitize_view_name(value):
    text = normalize_for_label(value) or "Piping System"
    text = re.sub(r"[\\:{}\[\]|;<>?`~]", "-", text)
    text = text.strip(" .-")
    if len(text) > 90:
        text = text[:90].strip()
    return text or "Piping System"


def get_drafting_view_type_id(active_doc):
    try:
        default_id = active_doc.GetDefaultElementTypeId(ElementTypeGroup.ViewTypeDrafting)
        if default_id and default_id != ElementId.InvalidElementId:
            return default_id
    except:
        pass
    for view_type in FilteredElementCollector(active_doc).OfClass(ViewFamilyType).ToElements():
        try:
            if view_type.ViewFamily == ViewFamily.Drafting:
                return view_type.Id
        except:
            pass
    raise Exception("No drafting view type was found in this Revit document.")


def find_drafting_view_by_name(active_doc, view_name):
    for view in FilteredElementCollector(active_doc).OfClass(ViewDrafting).ToElements():
        try:
            if not view.IsTemplate and view.Name == view_name:
                return view
        except:
            pass
    return None


def ensure_diagram_view(active_doc, payload):
    system_name = sanitize_view_name(payload.get("systemName"))
    system_id = safe_str(payload.get("systemId") or "NoId")
    view_name = "{0} - {1} ({2})".format(DRAWING_VIEW_PREFIX, system_name, system_id)
    existing_view = find_drafting_view_by_name(active_doc, view_name)
    if existing_view is not None:
        return existing_view
    view = ViewDrafting.Create(active_doc, get_drafting_view_type_id(active_doc))
    view.Name = view_name
    try:
        view.Scale = 12
    except:
        pass
    return view


def clear_view_contents(active_doc, view):
    element_ids = List[ElementId]()
    try:
        existing_ids = FilteredElementCollector(active_doc, view.Id).WhereElementIsNotElementType().ToElementIds()
    except:
        existing_ids = []
    for element_id in existing_ids:
        try:
            if element_id == view.Id:
                continue
        except:
            pass
        element_ids.Add(element_id)
    if element_ids.Count > 0:
        active_doc.Delete(element_ids)


def svg_point_to_revit(point):
    try:
        x = float(point.get("x", 0))
        y = float(point.get("y", 0))
    except:
        x = 0.0
        y = 0.0
    return XYZ(x * SVG_TO_FEET, -y * SVG_TO_FEET, 0)


def points_are_far_enough(point_a, point_b):
    try:
        dx = point_a.X - point_b.X
        dy = point_a.Y - point_b.Y
        dz = point_a.Z - point_b.Z
        return math.sqrt(dx * dx + dy * dy + dz * dz) > MIN_DETAIL_LINE_LENGTH
    except:
        return False


def draw_detail_line(active_doc, view, point_a, point_b):
    if not points_are_far_enough(point_a, point_b):
        return None
    line = Line.CreateBound(point_a, point_b)
    return active_doc.Create.NewDetailCurve(view, line)


def draw_svg_polyline(active_doc, view, points):
    if not points or len(points) < 2:
        return
    for index in range(len(points) - 1):
        point_a = svg_point_to_revit(points[index])
        point_b = svg_point_to_revit(points[index + 1])
        draw_detail_line(active_doc, view, point_a, point_b)


def get_flow_arrow_points(points):
    if not points or len(points) < 2:
        return None

    normalized = []
    for point in points:
        try:
            normalized.append({"x": float(point.get("x", 0)), "y": float(point.get("y", 0))})
        except:
            pass
    if len(normalized) < 2:
        return None

    lengths = []
    total_length = 0.0
    for index in range(len(normalized) - 1):
        point_a = normalized[index]
        point_b = normalized[index + 1]
        dx = point_b["x"] - point_a["x"]
        dy = point_b["y"] - point_a["y"]
        length = math.sqrt(dx * dx + dy * dy)
        lengths.append(length)
        total_length += length

    if total_length < 1.0:
        return None

    target = total_length / 2.0
    distance_so_far = 0.0
    for index, length in enumerate(lengths):
        if length <= 0.0:
            continue
        if distance_so_far + length >= target:
            point_a = normalized[index]
            point_b = normalized[index + 1]
            ratio = (target - distance_so_far) / length
            tip_x = point_a["x"] + (point_b["x"] - point_a["x"]) * ratio
            tip_y = point_a["y"] + (point_b["y"] - point_a["y"]) * ratio
            dir_x = (point_b["x"] - point_a["x"]) / length
            dir_y = (point_b["y"] - point_a["y"]) / length
            base_x = tip_x - dir_x * 12.0
            base_y = tip_y - dir_y * 12.0
            perp_x = -dir_y
            perp_y = dir_x
            return [
                {"x": base_x + perp_x * 5.0, "y": base_y + perp_y * 5.0},
                {"x": tip_x, "y": tip_y},
                {"x": base_x - perp_x * 5.0, "y": base_y - perp_y * 5.0},
            ]
        distance_so_far += length
    return None


def draw_flow_arrow(active_doc, view, points):
    arrow_points = get_flow_arrow_points(points)
    if not arrow_points:
        return
    draw_svg_polyline(active_doc, view, arrow_points)


def get_default_text_note_type_id(active_doc):
    try:
        type_id = active_doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType)
        if type_id and type_id != ElementId.InvalidElementId:
            return type_id
    except:
        pass
    text_type = FilteredElementCollector(active_doc).OfClass(TextNoteType).FirstElement()
    if text_type is not None:
        return text_type.Id
    raise Exception("No Text Note type was found in this Revit document.")


def create_text_note(active_doc, view, text, x, y, text_type_id):
    clean_text = normalize_for_label(text)
    if not clean_text:
        return None
    point = svg_point_to_revit({"x": x, "y": y})
    return TextNote.Create(active_doc, view.Id, point, clean_text, text_type_id)


def polygon_points(cx, cy, radius, sides):
    points = []
    for index in range(sides):
        angle = (math.pi * 2.0 * index / float(sides)) - (math.pi / 2.0)
        points.append({"x": cx + math.cos(angle) * radius, "y": cy + math.sin(angle) * radius})
    points.append(points[0])
    return points


def draw_box(active_doc, view, cx, cy, width, height):
    left = cx - width / 2.0
    right = cx + width / 2.0
    top = cy - height / 2.0
    bottom = cy + height / 2.0
    draw_svg_polyline(active_doc, view, [
        {"x": left, "y": top},
        {"x": right, "y": top},
        {"x": right, "y": bottom},
        {"x": left, "y": bottom},
        {"x": left, "y": top},
    ])


def draw_symbol(active_doc, view, kind, cx, cy):
    kind = safe_str(kind) or "junction"
    if kind == "valve":
        draw_svg_polyline(active_doc, view, [{"x": cx - 18, "y": cy - 10}, {"x": cx, "y": cy}, {"x": cx - 18, "y": cy + 10}, {"x": cx - 18, "y": cy - 10}])
        draw_svg_polyline(active_doc, view, [{"x": cx + 18, "y": cy - 10}, {"x": cx, "y": cy}, {"x": cx + 18, "y": cy + 10}, {"x": cx + 18, "y": cy - 10}])
        draw_svg_polyline(active_doc, view, [{"x": cx, "y": cy - 16}, {"x": cx, "y": cy - 26}, {"x": cx + 12, "y": cy - 26}])
    elif kind == SIZE_CHANGE_KIND:
        draw_svg_polyline(active_doc, view, [{"x": cx - 18, "y": cy - 9}, {"x": cx + 18, "y": cy - 9}, {"x": cx + 8, "y": cy + 9}, {"x": cx - 8, "y": cy + 9}, {"x": cx - 18, "y": cy - 9}])
        draw_svg_polyline(active_doc, view, [{"x": cx - 20, "y": cy + 14}, {"x": cx + 20, "y": cy - 14}])
    elif kind == "accessory":
        draw_box(active_doc, view, cx, cy, 32, 20)
        draw_svg_polyline(active_doc, view, [{"x": cx - 12, "y": cy + 8}, {"x": cx + 12, "y": cy - 8}])
    elif kind == "pump":
        draw_svg_polyline(active_doc, view, polygon_points(cx, cy, 18, 12))
        draw_svg_polyline(active_doc, view, [{"x": cx - 7, "y": cy - 10}, {"x": cx + 12, "y": cy}, {"x": cx - 7, "y": cy + 10}, {"x": cx - 7, "y": cy - 10}])
    elif kind == "strainer":
        draw_svg_polyline(active_doc, view, [{"x": cx, "y": cy - 18}, {"x": cx + 18, "y": cy}, {"x": cx, "y": cy + 18}, {"x": cx - 18, "y": cy}, {"x": cx, "y": cy - 18}])
        draw_svg_polyline(active_doc, view, [{"x": cx - 9, "y": cy + 9}, {"x": cx + 9, "y": cy - 9}])
        draw_svg_polyline(active_doc, view, [{"x": cx - 3, "y": cy + 15}, {"x": cx + 15, "y": cy - 3}])
    elif kind == "meter":
        draw_box(active_doc, view, cx, cy, 32, 20)
        draw_svg_polyline(active_doc, view, [{"x": cx - 10, "y": cy}, {"x": cx - 2, "y": cy - 7}, {"x": cx + 2, "y": cy + 7}, {"x": cx + 10, "y": cy}])
    elif kind == "equipment":
        draw_box(active_doc, view, cx, cy, 42, 26)
    elif kind == "pipe":
        draw_svg_polyline(active_doc, view, [{"x": cx - 10, "y": cy}, {"x": cx + 10, "y": cy}])
    elif kind == "branch":
        draw_svg_polyline(active_doc, view, [{"x": cx - 7, "y": cy - 7}, {"x": cx + 7, "y": cy + 7}])
        draw_svg_polyline(active_doc, view, [{"x": cx - 7, "y": cy + 7}, {"x": cx + 7, "y": cy - 7}])
    else:
        draw_svg_polyline(active_doc, view, polygon_points(cx, cy, 8, 8))


def draw_edge_label(active_doc, view, edge, text_type_id):
    label = normalize_for_label(edge.get("label"))
    if not label:
        return
    points = edge.get("points") or []
    if not points:
        return
    mid_index = int(len(points) / 2)
    try:
        point = points[mid_index]
        create_text_note(active_doc, view, label, float(point.get("x", 0)) + 8, float(point.get("y", 0)) - 10, text_type_id)
    except:
        pass


def draw_diagram_to_view(active_doc, view, payload):
    text_type_id = get_default_text_note_type_id(active_doc)
    canvas = payload.get("canvas") or {}
    width = float(canvas.get("width") or 900)
    height = float(canvas.get("height") or 600)

    draw_svg_polyline(active_doc, view, [
        {"x": 20, "y": 20},
        {"x": width - 20, "y": 20},
        {"x": width - 20, "y": height - 20},
        {"x": 20, "y": height - 20},
        {"x": 20, "y": 20},
    ])

    for edge in payload.get("edges") or []:
        points = edge.get("points") or []
        if not points:
            from_node = find_payload_node(payload, edge.get("from"))
            to_node = find_payload_node(payload, edge.get("to"))
            if from_node and to_node:
                x1 = float(from_node.get("x") or 0)
                y1 = float(from_node.get("y") or 0)
                x2 = float(to_node.get("x") or 0)
                y2 = float(to_node.get("y") or 0)
                mid_x = (x1 + x2) / 2.0
                points = [{"x": x1, "y": y1}, {"x": mid_x, "y": y1}, {"x": mid_x, "y": y2}, {"x": x2, "y": y2}]
        draw_svg_polyline(active_doc, view, points)
        draw_flow_arrow(active_doc, view, points)
        draw_edge_label(active_doc, view, edge, text_type_id)

    for node in payload.get("nodes") or []:
        try:
            cx = float(node.get("x") or 0)
            cy = float(node.get("y") or 0)
        except:
            continue
        draw_symbol(active_doc, view, node.get("kind"), cx, cy)
        label = normalize_for_label(node.get("label"))
        if label:
            create_text_note(active_doc, view, label, cx - 24, cy - 28, text_type_id)
        diameter = normalize_for_label(node.get("diameter"))
        if diameter and node.get("kind") == "pipe":
            create_text_note(active_doc, view, diameter, cx - 24, cy + 20, text_type_id)

    for symbol in payload.get("symbols") or []:
        try:
            cx = float(symbol.get("x") or 0)
            cy = float(symbol.get("y") or 0)
        except:
            continue
        draw_symbol(active_doc, view, symbol.get("kind"), cx, cy)
        label = normalize_for_label(symbol.get("label"))
        if label:
            create_text_note(active_doc, view, label, cx - 24, cy - 28, text_type_id)

    for label in payload.get("labels") or []:
        try:
            create_text_note(active_doc, view, label.get("text"), float(label.get("x") or 0), float(label.get("y") or 0), text_type_id)
        except:
            pass


def find_payload_node(payload, node_id):
    node_id = safe_str(node_id)
    for node in payload.get("nodes") or []:
        if safe_str(node.get("id")) == node_id:
            return node
    return None


def save_payload_to_revit(active_doc, payload):
    if not isinstance(payload, dict):
        raise Exception("The web editor did not send a valid diagram.")
    if not payload.get("systemUniqueId"):
        raise Exception("No piping system is loaded. Use Select System before saving.")
    payload["schemaVersion"] = SCHEMA_VERSION
    payload["generationMode"] = GENERATION_MODE

    transaction = Transaction(active_doc, "Save Pipe One-Line Diagram")
    transaction.Start()
    try:
        view = ensure_diagram_view(active_doc, payload)
        clear_view_contents(active_doc, view)
        payload["viewId"] = element_id_value(view.Id)
        payload["documentTitle"] = safe_str(getattr(active_doc, "Title", ""))
        draw_diagram_to_view(active_doc, view, payload)
        save_diagram_storage(active_doc, payload)
        transaction.Commit()
    except Exception:
        try:
            transaction.RollBack()
        except:
            pass
        raise

    return {
        "status": "ready",
        "message": "Saved one-line diagram to drafting view '{0}'.".format(view.Name),
        "payload": payload,
        "viewId": payload.get("viewId"),
        "viewName": view.Name,
    }


# ____________________________________________________________________ EXTERNAL EVENT
class PipeOneLineEventHandler(IExternalEventHandler):
    def __init__(self):
        self.window = None
        self.pending_action = None
        self.pending_payload = None

    def GetName(self):
        return "FFE Pipe One-Line Revit Bridge"

    def queue_select(self):
        self.pending_action = "select"
        self.pending_payload = None

    def queue_refresh(self, payload):
        self.pending_action = "refresh"
        self.pending_payload = payload or {}

    def queue_save(self, payload):
        self.pending_action = "save"
        self.pending_payload = payload or {}

    def Execute(self, uiapp):
        action = self.pending_action
        payload = self.pending_payload
        self.pending_action = None
        self.pending_payload = None

        if self.window is None or action is None:
            return

        try:
            active_uidoc = uiapp.ActiveUIDocument
            active_doc = active_uidoc.Document

            if action == "select":
                result_payload = build_payload_from_picked_element(active_uidoc)
                self.window.send_refresh_result({
                    "status": "ready",
                    "message": "Loaded piping system '{0}'.".format(result_payload.get("systemName") or ""),
                    "payload": result_payload,
                })
                self.window.load_diagram(result_payload)
                return

            if action == "refresh":
                force_regenerate = bool(payload.get("forceRegenerate"))
                result_payload = build_payload_from_selection(active_uidoc, force_regenerate)
                message = "Regenerated one-line from the current selection." if force_regenerate else "Loaded one-line from the current selection."
                self.window.send_refresh_result({
                    "status": "ready" if result_payload.get("systemUniqueId") else "warning",
                    "message": message,
                    "payload": result_payload,
                })
                self.window.load_diagram(result_payload)
                return

            if action == "save":
                save_result = save_payload_to_revit(active_doc, payload)
                self.window.send_save_result(save_result)
                if save_result.get("payload"):
                    self.window.load_diagram(save_result.get("payload"))
                return

        except Exception as exc:
            result = {
                "status": "error",
                "message": safe_str(exc) or "Pipe one-line action failed.",
            }
            try:
                LOGGER.debug(traceback.format_exc())
            except:
                pass
            if action == "save":
                self.window.send_save_result(result)
            else:
                self.window.send_refresh_result(result)


# ____________________________________________________________________ WEBVIEW WINDOW
class PipeOneLineWindow(Window):
    def __init__(self, webview_type, creation_properties_type, initial_payload, event_handler, external_event):
        Window.__init__(self)

        self.initial_payload = initial_payload
        self.event_handler = event_handler
        self.external_event = external_event
        self.has_sent_initial_payload = False
        self.index_uri = make_file_uri(PATH_INDEX)

        self.Title = APP_NAME
        self.Width = 1280
        self.Height = 820
        self.MinWidth = 980
        self.MinHeight = 640

        self.root = Grid()
        self.Content = self.root

        self.browser = webview_type()
        self.status_text = TextBlock()
        self.status_text.Text = "Initializing Pipe One-Line WebView2..."
        self.status_text.Margin = Thickness(16)
        self.status_text.Foreground = Brushes.Black

        self.root.Children.Add(self.browser)
        self.root.Children.Add(self.status_text)

        try:
            creation_properties = creation_properties_type()
            creation_properties.UserDataFolder = get_webview2_user_data_folder()
            self.browser.CreationProperties = creation_properties
        except Exception as exc:
            self.status_text.Text = "Could not configure WebView2 user data folder:\n{0}".format(exc)

        self.Loaded += self.on_loaded
        self.Closed += self.on_closed
        self.browser.CoreWebView2InitializationCompleted += self.on_core_webview2_initialized
        self.browser.NavigationCompleted += self.on_navigation_completed

    def on_loaded(self, sender, args):
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
                self.status_text.Text = "The Pipe One-Line web app could not be loaded."
                return
        except:
            pass
        self.status_text.Visibility = Visibility.Collapsed
        self.send_initial_payload()

    def on_closed(self, sender, args):
        try:
            WINDOW_REFS.remove(self)
        except:
            pass
        try:
            self.event_handler.window = None
        except:
            pass

    def execute_script(self, script_text):
        try:
            if self.browser.CoreWebView2 is not None:
                self.browser.CoreWebView2.ExecuteScriptAsync(script_text)
        except:
            try:
                LOGGER.debug(traceback.format_exc())
            except:
                pass

    def call_web_app(self, method_name, payload):
        script_text = "window.ffePipeOneLine && window.ffePipeOneLine.{0}({1});".format(
            method_name,
            json_dumps(payload)
        )
        self.execute_script(script_text)

    def send_initial_payload(self):
        if self.has_sent_initial_payload:
            return
        self.has_sent_initial_payload = True
        self.load_diagram(self.initial_payload)

    def load_diagram(self, payload):
        self.initial_payload = payload
        self.call_web_app("loadDiagram", payload)

    def send_status(self, status, message):
        self.call_web_app("setStatus", {"status": status, "message": message})

    def send_save_result(self, result):
        self.call_web_app("handleSaveResult", result)

    def send_refresh_result(self, result):
        self.call_web_app("handleRefreshResult", result)

    def request_refresh_from_app(self):
        self.send_status("warning", "Refreshing from the active Revit selection...")
        self.event_handler.queue_refresh({"forceRegenerate": False})
        self.external_event.Raise()

    def on_web_message_received(self, sender, args):
        raw_message = ""
        try:
            raw_message = args.TryGetWebMessageAsString()
        except:
            try:
                raw_message = args.WebMessageAsJson
            except:
                raw_message = ""

        try:
            message = json.loads(raw_message)
        except:
            return

        message_type = message.get("type")
        if message_type == "appReady":
            self.send_initial_payload()
            return

        if message_type == "dirtyStateChanged":
            return

        if message_type == "closeWindow":
            self.Close()
            return

        if message_type == "selectSystem":
            self.event_handler.queue_select()
            self.send_status("warning", "Select a piping system element in Revit...")
            self.external_event.Raise()
            return

        if message_type == "refreshFromSelection":
            self.event_handler.queue_refresh(message.get("payload") or {})
            self.send_status("warning", "Reading the active Revit selection...")
            self.external_event.Raise()
            return

        if message_type == "saveDiagram":
            self.event_handler.queue_save(message.get("payload") or {})
            self.send_status("warning", "Saving one-line diagram to Revit...")
            self.external_event.Raise()
            return


# ____________________________________________________________________ WINDOW LIFETIME
def focus_existing_window():
    for window in list(WINDOW_REFS):
        try:
            if window.IsVisible:
                window.Activate()
                window.request_refresh_from_app()
                return True
        except:
            try:
                WINDOW_REFS.remove(window)
            except:
                pass
    return False


if not os.path.exists(PATH_INDEX):
    forms.alert(
        "The Pipe One-Line web app was not found:\n{0}".format(PATH_INDEX),
        title=APP_NAME,
        exitscript=True
    )

if not focus_existing_window():
    try:
        initial_payload = build_payload_from_selection(uidoc, force_regenerate=False)
        WebView2, CoreWebView2CreationProperties = load_webview2_types()
    except Exception as startup_error:
        forms.alert(
            "Could not start the Pipe One-Line editor.\n\n{0}".format(startup_error),
            title=APP_NAME,
            exitscript=True
        )

    handler = PipeOneLineEventHandler()
    external_event = ExternalEvent.Create(handler)
    window = PipeOneLineWindow(
        WebView2,
        CoreWebView2CreationProperties,
        initial_payload,
        handler,
        external_event
    )
    handler.window = window
    WINDOW_REFS.append(window)
    window.Show()
