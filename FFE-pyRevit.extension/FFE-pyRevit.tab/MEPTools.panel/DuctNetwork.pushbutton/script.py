# -*- coding: utf-8 -*-
__title__     = "Duct Network \nSummary"
__version__   = "Version = v2.0"
__doc__       = """Version = v2.0
Date    = 04.22.2026
______________________________________________________________
Description:
-> Creates tables for the selected duct system and traces airflow
   paths using Revit sections plus filtered physical connectors.
______________________________________________________________
How-to:
-> Select any element in a supply, return, or exhaust duct system
______________________________________________________________
Last update:
- [07.08.2025] - v0.1 BETA RELEASE
- [01.08.2026] - v0.2 BETA - Changed to Duct Network Summary
- [02.02.2026] - v1.0 RELEASE
- [02.09.2026] - v1.1 WORKS ON SUPPLY, RETURN, & EXHAUST
- [04.22.2026] - v2.0 REWRITE - Section graph tracing
______________________________________________________________
Author: Kyle Guggenheim"""

# ____________________________________________________________________ IMPORTS (SYSTEM)
from System import String
from collections import defaultdict, deque
import cgi
import time


# ____________________________________________________________________ IMPORTS (AUTODESK)
import clr
clr.AddReference("System")

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.ExtensibleStorage import Schema
from Autodesk.Revit.DB.Mechanical import DuctSystemType


# ____________________________________________________________________ IMPORTS (PYREVIT)
from pyrevit import revit, UI, script, forms
from pyrevit.script import output


# ____________________________________________________________________ VARIABLES
app         = __revit__.Application
uidoc       = __revit__.ActiveUIDocument
doc         = __revit__.ActiveUIDocument.Document
selection   = uidoc.Selection

output_window = output.get_output()

DEBUG_OUTPUT = False

SUPPORTED_SYSTEM_TYPES = (
    DuctSystemType.SupplyAir,
    DuctSystemType.ReturnAir,
    DuctSystemType.ExhaustAir,
)

LOGICAL_CONNECTOR_TYPES = set([
    ConnectorType.Logical,
    ConnectorType.Reference,
    ConnectorType.NodeReference,
    ConnectorType.Invalid,
])

HTML_TABLE_STYLE = """
<style>
table {
    border-collapse: collapse;
    width: 100%;
    margin-bottom: 16px;
}
th {
    background: #1f2f3f;
    color: white;
    padding: 4px;
    text-align: left;
}
td {
    border: 1px solid #999;
    padding: 4px;
    text-align: left;
}
caption {
    font-size: 16px;
    font-weight: bold;
    text-align: left;
    margin: 6px 0;
}
</style>
"""


# Revit 2026 behavior used by this rewrite:
# - MEPSystem.GetCriticalPathSectionNumbers() returns sections in flow order.
# - Connector.AllRefs includes logical refs, so physical filtering is required.
# - Connector.Direction is calculated from the system and may be unavailable.
# - MEPSection members can belong to multiple sections.


# ____________________________________________________________________ HELPERS
def eid_key(eid_or_elem):
    """Return a stable string key for an element or ElementId."""
    try:
        return eid_or_elem.Id.ToString()
    except:
        try:
            return eid_or_elem.ToString()
        except:
            return str(eid_or_elem)


def safe_str(value):
    """Convert a value to string safely."""
    if value is None:
        return ""
    try:
        return str(value)
    except:
        try:
            return value.ToString()
        except:
            return ""


def html_escape(value):
    """Escape plain-text HTML values."""
    return cgi.escape(safe_str(value), True)


def ordered_unique(sequence):
    """Return ordered unique values from a sequence."""
    seen = set()
    output_values = []
    for item in sequence:
        if item in seen:
            continue
        seen.add(item)
        output_values.append(item)
    return output_values


def is_close(x, y, epsilon=1e-9):
    """Return True when two numeric values are close."""
    try:
        return abs(x - y) <= epsilon
    except:
        return False


def safe_get_category_name(element):
    """Return the Revit category name or an empty string."""
    try:
        if element and element.Category:
            return element.Category.Name
    except:
        pass
    return ""


def safe_get_parameter_string(param):
    """Return parameter AsString safely."""
    try:
        value = param.AsString()
        return value if value is not None else ""
    except:
        return ""


def get_mark(element):
    """Return the element mark value."""
    try:
        return safe_get_parameter_string(element.get_Parameter(BuiltInParameter.ALL_MODEL_MARK))
    except:
        return ""


def get_comments(element):
    """Return the instance comments value."""
    try:
        return safe_get_parameter_string(element.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS))
    except:
        return ""


def get_display_name(element):
    """Prefer mark, then element name, then element id."""
    mark = get_mark(element)
    if mark:
        return mark

    try:
        name = element.Name
        if name:
            return name
    except:
        pass

    return eid_key(element)


def get_size_text(element):
    """Return a display size value for the current element."""
    category_name = safe_get_category_name(element)
    parameter_name = None

    if category_name == "Flex Ducts":
        parameter_name = "Overall Size"
    elif category_name != "Mechanical Equipment":
        parameter_name = "Size"

    if not parameter_name:
        return ""

    try:
        param = element.LookupParameter(parameter_name)
        if param:
            value = param.AsString()
            return value if value is not None else ""
    except:
        pass

    return ""


def convert_units(value, units_name):
    """
    Convert Revit internal units to imperial display units.

    units_name:
    - pressure
    - air flow
    - length
    - velocity
    - friction
    """
    if value is None:
        return None

    if units_name == "pressure":
        target_units = UnitTypeId.InchesOfWater60DegreesFahrenheit
    elif units_name == "air flow":
        target_units = UnitTypeId.CubicFeetPerMinute
    elif units_name == "length":
        target_units = UnitTypeId.Feet
    elif units_name == "velocity":
        target_units = UnitTypeId.FeetPerMinute
    elif units_name == "friction":
        target_units = UnitTypeId.InchesOfWater60DegreesFahrenheitPer100Feet
    else:
        target_units = None

    if target_units is None:
        return value

    try:
        return UnitUtils.ConvertFromInternalUnits(value, target_units)
    except:
        return None


def format_number(value, precision):
    """Format a numeric value to the requested precision."""
    if value is None:
        return ""
    try:
        return ("{0:." + str(precision) + "f}").format(value)
    except:
        return ""


def safe_get_pressure_drop(section, element_id):
    """Return section pressure drop for a specific element."""
    try:
        return convert_units(section.GetPressureDrop(element_id), "pressure")
    except:
        return None


def safe_get_segment_length(section, element_id):
    """Return section segment length for a specific element."""
    try:
        return convert_units(section.GetSegmentLength(element_id), "length")
    except:
        return None


def find_coefficient_schema():
    """Return the CoefficientFromTable schema if it exists."""
    for schema in Schema.ListSchemas():
        if schema.SchemaName == "CoefficientFromTable":
            return schema
    return None


def get_ashrae_code(fitting, coeff_schema):
    """Return the ASHRAE table code stored in extensible storage."""
    if coeff_schema is None or fitting is None:
        return None

    try:
        entity = fitting.GetEntity(coeff_schema)
    except:
        return None

    try:
        if not (entity and entity.IsValid()):
            return None
    except:
        return None

    field = coeff_schema.GetField("ASHRAETableName")
    if field is None:
        return None

    try:
        table_name = entity.Get[String](field)
    except:
        try:
            table_name = entity.Get(field)
        except:
            table_name = None

    return table_name if table_name else None


def get_part_type_text(element):
    """Return a duct fitting part type when available."""
    category_name = safe_get_category_name(element)
    if category_name != "Duct Fittings":
        return ""

    try:
        return safe_str(element.MEPModel.PartType)
    except:
        return ""


def get_connectors_from_element(element):
    """Return connectors exposed through MEPModel or ConnectorManager."""
    connectors = []
    seen_connector_ids = set()

    if element is None:
        return connectors

    mep_model = getattr(element, "MEPModel", None)
    if mep_model is not None:
        try:
            manager = mep_model.ConnectorManager
            if manager:
                for connector in manager.Connectors:
                    key = safe_str(getattr(connector, "Id", None))
                    if key not in seen_connector_ids:
                        connectors.append(connector)
                        seen_connector_ids.add(key)
        except:
            pass

    manager = getattr(element, "ConnectorManager", None)
    if manager is not None:
        try:
            for connector in manager.Connectors:
                key = safe_str(getattr(connector, "Id", None))
                if key not in seen_connector_ids:
                    connectors.append(connector)
                    seen_connector_ids.add(key)
        except:
            pass

    return connectors


def try_get_connector_system(connector):
    """Return the connector MEP system when available."""
    try:
        return connector.MEPSystem
    except:
        return None


def safe_get_system_type(mep_system):
    """Return the system type safely."""
    try:
        return mep_system.SystemType
    except:
        return None


def get_mep_system(element):
    """Resolve a single mechanical duct system from an element."""
    systems = {}

    try:
        direct_system = element.MEPSystem
    except:
        direct_system = None

    if direct_system is not None:
        systems[eid_key(direct_system)] = direct_system

    for connector in get_connectors_from_element(element):
        connector_system = try_get_connector_system(connector)
        if connector_system is not None:
            systems[eid_key(connector_system)] = connector_system

    if not systems:
        return None

    supported = {}
    for system_key, mep_system in systems.items():
        if safe_get_system_type(mep_system) in SUPPORTED_SYSTEM_TYPES:
            supported[system_key] = mep_system

    if len(supported) == 1:
        return list(supported.values())[0]

    if len(supported) > 1:
        return None

    if len(systems) == 1:
        return list(systems.values())[0]

    return None


def connector_is_supported_physical(connector):
    """Return True when the connector can represent a physical connection."""
    if connector is None:
        return False

    try:
        connector_type = connector.ConnectorType
    except:
        return False

    if connector_type in LOGICAL_CONNECTOR_TYPES:
        return False

    try:
        return connector.IsConnected
    except:
        return False


def connector_belongs_to_system(connector, mep_system):
    """Return True when the connector belongs to the requested system."""
    if connector is None or mep_system is None:
        return False

    connector_system = try_get_connector_system(connector)
    if connector_system is not None:
        try:
            return connector_system.Id.IntegerValue == mep_system.Id.IntegerValue
        except:
            try:
                return connector_system.Name == mep_system.Name
            except:
                pass

    try:
        return connector.DuctSystemType == mep_system.SystemType
    except:
        return False


def get_connector_flow_direction(connector):
    """Return FlowDirectionType.In/Out when available."""
    direction = None

    try:
        direction = connector.Direction
    except:
        direction = None

    if direction in (FlowDirectionType.In, FlowDirectionType.Out):
        return direction

    try:
        direction = connector.AssignedFlowDirection
    except:
        direction = None

    if direction in (FlowDirectionType.In, FlowDirectionType.Out):
        return direction

    return None


def iter_physical_refs(connector, mep_system):
    """Yield physical connector references on the same duct system."""
    if not connector_is_supported_physical(connector):
        return

    if mep_system is not None and not connector_belongs_to_system(connector, mep_system):
        return

    try:
        references = connector.AllRefs
    except:
        references = None

    if not references:
        return

    current_owner_key = ""
    try:
        current_owner_key = eid_key(connector.Owner)
    except:
        current_owner_key = ""

    for ref_connector in references:
        if not connector_is_supported_physical(ref_connector):
            continue

        if mep_system is not None and not connector_belongs_to_system(ref_connector, mep_system):
            continue

        try:
            neighbor_owner = ref_connector.Owner
        except:
            neighbor_owner = None

        if neighbor_owner is None:
            continue

        if eid_key(neighbor_owner) == current_owner_key:
            continue

        yield ref_connector


def get_system_sections(mep_system):
    """Collect system sections using GetSectionByIndex."""
    sections = []

    try:
        count = mep_system.SectionsCount
    except:
        return sections

    for index in range(0, count):
        try:
            section = mep_system.GetSectionByIndex(index)
        except:
            section = None
        if section is not None:
            sections.append(section)

    return sections


def build_system_snapshot(mep_system):
    """Build section and element catalogs for the selected system."""
    system_sections = get_system_sections(mep_system)
    if not system_sections:
        return None

    sections_by_number = {}
    elements_by_section = {}
    elem_sections = defaultdict(set)
    element_cache = {}

    airflow_by_section = {}
    pressure_drop_by_section = {}
    velocity_by_section = {}
    friction_by_section = {}
    section_graph = defaultdict(set)

    for section in system_sections:
        section_number = section.Number
        sections_by_number[section_number] = section

        element_ids = []
        try:
            element_ids = section.GetElementIds()
        except:
            element_ids = []

        section_elements = []
        for element_id in element_ids:
            element = doc.GetElement(element_id)
            if element is None:
                continue
            section_elements.append(element)
            element_cache[eid_key(element)] = element
            elem_sections[eid_key(element)].add(section_number)

        elements_by_section[section_number] = section_elements
        airflow_by_section[section_number] = convert_units(section.Flow, "air flow")
        pressure_drop_by_section[section_number] = convert_units(section.TotalPressureLoss, "pressure")
        velocity_by_section[section_number] = convert_units(section.Velocity, "velocity")
        friction_by_section[section_number] = convert_units(section.Friction, "friction")

    for element_key, section_numbers in elem_sections.items():
        ordered_sections = sorted(list(section_numbers))
        if len(ordered_sections) < 2:
            continue

        element = element_cache.get(element_key)
        category_name = safe_get_category_name(element)

        if len(ordered_sections) == 2:
            section_a = ordered_sections[0]
            section_b = ordered_sections[1]
            section_graph[section_a].add(section_b)
            section_graph[section_b].add(section_a)
            continue

        section_flows = []
        for section_number in ordered_sections:
            section_flows.append((airflow_by_section.get(section_number), section_number))

        valid_flow_pairs = [pair for pair in section_flows if pair[0] is not None]

        # Multi-section ducts usually represent consecutive physical runs, while
        # multi-section fittings usually represent one main section feeding one
        # or more branch sections. They need different graph shapes.
        if len(valid_flow_pairs) == len(ordered_sections):
            reverse_sort = safe_get_system_type(mep_system) == DuctSystemType.SupplyAir
            ordered_by_flow = sorted(
                valid_flow_pairs,
                key=lambda pair: ((-pair[0]) if reverse_sort else pair[0], pair[1])
            )

            ordered_flow_sections = [section_number for flow_value, section_number in ordered_by_flow]
            if category_name in ["Ducts", "Flex Ducts"]:
                for index in range(len(ordered_flow_sections) - 1):
                    section_a = ordered_flow_sections[index]
                    section_b = ordered_flow_sections[index + 1]
                    section_graph[section_a].add(section_b)
                    section_graph[section_b].add(section_a)
                continue

            dominant_flow = ordered_by_flow[0][0]
            dominant_sections = sorted([
                section_number
                for flow_value, section_number in ordered_by_flow
                if is_close(flow_value, dominant_flow)
            ])

            if len(dominant_sections) == 1:
                dominant_section = dominant_sections[0]
                for section_number in ordered_flow_sections:
                    if section_number == dominant_section:
                        continue
                    section_graph[dominant_section].add(section_number)
                    section_graph[section_number].add(dominant_section)
                continue

        for index in range(len(ordered_sections) - 1):
            section_a = ordered_sections[index]
            section_b = ordered_sections[index + 1]
            section_graph[section_a].add(section_b)
            section_graph[section_b].add(section_a)

    for section_number in sections_by_number.keys():
        if section_number not in section_graph:
            section_graph[section_number] = set()

    base_equipment = None
    base_equipment_connector = None
    critical_path_sections = []
    critical_path_pressure_loss = None
    system_airflow = None
    is_well_connected = False
    is_multiple_network = False
    physical_networks = None

    try:
        base_equipment = mep_system.BaseEquipment
    except:
        base_equipment = None

    try:
        base_equipment_connector = mep_system.BaseEquipmentConnector
    except:
        base_equipment_connector = None

    try:
        critical_path_sections = list(mep_system.GetCriticalPathSectionNumbers())
    except:
        critical_path_sections = []

    try:
        critical_path_pressure_loss = convert_units(mep_system.PressureLossOfCriticalPath, "pressure")
    except:
        critical_path_pressure_loss = None

    try:
        system_airflow = convert_units(mep_system.GetFlow(), "air flow")
    except:
        system_airflow = None

    try:
        is_well_connected = bool(mep_system.IsWellConnected)
    except:
        is_well_connected = False

    try:
        is_multiple_network = bool(mep_system.IsMultipleNetwork)
    except:
        is_multiple_network = False

    try:
        physical_networks = mep_system.GetPhysicalNetworksNumber()
    except:
        physical_networks = None

    ordered_element_keys = sorted(
        list(element_cache.keys()),
        key=lambda key: int(key) if safe_str(key).isdigit() else safe_str(key)
    )

    all_elements = [element_cache[key] for key in ordered_element_keys]

    snapshot = {
        "system": mep_system,
        "system_name": mep_system.Name,
        "system_type": safe_get_system_type(mep_system),
        "sections_by_number": sections_by_number,
        "elements_by_section": elements_by_section,
        "elem_sections": elem_sections,
        "section_graph": section_graph,
        "airflow_by_section": airflow_by_section,
        "pressure_drop_by_section": pressure_drop_by_section,
        "velocity_by_section": velocity_by_section,
        "friction_by_section": friction_by_section,
        "elements": all_elements,
        "allowed_element_ids": set(elem_sections.keys()),
        "base_equipment": base_equipment,
        "base_equipment_id": eid_key(base_equipment) if base_equipment is not None else None,
        "base_equipment_connector": base_equipment_connector,
        "critical_path_sections": critical_path_sections,
        "critical_path_index": dict(
            (section_number, index) for index, section_number in enumerate(critical_path_sections)
        ),
        "critical_path_pressure_loss": critical_path_pressure_loss,
        "system_airflow": system_airflow,
        "is_well_connected": is_well_connected,
        "is_multiple_network": is_multiple_network,
        "physical_networks": physical_networks,
    }

    return snapshot


def get_element_data(element, system_name):
    """Extract element table data that does not depend on section membership."""
    data = {}
    data["System Name"] = system_name
    data["Category"] = safe_get_category_name(element) or "N/A"
    data["Part Type"] = get_part_type_text(element)
    data["Element ID"] = output_window.linkify(element.Id)
    data["Section"] = ""
    data["Mark"] = get_mark(element)
    data["ASHRAE Table"] = ""
    data["Comments"] = get_comments(element)
    data["Flow (CFM)"] = ""
    data["Size"] = get_size_text(element)
    data["Length (ft)"] = ""
    data["Velocity (FPM)"] = ""
    data["Friction (in-wg/100ft)"] = ""
    data["Pressure Drop (in-wg)"] = ""
    return data


def build_element_table_rows(snapshot, coeff_schema):
    """Build one table row per element-section membership."""
    column_order = [
        "System Name",
        "Category",
        "Part Type",
        "Element ID",
        "Section",
        "Mark",
        "ASHRAE Table",
        "Comments",
        "Flow (CFM)",
        "Size",
        "Length (ft)",
        "Velocity (FPM)",
        "Friction (in-wg/100ft)",
        "Pressure Drop (in-wg)",
    ]

    rows = []

    ordered_sections = sorted(list(snapshot["elements_by_section"].keys()))
    for section_number in ordered_sections:
        section = snapshot["sections_by_number"][section_number]
        section_elements = sorted(
            snapshot["elements_by_section"][section_number],
            key=lambda elem: eid_key(elem)
        )

        for element in section_elements:
            element_data = get_element_data(element, snapshot["system_name"])
            category_name = safe_get_category_name(element)

            element_data["Section"] = section_number
            element_data["Flow (CFM)"] = format_number(snapshot["airflow_by_section"].get(section_number), 0)

            if category_name in ["Ducts", "Flex Ducts"]:
                element_data["Velocity (FPM)"] = format_number(snapshot["velocity_by_section"].get(section_number), 2)
                element_data["Friction (in-wg/100ft)"] = format_number(snapshot["friction_by_section"].get(section_number), 4)

            pressure_drop = safe_get_pressure_drop(section, element.Id)
            element_data["Pressure Drop (in-wg)"] = format_number(pressure_drop, 4)

            segment_length = safe_get_segment_length(section, element.Id)
            element_data["Length (ft)"] = format_number(segment_length, 4)

            if category_name in ["Duct Fittings", "Duct Accessories"]:
                ashrae_code = get_ashrae_code(element, coeff_schema)
                element_data["ASHRAE Table"] = ashrae_code if ashrae_code else "<no ASHRAE table set>"

            rows.append([element_data.get(column_name, "") for column_name in column_order])

    return column_order, rows


def is_root_direction(direction, system_type):
    """Return True when the direction points away from the root equipment."""
    if direction is None:
        return False

    if system_type == DuctSystemType.SupplyAir:
        return direction == FlowDirectionType.Out

    return direction == FlowDirectionType.In


def is_endpoint_direction(direction, system_type):
    """Return True when the direction corresponds to a downstream sink endpoint."""
    if direction is None:
        return False

    if system_type == DuctSystemType.SupplyAir:
        return direction == FlowDirectionType.In

    return direction == FlowDirectionType.Out


def get_connector_anchor_sections(element, connector, snapshot, allow_neighbor_fallback):
    """
    Map a connector to likely section numbers by intersecting the owner's
    sections with connected neighbor sections.
    """
    current_sections = set(snapshot["elem_sections"].get(eid_key(element), set()))
    matched_sections = set()
    neighbor_sections_seen = set()

    for ref_connector in iter_physical_refs(connector, snapshot["system"]):
        try:
            neighbor_owner = ref_connector.Owner
        except:
            neighbor_owner = None

        if neighbor_owner is None:
            continue

        neighbor_key = eid_key(neighbor_owner)
        if neighbor_key not in snapshot["allowed_element_ids"]:
            continue

        neighbor_sections = set(snapshot["elem_sections"].get(neighbor_key, set()))
        if not neighbor_sections:
            continue

        neighbor_sections_seen.update(neighbor_sections)

        if current_sections:
            shared_sections = current_sections & neighbor_sections
            matched_sections.update(shared_sections)

    if matched_sections:
        return set([section for section in matched_sections if section in snapshot["sections_by_number"]])

    if not current_sections and allow_neighbor_fallback and neighbor_sections_seen:
        return set([section for section in neighbor_sections_seen if section in snapshot["sections_by_number"]])

    return set()


def resolve_root_sections(snapshot, diagnostics):
    """Resolve root sections from base equipment or a unique inferred equipment root."""
    base_equipment = snapshot["base_equipment"]
    base_connector = snapshot["base_equipment_connector"]

    if base_equipment is not None and base_connector is not None:
        root_sections = get_connector_anchor_sections(
            base_equipment,
            base_connector,
            snapshot,
            allow_neighbor_fallback=True
        )
        if root_sections:
            return sorted(list(root_sections)), base_equipment, "base equipment connector"

        diagnostics.append(("System root", "Base equipment connector could not be mapped to any section."))
    else:
        if base_equipment is None:
            diagnostics.append(("System root", "System has no base equipment."))
        if base_connector is None:
            diagnostics.append(("System root", "System has no base equipment connector."))

    inferred_candidates = []
    for element in snapshot["elements"]:
        if safe_get_category_name(element) != "Mechanical Equipment":
            continue

        candidate_sections = set()
        for connector in get_connectors_from_element(element):
            if not connector_belongs_to_system(connector, snapshot["system"]):
                continue

            if not is_root_direction(get_connector_flow_direction(connector), snapshot["system_type"]):
                continue

            candidate_sections.update(
                get_connector_anchor_sections(
                    element,
                    connector,
                    snapshot,
                    allow_neighbor_fallback=True
                )
            )

        if candidate_sections:
            inferred_candidates.append((element, sorted(list(candidate_sections))))

    if len(inferred_candidates) == 1:
        element, sections = inferred_candidates[0]
        diagnostics.append(("System root", "Using uniquely inferred equipment root: {}".format(get_display_name(element))))
        return sections, element, "inferred equipment root"

    if len(inferred_candidates) == 0:
        diagnostics.append(("System root", "No uniquely inferable equipment root was found."))
    else:
        diagnostics.append(("System root", "Multiple possible equipment roots were found."))

    return [], None, None


def collect_endpoint_candidates(snapshot):
    """Collect air terminals and downstream mechanical equipment candidates."""
    candidates = []
    seen = set()

    for element in snapshot["elements"]:
        category_name = safe_get_category_name(element)
        element_key = eid_key(element)

        if category_name == "Air Terminals":
            if element_key not in seen:
                candidates.append(element)
                seen.add(element_key)
        elif category_name == "Mechanical Equipment":
            if snapshot["base_equipment_id"] == element_key:
                continue
            if element_key not in seen:
                candidates.append(element)
                seen.add(element_key)

    return candidates


def resolve_endpoint_sections(element, snapshot):
    """Resolve endpoint anchor sections from sink-direction connectors."""
    element_sections = set(snapshot["elem_sections"].get(eid_key(element), set()))
    anchor_sections = set()
    has_same_system_connector = False
    has_sink_connector = False

    for connector in get_connectors_from_element(element):
        if not connector_belongs_to_system(connector, snapshot["system"]):
            continue

        has_same_system_connector = True
        direction = get_connector_flow_direction(connector)
        if not is_endpoint_direction(direction, snapshot["system_type"]):
            continue

        has_sink_connector = True
        anchor_sections.update(
            get_connector_anchor_sections(
                element,
                connector,
                snapshot,
                allow_neighbor_fallback=False
            )
        )

    anchor_sections = set([section for section in anchor_sections if section in snapshot["sections_by_number"]])
    if anchor_sections:
        return sorted(list(anchor_sections)), None

    if len(element_sections) == 1 and has_same_system_connector:
        return sorted(list(element_sections)), "single-section fallback"

    if not element_sections:
        return [], "Element is not assigned to any analyzed section."

    if not has_same_system_connector:
        return [], "Element has no connectors on the selected system."

    if not has_sink_connector:
        return [], "Element has no sink-direction connectors on the selected system."

    return [], "Endpoint connector could not be mapped to a section."


def backtrack_shortest_paths(node, predecessors, source_sections, partial_path, results, limit):
    """Collect up to limit shortest paths from any source to node."""
    if len(results) >= limit:
        return

    if node in source_sections:
        candidate = list(reversed(partial_path + [node]))
        results.append(candidate)
        return

    for previous_node in sorted(list(predecessors.get(node, set()))):
        backtrack_shortest_paths(
            previous_node,
            predecessors,
            source_sections,
            partial_path + [node],
            results,
            limit
        )
        if len(results) >= limit:
            return


def find_unique_shortest_path(section_graph, source_sections, target_sections):
    """Return a unique shortest path or an ambiguity/error reason."""
    source_sections = set(source_sections)
    target_sections = set(target_sections)

    if not source_sections:
        return None, "No root sections were resolved."

    if not target_sections:
        return None, "No endpoint sections were resolved."

    distances = {}
    predecessors = defaultdict(set)
    queue = deque()

    for section_number in sorted(list(source_sections)):
        distances[section_number] = 0
        queue.append(section_number)

    while queue:
        current = queue.popleft()
        for neighbor in sorted(list(section_graph.get(current, set()))):
            next_distance = distances[current] + 1
            if neighbor not in distances:
                distances[neighbor] = next_distance
                predecessors[neighbor].add(current)
                queue.append(neighbor)
            elif distances[neighbor] == next_distance:
                predecessors[neighbor].add(current)

    reachable_targets = [section for section in sorted(list(target_sections)) if section in distances]
    if not reachable_targets:
        return None, "No path from the root network reaches the endpoint anchor."

    shortest_distance = min([distances[section] for section in reachable_targets])
    best_targets = [section for section in reachable_targets if distances[section] == shortest_distance]

    all_paths = []
    for target_section in best_targets:
        backtrack_shortest_paths(
            target_section,
            predecessors,
            source_sections,
            [],
            all_paths,
            2
        )
        if len(all_paths) >= 2:
            break

    unique_paths = []
    seen_paths = set()
    for path in all_paths:
        key = tuple(path)
        if key not in seen_paths:
            unique_paths.append(path)
            seen_paths.add(key)

    if len(unique_paths) != 1:
        return None, "Multiple equally short section paths were found."

    return unique_paths[0], None


def orient_path_for_flow(path_sections, system_type):
    """Orient path sections to match the direction of airflow."""
    if system_type == DuctSystemType.SupplyAir:
        return list(path_sections)
    return list(reversed(path_sections))


def validate_flow_monotonicity(path_sections, airflow_by_section, system_type):
    """Validate section airflow ordering along the flow path."""
    section_flows = [airflow_by_section.get(section) for section in path_sections]
    if None in section_flows:
        return False, "One or more sections do not have airflow values."

    for index in range(len(section_flows) - 1):
        current_flow = section_flows[index]
        next_flow = section_flows[index + 1]

        if system_type == DuctSystemType.SupplyAir:
            if next_flow > current_flow and not is_close(next_flow, current_flow):
                return False, "Supply airflow increases along the traced path."
        else:
            if next_flow < current_flow and not is_close(next_flow, current_flow):
                return False, "Return or exhaust airflow decreases along the traced path."

    return True, None


def validate_critical_path_overlap(path_sections, critical_path_index):
    """Validate that any overlap preserves Revit critical path order."""
    overlapping_indexes = []
    for section_number in path_sections:
        if section_number in critical_path_index:
            overlapping_indexes.append(critical_path_index[section_number])

    if len(overlapping_indexes) < 2:
        return True, None

    for index in range(len(overlapping_indexes) - 1):
        if overlapping_indexes[index + 1] <= overlapping_indexes[index]:
            return False, "Critical path overlap does not preserve Revit flow order."

    return True, None


def build_path_records(snapshot, diagnostics):
    """Build validated flow-ordered section paths for endpoint elements."""
    path_records = []
    root_sections, root_element, root_method = resolve_root_sections(snapshot, diagnostics)
    endpoint_candidates = collect_endpoint_candidates(snapshot)

    if not root_sections:
        return {
            "records": path_records,
            "root_sections": [],
            "root_element": root_element,
            "root_method": root_method,
            "endpoint_candidates": endpoint_candidates,
            "reachable_sections": set(),
        }

    reachable_sections = set()
    queue = deque()
    for section_number in root_sections:
        queue.append(section_number)
        reachable_sections.add(section_number)

    while queue:
        current = queue.popleft()
        for neighbor in snapshot["section_graph"].get(current, set()):
            if neighbor in reachable_sections:
                continue
            reachable_sections.add(neighbor)
            queue.append(neighbor)

    root_label = get_display_name(root_element) if root_element is not None else "<unresolved>"

    for endpoint in endpoint_candidates:
        if root_element is not None and eid_key(endpoint) == eid_key(root_element):
            continue

        endpoint_label = get_display_name(endpoint)
        endpoint_sections, endpoint_note = resolve_endpoint_sections(endpoint, snapshot)

        if not endpoint_sections:
            diagnostics.append((endpoint_label, endpoint_note))
            continue

        path_sections, path_error = find_unique_shortest_path(
            snapshot["section_graph"],
            root_sections,
            endpoint_sections
        )

        if path_sections is None:
            diagnostics.append((endpoint_label, path_error))
            continue

        flow_path_sections = ordered_unique(
            orient_path_for_flow(path_sections, snapshot["system_type"])
        )

        valid, validation_error = validate_flow_monotonicity(
            flow_path_sections,
            snapshot["airflow_by_section"],
            snapshot["system_type"]
        )
        if not valid:
            diagnostics.append((endpoint_label, validation_error))
            continue

        valid, validation_error = validate_critical_path_overlap(
            flow_path_sections,
            snapshot["critical_path_index"]
        )
        if not valid:
            diagnostics.append((endpoint_label, validation_error))
            continue

        total_pressure_loss = 0.0
        missing_pressure_loss = False
        for section_number in flow_path_sections:
            section_pressure_loss = snapshot["pressure_drop_by_section"].get(section_number)
            if section_pressure_loss is None:
                missing_pressure_loss = True
                break
            total_pressure_loss += section_pressure_loss

        if missing_pressure_loss:
            diagnostics.append((endpoint_label, "One or more sections in the path do not have pressure loss values."))
            continue

        if snapshot["system_type"] == DuctSystemType.SupplyAir:
            from_label = root_label
            to_label = endpoint_label
        else:
            from_label = endpoint_label
            to_label = root_label

        path_records.append({
            "endpoint": endpoint,
            "endpoint_label": endpoint_label,
            "from_label": from_label,
            "to_label": to_label,
            "sections": flow_path_sections,
            "root_sections": list(root_sections),
            "endpoint_sections": list(endpoint_sections),
            "pressure_loss": total_pressure_loss,
            "endpoint_note": endpoint_note,
        })

    path_records = sorted(
        path_records,
        key=lambda item: (len(item["sections"]), item["to_label"], item["endpoint_label"]),
        reverse=True
    )

    return {
        "records": path_records,
        "root_sections": list(root_sections),
        "root_element": root_element,
        "root_method": root_method,
        "endpoint_candidates": endpoint_candidates,
        "reachable_sections": reachable_sections,
    }


def ashrae_color(value):
    """Highlight one specific ASHRAE code used by the original script."""
    if value == "CD3-11":
        return "#E58D33"
    return None


def tpl_to_red_hex(value, min_value, max_value):
    """Return a red color ramp for pressure loss values."""
    if max_value == min_value:
        weight = 1.0
    else:
        weight = (value - min_value) / float(max_value - min_value)

    red = 255
    green = int(255 * (1 - weight))
    blue = int(255 * (1 - weight))
    return "#{:02X}{:02X}{:02X}".format(red, green, blue)


def render_elements_table(headers, rows):
    """Render the duct network elements table as HTML."""
    html = [HTML_TABLE_STYLE]
    html.append("<table>")
    html.append("<caption>{}</caption>".format(html_escape("Duct Network Elements: {}".format(len(rows)))))
    html.append("<tr>")
    for header in headers:
        html.append("<th>{}</th>".format(html_escape(header)))
    html.append("</tr>")

    for row in rows:
        html.append("<tr>")
        for index, cell in enumerate(row):
            header = headers[index]
            style_bits = []

            if header == "ASHRAE Table":
                color = ashrae_color(cell)
                if color:
                    style_bits.append("background:{}".format(color))
                    style_bits.append("font-weight:bold")

            if style_bits:
                style_attr = " style='{}'".format("; ".join(style_bits))
            else:
                style_attr = ""

            if header == "Element ID":
                html.append("<td{}>{}</td>".format(style_attr, cell if cell else ""))
            else:
                html.append("<td{}>{}</td>".format(style_attr, html_escape(cell)))
        html.append("</tr>")

    html.append("</table>")
    output_window.print_html("".join(html))


def render_paths_table(path_records):
    """Render the validated network paths table as HTML."""
    if not path_records:
        output_window.print_md("### Duct Network Paths: 0")
        output_window.print_md("- No validated airflow paths were produced.")
        return

    max_path_length = max([len(record["sections"]) for record in path_records])
    headers = ["From", "To", "TPL"] + ["path{}".format(index + 1) for index in range(max_path_length)]

    tpl_values = [record["pressure_loss"] for record in path_records]
    tpl_min = min(tpl_values)
    tpl_max = max(tpl_values)

    html = [HTML_TABLE_STYLE]
    html.append("<table>")
    html.append("<caption>{}</caption>".format(html_escape("Duct Network Paths: {}".format(len(path_records)))))
    html.append("<tr>")
    for header in headers:
        html.append("<th>{}</th>".format(html_escape(header)))
    html.append("</tr>")

    for record in path_records:
        row = [
            record["from_label"],
            record["to_label"],
            record["pressure_loss"],
        ] + list(record["sections"]) + [""] * (max_path_length - len(record["sections"]))

        html.append("<tr>")
        for index, cell in enumerate(row):
            header = headers[index]
            if header == "TPL":
                cell_color = tpl_to_red_hex(record["pressure_loss"], tpl_min, tpl_max)
                html.append(
                    "<td style='background:{}; font-weight:bold;'>{}</td>".format(
                        cell_color,
                        html_escape(format_number(cell, 4))
                    )
                )
            else:
                html.append("<td>{}</td>".format(html_escape(cell)))
        html.append("</tr>")

    html.append("</table>")
    output_window.print_html("".join(html))


def render_system_summary(snapshot, path_context, diagnostics):
    """Render markdown summary and diagnostics."""
    output_window.print_md("# Duct Network Summary")
    output_window.print_md("### Timestamp: {}".format(time.strftime("%Y-%m-%d %H:%M:%S")))
    output_window.print_md("### System Name: {}".format(snapshot["system_name"]))
    output_window.print_md("### System Type: {}".format(snapshot["system_type"]))
    output_window.print_md(
        "### System Connection Status: {}".format(
            "Well Connected" if snapshot["is_well_connected"] else "Not Well Connected"
        )
    )

    if snapshot["physical_networks"] is not None:
        output_window.print_md("### Physical Networks: {}".format(snapshot["physical_networks"]))

    if snapshot["critical_path_sections"]:
        output_window.print_md("### Critical Path Sections: {}".format(snapshot["critical_path_sections"]))
    else:
        output_window.print_md("### Critical Path Sections: <unavailable>")

    if snapshot["critical_path_pressure_loss"] is not None:
        output_window.print_md(
            "### Critical Path Pressure Loss: {} in-wg".format(
                format_number(snapshot["critical_path_pressure_loss"], 4)
            )
        )
    else:
        output_window.print_md("### Critical Path Pressure Loss: <unavailable>")

    if snapshot["system_airflow"] is not None:
        output_window.print_md("### System Air Flow: {} CFM".format(format_number(snapshot["system_airflow"], 0)))
    else:
        output_window.print_md("### System Air Flow: <unavailable>")

    output_window.print_md("### Sections in Snapshot: {}".format(len(snapshot["sections_by_number"])))

    if path_context["root_element"] is not None:
        output_window.print_md("### Root Element: {}".format(get_display_name(path_context["root_element"])))
    else:
        output_window.print_md("### Root Element: <unresolved>")

    if path_context["root_sections"]:
        output_window.print_md("### Root Sections: {}".format(path_context["root_sections"]))
    else:
        output_window.print_md("### Root Sections: <unresolved>")

    output_window.print_md("---")
    output_window.print_md("### Trace Diagnostics")
    output_window.print_md("- Endpoint candidates: {}".format(len(path_context["endpoint_candidates"])))
    output_window.print_md("- Validated paths: {}".format(len(path_context["records"])))
    output_window.print_md(
        "- Reachable sections from root: {} / {}".format(
            len(path_context["reachable_sections"]),
            len(snapshot["sections_by_number"])
        )
    )

    if path_context["root_method"]:
        output_window.print_md("- Root resolution method: {}".format(path_context["root_method"]))

    if snapshot["is_multiple_network"]:
        output_window.print_md("- Revit reports multiple networks in this system.")

    if not snapshot["is_well_connected"]:
        output_window.print_md("- Revit reports this system is not well connected, so some calculated values may be invalid.")

    if diagnostics:
        for label, reason in diagnostics:
            output_window.print_md("- {}: {}".format(label, reason))
    else:
        output_window.print_md("- No skipped endpoints or root-resolution issues were detected.")

    if DEBUG_OUTPUT:
        total_edges = sum([len(neighbors) for neighbors in snapshot["section_graph"].values()]) / 2.0
        output_window.print_md("### Debug")
        output_window.print_md("- Section graph edges: {}".format(int(total_edges)))
        output_window.print_md("- Allowed elements in section catalog: {}".format(len(snapshot["allowed_element_ids"])))


def validate_selected_system(start_element):
    """Resolve and validate the selected duct system."""
    mep_system = get_mep_system(start_element)
    if mep_system is None:
        forms.alert(
            "Select an element that belongs to a single supply, return, or exhaust duct system.",
            title="Invalid Selection",
            exitscript=True
        )

    system_type = safe_get_system_type(mep_system)
    if system_type not in SUPPORTED_SYSTEM_TYPES:
        forms.alert(
            "This tool only supports supply, return, and exhaust duct systems.",
            title="Unsupported System",
            exitscript=True
        )

    return mep_system


# ____________________________________________________________________ MAIN
try:
    picked_ref = uidoc.Selection.PickObject(
        UI.Selection.ObjectType.Element,
        "Select any duct system element."
    )
    start_element = doc.GetElement(picked_ref)
except:
    script.exit()

mep_system = validate_selected_system(start_element)
snapshot = build_system_snapshot(mep_system)

if snapshot is None:
    forms.alert(
        "The selected duct system does not expose any Revit sections.",
        title="No Sections Found",
        exitscript=True
    )

coeff_schema = find_coefficient_schema()

diagnostics = []
if coeff_schema is None:
    diagnostics.append((
        "ASHRAE schema",
        "Could not find the 'CoefficientFromTable' extensible storage schema."
    ))

element_headers, element_rows = build_element_table_rows(snapshot, coeff_schema)
path_context = build_path_records(snapshot, diagnostics)

render_system_summary(snapshot, path_context, diagnostics)
render_elements_table(element_headers, element_rows)
render_paths_table(path_context["records"])
