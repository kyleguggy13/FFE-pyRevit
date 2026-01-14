# -*- coding: utf-8 -*-
__title__     = "Duct Network \nSummary"
__version__   = 'Version = 0.2'
__doc__       = """Version = 0.2
Date    = 01.08.2026
______________________________________________________________
Description:
-> Creates a table of the duct network selected for pressure loss calculations.
______________________________________________________________
How-to:
-> Select Straight or Flex duct
______________________________________________________________
Last update:
- [07.08.2025] - v0.1 BETA RELEASE
- [01.08.2026] - v0.2 BETA - Changed to Duct Network Summary
______________________________________________________________
Author: Kyle Guggenheim"""

#____________________________________________________________________ IMPORTS (SYSTEM)
from System import String
from collections import defaultdict



#____________________________________________________________________ IMPORTS (AUTODESK)
import sys
import clr
clr.AddReference("System")
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import FilteredElementCollector, Mechanical, Family, BuiltInParameter, ElementType, UnitTypeId
from Autodesk.Revit.DB import BuiltInCategory, ElementCategoryFilter, ElementId, FamilyInstance
from Autodesk.Revit.DB.ExtensibleStorage import Schema


#____________________________________________________________________ IMPORTS (PYREVIT)
from pyrevit import revit, DB, UI, script
from pyrevit.script import output
from pyrevit import forms

#____________________________________________________________________ VARIABLES
app         = __revit__.Application
uidoc       = __revit__.ActiveUIDocument
doc         = __revit__.ActiveUIDocument.Document   #type: Document
selection   = uidoc.Selection                       #type: Selection


output_window = output.get_output()
"""Output window for displaying results."""


#____________________________________________________________________ FUNCTIONS
def get_MEPSystem(element):
    """Get the MEPSystem name of the given element, if it exists."""
    if hasattr(element, 'MEPSystem'):
        mep_system = element.MEPSystem
        if mep_system:
            return mep_system
    return "N/A"


def get_MEPSystem_sections(element):
    """Get all sections in the duct network starting from the given element."""
    if hasattr(element, 'MEPSystem'):
        mep_system = element.MEPSystem            # Get the MEPSystem of the element
        if mep_system:
            sections_count = mep_system.SectionsCount   # Get Total number of sections
            
            # Get all sections related to MEPSystem
            system_sections = [mep_system.GetSectionByNumber(i) for i in range(1, sections_count + 1)]
            
            return system_sections
    
    return []


def get_MEPSection_elements(section):
    """Get all elements in a given duct section."""
    element_ids_list = section.GetElementIds()                          # Get all element IDs in section
    elements = [doc.GetElement(eid) for eid in element_ids_list]        # Retrieve elements from IDs
    return elements


def convertUnits(value, units):
    """
    Function to convert internal units to Imperial
    
    :param value:   Value to convert
    :param units:   ["pressure", "air flow", "length", "velocity", "friction"]
    :return:        Specified Unit
    """

    if      units == "pressure" : units = UnitTypeId.InchesOfWater60DegreesFahrenheit
    elif    units == "air flow" : units = UnitTypeId.CubicFeetPerMinute
    elif    units == "length"   : units = UnitTypeId.Feet
    elif    units == "velocity" : units = UnitTypeId.FeetPerMinute
    elif    units == "friction" : units = UnitTypeId.InchesOfWater60DegreesFahrenheitPer100Feet

    return UnitUtils.ConvertFromInternalUnits(value, units)


def get_MEPSection_PressureDrop(section, element):
    """Get pressure drop values for each element per section"""
    try:
        pressuredrop = section.GetPressureDrop(element)
    except:
        pressuredrop = "Invalid Section"

    return convertUnits(pressuredrop, "pressure")


def get_MEPSection_SegmentLength(section, element):
    """Get segment length values for each straigt/flex duct element per section"""
    try:
        segmentlength = section.GetSegmentLength(element)
        segmentlength = "{:.4f}".format(convertUnits(segmentlength, "length"))
        # print(segmentlength) # <- TESTING
    except:
        segmentlength = ""

    return segmentlength


def find_coefficient_schema():
    """Return the Schema whose name is 'CoefficientFromTable', or None."""
    for s in Schema.ListSchemas():
        if s.SchemaName == "CoefficientFromTable":
            return s
    return None


def get_ashrae_code(fitting, coeff_schema):
    """
    Return the ASHRAE table code for a duct fitting (e.g. 'SD5-3').

    Stored in ExtensibleStorage:
      SchemaName: 'CoefficientFromTable'
      Field:      'ASHRAETableName'
    """
    if coeff_schema is None or fitting is None:
        return None

    # Get the Entity attached to this element for that schema
    entity = fitting.GetEntity(coeff_schema)

    if not (entity and entity.IsValid()):
        return None

    # Get the field object
    field = coeff_schema.GetField("ASHRAETableName")
    if field is None:
        return None

    try:
        # Generic Get<T>(Field) – T is System.String
        table_name = entity.Get[String](field)
    except:
        # Fallback to non-generic overload if needed
        try:
            table_name = entity.Get(field)
        except:
            return None

    if not table_name:
        return None

    return table_name


def get_Mark(element):
    """
    Docstring for get_Mark
    
    :param element: Element (object)
    """
    mark = element.get_Parameter(BuiltInParameter.ALL_MODEL_MARK).AsString()
    
    if mark != None:
        value = mark
    else:
        value = ""

    return value


def get_element_data(element, SystemName):
    """Extract relevant data from a duct element."""
    # Initialize dict
    data = {}
    
    # Get Category
    data['Category'] = element.Category.Name if element.Category else "N/A"
    
    # Get Element ID
    data['Element ID'] = output_window.linkify(element.Id)

    # Get Comments
    comments = element.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString()
    if comments != None:
        data["Comments"] = comments
    else:
        data["Comments"] = ""
    
    # Get System Name
    data['System Name'] = SystemName
    
    # Get Size
    if element.Category.Name == "Flex Ducts":
        data['Size'] = element.LookupParameter("Overall Size").AsString()
    elif element.Category.Name != "Mechanical Equipment":
        data['Size'] = element.LookupParameter("Size").AsString()
    else:
        data["Size"] = ""
    # Add more parameters as needed
    return data


def get_connectors_from_element(elem):
    """
    Return a list of Revit MEP Connectors for an element (duct, fitting,
    accessory, terminal, etc.).
    Handles both MEPModel.ConnectorManager and direct ConnectorManager.
    """
    connectors = []

    # Many family instances (fittings, accessories, terminals) use MEPModel.ConnectorManager
    mep_model = getattr(elem, "MEPModel", None)
    if mep_model:
        try:
            conn_mgr = mep_model.ConnectorManager
            if conn_mgr:
                for c in conn_mgr.Connectors:
                    connectors.append(c)
        except:
            pass

    # Ducts / flex ducts expose ConnectorManager directly
    conn_mgr2 = getattr(elem, "ConnectorManager", None)
    if conn_mgr2:
        try:
            for c in conn_mgr2.Connectors:
                connectors.append(c)
        except:
            pass

    return connectors




#____________________________________________________________________ MAIN
uidoc = revit.uidoc
doc = revit.doc

# Select one duct element first
try:
    ref = uidoc.Selection.PickObject(UI.Selection.ObjectType.Element, "Select a duct.")
    start_element = doc.GetElement(ref)
    # print("start element: ", start_element, type(start_element)) # <- TESTING 
except:
    script.exit()



output_window.print_md("# 📊 Duct Network Summary:")


# Collect MEPSystem Data
MEPSystem_Obj = get_MEPSystem(start_element)
# print("MEPSystem_Obj: ", MEPSystem_Obj, type(MEPSystem_Obj)) # <- TESTING


# Collect System Name
SystemName = MEPSystem_Obj.Name if MEPSystem_Obj != "N/A" else "N/A"
output_window.print_md("### MEP System Name: {}".format(SystemName))


# Check System Connection Status
try:
    System_ConnectionStatus = MEPSystem_Obj.IsWellConnected
    if System_ConnectionStatus == True:
        output_window.print_md("### System Connection Status: ✅")
    else:
        output_window.print_md("### System Connection Status: ❌")
except:
    forms.alert("Please Select Straight or Flex Duct", title="Invalid Selection", exitscript=True)


# Collect Critical Path Sections
SystemCriticalPath = MEPSystem.GetCriticalPathSectionNumbers(MEPSystem_Obj)
SystemCriticalPath = list(map(str, SystemCriticalPath))
# SystemCriticalPath = MEPSystem_Obj.GetCriticalPathSectionNumbers()
output_window.print_md("### Critical Path Sections: {}".format(SystemCriticalPath))


# Get Critical Path Total Static Pressure Loss
CriticalPath_PressureLoss_Internal = MEPSystem_Obj.PressureLossOfCriticalPath
CriticalPath_PressureLoss = convertUnits(CriticalPath_PressureLoss_Internal, "pressure")
output_window.print_md("### Critical Path Pressure Loss: {:.4f} in-wg".format(CriticalPath_PressureLoss))


# Get Air Flow of System
System_AirFlow_Internal = MEPSystem_Obj.GetFlow()
System_AirFlow = convertUnits(System_AirFlow_Internal, "air flow")
output_window.print_md("### System Air Flow: {:.0f} CFM".format(System_AirFlow))


# Collect all sections in the System
system_sections = get_MEPSystem_sections(start_element)


Elements_BySection = {}
AirFlow_BySection = {}
PressureDrop_BySection = {}
Velocity_BySection = {}
Friction_BySection = {}

# Get all elements in each section
for section in system_sections:
    elements = get_MEPSection_elements(section)
    Elements_BySection[system_sections.index(section) + 1] = elements

    # Get Air Flow per section
    AirFlow_BySection[system_sections.index(section) + 1] = convertUnits(section.Flow, "air flow")

    # Get Velocity per section
    Velocity_BySection[system_sections.index(section) + 1] = convertUnits(section.Velocity, "velocity")

    # Get Friction per section
    Friction_BySection[system_sections.index(section) + 1] = convertUnits(section.Friction, "friction")


# Get coefficient table schema
coeff_schema = find_coefficient_schema()
if coeff_schema is None:
    output_window.print_md(
        "**Could not find ExtensibleStorage schema 'CoefficientFromTable'.**\n"
        "Make sure you have ASHRAE loss coefficient data assigned in the model."
    )


DuctNetworkData = []
""" 
List of Dictionaries containing data for each element in the duct network.
## HEADERS FOR DATA TABLE
| Headers               | Status    | Categories    |
| ----------            | ----      | ----          |
|System Name            |✅         |                                       |
|Category               |✅         |                                       |
|Element ID             |✅         |                                       |
|Section                |✅         |                                       |
|Mark                   |✅         |                                       |
|ASHRAE Table           |✅         | Duct, Fitting, Accessory              |
|Comments               |✅         |                                       |
|Flow (CFM)             |✅         | Duct, Flex Duct, Fitting, Accessory   |
|Size                   |✅         | Duct, Flex Duct, Fitting, Accessory   |
|Length (ft)            |✅         | Duct, Flex Duct                       |
|Velocity (FPM)         |✅         | Duct, Flex Duct                       |
|Friction (in-wg/100ft) |✅         | Duct, Flex Duct                       |
|Pressure Loss          |✅         |                                       |
"""

ColumnOrder = [
    'System Name',
    'Category',
    'Element ID',
    'Section',
    'Mark',
    'ASHRAE Table',
    'Comments',
    'Flow (CFM)',
    'Size',
    'Length (ft)',
    'Velocity (FPM)',
    'Friction (in-wg/100ft)',
    'Pressure Drop (in-wg)',
]


#____________________________________________________________________ RUN: DUCT NETWORK SUMMARY
# Compile data for all elements in the duct network
for section_num, elements in Elements_BySection.items():    # Iterate over dict of Section and Element list
    for elem in elements:                                   # Iterate over Elements per Section
        elem_category = elem.Category.Name

        elem_data = get_element_data(elem, SystemName)      # Must be first item to start dict

        elem_data["Section"] = section_num
        
        # elem_data["Flow (CFM)"] = "{:.0f}".format(AirFlow_BySection[section_num])
        elem_data["Flow (CFM)"] = "{:.0f} ({:.4f})".format(AirFlow_BySection[section_num], system_sections[section_num - 1].Flow)

        # Get Velocity and Friction Values
        if elem_category in ["Ducts", "Flex Ducts"]:
            velocity = Velocity_BySection[section_num]
            elem_data["Velocity (FPM)"] = "{:.2f}".format(velocity)
            
            friction = Friction_BySection[section_num]
            elem_data["Friction (in-wg/100ft)"] = "{:.4f}".format(friction)
        else:
            elem_data["Velocity (FPM)"] = ""
            elem_data["Friction (in-wg/100ft)"] = ""

        pd = get_MEPSection_PressureDrop(system_sections[section_num - 1], elem.Id)
        elem_data["Pressure Drop (in-wg)"] = "{:.4f}".format(pd)

        length = get_MEPSection_SegmentLength(system_sections[section_num - 1], elem.Id)
        elem_data["Length (ft)"] = length

        if elem_category in ["Duct Fittings", "Duct Accessories"]:
            code = get_ashrae_code(elem, coeff_schema) or "<no ASHRAE table set>"
        else:
            code = ""
        elem_data["ASHRAE Table"] = code

        elem_data["Mark"] = get_Mark(elem)

        DuctNetworkData.append(elem_data)


# Prepare data for table display
TableRows = []
for data in DuctNetworkData:
    row = [data.get(col, "") for col in ColumnOrder]
    TableRows.append(row)

TableTitle = "Duct Network Elements: {}".format(len(TableRows))

output_window.print_table(table_data=TableRows, columns=ColumnOrder, title=TableTitle)




#____________________________________________________________________ PATH FINDING - GATHER DUCT PATHS (DIRECTED BY FLOW)






############################################################
############################################################
############################################################
#____________________________________________ PATH FINDING - 1) MAP ELEMENTS TO SECTIONS & COLLECT EQUIPMENT + TERMINALS
# 1) Map each element to its section number and collect equipment + air terminals
AirTerminals = []
Equipment = []
elem_section = {}     # {ElementId: section_number}

for section_num, elements in Elements_BySection.items():    # Iterate over dict of Section and Element list
    for elem in elements:                                   # Iterate over Elements per Section
        elem_section[elem.Id] = section_num
        # print("element: {}, section: {}".format(elem.Id, section_num)) # <- TESTING

        cat_name = elem.Category.Name if elem.Category else ""

        if cat_name == "Air Terminals":
            AirTerminals.append(elem)
        elif cat_name == "Mechanical Equipment":
            Equipment.append(elem)

# Build quick sets of section numbers containing terminals and equipment
terminal_sections = set()
for term in AirTerminals:
    sec = elem_section.get(term.Id, None)
    if sec is not None:
        terminal_sections.add(sec)

equipment_sections = set()
for eq in Equipment:
    sec = elem_section.get(eq.Id, None)
    if sec is not None:
        equipment_sections.add(sec)


# 2) Build an undirected graph of sections based on connector connectivity
#    section_graph: {section_number: set([neighbor_section_number, ...])}

# print("elem_section: ", elem_section)# <- TESTING

"""
################################
### BUILD DICT OF CONNECTORS ###
section_graph = defaultdict(set)

for A_elem_id, A_sec in elem_section.items():
    output_window.print_md("## A_elem_id (section): {} ({})".format(A_elem_id, A_sec)) # <- TESTING
    A_elem = doc.GetElement(A_elem_id)
    if A_elem is None:
        continue

    connectors = get_connectors_from_element(A_elem)
    # print("connectors: ", connectors) # <- TESTING
    if not connectors:
        continue

    for c in connectors:
        try:
            output_window.print_md("### c: {}, {}".format(c.Direction, c.ConnectorType)) # <- TESTING
            # c.AllRefs is a set of ConnectorRefs; each has an Owner element
            for ref in c.AllRefs:
                B_element = ref.Owner
                
                output_window.print_md("- B element: {}".format(B_element.Id.ToString())) # <- TESTING
                if B_element is None:
                    continue

                B_element_id = B_element.Id
                if B_element_id == A_elem_id:
                    continue
                if B_element_id not in elem_section:
                    # connected to something outside the analyzed network
                    continue

                B_sec = elem_section[B_element_id]
                if B_sec is None or B_sec == A_sec:
                    continue

                # Undirected edge between section numbers
                section_graph[A_sec].add(B_sec)
                section_graph[B_sec].add(A_sec)

                # output_window.print_md("### B_sec: {}".format(B_sec)) # <- TESTING
                # output_window.print_md("### A_sec: {}".format(A_sec)) # <- TESTING
                # output_window.print_md("---") # <- TESTING
            print("Section Graph: ", section_graph)
        except:
            # Some connectors may not expose AllRefs properly; ignore and continue
            pass
### BUILD DICT OF CONNECTORS ###
################################
# """


section_graph = defaultdict(set)

dict_Path = {}  # { Element Id : c_Id: dict }
Elements_All = []
# Connectors_All = []

for A_section, elements in Elements_BySection.items():

    for A_elem in elements:
        # output_window.print_md("### A_elem_id (section): {} ({})".format(A_elem.Id, A_section)) # <- TESTING
        # A_elem = doc.GetElement(A_elem_id)
        if A_elem is None:
            continue

        Elements_All.append(A_elem.Id)
        
        dict_Connectors = {}    # { connector Id : section, c_Id, c_Direction, c_ConnectorType, c_Flow, c_Owner, c_AllRefs, connector }

        connectors = get_connectors_from_element(A_elem)
        # print("connectors: ", connectors) # <- TESTING
        if not connectors:
            continue

        for c in connectors:
            section =   A_section
            c_Id =      c.Id
            c_Owner =   c.Owner
            
            # Connectors_All.append(c)
            
            try:
                c_Direction = c.Direction if c.Direction else "N/A"
                c_ConnectorType = c.ConnectorType if c.ConnectorType else "N/A"
                c_Flow = convertUnits(c.Flow, "air flow") if c.Flow else "N/A"
                c_AllRefs = c.AllRefs
            except:
                pass
            
            ref_list = []
            for ref in c_AllRefs:
                if ref:
                    ref_list.append(ref)

            
            dict_Connectors[c.Id] = {
                "section":          section, 
                "c_Id":             c_Id, 
                "c_Direction":      c_Direction, 
                "c_ConnectorType":  c_ConnectorType, 
                "c_Flow":           c_Flow, 
                "c_Owner":          c_Owner, 
                "c_AllRefs":        ref_list,
                "connector":        c
                }
            
            
            
            # print(dict_Connectors[c.Id])
            try:
                # output_window.print_md("Element: {} ({}) | c: {}, {}, {}, {}".format(A_elem.Id, A_section, c.Id, c.Direction, c.ConnectorType, c.Origin)) # <- TESTING
                # c.AllRefs is a set of ConnectorRefs; each has an Owner element
                for ref in c.AllRefs:
                    B_element = ref.Owner
                    
                    # output_window.print_md("- B element: {} (connector: {} {} {} {})".format(B_element.Id.ToString(), ref.Id, ref.Direction, ref.ConnectorType, ref.Origin)) # <- TESTING
                    if B_element is None:
                        continue

                    B_element_id = B_element.Id
                    if B_element_id == A_elem.Id:
                        continue
                    if B_element_id not in elem_section:
                        # connected to something outside the analyzed network
                        continue

                    B_sec = elem_section[B_element_id]
                    if B_sec is None or B_sec == A_section:
                        continue

                    # Undirected edge between section numbers
                    section_graph[A_section].add(B_sec)
                    section_graph[B_sec].add(A_section)

                    # output_window.print_md("### B_sec: {}".format(B_sec)) # <- TESTING
                    # output_window.print_md("### A_sec: {}".format(A_section)) # <- TESTING
                    # output_window.print_md("---") # <- TESTING
                # print("Section Graph: ", section_graph)
            except:
                # Some connectors may not expose AllRefs properly; ignore and continue
                pass
        dict_Path[A_elem.Id.ToString()] = dict_Connectors
        # print(dict_Connectors)

# print(dict_Path)



# output_window.print_md("## TESTING IsConnectedTo on 2457958")
# print("dict_Path: ", dict_Path.keys())
# print("dict_Path (length): ", len(dict_Path.keys()))

elem_2457958 = dict_Path["2457958"] # <- dict of element's connectors
elem_2457958_c1 = elem_2457958[1]   # <- dict of connector 1
elem_2457958_c2 = elem_2457958[2]   # <- dict of connector 2


# print("elem_2457958: {}".format(elem_2457958))
# output_window.print_md("---")
# print("elem_2457958_c1: {}".format(elem_2457958_c1['connector']))
# output_window.print_md("---")
# print("elem_2457958_c2: {}".format(elem_2457958_c2['connector']))


# print("elem_2457958_c1: {}".format(elem_2457958_c1['connector'].IsConnectedTo(elem_2457958_c2['connector'])))


Elements_All = list(set(Elements_All))

print("Elements (length): {}".format(len(Elements_All)))
# print("Elements: {}".format(Elements_All))


Connectors_All = []
for elem_id in Elements_All:
    elem = doc.GetElement(elem_id)
    connectors = get_connectors_from_element(elem)
    for c in connectors:
        try:
            c_Flow = convertUnits(c.Flow, "air flow") if c.Flow else "N/A"
            c_Direction = c.Direction if c.Direction else "N/A"
            c_ConnectorType = c.ConnectorType if c.ConnectorType else "N/A"
        except:
            pass
        
        Connectors_All.append(c)
        # print("connector: {}, Owner: {}, Id: {}, Direction: {}, Flow: {}, ConnectorType: {}".format(c, c.Owner.Id, c.Id, c_Direction, c_Flow, c_ConnectorType))


print("Connectors: {}".format(len(Connectors_All)))


# for element, conn in dict_Path.items():
#     for c_id in conn.keys():
#         connector_dict = conn[c_id]
#         connector_A = connector_dict['connector']

#         for connector_B in Connectors_All:
#             IsConnected = connector_A.IsConnectedTo(connector_B)
#             if IsConnected == True and connector_A.Owner.Id.ToString() != connector_B.Owner.Id.ToString():
#                 output_window.print_md("{} --> {}".format(connector_A.Owner.Id.ToString(), connector_B.Owner.Id.ToString()))



FlowPath = []
source_id = "2633093"
source = Equipment
source_ids = {eq.Id for eq in Equipment}



# print("AirTerminals: {}".format(AirTerminals))
# print("Equipment: {}".format(Equipment))


# def isConnectorConnected(conn_A, conn_All):
#     for conn_B in conn_All:
#         isconnected = conn_A.IsConnectedTo(conn_B)
#         if isconnected == True and conn_A.Owner.Id.ToString() != conn_B.Owner.Id.ToString():
#             outputValue = conn_B
    
#     return outputValue


# for terminal in AirTerminals:
#     FlowPath = [terminal]

#     while FlowPath[-1].Id not in source_ids:
#         current_elem = FlowPath[-1]

#         next_elem = None

#         # TODO: find upstream-connected element via connectors
#         # next_elem = ...

#         if next_elem is None:
#             # dead end, prevent infinite loop
#             break

#         FlowPath.append(next_elem)
            
# output_window.print_md("---")
# print("FlowPath: {}".format(FlowPath))



######################################################
######################################################
######################################################
from Autodesk.Revit.DB import FlowDirectionType

def _as_flow_dir(d):
    """Return FlowDirectionType or None safely."""
    try:
        return d
    except:
        return None


def iter_connected_neighbors(current_elem):
    """
    Yield neighbor relationships as tuples:
        (neighbor_elem, current_connector, neighbor_connector)

    neighbor_connector is a Connector object from current_connector.AllRefs.
    """
    for c in get_connectors_from_element(current_elem):
        if c is None:
            continue
        try:
            refs = c.AllRefs
        except:
            refs = None

        if not refs:
            continue

        for nconn in refs:
            try:
                nowner = nconn.Owner
            except:
                nowner = None

            if nowner is None:
                continue

            # skip self-connection
            if nowner.Id.ToString() == current_elem.Id.ToString():
                continue

            yield (nowner, c, nconn)


def pick_upstream_neighbor(current_elem, visited_ids, allowed_elem_ids_set, mode="terminal_to_equipment_supply"):
    """
    Choose ONE upstream neighbor element based on connector directions.

    Parameters
    ----------
    current_elem : Revit element
    visited_ids : set[int]     (integer ElementId values)
    allowed_elem_ids_set : set[int]  only consider neighbors inside your analyzed network
    mode : str
        "terminal_to_equipment_supply"  (default)
            Walking from Air Terminal upstream to Equipment on a SUPPLY system.
            Prefer: current_conn.Direction == In and neighbor_conn.Direction == Out.

        If you later want return/exhaust walking from terminal to equipment, you likely want:
            Prefer: current_conn.Direction == Out and neighbor_conn.Direction == In
    """
    candidates = []

    for neigh, c_cur, c_neigh in iter_connected_neighbors(current_elem):
        nid = neigh.Id.ToString()

        # must be in the analyzed network
        if nid not in allowed_elem_ids_set:
            continue

        # avoid cycles
        if nid in visited_ids:
            continue

        # direction heuristic
        cur_dir = _as_flow_dir(getattr(c_cur, "Direction", None))
        n_dir   = _as_flow_dir(getattr(c_neigh, "Direction", None))

        score = 0

        if mode == "terminal_to_equipment_supply":
            # Upstream (opposite flow) from terminal:
            # current element typically receives flow -> In
            # upstream neighbor typically sends flow -> Out
            if cur_dir == FlowDirectionType.In:
                score += 10
            if n_dir == FlowDirectionType.Out:
                score += 10

        elif mode == "terminal_to_equipment_return":
            # Return/exhaust case (often opposite):
            if cur_dir == FlowDirectionType.Out:
                score += 10
            if n_dir == FlowDirectionType.In:
                score += 10

        # tie-breaker: favor higher connector flow if available
        try:
            nflow = c_neigh.Flow
            # note: internal units, but relative magnitude is sufficient for scoring
            if nflow:
                score += float(nflow)
        except:
            pass

        candidates.append((score, neigh))

    if not candidates:
        return None

    # pick best-scoring candidate
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]

###################################################
###################################################

# Build equipment id set for termination
equipment_ids = set([eq.Id.ToString() for eq in Equipment])

# Build analyzed-network id set (Elements_All currently contains ElementId objects)
allowed_ids = set([eid.ToString() for eid in Elements_All])

all_flow_paths = []  # optional: collect each terminal's path
dict_all_flow_paths = {}

for terminal in AirTerminals:
    FlowPath = [terminal]
    visited = set([terminal.Id.ToString()])

    max_hops = 500  # safety cap for weird/cyclic networks

    hops = 0
    while FlowPath[-1].Id.ToString() not in equipment_ids:
        hops += 1
        if hops > max_hops:
            # prevent hang
            break

        current_elem = FlowPath[-1]

        next_elem = pick_upstream_neighbor(
            current_elem=current_elem,
            visited_ids=visited,
            allowed_elem_ids_set=allowed_ids,
            mode="terminal_to_equipment_supply"  # change if needed
        )

        if next_elem is None:
            # dead end: no upstream neighbor found
            break

        FlowPath.append(next_elem)
        visited.add(next_elem.Id.ToString())

    all_flow_paths.append(FlowPath)

    # Optional: print a compact result
    try:
        term_mark = get_Mark(terminal) or "<no mark>"
    except:
        term_mark = "<no mark>"
    
    dict_all_flow_paths[term_mark] = FlowPath   # Main Output of Flow Paths

    path_ids = [e.Id.ToString() for e in FlowPath]
    output_window.print_md("- Terminal {} path: {}".format(term_mark, " --> ".join([str(i) for i in path_ids])))

######################################################
######################################################
######################################################

# output_window.print_md("## all_flow_paths:")
# print(all_flow_paths)

# output_window.print_md("## dict_all_flow_paths:")
# print(dict_all_flow_paths)

















# # 3) Determine sources and sinks based on flow direction
# flow_from_equipment = True
# if flow_from_equipment:
#     # Supply: Equipment -> Terminals
#     sources = set(equipment_sections)
#     sinks   = set(terminal_sections)
# else:
#     # Return/Exhaust: Terminals -> Equipment
#     sources = set(terminal_sections)
#     sinks   = set(equipment_sections)

# # Fallbacks if something is missing (e.g. user started from a branch)
# if not sources:
#     # Use endpoints (degree 1) as sources
#     endpoints = [s for s, nbs in section_graph.items() if len(nbs) == 1]
#     sources = set(endpoints)

# if not sinks:
#     # Use endpoints as sinks if none detected
#     endpoints = [s for s, nbs in section_graph.items() if len(nbs) == 1]
#     sinks = set(endpoints)


# def find_directed_flow_paths(section_graph_dict, sources_set, sinks_set):
#     """
#     Generate directed paths from sources to sinks over an undirected section graph.

#     - section_graph_dict: dict[int, set[int]]
#     - sources_set: set[int]
#     - sinks_set: set[int]

#     Returns:
#         List[List[int]]  e.g. [[10, 9, 8], [10, 11, 12, 13], ...]
#         where order follows the logical flow direction.
#     """
#     paths = []
#     paths_set = set()

#     for src in sources_set:
#         if src not in section_graph_dict:
#             continue

#         # Iterative DFS stack: (current_section, path_list)
#         stack = [(src, [src])]

#         while stack:
#             current, path = stack.pop()

#             # If we've reached a sink (and it's not the trivial src-only path), record the path
#             if current in sinks_set and current != src:
#                 path_tuple = tuple(path)
#                 if path_tuple not in paths_set:
#                     paths_set.add(path_tuple)
#                     paths.append(list(path))
#                     # print("paths: ", paths) # <- TESTING
#                 # Don't continue past sink for flow paths
#                 continue

#             for nxt in section_graph_dict.get(current, set()):
#                 if nxt in path:
#                     # avoid cycles
#                     continue
#                 stack.append((nxt, path + [nxt]))

#     return paths


# SectionPaths = find_directed_flow_paths(section_graph, sources, sinks)

# # 4) Output: directed list of section numbers for each flow path
# output_window.print_md("## 🔀 Flow Paths by Section (Directed)")

# print(AirFlow_BySection)


# if not SectionPaths:
#     output_window.print_md(
#         "- No directed section-level flow paths could be determined.\n"
#         "  This can happen if the network is extremely small or disconnected."
#     )
# else:
#     direction_label = "Equipment → Terminals" if flow_from_equipment else "Terminals → Equipment"
#     output_window.print_md("Flow orientation: **{}**".format(direction_label))

#     for i, path in enumerate(SectionPaths, 1):
#         str_path = [s for s in path]
#         airflow_path = []
#         for p in str_path:
#             airflow_path.append(AirFlow_BySection[p])

#         outputPath = []
#         for sec, air in zip(str_path, airflow_path):
#             outputPath.append("{} ({:.0f})".format(sec, air))

        # output_window.print_md("**Path {0}:** `{1}`".format(i, " -> ".join(str(s) for s in outputPath))) # <- MAIN PATH OUTPUT

        # output_window.print_md(
        #     "**Path {0}:** `{1}`".format(i, " -> ".join(str(s) for s in path))
        # )

# 5) Optional: show which paths contain air terminals and equipment

# if SectionPaths:
#     # Build a quick map: section_number -> list of path indices (1-based) that contain it
#     section_to_paths = defaultdict(list)
#     for idx, path in enumerate(SectionPaths):
#         for sec in path:
#             section_to_paths[sec].append(idx + 1)

#     output_window.print_md("### Section Membership by Flow Path")

#     if equipment_sections:
#         output_window.print_md("**Equipment Sections:** {}".format(
#             ", ".join(str(s) for s in sorted(equipment_sections))
#         ))
#     if terminal_sections:
#         output_window.print_md("**Terminal Sections:** {}".format(
#             ", ".join(str(s) for s in sorted(terminal_sections))
#         ))

#     # Example: detail each air terminal with its flow path(s)
#     if AirTerminals:
#         output_window.print_md("#### Air Terminals by Path")
#         for term in AirTerminals:
#             sec_num = elem_section.get(term.Id, None)
#             term_mark = get_Mark(term)
#             path_indices = section_to_paths.get(sec_num, [])

#             if sec_num is None:
#                 info = "- Air Terminal {} (Id {}) is not in any section.\n".format(
#                     term_mark or "<no mark>", term.Id.IntegerValue
#                 )
#             elif not path_indices:
#                 info = "- Air Terminal {} (Section {}) is not on any detected path.\n".format(
#                     term_mark or "<no mark>", sec_num
#                 )
#             else:
#                 info = "- Air Terminal {} (Section {}) → Paths: {}\n".format(
#                     term_mark or "<no mark>",
#                     sec_num,
#                     ", ".join("Path {0}".format(p) for p in path_indices)
#                 )
#             output_window.print_md(info)
