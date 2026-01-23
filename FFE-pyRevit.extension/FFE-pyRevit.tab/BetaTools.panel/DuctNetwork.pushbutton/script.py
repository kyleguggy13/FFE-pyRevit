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
import time
import math


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
def eid_key(eid_or_elem):
    """
    Return a stable key for an element or ElementId.
    Prefers .Id.ToString() if given an element, otherwise uses .ToString().
    """
    try:
        # If it's an element, it has .Id
        return eid_or_elem.Id.ToString()
    except:
        # If it's an ElementId already
        try:
            return eid_or_elem.ToString()
        except:
            return str(eid_or_elem)

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

def get_connector_properties(connectors):
    """
    Function to get properties from connectors.
    
    [Direction, ConnectorType, Flow, AllRefs]
    
    :param connectors:  Object to collect from
    :return:            List
    """
    c_properties = []
    for c in connectors:
        try:
            c_Direction = c.Direction if c.Direction else "N/A"
            c_ConnectorType = c.ConnectorType if c.ConnectorType else "N/A"
            c_Flow = convertUnits(c.Flow, "air flow") if c.Flow else "N/A"
            c_AllRefs = c.AllRefs
            c_properties.append(c_Direction, c_ConnectorType, c_Flow, c_AllRefs)
        except:
            pass
    return c_properties


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
print("Timestamp: {}".format(time.strftime("%Y-%m-%d %H:%M:%S")))


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
# print("system_sections: {}".format(system_sections))   # <- TESTING


Elements_BySection = {}
"""Dict of Lists containing elements per section number."""
ElementIds_BySection = {}
"""Dict of Lists containing element ids per section number."""
AirFlow_BySection = {}
"""Dict of Air Flow values per section number."""
PressureDrop_BySection = {}
"""Dict of Pressure Drop values per section number."""
Velocity_BySection = {}
"""Dict of Velocity values per section number."""
Friction_BySection = {}
"""Dict of Friction values per section number."""

# Get all elements in each section
for section in system_sections:
    elements = get_MEPSection_elements(section)
    Elements_BySection[system_sections.index(section) + 1] = elements
    
    ElementIds_BySection[system_sections.index(section) + 1] = [eid_key(elem) for elem in elements] # <- TESTING

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
|Part Type              |✅         |                                       |
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
    'Part Type',
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
"""
Column order for output_window.print_table.
"""


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
            if elem_category == "Duct Fittings":
                parttype = str(elem.MEPModel.PartType)
            else:
                parttype = ""
        else:
            code = ""
            parttype = ""
        elem_data["ASHRAE Table"] = code
        elem_data["Part Type"] = parttype

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
"""
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
"""


# 1) Map each element to its section number and collect equipment + air terminals
AirTerminals = []
Equipment = []

# elementId -> set(SectionNumbers)
elem_sections = defaultdict(set)
"""
defaultdict(set): maps element IDs to sets of section numbers.
- elementId -> set(SectionNumbers)
"""

elem_sections_dict = {}
"""
Dict mapping element IDs to sets of section numbers.
"""

# AirTerminals = []
# Equipment = []

for section_num, elements in Elements_BySection.items():
    elem_list = []
    # output_window.print_md("## Section Number: {}".format(section_num))  # <- TESTING
    for elem in elements:
        eid = eid_key(elem)
        elem_category = elem.Category.Name
        parttype = str(elem.MEPModel.PartType) if hasattr(elem, 'MEPModel') and hasattr(elem.MEPModel, 'PartType') else "N/A"
        # output_window.print_md("### - Element ID: {} | Category: {} | Part Type: {}".format(eid, elem_category, parttype))  # <- TESTING
        connectors = get_connectors_from_element(elem)
        for c in connectors:
            c_Id = c.Id
            try:
                c_Direction = c.Direction if c.Direction else "N/A"
                c_ConnectorType = c.ConnectorType if c.ConnectorType else "N/A"
                c_Flow = convertUnits(c.Flow, "air flow") if c.Flow else "N/A"
                c_PD = "{:.4f}".format(convertUnits(c.PressureDrop, "pressure")) if c.PressureDrop else "N/A"
                c_AllRefs = c.AllRefs
            except:
                pass
            # output_window.print_md("- Connector ID: {} | Direction: {} | ConnectorType: {} | Flow: {}".format(c_Id, c_Direction, c_ConnectorType, c_Flow))  # <- TESTING
            connector_obj = "{}-{}-{}-{}-{}-{}".format(eid, c_Id, c_Direction, c_ConnectorType, c_Flow, c_PD)
            # print("Section{} --> {}-{}-{}-{}-{}-{}".format(section_num, eid, c_Id, c_Direction, c_ConnectorType, c_Flow, c_PD))  # <- TESTING

            for ref in c_AllRefs:
                # print("ref: {}".format(ref))  # <- TESTING
                try:
                    ref_Direction = ref.Direction if ref.Direction else "N/A"
                    ref_ConnectorType = ref.ConnectorType if ref.ConnectorType else "N/A"
                    ref_Flow = convertUnits(ref.Flow, "air flow") if ref.Flow else "N/A"
                    ref_PD = "{:.4f}".format(convertUnits(ref.PressureDrop, "pressure")) if ref.PressureDrop else "N/A"
                    # print("{} --> {}-{}-{}-{}-{}-{}".format(connector_obj, eid_key(ref.Owner), ref.Id, ref_Direction, ref_ConnectorType, ref_Flow, ref_PD))  # <- TESTING
                except:
                    pass
                ref_Id = ref.Id
                ref_Owner = eid_key(ref.Owner)
                # print("{}".format(ref_Direction))  # <- TESTING
        
        
        elem_sections[eid].add(section_num)
        if eid not in elem_sections_dict:
            elem_sections_dict[eid] = [section_num]
        else:
            elem_sections_dict[eid].append(section_num)



        cat_name = elem.Category.Name if elem.Category else ""
        if cat_name == "Air Terminals":
            AirTerminals.append(elem)
        elif cat_name == "Mechanical Equipment":
            Equipment.append(elem)
        # print("element: {}, section: {}".format(elem.Id, section_num)) # <- TESTING


# print("elem_sections: {}".format(elem_sections))                  # <- TESTING

# print("elem_sections_dict: {}".format(elem_sections_dict))        # <- TESTING

# print("Elements_BySection: {}".format(Elements_BySection))        # <- TESTING

# print("ElementIds_BySection: {}".format(ElementIds_BySection))    # <- TESTING



# Build quick sets of section numbers containing terminals and equipment
# terminal_sections = set()
# for term in AirTerminals:
#     sec = elem_section.get(term.Id, None)
#     if sec is not None:
#         terminal_sections.add(sec)

# equipment_sections = set()
# for eq in Equipment:
#     sec = elem_section.get(eq.Id, None)
#     if sec is not None:
#         equipment_sections.add(sec)





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
                    if B_element_id not in elem_sections:
                        # connected to something outside the analyzed network
                        continue

                    B_sec = elem_sections[B_element_id]
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

# elem_2457958 = dict_Path["2457958"] # <- dict of element's connectors
# elem_2457958_c1 = elem_2457958[1]   # <- dict of connector 1
# elem_2457958_c2 = elem_2457958[2]   # <- dict of connector 2


# print("elem_2457958: {}".format(elem_2457958))
# output_window.print_md("---")
# print("elem_2457958_c1: {}".format(elem_2457958_c1['connector']))
# output_window.print_md("---")
# print("elem_2457958_c2: {}".format(elem_2457958_c2['connector']))


# print("elem_2457958_c1: {}".format(elem_2457958_c1['connector'].IsConnectedTo(elem_2457958_c2['connector'])))




Elements_All = list(set(Elements_All))

# output_window.print_md("---")   # <- TESTING
# print("Elements_All: {}".format(Elements_All))  # <- TESTING
# output_window.print_md("---")   # <- TESTING

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




FlowPath = []
source = Equipment
source_ids = {eq.Id for eq in Equipment}




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

def is_close(x, y, epsilon=1e-9):
    """Return True if two values are close in numeric value within a given epsilon."""
    return abs(x - y) <= epsilon

# a = 0.1 + 0.2
# b = 0.3

# print(f"Using custom function: {is_close(a, b)}")
# print(f"Using custom function with different epsilon: {is_close(a, b, epsilon=1e-15)}") # May be False



def find_sum_object(a, b, c):
    a_flow = AirFlow_BySection[a]
    b_flow = AirFlow_BySection[b]
    c_flow = AirFlow_BySection[c]

    print("Analyzing sections {} ({}), {} ({}), {} ({})".format(a_flow, type(a_flow), b_flow, type(b_flow), c_flow, type(c_flow)))  # <- TESTING
    print("b + c = {}".format(int(b_flow)+int(c_flow) == int(a_flow)))
    print("a + c = {}".format(int(a_flow)+int(c_flow) == int(b_flow)))
    print("a + b = {}".format(int(a_flow)+int(b_flow) == int(c_flow)))

    b_c = b_flow + c_flow
    a_c = a_flow + c_flow
    a_b = a_flow + b_flow
    # Use math.isclose() to check for approximate equality
    # if is_close(a_flow, b_c):
    #     print("The numbers are close enough to be considered equal.")
    # else:
    #     print("The numbers are not close enough.")

    if is_close(a_flow, b_c):
        print("  > MATCH FOUND: Section {} equals sum of {} + {}".format(a, b, c))  # <- TESTING
        return a
    if is_close(b_flow, a_c):
        print("  > MATCH FOUND: Section {} equals sum of {} + {}".format(b, a, c))  # <- TESTING
        return b
    if is_close(c_flow, a_b):
        print("  > MATCH FOUND: Section {} equals sum of {} + {}".format(c, a, b))  # <- TESTING
        return c
    # if a_flow == b_flow + c_flow:
    #     print("  > MATCH FOUND: Section {} equals sum of {} + {}".format(a, b, c))  # <- TESTING
    #     return a
    # if b_flow == a_flow + c_flow:
    #     print("  > MATCH FOUND: Section {} equals sum of {} + {}".format(b, a, c))  # <- TESTING
    #     return b
    # if c_flow == a_flow + b_flow:
    #     print("  > MATCH FOUND: Section {} equals sum of {} + {}".format(c, a, b))  # <- TESTING
    #     return c
    return None


# print("AirFlow_BySection: {}".format(AirFlow_BySection))    # <- TESTING
# output_window.print_md("---")   # <- TESTING

# print("Elements_BySection: {}".format(Elements_BySection))    # <- TESTING

def element_path_to_section_path(element_path, elem_sections_map):
    """
    Given elem_path = [Element, Element, ...]
    return a compact section path like [1,2,3,4,5]
    by selecting the shared section between each adjacent pair.

    If an adjacent pair shares multiple sections (rare but possible),
    we prefer continuity with the last chosen section.
    """
    if not element_path or len(element_path) < 2:
        return []

    sec_path = []
    last_sec = None
    output_window.print_md("---")
    output_window.print_md("## === STARTING element_path_to_section_path ===")                                                      # <- TESTING
    print("Total elements in path: {}".format(len(element_path)))                                                                   # <- TESTING

    for i in range(len(element_path) - 1):
        output_window.print_md("### --- ITERATION {}/{} ---".format(i, len(element_path) - 2))                                      # <- TESTING
        a_id = eid_key(element_path[i])
        b_id = eid_key(element_path[i + 1])
        print("Current element A: {} ({})".format(a_id, element_path[i].Category.Name if element_path[i].Category else "N/A"))      # <- TESTING
        print("Next element B: {} ({})".format(b_id, element_path[i + 1].Category.Name if element_path[i + 1].Category else "N/A")) # <- TESTING

        a_secs = elem_sections_map.get(a_id, set())
        b_secs = elem_sections_map.get(b_id, set())
        print("Sections in element A: {}".format(sorted(a_secs)))                                                                   # <- TESTING
        print("Sections in element B: {}".format(sorted(b_secs)))                                                                   # <- TESTING

        shared = set(a_secs) & set(b_secs)
        print("Shared sections: {}".format(sorted(shared)))                                                                         # <- TESTING
        print("Last section (continuity): {}".format(last_sec))                                                                     # <- TESTING

        ### NEED TO USE THESE IF STATEMENTS TO CORRECTLY SELECT THE SECTION
        chosen = None
        if shared:
            print("BRANCH: Shared sections exist")                                                                                  # <- TESTING
            # Prefer to keep continuity if possible
            if last_sec in shared:
                print("  > CONTINUITY: Last section {} is in shared set".format(last_sec))                                          # <- TESTING
                chosen = last_sec
                last_sec_flow = AirFlow_BySection[last_sec]
                print("  > CHOSEN: {} (flow: {})".format(chosen, last_sec_flow))                                                    # <- TESTING
                print("element: {}, if last_sec in shared: {} ({})".format(element_path[i].Id.ToString(), chosen, last_sec_flow))
            

            elif len(shared) >= 2:
                print("  > MULTIPLE SHARED SECTIONS: {} sections shared".format(len(shared)))                                       # <- TESTING
                chosen = sorted(shared)[0]  # stable deterministic pick
                print("  > Initial choice (sorted): {}".format(chosen))                                                             # <- TESTING
                print("element: {}, # of shared: {}, sorted(shared)[0]: {}".format(element_path[i].Id.ToString(), len(shared), chosen))
                
                # Correctly chose branch
                print("  > Attempting to find best fit using 'find_sum_object'")                                                    # <- TESTING
                shared_list = list(shared)
                shared_a = shared_list[0]
                shared_b = shared_list[1]
                print("  > Testing sections {} and {} against last_sec {}".format(shared_a, shared_b, last_sec))                    # <- TESTING
                
                chosen = find_sum_object(shared_a, shared_b, last_sec)
                if chosen:                                                                                                          # <- TESTING
                    print("  > FLOW ANALYSIS RESULT: {} (matches flow sum)".format(chosen))                                         # <- TESTING
                else:                                                                                                               # <- TESTING
                    print("  > FLOW ANALYSIS: No match found, keeping sorted choice")                                               # <- TESTING


                # for sec in shared:
                #     sec_flow = AirFlow_BySection[sec]
                #     # print("  > Analyzing section {} (flow: {})".format(sec, sec_flow))                                          # <- TESTING
                #     connectors = get_connectors_from_element(element_path[i])
                #     for c in connectors:
                #         try:
                #             c_Direction = c.Direction if c.Direction else "N/A"
                #             c_ConnectorType = c.ConnectorType if c.ConnectorType else "N/A"
                #             c_Flow = convertUnits(c.Flow, "air flow") if c.Flow else "N/A"
                #             c_AllRefs = c.AllRefs
                #         except:
                #             pass

                        # print("- connector: {}, {}, {}, {}".format(c_Direction, c_ConnectorType, c_Flow, [c_ref.Owner.Id.ToString() for c_ref in c_AllRefs]))
                    # print("EXTRA PRINT: element: {}, sec (flow): {} ({})".format(element_path[i].Id.ToString(), sec, sec_flow))
                #     if sec_flow == last_sec_flow:
                #         chosen = sec
                #     print("section (flow): {} ({}), if sec_flow == last_sec_flow: {} ({})".format(sec, sec_flow, last_sec, last_sec_flow))

            elif len(shared) == 1 and element_path[i].Category.Name == "Ducts" and len(list(elem_sections[eid_key(element_path[i])])) > 1:
                print("  > SPECIAL CASE: Single shared section for Duct element")                                                   # <- TESTING
                # Check if Duct has multiple sections with increasing flow
                duct_sections = list(elem_sections[eid_key(element_path[i])])
                print("  > Duct sections before sorting: {}".format(duct_sections))                                                 # <- TESTING
                
                duct_airflows = [AirFlow_BySection[sec] for sec in duct_sections]
                print("  > Duct airflows before sorting: {}".format(duct_airflows))                                                 # <- TESTING
                
                zipped_pairs = zip(duct_airflows, duct_sections)
                sorted_pairs = sorted(zipped_pairs)
                duct_airflows, duct_sections = zip(*sorted_pairs)

                for airflow in duct_airflows:
                    if airflow > AirFlow_BySection[last_sec] and airflow != duct_airflows[-1]:
                        duct_index = duct_airflows.index(airflow)
                        chosen = list(duct_sections[duct_index:])
                        print("  > SPECIAL CASE CHOSEN: {} (flow: {})".format(chosen, airflow))                                     # <- TESTING
                        break


                # print("  > Duct sections: {}, flows: {}".format(duct_sections, duct_airflows))                                      # <- TESTING
                # print("element: {}, sections: {}, air flows: {}".format(eid_key(element_path[i]), duct_sections, duct_airflows))
                # connectors = get_connectors_from_element(element_path[i])
                # print("  > Duct has {} connectors".format(len(connectors)))                                                         # <- TESTING
                # # print(" - connectors: {}".format(len(connectors)))
                # for c in connectors:
                #     try:
                #         c_Direction = c.Direction if c.Direction else "N/A"
                #         c_ConnectorType = c.ConnectorType if c.ConnectorType else "N/A"
                #         c_Flow = convertUnits(c.Flow, "air flow") if c.Flow else "N/A"
                #         c_AllRefs = c.AllRefs
                #     except:
                #         pass
                #     c_AllRefs = [c_ref.Owner.Id.ToString() for c_ref in c_AllRefs]
                #     if eid_key(element_path[i - 1]) in c_AllRefs:
                #         print("  > Found connector connected to previous element")                                                  # <- TESTING
                #         print("- last element: {}, connector: {}, {}, {}, {}".format(element_path[i - 1].Id.ToString(), c_Direction, c_ConnectorType, c_Flow, c_AllRefs))
                #         # section_index = duct_airflows.index(c_Flow)
                #         # chosen = duct_sections[section_index]
            else:
                print("  > FLOW ANALYSIS: No match found, keeping sorted choice")                                                   # <- TESTING
                chosen = sorted(shared)[0]  # stable deterministic pick

            """
            else:
                print("  > MULTIPLE SHARED SECTIONS: {} sections shared".format(len(shared)))                                       # <- TESTING
                chosen = sorted(shared)[0]  # stable deterministic pick
                print("  > Initial choice (sorted): {}".format(chosen))                                                             # <- TESTING
                # print("element: {}, # of shared: {}, sorted(shared)[0]: {}".format(element_path[i].Id.ToString(), len(shared), chosen))
                
                # Correctly chose branch
                if len(shared) >= 2:
                    print("  > Attempting to find best fit using flow analysis")                                                    # <- TESTING
                    shared_list = list(shared)
                    shared_a = shared_list[0]
                    shared_b = shared_list[1]
                    print("  > Testing sections {} and {} against last_sec {}".format(shared_a, shared_b, last_sec))                # <- TESTING
                    
                    chosen = find_sum_object(shared_a, shared_b, last_sec)
                    if chosen:                                                                                                      # <- TESTING
                        print("  > FLOW ANALYSIS RESULT: {} (matches flow sum)".format(chosen))                                     # <- TESTING
                    else:                                                                                                           # <- TESTING
                        print("  > FLOW ANALYSIS: No match found, keeping sorted choice")                                           # <- TESTING


                for sec in shared:
                    sec_flow = AirFlow_BySection[sec]
                    print("  > Analyzing section {} (flow: {})".format(sec, sec_flow))                                              # <- TESTING
                    connectors = get_connectors_from_element(element_path[i])
                    for c in connectors:
                        try:
                            c_Direction = c.Direction if c.Direction else "N/A"
                            c_ConnectorType = c.ConnectorType if c.ConnectorType else "N/A"
                            c_Flow = convertUnits(c.Flow, "air flow") if c.Flow else "N/A"
                            c_AllRefs = c.AllRefs
                        except:
                            pass

                        print("- connector: {}, {}, {}, {}".format(c_Direction, c_ConnectorType, c_Flow, [c_ref.Owner.Id.ToString() for c_ref in c_AllRefs]))
                    print("EXTRA PRINT: element: {}, sec (flow): {} ({})".format(element_path[i].Id.ToString(), sec, sec_flow))
                #     if sec_flow == last_sec_flow:
                #         chosen = sec
                #     print("section (flow): {} ({}), if sec_flow == last_sec_flow: {} ({})".format(sec, sec_flow, last_sec, last_sec_flow))
                """

        else:
            print("BRANCH: NO SHARED SECTIONS")                                                                                     # <- TESTING
            # No shared section found. This can happen if:
            # - one element isn't in Elements_BySection
            # - Revit sectioning produced a gap at this adjacency
            # Fallback: pick something deterministic so output still exists.
            union_secs = set(a_secs) | set(b_secs)
            print("  > Union of sections: {}".format(sorted(union_secs)))                                                           # <- TESTING
            if union_secs:
                chosen = sorted(union_secs)[0]
                print("  > FALLBACK: Using first element from union: {}".format(chosen))                                            # <- TESTING
                # print("element: {}, union_secs: {}".format(element_path[i].Id.ToString(), chosen))
            else:                                                                                                                   # <- TESTING
                print("  > WARNING: No sections available at all!")                                                                 # <- TESTING

        print("\nFinal decision for iteration {}:".format(i))                                                                       # <- TESTING
        print("  > chosen section: {}".format(chosen))                                                                              # <- TESTING
        print("  > last_sec: {}".format(last_sec))                                                                                  # <- TESTING

        if chosen is not None and chosen != last_sec:
            if isinstance(chosen, list):
                sec_path.extend(chosen)
                print("  > APPENDED to sec_path (list): {}".format(chosen))
                last_sec = chosen[-1]
                
            else:
                sec_path.append(chosen)
                print("  > APPENDED to sec_path: {}".format(chosen))                                                                    # <- TESTING
                last_sec = chosen
            output_window.print_md("**CHOSEN SECTION: {}**".format(chosen))
        else:                                                                                                                       # <- TESTING
            if chosen is None:                                                                                                      # <- TESTING
                print("  > SKIPPED: chosen is None")                                                                                # <- TESTING
            else:                                                                                                                   # <- TESTING
                print("  > SKIPPED: chosen == last_sec (no change needed)")                                                         # <- TESTING
        output_window.print_md("---")

    print("\n=== COMPLETED element_path_to_section_path ===")                                                                       # <- TESTING
    print("Final section path: {}".format(sec_path))                                                                                # <- TESTING
    return sec_path


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
    visited_ids : set[str]     (string ElementId values)
    allowed_elem_ids_set : set[str]  only consider neighbors inside your analyzed network
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
                # print("Neighbor flow: {} ({})".format(convertUnits(nflow, "air flow"), nflow))   # <- TESTING
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
dict_all_flow_paths_sections = {}

for terminal in AirTerminals:
    FlowPath = [terminal]
    FlowPath_Sections = [elem_sections[eid_key(terminal)]]

    visited = set([terminal.Id.ToString()])

    max_hops = 500  # safety cap for weird/cyclic networks

    hops = 0
    while FlowPath[-1].Id.ToString() not in equipment_ids:
        hops += 1
        if hops > max_hops:
            # prevent hang
            break

        current_elem = FlowPath[-1]
        current_elem_sections = elem_sections_dict.get(eid_key(current_elem))
        current_elem_flow = []
        for sec in current_elem_sections:
            current_elem_flow.append(AirFlow_BySection.get(sec, None))

        # print("Current Element: {} | Sections: {} | Air Flows: {}".format(
        #     eid_key(current_elem),
        #     current_elem_sections,
        #     current_elem_flow
        # ))  # <- TESTING

        next_elem = pick_upstream_neighbor(
            current_elem=current_elem,
            visited_ids=visited,
            allowed_elem_ids_set=allowed_ids,
            mode="terminal_to_equipment_supply"  # change if needed
        )

        if next_elem is None:
            # dead end: no upstream neighbor found
            break

        # print("Next Element: {} | Sections: {} | Air Flows: {}".format(eid_key(next_elem), current_elem_sections, next_elem_flow))  # <- TESTING

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
    # output_window.print_md("- Terminal {} path: {}".format(term_mark, " --> ".join([str(i) for i in path_ids])))


    # Build section path from element path
    sec_path = element_path_to_section_path(FlowPath, elem_sections)
    dict_all_flow_paths_sections[term_mark] = sec_path   # Main Output of Flow Paths
    
    # Print results
    try:
        term_mark = get_Mark(terminal) or "<no mark>"
    except:
        term_mark = "<no mark>"

    path_ids = [eid_key(e) for e in FlowPath]
    output_window.print_md("- Terminal {} element path: {}".format(
        term_mark, " -> ".join(str(i) for i in path_ids)
    ))

    if sec_path:
        output_window.print_md("  - Terminal {} section path: {}".format(
            term_mark, " -> ".join(str(s) for s in sec_path)
        ))
    else:
        output_window.print_md("  - Terminal {} section path: <none determined>".format(term_mark))



######################################################
######################################################
######################################################

output_window.print_md("---")
output_window.print_md("---")
output_window.print_md("## Results: Terminal section paths")
for term_mark, sec_path in dict_all_flow_paths_sections.items():
    if sec_path:
        output_window.print_md("- Terminal {} section path: {}".format(
            term_mark, " -> ".join(str(s) for s in sec_path)
        ))
    else:
        output_window.print_md("- Terminal {} section path: <none determined>".format(term_mark))





