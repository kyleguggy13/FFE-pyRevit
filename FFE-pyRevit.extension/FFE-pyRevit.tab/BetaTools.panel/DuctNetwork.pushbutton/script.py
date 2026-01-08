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
    :param units:   ["pressure", "air flow"]
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


def get_element_data(element, SystemName):
    """Extract relevant data from a duct element."""
    data = {}
    data['Category'] = element.Category.Name if element.Category else "N/A"
    data['Element ID'] = output_window.linkify(element.Id)
    data['Comments'] = element.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString() if element.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS) else "N/A"
    data['System Name'] = SystemName
    
    if element.Category.Name == "Flex Ducts":
        data['Size'] = element.LookupParameter("Overall Size").AsString()
    elif element.Category.Name != "Mechanical Equipment":
        data['Size'] = element.LookupParameter("Size").AsString()
    else:
        data["Size"] = ""
    # Add more parameters as needed
    return data


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
output_window.print_md("### System Air Flow: {} CFM".format(System_AirFlow))


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
| Headers       | Status    | Categories    |
| ----------    | ----      | ----          |
|Section        |✅     |                                       |
|Category       |✅     |                                       |
|Element ID     |✅     |                                       |
|Type Mark      |❌     |                                       |
|ASHRAE Table   |✅     | Duct, Fitting, Accessory              |
|Comments       |✅     |                                       |
|Size           |✅     | Duct, Flex Duct, Fitting, Accessory   |
|Flow           |✅     | Duct, Flex Duct, Fitting, Accessory   |
|Length         |✅     | Duct, Flex Duct                       |
|Velocity       |✅     | Duct, Flex Duct                       |
|Friction       |✅     | Duct, Flex Duct                       |
|System Name    |✅     |                                       |
|Pressure Loss  |✅     |                                       |
"""



#____________________________________________________________________ RUN
# Compile data for all elements in the duct network
for section_num, elements in Elements_BySection.items():
    for elem in elements:
        elem_data = get_element_data(elem, SystemName)
        elem_data["Section"] = section_num

        elem_data["Flow (CFM)"] = AirFlow_BySection[section_num]

        if elem.Category.Name in ["Ducts", "Flex Ducts"]:

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

        if elem.Category.Name in ["Duct Fittings", "Duct Accessories"]:
            code = get_ashrae_code(elem, coeff_schema) or "<no ASHRAE table set>"
        else:
            code = ""
        elem_data["ASHRAE Table"] = code


        DuctNetworkData.append(elem_data)


# Prepare data for table display
TableRows = []
for data in DuctNetworkData:
    row = data.values()
    TableRows.append(row)

output_window.print_table(table_data=TableRows, columns=DuctNetworkData[0].keys(), title="Duct Network Elements")

