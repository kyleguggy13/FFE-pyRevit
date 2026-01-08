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

#____________________________________________________________________ IMPORTS (AUTODESK)

import sys
from webbrowser import get
import clr
clr.AddReference("System")
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import FilteredElementCollector, Mechanical, Family, BuiltInParameter, ElementType, UnitTypeId
from Autodesk.Revit.DB import BuiltInCategory, ElementCategoryFilter, ElementId, GeometryElement
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



############## LINKIFY TESTING
"""
selection = revit.get_selection()

# Set up the output panel
output_window.set_title("Duct Estimation")
output_window.print_md("## 🛠 Duct & Insulation Estimation Tool")
output_window.print_md("### ⚠️ This tool estimates the total surface area and volume of sheet metal and insulation in the project.")

output_window.print_md("### 📋 Selected Elements:"
                       "\n- {}".format(len(selection.elements)))
# output_window.print_md("### 📋 Selected Element IDs:"
#                        "\n- {}".format(', '.join([str(el.Id) for el in selection.elements])))

for elem in selection.elements:
    elid = elem.Id
    output_window.print_md("### 📋 Element ID: {}".format(output_window.linkify(elid)))
# """
############## LINKIFY TESTING



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



def get_element_data(element, SystemName):
    """Extract relevant data from a duct element."""
    # print("element: ", element, type(element))  # <- TESTING
    data = {}
    data['Category'] = element.Category.Name if element.Category else "N/A"
    # data['Element ID'] = element.Id.ToString()
    data['Element ID'] = output_window.linkify(element.Id)
    data['Comments'] = element.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString() if element.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS) else "N/A"
    data['System Name'] = SystemName
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

"""
# # If the element belongs to a system, this is the easiest way to get all connected elements:
# if hasattr(start_element, 'MEPSystem'):
#     mep_system = start_element.MEPSystem
#     print("mep_system: ", mep_system, type(mep_system)) # <- TESTING
    
#     # If the element is not part of a system, mep_system will be None
#     if mep_system:
#         sections_count = mep_system.SectionsCount                                                   # Get all sections related to MEPSystem
#         system_sections = [mep_system.GetSectionByNumber(i) for i in range(1, sections_count + 1)]
#         output_window.print_md("Found {} sections in the system '{}'.".format(len(system_sections), mep_system.Name))
        
#         # Get all element IDs in section
#         for section in system_sections:
#             output_window.print_md("### Section: {}".format(system_sections.index(section) + 1))    # Print Section Number
#             element_ids_list = section.GetElementIds()
#             for element_id in element_ids_list:
#                 output_window.print_md(" - {}".format(element_id))                                  # Print Element IDs in Section
"""



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
SystemCriticalPath = MEPSystem_Obj.GetCriticalPathSectionNumbers()
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

# Get all elements in each section
for section in system_sections:
    elements = get_MEPSection_elements(section)
    Elements_BySection[system_sections.index(section) + 1] = elements

    # Get Air Flow per section
    AirFlow_BySection[system_sections.index(section) + 1] = convertUnits(section.Flow,"air flow")





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
|ASHRAE Table   |❌     | Duct, Fitting, Accessory              |
|Comments       |✅     |                                       |
|Size           |❌     | Duct, Flex Duct, Fitting, Accessory   |
|Flow           |✅     | Duct, Flex Duct, Fitting, Accessory   |
|Length         |✅     | Duct, Flex Duct                       |
|Velocity       |❌     | Duct, Flex Duct                       |
|Friction       |❌     | Duct, Flex Duct                       |
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
        
        pd = get_MEPSection_PressureDrop(system_sections[section_num - 1], elem.Id)
        elem_data["Pressure Drop (in-wg)"] = "{:.4f}".format(pd)

        length = get_MEPSection_SegmentLength(system_sections[section_num - 1], elem.Id)
        elem_data["Length (ft)"] = length


        DuctNetworkData.append(elem_data)


# Prepare data for table display
TableRows = []
for data in DuctNetworkData:
    row = data.values()
    TableRows.append(row)

output_window.print_table(table_data=TableRows, columns=DuctNetworkData[0].keys(), title="Duct Network Elements")

