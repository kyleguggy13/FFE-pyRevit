# -*- coding: utf-8 -*-
__title__     = "Duct Estimation"
__version__   = 'Version = 0.1'
__doc__       = """Version = 0.1
Date    = 07.09.2025
______________________________________________________________
Description:
-> Estimates total surface area & volume of sheet metal.

-> Estimates total surface area & volume of insulation.
______________________________________________________________
How-to:
-> Click the button
______________________________________________________________
Last update:
- [07.08.2025] - v0.1 BETA RELEASE
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
"""
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


    # ConvertedValue = DB.UnitUtils.ConvertFromInternalUnits(Val, DB.UnitTypeId.InchesOfWater60DegreesFahrenheit)
    return UnitUtils.ConvertFromInternalUnits(value, units)


def get_MEPSection_PressureDrop(section, element):
    """Get pressure drop values for each element per section"""
    # PressureDrop_dict = {}
    try:
        pressuredrop = section.GetPressureDrop(element)
        # print(pressuredrop) # <- TESTING
        
    except:
        pressuredrop = "Invalid Section"

    # PressureDrop_dict[element] = pressuredrop
    return convertUnits(pressuredrop, "pressure")




def get_element_data(element, SystemName):
    """Extract relevant data from a duct element."""
    # print("element: ", element, type(element))  # <- TESTING
    data = {}
    data['Category'] = element.Category.Name if element.Category else "N/A"
    data['Element ID'] = element.Id.ToString()
    data['Comments'] = element.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).AsString() if element.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS) else "N/A"
    data['System Name'] = SystemName
    # Add more parameters as needed
    return data


#____________________________________________________________________ MAIN
uidoc = revit.uidoc
doc = revit.doc

# Select one duct element first
try:
    ref = uidoc.Selection.PickObject(UI.Selection.ObjectType.Element, "Select a duct or fitting.")
    start_element = doc.GetElement(ref)
    print("start element: ", start_element, type(start_element)) # <- TESTING 
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
print("MEPSystem_Obj: ", MEPSystem_Obj, type(MEPSystem_Obj)) # <- TESTING


# Collect System Name
SystemName = MEPSystem_Obj.Name if MEPSystem_Obj != "N/A" else "N/A"
output_window.print_md("### MEP System Name: {}".format(SystemName))


# Check System Connection Status
System_ConnectionStatus = MEPSystem_Obj.IsWellConnected
if System_ConnectionStatus == True:
    output_window.print_md("### System Connection Status: ✅")
else:
    output_window.print_md("### System Connection Status: ❌")


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

    # output_window.print_md("Section {}: {} CFM".format(section, convertUnits(section.Flow,"air flow")))




# print("section 1: ", system_sections[1])

# print("section 14: ", system_sections[14].GetPressureDrop(ElementId.Parse("2457552")))
# print("section 14[0]: ", system_sections[14].GetPressureDrop(Elements_BySection[14][0].Id))
# print("section 14[1]: ", system_sections[14].GetPressureDrop(Elements_BySection[14][1].Id))
# print("section 14[2]: ", system_sections[14].GetPressureDrop(Elements_BySection[14][2].Id))
# print("section 14, Element 0: ", Elements_BySection[14][0].Id)
# print("section 14, Element 1: ", Elements_BySection[14][1].Id)
# print("section 14, Element 2: ", Elements_BySection[14][2].Id)


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
|Length         |❌     | Duct, Flex Duct                       |
|Velocity       |❌     | Duct, Flex Duct                       |
|Friction       |❌     | Duct, Flex Duct                       |
|System Name    |✅     |                                       |
|Pressure Loss  |✅     |                                       |
"""




# Compile data for all elements in the duct network

for section_num, elements in Elements_BySection.items():
    for elem in elements:
        elem_data = get_element_data(elem, SystemName)
        elem_data["Section"] = section_num

        elem_data["Flow (CFM)"] = AirFlow_BySection[section_num]
        
        pd = get_MEPSection_PressureDrop(system_sections[section_num - 1], elem.Id)
        elem_data["Pressure Drop (in-wg)"] = "{:.4f}".format(pd)

        DuctNetworkData.append(elem_data)


# Prepare data for table display
TableRows = []
for data in DuctNetworkData:
    row = data.values()
    TableRows.append(row)

output_window.print_table(table_data=TableRows, columns=DuctNetworkData[0].keys(), title="Duct Network Elements")





### Attempt to get GeometryObject from selected reference
"""
if ref:
    # 1. Get the host Element (the wall, floor, etc. that the geometry belongs to)
    element = doc.GetElement(ref.ElementId)

    # 2. Use the GetGeometryObjectFromReference method to get the specific GeometryObject
    # This method takes the reference and returns the specific geometry part (Solid, Face, Edge, etc.)
    geometry_object = element.GetGeometryObjectFromReference(ref)
    
    if geometry_object:
        output_window.print_md("Type: {}".format(type(geometry_object).__name__))
        output_window.print_md("Internal Id: {}".format(geometry_object.Id))
        output_window.print_md("Geometry Object: {}".format(geometry_object))
        # forms.alert(
        #     "Successfully retrieved GeometryObject:\nType: {}\nInternal Id: {}"
        #     .format(type(geometry_object).__name__, geometry_object.Id),
        #     "Geometry Object Found"
        # )
        # You can now work with the geometry_object (e.g., get area if it's a face)
        if isinstance(geometry_object, DB.Solid):
            output_window.print_md("Type: Solid")
        else:
            output_window.print_md("The geometry object is not a Solid.")
            
    else:
        output_window.print_md("Could not retrieve the specific geometry object.")
else:
    output_window.print_md("No reference was selected.")
"""




### Revit Schedule Calculations (Reference)
# Density               = 0.294
# Gauge_Sq              = if(not((Width + Height) > 2'  6"), 24, if(not((Width + Height) > 4'  6"), 22, if(not((Width + Height) > 7'), 20, if((Width + Height) > 7', 18, 0))))
# Gauge_Thickness_Sq    = if(Gauge_Sq = 24, 0.028, if(Gauge_Sq = 22, 0.034, if(Gauge_Sq = 20, 0.04, if(Gauge_Sq = 18, 0.052, 0))))
# Weight_Sq             = Density * (2 * (Width + Height) / 0'  1") * Gauge_Thickness_Sq * (Length / 0'  1")
# SqFt_Sq               = (2 * (Width + Height) / 1') * (Length / 1')
# Gauge_Rnd             = if(not(Diameter > 2'  4"), 24, if(not(Diameter > 3'  2"), 22, if(not(Diameter > 4'), 20, if(Diameter > 4', 18, 0))))
# Gauge_Thickness_Rnd   = if(Gauge_Rnd = 24, 0.028, if(Gauge_Rnd = 22, 0.034, if(Gauge_Rnd = 20, 0.04, if(Gauge_Rnd = 18, 0.052, 0))))
# Weight_Rnd            = Density * (pi() * Diameter / 0'  1") * Gauge_Thickness_Rnd * (Length / 0'  1")
# SqFt_Rnd              = (pi() * Diameter / 1') * (Length / 1')

#_____________________________________________________________________ 🏃‍➡️ RUN 
"""
# Collect all duct elements
duct_collector = FilteredElementCollector(doc)\
                    .OfCategory(BuiltInCategory.OST_DuctCurves)\
                    .WhereElementIsNotElementType()
MEPCurve_collector = FilteredElementCollector(doc).OfClass(MEPCurve).ToElements()

fitting_collector = FilteredElementCollector(doc)\
                    .OfCategory(BuiltInCategory.OST_DuctFitting)\
                    .WhereElementIsNotElementType()

# Group ducts by system
DuctbySystem = {}
for duct in duct_collector:
    system_name = duct.MEPSystem.Name if duct.MEPSystem else "No System"
    if system_name not in DuctbySystem:
        DuctbySystem[system_name] = []
    DuctbySystem[system_name].append(duct)

# Group MEP Curves by category
MEPCurve_Categories = {}
for curve in MEPCurve_collector:
    c_category = curve.Category.Name
    if c_category not in MEPCurve_Categories:
        MEPCurve_Categories[c_category] = []
    MEPCurve_Categories[c_category].append(curve)



# options = Options()
# options.DetailLevel = ViewDetailLevel.Fine

# Group fittings by Family Name
fitting_geometry = []
Fitting_FamilyNames = {}
for fitting in fitting_collector:
    family_name = fitting.Symbol.FamilyName
    fitting_val = fitting.get_Geometry(DB.Options())
    fitting_geometry.append(fitting_val)
    if family_name not in Fitting_FamilyNames:
        Fitting_FamilyNames[family_name] = []
    Fitting_FamilyNames[family_name].append(fitting)




# print(fitting_geometry[:20])

output_window.print_md("### 📊 MEPCurves Grouped by Category:")
for cat, cur in MEPCurve_Categories.items():
    output_window.print_md("### 📊 {}: {}".format(cat, len(cur)))

output_window.print_md("---")

output_window.print_md("### 📊 Duct Fittings Grouped by Family Name:")
for fam, fits in Fitting_FamilyNames.items():
    output_window.print_md("### 📊 {}: {}".format(fam, len(fits)))


# output_window.print_md("### 📊 Ducts Grouped by System:")
# for system, ducts in DuctbySystem.items():
#     if system == "SA_AHU-7_7-VAV-5":  # Example: filter for a specific system
#         output_window.print_md("#### System: {} | Duct Count: {}".format(system, len(ducts)))

# output_window.print_md("### 📊 Fittings Grouped by System:")
# for system, fittings in FittingbySystem.items():
#     if system == "SA_AHU-7_7-VAV-5":  # Example: filter for a specific system
#         output_window.print_md("#### System: {} | Fitting Count: {}".format(system, len(fittings)))

#_____________________________________________________________________ 📊 TABLE
table_columns = ["Id", "Type Name", "Geometry", "System Name", "Length (ft)", "Level"]


table_rows=[
    [
        duct.Id.ToString(),     # ID to string
        duct.Name,              # Duct Name
        duct.Geometry,
        duct.MEPSystem.Name if duct.MEPSystem else "N/A",       # Duct's System Name
        "{:.2f}".format(duct.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble()) if duct.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH) else "N/A",    # Duct Length
        doc.GetElement(duct.LevelId).Name if duct.LevelId != ElementId.InvalidElementId else "N/A"  # Duct Reference Level
    ]
    for duct in duct_collector  
]



# Print Table
# output_window.print_md("### Duct Elements Found")

# output_window.print_table(table_data=table_rows, columns=table_columns, title="Ducts in Model")
"""