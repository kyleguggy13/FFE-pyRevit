# -*- coding: utf-8 -*-
__title__     = "Keynote Update"
__version__   = 'Version = 1.0'
__doc__       = """Version = 1.0
Date    = 06.09.2025
# _____________________________________________________________________
# Description:
#
# 
# _____________________________________________________________________
# How-to:
#
# -> Click the button
# -> 
# _____________________________________________________________________
# Last update:
# - [06.09.2025] - 1.0 RELEASE
# _____________________________________________________________________
Author: Kyle Guggenheim"""


#____________________________________________________________________ IMPORTS (AUTODESK)

from logging import Filter
from requests import get
import clr
clr.AddReference("System")
from System.Collections.Generic import List
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import Selection


#____________________________________________________________________ IMPORTS (PYREVIT)

from pyrevit import revit, DB, forms


#____________________________________________________________________ VARIABLES
app         = __revit__.Application
uidoc       = __revit__.ActiveUIDocument
doc         = __revit__.ActiveUIDocument.Document   #type: Document
selection   = uidoc.Selection                       #type: Selection


#____________________________________________________________________ MAIN

### NOTES ###
# LeaderAtachement Enumeration
# https://www.revitapidocs.com/2026/82ed0368-6da3-53a3-8c07-4061efd0be56.htm



# GET SELECTION
selected_elements = selection.GetElementIds()

# CONTAINER
elements_list = []
param_Left_List = []
param_Right_List = []

# LOOP THROUGH SELECTED ELEMENTS
for element_id in selected_elements:
    element = doc.GetElement(element_id)
    elements_list.append(element)
    param_Left_List.append(element.get_Parameter(BuiltInParameter.LEADER_LEFT_ATTACHMENT))
    param_Right_List.append(element.get_Parameter(BuiltInParameter.LEADER_RIGHT_ATTACHMENT))

    # print("Element ID: ", element.Id, " Element Parameters: ", element.GetParameters('Left Attachment')[0].AsValueString())
    print("get_Parameter: ", element.get_Parameter(BuiltInParameter.LEADER_LEFT_ATTACHMENT).AsValueString())




# Step 1: Ask user for Top, Middle, or Bottom attachment
options = ["Top", "Middle", "Bottom"]
attachment_choice_left = forms.SelectFromList.show(options, title="Select Left Leader Attachment Position", button_name="Apply")
attachment_choice_right = forms.SelectFromList.show(options, title="Select Right Leader Attachment Position", button_name="Apply")

print("Selected Attachment Position: ", attachment_choice_left)
print("Selected Attachment Position: ", attachment_choice_right)


# If not selected, exit the script
if not attachment_choice_left or not attachment_choice_right:
    forms.alert("No attachment selected. Exiting.", exitscript=True)



# Step 2: Use raw int for AttachmentType enum
attachment_map = {
    "Top": 0,
    "Middle": 1,
    "Bottom": 2
}


attachment_value_left = attachment_map[attachment_choice_left]

attachment_value_right = attachment_map[attachment_choice_right]


#____________________________________________________________________ 🤖 Transaction

# Set Attachment Parameters
transaction = Transaction(doc, "Text Leader")
transaction.Start()
try:
    print("Changing text leader position...")
    for param_left, param_right in zip(param_Left_List, param_Right_List):
        ## Set the Left Attachment parameter to the selected attachment position
        param_left.Set(attachment_value_left)
        param_right.Set(attachment_value_right)

        ## Optionally, you can also set the Right Attachment if needed

        ## If you want to set a specific attachment type, you can do so here
        ## For example, if you want to set it to TopLine:
        ##
        print("Left Leader position changed to: ", attachment_choice_left)
        print("Right Leader position changed to: ", attachment_choice_right)

    transaction.Commit()
except Exception as e:

    transaction.RollBack()
    print("Error ", "Failed to change leader position: ", str(e))



#==================================================
