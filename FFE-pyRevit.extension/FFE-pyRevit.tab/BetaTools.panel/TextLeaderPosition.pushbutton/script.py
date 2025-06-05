# -*- coding: utf-8 -*-
__title__     = "Text Leader Position"
__version__   = 'Version = 1.0'
__doc__       = """Version = 1.0
Date    = 06.05.2025
# _____________________________________________________________________
# Description:
#
# 
# _____________________________________________________________________
# How-to:
#
# -> Click on the button
# -> 
# _____________________________________________________________________
# Last update:
# - [06.05.2025] - 1.0 RELEASE
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


#____________________________________________________________________ IMPORTS (PYREVIT)

from pyrevit import revit, DB, forms


#____________________________________________________________________ VARIABLES
# app    = __revit__.Application
doc    = __revit__.ActiveUIDocument.Document #type:Document
uidoc  = __revit__.ActiveUIDocument


#____________________________________________________________________ MAIN

### NOTES ###
# LeaderAtachement Enumeration
# https://www.revitapidocs.com/2026/82ed0368-6da3-53a3-8c07-4061efd0be56.htm


if __name__ == '__main__':
    # GET SELECTION
    selected_elements = uidoc.Selection.GetElementIds()


    # CONTAINER
    elements_list = []

    # LOOP THROUGH SELECTED ELEMENTS
    for element_id in selected_elements:
        element = doc.GetElement(element_id)
        elements_list.append(element)
        print("Element ID: ", element.Id, " Element Parameters: ", element.GetParameters('Left Attachment')[0].AsValueString())





# Step 1: Ask user for Top, Middle, or Bottom attachment
options = ["Top", "Middle", "Bottom"]
attachment_choice = forms.SelectFromList.show(options, title="Select Leader Attachment Position", button_name="Apply")


print("Selected Attachment Position: ", attachment_choice, " Type: ", type(attachment_choice))


# If not selected, exit the script
if not attachment_choice:
    forms.alert("No attachment selected. Exiting.", exitscript=True)


# Step 2: Use raw int for AttachmentType enum
attachment_map = {
    "Top": 'Top',
    "Middle": 'Middle',
    "Bottom": 'Bottom'
}
attachment_value = attachment_map[attachment_choice]




#🤖 Automate Your Boring Work Here





# # Set name for the print set
# transaction = Transaction(doc, "Text Leader")
# transaction.Start()
# try:
#     print("Changing text leader position...")


#     transaction.Commit()
# except Exception as e:

#     transaction.RollBack()
#     print("Error ", "Failed to change leader position: ", str(e))


#🤖 Automate Your Boring Work Here


#==================================================
