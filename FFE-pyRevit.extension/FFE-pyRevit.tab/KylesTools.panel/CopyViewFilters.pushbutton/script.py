# -*- coding: utf-8 -*-
__title__   = "Copy View Filters"
__doc__     = """Version = 1.0
Date    = 03.06.2025
________________________________________________________________
Description:

This is my test button.
# # -*- coding: utf-8 -*-

________________________________________________________________
How-To:

1. [Hold ALT + CLICK] on the button to open its source folder.
You will be able to override this placeholder.
________________________________________________________________
Author: Kyle Guggenheim"""

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝
#==================================================

from Autodesk.Revit.DB import *


# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝
#==================================================
# app    = __revit__.Application
uidoc  = __revit__.ActiveUIDocument
doc    = __revit__.ActiveUIDocument.Document #type:Document


# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝
#==================================================

# 1️⃣ Get Views
from pyrevit import forms

all_views_and_vt = FilteredElementCollector(doc)\
                    .OfCategory(BuiltInCategory.OST_Views)\
                    .WhereElementIsNotElementType()\
                    .ToElements()

dict_views = {}
for v in all_views_and_vt:
    view_or_template = 'View' if not v.IsTemplate else 'ViewTemplate'
    key = '[{a}] ({b}) {c}'.format(a=view_or_template, b=v.ViewType, c=v.Name)
    dict_views[key] = v

# View From
views_from      = sorted(dict_views.keys())
sel_view_from   = forms.SelectFromList.show(views_from, multiselect=False, button_name='Select View with Filters')
view_from       = dict_views[sel_view_from]

# View To
views_to            = sorted(dict_views.keys())
sel_list_views_to   = forms.SelectFromList.show(views_to, multiselect=True, button_name='Select Destination Views')
list_views_to       = [dict_views[view_name] for view_name in sel_list_views_to]



# 2️⃣ Extract Filters
filter_ids  = view_from.GetOrderedFilters()
filters     = [doc.GetElement(f_id) for f_id in filter_ids]


# This was copied from PyRevit Docs: Select From List (Single, Multiple)
# https://pyrevitlabs.notion.site/Effective-Input-ea95e95282a24ba9b154ef88f4f8d056
sel_filters = forms.SelectFromList.show(filters,
                                        multiselect=True,
                                        name_attr='Name', # This needs to be something that exists in the objects in the list
                                        button_name='Select View Filters')



# 3️⃣ Paste Filters
with Transaction(doc, 'Copy View Filters') as t:
    t.Start() # 🔓
    # ALL CHANGES HAVE TO BE HERE

    for view_to in list_views_to:

        if view_to.ViewTemplateId == ElementId(-1): # Only applies view filters if no template.
            for v_filter in sel_filters:
            
                # 1. Copy Filters
                overrides = view_from.GetFilterOverrides(v_filter.Id)
                view_to.SetFilterOverrides(v_filter.Id, overrides)

                # 2. Copy Settings
                visibility = view_from.GetFilterVisibility(v_filter.Id)
                enable = view_from.GetIsFilterEnabled(v_filter.Id)

                view_to.SetFilterVisibility(v_filter.Id, visibility)
                view_to.SetIsFilterEnabled(v_filter.Id, enable)
        else:
            vt = doc.GetElement(view_to.ViewTemplateId)

            print("⚠️View ({v}) has a ViewTemplate({vt}) assigned. It will be skipped".format(v=view_to.Name, vt=vt.Name))


    t.Commit() # 🔒

#🤖 Automate Your Boring Work Here


#==================================================
