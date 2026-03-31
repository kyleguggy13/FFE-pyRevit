# -*- coding: UTF-8 -*-
# -------------------------------------------
"""
__title__   = "doc-opened"
__doc__     = Version = v0.1
Date    = MM.DD.YYYY
________________________________________________________________
Tested Revit Versions: 
________________________________________________________________
Description:
# This hook runs after a document is opened.
# It logs the sync event to a JSON file.
________________________________________________________________
Last update:
- [MM.DD.YYYY] - v0.1 BETA
________________________________________________________________
"""

# from pyrevit import forms, revit
# from pyrevit.script import output


# # from Autodesk.Revit.DB import ViewDrafting, ViewSchedule

# output_window = output.get_output()


import clr
clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import ViewDrafting, ViewSchedule, View, ViewType
from Autodesk.Revit.DB.Events import DocumentChangedEventArgs
from Autodesk.Revit.UI import TaskDialog

# Keep handler alive
try:
    from pyrevit import script, HOST_APP
    STICKY = script.get_sticky_dict()
    app = HOST_APP.app                 # ✅ Revit Application (no extra .Application)
except Exception:
    STICKY = globals()
    app = __revit__.Application        # ✅ also fine

HANDLER_KEY = "ffe_doc_changed_handler"

def _is_target_view(el):
    if el is None: return False
    if isinstance(el, (ViewDrafting, ViewSchedule)): return True
    if isinstance(el, View) and el.ViewType == ViewType.Legend: return True
    return False

def on_doc_changed(sender, args):
    doc = args.GetDocument()
    txn_names = list(args.GetTransactionNames() or [])
    # Broadened text just in case your locale differs
    if not any(("Insert Views from File" in n) or ("Insert Views" in n) for n in txn_names):
        return

    hits = []
    for elid in args.GetAddedElementIds() or []:
        el = doc.GetElement(elid)
        if _is_target_view(el):
            if isinstance(el, ViewDrafting):
                hits.append(u"📄 Drafting View: {}".format(el.Name))
            elif isinstance(el, ViewSchedule):
                hits.append(u"📊 Schedule: {}".format(el.Name))
            else:
                hits.append(u"🏷 View: {} ({})".format(el.Name, el.ViewType))

    if hits:
        TaskDialog.Show("View Import Detected", "\n".join(hits))

def subscribe_once():
    if STICKY.get(HANDLER_KEY):
        return
    import System
    from System import EventHandler
    handler = EventHandler[DocumentChangedEventArgs](on_doc_changed)
    app.DocumentChanged += handler
    STICKY[HANDLER_KEY] = handler

subscribe_once()
