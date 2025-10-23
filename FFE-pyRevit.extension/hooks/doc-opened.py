# doc-opened.py

from pyrevit import forms, revit
from pyrevit.script import output


# from Autodesk.Revit.DB import ViewDrafting, ViewSchedule

output_window = output.get_output()



# def on_element_added(sender, args):
#     doc = args.GetDocument()
#     for id in args.GetAddedElementIds():
#         el = doc.GetElement(id)
#         if isinstance(el, ViewDrafting):
#             output_window.print_md("📄 Drafting View imported:", el.Name)
#         elif isinstance(el, ViewSchedule):
#             output_window.print_md("📊 Schedule imported:", el.Name)

# __revit__.Application.DocumentChanged += on_element_added



# hook: app-startup.py or doc-opened.py
import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitServices")
from Autodesk.Revit.DB import ViewDrafting, ViewSchedule, Element, ElementId
from Autodesk.Revit.ApplicationServices import Application

# event handler
def on_element_added(sender, args):
    doc = args.GetDocument()
    for id in args.GetAddedElementIds():
        el = doc.GetElement(id)
        if isinstance(el, ViewDrafting):
            output_window.print_md("Drafting View imported:", el.Name)
        elif isinstance(el, ViewSchedule):
            output_window.print_md("Schedule imported:", el.Name)

# subscribe to element added
app = __revit__.Application
app.DocumentChanged += on_element_added