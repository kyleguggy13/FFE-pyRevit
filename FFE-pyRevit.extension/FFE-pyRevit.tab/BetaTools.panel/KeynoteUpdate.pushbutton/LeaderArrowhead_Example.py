# encoding: utf-8
from pyrevit import revit, DB, forms, script
from collections import OrderedDict

script.get_output().close_others()

doc = revit.doc


def get_name(element):
    """Get the name of the element."""
    return DB.Element.Name.__get__(element)


generic_annotation_collector = [
    _
    for _ in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_GenericAnnotation).WhereElementIsElementType().ToElements()
    if _.get_Parameter(DB.BuiltInParameter.LEADER_ARROWHEAD) is not None
]

generic_annotation_sorted_dict = OrderedDict(
    sorted(
        {
            get_name(generic_annotation): generic_annotation
            for generic_annotation in generic_annotation_collector
        }.items()
    )
)

generic_annotation_selection = forms.SelectFromList.show(
    generic_annotation_sorted_dict.keys(),
    title="Select Generic Annotation Type",
    multiselect=True,
    width=400,
    height=300,
    exitscript=True,
)

if not generic_annotation_collector:
    forms.alert("No generic annotation types found in the document.", exitscript=True)
generic_annotation_types = []

for generic_annotation in generic_annotation_collector:
    if str(get_name(generic_annotation)) in generic_annotation_selection:
        generic_annotation_types.append(generic_annotation)

arrow_heads_types = [
    _
    for _ in DB.FilteredElementCollector(doc).OfClass(DB.ElementType).WhereElementIsElementType().ToElements()
    if _.FamilyName == "Arrowhead"
]

available_arrowhead_types = OrderedDict(
    sorted(
        {
            get_name(arrow_head): arrow_head
            for arrow_head in arrow_heads_types
            if get_name(arrow_head) != "Arrowhead"
        }.items()
    )
)

picked_arrow_head = forms.SelectFromList.show(
    available_arrowhead_types.keys(),
    title="Select Arrowhead Type",
    multiselect=False,
    width=400,
    height=300,
    exitscript=True,
)

if not picked_arrow_head:
    forms.alert("No arrowhead type selected.", exitscript=True)

chosen_arrowhead = available_arrowhead_types.get(picked_arrow_head)

with revit.Transaction("Set Arrowhead Type for Generic Annotations"):
    for generic_annotation_type in generic_annotation_types:
        param = generic_annotation_type.get_Parameter(DB.BuiltInParameter.LEADER_ARROWHEAD).Set(chosen_arrowhead.Id)