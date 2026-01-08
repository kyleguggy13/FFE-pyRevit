# -*- coding: utf-8 -*-
# Gets ASHRAE fitting code (ASHRAE Table) from duct fittings via ExtensibleStorage

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    FamilyInstance,
    ElementId
)
from Autodesk.Revit.DB.ExtensibleStorage import Schema
from pyrevit import revit, script
from System import String

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()


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


# ----------------------------------------------------------------------------- 
# Collect fittings (selection first; fallback to all duct fittings)
# -----------------------------------------------------------------------------
selection_ids = [el.Id for el in revit.get_selection().elements]

if selection_ids:
    elems = [doc.GetElement(eid) for eid in selection_ids]
    fittings = [
        e for e in elems
        if isinstance(e, FamilyInstance)
        and e.Category
    ]
else:
    fittings = (
        FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_DuctFitting)
        .WhereElementIsNotElementType()
        .ToElements()
    )

if not fittings:
    output.print_md("**No duct fittings found (in selection or model).**")
else:
    coeff_schema = find_coefficient_schema()
    if coeff_schema is None:
        output.print_md(
            "**Could not find ExtensibleStorage schema 'CoefficientFromTable'.**\n"
            "Make sure you have ASHRAE loss coefficient data assigned in the model."
        )
    else:
        output.print_md("### Duct Fittings – ASHRAE Table Codes")

        for f in fittings:
            code = get_ashrae_code(f, coeff_schema) or "<no ASHRAE table set>"

            fam_name = (
                f.Symbol.Family.Name
                if f.Symbol and f.Symbol.Family else "<no family>"
            )
            type_name = f.Symbol if f.Symbol else "<no type>"

            output.print_md(
                "ID: {0:7} | Family: {1} | Type: {2} | ASHRAE Table: {3}".format(
                    f.Id, fam_name, type_name, code
                )
            )
