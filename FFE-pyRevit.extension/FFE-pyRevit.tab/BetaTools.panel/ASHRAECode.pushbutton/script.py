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






### CARD TESTING ###
"""
Colorized table demo using pyRevit output cards.

Each "row" is a frame with a title and a set of colorized cards.
Card color is driven by value/limit (handled internally by cards.card_builder).

Drop this into a .pushbutton script bundle.
"""

from pyrevit import script
from pyrevit.output import cards


# --------------------------------------------------------------------------
# Sample "table" data
#   - name: row label (e.g. Level name, System name, etc.)
#   - current: measured value
#   - limit: threshold value used to color the card
# --------------------------------------------------------------------------
# TABLE_ROWS = [
#     {"name": "Zone 1", "current": 10, "limit": 50},
#     {"name": "Zone 2", "current": 25, "limit": 50},
#     {"name": "Zone 3", "current": 40, "limit": 50},
#     {"name": "Zone 4", "current": 55, "limit": 50},
# ]
TABLE_ROWS = [
    {"Element ID":2447626, "Pressure Drop (in-wg)":0.0022,"Flow (CFM)":0.08},
    {"Element ID":2457991, "Pressure Drop (in-wg)":0.0400,"Flow (CFM)":0.08},
    {"Element ID":2457987, "Pressure Drop (in-wg)":0.0020,"Flow (CFM)":0.08},
    {"Element ID":2457991, "Pressure Drop (in-wg)":0.1000,"Flow (CFM)":0.08},
    {"Element ID":3012904, "Pressure Drop (in-wg)":0.0000,"Flow (CFM)":0.08},
    {"Element ID":2457968, "Pressure Drop (in-wg)":0.0001,"Flow (CFM)":0.08}
]

def make_legend_html():
    """Optional legend showing what the colors mean."""
    legend_items = [
        ("#d0d3d4", "0 (or no data)"),
        ("#D0E6A5", "0–50% of limit"),
        ("#FFDD94", "50–100% of limit"),
        ("#FA897B", "> 100% of limit"),
    ]
    parts = ['<div style="margin-bottom:10px;font-family:sans-serif;font-size:0.8rem;">']
    parts.append("<b>Legend</b><br/>")
    for color, label in legend_items:
        parts.append(
            '<span style="display:inline-block;width:14px;height:14px;'
            'border-radius:3px;margin-right:4px;background:{0};"></span>{1}<br/>'
            .format(color, label)
        )
    parts.append("</div>")
    return "".join(parts)


def print_colorized_table(rows):
    """Render a colorized 'table' using frames and cards."""
    out = script.get_output()
    out.set_title("Colorized Table (Output Cards Demo)")

    # Optional: show a legend above the table
    out.print_html(make_legend_html())

    # One frame per row (like a table row)
    for row in rows:
        # name = row["name"]
        # current = row["current"]
        # limit = row["limit"]
        
        name = row["Element ID"]
        current = row["Pressure Drop (in-wg)"]
        limit = row["Flow (CFM)"]

        # Build cards for this "row"
        cards_for_row = []

        # Card for current value (color is based on current vs. limit)
        desc_current = "Current (limit {0})".format(limit)
        current_card = cards.card_builder(limit, current, desc_current)
        cards_for_row.append(current_card)

        # Card for the limit itself (always at limit / limit ratio = 1 → orange)
        desc_limit = "Limit"
        limit_card = cards.card_builder(limit, limit, desc_limit)
        cards_for_row.append(limit_card)

        # Create a frame with the row name as the title and the cards inside
        frame_html = cards.create_frame(name, *cards_for_row)

        # Print the frame to the output window
        out.print_html(frame_html)


if __name__ == "__main__":
    print_colorized_table(TABLE_ROWS)
