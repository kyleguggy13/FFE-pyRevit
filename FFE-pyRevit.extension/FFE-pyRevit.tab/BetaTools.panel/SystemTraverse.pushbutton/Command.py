# -*- coding: utf-8 -*-
"""
Python translation of the C# ExternalCommand that:
- Finds "desirable" MEP systems
- Traverses them with TraversalTree
- Dumps each traversal graph to an XML file in a temp folder
- Shows a TaskDialog summary

Assumes TreeNode and TraversalTree classes are defined in this file
(or imported) per your previous snippet.
"""

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    MEPSystem,
    MechanicalSystem,
    PipingSystem,
)
from Autodesk.Revit.UI import TaskDialog
from System.IO import Path, Directory
from System.Diagnostics import Debug

# pyRevit context
from pyrevit import revit

uidoc = revit.uidoc
doc = revit.doc
app = doc.Application


def is_desirable_system(system):
    """
    Return True to include this system in the exported system graphs.

    Mirrors the intended C# logic:
    (Mechanical OR Piping) AND name != "unassigned" AND element count > 1
    """
    return (
        isinstance(system, (MechanicalSystem, PipingSystem))
        and system.Name != "unassigned"
        and system.Elements.Size > 1
    )


def get_temporary_directory():
    """
    Create and return the path of a random temporary directory.
    Mirrors C# GetTemporaryDirectory().
    """
    temp_directory = Path.Combine(
        Path.GetTempPath(),
        Path.GetRandomFileName()
    )
    Directory.CreateDirectory(temp_directory)
    return temp_directory


def main():
    # Collect all MEP systems
    all_systems_collector = FilteredElementCollector(doc).OfClass(MEPSystem)
    all_systems = list(all_systems_collector)
    n_all_systems = len(all_systems)

    # Filter desirable systems
    desirable_systems = [s for s in all_systems if is_desirable_system(s)]
    n_desirable_systems = len(desirable_systems)

    # Temporary output folder
    output_folder = get_temporary_directory()

    n = 0  # count how many XML files we actually generate

    for system in desirable_systems:
        Debug.WriteLine(system.Name)

        # Root equipment, not used further but kept for parity
        root = system.BaseEquipment  # type: FamilyInstance or None

        # Build traversal tree and traverse the system
        # NOTE: updated to use TraversalTree(Document, MEPSystem)
        tree = TraversalTree(doc, system)
        tree.Traverse()   # no return value in your latest C# pattern

        # Dump traversal graph into an XML file
        filename = str(system.Id.IntegerValue)
        filename = Path.ChangeExtension(
            Path.Combine(output_folder, filename),
            "xml"
        )

        tree.DumpIntoXML(filename)

        # To preview the XML structure you could:
        # System.Diagnostics.Process.Start(filename)

        n += 1

    # Build dialog text
    main_text = (
        "{0} XML files generated in {1} ({2} total systems, {3} desirable):"
        .format(n, output_folder, n_all_systems, n_desirable_systems)
    )

    system_list = [
        "{0}({1})".format(s.Id, s.Name)
        for s in desirable_systems
    ]
    system_list.sort()

    detail_text = ", ".join(system_list)

    dlg = TaskDialog("{0} Systems".format(n))
    dlg.MainInstruction = main_text
    dlg.MainContent = detail_text
    dlg.Show()


# pyRevit / RPS entry point
if __name__ == "__main__":
    main()
