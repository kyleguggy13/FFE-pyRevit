# -*- coding: utf-8 -*-
__title__     = "Traverse System"
__version__   = 'Version = 0.1'
__doc__       = """Version = 0.1
Date    = 01.20.2026
______________________________________________________________
Description:
-> 
______________________________________________________________
How-to:
-> Select Straight or Flex duct
______________________________________________________________
Last update:
- [01.20.2026] - v0.1 BETA RELEASE
______________________________________________________________
Author: Kyle Guggenheim"""

"""
# Links:
- Traversing and Exporting all MEP System Graphs
    https://jeremytammik.github.io/tbc/a/1449_traverse_mep_system.html

- MEP System Structure in Hierarchical JSON Graph
    https://jeremytammik.github.io/tbc/a/1450_mep_system_json_graph.html

- Exporting RVT BIM to JSON, WebGL and Forge
    https://jeremytammik.github.io/tbc/a/1451_custom_export_adn_2017.html
"""



#____________________________________________________________________ IMPORTS (SYSTEM)
# from collections import defaultdict
# import time


#____________________________________________________________________ IMPORTS (AUTODESK)
# import sys
# import clr
# clr.AddReference("System")
# from Autodesk.Revit.DB import *
# from Autodesk.Revit.UI import *
# from Autodesk.Revit.DB import FilteredElementCollector, Mechanical, Family, BuiltInParameter, ElementType, UnitTypeId
# from Autodesk.Revit.DB import BuiltInCategory, ElementCategoryFilter, ElementId, FamilyInstance
# from Autodesk.Revit.DB.ExtensibleStorage import Schema


#____________________________________________________________________ IMPORTS (PYREVIT)
# from pyrevit import revit, DB, UI, script
from math import log
from pyrevit.script import output
# from pyrevit import forms

#____________________________________________________________________ VARIABLES
# app         = __revit__.Application
# uidoc       = __revit__.ActiveUIDocument
# doc         = __revit__.ActiveUIDocument.Document   #type: Document
# selection   = uidoc.Selection                       #type: Selection


output_window = output.get_output()
"""Output window for displaying results."""


#____________________________________________________________________ ChatGPT CONVERSION OF C# TO PYTHON

import json
import os

from Autodesk.Revit.DB import (
    Element,
    ElementId,
    FamilyInstance,
    MEPCurve,
    MEPSystem,
    Connector,
    ConnectorType,
    FlowDirectionType,
)

# IMPORTANT: these must come from the Mechanical / Plumbing namespaces
from Autodesk.Revit.DB.Mechanical import MechanicalSystem, MechanicalEquipment, MechanicalFitting
from Autodesk.Revit.DB.Plumbing import PipingSystem

from Autodesk.Revit.UI import TaskDialog
from System.IO import Path, Directory
from System.Diagnostics import Debug

from pyrevit import revit


uidoc = revit.uidoc
doc = revit.doc


def eid_key(eid_or_elem):
    """
    Return a stable key for an element or ElementId.
    Prefers .Id.ToString() if given an element, otherwise uses .ToString().
    """
    try:
        # If it's an element, it has .Id
        return eid_or_elem.Id.ToString()
    except:
        # If it's an ElementId already
        try:
            return eid_or_elem.ToString()
        except:
            return str(eid_or_elem)


# =====================================================================
# TreeNode
# =====================================================================
class TreeNode(object):
    """A TreeNode object represents an element in the system."""

    def __init__(self, doc, element_id):
        self._document = doc
        self._id = element_id
        try:
            self._direction = FlowDirectionType
        except:
            self._direction = None
        # self._direction = FlowDirectionType.Undefined
        self._parent = None
        self._input_connector = None
        self._child_nodes = []

    @property
    def Id(self):
        return self._id

    @property
    def Direction(self):
        return self._direction

    @Direction.setter
    def Direction(self, value):
        self._direction = value

    @property
    def Parent(self):
        return self._parent

    @Parent.setter
    def Parent(self, value):
        self._parent = value

    @property
    def ChildNodes(self):
        return self._child_nodes

    @ChildNodes.setter
    def ChildNodes(self, value):
        self._child_nodes = value

    @property
    def InputConnector(self):
        return self._input_connector

    @InputConnector.setter
    def InputConnector(self, value):
        self._input_connector = value

    def _get_element_by_id(self, eid):
        return self._document.GetElement(eid)

    # ---------- JSON serialization ----------
    def to_dict(self):
        element = self._get_element_by_id(self._id)
        fi = element if isinstance(element, FamilyInstance) else None

        node = {
            "Id": eid_key(element.Id),
            "Name": element.Name,
            "Direction": str(self._direction),
            "RevitType": element.GetType().Name,
        }

        if fi is not None:
            mep_model = fi.MEPModel

            # mechanical types already imported at top: MechanicalEquipment, MechanicalFitting
            if isinstance(mep_model, MechanicalEquipment):
                node["NodeType"] = "MechanicalEquipment"

            elif isinstance(mep_model, MechanicalFitting):
                node["NodeType"] = "MechanicalFitting"
                node["Category"] = element.Category.Name
                node["PartType"] = str(mep_model.PartType)

            else:
                node["NodeType"] = "FamilyInstance"
                node["Category"] = element.Category.Name
        else:
            node["NodeType"] = "NonFamilyElement"

        if self._child_nodes:
            node["Children"] = [child.to_dict() for child in self._child_nodes]

        return node



# =====================================================================
# TraversalTree
# =====================================================================
class TraversalTree(object):
    """Data structure of the traversal."""

    def __init__(self, active_document, system):
        self._document = active_document
        self._system = system
        self._is_mechanical_system = isinstance(system, MEPSystem) and system.GetType().Name == "MechanicalSystem"
        self._starting_element_node = None

    # ---------- main traversal ----------
    def Traverse(self):
        """
        Traverse the system.

        Safely handles cases where we can't determine a valid starting element.
        """
        self._starting_element_node = self._get_starting_element_node()

        # If we fail to determine a starting node, just stop gracefully.
        if self._starting_element_node is None:
            from Autodesk.Revit.UI import TaskDialog
            TaskDialog.Show(
                "Traverse System",
                "Could not determine a valid starting element for this system."
            )
            return

        # Traverse the system recursively
        self._traverse_node(self._starting_element_node)


    # ---------- JSON serialization ----------
    def to_dict(self):
        """
        Build a JSON-serializable object representing the traversal.

        Structure:
        {
            "SystemClass": "MechanicalSystem" | "PipingSystem",
            "BasicInformation": {...},
            "Path": { TreeNode dict ... }
        }
        """
        system_class = "MechanicalSystem" if self._is_mechanical_system else "PipingSystem"
        return {
            "SystemClass": system_class,
            "BasicInformation": self._basic_info_dict(),
            "Path": self._starting_element_node.to_dict() if self._starting_element_node else None,
        }

    def DumpIntoJSON(self, file_name):
        """
        Dump the traversal into a JSON file.
        """
        data = self.to_dict()
        # IronPython can use normal Python file IO here
        with open(file_name, "w") as f:
            json.dump(data, f, indent=4)

    # ---------- internal helpers ----------
    def _get_starting_element_node(self):
        """
        Get the starting element node.

        If the system has base equipment, use that.
        Otherwise, try owner of open connector.
        If that also fails, fall back to the first element in the system, if any.
        """
        equipment = self._system.BaseEquipment
        if equipment is not None:
            starting_element_node = TreeNode(self._document, equipment.Id)
        else:
            owner = self._get_owner_of_open_connector()
            if owner is not None:
                starting_element_node = TreeNode(self._document, owner.Id)
            else:
                # Fallback: just pick the first element in the system, if any
                elements = self._system.Elements
                first_elem = None
                for ele in elements:
                    first_elem = ele
                    break

                if first_elem is None:
                    # No elements at all; nothing we can do
                    return None

                starting_element_node = TreeNode(self._document, first_elem.Id)

        starting_element_node.Parent = None
        starting_element_node.InputConnector = None
        return starting_element_node


    def _get_owner_of_open_connector(self):
        """
        Get the owner of the open connector as the starting element.

        Returns:
            Element or None if we can't find a suitable open connector.
        """
        element = None

        # Get an element from the system's terminals
        elements = self._system.Elements
        for ele in elements:
            element = ele
            break

        if element is None:
            # System has no elements
            return None

        # Get the open connector recursively
        open_connector = self._get_open_connector(element, None)
        if open_connector is None:
            # No open connector found; signal failure
            return None

        return open_connector.Owner


    def _get_open_connector(self, element, input_connector):
        open_connector = None

        if isinstance(element, FamilyInstance):
            fi = element
            cm = fi.MEPModel.ConnectorManager
        else:
            mep_curve = element
            cm = mep_curve.ConnectorManager

        for conn in cm.Connectors:
            if conn.MEPSystem is None or conn.MEPSystem.Id != self._system.Id:
                continue

            if input_connector is not None and conn.IsConnectedTo(input_connector):
                continue

            if not conn.IsConnected:
                open_connector = conn
                break

            for ref_connector in conn.AllRefs:
                if (ref_connector.ConnectorType != ConnectorType.End or
                        ref_connector.Owner.Id == conn.Owner.Id):
                    continue

                if (input_connector is not None and
                        ref_connector.Owner.Id == input_connector.Owner.Id):
                    continue

                open_connector = self._get_open_connector(ref_connector.Owner, conn)
                if open_connector is not None:
                    return open_connector

        return open_connector

    def _traverse_node(self, element_node):
        self._append_children(element_node)
        for node in element_node.ChildNodes:
            self._traverse_node(node)

    def _append_children(self, element_node):
        nodes = element_node.ChildNodes

        element = self._get_element_by_id(element_node.Id)
        fi = element if isinstance(element, FamilyInstance) else None

        if fi is not None:
            connectors = fi.MEPModel.ConnectorManager.Connectors
        else:
            mep_curve = element
            connectors = mep_curve.ConnectorManager.Connectors

        for connector in connectors:
            mep_system = connector.MEPSystem
            if mep_system is None or mep_system.Id != self._system.Id:
                continue

            if element_node.Parent is None:
                if connector.IsConnected:
                    element_node.Direction = connector.Direction
            else:
                if connector.IsConnectedTo(element_node.InputConnector):
                    element_node.Direction = connector.Direction
                    continue

            connected_connector = self._get_connected_connector(connector)
            if connected_connector is not None:
                node = TreeNode(self._document, connected_connector.Owner.Id)
                node.InputConnector = connector
                node.Parent = element_node
                nodes.append(node)

        def sort_key(tn):
            return tn.Id.ToString()

        nodes.sort(key=sort_key)

    @staticmethod
    def _get_connected_connector(connector):
        connected_connector = None
        all_refs = connector.AllRefs

        for conn in all_refs:
            if conn.ConnectorType != ConnectorType.End or conn.Owner.Id == connector.Owner.Id:
                continue
            connected_connector = conn
            break

        return connected_connector

    def _get_element_by_id(self, eid):
        return self._document.GetElement(eid)

    def _basic_info_dict(self):
        """
        Build a dict equivalent of the BasicInformation XML block.
        """
        if self._is_mechanical_system:
            ms = self._system  # MechanicalSystem
            sys_type = ms.SystemType.ToString()
            is_well_connected = ms.IsWellConnected
            flow_val = ms.GetFlow()
        else:
            ps = self._system  # PipingSystem
            sys_type = ps.SystemType.ToString()
            is_well_connected = ps.IsWellConnected
            flow_val = ps.GetFlow()

        return {
            "Name": self._system.Name,
            "Id": self._system.Id.ToString(),
            "UniqueId": self._system.UniqueId,
            "SystemType": sys_type,
            "Category": {
                "Id": self._system.Category.Id.ToString(),
                "Name": self._system.Category.Name,
            },
            "IsWellConnected": bool(is_well_connected),
            "HasBaseEquipment": (self._system.BaseEquipment is not None),
            "TerminalElementsCount": self._system.Elements.Size,
            "Flow": flow_val,
        }


# =====================================================================
# Helper: find a "desirable" system from selection
# =====================================================================
def _get_systems_for_element(elem):
    """Collect MEP systems associated with the element via connectors."""
    systems = set()

    if isinstance(elem, MEPSystem):
        systems.add(elem)
        return list(systems)

    connectors = None

    if isinstance(elem, FamilyInstance):
        mep_model = elem.MEPModel
        if mep_model:
            connectors = mep_model.ConnectorManager.Connectors
    elif isinstance(elem, MEPCurve):
        connectors = elem.ConnectorManager.Connectors

    if connectors:
        for conn in connectors:
            if conn.MEPSystem:
                systems.add(conn.MEPSystem)

    return list(systems)


def _pick_best_system(candidates):
    """Pick the well-connected mechanical/piping system with most elements."""
    best = None
    best_size = -1

    for sys in candidates:
        if not isinstance(sys, (MechanicalSystem, PipingSystem)):
            continue

        if isinstance(sys, MechanicalSystem) and not sys.IsWellConnected:
            continue
        if isinstance(sys, PipingSystem) and not sys.IsWellConnected:
            continue

        size = sys.Elements.Size
        if size > best_size:
            best = sys
            best_size = size

    return best


def _get_selected_system():
    """Get MEPSystem from current selection per README behavior."""
    sel_ids = uidoc.Selection.GetElementIds()
    if sel_ids.Count == 0:
        TaskDialog.Show("Traverse System", "Please select a system or an element in a well-connected system.")
        return None

    first_id = list(sel_ids)[0]
    elem = doc.GetElement(first_id)

    candidates = _get_systems_for_element(elem)
    if not candidates:
        TaskDialog.Show("Traverse System", "Selected element is not part of a mechanical or piping system.")
        return None

    system = _pick_best_system(candidates)
    if system is None:
        TaskDialog.Show("Traverse System", "No well-connected mechanical or piping system found.")
        return None

    return system


def _get_temporary_directory():
    # temp_dir = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName())
    # Directory.CreateDirectory(temp_dir)

    log_dir = os.path.join(os.path.expanduser("~"), "Downloads")

    # log_file = os.path.join(log_dir, username + "_revit_log.json")
    return log_dir


# =====================================================================
# Entry point
# =====================================================================
def main():
    system = _get_selected_system()
    if system is None:
        return

    Debug.WriteLine("Traversing system: {0} ({1})".format(system.Name, system.Id))

    tree = TraversalTree(doc, system)
    tree.Traverse()

    output_folder = _get_temporary_directory()
    filename = Path.Combine(output_folder, "traversal.json")

    tree.DumpIntoJSON(filename)

    msg = "Traversal dumped to:\n{0}".format(filename)
    TaskDialog.Show("Traverse System", msg)


if __name__ == "__main__":
    main()
