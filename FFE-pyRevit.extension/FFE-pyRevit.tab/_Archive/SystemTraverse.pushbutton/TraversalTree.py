# -*- coding: utf-8 -*-
"""
Python translation of:
- TreeNode
- TraversalTree

from Revit SDK Samples / TraverseSystem
"""

from Autodesk.Revit.DB import (
    Element,
    ElementId,
    FamilyInstance,
    MEPCurve,
    MEPSystem,
    MechanicalSystem,
    PipingSystem,
    Connector,
    ConnectorType,
    FlowDirectionType
)
from System.Xml import XmlWriter, XmlWriterSettings


class TreeNode(object):
    """
    A TreeNode object represents an element in the system.
    """

    def __init__(self, doc, element_id):
        """
        Constructor

        :param doc: Revit Document
        :param element_id: Autodesk.Revit.DB.ElementId
        """
        self._document = doc
        self._id = element_id
        self._direction = FlowDirectionType.Undefined
        self._parent = None
        self._input_connector = None
        self._child_nodes = []

    # ------------------- Properties -------------------

    @property
    def Id(self):
        """Id of the element."""
        return self._id

    @property
    def Direction(self):
        """Flow direction of the node."""
        return self._direction

    @Direction.setter
    def Direction(self, value):
        self._direction = value

    @property
    def Parent(self):
        """Parent node of the current node."""
        return self._parent

    @Parent.setter
    def Parent(self, value):
        self._parent = value

    @property
    def ChildNodes(self):
        """First-level child nodes of the current node."""
        return self._child_nodes

    @ChildNodes.setter
    def ChildNodes(self, value):
        self._child_nodes = value

    @property
    def InputConnector(self):
        """
        The connector of the previous element
        to which current element is connected.
        """
        return self._input_connector

    @InputConnector.setter
    def InputConnector(self, value):
        self._input_connector = value

    # ------------------- Internal helpers -------------------

    def _get_element_by_id(self, eid):
        """Get Element by its Id."""
        return self._document.GetElement(eid)

    # ------------------- XML Output -------------------

    def DumpIntoXML(self, writer):
        """
        Dump the node into XML.

        :param writer: System.Xml.XmlWriter
        """
        element = self._get_element_by_id(self._id)
        fi = element if isinstance(element, FamilyInstance) else None

        if fi is not None:
            mep_model = fi.MEPModel
            type_name = ""

            if isinstance(mep_model, MechanicalEquipment):
                type_name = "MechanicalEquipment"
                writer.WriteStartElement(type_name)

            elif isinstance(mep_model, MechanicalFitting):
                mf = mep_model  # type: MechanicalFitting
                type_name = "MechanicalFitting"
                writer.WriteStartElement(type_name)
                writer.WriteAttributeString("Category", element.Category.Name)
                writer.WriteAttributeString("PartType", str(mf.PartType))

            else:
                type_name = "FamilyInstance"
                writer.WriteStartElement(type_name)
                writer.WriteAttributeString("Category", element.Category.Name)

            writer.WriteAttributeString("Name", element.Name)
            writer.WriteAttributeString("Id", element.Id.ToString())
            writer.WriteAttributeString("Direction", str(self._direction))
            writer.WriteEndElement()

        else:
            type_name = element.GetType().Name
            writer.WriteStartElement(type_name)
            writer.WriteAttributeString("Name", element.Name)
            writer.WriteAttributeString("Id", element.Id.ToString())
            writer.WriteAttributeString("Direction", str(self._direction))
            writer.WriteEndElement()

        # Write children / paths
        for node in self._child_nodes:
            if len(self._child_nodes) > 1:
                writer.WriteStartElement("Path")

            node.DumpIntoXML(writer)

            if len(self._child_nodes) > 1:
                writer.WriteEndElement()


class TraversalTree(object):
    """
    Data structure of the traversal.
    """

    def __init__(self, active_document, system):
        """
        Constructor

        :param active_document: Revit Document
        :param system: The MEPSystem to traverse
        """
        self._document = active_document
        self._system = system
        self._is_mechanical_system = isinstance(system, MechanicalSystem)
        self._starting_element_node = None

    # ------------------- Public API -------------------

    def Traverse(self):
        """
        Traverse the system.
        """
        # Get the starting element node
        self._starting_element_node = self._get_starting_element_node()
        # Traverse the system recursively
        self._traverse_node(self._starting_element_node)

    def DumpIntoXML(self, file_name):
        """
        Dump the traversal into an XML file.

        :param file_name: Name (path) of the XML file
        """
        settings = XmlWriterSettings()
        settings.Indent = True
        settings.IndentChars = "    "

        writer = XmlWriter.Create(file_name, settings)

        # Root element: MechanicalSystem or PipingSystem
        mep_system_type = "MechanicalSystem" if self._is_mechanical_system else "PipingSystem"
        writer.WriteStartElement(mep_system_type)

        # Basic information of the MEP system
        self._write_basic_info(writer)
        # Paths of the traversal
        self._write_paths(writer)

        # Close root
        writer.WriteEndElement()

        writer.Flush()
        writer.Close()

    # ------------------- Internal helpers -------------------

    def _get_starting_element_node(self):
        """
        Get the starting element node.
        If the system has base equipment then get it;
        otherwise get the owner of the open connector in the system.
        """
        equipment = self._system.BaseEquipment  # FamilyInstance
        if equipment is not None:
            starting_element_node = TreeNode(self._document, equipment.Id)
        else:
            starting_element_node = TreeNode(self._document, self._get_owner_of_open_connector().Id)

        starting_element_node.Parent = None
        starting_element_node.InputConnector = None

        return starting_element_node

    def _get_owner_of_open_connector(self):
        """
        Get the owner of the open connector as the starting element.
        """
        element = None

        # Get an element from the system's terminals
        elements = self._system.Elements
        for ele in elements:
            element = ele
            break

        # Get the open connector recursively
        open_connector = self._get_open_connector(element, None)
        return open_connector.Owner

    def _get_open_connector(self, element, input_connector):
        """
        Get the open connector of the system if the system has no base equipment.

        :param element: An element in the system
        :param input_connector: Connector of the previous element
        :return: Connector or None
        """
        open_connector = None

        # Get connector manager of the element
        if isinstance(element, FamilyInstance):
            fi = element  # type: FamilyInstance
            cm = fi.MEPModel.ConnectorManager
        else:
            mep_curve = element  # type: MEPCurve
            cm = mep_curve.ConnectorManager

        for conn in cm.Connectors:
            # Ignore connector not in this system
            if conn.MEPSystem is None or conn.MEPSystem.Id != self._system.Id:
                continue

            # If connected to the input connector (opposite flow), skip
            if input_connector is not None and conn.IsConnectedTo(input_connector):
                continue

            # If the connector is not connected, it is the open connector
            if not conn.IsConnected:
                open_connector = conn
                break

            # If open connector not found, look from connected elements
            for ref_connector in conn.AllRefs:
                # Ignore non-EndConn and connectors of current element
                if (ref_connector.ConnectorType != ConnectorType.End or
                        ref_connector.Owner.Id == conn.Owner.Id):
                    continue

                # Ignore connectors of the previous element
                if (input_connector is not None and
                        ref_connector.Owner.Id == input_connector.Owner.Id):
                    continue

                open_connector = self._get_open_connector(ref_connector.Owner, conn)
                if open_connector is not None:
                    return open_connector

        return open_connector

    def _traverse_node(self, element_node):
        """
        Traverse the system recursively by analyzing each element.

        :param element_node: TreeNode to be analyzed
        """
        # Find all child nodes and analyze recursively
        self._append_children(element_node)
        for node in element_node.ChildNodes:
            self._traverse_node(node)

    def _append_children(self, element_node):
        """
        Find all child nodes of the specified element node.

        :param element_node: TreeNode
        """
        nodes = element_node.ChildNodes

        # Get connectors
        element = self._get_element_by_id(element_node.Id)
        fi = element if isinstance(element, FamilyInstance) else None

        if fi is not None:
            connectors = fi.MEPModel.ConnectorManager.Connectors
        else:
            mep_curve = element  # type: MEPCurve
            connectors = mep_curve.ConnectorManager.Connectors

        # Find connected connector for each connector
        for connector in connectors:
            mep_system = connector.MEPSystem

            # Ignore connector not in this system
            if mep_system is None or mep_system.Id != self._system.Id:
                continue

            # Direction of the TreeNode
            if element_node.Parent is None:
                if connector.IsConnected:
                    element_node.Direction = connector.Direction
            else:
                # If connector is connected to the input connector, they
                # have opposite flow directions; skip it.
                if connector.IsConnectedTo(element_node.InputConnector):
                    element_node.Direction = connector.Direction
                    continue

            # Connector connected to current connector
            connected_connector = self._get_connected_connector(connector)
            if connected_connector is not None:
                node = TreeNode(self._document, connected_connector.Owner.Id)
                node.InputConnector = connector
                node.Parent = element_node
                nodes.append(node)

        # Sort nodes by element Id (IntegerValue)
        def sort_key(tn):
            return tn.Id.IntegerValue

        nodes.sort(key=sort_key)

    @staticmethod
    def _get_connected_connector(connector):
        """
        Get the connected connector of one connector.

        :param connector: Connector
        :return: Connector or None
        """
        connected_connector = None
        all_refs = connector.AllRefs

        for conn in all_refs:
            # Ignore non-EndConn and connectors of current element
            if conn.ConnectorType != ConnectorType.End or conn.Owner.Id == connector.Owner.Id:
                continue

            connected_connector = conn
            break

        return connected_connector

    def _get_element_by_id(self, eid):
        """Get element by its Id."""
        return self._document.GetElement(eid)

    def _write_basic_info(self, writer):
        """
        Write basic information of the MEP system into the XML file.
        """
        ms = None
        ps = None
        if self._is_mechanical_system:
            ms = self._system  # type: MechanicalSystem
        else:
            ps = self._system  # type: PipingSystem

        writer.WriteStartElement("BasicInformation")

        # Name
        writer.WriteStartElement("Name")
        writer.WriteString(self._system.Name)
        writer.WriteEndElement()

        # Id
        writer.WriteStartElement("Id")
        writer.WriteValue(self._system.Id.ToString())
        writer.WriteEndElement()

        # UniqueId
        writer.WriteStartElement("UniqueId")
        writer.WriteString(self._system.UniqueId)
        writer.WriteEndElement()

        # SystemType
        writer.WriteStartElement("SystemType")
        if self._is_mechanical_system:
            writer.WriteString(ms.SystemType.ToString())
        else:
            writer.WriteString(ps.SystemType.ToString())
        writer.WriteEndElement()

        # Category
        writer.WriteStartElement("Category")
        writer.WriteAttributeString("Id", self._system.Category.Id.ToString())
        writer.WriteAttributeString("Name", self._system.Category.Name)
        writer.WriteEndElement()

        # IsWellConnected
        writer.WriteStartElement("IsWellConnected")
        if self._is_mechanical_system:
            writer.WriteValue(ms.IsWellConnected)
        else:
            writer.WriteValue(ps.IsWellConnected)
        writer.WriteEndElement()

        # HasBaseEquipment
        writer.WriteStartElement("HasBaseEquipment")
        has_base_equipment = (self._system.BaseEquipment is not None)
        writer.WriteValue(has_base_equipment)
        writer.WriteEndElement()

        # TerminalElementsCount
        writer.WriteStartElement("TerminalElementsCount")
        writer.WriteValue(self._system.Elements.Size)
        writer.WriteEndElement()

        # Flow
        writer.WriteStartElement("Flow")
        if self._is_mechanical_system:
            writer.WriteValue(ms.GetFlow())
        else:
            writer.WriteValue(ps.GetFlow())
        writer.WriteEndElement()

        writer.WriteEndElement()  # BasicInformation

    def _write_paths(self, writer):
        """
        Write paths of the traversal into the XML file.
        """
        writer.WriteStartElement("Path")
        self._starting_element_node.DumpIntoXML(writer)
        writer.WriteEndElement()
