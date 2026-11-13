"""Verify the standalone tagging command without loading Revit.

Key behaviors: validate/cancel settings, preserve classification and scope,
isolate per-tag failures, and publish IDs only after a successful group commit.
Revit API notes: mock the 2025/2026 transaction and ElementId.Value boundaries.
Design decisions: execute the real functions; substitute only host APIs and UI.
"""
import ast
import contextlib
import io
import math
import os
from pathlib import Path
import traceback
import types
import unittest
from unittest.mock import Mock
import xml.etree.ElementTree as ET


BUNDLE = (Path(__file__).resolve().parents[1] / "FFE-pyRevit.extension"
          / "FFE-pyRevit.tab/BetaTools.panel/TagReceptacles.pushbutton")
SCRIPT = BUNDLE / "script.py"
SOURCE = SCRIPT.read_text(encoding="utf-8")
NS = types.SimpleNamespace


class FakeList(list):
    """Supply the .NET generic List surface used by selection."""

    @classmethod
    def __class_getitem__(cls, item):
        return cls

    def Add(self, item):
        self.append(item)

    @property
    def Count(self):
        return len(self)


class Vector:
    """Minimal immutable XYZ operations for actual placement calculations."""

    def __init__(self, x, y, z=0):
        self.X, self.Y, self.Z = x, y, z

    def Add(self, other):
        return Vector(self.X + other.X, self.Y + other.Y, self.Z + other.Z)

    def Subtract(self, other):
        return Vector(self.X - other.X, self.Y - other.Y, self.Z - other.Z)

    def Multiply(self, factor):
        return Vector(self.X * factor, self.Y * factor, self.Z * factor)

    def DotProduct(self, other):
        return self.X * other.X + self.Y * other.Y + self.Z * other.Z


class FakeLocation:
    def __init__(self, point):
        self.Point = point


class FakeSheet:
    ViewType = "Sheet"

    def __init__(self, viewportIds):
        self.viewportIds = viewportIds

    def GetAllViewports(self):
        return self.viewportIds


class FakeWindow:
    """Expose controls and track whether settings close the dialog."""

    def __init__(self, path):
        self.xamlPath = path
        self.closed = False
        for name in ("landingText", "offsetText", "errorText"):
            setattr(self, name, NS(Text=""))
        self.runButton = NS(IsEnabled=True)
        self.writeMode = NS(IsChecked=False)

    def Close(self):
        self.closed = True


class TagReceptaclesTests(unittest.TestCase):
    def setUp(self):
        self.events = []
        self.scopes = []
        self.failKind = None
        self.failOperation = None
        self.failResult = "RolledBack"
        self.raiseKind = None
        self.raiseOperation = None
        self.failElements = set()
        self.failLeaderElements = set()
        self.elementId = 100
        self.view = self.MakeView(10)
        self.doc = NS(ActiveView=self.view, IsReadOnly=False, IsModifiable=False,
                      IsFamilyDocument=False, Regenerate=Mock(), tags=[], symbols=[],
                      fixtures={}, equipment={}, existing={}, elements={})
        self.doc.GetElement = lambda key: self.doc.elements[key.Value]
        self.uidoc = NS(Document=self.doc, Selection=NS(SetElementIds=Mock()))
        self.status = NS(Started="Started", Committed="Committed", RolledBack="RolledBack")
        outer = self

        class Scope:
            def __init__(self, kind):
                self.kind, self.status = kind, "Uninitialized"
                self.disposed = False
                outer.scopes.append(self)

            def GetStatus(self):
                return self.status

            def Start(self):
                self.before = list(outer.doc.tags)
                return self.Finish("Start", "Started")

            def Finish(self, operation, success):
                outer.events.append((self.kind, operation))
                if outer.raiseKind == self.kind and outer.raiseOperation == operation:
                    raise RuntimeError("Simulated {} {} failure".format(self.kind, operation))
                self.status = (outer.failResult if outer.failKind == self.kind
                               and outer.failOperation == operation else success)
                if self.status == "RolledBack":
                    outer.doc.tags[:] = self.before
                return self.status

            def Commit(self):
                return self.Finish("Commit", "Committed")

            def Assimilate(self):
                return self.Finish("Assimilate", "Committed")

            def RollBack(self):
                outer.doc.tags[:] = self.before
                return self.Finish("RollBack", "RolledBack")

            def Dispose(self):
                self.disposed = True

            def GetFailureHandlingOptions(self):
                return NS(SetForcedModalHandling=lambda value: None)

            def SetFailureHandlingOptions(self, options):
                pass

        class Collector:
            def __init__(self, doc, viewId=None):
                self.doc = doc
                self.viewId = viewId.Value if viewId else None
                self.category = None

            def OfClass(self, cls):
                self.key = cls.__name__
                return self

            def OfCategory(self, category):
                self.category = category
                self.key = category
                return self

            def WhereElementIsNotElementType(self):
                return self

            def ToElements(self):
                if self.key == "FamilySymbol" and self.category is not None:
                    return [symbol for symbol in self.doc.symbols
                            if symbol.Category == self.category]
                return {"FamilySymbol": self.doc.symbols,
                        "IndependentTag": self.doc.existing.get(self.viewId, []),
                        "fixtures": self.doc.fixtures.get(self.viewId, []),
                        "equipment": self.doc.equipment.get(self.viewId, [])}.get(self.key, [])

        def CreateTag(doc, symbolId, viewId, element, leader, orientation, head):
            if element.Id.Value in outer.failElements:
                raise RuntimeError("Element cannot be tagged")
            outer.elementId += 1
            tag = Mock(Id=NS(Value=outer.elementId), TagHeadPosition=head)
            tag.element, tag.viewId, tag.symbolId = element, viewId, symbolId
            tag.CanLeaderEndConditionBeAssigned.return_value = True
            if element.Id.Value in outer.failLeaderElements:
                tag.SetLeaderElbow.side_effect = RuntimeError("Cannot set leader")
            doc.tags.append(tag)
            return tag

        tagClass = type("IndependentTag", (), {"Create": staticmethod(CreateTag)})
        self.api = {
            "__file__": str(SCRIPT), "math": math, "os": os, "traceback": traceback,
            "forms": NS(WPFWindow=FakeWindow), "script": NS(get_output=Mock()),
            "__revit__": NS(ActiveUIDocument=self.uidoc),
            "List": FakeList, "ElementId": object, "XYZ": Vector,
            "LocationPoint": FakeLocation, "ViewSheet": FakeSheet,
            "ViewType": NS(FloorPlan="FloorPlan", EngineeringPlan="EngineeringPlan"),
            "StorageType": NS(Double="Double"),
            "BuiltInCategory": NS(OST_Rooms="rooms", OST_ElectricalFixtures="fixtures",
                                  OST_ElectricalEquipment="equipment",
                                  OST_ElectricalFixtureTags="fixtureTags"),
            "BuiltInParameter": NS(SYMBOL_NAME_PARAM="typeName"),
            "FamilySymbol": type("FamilySymbol", (), {}), "IndependentTag": tagClass,
            "RevitLinkInstance": type("RevitLinkInstance", (), {}),
            "FilteredElementCollector": Collector, "SpatialElementBoundaryOptions": object,
            "Reference": lambda element: element,
            "LeaderEndCondition": NS(Free="Free"), "TagOrientation": NS(Horizontal="Horizontal"),
            "TransactionStatus": self.status,
            "TransactionGroup": lambda *args: Scope("group"),
            "Transaction": lambda *args: Scope("transaction"),
            "SubTransaction": lambda *args: Scope("sub"),
        }
        tree = ast.parse(SOURCE, filename=str(SCRIPT))
        tree.body = [node for node in tree.body
                     if isinstance(node, (ast.FunctionDef, ast.ClassDef, ast.Assign))]
        exec(compile(tree, str(SCRIPT), "exec"), self.api)
        self.settings = self.api["ValidateSettings"](False, 0.074, 0.375)
        for direction in ("LA", "RA"):
            for name in ("PANEL/CIRCUIT", "PANEL/CIRCUIT + AFF", "TYPE NAME", "TYPE NAME + AFF"):
                self.doc.symbols.append(self.MakeSymbol(direction, name))
        self.doc.fixtures[10] = [self.MakeFixture(1)]

    def MakeView(self, value, viewType="FloorPlan"):
        return NS(Id=NS(Value=value), ViewType=viewType, Scale=48,
                  Name="View {}".format(value), RightDirection=Vector(1, 0),
                  UpDirection=Vector(0, 1))

    def MakeSymbol(self, direction, name):
        return NS(Id=NS(Value=len(self.doc.symbols) + 200), IsActive=True, Category="fixtureTags",
                  Family=NS(Name="FFE_Tag_EF_Receptacle ({})".format(direction)),
                  get_Parameter=lambda key: NS(AsString=lambda: name))

    def MakeFixture(self, value, height=18.0, name="Standard Duplex", family="FFE_EF_RECEPTACLES"):
        return NS(Id=NS(Value=value), Symbol=NS(
            Family=NS(Name=family), get_Parameter=lambda key: NS(AsString=lambda: name)),
            LookupParameter=lambda key: NS(StorageType="Double", AsDouble=lambda: height / 12.0),
            Location=FakeLocation(Vector(value * 20, 0)), get_BoundingBox=lambda view: None)

    def RunTagging(self, write=False):
        self.settings["write"] = write
        return "\n".join(self.api["RunTagging"](self.doc, self.uidoc, self.settings))

    def RunCommand(self):
        output = io.StringIO()
        with contextlib.redirect_stdout(output):
            self.api["RunCommand"]()
        return output.getvalue()

    def test_dialog_defaults_validation_and_cancel(self):
        dialog = self.api["SettingsDialog"]()
        self.assertEqual(dialog.landingText.Text, "0.074")
        self.assertEqual(dialog.offsetText.Text, "0.375")
        self.assertTrue(dialog.writeMode.IsChecked)
        dialog.landingText.Text = "nan"
        dialog.RunClick(None, None)
        self.assertFalse(dialog.closed)
        self.assertIsNone(dialog.Settings)
        self.assertIn("finite", dialog.errorText.Text)
        dialog.CancelClick(None, None)
        self.assertTrue(dialog.closed)
        self.assertIsNone(dialog.Settings)

    def test_dialog_accepts_preview_and_resets_on_next_launch(self):
        dialog = self.api["SettingsDialog"]()
        dialog.writeMode.IsChecked = False
        dialog.offsetText.Text = "0.5"
        dialog.RunClick(None, None)
        self.assertFalse(dialog.Settings["write"])
        self.assertEqual(dialog.Settings["offset"], 0.5)
        self.assertTrue(dialog.closed)
        self.assertEqual(self.api["SettingsDialog"]().offsetText.Text, "0.375")

    def test_fixed_ffe_families_ignore_other_loaded_families(self):
        otherTag = self.MakeSymbol("LA", "PANEL/CIRCUIT")
        otherTag.Family.Name = "Other Electrical Tag (LA)"
        self.doc.symbols.insert(0, otherTag)
        # An old or externally supplied family setting cannot override the
        # fixed family pair now that the command selects alignment itself.
        self.settings["laFamily"] = "Other Electrical Tag (LA)"
        self.RunTagging(write=True)
        self.assertEqual(self.doc.tags[0].symbolId.Value, self.doc.symbols[1].Id.Value)

    def test_missing_ffe_families_are_reported_without_falling_back(self):
        for symbol in self.doc.symbols:
            symbol.Family.Name = symbol.Family.Name.replace("FFE_Tag_EF_Receptacle", "Other")
        report = self.RunTagging(write=True)
        self.assertIn("Required families: FFE_Tag_EF_Receptacle (LA) and "
                      "FFE_Tag_EF_Receptacle (RA).", report)
        self.assertIn("Missing tag types", report)
        self.assertEqual(self.events, [])

    def test_right_aligned_family_used_when_leader_meets_right_text_edge(self):
        for height, name, index in ((18, "Standard Duplex", 4),
                                    (24, "Standard Duplex", 5),
                                    (18, "USB", 6), (24, "USB", 7)):
            with self.subTest(height=height, name=name):
                self.setUp()
                self.doc.fixtures[10] = [self.MakeFixture(1, height=height, name=name)]
                # Occupy the right-hand placement so the optimizer chooses a
                # left-hand tag, whose right-aligned text faces the leader.
                obstacle = self.MakeFixture(90)
                obstacle.Location = FakeLocation(Vector(24, 2))
                obstacle.get_BoundingBox = lambda view: NS(
                    Min=Vector(21, 1), Max=Vector(27, 3))
                self.doc.equipment[10] = [obstacle]
                report = self.RunTagging(write=True)
                self.assertIn("RA candidates: 1", report)
                tag = self.doc.tags[0]
                self.assertEqual(tag.symbolId.Value, self.doc.symbols[index].Id.Value)
                self.assertEqual(self.doc.symbols[index].Family.Name,
                                 "FFE_Tag_EF_Receptacle (RA)")
                headX = tag.TagHeadPosition.X
                elbowX = tag.SetLeaderElbow.call_args.args[1].X
                endX = tag.SetLeaderEnd.call_args.args[1].X
                self.assertLess(headX, elbowX)
                self.assertLess(elbowX, endX)

    def test_invalid_settings_rejected(self):
        validate = self.api["ValidateSettings"]
        for invalid in ("", "text", 0, -1, "nan", "inf", "-inf", "1e999"):
            for landing, offset in ((invalid, 0.375), (0.074, invalid)):
                with self.subTest(landing=landing, offset=offset), self.assertRaises(ValueError):
                    validate(False, landing, offset)

    def test_cancel_and_window_close_do_not_run_tagging(self):
        self.api["SettingsDialog"] = Mock(return_value=NS(Settings=None, ShowDialog=Mock()))
        self.api["RunTagging"] = Mock()
        self.assertEqual(self.RunCommand(), "")
        self.api["RunTagging"].assert_not_called()
        self.assertEqual(self.events, [])

    def test_preview_classifies_all_four_types_without_writes_or_selection(self):
        self.doc.fixtures[10] = [self.MakeFixture(1), self.MakeFixture(2, height=24),
                                 self.MakeFixture(3, name="USB"),
                                 self.MakeFixture(4, height=24, name="USB")]
        report = self.RunTagging()
        for line in ("Tag candidates: 4", "Panel/Circuit: 1", "Panel/Circuit + AFF: 1",
                     "Type Name: 1", "Type Name + AFF: 1", "Tags created: 0", "Tags selected: 0"):
            self.assertIn(line, report)
        self.assertEqual(self.events, [])
        self.assertEqual(self.doc.tags, [])
        self.uidoc.Selection.SetElementIds.assert_not_called()

    def test_missing_tag_types_prevents_writes(self):
        self.doc.symbols.pop()
        self.assertIn("RA / TYPE NAME + AFF", self.RunTagging(write=True))
        self.assertEqual(self.events, [])

    def test_existing_tags_and_non_ffe_fixtures_are_skipped(self):
        existingTag = NS(Id=NS(Value=500), get_BoundingBox=lambda view: None,
                         GetTaggedLocalElementIds=lambda: [NS(Value=1)])
        self.doc.existing[10] = [existingTag]
        self.doc.fixtures[10].append(self.MakeFixture(2, family="Other receptacle"))
        report = self.RunTagging(write=True)
        self.assertIn("Already tagged", report)
        self.assertIn("Skipped: 2", report)
        self.assertIn("Tag candidates: 0", report)
        self.assertEqual(self.events, [])

    def test_sheet_only_processes_eligible_placed_plans(self):
        self.doc.ActiveView = FakeSheet([NS(Value=21), NS(Value=22), NS(Value=23)])
        self.doc.elements = {21: NS(ViewId=NS(Value=10)), 22: NS(ViewId=NS(Value=11)),
                             23: NS(ViewId=NS(Value=12)), 10: self.view,
                             11: self.MakeView(11, "EngineeringPlan"),
                             12: self.MakeView(12, "ThreeD")}
        self.doc.fixtures[11] = [self.MakeFixture(2)]
        self.doc.fixtures[12] = [self.MakeFixture(3)]
        report = self.RunTagging()
        self.assertIn("Views processed: 2", report)
        self.assertIn("Tag candidates: 2", report)

    def test_ineligible_view_or_sheet_has_actionable_report(self):
        self.doc.ActiveView = self.MakeView(10, "ThreeD")
        self.assertIn("ERROR: Open a floor plan or sheet.", self.RunTagging())
        self.doc.ActiveView = FakeSheet([])
        self.assertIn("ERROR: No eligible floor plans.", self.RunTagging())
        self.assertEqual(self.events, [])

    def test_write_commits_before_selection_and_preserves_leader_geometry(self):
        def CheckSelection(ids):
            self.assertEqual(self.events[-1], ("group", "Assimilate"))
            self.assertEqual([i.Value for i in ids], [101])
        self.uidoc.Selection.SetElementIds.side_effect = CheckSelection
        report = self.RunTagging(write=True)
        self.assertIn("Tags created: 1", report)
        self.assertIn("Tags selected: 1", report)
        tag = self.doc.tags[0]
        self.assertEqual(tag.viewId.Value, 10)
        self.assertEqual(tag.symbolId.Value, self.doc.symbols[0].Id.Value)
        self.assertEqual(self.doc.symbols[0].Family.Name, "FFE_Tag_EF_Receptacle (LA)")
        self.assertAlmostEqual(tag.TagHeadPosition.X, 21.796)
        self.assertAlmostEqual(tag.TagHeadPosition.Y, 1.7988)
        elbow = tag.SetLeaderElbow.call_args.args[1]
        self.assertAlmostEqual(elbow.X, 21.5)
        self.assertAlmostEqual(elbow.Y, 1.5)
        self.assertEqual(tag.SetLeaderEnd.call_args.args[1].X, 20)
        self.assertTrue(all(scope.disposed for scope in self.scopes))
        self.assertEqual([s.kind for s in self.scopes], ["group", "transaction", "sub"])

    def test_failed_tag_rolls_back_partial_tag_and_continues(self):
        self.doc.fixtures[10].append(self.MakeFixture(2))
        self.failLeaderElements.add(1)
        report = self.RunTagging(write=True)
        self.assertIn("ERROR 1 | Cannot set leader", report)
        self.assertIn("Tags created: 1", report)
        self.assertEqual([tag.element.Id.Value for tag in self.doc.tags], [2])
        self.assertIn(("sub", "RollBack"), self.events)
        self.assertNotIn("CREATED | Element 1", report)

    def test_all_tag_failures_do_not_select_anything(self):
        self.failElements.add(1)
        self.assertIn("Tags created: 0", self.RunTagging(write=True))
        self.uidoc.Selection.SetElementIds.assert_not_called()
        self.assertEqual(self.doc.tags, [])

    def test_selection_warning_preserves_committed_result(self):
        self.uidoc.Selection.SetElementIds.side_effect = RuntimeError("Selection unavailable")
        report = self.RunTagging(write=True)
        self.assertIn("Tags created: 1", report)
        self.assertIn("Tags selected: 0", report)
        self.assertIn("SELECTION WARNING: Selection unavailable", report)
        self.assertEqual(len(self.doc.tags), 1)
        self.assertNotIn(("group", "RollBack"), self.events)

    def test_inactive_symbol_is_activated_only_in_write_mode(self):
        symbol = self.doc.symbols[0]
        symbol.IsActive = False
        symbol.Activate = Mock(side_effect=lambda: setattr(symbol, "IsActive", True))
        self.RunTagging()
        symbol.Activate.assert_not_called()
        self.RunTagging(write=True)
        symbol.Activate.assert_called_once_with()

    def test_rerun_skips_new_tags(self):
        self.RunTagging(write=True)
        self.doc.existing[10] = [NS(
            Id=tag.Id, get_BoundingBox=lambda view: None,
            GetTaggedLocalElementIds=lambda: [NS(Value=1)]) for tag in self.doc.tags]
        self.events.clear()
        report = self.RunTagging(write=True)
        self.assertIn("Tag candidates: 0", report)
        self.assertEqual(len(self.doc.tags), 1)
        self.assertEqual(self.events, [])

    def test_commit_failures_roll_back_and_never_publish_success(self):
        for kind, operation in (("transaction", "Commit"), ("group", "Assimilate")):
            for status in ("RolledBack", "Started"):
                with self.subTest(kind=kind, status=status):
                    self.setUp()
                    self.failKind, self.failOperation, self.failResult = kind, operation, status
                    logs, errors = [], []
                    fixture = self.doc.fixtures[10][0]
                    job = (self.view, fixture, self.doc.symbols[0], Vector(20, 0),
                           Vector(21.5, 1.5), Vector(21.796, 1.7988), "LA", "PANEL/CIRCUIT", [])
                    with self.assertRaisesRegex(RuntimeError, "did not commit"):
                        self.api["CreateTags"](self.doc, [job], logs, errors)
                    self.assertEqual(self.doc.tags, [])
                    self.assertEqual(logs, [])
                    self.assertTrue(all(scope.disposed for scope in self.scopes))
                    self.assertFalse(any(scope.GetStatus() == "Started" for scope in self.scopes))
                    self.uidoc.Selection.SetElementIds.assert_not_called()

    def test_transaction_exceptions_roll_back_every_started_scope(self):
        for kind, operation in (("group", "Start"), ("transaction", "Start"),
                                ("transaction", "Commit"), ("group", "Assimilate")):
            with self.subTest(kind=kind, operation=operation):
                self.setUp()
                self.raiseKind, self.raiseOperation = kind, operation
                self.settings["write"] = True
                self.api["SettingsDialog"] = Mock(
                    return_value=NS(Settings=self.settings, ShowDialog=Mock()))
                self.assertIn("FATAL SCRIPT ERROR", self.RunCommand())
                self.assertEqual(self.doc.tags, [])
                self.assertTrue(all(scope.disposed for scope in self.scopes))
                self.assertFalse(any(scope.GetStatus() == "Started" for scope in self.scopes))
                self.uidoc.Selection.SetElementIds.assert_not_called()

    def test_read_only_write_is_blocked_but_preview_runs(self):
        self.doc.IsReadOnly = True
        self.settings["write"] = True
        self.api["SettingsDialog"] = Mock(return_value=NS(Settings=self.settings, ShowDialog=Mock()))
        self.assertIn("read-only", self.RunCommand())
        self.settings["write"] = False
        self.assertIn("Mode: PREVIEW", self.RunCommand())
        self.assertEqual(self.events, [])

    def test_fatal_error_is_printed_and_tags_are_not_selected(self):
        self.settings["write"] = True
        self.failKind, self.failOperation = "transaction", "Commit"
        self.api["SettingsDialog"] = Mock(return_value=NS(Settings=self.settings, ShowDialog=Mock()))
        report = self.RunCommand()
        self.assertIn("FATAL SCRIPT ERROR", report)
        self.assertIn("Traceback", report)
        self.assertNotIn("CREATED |", report)
        self.assertEqual(self.doc.tags, [])
        self.uidoc.Selection.SetElementIds.assert_not_called()

    def test_no_active_document_is_reported_without_dialog(self):
        self.api["__revit__"].ActiveUIDocument = None
        self.api["SettingsDialog"] = Mock()
        self.assertIn("Open a Revit project", self.RunCommand())
        self.api["SettingsDialog"].assert_not_called()

    def test_xaml_controls_and_handlers_match_dialog(self):
        root = ET.parse(BUNDLE / "Settings.xaml").getroot()
        names = {node.get("{http://schemas.microsoft.com/winfx/2006/xaml}Name")
                 for node in root.iter()}
        self.assertTrue({"runButton", "landingText", "offsetText", "writeMode",
                         "previewMode", "errorText"}.issubset(names))
        self.assertFalse(any(node.tag.endswith("ComboBox") for node in root.iter()))
        self.assertTrue(any(node.get("Text") == "FFE_Tag_EF_Receptacle" for node in root.iter()))
        for node in root.iter():
            if node.get("Click"):
                self.assertTrue(hasattr(self.api["SettingsDialog"], node.get("Click")))


if __name__ == "__main__":
    unittest.main()
