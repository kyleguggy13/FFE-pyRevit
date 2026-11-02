"""Verify diagnostic and transaction failure paths without starting Revit.

Key behaviors: exercise startup failures, parameter errors, and usage-log errors.
Revit API notes: mock the Revit 2025/2026 API boundary, including ElementId.Value.
Design decisions: run the actual command with only its external APIs substituted.
"""
import ast
import contextlib
import io
from pathlib import Path
import sys
import types
import unittest
from unittest.mock import Mock, patch


SCRIPT = (Path(__file__).resolve().parents[1]
          / "FFE-pyRevit.extension/FFE-pyRevit.tab/DesignTools.panel"
          / "TextLeaderPosition.pushbutton/script.py")
SOURCE = SCRIPT.read_text(encoding="utf-8")


class TextLeaderPositionTests(unittest.TestCase):
    def setUp(self):
        self.transaction = Mock()
        self.transaction.Start.return_value = "Started"
        self.transaction.Commit.return_value = "Committed"
        self.transaction.GetStatus.return_value = "Started"
        self.left = self.MakeParameter()
        self.right = self.MakeParameter()
        element = Mock()
        element.GetType.return_value.Name = "TextNote"
        element.get_Parameter.side_effect = [self.left, self.right]
        self.doc = Mock(Title="Test model", IsReadOnly=False)
        self.doc.GetElement.return_value = element
        self.uidoc = Mock(Document=self.doc)
        self.uidoc.Selection.GetElementIds.return_value = [types.SimpleNamespace(Value=123)]
        app = types.SimpleNamespace(VersionNumber="2026", VersionBuild="test-build")
        self.host = types.SimpleNamespace(Application=app, ActiveUIDocument=self.uidoc)
        self.forms = Mock()
        self.forms.SelectFromList.show.side_effect = ["Top", "Bottom"]
        db = types.SimpleNamespace(
            Transaction=Mock(return_value=self.transaction),
            TransactionStatus=types.SimpleNamespace(Started="Started", Committed="Committed"),
            StorageType=types.SimpleNamespace(Integer="Integer"),
            BuiltInParameter=types.SimpleNamespace(
                LEADER_LEFT_ATTACHMENT="left", LEADER_RIGHT_ATTACHMENT="right"))
        self.pyrevit = types.ModuleType("pyrevit")
        self.pyrevit.DB = db
        self.pyrevit.forms = self.forms
        self.output = io.StringIO()
        tree = ast.parse(SOURCE, filename=str(SCRIPT))
        # Define the real functions while allowing each test to configure the
        # command's external state before invoking Main.
        tree.body = tree.body[:-1]
        self.namespace = {"__file__": str(SCRIPT), "__revit__": self.host}
        with patch.dict(sys.modules, {"pyrevit": self.pyrevit}), contextlib.redirect_stdout(self.output):
            exec(compile(tree, str(SCRIPT), "exec"), self.namespace)
        self.logAction = Mock()
        self.namespace["LogAction"] = self.logAction

    def MakeParameter(self):
        parameter = Mock(IsReadOnly=False, StorageType="Integer")
        parameter.Set.return_value = True
        return parameter

    def RunCommand(self):
        with contextlib.redirect_stdout(self.output):
            self.namespace["Main"]()
        return self.output.getvalue()

    def test_success_reports_versions_and_committed_updates(self):
        output = self.RunCommand()
        self.assertIn("Text Leader Position 1.1: starting", output)
        self.assertIn("Revit: 2026; build: test-build", output)
        self.assertIn("2 attachment updates committed", output)
        self.left.Set.assert_called_once_with(0)
        self.right.Set.assert_called_once_with(2)
        self.logAction.assert_called_once_with(self.doc, "Success")

    def test_import_failure_prints_startup_and_traceback(self):
        output = io.StringIO()
        with patch.dict(sys.modules, {"pyrevit": None}), contextlib.redirect_stdout(output):
            with self.assertRaises(ImportError):
                exec(compile(SOURCE, str(SCRIPT), "exec"), {"__file__": str(SCRIPT)})
        self.assertIn("starting", output.getvalue())
        self.assertIn("Loading pyRevit/Revit dependencies", output.getvalue())
        self.assertIn("Traceback", output.getvalue())

    def test_transaction_start_exception_has_stage_and_traceback(self):
        self.transaction.Start.side_effect = RuntimeError("document locked")
        self.transaction.GetStatus.return_value = "Uninitialized"
        output = self.RunCommand()
        self.assertIn("Failed during: Starting the Text Leader transaction", output)
        self.assertIn("document locked", output)
        self.assertIn("Traceback", output)
        self.transaction.RollBack.assert_not_called()

    def test_commit_exception_rolls_back_without_reporting_success(self):
        self.transaction.Commit.side_effect = RuntimeError("commit rejected")
        output = self.RunCommand()
        self.assertIn("Failed during: Committing", output)
        self.assertIn("Rolled back. No attachment updates committed", output)
        self.transaction.RollBack.assert_called_once()
        self.logAction.assert_called_once_with(self.doc, "Failed")

    def test_unsuccessful_commit_status_is_reported(self):
        self.transaction.Commit.return_value = "RolledBack"
        self.transaction.GetStatus.return_value = "RolledBack"
        output = self.RunCommand()
        self.assertIn("Changes were not committed", output)
        self.assertNotIn("[Result] Success", output)
        self.transaction.RollBack.assert_not_called()

    def test_parameter_error_reports_element_and_continues_other_side(self):
        self.left.Set.side_effect = RuntimeError("owned by another user")
        output = self.RunCommand()
        self.assertIn("Element 123 (TextNote), Left attachment", output)
        self.assertIn("owned by another user", output)
        self.assertIn("1 attachment updates committed", output)
        self.assertIn("1 attachment errors", output)
        self.right.Set.assert_called_once_with(2)
        self.logAction.assert_called_once_with(self.doc, "Partial Success")

    def test_missing_and_readonly_parameters_explain_skips(self):
        self.doc.GetElement.return_value.get_Parameter.side_effect = [None, self.right]
        self.right.IsReadOnly = True
        output = self.RunCommand()
        self.assertIn("parameter is missing", output)
        self.assertIn("parameter is read-only", output)
        self.assertIn("2 attachments skipped", output)
        self.right.Set.assert_not_called()

    def test_log_failure_preserves_success_result(self):
        self.logAction.side_effect = OSError("access denied")
        output = self.RunCommand()
        self.assertIn("Usage log could not be saved. Document-edit result: Success", output)
        self.assertIn("access denied", output)
        self.assertIn("finished (Success)", output)

    def test_empty_selection_does_not_open_dialogs_or_start_transaction(self):
        self.uidoc.Selection.GetElementIds.return_value = []
        output = self.RunCommand()
        self.assertIn("Nothing selected", output)
        self.forms.SelectFromList.show.assert_not_called()
        self.transaction.Start.assert_not_called()

    def test_cancelling_both_choices_does_not_start_transaction(self):
        self.forms.SelectFromList.show.side_effect = [None, None]
        output = self.RunCommand()
        self.assertIn("Both attachment choices cancelled", output)
        self.transaction.Start.assert_not_called()

    def test_cancelling_one_choice_preserves_other_side(self):
        self.forms.SelectFromList.show.side_effect = [None, "Middle"]
        self.RunCommand()
        self.left.Set.assert_not_called()
        self.right.Set.assert_called_once_with(1)

    def test_no_document_prints_actionable_message(self):
        self.host.ActiveUIDocument = None
        output = self.RunCommand()
        self.assertIn("No active document", output)
        self.transaction.Start.assert_not_called()
        self.logAction.assert_not_called()


if __name__ == "__main__":
    unittest.main()
