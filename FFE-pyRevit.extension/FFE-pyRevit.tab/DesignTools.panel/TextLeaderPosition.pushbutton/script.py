# -*- coding: utf-8 -*-
# -*- coding: utf-8 -*-
__title__     = "Text Leader \nPosition"
__version__   = 'Version = 1.1'
__doc__       = """Version = 1.1
Date    = 11.02.2026
_____________________________________________________________________
Description:
- This script allows you to change the left/right 
  attachment position of text leaders in Revit.
- The script will prompt you to select the 
  attachment position
- The available positions are Top, Middle, and Bottom.
_____________________________________________________________________
How-to:
-> Select the text elements you want to change
-> Click the button
-> Choose the Left Attachment position from the list
-> Choose the Right Attachment position from the list
_____________________________________________________________________
Last update:
- [06.05.2025] - 1.0 RELEASE
- [11.02.2026] - 1.1 Added error handling and reporting
_____________________________________________________________________
Inspiration: Olivia Bates
Author: Kyle Guggenheim"""


# """Change text leader attachments and report failures in the pyRevit output.

# Key behaviors: report startup, selection, parameter, transaction, and log errors;
# continue past individual unsupported parameters; preserve left/right choices.
# Revit API notes: targets Revit 2025/2026; uses ElementId.Value for diagnostics.
# Design decisions: import only required pyRevit APIs, keep dialogs outside the
# transaction, and keep usage-log failures separate from document-edit results.
# """
# __title__ = "Text Leader \nPosition"
# __version__ = "1.1"
# __doc__ = """Change the left/right attachment position of selected text leaders.

# Select text notes, then choose Top, Middle, or Bottom for each side.
# Cancel either list to leave that side unchanged. Diagnostic details and a final
# report are printed to the pyRevit output window.

# Author: Kyle Guggenheim
# Inspiration: Olivia Bates
# """

import json
import os
import sys
import time
import traceback


def ReportError(context):
    """Print the failing operation and the active exception's full traceback."""
    print("ERROR: {0}".format(context))
    print(traceback.format_exc())


# Print before third-party imports so dependency failures cannot leave a blank
# output window. Unused imports (including requests) are intentionally omitted.
print("Text Leader Position {0}: starting".format(__version__))
print("Script: {0}".format(__file__))
print("Python: {0}".format(sys.version))
try:
    import pyrevit
    from pyrevit import DB, forms
    print("pyRevit: {0}".format(getattr(pyrevit, "__version__", "unknown")))
    print("[Startup] Required imports loaded.")
except Exception:
    ReportError("Loading pyRevit/Revit dependencies. Share this output with support.")
    raise


def LogAction(doc, logStatus):
    """Append the existing usage-log entry; let the caller report I/O failures."""
    app = doc.Application
    logDir = os.path.join(
        os.path.expanduser("~"), "FFE Inc", "FFE Revit Users - Documents",
        "00-General", "Revit_Add-Ins", "FFE-pyRevit", "Logs")
    logFile = os.path.join(logDir, app.Username + "_revit_log.json")
    print("[Usage log] Writing: {0}".format(logFile))
    dataEntry = {
        "datetime": time.strftime("%Y-%m-%d %H:%M:%S"),
        "username": app.Username,
        "doc_title": doc.Title,
        "doc_path": doc.PathName or "<Untitled>",
        "revit_version_number": app.VersionNumber,
        "revit_build": app.VersionBuild,
        "action": "Text Leader Position",
        "status": logStatus
    }
    if not os.path.exists(logDir):
        os.makedirs(logDir)
    if not os.path.exists(logFile):
        with open(logFile, "w") as logStream:
            json.dump({"action": [dataEntry]}, logStream, indent=4)
    else:
        with open(logFile, "r+") as logStream:
            fileData = json.load(logStream)
            fileData.setdefault("action", []).append(dataEntry)
            logStream.seek(0)
            json.dump(fileData, logStream, indent=4)
            logStream.truncate()
    print("[Usage log] Entry saved.")


def Main():
    """Run the command with stage diagnostics and guarded transaction cleanup."""
    stage = "Reading the active Revit document"
    doc = None
    transaction = None
    logStatus = "Failed"
    try:
        app = __revit__.Application
        print("Revit: {0}; build: {1}".format(app.VersionNumber, app.VersionBuild))
        uidoc = __revit__.ActiveUIDocument
        if uidoc is None:
            print("[Stopped] No active document. Open a model and select text notes.")
            return
        doc = uidoc.Document
        print("Document: {0}; read-only: {1}".format(doc.Title, doc.IsReadOnly))

        stage = "Reading selected elements"
        selectedIds = list(uidoc.Selection.GetElementIds())
        print("[Selection] Selected elements: {0}".format(len(selectedIds)))
        if not selectedIds:
            logStatus = "Cancelled"
            print("[Stopped] Nothing selected. Select text notes before running the tool.")
            return

        stage = "Inspecting leader attachment parameters"
        elementParameters = []
        skippedElements = 0
        for elementId in selectedIds:
            context = "Element {0}".format(elementId.Value)
            try:
                element = doc.GetElement(elementId)
                if element is None:
                    skippedElements += 1
                    print("[Skipped] {0}: could not retrieve element.".format(context))
                    continue
                context = "{0} ({1})".format(context, element.GetType().Name)
                leftParam = element.get_Parameter(DB.BuiltInParameter.LEADER_LEFT_ATTACHMENT)
                rightParam = element.get_Parameter(DB.BuiltInParameter.LEADER_RIGHT_ATTACHMENT)
                print("[Selection] {0}; left parameter: {1}; right parameter: {2}".format(
                    context, leftParam is not None, rightParam is not None))
                if leftParam is None and rightParam is None:
                    skippedElements += 1
                    print("[Skipped] {0}: no leader attachment parameters; select text notes.".format(context))
                    continue
                elementParameters.append((context, leftParam, rightParam))
            except Exception:
                skippedElements += 1
                ReportError("Inspecting {0}".format(context))

        print("[Selection] Usable elements: {0}; skipped elements: {1}".format(
            len(elementParameters), skippedElements))
        if not elementParameters:
            print("[Stopped] No usable leader attachment parameters found.")
            return

        stage = "Showing left attachment selection dialog"
        options = ["Top", "Middle", "Bottom"]
        print("[Dialog] Waiting for left attachment choice...")
        leftChoice = forms.SelectFromList.show(
            options, title="Select Left Leader Attachment Position", button_name="Apply")
        print("[Dialog] Left attachment: {0}".format(leftChoice or "unchanged (cancelled)"))
        stage = "Showing right attachment selection dialog"
        print("[Dialog] Waiting for right attachment choice...")
        rightChoice = forms.SelectFromList.show(
            options, title="Select Right Leader Attachment Position", button_name="Apply")
        print("[Dialog] Right attachment: {0}".format(rightChoice or "unchanged (cancelled)"))
        if not leftChoice and not rightChoice:
            logStatus = "Cancelled"
            print("[Stopped] Both attachment choices cancelled. No changes made.")
            return

        stage = "Starting the Text Leader transaction"
        print("[Transaction] Starting...")
        transaction = DB.Transaction(doc, "Text Leader")
        startStatus = transaction.Start()
        if startStatus != DB.TransactionStatus.Started:
            raise RuntimeError("Transaction did not start: {0}".format(startStatus))

        attachmentMap = {"Top": 0, "Middle": 1, "Bottom": 2}
        changedParameters = 0
        skippedParameters = 0
        failedParameters = 0
        stage = "Setting leader attachment parameters"
        for context, leftParam, rightParam in elementParameters:
            for side, parameter, choice in (
                    ("Left", leftParam, leftChoice), ("Right", rightParam, rightChoice)):
                if not choice:
                    continue
                parameterContext = "{0}, {1} attachment".format(context, side)
                try:
                    if parameter is None:
                        skippedParameters += 1
                        print("[Skipped] {0}: parameter is missing.".format(parameterContext))
                        continue
                    if parameter.IsReadOnly:
                        skippedParameters += 1
                        print("[Skipped] {0}: parameter is read-only.".format(parameterContext))
                        continue
                    if parameter.StorageType != DB.StorageType.Integer:
                        skippedParameters += 1
                        print("[Skipped] {0}: expected Integer storage, found {1}.".format(
                            parameterContext, parameter.StorageType))
                        continue
                    if not parameter.Set(attachmentMap[choice]):
                        raise RuntimeError("Parameter.Set returned False")
                    changedParameters += 1
                    print("[Updated] {0}: {1}".format(parameterContext, choice))
                except Exception:
                    failedParameters += 1
                    ReportError("Setting {0} to {1}".format(parameterContext, choice))

        stage = "Committing the Text Leader transaction"
        commitStatus = transaction.Commit()
        print("[Transaction] Commit result: {0}".format(commitStatus))
        if commitStatus != DB.TransactionStatus.Committed:
            raise RuntimeError("Changes were not committed; transaction status: {0}".format(commitStatus))
        if changedParameters:
            logStatus = "Partial Success" if (skippedElements or skippedParameters or failedParameters) else "Success"
        print("[Result] {0}: {1} attachment updates committed; {2} elements skipped; "
              "{3} attachments skipped; {4} attachment errors.".format(
                  logStatus, changedParameters, skippedElements, skippedParameters, failedParameters))
    except Exception:
        ReportError(stage)
        if transaction is not None:
            try:
                if transaction.GetStatus() == DB.TransactionStatus.Started:
                    transaction.RollBack()
                    print("[Transaction] Rolled back. No attachment updates committed.")
            except Exception:
                ReportError("Rolling back the Text Leader transaction")
        print("[Result] Failed during: {0}. Share the full output above for troubleshooting.".format(stage))
    finally:
        if doc is not None:
            try:
                LogAction(doc, logStatus)
            except Exception:
                print("[Warning] Usage log could not be saved. Document-edit result: {0}.".format(logStatus))
                ReportError("Writing the usage log (check the printed path, folder access, and JSON file)")
        print("Text Leader Position: finished ({0}).".format(logStatus))


Main()
