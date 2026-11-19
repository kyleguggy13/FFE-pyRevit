# Revit Data Table — v0.1.1

Open **FFE-pyRevit → MEPTools → Revit Data Table** in a Revit 2025 or 2026 project. Reload pyRevit after installing the command so its ribbon button appears.

The modeless window displays host-model Spaces, including unplaced and zero-area Spaces. The initial columns are Level, Number, Name, Area, and Height. Height is `Space.UnboundedHeight`, not an average from volume/area. Measurements use Revit project units; sorting uses raw numbers. Space numbers remain text, including leading zeros.

Space phases are read from `ROOM_PHASE_ID`, falling back to `ROOM_PHASE`, using `AsElementId()`. The general element `CreatedPhaseId` is not used for Space phase filtering. If neither parameter resolves to a project phase, the row remains visible under All phases with a warning.

## Working with the table

- Choose a Phase and Level, or search by number, name, level, or phase. The initial phase matches the launching view; views without a phase use the last project phase. All phases reveals a Phase column.
- Click column headers to sort. Ctrl+click or Shift+click another header adds sorting. Drag headers to reorder columns and drag column edges to resize.
- Click rows to toggle selection; Shift+click selects a range. Focus the grid and use Up/Down, Home/End, and Space for keyboard selection. Changing filters clears selections hidden by the new filters.
- **Select in Revit** changes Revit's selection. **Show in Revit** asks Revit to locate selected Spaces in an available view. Unplaced Spaces have no viewable geometry and are skipped; individual row warnings explain unavailable values.
- **Refresh** reloads model data while preserving valid filters, column sorting/order/widths, and selected rows. **Clear filters** returns to the snapshot's default phase, All levels, and an empty search.
- Model edits, undo, and redo mark the table **Out of date**. They do not automatically replace its rows. Deleted or replaced Spaces are checked again before navigation.
- Each document has its own window. Reopening the command focuses that model's existing window. Activate the corresponding model before using its Refresh, Select, or Show actions. Closing the document closes its table.

The warning panel reports unreadable and skipped rows. This version has no cell editing, model writes, transactions, linked-model collection, cloud storage, or network requests.

## Architecture and dependencies

The native WPF host owns a WebView2 control and a queued Revit `ExternalEvent` bridge. The browser receives structured JSON snapshots; it never receives Revit objects. This pushbutton's `lib/revit_data_table.py` owns data collection, request validation, and navigation. Its `lib/ffe_webview2.py` adapts Keynotes' tested assembly compatibility loader without changing Keynotes. The command adds its own `lib` directory to the import path.

Tabulator **6.3.1** is pinned and bundled under `support/vendor`; runtime loading is fully offline. The source files are the published `tabulator-tables@6.3.1` package's `dist/js/tabulator.min.js`, `dist/css/tabulator.min.css`, and `LICENSE`, downloaded from jsDelivr. Keep `LICENSE.tabulator` when updating the assets. No Node/npm installation or frontend build is required for users.

WebView2 browser data lives under `%LOCALAPPDATA%\FFE-pyRevit\RevitDataTable\WebView2`. The loader reuses compatible loaded SDK assemblies, otherwise preferring pyRevit's runtime-specific pair before Revit's pair. Initialization errors report versions/paths. Fully restart Revit after changing SDK/add-in installations; reloading pyRevit cannot replace already-loaded DLLs.

## Validation

From the `RevitDataTable.pushbutton` directory, with CPython 3 and Node 22+:

```powershell
python -m unittest discover -s tests -p 'test_revit_data_table.py' -v
node --test tests/revit_data_table.test.cjs
node tests/revit_data_table_browser.cjs
```

All table-specific tests live in this pushbutton's `tests` directory and run separately from repository-level test discovery. The Python loader tests reuse the shared Keynotes fixture at the repository's `tests/test_keynote_webview2.py`.

The browser test uses installed Microsoft Edge headlessly with a simulated Revit bridge and no npm dependencies. Set `EDGE_PATH` to override the executable path. It writes a screenshot to a temporary browser profile. It validates browser behavior; actual Revit integration requires the local [acceptance checks](tests/revit_data_table_acceptance.md).
