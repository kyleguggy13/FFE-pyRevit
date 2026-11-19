# Revit Data Table acceptance checks

Run in both Revit 2025 and Revit 2026 with the installed extension and supported pyRevit IronPython engine. These checks require an actual Revit project; the automated browser test uses a simulated bridge.

## Startup and coexistence

1. Reload pyRevit and launch **MEPTools → Revit Data Table** from a project. Verify the window is modeless and owned by Revit, with Level, Number, Name, Area, and Height columns.
2. Launch it again for the same model, including while the window is minimized. Verify the existing window is restored and focused.
3. In fresh Revit sessions, test opening this tool before and after Keynotes and pyRevit's output window. All tools must continue working without mixed WebView2 assembly versions.
4. Disconnect the network and reopen the tool. Its local assets and full grid functionality must still load.
5. Launch with no document, a family document, and a project containing no Spaces. Verify a clear prompt or a valid empty table, without an unhandled exception.

## Data and units

1. Compare several rows with a Revit Space schedule: level, number, plain name, area, and unbounded height. Include leading-zero and alphanumeric numbers, repeated numbers on different phases, and non-English names.
2. Compare an imperial project and a metric project. Values must follow project formatting; sorting Area/Height must use numeric order. Change project units while the window is open, then Refresh.
3. Include placed, unplaced, unenclosed/redundant, and zero-area Spaces. All remain discoverable; zero is distinct from unavailable values, and row warnings explain missing data. An unreadable Space must not hide all remaining rows.
4. Verify the initial view phase, a view with no phase parameter, switching phases, All phases revealing the Phase column, and combined Phase/Level/search filtering.
5. Inspect a Space in RevitLookup whose `ROOM_PHASE_ID`/`ROOM_PHASE` references an existing project phase while `CreatedPhaseId` is invalid or different. Verify the table shows the spatial phase from `AsElementId()` and includes the Space when that phase is selected. Spaces with unreadable phase parameters must remain under All phases with a warning.

## Interaction and freshness

1. Verify initial ordering by level elevation, then natural number order. Sort multiple columns, resize/reorder headers, and select a range of rows. Refresh must retain valid filters, sorting, column layout, and selection.
2. Verify Up/Down, Home/End, and Space after focusing the grid. Tab must reach filters and navigation buttons, with visible focus.
3. Select several rows and use Select in Revit. Show placed Spaces in a suitable plan/3D view. Try Show for unplaced Spaces and from a view that cannot display them; verify an actionable skipped/failure report.
4. Rename, add, or delete Spaces in Revit while the table is open. Verify Out of date appears and rows remain unchanged until Refresh. Try navigation to a deleted Space before Refresh; it must skip that row safely. Repeat after undo/redo.
5. Change a phase/level filter to hide a selected row. Navigation must act only on currently visible selected rows.
6. Test several thousand Spaces: scrolling should remain responsive, headers remain visible, and the grid should occupy available window space. Check both the default size and minimum window size at common Windows display scaling settings.

## Lifecycle and document safety

1. Open two projects, and open a separate table for each. Switching the active project must disable Refresh/Select/Show in the other table while leaving snapshot browsing available.
2. Queue navigation while Revit is busy, then switch documents before it executes. The request must fail clearly without navigating the other document.
3. Close one project and verify only its table closes. Repeat with Revit shutdown and with a table closed before a pending event executes. No event errors, orphan windows, or retained document actions should remain.
4. Reopen/close the table repeatedly, then edit the model. Verify no duplicated callbacks or stale windows appear.
5. In an otherwise unchanged project, use Refresh, sorting, filters, selection, and Show. Confirm this command does not mark the model modified or add transactions to Undo.
