# FFE Keynote Manager User Guide

The FFE Keynote Manager is a Revit tool for reviewing, editing, saving, and placing keynotes from the keynote file assigned to the active Revit document. It is intended for users who already understand the basics of Revit keynotes and generic annotations.

The manager edits the shared keynote text file used by the model. Use it with the same care you would use when editing the keynote file directly.

## Opening the Manager

Open the tool from the FFE-pyRevit ribbon. The manager opens in its own window and loads the keynote file assigned to the active Revit document.

At the top of the window, confirm that the correct Revit document is listed. If the wrong document is active, switch back to Revit, activate the correct model, and reopen or refresh the manager.

## Main Areas

### Status and Warnings

The message bar shows the current tool status, such as ready, syncing, validation required, or error. The `Warnings` pill shows the current number of warnings or errors.

Click `Warnings` to open the warnings sidebar. When a warning can be tied to a keynote row, selecting it will jump to the related keynote.

### Divisions

The `DIVISIONS` panel lists the top-level keynote divisions. Select a division to view and edit the keynote rows inside it.

On wider windows, the Divisions panel can be collapsed. On narrower windows, the Divisions panel is hidden so the keynote table has more room.

### Action Buttons

The action row contains the main editing commands:

- `Add Parent` creates a new top-level division.
- `Add Note in Sequence` adds a keynote under the same parent as the selected keynote.
- `Add Sub-Note` adds a child keynote under the selected keynote.
- `Duplicate` copies the selected division or keynote row.
- `Delete` removes the selected row if it does not still have child keynotes.

Some buttons are disabled until a valid row is selected.

### Search, Show, and Place As

Use `Search` to find keynotes by key, description, or parent key.

Use `Show` to control which keynotes are visible:

- `All Keynotes` shows the full list.
- `Placed Keynotes` shows keynotes that are currently placed in the model.
- `Unused Keynotes` shows keynotes that are not currently placed in the model.

The placed and unused filters use placement information collected from the model. Use `Collect Analytics` when you need the filters to reflect current model placement.

Use `Place As` to choose how the place button in the keynote table behaves:

- `User Keynote` places a standard Revit user keynote.
- `Generic Annotation` places the keynote as an FFE generic annotation keynote symbol.

## Editing Keynotes

Select a division, then edit rows directly in the keynote table.

Each row has a keynote key and description. Child keynotes are shown in a tree under their parent. Use the expand and collapse controls in the key column to show or hide child rows.

The selected division header shows the current division key and description. Division descriptions can be edited there. Division keys are protected during normal editing; if key editing is enabled in your version, use extra care because changing keys can affect existing placed references.

Edits are not written to the keynote file until you click `Save`.

## Keynote Text File Format

The shared keynote file is tab-delimited text. Data rows use either `Key<TAB>Text` for parent rows or `Key<TAB>Text<TAB>ParentKey` for child rows.

Rows whose first non-space character is `#` are treated as comments. Revit does not load those rows as keynotes, and the manager ignores them while reading the file.

On save, the manager writes a metadata comment, then a `categories` table containing only root parent rows, then a `keynotes` table containing every non-root keynote row. Sub-groups are normal rows in the `keynotes` table; child rows use the sub-group key as their parent key.

Blank text is allowed when the tab-delimited columns are still present. For example, a child row with no text should be written as `Key<TAB><TAB>ParentKey`.

## Saving and Refreshing

Click `Save` to merge your edits into the shared keynote file and reload Revit's keynote table.

Save may be blocked if the manager finds errors such as duplicate keys, empty keys, missing parents, malformed source lines, or a file access problem. Open `Warnings` and fix the listed items before saving again.

Click `Refresh` to reload the keynote file from disk. Refreshing discards unsaved edits in the manager, so save first if you want to keep your changes.

Click `Close` to close the manager. If you have unsaved edits, the manager will ask you to confirm before discarding them.

## Placing Keynotes

Use the arrow button in a keynote row to place that keynote in the active Revit view.

Before placing, make sure:

- The row has been saved.
- The correct Revit document and view are active.
- `Place As` is set to the placement type you want.

For `User Keynote`, Revit starts standard keynote placement for the selected key.

For `Generic Annotation`, the manager prepares the matching generic annotation keynote type and starts placement in the active view. If required content is missing, the manager will show a warning or error.

## Collecting Analytics

Click `Collect Analytics` to scan the active Revit document for keynote placement information. This updates the placed-keynote markers and helps the `Placed Keynotes` and `Unused Keynotes` filters show useful results.

Use this before reviewing unused keynotes, especially after other users have added or removed keynote annotations.

Supabase stores only the latest collected analytics for each keynote library and Revit document. Collecting again updates the existing document summary and keynote rows in place; it does not create a historical analytics run.

## Warnings and Common Issues

The manager validates the keynote file before saving. Common warnings and errors include:

- Duplicate keynote keys.
- Empty keys.
- Parent keys that do not exist.
- Parent/child cycles.
- Keynote file missing or unavailable.
- Shared-file or sync conflicts.
- Rows being edited by another user.

Errors must be fixed before saving. Warnings may not always block saving, but they should be reviewed.

## Safe Mode

Safe Mode pauses editing and saving when the model health scan finds a significant number of placed keynote keys that are missing from the active keynote file. Open `Model Issues` to review the affected keynotes.

For a placed Generic Annotation key that is missing from the text file, choose `Use Family Type` to add its key and text to the file. Choose a replacement keynote and `Use Text File` to overwrite the family type and migrate its placed instances to that file entry. Resolve every available missing-key choice, unlock editing, and click `Save` to apply the selections.

## Best Practices

- Save before placing keynotes you have just edited.
- Refresh before starting work if you know other users have been editing the same keynote file.
- Use `Collect Analytics` before relying on placed or unused keynote filters.
- Avoid deleting parent keynotes until child keynotes have been moved or deleted.
- Keep keynote keys consistent with the project's established numbering system.
- Review the `Warnings` panel before saving, especially after large edits.

## Quick Workflow

1. Open the Keynote Manager from the FFE-pyRevit ribbon.
2. Confirm the correct Revit document is loaded.
3. Select a division or use `Search`.
4. Use `Show` if you only want placed or unused keynotes.
5. Edit, add, duplicate, or delete keynote rows as needed.
6. Review `Warnings`.
7. Click `Save`.
8. Set `Place As`, then use the row arrow button to place keynotes in Revit.
