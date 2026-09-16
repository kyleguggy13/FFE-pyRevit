# FFE Keynote Manager User Guide

The FFE Keynote Manager reviews, edits, saves, and places keynotes from either the assigned text file or a Supabase library. Generic Annotation Only mode supports keynotes without a text file.

Both storage modes use the same Key, Text, and Parent structure and the same editing tools. The text file is authoritative in Text File mode. Supabase is authoritative in Generic Annotation Only mode.

## Opening the Manager

Open the tool from the FFE-pyRevit ribbon. The manager opens in its own window and loads the keynote file assigned to the active Revit document.

At the top of the window, confirm that the correct Revit document is listed. If the wrong document is active, switch back to Revit, activate the correct model, and reopen or refresh the manager.

## Choosing Storage

Open **Settings** at the top of the manager to choose **Text File** or **Generic Annotation Only** under **Storage Mode**. The popup also contains **Create Text File**, **Supabase Project URL**, and **Supabase Publishable Key**. Edit the connection fields and click **Save Connection** to save them in your user settings and reload the library. Legacy anon keys remain supported. Closing Settings without saving leaves the connection unchanged. Existing projects default to Text File. The mode choice is remembered in your local user settings; no keynote library or library identity is written to RVT storage.

- **Text File** uses the assigned text file with Supabase as a mirror. **Create Text File** downloads the 21-division template, asks for a new filename, writes UTF-16 text, and assigns it to Revit.
- **Generic Annotation Only** uses **Supabase as the source of truth**. On first setup, the cloud library is seeded from the valid assigned file, or from the Supabase division template. Existing cloud libraries are never reseeded from a file or family types. The editor retains divisions, subnotes, blank descriptions, and unplaced notes.

Save the Revit project before setup. Worksharing locals use the central/cloud document path to identify the same Supabase library. A different project path (including Save As to a new project) identifies a separate library. Only the mode preference is stored locally.

Switching from **Generic Annotation Only** to **Text File** opens Save As and exports the entire saved Supabase library, including divisions, nested notes, blank descriptions, and unplaced notes. **Export Library to Text File** in Settings performs the same conversion. The manager writes a new UTF-16 file, checks Revit assignment, and converts the **existing Supabase library record** to Text File mode with the file path, hash, encoding, and timestamp. Its library ID, keynote row IDs, claims, and analytics are retained. No second library record is created. Canceling Save As changes nothing. Existing files are not overwritten.

Save pending edits before converting; switching sources asks before discarding unsaved edits. A concurrent library change blocks conversion. If Supabase conversion cannot be confirmed, the Revit assignment rolls back and the mode stays unchanged; the exported file remains on disk. Refresh before retrying because an interrupted response may have followed a successful database update. Returning to Generic Annotation Only uses the same record through its retained project identity. Both modes can update matching FFE family types, so review Model Issues when switching.

### Saving and placing Generic Annotations

**Save** commits directly to Supabase using database version and edit-claim checks. Other users can load saved notes immediately, without Revit Sync/Reload Latest. Concurrent changes produce a conflict and retain your unsaved edits until you refresh or resolve them.

After the database save, the manager updates matching Revit family types. If a type update fails, the Supabase library remains saved; failed Revit updates roll back and the manager reports the failure separately. Use **Update Family Types** to retry. Revit Save/Sync is still needed to persist/share placed annotations and family changes, but never to persist the keynote library.

Placement and analytics fetch current notes from Supabase. User Keynote placement is disabled in Generic Annotation Only mode, and native keynote tables/references are not modified. Missing family notes can be adopted through Model Issues; choose **Use Supabase Library** to use the saved cloud description. Existing placed instances are preserved when their library row is deleted.

A Supabase connection is required to load, save, or place notes in Generic Annotation Only mode. There is no RVT or text-file fallback. If a save cannot be confirmed, edits remain in the editor; Refresh checks the server before retrying. Text File mode retains its existing behavior.

## Deployment (v1.3)

Apply `supabase/keynote_annotation_library.sql` after the existing `supabase/keynote_manager.sql` schema and before releasing this client. It uploads the read-only division template and adds the canonical annotation-library APIs. Run `supabase/test_keynote_annotation_library.sql` for database acceptance tests; fixtures are rolled back.

The previous draft `keynote_model_storage.sql` is superseded and must not be deployed. This implementation never reads or writes its RVT library storage. The migration removes its obsolete mirror APIs if they were installed, without deleting any stored library data.

See `tests/keynote_annotation_acceptance.md` for Revit 2025/2026 checks.

## TODO:
- Add realtime updates when a keynote is added or removed.
- Figure out how to work around Keynote's workset.
- When user 1 deletes note and saves it deletes the note for user 2 without refresh.
- If user 1 is editing note prevent other users from placing that note until it's released.

### Future Features:
- Tagging/Reference System
    - Connect Sheet Numbers, Note Numbers, View ID Numbers to automatically update in the keynotes.


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

Keynote keys show a green `Nx` badge when they are placed in the active model. If the same library is tracked in other Revit models, a blue `NM` badge marks keys placed in those other models; hover over that badge to see the model names and placement counts. Other-model usage is informational and does not change the `Placed Keynotes` or `Unused Keynotes` filters.

Use `Place As` to choose how the place button in the keynote table behaves:

- `User Keynote` places a standard Revit user keynote.
- `Generic Annotation` places the keynote as an FFE generic annotation keynote symbol.

## Editing Keynotes

Select a division, then edit rows directly in the keynote table.

Each row has a keynote key and description. Child keynotes are shown in a tree under their parent. Use the expand and collapse controls in the key column to show or hide child rows.

Use the ellipsis button in a keynote row for actions that apply directly to that row:

- `Copy Text` copies the keynote description to the clipboard.
- `Delete Note` removes the note when it has no child keynotes.
- `Move Note to Division` reparents the note under the selected top-level division. Any children of the moved note remain attached to it.
- `Promote to Parent` makes the note a new top-level parent. Any children of the promoted note remain attached to it.
- `Add Note in Sequence` creates a new note under the same parent.
- `UPPER CASE` converts the keynote description to uppercase.

Each parent also has an ellipsis menu in the Divisions list and selected-parent header:

- `Copy Text` copies the parent description.
- `Demote to a Note` moves the parent under another top-level parent while retaining its subnotes.
- `Delete Parent` asks whether its complete subnote tree should also be deleted or moved to another parent. Moving reparents the direct subnotes so their nested descendants remain attached.

The selected division header shows the current division key and description. Division descriptions can be edited there. Division keys are protected during normal editing; if key editing is enabled in your version, use extra care because changing keys can affect existing placed references.

Edits are not written to the selected storage until you click `Save`.

## Keynote Text File Format

The shared keynote file is tab-delimited text. Data rows use either `Key<TAB>Text` for parent rows or `Key<TAB>Text<TAB>ParentKey` for child rows.

Rows whose first non-space character is `#` are treated as comments. Revit does not load those rows as keynotes, and the manager ignores them while reading the file.

On save, the manager writes a metadata comment, then a `categories` table containing only root parent rows, then a `keynotes` table containing every non-root keynote row. Sub-groups are normal rows in the `keynotes` table; child rows use the sub-group key as their parent key.

Blank text is allowed when the tab-delimited columns are still present. For example, a child row with no text should be written as `Key<TAB><TAB>ParentKey`.

## Saving and Refreshing

In Text File mode, click `Save` to merge your edits into the shared keynote file and reload Revit's keynote table. In Generic Annotation Only mode, Save writes directly to Supabase as described above.

Save may be blocked if the manager finds errors such as duplicate keys, empty keys, missing parents, malformed source lines, or a file access problem. Open `Warnings` and fix the listed items before saving again.

Click `Refresh` to reload the selected source. Refreshing discards unsaved edits in the manager, so save first if you want to keep your changes.

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

The manager automatically scans the active Revit document for keynote placement information when it opens and syncs the results to Supabase. This updates the placed-keynote markers and helps the `Placed Keynotes` and `Unused Keynotes` filters show useful results.

Click `Collect Analytics` to run the scan again, especially after other users have added or removed keynote annotations while the manager is open.

Supabase stores only the latest collected analytics for each keynote library and Revit document. Collecting again updates the existing document summary and keynote rows in place; it does not create a historical analytics run. After syncing, the manager also compares the library's other document snapshots and marks keys placed in those models.

## Warnings and Common Issues

The manager validates the selected library before saving. Common warnings and errors include:

- Duplicate keynote keys.
- Empty keys.
- Parent keys that do not exist.
- Parent/child cycles.
- Keynote file missing or unavailable.
- Shared-file or sync conflicts.
- Rows being edited by another user.

Errors must be fixed before saving. Warnings may not always block saving, but they should be reviewed.

## Safe Mode

Safe Mode pauses editing and saving when the model health scan finds a significant number of placed keynote keys that are missing from the selected library. Open `Model Issues` to review the affected keynotes. In Generic Annotation Only mode, these checks cover Generic Annotations, and `Use Supabase Library` replaces `Use Text File` in the choices below.

For a placed Generic Annotation key that is missing from the text file, choose `Use Family Type` to add its key and text to the file. Choose a replacement keynote and `Use Text File` to overwrite the family type and migrate its placed instances to that file entry. Resolve every available missing-key choice, unlock editing, and click `Save` to apply the selections.

When a Generic Annotation family type and its matching text-file keynote have different descriptions, the Model Issues card offers three choices. `Use Family Type` updates the existing file row, `Use Text File` updates the family type, and `Keep Both + New Note` creates the next keynote in sequence from the family description while preserving the existing text-file row. Click `Save` to apply the selected resolution to the file and model.

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
