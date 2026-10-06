# FFE Keynote Manager User Guide

The FFE Keynote Manager reviews, edits, saves, and places keynotes from either the assigned text file or a Supabase library. Generic Annotation Only mode supports keynotes without a text file.

Both storage modes use the same Key, Text, and Parent structure and the same editing tools. The text file is authoritative in Text File mode. Supabase is authoritative in Generic Annotation Only mode.

## Opening the Manager

Open the tool from the FFE-pyRevit ribbon. The manager opens in its own window and loads the keynote source for the active Revit document. Projects without an assigned keynote reference or an existing cloud library open on **Set up project keynotes**.

Click **About** beside **Settings** to view the manager version, Revit document name, keynote source path, encoding/storage information, and entry count. The change indicator stays in the main header. If the wrong document is listed, switch back to Revit, activate the correct model, and reopen or refresh the manager.

### New Project Setup

Save the Revit project first, then click **Refresh** if the setup page asks you to save. Choose a keynote type and click **Continue**:

- **Revit Keynotes** downloads the FFE division template and opens Save As with `RevitKeynotes.txt` as the suggested name. Choose your project's shared file location. The manager creates a UTF-16 text file, assigns it to Revit's keynote table, and opens the editor after loading succeeds. Existing files are not overwritten.
- **Generic Annotation Keynotes** creates the project library in Supabase from the FFE division template and opens the editor with generic annotation placement selected. Setup does not create a text file or prepare a Revit family; existing placement tools handle the family requirements.

Settings, Refresh, and Close remain available on the setup page. Canceling Save As or a failed setup keeps your choice so you can retry. A Supabase connection is required for both setup choices because the division template is stored there. If the cloud library check fails, open **Settings** to check the connection, then **Refresh**; creation is disabled until the manager confirms that no library exists.

An existing library is recovered by its model-stored Supabase library ID, without replacing its notes. Models without an association first use the legacy file or central/cloud project path, then adopt the verified library ID. An explicit storage preference is preserved. An assigned reference that is missing, inaccessible, malformed, or unsupported stays in the existing error view instead of starting new-project setup.

### Recovering a Library or File Connection

When the library fails to load, the status bar offers recovery actions. A missing assigned text file shows **Reconnect Keynote File** and hides **Set Up Project**.

- **Set Up Project** opens the setup choices directly, even when automatic reference detection fails. Save the model before continuing. Revit Keynotes creates and assigns a new file; if a project cloud library exists, its saved notes are exported into that file instead of being replaced by the division template. Generic Annotation Keynotes reuses the existing project cloud library or initializes it when none exists. Supabase must be available for these setup operations. **Back to Manager** returns to the previous view without changing the source.
- **Reconnect Keynote File** opens a picker for an existing `.txt` file, validates it, assigns it to Revit, and loads it in Text File mode. The selected file is not rewritten. When Supabase is available, a bound library keeps its ID and records the new active file path; a path associated with another library is rejected. Offline reconnection retains the association for later reconciliation. If assignment or loading fails, the previous assignment is restored. Cancellation retains the current source and unsaved edits.

Reconnect is also available on the setup page. Recovery actions remain available after failures and are hidden after the source loads successfully. Save/synchronize the RVT to persist the new keynote assignment.

## Choosing Storage

Open **Settings** at the top of the manager to choose **Text File** or **Generic Annotation Only** under **Storage Mode**, then click **Apply** beside the dropdown. Apply is disabled until you select a different mode. Closing Settings discards an unapplied choice. The popup also contains **Create Text File**, **Supabase Project URL**, and **Supabase Publishable Key**. During new-project setup, choose the keynote type on the setup page; the storage controls in Settings stay disabled. Edit the connection fields and click **Save Connection** to save them in your user settings and reload the library. Legacy anon keys remain supported. Closing Settings without saving leaves the connection unchanged. Projects with assigned keynote references default to Text File unless you have selected another mode. The mode choice is remembered in your local user settings and verified library association.

- **Text File** uses the assigned text file with Supabase as a mirror. **Create Text File** downloads the 21-division template, asks for a new filename, writes UTF-16 text, and assigns it to Revit.
- **Generic Annotation Only** uses **Supabase as the source of truth**. Switching from Text File converts its existing Supabase mirror in place, retaining the library ID, surviving keynote row IDs, claims, and analytics. The original file key remains associated with the record, and the Revit project identity becomes its annotation alias. Current file edits are imported during conversion; subsequent cloud edits are preserved on retries. If the file has never been mirrored, setup initializes the library from its valid contents. Projects without a file use the Supabase division template. The editor retains divisions, subnotes, blank descriptions, and unplaced notes.

Save the Revit project before setup. Worksharing locals share the central/cloud document identity and the model-stored library ID. A changed model path opens a choice to **Keep existing library**, **Create independent library**, or cancel. Keep the existing library for a renamed or moved project. An independent copy receives a new library ID and keynote row IDs; its source library, claims, and analytics remain unchanged. Text File copies ask for a new file location and keep Text File mode. Generic Annotation copies use the saved cloud notes.

### Library identity and recovery

The manager stores the Supabase library UUID, project URL, storage mode, and reconciliation paths in a dedicated Revit Extensible Storage entity. It stores no keynote contents, database versions, or connection keys there. The schema GUID identifies the metadata format; each entity's library UUID identifies its Supabase record. Save/synchronize Revit after the association is created or changed so other users and later sessions can recover it.

Library ID lookup takes precedence over names and paths. Reconnecting a moved or renamed keynote file keeps its library ID, note IDs, claims, and analytics. Historical file paths remain aliases. Revit still needs the physical file reconnected; a model using an older file copy must reconnect to the active path before it can overwrite the shared mirror.

Under **Settings > Library Association**, **Change Library...** selects an existing Supabase library and repairs a missing or incorrect association. **Create Independent Library...** copies the current library into a separate record. A missing stored UUID or a different Supabase project blocks automatic fallback instead of creating a replacement library. If Revit ownership prevents saving the association, the manager reports it separately; resolve ownership, Refresh, and Save/Sync.

Switching from **Generic Annotation Only** to **Text File** opens Save As and exports the entire saved Supabase library, including divisions, nested notes, blank descriptions, and unplaced notes. **Export Library to Text File** in Settings performs the same conversion. The manager writes a new UTF-16 file, checks Revit assignment, and converts the **existing Supabase library record** to Text File mode with the file path, hash, encoding, and timestamp. Its library ID, keynote row IDs, claims, and analytics are retained. No second library record is created. Canceling Save As changes nothing. Existing files are not overwritten.

Save pending edits before converting; switching sources asks before discarding unsaved edits. A concurrent library change blocks conversion. If Supabase conversion cannot be confirmed, the Revit assignment rolls back and the mode stays unchanged; the exported file remains on disk. Refresh before retrying because an interrupted response may have followed a successful database update. Returning to Generic Annotation Only uses the same record through its retained project identity. Both modes can update matching FFE family types, so review Model Issues when switching.

### Saving and placing Generic Annotations

**Save** commits directly to Supabase using database version and edit-claim checks. Other users can load saved notes immediately, without Revit Sync/Reload Latest. Concurrent changes produce a conflict and retain your unsaved edits until you refresh or resolve them.

After the database save, the manager updates matching Revit family types. If a type update fails, the Supabase library remains saved; failed Revit updates roll back and the manager reports the failure separately. Use **Update Family Types** to retry. Revit Save/Sync persists/shares the association, placed annotations, and family changes. Note contents are saved directly in Supabase.

Placement and analytics fetch current notes from Supabase. User Keynote placement is disabled in Generic Annotation Only mode, and native keynote tables/references are not modified. Missing family notes can be adopted through Model Issues; choose **Use Supabase Library** to use the saved cloud description. Existing placed instances are preserved when their library row is deleted.

A Supabase connection is required to load, save, or place notes in Generic Annotation Only mode. There is no RVT or text-file fallback. If a save cannot be confirmed, edits remain in the editor; Refresh checks the server before retrying. Text File mode retains its existing behavior.

## Deployment (v1.3)

Apply the migrations in this order before releasing this client: `supabase/keynote_manager.sql`, `supabase/keynote_annotation_library.sql`, then `supabase/keynote_library_identity.sql`. The identity migration adds UUID lookup, protected path aliases, active-file checks, and independent-library creation without replacing existing records. Run both `supabase/test_keynote_annotation_library.sql` and `supabase/test_keynote_library_identity.sql`; fixtures are rolled back.

For existing installations, apply the updated annotation migration followed by the identity migration. Reapply the identity migration last whenever older migrations are rerun, because it updates shared snapshot APIs. File-to-annotation conversion changes the existing file library. An unreadable assigned file must be reconnected before conversion; conflicting project associations stop conversion without replacing either library.

The previous draft `keynote_model_storage.sql` is superseded and must not be deployed. This implementation uses a separate identity-only schema and never reads or writes the draft's RVT note snapshots. The annotation migration removes its obsolete mirror APIs if they were installed, without deleting any stored library data.

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
