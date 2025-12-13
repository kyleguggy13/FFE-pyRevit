Understood. Option 1 is a strong pattern: **export a Revit “documentation + spatial + MEP connectivity” graph to JSON**, then render it in your Pixi viewer.

Below is a concrete, Revit-native way to build exactly the nodes/edges you listed, including pragmatic implementation choices (because “visible within crop” is the only part that gets mathematically fussy).

---

## 1) Data model (JSON schema)

This is the minimal schema that will keep your Pixi viewer simple and your Revit exporter deterministic.

```json
{
  "meta": {
    "sourceModel": "C:\\path\\model.rvt",
    "exportedAtUtc": "2025-12-13T16:30:00Z",
    "revitVersion": "2026"
  },
  "nodes": [
    {
      "key": "sheet:12345",
      "type": "sheet",
      "revit": { "elementId": 12345, "uniqueId": "..." },
      "label": "A101 - Floor Plan",
      "props": { "sheetNumber": "A101", "sheetName": "Floor Plan" },
      "pos": { "x": -300, "y": 0 }
    }
  ],
  "edges": [
    {
      "type": "sheet_to_view",
      "from": "sheet:12345",
      "to": "view:67890",
      "props": { "via": "Viewport" }
    }
  ]
}
```

### Node keys

Use stable, readable keys:

* `sheet:<ElementId>`
* `view:<ElementId>`
* `room:<ElementId>`
* `equip:<ElementId>`
* `system:<ElementId>`

This makes the viewer and debugging much easier.

---

## 2) Export logic in Revit (pyRevit / Revit API)

### A. Sheets → Views (placed on)

This is straightforward and reliable.

**Sources:**

* `Viewport` elements: `Viewport.SheetId`, `Viewport.ViewId`
* `ScheduleSheetInstance` if you want schedules too (optional)

**Edges:**

* `sheet_to_view` with `props.via = "Viewport"` (or “Schedule”)

---

### B. Views → Rooms (visible / within crop)

This is the “hard part.” The most robust practical approach is:

1. Only consider **views that have a meaningful crop**:

   * Floor Plans / Ceiling Plans / Area Plans
   * Sections / Elevations (optional, but more work)
2. For each view, compute a **2D crop polygon** in the view’s plane.
3. For each room, test its “representative point” against that polygon.

#### Crop polygon

Use:

* `var mgr = view.GetCropRegionShapeManager()`
* `mgr.GetCropShape()` → returns one or more `CurveLoop`s

Those curves are in model coordinates, but they lie on the view plane.

#### Room representative point

Use (in order of preference):

* Room location point (if present)
* Otherwise room bounding box center

Then:

* Project the 3D point into the view plane’s 2D coordinate system
* Use point-in-polygon

This gives a correct “inside crop” test for most plan views and typical workflows.

**Edge:**

* `view_to_room`

**Notes:**

* If a view has `CropBoxActive == false`, you can either:

  * skip view→rooms edges, or
  * treat the view as “uncropped” and link to all rooms on the view’s level (coarser but usable)

---

### C. Rooms → Equipment (contained)

This can be done cleanly using `Room.IsPointInRoom(point)`.

For each candidate equipment instance:

* Get a point:

  * `LocationPoint.Point`, or
  * midpoint of `LocationCurve`, or
  * bounding box center as fallback
* Test each room on that level: `room.IsPointInRoom(pt)`

**Edge:**

* `room_to_equip`

**Candidate equipment filtering**
To keep it “key equipment,” filter by:

* Categories: `OST_MechanicalEquipment`, `OST_PlumbingFixtures`, `OST_ElectricalEquipment`, etc.
* And/or a shared parameter like `IsKeyEquipment = Yes`
* And/or family name contains a whitelist

---

### D. Equipment → Systems (connected to)

This is best done via connectors.

For each equipment `FamilyInstance`:

* `fi.MEPModel?.ConnectorManager?.Connectors`
* For each connector, read:

  * `connector.MEPSystem` (if available) or
  * `connector.AllRefs` to traverse to connected system elements

Depending on discipline:

* Electrical: `ElectricalSystem` objects / circuits
* Mechanical piping/duct: `MEPSystem` (e.g., `MechanicalSystem`, `PipingSystem`)

**Edge:**

* `equip_to_system`

**System node**
Create a node for each distinct system:

* key: `system:<SystemId>` (use ElementId if it exists; for some system objects you may need a synthetic key)
* label: `System.Name` plus type/classification if available

---

## 3) A production-friendly coordinate strategy for the Pixi viewer

Instead of trying to spatially lay out everything in true XY coordinates (which will look cluttered), use a **layered “swimlane” layout**:

* X positions by type:

  * Sheets: `x = -600`
  * Views: `x = -200`
  * Rooms: `x = +200`
  * Equipment: `x = +600`
  * Systems: `x = +1000`
* Y positions: stack within each type (or by Level)

This produces a readable graph immediately. You can later add a force-layout in JS if you want.

---

## 4) What I recommend you implement first (fast, reliable)

**Phase 1 (1–2 hours, very doable)**

* Sheets nodes
* Views placed on sheets
* Rooms nodes (all rooms)
* Equipment nodes (filtered)
* Systems nodes (from equipment connectors)
* Edges:

  * sheet→view
  * room→equipment
  * equipment→system
* Optional: view→rooms as “rooms on same level as view” (fallback)

**Phase 2 (accuracy upgrade)**

* Replace fallback view→rooms with the crop-polygon test described above.

This avoids getting blocked on geometry nuances while still producing a useful graph.

---

## 5) If you want, I can give you the full pyRevit exporter script next

If you confirm these two implementation decisions, I will write the complete working pyRevit script (export JSON + a ready-to-open Pixi HTML viewer that reads it):

1. “Key equipment” filter:

   * Category whitelist only, or
   * Category + shared parameter (tell me the parameter name), or
   * Family/type name whitelist

2. Views included for view→room edges:

   * Plans only (recommended), or
   * Plans + sections/elevations

If you do not want to decide, I will default to:

* **Plans only**
* **Category whitelist** for key equipment

and proceed accordingly.
