/* Pure table behavior shared by the offline UI and Node tests. */
(function (scope) {
  "use strict";
  const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: "base" });

  function compareNatural(a, b) {
    return collator.compare(String(a ?? ""), String(b ?? ""));
  }

  function compareNumeric(a, b) {
    if (a == null) return b == null ? 0 : 1;
    if (b == null) return -1;
    return a - b;
  }

  function compareLevel(a, b) {
    return compareNumeric(a.levelElevation, b.levelElevation) || compareNatural(a.level, b.level);
  }

  function matches(row, filters) {
    if (filters.phase !== "all" && row.phaseId !== filters.phase) return false;
    if (filters.level !== "all" && row.levelId !== filters.level) return false;
    const query = filters.query.trim().toLocaleLowerCase();
    return !query || [row.number, row.name, row.level, row.phase].some(
      value => String(value ?? "").toLocaleLowerCase().includes(query));
  }

  function reconcileFilters(snapshot, previous) {
    const phases = new Set(snapshot.phases.map(phase => phase.id));
    const levels = new Set(snapshot.levels.map(level => level.id));
    const phase = previous?.phase === "all" || phases.has(previous?.phase)
      ? previous.phase : snapshot.defaultPhaseId;
    return {
      phase: phase === "all" || phases.has(phase) ? phase : "all",
      level: levels.has(previous?.level) ? previous.level : "all",
      query: previous?.query || "",
    };
  }

  function selectedVisibleIds(rows, selectedIds, filters) {
    const ids = new Set(selectedIds);
    return rows.filter(row => ids.has(row.id) && matches(row, filters)).map(row => row.id);
  }

  function groupValue(row, field) {
    if (field === "level") return row.levelId ?? row.level;
    if (field === "phase") return row.phaseId ?? row.phase;
    if (field.startsWith("parameterValues.")) return row.parameterValues?.[field.slice("parameterValues.".length)];
    return row[field];
  }

  function groupKey(row, field) {
    const value = groupValue(row, field);
    // Typed keys preserve zero, leading zeros, exact IDs, and literal text without
    // colliding with missing values or object property names inside the grid.
    return JSON.stringify(value == null || value === "" ? ["missing"] : [typeof value, value]);
  }

  function groupLabel(row, field) {
    const value = groupValue(row, field);
    if (value == null || value === "") return "No value";
    let text = row[field];
    if (field === "areaValue") text = row.areaText;
    else if (field === "heightValue") text = row.heightText;
    else if (field.startsWith("parameterValues.")) text = row.parameterTexts?.[field.slice("parameterValues.".length)];
    return String(text ?? value);
  }

  /** Keep surviving header bands disjoint and arrange their members together. */
  function headerLayout(fields, groups) {
    const available = new Set(fields);
    const assigned = new Set();
    const bands = [];
    for (const group of groups) {
      const members = fields.filter(field => available.has(field) && group.fields.includes(field) && !assigned.has(field));
      members.forEach(field => assigned.add(field));
      if (members.length) bands.push({ name: group.name, fields: members });
    }
    const membership = new Map(bands.flatMap(group => group.fields.map(field => [field, group])));
    const emitted = new Set();
    const ordered = [];
    const orderedBands = [];
    for (const field of fields) {
      const band = membership.get(field);
      if (!band) ordered.push(field);
      else if (!emitted.has(band)) {
        emitted.add(band);
        ordered.push(...band.fields);
        orderedBands.push(band);
      }
    }
    return { fields: ordered, groups: orderedBands };
  }

  const model = { compareNatural, compareNumeric, compareLevel, matches, reconcileFilters,
    selectedVisibleIds, groupKey, groupLabel, headerLayout };
  if (typeof module !== "undefined" && module.exports) module.exports = model;
  else scope.RevitTableModel = model;
})(typeof globalThis === "undefined" ? this : globalThis);
