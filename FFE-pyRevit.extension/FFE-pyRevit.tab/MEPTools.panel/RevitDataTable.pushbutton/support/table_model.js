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

  const model = { compareNatural, compareNumeric, compareLevel, matches, reconcileFilters, selectedVisibleIds };
  if (typeof module !== "undefined" && module.exports) module.exports = model;
  else scope.RevitTableModel = model;
})(typeof globalThis === "undefined" ? this : globalThis);
