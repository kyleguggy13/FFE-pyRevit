/* Model text only reaches textContent/DOM nodes, never HTML or executable scripts. */
(function () {
  "use strict";
  const model = window.RevitTableModel;
  const byId = id => document.getElementById(id);
  const state = { snapshot: null, filters: null, active: true, outdated: false,
    busy: false, loading: false, table: null, focusId: null };
  const bridge = window.chrome?.webview;

  function status(message, kind = "info") {
    byId("status").textContent = message;
    byId("status").dataset.kind = kind;
  }

  function selectedIds() {
    return model.selectedVisibleIds(state.snapshot?.rows || [],
      state.table?.getSelectedData().map(row => row.id) || [], state.filters);
  }

  function updateControls() {
    const loaded = Boolean(state.snapshot) && !state.loading;
    const canRequest = loaded && state.active && !state.busy && Boolean(bridge);
    byId("refresh").disabled = !canRequest;
    byId("refresh").textContent = state.busy ? "Waiting…" : "Refresh";
    for (const id of ["search", "phase", "level", "clear-filters"]) byId(id).disabled = !loaded;
    const count = loaded ? selectedIds().length : 0;
    byId("select-revit").disabled = !canRequest || !count;
    byId("show-revit").disabled = !canRequest || !count;
    if (loaded) {
      const visible = state.table.getDataCount("active");
      byId("counts").textContent = `${visible} of ${state.snapshot.rows.length} Spaces · ${count} selected`;
    }
    const messages = [];
    if (!state.active) messages.push("Activate this table's model in Revit to refresh, select, or show Spaces.");
    if (state.outdated) messages.push("Out of date — the model has changed. Click Refresh to load current data.");
    byId("context").textContent = messages.join(" ");
    byId("context").hidden = !messages.length;
  }

  function request(action) {
    if (!bridge || state.busy || state.loading || !state.active || !state.snapshot) return;
    const message = { action, documentKey: state.snapshot.documentKey };
    if (action !== "refresh") {
      message.elementIds = selectedIds();
      if (!message.elementIds.length) return;
    }
    state.busy = true;
    status("Waiting for Revit…");
    updateControls();
    try {
      bridge.postMessage(message);
    } catch (error) {
      state.busy = false;
      status(`Could not contact Revit: ${error.message}`, "error");
      updateControls();
    }
  }

  function textFormatter(cell) {
    const node = document.createElement("span");
    node.textContent = cell.getValue() ?? "—";
    return node;
  }

  function measurementFormatter(textField) {
    return cell => {
      const node = document.createElement("span");
      node.textContent = cell.getRow().getData()[textField];
      return node;
    };
  }

  function levelSorter(a, b, aRow, bRow) {
    return model.compareLevel(aRow.getData(), bRow.getData());
  }

  function makeGrid() {
    const table = new Tabulator("#grid", {
      index: "id", height: "100%", layout: "fitColumns", rowHeight: 38,
      movableColumns: true, selectableRows: true, selectableRowsRangeMode: "click",
      placeholder: "No Spaces match the current filters.",
      columnDefaults: { minWidth: 110, formatter: textFormatter, resizable: true,
        headerSortTristate: true },
      initialSort: [{ column: "level", dir: "asc" }, { column: "number", dir: "asc" }],
      columns: [
        { title: "Level", field: "level", sorter: levelSorter, widthGrow: 1.3 },
        { title: "Number", field: "number", sorter: model.compareNatural, widthGrow: 1 },
        { title: "Name", field: "name", sorter: model.compareNatural, widthGrow: 2.5 },
        { title: "Area", field: "areaValue", sorter: "number", hozAlign: "right",
          sorterParams: { alignEmptyValues: "bottom" }, formatter: measurementFormatter("areaText"),
          headerTooltip: "Space area, formatted using project units" },
        { title: "Height", field: "heightValue", sorter: "number", hozAlign: "right",
          sorterParams: { alignEmptyValues: "bottom" }, formatter: measurementFormatter("heightText"),
          headerTooltip: "Space.UnboundedHeight, formatted using project units" },
        { title: "Phase", field: "phase", sorter: model.compareNatural, visible: false },
      ],
      rowFormatter(row) {
        const element = row.getElement();
        const data = row.getData();
        element.classList.toggle("space-warning", Boolean(data.warnings.length));
        element.classList.toggle("keyboard-focus", data.id === state.focusId);
        element.id = `space-row-${data.id}`;
        element.title = data.warnings.join("\n");
        element.setAttribute("aria-selected", String(row.isSelected()));
      },
    });
    table.on("rowSelectionChanged", (data, rows, selected, deselected) => {
      for (const row of [...selected, ...deselected]) {
        row.getElement().setAttribute("aria-selected", String(row.isSelected()));
      }
      updateControls();
    });
    table.on("dataFiltered", updateControls);
    table.on("rowClick", (event, row) => { setFocus(row.getData().id); byId("grid").focus(); });
    state.table = table;
    return new Promise(resolve => table.on("tableBuilt", resolve));
  }

  function setOptions(select, options, firstLabel, value) {
    const nodes = [{ id: "all", name: firstLabel }, ...options].map(item => {
      const option = document.createElement("option");
      option.value = item.id;
      option.textContent = item.name;
      return option;
    });
    select.replaceChildren(...nodes);
    select.value = value;
  }

  function applyFilters() {
    const filters = state.filters;
    state.table.setFilter(row => model.matches(row, filters));
    const selected = state.table.getSelectedData().map(row => row.id);
    const visible = new Set(model.selectedVisibleIds(state.snapshot.rows, selected, filters));
    const hidden = selected.filter(id => !visible.has(id));
    if (hidden.length) state.table.deselectRow(hidden);
    if (filters.phase === "all") state.table.showColumn("phase");
    else state.table.hideColumn("phase");
    if (state.focusId && !state.table.getRows("active").some(row => row.getData().id === state.focusId)) {
      setFocus(null);
    }
    updateControls();
  }

  function renderIssues(snapshot) {
    const warnings = snapshot.rows.filter(row => row.warnings.length);
    const count = warnings.length + snapshot.skippedCount;
    byId("issues").hidden = !count;
    byId("issues-summary").textContent = `${warnings.length} Space(s) with warnings · ${snapshot.skippedCount} skipped`;
    const messages = [...snapshot.issues, ...warnings.map(row =>
      `${row.number || "(no number)"} — ${row.name}: ${row.warnings.join(" ")}`)];
    const items = messages.slice(0, 100).map(message => {
      const item = document.createElement("li");
      item.textContent = message;
      return item;
    });
    if (messages.length > 100) {
      const item = document.createElement("li");
      item.textContent = `${messages.length - 100} additional warnings. Narrow the table to inspect individual rows.`;
      items.push(item);
    }
    byId("issues-list").replaceChildren(...items);
  }

  async function loadSnapshot(snapshot) {
    state.loading = true;
    updateControls();
    try {
      const sameDocument = state.snapshot?.documentKey === snapshot.documentKey;
      const selected = sameDocument ? state.table.getSelectedData().map(row => row.id) : [];
      state.filters = model.reconcileFilters(snapshot, sameDocument ? state.filters : null);
      state.snapshot = snapshot;
      byId("document-title").textContent = snapshot.documentTitle;
      setOptions(byId("phase"), snapshot.phases, "All phases", state.filters.phase);
      const levels = [...snapshot.levels].sort((a, b) =>
        model.compareNumeric(a.elevation, b.elevation) || model.compareNatural(a.name, b.name));
      setOptions(byId("level"), levels, "All levels", state.filters.level);
      byId("search").value = state.filters.query;
      await state.table.replaceData(snapshot.rows);
      applyFilters();
      const valid = model.selectedVisibleIds(snapshot.rows, selected, state.filters);
      if (valid.length) state.table.selectRow(valid);
      renderIssues(snapshot);
      const refreshed = new Date(snapshot.refreshedAt);
      byId("refreshed").textContent = `Refreshed ${refreshed.toLocaleTimeString()}`;
      byId("refreshed").title = refreshed.toLocaleString();
      if (!state.busy) status(`Loaded ${snapshot.rows.length} Space(s). Skipped ${snapshot.skippedCount}.`,
        snapshot.skippedCount ? "warning" : "success");
    } finally {
      state.loading = false;
      updateControls();
    }
  }

  function setFocus(id) {
    const previous = state.focusId && state.table.getRow(state.focusId);
    if (previous) previous.getElement().classList.remove("keyboard-focus");
    state.focusId = id;
    const current = id && state.table.getRow(id);
    if (current) {
      const element = current.getElement();
      element.classList.add("keyboard-focus");
      element.id = `space-row-${id}`;
      byId("grid").setAttribute("aria-activedescendant", element.id);
    } else byId("grid").removeAttribute("aria-activedescendant");
  }

  function gridKeydown(event) {
    if (event.target !== byId("grid") || !state.snapshot || state.loading) return;
    if (!["ArrowUp", "ArrowDown", " ", "Home", "End"].includes(event.key)) return;
    const rows = state.table.getRows("active");
    if (!rows.length) return;
    event.preventDefault();
    let index = rows.findIndex(row => row.getData().id === state.focusId);
    if (event.key === " ") {
      if (index < 0) index = 0;
      rows[index].toggleSelect();
    } else if (event.key === "Home") index = 0;
    else if (event.key === "End") index = rows.length - 1;
    else index = Math.min(rows.length - 1, Math.max(0, index + (event.key === "ArrowDown" ? 1 : -1)));
    setFocus(rows[index].getData().id);
    state.table.scrollToRow(rows[index], "nearest", false).catch(() => {});
  }

  async function initialize() {
    try {
      await makeGrid();
      byId("refresh").addEventListener("click", () => request("refresh"));
      byId("select-revit").addEventListener("click", () => request("select"));
      byId("show-revit").addEventListener("click", () => request("show"));
      byId("grid").addEventListener("keydown", gridKeydown);
      byId("search").addEventListener("input", () => {
        state.filters.query = byId("search").value; applyFilters();
      });
      for (const field of ["phase", "level"]) byId(field).addEventListener("change", () => {
        state.filters[field] = byId(field).value; applyFilters();
      });
      byId("clear-filters").addEventListener("click", () => {
        state.filters = { phase: state.snapshot.defaultPhaseId, level: "all", query: "" };
        byId("phase").value = state.filters.phase;
        byId("level").value = "all";
        byId("search").value = "";
        applyFilters();
      });
      if (!bridge) {
        status("Open this tool from the FFE-pyRevit ribbon to connect to a Revit model.");
        return;
      }
      // Serialize snapshots and results; a slow grid render cannot reorder host messages.
      let messages = Promise.resolve();
      bridge.addEventListener("message", event => {
        messages = messages.then(async () => {
          const message = event.data;
          if (message.type === "snapshot") await loadSnapshot(message.snapshot);
          else if (message.type === "context") {
            state.active = message.active;
            state.outdated = message.outdated;
            updateControls();
          } else if (message.type === "result") {
            state.busy = false;
            status(message.message, message.kind);
            updateControls();
          }
        }).catch(error => {
          state.busy = false;
          status(`Could not update the table: ${error.message}`, "error");
          updateControls();
        });
      });
      bridge.postMessage({ action: "ready" });
    } catch (error) {
      status(`Could not start the table: ${error.message}`, "error");
    }
  }

  initialize();
})();
