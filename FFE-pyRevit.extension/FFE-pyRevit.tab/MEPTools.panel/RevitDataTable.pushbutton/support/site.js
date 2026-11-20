/* Model text only reaches textContent/DOM nodes, never HTML or executable scripts. */
(function () {
  "use strict";
  const model = window.RevitTableModel;
  const byId = id => document.getElementById(id);
  const state = { snapshot: null, filters: null, active: true, outdated: false,
    busy: false, loading: false, table: null, focusId: null,
    draftParameterIds: new Set(), columnMetadata: new Map(), columnColors: new Map(), groupFields: [], headerGroups: [] };
  const bridge = window.chrome?.webview;
  let savedTables = null;
  let columnGrouping = null;
  let headerColors = null;
  let cloudPublish = null;

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
    byId("parameters").disabled = !loaded || state.busy;
    byId("column-groups").disabled = !loaded || state.busy;
    byId("parameter-apply").disabled = !canRequest;
    byId("parameter-apply").title = state.active ? "Load the selected parameter columns" : "Activate this table's model in Revit to apply";
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
    if (savedTables) savedTables.update();
    if (columnGrouping) columnGrouping.update();
    if (headerColors) headerColors.update();
    if (cloudPublish) cloudPublish.update();
    if (state.table) updateSortPriorities();
  }

  // The grid is configured so the first selected sort is the primary key.
  function updateSortPriorities() {
    const sorters = state.table.getSorters().filter(sorter => ["asc", "desc"].includes(sorter.dir));
    const priorities = new Map(sorters.map((sorter, index) => [sorter.field, index + 1]));
    for (const column of state.table.getColumns()) {
      const header = column.getElement();
      const sorter = header.querySelector(".tabulator-col-sorter");
      const priority = sorters.length > 1 ? priorities.get(column.getField()) : null;
      const existing = sorter?.querySelector(".sort-priority");
      header.classList.toggle("multi-sort", Boolean(priority));
      if (!priority || !sorter) {
        existing?.remove();
        continue;
      }
      const number = existing || document.createElement("sub");
      number.className = "sort-priority";
      number.textContent = String(priority);
      number.setAttribute("aria-label", `Sort priority ${priority} of ${sorters.length}`);
      number.title = `Sort priority ${priority} of ${sorters.length}`;
      if (!existing) sorter.append(number);
    }
  }

  // Tabulator normally moves a changed multi-sort to the end. Keep its original slot.
  function preserveSortPriority(event) {
    if (!(event.ctrlKey || event.shiftKey) || event.button !== 0 ||
        event.target.closest(".tabulator-col-resize-handle")) return;
    const header = event.target.closest(".tabulator-col[tabulator-field]");
    if (!header) return;
    const field = header.getAttribute("tabulator-field");
    const sorters = state.table.getSorters().map(sorter => ({ column: sorter.field, dir: sorter.dir }));
    const index = sorters.findIndex(sorter => sorter.column === field);
    if (index < 0) return;
    const definition = state.table.getColumn(field).getDefinition();
    if (definition.headerSort === false) return;
    const starting = definition.headerSortStartingDir || "asc";
    const next = sorters[index].dir === starting ? (starting === "asc" ? "desc" : "asc")
      : definition.headerSortTristate ? "none" : starting;
    event.stopPropagation();
    if (next === "none") sorters.splice(index, 1);
    else sorters[index].dir = next;
    state.table.setSort(sorters);
  }

  function request(action, options = {}) {
    if (!bridge || state.busy || state.loading || !state.active || !state.snapshot) return;
    const message = { ...options, action, documentKey: state.snapshot.documentKey };
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

  function parameterColumn(parameter) {
    const numeric = ["Double", "Integer"].includes(parameter.storageType);
    return {
      title: parameter.name, field: `parameterValues.${parameter.key}`, width: 180,
      sorter: numeric ? "number" : model.compareNatural, hozAlign: numeric ? "right" : "left",
      sorterParams: numeric ? { alignEmptyValues: "bottom" } : {},
      titleFormatter() {
        const node = document.createElement("span");
        node.textContent = parameter.name;
        return node;
      },
      headerTooltip() {
        const node = document.createElement("span");
        node.textContent = `${parameter.name} · ${parameter.dataType} · ID ${parameter.id}`;
        return node;
      },
      formatter(cell) {
        const node = document.createElement("span");
        node.textContent = cell.getRow().getData().parameterTexts?.[parameter.key] ?? "—";
        return node;
      },
    };
  }

  async function syncParameterColumns(snapshot) {
    const selected = new Set(snapshot.selectedParameterIds || []);
    const parameters = (snapshot.parameters || []).filter(parameter => !parameter.baseField && selected.has(parameter.id));
    const fields = new Set(parameters.map(parameter => `parameterValues.${parameter.key}`));
    for (const column of state.table.getColumns()) {
      const field = column.getField();
      if (field.startsWith("parameterValues.") && !fields.has(field)) {
        await column.delete();
        state.columnMetadata.delete(field);
        state.columnColors.delete(field);
      }
    }
    for (const parameter of parameters) {
      const field = `parameterValues.${parameter.key}`;
      const signature = JSON.stringify([parameter.name, parameter.dataType, parameter.storageType]);
      const column = state.table.getColumns().find(candidate => candidate.getField() === field);
      if (!column) await state.table.addColumn(parameterColumn(parameter));
      else if (state.columnMetadata.get(field) !== signature) {
        await column.updateDefinition({ ...parameterColumn(parameter), width: column.getWidth() });
      }
      state.columnMetadata.set(field, signature);
    }
    const available = new Set(state.table.getColumns().map(column => column.getField()));
    const sorters = state.table.getSorters();
    const validSorters = sorters.filter(sorter => available.has(sorter.field));
    if (validSorters.length !== sorters.length) state.table.setSort(validSorters.length
      ? validSorters : [{ column: "level", dir: "asc" }, { column: "number", dir: "asc" }]);
  }

  function filteredParameters() {
    const query = byId("parameter-search").value.trim().toLocaleLowerCase();
    return (state.snapshot.parameters || []).filter(parameter =>
      (!byId("parameter-checked-only").checked || state.draftParameterIds.has(parameter.id)) &&
      `${parameter.name} ${parameter.dataType} ${parameter.parameterType || "Other"} ${parameter.id}`.toLocaleLowerCase().includes(query));
  }

  function updateParameterSelection() {
    const parameters = filteredParameters();
    const editable = parameters.filter(parameter => !parameter.baseField);
    const checked = editable.filter(parameter => state.draftParameterIds.has(parameter.id)).length;
    byId("parameters-all").checked = Boolean(editable.length) && checked === editable.length;
    byId("parameters-all").indeterminate = checked > 0 && checked < editable.length;
    byId("parameters-all").disabled = !editable.length;
    byId("parameter-count").textContent = `${parameters.length} of ${(state.snapshot.parameters || []).length} parameters · ${state.draftParameterIds.size} selected`;
  }

  function renderParameters() {
    const names = new Map();
    for (const parameter of state.snapshot.parameters || []) names.set(parameter.name, (names.get(parameter.name) || 0) + 1);
    const rows = filteredParameters().map(parameter => {
      const row = document.createElement("tr");
      const select = document.createElement("td");
      select.className = "checkbox-column";
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.dataset.parameterId = parameter.id;
      checkbox.checked = state.draftParameterIds.has(parameter.id);
      checkbox.disabled = Boolean(parameter.baseField);
      if (parameter.baseField) checkbox.title = parameter.baseField === "phase"
        ? "The Phase column is shown when All phases is selected."
        : "This starter column stays visible.";
      checkbox.setAttribute("aria-label", `${parameter.name}, ${parameter.dataType}, ID ${parameter.id}`);
      select.append(checkbox);
      const name = document.createElement("td");
      name.textContent = parameter.name;
      if (names.get(parameter.name) > 1) {
        const identity = document.createElement("small");
        identity.textContent = `ID ${parameter.id}`;
        name.append(identity);
      }
      const dataType = document.createElement("td");
      dataType.className = "parameter-data-type";
      dataType.textContent = parameter.dataType;
      const parameterType = document.createElement("td");
      parameterType.className = "parameter-origin";
      parameterType.textContent = parameter.parameterType || "Other";
      row.append(select, name, dataType, parameterType);
      return row;
    });
    byId("parameter-list").replaceChildren(...rows);
    byId("parameters-empty").hidden = Boolean(rows.length);
    byId("parameters-empty").textContent = byId("parameter-checked-only").checked
      ? "No checked parameters match your search." : "No parameters match your search.";
    updateParameterSelection();
  }

  function openParameters() {
    const visibleFields = new Set(state.table.getColumns().filter(column => column.isVisible())
      .map(column => column.getField()));
    state.draftParameterIds = new Set((state.snapshot.parameters || []).filter(parameter =>
      visibleFields.has(parameter.baseField || `parameterValues.${parameter.key}`)).map(parameter => parameter.id));
    byId("parameter-search").value = "";
    renderParameters();
    byId("parameter-dialog").showModal();
  }

  function applyGrouping() {
    state.table.setGroupBy(state.groupFields.length
      ? state.groupFields.map(field => row => model.groupKey(row, field)) : false);
    updateControls();
  }

  function selectGrouping(fields) {
    const visible = new Set(state.table.getColumns().filter(column => column.isVisible()).map(column => column.getField()));
    state.groupFields = [...new Set(fields)].filter(field => visible.has(field));
    syncGrouping();
    applyGrouping();
  }

  // Resolve chip labels directly from the visible columns, including duplicate names.
  function groupingCaption(field) {
    const columns = state.table.getColumns().filter(column => column.isVisible());
    const title = columns.find(column => column.getField() === field)?.getDefinition().title || field;
    const parameter = (state.snapshot?.parameters || []).find(item => `parameterValues.${item.key}` === field);
    return parameter && columns.filter(column => column.getDefinition().title === title).length > 1
      ? `${title} (ID ${parameter.id})` : title;
  }

  function syncGrouping() {
    const columns = state.table.getColumns().filter(column => column.isVisible());
    const previous = state.groupFields;
    state.groupFields = previous.filter(field => columns.some(column => column.getField() === field));
    if (previous.length !== state.groupFields.length) applyGrouping();
  }

  // Rebuild from live leaf definitions so dynamic parameters, widths and formatters survive.
  function rebuildHeaders(groups, orderedFields = null) {
    const columns = state.table.getColumns();
    const definitions = new Map(columns.map(column => [column.getField(), {
      ...column.getDefinition(), width: column.getWidth(), visible: column.isVisible(),
    }]));
    const layout = model.headerLayout(orderedFields || columns.map(column => column.getField()), groups);
    const sorters = state.table.getSorters().map(sorter => ({ column: sorter.field, dir: sorter.dir }));
    state.headerGroups = layout.groups;
    const membership = new Map(layout.groups.flatMap(group => group.fields.map(field => [field, group])));
    const emitted = new Set();
    const headers = [];
    for (const field of layout.fields) {
      const group = membership.get(field);
      if (!group) headers.push(definitions.get(field));
      else if (!emitted.has(group)) {
        emitted.add(group);
        headers.push({ title: group.name, headerHozAlign: "center", columns: group.fields.map(item => definitions.get(item)),
          titleFormatter() {
            const text = document.createElement("span");
            text.textContent = group.name;
            return text;
          },
        });
      }
    }
    state.table.setColumns(headers);
    state.table.setSort(sorters);
  }

  function setHeaderGroups(groups, orderedFields = null) {
    const loading = state.loading;
    state.loading = true;
    updateControls();
    try {
      state.table.setGroupBy(false);
      rebuildHeaders(groups, orderedFields);
      applyGrouping();
    } finally {
      state.loading = loading;
      updateControls();
    }
  }

  // Crossing a band boundary moves the leaf into the target's band (or out to a loose column).
  function moveHeaderColumn(source, target, after) {
    if (!state.headerGroups.length) {
      state.table.moveColumn(source, target, after);
      return;
    }
    const fields = state.table.getColumns().map(column => column.getField()).filter(field => field !== source);
    fields.splice(fields.indexOf(target) + (after ? 1 : 0), 0, source);
    const groups = state.headerGroups.map(group => ({ ...group,
      fields: group.fields.filter(field => field !== source),
    }));
    const targetGroup = groups.find(group => group.fields.includes(target));
    if (targetGroup) targetGroup.fields.push(source);
    setHeaderGroups(groups, fields);
  }

  // Move the complete band around whole target bands; memberships never change here.
  function moveHeaderBand(source, target, after) {
    const sourceGroup = state.headerGroups.find(group => group.fields.includes(source));
    const targetGroup = state.headerGroups.find(group => group.fields.includes(target));
    if (!sourceGroup || sourceGroup === targetGroup) return;
    const fields = state.table.getColumns().map(column => column.getField());
    const members = new Set(sourceGroup.fields);
    const moving = fields.filter(field => members.has(field));
    const remaining = fields.filter(field => !members.has(field));
    const targets = targetGroup ? targetGroup.fields : [target];
    const anchor = after ? targets[targets.length - 1] : targets[0];
    const position = remaining.indexOf(anchor);
    if (position < 0) return;
    remaining.splice(position + (after ? 1 : 0), 0, ...moving);
    setHeaderGroups(state.headerGroups, remaining);
  }

  function makeGrid() {
    const table = new Tabulator("#grid", {
      index: "id", height: "100%", layout: "fitColumns", rowHeight: 24,
      // Native header drag/drop handles grouping and reordering together; the
      // grid's separate mouse-based mover would compete with those gestures.
      movableColumns: false, selectableRows: true, selectableRowsRangeMode: "click",
      sortOrderReverse: true,
      groupStartOpen: true, groupToggleElement: "header",
      groupHeader(value, count, rows, group) {
        // Functions used for typed keys have no column field; parent depth identifies it.
        let depth = 0;
        for (let parent = group.getParentGroup(); parent; parent = parent.getParentGroup()) depth++;
        const field = state.groupFields[depth];
        const column = state.table.getColumns().find(item => item.getField() === field);
        const title = column?.getDefinition().title || field;
        const button = document.createElement("button");
        button.type = "button";
        button.className = "group-toggle";
        button.textContent = `${title}: ${model.groupLabel(rows[0] || {}, field)} · ${count} ${count === 1 ? "Space" : "Spaces"}`;
        button.setAttribute("aria-expanded", String(group.isVisible()));
        return button;
      },
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
    table.on("columnMoved", () => { if (state.snapshot && !state.loading) syncGrouping(); });
    table.on("columnMoved", updateControls);
    table.on("columnResized", updateControls);
    table.on("dataSorted", updateControls);
    table.on("groupVisibilityChanged", (group, visible) => {
      const element = group.getElement();
      const button = element && typeof element.querySelector === "function"
        ? element.querySelector(".group-toggle") : null;
      if (button) button.setAttribute("aria-expanded", String(visible));
    });
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
    syncGrouping();
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

  async function loadSnapshot(snapshot, layout = null) {
    state.loading = true;
    updateControls();
    try {
      const sameDocument = state.snapshot?.documentKey === snapshot.documentKey;
      if (layout || !sameDocument) state.columnColors = new Map(
        (layout?.columns || []).filter(column => column.color).map(column => [column.field, column.color]));
      const headerGroups = layout ? layout.headerGroups || [] : sameDocument ? state.headerGroups : [];
      state.table.setGroupBy(false);
      if (state.headerGroups.length) rebuildHeaders([]);
      if (!sameDocument && state.groupFields.length) {
        state.groupFields = [];
        state.table.setGroupBy(false);
      }
      const selected = sameDocument && !layout ? state.table.getSelectedData().map(row => row.id) : [];
      if (layout) {
        state.table.setGroupBy(false);
        state.groupFields = layout.groupFields || (layout.groupField ? [layout.groupField] : []);
      }
      state.filters = model.reconcileFilters(snapshot, layout ? layout.filters : sameDocument ? state.filters : null);
      state.snapshot = snapshot;
      setOptions(byId("phase"), snapshot.phases, "All phases", state.filters.phase);
      const levels = [...snapshot.levels].sort((a, b) =>
        model.compareNumeric(a.elevation, b.elevation) || model.compareNatural(a.name, b.name));
      setOptions(byId("level"), levels, "All levels", state.filters.level);
      byId("search").value = state.filters.query;
      await syncParameterColumns(snapshot);
      await state.table.replaceData(snapshot.rows);
      if (layout) {
        // Tabulator's persistence layout loader uses its initial definitions and
        // drops columns added later. Move/resize the live components instead.
        let previous = null;
        for (const column of layout.columns) {
          if (previous) state.table.moveColumn(column.field, previous, true);
          else {
            const first = state.table.getColumns()[0].getField();
            if (first !== column.field) state.table.moveColumn(column.field, first, false);
          }
          state.table.getColumn(column.field).setWidth(column.width);
          previous = column.field;
        }
        state.table.setSort(layout.sorters.map(sorter => ({ column: sorter.field, dir: sorter.dir })));
      }
      if (headerGroups.length) rebuildHeaders(headerGroups);
      applyFilters();
      applyGrouping();
      const valid = model.selectedVisibleIds(snapshot.rows, selected, state.filters);
      if (valid.length) state.table.selectRow(valid);
      renderIssues(snapshot);
      const refreshed = new Date(snapshot.refreshedAt);
      byId("refreshed").textContent = `Refreshed ${refreshed.toLocaleTimeString()}`;
      byId("refreshed").title = refreshed.toLocaleString();
      if (!state.busy) status(`Loaded ${snapshot.rows.length} Space(s). Skipped ${snapshot.skippedCount}.`,
        snapshot.skippedCount ? "warning" : "success");
      if (savedTables) savedTables.onSnapshot();
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
    const rows = state.table.getRows("active").filter(row => {
      for (let group = row.getGroup(); group; group = group.getParentGroup()) if (!group.isVisible()) return false;
      return true;
    });
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
      columnGrouping = new window.ColumnGroupingDrag({ table: state.table,
        isReady: () => Boolean(state.snapshot) && !state.loading,
        getGroups: () => state.groupFields, setGroups: selectGrouping,
        caption: groupingCaption,
        moveColumn: moveHeaderColumn,
        moveBand: moveHeaderBand,
      });
      state.table.on("tableDestroyed", () => columnGrouping.destroy());
      headerColors = new window.HeaderColorControls({ table: state.table,
        isReady: () => Boolean(state.snapshot) && !state.loading && !state.busy,
        getGroups: () => state.headerGroups, getColumnColors: () => state.columnColors,
        setColor(target, color) {
          if (target.group) {
            const group = state.headerGroups.find(group => group.fields.includes(target.field));
            if (!group) return;
            if (color) group.color = color;
            else delete group.color;
          } else if (color) state.columnColors.set(target.field, color);
          else state.columnColors.delete(target.field);
          updateControls();
        },
      });
      state.table.on("tableDestroyed", () => headerColors.destroy());
      new window.HeaderGroupControls({
        isReady: () => Boolean(state.snapshot) && !state.loading && !state.busy,
        getGroups: () => state.headerGroups, setGroups: setHeaderGroups,
        getColumns: () => state.table.getColumns().map(column => ({ field: column.getField(),
          title: groupingCaption(column.getField()), visible: column.isVisible(),
        })),
      });
      savedTables = new window.SavedTableControls({ bridge, status, loadSnapshot,
        getSnapshot: () => state.snapshot, getTable: () => state.table,
        isLoading: () => state.loading, isActive: () => state.active, isBusy: () => state.busy,
        setBusy: value => { state.busy = value; updateControls(); },
        getLayout: () => ({
          columns: state.table.getColumns().map(column => ({ field: column.getField(),
            width: Math.max(110, Math.min(3000, Math.round(column.getWidth() || column.getDefinition().width || 140))),
            ...(state.columnColors.has(column.getField()) ? { color: state.columnColors.get(column.getField()) } : {}) })),
          sorters: state.table.getSorters().map(sorter => ({ field: sorter.field, dir: sorter.dir })),
          groupFields: [...state.groupFields], headerGroups: state.headerGroups.map(group => ({ ...group, fields: [...group.fields] })),
          filters: { ...state.filters },
        }),
      });
      cloudPublish = new window.CloudPublishControls({ bridge, status,
        getSnapshot: () => state.snapshot, isLoading: () => state.loading,
        isActive: () => state.active, isBusy: () => state.busy,
        setBusy: value => { state.busy = value; updateControls(); },
      });
      byId("refresh").addEventListener("click", () => request("refresh"));
      byId("grid").addEventListener("click", preserveSortPriority, true);
      byId("parameters").addEventListener("click", openParameters);
      byId("parameter-search").addEventListener("input", renderParameters);
      byId("parameter-checked-only").addEventListener("change", renderParameters);
      byId("parameter-cancel").addEventListener("click", () => byId("parameter-dialog").close());
      byId("parameter-list").addEventListener("change", event => {
        const checkbox = event.target;
        if (!checkbox.matches("input[data-parameter-id]")) return;
        if (checkbox.checked) state.draftParameterIds.add(checkbox.dataset.parameterId);
        else state.draftParameterIds.delete(checkbox.dataset.parameterId);
        if (byId("parameter-checked-only").checked) renderParameters();
        else updateParameterSelection();
      });
      byId("parameters-all").addEventListener("change", event => {
        for (const parameter of filteredParameters()) {
          if (parameter.baseField) continue;
          if (event.target.checked) state.draftParameterIds.add(parameter.id);
          else state.draftParameterIds.delete(parameter.id);
        }
        renderParameters();
      });
      byId("parameter-form").addEventListener("submit", event => {
        event.preventDefault();
        if (!state.active || state.busy || state.loading) return;
        const parameterIds = (state.snapshot.parameters || []).filter(parameter =>
          !parameter.baseField && state.draftParameterIds.has(parameter.id)).map(parameter => parameter.id);
        byId("parameter-dialog").close();
        request("refresh", { parameterIds });
      });
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
          else if (message.type === "tableResult") await savedTables.handleResult(message);
          else if (message.type === "publishResult") cloudPublish.handleResult(message);
          else if (message.type === "closeRequest") savedTables.close();
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
