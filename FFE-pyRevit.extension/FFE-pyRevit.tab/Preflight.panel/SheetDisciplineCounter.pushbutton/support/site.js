(function attachSheetOrderApp(globalScope) {
  "use strict";

  var refs = null;
  var state = {
    documentTitle: "",
    sheets: [],
    disciplines: [],
    sheetsByDiscipline: Object.create(null),
    selectedDiscipline: null,
    dirtyDisciplines: Object.create(null),
    editedOrders: Object.create(null),
    busy: false
  };

  function hasBridge() {
    return !!(
      globalScope.chrome &&
      globalScope.chrome.webview &&
      typeof globalScope.chrome.webview.postMessage === "function"
    );
  }

  function postRevitMessage(message) {
    if (!hasBridge()) {
      setStatus("error", "Revit bridge is not available.");
      return false;
    }

    globalScope.chrome.webview.postMessage(JSON.stringify(message));
    return true;
  }

  function cleanText(value) {
    if (value === null || value === undefined) {
      return "";
    }
    return String(value);
  }

  function disciplineLabel(value) {
    var text = cleanText(value).trim();
    return text || "Unassigned";
  }

  function pluralize(count, singular, plural) {
    return count === 1 ? singular : plural;
  }

  function toId(value) {
    var numberValue = Number(value);
    return Number.isFinite(numberValue) ? numberValue : value;
  }

  function idKey(value) {
    return String(value);
  }

  function createMap() {
    return Object.create(null);
  }

  function arraysEqual(left, right) {
    if (!left || !right || left.length !== right.length) {
      return false;
    }

    for (var index = 0; index < left.length; index += 1) {
      if (idKey(left[index]) !== idKey(right[index])) {
        return false;
      }
    }

    return true;
  }

  function countDirtyDisciplines() {
    return Object.keys(state.dirtyDisciplines).length;
  }

  function setBusy(isBusy) {
    state.busy = !!isBusy;
    updateButtons();
  }

  function setStatus(statusOrState, message) {
    var statusState = typeof statusOrState === "object"
      ? statusOrState
      : { status: statusOrState || "ready", message: message || "" };

    if (!refs || !refs.statusBanner || !refs.statusText) {
      return;
    }

    var status = cleanText(statusState.status || "ready").toLowerCase();
    var text = cleanText(statusState.message || "");
    var issues = Array.isArray(statusState.issues) ? statusState.issues : [];

    refs.statusBanner.className = "status-banner status-" + status;
    refs.statusText.textContent = text || "Ready.";

    if (refs.issueList) {
      refs.issueList.innerHTML = "";
      refs.issueList.hidden = !issues.length;

      issues.slice(0, 8).forEach(function appendIssue(issue) {
        var item = document.createElement("li");
        item.textContent = cleanText(issue);
        refs.issueList.appendChild(item);
      });

      if (issues.length > 8) {
        var overflowItem = document.createElement("li");
        overflowItem.textContent = (issues.length - 8) + " more issues.";
        refs.issueList.appendChild(overflowItem);
      }
    }
  }

  function getOriginalSheets(discipline) {
    return state.sheetsByDiscipline[cleanText(discipline)] || [];
  }

  function getOriginalIds(discipline) {
    return getOriginalSheets(discipline).map(function mapSheet(sheet) {
      return sheet.id;
    });
  }

  function getSheetMap() {
    var byId = createMap();
    state.sheets.forEach(function addSheet(sheet) {
      byId[idKey(sheet.id)] = sheet;
    });
    return byId;
  }

  function getDisplaySheets(discipline) {
    var disciplineValue = cleanText(discipline);
    var editedOrder = state.editedOrders[disciplineValue];
    var originalSheets = getOriginalSheets(disciplineValue);

    if (!editedOrder) {
      return originalSheets.slice();
    }

    var byId = getSheetMap();
    var used = createMap();
    var ordered = [];

    editedOrder.forEach(function addEdited(id) {
      var sheet = byId[idKey(id)];
      if (sheet && cleanText(sheet.discipline) === disciplineValue && !used[idKey(sheet.id)]) {
        ordered.push(sheet);
        used[idKey(sheet.id)] = true;
      }
    });

    originalSheets.forEach(function appendMissing(sheet) {
      if (!used[idKey(sheet.id)]) {
        ordered.push(sheet);
      }
    });

    return ordered;
  }

  function groupSheetsByDiscipline(sheets) {
    var grouped = createMap();

    sheets.forEach(function addSheet(sheet) {
      var discipline = cleanText(sheet.discipline);
      if (!grouped[discipline]) {
        grouped[discipline] = [];
      }
      grouped[discipline].push(sheet);
    });

    return grouped;
  }

  function normalizePayload(payload) {
    var data = payload || {};
    var sheets = Array.isArray(data.sheets) ? data.sheets.slice() : [];
    var disciplines = Array.isArray(data.disciplines) ? data.disciplines.slice() : [];

    sheets.forEach(function normalizeSheet(sheet) {
      sheet.id = toId(sheet.id);
      sheet.discipline = cleanText(sheet.discipline);
      sheet.disciplineLabel = cleanText(sheet.disciplineLabel) || disciplineLabel(sheet.discipline);
      sheet.sheetNumber = cleanText(sheet.sheetNumber);
      sheet.name = cleanText(sheet.name);
      sheet.orderValue = cleanText(sheet.orderValue);
    });

    disciplines.forEach(function normalizeDiscipline(discipline) {
      discipline.value = cleanText(discipline.value);
      discipline.label = cleanText(discipline.label) || disciplineLabel(discipline.value);
      discipline.sheetCount = Number(discipline.sheetCount || 0);
    });

    if (!disciplines.length && sheets.length) {
      var grouped = groupSheetsByDiscipline(sheets);
      disciplines = Object.keys(grouped).map(function createDiscipline(value) {
        return {
          value: value,
          label: disciplineLabel(value),
          sheetCount: grouped[value].length
        };
      }).sort(function sortDisciplines(left, right) {
        return left.label.localeCompare(right.label, undefined, { numeric: true, sensitivity: "base" });
      });
    }

    return {
      documentTitle: cleanText(data.documentTitle),
      sheets: sheets,
      disciplines: disciplines,
      sheetsByDiscipline: groupSheetsByDiscipline(sheets)
    };
  }

  function loadData(payload) {
    var normalized = normalizePayload(payload);
    var previousSelection = state.selectedDiscipline;

    state.documentTitle = normalized.documentTitle;
    state.sheets = normalized.sheets;
    state.disciplines = normalized.disciplines;
    state.sheetsByDiscipline = normalized.sheetsByDiscipline;
    state.dirtyDisciplines = createMap();
    state.editedOrders = createMap();

    var disciplineValues = state.disciplines.map(function getValue(discipline) {
      return discipline.value;
    });

    if (previousSelection !== null && disciplineValues.indexOf(previousSelection) !== -1) {
      state.selectedDiscipline = previousSelection;
    } else if (disciplineValues.length) {
      state.selectedDiscipline = disciplineValues[0];
    } else {
      state.selectedDiscipline = null;
    }

    setBusy(false);
    render();
    setStatus("ready", "Loaded " + state.sheets.length + " " + pluralize(state.sheets.length, "sheet", "sheets") + ".");
  }

  function updateButtons() {
    if (!refs) {
      return;
    }

    var dirtyCount = countDirtyDisciplines();
    refs.saveButton.disabled = state.busy || dirtyCount === 0;
    refs.refreshButton.disabled = state.busy;
    refs.disciplineSelect.disabled = state.busy || !state.disciplines.length;
  }

  function renderDocumentTitle() {
    if (!refs) {
      return;
    }

    refs.documentTitle.textContent = state.documentTitle || "";
    refs.documentTitle.hidden = !state.documentTitle;
  }

  function renderDisciplineSelect() {
    if (!refs) {
      return;
    }

    refs.disciplineSelect.innerHTML = "";

    state.disciplines.forEach(function appendDiscipline(discipline) {
      var option = document.createElement("option");
      option.value = discipline.value;
      option.textContent = discipline.label + " (" + discipline.sheetCount + ")";
      refs.disciplineSelect.appendChild(option);
    });

    if (state.selectedDiscipline !== null) {
      refs.disciplineSelect.value = state.selectedDiscipline;
    }
  }

  function renderToolbar() {
    if (!refs) {
      return;
    }

    var selectedSheets = state.selectedDiscipline === null ? [] : getDisplaySheets(state.selectedDiscipline);
    var dirtyCount = countDirtyDisciplines();

    refs.sheetCount.textContent = selectedSheets.length + " " + pluralize(selectedSheets.length, "sheet", "sheets");
    refs.dirtyCount.textContent = dirtyCount
      ? dirtyCount + " edited " + pluralize(dirtyCount, "discipline", "disciplines")
      : "";

    updateButtons();
  }

  function createSheetRow(sheet, index) {
    var row = document.createElement("li");
    row.className = "sheet-row";
    row.draggable = true;
    row.dataset.sheetId = idKey(sheet.id);

    if (sheet.canWriteOrder === false) {
      row.className += " read-only";
    }

    var handle = document.createElement("span");
    handle.className = "drag-handle";
    handle.setAttribute("aria-hidden", "true");
    handle.textContent = "::";

    var position = document.createElement("span");
    position.className = "position";
    position.textContent = String(index + 1).padStart(2, "0");

    var main = document.createElement("div");
    main.className = "sheet-main";

    var number = document.createElement("span");
    number.className = "sheet-number";
    number.textContent = sheet.sheetNumber || "(No number)";

    var name = document.createElement("span");
    name.className = "sheet-name";
    name.textContent = sheet.name || "(No name)";

    main.appendChild(number);
    main.appendChild(name);

    var meta = document.createElement("span");
    meta.className = "sheet-meta";
    meta.textContent = sheet.orderValue ? "Order " + sheet.orderValue : "No order";

    row.appendChild(handle);
    row.appendChild(position);
    row.appendChild(main);
    row.appendChild(meta);

    row.addEventListener("dragstart", onDragStart);
    row.addEventListener("dragend", onDragEnd);

    return row;
  }

  function renderList() {
    if (!refs) {
      return;
    }

    refs.sheetList.innerHTML = "";

    var selectedSheets = state.selectedDiscipline === null ? [] : getDisplaySheets(state.selectedDiscipline);
    refs.emptyState.hidden = selectedSheets.length > 0;
    refs.emptyState.textContent = state.disciplines.length ? "No sheets in this discipline." : "No sheets found.";

    selectedSheets.forEach(function appendSheet(sheet, index) {
      refs.sheetList.appendChild(createSheetRow(sheet, index));
    });
  }

  function render() {
    renderDocumentTitle();
    renderDisciplineSelect();
    renderList();
    renderToolbar();
  }

  function getDragAfterElement(container, y) {
    var rows = Array.prototype.slice.call(container.querySelectorAll(".sheet-row:not(.dragging)"));

    return rows.reduce(function findClosest(closest, child) {
      var box = child.getBoundingClientRect();
      var offset = y - box.top - (box.height / 2);

      if (offset < 0 && offset > closest.offset) {
        return { offset: offset, element: child };
      }

      return closest;
    }, { offset: Number.NEGATIVE_INFINITY, element: null }).element;
  }

  function onDragStart(event) {
    if (state.busy) {
      event.preventDefault();
      return;
    }

    event.currentTarget.classList.add("dragging");
    event.dataTransfer.effectAllowed = "move";
    event.dataTransfer.setData("text/plain", event.currentTarget.dataset.sheetId);
  }

  function onDragEnd(event) {
    event.currentTarget.classList.remove("dragging");
    updateEditedOrderFromDom();
  }

  function onListDragOver(event) {
    var dragging = refs.sheetList.querySelector(".dragging");
    if (!dragging) {
      return;
    }

    event.preventDefault();
    var afterElement = getDragAfterElement(refs.sheetList, event.clientY);

    if (afterElement === null) {
      refs.sheetList.appendChild(dragging);
    } else {
      refs.sheetList.insertBefore(dragging, afterElement);
    }
  }

  function updateEditedOrderFromDom() {
    if (state.selectedDiscipline === null || !refs) {
      return;
    }

    var ids = Array.prototype.slice.call(refs.sheetList.querySelectorAll(".sheet-row")).map(function readId(row) {
      return toId(row.dataset.sheetId);
    });

    var originalIds = getOriginalIds(state.selectedDiscipline);
    if (arraysEqual(ids, originalIds)) {
      delete state.editedOrders[state.selectedDiscipline];
      delete state.dirtyDisciplines[state.selectedDiscipline];
    } else {
      state.editedOrders[state.selectedDiscipline] = ids;
      state.dirtyDisciplines[state.selectedDiscipline] = true;
      setStatus("warning", "Unsaved sheet order changes.");
    }

    renderList();
    renderToolbar();
  }

  function saveOrder() {
    var dirtyKeys = Object.keys(state.dirtyDisciplines);
    if (!dirtyKeys.length) {
      setStatus("ready", "No changes to save.");
      return;
    }

    var disciplines = dirtyKeys.map(function buildDisciplinePayload(discipline) {
      return {
        discipline: discipline,
        sheetIds: state.editedOrders[discipline] || getOriginalIds(discipline)
      };
    });

    setBusy(true);
    setStatus("busy", "Saving sheet order in Revit...");

    if (!postRevitMessage({
      type: "saveOrder",
      payload: {
        disciplines: disciplines
      }
    })) {
      setBusy(false);
    }
  }

  function refreshSheets() {
    if (countDirtyDisciplines() && !globalScope.confirm("Discard unsaved changes and refresh from Revit?")) {
      return;
    }

    setBusy(true);
    setStatus("busy", "Refreshing sheets from Revit...");

    if (!postRevitMessage({ type: "refreshSheets" })) {
      setBusy(false);
    }
  }

  function handleSaveResult(result) {
    setBusy(false);
    setStatus(result || { status: "ready", message: "Save complete." });
  }

  function handleRefreshResult(result) {
    setBusy(false);
    setStatus(result || { status: "ready", message: "Refresh complete." });
  }

  function bindEvents() {
    refs.disciplineSelect.addEventListener("change", function onDisciplineChanged() {
      state.selectedDiscipline = refs.disciplineSelect.value;
      renderList();
      renderToolbar();
    });

    refs.saveButton.addEventListener("click", saveOrder);
    refs.refreshButton.addEventListener("click", refreshSheets);
    refs.sheetList.addEventListener("dragover", onListDragOver);
  }

  function init() {
    refs = {
      documentTitle: document.getElementById("document-title"),
      statusBanner: document.getElementById("status-banner"),
      statusText: document.getElementById("status-text"),
      issueList: document.getElementById("issue-list"),
      disciplineSelect: document.getElementById("discipline-select"),
      sheetCount: document.getElementById("sheet-count"),
      dirtyCount: document.getElementById("dirty-count"),
      saveButton: document.getElementById("save-button"),
      refreshButton: document.getElementById("refresh-button"),
      sheetList: document.getElementById("sheet-list"),
      emptyState: document.getElementById("empty-state")
    };

    bindEvents();
    render();
    setBusy(true);
    setStatus("busy", "Loading sheets from Revit...");

    if (!postRevitMessage({ type: "appReady" })) {
      setBusy(false);
    }
  }

  globalScope.ffeSheets = {
    loadData: loadData,
    handleSaveResult: handleSaveResult,
    handleRefreshResult: handleRefreshResult,
    setStatus: setStatus
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
}(window));
