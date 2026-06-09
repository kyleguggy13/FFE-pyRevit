(function attachSheetOrderApp(globalScope) {
  "use strict";

  var refs = null;
  var sortableList = null;
  var state = {
    documentTitle: "",
    sheets: [],
    disciplines: [],
    sheetsByDiscipline: Object.create(null),
    selectedDiscipline: null,
    dirtyDisciplines: Object.create(null),
    editedOrders: Object.create(null),
    showNonPrintable: false,
    includeNonPrintableInIndex: false,
    busy: false
  };

  function hasBridge() {
    return !!(
      globalScope.chrome &&
      globalScope.chrome.webview &&
      typeof globalScope.chrome.webview.postMessage === "function"
    );
  }

  function hasSortableJs() {
    return !!(
      globalScope.Sortable &&
      typeof globalScope.Sortable.create === "function"
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

  function toBool(value) {
    if (value === true || value === 1) {
      return true;
    }
    var text = cleanText(value).trim().toLowerCase();
    return text === "1" || text === "true" || text === "yes" || text === "on";
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

  function destroySortable() {
    if (!sortableList) {
      return;
    }

    try {
      sortableList.destroy();
    } catch (error) {
      // Sortable owns DOM listeners; stale destroy failures should not block rendering.
    }

    sortableList = null;
  }

  function updateSortableState() {
    if (!sortableList || !refs) {
      return;
    }

    var selectedSheets = state.selectedDiscipline === null ? [] : getDisplaySheets(state.selectedDiscipline);
    sortableList.option("disabled", state.busy || selectedSheets.length < 2);
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

  function getSheetMap() {
    var byId = createMap();
    state.sheets.forEach(function addSheet(sheet) {
      byId[idKey(sheet.key)] = sheet;
    });
    return byId;
  }

  function shouldShowSheet(sheet) {
    return state.showNonPrintable || !sheet.isNonPrintable;
  }

  function shouldIndexSheet(sheet) {
    return state.includeNonPrintableInIndex || !sheet.isNonPrintable;
  }

  function getOrderedSheets(discipline, predicate) {
    var disciplineValue = cleanText(discipline);
    var editedOrder = state.editedOrders[disciplineValue];
    var originalSheets = getOriginalSheets(disciplineValue);

    if (!editedOrder) {
      return originalSheets.filter(predicate);
    }

    var byId = getSheetMap();
    var used = createMap();
    var ordered = [];

    editedOrder.forEach(function addEdited(id) {
      var sheet = byId[idKey(id)];
      if (
        sheet &&
        cleanText(sheet.discipline) === disciplineValue &&
        predicate(sheet) &&
        !used[idKey(sheet.key)]
      ) {
        ordered.push(sheet);
        used[idKey(sheet.key)] = true;
      }
    });

    originalSheets.forEach(function appendMissing(sheet) {
      if (predicate(sheet) && !used[idKey(sheet.key)]) {
        ordered.push(sheet);
      }
    });

    return ordered;
  }

  function getDisplaySheets(discipline) {
    return getOrderedSheets(discipline, shouldShowSheet);
  }

  function getIndexSheets(discipline) {
    return getOrderedSheets(discipline, shouldIndexSheet);
  }

  function getOriginalDisplayKeys(discipline) {
    return getOriginalSheets(discipline).filter(shouldShowSheet).map(function mapSheet(sheet) {
      return sheet.key;
    });
  }

  function getIndexKeys(discipline) {
    return getIndexSheets(discipline).map(function mapSheet(sheet) {
      return sheet.key;
    });
  }

  function markVisibleDisciplinesDirty() {
    getVisibleDisciplines().forEach(function markDirty(discipline) {
      var displayKeys = getDisplaySheets(discipline.value).map(function mapSheet(sheet) {
        return sheet.key;
      });
      state.editedOrders[discipline.value] = displayKeys;
      state.dirtyDisciplines[discipline.value] = true;
    });
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
      sheet.key = cleanText(sheet.key || sheet.id);
      sheet.discipline = cleanText(sheet.discipline);
      sheet.disciplineLabel = cleanText(sheet.disciplineLabel) || disciplineLabel(sheet.discipline);
      sheet.sheetNumber = cleanText(sheet.sheetNumber);
      sheet.name = cleanText(sheet.name);
      sheet.orderValue = cleanText(sheet.orderValue);
      sheet.sourceKey = cleanText(sheet.sourceKey);
      sheet.sourceLabel = cleanText(sheet.sourceLabel);
      sheet.isLinked = toBool(sheet.isLinked);
      sheet.isPlaceholder = toBool(sheet.isPlaceholder);
      sheet.isScheduled = sheet.isScheduled === undefined ? true : toBool(sheet.isScheduled);
      sheet.canBePrinted = sheet.canBePrinted === undefined ? true : toBool(sheet.canBePrinted);
      sheet.isNonPrintable = toBool(sheet.isNonPrintable) || sheet.isLinked || sheet.isPlaceholder || !sheet.isScheduled || !sheet.canBePrinted;
      sheet.canWriteOrder = sheet.canWriteOrder !== false && !sheet.isLinked;
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

  function getVisibleDisciplines() {
    return state.disciplines.map(function countVisible(discipline) {
      var value = cleanText(discipline.value);
      var visibleCount = getOriginalSheets(value).filter(shouldShowSheet).length;
      return {
        value: value,
        label: cleanText(discipline.label) || disciplineLabel(value),
        sheetCount: visibleCount
      };
    }).filter(function hasVisibleSheets(discipline) {
      return discipline.sheetCount > 0;
    });
  }

  function ensureSelectedDiscipline(visibleDisciplines) {
    var values = visibleDisciplines.map(function getValue(discipline) {
      return discipline.value;
    });

    if (state.selectedDiscipline !== null && values.indexOf(state.selectedDiscipline) !== -1) {
      return;
    }

    state.selectedDiscipline = values.length ? values[0] : null;
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

    state.selectedDiscipline = previousSelection;
    ensureSelectedDiscipline(getVisibleDisciplines());

    setBusy(false);
    render();
    var visibleCount = getVisibleDisciplines().reduce(function sumSheets(total, discipline) {
      return total + discipline.sheetCount;
    }, 0);
    setStatus("ready", "Loaded " + visibleCount + " visible " + pluralize(visibleCount, "sheet", "sheets") + ".");
  }

  function updateButtons() {
    if (!refs) {
      return;
    }

    var dirtyCount = countDirtyDisciplines();
    refs.saveButton.disabled = state.busy || dirtyCount === 0;
    refs.refreshButton.disabled = state.busy;
    refs.disciplineSelect.disabled = state.busy || !getVisibleDisciplines().length;
    refs.showNonPrintableCheckbox.disabled = state.busy;
    refs.showNonPrintableCheckbox.checked = state.showNonPrintable;
    refs.includeNonPrintableCheckbox.disabled = state.busy || !state.showNonPrintable;
    refs.includeNonPrintableCheckbox.checked = state.includeNonPrintableInIndex && state.showNonPrintable;
    updateSortableState();
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
    var visibleDisciplines = getVisibleDisciplines();
    ensureSelectedDiscipline(visibleDisciplines);

    visibleDisciplines.forEach(function appendDiscipline(discipline) {
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
    row.draggable = !hasSortableJs(); // Use native HTML5 drag-and-drop if SortableJS is not available
    row.dataset.sheetKey = idKey(sheet.key);

    if (sheet.canWriteOrder === false) {
      row.className += " read-only";
    }

    if (sheet.isNonPrintable) {
      row.className += " non-printable";
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
    var metaParts = [sheet.orderValue ? "Order " + sheet.orderValue : "No order"];
    if (sheet.isLinked) {
      metaParts.push("Linked: " + (sheet.sourceLabel || "model"));
    }
    if (sheet.isPlaceholder) {
      metaParts.push("Placeholder");
    }
    if (sheet.isScheduled === false) {
      metaParts.push("Not in sheet list");
    }
    if (sheet.canBePrinted === false) {
      metaParts.push("Not printable");
    }
    if (sheet.canWriteOrder === false) {
      metaParts.push("Read-only");
    }
    meta.textContent = metaParts.join(" | ");

    row.appendChild(handle);
    row.appendChild(position);
    row.appendChild(main);
    row.appendChild(meta);

    if (!hasSortableJs()) {
      row.addEventListener("dragstart", onDragStart);
      row.addEventListener("dragend", onDragEnd);
    }

    return row;
  }

  function setupSortable() {
    destroySortable();

    if (!refs || !hasSortableJs()) {
      return;
    }

    sortableList = globalScope.Sortable.create(refs.sheetList, {
      animation: 150,
      chosenClass: "sortable-chosen",
      dragClass: "sortable-drag",
      draggable: ".sheet-row",
      ghostClass: "sortable-ghost",
      handle: ".drag-handle",
      onEnd: updateEditedOrderFromDom
    });

    updateSortableState();
  }

  function renderList() {
    if (!refs) {
      return;
    }

    destroySortable();
    refs.sheetList.innerHTML = "";

    var selectedSheets = state.selectedDiscipline === null ? [] : getDisplaySheets(state.selectedDiscipline);
    refs.emptyState.hidden = selectedSheets.length > 0;
    refs.emptyState.textContent = getVisibleDisciplines().length ? "No sheets in this discipline." : "No sheets found.";

    selectedSheets.forEach(function appendSheet(sheet, index) {
      refs.sheetList.appendChild(createSheetRow(sheet, index));
    });

    setupSortable();
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
    event.dataTransfer.setData("text/plain", event.currentTarget.dataset.sheetKey);
  }

  function onDragEnd(event) {
    event.currentTarget.classList.remove("dragging");
    updateEditedOrderFromDom();
  }

  function onListDragOver(event) {
    if (hasSortableJs()) {
      return;
    }
    
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

    var keys = Array.prototype.slice.call(refs.sheetList.querySelectorAll(".sheet-row")).map(function readKey(row) {
      return cleanText(row.dataset.sheetKey);
    });

    var originalKeys = getOriginalDisplayKeys(state.selectedDiscipline);
    if (arraysEqual(keys, originalKeys)) {
      delete state.editedOrders[state.selectedDiscipline];
      delete state.dirtyDisciplines[state.selectedDiscipline];
    } else {
      state.editedOrders[state.selectedDiscipline] = keys;
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
        sheetKeys: getIndexKeys(discipline)
      };
    });

    setBusy(true);
    setStatus("busy", "Saving sheet order in Revit...");

    if (!postRevitMessage({
      type: "saveOrder",
      payload: {
        includeNonPrintableInIndex: state.includeNonPrintableInIndex,
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

    refs.showNonPrintableCheckbox.addEventListener("change", function onShowNonPrintableChanged() {
      state.showNonPrintable = refs.showNonPrintableCheckbox.checked;
      if (!state.showNonPrintable) {
        state.includeNonPrintableInIndex = false;
      }
      render();
    });

    refs.includeNonPrintableCheckbox.addEventListener("change", function onIncludeNonPrintableChanged() {
      state.includeNonPrintableInIndex = refs.includeNonPrintableCheckbox.checked && state.showNonPrintable;
      markVisibleDisciplinesDirty();
      setStatus("warning", "Unsaved sheet order changes.");
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
      showNonPrintableCheckbox: document.getElementById("show-non-printable-checkbox"),
      includeNonPrintableCheckbox: document.getElementById("include-non-printable-checkbox"),
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
