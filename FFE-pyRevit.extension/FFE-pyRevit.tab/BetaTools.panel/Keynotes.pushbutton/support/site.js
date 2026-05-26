(function attachKeynoteManager(globalScope) {
  "use strict";

  var UNGROUPED_ID = "__ungrouped__";

  var state = {
    payload: null,
    entries: [],
    sourceIssues: [],
    selectedDivisionId: null,
    selectedNoteId: null,
    selectedId: null,
    dirty: false,
    saving: false,
    dbReady: false,
    dbInitializing: false,
    dbSnapshot: null,
    baselineEntries: [],
    syncIssues: [],
    remotePending: false,
    allowNextLoad: false,
    nextLocalId: 1
  };

  var STATUS_TITLES = {
    ready: "Ready",
    syncing: "Syncing",
    conflict: "Conflict",
    invalidFormat: "Validation Required",
    missingFile: "File Missing",
    unsupported: "Unsupported Keynote Reference",
    error: "Error",
    warning: "Working",
    idle: "<Messages>"
  };

  function hasWebViewBridge() {
    return Boolean(
      globalScope.chrome &&
      globalScope.chrome.webview &&
      typeof globalScope.chrome.webview.postMessage === "function"
    );
  }

  function postWebViewMessage(message) {
    if (!hasWebViewBridge()) {
      setStatus({
        status: "error",
        message: "Open this manager from pyRevit to connect it to Revit."
      });
      return false;
    }

    globalScope.chrome.webview.postMessage(JSON.stringify(message));
    return true;
  }

  function byId(id) {
    return document.getElementById(id);
  }

  function text(value) {
    if (value === null || value === undefined) {
      return "";
    }
    return String(value);
  }

  function trim(value) {
    return text(value).replace(/^\s+|\s+$/g, "");
  }

  function setText(target, value) {
    var element = typeof target === "string" ? byId(target) : target;
    if (element) {
      element.textContent = text(value);
    }
  }

  function clearElement(element) {
    if (!element) {
      return;
    }
    while (element.firstChild) {
      element.removeChild(element.firstChild);
    }
  }

  function formatNumber(value) {
    var number = Number(value || 0);
    try {
      return number.toLocaleString();
    } catch (ignore) {
      return String(number);
    }
  }

  function makeLocalId() {
    var id = "local-" + state.nextLocalId;
    state.nextLocalId += 1;
    return id;
  }

  function normalizeEntry(entry, index) {
    var dbId = text(entry.dbId || "");
    return {
      id: text(entry.id || makeLocalId()),
      dbId: dbId,
      key: trim(entry.key),
      text: trim(entry.text),
      parentKey: trim(entry.parentKey),
      rowVersion: entry.rowVersion || null,
      sortOrder: entry.sortOrder === 0 ? 0 : (entry.sortOrder || index),
      lineNumber: entry.lineNumber || null,
      originalIndex: index
    };
  }

  function sourceIssueBlocksSave(issue) {
    var code = text(issue.code);
    return (
      code === "malformedLine" ||
      code === "unsupportedReference" ||
      code === "missingFile" ||
      code === "loadError" ||
      code === "writeUnavailable"
    );
  }

  function hasSourceIssueCode(code) {
    return state.sourceIssues.some(function (issue) {
      return text(issue.code) === code;
    });
  }

  function makeIssue(severity, message, key, code) {
    return {
      severity: severity || "error",
      message: message || "",
      key: key || "",
      code: code || ""
    };
  }

  function findEntryById(id) {
    var index;
    for (index = 0; index < state.entries.length; index += 1) {
      if (state.entries[index].id === id) {
        return state.entries[index];
      }
    }
    return null;
  }

  function selectedDivisionEntry() {
    if (state.selectedDivisionId === UNGROUPED_ID) {
      return null;
    }
    return findEntryById(state.selectedDivisionId);
  }

  function selectedNoteEntry() {
    return findEntryById(state.selectedNoteId);
  }

  function selectedEntry() {
    return selectedNoteEntry() || selectedDivisionEntry();
  }

  function entriesByKey() {
    var map = {};
    state.entries.forEach(function (entry) {
      if (entry.key && !map[entry.key]) {
        map[entry.key] = entry;
      }
    });
    return map;
  }

  function childCount(key) {
    var count = 0;
    state.entries.forEach(function (entry) {
      if (entry.parentKey === key) {
        count += 1;
      }
    });
    return count;
  }

  function rootEntries() {
    return state.entries.filter(function (entry) {
      return !trim(entry.parentKey);
    });
  }

  function buildChildrenMap() {
    var children = {};
    var byKey = entriesByKey();

    state.entries.forEach(function (entry) {
      var parentKey = entry.parentKey && byKey[entry.parentKey] ? entry.parentKey : "";
      if (!children[parentKey]) {
        children[parentKey] = [];
      }
      children[parentKey].push(entry);
    });

    return children;
  }

  function appendDescendants(children, parentKey, depth, rows, visited) {
    (children[parentKey] || []).forEach(function (entry) {
      if (visited[entry.id]) {
        return;
      }
      visited[entry.id] = true;
      rows.push({ entry: entry, depth: depth });
      appendDescendants(children, entry.key, depth + 1, rows, visited);
    });
  }

  function buildRowsForDivision(entry) {
    var rows = [];
    var children = buildChildrenMap();
    appendDescendants(children, entry.key, 0, rows, {});
    return rows;
  }

  function buildUngroupedRows() {
    var rows = [];
    var byKey = entriesByKey();
    var children = buildChildrenMap();
    var roots = rootEntries();
    var visited = {};

    state.entries.forEach(function (entry) {
      var shouldStart = false;

      if (!roots.length) {
        shouldStart = !entry.parentKey || !byKey[entry.parentKey];
      } else {
        shouldStart = Boolean(entry.parentKey && !byKey[entry.parentKey]);
      }

      if (shouldStart && !visited[entry.id]) {
        visited[entry.id] = true;
        rows.push({ entry: entry, depth: 0 });
        appendDescendants(children, entry.key, 1, rows, visited);
      }
    });

    if (!roots.length) {
      state.entries.forEach(function (entry) {
        if (!visited[entry.id]) {
          visited[entry.id] = true;
          rows.push({ entry: entry, depth: 0 });
        }
      });
    }

    return rows;
  }

  function divisionCode(key) {
    var value = trim(key);
    var match = value.match(/^\d+$/);
    if (match && value.length === 1) {
      return "0" + value;
    }
    return value || "--";
  }

  function divisionTitle(entry) {
    if (!entry) {
      return "UNGROUPED";
    }
    // return "DIVISION " + divisionCode(entry.key);
    return divisionCode(entry.key);
  }

  function getDivisionModels() {
    var models = rootEntries().map(function (entry) {
      return {
        id: entry.id,
        entry: entry,
        title: divisionTitle(entry),
        text: entry.text || "Untitled division",
        rows: buildRowsForDivision(entry)
      };
    });
    var ungroupedRows = buildUngroupedRows();

    if (!models.length || ungroupedRows.length) {
      models.push({
        id: UNGROUPED_ID,
        entry: null,
        title: "UNGROUPED",
        text: "Keynotes without a root division",
        rows: ungroupedRows
      });
    }

    return models;
  }

  function selectedDivisionModel() {
    var models = getDivisionModels();
    var index;

    for (index = 0; index < models.length; index += 1) {
      if (models[index].id === state.selectedDivisionId) {
        return models[index];
      }
    }
    return models[0] || null;
  }

  function rootAncestorFor(entry) {
    var byKey = entriesByKey();
    var seen = {};
    var cursor = entry;
    var parent;

    while (cursor && cursor.parentKey && byKey[cursor.parentKey]) {
      if (seen[cursor.id]) {
        return null;
      }
      seen[cursor.id] = true;
      parent = byKey[cursor.parentKey];
      if (!parent.parentKey) {
        return parent;
      }
      cursor = parent;
    }

    return cursor && !cursor.parentKey ? cursor : null;
  }

  function setSelectionForEntry(entry) {
    var root = entry ? rootAncestorFor(entry) : null;

    if (!entry) {
      state.selectedDivisionId = null;
      state.selectedNoteId = null;
      state.selectedId = null;
      return;
    }

    if (root && root.id === entry.id) {
      state.selectedDivisionId = root.id;
      state.selectedNoteId = null;
      state.selectedId = root.id;
      return;
    }

    state.selectedDivisionId = root ? root.id : UNGROUPED_ID;
    state.selectedNoteId = entry.id;
    state.selectedId = entry.id;
  }

  function ensureSelection() {
    var models = getDivisionModels();
    var selectedModel = null;
    var selectedNote = selectedNoteEntry();
    var noteIsVisible = false;
    var index;

    if (!models.length) {
      state.selectedDivisionId = null;
      state.selectedNoteId = null;
      state.selectedId = null;
      return null;
    }

    for (index = 0; index < models.length; index += 1) {
      if (models[index].id === state.selectedDivisionId) {
        selectedModel = models[index];
        break;
      }
    }

    if (!selectedModel) {
      selectedModel = models[0];
      state.selectedDivisionId = selectedModel.id;
      state.selectedNoteId = null;
    }

    if (selectedNote) {
      noteIsVisible = selectedModel.rows.some(function (row) {
        return row.entry.id === selectedNote.id;
      });
      if (!noteIsVisible) {
        state.selectedNoteId = null;
      }
    }

    state.selectedId = state.selectedNoteId || (selectedModel.entry ? selectedModel.entry.id : null);
    return selectedModel;
  }

  function validateAll() {
    var issues = state.sourceIssues.concat(state.syncIssues || []);
    var keyCounts = {};
    var keyMap = {};
    var parentMap = {};

    state.entries.forEach(function (entry) {
      var key = trim(entry.key);
      var valueText = trim(entry.text);
      var parentKey = trim(entry.parentKey);

      if (!key) {
        issues.push(makeIssue("error", "Key is required.", key, "emptyKey"));
      }
      if (!valueText) {
        issues.push(makeIssue("error", "Text is required.", key, "emptyText"));
      }

      [
        { label: "Key", value: key },
        { label: "Text", value: valueText },
        { label: "Parent", value: parentKey }
      ].forEach(function (field) {
        if (field.value.indexOf("\t") !== -1 || /[\r\n]/.test(field.value)) {
          issues.push(makeIssue(
            "error",
            field.label + " cannot contain tabs or line breaks.",
            key,
            "invalidFieldCharacter"
          ));
        }
      });

      if (key) {
        keyCounts[key] = (keyCounts[key] || 0) + 1;
        if (!keyMap[key]) {
          keyMap[key] = entry;
        }
        parentMap[key] = parentKey;
      }
    });

    Object.keys(keyCounts).forEach(function (key) {
      if (keyCounts[key] > 1) {
        issues.push(makeIssue("error", "Duplicate keynote key: " + key, key, "duplicateKey"));
      }
    });

    state.entries.forEach(function (entry) {
      if (entry.parentKey && !keyMap[entry.parentKey]) {
        issues.push(makeIssue(
          "error",
          "Parent key '" + entry.parentKey + "' was not found.",
          entry.key,
          "missingParent"
        ));
      }
    });

    state.entries.forEach(function (entry) {
      var seen = {};
      var cursor = entry.key;

      while (cursor) {
        if (seen[cursor]) {
          issues.push(makeIssue(
            "error",
            "Parent cycle detected at key '" + cursor + "'.",
            entry.key,
            "parentCycle"
          ));
          return;
        }
        seen[cursor] = true;
        cursor = parentMap[cursor] || "";
      }
    });

    if (!state.entries.length) {
      issues.push(makeIssue("warning", "The keynote file contains no entries.", "", "emptyFile"));
    }

    return issues;
  }

  function hasErrorIssues(issues) {
    return (issues || []).some(function (issue) {
      return text(issue.severity).toLowerCase() === "error";
    });
  }

  function hasBlockingSourceIssue() {
    return state.sourceIssues.some(sourceIssueBlocksSave);
  }

  function rowMatchesQuery(row, query) {
    var entry = row.entry;
    var haystack = (entry.key + " " + entry.text + " " + entry.parentKey).toLowerCase();
    return !query || haystack.indexOf(query) !== -1;
  }

  function divisionMatchesQuery(model, query) {
    var haystack = (model.title + " " + model.text).toLowerCase();

    if (!query || haystack.indexOf(query) !== -1) {
      return true;
    }

    return model.rows.some(function (row) {
      return rowMatchesQuery(row, query);
    });
  }

  function selectedRows() {
    var model = ensureSelection();
    var query = trim(byId("search-input") ? byId("search-input").value : "").toLowerCase();

    if (!model) {
      return [];
    }

    return model.rows.filter(function (row) {
      return rowMatchesQuery(row, query);
    });
  }

  function setStatus(statusOrState, message) {
    var statusState = typeof statusOrState === "object"
      ? statusOrState
      : { status: statusOrState || "idle", message: message || "" };
    var status = statusState.status || "idle";
    var banner = byId("state-banner");

    if (banner) {
      banner.setAttribute("data-tone", status);
    }

    setText("state-title", STATUS_TITLES[status] || "Status");
    setText("state-message", statusState.message || "");
  }

  function confirmDiscardChanges(message) {
    if (!state.dirty) {
      return true;
    }
    return globalScope.confirm(message || "Discard unsaved keynote edits?");
  }

  function setDirty(isDirty) {
    var nextDirty = Boolean(isDirty);
    var didChange = state.dirty !== nextDirty;

    state.dirty = nextDirty;
    setText("dirty-state", state.dirty ? "UNSAVED CHANGES" : "NO CHANGES");
    renderSaveState();

    if (didChange && hasWebViewBridge()) {
      postWebViewMessage({ type: "dirtyStateChanged", dirty: state.dirty });
    }
  }

  function markDirty() {
    state.allowNextLoad = false;
    setDirty(true);
  }

  function renderMeta() {
    var payload = state.payload || {};
    setText("doc-title", payload.docTitle || "No Revit document");
    setText("keynote-path", payload.displayPath || payload.keynotePath || "No keynote file loaded");
    setText("encoding-label", payload.encoding || "-");
    setText("entry-count", formatNumber(state.entries.length));
  }

  function renderDivisions() {
    var list = byId("division-list");
    var query = trim(byId("search-input") ? byId("search-input").value : "").toLowerCase();
    var models = getDivisionModels().filter(function (model) {
      return divisionMatchesQuery(model, query);
    });
    var selectedModel = selectedDivisionModel();

    if (!list) {
      return;
    }

    clearElement(list);
    setText("filter-summary", formatNumber(models.length) + " divisions");

    if (!models.length) {
      var empty = document.createElement("div");
      empty.className = "empty-cell";
      empty.textContent = state.entries.length ? "No divisions match the search." : "No divisions loaded.";
      list.appendChild(empty);
      return;
    }

    models.forEach(function (model) {
      var item = document.createElement("button");
      var copy = document.createElement("span");
      var title = document.createElement("strong");
      var subtitle = document.createElement("span");
      var badge = document.createElement("span");
      var isSelected = selectedModel && model.id === selectedModel.id;

      item.type = "button";
      item.className = "list-group-item list-group-item-action division-row" + (isSelected ? " active is-selected" : "");
      item.setAttribute("role", "option");
      item.setAttribute("aria-selected", isSelected ? "true" : "false");
      item.setAttribute("data-entry-id", model.id);
      item.setAttribute("title", model.title + " - " + model.text);
      item.tabIndex = -1;

      copy.className = "division-row-copy";
      title.textContent = model.title;
      subtitle.textContent = model.text || "Untitled division";
      badge.className = "badge rounded-pill division-note-badge";
      badge.textContent = formatNumber(model.rows.length);
      badge.setAttribute("aria-label", formatNumber(model.rows.length) + (model.rows.length === 1 ? " note" : " notes"));

      copy.appendChild(title);
      copy.appendChild(subtitle);
      item.appendChild(copy);
      item.appendChild(badge);
      item.addEventListener("click", function () {
        selectDivision(model.id);
      });

      list.appendChild(item);
    });
  }

  function renderDivisionSelect() {
    var toggle = byId("division-select-toggle");
    var menu = byId("division-select-menu");
    var models = getDivisionModels();
    var selectedModel = selectedDivisionModel();

    if (!toggle || !menu) {
      return;
    }

    toggle.disabled = !models.length;
    toggle.setAttribute(
      "aria-label",
      selectedModel ? "Select division, current " + selectedModel.title : "Select division"
    );

    clearElement(menu);
    models.forEach(function (model) {
      var item = document.createElement("li");
      var button = document.createElement("button");
      var isSelected = selectedModel && model.id === selectedModel.id;

      button.type = "button";
      button.className = "dropdown-item" + (isSelected ? " active" : "");
      button.setAttribute("data-division-id", model.id);
      button.setAttribute("title", model.title + " - " + model.text);
      if (isSelected) {
        button.setAttribute("aria-current", "true");
      }
      button.textContent = model.title + " - " + model.text;

      item.appendChild(button);
      menu.appendChild(item);
    });
  }

  function renderDivisionHeader() {
    var model = selectedDivisionModel();
    var keyInput = byId("division-key-input");
    var textInput = byId("division-text-input");
    var hasEntry = Boolean(model && model.entry);

    if (keyInput) {
      keyInput.disabled = !hasEntry;
      keyInput.value = hasEntry ? model.entry.key : "UNGROUPED";
    }
    if (textInput) {
      textInput.disabled = !hasEntry;
      textInput.value = hasEntry ? model.entry.text : "Keynotes without a root division";
    }
  }

  function resizeNoteTextInput(input) {
    var borderSize;

    if (!input) {
      return;
    }

    input.style.height = "auto";
    borderSize = input.offsetHeight - input.clientHeight;
    input.style.height = (input.scrollHeight + borderSize) + "px";
  }

  function renderNotes() {
    var body = byId("keynote-table-body");
    var rows = selectedRows();

    if (!body) {
      return;
    }

    clearElement(body);

    if (!rows.length) {
      var emptyRow = document.createElement("tr");
      var empty = document.createElement("td");
      empty.className = "empty-cell";
      empty.setAttribute("colspan", "2");
      empty.textContent = state.entries.length ? "No notes in this division." : "No keynote notes loaded.";
      emptyRow.appendChild(empty);
      body.appendChild(emptyRow);
      return;
    }

    rows.forEach(function (row) {
      var entry = row.entry;
      var item = document.createElement("tr");
      var keyCell = document.createElement("td");
      var textCell = document.createElement("td");
      var keyInput = document.createElement("input");
      var textInput = document.createElement("textarea");

      item.className = "note-row" + (entry.id === state.selectedNoteId ? " is-selected" : "");
      item.setAttribute("data-entry-id", entry.id);
      item.tabIndex = -1;
      item.style.setProperty("--depth", row.depth);

      keyCell.className = "note-cell note-key-cell";
      textCell.className = "note-cell note-text-cell";

      keyInput.type = "text";
      keyInput.className = "form-control form-control-sm note-input note-key-input";
      keyInput.value = entry.key;
      keyInput.setAttribute("aria-label", "Key for " + (entry.key || "new keynote"));

      textInput.className = "form-control form-control-sm note-input note-text-input";
      textInput.rows = 2;
      textInput.value = entry.text;
      textInput.setAttribute("aria-label", "Description for " + (entry.key || "new keynote"));

      [keyInput, textInput].forEach(function (input) {
        input.addEventListener("focus", function () {
          selectNote(entry.id, false);
        });
        input.addEventListener("blur", function () {
          deferRenderAllWhenEditingSettles();
        });
      });

      keyInput.addEventListener("input", function () {
        updateEntryField(entry.id, "key", keyInput.value, false);
      });

      textInput.addEventListener("input", function () {
        updateEntryField(entry.id, "text", textInput.value, false);
        resizeNoteTextInput(textInput);
      });

      item.addEventListener("click", function (event) {
        if (event.target.tagName !== "INPUT" && event.target.tagName !== "TEXTAREA") {
          selectNote(entry.id, true);
        }
      });

      keyCell.appendChild(keyInput);
      textCell.appendChild(textInput);
      item.appendChild(keyCell);
      item.appendChild(textCell);
      body.appendChild(item);
      resizeNoteTextInput(textInput);
    });
  }

  function renderValidation() {
    var container = byId("validation-list");
    var issues = validateAll();
    var errors = issues.filter(function (issue) {
      return text(issue.severity).toLowerCase() === "error";
    });
    var warnings = issues.filter(function (issue) {
      return text(issue.severity).toLowerCase() !== "error";
    });
    var total = errors.length + warnings.length;

    setText("validation-summary", formatNumber(total));

    if (!container) {
      return;
    }

    clearElement(container);

    if (!issues.length) {
      var empty = document.createElement("div");
      empty.className = "empty-state";
      empty.textContent = "No validation messages.";
      container.appendChild(empty);
      return;
    }

    issues.forEach(function (issue) {
      var item = document.createElement("button");
      var label = document.createElement("span");
      var message = document.createElement("strong");
      var detail = document.createElement("span");

      item.type = "button";
      item.className = "validation-item";
      item.setAttribute("data-severity", issue.severity || "error");

      label.className = "validation-label";
      label.textContent = text(issue.severity || "error").toUpperCase();
      message.textContent = issue.message || "";
      detail.className = "validation-detail";
      detail.textContent = issue.key ? "Key: " + issue.key : issue.lineNumber ? "Line: " + issue.lineNumber : "";

      item.appendChild(label);
      item.appendChild(message);
      item.appendChild(detail);

      if (issue.key) {
        item.addEventListener("click", function () {
          selectEntryByKey(issue.key, true);
        });
      }

      container.appendChild(item);
    });
  }

  function renderRowActions() {
    var target = selectedNoteEntry() || selectedDivisionEntry();
    var division = selectedDivisionEntry();
    var deleteButton = byId("delete-row");
    var addChildButton = byId("add-child");
    var duplicateButton = byId("duplicate-row");

    if (addChildButton) {
      addChildButton.disabled = !division;
    }
    if (duplicateButton) {
      duplicateButton.disabled = !target;
    }
    if (deleteButton) {
      deleteButton.disabled = !target;
    }
  }

  function renderSaveState() {
    var saveButton = byId("save-data");
    var issues = validateAll();
    var canSave = Boolean(
      state.payload &&
      state.payload.keynotePath &&
      state.dirty &&
      !state.saving &&
      !hasBlockingSourceIssue() &&
      !hasErrorIssues(issues)
    );

    if (saveButton) {
      saveButton.disabled = !canSave;
      saveButton.textContent = state.saving ? "Saving..." : "Save";
    }
  }

  function renderAll() {
    ensureSelection();
    renderMeta();
    renderDivisions();
    renderDivisionSelect();
    renderDivisionHeader();
    renderNotes();
    renderValidation();
    renderSaveState();
    renderRowActions();
  }

  function deferRenderAllWhenEditingSettles() {
    globalScope.setTimeout(function () {
      var active = document.activeElement;
      if (
        active &&
        active.closest &&
        active.closest(".note-table-body, .selected-division-card")
      ) {
        return;
      }
      renderAll();
    }, 0);
  }

  function syncSelectionClasses() {
    Array.prototype.forEach.call(document.querySelectorAll(".division-row"), function (row) {
      var isSelected = row.getAttribute("data-entry-id") === state.selectedDivisionId;
      row.classList.toggle("is-selected", isSelected);
      row.classList.toggle("active", isSelected);
      row.setAttribute("aria-selected", isSelected ? "true" : "false");
    });

    Array.prototype.forEach.call(document.querySelectorAll(".note-row"), function (row) {
      row.classList.toggle("is-selected", row.getAttribute("data-entry-id") === state.selectedNoteId);
    });

    state.selectedId = state.selectedNoteId || (selectedDivisionEntry() ? selectedDivisionEntry().id : null);
    renderRowActions();
  }

  function selectDivision(id) {
    state.selectedDivisionId = id;
    state.selectedNoteId = null;
    state.selectedId = selectedDivisionEntry() ? selectedDivisionEntry().id : null;
    renderAll();
  }

  function selectNote(id, shouldRender) {
    var entry = findEntryById(id);
    if (!entry) {
      return;
    }
    setSelectionForEntry(entry);
    if (shouldRender) {
      renderAll();
    } else {
      syncSelectionClasses();
    }
  }

  function selectEntryByKey(key, shouldScroll) {
    var match = state.entries.filter(function (entry) {
      return entry.key === key;
    })[0];

    if (!match) {
      return;
    }

    setSelectionForEntry(match);
    renderAll();

    if (shouldScroll) {
      globalScope.setTimeout(function () {
        var selector = match.parentKey
          ? '.note-row[data-entry-id="' + match.id + '"]'
          : '.division-row[data-entry-id="' + match.id + '"]';
        var element = document.querySelector(selector);
        if (element && typeof element.scrollIntoView === "function") {
          element.scrollIntoView({ block: "nearest" });
        }
        if (element && typeof element.focus === "function") {
          element.focus();
        }
      }, 0);
    }
  }

  function dbManager() {
    return globalScope.ffeKeynoteDb || null;
  }

  function currentClient() {
    var settings = (state.payload && state.payload.supabase) || {};
    return {
      clientId: text(settings.clientId),
      clientName: text(settings.clientName)
    };
  }

  function rememberBaseline() {
    state.baselineEntries = state.entries.map(function (entry) {
      return {
        id: entry.id,
        dbId: entry.dbId,
        key: entry.key,
        text: entry.text,
        parentKey: entry.parentKey,
        rowVersion: entry.rowVersion,
        sortOrder: entry.sortOrder,
        lineNumber: entry.lineNumber || null
      };
    });
  }

  function subscribeToLibrary(snapshot) {
    var db = dbManager();
    var client = currentClient();

    if (!db || !snapshot || !snapshot.libraryId) {
      return;
    }

    db.subscribeLibrary(snapshot.libraryId, client.clientId, {
      onRemoteChange: function () {
        if (state.dirty) {
          state.remotePending = true;
          state.syncIssues = [makeIssue(
            "warning",
            "Remote Supabase changes are available. Save will check row conflicts; Refresh discards local edits and loads the latest data.",
            "",
            "remotePending"
          )];
          setStatus({
            status: "warning",
            message: "Remote Supabase changes are available while you have unsaved edits."
          });
          renderValidation();
          renderSaveState();
          return;
        }
        state.allowNextLoad = false;
        postWebViewMessage({ type: "refreshData" });
      },
      onStatus: function (status) {
        if (status === "SUBSCRIBED" && !state.dirty) {
          setStatus({ status: "ready", message: "Listening for shared keynote updates." });
        }
      }
    });
  }

  function mirrorFileSnapshot(payload, reason) {
    var db = dbManager();
    var settings = (payload && payload.supabase) || {};
    var client = currentClient();

    if (!payload || !payload.libraryKey) {
      state.dbReady = false;
      return;
    }

    if (!settings.configured) {
      state.dbReady = false;
      state.syncIssues = [makeIssue(
        "warning",
        "Supabase is not configured, so realtime sync is disabled. Click Supabase to set the project URL and publishable key.",
        "",
        "supabaseConfigMissing"
      )];
      renderValidation();
      renderSaveState();
      return;
    }

    state.dbInitializing = true;

    if (!db) {
      state.dbInitializing = false;
      state.syncIssues = [makeIssue("warning", "The Supabase database manager did not load.", "", "supabaseClientMissing")];
      renderValidation();
      renderSaveState();
      return;
    }

    try {
      db.configure(settings);
    } catch (error) {
      state.dbInitializing = false;
      state.syncIssues = [makeIssue("warning", error.message, "", "supabaseConfigError")];
      renderValidation();
      renderSaveState();
      return;
    }

    db.syncFileSnapshot({
        libraryKey: payload.libraryKey,
        displayPath: payload.displayPath || payload.keynotePath,
        keynotePath: payload.keynotePath,
        encoding: payload.encoding || "utf-8",
        lineEnding: payload.lineEnding || "\r\n",
        fileHash: payload.fileHash || "",
        lastWriteUtc: payload.lastWriteUtc || null,
        entries: payload.entries || [],
        clientId: client.clientId,
        clientName: client.clientName
    }).then(function (snapshot) {
      state.dbInitializing = false;
      state.dbReady = true;
      state.dbSnapshot = snapshot;
      subscribeToLibrary(snapshot);
      if (!state.dirty && reason === "save") {
        setStatus({ status: "ready", message: "Saved shared keynote file and mirrored to Supabase." });
      }
    }).catch(function (error) {
      state.dbInitializing = false;
      state.dbReady = false;
      state.syncIssues = [makeIssue("warning", error.message, "", "supabaseMirrorError")];
      renderValidation();
      renderSaveState();
    });
  }

  function setStatusFromPayload(payload) {
    var status = payload.status || "idle";
    var message = payload.message || "";
    if (status === "invalidFormat") {
      message = message || "Validation errors must be fixed before saving.";
    }
    setStatus({ status: status, message: message });
  }

  function applyData(payload) {
    var previousSelection = selectedEntry();
    var previousKey = previousSelection ? previousSelection.key : "";
    var selected = null;

    state.payload = payload || {};
    state.entries = (state.payload.entries || []).map(normalizeEntry);
    state.sourceIssues = (state.payload.issues || []).filter(sourceIssueBlocksSave);
    state.saving = false;

    if (previousKey) {
      selected = state.entries.filter(function (entry) {
        return entry.key === previousKey;
      })[0] || null;
    }

    if (selected) {
      setSelectionForEntry(selected);
    } else {
      state.selectedDivisionId = null;
      state.selectedNoteId = null;
      state.selectedId = null;
      ensureSelection();
    }

    setDirty(false);
    setStatusFromPayload(state.payload);
    renderAll();
  }

  function loadData(payload) {
    if (state.dirty) {
      if (state.allowNextLoad) {
        state.allowNextLoad = false;
      } else if (!confirmDiscardChanges("Reloading keynote data will discard unsaved edits. Continue?")) {
        setStatus({ status: "warning", message: "Reload canceled; unsaved changes were kept." });
        return;
      }
    } else {
      state.allowNextLoad = false;
    }

    state.syncIssues = [];
    state.remotePending = false;
    applyData(payload);
    rememberBaseline();
    mirrorFileSnapshot(state.payload, "load");
  }

  function updateEntryField(entryId, fieldName, value, shouldRender) {
    var entry = findEntryById(entryId);
    var oldKey;

    if (!entry) {
      return;
    }

    if (fieldName === "key") {
      value = trim(value);
      oldKey = entry.key;
      entry.key = value;
      if (oldKey && oldKey !== value) {
        state.entries.forEach(function (candidate) {
          if (candidate.parentKey === oldKey) {
            candidate.parentKey = value;
          }
        });
      }
    } else if (fieldName === "text") {
      entry.text = text(value);
    } else if (fieldName === "parentKey") {
      entry.parentKey = trim(value);
    }

    markDirty();

    if (shouldRender) {
      renderAll();
      return;
    }

    renderMeta();
    renderValidation();
    renderSaveState();
    syncSelectionClasses();
  }

  function makeUniqueKey(baseKey) {
    var base = trim(baseKey) || "NEW";
    var existing = {};
    var candidate = base;
    var index = 1;

    state.entries.forEach(function (entry) {
      existing[entry.key] = true;
    });

    while (existing[candidate]) {
      candidate = base + "-" + index;
      index += 1;
    }

    return candidate;
  }

  function makeRootKey() {
    var max = 0;
    rootEntries().forEach(function (entry) {
      var match = trim(entry.key).match(/^\d+$/);
      var number;
      if (match) {
        number = Number(entry.key);
        if (number > max) {
          max = number;
        }
      }
    });
    return makeUniqueKey(("0" + (max + 1)).slice(-2));
  }

  function makeChildKey(parentKey) {
    var prefix = trim(parentKey) ? trim(parentKey) + "." : "NEW.";
    var existing = {};
    var index;
    var candidate;

    state.entries.forEach(function (entry) {
      existing[entry.key] = true;
    });

    for (index = 1; index < 1000; index += 1) {
      candidate = prefix + ("0" + index).slice(-2);
      if (!existing[candidate]) {
        return candidate;
      }
    }

    return makeUniqueKey(prefix + "NEW");
  }

  function addRoot() {
    var entry = {
      id: makeLocalId(),
      key: makeRootKey(),
      text: "New division",
      parentKey: "",
      lineNumber: null,
      originalIndex: state.entries.length
    };
    state.entries.push(entry);
    state.selectedDivisionId = entry.id;
    state.selectedNoteId = null;
    state.selectedId = entry.id;
    markDirty();
    renderAll();
  }

  function addChild() {
    var parent = selectedDivisionEntry();
    var entry;

    if (!parent) {
      setStatus({
        status: "error",
        message: "Add a parent division before adding a keynote note."
      });
      return;
    }

    entry = {
      id: makeLocalId(),
      key: makeChildKey(parent.key),
      text: "New keynote",
      parentKey: parent.key,
      lineNumber: null,
      originalIndex: state.entries.length
    };
    state.entries.push(entry);
    state.selectedNoteId = entry.id;
    state.selectedId = entry.id;
    markDirty();
    renderAll();
  }

  function actionTargetEntry() {
    return selectedNoteEntry() || selectedDivisionEntry();
  }

  function duplicateSelected() {
    var entry = actionTargetEntry();
    var copy;

    if (!entry) {
      return;
    }

    copy = {
      id: makeLocalId(),
      key: makeUniqueKey(entry.key + "-COPY"),
      text: entry.text,
      parentKey: entry.parentKey,
      lineNumber: null,
      originalIndex: state.entries.length
    };
    state.entries.push(copy);
    setSelectionForEntry(copy);
    markDirty();
    renderAll();
  }

  function deleteSelected() {
    var entry = actionTargetEntry();

    if (!entry) {
      return;
    }

    if (childCount(entry.key)) {
      setStatus({
        status: "error",
        message: "Delete child keynotes before deleting parent key '" + entry.key + "'."
      });
      return;
    }

    state.entries = state.entries.filter(function (candidate) {
      return candidate.id !== entry.id;
    });

    if (state.selectedNoteId === entry.id) {
      state.selectedNoteId = null;
    }
    if (state.selectedDivisionId === entry.id) {
      state.selectedDivisionId = null;
    }
    state.selectedId = null;
    ensureSelection();
    markDirty();
    renderAll();
  }

  function refreshData() {
    var shouldAllowLoad = state.dirty;

    if (!confirmDiscardChanges("Refresh will discard unsaved keynote edits in this window. Continue?")) {
      return;
    }

    setStatus({ status: "warning", message: "Refreshing keynote file..." });
    if (postWebViewMessage({ type: "refreshData" })) {
      state.allowNextLoad = shouldAllowLoad;
    } else {
      state.allowNextLoad = false;
    }
  }

  function requestRefresh() {
    refreshData();
  }

  function saveData() {
    var issues = validateAll();
    var payload;

    renderValidation();

    if (hasErrorIssues(issues)) {
      setStatus({ status: "error", message: "Fix validation errors before saving." });
      return;
    }

    if (hasBlockingSourceIssue()) {
      setStatus({ status: "error", message: "The shared keynote file is not available for saving. Check the warning details and refresh." });
      return;
    }

    if (!state.payload || !state.payload.keynotePath) {
      setStatus({ status: "error", message: "No keynote file is available to save." });
      return;
    }

    payload = {
      keynotePath: state.payload.keynotePath,
      encoding: state.payload.encoding || "utf-8",
      lineEnding: state.payload.lineEnding || "\r\n",
      lastWriteUtc: state.payload.lastWriteUtc,
      fileHash: state.payload.fileHash,
      sourceHasMalformed: hasSourceIssueCode("malformedLine"),
      baselineEntries: state.baselineEntries.map(function (entry) {
        return {
          id: entry.id,
          key: trim(entry.key),
          text: trim(entry.text),
          parentKey: trim(entry.parentKey),
          lineNumber: entry.lineNumber || null
        };
      }),
      entries: state.entries.map(function (entry) {
        return {
          id: entry.id,
          key: trim(entry.key),
          text: trim(entry.text),
          parentKey: trim(entry.parentKey),
          lineNumber: entry.lineNumber
        };
      })
    };

    state.saving = true;
    state.syncIssues = [];
    renderSaveState();
    setStatus({ status: "syncing", message: "Merging edits into the shared keynote file..." });

    if (!postWebViewMessage({ type: "saveKeynotes", payload: payload })) {
      state.saving = false;
      renderSaveState();
    }
  }

  function handleSaveResult(result) {
    result = result || {};
    state.saving = false;

    if ((result.status || "") === "conflict") {
      state.syncIssues = result.issues || [makeIssue(
        "error",
        result.message || "Some keynote rows changed in the shared file before your save.",
        "",
        "rowConflict"
      )];
      setStatus({
        status: "conflict",
        message: result.message || "Some keynote rows changed in the shared file before your save."
      });
      renderValidation();
      renderSaveState();
      return;
    }

    if ((result.status || "") === "ready" && result.payload) {
      applyData(result.payload);
      rememberBaseline();
      mirrorFileSnapshot(result.payload, "save");
    } else if ((result.status || "") !== "ready") {
      state.syncIssues = result.issues || [];
    }

    setStatus({
      status: result.status || "idle",
      message: result.backupPath
        ? (result.message || "") + " Backup: " + result.backupPath
        : (result.message || "")
    });

    if ((result.status || "") !== "ready") {
      renderValidation();
      renderSaveState();
    }
  }

  function closeWindow() {
    var discardConfirmed = state.dirty;

    if (!confirmDiscardChanges("Close the manager and discard unsaved keynote edits in this window?")) {
      return;
    }

    if (postWebViewMessage({ type: "closeWindow", discardConfirmed: discardConfirmed })) {
      return;
    }

    try {
      globalScope.close();
    } catch (ignore) {
      // Browser fallback only.
    }
  }

  function configureSupabase() {
    if (state.dirty && !confirmDiscardChanges("Changing Supabase settings will reload keynote data and discard unsaved edits. Continue?")) {
      return;
    }

    setStatus({ status: "warning", message: "Opening Supabase settings..." });
    if (postWebViewMessage({ type: "configureSupabase" })) {
      state.allowNextLoad = state.dirty;
    }
  }

  function bindDivisionInput(id, fieldName) {
    var input = byId(id);
    if (!input) {
      return;
    }

    input.addEventListener("input", function () {
      var entry = selectedDivisionEntry();
      if (entry) {
        updateEntryField(entry.id, fieldName, input.value, false);
      }
    });
    input.addEventListener("blur", function () {
      deferRenderAllWhenEditingSettles();
    });
  }

  function bindClick(id, handler) {
    var button = byId(id);
    if (button) {
      button.addEventListener("click", handler);
    }
  }

  function init() {
    var searchInput = byId("search-input");
    var divisionSelectMenu = byId("division-select-menu");

    if (searchInput) {
      searchInput.addEventListener("input", function () {
        renderAll();
      });
    }

    if (divisionSelectMenu) {
      divisionSelectMenu.addEventListener("click", function (event) {
        var target = event.target.closest
          ? event.target.closest(".dropdown-item[data-division-id]")
          : null;
        if (!target) {
          return;
        }
        event.preventDefault();
        selectDivision(target.getAttribute("data-division-id"));
      });
    }

    bindDivisionInput("division-key-input", "key");
    bindDivisionInput("division-text-input", "text");

    bindClick("add-root", addRoot);
    bindClick("add-child", addChild);
    bindClick("duplicate-row", duplicateSelected);
    bindClick("delete-row", deleteSelected);
    bindClick("configure-supabase", configureSupabase);
    bindClick("refresh-data", refreshData);
    bindClick("save-data", saveData);
    bindClick("close-window", closeWindow);

    renderAll();
    postWebViewMessage({ type: "appReady" });
  }

  globalScope.ffeKeynotes = {
    loadData: loadData,
    setStatus: setStatus,
    handleSaveResult: handleSaveResult,
    requestRefresh: requestRefresh
  };

  if (typeof document !== "undefined") {
    document.addEventListener("DOMContentLoaded", init);
  }
}(typeof window !== "undefined" ? window : this));
