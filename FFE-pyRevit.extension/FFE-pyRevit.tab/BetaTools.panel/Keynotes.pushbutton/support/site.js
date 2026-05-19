(function attachKeynoteManager(globalScope) {
  "use strict";

  var state = {
    payload: null,
    entries: [],
    sourceIssues: [],
    selectedId: null,
    dirty: false,
    saving: false,
    nextLocalId: 1,
    collapsedSections: {}
  };

  var STATUS_TITLES = {
    ready: "Ready",
    invalidFormat: "Validation Required",
    missingFile: "File Missing",
    unsupported: "Unsupported Keynote Reference",
    error: "Error",
    warning: "Working",
    idle: "Idle"
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

  function shortPath(path) {
    var value = text(path);
    var parts;
    if (!value) {
      return "";
    }
    parts = value.split(/[\\/]/);
    return parts[parts.length - 1] || value;
  }

  function collapsedStorageKey() {
    return "ffe-keynote-manager-collapsed-sections";
  }

  function readCollapsedSections() {
    try {
      return JSON.parse(globalScope.localStorage.getItem(collapsedStorageKey())) || {};
    } catch (ignore) {
      return {};
    }
  }

  function writeCollapsedSections() {
    try {
      globalScope.localStorage.setItem(
        collapsedStorageKey(),
        JSON.stringify(state.collapsedSections || {})
      );
    } catch (ignore) {
      // Collapse state is only a convenience preference.
    }
  }

  function getCollapsibleSection(sectionId) {
    return document.querySelector('[data-section-id="' + sectionId + '"]');
  }

  function getCollapseToggle(sectionId) {
    return document.querySelector('[data-collapse-section="' + sectionId + '"]');
  }

  function sectionTitle(section) {
    var heading = section ? section.querySelector("h2") : null;
    return heading ? trim(heading.textContent) : "Section";
  }

  function setSectionCollapsed(sectionId, isCollapsed, shouldSave) {
    var section = getCollapsibleSection(sectionId);
    var button = getCollapseToggle(sectionId);
    var title = sectionTitle(section);

    state.collapsedSections[sectionId] = Boolean(isCollapsed);

    if (section) {
      section.classList.toggle("is-collapsed", Boolean(isCollapsed));
    }

    if (button) {
      button.textContent = isCollapsed ? ">" : "v";
      button.setAttribute("aria-expanded", isCollapsed ? "false" : "true");
      button.setAttribute("title", (isCollapsed ? "Expand " : "Collapse ") + title);
      button.setAttribute("aria-label", (isCollapsed ? "Expand " : "Collapse ") + title);
    }

    if (shouldSave) {
      writeCollapsedSections();
    }
  }

  function initCollapsibleSections() {
    state.collapsedSections = readCollapsedSections();

    Array.prototype.forEach.call(
      document.querySelectorAll("[data-collapse-section]"),
      function (button) {
        var sectionId = button.getAttribute("data-collapse-section");
        setSectionCollapsed(sectionId, Boolean(state.collapsedSections[sectionId]), false);
        button.addEventListener("click", function () {
          setSectionCollapsed(
            sectionId,
            !Boolean(state.collapsedSections[sectionId]),
            true
          );
        });
      }
    );
  }

  function makeLocalId() {
    var id = "local-" + state.nextLocalId;
    state.nextLocalId += 1;
    return id;
  }

  function normalizeEntry(entry, index) {
    return {
      id: text(entry.id || makeLocalId()),
      key: trim(entry.key),
      text: trim(entry.text),
      parentKey: trim(entry.parentKey),
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
      code === "loadError"
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

  function selectedEntry() {
    var index;
    for (index = 0; index < state.entries.length; index += 1) {
      if (state.entries[index].id === state.selectedId) {
        return state.entries[index];
      }
    }
    return null;
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

  function validateAll() {
    var issues = state.sourceIssues.slice();
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

  function setDirty(isDirty) {
    state.dirty = Boolean(isDirty);
    setText("dirty-state", state.dirty ? "Unsaved changes" : "No changes");
    renderSaveState();
  }

  function markDirty() {
    setDirty(true);
  }

  function buildTreeRows() {
    var rows = [];
    var children = {};
    var byKey = entriesByKey();
    var visited = {};

    state.entries.forEach(function (entry) {
      var parentKey = entry.parentKey && byKey[entry.parentKey] ? entry.parentKey : "";
      if (!children[parentKey]) {
        children[parentKey] = [];
      }
      children[parentKey].push(entry);
    });

    function appendChildren(parentKey, depth) {
      (children[parentKey] || []).forEach(function (entry) {
        if (visited[entry.id]) {
          return;
        }
        visited[entry.id] = true;
        rows.push({ entry: entry, depth: depth });
        appendChildren(entry.key, depth + 1);
      });
    }

    appendChildren("", 0);

    state.entries.forEach(function (entry) {
      if (!visited[entry.id]) {
        visited[entry.id] = true;
        rows.push({ entry: entry, depth: 0 });
      }
    });

    return rows;
  }

  function rowMatchesQuery(row, query) {
    var entry = row.entry;
    var haystack = (entry.key + " " + entry.text + " " + entry.parentKey).toLowerCase();
    return !query || haystack.indexOf(query) !== -1;
  }

  function renderTable() {
    var list = byId("keynote-table-body");
    var query = trim(byId("search-input") ? byId("search-input").value : "").toLowerCase();
    var rows = buildTreeRows().filter(function (row) {
      return rowMatchesQuery(row, query);
    });

    if (!list) {
      return;
    }

    clearElement(list);
    setText("filter-summary", formatNumber(rows.length) + " visible rows");

    if (!rows.length) {
      var emptyState = document.createElement("div");
      emptyState.className = "empty-cell";
      emptyState.textContent = state.entries.length ? "No keynotes match the search." : "No keynote rows loaded.";
      list.appendChild(emptyState);
      return;
    }

    rows.forEach(function (row) {
      var entry = row.entry;
      var item = document.createElement("button");
      var main = document.createElement("span");
      var keyText = document.createElement("span");
      var childBadge = document.createElement("span");
      var preview = document.createElement("span");
      var meta = document.createElement("span");
      var children = childCount(entry.key);
      var metaText = [];

      item.type = "button";
      item.className = entry.id === state.selectedId ? "sidebar-keynote is-selected" : "sidebar-keynote";
      item.style.setProperty("--depth", row.depth);
      item.setAttribute("role", "option");
      item.setAttribute("aria-selected", entry.id === state.selectedId ? "true" : "false");
      item.setAttribute("aria-label", "Select keynote " + (entry.key || "without key"));
      item.setAttribute("title", (entry.key || "(No key)") + " - " + (entry.text || ""));

      main.className = "sidebar-keynote__main";
      keyText.className = "sidebar-keynote__key";
      keyText.textContent = entry.key || "(No key)";
      if (children) {
        childBadge.className = "child-badge";
        childBadge.textContent = formatNumber(children);
      }

      preview.className = "sidebar-keynote__text";
      preview.textContent = entry.text || "-";

      if (entry.parentKey) {
        metaText.push("Parent " + entry.parentKey);
      }
      metaText.push(entry.lineNumber ? "Line " + entry.lineNumber : "New row");
      meta.className = "sidebar-keynote__meta";
      meta.textContent = metaText.join(" | ");

      main.appendChild(keyText);
      if (children) {
        main.appendChild(childBadge);
      }
      item.appendChild(main);
      item.appendChild(preview);
      item.appendChild(meta);

      item.addEventListener("click", function () {
        selectEntry(entry.id);
      });

      list.appendChild(item);
    });
  }

  function renderParentOptions() {
    var options = byId("parent-options");
    var selected = selectedEntry();

    if (!options) {
      return;
    }

    clearElement(options);
    state.entries.forEach(function (entry) {
      var option;
      if (selected && entry.id === selected.id) {
        return;
      }
      if (!entry.key) {
        return;
      }
      option = document.createElement("option");
      option.value = entry.key;
      options.appendChild(option);
    });
  }

  function renderEditor() {
    var entry = selectedEntry();
    var keyInput = byId("key-input");
    var textInput = byId("text-input");
    var parentInput = byId("parent-input");
    var hasSelection = Boolean(entry);

    if (keyInput) {
      keyInput.disabled = !hasSelection;
      keyInput.value = entry ? entry.key : "";
    }
    if (textInput) {
      textInput.disabled = !hasSelection;
      textInput.value = entry ? entry.text : "";
    }
    if (parentInput) {
      parentInput.disabled = !hasSelection;
      parentInput.value = entry ? entry.parentKey : "";
    }

    if (entry) {
      setText(
        "selected-summary",
        entry.key ? entry.key + " | " + shortPath(state.payload ? state.payload.keynotePath : "") : "New keynote row"
      );
    } else {
      setText("selected-summary", "Select a keynote row.");
    }

    renderParentOptions();
    renderRowActions();
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

    setText(
      "validation-summary",
      errors.length
        ? formatNumber(errors.length) + " errors"
        : warnings.length
          ? formatNumber(warnings.length) + " warnings"
          : "Ready to save"
    );

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
          var match = state.entries.filter(function (entry) {
            return entry.key === issue.key;
          })[0];
          if (match) {
            selectEntry(match.id);
          }
        });
      }

      container.appendChild(item);
    });
  }

  function renderRowActions() {
    var entry = selectedEntry();
    var hasSelection = Boolean(entry);
    var deleteButton = byId("delete-row");
    var addChildButton = byId("add-child");
    var duplicateButton = byId("duplicate-row");

    if (addChildButton) {
      addChildButton.disabled = !hasSelection;
    }
    if (duplicateButton) {
      duplicateButton.disabled = !hasSelection;
    }
    if (deleteButton) {
      deleteButton.disabled = !hasSelection;
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
      !hasErrorIssues(issues)
    );

    if (saveButton) {
      saveButton.disabled = !canSave;
      saveButton.textContent = state.saving ? "Saving..." : "Save";
    }
  }

  function renderMeta() {
    var payload = state.payload || {};
    setText("doc-title", payload.docTitle || "No Revit document");
    setText("keynote-path", payload.displayPath || payload.keynotePath || "No keynote file loaded");
    setText("encoding-label", payload.encoding || "-");
    setText("entry-count", formatNumber(state.entries.length));
  }

  function renderAll() {
    renderMeta();
    renderTable();
    renderEditor();
    renderValidation();
    renderSaveState();
  }

  function selectEntry(id) {
    state.selectedId = id;
    renderAll();
  }

  function setStatusFromPayload(payload) {
    var status = payload.status || "idle";
    var message = payload.message || "";
    if (status === "invalidFormat") {
      message = message || "Validation errors must be fixed before saving.";
    }
    setStatus({ status: status, message: message });
  }

  function loadData(payload) {
    var previousSelection = selectedEntry();
    var previousKey = previousSelection ? previousSelection.key : "";

    state.payload = payload || {};
    state.entries = (state.payload.entries || []).map(normalizeEntry);
    state.sourceIssues = (state.payload.issues || []).filter(sourceIssueBlocksSave);
    state.saving = false;

    var selected = null;
    if (previousKey) {
      selected = state.entries.filter(function (entry) {
        return entry.key === previousKey;
      })[0] || null;
    }
    if (!selected && state.entries.length) {
      selected = state.entries[0];
    }
    state.selectedId = selected ? selected.id : null;

    setDirty(false);
    setStatusFromPayload(state.payload);
    renderAll();
  }

  function updateSelectedField(fieldName, value) {
    var entry = selectedEntry();
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
    renderTable();
    renderParentOptions();
    renderValidation();
    renderSaveState();
    setText(
      "selected-summary",
      entry.key ? entry.key + " | " + shortPath(state.payload ? state.payload.keynotePath : "") : "New keynote row"
    );
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

  function addRoot() {
    var entry = {
      id: makeLocalId(),
      key: makeUniqueKey("NEW"),
      text: "New keynote",
      parentKey: "",
      lineNumber: null,
      originalIndex: state.entries.length
    };
    state.entries.push(entry);
    state.selectedId = entry.id;
    markDirty();
    renderAll();
  }

  function addChild() {
    var parent = selectedEntry();
    var entry;

    if (!parent) {
      return;
    }

    entry = {
      id: makeLocalId(),
      key: makeUniqueKey(parent.key + ".A"),
      text: "New keynote",
      parentKey: parent.key,
      lineNumber: null,
      originalIndex: state.entries.length
    };
    state.entries.push(entry);
    state.selectedId = entry.id;
    markDirty();
    renderAll();
  }

  function duplicateSelected() {
    var entry = selectedEntry();
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
    state.selectedId = copy.id;
    markDirty();
    renderAll();
  }

  function deleteSelected() {
    var entry = selectedEntry();
    var nextSelection = null;

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

    if (state.entries.length) {
      nextSelection = state.entries[Math.min(entry.originalIndex || 0, state.entries.length - 1)];
    }
    state.selectedId = nextSelection ? nextSelection.id : null;
    markDirty();
    renderAll();
  }

  function refreshData() {
    if (state.dirty) {
      var shouldRefresh = globalScope.confirm(
        "Refresh will discard unsaved keynote edits in this window. Continue?"
      );
      if (!shouldRefresh) {
        return;
      }
    }

    setStatus({ status: "warning", message: "Refreshing keynote file..." });
    postWebViewMessage({ type: "refreshData" });
  }

  function saveData() {
    var issues = validateAll();
    var payload;

    renderValidation();

    if (hasErrorIssues(issues)) {
      setStatus({ status: "error", message: "Fix validation errors before saving." });
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
    renderSaveState();
    setStatus({ status: "warning", message: "Saving keynote file and reloading Revit..." });

    if (!postWebViewMessage({ type: "saveKeynotes", payload: payload })) {
      state.saving = false;
      renderSaveState();
    }
  }

  function handleSaveResult(result) {
    result = result || {};
    state.saving = false;

    if (result.payload) {
      loadData(result.payload);
    }

    setStatus({
      status: result.status || "idle",
      message: result.backupPath
        ? (result.message || "") + " Backup: " + result.backupPath
        : (result.message || "")
    });

    if (!result.payload) {
      renderValidation();
      renderSaveState();
    }
  }

  function closeWindow() {
    if (state.dirty) {
      var shouldClose = globalScope.confirm(
        "Close the manager and discard unsaved keynote edits in this window?"
      );
      if (!shouldClose) {
        return;
      }
    }

    if (postWebViewMessage({ type: "closeWindow" })) {
      return;
    }

    try {
      globalScope.close();
    } catch (ignore) {
      // Browser fallback only.
    }
  }

  function bindInput(id, fieldName) {
    var input = byId(id);
    if (!input) {
      return;
    }

    input.addEventListener("input", function () {
      updateSelectedField(fieldName, input.value);
    });
  }

  function init() {
    var searchInput = byId("search-input");

    initCollapsibleSections();

    if (searchInput) {
      searchInput.addEventListener("input", function () {
        renderTable();
      });
    }

    bindInput("key-input", "key");
    bindInput("text-input", "text");
    bindInput("parent-input", "parentKey");

    byId("add-root").addEventListener("click", addRoot);
    byId("add-child").addEventListener("click", addChild);
    byId("duplicate-row").addEventListener("click", duplicateSelected);
    byId("delete-row").addEventListener("click", deleteSelected);
    byId("refresh-data").addEventListener("click", refreshData);
    byId("save-data").addEventListener("click", saveData);
    byId("close-window").addEventListener("click", closeWindow);

    renderAll();
    postWebViewMessage({ type: "appReady" });
  }

  globalScope.ffeKeynotes = {
    loadData: loadData,
    setStatus: setStatus,
    handleSaveResult: handleSaveResult
  };

  if (typeof document !== "undefined") {
    document.addEventListener("DOMContentLoaded", init);
  }
}(typeof window !== "undefined" ? window : globalThis));
