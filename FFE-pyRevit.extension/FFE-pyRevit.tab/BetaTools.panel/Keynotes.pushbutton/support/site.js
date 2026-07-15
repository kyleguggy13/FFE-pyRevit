(function attachKeynoteManager(globalScope) {
  "use strict";

  var UNGROUPED_ID = "__ungrouped__";
  var REMOTE_ENTRY_REFRESH_DELAY_MS = 600;
  var EDIT_CLAIM_HEARTBEAT_MS = 60000;
  var activeRowActionMenu = null;

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
    pendingDbChanges: null,
    remoteEditClaims: {},
    remoteEntrySnapshot: null,
    remoteEntriesPending: false,
    remoteEntriesTimer: null,
    localEditClaimSignature: "",
    localEditClaimHeartbeat: null,
    syncIssues: [],
    operationIssues: [],
    analyticsCollecting: false,
    analyticsRequestedOnOpen: false,
    lastAnalyticsResult: null,
    modelHealth: null,
    modelIssuesOpen: false,
    reviewedModelHealthSignature: "",
    modelIssueResolutions: {},
    remotePending: false,
    allowNextLoad: false,
    collapsedEntryIds: {},
    sheetVisibleKeynotes: {},
    placementMode: "userKeynote",
    placementFilter: "all",
    divisionsCollapsed: false,
    warningSidebarOpen: false,
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

  function normalizePlacementMode(value) {
    return value === "genericAnnotation" ? "genericAnnotation" : "userKeynote";
  }

  function normalizePlacementFilter(value) {
    return value === "placed" || value === "unused" ? value : "all";
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

  function naturalCompareText(first, second) {
    var firstParts = trim(first).toLowerCase().match(/\d+|\D+/g) || [""];
    var secondParts = trim(second).toLowerCase().match(/\d+|\D+/g) || [""];
    var length = Math.max(firstParts.length, secondParts.length);
    var index;
    var firstPart;
    var secondPart;
    var firstNumber;
    var secondNumber;

    for (index = 0; index < length; index += 1) {
      firstPart = firstParts[index] || "";
      secondPart = secondParts[index] || "";

      if (/^\d+$/.test(firstPart) && /^\d+$/.test(secondPart)) {
        firstNumber = Number(firstPart);
        secondNumber = Number(secondPart);
        if (firstNumber !== secondNumber) {
          return firstNumber - secondNumber;
        }
        if (firstPart.length !== secondPart.length) {
          return firstPart.length - secondPart.length;
        }
      } else if (firstPart !== secondPart) {
        return firstPart < secondPart ? -1 : 1;
      }
    }

    return 0;
  }

  function compareEntriesByKey(first, second) {
    return naturalCompareText(first && first.key, second && second.key) ||
      naturalCompareText(first && first.text, second && second.text) ||
      ((first && first.originalIndex) || 0) - ((second && second.originalIndex) || 0);
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

  function normalizeSheetVisibleKeynotes(value) {
    var result = {};
    Object.keys(value || {}).forEach(function (key) {
      var normalizedKey = trim(key);
      if (normalizedKey && value[key]) {
        result[normalizedKey] = true;
      }
    });
    return result;
  }

  function normalizeModelHealth(value) {
    value = value || {};
    return {
      status: text(value.status || "notScanned"),
      message: text(value.message || "Model health has not been scanned."),
      scannedAt: text(value.scannedAt),
      safeModeRecommended: Boolean(value.safeModeRecommended),
      signature: text(value.signature),
      placedKeyCount: Number(value.placedKeyCount || 0),
      placedCount: Number(value.placedCount || 0),
      missingKeyCount: Number(value.missingKeyCount || 0),
      missingPlacedCount: Number(value.missingPlacedCount || 0),
      missingRatio: Number(value.missingRatio || 0),
      userKeynoteCount: Number(value.userKeynoteCount || 0),
      genericAnnotationCount: Number(value.genericAnnotationCount || 0),
      sheetCount: Number(value.sheetCount || 0),
      unsheetedCount: Number(value.unsheetedCount || 0),
      skippedCount: Number(value.skippedCount || 0),
      placedKeyMap: normalizeSheetVisibleKeynotes(value.placedKeyMap || {}),
      issues: (value.issues || []).map(function (issue) {
        return {
          severity: text(issue.severity || "warning"),
          code: text(issue.code),
          key: trim(issue.key),
          message: text(issue.message),
          details: text(issue.details),
          placedCount: Number(issue.placedCount || 0),
          userKeynoteCount: Number(issue.userKeynoteCount || 0),
          genericAnnotationCount: Number(issue.genericAnnotationCount || 0),
          sheetCount: Number(issue.sheetCount || 0),
          unsheetedCount: Number(issue.unsheetedCount || 0),
          sheets: issue.sheets || [],
          typeNames: issue.typeNames || [],
          resolution: issue.resolution ? {
            resolutionType: text(issue.resolution.resolutionType),
            familyTypeName: text(issue.resolution.familyTypeName),
            familyTypeText: text(issue.resolution.familyTypeText),
            fileText: text(issue.resolution.fileText)
          } : null
        };
      })
    };
  }

  function currentModelHealth() {
    return state.modelHealth || normalizeModelHealth(null);
  }

  function modelHealthSignature() {
    var health = currentModelHealth();
    if (health.signature) {
      return health.signature;
    }
    return [
      health.status,
      health.safeModeRecommended ? "safe" : "ok",
      health.placedKeyCount,
      health.missingKeyCount,
      health.issues.map(function (issue) {
        return [
          issue.severity,
          issue.code,
          issue.key,
          issue.placedCount,
          issue.resolution && issue.resolution.familyTypeName,
          issue.resolution && issue.resolution.familyTypeText,
          issue.resolution && issue.resolution.fileText
        ].join(":");
      }).join("|")
    ].join("|");
  }

  function modelIssueCount() {
    return currentModelHealth().issues.length;
  }

  function modelIssueResolutionId(issue) {
    var resolution = issue && issue.resolution;
    return [
      text(issue && issue.code),
      trim(issue && issue.key),
      trim(resolution && resolution.familyTypeName)
    ].join("|");
  }

  function inferParentKeyForMissingKey(key) {
    var targetKey = trim(key);
    var bestMatch = "";

    state.entries.forEach(function (entry) {
      var candidateKey = trim(entry.key);
      if (
        candidateKey &&
        targetKey.indexOf(candidateKey + ".") === 0 &&
        candidateKey.length > bestMatch.length
      ) {
        bestMatch = candidateKey;
      }
    });
    return bestMatch;
  }

  function unresolvedModelIssueCount() {
    return currentModelHealth().issues.filter(function (issue) {
      return Boolean(
        issue.resolution &&
        !state.modelIssueResolutions[modelIssueResolutionId(issue)]
      );
    }).length;
  }

  function isModelSafeModeActive() {
    var health = currentModelHealth();
    var signature = modelHealthSignature();
    return Boolean(
      health.safeModeRecommended &&
      signature &&
      state.reviewedModelHealthSignature !== signature
    );
  }

  function modelSafeModeMessage() {
    var health = currentModelHealth();
    if (!health.safeModeRecommended) {
      return "";
    }
    return "Review " + formatNumber(health.missingKeyCount) +
      " missing placed keynote key(s) before editing or saving.";
  }

  function blockForSafeMode(actionLabel) {
    if (!isModelSafeModeActive()) {
      return false;
    }
    setModelIssuesOpen(true);
    setStatus({
      status: "warning",
      message: (actionLabel || "This action") + " is paused until model issues are reviewed."
    });
    renderAll();
    return true;
  }

  function keyIsPlaced(key) {
    return Boolean(state.sheetVisibleKeynotes[trim(key)]);
  }

  function createPlacedKeyBadge(key) {
    var badge;
    var path;
    if (!keyIsPlaced(key)) {
      return null;
    }

    badge = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    badge.setAttribute("class", "sheet-visible-marker");
    badge.setAttribute("viewBox", "0 0 16 16");
    badge.setAttribute("fill", "currentColor");
    badge.setAttribute("focusable", "false");
    badge.setAttribute("role", "img");
    badge.setAttribute("aria-label", "Placed keynote annotation");
    badge.setAttribute("title", "Placed keynote annotation");
    path.setAttribute("d", "M4 0h5.293A1 1 0 0 1 10 .293L13.707 4a1 1 0 0 1 .293.707V14a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V2a2 2 0 0 1 2-2m5.5 1.5v2a1 1 0 0 0 1 1h2z");
    badge.appendChild(path);
    return badge;
  }

  function appendPlacedKeyBadge(parent, key) {
    var badge = createPlacedKeyBadge(key);
    if (badge && parent) {
      parent.classList.add("sheet-marker-anchor");
      parent.appendChild(badge);
    }
    return badge;
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

  function issueHasCode(issue, code) {
    return text(issue && issue.code) === code;
  }

  function clearSyncIssueCode(code) {
    state.syncIssues = (state.syncIssues || []).filter(function (issue) {
      return !issueHasCode(issue, code);
    });
  }

  function upsertSyncIssue(issue) {
    if (!issue || !issue.code) {
      return;
    }
    clearSyncIssueCode(issue.code);
    state.syncIssues.push(issue);
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
    var parentKey = trim(key);
    state.entries.forEach(function (entry) {
      if (trim(entry.parentKey) === parentKey) {
        count += 1;
      }
    });
    return count;
  }

  function entryHasChildren(entry) {
    return Boolean(entry && childCount(entry.key));
  }

  function rootEntries() {
    return state.entries.filter(function (entry) {
      return !trim(entry.parentKey);
    }).sort(compareEntriesByKey);
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
    Object.keys(children).forEach(function (parentKey) {
      children[parentKey].sort(compareEntriesByKey);
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

  function entryIsDescendantOf(entry, ancestor) {
    var byKey = entriesByKey();
    var seen = {};
    var cursor = entry;
    var parent;

    if (!entry || !ancestor || entry.id === ancestor.id) {
      return false;
    }

    while (cursor && cursor.parentKey) {
      parent = byKey[cursor.parentKey];
      if (!parent || seen[parent.id]) {
        return false;
      }
      if (parent.id === ancestor.id) {
        return true;
      }
      seen[parent.id] = true;
      cursor = parent;
    }

    return false;
  }

  function expandAncestorsForEntry(entry) {
    var byKey = entriesByKey();
    var seen = {};
    var cursor = entry;
    var parent;

    while (cursor && cursor.parentKey) {
      parent = byKey[cursor.parentKey];
      if (!parent || seen[parent.id]) {
        return;
      }
      delete state.collapsedEntryIds[parent.id];
      seen[parent.id] = true;
      cursor = parent;
    }
  }

  function setSelectionForEntry(entry) {
    var root = entry ? rootAncestorFor(entry) : null;

    if (!entry) {
      state.selectedDivisionId = null;
      state.selectedNoteId = null;
      state.selectedId = null;
      return;
    }

    expandAncestorsForEntry(entry);

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
      } else if (!currentSearchQuery() && !visibleRowIdMapForModel(selectedModel)[selectedNote.id]) {
        selectedNote = collapsedAncestorForEntry(selectedNote, selectedModel);
        state.selectedNoteId = selectedNote ? selectedNote.id : null;
      }
    }

    state.selectedId = state.selectedNoteId || (selectedModel.entry ? selectedModel.entry.id : null);
    return selectedModel;
  }

  function currentSearchQuery() {
    return trim(byId("search-input") ? byId("search-input").value : "").toLowerCase();
  }

  function currentPlacementFilter() {
    return normalizePlacementFilter(state.placementFilter);
  }

  function currentRemoteEntrySnapshot() {
    if (!state.remoteEntriesPending || !state.remoteEntrySnapshot) {
      return null;
    }
    return state.remoteEntrySnapshot;
  }

  function remoteClaimForKey(key) {
    return state.remoteEditClaims["key:" + trim(key)] || null;
  }

  function localEntryNeedsKeyReservationCheck(entry) {
    var baselineEntry = findBaselineEntry(entry);
    var key = trim(entry && entry.key);

    if (!key) {
      return false;
    }
    return !baselineEntry || key !== trim(baselineEntry.key);
  }

  function localEntryMatchesRemoteSavedEntry(entry, remoteEntry) {
    var baselineEntry = findBaselineEntry(entry);
    var baselineDbId = text(baselineEntry && baselineEntry.dbId);
    var remoteDbId = text(remoteEntry && (remoteEntry.dbId || remoteEntry.id));

    if (baselineDbId && remoteDbId) {
      return baselineDbId === remoteDbId;
    }
    return Boolean(
      baselineEntry &&
      trim(baselineEntry.key) === trim(remoteEntry && remoteEntry.key) &&
      trim(entry && entry.key) === trim(remoteEntry && remoteEntry.key)
    );
  }

  function appendRealtimeConflictIssues(issues) {
    var remoteSnapshot = currentRemoteEntrySnapshot();
    var remoteByKey = remoteSnapshot ? indexEntriesByKey(remoteSnapshot.entries) : {};

    state.entries.forEach(function (entry) {
      var key = trim(entry.key);
      var claim;
      var remoteEntry;
      var clientName;

      if (!localEntryNeedsKeyReservationCheck(entry)) {
        return;
      }

      claim = remoteClaimForKey(key);
      if (claim) {
        clientName = trim(claim.clientName) || "another user";
        issues.push(makeIssue(
          "error",
          "Key '" + key + "' is currently reserved by " + clientName + ". Choose another key or wait for them to save/close.",
          key,
          "remoteReservedKey"
        ));
      }

      remoteEntry = remoteByKey[key];
      if (remoteEntry && !localEntryMatchesRemoteSavedEntry(entry, remoteEntry)) {
        issues.push(makeIssue(
          "error",
          "Key '" + key + "' was added by another user. Refresh before saving or choose another key.",
          key,
          "remoteSavedKey"
        ));
      }
    });
  }

  function validateAll() {
    var issues = state.sourceIssues.concat(state.operationIssues || [], state.syncIssues || []);
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

    appendRealtimeConflictIssues(issues);

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

  function entryMatchesQuery(entry, query) {
    var haystack = (entry.key + " " + entry.text + " " + entry.parentKey).toLowerCase();
    return !query || haystack.indexOf(query) !== -1;
  }

  function rowMatchesQuery(row, query) {
    return entryMatchesQuery(row.entry, query);
  }

  function entryMatchesPlacementFilter(entry, placementFilter) {
    placementFilter = normalizePlacementFilter(placementFilter);
    if (placementFilter === "placed") {
      return keyIsPlaced(entry && entry.key);
    }
    if (placementFilter === "unused") {
      return !keyIsPlaced(entry && entry.key);
    }
    return true;
  }

  function rowMatchesFilters(row, query, placementFilter) {
    return rowMatchesQuery(row, query) && entryMatchesPlacementFilter(row.entry, placementFilter);
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

  function makeTreeRow(row, options) {
    var hasChildren = entryHasChildren(row.entry);
    var isSearchMode = Boolean(options && options.isSearchMode);
    var isSearchMatch = Boolean(options && options.matchIds && options.matchIds[row.entry.id]);
    var isSearchContext = Boolean(
      options &&
      options.contextIds &&
      options.contextIds[row.entry.id] &&
      !isSearchMatch
    );

    return {
      entry: row.entry,
      depth: row.depth,
      hasChildren: hasChildren,
      isCollapsed: hasChildren && !isSearchMode && Boolean(state.collapsedEntryIds[row.entry.id]),
      isSearchMatch: isSearchMatch,
      isSearchContext: isSearchContext
    };
  }

  function expandedRowsForModel(model) {
    var rows = [];
    var hiddenDepth = null;

    (model.rows || []).forEach(function (row) {
      var treeRow;

      if (hiddenDepth !== null) {
        if (row.depth > hiddenDepth) {
          return;
        }
        hiddenDepth = null;
      }

      treeRow = makeTreeRow(row, {});
      rows.push(treeRow);

      if (treeRow.isCollapsed) {
        hiddenDepth = row.depth;
      }
    });

    return rows;
  }

  function searchRowsForModel(model, query) {
    return filteredRowsForModel(model, query, "all");
  }

  function filteredRowsForModel(model, query, placementFilter) {
    var byKey = entriesByKey();
    var rowIds = {};
    var matchIds = {};
    var contextIds = {};
    var includeIds = {};
    placementFilter = normalizePlacementFilter(placementFilter);

    (model.rows || []).forEach(function (row) {
      rowIds[row.entry.id] = true;
    });

    (model.rows || []).forEach(function (row) {
      var cursor;
      var parent;
      var seen = {};

      if (!rowMatchesFilters(row, query, placementFilter)) {
        return;
      }

      matchIds[row.entry.id] = true;
      includeIds[row.entry.id] = true;
      cursor = row.entry;

      while (cursor && cursor.parentKey) {
        parent = byKey[cursor.parentKey];
        if (!parent || seen[parent.id]) {
          return;
        }

        if (rowIds[parent.id]) {
          includeIds[parent.id] = true;
          contextIds[parent.id] = true;
        }

        seen[parent.id] = true;
        cursor = parent;
      }
    });

    return (model.rows || []).filter(function (row) {
      return includeIds[row.entry.id];
    }).map(function (row) {
      return makeTreeRow(row, {
        isSearchMode: true,
        matchIds: matchIds,
        contextIds: contextIds
      });
    });
  }

  function visibleRowIdMapForModel(model) {
    var ids = {};
    expandedRowsForModel(model).forEach(function (row) {
      ids[row.entry.id] = true;
    });
    return ids;
  }

  function collapsedAncestorForEntry(entry, model) {
    var byKey = entriesByKey();
    var modelRowIds = {};
    var cursor = entry;
    var parent;
    var result = null;
    var seen = {};

    (model.rows || []).forEach(function (row) {
      modelRowIds[row.entry.id] = true;
    });

    while (cursor && cursor.parentKey) {
      parent = byKey[cursor.parentKey];
      if (!parent || seen[parent.id]) {
        return result;
      }

      if (modelRowIds[parent.id] && state.collapsedEntryIds[parent.id]) {
        result = parent;
      }

      seen[parent.id] = true;
      cursor = parent;
    }

    return result;
  }

  function selectedRows() {
    var model = ensureSelection();
    var query = currentSearchQuery();
    var placementFilter = currentPlacementFilter();

    if (!model) {
      return [];
    }

    if (query || placementFilter !== "all") {
      return filteredRowsForModel(model, query, placementFilter);
    }

    return expandedRowsForModel(model);
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
    var dirtyElement = byId("dirty-state");

    state.dirty = nextDirty;
    setText(dirtyElement, state.dirty ? "UNSAVED CHANGES" : "NO CHANGES");
    if (dirtyElement) {
      dirtyElement.setAttribute("data-dirty", state.dirty ? "true" : "false");
    }
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
      var item = document.createElement("div");
      var copy = document.createElement("span");
      var title = document.createElement("strong");
      var subtitle = document.createElement("span");
      var badge = document.createElement("span");
      var actionControls = document.createElement("span");
      var menuButton;
      var isSelected = selectedModel && model.id === selectedModel.id;
      var claimTitle = model.entry ? editClaimTitle(model.entry) : "";

      item.className = "list-group-item list-group-item-action division-row" +
        (isSelected ? " active is-selected" : "") +
        (claimTitle ? " is-claimed" : "");
      item.setAttribute("role", "option");
      item.setAttribute("aria-selected", isSelected ? "true" : "false");
      item.setAttribute("data-entry-id", model.id);
      item.setAttribute("title", claimTitle || (model.title + " - " + model.text));
      item.tabIndex = -1;

      copy.className = "division-row-copy";
      title.textContent = model.title;
      appendPlacedKeyBadge(title, model.entry ? model.entry.key : "");
      subtitle.textContent = model.text || "Untitled division";
      badge.className = "badge rounded-pill division-note-badge";
      badge.textContent = formatNumber(model.rows.length);
      badge.setAttribute("aria-label", formatNumber(model.rows.length) + (model.rows.length === 1 ? " note" : " notes"));
      actionControls.className = "division-row-actions";

      copy.appendChild(title);
      copy.appendChild(subtitle);
      item.appendChild(copy);
      actionControls.appendChild(badge);
      if (model.entry) {
        menuButton = document.createElement("button");
        menuButton.type = "button";
        menuButton.className = "division-more-button";
        menuButton.appendChild(createEllipsisVerticalIcon());
        menuButton.setAttribute("aria-label", "More actions for parent " + (model.entry.key || "new parent"));
        menuButton.setAttribute("aria-haspopup", "menu");
        menuButton.setAttribute("aria-expanded", "false");
        menuButton.setAttribute("title", "More parent actions");
        menuButton.addEventListener("click", function (event) {
          event.preventDefault();
          event.stopPropagation();
          openParentActionMenu(model.entry, menuButton, false);
        });
        actionControls.appendChild(menuButton);
      }
      item.appendChild(actionControls);
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
      var label = document.createElement("span");
      var isSelected = selectedModel && model.id === selectedModel.id;
      var claimTitle = model.entry ? editClaimTitle(model.entry) : "";

      button.type = "button";
      button.className = "dropdown-item" + (isSelected ? " active" : "") + (claimTitle ? " is-claimed" : "");
      button.setAttribute("data-division-id", model.id);
      button.setAttribute("title", claimTitle || (model.title + " - " + model.text));
      if (isSelected) {
        button.setAttribute("aria-current", "true");
      }
      label.textContent = model.title + " - " + model.text;
      button.appendChild(label);
      appendPlacedKeyBadge(button, model.entry ? model.entry.key : "");

      item.appendChild(button);
      menu.appendChild(item);
    });
  }

  function renderDivisionHeader() {
    var model = selectedDivisionModel();
    var keyInput = byId("division-key-input");
    var textInput = byId("division-text-input");
    var menuButton = byId("division-more-button");
    var hasEntry = Boolean(model && model.entry);
    var claimTitle = hasEntry ? editClaimTitle(model.entry) : "";
    var safeMode = isModelSafeModeActive();
    var card = document.querySelector(".selected-division-card");
    var keyField = document.querySelector(".division-key-field");
    var existingBadge = keyField ? keyField.querySelector(".sheet-visible-marker") : null;

    if (existingBadge) {
      existingBadge.parentNode.removeChild(existingBadge);
    }
    if (keyField) {
      keyField.classList.remove("sheet-marker-anchor");
    }

    if (card) {
      card.classList.toggle("is-claimed", Boolean(claimTitle));
      card.setAttribute("title", claimTitle || "");
    }

    if (keyInput) {
      keyInput.disabled = safeMode || !hasEntry || Boolean(claimTitle);
      lockKeyInput(keyInput);
      keyInput.value = hasEntry ? model.entry.key : "UNGROUPED";
      keyInput.setAttribute("title", safeMode ? "Review model issues before editing." : (claimTitle || "Double-click to edit division key"));
    }
    if (keyField && hasEntry) {
      appendPlacedKeyBadge(keyField, model.entry.key);
    }
    if (textInput) {
      textInput.disabled = safeMode || !hasEntry || Boolean(claimTitle);
      textInput.value = hasEntry ? model.entry.text : "Keynotes without a root division";
      textInput.setAttribute("title", safeMode ? "Review model issues before editing." : (claimTitle || ""));
    }
    if (menuButton) {
      menuButton.disabled = !hasEntry;
      menuButton.setAttribute(
        "aria-label",
        hasEntry ? "More actions for parent " + (model.entry.key || "new parent") : "No parent actions available"
      );
    }
  }

  function lockKeyInput(input) {
    if (!input) {
      return;
    }

    input.readOnly = true;
    input.setAttribute("aria-readonly", "true");
    input.classList.remove("is-key-editing");
  }

  function unlockKeyInput(input) {
    if (!input || input.disabled) {
      return;
    }

    input.readOnly = false;
    input.setAttribute("aria-readonly", "false");
    input.classList.add("is-key-editing");
    input.focus();
    if (typeof input.select === "function") {
      input.select();
    }
  }

  function createPlaceKeynoteIcon() {
    var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    var path = document.createElementNS("http://www.w3.org/2000/svg", "path");

    svg.setAttribute("class", "bi bi-arrow-right-square-fill");
    svg.setAttribute("viewBox", "0 0 16 16");
    svg.setAttribute("fill", "currentColor");
    svg.setAttribute("focusable", "false");
    svg.setAttribute("aria-hidden", "true");
    path.setAttribute("d", "M0 14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V2a2 2 0 0 0-2-2H2a2 2 0 0 0-2 2zm4.5-6.5h5.793L8.146 5.354a.5.5 0 1 1 .708-.708l3 3a.5.5 0 0 1 0 .708l-3 3a.5.5 0 0 1-.708-.708L10.293 8.5H4.5a.5.5 0 0 1 0-1");
    svg.appendChild(path);
    return svg;
  }

  function createEllipsisVerticalIcon() {
    var svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");

    svg.setAttribute("class", "lucide lucide-ellipsis-vertical");
    svg.setAttribute("viewBox", "0 0 24 24");
    svg.setAttribute("fill", "none");
    svg.setAttribute("stroke", "currentColor");
    svg.setAttribute("stroke-width", "2");
    svg.setAttribute("stroke-linecap", "round");
    svg.setAttribute("stroke-linejoin", "round");
    svg.setAttribute("focusable", "false");
    svg.setAttribute("aria-hidden", "true");

    [5, 12, 19].forEach(function (cy) {
      var circle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
      circle.setAttribute("cx", "12");
      circle.setAttribute("cy", String(cy));
      circle.setAttribute("r", "1");
      svg.appendChild(circle);
    });
    return svg;
  }

  function closeRowActionMenu(restoreFocus) {
    var trigger;

    if (!activeRowActionMenu) {
      return;
    }

    trigger = activeRowActionMenu.trigger;
    if (trigger) {
      trigger.setAttribute("aria-expanded", "false");
    }
    if (activeRowActionMenu.element && activeRowActionMenu.element.parentNode) {
      activeRowActionMenu.element.parentNode.removeChild(activeRowActionMenu.element);
    }
    activeRowActionMenu = null;

    if (restoreFocus && trigger && document.documentElement.contains(trigger)) {
      trigger.focus();
    }
  }

  function positionRowActionMenu(menu, trigger) {
    var triggerRect;
    var menuRect;
    var left;
    var top;
    var viewportPadding = 8;

    if (!menu || !trigger) {
      return;
    }

    menu.style.visibility = "hidden";
    menu.style.left = "0px";
    menu.style.top = "0px";
    triggerRect = trigger.getBoundingClientRect();
    menuRect = menu.getBoundingClientRect();
    left = triggerRect.right - menuRect.width;
    top = triggerRect.bottom + 4;

    if (left < viewportPadding) {
      left = viewportPadding;
    }
    if (left + menuRect.width > globalScope.innerWidth - viewportPadding) {
      left = globalScope.innerWidth - menuRect.width - viewportPadding;
    }
    if (top + menuRect.height > globalScope.innerHeight - viewportPadding) {
      top = triggerRect.top - menuRect.height - 4;
    }
    if (top < viewportPadding) {
      top = viewportPadding;
    }

    menu.style.left = Math.round(left) + "px";
    menu.style.top = Math.round(top) + "px";
    menu.style.visibility = "visible";
  }

  function createRowActionButton(label, handler, options) {
    var button = document.createElement("button");

    options = options || {};
    button.type = "button";
    button.className = 
    "row-action-menu-item" + 
    (options.danger ? " is-danger" : "") + 
    (options.addition ? " is-addition" : "");
    button.textContent = label;
    button.setAttribute("role", "menuitem");
    button.disabled = Boolean(options.disabled);
    if (options.title) {
      button.setAttribute("title", options.title);
    }
    button.addEventListener("click", function (event) {
      event.preventDefault();
      event.stopPropagation();
      if (button.disabled) {
        return;
      }
      handler();
    });
    return button;
  }

  function copyNoteText(entry) {
    var noteText = text(entry && entry.text);
    var clipboard = globalScope.navigator && globalScope.navigator.clipboard;
    var copyPromise;

    function fallbackCopy() {
      var textarea = document.createElement("textarea");
      var copied;

      textarea.value = noteText;
      textarea.setAttribute("readonly", "readonly");
      textarea.className = "clipboard-copy-source";
      document.body.appendChild(textarea);
      textarea.select();
      copied = document.execCommand && document.execCommand("copy");
      document.body.removeChild(textarea);
      if (!copied) {
        throw new Error("Clipboard copy was not available.");
      }
    }

    closeRowActionMenu(false);
    try {
      copyPromise = clipboard && typeof clipboard.writeText === "function"
        ? clipboard.writeText(noteText)
        : Promise.resolve().then(fallbackCopy);
      copyPromise.then(function () {
        setStatus({
          status: "ready",
          message: "Copied text for keynote '" + (entry.key || "new keynote") + "'."
        });
      }).catch(function () {
        try {
          fallbackCopy();
          setStatus({
            status: "ready",
            message: "Copied text for keynote '" + (entry.key || "new keynote") + "'."
          });
        } catch (ignore) {
          setStatus({ status: "error", message: "Could not copy the keynote text to the clipboard." });
        }
      });
    } catch (ignore) {
      setStatus({ status: "error", message: "Could not copy the keynote text to the clipboard." });
    }
  }

  function renderMoveToDivisionMenu(menu, entry, trigger) {
    var heading = document.createElement("div");
    var currentRoot = rootAncestorFor(entry);
    var roots = rootEntries();

    clearElement(menu);
    menu.classList.add("is-division-picker");
    menu.appendChild(createRowActionButton("Back", function () {
      renderRowActionMenuContent(menu, entry, trigger);
    }));

    heading.className = "row-action-menu-heading";
    heading.textContent = "Move Note to Division";
    menu.appendChild(heading);

    roots.forEach(function (division) {
      var isDirectParent = trim(entry.parentKey) === trim(division.key);
      var isCurrentDivision = currentRoot && currentRoot.id === division.id;
      var label = divisionTitle(division) + " - " + (division.text || "Untitled division");
      if (isCurrentDivision) {
        label += " (current)";
      }
      menu.appendChild(createRowActionButton(label, function () {
        closeRowActionMenu(false);
        moveNoteToDivision(entry, division);
      }, {
        disabled: isDirectParent,
        title: isDirectParent ? "This note is already directly under this division." : ""
      }));
    });

    if (!roots.length) {
      heading = document.createElement("div");
      heading.className = "row-action-menu-empty";
      heading.textContent = "No divisions are available.";
      menu.appendChild(heading);
    }

    positionRowActionMenu(menu, trigger);
    if (menu.querySelector(".row-action-menu-item:not(:disabled)")) {
      menu.querySelector(".row-action-menu-item:not(:disabled)").focus();
    }
  }

  function renderRowActionMenuContent(menu, entry, trigger) {
    var safeMode = isModelSafeModeActive();
    var claimed = isEntryRemotelyClaimed(entry);
    var parentKey = trim(entry && entry.parentKey);
    var sequenceParent = parentKey ? entriesByKey()[parentKey] || null : null;
    var sequenceParentClaimed = sequenceParent && isEntryRemotelyClaimed(sequenceParent);
    var editingDisabled = safeMode || claimed;
    var sequenceDisabled = safeMode || !sequenceParent || sequenceParentClaimed;
    var disabledTitle = safeMode
      ? "Review model issues before editing."
      : (claimed ? editClaimTitle(entry) : "");
    var sequenceDisabledTitle = safeMode
      ? "Review model issues before editing."
      : (sequenceParentClaimed
        ? editClaimTitle(sequenceParent)
        : (!sequenceParent ? "This keynote does not have a valid parent." : ""));

    clearElement(menu);
    menu.classList.remove("is-division-picker");
    menu.appendChild(createRowActionButton("Copy Text", function () {
      copyNoteText(entry);
    }));
    menu.appendChild(createRowActionButton("Delete Note", function () {
      closeRowActionMenu(false);
      deleteEntry(entry);
    }, { disabled: editingDisabled, title: disabledTitle, danger: true }));
    menu.appendChild(createRowActionButton("Move Note to Division", function () {
      renderMoveToDivisionMenu(menu, entry, trigger);
    }, { disabled: editingDisabled, title: disabledTitle }));
    menu.appendChild(createRowActionButton("Promote to Parent", function () {
      closeRowActionMenu(false);
      promoteNoteToParent(entry);
    }, { disabled: editingDisabled, title: disabledTitle }));
    menu.appendChild(createRowActionButton("Add Note in Sequence", function () {
      closeRowActionMenu(false);
      addNoteInSequenceForEntry(entry);
    }, { disabled: sequenceDisabled, title: sequenceDisabledTitle, addition: true }));
    menu.appendChild(createRowActionButton("UPPER CASE", function () {
      closeRowActionMenu(false);
      uppercaseNoteText(entry);
    }, { disabled: editingDisabled, title: disabledTitle }));
    positionRowActionMenu(menu, trigger);
    if (menu.querySelector(".row-action-menu-item:not(:disabled)")) {
      menu.querySelector(".row-action-menu-item:not(:disabled)").focus();
    }
  }

  function openRowActionMenu(entry, trigger) {
    var menu;

    if (
      activeRowActionMenu &&
      activeRowActionMenu.entryId === entry.id &&
      activeRowActionMenu.menuKind === "note"
    ) {
      closeRowActionMenu(true);
      return;
    }

    closeRowActionMenu(false);
    selectNote(entry.id, false);
    menu = document.createElement("div");
    menu.className = "row-action-menu";
    menu.id = "row-action-menu";
    menu.setAttribute("role", "menu");
    menu.setAttribute("aria-label", "Actions for keynote " + (entry.key || "new keynote"));
    menu.addEventListener("click", function (event) {
      event.stopPropagation();
    });
    document.body.appendChild(menu);
    activeRowActionMenu = {
      element: menu,
      entryId: entry.id,
      menuKind: "note",
      trigger: trigger
    };
    trigger.setAttribute("aria-expanded", "true");
    trigger.setAttribute("aria-controls", menu.id);
    renderRowActionMenuContent(menu, entry, trigger);
  }

  function descendantEntriesFor(entry) {
    if (!entry) {
      return [];
    }
    return state.entries.filter(function (candidate) {
      return entryIsDescendantOf(candidate, entry);
    });
  }

  function directChildEntriesFor(entry) {
    var parentKey = trim(entry && entry.key);
    return state.entries.filter(function (candidate) {
      return trim(candidate.parentKey) === parentKey;
    });
  }

  function firstRemotelyClaimedEntry(entries) {
    var index;
    for (index = 0; index < entries.length; index += 1) {
      if (isEntryRemotelyClaimed(entries[index])) {
        return entries[index];
      }
    }
    return null;
  }

  function renderParentDestinationMenu(menu, entry, trigger, mode) {
    var heading = document.createElement("div");
    var roots = rootEntries().filter(function (candidate) {
      return candidate.id !== entry.id;
    });
    var isDeleteMove = mode === "deleteMove";

    clearElement(menu);
    menu.classList.add("is-division-picker");
    menu.appendChild(createRowActionButton("Back", function () {
      if (isDeleteMove) {
        renderDeleteParentMenu(menu, entry, trigger);
      } else {
        renderParentActionMenuContent(menu, entry, trigger);
      }
    }));

    heading.className = "row-action-menu-heading";
    heading.textContent = isDeleteMove ? "Move Subnotes to Parent" : "Demote Under Parent";
    menu.appendChild(heading);

    roots.forEach(function (destination) {
      var label = divisionTitle(destination) + " - " + (destination.text || "Untitled division");
      menu.appendChild(createRowActionButton(label, function () {
        closeRowActionMenu(false);
        if (isDeleteMove) {
          deleteParentAndMoveSubnotes(entry, destination);
        } else {
          demoteParentToNote(entry, destination);
        }
      }));
    });

    if (!roots.length) {
      heading = document.createElement("div");
      heading.className = "row-action-menu-empty";
      heading.textContent = "Create another parent before using this action.";
      menu.appendChild(heading);
    }

    positionRowActionMenu(menu, trigger);
    if (menu.querySelector(".row-action-menu-item:not(:disabled)")) {
      menu.querySelector(".row-action-menu-item:not(:disabled)").focus();
    }
  }

  function renderDeleteParentMenu(menu, entry, trigger) {
    var descendants = descendantEntriesFor(entry);
    var directChildren = directChildEntriesFor(entry);
    var safeMode = isModelSafeModeActive();
    var deleteAllClaimedEntry = firstRemotelyClaimedEntry([entry].concat(descendants));
    var moveClaimedEntry = firstRemotelyClaimedEntry([entry].concat(directChildren));
    var alternateParentCount = rootEntries().filter(function (candidate) {
      return candidate.id !== entry.id;
    }).length;
    var heading = document.createElement("div");
    var summary = document.createElement("div");
    var deleteLabel = descendants.length
      ? "Delete Parent + All " + formatNumber(descendants.length) + " Subnotes"
      : "Delete Parent";
    var deleteDisabledTitle = safeMode
      ? "Review model issues before editing."
      : (deleteAllClaimedEntry ? editClaimTitle(deleteAllClaimedEntry) : "");
    var moveDisabledTitle = safeMode
      ? "Review model issues before editing."
      : (moveClaimedEntry
        ? editClaimTitle(moveClaimedEntry)
        : (!alternateParentCount ? "Create another parent before moving these subnotes." : ""));

    clearElement(menu);
    menu.classList.remove("is-division-picker");
    menu.appendChild(createRowActionButton("Back", function () {
      renderParentActionMenuContent(menu, entry, trigger);
    }));

    heading.className = "row-action-menu-heading";
    heading.textContent = "Delete Parent " + (entry.key || "");
    menu.appendChild(heading);

    summary.className = "row-action-menu-summary";
    summary.textContent = descendants.length
      ? "Choose whether to delete every subnote or move them before deleting this parent."
      : "This parent has no subnotes.";
    menu.appendChild(summary);

    menu.appendChild(createRowActionButton(deleteLabel, function () {
      closeRowActionMenu(false);
      deleteParentAndSubnotes(entry);
    }, {
      disabled: safeMode || Boolean(deleteAllClaimedEntry),
      title: deleteDisabledTitle,
      danger: true
    }));

    if (descendants.length) {
      menu.appendChild(createRowActionButton("Move Subnotes to Another Parent", function () {
        renderParentDestinationMenu(menu, entry, trigger, "deleteMove");
      }, {
        disabled: safeMode || Boolean(moveClaimedEntry) || !alternateParentCount,
        title: moveDisabledTitle
      }));
    }

    positionRowActionMenu(menu, trigger);
    if (menu.querySelector(".row-action-menu-item:not(:disabled)")) {
      menu.querySelector(".row-action-menu-item:not(:disabled)").focus();
    }
  }

  function renderParentActionMenuContent(menu, entry, trigger) {
    var safeMode = isModelSafeModeActive();
    var claimed = isEntryRemotelyClaimed(entry);
    var editingDisabled = safeMode || claimed;
    var alternateParentCount = rootEntries().filter(function (candidate) {
      return candidate.id !== entry.id;
    }).length;
    var disabledTitle = safeMode
      ? "Review model issues before editing."
      : (claimed ? editClaimTitle(entry) : "");

    clearElement(menu);
    menu.classList.remove("is-division-picker");
    menu.appendChild(createRowActionButton("Copy Text", function () {
      copyNoteText(entry);
    }));
    menu.appendChild(createRowActionButton("Demote to a Note", function () {
      renderParentDestinationMenu(menu, entry, trigger, "demote");
    }, {
      disabled: editingDisabled || !alternateParentCount,
      title: disabledTitle || (!alternateParentCount ? "Create another parent before demoting this parent." : "")
    }));
    menu.appendChild(createRowActionButton("Delete Parent", function () {
      renderDeleteParentMenu(menu, entry, trigger);
    }, { disabled: editingDisabled, title: disabledTitle, danger: true }));
    positionRowActionMenu(menu, trigger);
    if (menu.querySelector(".row-action-menu-item:not(:disabled)")) {
      menu.querySelector(".row-action-menu-item:not(:disabled)").focus();
    }
  }

  function openParentActionMenu(entry, trigger, showDeleteOptions) {
    var menu;

    if (!entry || trim(entry.parentKey)) {
      return;
    }
    if (
      activeRowActionMenu &&
      activeRowActionMenu.entryId === entry.id &&
      activeRowActionMenu.menuKind === "parent"
    ) {
      closeRowActionMenu(true);
      return;
    }

    closeRowActionMenu(false);
    menu = document.createElement("div");
    menu.className = "row-action-menu";
    menu.id = "parent-action-menu";
    menu.setAttribute("role", "menu");
    menu.setAttribute("aria-label", "Actions for parent " + (entry.key || "new parent"));
    menu.addEventListener("click", function (event) {
      event.stopPropagation();
    });
    document.body.appendChild(menu);
    activeRowActionMenu = {
      element: menu,
      entryId: entry.id,
      menuKind: "parent",
      trigger: trigger
    };
    trigger.setAttribute("aria-expanded", "true");
    trigger.setAttribute("aria-controls", menu.id);
    if (showDeleteOptions) {
      renderDeleteParentMenu(menu, entry, trigger);
    } else {
      renderParentActionMenuContent(menu, entry, trigger);
    }
  }

  function placementModeLabel() {
    return state.placementMode === "genericAnnotation" ? "Generic Annotation keynote" : "User Keynote";
  }

  function syncPlacementModeSelect() {
    var placementModeSelect = byId("placement-mode-select");
    if (placementModeSelect) {
      placementModeSelect.value = state.placementMode;
    }
  }

  function syncPlacementFilterSelect() {
    var placementFilterSelect = byId("placement-filter-select");
    if (placementFilterSelect) {
      placementFilterSelect.value = currentPlacementFilter();
    }
  }

  function setPlacementMode(value) {
    state.placementMode = normalizePlacementMode(value);
    if (state.payload) {
      state.payload.preferences = state.payload.preferences || {};
      state.payload.preferences.placementMode = state.placementMode;
    }
    syncPlacementModeSelect();
  }

  function setPlacementFilter(value) {
    state.placementFilter = normalizePlacementFilter(value);
    syncPlacementFilterSelect();
  }

  function entryCanPlaceKeynote(entry) {
    var baselineEntry = findBaselineEntry(entry);
    var key = trim(entry && entry.key);

    return Boolean(
      key &&
      baselineEntry &&
      entryFieldsEqual(entry, baselineEntry) &&
      remoteSnapshotAllowsPlacement(entry)
    );
  }

  function remoteSnapshotAllowsPlacement(entry) {
    var remoteSnapshot = currentRemoteEntrySnapshot();
    var key = trim(entry && entry.key);
    var dbId = text(entry && entry.dbId);
    var remoteEntry;

    if (!key) {
      return true;
    }
    if (state.remoteEntriesPending && !remoteSnapshot) {
      return false;
    }
    if (!remoteSnapshot) {
      return true;
    }

    if (dbId) {
      remoteEntry = indexEntriesByDbId(remoteSnapshot.entries)[dbId];
      return Boolean(remoteEntry && trim(remoteEntry.key) === key);
    }

    return Boolean(indexEntriesByKey(remoteSnapshot.entries)[key]);
  }

  function placementBlockedByRemoteDelete(entry) {
    return Boolean(
      state.remoteEntriesPending &&
      entry &&
      trim(entry.key) &&
      !remoteSnapshotAllowsPlacement(entry)
    );
  }

  function placeKeynote(entry) {
    var key = trim(entry && entry.key);
    var messageType = state.placementMode === "genericAnnotation"
      ? "placeGenericAnnotation"
      : "placeUserKeynote";
    var label = placementModeLabel();

    if (!entry) {
      return;
    }

    selectNote(entry.id, false);

    if (blockForSafeMode("Placement")) {
      return;
    }

    if (!key) {
      setStatus({
        status: "warning",
        message: "Save this keynote with a key before placing it in Revit."
      });
      return;
    }

    if (!entryCanPlaceKeynote(entry)) {
      if (placementBlockedByRemoteDelete(entry)) {
        setStatus({
          status: "warning",
          message: "Refresh before placing this keynote. It may have been deleted by another user."
        });
        return;
      }
      setStatus({
        status: "warning",
        message: "Save this keynote before placing it in Revit."
      });
      return;
    }

    if (postWebViewMessage({
      type: messageType,
      payload: {
        id: entry.id,
        key: key
      }
    })) {
      setStatus({
        status: "warning",
        message: "Starting Revit " + label + " placement for key '" + key + "'..."
      });
    }
  }

  // function resizeNoteTextInput(input) {
  //   var borderSize;

  //   if (!input) {
  //     return;
  //   }

  //   input.style.height = "auto";
  //   borderSize = input.offsetHeight - input.clientHeight;
  //   input.style.height = (input.scrollHeight + borderSize) + "px";
  // }

  function toggleTreeRow(entryId) {
    var entry = findEntryById(entryId);
    var selected = selectedNoteEntry();

    if (!entry || !entryHasChildren(entry)) {
      return;
    }

    if (state.collapsedEntryIds[entry.id]) {
      delete state.collapsedEntryIds[entry.id];
    } else {
      state.collapsedEntryIds[entry.id] = true;
      if (selected && entryIsDescendantOf(selected, entry)) {
        setSelectionForEntry(entry);
      }
    }

    if (!selected || selected.id !== entry.id) {
      setSelectionForEntry(entry);
    }

    renderAll();
  }

  function renderNotes() {
    var body = byId("keynote-table-body");
    var rows = selectedRows();
    var placementFilter = currentPlacementFilter();

    if (!body) {
      return;
    }

    closeRowActionMenu(false);
    clearElement(body);

    if (!rows.length) {
      var emptyRow = document.createElement("tr");
      var empty = document.createElement("td");
      empty.className = "empty-cell";
      empty.setAttribute("colspan", "3");
      if (!state.entries.length) {
        empty.textContent = "No keynote notes loaded.";
      } else if (placementFilter === "placed") {
        empty.textContent = "No placed keynotes in this division.";
      } else if (placementFilter === "unused") {
        empty.textContent = "No unused keynotes in this division.";
      } else {
        empty.textContent = "No notes in this division.";
      }
      emptyRow.appendChild(empty);
      body.appendChild(emptyRow);
      return;
    }

    rows.forEach(function (row) {
      var entry = row.entry;
      var item = document.createElement("tr");
      var actionCell = document.createElement("td");
      var actionWrap = document.createElement("div");
      var keyCell = document.createElement("td");
      var textCell = document.createElement("td");
      var keyWrap = document.createElement("div");
      var placeButton = document.createElement("button");
      var menuButton = document.createElement("button");
      var treeControl;
      var keyInput = document.createElement("input");
      var textInput = document.createElement("textarea");
      var claimTitle = editClaimTitle(entry);
      var canPlaceKeynote = entryCanPlaceKeynote(entry);
      var placementLabel = placementModeLabel();
      var safeMode = isModelSafeModeActive();
      var placeTitle = canPlaceKeynote
        ? "Place " + placementLabel + " " + (entry.key || "")
        : "Save this keynote before placing it in Revit";
      if (safeMode) {
        placeTitle = "Review model issues before placing keynotes.";
      }

      item.className = "note-row" +
        (row.hasChildren ? " is-parent-row" : "") +
        (row.isCollapsed ? " is-collapsed" : "") +
        (row.isSearchContext ? " is-search-context" : "") +
        (row.isSearchMatch ? " is-search-match" : "") +
        (entry.id === state.selectedNoteId ? " is-selected" : "") +
        (claimTitle ? " is-claimed" : "");
      item.setAttribute("data-entry-id", entry.id);
      item.setAttribute("title", claimTitle || "");
      item.setAttribute("aria-level", String(row.depth + 1));
      if (row.hasChildren) {
        item.setAttribute("aria-expanded", row.isCollapsed ? "false" : "true");
      }
      item.tabIndex = -1;
      item.style.setProperty("--depth", row.depth);

      actionCell.className = "note-cell note-action-cell";
      actionWrap.className = "note-action-wrap";
      keyCell.className = "note-cell note-key-cell";
      textCell.className = "note-cell note-text-cell";
      keyWrap.className = "note-key-wrap";
      keyWrap.style.setProperty("--depth", row.depth);

      placeButton.type = "button";
      placeButton.className = "note-place-button" + (canPlaceKeynote && !safeMode ? "" : " needs-save");
      placeButton.setAttribute("aria-label", "Place " + placementLabel + " " + (entry.key || "new keynote"));
      placeButton.setAttribute("title", placeTitle);
      placeButton.disabled = safeMode;
      placeButton.appendChild(createPlaceKeynoteIcon());
      placeButton.addEventListener("click", function (event) {
        event.preventDefault();
        event.stopPropagation();
        placeKeynote(entry);
      });

      menuButton.type = "button";
      menuButton.className = "note-more-button";
      menuButton.appendChild(createEllipsisVerticalIcon());
      menuButton.setAttribute("aria-label", "More actions for keynote " + (entry.key || "new keynote"));
      menuButton.setAttribute("aria-haspopup", "menu");
      menuButton.setAttribute("aria-expanded", "false");
      menuButton.setAttribute("title", "More actions");
      menuButton.addEventListener("click", function (event) {
        event.preventDefault();
        event.stopPropagation();
        openRowActionMenu(entry, menuButton);
      });

      if (row.hasChildren) {
        treeControl = document.createElement("button");
        treeControl.type = "button";
        treeControl.className = "note-tree-toggle";
        treeControl.textContent = row.isCollapsed ? ">" : "v";
        treeControl.setAttribute("aria-expanded", row.isCollapsed ? "false" : "true");
        treeControl.setAttribute(
          "aria-label",
          (row.isCollapsed ? "Expand " : "Collapse ") + (entry.key || "keynote group")
        );
        treeControl.addEventListener("click", function (event) {
          event.preventDefault();
          event.stopPropagation();
          toggleTreeRow(entry.id);
        });
      } else {
        treeControl = document.createElement("span");
        treeControl.className = "note-tree-spacer";
        treeControl.setAttribute("aria-hidden", "true");
      }

      keyInput.type = "text";
      keyInput.className = "form-control form-control-sm note-input note-key-input";
      keyInput.value = entry.key;
      keyInput.setAttribute("aria-label", "Key for " + (entry.key || "new keynote"));
      keyInput.disabled = safeMode || Boolean(claimTitle);
      lockKeyInput(keyInput);
      keyInput.setAttribute("title", safeMode ? "Review model issues before editing." : (claimTitle || "Double-click to edit key"));

      textInput.className = "form-control form-control-sm note-input note-text-input";
      textInput.rows = 2;
      textInput.value = entry.text;
      textInput.setAttribute("aria-label", "Description for " + (entry.key || "new keynote"));
      textInput.disabled = safeMode || Boolean(claimTitle);
      textInput.setAttribute("title", safeMode ? "Review model issues before editing." : (claimTitle || ""));

      [keyInput, textInput].forEach(function (input) {
        input.addEventListener("focus", function () {
          selectNote(entry.id, false);
        });
        input.addEventListener("blur", function () {
          if (input === keyInput) {
            lockKeyInput(keyInput);
          }
          deferRenderAllWhenEditingSettles();
        });
      });

      keyInput.addEventListener("dblclick", function (event) {
        event.preventDefault();
        unlockKeyInput(keyInput);
      });

      keyInput.addEventListener("input", function () {
        updateEntryField(entry.id, "key", keyInput.value, false);
      });

      textInput.addEventListener("input", function () {
        updateEntryField(entry.id, "text", textInput.value, false);
        // resizeNoteTextInput(textInput);
      });

      item.addEventListener("click", function (event) {
        if (
          event.target.tagName !== "INPUT" &&
          event.target.tagName !== "TEXTAREA" &&
          event.target.tagName !== "BUTTON"
        ) {
          selectNote(entry.id, true);
        }
      });

      keyWrap.appendChild(treeControl);
      keyWrap.appendChild(keyInput);
      appendPlacedKeyBadge(keyWrap, entry.key);
      actionWrap.appendChild(placeButton);
      actionWrap.appendChild(menuButton);
      actionCell.appendChild(actionWrap);
      keyCell.appendChild(keyWrap);
      textCell.appendChild(textInput);
      item.appendChild(actionCell);
      item.appendChild(keyCell);
      item.appendChild(textCell);
      body.appendChild(item);
      // resizeNoteTextInput(textInput);
    });
  }

  function renderValidation() {
    var container = byId("validation-list");
    var warningPill = byId("warning-pill");
    var issues = validateAll();
    var errors = issues.filter(function (issue) {
      return text(issue.severity).toLowerCase() === "error";
    });
    var warnings = issues.filter(function (issue) {
      return text(issue.severity).toLowerCase() !== "error";
    });
    var total = errors.length + warnings.length;

    setText("validation-summary", formatNumber(total));
    if (warningPill) {
      warningPill.setAttribute("data-severity", errors.length ? "error" : (warnings.length ? "warning" : "none"));
    }

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

  function formatPercent(value) {
    var number = Number(value || 0) * 100;
    if (!isFinite(number)) {
      number = 0;
    }
    return number.toFixed(number >= 10 ? 0 : 1) + "%";
  }

  function modelIssueDetail(issue) {
    var parts = [];
    var sheetLabels = [];

    if (issue.key) {
      parts.push("Key: " + issue.key);
    }
    if (issue.placedCount) {
      parts.push(formatNumber(issue.placedCount) + " placed");
    }
    if (issue.userKeynoteCount) {
      parts.push(formatNumber(issue.userKeynoteCount) + " user keynote");
    }
    if (issue.genericAnnotationCount) {
      parts.push(formatNumber(issue.genericAnnotationCount) + " generic annotation");
    }
    if (issue.sheetCount) {
      parts.push(formatNumber(issue.sheetCount) + " sheet(s)");
    }
    if (issue.unsheetedCount) {
      parts.push(formatNumber(issue.unsheetedCount) + " unsheeted");
    }
    (issue.sheets || []).slice(0, 4).forEach(function (sheet) {
      var number = trim(sheet && sheet.number);
      var name = trim(sheet && sheet.name);
      if (number || name) {
        sheetLabels.push(trim(number + " " + name));
      }
    });
    if (sheetLabels.length) {
      parts.push("Sheets: " + sheetLabels.join(", "));
    }
    if ((issue.sheets || []).length > sheetLabels.length) {
      parts.push("+" + formatNumber((issue.sheets || []).length - sheetLabels.length) + " more sheet(s)");
    }
    if ((issue.typeNames || []).length) {
      parts.push("Types: " + issue.typeNames.join(", "));
    }
    if (issue.details) {
      parts.push(issue.details);
    }

    return parts.join(" | ");
  }

  function createModelHealthStat(label, value) {
    var item = document.createElement("div");
    var strong = document.createElement("strong");
    var span = document.createElement("span");

    strong.textContent = value;
    span.textContent = label;
    item.appendChild(strong);
    item.appendChild(span);
    return item;
  }

  function modelIssueGroupTitle(issue) {
    var code = text(issue && issue.code);
    if (code === "placedKeyMissingFromLibrary") {
      return "Missing Placed Keys";
    }
    if (code === "genericAnnotationTextMismatch") {
      return "Generic Annotation Text";
    }
    if (code === "genericAnnotationDuplicateTypes") {
      return "Duplicate Generic Annotation Types";
    }
    if (code.indexOf("genericAnnotation") === 0) {
      return "Generic Annotation Setup";
    }
    return "Model Health";
  }

  function applyModelIssueResolution(issue, source, replacementKey) {
    var resolution = issue && issue.resolution;
    var entry = issue && entriesByKey()[issue.key];
    var replacementEntry;

    if (!resolution || (source !== "familyType" && source !== "textFile")) {
      return;
    }
    if (entry && isEntryRemotelyClaimed(entry)) {
      setStatus({
        status: "warning",
        message: editClaimTitle(entry) || "This keynote is being edited by another user."
      });
      renderAll();
      return;
    }

    if (source === "familyType") {
      if (!entry) {
        entry = normalizeEntry({
          id: makeLocalId(),
          key: issue.key,
          text: resolution.familyTypeText,
          parentKey: inferParentKeyForMissingKey(issue.key),
          sortOrder: state.entries.length,
          lineNumber: null
        }, state.entries.length);
        state.entries.push(entry);
      } else {
        entry.text = text(resolution.familyTypeText);
      }
      replacementKey = "";
    } else {
      replacementKey = trim(replacementKey);
      replacementEntry = entriesByKey()[replacementKey];
      if (!replacementEntry) {
        setStatus({ status: "warning", message: "Select a keynote entry from the text file first." });
        return;
      }
      if (entry && !state.baselineEntries.some(function (baselineEntry) {
        return baselineEntry.key === issue.key;
      })) {
        state.entries = state.entries.filter(function (candidate) {
          return candidate.id !== entry.id;
        });
      }
    }

    state.modelIssueResolutions[modelIssueResolutionId(issue)] = {
      issueKey: issue.key,
      source: source,
      replacementKey: replacementKey
    };
    markDirty();
    syncLocalEditClaims();
    setStatus({
      status: "warning",
      message: source === "familyType"
        ? "Keynote " + issue.key + " will be added to the text file from the family type when saved."
        : "Family keynote " + issue.key + " will be replaced by text-file keynote " + replacementKey + " when saved."
    });
    renderAll();
  }

  function createModelIssueResolutionControls(issue) {
    var resolution = issue.resolution;
    var selectedResolution = state.modelIssueResolutions[modelIssueResolutionId(issue)] || {};
    var selectedSource = selectedResolution.source || "";
    var controls = document.createElement("div");
    var familyButton = document.createElement("button");
    var fileChoice = document.createElement("div");
    var fileSelect = document.createElement("select");
    var fileButton = document.createElement("button");
    var placeholderOption = document.createElement("option");

    controls.className = "model-issue-resolution";
    familyButton.type = "button";
    familyButton.className = "model-resolution-button";
    familyButton.setAttribute("data-selected", selectedSource === "familyType" ? "true" : "false");
    familyButton.textContent = "Use Family Type: " + (resolution.familyTypeText || "(blank)");
    familyButton.title = "Write the family type text to the keynote file.";
    familyButton.addEventListener("click", function () {
      applyModelIssueResolution(issue, "familyType");
    });

    fileChoice.className = "model-resolution-file-choice";
    fileSelect.className = "model-resolution-select";
    fileSelect.setAttribute("aria-label", "Replacement keynote from text file for " + issue.key);
    placeholderOption.value = "";
    placeholderOption.textContent = "Choose text-file keynote...";
    fileSelect.appendChild(placeholderOption);
    state.baselineEntries.slice().sort(compareEntriesByKey).forEach(function (entry) {
      var option;
      if (!trim(entry.key) || entry.key === issue.key) {
        return;
      }
      option = document.createElement("option");
      option.value = entry.key;
      option.textContent = entry.key + " - " + (entry.text || "(blank)");
      fileSelect.appendChild(option);
    });
    fileSelect.value = selectedResolution.replacementKey || "";

    fileButton.type = "button";
    fileButton.className = "model-resolution-button";
    fileButton.setAttribute("data-selected", selectedSource === "textFile" ? "true" : "false");
    fileButton.textContent = "Use Text File";
    fileButton.disabled = !fileSelect.value;
    fileButton.title = "Overwrite the family type and migrate its placed instances to the selected text-file keynote.";
    fileSelect.addEventListener("change", function () {
      fileButton.disabled = !fileSelect.value;
    });
    fileButton.addEventListener("click", function () {
      applyModelIssueResolution(issue, "textFile", fileSelect.value);
    });

    fileChoice.appendChild(fileSelect);
    fileChoice.appendChild(fileButton);
    controls.appendChild(familyButton);
    controls.appendChild(fileChoice);
    return controls;
  }

  function renderModelHealth() {
    var health = currentModelHealth();
    var issueCount = modelIssueCount();
    var isSafeMode = isModelSafeModeActive();
    var pill = byId("model-health-pill");
    var summary = byId("model-health-summary");
    var safeStrip = byId("safe-mode-strip");
    var safeMessage = byId("safe-mode-message");
    var overlay = byId("model-issues-overlay");
    var status = byId("model-health-status");
    var stats = byId("model-health-stats");
    var list = byId("model-issues-list");
    var acknowledge = byId("acknowledge-model-health");
    var severity = isSafeMode ? "error" : (issueCount ? "warning" : "none");

    if (summary) {
      summary.textContent = formatNumber(issueCount);
    }
    if (pill) {
      pill.setAttribute("data-severity", severity);
      pill.setAttribute("aria-expanded", state.modelIssuesOpen ? "true" : "false");
      pill.setAttribute("aria-pressed", state.modelIssuesOpen ? "true" : "false");
      pill.setAttribute("title", health.message || "Show model issues");
    }

    if (safeStrip) {
      safeStrip.hidden = !isSafeMode;
    }
    if (safeMessage) {
      safeMessage.textContent = modelSafeModeMessage();
    }

    if (overlay) {
      overlay.hidden = !state.modelIssuesOpen;
    }
    if (status) {
      status.textContent = health.message || "Model health has not been scanned.";
    }

    if (stats) {
      clearElement(stats);
      stats.appendChild(createModelHealthStat("Placed keys", formatNumber(health.placedKeyCount)));
      stats.appendChild(createModelHealthStat("Missing keys", formatNumber(health.missingKeyCount)));
      stats.appendChild(createModelHealthStat("Missing ratio", formatPercent(health.missingRatio)));
      stats.appendChild(createModelHealthStat("Sheets", formatNumber(health.sheetCount)));
    }

    if (list) {
      clearElement(list);
      if (!issueCount) {
        var empty = document.createElement("div");
        empty.className = "empty-state";
        empty.textContent = "No model issues.";
        list.appendChild(empty);
      } else {
        var lastGroupTitle = "";
        health.issues.forEach(function (issue) {
          var canJump = Boolean(issue.key && entriesByKey()[issue.key]);
          var canResolve = Boolean(isSafeMode && issue.resolution);
          var item = document.createElement(canResolve ? "div" : "button");
          var label = document.createElement("span");
          var message = document.createElement("strong");
          var detail = document.createElement("span");
          var groupTitle = modelIssueGroupTitle(issue);

          if (groupTitle !== lastGroupTitle) {
            var group = document.createElement("h3");
            group.className = "model-issue-group";
            group.textContent = groupTitle;
            list.appendChild(group);
            lastGroupTitle = groupTitle;
          }

          if (!canResolve) {
            item.type = "button";
          }
          item.className = "model-issue-item";
          item.setAttribute("data-severity", issue.severity || "warning");
          if (!canResolve) {
            item.disabled = !canJump;
            item.setAttribute(
              "title",
              canJump ? "Jump to keynote " + issue.key : "This issue is not tied to a row in the keynote file."
            );
          }

          label.className = "validation-label";
          label.textContent = text(issue.severity || "warning").toUpperCase();
          message.textContent = issue.message || "";
          detail.className = "validation-detail";
          detail.textContent = modelIssueDetail(issue);

          item.appendChild(label);
          item.appendChild(message);
          item.appendChild(detail);

          if (canResolve) {
            item.appendChild(createModelIssueResolutionControls(issue));
          } else if (canJump) {
            item.addEventListener("click", function () {
              setModelIssuesOpen(false);
              selectEntryByKey(issue.key, true);
            });
          }

          list.appendChild(item);
        });
      }
    }

    if (acknowledge) {
      var unresolvedCount = unresolvedModelIssueCount();
      acknowledge.disabled = !health.safeModeRecommended || (isSafeMode && unresolvedCount > 0);
      acknowledge.textContent = isSafeMode && unresolvedCount
        ? "Resolve " + formatNumber(unresolvedCount) + " missing keynote(s) to unlock"
        : (isSafeMode ? "Reviewed - Unlock Editing" : "Reviewed");
    }
  }

  function setModelIssuesOpen(isOpen) {
    state.modelIssuesOpen = Boolean(isOpen);
    renderModelHealth();
    syncSidebarState();
  }

  function acknowledgeModelHealthReview() {
    if (!currentModelHealth().safeModeRecommended) {
      setModelIssuesOpen(false);
      return;
    }
    if (isModelSafeModeActive() && unresolvedModelIssueCount()) {
      setStatus({
        status: "warning",
        message: "Choose the family type or a replacement text-file keynote for each available missing-key error."
      });
      return;
    }
    state.reviewedModelHealthSignature = modelHealthSignature();
    setModelIssuesOpen(false);
    setStatus({
      status: "ready",
      message: "Model issues reviewed. Editing is unlocked for this scan."
    });
    renderAll();
  }

  function renderRowActions() {
    var selectedNote = selectedNoteEntry();
    var target = selectedNote || selectedDivisionEntry();
    var sequenceParent = selectedSequenceParentEntry();
    var addRootButton = byId("add-root");
    var deleteButton = byId("delete-row");
    var addSequenceButton = byId("add-sequence");
    var addSubNoteButton = byId("add-sub-note");
    var duplicateButton = byId("duplicate-row");
    var targetClaimed = target && isEntryRemotelyClaimed(target);
    var sequenceParentClaimed = sequenceParent && isEntryRemotelyClaimed(sequenceParent);
    var subNoteParentClaimed = selectedNote && isEntryRemotelyClaimed(selectedNote);
    var safeMode = isModelSafeModeActive();

    if (addRootButton) {
      addRootButton.disabled = safeMode;
    }
    if (addSequenceButton) {
      addSequenceButton.disabled = safeMode || !sequenceParent || sequenceParentClaimed;
    }
    if (addSubNoteButton) {
      addSubNoteButton.disabled = safeMode || !selectedNote || subNoteParentClaimed;
    }
    if (duplicateButton) {
      duplicateButton.disabled = safeMode || !target || targetClaimed;
    }
    if (deleteButton) {
      deleteButton.disabled = safeMode || !target || targetClaimed;
    }
  }

  function renderSaveState() {
    var saveButton = byId("save-data");
    var analyticsButton = byId("collect-analytics");
    var issues = validateAll();
    var safeMode = isModelSafeModeActive();
    var canSave = Boolean(
      state.payload &&
      state.payload.keynotePath &&
      state.dirty &&
      !state.saving &&
      !safeMode &&
      !hasBlockingSourceIssue() &&
      !hasErrorIssues(issues)
    );

    if (saveButton) {
      saveButton.disabled = !canSave;
      saveButton.textContent = state.saving ? "Saving..." : "Save";
      saveButton.setAttribute(
        "title",
        safeMode ? "Review model issues before saving." : "Save data to Keynote file"
      );
    }
    if (analyticsButton) {
      analyticsButton.disabled = state.analyticsCollecting || !state.payload || !state.payload.libraryKey;
      analyticsButton.textContent = state.analyticsCollecting ? "Collecting..." : "Collect Analytics";
    }
  }

  function syncSidebarState() {
    var workspace = document.querySelector(".workspace");
    var divisionPanel = byId("division-panel");
    var divisionBody = byId("keynotes-section-body");
    var divisionToggle = byId("toggle-divisions");
    var warningPill = byId("warning-pill");
    var modelHealthPill = byId("model-health-pill");
    var modelIssuesOverlay = byId("model-issues-overlay");
    var validationPanel = byId("validation-panel");
    var divisionsExpanded = !state.divisionsCollapsed;
    var warningsOpen = Boolean(state.warningSidebarOpen);
    var divisionToggleIcon;

    if (workspace) {
      workspace.classList.toggle("is-divisions-collapsed", state.divisionsCollapsed);
      workspace.classList.toggle("has-warning-sidebar", warningsOpen);
    }

    if (divisionPanel) {
      divisionPanel.hidden = !divisionsExpanded;
      divisionPanel.setAttribute("aria-hidden", divisionsExpanded ? "false" : "true");
    }

    if (divisionBody) {
      divisionBody.hidden = !divisionsExpanded;
    }

    if (divisionToggle) {
      divisionToggle.setAttribute("aria-expanded", divisionsExpanded ? "true" : "false");
      divisionToggle.setAttribute("aria-label", divisionsExpanded ? "Collapse divisions" : "Expand divisions");
      divisionToggle.setAttribute("title", divisionsExpanded ? "Collapse divisions" : "Expand divisions");
      divisionToggleIcon = divisionToggle.querySelector("[aria-hidden='true']");
      if (divisionToggleIcon) {
        divisionToggleIcon.textContent = divisionsExpanded ? "<" : ">";
      }
    }

    if (warningPill) {
      warningPill.setAttribute("aria-expanded", warningsOpen ? "true" : "false");
      warningPill.setAttribute("aria-pressed", warningsOpen ? "true" : "false");
      warningPill.setAttribute("aria-label", warningsOpen ? "Hide warnings" : "Show warnings");
    }

    if (modelHealthPill) {
      modelHealthPill.setAttribute("aria-expanded", state.modelIssuesOpen ? "true" : "false");
      modelHealthPill.setAttribute("aria-pressed", state.modelIssuesOpen ? "true" : "false");
      modelHealthPill.setAttribute("aria-label", state.modelIssuesOpen ? "Hide model issues" : "Show model issues");
    }

    if (modelIssuesOverlay) {
      modelIssuesOverlay.hidden = !state.modelIssuesOpen;
    }

    if (validationPanel) {
      validationPanel.hidden = !warningsOpen;
      validationPanel.setAttribute("aria-hidden", warningsOpen ? "false" : "true");
    }
  }

  function setDivisionsCollapsed(isCollapsed) {
    state.divisionsCollapsed = Boolean(isCollapsed);
    syncSidebarState();
  }

  function toggleDivisionsSidebar() {
    setDivisionsCollapsed(!state.divisionsCollapsed);
  }

  function setWarningSidebarOpen(isOpen) {
    state.warningSidebarOpen = Boolean(isOpen);
    syncSidebarState();
  }

  function toggleWarningSidebar() {
    setWarningSidebarOpen(!state.warningSidebarOpen);
  }

  function renderAll() {
    ensureSelection();
    renderMeta();
    renderDivisions();
    renderDivisionSelect();
    renderDivisionHeader();
    renderNotes();
    renderValidation();
    renderModelHealth();
    renderSaveState();
    renderRowActions();
    syncSidebarState();
  }

  function deferRenderAllWhenEditingSettles() {
    globalScope.setTimeout(function () {
      var active = document.activeElement;
      if (
        active &&
        active.closest &&
        active.closest(".note-table-body, .selected-division-card, .row-action-menu")
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

  function analyticsRowsFromResult(analytics) {
    return (analytics && (analytics.analyticsRows || analytics.rows)) || [];
  }

  function analyticsPlacedKeyMap(analytics) {
    var map = {};
    if (analytics && analytics.placedKeyMap) {
      Object.keys(analytics.placedKeyMap).forEach(function (key) {
        var normalizedKey = trim(key);
        if (normalizedKey && analytics.placedKeyMap[key]) {
          map[normalizedKey] = true;
        }
      });
      return map;
    }

    analyticsRowsFromResult(analytics).forEach(function (row) {
      var key = trim(row && row.keynoteKey);
      if (key && row && Number(row.placedCount || 0) > 0) {
        map[key] = true;
      }
    });
    return map;
  }

  function analyticsSummary(analytics) {
    analytics = analytics || {};
    return {
      entryCount: Number(analytics.entryCount || 0),
      analyticsRowCount: Number(analytics.analyticsRowCount || analyticsRowsFromResult(analytics).length || 0),
      placedKeyCount: Number(analytics.placedKeyCount || 0),
      placedCount: Number(analytics.placedCount || 0),
      userKeynoteCount: Number(analytics.userKeynoteCount || 0),
      genericAnnotationCount: Number(analytics.genericAnnotationCount || 0),
      sheetCount: Number(analytics.sheetCount || 0),
      unsheetedCount: Number(analytics.unsheetedCount || 0),
      orphanKeyCount: Number(analytics.orphanKeyCount || 0),
      skippedCount: Number(analytics.skippedCount || 0),
      collectedAt: text(analytics.collectedAt)
    };
  }

  function applyAnalyticsResultToUi(analytics, modelHealth) {
    state.lastAnalyticsResult = analytics || null;
    if (modelHealth) {
      state.modelHealth = normalizeModelHealth(modelHealth);
    }
    state.sheetVisibleKeynotes = analyticsPlacedKeyMap(analytics);
    if (state.modelHealth && Object.keys(state.modelHealth.placedKeyMap || {}).length) {
      state.sheetVisibleKeynotes = state.modelHealth.placedKeyMap;
    }
    renderAll();
  }

  function analyticsSyncPayload(analytics) {
    var client = currentClient();
    analytics = analytics || {};
    return {
      libraryKey: analytics.libraryKey || (state.payload && state.payload.libraryKey) || "",
      documentKey: analytics.documentKey || "",
      documentTitle: analytics.documentTitle || "",
      documentPath: analytics.documentPath || "",
      centralPath: analytics.centralPath || "",
      documentKeySource: analytics.documentKeySource || "",
      summary: analyticsSummary(analytics),
      entries: analyticsRowsFromResult(analytics),
      clientId: client.clientId,
      clientName: client.clientName
    };
  }

  function ensureLibraryBeforeAnalyticsSync(db, analytics) {
    var client = currentClient();

    if (!db || typeof db.ensureLibrary !== "function") {
      return Promise.reject(new Error("The Supabase library API was not available."));
    }

    return db.ensureLibrary({
      libraryKey: analytics.libraryKey || (state.payload && state.payload.libraryKey) || "",
      displayPath: analytics.displayPath || analytics.keynotePath || "",
      keynotePath: analytics.keynotePath || "",
      encoding: analytics.encoding || "utf-8",
      lineEnding: analytics.lineEnding || "\r\n",
      entries: analytics.entries || [],
      clientId: client.clientId,
      clientName: client.clientName
    });
  }

  function syncAnalyticsResult(analytics) {
    var db = dbManager();
    var payload = analyticsSyncPayload(analytics);

    if (!db || typeof db.syncAnalytics !== "function") {
      return Promise.reject(new Error("The Supabase analytics API was not available."));
    }

    if (!payload.libraryKey) {
      return Promise.reject(new Error("No keynote library key was available for analytics sync."));
    }

    if (!payload.documentKey) {
      return Promise.reject(new Error("No Revit document key was available for analytics sync."));
    }

    return ensureLibraryBeforeAnalyticsSync(db, analytics).then(function () {
      return db.syncAnalytics(payload);
    });
  }

  function findBaselineEntry(entry) {
    var index;
    if (!entry) {
      return null;
    }
    for (index = 0; index < state.baselineEntries.length; index += 1) {
      if (state.baselineEntries[index].id === entry.id) {
        return state.baselineEntries[index];
      }
    }
    return null;
  }

  function claimKeyForEntry(entry, baselineEntry) {
    var dbId = text((baselineEntry && baselineEntry.dbId) || (entry && entry.dbId));
    var key = trim(baselineEntry && baselineEntry.key);

    if (dbId) {
      return "db:" + dbId;
    }
    if (key) {
      return "key:" + key;
    }
    return "";
  }

  function claimKeysForEntry(entry) {
    var baselineEntry = findBaselineEntry(entry);
    var keys = [];
    var dbId = text((baselineEntry && baselineEntry.dbId) || (entry && entry.dbId));
    var baseKey = trim(baselineEntry && baselineEntry.key);
    var currentKey = trim(entry && entry.key);

    if (dbId) {
      keys.push("db:" + dbId);
    } else if (baseKey) {
      keys.push("key:" + baseKey);
    }
    if (!baselineEntry && currentKey) {
      keys.push("key:" + currentKey);
    } else if (currentKey && currentKey !== baseKey) {
      keys.push("key:" + currentKey);
    }
    return keys;
  }

  function remoteClaimForEntry(entry) {
    var keys = claimKeysForEntry(entry);
    var index;
    for (index = 0; index < keys.length; index += 1) {
      if (state.remoteEditClaims[keys[index]]) {
        return state.remoteEditClaims[keys[index]];
      }
    }
    return null;
  }

  function isEntryRemotelyClaimed(entry) {
    return Boolean(remoteClaimForEntry(entry));
  }

  function editClaimTitle(entry) {
    var claim = remoteClaimForEntry(entry);
    var name;
    if (!claim) {
      return "";
    }
    name = trim(claim.clientName) || "another user";
    return "Being edited by " + name;
  }

  function normalizeEditClaimsData(data) {
    data = data || {};
    data.claims = data.claims || [];
    return data;
  }

  function remoteEditClaimSignature(claims) {
    return Object.keys(claims || {}).sort().map(function (claimKey) {
      var claim = claims[claimKey] || {};
      return claimKey + ":" + text(claim.clientId) + ":" + text(claim.clientName);
    }).join("|");
  }

  function applyEditClaimsData(data) {
    var client = currentClient();
    var remote = {};
    var previousSignature = remoteEditClaimSignature(state.remoteEditClaims);
    var nextSignature;

    data = normalizeEditClaimsData(data);
    data.claims.forEach(function (claim) {
      var claimKey = text(claim.claimKey);
      if (!claimKey || (client.clientId && text(claim.clientId) === client.clientId)) {
        return;
      }
      remote[claimKey] = {
        claimKey: claimKey,
        dbId: text(claim.dbId),
        key: text(claim.key),
        clientId: text(claim.clientId),
        clientName: text(claim.clientName),
        updatedAt: text(claim.updatedAt)
      };
    });

    nextSignature = remoteEditClaimSignature(remote);
    state.remoteEditClaims = remote;
    if (previousSignature !== nextSignature) {
      renderAll();
    }
  }

  function entryHasClaimableEdit(entry, baselineEntry) {
    if (!entry) {
      return false;
    }
    if (!baselineEntry) {
      return Boolean(trim(entry.key));
    }
    return (
      trim(entry.key) !== trim(baselineEntry.key) ||
      trim(entry.text) !== trim(baselineEntry.text)
    );
  }

  function buildLocalEditClaims() {
    var claims = [];
    var seen = {};

    state.entries.forEach(function (entry) {
      var baselineEntry = findBaselineEntry(entry);
      var dbId = text((baselineEntry && baselineEntry.dbId) || entry.dbId || "");

      if (!entryHasClaimableEdit(entry, baselineEntry)) {
        return;
      }

      claimKeysForEntry(entry).forEach(function (claimKey) {
        var key = trim((baselineEntry && baselineEntry.key) || entry.key);

        if (!claimKey || seen[claimKey]) {
          return;
        }
        seen[claimKey] = true;
        if (claimKey.indexOf("key:") === 0) {
          key = claimKey.slice(4);
        }
        claims.push({
          claimKey: claimKey,
          dbId: claimKey.indexOf("db:") === 0 ? dbId : "",
          key: key
        });
      });
    });

    return claims;
  }

  function editClaimSignature(claims) {
    return (claims || []).map(function (claim) {
      return claim.claimKey;
    }).sort().join("|");
  }

  function claimsAreConfigured() {
    var settings = (state.payload && state.payload.supabase) || {};
    var db = dbManager();
    return Boolean(
      state.payload &&
      state.payload.libraryKey &&
      state.dbReady &&
      settings.configured &&
      db &&
      typeof db.setEditClaims === "function"
    );
  }

  function stopLocalEditClaimHeartbeat() {
    if (state.localEditClaimHeartbeat) {
      globalScope.clearInterval(state.localEditClaimHeartbeat);
      state.localEditClaimHeartbeat = null;
    }
  }

  function updateLocalEditClaimHeartbeat() {
    if (!state.localEditClaimSignature || !claimsAreConfigured()) {
      stopLocalEditClaimHeartbeat();
      return;
    }
    if (state.localEditClaimHeartbeat) {
      return;
    }
    state.localEditClaimHeartbeat = globalScope.setInterval(function () {
      if (!state.localEditClaimSignature) {
        stopLocalEditClaimHeartbeat();
        return;
      }
      syncLocalEditClaims({ force: true });
    }, EDIT_CLAIM_HEARTBEAT_MS);
  }

  function syncLocalEditClaims(options) {
    var db = dbManager();
    var client = currentClient();
    var claims = buildLocalEditClaims();
    var signature = editClaimSignature(claims);
    var previousSignature = state.localEditClaimSignature;

    options = options || {};

    if (signature === state.localEditClaimSignature && !options.force) {
      updateLocalEditClaimHeartbeat();
      return Promise.resolve(null);
    }

    if (!claimsAreConfigured()) {
      if (!signature) {
        stopLocalEditClaimHeartbeat();
      }
      return Promise.resolve(null);
    }
    state.localEditClaimSignature = signature;

    return db.setEditClaims({
      libraryKey: state.payload.libraryKey,
      clientId: client.clientId,
      clientName: client.clientName,
      claims: claims
    }).then(function (data) {
      clearSyncIssueCode("editClaimUpdateFailed");
      applyEditClaimsData(data);
      renderValidation();
      renderSaveState();
      updateLocalEditClaimHeartbeat();
      return data;
    }).catch(function (error) {
      state.localEditClaimSignature = previousSignature;
      upsertSyncIssue(makeIssue(
        "warning",
        "Could not update unsaved edit locks: " + (error.message || error),
        "",
        "editClaimUpdateFailed"
      ));
      renderValidation();
      renderSaveState();
      updateLocalEditClaimHeartbeat();
      return null;
    });
  }

  function clearLocalEditClaims() {
    var db = dbManager();
    var client = currentClient();
    var shouldClearRemote = state.localEditClaimSignature;

    if (!claimsAreConfigured() || !shouldClearRemote) {
      if (!shouldClearRemote) {
        stopLocalEditClaimHeartbeat();
      }
      return Promise.resolve(null);
    }
    state.localEditClaimSignature = "";
    stopLocalEditClaimHeartbeat();

    return db.setEditClaims({
      libraryKey: state.payload.libraryKey,
      clientId: client.clientId,
      clientName: client.clientName,
      claims: []
    }).then(function (data) {
      applyEditClaimsData(data);
      return data;
    }).catch(function () {
      state.localEditClaimSignature = shouldClearRemote;
      updateLocalEditClaimHeartbeat();
      return null;
    });
  }

  function refreshEditClaims() {
    var db = dbManager();
    if (
      !state.payload ||
      !state.payload.libraryKey ||
      !db ||
      typeof db.getEditClaims !== "function"
    ) {
      return Promise.resolve(null);
    }

    return db.getEditClaims(state.payload.libraryKey).then(function (data) {
      applyEditClaimsData(data);
      return data;
    }).catch(function () {
      return null;
    });
  }

  function subscribeToEditClaims(snapshot) {
    var db = dbManager();
    var client = currentClient();

    if (!db || !snapshot || !snapshot.libraryId || typeof db.subscribeEditClaims !== "function") {
      return;
    }

    db.subscribeEditClaims(snapshot.libraryId, client.clientId, {
      onClaimsChanged: function () {
        refreshEditClaims();
      }
    });
  }

  function indexEntriesById(entries) {
    var map = {};
    (entries || []).forEach(function (entry) {
      if (entry.id) {
        map[entry.id] = entry;
      }
    });
    return map;
  }

  function indexEntriesByKey(entries) {
    var map = {};
    (entries || []).forEach(function (entry) {
      var key = trim(entry.key);
      if (key && !map[key]) {
        map[key] = entry;
      }
    });
    return map;
  }

  function indexEntriesByDbId(entries) {
    var map = {};
    (entries || []).forEach(function (entry) {
      var dbId = text(entry.dbId || entry.id);
      if (dbId && !map[dbId]) {
        map[dbId] = entry;
      }
    });
    return map;
  }

  function entryFieldsEqual(first, second) {
    return Boolean(first && second) &&
      trim(first.key) === trim(second.key) &&
      trim(first.text) === trim(second.text) &&
      trim(first.parentKey) === trim(second.parentKey);
  }

  function numericVersion(value) {
    var number = Number(value || 0);
    return isNaN(number) ? 0 : number;
  }

  function nullableDbId(value) {
    // Supabase assigns UUIDs to new rows; JSON null prevents a blank value from reaching a UUID cast.
    var dbId = trim(value);
    return dbId || null;
  }

  function sortOrderFor(entry, fallbackIndex, baseEntry) {
    if (baseEntry && (baseEntry.sortOrder || baseEntry.sortOrder === 0)) {
      return Number(baseEntry.sortOrder);
    }
    if (entry && (entry.sortOrder || entry.sortOrder === 0)) {
      return Number(entry.sortOrder);
    }
    return fallbackIndex || 0;
  }

  function makeDbUpsert(entry, baseEntry, index) {
    return {
      dbId: nullableDbId((baseEntry && baseEntry.dbId) || entry.dbId),
      key: trim(entry.key),
      text: trim(entry.text),
      parentKey: trim(entry.parentKey),
      sortOrder: sortOrderFor(entry, index, baseEntry),
      baseVersion: numericVersion((baseEntry && baseEntry.rowVersion) || entry.rowVersion),
      previousKey: baseEntry ? trim(baseEntry.key) : ""
    };
  }

  function makeDbDelete(baseEntry) {
    return {
      dbId: nullableDbId(baseEntry.dbId),
      key: trim(baseEntry.key),
      baseVersion: numericVersion(baseEntry.rowVersion)
    };
  }

  function buildPendingDbChanges() {
    var baselineById = indexEntriesById(state.baselineEntries);
    var currentById = indexEntriesById(state.entries);
    var seen = {};
    var changes = {
      upserts: [],
      deletes: []
    };

    state.baselineEntries.forEach(function (baseEntry) {
      var currentEntry = currentById[baseEntry.id];
      seen[baseEntry.id] = true;

      if (!currentEntry) {
        changes.deletes.push(makeDbDelete(baseEntry));
        return;
      }

      if (!entryFieldsEqual(baseEntry, currentEntry)) {
        changes.upserts.push(makeDbUpsert(
          currentEntry,
          baseEntry,
          state.entries.indexOf(currentEntry)
        ));
      }
    });

    state.entries.forEach(function (entry, index) {
      if (!seen[entry.id]) {
        changes.upserts.push(makeDbUpsert(entry, null, index));
      }
    });

    return changes;
  }

  function addFileMetadataToChanges(changes, payload) {
    changes = changes || { upserts: [], deletes: [] };
    payload = payload || {};
    changes.metadata = {
      displayPath: payload.displayPath || payload.keynotePath || "",
      encoding: payload.encoding || "utf-8",
      lineEnding: payload.lineEnding || "\r\n",
      fileHash: payload.fileHash || "",
      lastWriteUtc: payload.lastWriteUtc || null
    };
    return changes;
  }

  function mergeSnapshotMetadataInto(entries, snapshot) {
    var snapshotByKey = indexEntriesByKey(snapshot && snapshot.entries);
    var count = 0;

    (entries || []).forEach(function (entry) {
      var match = snapshotByKey[trim(entry.key)];
      if (!match) {
        return;
      }

      entry.dbId = text(match.dbId || match.id || "");
      entry.rowVersion = match.rowVersion || null;
      entry.sortOrder = match.sortOrder === 0 ? 0 : (match.sortOrder || entry.sortOrder);
      count += 1;
    });

    return count;
  }

  function applySnapshotMetadata(snapshot) {
    if (!snapshot) {
      return;
    }

    state.dbSnapshot = snapshot;
    if (state.payload) {
      state.payload.libraryId = snapshot.libraryId || state.payload.libraryId || "";
      state.payload.datasetVersion = snapshot.datasetVersion || 0;
    }

    mergeSnapshotMetadataInto(state.entries, snapshot);
    mergeSnapshotMetadataInto(state.baselineEntries, snapshot);

    if (!state.dirty) {
      rememberBaseline();
    }

    renderAll();
  }

  function clearRemoteEntryTimer() {
    if (state.remoteEntriesTimer) {
      globalScope.clearTimeout(state.remoteEntriesTimer);
      state.remoteEntriesTimer = null;
    }
  }

  function clearRemoteEntryState() {
    clearRemoteEntryTimer();
    state.remoteEntrySnapshot = null;
    state.remoteEntriesPending = false;
    state.remotePending = false;
    clearSyncIssueCode("remoteEntriesPending");
    clearSyncIssueCode("remoteEntrySnapshotFailed");
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

  function markRemoteEntriesPending() {
    state.remotePending = true;
    state.remoteEntriesPending = true;
    upsertSyncIssue(makeIssue(
      "warning",
      "Remote keynote add/delete changes are available. Save will check conflicts; Refresh discards local edits and loads the latest data.",
      "",
      "remoteEntriesPending"
    ));
  }

  function fetchRemoteEntrySnapshotForDirtyWindow() {
    var db = dbManager();

    if (
      !state.payload ||
      !state.payload.libraryKey ||
      !db ||
      typeof db.getSnapshot !== "function"
    ) {
      markRemoteEntriesPending();
      setStatus({
        status: "warning",
        message: "Remote keynote add/delete changes are available while you have unsaved edits."
      });
      renderValidation();
      renderSaveState();
      return Promise.resolve(null);
    }

    return db.getSnapshot(state.payload.libraryKey).then(function (snapshot) {
      state.remoteEntrySnapshot = snapshot;
      clearSyncIssueCode("remoteEntrySnapshotFailed");
      markRemoteEntriesPending();
      setStatus({
        status: "warning",
        message: "Remote keynote add/delete changes are available while you have unsaved edits."
      });
      renderValidation();
      renderSaveState();
      return snapshot;
    }).catch(function (error) {
      markRemoteEntriesPending();
      upsertSyncIssue(makeIssue(
        "warning",
        "Remote keynote changes are available, but the latest Supabase snapshot could not be loaded: " + (error.message || error),
        "",
        "remoteEntrySnapshotFailed"
      ));
      setStatus({
        status: "warning",
        message: "Remote keynote changes are available, but the latest Supabase snapshot could not be loaded."
      });
      renderValidation();
      renderSaveState();
      return null;
    });
  }

  function processRemoteEntryChange() {
    state.remoteEntriesTimer = null;

    if (state.saving || state.pendingDbChanges) {
      return;
    }

    if (state.dirty) {
      fetchRemoteEntrySnapshotForDirtyWindow();
      return;
    }

    clearRemoteEntryState();
    state.allowNextLoad = false;
    postWebViewMessage({ type: "refreshData" });
  }

  function scheduleRemoteEntryChange() {
    if (state.saving || state.pendingDbChanges) {
      return;
    }

    clearRemoteEntryTimer();
    state.remoteEntriesTimer = globalScope.setTimeout(
      processRemoteEntryChange,
      REMOTE_ENTRY_REFRESH_DELAY_MS
    );
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
          upsertSyncIssue(makeIssue(
            "warning",
            "Remote Supabase changes are available. Save will check row conflicts; Refresh discards local edits and loads the latest data.",
            "",
            "remotePending"
          ));
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

  function subscribeToEntries(snapshot) {
    var db = dbManager();
    var client = currentClient();

    if (!db || !snapshot || !snapshot.libraryId || typeof db.subscribeEntries !== "function") {
      return;
    }

    db.subscribeEntries(snapshot.libraryId, client.clientId, {
      onEntriesChanged: function () {
        scheduleRemoteEntryChange();
      }
    });
  }

  function filePayloadDiffersFromSnapshot(payload, snapshot) {
    var fileEntries = (payload && payload.entries) || [];
    var snapshotEntries = (snapshot && snapshot.entries) || [];
    var snapshotByKey;
    var differs = false;

    if (!snapshot || fileEntries.length !== snapshotEntries.length) {
      return true;
    }

    snapshotByKey = indexEntriesByKey(snapshotEntries);
    fileEntries.forEach(function (entry) {
      var match = snapshotByKey[trim(entry.key)];
      if (!match || !entryFieldsEqual(entry, match)) {
        differs = true;
      }
    });
    return differs;
  }

  function attachSupabaseLibrary(payload, reason) {
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

    db.ensureLibrary({
      libraryKey: payload.libraryKey,
      displayPath: payload.displayPath || payload.keynotePath,
      keynotePath: payload.keynotePath,
      encoding: payload.encoding || "utf-8",
      lineEnding: payload.lineEnding || "\r\n",
      entries: payload.entries || [],
      clientId: client.clientId,
      clientName: client.clientName
    }).then(function (snapshot) {
      state.dbInitializing = false;
      state.dbReady = true;
      state.syncIssues = [];
      if (filePayloadDiffersFromSnapshot(payload, snapshot)) {
        return syncSavedFileSnapshot(
          db,
          payload,
          "Loaded shared keynote file.",
          "Supabase mirror recovered from the shared keynote file."
        ).then(function () {
          collectAnalyticsOnOpen();
        });
      }
      applySnapshotMetadata(snapshot);
      subscribeToLibrary(snapshot);
      subscribeToEntries(snapshot);
      subscribeToEditClaims(snapshot);
      refreshEditClaims();
      syncLocalEditClaims();
      if (!state.dirty && reason === "load") {
        setStatus({ status: "ready", message: "Loaded shared keynote file and attached Supabase row metadata." });
      }
      collectAnalyticsOnOpen();
    }).catch(function (error) {
      state.dbInitializing = false;
      state.dbReady = false;
      state.syncIssues = [makeIssue("warning", error.message, "", "supabaseMirrorError")];
      renderValidation();
      renderSaveState();
    });
  }

  function makeSupabaseSyncIssue(message, code) {
    return makeIssue("warning", message || "Supabase delta sync failed.", "", code || "supabaseDeltaSyncFailed");
  }

  function makeSupabaseConflictIssues(result) {
    var conflicts = (result && result.conflicts) || [];
    if (!conflicts.length) {
      return [makeSupabaseSyncIssue(
        (result && result.message) || "Some keynotes changed in Supabase before the delta save completed.",
        "supabaseConflict"
      )];
    }

    return conflicts.map(function (conflict) {
      return makeIssue(
        "warning",
        conflict.message || "This keynote changed in Supabase before the delta save completed.",
        conflict.key || "",
        "supabaseConflict"
      );
    });
  }

  function describeSupabaseDeltaResult(result) {
    var inserted = Number(result.insertedCount || 0);
    var updated = Number(result.updatedCount || 0);
    var deleted = Number(result.deletedCount || 0);
    var parts = [];

    if (inserted) {
      parts.push(inserted + " inserted");
    }
    if (updated) {
      parts.push(updated + " updated");
    }
    if (deleted) {
      parts.push(deleted + " deleted");
    }
    if (!parts.length) {
      return "metadata updated";
    }
    return parts.join(", ");
  }

  function isMissingSupabaseLibraryError(error) {
    return text(error && (error.message || error)).toLowerCase().indexOf("keynote library was not found") !== -1;
  }

  function isInvalidSupabaseUuidError(error) {
    return text(error && (error.message || error)).toLowerCase().indexOf("invalid input syntax for type uuid") !== -1;
  }

  function applySupabaseSnapshotResult(result, fileMessage, suffix) {
    state.pendingDbChanges = null;
    state.dbReady = true;
    state.syncIssues = [];
    clearRemoteEntryState();
    applySnapshotMetadata(result);
    subscribeToLibrary(result);
    subscribeToEntries(result);
    subscribeToEditClaims(result);
    refreshEditClaims();
    syncLocalEditClaims();
    setStatus({
      status: "ready",
      message: (fileMessage || "Saved shared keynote file.") + " " + suffix
    });
  }

  function syncSavedFileSnapshot(db, filePayload, fileMessage, fallbackReason) {
    var client = currentClient();

    if (!db || typeof db.syncFileSnapshot !== "function") {
      state.pendingDbChanges = null;
      state.dbReady = false;
      state.syncIssues = [makeSupabaseSyncIssue(
        "Saved the shared keynote file, but the Supabase full snapshot sync API was not available.",
        "supabaseClientMissing"
      )];
      renderValidation();
      renderSaveState();
      setStatus({
        status: "warning",
        message: (fileMessage || "Saved shared keynote file.") + " Supabase snapshot sync failed."
      });
      return Promise.resolve(null);
    }

    setStatus({
      status: "syncing",
      message: "Saved shared keynote file. Mirroring the full keynote file to Supabase..."
    });

    return db.syncFileSnapshot({
      libraryKey: filePayload.libraryKey,
      displayPath: filePayload.displayPath || filePayload.keynotePath || "",
      keynotePath: filePayload.keynotePath || "",
      encoding: filePayload.encoding || "utf-8",
      lineEnding: filePayload.lineEnding || "\r\n",
      fileHash: filePayload.fileHash || "",
      lastWriteUtc: filePayload.lastWriteUtc || null,
      entries: filePayload.entries || [],
      clientId: client.clientId,
      clientName: client.clientName
    }).then(function (result) {
      applySupabaseSnapshotResult(
        result,
        fileMessage,
        fallbackReason || "Supabase snapshot sync complete."
      );
      return result;
    }).catch(function (error) {
      state.pendingDbChanges = null;
      state.dbReady = false;
      state.syncIssues = [makeSupabaseSyncIssue(
        "Saved the shared keynote file, but Supabase snapshot sync failed: " + (error.message || error),
        "supabaseSnapshotSyncFailed"
      )];
      renderValidation();
      renderSaveState();
      setStatus({
        status: "warning",
        message: (fileMessage || "Saved shared keynote file.") + " Supabase snapshot sync failed."
      });
      return null;
    });
  }

  function savePendingDbChanges(filePayload, pendingChanges, fileMessage) {
    var db = dbManager();
    var settings = (state.payload && state.payload.supabase) || {};
    var client = currentClient();
    var baseDatasetVersion = (state.dbSnapshot && state.dbSnapshot.datasetVersion) ||
      (state.payload && state.payload.datasetVersion) ||
      0;

    filePayload = filePayload || {};
    pendingChanges = addFileMetadataToChanges(pendingChanges || { upserts: [], deletes: [] }, filePayload);

    if (!filePayload.libraryKey) {
      state.pendingDbChanges = null;
      return;
    }

    if (!settings.configured) {
      state.pendingDbChanges = null;
      state.dbReady = false;
      state.syncIssues = [makeSupabaseSyncIssue(
        "Saved the shared keynote file, but Supabase is not configured for delta sync.",
        "supabaseConfigMissing"
      )];
      renderValidation();
      renderSaveState();
      return;
    }

    if (!db || typeof db.saveChanges !== "function") {
      state.pendingDbChanges = null;
      state.dbReady = false;
      state.syncIssues = [makeSupabaseSyncIssue(
        "Saved the shared keynote file, but the Supabase delta save API was not available.",
        "supabaseClientMissing"
      )];
      renderValidation();
      renderSaveState();
      return;
    }

    if (!state.dbReady || !state.dbSnapshot || !state.dbSnapshot.libraryId) {
      syncSavedFileSnapshot(
        db,
        filePayload,
        fileMessage,
        "Supabase library was reattached from the saved file."
      );
      return;
    }

    setStatus({
      status: "syncing",
      message: "Saved shared keynote file. Updating changed Supabase rows..."
    });

    db.saveChanges({
      libraryKey: filePayload.libraryKey,
      clientId: client.clientId,
      clientName: client.clientName,
      baseDatasetVersion: baseDatasetVersion,
      changes: pendingChanges
    }).then(function (result) {
      state.pendingDbChanges = null;

      if ((result.status || "") === "conflict") {
        state.dbReady = false;
        state.syncIssues = makeSupabaseConflictIssues(result);
        if (result.snapshot) {
          applySnapshotMetadata(result.snapshot);
        }
        setStatus({
          status: "conflict",
          message: (fileMessage || "Saved shared keynote file.") + " Supabase delta sync has conflicts."
        });
        renderValidation();
        renderSaveState();
        return;
      }

      state.dbReady = true;
      state.syncIssues = [];
      applySupabaseSnapshotResult(
        result,
        fileMessage,
        "Supabase delta sync complete: " + describeSupabaseDeltaResult(result) + "."
      );
    }).catch(function (error) {
      if (isMissingSupabaseLibraryError(error)) {
        syncSavedFileSnapshot(
          db,
          filePayload,
          fileMessage,
          "Supabase library was recreated from the saved file."
        );
        return;
      }
      if (isInvalidSupabaseUuidError(error)) {
        syncSavedFileSnapshot(
          db,
          filePayload,
          fileMessage,
          "Supabase rejected a missing row ID, so the mirror was recovered from the saved file."
        );
        return;
      }

      state.pendingDbChanges = null;
      state.dbReady = false;
      state.syncIssues = [makeSupabaseSyncIssue(
        "Saved the shared keynote file, but Supabase delta sync failed: " + (error.message || error),
        "supabaseDeltaSyncFailed"
      )];
      renderValidation();
      renderSaveState();
      setStatus({
        status: "warning",
        message: (fileMessage || "Saved shared keynote file.") + " Supabase delta sync failed."
      });
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
    var preferences;

    state.payload = payload || {};
    preferences = state.payload.preferences || {};
    setPlacementMode(preferences.placementMode || state.payload.placementMode || state.placementMode);
    state.entries = (state.payload.entries || []).map(normalizeEntry);
    state.modelHealth = normalizeModelHealth(state.payload.modelHealth || {});
    state.modelIssueResolutions = {};
    state.sheetVisibleKeynotes = Object.keys(state.modelHealth.placedKeyMap || {}).length
      ? state.modelHealth.placedKeyMap
      : normalizeSheetVisibleKeynotes(state.payload.sheetVisibleKeynotes || {});
    state.sourceIssues = (state.payload.issues || []).filter(sourceIssueBlocksSave);
    state.operationIssues = [];
    state.collapsedEntryIds = {};
    state.saving = false;
    clearRemoteEntryState();

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
    rememberBaseline();
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

    clearLocalEditClaims();
    state.syncIssues = [];
    state.operationIssues = [];
    state.remotePending = false;
    state.remoteEntrySnapshot = null;
    state.remoteEntriesPending = false;
    state.remoteEditClaims = {};
    applyData(payload);
    rememberBaseline();
    attachSupabaseLibrary(state.payload, "load");
  }

  function updateEntryField(entryId, fieldName, value, shouldRender) {
    var entry = findEntryById(entryId);
    var oldKey;

    if (!entry) {
      return;
    }
    if (blockForSafeMode("Editing")) {
      return;
    }
    if (isEntryRemotelyClaimed(entry)) {
      setStatus({ status: "warning", message: editClaimTitle(entry) || "This keynote is being edited by another user." });
      renderAll();
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
    if (fieldName === "key" || fieldName === "text") {
      syncLocalEditClaims();
    }

    if (shouldRender) {
      renderAll();
      return;
    }

    renderMeta();
    renderValidation();
    renderSaveState();
    syncSelectionClasses();
  }

  function addRemoteReservedKeys(map) {
    var remoteSnapshot = currentRemoteEntrySnapshot();

    Object.keys(state.remoteEditClaims || {}).forEach(function (claimKey) {
      var key;
      if (claimKey.indexOf("key:") !== 0) {
        return;
      }
      key = trim(claimKey.slice(4));
      if (key) {
        map[key] = true;
      }
    });

    ((remoteSnapshot && remoteSnapshot.entries) || []).forEach(function (entry) {
      var key = trim(entry.key);
      if (key) {
        map[key] = true;
      }
    });
  }

  function makeUniqueKey(baseKey) {
    var base = trim(baseKey) || "NEW";
    var existing = {};
    var candidate = base;
    var index = 1;

    state.entries.forEach(function (entry) {
      existing[entry.key] = true;
    });
    addRemoteReservedKeys(existing);

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

  function normalizeChildKeyPrefix(parentKey) {
    var value = trim(parentKey);
    var number;

    if (!value) {
      return "NEW";
    }

    if (/^\d+$/.test(value)) {
      number = Number(value);
      if (!isNaN(number)) {
        return String(number);
      }
    }

    return value;
  }

  function padNumber(value, width) {
    var result = String(value);

    while (result.length < width) {
      result = "0" + result;
    }

    return result;
  }

  function makeChildKey(parentKey) {
    var normalizedParentKey = trim(parentKey);
    var prefix = normalizeChildKeyPrefix(parentKey);
    var existing = {};
    var max = 0;
    var width = 2;
    var bestPrefix = prefix;
    var nextNumber;
    var candidate;
    var match;

    state.entries.forEach(function (entry) {
      var key = trim(entry.key);
      var suffixNumber;
      existing[key] = true;

      if (trim(entry.parentKey) !== normalizedParentKey) {
        return;
      }

      match = key.match(/^(.+)\.(\d+)$/);
      if (!match) {
        return;
      }

      suffixNumber = Number(match[2]);
      if (isNaN(suffixNumber)) {
        return;
      }

      if (suffixNumber > max) {
        max = suffixNumber;
        bestPrefix = match[1];
        width = match[2].length;
      }
    });
    addRemoteReservedKeys(existing);

    nextNumber = max + 1;
    while (nextNumber < 10000) {
      candidate = bestPrefix + "." + padNumber(nextNumber, width);
      if (!existing[candidate]) {
        return candidate;
      }
      nextNumber += 1;
    }

    return makeUniqueKey(bestPrefix + ".NEW");
  }

  function addRoot() {
    if (blockForSafeMode("Adding keynotes")) {
      return;
    }

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
    syncLocalEditClaims();
    renderAll();
  }

  function selectedSequenceParentEntry() {
    var selectedNote = selectedNoteEntry();
    var parentKey;

    if (!selectedNote) {
      return selectedDivisionEntry();
    }

    parentKey = trim(selectedNote.parentKey);
    return parentKey ? entriesByKey()[parentKey] || null : null;
  }

  function addNoteUnderParent(parent, missingParentMessage) {
    var entry;

    if (blockForSafeMode("Adding keynotes")) {
      return;
    }

    if (!parent) {
      setStatus({
        status: "error",
        message: missingParentMessage || "Select a division or keynote parent before adding a note."
      });
      return;
    }
    if (isEntryRemotelyClaimed(parent)) {
      setStatus({ status: "warning", message: editClaimTitle(parent) || "This keynote parent is being edited by another user." });
      renderAll();
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
    setSelectionForEntry(entry);
    markDirty();
    syncLocalEditClaims();
    renderAll();
  }

  function addNoteInSequence() {
    var selectedNote = selectedNoteEntry();
    var missingParentMessage = "Select a division or keynote row before adding a note in sequence.";

    if (!selectedNote) {
      addNoteUnderParent(selectedDivisionEntry(), missingParentMessage);
      return;
    }

    addNoteInSequenceForEntry(selectedNote);
  }

  function addNoteInSequenceForEntry(entry) {
    var parentKey = trim(entry && entry.parentKey);
    var parent = parentKey ? entriesByKey()[parentKey] || null : null;
    var missingParentMessage = "Select a keynote with a parent before adding a note in sequence.";

    if (entry) {
      parentKey = trim(entry.parentKey);
      missingParentMessage = parentKey
        ? "Parent key '" + parentKey + "' was not found. Fix the parent before adding a note in sequence."
        : "Select a keynote with a parent before adding a note in sequence.";
    }

    addNoteUnderParent(parent, missingParentMessage);
  }

  function addSubNote() {
    addNoteUnderParent(
      selectedNoteEntry(),
      "Select a keynote row before adding a sub-note."
    );
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
    if (blockForSafeMode("Duplicating keynotes")) {
      return;
    }
    if (isEntryRemotelyClaimed(entry)) {
      setStatus({ status: "warning", message: editClaimTitle(entry) || "This keynote is being edited by another user." });
      renderAll();
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
    syncLocalEditClaims();
    renderAll();
  }

  function deleteEntry(entry) {
    if (!entry) {
      return;
    }
    if (blockForSafeMode("Deleting keynotes")) {
      return;
    }
    if (isEntryRemotelyClaimed(entry)) {
      setStatus({ status: "warning", message: editClaimTitle(entry) || "This keynote is being edited by another user." });
      renderAll();
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
    delete state.collapsedEntryIds[entry.id];

    if (state.selectedNoteId === entry.id) {
      state.selectedNoteId = null;
    }
    if (state.selectedDivisionId === entry.id) {
      state.selectedDivisionId = null;
    }
    state.selectedId = null;
    ensureSelection();
    markDirty();
    syncLocalEditClaims();
    renderAll();
  }

  function deleteSelected() {
    var entry = actionTargetEntry();
    var deleteButton = byId("delete-row");

    if (entry && !trim(entry.parentKey) && deleteButton) {
      openParentActionMenu(entry, deleteButton, true);
      return;
    }
    deleteEntry(entry);
  }

  function moveNoteToDivision(entry, division) {
    if (!entry || !division || trim(division.parentKey)) {
      setStatus({ status: "error", message: "Select a valid division before moving this keynote." });
      return;
    }
    if (blockForSafeMode("Moving keynotes")) {
      return;
    }
    if (isEntryRemotelyClaimed(entry)) {
      setStatus({ status: "warning", message: editClaimTitle(entry) || "This keynote is being edited by another user." });
      renderAll();
      return;
    }
    if (trim(entry.parentKey) === trim(division.key)) {
      setStatus({
        status: "ready",
        message: "Keynote '" + (entry.key || "new keynote") + "' is already in division " + divisionTitle(division) + "."
      });
      return;
    }

    entry.parentKey = division.key;
    setSelectionForEntry(entry);
    markDirty();
    syncLocalEditClaims();
    renderAll();
    setStatus({
      status: "ready",
      message: "Moved keynote '" + (entry.key || "new keynote") + "' to division " + divisionTitle(division) + "."
    });
  }

  function promoteNoteToParent(entry) {
    if (!entry || !trim(entry.parentKey)) {
      setStatus({ status: "warning", message: "This keynote is already a parent." });
      return;
    }
    if (blockForSafeMode("Promoting keynotes")) {
      return;
    }
    if (isEntryRemotelyClaimed(entry)) {
      setStatus({ status: "warning", message: editClaimTitle(entry) || "This keynote is being edited by another user." });
      renderAll();
      return;
    }

    entry.parentKey = "";
    setSelectionForEntry(entry);
    markDirty();
    syncLocalEditClaims();
    renderAll();
    setStatus({
      status: "ready",
      message: "Promoted keynote '" + (entry.key || "new keynote") + "' to a parent."
    });
  }

  function demoteParentToNote(entry, destination) {
    if (!entry || trim(entry.parentKey) || !destination || trim(destination.parentKey) || entry.id === destination.id) {
      setStatus({ status: "error", message: "Select another valid parent before demoting this parent." });
      return;
    }
    if (blockForSafeMode("Demoting parents")) {
      return;
    }
    if (isEntryRemotelyClaimed(entry)) {
      setStatus({ status: "warning", message: editClaimTitle(entry) || "This parent is being edited by another user." });
      renderAll();
      return;
    }

    entry.parentKey = destination.key;
    setSelectionForEntry(entry);
    markDirty();
    syncLocalEditClaims();
    renderAll();
    setStatus({
      status: "ready",
      message: "Demoted parent '" + (entry.key || "new parent") + "' to a note under parent " + divisionTitle(destination) + "."
    });
  }

  function deleteParentAndSubnotes(entry) {
    var entriesToDelete;
    var claimedEntry;
    var deleteIds = {};

    if (!entry || trim(entry.parentKey)) {
      return;
    }
    if (blockForSafeMode("Deleting parents")) {
      return;
    }

    entriesToDelete = [entry].concat(descendantEntriesFor(entry));
    claimedEntry = firstRemotelyClaimedEntry(entriesToDelete);
    if (claimedEntry) {
      setStatus({
        status: "warning",
        message: editClaimTitle(claimedEntry) || "A keynote in this parent is being edited by another user."
      });
      renderAll();
      return;
    }

    entriesToDelete.forEach(function (candidate) {
      deleteIds[candidate.id] = true;
      delete state.collapsedEntryIds[candidate.id];
    });
    state.entries = state.entries.filter(function (candidate) {
      return !deleteIds[candidate.id];
    });
    if (deleteIds[state.selectedDivisionId]) {
      state.selectedDivisionId = null;
    }
    if (deleteIds[state.selectedNoteId]) {
      state.selectedNoteId = null;
    }
    state.selectedId = state.selectedNoteId || state.selectedDivisionId;
    ensureSelection();
    markDirty();
    syncLocalEditClaims();
    renderAll();
    setStatus({
      status: "ready",
      message: "Deleted parent '" + (entry.key || "new parent") + "' and " +
        formatNumber(entriesToDelete.length - 1) + " subnote(s)."
    });
  }

  function deleteParentAndMoveSubnotes(entry, destination) {
    var directChildren;
    var claimedEntry;

    if (
      !entry ||
      trim(entry.parentKey) ||
      !destination ||
      trim(destination.parentKey) ||
      entry.id === destination.id
    ) {
      setStatus({ status: "error", message: "Select another valid parent before moving these subnotes." });
      return;
    }
    if (blockForSafeMode("Deleting parents")) {
      return;
    }

    directChildren = directChildEntriesFor(entry);
    claimedEntry = firstRemotelyClaimedEntry([entry].concat(directChildren));
    if (claimedEntry) {
      setStatus({
        status: "warning",
        message: editClaimTitle(claimedEntry) || "A keynote being moved is currently edited by another user."
      });
      renderAll();
      return;
    }

    directChildren.forEach(function (child) {
      child.parentKey = destination.key;
    });
    state.entries = state.entries.filter(function (candidate) {
      return candidate.id !== entry.id;
    });
    delete state.collapsedEntryIds[entry.id];
    setSelectionForEntry(destination);
    markDirty();
    syncLocalEditClaims();
    renderAll();
    setStatus({
      status: "ready",
      message: "Deleted parent '" + (entry.key || "new parent") + "' and moved its " +
        formatNumber(directChildren.length) + " direct subnote(s) to parent " + divisionTitle(destination) + "."
    });
  }

  function uppercaseNoteText(entry) {
    var upperText;

    if (!entry) {
      return;
    }
    upperText = text(entry.text).toUpperCase();
    if (upperText === entry.text) {
      setStatus({
        status: "ready",
        message: "Text for keynote '" + (entry.key || "new keynote") + "' is already uppercase."
      });
      return;
    }

    updateEntryField(entry.id, "text", upperText, true);
    setStatus({
      status: "ready",
      message: "Converted text for keynote '" + (entry.key || "new keynote") + "' to uppercase."
    });
  }

  function refreshData() {
    var shouldAllowLoad = state.dirty;

    if (!confirmDiscardChanges("Refresh will discard unsaved keynote edits in this window. Continue?")) {
      return;
    }

    setStatus({ status: "warning", message: "Refreshing keynote file..." });
    clearLocalEditClaims().then(function () {
      if (postWebViewMessage({ type: "refreshData" })) {
        state.allowNextLoad = shouldAllowLoad;
      } else {
        state.allowNextLoad = false;
      }
    });
  }

  function requestRefresh() {
    refreshData();
  }

  function collectAnalytics() {
    var settings = (state.payload && state.payload.supabase) || {};
    var db = dbManager();

    if (state.analyticsCollecting) {
      return;
    }

    if (!state.payload || !state.payload.libraryKey) {
      setStatus({ status: "error", message: "No keynote library is available for analytics." });
      return;
    }

    if (!settings.configured) {
      setStatus({ status: "warning", message: "Supabase is not configured, so analytics cannot be collected." });
      return;
    }

    if (state.dbInitializing) {
      setStatus({ status: "warning", message: "Supabase is still attaching. Try collecting analytics again in a moment." });
      return;
    }

    if (!state.dbReady) {
      setStatus({ status: "warning", message: "Supabase is not ready. Refresh or configure Supabase before collecting analytics." });
      return;
    }

    if (!db || typeof db.syncAnalytics !== "function") {
      setStatus({ status: "warning", message: "The Supabase analytics API did not load." });
      return;
    }

    state.analyticsCollecting = true;
    renderSaveState();
    setStatus({ status: "warning", message: "Collecting keynote analytics from Revit..." });

    if (!postWebViewMessage({ type: "collectAnalytics" })) {
      state.analyticsCollecting = false;
      renderSaveState();
    }
  }

  function collectAnalyticsOnOpen() {
    if (state.analyticsRequestedOnOpen || !state.dbReady) {
      return;
    }

    state.analyticsRequestedOnOpen = true;
    collectAnalytics();
  }

  function handleAnalyticsResult(result) {
    var analytics;

    result = result || {};
    analytics = result.analytics || null;

    if ((result.status || "") !== "ready" || !analytics) {
      state.analyticsCollecting = false;
      state.operationIssues = (result.issues || []).map(function (issue) {
        return makeIssue(
          "warning",
          issue.message || result.message || "Could not collect keynote analytics.",
          issue.key || "",
          issue.code || "analyticsCollectionFailed"
        );
      });
      renderValidation();
      renderSaveState();
      setStatus({
        status: result.status || "error",
        message: result.message || "Could not collect keynote analytics."
      });
      return;
    }

    applyAnalyticsResultToUi(analytics, result.modelHealth || (analytics && analytics.modelHealth));
    setStatus({ status: "syncing", message: "Syncing keynote analytics to Supabase..." });

    syncAnalyticsResult(analytics).then(function (syncResult) {
      state.analyticsCollecting = false;
      state.operationIssues = [];
      renderValidation();
      renderSaveState();
      setStatus({
        status: "ready",
        message: (result.message || "Collected keynote analytics.") + " Supabase sync complete."
      });
      return syncResult;
    }).catch(function (error) {
      state.analyticsCollecting = false;
      state.operationIssues = [makeIssue(
        "warning",
        "Collected keynote analytics, but Supabase sync failed: " + (error.message || error),
        "",
        "analyticsSyncFailed"
      )];
      renderValidation();
      renderSaveState();
      setStatus({
        status: "warning",
        message: "Collected keynote analytics, but Supabase sync failed."
      });
    });
  }

  function saveData() {
    var issues = validateAll();
    var payload;

    renderValidation();

    if (blockForSafeMode("Save")) {
      return;
    }

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

    state.pendingDbChanges = buildPendingDbChanges();

    payload = {
      keynotePath: state.payload.keynotePath,
      encoding: state.payload.encoding || "utf-8",
      lineEnding: state.payload.lineEnding || "\r\n",
      lastWriteUtc: state.payload.lastWriteUtc,
      fileHash: state.payload.fileHash,
      sourceHasMalformed: hasSourceIssueCode("malformedLine"),
      modelIssueResolutions: Object.keys(state.modelIssueResolutions).map(function (resolutionId) {
        return state.modelIssueResolutions[resolutionId];
      }),
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
    state.operationIssues = [];
    renderSaveState();
    setStatus({ status: "syncing", message: "Merging edits into the shared keynote file..." });

    if (!postWebViewMessage({ type: "saveKeynotes", payload: payload })) {
      state.saving = false;
      state.pendingDbChanges = null;
      renderSaveState();
    }
  }

  function handleSaveResult(result) {
    var pendingChanges = state.pendingDbChanges;
    var fileMessage;

    result = result || {};
    state.saving = false;

    if ((result.status || "") === "conflict") {
      state.pendingDbChanges = null;
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
      state.operationIssues = (result.issues || []).filter(function (issue) {
        return !sourceIssueBlocksSave(issue);
      });
      rememberBaseline();
      fileMessage = result.backupPath
        ? (result.message || "") + " Backup: " + result.backupPath
        : (result.message || "");
      setStatus({
        status: "ready",
        message: fileMessage
      });
      clearLocalEditClaims();
      renderValidation();
      savePendingDbChanges(result.payload, pendingChanges, fileMessage);
      return;
    } else if ((result.status || "") !== "ready") {
      state.pendingDbChanges = null;
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

    clearLocalEditClaims().then(function () {
      if (postWebViewMessage({ type: "closeWindow", discardConfirmed: discardConfirmed })) {
        return;
      }

      try {
        globalScope.close();
      } catch (ignore) {
        // Browser fallback only.
      }
    });
  }

  function configureSupabase() {
    if (state.dirty && !confirmDiscardChanges("Changing Supabase settings will reload keynote data and discard unsaved edits. Continue?")) {
      return;
    }

    setStatus({ status: "warning", message: "Opening Supabase settings..." });
    clearLocalEditClaims().then(function () {
      if (postWebViewMessage({ type: "configureSupabase" })) {
        state.allowNextLoad = state.dirty;
      }
    });
  }

  function bindDivisionInput(id, fieldName) {
    var input = byId(id);
    if (!input) {
      return;
    }

    if (fieldName === "key") {
      input.addEventListener("dblclick", function (event) {
        event.preventDefault();
        unlockKeyInput(input);
      });
    }

    input.addEventListener("input", function () {
      var entry = selectedDivisionEntry();
      if (entry) {
        updateEntryField(entry.id, fieldName, input.value, false);
      }
    });
    input.addEventListener("blur", function () {
      if (fieldName === "key") {
        lockKeyInput(input);
      }
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
    var placementFilterSelect = byId("placement-filter-select");
    var placementModeSelect = byId("placement-mode-select");

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

    if (placementModeSelect) {
      syncPlacementModeSelect();
      placementModeSelect.addEventListener("change", function () {
        setPlacementMode(placementModeSelect.value);
        postWebViewMessage({
          type: "placementModeChanged",
          placementMode: state.placementMode
        });
        renderNotes();
      });
    }

    if (placementFilterSelect) {
      syncPlacementFilterSelect();
      placementFilterSelect.addEventListener("change", function () {
        setPlacementFilter(placementFilterSelect.value);
        renderAll();
      });
    }

    document.addEventListener("click", function () {
      closeRowActionMenu(false);
    });
    document.addEventListener("keydown", function (event) {
      if (event.key === "Escape" && activeRowActionMenu) {
        event.preventDefault();
        closeRowActionMenu(true);
      }
    });
    globalScope.addEventListener("resize", function () {
      closeRowActionMenu(false);
    });
    if (byId("notes-section-body")) {
      byId("notes-section-body").addEventListener("scroll", function () {
        closeRowActionMenu(false);
      });
    }
    if (document.querySelector(".division-list-wrap")) {
      document.querySelector(".division-list-wrap").addEventListener("scroll", function () {
        closeRowActionMenu(false);
      });
    }

    bindDivisionInput("division-key-input", "key");
    bindDivisionInput("division-text-input", "text");

    bindClick("division-more-button", function (event) {
      var entry = selectedDivisionEntry();
      if (entry) {
        event.preventDefault();
        event.stopPropagation();
        openParentActionMenu(entry, byId("division-more-button"), false);
      }
    });

    bindClick("add-root", addRoot);
    bindClick("add-sequence", addNoteInSequence);
    bindClick("add-sub-note", addSubNote);
    bindClick("duplicate-row", duplicateSelected);
    bindClick("delete-row", deleteSelected);
    bindClick("toggle-divisions", toggleDivisionsSidebar);
    bindClick("warning-pill", toggleWarningSidebar);
    bindClick("model-health-pill", function () {
      setModelIssuesOpen(!state.modelIssuesOpen);
    });
    bindClick("review-model-issues", function () {
      setModelIssuesOpen(true);
    });
    bindClick("close-model-issues", function () {
      setModelIssuesOpen(false);
    });
    bindClick("model-issues-scrim", function () {
      setModelIssuesOpen(false);
    });
    bindClick("acknowledge-model-health", acknowledgeModelHealthReview);
    bindClick("close-validation-sidebar", function () {
      setWarningSidebarOpen(false);
    });
    bindClick("configure-supabase", configureSupabase);
    bindClick("refresh-data", refreshData);
    bindClick("collect-analytics", collectAnalytics);
    bindClick("save-data", saveData);
    bindClick("close-window", closeWindow);

    renderAll();
    postWebViewMessage({ type: "appReady" });
  }

  globalScope.ffeKeynotes = {
    loadData: loadData,
    setStatus: setStatus,
    handleSaveResult: handleSaveResult,
    handleAnalyticsResult: handleAnalyticsResult,
    requestRefresh: requestRefresh
  };

  if (typeof document !== "undefined") {
    document.addEventListener("DOMContentLoaded", init);
  }
}(typeof window !== "undefined" ? window : this));
