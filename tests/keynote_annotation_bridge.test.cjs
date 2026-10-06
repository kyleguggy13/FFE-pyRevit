// Run: node --test tests/keynote_annotation_bridge.test.cjs
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const { test } = require('node:test');
const support = path.join(__dirname, '../FFE-pyRevit.extension/FFE-pyRevit.tab/DesignTools.panel/Keynotes.pushbutton/support');

function harness(db = {}) {
  const context = { Promise, console, setTimeout, clearTimeout, setInterval, clearInterval };
  vm.createContext(context);
  const hooks = `
    globalScope.testApi = { state: state, saveData: saveData, setPlacementMode: setPlacementMode,
      attachSupabaseLibrary: attachSupabaseLibrary, processRemoteEntryChange: processRemoteEntryChange,
      ensureLibraryBeforeAnalyticsSync: ensureLibraryBeforeAnalyticsSync,
      subscribeToLibrary: subscribeToLibrary,
      normalizeModelHealth: normalizeModelHealth, renderModelHealth: renderModelHealth,
      createModelIssueRepairControls: createModelIssueRepairControls,
      requestModelIssueRepair: requestModelIssueRepair, handleModelIssueRepairResult: handleModelIssueRepairResult,
      handleFamilySyncResult: handleFamilySyncResult, requestFamilySync: requestFamilySync,
      requestProjectSetup: requestProjectSetup, handleStorageResult: handleStorageResult,
      bindStorageModeControls: bindStorageModeControls, renderStorageModeControls: renderStorageModeControls,
      setSettingsOpen: setSettingsOpen,
      requestLibraryAssociation: requestLibraryAssociation,
      projectSetupActive: projectSetupActive, renderProjectSetup: renderProjectSetup,
      openProjectSetup: openProjectSetup, closeProjectSetup: closeProjectSetup,
      requestReconnectFile: requestReconnectFile, renderKeynoteRecovery: renderKeynoteRecovery,
      useSetupLoadBoundary: function () {
        loadData = function (payload) { globalScope.loadedPayload = payload; state.payload = payload; };
      },
      setDb: function (db) { dbManager = function () { return db; }; }
    };
    globalScope.messages = [];
    postWebViewMessage = function (message) { globalScope.messages.push(message); return true; };
    setStatus = function (message) { globalScope.status = message; };
    renderAll = renderMeta = renderValidation = renderSaveState = syncPlacementModeControl = function () {};
    applySnapshotMetadata = function (snapshot) { state.dbSnapshot = snapshot; };
    applyData = function (payload) {
      state.loadGeneration += 1; state.payload = payload; state.entries = payload.entries;
      state.baselineEntries = payload.entries; state.dirty = false;
    };
    subscribeToLibrary = subscribeToEntries = subscribeToAnalytics = subscribeToEditClaims = function () {};
    refreshEditClaims = syncLocalEditClaims = refreshOtherModelUsage = collectAnalyticsOnOpen = collectAnalytics = function () {};
    clearLocalEditClaims = function () { return Promise.resolve(); };
    globalScope.testApi.flush = function () { return new Promise(function (resolve) { setTimeout(resolve, 0); }); };
  `;
  const source = fs.readFileSync(path.join(support, 'site.js'), 'utf8');
  vm.runInContext(source.replace('  globalScope.ffeKeynotes = {', hooks + '\n  globalScope.ffeKeynotes = {'), context);
  context.testApi.setDb(db);
  const state = context.testApi.state;
  state.payload = { storageMode: 'annotation', libraryKey: 'annotation:central', libraryId: 'library', datasetVersion: 1,
    supabase: { configured: true }, entries: [{ id: 'row-id', dbId: 'row-id', rowVersion: 1, key: 'A', text: 'Original', parentKey: '' }] };
  state.dbSnapshot = state.payload;
  state.dbReady = true;
  state.entries = [Object.assign({}, state.payload.entries[0], { text: 'Updated' })];
  state.baselineEntries = state.payload.entries;
  state.dirty = true;
  return context;
}

function repairHarness() {
  const h = harness();
  h.testApi.state.payload.documentKey = 'central.rvt';
  function element(tag) {
    return { tag, children: [], attributes: {}, events: {}, value: '', disabled: false, scrollTop: 0,
      appendChild(child) { this.children.push(child); },
      removeChild(child) { this.children.splice(this.children.indexOf(child), 1); },
      get firstChild() { return this.children[0] || null; },
      setAttribute(name, value) { this.attributes[name] = value; },
      addEventListener(name, callback) { this.events[name] = callback; } };
  }
  h.list = element('div');
  h.document = { createElement: element, getElementById: id => id === 'model-issues-list' ? h.list : null };
  return h;
}

function settingsHarness(mode = 'file') {
  const h = harness();
  h.testApi.state.payload.storageMode = mode;
  h.testApi.state.dirty = false;
  h.controls = {};
  for (const id of ['storage-mode', 'apply-storage-mode', 'settings-dialog',
                    'supabase-project-url', 'supabase-publishable-key']) {
    h.controls[id] = { value: '', disabled: false, events: {},
      addEventListener(name, callback) { this.events[name] = callback; } };
  }
  const dialog = h.controls['settings-dialog'];
  dialog.showModal = () => { dialog.open = true; };
  dialog.close = () => { dialog.open = false; };
  h.document = { getElementById: id => h.controls[id] || null };
  h.testApi.bindStorageModeControls();
  h.testApi.setSettingsOpen(true);
  h.testApi.renderStorageModeControls();
  return h;
}

test('Storage Mode selection enables Apply without starting conversion, and survives redraws', () => {
  for (const mode of ['file', 'annotation']) {
    const h = settingsHarness(mode);
    const select = h.controls['storage-mode'];
    const apply = h.controls['apply-storage-mode'];
    assert.equal(select.value, mode);
    assert.equal(apply.disabled, true);
    apply.events.click();
    assert.equal(h.messages.length, 0);
    select.value = mode === 'file' ? 'annotation' : 'file';
    select.events.change();
    assert.equal(apply.disabled, false);
    assert.equal(h.testApi.state.payload.storageMode, mode);
    assert.equal(h.messages.length, 0);
    h.testApi.renderStorageModeControls();
    assert.equal(select.value, mode === 'file' ? 'annotation' : 'file');
    assert.equal(apply.disabled, false);
    select.value = mode;
    select.events.change();
    assert.equal(apply.disabled, true);
  }
});

test('Apply sends the selected mode once and clears the staged selection after success', async () => {
  const h = settingsHarness();
  h.testApi.useSetupLoadBoundary();
  const select = h.controls['storage-mode'];
  const apply = h.controls['apply-storage-mode'];
  select.value = 'annotation';
  select.events.change();
  apply.events.click();
  apply.events.click();
  await h.testApi.flush();
  assert.equal(h.messages.length, 1);
  assert.equal(h.messages[0].payload.storageMode, 'annotation');
  assert.equal(h.messages[0].payload.createFile, false);
  assert.equal(h.controls['settings-dialog'].open, false);
  h.testApi.renderStorageModeControls();
  assert.equal(apply.disabled, true);
  assert.equal(select.disabled, true);
  h.testApi.handleStorageResult({ status: 'ready', payload: {
    status: 'ready', storageMode: 'annotation', entries: [] } });
  h.testApi.setSettingsOpen(true);
  h.testApi.renderStorageModeControls();
  assert.equal(select.value, 'annotation');
  assert.equal(apply.disabled, true);
});

test('Apply stays disabled during work and project setup', () => {
  for (const flag of ['storageBusy', 'saving', 'familySyncing', 'dbInitializing', 'analyticsCollecting', 'manualSetup']) {
    const h = settingsHarness();
    h.controls['storage-mode'].value = 'annotation';
    h.controls['storage-mode'].events.change();
    h.testApi.state[flag] = true;
    h.testApi.renderStorageModeControls();
    assert.equal(h.controls['apply-storage-mode'].disabled, true);
    assert.equal(h.controls['storage-mode'].disabled, true);
    h.controls['apply-storage-mode'].events.click();
    assert.equal(h.messages.length, 0);
  }
});

test('closing Settings discards an unapplied storage choice; canceling discard retains it for retry', () => {
  const h = settingsHarness();
  const select = h.controls['storage-mode'];
  const apply = h.controls['apply-storage-mode'];
  select.value = 'annotation';
  select.events.change();
  h.testApi.state.dirty = true;
  h.confirm = () => false;
  apply.events.click();
  h.testApi.renderStorageModeControls();
  assert.equal(h.messages.length, 0);
  assert.equal(h.testApi.state.pendingStorageMode, 'annotation');
  assert.equal(apply.disabled, false);
  assert.equal(h.controls['settings-dialog'].open, true);
  h.testApi.setSettingsOpen(false);
  h.testApi.setSettingsOpen(true);
  h.testApi.renderStorageModeControls();
  assert.equal(select.value, 'file');
  assert.equal(apply.disabled, true);
});

function repairIssue(action = 'renameType', secondText = 'Same description') {
  return { severity: 'warning', key: '00.00', code: action === 'renameType'
    ? 'genericAnnotationTypeNameMismatch' : 'genericAnnotationDuplicateTypes',
    repair: { action, types: action === 'renameType'
      ? [{ id: '9876543210', name: 'Long descriptive name', key: '00.00', text: 'Same description' }]
      : [{ id: '1', name: '00.00', key: '00.00', text: 'Same description' },
         { id: '2', name: 'Alternate descriptive name', key: '00.00', text: secondText }] } };
}

test('bound file attachment passes the UUID and uses the resolved canonical key before mirroring', async () => {
  let ensured;
  let mirrored;
  const h = harness({ configure() {}, ensureLibrary: async payload => {
    ensured = payload;
    return { status: 'ready', libraryId: 'existing-uuid', libraryKey: 'original.txt', fileHash: 'old-hash', entries: [] };
  }, syncFileSnapshot: async payload => {
    mirrored = payload;
    return { status: 'ready', libraryId: 'existing-uuid', libraryKey: 'original.txt', fileHash: payload.fileHash, entries: payload.entries };
  } });
  const payload = { status: 'ready', storageMode: 'file', libraryId: 'existing-uuid', libraryKey: 'cached-key.txt',
    fileLibraryKey: 'moved.txt', keynotePath: 'Moved.txt', fileHash: 'new-hash', entries: [{ key: 'A', text: 'Note', parentKey: '' }],
    issues: [], supabase: { configured: true } };
  h.testApi.state.payload = payload;
  h.testApi.attachSupabaseLibrary(payload, 'load');
  await h.testApi.flush();
  assert.equal(ensured.libraryId, 'existing-uuid');
  assert.equal(ensured.fileLibraryKey, 'moved.txt');
  assert.equal(mirrored.libraryKey, 'original.txt');
  assert.equal(payload.libraryKey, 'original.txt');
  assert.equal(payload.libraryId, 'existing-uuid');
});

test('analytics file attachment keeps the UUID and active path after a file move', async () => {
  let ensured;
  const db = { ensureLibrary: async payload => { ensured = payload; return {}; } };
  const h = harness(db);
  h.testApi.state.payload = { storageMode: 'file', libraryId: 'model-uuid',
    libraryKey: 'original.txt', fileLibraryKey: 'moved.txt' };
  for (const analytics of [
    { storageMode: 'file', libraryId: 'scan-uuid', libraryKey: 'original.txt', fileLibraryKey: 'moved.txt' },
    { storageMode: 'file', libraryKey: 'original.txt' }
  ]) {
    await h.testApi.ensureLibraryBeforeAnalyticsSync(db, analytics);
    assert.equal(ensured.libraryId, analytics.libraryId || 'model-uuid');
    assert.equal(ensured.fileLibraryKey, 'moved.txt');
    assert.equal(ensured.libraryKey, 'original.txt');
  }
});

test('adapter resolves file attachment by UUID with a separate current file identity', async () => {
  let request;
  const context = { Promise, supabase: { createClient: () => ({ rpc: async (name, args) => {
    request = { name, args };
    return { data: { status: 'ready', libraryId: 'uuid', libraryKey: 'original.txt' } };
  } }) } };
  vm.createContext(context);
  vm.runInContext(fs.readFileSync(path.join(support, 'db_manager.js'), 'utf8'), context);
  context.ffeKeynoteDb.configure({ url: 'https://example.supabase.co', anonKey: 'test' });
  await context.ffeKeynoteDb.ensureLibrary({ libraryId: 'uuid', libraryKey: 'original.txt', fileLibraryKey: 'renamed.txt' });
  assert.equal(request.name, 'ensure_keynote_library_association');
  assert.equal(request.args.p_library_id, 'uuid');
  assert.equal(request.args.p_library_key, 'original.txt');
  assert.equal(request.args.p_file_key, 'renamed.txt');
  await context.ffeKeynoteDb.ensureLibrary({ libraryKey: 'unbound.txt' });
  assert.equal(request.args.p_library_id, null);
});

test('library association actions preserve edits on cancellation and post one native request when accepted', async () => {
  const h = settingsHarness();
  h.testApi.state.dirty = true;
  const entries = h.testApi.state.entries;
  h.confirm = () => false;
  h.testApi.requestLibraryAssociation('choose');
  await h.testApi.flush();
  assert.equal(h.messages.length, 0);
  assert.equal(h.testApi.state.entries, entries);
  h.confirm = () => true;
  h.testApi.requestLibraryAssociation('choose');
  h.testApi.requestLibraryAssociation('choose');
  await h.testApi.flush();
  assert.equal(h.messages.length, 1);
  assert.equal(h.messages[0].type, 'changeLibraryAssociation');
  assert.equal(h.messages[0].payload.operation, 'choose');
  h.testApi.handleStorageResult({ status: 'canceled', message: 'Canceled' });
  assert.equal(h.testApi.state.storageBusy, false);
  assert.equal(h.testApi.state.entries, entries);
  assert.equal(h.testApi.state.dirty, true);
});

test('type-name warning renders a repair button even without a corresponding library row', () => {
  const h = repairHarness();
  h.testApi.state.modelHealth = h.testApi.normalizeModelHealth({ issues: [repairIssue()] });
  h.testApi.renderModelHealth();
  const card = h.list.children.find(child => child.className === 'model-issue-item');
  assert.equal(card.tag, 'div');
  const button = card.children.at(-1).children.at(-1);
  assert.equal(button.textContent, 'Rename Type to 00.00');
  button.events.click();
  button.events.click();
  assert.equal(h.messages.length, 1);
  assert.equal(h.messages[0].type, 'repairModelIssue');
  assert.equal(h.messages[0].payload.documentKey, 'central.rvt');
  assert.equal(h.messages[0].payload.types[0].id, '9876543210');
  assert.equal(h.testApi.state.storageBusy, true);
  assert.equal(h.testApi.state.dirty, true);
});

test('duplicate repair asks for distinct keys and preserves each description in the request', () => {
  const h = repairHarness();
  const controls = h.testApi.createModelIssueRepairControls(repairIssue('renumberDuplicateTypes', 'Different text'));
  const labels = controls.children.filter(child => child.tag === 'label');
  const first = labels[0].children.at(-1);
  const second = labels[1].children.at(-1);
  const button = controls.children.at(-1);
  assert.equal(first.value, '00.00');
  assert.equal(second.value, '00.00');
  assert.equal(button.disabled, true);
  button.events.click();
  assert.equal(h.messages.length, 0);
  assert.match(labels[1].children[0].textContent, /Different text/);
  second.value = '00.01';
  second.events.input();
  assert.equal(button.disabled, false);
  button.events.click();
  assert.equal(h.messages[0].payload.action, 'renumberDuplicateTypes');
  assert.deepEqual(JSON.parse(JSON.stringify(h.messages[0].payload.newKeys)),
    [{ id: '1', key: '00.00' }, { id: '2', key: '00.01' }]);
  assert.equal(h.messages[0].payload.types[1].text, 'Different text');
  assert.equal(h.messages[0].payload.keepTypeId, undefined);
});

test('identical descriptions still require user-entered unique keys', () => {
  const h = repairHarness();
  const issue = repairIssue('renumberDuplicateTypes');
  const controls = h.testApi.createModelIssueRepairControls(issue);
  const labels = controls.children.filter(child => child.tag === 'label');
  const second = labels[1].children.at(-1);
  const button = controls.children.at(-1);
  assert.equal(button.disabled, true);
  for (const invalid of ['', ' ', ' 00.00 ', '00\t01']) {
    second.value = invalid;
    second.events.input();
    assert.equal(button.disabled, true);
    button.events.click();
  }
  assert.equal(h.messages.length, 0);
  labels[0].children.at(-1).value = '00.02';
  second.value = '00.03';
  second.events.input();
  assert.equal(button.disabled, false);
  button.events.click();
  assert.deepEqual(Array.from(h.messages[0].payload.newKeys, item => item.key), ['00.02', '00.03']);
});

test('legacy destructive duplicate repair cannot be sent to Revit', () => {
  const h = repairHarness();
  h.testApi.requestModelIssueRepair(repairIssue('mergeDuplicateTypes'), '1');
  assert.equal(h.messages.length, 0);
});

test('repairs are blocked during other operations and remote refresh cannot replace a pending repair', () => {
  const h = repairHarness();
  h.testApi.state.storageBusy = true;
  h.testApi.requestModelIssueRepair(repairIssue());
  h.testApi.processRemoteEntryChange();
  assert.equal(h.messages.length, 0);
  assert.equal(h.testApi.state.remoteEntriesPending, true);
  assert.equal(h.testApi.createModelIssueRepairControls(repairIssue()).children.at(-1).disabled, true);
  h.testApi.handleModelIssueRepairResult({ status: 'ready', modelHealth: { issues: [] } });
  assert.notEqual(h.testApi.state.remoteEntriesTimer, null);
  clearTimeout(h.testApi.state.remoteEntriesTimer);
});

test('repair responses refresh issues while preserving unsaved library edits and open panel', () => {
  const h = repairHarness();
  const entries = h.testApi.state.entries;
  h.testApi.state.modelIssuesOpen = true;
  h.testApi.requestModelIssueRepair(repairIssue());
  h.testApi.handleModelIssueRepairResult({ status: 'warning', message: 'Ownership failure',
    modelHealth: { issues: [repairIssue()] } });
  assert.equal(h.testApi.state.storageBusy, false);
  assert.equal(h.testApi.state.modelHealth.issues.length, 1);
  assert.equal(h.testApi.state.syncIssues[0].code, 'modelIssueRepairFailed');
  h.testApi.handleModelIssueRepairResult({ status: 'ready', message: 'Renamed type', modelHealth: { issues: [] } });
  assert.equal(h.testApi.state.modelHealth.issues.length, 0);
  assert.equal(h.testApi.state.syncIssues.length, 0);
  assert.equal(h.testApi.state.entries, entries);
  assert.equal(h.testApi.state.dirty, true);
  assert.equal(h.testApi.state.modelIssuesOpen, true);
  assert.equal(h.messages.length, 1);
});

test('database commit occurs before any family update and never invokes local save', async () => {
  let resolveSave;
  let request;
  const h = harness({ saveAnnotationChanges: payload => { request = payload; return new Promise(resolve => { resolveSave = resolve; }); } });
  h.testApi.saveData();
  assert.equal(h.messages.length, 0);
  assert.equal(h.testApi.state.dirty, true);
  assert.equal(request.baseDatasetVersion, 1);
  assert.equal(request.changes.upserts[0].text, 'Updated');
  resolveSave(Object.assign({}, h.testApi.state.payload, { status: 'ready', datasetVersion: 2, entries: h.testApi.state.entries }));
  await h.testApi.flush();
  assert.equal(h.testApi.state.dirty, false);
  assert.equal(h.messages.length, 1);
  assert.equal(h.messages[0].type, 'syncAnnotationFamily');
  assert.equal(h.messages[0].payload.datasetVersion, 2);
  assert.equal(h.messages[0].payload.baselineEntries[0].text, 'Original');
});

test('failed cloud save retains unsaved edits and never saves into RVT', async () => {
  const h = harness({ saveAnnotationChanges: async () => { throw new Error('offline'); } });
  h.testApi.saveData();
  await h.testApi.flush();
  assert.equal(h.testApi.state.dirty, true);
  assert.equal(h.testApi.state.entries[0].text, 'Updated');
  assert.equal(h.testApi.state.baselineEntries[0].text, 'Original');
  assert.equal(h.messages.length, 0);
});

test('unavailable connection cannot use a local save fallback', () => {
  const h = harness();
  h.testApi.state.dbReady = false;
  h.testApi.saveData();
  assert.equal(h.messages.length, 0);
  assert.equal(h.testApi.state.dirty, true);
  assert.match(h.status.message, /Connect to Supabase/);
});

test('concurrent database edit returns a conflict without losing local changes', async () => {
  const h = harness({ saveAnnotationChanges: async () => ({ status: 'conflict', message: 'Other user saved.' }) });
  h.testApi.saveData();
  await h.testApi.flush();
  assert.equal(h.testApi.state.dirty, true);
  assert.equal(h.messages.length, 0);
  assert.equal(h.status.status, 'conflict');
});

test('failed family update leaves cloud save intact and keeps retry request', () => {
  const h = harness();
  h.testApi.state.dirty = false;
  h.testApi.state.pendingFamilySync = { libraryKey: 'annotation:central', datasetVersion: 2 };
  h.testApi.handleFamilySyncResult({ status: 'warning', message: 'Type is not editable.' });
  assert.equal(h.testApi.state.dirty, false);
  assert.equal(h.testApi.state.pendingFamilySync.datasetVersion, 2);
  h.testApi.requestFamilySync();
  assert.equal(h.messages[0].type, 'syncAnnotationFamily');
});

test('cloud attachment never seeds or repairs from local entries', () => {
  const h = harness({ configure() {}, ensureLibrary() { assert.fail('Local seed attempted'); },
    syncFileSnapshot() { assert.fail('File mirror attempted'); } });
  h.testApi.attachSupabaseLibrary(h.testApi.state.payload, 'load');
  assert.equal(h.testApi.state.dbReady, true);
  h.testApi.setPlacementMode('userKeynote');
  assert.equal(h.testApi.state.placementMode, 'genericAnnotation');
});

test('family retry uses refreshed cloud revision and retains the original rename baseline', () => {
  const h = harness();
  h.testApi.state.dirty = false;
  h.testApi.state.payload.datasetVersion = 3;
  h.testApi.state.pendingFamilySync = { libraryKey: 'annotation:central', datasetVersion: 2,
    baselineEntries: [{ id: 'row-id', key: 'OLD' }], modelIssueResolutions: [] };
  h.testApi.requestFamilySync();
  assert.equal(h.messages[0].payload.datasetVersion, 3);
  assert.equal(h.messages[0].payload.baselineEntries[0].key, 'OLD');
});

test('saving again after failed family updates preserves earlier pending renames', async () => {
  let h;
  h = harness({ saveAnnotationChanges: async () => Object.assign({}, h.testApi.state.payload,
    { status: 'ready', datasetVersion: 3, entries: h.testApi.state.entries }) });
  h.testApi.state.pendingFamilySync = { libraryKey: 'annotation:central', datasetVersion: 2,
    baselineEntries: [{ id: 'row-id', key: 'OLD' }], modelIssueResolutions: [] };
  h.testApi.saveData();
  await h.testApi.flush();
  assert.equal(h.messages[0].payload.baselineEntries[0].key, 'OLD');
  assert.equal(h.messages[0].payload.datasetVersion, 3);
});

test('realtime changes reload the authoritative library without a Revit worksharing gate', () => {
  const h = harness();
  h.testApi.state.dirty = false;
  h.testApi.processRemoteEntryChange();
  assert.equal(h.messages[0].type, 'refreshData');
});

test('missing text file does not erase its Supabase mirror', () => {
  const h = harness({ ensureLibrary() { assert.fail('Seeded unreadable source'); } });
  h.testApi.attachSupabaseLibrary({ storageMode: 'file', libraryKey: 'file.txt',
    issues: [{ severity: 'error', code: 'missingFile' }] }, 'load');
  assert.equal(h.testApi.state.dbReady, false);
});

test('file attachment uses the matching hash with metadata rows without repeatedly mirroring', async () => {
  const metadata = { libraryId: 'library', datasetVersion: 7, fileHash: 'same-hash',
    entries: [{ key: 'A', dbId: 'row-id', rowVersion: 2, sortOrder: 0 }] };
  let mirrors = 0;
  const h = harness({ configure() {}, ensureLibrary: async () => metadata,
    syncFileSnapshot: async () => { mirrors += 1; return metadata; } });
  const payload = { storageMode: 'file', libraryKey: 'notes.txt', fileHash: 'same-hash',
    entries: [{ key: 'A', text: 'Actual description', parentKey: 'DIVISION' }], supabase: { configured: true } };
  h.testApi.state.payload = payload;
  h.testApi.state.dirty = false;
  for (let refresh = 0; refresh < 3; refresh += 1) {
    h.testApi.attachSupabaseLibrary(payload, 'load');
    await h.testApi.flush();
    assert.equal(h.testApi.state.dbReady, true);
  }
  assert.equal(mirrors, 0);
  assert.equal(h.testApi.state.dbSnapshot, metadata);
});

test('file attachment still mirrors changed hashes, missing keys, and different full row content', async () => {
  const payload = { storageMode: 'file', libraryKey: 'notes.txt', fileHash: 'file-hash',
    entries: [{ key: 'A', text: 'Current text', parentKey: '' }], supabase: { configured: true } };
  const snapshots = [
    { fileHash: 'older-hash', entries: [{ key: 'A' }] },
    { fileHash: 'file-hash', entries: [{ key: 'OTHER' }] },
    { fileHash: 'file-hash', entries: [] },
    { fileHash: 'file-hash', entries: [{ key: 'A', text: 'Old text', parentKey: '' }] },
    { entries: [{ key: 'A' }] }
  ];
  for (const snapshot of snapshots) {
    let mirrors = 0;
    const h = harness({ configure() {}, ensureLibrary: async () => snapshot,
      syncFileSnapshot: async () => { mirrors += 1; return { libraryId: 'library', entries: [] }; } });
    h.testApi.state.payload = payload;
    h.testApi.attachSupabaseLibrary(payload, 'load');
    await h.testApi.flush();
    assert.equal(mirrors, 1);
  }
});

test('library subscription passes the loaded revision to the realtime adapter', () => {
  let handlers;
  const h = harness({ subscribeLibrary: (id, client, value) => { handlers = value; } });
  h.testApi.subscribeToLibrary({ libraryId: 'library', datasetVersion: 7 });
  assert.equal(handlers.datasetVersion, 7);
});

test('analytics metadata cannot start a refresh loop while remote keynote saves still notify', () => {
  let receive;
  let refreshes = 0;
  const channel = { on(event, filter, callback) { receive = callback; return this; }, subscribe() { return this; } };
  const context = { supabase: { createClient: () => ({ channel: () => channel, removeChannel() {} }) } };
  vm.createContext(context);
  vm.runInContext(fs.readFileSync(path.join(support, 'db_manager.js'), 'utf8'), context);
  const db = context.ffeKeynoteDb;
  db.configure({ url: 'https://example.supabase.co', anonKey: 'test' });
  db.subscribeLibrary('library', 'local', { datasetVersion: 7, onRemoteChange() { refreshes += 1; } });
  const notify = (version, client = 'other') => receive({ old: { id: 'library' },
    new: { id: 'library', dataset_version: version, last_saved_by_client_id: client, document_title: 'Project' } });
  // Analytics retains the last keynote saver and only changes document metadata.
  for (let update = 0; update < 5; update += 1) { notify(7); }
  assert.equal(refreshes, 0);
  notify(8);
  assert.equal(refreshes, 1);
  notify(8);
  notify(7);
  assert.equal(refreshes, 1);
  notify(9, 'local');
  notify(9, 'other');
  assert.equal(refreshes, 1);
  db.subscribeLibrary('library', 'local', { datasetVersion: 10 });
  notify(10);
  assert.equal(refreshes, 1);
  notify(11);
  assert.equal(refreshes, 2);
  db.unsubscribe();
  db.subscribeLibrary('different-library', 'local', { datasetVersion: 1, onRemoteChange() { refreshes += 1; } });
  notify(2);
  assert.equal(refreshes, 3);
});

test('adapter invokes authoritative save endpoint with database version', async () => {
  const calls = [];
  const context = { Promise, supabase: { createClient: () => ({ rpc: async (name, args) => {
    calls.push({ name, args }); return { data: { status: 'ready', libraryId: 'library', entries: [] } };
  } }) } };
  vm.createContext(context);
  vm.runInContext(fs.readFileSync(path.join(support, 'db_manager.js'), 'utf8'), context);
  context.ffeKeynoteDb.configure({ url: 'https://example.supabase.co', anonKey: 'test' });
  await context.ffeKeynoteDb.saveAnnotationChanges({ libraryKey: 'annotation:x', baseDatasetVersion: 4 });
  assert.equal(calls[0].name, 'save_annotation_keynote_changes');
  assert.equal(calls[0].args.p_base_dataset_version, 4);
});

function setupHarness(db = {}) {
  const h = harness(db);
  h.document = { getElementById: () => null, querySelector: () => null };
  h.testApi.state.payload = { documentKey: 'central.rvt', storageMode: 'file', entries: [],
    projectSetup: { state: 'required', canContinue: true, message: 'Choose how this project will use keynotes.' } };
  h.testApi.state.dirty = false;
  h.testApi.state.dbReady = false;
  h.testApi.useSetupLoadBoundary();
  return h;
}

test('each setup choice posts the appropriate storage action and ignores duplicate submissions', async () => {
  for (const mode of ['file', 'annotation']) {
    const h = setupHarness();
    h.testApi.state.setupMode = mode;
    h.testApi.requestProjectSetup();
    h.testApi.requestProjectSetup();
    await h.testApi.flush();
    assert.equal(h.messages.length, 1);
    assert.equal(h.messages[0].type, 'setupStorage');
    assert.equal(h.messages[0].payload.storageMode, mode);
    assert.equal(h.messages[0].payload.createFile, mode === 'file');
    assert.equal(h.messages[0].payload.projectSetup, true);
    assert.equal(h.testApi.state.storageBusy, true);
  }
});

test('unsaved, blocked, and unselected projects cannot start setup', async () => {
  for (const setup of [{ state: 'required', canContinue: false }, { state: 'blocked', canContinue: false }]) {
    const h = setupHarness();
    h.testApi.state.payload.projectSetup = setup;
    h.testApi.state.setupMode = 'file';
    h.testApi.requestProjectSetup();
    await h.testApi.flush();
    assert.equal(h.messages.length, 0);
  }
  const h = setupHarness();
  h.testApi.requestProjectSetup();
  assert.equal(h.messages.length, 0);
});

test('cancellation and failure retain setup selection and allow a retry', async () => {
  for (const status of ['canceled', 'error']) {
    const h = setupHarness();
    h.testApi.state.setupMode = 'file';
    h.testApi.state.storageBusy = true;
    h.testApi.handleStorageResult({ status, message: 'Try again' });
    assert.equal(h.testApi.projectSetupActive(), true);
    assert.equal(h.testApi.state.setupMode, 'file');
    assert.equal(h.testApi.state.storageBusy, false);
    assert.equal(h.status.message, 'Try again');
    h.testApi.requestProjectSetup();
    await h.testApi.flush();
    assert.equal(h.messages.length, 1);
  }
});

test('successful setup opens the editor only after a ready payload is returned', () => {
  const h = setupHarness();
  h.testApi.state.setupMode = 'annotation';
  h.testApi.state.storageBusy = true;
  const payload = { status: 'ready', storageMode: 'annotation', libraryId: 'library', entries: [],
    projectSetup: { state: 'complete', canContinue: false } };
  h.testApi.handleStorageResult({ status: 'ready', payload });
  assert.equal(h.loadedPayload, payload);
  assert.equal(h.testApi.projectSetupActive(), false);
  assert.equal(h.testApi.state.storageBusy, false);
});

test('a successful envelope with a failed source payload stays on setup', () => {
  const h = setupHarness();
  h.testApi.state.storageBusy = true;
  h.testApi.handleStorageResult({ status: 'ready', payload: { status: 'error', message: 'Could not load source' } });
  assert.equal(h.loadedPayload, undefined);
  assert.equal(h.testApi.projectSetupActive(), true);
  assert.equal(h.status.status, 'error');
  assert.equal(h.status.message, 'Could not load source');
});

test('template download preserves the setup request and remains busy through Save As', async () => {
  const h = setupHarness({ configure() {}, getTemplate: async () => ({ content: 'A\tAlpha\n' }) });
  h.testApi.state.setupMode = 'file';
  h.testApi.state.storageBusy = true;
  h.testApi.handleStorageResult({ status: 'needsTemplate', request: {
    storageMode: 'file', createFile: true, projectSetup: true } });
  await h.testApi.flush();
  assert.equal(h.messages.length, 1);
  assert.equal(h.messages[0].payload.template.content, 'A\tAlpha\n');
  assert.equal(h.messages[0].payload.projectSetup, true);
  assert.equal(h.testApi.state.storageBusy, true);
});

test('failed template download keeps setup available for retry', async () => {
  const h = setupHarness({ configure() {}, getTemplate: async () => { throw new Error('Offline'); } });
  h.testApi.state.setupMode = 'file';
  h.testApi.state.storageBusy = true;
  h.testApi.handleStorageResult({ status: 'needsTemplate', request: { storageMode: 'file', createFile: true } });
  await h.testApi.flush();
  assert.equal(h.testApi.state.storageBusy, false);
  assert.equal(h.testApi.state.setupMode, 'file');
  assert.equal(h.testApi.projectSetupActive(), true);
  assert.equal(h.messages.length, 0);
});

test('annotation setup can download a division template before Revit asks to merge or remove model notes', async () => {
  const h = setupHarness({ configure() {}, getTemplate: async () => ({ content: 'DIVISION 22\tPLUMBING\n' }) });
  h.testApi.state.setupMode = 'annotation';
  h.testApi.state.storageBusy = true;
  h.testApi.handleStorageResult({ status: 'needsTemplate', request: {
    storageMode: 'annotation', createFile: false, projectSetup: true } });
  await h.testApi.flush();
  assert.equal(h.messages.length, 1);
  assert.equal(h.messages[0].type, 'setupStorage');
  assert.equal(h.messages[0].payload.storageMode, 'annotation');
  assert.equal(h.messages[0].payload.createFile, false);
  assert.equal(h.messages[0].payload.projectSetup, true);
  assert.match(h.messages[0].payload.template.content, /DIVISION 22/);
  assert.equal(h.testApi.state.storageBusy, true);
});

test('setup rendering hides the workspace and enables Continue only when eligible', () => {
  const h = setupHarness();
  function element(value) {
    return { value, hidden: false, disabled: false, attrs: {},
      setAttribute(name, value) { this.attrs[name] = value; },
      classList: { toggle() {} } };
  }
  const ids = Object.fromEntries(['project-setup', 'setup-revit-keynotes', 'setup-annotation-keynotes',
    'project-setup-continue', 'project-setup-message', 'project-setup-next'].map(id => [id, element()]));
  ids['setup-revit-keynotes'].value = 'file';
  ids['setup-annotation-keynotes'].value = 'annotation';
  const workspace = element();
  h.document = { getElementById: id => ids[id], querySelector: selector => selector === '.workspace' ? workspace : element() };
  h.testApi.renderProjectSetup();
  assert.equal(ids['project-setup'].hidden, false);
  assert.equal(workspace.hidden, true);
  assert.equal(workspace.inert, true);
  assert.equal(ids['project-setup-continue'].disabled, true);
  h.testApi.state.setupMode = 'file';
  h.testApi.renderProjectSetup();
  assert.equal(ids['project-setup-continue'].disabled, false);
  assert.equal(ids['setup-revit-keynotes'].checked, true);
  h.testApi.state.storageBusy = true;
  h.testApi.renderProjectSetup();
  assert.equal(ids['project-setup-continue'].disabled, true);
  assert.equal(ids['setup-annotation-keynotes'].disabled, true);
  h.testApi.state.storageBusy = false;
  h.testApi.state.payload.projectSetup = { state: 'complete' };
  h.testApi.renderProjectSetup();
  assert.equal(ids['project-setup'].hidden, true);
  assert.equal(workspace.hidden, false);
});

function recoveryHarness(status = 'error') {
  const h = setupHarness();
  h.testApi.state.payload.status = status;
  h.testApi.state.payload.documentKeySource = 'centralPath';
  h.testApi.state.payload.message = 'Keynote reference could not be read.';
  h.testApi.state.payload.projectSetup = { state: 'complete', canContinue: false };
  return h;
}

test('missing files offer reconnect while other load failures also offer setup', () => {
  const h = recoveryHarness();
  const controls = Object.fromEntries(['keynote-recovery-actions', 'open-project-setup', 'reconnect-keynote-file']
    .map(id => [id, {}]));
  h.document.getElementById = id => controls[id];
  for (const status of ['error', 'missingFile', 'unsupported', 'invalidFormat']) {
    h.testApi.state.payload.status = status;
    h.testApi.renderKeynoteRecovery();
    assert.equal(controls['keynote-recovery-actions'].hidden, false);
    assert.equal(controls['open-project-setup'].hidden, status === 'missingFile');
    assert.equal(controls['reconnect-keynote-file'].disabled, false);
    if (status === 'missingFile') {
      h.testApi.openProjectSetup();
      assert.equal(h.testApi.state.manualSetup, false);
    }
  }
  h.testApi.state.storageBusy = true;
  h.testApi.renderKeynoteRecovery();
  assert.equal(controls['open-project-setup'].disabled, true);
  assert.equal(controls['reconnect-keynote-file'].disabled, true);
  h.testApi.state.payload.status = 'ready';
  h.testApi.state.storageBusy = false;
  h.testApi.renderKeynoteRecovery();
  assert.equal(controls['keynote-recovery-actions'].hidden, true);
});

test('manual setup bypasses failed automatic detection for either keynote choice', async () => {
  for (const mode of ['file', 'annotation']) {
    const h = recoveryHarness();
    h.testApi.openProjectSetup();
    assert.equal(h.testApi.projectSetupActive(), true);
    h.testApi.state.setupMode = mode;
    h.testApi.requestProjectSetup();
    h.testApi.requestProjectSetup();
    await h.testApi.flush();
    assert.equal(h.messages.length, 1);
    assert.equal(h.messages[0].payload.recoverySetup, true);
    assert.equal(h.messages[0].payload.projectSetup, undefined);
    assert.equal(h.messages[0].payload.storageMode, mode);
    assert.equal(h.messages[0].payload.createFile, mode === 'file');
  }
});

test('manual setup can open after a blocked automatic check and supports Back without changing data', () => {
  const h = recoveryHarness();
  h.testApi.state.payload.projectSetup = { state: 'blocked', canContinue: false };
  const payload = h.testApi.state.payload;
  h.testApi.openProjectSetup();
  assert.equal(h.testApi.state.manualSetup, true);
  h.testApi.closeProjectSetup();
  assert.equal(h.testApi.state.manualSetup, false);
  assert.equal(h.testApi.state.payload, payload);
  assert.equal(h.messages.length, 0);
});

test('manual setup in unsaved projects cannot submit until the model is saved', () => {
  const h = recoveryHarness();
  h.testApi.state.payload.documentKeySource = 'title';
  h.testApi.openProjectSetup();
  h.testApi.state.setupMode = 'file';
  h.testApi.requestProjectSetup();
  assert.equal(h.messages.length, 0);
  assert.match(h.status.message, /Save the Revit project/);
});

test('Reconnect sends one dedicated action without requiring a cloud connection', async () => {
  const h = recoveryHarness('missingFile');
  h.testApi.requestReconnectFile();
  h.testApi.requestReconnectFile();
  await h.testApi.flush();
  assert.equal(h.messages.length, 1);
  assert.equal(h.messages[0].type, 'reconnectFile');
  assert.equal(h.testApi.state.storageBusy, true);
});

test('reconnect is available during setup and cancellation retains the selected mode', async () => {
  const h = recoveryHarness();
  h.testApi.openProjectSetup();
  h.testApi.state.setupMode = 'file';
  h.testApi.requestReconnectFile();
  await h.testApi.flush();
  h.testApi.handleStorageResult({ status: 'canceled', message: 'Reconnection canceled' });
  assert.equal(h.testApi.state.manualSetup, true);
  assert.equal(h.testApi.state.setupMode, 'file');
  assert.equal(h.testApi.state.storageBusy, false);
  h.testApi.requestReconnectFile();
  await h.testApi.flush();
  assert.equal(h.messages.length, 2);
});

test('successful reconnect exits manual setup and failed loading retains recovery', () => {
  const h = recoveryHarness();
  h.testApi.openProjectSetup();
  h.testApi.handleStorageResult({ status: 'error', message: 'File unavailable' });
  assert.equal(h.testApi.state.manualSetup, true);
  const payload = { status: 'ready', storageMode: 'file', keynotePath: 'notes.txt', entries: [],
    projectSetup: { state: 'complete' } };
  h.testApi.handleStorageResult({ status: 'ready', payload });
  assert.equal(h.testApi.state.manualSetup, false);
  assert.equal(h.testApi.projectSetupActive(), false);
  assert.equal(h.loadedPayload, payload);
});

test('canceling discard before reconnection preserves edits and never opens a picker', async () => {
  const h = recoveryHarness();
  h.testApi.state.dirty = true;
  h.confirm = () => false;
  h.testApi.requestReconnectFile();
  await h.testApi.flush();
  assert.equal(h.messages.length, 0);
  assert.equal(h.testApi.state.dirty, true);
  assert.equal(h.testApi.state.storageBusy, false);
});
