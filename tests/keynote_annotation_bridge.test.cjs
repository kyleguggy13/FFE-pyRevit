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
      handleFamilySyncResult: handleFamilySyncResult, requestFamilySync: requestFamilySync,
      requestProjectSetup: requestProjectSetup, handleStorageResult: handleStorageResult,
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

test('load failures and missing files expose both recovery actions, ready libraries hide them', () => {
  const h = recoveryHarness();
  const controls = Object.fromEntries(['keynote-recovery-actions', 'open-project-setup', 'reconnect-keynote-file']
    .map(id => [id, {}]));
  h.document.getElementById = id => controls[id];
  for (const status of ['error', 'missingFile', 'unsupported', 'invalidFormat']) {
    h.testApi.state.payload.status = status;
    h.testApi.renderKeynoteRecovery();
    assert.equal(controls['keynote-recovery-actions'].hidden, false);
    assert.equal(controls['open-project-setup'].hidden, false);
    assert.equal(controls['reconnect-keynote-file'].disabled, false);
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
    const h = recoveryHarness('missingFile');
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
