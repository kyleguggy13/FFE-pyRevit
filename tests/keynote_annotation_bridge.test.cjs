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
