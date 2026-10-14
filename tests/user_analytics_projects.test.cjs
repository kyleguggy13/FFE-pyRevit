// Run: node --test tests/user_analytics_projects.test.cjs
// Exercise the production rendering and bridge handlers with a small DOM fixture.
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const { test } = require('node:test');

const support = path.join(__dirname,
  '../FFE-pyRevit.extension/FFE-pyRevit.tab/Analytics.panel/UserAnalytics.pushbutton/support');
const html = fs.readFileSync(path.join(support, 'index.html'), 'utf8');
const source = fs.readFileSync(path.join(support, 'site.js'), 'utf8');

function element() {
  let content = '';
  return {
    children: [], events: {}, attributes: {}, style: {}, value: '', hidden: false,
    disabled: false, clientWidth: 900, scrollTop: 0,
    classList: { add() {}, remove() {} },
    appendChild(child) { this.children.push(child); return child; },
    removeChild(child) {
      this.children.splice(this.children.indexOf(child), 1);
      if (!this.children.length) this.scrollTop = 0;
    },
    get firstChild() { return this.children[0] || null; },
    get textContent() { return content + this.children.map(child => child.textContent).join(''); },
    set textContent(value) { content = String(value); this.children = []; },
    addEventListener(name, callback) { this.events[name] = callback; },
    setAttribute(name, value) { this.attributes[name] = String(value); },
    getAttribute(name) { return this.attributes[name]; },
    focus() { this.focused = true; },
    click() { if (!this.disabled) this.events.click(); }
  };
}

function harness() {
  const controls = {};
  for (const match of html.matchAll(/id="([^"]+)"/g)) controls[match[1]] = element();
  const pages = Array.from(html.matchAll(/<section[^>]*data-analytics-page="([^"]+)"[^>]*>/g), match => {
    const section = element();
    section.setAttribute('data-analytics-page', match[1]);
    section.hidden = match[0].includes(' hidden');
    return section;
  });
  const document = {
    events: {},
    getElementById: id => controls[id] || null,
    createElement: element,
    querySelector: () => element(),
    querySelectorAll: selector => {
      assert.equal(selector, '[data-analytics-page]');
      return pages;
    },
    addEventListener(name, callback) { this.events[name] = callback; }
  };
  const messages = [];
  const windowEvents = {};
  let chartRenders = 0;
  function assertChartVisible() {
    assert.ok(pages.filter(page => page.getAttribute('data-analytics-page') === 'dashboard')
      .every(page => !page.hidden), 'charts must render after their containers become visible');
    chartRenders++;
  }
  const context = {
    document,
    Option: function(label, value) {
      const option = element(); option.textContent = label; option.value = value; return option;
    },
    chrome: { webview: { postMessage(message) { messages.push(JSON.parse(message)); } } },
    Chart: function() {
      assertChartVisible(); this.resize = assertChartVisible; this.update = assertChartVisible; this.destroy = () => {};
    },
    addEventListener(name, callback) { windowEvents[name] = callback; },
    scrollTo() {}, close() {},
    setTimeout(callback) { callback(); return 1; }, clearTimeout() {}
  };
  context.window = context;
  const hooks = 'globalScope.testApi = { state, buildModelIdentity, normalizeEntries, buildProjects };\n';
  vm.runInNewContext(source.replace('  globalScope.ffeAnalytics = {', hooks + '  globalScope.ffeAnalytics = {'), context);
  controls['range-filter'].value = 'all';
  controls['action-filter'].value = 'all';
  document.events.DOMContentLoaded();
  return {
    controls, pages, messages, windowEvents, api: context.testApi,
    load(entries, status = 'ready') {
      context.ffeAnalytics.loadData({ username: 'Tester', generatedAt: '2026-10-14 12:00:00', status, entries });
    },
    filter(id, value) {
      controls[id].value = value;
      controls[id].events[id === 'document-search' ? 'input' : 'change']();
    },
    select(name) {
      const button = controls['project-list'].children.find(button => button.children[0].textContent === name);
      assert.ok(button, 'project button exists: ' + name);
      button.click();
    },
    chartRenders: () => chartRenders
  };
}

function entry(docPath, overrides = {}) {
  return Object.assign({ datetime: '2026-10-14 11:00:00', doc_path: docPath,
    doc_title: 'Model', action: 'sync', status: 'success' }, overrides);
}

function history(count = 205) {
  return Array.from({ length: count }, (_, index) => entry('Autodesk Docs://Alpha/' + (index % 2 ? 'A.rvt' : 'B.rvt'), {
    doc_title: index % 2 ? 'A' : 'B', action: index % 2 ? 'sync' : 'open', family_name: 'Event ' + index
  }));
}

test('cloud identities normalize prefixes, case, whitespace and separators without guessing membership', () => {
  const { api } = harness();
  const first = api.buildModelIdentity(' Autodesk Docs://  Alpha   Project / Folder / Model.rvt ', 'Model');
  const second = api.buildModelIdentity('bim 360:\\\\alpha project\\folder\\MODEL.RVT', 'Renamed title');
  assert.equal(first.projectName, 'Alpha Project');
  assert.equal(first.projectKey, second.projectKey);
  assert.equal(first.modelKey, second.modelKey);
  assert.notEqual(first.modelKey, api.buildModelIdentity('Autodesk Docs://Alpha Project/Other/Model.rvt', 'Model').modelKey);
  for (const value of ['', '<Untitled>', 'C:\\Models\\Alpha Project\\Model.rvt',
    'RSN://Server/Alpha Project/Model.rvt', 'Autodesk Docs://Alpha', 'BIM 360:///Model.rvt',
    'Autodesk Docs://Alpha/', 'Autodesk Docs://Alpha/../Model.rvt']) {
    assert.equal(api.buildModelIdentity(value, 'Model').projectKey, 'unassigned', value);
  }
  assert.equal(api.buildModelIdentity('', '').modelLabel, 'Unknown Model');
  assert.equal(api.buildModelIdentity('<Untitled>', '  Office  ').modelKey,
    api.buildModelIdentity('', 'office').modelKey);
  assert.equal(api.buildModelIdentity('C:\\Models\\Office.rvt', '').modelLabel, 'Office.rvt');
});

test('project grouping keeps full model paths, title fallbacks and Unassigned separate', () => {
  const { api } = harness();
  const projects = api.buildProjects(api.normalizeEntries([
    entry('Autodesk Docs://Zulu/A.rvt'),
    entry('Autodesk Docs://Alpha/Folder1/Same.rvt'),
    entry('BIM 360://alpha/folder1/same.rvt'),
    entry('Autodesk Docs://Alpha/Folder2/Same.rvt'),
    entry('Autodesk Docs://Unassigned/A.rvt'),
    entry('C:\\Models\\Same.rvt'),
    entry('', { doc_title: 'Unsaved' }), entry('<Untitled>', { doc_title: 'unsaved' }),
    entry('', { doc_title: '' }), entry('Autodesk Docs://__proto__/A.rvt')
  ]));
  assert.equal(projects.at(-1).key, 'unassigned');
  const alpha = projects.find(project => project.key === 'cloud:alpha');
  assert.equal(alpha.entries.length, 3);
  assert.equal(alpha.models.length, 2);
  assert.equal(alpha.models.find(model => model.key.includes('folder1')).count, 2);
  assert.equal(projects.find(project => project.key === 'unassigned').models.length, 3);
  assert.ok(projects.find(project => project.key === 'cloud:unassigned'));
  assert.ok(projects.find(project => project.key === 'cloud:__proto__'));
});

test('project activity exposes every action through pagination and keeps Recent Activity capped', () => {
  const h = harness();
  h.load(history());
  assert.equal(h.api.state.selectedProjectKey, 'cloud:alpha');
  assert.equal(h.controls['project-action-total'].textContent, '205');
  assert.equal(h.controls['project-model-total'].textContent, '2');
  assert.equal(h.controls['project-actions-breakdown'].children.length, 2);
  assert.equal(h.controls['recent-table-body'].children.length, 80);
  const seen = [];
  for (const [page, expected] of [[0, '1-100 of 205 actions'], [1, '101-200 of 205 actions'], [2, '201-205 of 205 actions']]) {
    assert.equal(h.controls['project-page-range'].textContent, expected);
    seen.push(...h.controls['project-table-body'].children.map(row => row.children[4].textContent));
    if (page < 2) h.controls['project-next'].click();
  }
  assert.equal(seen.length, 205);
  assert.equal(new Set(seen).size, 205);
  assert.equal(seen[0], 'Event 204');
  assert.equal(seen.at(-1), 'Event 0');
  assert.ok(h.controls['project-next'].disabled);
  h.controls['project-previous'].click();
  assert.equal(h.controls['project-page-range'].textContent, '101-200 of 205 actions');
  h.filter('project-model-filter', h.api.state.projects[0].models.find(model => model.label === 'A').key);
  assert.equal(h.controls['project-page-range'].textContent, '1-100 of 102 actions');
  assert.equal(h.controls['project-actions-breakdown'].textContent, 'sync: 102');
  assert.equal(h.controls['project-action-total'].textContent, '205');
});

test('shared date, action and document filters update project counts and reset pagination', () => {
  const h = harness();
  h.load(history().concat([
    entry('Autodesk Docs://Beta/Office.rvt', { doc_title: 'Office' }),
    entry('Autodesk Docs://Old/Office.rvt', { doc_title: 'Office', datetime: '2025-01-01 12:00:00' })
  ]));
  h.controls['project-next'].click();
  h.filter('action-filter', 'sync');
  assert.equal(h.api.state.projectPage, 0);
  assert.equal(h.controls['project-action-total'].textContent, '102');
  assert.equal(h.controls['project-model-total'].textContent, '1');
  h.filter('document-search', 'Office');
  assert.equal(h.api.state.selectedProjectKey, 'cloud:beta');
  assert.equal(h.controls['projects-subtitle'].textContent, '2 projects');
  h.filter('range-filter', '7');
  assert.equal(h.controls['projects-subtitle'].textContent, '1 projects');
  assert.equal(h.controls['project-action-total'].textContent, '1');
  assert.equal(h.controls['project-model-filter'].children.length, 2);
  h.filter('document-search', 'no matching document');
  assert.equal(h.api.state.selectedProjectKey, null);
  assert.equal(h.controls['projects-subtitle'].textContent, '0 projects');
  assert.equal(h.controls['project-page-range'].textContent, '0 of 0 actions');
  assert.ok(h.controls['project-model-filter'].disabled);
  assert.ok(h.controls['project-next'].disabled);
  assert.ok(h.controls['project-previous'].disabled);
  assert.equal(h.controls['project-table-body'].textContent, 'No activity in current filter.');
});

test('project selection resets model and pagination; refresh preserves or repairs selections', () => {
  const h = harness();
  const entries = history().concat(entry('Autodesk Docs://Beta/Office.rvt', { doc_title: 'Office' }));
  h.load(entries);
  h.controls['project-next'].click();
  h.controls['project-list'].scrollTop = 240;
  h.select('Beta');
  assert.equal(h.controls['project-list'].scrollTop, 240);
  assert.equal(h.api.state.projectPage, 0);
  assert.equal(h.controls['project-title'].textContent, 'Beta');
  assert.equal(h.controls['project-list'].children[1].getAttribute('aria-pressed'), 'true');
  h.filter('project-model-filter', h.api.state.projects[1].models[0].key);
  h.load(entries);
  assert.equal(h.api.state.selectedProjectKey, 'cloud:beta');
  assert.notEqual(h.api.state.selectedModelKey, 'all');
  h.load(history());
  assert.equal(h.api.state.selectedProjectKey, 'cloud:alpha');
  assert.equal(h.api.state.selectedModelKey, 'all');
  const modelA = h.api.state.projects[0].models.find(model => model.label === 'A').key;
  h.filter('project-model-filter', modelA);
  h.load(history().filter(row => row.doc_title === 'B'));
  assert.equal(h.api.state.selectedModelKey, 'all');
  h.load(history());
  h.controls['project-next'].click();
  h.load(history());
  assert.equal(h.api.state.projectPage, 1);
  h.load(history(10));
  assert.equal(h.api.state.projectPage, 0);
  assert.equal(h.controls['project-page-range'].textContent, '1-10 of 10 actions');
});

test('navigation preserves shared filters and uses the unchanged refresh/close bridge', () => {
  const h = harness();
  assert.equal(h.messages[0].type, 'appReady');
  assert.match(html, /id="recent-activity"[^\n]*\n\s*<button id="projects-page"/);
  h.load(history());
  h.filter('document-search', 'A');
  h.controls['projects-page'].click();
  assert.equal(h.controls['page-title'].textContent, 'Projects');
  assert.equal(h.controls['projects-page'].textContent, 'Back to Analytics');
  assert.equal(h.controls['recent-activity'].textContent, 'Recent Activity');
  assert.ok(h.pages.every(page => page.hidden === (page.getAttribute('data-analytics-page') !== 'projects')));
  const renders = h.chartRenders();
  h.windowEvents.resize();
  h.controls['refresh-data'].click();
  assert.equal(h.messages.at(-1).type, 'refreshData');
  h.load(history());
  assert.equal(h.chartRenders(), renders);
  assert.equal(h.api.state.currentPage, 'projects');
  assert.equal(h.controls['document-search'].value, 'A');
  assert.equal(h.controls['refresh-data'].textContent, 'Refresh');
  h.controls['recent-activity'].click();
  assert.equal(h.api.state.currentPage, 'recent');
  assert.equal(h.controls['projects-page'].textContent, 'Projects');
  h.controls['projects-page'].click();
  assert.equal(h.api.state.currentPage, 'projects');
  h.controls['projects-page'].click();
  assert.equal(h.api.state.currentPage, 'dashboard');
  assert.equal(h.controls['page-title'].textContent, 'User Analytics');
  assert.ok(h.chartRenders() > renders);
  assert.equal(h.controls['document-search'].value, 'A');
  h.controls['close-window'].click();
  assert.equal(h.messages.at(-1).type, 'closeWindow');
});

test('empty and missing logs render a usable Projects page', () => {
  const h = harness();
  h.load([], 'missingFile');
  h.controls['projects-page'].click();
  assert.equal(h.controls['project-action-total'].textContent, '0');
  assert.equal(h.controls['project-model-total'].textContent, '0');
  assert.equal(h.controls['project-list'].textContent, 'No projects in current filter.');
  assert.equal(h.controls['project-actions-breakdown'].textContent, 'No actions in current filter.');
  assert.equal(h.controls['state-banner'].hidden, false);
  h.load([entry('', { doc_title: '' })]);
  assert.equal(h.controls['project-title'].textContent, 'Unassigned');
  assert.equal(h.controls['project-table-body'].children[0].children[2].children[0].textContent, 'Unknown Model');
});

test('action breakdown toggles filter only model activity and preserve all available action counts', () => {
  const h = harness();
  const actionButton = action => h.controls['project-actions-breakdown'].children
    .find(button => button.textContent.startsWith(action + ':'));
  h.load(history());
  h.controls['project-next'].click();
  assert.equal(actionButton('sync').type, 'button');
  assert.equal(actionButton('sync').getAttribute('aria-pressed'), 'false');
  actionButton('sync').click();
  assert.equal(h.controls['project-page-range'].textContent, '1-100 of 102 actions');
  assert.equal(actionButton('sync').getAttribute('aria-pressed'), 'true');
  assert.equal(actionButton('sync').getAttribute('aria-controls'), 'project-table-body');
  assert.ok(actionButton('sync').focused);
  assert.equal(actionButton('open').textContent, 'open: 103');
  assert.equal(actionButton('sync').textContent, 'sync: 102');
  assert.ok(h.controls['project-table-body'].children.every(row => row.children[1].textContent === 'sync'));
  assert.equal(h.controls['action-filter'].value, 'all');
  assert.equal(h.controls['project-action-total'].textContent, '205');
  assert.equal(h.controls['recent-subtitle'].textContent, '205 entries');
  h.controls['project-next'].click();
  assert.equal(h.controls['project-page-range'].textContent, '101-102 of 102 actions');
  actionButton('open').click();
  assert.equal(h.controls['project-page-range'].textContent, '1-100 of 205 actions');
  assert.equal(actionButton('open').getAttribute('aria-pressed'), 'true');
  assert.equal(actionButton('sync').getAttribute('aria-pressed'), 'true');
  actionButton('sync').click();
  assert.equal(h.controls['project-activity-subtitle'].textContent, '103 entries');
  assert.ok(h.controls['project-table-body'].children.every(row => row.children[1].textContent === 'open'));
  actionButton('open').click();
  assert.equal(h.controls['project-activity-subtitle'].textContent, '205 entries');
  assert.ok(h.controls['project-actions-breakdown'].children.every(button => button.getAttribute('aria-pressed') === 'false'));
});

test('action selection survives refresh and navigation, and clears actions unavailable in the new scope', () => {
  const h = harness();
  const rows = history().concat(entry('Autodesk Docs://Beta/Office.rvt', { action: 'sync' }));
  const actionButton = action => h.controls['project-actions-breakdown'].children
    .find(button => button.textContent.startsWith(action + ':'));
  h.load(rows);
  actionButton('sync').click();
  h.controls['project-next'].click();
  h.load(rows);
  assert.equal(h.controls['project-page-range'].textContent, '101-102 of 102 actions');
  assert.equal(actionButton('sync').getAttribute('aria-pressed'), 'true');
  h.controls['projects-page'].click();
  h.controls['recent-activity'].click();
  h.controls['projects-page'].click();
  assert.equal(actionButton('sync').getAttribute('aria-pressed'), 'true');
  h.filter('project-model-filter', h.api.state.projects[0].models.find(model => model.label === 'B').key);
  assert.equal(h.api.state.selectedProjectActions.length, 0);
  assert.equal(h.controls['project-activity-subtitle'].textContent, '103 entries');
  h.filter('project-model-filter', 'all');
  actionButton('sync').click();
  h.filter('action-filter', 'open');
  assert.equal(h.api.state.selectedProjectActions.length, 0);
  assert.equal(h.controls['project-activity-subtitle'].textContent, '103 entries');
  h.filter('action-filter', 'all');
  actionButton('sync').click();
  h.select('Beta');
  assert.equal(actionButton('sync').getAttribute('aria-pressed'), 'false');
  assert.equal(h.api.state.selectedProjectActions.length, 0);
  actionButton('sync').click();
  h.load(history().filter(row => row.action === 'open'));
  assert.equal(h.api.state.selectedProjectKey, 'cloud:alpha');
  assert.equal(h.api.state.selectedProjectActions.length, 0);
  assert.equal(h.controls['project-activity-subtitle'].textContent, '103 entries');
  h.load([]);
  assert.equal(h.controls['project-actions-breakdown'].textContent, 'No actions in current filter.');
});
