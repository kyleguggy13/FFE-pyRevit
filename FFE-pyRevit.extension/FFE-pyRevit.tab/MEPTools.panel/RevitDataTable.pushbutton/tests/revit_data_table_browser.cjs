// Offline integration test using installed Edge and its DevTools protocol; no npm dependencies.
// Run: node tests/revit_data_table_browser.cjs
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const { spawn } = require('node:child_process');
const { randomUUID } = require('node:crypto');

const edge = process.env.EDGE_PATH || 'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe';
const support = path.resolve(__dirname, '../support');
const pause = ms => new Promise(resolve => setTimeout(resolve, ms));

async function main() {
  const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'ffe-data-table-'));
  const browser = spawn(edge, ['--headless=new', '--disable-gpu', '--no-first-run',
    '--no-default-browser-check', '--disable-extensions', '--remote-debugging-port=0',
    `--user-data-dir=${profile}`, 'about:blank'], { windowsHide: true, stdio: 'ignore' });
  let socket;
  let command;
  browser.on('error', error => { console.error(error); process.exitCode = 1; });
  try {
    const activePort = path.join(profile, 'DevToolsActivePort');
    for (let attempt = 0; !fs.existsSync(activePort) && attempt < 150; attempt++) await pause(100);
    assert.ok(fs.existsSync(activePort), 'Edge must expose a DevTools endpoint');
    const port = fs.readFileSync(activePort, 'utf8').split('\n')[0];
    const pages = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json();
    socket = new WebSocket(pages.find(page => page.type === 'page').webSocketDebuggerUrl);
    await new Promise((resolve, reject) => { socket.onopen = resolve; socket.onerror = reject; });
    let serial = 0;
    const pending = new Map();
    const errors = [];
    let interceptedDrag = null;
    socket.onmessage = event => {
      const message = JSON.parse(event.data);
      if (message.method === 'Input.dragIntercepted') interceptedDrag = message.params.data;
      if (message.method === 'Runtime.exceptionThrown') errors.push(message.params.exceptionDetails.text);
      if (message.id && pending.has(message.id)) {
        const { resolve, reject, timer } = pending.get(message.id);
        clearTimeout(timer);
        pending.delete(message.id);
        if (message.error) reject(new Error(message.error.message));
        else resolve(message.result);
      }
    };
    command = (method, params = {}) => new Promise((resolve, reject) => {
      const id = ++serial;
      const timer = setTimeout(() => { pending.delete(id); reject(new Error(`Timed out: ${method}`)); }, 15000);
      pending.set(id, { resolve, reject, timer });
      socket.send(JSON.stringify({ id, method, params }));
    });
    const evaluate = async expression => {
      const result = await command('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true })
        .catch(error => { throw new Error(`${error.message}: ${expression.slice(0, 140)}`); });
      if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
      return result.result.value;
    };
    const waitFor = async expression => {
      for (let attempt = 0; attempt < 150; attempt++) {
        if (await evaluate(expression)) return;
        await pause(50);
      }
      throw new Error(`UI did not settle: ${expression}\n${await evaluate('document.getElementById("status").textContent')}`);
    };
    const grid = 'Tabulator.findTable("#grid")[0]';
    const groupBy = async field => {
      await evaluate('document.getElementById("group-clear").click()');
      if (field) await dragColumn(field, '#group-drop-zone');
    };
    const grouped = 'Array.from(document.querySelectorAll(".group-chip"), chip => chip.dataset.field)';
    const sortPriorities = 'Array.from(document.querySelectorAll("#grid .sort-priority"), number => ({field: number.closest(".tabulator-col").getAttribute("tabulator-field"), priority: number.textContent}))';
    const bands = `${grid}.getColumns(true).filter(column => column.getSubColumns().length).map(column => ({name: column.getDefinition().title, fields: column.getSubColumns().map(child => child.getField())}))`;
    const configureHeaders = async (groups, apply = true) => {
      await evaluate(`document.getElementById('column-groups').click();
        while (document.querySelector('.header-group-remove')) document.querySelector('.header-group-remove').click();
        for (const group of ${JSON.stringify(groups)}) {
          document.getElementById('header-group-add').click();
          const index = document.querySelectorAll('#header-group-list fieldset').length - 1;
          const input = document.querySelectorAll('#header-group-list input[type=text]')[index];
          input.value = group.name; input.dispatchEvent(new Event('input', {bubbles:true}));
          for (const field of group.fields) Array.from(document.querySelectorAll('#header-group-list fieldset')[index].querySelectorAll('input[type=checkbox]'))
            .find(checkbox => checkbox.dataset.field === field).click();
        }
        document.getElementById(${JSON.stringify(apply ? 'header-groups-apply' : 'header-groups-cancel')}).click();`);
      await waitFor('!document.getElementById("header-groups-dialog").open');
    };
    const dragColumn = async (field, target, after = true, cancel = false, sourceKind = 'header') => {
      await command('Page.bringToFront');
      await evaluate(`window.__dragTrace = []; if (!window.__dragTraceInstalled) {
        window.__dragTraceInstalled = true;
        for (const type of ['mousedown','mousemove','dragstart','dragend']) window.addEventListener(type, event =>
          window.__dragTrace.push({type, prevented:event.defaultPrevented, tag:event.target.tagName, cls:event.target.className}));
      }`);
      if (sourceKind === 'header' || sourceKind === 'band') await evaluate(`${grid}.getColumn(${JSON.stringify(field)}).scrollTo('center', true)`);
      if (sourceKind === 'band' && target.startsWith('band:')) {
        await evaluate(`${grid}.getColumn(${JSON.stringify(target.slice(5))}).getParentColumn().getSubColumns().filter(column => column.isVisible()).at(${after ? -1 : 0}).scrollTo('center', true)`);
      }
      const source = await evaluate(`(() => {
        const element = ${sourceKind === 'chip' ? `Array.from(document.querySelectorAll('.group-chip')).find(chip => chip.dataset.field === ${JSON.stringify(field)})` : sourceKind === 'band' ? `${grid}.getColumn(${JSON.stringify(field)}).getParentColumn().getElement().querySelector(':scope > .tabulator-col-content')` : `${grid}.getColumn(${JSON.stringify(field)}).getElement()`};
        const bounds = element.getBoundingClientRect();
        const grid = document.getElementById('grid').getBoundingClientRect();
        return {x: Math.min(bounds.right - 15, Math.max(bounds.left + 35, grid.left + 20)), y: bounds.top + bounds.height / 2};
      })()`);
      const destination = await evaluate(`(() => {
        const element = ${target === '#group-drop-zone' ? 'document.getElementById("group-drop-zone")' : target.startsWith('band:') ? `${grid}.getColumn(${JSON.stringify(target.slice(5))}).getParentColumn().getElement().querySelector(':scope > .tabulator-col-content')` : target.startsWith('chip:') ? `Array.from(document.querySelectorAll('.group-chip')).find(chip => chip.dataset.field === ${JSON.stringify(target.slice(5))})` : `${grid}.getColumn(${JSON.stringify(target)}).getElement()`};
        const bounds = element.getBoundingClientRect();
        return {x: bounds.left + bounds.width * ${after ? '0.75' : '0.25'}, y: bounds.top + bounds.height / 2};
      })()`);
      interceptedDrag = null;
      await command('Input.setInterceptDrags', { enabled: true });
      try {
        await command('Input.dispatchMouseEvent', { type: 'mouseMoved', ...source });
        await command('Input.dispatchMouseEvent', { type: 'mousePressed', ...source, button: 'left', buttons: 1, clickCount: 1 });
        await pause(60);
        await command('Input.dispatchMouseEvent', { type: 'mouseMoved', x: source.x + 12, y: source.y, button: 'left', buttons: 1 });
        await pause(20);
        await command('Input.dispatchMouseEvent', { type: 'mouseMoved', x: source.x + 20, y: source.y - 5, button: 'left', buttons: 1 });
        for (let attempt = 0; !interceptedDrag && attempt < 75; attempt++) await pause(20);
        assert.ok(interceptedDrag, `native header drag must start for ${field}: ${JSON.stringify(await evaluate('window.__dragTrace'))}; ${JSON.stringify(source)}; ${await evaluate(`${grid}.getColumn(${JSON.stringify(field)}).getElement().outerHTML.slice(0, 450)`)}`);
        assert.ok(interceptedDrag.items.some(item => item.mimeType === 'application/x-ffe-table-column'));
        if (!cancel) {
          await command('Input.dispatchDragEvent', { type: 'dragEnter', ...destination, data: interceptedDrag });
          await command('Input.dispatchDragEvent', { type: 'dragOver', ...destination, data: interceptedDrag });
          if (target === '#group-drop-zone') {
            assert.equal(await evaluate('document.getElementById("group-drop-zone").classList.contains("group-drop-over")'), sourceKind !== 'band');
          }
          await command('Input.dispatchDragEvent', { type: 'drop', ...destination, data: interceptedDrag });
        } else {
          await command('Input.dispatchDragEvent', { type: 'dragCancel', ...destination, data: interceptedDrag });
        }
      } finally {
        await command('Input.cancelDragging').catch(() => {});
        await command('Input.dispatchMouseEvent', { type: 'mouseReleased', ...destination, button: 'left', buttons: 0, clickCount: 1 });
        await command('Input.setInterceptDrags', { enabled: false });
      }
      await waitFor('!document.querySelector(".column-drag-source")');
    };
    await command('Runtime.enable');
    await command('Page.enable');
    await command('Emulation.setDeviceMetricsOverride', { width: 1160, height: 740, deviceScaleFactor: 1, mobile: false });
    await command('Page.addScriptToEvaluateOnNewDocument', { source: `
      window.__requests = [];
      window.chrome = window.chrome || {};
      window.chrome.webview = {
        postMessage(message) { window.__requests.push(message); },
        addEventListener(type, callback) { window.__hostMessage = data => callback({ data }); }
      };` });
    await command('Page.navigate', { url: pathToFileURL(path.join(support, 'index.html')).href });
    await waitFor('window.__requests.some(message => message.action === "ready")');
    const makeRow = (id, number, phaseId, levelId, elevation, area, name = 'Office') => ({
      id, uniqueId: 'uid-' + id, number, name, phaseId, phase: phaseId === '1' ? 'Existing' : 'New',
      levelId, level: levelId === '10' ? 'Ground' : 'Basement', levelElevation: elevation,
      areaValue: area, areaText: `${area} ft²`, heightValue: 9, heightText: `9' 0"`, warnings: [],
      parameterValues: { 'p-200': area * 10, p201: area === 0 ? null : `Comment ${number}`, p204: '<img src=x onerror="window.injected=true">' },
      parameterTexts: { 'p-200': `${area * 10} CFM`, p201: area === 0 ? '—' : `Comment ${number}`, p204: '<img src=x onerror="window.injected=true">' },
    });
    const snapshot = {
      documentKey: 'model-session', documentTitle: 'Mechanical Model', defaultPhaseId: '2',
      refreshedAt: new Date().toISOString(), skippedCount: 0, issues: [],
      savedTables: [], tableIssues: [], canSaveTables: true,
      phases: [{ id: '1', name: 'Existing' }, { id: '2', name: 'New' }],
      levels: [{ id: '10', name: 'Ground', elevation: 0 }, { id: '11', name: 'Basement', elevation: -10 }],
      parameters: [
        { id: '-200', key: 'p-200', name: 'Supply Airflow', dataType: 'Air Flow', storageType: 'Double' },
        { id: '201', key: 'p201', name: 'Comments', dataType: 'Text', storageType: 'String' },
        { id: '202', key: 'p202', name: 'Conditioned', dataType: 'Yes/No', storageType: 'Integer' },
        { id: '203', key: 'p203', name: 'Comments', dataType: 'Text', storageType: 'String' },
        { id: '204', key: 'p204', name: '<img src=x onerror="window.injected=true">', dataType: 'Text', storageType: 'String' },
        { id: '-1006916', key: 'p-1006916', name: 'Level', dataType: 'Element reference', storageType: 'ElementId', baseField: 'level' },
        { id: '-1006901', key: 'p-1006901', name: 'Number', dataType: 'Text', storageType: 'String', baseField: 'number' },
        { id: '-1006900', key: 'p-1006900', name: 'Name', dataType: 'Text', storageType: 'String', baseField: 'name' },
        { id: '-1006902', key: 'p-1006902', name: 'Area', dataType: 'Area', storageType: 'Double', baseField: 'areaValue' },
        { id: '-1006920', key: 'p-1006920', name: 'Unbounded Height', dataType: 'Length', storageType: 'Double', baseField: 'heightValue' },
        { id: '-1012112', key: 'p-1012112', name: 'Phase', dataType: 'Element reference', storageType: 'ElementId', baseField: 'phase' },
      ],
      selectedParameterIds: [],
      rows: [makeRow('9007199254740993', '10', '2', '10', 0, 100),
        makeRow('9007199254740994', '002', '2', '10', 0, 20, '<img src=x onerror="window.injected=true">'),
        makeRow('3', '1', '2', '11', -10, 0), makeRow('4', '1', '1', '10', 0, 50)],
    };
    const deliver = message => evaluate(`window.__hostMessage(${JSON.stringify(message)})`);
    const records = new Map();
    let lastTableRequest = null;
    const copy = value => JSON.parse(JSON.stringify(value));
    const metadata = record => Object.fromEntries(Object.entries(record).filter(([key]) => key !== 'layout'));
    const catalog = () => {
      snapshot.savedTables = Array.from(records.values(), metadata);
      return { savedTables: snapshot.savedTables, tableIssues: snapshot.tableIssues, canSaveTables: snapshot.canSaveTables };
    };
    const respondToTable = async (failure = null, transform = null) => {
      await waitFor(`['loadTable','saveTable','renameTable','deleteTable'].includes(window.__requests.at(-1).action) && window.__requests.at(-1).requestId !== ${JSON.stringify(lastTableRequest)}`);
      const request = await evaluate('window.__requests.at(-1)');
      lastTableRequest = request.requestId;
      const response = { type: 'tableResult', requestId: request.requestId, action: request.action, success: !failure };
      if (failure) {
        await deliver({ ...response, message: failure });
        await waitFor('!document.getElementById("refresh").disabled');
        return request;
      }
      let record = request.tableId ? records.get(request.tableId) : null;
      if (request.action === 'loadTable') {
        assert.ok(!request.tableId || record, 'requested saved table must exist');
        snapshot.selectedParameterIds = record ? record.layout.columns.filter(column => column.field.startsWith('parameterValues.'))
          .map(column => snapshot.parameters.find(parameter => `parameterValues.${parameter.key}` === column.field)?.id).filter(Boolean) : [];
        catalog();
        Object.assign(response, { table: record ? metadata(record) : null, snapshot, layout: record ? copy(record.layout) : null, warnings: [] });
      } else {
        assert.ok(!record || request.revision === record.revision, 'writes must carry the loaded revision');
        if (request.action === 'deleteTable') {
          records.delete(record.id);
          record = null;
        } else if (request.action === 'renameTable') {
          record = { ...record, name: request.name, revision: randomUUID() };
          records.set(record.id, record);
        } else {
          assert.ok(request.layout.columns.every(column => Number.isFinite(column.width) && column.width >= 60 && column.width <= 3000));
          record = { id: record?.id || randomUUID(), revision: randomUUID(), name: request.name,
            modifiedBy: 'Tester', modifiedAt: new Date().toISOString(), layout: copy(request.layout) };
          records.set(record.id, record);
        }
        Object.assign(response, { table: record ? metadata(record) : null, catalog: catalog(), layout: request.layout || null });
      }
      if (transform) transform(response);
      await deliver(response);
      await waitFor(`!document.getElementById("refresh").disabled || (window.__requests.at(-1).requestId && window.__requests.at(-1).requestId !== ${JSON.stringify(request.requestId)})`);
      return { request, record };
    };
    const chooseTable = id => evaluate(`document.getElementById("saved-table").value=${JSON.stringify(id)}; document.getElementById("saved-table").dispatchEvent(new Event("change"))`);
    const nameDialog = async name => {
      await waitFor('document.getElementById("table-dialog").open && !document.getElementById("table-name").disabled');
      await evaluate(`document.getElementById("table-name").value=${JSON.stringify(name)}; document.getElementById("table-dialog-confirm").click()`);
    };
    for (const parameter of snapshot.parameters) {
      parameter.parameterType = parameter.id.startsWith('-') ? 'Built-In'
        : parameter.id === '201' ? 'Shared Parameter' : 'Project Parameter';
    }
    await deliver({ type: 'snapshot', snapshot });
    await waitFor('document.getElementById("counts").textContent.startsWith("3 of 4")');
    assert.deepEqual(await evaluate(`${grid}.getRows("active").map(row => row.getData().id)`),
      ['3', '9007199254740994', '9007199254740993'], 'level elevation then natural number ordering');
    assert.equal(await evaluate('document.querySelectorAll("#grid img").length'), 0, 'model text cannot inject markup');
    assert.equal(await evaluate(`${grid}.getColumn("phase").isVisible()`), false);
    assert.equal(await evaluate('Boolean(document.getElementById("group-by") || document.getElementById("document-title"))'), false);
    const originalOrder = await evaluate(`${grid}.getColumns().map(column => column.getField())`);
    const originalSort = await evaluate(`${grid}.getSorters().map(sorter => ({field:sorter.field, dir:sorter.dir}))`);
    assert.deepEqual(await evaluate(sortPriorities), [{field:'level',priority:'1'},{field:'number',priority:'2'}], 'multi-sort badges reflect actual comparator priority');
    await dragColumn('level', '#group-drop-zone');
    assert.equal(await evaluate('document.querySelector(".group-chip")?.dataset.field || ""'), 'level');
    assert.equal(await evaluate('document.querySelector(".group-chip-name").textContent'), 'Level');
    assert.equal(await evaluate(`${grid}.getGroups().length`), 2);
    await dragColumn('name', '#group-drop-zone');
    await dragColumn('number', '#group-drop-zone');
    assert.deepEqual(await evaluate(grouped), ['level', 'name', 'number'], 'header drops append nested grouping levels');
    assert.equal(await evaluate('document.querySelectorAll(".group-arrow").length'), 2, 'arrows connect successive hierarchy levels');
    assert.equal(await evaluate(`${grid}.getGroups().every(group => group.getSubGroups().every(child => child.getSubGroups().length > 0))`), true, 'three nested levels reach the leaf groups');
    assert.equal(await evaluate(`${grid}.getGroups()[0].getElement().querySelector('.group-toggle').textContent.startsWith('Level:')`), true);
    assert.equal(await evaluate(`${grid}.getGroups()[0].getSubGroups()[0].getElement().querySelector('.group-toggle').textContent.startsWith('Name:')`), true);
    assert.equal(await evaluate(`${grid}.getGroups()[0].getSubGroups()[0].getSubGroups()[0].getElement().querySelector('.group-toggle').textContent.startsWith('Number:')`), true);
    const hierarchyScreenshot = await command('Page.captureScreenshot', { format: 'png' });
    fs.writeFileSync(path.join(profile, 'grouping-hierarchy.png'), Buffer.from(hierarchyScreenshot.data, 'base64'));
    await dragColumn('level', '#group-drop-zone');
    assert.deepEqual(await evaluate(grouped), ['level', 'name', 'number'], 'repeated header drops never duplicate groups');
    await dragColumn('number', 'chip:level', false, false, 'chip');
    assert.deepEqual(await evaluate(grouped), ['number', 'level', 'name'], 'chip drops reorder the hierarchy');
    assert.deepEqual(await evaluate(`${grid}.getColumns().map(column => column.getField())`), originalOrder, 'chip reordering leaves columns in place');
    await dragColumn('number', 'chip:name', true, true, 'chip');
    assert.deepEqual(await evaluate(grouped), ['number', 'level', 'name'], 'cancelled chip drags preserve hierarchy order');
    await evaluate('document.querySelector(".group-chip").dispatchEvent(new KeyboardEvent("keydown", {key:"ArrowRight", altKey:true, bubbles:true}))');
    assert.deepEqual(await evaluate(grouped), ['level', 'number', 'name'], 'keyboard reordering updates hierarchy');
    await evaluate('document.querySelectorAll(".group-remove")[1].click()');
    assert.deepEqual(await evaluate(grouped), ['level', 'name'], 'removing an inner group preserves the remaining order');
    // Collapsed ancestors must hide leaf rows from keyboard focus/navigation.
    await evaluate(`${grid}.getGroups()[0].hide(); document.getElementById('grid').focus(); document.getElementById('grid').dispatchEvent(new KeyboardEvent('keydown', {key:'Home', bubbles:true}))`);
    assert.equal(await evaluate(`(() => {
      const id = document.getElementById('grid').getAttribute('aria-activedescendant').replace('space-row-', '');
      for (let group = ${grid}.getRow(id).getGroup(); group; group = group.getParentGroup()) if (!group.isVisible()) return false;
      return true;
    })()`), true, 'keyboard navigation skips descendants of collapsed groups');
    await evaluate(`${grid}.getGroups()[0].show(); document.querySelectorAll('.group-remove')[1].click()`);
    assert.deepEqual(await evaluate(`${grid}.getColumns().map(column => column.getField())`), originalOrder, 'grouping a header never moves it');
    assert.deepEqual(await evaluate(`${grid}.getSorters().map(sorter => ({field:sorter.field, dir:sorter.dir}))`), originalSort, 'dragging does not sort the column');
    await dragColumn('number', '#group-drop-zone', true, true);
    assert.equal(await evaluate('document.querySelector(".group-chip")?.dataset.field || ""'), 'level', 'cancelled drags preserve grouping');
    assert.deepEqual(await evaluate(`${grid}.getColumns().map(column => column.getField())`), originalOrder);
    const headerGroups = [{name:'Identity', fields:['level','name']}, {name:'Dimensions', fields:['areaValue','heightValue']}];
    await evaluate(`${grid}.selectRow('9007199254740993')`);
    await configureHeaders(headerGroups);
    assert.deepEqual(await evaluate(bands), headerGroups, 'named native bands contain the chosen leaf columns');
    assert.deepEqual(await evaluate(sortPriorities), [{field:'level',priority:'1'},{field:'number',priority:'2'}], 'rebuilding grouped headers retains priority badges on leaf columns');
    assert.deepEqual(await evaluate(`${grid}.getColumns().map(column => column.getField())`),
      ['level','name','number','areaValue','heightValue','phase'], 'nonadjacent members move together');
    assert.deepEqual(await evaluate(`${grid}.getSelectedData().map(row => row.id)`), ['9007199254740993'], 'applying header bands retains row selection');
    assert.deepEqual(await evaluate(`${grid}.getSorters().map(sorter => ({field:sorter.field,dir:sorter.dir}))`), originalSort, 'applying bands retains sorting');
    const originalBandWidths = await evaluate(`${grid}.getColumns().map(column => ({field:column.getField(),width:column.getWidth()}))`);
    await dragColumn('areaValue', 'band:level', false, false, 'band');
    assert.deepEqual(await evaluate(bands), [headerGroups[1], headerGroups[0]], 'dragging a band title moves the entire group before another band');
    assert.deepEqual(await evaluate(`${grid}.getColumns().map(column => column.getField())`),
      ['areaValue','heightValue','level','name','number','phase'], 'whole-band moves retain contiguous member order');
    assert.deepEqual(await evaluate(`${grid}.getSelectedData().map(row => row.id)`), ['9007199254740993']);
    assert.deepEqual(await evaluate(`${grid}.getSorters().map(sorter => ({field:sorter.field,dir:sorter.dir}))`), originalSort);
    assert.deepEqual(await evaluate(`${grid}.getColumns().map(column => ({field:column.getField(),width:column.getWidth()})).sort((a,b) => a.field.localeCompare(b.field))`),
      [...originalBandWidths].sort((a,b) => a.field.localeCompare(b.field)), 'band dragging preserves every leaf width');
    await dragColumn('areaValue', 'band:level', true, true, 'band');
    assert.deepEqual(await evaluate(bands), [headerGroups[1], headerGroups[0]], 'cancelled band drags preserve order');
    await dragColumn('areaValue', '#group-drop-zone', true, false, 'band');
    assert.deepEqual(await evaluate(grouped), ['level'], 'whole bands cannot become row grouping fields');
    assert.deepEqual(await evaluate(bands), [headerGroups[1], headerGroups[0]]);
    await dragColumn('areaValue', 'number', true, false, 'band');
    assert.deepEqual(await evaluate(`${grid}.getColumns().map(column => column.getField())`),
      ['level','name','number','areaValue','heightValue','phase'], 'bands can move beside ungrouped columns');
    await dragColumn('areaValue', 'level', false, false, 'band');
    assert.deepEqual(await evaluate(`${grid}.getColumns().map(column => column.getField())`),
      ['areaValue','heightValue','level','name','number','phase'], 'dropping on a grouped leaf moves beside its entire band');
    await dragColumn('areaValue', 'band:level', true, false, 'band');
    assert.deepEqual(await evaluate(`${grid}.getColumns().map(column => column.getField())`),
      ['level','name','areaValue','heightValue','number','phase'], 'after drops place whole groups after the target band');
    await dragColumn('areaValue', 'number', true, false, 'band');
    assert.deepEqual(await evaluate(bands), headerGroups, 'band dragging never transfers leaf membership');
    await evaluate('document.getElementById("column-groups").click()');
    assert.equal(await evaluate(`Array.from(document.querySelectorAll('#header-group-list fieldset')[1].querySelectorAll('input[type=checkbox]')).find(input => input.dataset.field === 'level').disabled`), true, 'a column cannot be assigned to two header bands');
    await evaluate('document.getElementById("header-groups-cancel").click()');
    await configureHeaders([{name:'Discarded', fields:['number']}], false);
    assert.deepEqual(await evaluate(bands), headerGroups, 'Cancel discards the header group draft');
    await groupBy('level');
    await dragColumn('name', 'level', false);
    assert.deepEqual(await evaluate(bands), [{name:'Identity',fields:['name','level']}, headerGroups[1]], 'leaf headers reorder inside their band');
    await dragColumn('number', 'heightValue');
    assert.deepEqual(await evaluate(bands), [{name:'Identity',fields:['name','level']}, {name:'Dimensions',fields:['areaValue','heightValue','number']}], 'cross-band header drops update membership and leaf order together');
    assert.deepEqual(await evaluate(grouped), ['level'], 'header bands and row grouping are independent');
    await deliver({type:'snapshot',snapshot});
    await waitFor('!document.getElementById("column-groups").disabled');
    assert.deepEqual((await evaluate(bands))[1].fields, ['areaValue','heightValue','number'], 'Refresh retains bands');
    await configureHeaders([{name:'Context',fields:['level','phase']}, headerGroups[1]]);
    await dragColumn('level', 'band:areaValue', true, false, 'band');
    assert.deepEqual((await evaluate(bands)).find(group => group.name === 'Context').fields, ['level','phase'], 'hidden members move with their visible band');
    assert.equal(await evaluate(`${grid}.getColumn('phase').isVisible()`), false, 'moving a band does not reveal hidden members');
    const ownBandOrder = await evaluate(`${grid}.getColumns().map(column => column.getField())`);
    await dragColumn('level', 'level', true, false, 'band');
    assert.deepEqual(await evaluate(`${grid}.getColumns().map(column => column.getField())`), ownBandOrder, 'dropping a band on itself is a no-op');
    const unsafeName = '<img src=x onerror="window.injected=true">';
    await configureHeaders([{name:unsafeName,fields:['areaValue','heightValue']}]);
    assert.equal(await evaluate('document.querySelectorAll("#grid img").length'), 0, 'custom header names stay literal text');
    assert.equal((await evaluate(bands))[0].name, unsafeName);
    await configureHeaders([]);
    assert.deepEqual(await evaluate(bands), [], 'removing header groups leaves the leaf columns');
    await groupBy('');
    await evaluate(`${grid}.deselectRow(); for (let i=1; i<${JSON.stringify(originalOrder)}.length; i++) ${grid}.moveColumn(${JSON.stringify(originalOrder)}[i], ${JSON.stringify(originalOrder)}[i-1], true)`);
    await evaluate('document.getElementById("group-clear").click()');
    assert.equal(await evaluate(`${grid}.getGroups().length`), 0);
    assert.equal(await evaluate('!document.querySelector(".group-chip")'), true);
    await dragColumn('number', 'name');
    assert.deepEqual(await evaluate(`${grid}.getColumns().map(column => column.getField())`),
      ['level', 'name', 'number', 'areaValue', 'heightValue', 'phase'], 'native header drops still reorder columns');
    await dragColumn('number', 'level');
    assert.deepEqual(await evaluate(`${grid}.getColumns().map(column => column.getField())`), originalOrder);
    const clickPoint = await evaluate(`(() => {
      const bounds = ${grid}.getColumn('name').getElement().getBoundingClientRect();
      return {x: bounds.left + 35, y: bounds.top + bounds.height / 2};
    })()`);
    await command('Input.dispatchMouseEvent', { type: 'mousePressed', ...clickPoint, button: 'left', buttons: 1, clickCount: 1 });
    await command('Input.dispatchMouseEvent', { type: 'mouseReleased', ...clickPoint, button: 'left', buttons: 0, clickCount: 1 });
    assert.equal(await evaluate(`${grid}.getSorters()[0].field`), 'name', 'clicking draggable headers still sorts');
    assert.deepEqual(await evaluate(sortPriorities), [], 'single-column sorting has no priority badge');
    const primaryNameOrder = await evaluate(`${grid}.getRows('active').map(row => row.getData().name)`);
    const areaSortPoint = await evaluate(`(() => {
      const bounds = ${grid}.getColumn('areaValue').getElement().getBoundingClientRect();
      return {x: bounds.left + 35, y: bounds.top + bounds.height / 2};
    })()`);
    const ctrlClickArea = async () => {
      await command('Input.dispatchMouseEvent', {type:'mousePressed', ...areaSortPoint, button:'left', buttons:1, clickCount:1, modifiers:2});
      await command('Input.dispatchMouseEvent', {type:'mouseReleased', ...areaSortPoint, button:'left', buttons:0, clickCount:1, modifiers:2});
    };
    await ctrlClickArea();
    assert.deepEqual(await evaluate(sortPriorities), [{field:'name',priority:'1'},{field:'areaValue',priority:'2'}], 'Ctrl+click updates sorting priorities immediately');
    assert.deepEqual(await evaluate(`${grid}.getRows('active').map(row => row.getData().name)`), primaryNameOrder, 'the first selection remains the primary sort');
    assert.deepEqual(await evaluate(`${grid}.getRows('active').filter(row => row.getData().name === 'Office').map(row => row.getData().areaValue)`), [0,100], 'the second selection breaks ties in the first');
    await ctrlClickArea();
    assert.deepEqual(await evaluate(sortPriorities), [{field:'name',priority:'1'},{field:'areaValue',priority:'2'}], 'direction changes retain priority');
    assert.deepEqual(await evaluate(`${grid}.getRows('active').map(row => row.getData().name)`), primaryNameOrder);
    assert.deepEqual(await evaluate(`${grid}.getRows('active').filter(row => row.getData().name === 'Office').map(row => row.getData().areaValue)`), [100,0]);
    const shiftClickNumber = async () => {
      const point = await evaluate(`(() => {
        const bounds = ${grid}.getColumn('number').getElement().getBoundingClientRect();
        return {x: bounds.left + 35, y: bounds.top + bounds.height / 2};
      })()`);
      await command('Input.dispatchMouseEvent', {type:'mousePressed', ...point, button:'left', buttons:1, clickCount:1, modifiers:8});
      await command('Input.dispatchMouseEvent', {type:'mouseReleased', ...point, button:'left', buttons:0, clickCount:1, modifiers:8});
    };
    await shiftClickNumber();
    assert.deepEqual(await evaluate(`${grid}.getSorters().map(sorter => sorter.field)`), ['name','areaValue','number'], 'Shift+click appends the third selection');
    assert.deepEqual(await evaluate(sortPriorities), [{field:'number',priority:'3'},{field:'name',priority:'1'},{field:'areaValue',priority:'2'}]);
    const ctrlClickName = async () => {
      await command('Input.dispatchMouseEvent', {type:'mousePressed', ...clickPoint, button:'left', buttons:1, clickCount:1, modifiers:2});
      await command('Input.dispatchMouseEvent', {type:'mouseReleased', ...clickPoint, button:'left', buttons:0, clickCount:1, modifiers:2});
    };
    await ctrlClickName();
    assert.deepEqual(await evaluate(`${grid}.getSorters().map(sorter => ({field:sorter.field,dir:sorter.dir}))`),
      [{field:'name',dir:'desc'},{field:'areaValue',dir:'desc'},{field:'number',dir:'asc'}], 'changing the first sort direction preserves its original priority');
    assert.deepEqual(await evaluate(sortPriorities), [{field:'number',priority:'3'},{field:'name',priority:'1'},{field:'areaValue',priority:'2'}]);
    await ctrlClickName();
    assert.deepEqual(await evaluate(sortPriorities), [{field:'number',priority:'2'},{field:'areaValue',priority:'1'}], 'removing the primary sort renumbers the remaining selections');
    await ctrlClickArea();
    assert.deepEqual(await evaluate(sortPriorities), [], 'removing the secondary sort clears all badges for the single remaining column');
    await ctrlClickName();
    assert.deepEqual(await evaluate(`${grid}.getSorters().map(sorter => sorter.field)`), ['number','name'], 'a removed column is appended when selected again');
    assert.deepEqual(await evaluate(sortPriorities), [{field:'number',priority:'1'},{field:'name',priority:'2'}]);
    await evaluate(`${grid}.clearSort()`);
    assert.deepEqual(await evaluate(sortPriorities), [], 'cleared sorting leaves no badges');
    await evaluate(`${grid}.setSort(${JSON.stringify(originalSort.map(sorter => ({ column: sorter.field, dir: sorter.dir })))})`);
    assert.deepEqual(await evaluate(sortPriorities), [{field:'level',priority:'1'},{field:'number',priority:'2'}]);
    const widthBeforeResize = await evaluate(`${grid}.getColumn('number').getWidth()`);
    const resizePoint = await evaluate(`(() => {
      const handle = ${grid}.getColumn('number').getElement().nextElementSibling;
      if (!handle.classList.contains('tabulator-col-resize-handle')) throw new Error('Expected resize handle');
      const bounds = handle.getBoundingClientRect();
      return {x: bounds.left + bounds.width / 2, y: bounds.top + bounds.height / 2};
    })()`);
    await command('Input.dispatchMouseEvent', { type: 'mouseMoved', ...resizePoint });
    await command('Input.dispatchMouseEvent', { type: 'mousePressed', ...resizePoint, button: 'left', buttons: 1, clickCount: 1 });
    await command('Input.dispatchMouseEvent', { type: 'mouseMoved', x: resizePoint.x + 40, y: resizePoint.y, button: 'left', buttons: 1 });
    await command('Input.dispatchMouseEvent', { type: 'mouseReleased', x: resizePoint.x + 40, y: resizePoint.y, button: 'left', buttons: 0, clickCount: 1 });
    assert.equal(await evaluate(`${grid}.getColumn('number').getWidth()`), widthBeforeResize + 40, 'edge dragging still resizes');
    assert.equal(await evaluate('Boolean(document.querySelector(".column-drag-source"))'), false, 'resizing does not start header dragging');
    assert.equal(await evaluate('document.querySelector(".group-chip")?.dataset.field || ""'), '', 'resizing does not group');
    await evaluate(`${grid}.getColumn('number').setWidth(${widthBeforeResize})`);
    for (const field of ['level', 'number', 'name', 'areaValue', 'heightValue']) {
      await groupBy(field);
      assert.equal(await evaluate(`${grid}.getGroups().reduce((count, group) => count + group.getRows().length, 0)`), 3,
        `all Spaces are present when grouping by ${field}`);
    }
    await groupBy('name');
    assert.equal(await evaluate('document.querySelectorAll("#grid img").length'), 0, 'group labels cannot inject HTML');
    await groupBy('level');
    await evaluate(`${grid}.getGroups()[0].show(); ${grid}.getGroups()[0].getElement().querySelector('.group-toggle').click()`);
    await waitFor(`!${grid}.getGroups()[0].isVisible()`);
    assert.equal(await evaluate(`${grid}.getGroups()[0].isVisible()`), false, 'group header collapses its rows');
    assert.equal(await evaluate(`${grid}.getGroups()[0].getElement().querySelector('.group-toggle').getAttribute('aria-expanded')`), 'false');
    await evaluate(`${grid}.getGroups()[0].getElement().querySelector('.group-toggle').click()`);
    await waitFor(`${grid}.getGroups()[0].isVisible()`);
    assert.equal(await evaluate(`${grid}.getGroups()[0].isVisible()`), true);
    await evaluate(`${grid}.setSort("areaValue", "asc")`);
    assert.deepEqual(await evaluate(`${grid}.getRows("active").map(row => row.getData().areaValue)`), [0, 20, 100]);
    await evaluate(`${grid}.selectRow("9007199254740994"); document.getElementById("select-revit").click()`);
    assert.deepEqual(await evaluate('window.__requests.at(-1).elementIds'), ['9007199254740994']);
    await deliver({ type: 'result', kind: 'success', message: 'Selected 1 Space.' });
    await deliver({ type: 'context', active: true, outdated: true });
    await waitFor('document.getElementById("context").textContent.includes("Out of date")');
    assert.equal(await evaluate(`${grid}.getDataCount()`), 4, 'stale notification leaves displayed data intact');
    await evaluate('document.getElementById("refresh").click()');
    snapshot.rows[0].areaValue = 200;
    await deliver({ type: 'snapshot', snapshot });
    await deliver({ type: 'context', active: true, outdated: false });
    await deliver({ type: 'result', kind: 'success', message: 'Refreshed.' });
    await waitFor('!document.getElementById("refresh").disabled');
    assert.deepEqual(await evaluate(`${grid}.getSelectedData().map(row => row.id)`), ['9007199254740994']);
    assert.equal(await evaluate(`${grid}.getSorters()[0].field`), 'areaValue', 'refresh retains sorting');
    assert.equal(await evaluate('document.querySelector(".group-chip")?.dataset.field || ""'), 'level', 'refresh retains grouping');
    assert.equal(await evaluate(`${grid}.getGroups().length`), 2);
    await groupBy('');
    assert.equal(await evaluate(`${grid}.getGroups().length`), 0, 'None restores the flat table');
    await deliver({ type: 'context', active: false, outdated: false });
    await waitFor('document.getElementById("select-revit").disabled');
    assert.equal(await evaluate('document.getElementById("refresh").disabled'), true);
    await deliver({ type: 'context', active: true, outdated: false });
    await waitFor('!document.getElementById("refresh").disabled');
    await evaluate('document.getElementById("phase").value = "all"; document.getElementById("phase").dispatchEvent(new Event("change"))');
    assert.equal(await evaluate(`${grid}.getColumn("phase").isVisible()`), true);
    assert.equal(await evaluate(`${grid}.getDataCount("active")`), 4);
    await groupBy('phase');
    assert.equal(await evaluate(`${grid}.getGroups().length`), 2, 'visible Phase is available for grouping');
    await evaluate('document.getElementById("search").value = "nonexistent"; document.getElementById("search").dispatchEvent(new Event("input"))');
    assert.equal(await evaluate(`${grid}.getDataCount("active")`), 0);
    assert.equal(await evaluate(`${grid}.getSelectedData().length`), 0, 'hidden rows cannot remain selected for navigation');
    await evaluate('document.getElementById("clear-filters").click(); document.getElementById("grid").focus(); document.getElementById("grid").dispatchEvent(new KeyboardEvent("keydown", {key:"ArrowDown"})); document.getElementById("grid").dispatchEvent(new KeyboardEvent("keydown", {key:" "}))');
    assert.equal(await evaluate(`${grid}.getSelectedData().length`), 1, 'keyboard row selection');
    assert.equal(await evaluate('document.querySelector(".group-chip")?.dataset.field || ""'), '', 'hiding a grouped Phase clears grouping');

    // The parameter window changes drafts until Apply requests a new API snapshot.
    await evaluate('document.getElementById("parameters").click()');
    assert.equal(await evaluate('document.getElementById("parameter-dialog").open'), true);
    assert.equal(await evaluate('document.activeElement.id'), 'parameter-search');
    assert.equal(await evaluate('document.querySelectorAll("#parameter-list tr").length'), 11);
    assert.deepEqual(await evaluate('Array.from(document.querySelectorAll("#parameter-list input:checked"), checkbox => checkbox.dataset.parameterId)'),
      ['-1006916', '-1006901', '-1006900', '-1006902', '-1006920'], 'starter columns are already checked');
    assert.equal(await evaluate('document.querySelectorAll("#parameter-list input:checked:disabled").length'), 5, 'starter columns remain fixed');
    assert.equal(await evaluate(`document.querySelector('[data-parameter-id="-1012112"]').checked`), false, 'hidden Phase column is not checked');
    assert.deepEqual(await evaluate('Array.from(document.querySelectorAll(".parameter-list thead th"), cell => cell.textContent.trim())'),
      ['', 'Parameter name', 'Data type', 'Parameter type']);
    assert.equal(await evaluate(`document.querySelector('[data-parameter-id="201"]').closest("tr").lastElementChild.textContent`), 'Shared Parameter');
    assert.equal(await evaluate(`document.querySelector('[data-parameter-id="203"]').closest("tr").lastElementChild.textContent`), 'Project Parameter');
    assert.equal(await evaluate('document.querySelectorAll("#parameter-list small").length'), 2, 'duplicate names have distinct IDs');
    assert.equal(await evaluate('document.querySelectorAll("#parameter-list img").length'), 0);
    await evaluate('document.getElementById("parameter-checked-only").click()');
    assert.equal(await evaluate('document.querySelectorAll("#parameter-list tr").length'), 5, 'checked-only includes locked starter parameters');
    assert.equal(await evaluate('Array.from(document.querySelectorAll("#parameter-list .parameter-origin")).every(cell => cell.textContent === "Built-In")'), true);
    await evaluate('document.getElementById("parameter-search").value="shared parameter"; document.getElementById("parameter-search").dispatchEvent(new Event("input"))');
    assert.equal(await evaluate('document.querySelectorAll("#parameter-list tr").length'), 0, 'search intersects checked-only');
    await evaluate('document.getElementById("parameter-checked-only").click()');
    assert.equal(await evaluate('document.querySelectorAll("#parameter-list tr").length'), 1, 'search includes parameter origin');
    await evaluate(`document.querySelector('[data-parameter-id="201"]').click(); document.getElementById("parameter-checked-only").click()`);
    assert.equal(await evaluate('document.querySelectorAll("#parameter-list tr").length'), 1);
    await evaluate(`document.querySelector('[data-parameter-id="201"]').click()`);
    assert.equal(await evaluate('document.querySelectorAll("#parameter-list tr").length'), 0, 'unchecking immediately removes the row');
    assert.equal(await evaluate('document.getElementById("parameters-empty").hidden'), false);
    await evaluate('document.getElementById("parameter-search").value=""; document.getElementById("parameter-search").dispatchEvent(new Event("input")); document.getElementById("parameter-checked-only").click()');
    assert.equal(await evaluate('document.querySelectorAll("#parameter-list tr").length'), 11, 'toggle does not alter draft selections');
    const beforeCancel = await evaluate('window.__requests.length');
    await evaluate(`document.querySelector('[data-parameter-id="201"]').click(); document.getElementById("parameter-cancel").click(); document.getElementById("parameters").click()`);
    assert.equal(await evaluate('window.__requests.length'), beforeCancel, 'Cancel never requests a model refresh');
    assert.equal(await evaluate(`document.querySelector('[data-parameter-id="201"]').checked`), false);
    await evaluate('document.getElementById("parameter-search").value="air flow"; document.getElementById("parameter-search").dispatchEvent(new Event("input"))');
    assert.equal(await evaluate('document.querySelectorAll("#parameter-list tr").length'), 1, 'search matches data type');
    await evaluate(`document.getElementById("parameters-all").click(); document.getElementById("parameter-search").value=""; document.getElementById("parameter-search").dispatchEvent(new Event("input")); document.querySelector('[data-parameter-id="201"]').click(); document.querySelector('[data-parameter-id="204"]').click(); document.getElementById("parameter-apply").click()`);
    assert.deepEqual(await evaluate('window.__requests.at(-1).parameterIds'), ['-200', '201', '204']);
    snapshot.selectedParameterIds = ['-200', '201', '204'];
    await deliver({ type: 'snapshot', snapshot });
    await deliver({ type: 'result', kind: 'success', message: 'Parameters loaded.' });
    await waitFor('!document.getElementById("refresh").disabled');
    assert.equal(await evaluate(`${grid}.getColumns().filter(column => column.getField().startsWith("parameterValues.")).length`), 3);
    assert.equal(await evaluate('document.querySelectorAll("#grid img").length'), 0, 'parameter names and values cannot inject HTML');
    for (const field of ['parameterValues.p-200', 'parameterValues.p201', 'parameterValues.p204']) {
      await groupBy('');
      await dragColumn(field, '#group-drop-zone');
      assert.equal(await evaluate('document.querySelector(".group-chip")?.dataset.field || ""'), field, 'added parameter headers can be dropped for grouping');
      assert.equal(await evaluate(`${grid}.getGroups().reduce((count, group) => count + group.getRows().length, 0)`), 3);
    }
    assert.equal(await evaluate('document.querySelectorAll("#grid img").length'), 0, 'parameter group labels stay plain text');
    await groupBy('parameterValues.p201');
    assert.equal(await evaluate('Array.from(document.querySelectorAll(".group-toggle")).some(header => header.textContent.includes("No value"))'), true);
    await groupBy('parameterValues.p-200');
    assert.equal(await evaluate('Array.from(document.querySelectorAll(".group-toggle")).some(header => header.textContent.includes("0 CFM"))'), true);
    await evaluate(`${grid}.setSort("parameterValues.p-200", "asc")`);
    assert.deepEqual(await evaluate(`${grid}.getRows("active").map(row => row.getData().parameterValues["p-200"])`), [0, 200, 1000]);
    assert.equal(await evaluate(`${grid}.getRow("9007199254740994").getCell("parameterValues.p-200").getElement().textContent`), '200 CFM');
    await evaluate('document.getElementById("refresh").click()');
    await deliver({ type: 'snapshot', snapshot });
    await deliver({ type: 'result', kind: 'success', message: 'Refreshed.' });
    await waitFor('!document.getElementById("refresh").disabled');
    assert.equal(await evaluate(`${grid}.getSorters()[0].field`), 'parameterValues.p-200');
    assert.equal(await evaluate('document.querySelector(".group-chip")?.dataset.field || ""'), 'parameterValues.p-200');
    await evaluate('document.getElementById("parameters").click()');
    assert.equal(await evaluate('document.querySelectorAll("#parameter-list input:checked").length'), 8, 'reopening checks starter and applied columns');
    await evaluate('document.getElementById("parameter-checked-only").click()');
    assert.equal(await evaluate('document.querySelectorAll("#parameter-list tr").length'), 8, 'checked-only includes applied extra columns');
    await evaluate(`document.querySelector('[data-parameter-id="201"]').click()`);
    assert.equal(await evaluate('document.querySelectorAll("#parameter-list tr").length'), 7);
    await evaluate(`document.getElementById("parameter-checked-only").click(); document.querySelector('[data-parameter-id="201"]').click()`);
    await deliver({ type: 'context', active: false, outdated: false });
    await waitFor('document.getElementById("parameter-apply").disabled');
    await deliver({ type: 'context', active: true, outdated: false });
    await waitFor('!document.getElementById("parameter-apply").disabled');
    const pickerScreenshot = await command('Page.captureScreenshot', { format: 'png' });
    fs.writeFileSync(path.join(profile, 'parameters-picker.png'), Buffer.from(pickerScreenshot.data, 'base64'));
    await evaluate('document.querySelectorAll("#parameter-list input:checked").forEach(checkbox => checkbox.click()); document.getElementById("parameter-apply").click()');
    snapshot.selectedParameterIds = [];
    await deliver({ type: 'snapshot', snapshot });
    await deliver({ type: 'result', kind: 'success', message: 'Parameters removed.' });
    await waitFor('!document.getElementById("refresh").disabled');
    assert.equal(await evaluate(`${grid}.getColumns().filter(column => column.getField().startsWith("parameterValues.")).length`), 0);
    assert.ok(await evaluate(`${grid}.getSorters().every(sorter => !sorter.field.startsWith("parameterValues."))`), 'removed columns cannot leave stale sorting');
    assert.equal(await evaluate('document.querySelector(".group-chip")?.dataset.field || ""'), '', 'removing a grouped parameter restores the flat table');
    assert.equal(await evaluate(`${grid}.getGroups().length`), 0);
    await evaluate('document.getElementById("phase").value="all"; document.getElementById("phase").dispatchEvent(new Event("change")); document.getElementById("parameters").click()');
    assert.equal(await evaluate(`document.querySelector('[data-parameter-id="-1012112"]').checked`), true, 'visible Phase column is checked');
    await evaluate('document.getElementById("parameter-cancel").click(); document.getElementById("clear-filters").click()');
    await groupBy('level');
    await evaluate('document.getElementById("phase").value="all"; document.getElementById("phase").dispatchEvent(new Event("change"))');
    await dragColumn('phase', '#group-drop-zone');
    await dragColumn('name', '#group-drop-zone');
    assert.deepEqual(await evaluate(grouped), ['level', 'phase', 'name'], 'header drops append grouping levels');
    await evaluate('document.getElementById("clear-filters").click()');
    assert.deepEqual(await evaluate(grouped), ['level', 'name'], 'hiding Phase removes only that level');
    await groupBy('');

    // Save complete views, resolve unsaved changes, and simulate the model-backed bridge.
    snapshot.selectedParameterIds = ['-200', '201'];
    await deliver({ type: 'snapshot', snapshot });
    await waitFor(`${grid}.getColumns().some(column => column.getField() === 'parameterValues.p201')`);
    await groupBy('level');
    await dragColumn('parameterValues.p201', '#group-drop-zone');
    await dragColumn('name', '#group-drop-zone');
    await configureHeaders([{name:'Notes',fields:['name','parameterValues.p201']}, {name:'Airflow',fields:['parameterValues.p-200']}]);
    snapshot.selectedParameterIds = ['-200'];
    await deliver({ type: 'snapshot', snapshot });
    await waitFor(`!${grid}.getColumns().some(column => column.getField() === 'parameterValues.p201')`);
    assert.deepEqual(await evaluate(grouped), ['level', 'name'], 'removing a parameter retains other hierarchy levels');
    assert.deepEqual(await evaluate(bands), [{name:'Notes',fields:['name']}, {name:'Airflow',fields:['parameterValues.p-200']}], 'parameter removal prunes just its band membership');
    snapshot.selectedParameterIds = ['-200', '201'];
    await deliver({ type: 'snapshot', snapshot });
    await waitFor(`${grid}.getColumns().some(column => column.getField() === 'parameterValues.p201')`);
    await configureHeaders([]);
    await groupBy('');
    await evaluate(`${grid}.moveColumn('parameterValues.p201', 'name', true); ${grid}.getColumn('name').setWidth(245)`);
    await evaluate('document.getElementById("phase").value="all"; document.getElementById("phase").dispatchEvent(new Event("change")); document.getElementById("level").value="10"; document.getElementById("level").dispatchEvent(new Event("change")); document.getElementById("search").value="Office"; document.getElementById("search").dispatchEvent(new Event("input"))');
    await dragColumn('level', '#group-drop-zone');
    await dragColumn('parameterValues.p-200', '#group-drop-zone');
    const savedBands = [{name:'Dimensions',fields:['areaValue','heightValue']},
      {name:'Ventilation',fields:['parameterValues.p201','parameterValues.p-200']}];
    await configureHeaders(savedBands);
    await dragColumn('areaValue', 'band:parameterValues.p201', false, false, 'band');
    assert.deepEqual((await evaluate(bands)).map(group => group.name), ['Dimensions','Ventilation'], 'whole-band drag order is the saved working order');
    const appliedSavedBands = await evaluate(bands);
    await evaluate(`${grid}.setSort([{column:'number',dir:'asc'},{column:'parameterValues.p-200',dir:'desc'}])`);
    const bandsScreenshot = await command('Page.captureScreenshot', {format:'png'});
    fs.writeFileSync(path.join(profile, 'column-header-groups.png'), Buffer.from(bandsScreenshot.data, 'base64'));
    await evaluate('document.getElementById("table-save-as").click()');
    await nameDialog('Ventilation Review');
    const firstSave = await respondToTable();
    const firstId = firstSave.record.id;
    assert.deepEqual(firstSave.request.layout.filters, { phase: 'all', level: '10', query: 'Office' });
    assert.deepEqual(firstSave.request.layout.groupFields, ['level', 'parameterValues.p-200']);
    assert.deepEqual(firstSave.request.layout.headerGroups, appliedSavedBands, 'saving includes custom header names and memberships');
    assert.equal(await evaluate('document.getElementById("saved-table").value'), firstId);
    assert.equal(await evaluate('document.getElementById("table-dirty").textContent'), '');
    await dragColumn('areaValue', 'band:parameterValues.p201', true, true, 'band');
    assert.equal(await evaluate('document.getElementById("table-dirty").textContent'), '', 'a cancelled band drag leaves a saved view clean');
    await dragColumn('areaValue', 'band:parameterValues.p201', true, false, 'band');
    assert.equal(await evaluate('document.getElementById("table-dirty").textContent'), 'Unsaved changes', 'a whole-band reorder marks a saved view dirty');
    await dragColumn('areaValue', 'band:parameterValues.p201', false, false, 'band');
    assert.equal(await evaluate('document.getElementById("table-dirty").textContent'), '', 'moving a band back restores the saved layout');
    await configureHeaders(appliedSavedBands.map((group,index) => ({...group,name:index ? group.name : 'Renamed band'})));
    assert.equal(await evaluate('document.getElementById("table-dirty").textContent'), 'Unsaved changes', 'a band rename alone marks the view dirty');
    await configureHeaders(appliedSavedBands);
    assert.equal(await evaluate('document.getElementById("table-dirty").textContent'), '', 'restoring just the band name restores the saved layout');
    await evaluate('document.getElementById("search").value="Changed working view"; document.getElementById("search").dispatchEvent(new Event("input"))');
    assert.equal(await evaluate('document.getElementById("table-dirty").textContent'), 'Unsaved changes');
    await chooseTable('');
    await waitFor('document.getElementById("table-dialog").open');
    await evaluate('document.getElementById("table-dialog-cancel").click()');
    assert.equal(await evaluate('document.getElementById("saved-table").value'), firstId);
    assert.equal(await evaluate('document.getElementById("search").value'), 'Changed working view');
    await chooseTable('');
    await waitFor('document.getElementById("table-dialog").open');
    await evaluate('document.getElementById("table-dialog-discard").click()');
    await respondToTable();
    assert.equal(await evaluate('document.getElementById("saved-table").value'), '');
    assert.equal(await evaluate('document.getElementById("search").value'), '');
    snapshot.rows[0].parameterValues['p-200'] = 1500;
    snapshot.rows[0].parameterTexts['p-200'] = '1500 CFM';
    await chooseTable(firstId);
    await respondToTable();
    assert.deepEqual(await evaluate(`${grid}.getColumns().map(column => column.getField())`), firstSave.request.layout.columns.map(column => column.field));
    assert.equal(await evaluate(`${grid}.getColumn('name').getWidth()`), 245);
    assert.equal(await evaluate('document.getElementById("phase").value'), 'all');
    assert.equal(await evaluate('document.getElementById("level").value'), '10');
    assert.equal(await evaluate('document.getElementById("search").value'), 'Office');
    assert.deepEqual(await evaluate(grouped), ['level', 'parameterValues.p-200'], 'saved definitions restore hierarchy order');
    assert.deepEqual(await evaluate(bands), appliedSavedBands, 'reopening restores named header bands and dynamic members');
    assert.deepEqual((await evaluate(sortPriorities)).sort((a,b) => a.field.localeCompare(b.field)),
      [{field:'number',priority:'1'},{field:'parameterValues.p-200',priority:'2'}], 'saved-table restoration recreates dynamic column sort priorities');
    assert.equal(await evaluate(`${grid}.getRow('9007199254740993').getCell('parameterValues.p-200').getElement().textContent`), '1500 CFM', 'opening a definition reads current values');
    await evaluate('document.getElementById("table-save-as").click()');
    await nameDialog('Ventilation copy');
    const secondSave = await respondToTable();
    assert.notEqual(firstId, secondSave.record.id);
    await evaluate('document.getElementById("search").value="draft"; document.getElementById("search").dispatchEvent(new Event("input")); document.getElementById("table-rename").click()');
    await nameDialog('Renamed copy');
    await respondToTable();
    assert.equal(await evaluate('document.getElementById("search").value'), 'draft', 'rename preserves unsaved layout changes');
    assert.equal(records.get(secondSave.record.id).layout.filters.query, 'Office', 'rename does not save the draft layout');
    records.get(secondSave.record.id).revision = randomUUID(); // A teammate saved a newer definition.
    await evaluate('document.getElementById("table-save").click()');
    await respondToTable('This saved table has a newer revision. Your changes have been retained.');
    assert.equal(await evaluate('document.getElementById("search").value'), 'draft');
    assert.equal(await evaluate('document.getElementById("table-dirty").textContent'), 'Unsaved changes');
    await evaluate('document.getElementById("table-reopen").click()');
    await waitFor('document.getElementById("table-dialog").open');
    await evaluate('document.getElementById("table-dialog-discard").click()');
    const reopened = await respondToTable();
    assert.equal(reopened.record.revision, records.get(secondSave.record.id).revision);
    assert.equal(await evaluate('document.getElementById("search").value'), 'Office', 'Reopen explicitly replaces the working settings');
    await evaluate('document.getElementById("search").value="draft"; document.getElementById("search").dispatchEvent(new Event("input"))');
    await deliver({ type: 'closeRequest' });
    await waitFor('document.getElementById("table-dialog").open');
    await evaluate('document.getElementById("table-dialog-cancel").click()');
    await deliver({ type: 'closeRequest' });
    await waitFor('document.getElementById("table-dialog").open');
    await evaluate('document.getElementById("table-dialog-confirm").click()');
    await respondToTable('Another user owns this saved table. Your changes have been retained.');
    assert.notEqual(await evaluate('window.__requests.at(-1).action'), 'closeWindow', 'failed save cannot close the window');
    await deliver({ type: 'closeRequest' });
    await waitFor('document.getElementById("table-dialog").open');
    await evaluate('document.getElementById("table-dialog-discard").click()');
    await waitFor('window.__requests.at(-1).action === "closeWindow"');
    await chooseTable('');
    await waitFor('document.getElementById("table-dialog").open');
    await evaluate('document.getElementById("table-dialog-discard").click()');
    await respondToTable();
    await evaluate(`${grid}.getColumn('name').setWidth(260)`);
    await chooseTable(firstId);
    await waitFor('document.getElementById("table-dialog").open');
    await evaluate('document.getElementById("table-dialog-confirm").click()');
    await nameDialog('Renamed copy'); // Names can collide; the picker must still distinguish definitions.
    await respondToTable();
    await respondToTable(); // Successful save continues the original switch request.
    assert.equal(await evaluate('document.getElementById("saved-table").value'), firstId);
    const duplicateLabels = await evaluate('Array.from(document.getElementById("saved-table").options).filter(option => option.textContent.startsWith("Renamed copy")).map(option => option.textContent)');
    assert.equal(duplicateLabels.length, 2);
    assert.equal(new Set(duplicateLabels).size, 2, 'same-named tables have distinct labels');
    await evaluate('document.getElementById("table-delete").click()');
    await waitFor('document.getElementById("table-dialog").open');
    await evaluate('document.getElementById("table-dialog-confirm").click()');
    await respondToTable();
    assert.equal(records.has(firstId), false);
    assert.equal(await evaluate('document.getElementById("saved-table").value'), '');
    assert.equal(await evaluate(`${grid}.getColumns().filter(column => column.getField().startsWith('parameterValues.')).length`), 2, 'deleting a definition retains the working view');
    snapshot.canSaveTables = false;
    await deliver({ type: 'snapshot', snapshot });
    await waitFor('document.getElementById("table-save-as").disabled');
    assert.equal(await evaluate('document.getElementById("saved-table").disabled'), false, 'read-only models can load saved definitions');
    await chooseTable(secondSave.record.id);
    await waitFor('document.getElementById("table-dialog").open');
    assert.equal(await evaluate('document.getElementById("table-dialog-confirm").disabled'), true, 'read-only dirty prompt cannot save');
    await evaluate('document.getElementById("table-dialog-discard").click()');
    await respondToTable();
    snapshot.canSaveTables = true;
    await deliver({ type: 'snapshot', snapshot });
    await chooseTable('');
    await respondToTable();
    assert.equal(await evaluate('document.getElementById("table-dirty").textContent'), '');
    await chooseTable(secondSave.record.id);
    await respondToTable(null, response => {
      response.snapshot = copy(response.snapshot);
      response.snapshot.selectedParameterIds = ['-200'];
      response.layout.columns = response.layout.columns.filter(column => column.field !== 'parameterValues.p201');
      response.layout.sorters = response.layout.sorters.filter(sorter => sorter.field !== 'parameterValues.p201');
      response.warnings = ['Unavailable column: Comments'];
    });
    assert.equal(await evaluate('document.getElementById("table-note").textContent.includes("Unavailable column")'), true);
    assert.equal(await evaluate(`${grid}.getColumns().some(column => column.getField() === 'parameterValues.p201')`), false);
    assert.ok(records.get(secondSave.record.id).layout.columns.some(column => column.field === 'parameterValues.p201'), 'loading missing columns never rewrites the definition');
    assert.equal(await evaluate('document.getElementById("table-dirty").textContent'), 'Unsaved changes', 'partial restoration needs an explicit save to update the definition');
    await chooseTable('');
    await waitFor('document.getElementById("table-dialog").open');
    await evaluate('document.getElementById("table-dialog-discard").click()');
    await respondToTable();

    // Exercise the real virtual DOM with thousands of rows, not just the pure model.
    await groupBy('heightValue');
    snapshot.rows = Array.from({ length: 6000 }, (_, index) =>
      makeRow(String(index + 100), String(index + 1), '2', '10', 0, index));
    const started = Date.now();
    await deliver({ type: 'snapshot', snapshot });
    await waitFor('document.getElementById("counts").textContent.startsWith("6000 of 6000")');
    assert.ok(await evaluate('document.querySelectorAll(".tabulator-row").length < 200'), 'rows must be virtualized');
    assert.equal(await evaluate(`${grid}.getGroups()[0].getRows().length`), 6000, 'large datasets remain grouped');
    const largeTableMs = Date.now() - started;
    const screenshot = await command('Page.captureScreenshot', { format: 'png' });
    const screenshotPath = path.join(profile, 'spaces-table.png');
    fs.writeFileSync(screenshotPath, Buffer.from(screenshot.data, 'base64'));
    await command('Emulation.setDeviceMetricsOverride', { width: 760, height: 480, deviceScaleFactor: 1, mobile: false });
    for (const field of ['level', 'name']) await dragColumn(field, '#group-drop-zone');
    await evaluate('new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))');
    const bounds = await evaluate('({gridBottom: document.getElementById("grid").getBoundingClientRect().bottom, gridHeight: document.getElementById("grid").getBoundingClientRect().height, footerTop: document.querySelector("footer").getBoundingClientRect().top})');
    assert.ok(bounds.gridHeight >= 100 && bounds.gridBottom <= bounds.footerTop, 'minimum window size keeps grid and footer usable');
    await evaluate('document.getElementById("parameters").click()');
    assert.equal(await evaluate('document.getElementById("parameter-apply").getBoundingClientRect().bottom < innerHeight'), true, 'parameter actions fit the minimum window');
    await evaluate('document.getElementById("parameter-cancel").click()');
    await evaluate('document.getElementById("column-groups").click(); document.getElementById("header-group-add").click()');
    assert.equal(await evaluate('document.getElementById("header-groups-apply").getBoundingClientRect().bottom < innerHeight'), true, 'column group editor actions fit the minimum window');
    await evaluate('document.getElementById("header-groups-cancel").click()');
    snapshot.rows = [];
    snapshot.parameters = [];
    await deliver({ type: 'snapshot', snapshot });
    await waitFor('document.getElementById("counts").textContent.startsWith("0 of 0")');
    assert.equal(await evaluate('document.getElementById("select-revit").disabled'), true);
    assert.equal(await evaluate('document.getElementById("show-revit").disabled'), true);
    assert.deepEqual(errors, [], 'no uncaught browser exceptions');
    console.log(`Browser checks passed: multi-sort priority subscripts, Ctrl+click sort/direction/clear, saved priorities, whole-band dragging, row grouping, picker/filtering, and 6,000 grouped rows (${largeTableMs}ms).`);
    console.log(`Screenshot: ${screenshotPath}`);
    console.log(`Grouping hierarchy: ${path.join(profile, 'grouping-hierarchy.png')}`);
    console.log(`Column header groups: ${path.join(profile, 'column-header-groups.png')}`);
    console.log(`Parameter picker: ${path.join(profile, 'parameters-picker.png')}`);
  } finally {
    if (command && socket?.readyState === WebSocket.OPEN) {
      try { await command('Browser.close'); } catch { /* Browser can close before its reply. */ }
    }
    socket?.close();
    browser.kill();
  }
}

main().catch(error => { console.error(error); process.exitCode = 1; });
