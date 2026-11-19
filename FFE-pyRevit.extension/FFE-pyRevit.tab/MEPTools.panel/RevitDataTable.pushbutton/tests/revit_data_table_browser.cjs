// Offline integration test using installed Edge and its DevTools protocol; no npm dependencies.
// Run: node tests/revit_data_table_browser.cjs
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { pathToFileURL } = require('node:url');
const { spawn } = require('node:child_process');

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
    socket.onmessage = event => {
      const message = JSON.parse(event.data);
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
      const result = await command('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true });
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
    });
    const snapshot = {
      documentKey: 'model-session', documentTitle: 'Mechanical Model', defaultPhaseId: '2',
      refreshedAt: new Date().toISOString(), skippedCount: 0, issues: [],
      phases: [{ id: '1', name: 'Existing' }, { id: '2', name: 'New' }],
      levels: [{ id: '10', name: 'Ground', elevation: 0 }, { id: '11', name: 'Basement', elevation: -10 }],
      rows: [makeRow('9007199254740993', '10', '2', '10', 0, 100),
        makeRow('9007199254740994', '002', '2', '10', 0, 20, '<img src=x onerror="window.injected=true">'),
        makeRow('3', '1', '2', '11', -10, 0), makeRow('4', '1', '1', '10', 0, 50)],
    };
    const deliver = message => evaluate(`window.__hostMessage(${JSON.stringify(message)})`);
    await deliver({ type: 'snapshot', snapshot });
    await waitFor('document.getElementById("counts").textContent.startsWith("3 of 4")');
    assert.deepEqual(await evaluate(`${grid}.getRows("active").map(row => row.getData().id)`),
      ['3', '9007199254740994', '9007199254740993'], 'level elevation then natural number ordering');
    assert.equal(await evaluate('document.querySelectorAll("#grid img").length'), 0, 'model text cannot inject markup');
    assert.equal(await evaluate(`${grid}.getColumn("phase").isVisible()`), false);
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
    await deliver({ type: 'context', active: false, outdated: false });
    await waitFor('document.getElementById("select-revit").disabled');
    assert.equal(await evaluate('document.getElementById("refresh").disabled'), true);
    await deliver({ type: 'context', active: true, outdated: false });
    await waitFor('!document.getElementById("refresh").disabled');
    await evaluate('document.getElementById("phase").value = "all"; document.getElementById("phase").dispatchEvent(new Event("change"))');
    assert.equal(await evaluate(`${grid}.getColumn("phase").isVisible()`), true);
    assert.equal(await evaluate(`${grid}.getDataCount("active")`), 4);
    await evaluate('document.getElementById("search").value = "nonexistent"; document.getElementById("search").dispatchEvent(new Event("input"))');
    assert.equal(await evaluate(`${grid}.getDataCount("active")`), 0);
    assert.equal(await evaluate(`${grid}.getSelectedData().length`), 0, 'hidden rows cannot remain selected for navigation');
    await evaluate('document.getElementById("clear-filters").click(); document.getElementById("grid").focus(); document.getElementById("grid").dispatchEvent(new KeyboardEvent("keydown", {key:"ArrowDown"})); document.getElementById("grid").dispatchEvent(new KeyboardEvent("keydown", {key:" "}))');
    assert.equal(await evaluate(`${grid}.getSelectedData().length`), 1, 'keyboard row selection');

    // Exercise the real virtual DOM with thousands of rows, not just the pure model.
    snapshot.rows = Array.from({ length: 6000 }, (_, index) =>
      makeRow(String(index + 100), String(index + 1), '2', '10', 0, index));
    const started = Date.now();
    await deliver({ type: 'snapshot', snapshot });
    await waitFor('document.getElementById("counts").textContent.startsWith("6000 of 6000")');
    assert.ok(await evaluate('document.querySelectorAll(".tabulator-row").length < 200'), 'rows must be virtualized');
    const largeTableMs = Date.now() - started;
    const screenshot = await command('Page.captureScreenshot', { format: 'png' });
    const screenshotPath = path.join(profile, 'spaces-table.png');
    fs.writeFileSync(screenshotPath, Buffer.from(screenshot.data, 'base64'));
    await command('Emulation.setDeviceMetricsOverride', { width: 760, height: 480, deviceScaleFactor: 1, mobile: false });
    await evaluate('new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))');
    const bounds = await evaluate('({gridBottom: document.getElementById("grid").getBoundingClientRect().bottom, gridHeight: document.getElementById("grid").getBoundingClientRect().height, footerTop: document.querySelector("footer").getBoundingClientRect().top})');
    assert.ok(bounds.gridHeight >= 100 && bounds.gridBottom <= bounds.footerTop, 'minimum window size keeps grid and footer usable');
    snapshot.rows = [];
    await deliver({ type: 'snapshot', snapshot });
    await waitFor('document.getElementById("counts").textContent.startsWith("0 of 0")');
    assert.equal(await evaluate('document.getElementById("select-revit").disabled'), true);
    assert.equal(await evaluate('document.getElementById("show-revit").disabled'), true);
    assert.deepEqual(errors, [], 'no uncaught browser exceptions');
    console.log(`Browser checks passed: offline load, sorting, identity, text escaping, selection, refresh, context, phases, keyboard, minimum window, empty model, and 6,000 rows (${largeTableMs}ms).`);
    console.log(`Screenshot: ${screenshotPath}`);
  } finally {
    if (command && socket?.readyState === WebSocket.OPEN) {
      try { await command('Browser.close'); } catch { /* Browser can close before its reply. */ }
    }
    socket?.close();
    browser.kill();
  }
}

main().catch(error => { console.error(error); process.exitCode = 1; });
