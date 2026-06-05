(function () {
  "use strict";

  var SVG_NS = "http://www.w3.org/2000/svg";
  var SNAP = 12;
  var MIN_ZOOM = 0.2;
  var MAX_ZOOM = 3.2;

  var dom = {};
  var state = {
    diagram: null,
    selected: null,
    mode: "select",
    dirty: false,
    view: { zoom: 1, panX: 0, panY: 0 },
    drag: null,
    history: { past: [], future: [] }
  };

  function hasWebViewBridge() {
    return !!(window.chrome && window.chrome.webview && window.chrome.webview.postMessage);
  }

  function postWebViewMessage(message) {
    if (!hasWebViewBridge()) {
      return false;
    }
    window.chrome.webview.postMessage(JSON.stringify(message));
    return true;
  }

  function clone(value) {
    return JSON.parse(JSON.stringify(value));
  }

  function asArray(value) {
    return Array.isArray(value) ? value : [];
  }

  function byId(id) {
    return document.getElementById(id);
  }

  function svgEl(tag, attrs) {
    var el = document.createElementNS(SVG_NS, tag);
    Object.keys(attrs || {}).forEach(function (key) {
      if (attrs[key] !== undefined && attrs[key] !== null) {
        el.setAttribute(key, String(attrs[key]));
      }
    });
    return el;
  }

  function textOr(value, fallback) {
    var text = value === undefined || value === null ? "" : String(value);
    return text.trim() || fallback || "";
  }

  function numberOr(value, fallback) {
    var number = Number(value);
    return Number.isFinite(number) ? number : fallback;
  }

  function snap(value) {
    return Math.round(value / SNAP) * SNAP;
  }

  function makeId(prefix) {
    return prefix + "-" + Date.now().toString(36) + "-" + Math.round(Math.random() * 9999).toString(36);
  }

  function normalizeDiagram(payload) {
    var diagram = payload && typeof payload === "object" ? clone(payload) : {};
    diagram.schemaVersion = diagram.schemaVersion || 1;
    diagram.generationMode = textOr(diagram.generationMode, "compact-significant-nodes");
    diagram.documentTitle = textOr(diagram.documentTitle, "");
    diagram.systemName = textOr(diagram.systemName, "");
    diagram.systemUniqueId = textOr(diagram.systemUniqueId, "");
    diagram.nodes = asArray(diagram.nodes);
    diagram.edges = asArray(diagram.edges);
    diagram.symbols = asArray(diagram.symbols);
    diagram.labels = asArray(diagram.labels);
    diagram.warnings = asArray(diagram.warnings);
    diagram.canvas = diagram.canvas || {};
    diagram.canvas.width = Math.max(720, numberOr(diagram.canvas.width, 900));
    diagram.canvas.height = Math.max(460, numberOr(diagram.canvas.height, 600));

    diagram.nodes.forEach(function (node) {
      node.id = textOr(node.id, makeId("node"));
      node.kind = textOr(node.kind, "junction");
      node.label = textOr(node.label, node.kind);
      node.x = snap(numberOr(node.x, 100));
      node.y = snap(numberOr(node.y, 100));
      node.diameter = textOr(node.diameter, "");
      node.flow = textOr(node.flow, "");
    });

    diagram.edges.forEach(function (edge) {
      edge.id = textOr(edge.id, makeId("edge"));
      edge.from = textOr(edge.from, "");
      edge.to = textOr(edge.to, "");
      edge.label = textOr(edge.label, "");
      edge.diameter = textOr(edge.diameter, "");
      edge.flow = textOr(edge.flow, "");
      edge.flowDirection = textOr(edge.flowDirection, "fromTo");
      edge.points = asArray(edge.points);
      edge.points.forEach(function (point) {
        point.x = snap(numberOr(point.x, 0));
        point.y = snap(numberOr(point.y, 0));
      });
    });

    diagram.symbols.forEach(function (symbol) {
      symbol.id = textOr(symbol.id, makeId("symbol"));
      symbol.kind = textOr(symbol.kind, "valve");
      symbol.label = textOr(symbol.label, symbol.kind);
      symbol.x = snap(numberOr(symbol.x, 160));
      symbol.y = snap(numberOr(symbol.y, 160));
      symbol.diameter = textOr(symbol.diameter, "");
      symbol.flow = textOr(symbol.flow, "");
    });

    diagram.labels.forEach(function (label) {
      label.id = textOr(label.id, makeId("label"));
      label.kind = textOr(label.kind, "note");
      label.text = textOr(label.text, "");
      label.x = snap(numberOr(label.x, 80));
      label.y = snap(numberOr(label.y, 80));
    });

    return diagram;
  }

  function fixtureDiagram() {
    return normalizeDiagram({
      schemaVersion: 1,
      generationMode: "compact-significant-nodes",
      documentTitle: "Standalone Preview",
      systemName: "Bldg 69 - Pipe One-Line",
      systemUniqueId: "fixture",
      systemId: 69,
      warnings: [],
      canvas: { width: 980, height: 640 },
      nodes: [
        { id: "n-pump", kind: "pump", label: "69-P4", x: 120, y: 430 },
        { id: "n-main-a", kind: "branch", label: "", x: 220, y: 430 },
        { id: "n-riser-1", kind: "branch", label: "", x: 220, y: 340 },
        { id: "n-riser-2", kind: "branch", label: "", x: 220, y: 250 },
        { id: "n-size-1", kind: "sizeChange", label: "Dia 20 in to Dia 12 in", x: 330, y: 250 },
        { id: "n-ch1", kind: "valve", label: "69-CH1", x: 420, y: 250 },
        { id: "n-ch2", kind: "valve", label: "69-CH2", x: 420, y: 340 },
        { id: "n-meter", kind: "meter", label: "WFMD-1", x: 620, y: 250 },
        { id: "n-return", kind: "accessory", label: "CHR", x: 780, y: 250 },
        { id: "n-strainer", kind: "strainer", label: "69-WS2", x: 420, y: 520 }
      ],
      edges: [
        { id: "e1", from: "n-pump", to: "n-main-a", label: "Dia 20 in", points: [{ x: 120, y: 430 }, { x: 220, y: 430 }] },
        { id: "e2", from: "n-main-a", to: "n-riser-1", label: "Dia 20 in", points: [{ x: 220, y: 430 }, { x: 220, y: 340 }] },
        { id: "e3", from: "n-riser-1", to: "n-riser-2", label: "Dia 20 in", points: [{ x: 220, y: 340 }, { x: 220, y: 250 }] },
        { id: "e4a", from: "n-riser-2", to: "n-size-1", label: "Dia 20 in", points: [{ x: 220, y: 250 }, { x: 330, y: 250 }] },
        { id: "e4b", from: "n-size-1", to: "n-ch1", label: "1760 GPM | Dia 12 in", points: [{ x: 330, y: 250 }, { x: 420, y: 250 }] },
        { id: "e5", from: "n-riser-1", to: "n-ch2", label: "1760 GPM | Dia 12 in", points: [{ x: 220, y: 340 }, { x: 420, y: 340 }] },
        { id: "e6", from: "n-ch1", to: "n-meter", label: "", points: [{ x: 420, y: 250 }, { x: 620, y: 250 }] },
        { id: "e7", from: "n-meter", to: "n-return", label: "", points: [{ x: 620, y: 250 }, { x: 780, y: 250 }] },
        { id: "e8", from: "n-main-a", to: "n-strainer", label: "Dia 20 in", points: [{ x: 220, y: 430 }, { x: 220, y: 520 }, { x: 420, y: 520 }] }
      ],
      symbols: [],
      labels: [
        { id: "title", kind: "title", text: "Bldg 69", x: 410, y: 52 },
        { id: "subtitle", kind: "note", text: "Top of Slab - 845'", x: 385, y: 72 }
      ]
    });
  }

  function resetHistory() {
    state.history.past = state.diagram ? [clone(state.diagram)] : [];
    state.history.future = [];
    updateHistoryButtons();
  }

  function pushSnapshot() {
    if (!state.diagram) {
      return;
    }
    var snapshot = clone(state.diagram);
    var past = state.history.past;
    var last = past[past.length - 1];
    if (last && JSON.stringify(last) === JSON.stringify(snapshot)) {
      updateHistoryButtons();
      return;
    }
    past.push(snapshot);
    if (past.length > 60) {
      past.shift();
    }
    state.history.future = [];
    updateHistoryButtons();
  }

  function undo() {
    if (state.history.past.length <= 1) {
      return;
    }
    var current = state.history.past.pop();
    state.history.future.push(current);
    state.diagram = clone(state.history.past[state.history.past.length - 1]);
    state.selected = null;
    setDirty(true);
    render();
    updateHistoryButtons();
  }

  function redo() {
    if (!state.history.future.length) {
      return;
    }
    var next = state.history.future.pop();
    state.history.past.push(clone(next));
    state.diagram = clone(next);
    state.selected = null;
    setDirty(true);
    render();
    updateHistoryButtons();
  }

  function updateHistoryButtons() {
    if (!dom.undo) {
      return;
    }
    dom.undo.disabled = state.history.past.length <= 1;
    dom.redo.disabled = state.history.future.length === 0;
  }

  function setDirty(isDirty) {
    var changed = state.dirty !== !!isDirty;
    state.dirty = !!isDirty;
    if (dom.dirtyPill) {
      dom.dirtyPill.textContent = state.dirty ? "Dirty" : "Clean";
      dom.dirtyPill.classList.toggle("is-dirty", state.dirty);
    }
    if (changed) {
      postWebViewMessage({ type: "dirtyStateChanged", dirty: state.dirty });
    }
  }

  function setStatus(status, message) {
    var statusName = textOr(status, "info");
    dom.status.textContent = textOr(message, "Ready");
    dom.status.className = "status is-" + statusName;
  }

  function nodeById(id) {
    if (!state.diagram) {
      return null;
    }
    return state.diagram.nodes.find(function (node) {
      return node.id === id;
    }) || null;
  }

  function edgeById(id) {
    if (!state.diagram) {
      return null;
    }
    return state.diagram.edges.find(function (edge) {
      return edge.id === id;
    }) || null;
  }

  function symbolById(id) {
    if (!state.diagram) {
      return null;
    }
    return state.diagram.symbols.find(function (symbol) {
      return symbol.id === id;
    }) || null;
  }

  function labelById(id) {
    if (!state.diagram) {
      return null;
    }
    return state.diagram.labels.find(function (label) {
      return label.id === id;
    }) || null;
  }

  function selectedObject() {
    if (!state.selected) {
      return null;
    }
    if (state.selected.type === "node") {
      return nodeById(state.selected.id);
    }
    if (state.selected.type === "edge") {
      return edgeById(state.selected.id);
    }
    if (state.selected.type === "symbol") {
      return symbolById(state.selected.id);
    }
    if (state.selected.type === "label") {
      return labelById(state.selected.id);
    }
    return null;
  }

  function edgePoints(edge) {
    var points = asArray(edge.points);
    if (points.length >= 2) {
      return points;
    }
    var from = nodeById(edge.from) || symbolById(edge.from);
    var to = nodeById(edge.to) || symbolById(edge.to);
    if (!from || !to) {
      return [];
    }
    var midX = (numberOr(from.x, 0) + numberOr(to.x, 0)) / 2;
    return [
      { x: numberOr(from.x, 0), y: numberOr(from.y, 0) },
      { x: midX, y: numberOr(from.y, 0) },
      { x: midX, y: numberOr(to.y, 0) },
      { x: numberOr(to.x, 0), y: numberOr(to.y, 0) }
    ];
  }

  function pathFromPoints(points) {
    if (!points.length) {
      return "";
    }
    return points.map(function (point, index) {
      return (index ? "L " : "M ") + numberOr(point.x, 0) + " " + numberOr(point.y, 0);
    }).join(" ");
  }

  function midpoint(points) {
    if (!points.length) {
      return { x: 0, y: 0 };
    }
    return points[Math.floor(points.length / 2)];
  }

  function flowArrowPoints(points) {
    if (!points || points.length < 2) {
      return null;
    }
    var lengths = [];
    var total = 0;
    for (var index = 0; index < points.length - 1; index += 1) {
      var a = points[index];
      var b = points[index + 1];
      var dx = numberOr(b.x, 0) - numberOr(a.x, 0);
      var dy = numberOr(b.y, 0) - numberOr(a.y, 0);
      var length = Math.sqrt(dx * dx + dy * dy);
      lengths.push(length);
      total += length;
    }
    if (total < 1) {
      return null;
    }
    var target = total / 2;
    var traveled = 0;
    for (var i = 0; i < lengths.length; i += 1) {
      var segmentLength = lengths[i];
      if (segmentLength <= 0) {
        continue;
      }
      if (traveled + segmentLength >= target) {
        var p1 = points[i];
        var p2 = points[i + 1];
        var ratio = (target - traveled) / segmentLength;
        var tipX = numberOr(p1.x, 0) + (numberOr(p2.x, 0) - numberOr(p1.x, 0)) * ratio;
        var tipY = numberOr(p1.y, 0) + (numberOr(p2.y, 0) - numberOr(p1.y, 0)) * ratio;
        var dirX = (numberOr(p2.x, 0) - numberOr(p1.x, 0)) / segmentLength;
        var dirY = (numberOr(p2.y, 0) - numberOr(p1.y, 0)) / segmentLength;
        var baseX = tipX - dirX * 13;
        var baseY = tipY - dirY * 13;
        var perpX = -dirY;
        var perpY = dirX;
        return [
          [tipX, tipY],
          [baseX + perpX * 5.5, baseY + perpY * 5.5],
          [baseX - perpX * 5.5, baseY - perpY * 5.5]
        ];
      }
      traveled += segmentLength;
    }
    return null;
  }

  function drawFlowArrow(points) {
    var arrow = flowArrowPoints(points);
    if (!arrow) {
      return;
    }
    dom.edges.appendChild(svgEl("polygon", {
      class: "flow-arrow",
      points: arrow.map(function (point) {
        return point[0] + "," + point[1];
      }).join(" ")
    }));
  }

  function clearGroup(group) {
    while (group.firstChild) {
      group.removeChild(group.firstChild);
    }
  }

  function drawFrame() {
    var canvas = state.diagram.canvas || {};
    var frame = svgEl("rect", {
      class: "diagram-frame",
      x: 20,
      y: 20,
      width: Math.max(200, numberOr(canvas.width, 900) - 40),
      height: Math.max(160, numberOr(canvas.height, 600) - 40)
    });
    dom.edges.appendChild(frame);
  }

  function drawEdges() {
    drawFrame();
    state.diagram.edges.forEach(function (edge) {
      var points = edgePoints(edge);
      var d = pathFromPoints(points);
      if (!d) {
        return;
      }
      var visible = svgEl("path", {
        class: "pipe-edge" + (state.selected && state.selected.type === "edge" && state.selected.id === edge.id ? " is-selected" : ""),
        d: d
      });
      var hit = svgEl("path", {
        class: "edge-hit",
        d: d,
        "data-type": "edge",
        "data-id": edge.id
      });
      dom.edges.appendChild(visible);
      dom.edges.appendChild(hit);
      drawFlowArrow(points);

      var labelText = textOr(edge.label, "");
      if (labelText) {
        var mid = midpoint(points);
        var label = svgEl("text", {
          class: "edge-label",
          x: numberOr(mid.x, 0) + 8,
          y: numberOr(mid.y, 0) - 8
        });
        label.textContent = labelText;
        dom.edges.appendChild(label);
      }
    });
  }

  function addLine(parent, x1, y1, x2, y2) {
    parent.appendChild(svgEl("line", {
      class: "symbol-line",
      x1: x1,
      y1: y1,
      x2: x2,
      y2: y2
    }));
  }

  function addPolygon(parent, points, className) {
    parent.appendChild(svgEl("polygon", {
      class: className || "symbol-line",
      points: points.map(function (point) {
        return point[0] + "," + point[1];
      }).join(" ")
    }));
  }

  function addPolyline(parent, points) {
    parent.appendChild(svgEl("polyline", {
      class: "symbol-line",
      points: points.map(function (point) {
        return point[0] + "," + point[1];
      }).join(" ")
    }));
  }

  function drawSymbolShape(parent, kind) {
    var type = textOr(kind, "junction");
    addLine(parent, -28, 0, 28, 0);

    if (type === "valve") {
      addPolygon(parent, [[-16, -10], [0, 0], [-16, 10]], "symbol-white");
      addPolygon(parent, [[16, -10], [0, 0], [16, 10]], "symbol-white");
      addLine(parent, 0, -14, 0, -24);
      addLine(parent, 0, -24, 12, -24);
      return;
    }

    if (type === "sizeChange") {
      addPolygon(parent, [[-18, -9], [18, -9], [8, 9], [-8, 9]], "symbol-white");
      addLine(parent, -20, 14, 20, -14);
      return;
    }

    if (type === "accessory") {
      parent.appendChild(svgEl("rect", { class: "symbol-white", x: -17, y: -10, width: 34, height: 20 }));
      addLine(parent, -11, 8, 11, -8);
      return;
    }

    if (type === "pump") {
      parent.appendChild(svgEl("circle", { class: "symbol-white", cx: 0, cy: 0, r: 18 }));
      addPolygon(parent, [[-7, -10], [12, 0], [-7, 10]], "symbol-line");
      return;
    }

    if (type === "strainer") {
      addPolygon(parent, [[0, -18], [18, 0], [0, 18], [-18, 0]], "symbol-white");
      addLine(parent, -9, 9, 9, -9);
      addLine(parent, -3, 15, 15, -3);
      return;
    }

    if (type === "meter") {
      parent.appendChild(svgEl("rect", { class: "symbol-white", x: -17, y: -10, width: 34, height: 20 }));
      addPolyline(parent, [[-10, 0], [-2, -7], [2, 7], [10, 0]]);
      return;
    }

    if (type === "equipment") {
      parent.appendChild(svgEl("rect", { class: "symbol-white", x: -22, y: -13, width: 44, height: 26 }));
      return;
    }

    if (type === "pipe") {
      parent.appendChild(svgEl("circle", { class: "symbol-fill", cx: 0, cy: 0, r: 3.5 }));
      return;
    }

    if (type === "branch") {
      parent.appendChild(svgEl("circle", { class: "symbol-fill", cx: 0, cy: 0, r: 5 }));
      return;
    }

    parent.appendChild(svgEl("circle", { class: "symbol-white", cx: 0, cy: 0, r: 8 }));
  }

  function drawNode(node) {
    var group = svgEl("g", {
      class: "node" + (state.selected && state.selected.type === "node" && state.selected.id === node.id ? " is-selected" : ""),
      transform: "translate(" + numberOr(node.x, 0) + " " + numberOr(node.y, 0) + ")",
      "data-type": "node",
      "data-id": node.id
    });
    drawSymbolShape(group, node.kind);
    if (textOr(node.label, "")) {
      var label = svgEl("text", { class: "diagram-text small", x: -28, y: -28 });
      label.textContent = node.label;
      group.appendChild(label);
    }
    if (textOr(node.diameter, "") && node.kind === "pipe") {
      var diameter = svgEl("text", { class: "diagram-text small", x: -24, y: 24 });
      diameter.textContent = node.diameter;
      group.appendChild(diameter);
    }
    dom.symbols.appendChild(group);
  }

  function drawUserSymbol(symbol) {
    var group = svgEl("g", {
      class: "symbol" + (state.selected && state.selected.type === "symbol" && state.selected.id === symbol.id ? " is-selected" : ""),
      transform: "translate(" + numberOr(symbol.x, 0) + " " + numberOr(symbol.y, 0) + ")",
      "data-type": "symbol",
      "data-id": symbol.id
    });
    drawSymbolShape(group, symbol.kind);
    if (textOr(symbol.label, "")) {
      var label = svgEl("text", { class: "diagram-text small", x: -28, y: -28 });
      label.textContent = symbol.label;
      group.appendChild(label);
    }
    dom.symbols.appendChild(group);
  }

  function drawLabels() {
    state.diagram.labels.forEach(function (label) {
      var group = svgEl("g", {
        class: "diagram-label" + (state.selected && state.selected.type === "label" && state.selected.id === label.id ? " is-selected" : ""),
        transform: "translate(" + numberOr(label.x, 0) + " " + numberOr(label.y, 0) + ")",
        "data-type": "label",
        "data-id": label.id
      });
      var text = svgEl("text", {
        class: "diagram-text " + (label.kind === "title" ? "title" : ""),
        x: 0,
        y: 0
      });
      text.textContent = textOr(label.text, "");
      group.appendChild(text);
      dom.labels.appendChild(group);
    });
  }

  function render() {
    if (!state.diagram) {
      return;
    }
    clearGroup(dom.edges);
    clearGroup(dom.symbols);
    clearGroup(dom.labels);
    applyViewport();
    drawEdges();
    state.diagram.nodes.forEach(drawNode);
    state.diagram.symbols.forEach(drawUserSymbol);
    drawLabels();
    updateInspector();
    updateHeader();
    updateWarnings();
  }

  function updateHeader() {
    dom.systemName.textContent = textOr(state.diagram.systemName, "No system loaded");
  }

  function updateWarnings() {
    var warnings = asArray(state.diagram.warnings).filter(Boolean);
    if (!warnings.length) {
      dom.warnings.hidden = true;
      dom.warnings.textContent = "";
      return;
    }
    dom.warnings.hidden = false;
    dom.warnings.textContent = warnings.join(" ");
  }

  function applyViewport() {
    dom.viewport.setAttribute(
      "transform",
      "translate(" + state.view.panX + " " + state.view.panY + ") scale(" + state.view.zoom + ")"
    );
  }

  function fitView() {
    if (!state.diagram) {
      return;
    }
    var rect = dom.svg.getBoundingClientRect();
    var width = numberOr(state.diagram.canvas.width, 900);
    var height = numberOr(state.diagram.canvas.height, 600);
    var zoom = Math.min((rect.width - 32) / width, (rect.height - 32) / height);
    state.view.zoom = Math.max(MIN_ZOOM, Math.min(MAX_ZOOM, zoom));
    state.view.panX = Math.max(8, (rect.width - width * state.view.zoom) / 2);
    state.view.panY = Math.max(8, (rect.height - height * state.view.zoom) / 2);
    applyViewport();
  }

  function updateInspector() {
    var selected = selectedObject();
    var hasSelection = !!selected;
    var selectedType = state.selected ? state.selected.type : "";

    dom.selectedId.value = hasSelection ? selected.id : "";
    dom.selectedKind.value = hasSelection ? (selected.kind || selectedType) : "junction";
    dom.selectedLabel.value = hasSelection ? (selected.label || selected.text || "") : "";
    dom.selectedDiameter.value = hasSelection ? (selected.diameter || "") : "";
    dom.selectedFlow.value = hasSelection ? (selected.flow || "") : "";
    dom.selectedX.value = hasSelection && selectedType !== "edge" ? Math.round(numberOr(selected.x, 0)) : "";
    dom.selectedY.value = hasSelection && selectedType !== "edge" ? Math.round(numberOr(selected.y, 0)) : "";

    [
      dom.selectedKind,
      dom.selectedLabel,
      dom.selectedDiameter,
      dom.selectedFlow,
      dom.selectedX,
      dom.selectedY
    ].forEach(function (input) {
      input.disabled = !hasSelection;
    });

    dom.selectedKind.disabled = !hasSelection || selectedType === "edge" || selectedType === "label";
    dom.selectedX.disabled = !hasSelection || selectedType === "edge";
    dom.selectedY.disabled = !hasSelection || selectedType === "edge";
    dom.selectedDiameter.disabled = !hasSelection || selectedType === "label";
    dom.selectedFlow.disabled = !hasSelection || selectedType === "label";
  }

  function updateSelectedFromInspector() {
    var selected = selectedObject();
    if (!selected) {
      return;
    }

    if (state.selected.type !== "edge" && state.selected.type !== "label") {
      selected.kind = dom.selectedKind.value;
    }

    if (state.selected.type === "label") {
      selected.text = dom.selectedLabel.value;
    } else {
      selected.label = dom.selectedLabel.value;
      selected.diameter = dom.selectedDiameter.value;
      selected.flow = dom.selectedFlow.value;
      if (selected.flow || selected.diameter) {
        selected.label = selected.label || selected.kind;
      }
      if (state.selected.type === "edge") {
        selected.label = [selected.flow, selected.diameter].filter(Boolean).join(" | ") || dom.selectedLabel.value;
      }
    }

    if (state.selected.type !== "edge") {
      selected.x = snap(numberOr(dom.selectedX.value, selected.x || 0));
      selected.y = snap(numberOr(dom.selectedY.value, selected.y || 0));
    }

    setDirty(true);
    pushSnapshot();
    render();
  }

  function selectObject(type, id) {
    state.selected = type && id ? { type: type, id: id } : null;
    render();
  }

  function objectFromTarget(target) {
    var current = target;
    while (current && current !== dom.svg) {
      if (current.dataset && current.dataset.type && current.dataset.id) {
        return { type: current.dataset.type, id: current.dataset.id };
      }
      current = current.parentNode;
    }
    return null;
  }

  function svgPointFromEvent(evt) {
    var point = dom.svg.createSVGPoint();
    point.x = evt.clientX;
    point.y = evt.clientY;
    return point.matrixTransform(dom.viewport.getScreenCTM().inverse());
  }

  function screenPoint(evt) {
    return { x: evt.clientX, y: evt.clientY };
  }

  function startPointer(evt) {
    var hit = objectFromTarget(evt.target);

    if (state.mode === "pan" || (!hit && evt.button === 1)) {
      state.drag = {
        type: "pan",
        start: screenPoint(evt),
        panX: state.view.panX,
        panY: state.view.panY
      };
      evt.preventDefault();
      return;
    }

    if (!hit) {
      selectObject(null, null);
      return;
    }

    selectObject(hit.type, hit.id);

    if (hit.type === "node" || hit.type === "symbol" || hit.type === "label") {
      var obj = selectedObject();
      var point = svgPointFromEvent(evt);
      state.drag = {
        type: "move",
        id: hit.id,
        itemType: hit.type,
        start: point,
        startX: numberOr(obj.x, 0),
        startY: numberOr(obj.y, 0),
        changed: false
      };
      evt.preventDefault();
    }
  }

  function movePointer(evt) {
    if (!state.drag) {
      return;
    }

    if (state.drag.type === "pan") {
      var current = screenPoint(evt);
      state.view.panX = state.drag.panX + (current.x - state.drag.start.x);
      state.view.panY = state.drag.panY + (current.y - state.drag.start.y);
      applyViewport();
      return;
    }

    if (state.drag.type === "move") {
      var obj = selectedObject();
      if (!obj) {
        return;
      }
      var point = svgPointFromEvent(evt);
      obj.x = snap(state.drag.startX + point.x - state.drag.start.x);
      obj.y = snap(state.drag.startY + point.y - state.drag.start.y);
      state.drag.changed = true;
      setDirty(true);
      render();
    }
  }

  function endPointer() {
    if (state.drag && state.drag.type === "move" && state.drag.changed) {
      pushSnapshot();
    }
    state.drag = null;
  }

  function zoomBy(factor) {
    state.view.zoom = Math.max(MIN_ZOOM, Math.min(MAX_ZOOM, state.view.zoom * factor));
    applyViewport();
  }

  function wheelZoom(evt) {
    evt.preventDefault();
    zoomBy(evt.deltaY < 0 ? 1.12 : 0.88);
  }

  function canvasCenter() {
    var rect = dom.svg.getBoundingClientRect();
    var point = dom.svg.createSVGPoint();
    point.x = rect.left + rect.width / 2;
    point.y = rect.top + rect.height / 2;
    var transformed = point.matrixTransform(dom.viewport.getScreenCTM().inverse());
    return { x: snap(transformed.x), y: snap(transformed.y) };
  }

  function titleCase(value) {
    var text = textOr(value, "Item");
    if (text === "branch") {
      return "Tee";
    }
    if (text === "sizeChange") {
      return "Size Change";
    }
    if (text === "accessory") {
      return "Accessory";
    }
    return text.charAt(0).toUpperCase() + text.slice(1);
  }

  function addSymbol(kind) {
    if (!state.diagram) {
      return;
    }
    var center = canvasCenter();
    var symbol = {
      id: makeId("symbol"),
      kind: kind,
      label: titleCase(kind),
      x: center.x,
      y: center.y,
      diameter: "",
      flow: ""
    };
    state.diagram.symbols.push(symbol);
    state.selected = { type: "symbol", id: symbol.id };
    setDirty(true);
    pushSnapshot();
    render();
  }

  function addLabel() {
    if (!state.diagram) {
      return;
    }
    var center = canvasCenter();
    var label = {
      id: makeId("label"),
      kind: "note",
      text: "Label",
      x: center.x,
      y: center.y
    };
    state.diagram.labels.push(label);
    state.selected = { type: "label", id: label.id };
    setDirty(true);
    pushSnapshot();
    render();
  }

  function deleteSelected() {
    if (!state.diagram || !state.selected) {
      return;
    }
    var type = state.selected.type;
    var id = state.selected.id;

    if (type === "node") {
      state.diagram.nodes = state.diagram.nodes.filter(function (node) { return node.id !== id; });
      state.diagram.edges = state.diagram.edges.filter(function (edge) { return edge.from !== id && edge.to !== id; });
    } else if (type === "edge") {
      state.diagram.edges = state.diagram.edges.filter(function (edge) { return edge.id !== id; });
    } else if (type === "symbol") {
      state.diagram.symbols = state.diagram.symbols.filter(function (symbol) { return symbol.id !== id; });
      state.diagram.edges = state.diagram.edges.filter(function (edge) { return edge.from !== id && edge.to !== id; });
    } else if (type === "label") {
      state.diagram.labels = state.diagram.labels.filter(function (label) { return label.id !== id; });
    }

    state.selected = null;
    setDirty(true);
    pushSnapshot();
    render();
  }

  function setMode(mode) {
    state.mode = mode;
    dom.selectMode.classList.toggle("is-active", mode === "select");
    dom.panMode.classList.toggle("is-active", mode === "pan");
  }

  function requestSelectSystem() {
    if (postWebViewMessage({ type: "selectSystem" })) {
      setStatus("warning", "Waiting for Revit selection...");
    }
  }

  function requestRefresh(forceRegenerate) {
    if (postWebViewMessage({ type: "refreshFromSelection", payload: { forceRegenerate: !!forceRegenerate } })) {
      setStatus("warning", forceRegenerate ? "Regenerating from Revit..." : "Reloading from Revit...");
    } else if (forceRegenerate) {
      state.diagram = fixtureDiagram();
      setDirty(false);
      resetHistory();
      fitView();
      render();
      setStatus("ready", "Standalone preview regenerated.");
    }
  }

  function saveDiagram() {
    if (!state.diagram) {
      return;
    }
    if (postWebViewMessage({ type: "saveDiagram", payload: state.diagram })) {
      setStatus("warning", "Saving to Revit...");
    } else {
      setStatus("ready", "Standalone payload is valid.");
      setDirty(false);
    }
  }

  function loadDiagram(payload) {
    state.diagram = normalizeDiagram(payload);
    state.selected = null;
    setDirty(false);
    resetHistory();
    render();
    fitView();
    setStatus(state.diagram.systemUniqueId ? "ready" : "warning", state.diagram.systemUniqueId ? "Diagram loaded." : "No piping system loaded.");
  }

  function handleSaveResult(result) {
    var status = textOr(result && result.status, "info");
    var message = textOr(result && result.message, "Save finished.");
    if (result && result.payload) {
      state.diagram = normalizeDiagram(result.payload);
      state.selected = null;
      setDirty(false);
      resetHistory();
      render();
    }
    setStatus(status, message);
  }

  function handleRefreshResult(result) {
    var status = textOr(result && result.status, "info");
    var message = textOr(result && result.message, "Refresh finished.");
    if (result && result.payload) {
      state.diagram = normalizeDiagram(result.payload);
      state.selected = null;
      setDirty(false);
      resetHistory();
      render();
      fitView();
    }
    setStatus(status, message);
  }

  function bindDom() {
    dom.svg = byId("diagram");
    dom.viewport = byId("viewport");
    dom.edges = byId("edges");
    dom.symbols = byId("symbols");
    dom.labels = byId("labels");
    dom.systemName = byId("system-name");
    dom.status = byId("status");
    dom.warnings = byId("warnings");
    dom.dirtyPill = byId("dirty-pill");
    dom.selectedId = byId("selected-id");
    dom.selectedKind = byId("selected-kind");
    dom.selectedLabel = byId("selected-label");
    dom.selectedDiameter = byId("selected-diameter");
    dom.selectedFlow = byId("selected-flow");
    dom.selectedX = byId("selected-x");
    dom.selectedY = byId("selected-y");
    dom.undo = byId("undo");
    dom.redo = byId("redo");
    dom.selectMode = byId("select-mode");
    dom.panMode = byId("pan-mode");
  }

  function bindEvents() {
    byId("select-system").addEventListener("click", requestSelectSystem);
    byId("refresh").addEventListener("click", function () { requestRefresh(false); });
    byId("regenerate").addEventListener("click", function () { requestRefresh(true); });
    byId("save").addEventListener("click", saveDiagram);
    byId("close-window").addEventListener("click", function () {
      if (!postWebViewMessage({ type: "closeWindow" })) {
        window.close();
      }
    });

    dom.selectMode.addEventListener("click", function () { setMode("select"); });
    dom.panMode.addEventListener("click", function () { setMode("pan"); });
    byId("zoom-out").addEventListener("click", function () { zoomBy(0.85); });
    byId("zoom-in").addEventListener("click", function () { zoomBy(1.18); });
    byId("fit-view").addEventListener("click", fitView);
    byId("add-tee").addEventListener("click", function () { addSymbol("branch"); });
    byId("add-accessory").addEventListener("click", function () { addSymbol("accessory"); });
    byId("add-equipment").addEventListener("click", function () { addSymbol("equipment"); });
    byId("add-label").addEventListener("click", addLabel);
    byId("delete-selected").addEventListener("click", deleteSelected);
    dom.undo.addEventListener("click", undo);
    dom.redo.addEventListener("click", redo);

    [
      dom.selectedKind,
      dom.selectedLabel,
      dom.selectedDiameter,
      dom.selectedFlow,
      dom.selectedX,
      dom.selectedY
    ].forEach(function (input) {
      input.addEventListener("change", updateSelectedFromInspector);
    });

    dom.svg.addEventListener("pointerdown", startPointer);
    dom.svg.addEventListener("pointermove", movePointer);
    dom.svg.addEventListener("pointerup", endPointer);
    dom.svg.addEventListener("pointerleave", endPointer);
    dom.svg.addEventListener("wheel", wheelZoom, { passive: false });

    window.addEventListener("keydown", function (evt) {
      if ((evt.ctrlKey || evt.metaKey) && evt.key.toLowerCase() === "z") {
        evt.preventDefault();
        undo();
      } else if ((evt.ctrlKey || evt.metaKey) && evt.key.toLowerCase() === "y") {
        evt.preventDefault();
        redo();
      } else if (evt.key === "Delete" || evt.key === "Backspace") {
        if (document.activeElement && ["INPUT", "SELECT"].indexOf(document.activeElement.tagName) >= 0) {
          return;
        }
        evt.preventDefault();
        deleteSelected();
      }
    });

    window.addEventListener("resize", function () {
      applyViewport();
    });
  }

  window.ffePipeOneLine = {
    loadDiagram: loadDiagram,
    setStatus: function (payload) {
      setStatus(payload && payload.status, payload && payload.message);
    },
    handleSaveResult: handleSaveResult,
    handleRefreshResult: handleRefreshResult
  };

  document.addEventListener("DOMContentLoaded", function () {
    bindDom();
    bindEvents();
    setMode("select");
    updateInspector();
    updateHistoryButtons();
    if (!postWebViewMessage({ type: "appReady" })) {
      loadDiagram(fixtureDiagram());
      setStatus("ready", "Standalone preview loaded.");
    }
  });
})();
