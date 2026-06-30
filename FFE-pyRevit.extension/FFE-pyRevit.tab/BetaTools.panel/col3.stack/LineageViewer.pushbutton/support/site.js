(function () {
  "use strict";

  var SVG_NS = "http://www.w3.org/2000/svg";

  var NODE_WIDTH = 400;
  var NODE_HEIGHT = 96;
  var X_GAP = 130;
  var Y_GAP = 24;
  var MARGIN_X = 50;
  var MARGIN_Y = 40;
  var MIN_SCALE = 0.18;
  var MAX_SCALE = 2.5;

  var styles = {
    host: {
      icon: "R",
      border: "#3b8d82",
      iconFill: "#eaf7f4",
      iconStroke: "#3b8d82",
      tag: "Active Model"
    },
    revitlink: {
      icon: "R",
      border: "#7e9cb2",
      iconFill: "#f3f7fa",
      iconStroke: "#7e9cb2",
      tag: "Revit Link"
    },
    cadlink: {
      icon: "C",
      border: "#9b8db9",
      iconFill: "#f5f2fb",
      iconStroke: "#9b8db9",
      tag: "CAD Link"
    },
    cadimport: {
      icon: "C",
      border: "#b69561",
      iconFill: "#fbf6ec",
      iconStroke: "#b69561",
      tag: "CAD Import"
    },
    image: {
      icon: "I",
      border: "#c68957",
      iconFill: "#fcf2ea",
      iconStroke: "#c68957",
      tag: "Image"
    },
    pdf: {
      icon: "P",
      border: "#d08a57",
      iconFill: "#fcf3ec",
      iconStroke: "#d08a57",
      tag: "PDF"
    },
    pointcloud: {
      icon: "P",
      border: "#6e9fb2",
      iconFill: "#eef7fa",
      iconStroke: "#6e9fb2",
      tag: "Point Cloud"
    }
  };

  var state = {
    payload: null,
    nodesById: {},
    positions: {},
    layout: { width: 0, height: 0 },
    selectedNodeId: null,
    contentGroup: null,
    transform: { x: 20, y: 20, scale: 1 },
    autoFit: true,
    drag: {
      active: false,
      startedOnNodeId: null,
      didMove: false,
      startClientX: 0,
      startClientY: 0,
      startX: 0,
      startY: 0
    }
  };

  function $(id) {
    return document.getElementById(id);
  }

  function hasHostBridge() {
    return !!(
      window.chrome &&
      window.chrome.webview &&
      typeof window.chrome.webview.postMessage === "function"
    );
  }

  function postHost(message) {
    if (!hasHostBridge()) {
      return;
    }
    window.chrome.webview.postMessage(JSON.stringify(message));
  }

  function clearElement(element) {
    while (element.firstChild) {
      element.removeChild(element.firstChild);
    }
  }

  function createSvgElement(tagName, attrs) {
    var element = document.createElementNS(SVG_NS, tagName);
    Object.keys(attrs || {}).forEach(function (key) {
      element.setAttribute(key, attrs[key]);
    });
    return element;
  }

  function addText(parent, x, y, value, attrs) {
    var text = createSvgElement("text", attrs || {});
    text.setAttribute("x", x);
    text.setAttribute("y", y);
    text.textContent = value || "";
    parent.appendChild(text);
    return text;
  }

  function clamp(value, minValue, maxValue) {
    return Math.max(minValue, Math.min(maxValue, value));
  }

  function shorten(value, maxLength) {
    var text = value == null ? "" : String(value);
    if (text.length <= maxLength) {
      return text;
    }
    return text.slice(0, Math.max(0, maxLength - 3)) + "...";
  }

  function nodeStyle(kind) {
    return styles[kind] || styles.revitlink;
  }

  function countValue(counts, key) {
    if (!counts || counts[key] == null) {
      return 0;
    }
    return counts[key];
  }

  function setText(id, value) {
    var element = $(id);
    if (element) {
      element.textContent = value == null || value === "" ? "-" : String(value);
    }
  }

  function renderMeta(model) {
    model = model || {};
    var title = model.title || "Untitled model";
    var path = model.path || "Unsaved model";
    var generated = model.generated || "";
    var note = model.note || "";
    var pieces = [
      "Model: " + title,
      "Path: " + path
    ];
    if (generated) {
      pieces.push("Generated: " + generated);
    }
    if (note) {
      pieces.push(note);
    }
    var element = $("modelMeta");
    if (element) {
      element.textContent = pieces.join(" | ");
    }
  }

  function renderCounts(counts) {
    setText("countRevit", countValue(counts, "revitlink"));
    setText("countCadLink", countValue(counts, "cadlink"));
    setText("countCadImport", countValue(counts, "cadimport"));
    setText("countImage", countValue(counts, "image"));
    setText("countPdf", countValue(counts, "pdf"));
    setText("countPointCloud", countValue(counts, "pointcloud"));
  }

  function computeLayout(nodes) {
    var columns = {};
    var maxDepth = 0;
    var positions = {};
    var maxColumnHeight = 0;

    nodes.forEach(function (node) {
      var depth = Number(node.depth) || 0;
      node.depth = depth;
      maxDepth = Math.max(maxDepth, depth);
      if (!columns[depth]) {
        columns[depth] = [];
      }
      columns[depth].push(node);
    });

    Object.keys(columns).forEach(function (depthKey) {
      columns[depthKey].sort(function (a, b) {
        var left = String(a.kind || "") + "|" + String(a.label || "").toLowerCase();
        var right = String(b.kind || "") + "|" + String(b.label || "").toLowerCase();
        if (left < right) {
          return -1;
        }
        if (left > right) {
          return 1;
        }
        return 0;
      });

      var columnHeight = columns[depthKey].length * NODE_HEIGHT +
        Math.max(0, columns[depthKey].length - 1) * Y_GAP;
      maxColumnHeight = Math.max(maxColumnHeight, columnHeight);
    });

    for (var depth = 0; depth <= maxDepth; depth += 1) {
      var columnNodes = columns[depth] || [];
      var x = MARGIN_X + depth * (NODE_WIDTH + X_GAP);
      var currentColumnHeight = columnNodes.length * NODE_HEIGHT +
        Math.max(0, columnNodes.length - 1) * Y_GAP;
      var startY = MARGIN_Y + Math.max(0, (maxColumnHeight - currentColumnHeight) / 2);

      columnNodes.forEach(function (node, index) {
        positions[node.id] = {
          x: x,
          y: startY + index * (NODE_HEIGHT + Y_GAP)
        };
      });
    }

    return {
      positions: positions,
      width: MARGIN_X * 2 + (maxDepth + 1) * NODE_WIDTH + maxDepth * X_GAP,
      height: Math.max(maxColumnHeight + MARGIN_Y * 2, 350)
    };
  }

  function viewportSize() {
    var viewport = $("graphViewport");
    return {
      width: Math.max(320, viewport ? viewport.clientWidth : 900),
      height: Math.max(320, viewport ? viewport.clientHeight : 520)
    };
  }

  function updateSvgViewport() {
    var svg = $("graphSvg");
    var size = viewportSize();
    if (!svg) {
      return size;
    }
    svg.setAttribute("viewBox", "0 0 " + size.width + " " + size.height);
    return size;
  }

  function applyTransform() {
    if (!state.contentGroup) {
      return;
    }
    state.contentGroup.setAttribute(
      "transform",
      "translate(" + state.transform.x + " " + state.transform.y + ") scale(" + state.transform.scale + ")"
    );
  }

  function fitGraph() {
    if (!state.layout.width || !state.layout.height) {
      return;
    }
    var size = updateSvgViewport();
    var availableWidth = Math.max(100, size.width - 80);
    var availableHeight = Math.max(100, size.height - 80);
    var scale = Math.min(availableWidth / state.layout.width, availableHeight / state.layout.height);
    scale = clamp(scale, MIN_SCALE, 1);
    state.transform.scale = scale;
    state.transform.x = (size.width - state.layout.width * scale) / 2;
    state.transform.y = (size.height - state.layout.height * scale) / 2;
    state.autoFit = true;
    applyTransform();
  }

  function resetGraph() {
    state.transform.scale = 1;
    state.transform.x = 20;
    state.transform.y = 20;
    state.autoFit = false;
    updateSvgViewport();
    applyTransform();
  }

  function zoomAt(clientX, clientY, scaleFactor) {
    var svg = $("graphSvg");
    if (!svg || !state.layout.width) {
      return;
    }
    var rect = svg.getBoundingClientRect();
    var pointX = clientX - rect.left;
    var pointY = clientY - rect.top;
    var oldScale = state.transform.scale;
    var newScale = clamp(oldScale * scaleFactor, MIN_SCALE, MAX_SCALE);
    var worldX = (pointX - state.transform.x) / oldScale;
    var worldY = (pointY - state.transform.y) / oldScale;

    state.transform.scale = newScale;
    state.transform.x = pointX - worldX * newScale;
    state.transform.y = pointY - worldY * newScale;
    state.autoFit = false;
    applyTransform();
  }

  function zoomFromCenter(scaleFactor) {
    var size = viewportSize();
    var svg = $("graphSvg");
    if (!svg) {
      return;
    }
    var rect = svg.getBoundingClientRect();
    zoomAt(rect.left + size.width / 2, rect.top + size.height / 2, scaleFactor);
  }

  function addDefs(svg) {
    var defs = createSvgElement("defs");
    var filter = createSvgElement("filter", {
      id: "nodeShadow",
      x: "-20%",
      y: "-20%",
      width: "160%",
      height: "160%"
    });
    filter.appendChild(createSvgElement("feDropShadow", {
      dx: "0",
      dy: "2",
      stdDeviation: "2.4",
      "flood-color": "#000000",
      "flood-opacity": "0.12"
    }));
    defs.appendChild(filter);
    svg.appendChild(defs);
  }

  function renderEdge(parentGroup, edge, nodesById, positions) {
    var parentNode = nodesById[edge.from];
    var childNode = nodesById[edge.to];
    var parentPos = positions[edge.from];
    var childPos = positions[edge.to];
    if (!parentNode || !childNode || !parentPos || !childPos) {
      return;
    }

    var x1;
    var x2;
    var y1 = parentPos.y + NODE_HEIGHT / 2;
    var y2 = childPos.y + NODE_HEIGHT / 2;

    if (parentNode.depth <= childNode.depth) {
      x1 = parentPos.x + NODE_WIDTH;
      x2 = childPos.x;
    } else {
      x1 = parentPos.x;
      x2 = childPos.x + NODE_WIDTH;
    }

    var dx = Math.max(50, Math.abs(x2 - x1) / 2);
    var c1x = x1 <= x2 ? x1 + dx : x1 - dx;
    var c2x = x1 <= x2 ? x2 - dx : x2 + dx;
    var pathD = "M " + x1 + " " + y1 + " C " + c1x + " " + y1 + ", " + c2x + " " + y2 + ", " + x2 + " " + y2;

    parentGroup.appendChild(createSvgElement("path", {
      class: "graph-edge",
      d: pathD,
      fill: "none",
      stroke: "#c4cad1",
      "stroke-width": "2.2",
      "data-from": edge.from,
      "data-to": edge.to
    }));
  }

  function renderNode(parentGroup, node, position) {
    var style = nodeStyle(node.kind);
    var group = createSvgElement("g", {
      class: "graph-node",
      "data-node-id": node.id
    });
    var title = createSvgElement("title");
    title.textContent = [
      node.label || "",
      style.tag || "",
      node.sublabel || "",
      node.status || ""
    ].filter(Boolean).join(" | ");
    group.appendChild(title);

    var shadowGroup = createSvgElement("g", {
      filter: "url(#nodeShadow)"
    });
    shadowGroup.appendChild(createSvgElement("rect", {
      class: "node-box",
      x: position.x,
      y: position.y,
      rx: "8",
      ry: "8",
      width: NODE_WIDTH,
      height: NODE_HEIGHT,
      fill: "#ffffff",
      stroke: style.border,
      "stroke-width": "1.8"
    }));
    group.appendChild(shadowGroup);

    group.appendChild(createSvgElement("rect", {
      x: position.x + 14,
      y: position.y + 14,
      rx: "6",
      ry: "6",
      width: "26",
      height: "26",
      fill: style.iconFill,
      stroke: style.iconStroke,
      "stroke-width": "1.2"
    }));

    addText(group, position.x + 27, position.y + 32, style.icon, {
      "font-family": "Segoe UI, Arial",
      "font-size": "12",
      "font-weight": "700",
      "text-anchor": "middle",
      fill: style.iconStroke
    });

    addText(group, position.x + 52, position.y + 28, shorten(node.label, 42), {
      "font-family": "Segoe UI, Arial",
      "font-size": "12",
      "font-weight": "700",
      fill: "#2f3438"
    });

    addText(group, position.x + 52, position.y + 50, style.tag, {
      "font-family": "Segoe UI, Arial",
      "font-size": "10",
      fill: "#676e75"
    });

    addText(group, position.x + 52, position.y + 71, shorten(node.sublabel, 48), {
      "font-family": "Segoe UI, Arial",
      "font-size": "10",
      fill: "#80878e"
    });

    if (node.status) {
      addText(group, position.x + 52, position.y + 88, shorten(node.status, 50), {
        "font-family": "Segoe UI, Arial",
        "font-size": "10",
        fill: "#7b8288"
      });
    }

    parentGroup.appendChild(group);
  }

  function relatedNodeIds(selectedNodeId, edges) {
    var ids = {};
    if (!selectedNodeId) {
      return ids;
    }
    ids[selectedNodeId] = true;
    edges.forEach(function (edge) {
      if (edge.from === selectedNodeId) {
        ids[edge.to] = true;
      }
      if (edge.to === selectedNodeId) {
        ids[edge.from] = true;
      }
    });
    return ids;
  }

  function updateSelectionClasses() {
    var payload = state.payload || {};
    var edges = payload.edges || [];
    var selectedNodeId = state.selectedNodeId;
    var relatedIds = relatedNodeIds(selectedNodeId, edges);
    var nodes = document.querySelectorAll(".graph-node");
    var edgeElements = document.querySelectorAll(".graph-edge");

    Array.prototype.forEach.call(nodes, function (nodeElement) {
      var nodeId = nodeElement.getAttribute("data-node-id");
      var classes = ["graph-node"];
      if (selectedNodeId && nodeId === selectedNodeId) {
        classes.push("is-selected");
      } else if (selectedNodeId && relatedIds[nodeId]) {
        classes.push("is-related");
      }
      nodeElement.setAttribute("class", classes.join(" "));
    });

    Array.prototype.forEach.call(edgeElements, function (edgeElement) {
      var fromId = edgeElement.getAttribute("data-from");
      var toId = edgeElement.getAttribute("data-to");
      var classes = ["graph-edge"];
      if (selectedNodeId) {
        if (fromId === selectedNodeId || toId === selectedNodeId) {
          classes.push("is-selected");
        } else {
          classes.push("is-muted");
        }
      }
      edgeElement.setAttribute("class", classes.join(" "));
    });
  }

  function updateDetails(node) {
    if (!node) {
      setText("detailKind", "Selected Item");
      setText("detailTitle", "No selection");
      setText("detailType", "-");
      setText("detailStatus", "-");
      setText("detailDepth", "-");
      setText("detailParent", "-");
      setText("detailReference", "-");
      return;
    }

    var style = nodeStyle(node.kind);
    var parentLabel = "-";
    if (node.parentId && state.nodesById[node.parentId]) {
      parentLabel = state.nodesById[node.parentId].label || "-";
    }

    setText("detailKind", style.tag);
    setText("detailTitle", node.label || "-");
    setText("detailType", style.tag);
    setText("detailStatus", node.status || "-");
    setText("detailDepth", node.depth);
    setText("detailParent", parentLabel);
    setText("detailReference", node.sublabel || "-");
  }

  function selectNode(nodeId) {
    if (!nodeId || !state.nodesById[nodeId]) {
      state.selectedNodeId = null;
      updateDetails(null);
      updateSelectionClasses();
      return;
    }

    state.selectedNodeId = nodeId;
    updateDetails(state.nodesById[nodeId]);
    updateSelectionClasses();
  }

  function centerNode(nodeId) {
    var position = state.positions[nodeId];
    if (!position) {
      return;
    }
    var size = viewportSize();
    state.transform.x = size.width / 2 - (position.x + NODE_WIDTH / 2) * state.transform.scale;
    state.transform.y = size.height / 2 - (position.y + NODE_HEIGHT / 2) * state.transform.scale;
    state.autoFit = false;
    applyTransform();
  }

  function firstNodeId(nodes) {
    var hostId = null;
    if (!nodes.length) {
      return null;
    }
    nodes.forEach(function (node) {
      if (!hostId && node.kind === "host") {
        hostId = node.id;
      }
    });
    return hostId || nodes[0].id;
  }

  function renderGraph(payload) {
    var svg = $("graphSvg");
    var emptyState = $("emptyState");
    var nodes = payload.nodes || [];
    var edges = payload.edges || [];
    var contentGroup;
    var edgeGroup;
    var nodeGroup;

    clearElement(svg);
    updateSvgViewport();

    state.nodesById = {};
    state.positions = {};
    state.layout = { width: 0, height: 0 };
    state.contentGroup = null;

    if (!nodes.length) {
      emptyState.hidden = false;
      updateDetails(null);
      return;
    }

    emptyState.hidden = true;

    nodes.forEach(function (node) {
      state.nodesById[node.id] = node;
    });

    state.layout = computeLayout(nodes.slice());
    state.positions = state.layout.positions;

    addDefs(svg);
    contentGroup = createSvgElement("g", { id: "graphContent" });
    edgeGroup = createSvgElement("g", { id: "edgeLayer" });
    nodeGroup = createSvgElement("g", { id: "nodeLayer" });
    contentGroup.appendChild(edgeGroup);
    contentGroup.appendChild(nodeGroup);
    svg.appendChild(contentGroup);
    state.contentGroup = contentGroup;

    edges.forEach(function (edge) {
      renderEdge(edgeGroup, edge, state.nodesById, state.positions);
    });

    nodes.forEach(function (node) {
      var position = state.positions[node.id];
      if (position) {
        renderNode(nodeGroup, node, position);
      }
    });

    if (!state.selectedNodeId || !state.nodesById[state.selectedNodeId]) {
      state.selectedNodeId = firstNodeId(nodes);
    }

    fitGraph();
    selectNode(state.selectedNodeId);
  }

  function render(payload) {
    payload = payload || {};
    state.payload = payload;
    renderMeta(payload.model || {});
    renderCounts(payload.counts || {});
    renderGraph(payload);
  }

  function nodeIdFromEvent(event) {
    var current = event.target;
    var svg = $("graphSvg");
    while (current && current !== svg) {
      if (current.getAttribute) {
        var nodeId = current.getAttribute("data-node-id");
        if (nodeId) {
          return nodeId;
        }
      }
      current = current.parentNode;
    }
    return null;
  }

  function onMouseDown(event) {
    if (event.button !== 0) {
      return;
    }
    state.drag.active = true;
    state.drag.startedOnNodeId = nodeIdFromEvent(event);
    state.drag.didMove = false;
    state.drag.startClientX = event.clientX;
    state.drag.startClientY = event.clientY;
    state.drag.startX = state.transform.x;
    state.drag.startY = state.transform.y;

    var viewport = $("graphViewport");
    if (viewport) {
      viewport.classList.add("is-panning");
    }
    event.preventDefault();
  }

  function onMouseMove(event) {
    if (!state.drag.active) {
      return;
    }
    var dx = event.clientX - state.drag.startClientX;
    var dy = event.clientY - state.drag.startClientY;
    if (Math.abs(dx) > 3 || Math.abs(dy) > 3) {
      state.drag.didMove = true;
    }
    state.transform.x = state.drag.startX + dx;
    state.transform.y = state.drag.startY + dy;
    state.autoFit = false;
    applyTransform();
  }

  function onMouseUp(event) {
    var startedOnNodeId = state.drag.startedOnNodeId;
    var wasClick = state.drag.active && !state.drag.didMove;
    state.drag.active = false;
    state.drag.startedOnNodeId = null;

    var viewport = $("graphViewport");
    if (viewport) {
      viewport.classList.remove("is-panning");
    }

    if (wasClick) {
      if (startedOnNodeId) {
        selectNode(startedOnNodeId);
      } else if (event.target && event.target.id === "graphSvg") {
        selectNode(null);
      }
    }
  }

  function onWheel(event) {
    var factor = event.deltaY < 0 ? 1.12 : 0.89;
    zoomAt(event.clientX, event.clientY, factor);
    event.preventDefault();
  }

  function onDoubleClick(event) {
    var nodeId = nodeIdFromEvent(event);
    if (nodeId) {
      selectNode(nodeId);
      centerNode(nodeId);
      event.preventDefault();
    }
  }

  function onKeyDown(event) {
    if (event.key === "+" || event.key === "=") {
      zoomFromCenter(1.14);
      event.preventDefault();
      return;
    }
    if (event.key === "-" || event.key === "_") {
      zoomFromCenter(0.88);
      event.preventDefault();
      return;
    }
    if (event.key === "0") {
      resetGraph();
      event.preventDefault();
      return;
    }
    if (event.key === "f" || event.key === "F") {
      fitGraph();
      event.preventDefault();
    }
  }

  function closeWindow() {
    postHost({ type: "closeWindow" });
    if (!hasHostBridge()) {
      window.close();
    }
  }

  function bindControl(id, handler) {
    var element = $(id);
    if (element) {
      element.addEventListener("click", handler);
    }
  }

  function init() {
    var closeButton = $("closeButton");
    var svg = $("graphSvg");
    var viewport = $("graphViewport");

    if (closeButton) {
      closeButton.addEventListener("click", closeWindow);
    }

    bindControl("zoomInButton", function () {
      zoomFromCenter(1.14);
    });
    bindControl("zoomOutButton", function () {
      zoomFromCenter(0.88);
    });
    bindControl("fitButton", fitGraph);
    bindControl("resetButton", resetGraph);

    if (svg) {
      svg.addEventListener("mousedown", onMouseDown);
      svg.addEventListener("wheel", onWheel, { passive: false });
      svg.addEventListener("dblclick", onDoubleClick);
    }
    document.addEventListener("mousemove", onMouseMove);
    document.addEventListener("mouseup", onMouseUp);

    if (viewport) {
      viewport.addEventListener("keydown", onKeyDown);
    }

    window.addEventListener("resize", function () {
      updateSvgViewport();
      if (state.autoFit) {
        fitGraph();
      } else {
        applyTransform();
      }
    });

    postHost({ type: "appReady" });
  }

  window.ffeLineage = {
    loadData: render,
    fit: fitGraph,
    reset: resetGraph
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
}());
