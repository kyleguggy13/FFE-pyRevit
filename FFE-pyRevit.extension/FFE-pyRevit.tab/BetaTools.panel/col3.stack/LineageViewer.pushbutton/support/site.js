(function () {
  "use strict";

  var SVG_NS = "http://www.w3.org/2000/svg";

  var NODE_WIDTH = 370;
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
    collapsedNodeIds: {},
    manualNodePositions: {},
    childIdsByNodeId: {},
    visibleNodeIds: {},
    renderEdges: [],
    renderEdgesById: {},
    contentGroup: null,
    transform: { x: 20, y: 20, scale: 1 },
    autoFit: true,
    nodeDrag: {
      active: false,
      nodeId: null,
      didMove: false,
      startClientX: 0,
      startClientY: 0,
      startNodeX: 0,
      startNodeY: 0,
      renderX: 0,
      renderY: 0
    },
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

  function measuredTextLength(textElement, value) {
    var lengthValue = 0;
    try {
      lengthValue = textElement.getComputedTextLength();
    } catch (ignore) {
      lengthValue = 0;
    }
    if (lengthValue > 0) {
      return lengthValue;
    }
    return String(value || "").length * 5.8;
  }

  function fitTextElement(textElement, value, maxWidth) {
    var text = String(value || "");
    var low = 0;
    var high = text.length;
    var best = "";
    var candidate;
    var mid;

    textElement.textContent = text;
    if (!maxWidth || measuredTextLength(textElement, text) <= maxWidth) {
      return;
    }

    while (low <= high) {
      mid = Math.floor((low + high) / 2);
      candidate = text.slice(0, mid).replace(/\s+$/, "") + "...";
      textElement.textContent = candidate;
      if (measuredTextLength(textElement, candidate) <= maxWidth) {
        best = candidate;
        low = mid + 1;
      } else {
        high = mid - 1;
      }
    }

    textElement.textContent = best || "...";
  }

  function addFittedText(parent, x, y, value, maxWidth, attrs) {
    var text = addText(parent, x, y, value, attrs);
    fitTextElement(text, value, maxWidth);
    return text;
  }

  function clamp(value, minValue, maxValue) {
    return Math.max(minValue, Math.min(maxValue, value));
  }

  function isPlaceholderReference(value) {
    var text = String(value || "").trim().toLowerCase();
    return !text || text === "-" || text === "no path available" || text === "unknown";
  }

  function cleanFileName(value) {
    var text = String(value || "").trim();
    var slashIndex;
    if (!text) {
      return "";
    }
    text = text.split("|")[0].trim();
    slashIndex = Math.max(text.lastIndexOf("\\"), text.lastIndexOf("/"));
    if (slashIndex >= 0) {
      text = text.slice(slashIndex + 1);
    }
    return text;
  }

  function nodeFileName(node) {
    var reference = cleanFileName(node.sublabel);
    var label = cleanFileName(node.label);
    if (!isPlaceholderReference(reference)) {
      return reference;
    }
    return label || "Untitled";
  }

  function normalizedLinkStatus(node) {
    var raw = String(node.status || "").trim();
    var compact = raw.toLowerCase().replace(/[^a-z0-9]+/g, "");

    if (node.kind === "host") {
      return "Loaded";
    }
    if (compact === "notfound" || compact === "notlocated" || compact === "missing") {
      return "Not Found";
    }
    if (compact.indexOf("unloaded") >= 0 || compact === "notloaded") {
      return "Unloaded";
    }
    if (node.kind === "revitlink" && (!raw || compact === "unknown")) {
      return "Unloaded";
    }
    if (compact.indexOf("loaded") >= 0 || compact === "linked") {
      return "Loaded";
    }
    if (!raw) {
      return "Unknown";
    }
    return raw;
  }

  function statusBadgeVariant(status) {
    var compact = String(status || "").toLowerCase().replace(/[^a-z0-9]+/g, "");
    if (compact === "loaded") {
      return "text-bg-success";
    }
    if (compact === "unloaded") {
      return "text-bg-secondary";
    }
    if (compact === "notfound") {
      return "text-bg-danger";
    }
    if (compact === "imported") {
      return "text-bg-warning";
    }
    if (compact === "linked") {
      return "text-bg-info";
    }
    return "text-bg-secondary";
  }

  function nodeStyle(kind) {
    return styles[kind] || styles.revitlink;
  }

  function safeClassToken(value) {
    return String(value || "normal").toLowerCase().replace(/[^a-z0-9_-]+/g, "-");
  }

  function edgeClasses(kind, style) {
    return [
      "graph-edge",
      "is-kind-" + safeClassToken(kind || "contains"),
      "is-style-" + safeClassToken(style || "normal")
    ];
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

  function compareNodes(a, b) {
    var left = String(a.kind || "") + "|" + String(a.label || "").toLowerCase();
    var right = String(b.kind || "") + "|" + String(b.label || "").toLowerCase();
    if (left < right) {
      return -1;
    }
    if (left > right) {
      return 1;
    }
    return 0;
  }

  function buildHierarchy(nodes, edges) {
    var nodesById = {};
    var incoming = {};
    var childrenById = {};
    var childSeenByParent = {};

    nodes.forEach(function (node) {
      var depth = Number(node.depth) || 0;
      node.depth = depth;
      nodesById[node.id] = node;
    });

    (edges || []).forEach(function (edge) {
      var fromId = edge.from;
      var toId = edge.to;
      if (!fromId || !toId || fromId === toId || edge.style === "reciprocal") {
        return;
      }
      if (!nodesById[fromId] || !nodesById[toId]) {
        return;
      }
      if (!childrenById[fromId]) {
        childrenById[fromId] = [];
        childSeenByParent[fromId] = {};
      }
      if (!childSeenByParent[fromId][toId]) {
        childrenById[fromId].push(toId);
        childSeenByParent[fromId][toId] = true;
      }
      incoming[toId] = true;
    });

    return {
      nodesById: nodesById,
      incoming: incoming,
      childrenById: childrenById
    };
  }

  function hierarchyRoots(nodes, hierarchy) {
    var roots = [];

    nodes.forEach(function (node) {
      if (node.kind === "host") {
        roots.push(node);
      }
    });

    if (!roots.length) {
      nodes.forEach(function (node) {
        if (!hierarchy.incoming[node.id]) {
          roots.push(node);
        }
      });
    }

    if (!roots.length && nodes.length) {
      roots.push(nodes[0]);
    }

    roots.sort(compareNodes);
    return roots;
  }

  function sortedNodesByDepth(nodes) {
    return nodes.slice().sort(function (a, b) {
      var depthDelta = (Number(a.depth) || 0) - (Number(b.depth) || 0);
      if (depthDelta !== 0) {
        return depthDelta;
      }
      return compareNodes(a, b);
    });
  }

  function computeLayout(nodes, edges) {
    var columns = {};
    var maxDepth = 0;
    var positions = {};
    var maxColumnHeight = 0;
    var hierarchy = buildHierarchy(nodes, edges);
    var nodesById = hierarchy.nodesById;
    var childrenById = hierarchy.childrenById;
    var visited = {};
    var orderedNodes = [];
    var roots = hierarchyRoots(nodes, hierarchy);

    function walk(node, depth) {
      var childIds;
      var childNodes;
      if (!node || visited[node.id]) {
        return;
      }
      visited[node.id] = true;
      node.depth = Number(depth) || 0;
      orderedNodes.push(node);

      childIds = childrenById[node.id] || [];
      childNodes = childIds.map(function (childId) {
        return nodesById[childId];
      }).filter(Boolean);
      childNodes.sort(compareNodes);
      childNodes.forEach(function (childNode) {
        walk(childNode, node.depth + 1);
      });
    }

    roots.forEach(function (node) {
      walk(node, Number(node.depth) || 0);
    });

    sortedNodesByDepth(nodes).forEach(function (node) {
      walk(node, Number(node.depth) || 0);
    });

    orderedNodes.forEach(function (node) {
      var depth = Number(node.depth) || 0;
      maxDepth = Math.max(maxDepth, depth);
      if (!columns[depth]) {
        columns[depth] = [];
      }
      columns[depth].push(node);
    });

    Object.keys(columns).forEach(function (depthKey) {
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

  function pruneCollapsedNodes(nodesById) {
    Object.keys(state.collapsedNodeIds).forEach(function (nodeId) {
      if (!nodesById[nodeId]) {
        delete state.collapsedNodeIds[nodeId];
      }
    });
  }

  function visibleGraph(nodes, edges, hierarchy) {
    var visibleIds = {};
    var roots = hierarchyRoots(nodes, hierarchy);
    var nodesById = hierarchy.nodesById;
    var childrenById = hierarchy.childrenById;

    function walk(node) {
      var childIds;
      var childNodes;
      if (!node || visibleIds[node.id]) {
        return;
      }
      visibleIds[node.id] = true;
      if (state.collapsedNodeIds[node.id]) {
        return;
      }

      childIds = childrenById[node.id] || [];
      childNodes = childIds.map(function (childId) {
        return nodesById[childId];
      }).filter(Boolean);
      childNodes.sort(compareNodes);
      childNodes.forEach(walk);
    }

    roots.forEach(walk);

    return {
      nodeIds: visibleIds,
      nodes: nodes.filter(function (node) {
        return !!visibleIds[node.id];
      }),
      edges: edges.filter(function (edge) {
        return !!(visibleIds[edge.from] && visibleIds[edge.to]);
      })
    };
  }

  function pruneManualNodePositions(nodesById) {
    Object.keys(state.manualNodePositions).forEach(function (nodeId) {
      if (!nodesById[nodeId]) {
        delete state.manualNodePositions[nodeId];
      }
    });
  }

  function applyManualNodePositions(nodes) {
    nodes.forEach(function (node) {
      var manual = state.manualNodePositions[node.id];
      var x;
      var y;
      if (!manual || !state.positions[node.id]) {
        return;
      }
      x = Number(manual.x);
      y = Number(manual.y);
      if (isNaN(x) || isNaN(y)) {
        return;
      }
      state.positions[node.id] = {
        x: x,
        y: y
      };
    });
  }

  function refreshLayoutBounds() {
    var minX = Infinity;
    var minY = Infinity;
    var maxX = -Infinity;
    var maxY = -Infinity;
    var hasPosition = false;

    Object.keys(state.visibleNodeIds || {}).forEach(function (nodeId) {
      var position = state.positions[nodeId];
      if (!position) {
        return;
      }
      hasPosition = true;
      minX = Math.min(minX, position.x);
      minY = Math.min(minY, position.y);
      maxX = Math.max(maxX, position.x + NODE_WIDTH);
      maxY = Math.max(maxY, position.y + NODE_HEIGHT);
    });

    if (!hasPosition) {
      state.layout.bounds = {
        minX: 0,
        minY: 0,
        width: state.layout.width || 0,
        height: state.layout.height || 0
      };
      return;
    }

    state.layout.bounds = {
      minX: minX - MARGIN_X,
      minY: minY - MARGIN_Y,
      width: Math.max(maxX - minX + MARGIN_X * 2, 350),
      height: Math.max(maxY - minY + MARGIN_Y * 2, 350)
    };
    state.layout.width = state.layout.bounds.width;
    state.layout.height = state.layout.bounds.height;
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
    refreshLayoutBounds();
    var size = updateSvgViewport();
    var bounds = state.layout.bounds || {
      minX: 0,
      minY: 0,
      width: state.layout.width,
      height: state.layout.height
    };
    var availableWidth = Math.max(100, size.width - 80);
    var availableHeight = Math.max(100, size.height - 80);
    var scale = Math.min(availableWidth / bounds.width, availableHeight / bounds.height);
    scale = clamp(scale, MIN_SCALE, 1);
    state.transform.scale = scale;
    state.transform.x = (size.width - bounds.width * scale) / 2 - bounds.minX * scale;
    state.transform.y = (size.height - bounds.height * scale) / 2 - bounds.minY * scale;
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

  function edgePairKey(edge) {
    return String(edge.from || "") + ">" + String(edge.to || "");
  }

  function copyEdge(edge) {
    var result = {};
    Object.keys(edge || {}).forEach(function (key) {
      result[key] = edge[key];
    });
    return result;
  }

  function decorateEdges(edges) {
    var groups = {};
    var decorated = [];

    edges.forEach(function (edge) {
      var copy = copyEdge(edge);
      var key = edgePairKey(copy);
      if (!groups[key]) {
        groups[key] = [];
      }
      groups[key].push(copy);
      decorated.push(copy);
    });

    Object.keys(groups).forEach(function (key) {
      var group = groups[key];
      group.forEach(function (edge, index) {
        edge.parallelIndex = index;
        edge.parallelCount = group.length;
        edge.parallelOffset = (index - (group.length - 1) / 2) * 16;
      });
    });

    return decorated;
  }

  function edgePathD(edge, nodesById, positions) {
    var parentNode = nodesById[edge.from];
    var childNode = nodesById[edge.to];
    var parentPos = positions[edge.from];
    var childPos = positions[edge.to];
    if (!parentNode || !childNode || !parentPos || !childPos) {
      return "";
    }

    var x1;
    var x2;
    var offset = Number(edge.parallelOffset) || 0;
    var y1 = parentPos.y + NODE_HEIGHT / 2 + offset;
    var y2 = childPos.y + NODE_HEIGHT / 2 + offset;

    if (parentPos.x + NODE_WIDTH / 2 <= childPos.x + NODE_WIDTH / 2) {
      x1 = parentPos.x + NODE_WIDTH;
      x2 = childPos.x;
    } else {
      x1 = parentPos.x;
      x2 = childPos.x + NODE_WIDTH;
    }

    var dx = Math.max(50, Math.abs(x2 - x1) / 2);
    var c1x = x1 <= x2 ? x1 + dx : x1 - dx;
    var c2x = x1 <= x2 ? x2 - dx : x2 + dx;
    return "M " + x1 + " " + y1 + " C " + c1x + " " + y1 + ", " + c2x + " " + y2 + ", " + x2 + " " + y2;
  }

  function renderEdge(parentGroup, edge, nodesById, positions) {
    var parentNode = nodesById[edge.from];
    var childNode = nodesById[edge.to];
    var pathD = edgePathD(edge, nodesById, positions);
    if (!pathD) {
      return;
    }
    var path = createSvgElement("path", {
      class: edgeClasses(edge.kind, edge.style).join(" "),
      d: pathD,
      fill: "none",
      "data-edge-id": edge.id || "",
      "data-from": edge.from,
      "data-to": edge.to,
      "data-kind": edge.kind || "contains",
      "data-style": edge.style || "normal"
    });

    var title = createSvgElement("title");
    title.textContent = [
      edge.label || "",
      parentNode.label && childNode.label ? parentNode.label + " -> " + childNode.label : "",
      edge.status || ""
    ].filter(Boolean).join(" | ");
    path.appendChild(title);
    parentGroup.appendChild(path);
  }

  function nodeElementById(nodeId) {
    var nodes = document.querySelectorAll(".graph-node");
    var found = null;
    Array.prototype.some.call(nodes, function (nodeElement) {
      if (nodeElement.getAttribute("data-node-id") === nodeId) {
        found = nodeElement;
        return true;
      }
      return false;
    });
    return found;
  }

  function setNodeElementOffset(nodeId, dx, dy) {
    var nodeElement = nodeElementById(nodeId);
    if (!nodeElement) {
      return;
    }
    nodeElement.setAttribute("transform", "translate(" + dx + " " + dy + ")");
  }

  function updateRenderedEdges() {
    var edgeElements = document.querySelectorAll(".graph-edge");
    Array.prototype.forEach.call(edgeElements, function (edgeElement) {
      var edgeId = edgeElement.getAttribute("data-edge-id");
      var edge = state.renderEdgesById[edgeId];
      var pathD = edge ? edgePathD(edge, state.nodesById, state.positions) : "";
      if (pathD) {
        edgeElement.setAttribute("d", pathD);
      }
    });
  }

  function childCount(nodeId) {
    return (state.childIdsByNodeId[nodeId] || []).length;
  }

  function renderNodeExpander(parentGroup, node, position) {
    var count = childCount(node.id);
    var collapsed = !!state.collapsedNodeIds[node.id];
    var cx = position.x + NODE_WIDTH - 25;
    var cy = position.y + 25;
    var group;
    var title;

    if (!count) {
      return;
    }

    group = createSvgElement("g", {
      class: "node-expander" + (collapsed ? " is-collapsed" : ""),
      "data-collapse-node-id": node.id
    });

    title = createSvgElement("title");
    title.textContent = (collapsed ? "Expand" : "Collapse") + " " + count + " child item" + (count === 1 ? "" : "s");
    group.appendChild(title);

    group.appendChild(createSvgElement("circle", {
      cx: cx,
      cy: cy,
      r: "10"
    }));
    group.appendChild(createSvgElement("line", {
      x1: cx - 4,
      y1: cy,
      x2: cx + 4,
      y2: cy
    }));
    if (collapsed) {
      group.appendChild(createSvgElement("line", {
        x1: cx,
        y1: cy - 4,
        x2: cx,
        y2: cy + 4
      }));
    }

    parentGroup.appendChild(group);
  }

  function renderBootstrapBadge(parentGroup, x, y, label) {
    var width = Math.max(62, label.length * 7 + 18);
    var variant = statusBadgeVariant(label);
    var textClass = "node-status-badge-text";
    if (variant === "text-bg-warning" || variant === "text-bg-info") {
      textClass += " is-dark-text";
    }

    parentGroup.appendChild(createSvgElement("rect", {
      class: "node-status-badge badge " + variant,
      x: x,
      y: y,
      rx: "5",
      ry: "5",
      width: width,
      height: "20"
    }));
    addText(parentGroup, x + width / 2, y + 14, label, {
      class: textClass,
      "font-family": "Segoe UI, Arial",
      "font-size": "10",
      "font-weight": "700",
      "text-anchor": "middle"
    });
  }

  function renderNode(parentGroup, node, position) {
    var style = nodeStyle(node.kind);
    var fileName = nodeFileName(node);
    var linkStatus = normalizedLinkStatus(node);
    var hasChildren = childCount(node.id) > 0;
    var titleMaxWidth = hasChildren ? NODE_WIDTH - 60 : NODE_WIDTH - 36;
    var group = createSvgElement("g", {
      class: "graph-node" + (hasChildren ? " has-children" : "") + (state.collapsedNodeIds[node.id] ? " is-collapsed" : ""),
      "data-node-id": node.id,
      "data-render-x": position.x,
      "data-render-y": position.y
    });
    var title = createSvgElement("title");
    title.textContent = [
      fileName || node.label || "",
      style.tag || "",
      "Link Status: " + linkStatus,
      node.sublabel || "",
      node.identityPath ? "Model Path: " + node.identityPath : "",
      node.modelGuid ? "Model GUID: " + node.modelGuid : "",
      node.identitySource || ""
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

    addFittedText(group, position.x + 18, position.y + 30, fileName, titleMaxWidth, {
      "font-family": "Segoe UI, Arial",
      "font-size": "10",
      "font-weight": "700",
      fill: "#2f3438"
    });

    addText(group, position.x + 18, position.y + 52, style.tag, {
      "font-family": "Segoe UI, Arial",
      "font-size": "8",
      "font-weight": "700",
      "text-transform": "uppercase",
      fill: "#676e75"
    });

    renderBootstrapBadge(group, position.x + 18, position.y + 65, linkStatus);

    renderNodeExpander(group, node, position);
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
      if (childCount(nodeId)) {
        classes.push("has-children");
      }
      if (state.collapsedNodeIds[nodeId]) {
        classes.push("is-collapsed");
      }
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
      var classes = edgeClasses(
        edgeElement.getAttribute("data-kind"),
        edgeElement.getAttribute("data-style")
      );
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
      setText("detailIdentity", "-");
      return;
    }

    var style = nodeStyle(node.kind);
    var parentLabel = "-";
    if (node.parentId && state.nodesById[node.parentId]) {
      parentLabel = state.nodesById[node.parentId].label || "-";
    }
    var identityText = "-";
    if (node.identityPath) {
      identityText = node.identityPath;
      if (node.identitySource) {
        identityText += " | " + node.identitySource;
      }
      if (node.modelGuid) {
        identityText += " | GUID: " + node.modelGuid;
      }
    } else if (node.modelGuid) {
      identityText = node.modelGuid;
      if (node.identitySource) {
        identityText += " | " + node.identitySource;
      }
    } else if (node.identitySource || node.identityKey) {
      identityText = [node.identitySource || "", node.identityKey || ""].filter(Boolean).join(" | ");
    }

    setText("detailKind", style.tag);
    setText("detailTitle", node.label || "-");
    setText("detailType", style.tag);
    setText("detailStatus", node.status || "-");
    setText("detailDepth", node.depth);
    setText("detailParent", parentLabel);
    setText("detailReference", node.sublabel || "-");
    setText("detailIdentity", identityText);
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

  function setNodeCollapsed(nodeId, collapsed) {
    if (!nodeId || !childCount(nodeId)) {
      return false;
    }
    if (collapsed) {
      state.collapsedNodeIds[nodeId] = true;
    } else {
      delete state.collapsedNodeIds[nodeId];
    }
    return true;
  }

  function toggleNodeCollapsed(nodeId) {
    if (!setNodeCollapsed(nodeId, !state.collapsedNodeIds[nodeId])) {
      return;
    }
    renderGraph(state.payload || {}, { preserveView: !state.autoFit });
    selectNode(nodeId);
  }

  function renderGraph(payload, options) {
    var svg = $("graphSvg");
    var emptyState = $("emptyState");
    var allNodes = payload.nodes || [];
    var allEdges = payload.edges || [];
    var hierarchy = buildHierarchy(allNodes, allEdges);
    var visible;
    var nodes;
    var edges;
    var renderEdges;
    var contentGroup;
    var edgeGroup;
    var nodeGroup;
    var preserveView = !!(options && options.preserveView);

    clearElement(svg);
    updateSvgViewport();

    state.nodesById = {};
    state.positions = {};
    state.layout = { width: 0, height: 0 };
    state.contentGroup = null;
    state.renderEdges = [];
    state.renderEdgesById = {};
    state.childIdsByNodeId = hierarchy.childrenById;
    pruneCollapsedNodes(hierarchy.nodesById);
    pruneManualNodePositions(hierarchy.nodesById);
    visible = visibleGraph(allNodes, allEdges, hierarchy);
    nodes = visible.nodes;
    edges = visible.edges;
    renderEdges = decorateEdges(edges);
    state.visibleNodeIds = visible.nodeIds;

    if (!allNodes.length) {
      emptyState.hidden = false;
      updateDetails(null);
      return;
    }

    emptyState.hidden = true;

    nodes.forEach(function (node) {
      state.nodesById[node.id] = node;
    });

    state.layout = computeLayout(nodes.slice(), edges);
    state.positions = state.layout.positions;
    applyManualNodePositions(nodes);
    refreshLayoutBounds();

    addDefs(svg);
    contentGroup = createSvgElement("g", { id: "graphContent" });
    edgeGroup = createSvgElement("g", { id: "edgeLayer" });
    nodeGroup = createSvgElement("g", { id: "nodeLayer" });
    contentGroup.appendChild(edgeGroup);
    contentGroup.appendChild(nodeGroup);
    svg.appendChild(contentGroup);
    state.contentGroup = contentGroup;
    state.renderEdges = renderEdges;
    renderEdges.forEach(function (edge) {
      state.renderEdgesById[edge.id] = edge;
    });

    renderEdges.forEach(function (edge) {
      renderEdge(edgeGroup, edge, state.nodesById, state.positions);
    });

    nodes.forEach(function (node) {
      var position = state.positions[node.id];
      if (position) {
        renderNode(nodeGroup, node, position);
      }
    });

    if (!state.selectedNodeId || !state.visibleNodeIds[state.selectedNodeId]) {
      state.selectedNodeId = firstNodeId(nodes);
    }

    if (preserveView) {
      updateSvgViewport();
      applyTransform();
    } else {
      fitGraph();
    }
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

  function collapseNodeIdFromEvent(event) {
    var current = event.target;
    var svg = $("graphSvg");
    while (current && current !== svg) {
      if (current.getAttribute) {
        var nodeId = current.getAttribute("data-collapse-node-id");
        if (nodeId) {
          return nodeId;
        }
      }
      current = current.parentNode;
    }
    return null;
  }

  function numberAttr(element, attrName, fallbackValue) {
    var value = element ? Number(element.getAttribute(attrName)) : NaN;
    return isNaN(value) ? fallbackValue : value;
  }

  function onMouseDown(event) {
    var nodeId;
    var nodeElement;
    var viewport;

    if (event.button !== 0) {
      return;
    }
    if (collapseNodeIdFromEvent(event)) {
      event.preventDefault();
      return;
    }
    nodeId = nodeIdFromEvent(event);
    if (nodeId && state.positions[nodeId]) {
      nodeElement = nodeElementById(nodeId);
      state.nodeDrag.active = true;
      state.nodeDrag.nodeId = nodeId;
      state.nodeDrag.didMove = false;
      state.nodeDrag.startClientX = event.clientX;
      state.nodeDrag.startClientY = event.clientY;
      state.nodeDrag.startNodeX = state.positions[nodeId].x;
      state.nodeDrag.startNodeY = state.positions[nodeId].y;
      state.nodeDrag.renderX = numberAttr(nodeElement, "data-render-x", state.positions[nodeId].x);
      state.nodeDrag.renderY = numberAttr(nodeElement, "data-render-y", state.positions[nodeId].y);
      selectNode(nodeId);
      if (nodeElement) {
        nodeElement.setAttribute("class", nodeElement.getAttribute("class") + " is-dragging");
      }
      viewport = $("graphViewport");
      if (viewport) {
        viewport.classList.add("is-node-dragging");
      }
      event.preventDefault();
      return;
    }

    state.drag.active = true;
    state.drag.startedOnNodeId = null;
    state.drag.didMove = false;
    state.drag.startClientX = event.clientX;
    state.drag.startClientY = event.clientY;
    state.drag.startX = state.transform.x;
    state.drag.startY = state.transform.y;

    viewport = $("graphViewport");
    if (viewport) {
      viewport.classList.add("is-panning");
    }
    event.preventDefault();
  }

  function onMouseMove(event) {
    var nodeId;
    var dx;
    var dy;
    var newX;
    var newY;

    if (state.nodeDrag.active) {
      nodeId = state.nodeDrag.nodeId;
      dx = (event.clientX - state.nodeDrag.startClientX) / state.transform.scale;
      dy = (event.clientY - state.nodeDrag.startClientY) / state.transform.scale;
      if (Math.abs(event.clientX - state.nodeDrag.startClientX) > 3 ||
          Math.abs(event.clientY - state.nodeDrag.startClientY) > 3) {
        state.nodeDrag.didMove = true;
      }
      if (state.nodeDrag.didMove && nodeId) {
        newX = state.nodeDrag.startNodeX + dx;
        newY = state.nodeDrag.startNodeY + dy;
        state.positions[nodeId] = {
          x: newX,
          y: newY
        };
        state.manualNodePositions[nodeId] = {
          x: newX,
          y: newY
        };
        setNodeElementOffset(
          nodeId,
          newX - state.nodeDrag.renderX,
          newY - state.nodeDrag.renderY
        );
        refreshLayoutBounds();
        updateRenderedEdges();
        state.autoFit = false;
      }
      event.preventDefault();
      return;
    }

    if (!state.drag.active) {
      return;
    }
    dx = event.clientX - state.drag.startClientX;
    dy = event.clientY - state.drag.startClientY;
    if (Math.abs(dx) > 3 || Math.abs(dy) > 3) {
      state.drag.didMove = true;
    }
    state.transform.x = state.drag.startX + dx;
    state.transform.y = state.drag.startY + dy;
    state.autoFit = false;
    applyTransform();
  }

  function onMouseUp(event) {
    var nodeId = state.nodeDrag.nodeId;
    var wasNodeClick = state.nodeDrag.active && !state.nodeDrag.didMove;
    var wasNodeMove = state.nodeDrag.active && state.nodeDrag.didMove;
    var startedOnNodeId = state.drag.startedOnNodeId;
    var wasClick = state.drag.active && !state.drag.didMove;
    var viewport = $("graphViewport");
    var nodeElement;

    if (state.nodeDrag.active) {
      state.nodeDrag.active = false;
      state.nodeDrag.nodeId = null;
      if (viewport) {
        viewport.classList.remove("is-node-dragging");
      }
      nodeElement = nodeElementById(nodeId);
      if (nodeElement) {
        nodeElement.setAttribute(
          "class",
          String(nodeElement.getAttribute("class") || "")
            .replace(/\bis-dragging\b/g, "")
            .replace(/\s+/g, " ")
            .replace(/^\s+|\s+$/g, "")
        );
      }
      if (wasNodeMove) {
        renderGraph(state.payload || {}, { preserveView: true });
        selectNode(nodeId);
      } else if (wasNodeClick) {
        selectNode(nodeId);
      }
      event.preventDefault();
      return;
    }

    state.drag.active = false;
    state.drag.startedOnNodeId = null;

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

  function onGraphClick(event) {
    var collapseNodeId = collapseNodeIdFromEvent(event);
    if (!collapseNodeId || event.detail > 1) {
      return;
    }
    toggleNodeCollapsed(collapseNodeId);
    event.preventDefault();
  }

  function onWheel(event) {
    var factor = event.deltaY < 0 ? 1.12 : 0.89;
    zoomAt(event.clientX, event.clientY, factor);
    event.preventDefault();
  }

  function onDoubleClick(event) {
    if (collapseNodeIdFromEvent(event)) {
      event.preventDefault();
      return;
    }
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
    if (event.key === "Enter" || event.key === " ") {
      toggleNodeCollapsed(state.selectedNodeId);
      event.preventDefault();
      return;
    }
    if (event.key === "ArrowLeft" && state.selectedNodeId) {
      if (setNodeCollapsed(state.selectedNodeId, true)) {
        renderGraph(state.payload || {}, { preserveView: !state.autoFit });
        selectNode(state.selectedNodeId);
      }
      event.preventDefault();
      return;
    }
    if (event.key === "ArrowRight" && state.selectedNodeId) {
      if (setNodeCollapsed(state.selectedNodeId, false)) {
        renderGraph(state.payload || {}, { preserveView: !state.autoFit });
        selectNode(state.selectedNodeId);
      }
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
      svg.addEventListener("click", onGraphClick);
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
