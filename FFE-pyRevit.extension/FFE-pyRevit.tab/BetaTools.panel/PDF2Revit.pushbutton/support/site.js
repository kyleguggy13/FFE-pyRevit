(function attachPDF2Revit(globalScope) {
  "use strict";

  var state = {
    payload: null,
    analysis: null,
    pointA: null,
    pointB: null,
    allSelected: true
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

  function setText(id, value) {
    var element = byId(id);
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

  function formatPoint(point) {
    if (!point) {
      return "Not set";
    }
    return Math.round(point.x) + ", " + Math.round(point.y);
  }

  function formatNumber(value) {
    var number = Number(value || 0);
    try {
      return number.toLocaleString();
    } catch (ignore) {
      return String(number);
    }
  }

  function getPage() {
    return (state.payload && state.payload.page) || {};
  }

  function getPageWidth() {
    return Number(getPage().width || 1);
  }

  function getPageHeight() {
    return Number(getPage().height || 1);
  }

  function getCanvas() {
    return byId("overlay");
  }

  function resizeOverlay() {
    var image = byId("pdf-preview");
    var canvas = getCanvas();
    var rect;

    if (!image || !canvas) {
      return;
    }

    rect = image.getBoundingClientRect();
    canvas.width = Math.max(1, Math.round(rect.width));
    canvas.height = Math.max(1, Math.round(rect.height));
    drawOverlay();
  }

  function pdfToCanvas(point) {
    var canvas = getCanvas();
    return {
      x: Number(point.x) / getPageWidth() * canvas.width,
      y: Number(point.y) / getPageHeight() * canvas.height
    };
  }

  function canvasToPdf(clientX, clientY) {
    var canvas = getCanvas();
    var rect = canvas.getBoundingClientRect();
    var x = clientX - rect.left;
    var y = clientY - rect.top;

    x = Math.max(0, Math.min(rect.width, x));
    y = Math.max(0, Math.min(rect.height, y));

    return {
      x: x / Math.max(1, rect.width) * getPageWidth(),
      y: y / Math.max(1, rect.height) * getPageHeight()
    };
  }

  function drawLine(ctx, points, color, width) {
    var a;
    var b;

    if (!points || points.length < 2) {
      return;
    }

    a = pdfToCanvas({ x: points[0][0], y: points[0][1] });
    b = pdfToCanvas({ x: points[1][0], y: points[1][1] });

    ctx.beginPath();
    ctx.moveTo(a.x, a.y);
    ctx.lineTo(b.x, b.y);
    ctx.strokeStyle = color;
    ctx.lineWidth = width;
    ctx.lineCap = "round";
    ctx.stroke();
  }

  function drawPoint(ctx, point, color, label) {
    var canvasPoint;

    if (!point) {
      return;
    }

    canvasPoint = pdfToCanvas(point);
    ctx.beginPath();
    ctx.arc(canvasPoint.x, canvasPoint.y, 5, 0, Math.PI * 2);
    ctx.fillStyle = color;
    ctx.fill();
    ctx.font = "700 12px Segoe UI, Arial, sans-serif";
    ctx.fillText(label, canvasPoint.x + 8, canvasPoint.y - 8);
  }

  function drawFloor(ctx, floor) {
    var loop = floor && floor.pdf_loop;
    var point;

    if (!loop || loop.length < 3) {
      return;
    }

    ctx.beginPath();
    for (index = 0; index < loop.length; index += 1) {
      point = pdfToCanvas({ x: loop[index][0], y: loop[index][1] });
      if (index === 0) {
        ctx.moveTo(point.x, point.y);
      } else {
        ctx.lineTo(point.x, point.y);
      }
    }
    ctx.closePath();
    ctx.fillStyle = "rgba(22, 135, 125, 0.12)";
    ctx.strokeStyle = "rgba(22, 135, 125, 0.85)";
    ctx.lineWidth = 2;
    ctx.fill();
    ctx.stroke();
  }

  function drawMarker(ctx, pointArray, color) {
    var point;

    if (!pointArray || pointArray.length < 2) {
      return;
    }

    point = pdfToCanvas({ x: pointArray[0], y: pointArray[1] });
    ctx.beginPath();
    ctx.arc(point.x, point.y, 6, 0, Math.PI * 2);
    ctx.fillStyle = color;
    ctx.fill();
    ctx.strokeStyle = "#ffffff";
    ctx.lineWidth = 2;
    ctx.stroke();
  }

  function drawOverlay() {
    var canvas = getCanvas();
    var ctx;
    var elements;
    var index;

    if (!canvas) {
      return;
    }

    ctx = canvas.getContext("2d");
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    if (state.analysis && state.analysis.elements) {
      elements = state.analysis.elements;

      (elements.floors || []).forEach(function (floor) {
        drawFloor(ctx, floor);
      });

      (elements.walls || []).forEach(function (wall) {
        drawLine(ctx, wall.pdf_points, "rgba(21, 57, 93, 0.95)", 3);
      });

      (elements.doors || []).forEach(function (door) {
        drawMarker(ctx, door.pdf_point, "rgba(169, 119, 19, 0.95)");
      });

      (elements.windows || []).forEach(function (windowItem) {
        drawMarker(ctx, windowItem.pdf_point, "rgba(89, 106, 120, 0.95)");
      });
    }

    if (state.pointA && state.pointB) {
      drawLine(ctx, [[state.pointA.x, state.pointA.y], [state.pointB.x, state.pointB.y]], "rgba(178, 56, 56, 0.95)", 2);
    }

    drawPoint(ctx, state.pointA, "#b23838", "A");
    drawPoint(ctx, state.pointB, "#b23838", "B");
  }

  function setStatus(value) {
    setText("status-pill", value);
  }

  function showMessage(title, body, isError) {
    var panel = byId("message-panel");
    var bodyElement = byId("message-body");

    if (!panel || !bodyElement) {
      return;
    }

    panel.hidden = false;
    panel.classList.toggle("error", Boolean(isError));
    setText("message-title", title);

    if (Array.isArray(body)) {
      clearElement(bodyElement);
      if (!body.length) {
        bodyElement.textContent = "";
        return;
      }
      bodyElement.appendChild(makeList(body));
    } else {
      bodyElement.textContent = text(body);
    }
  }

  function hideMessage() {
    var panel = byId("message-panel");
    if (panel) {
      panel.hidden = true;
      panel.classList.remove("error");
    }
  }

  function makeList(items) {
    var list = document.createElement("ul");
    items.forEach(function (item) {
      var li = document.createElement("li");
      li.textContent = text(item);
      list.appendChild(li);
    });
    return list;
  }

  function populateSelect(id, options) {
    var select = byId(id);

    if (!select) {
      return;
    }

    clearElement(select);
    (options || []).forEach(function (option) {
      var element = document.createElement("option");
      element.value = text(option.id);
      element.textContent = text(option.label);
      if (option.widthFeet !== undefined) {
        element.dataset.widthFeet = text(option.widthFeet);
      }
      if (option.selected) {
        element.selected = true;
      }
      select.appendChild(element);
    });
  }

  function selectedOption(selectId) {
    var select = byId(selectId);
    if (!select || select.selectedIndex < 0) {
      return null;
    }
    return select.options[select.selectedIndex];
  }

  function getSettings() {
    var wallOption = selectedOption("wall-type-select");
    return {
      levelId: byId("level-select").value,
      wallTypeId: byId("wall-type-select").value,
      floorTypeId: byId("floor-type-select").value,
      doorTypeId: byId("door-type-select").value,
      windowTypeId: byId("window-type-select").value,
      wallWidthFeet: Number((wallOption && wallOption.dataset.widthFeet) || 0.5)
    };
  }

  function getCalibration() {
    return {
      pointA: state.pointA,
      pointB: state.pointB,
      distance: Number(byId("known-distance").value || 0),
      unit: byId("distance-unit").value
    };
  }

  function validateCalibration() {
    var calibration = getCalibration();
    if (!calibration.pointA || !calibration.pointB) {
      showMessage("Calibration Needed", "Set Point A and Point B on the PDF preview.", true);
      return false;
    }
    if (!(calibration.distance > 0)) {
      showMessage("Calibration Needed", "Enter a known distance greater than zero.", true);
      return false;
    }
    return true;
  }

  function updatePointLabels() {
    setText("point-a", formatPoint(state.pointA));
    setText("point-b", formatPoint(state.pointB));
    if (!state.pointA && !state.pointB) {
      setStatus("Waiting for calibration");
    } else if (state.pointA && !state.pointB) {
      setStatus("Set Point B");
    } else {
      setStatus("Ready to analyze");
    }
  }

  function onCanvasClick(event) {
    var point = canvasToPdf(event.clientX, event.clientY);

    if (!state.pointA || (state.pointA && state.pointB)) {
      state.pointA = point;
      state.pointB = null;
      state.analysis = null;
      byId("create-elements").disabled = true;
      renderElementList();
      setMetrics();
    } else {
      state.pointB = point;
    }

    hideMessage();
    updatePointLabels();
    drawOverlay();
  }

  function clearCalibration() {
    state.pointA = null;
    state.pointB = null;
    state.analysis = null;
    byId("create-elements").disabled = true;
    updatePointLabels();
    renderElementList();
    setMetrics();
    drawOverlay();
    hideMessage();
  }

  function setMetrics(elements) {
    elements = elements || {};
    setText("metric-walls", formatNumber((elements.walls || []).length));
    setText("metric-floors", formatNumber((elements.floors || []).length));
    setText("metric-doors", formatNumber((elements.doors || []).length));
    setText("metric-windows", formatNumber((elements.windows || []).length));
  }

  function allElements() {
    var elements = (state.analysis && state.analysis.elements) || {};
    var rows = [];

    (elements.walls || []).forEach(function (item) {
      rows.push({ category: "walls", type: "wall", label: item.id, item: item });
    });
    (elements.floors || []).forEach(function (item) {
      rows.push({ category: "floors", type: "floor", label: item.id, item: item });
    });
    (elements.doors || []).forEach(function (item) {
      rows.push({ category: "doors", type: "door", label: item.id, item: item });
    });
    (elements.windows || []).forEach(function (item) {
      rows.push({ category: "windows", type: "window", label: item.id, item: item });
    });

    return rows;
  }

  function elementMeta(row) {
    var item = row.item || {};
    if (row.type === "wall" && item.length_feet !== undefined) {
      return Number(item.length_feet).toFixed(1) + " ft";
    }
    if ((row.type === "door" || row.type === "window") && item.host_wall_id) {
      return "Host " + item.host_wall_id;
    }
    if (row.type === "floor" && item.area_square_feet !== undefined) {
      return Number(item.area_square_feet).toFixed(0) + " sf";
    }
    return "";
  }

  function renderElementList() {
    var container = byId("element-list");
    var rows = allElements();

    if (!container) {
      return;
    }

    clearElement(container);
    if (!rows.length) {
      var empty = document.createElement("p");
      empty.className = "empty-state";
      empty.textContent = "No analysis results yet.";
      container.appendChild(empty);
      return;
    }

    rows.forEach(function (row) {
      var wrapper = document.createElement("label");
      var checkbox = document.createElement("input");
      var titleWrap = document.createElement("div");
      var title = document.createElement("div");
      var meta = document.createElement("div");
      var badge = document.createElement("span");

      wrapper.className = "element-row";
      checkbox.type = "checkbox";
      checkbox.checked = true;
      checkbox.dataset.category = row.category;
      checkbox.dataset.id = row.item.id;

      titleWrap.className = "element-title-wrap";
      title.className = "element-title";
      title.textContent = row.label;
      meta.className = "element-meta";
      meta.textContent = elementMeta(row);
      titleWrap.appendChild(title);
      titleWrap.appendChild(meta);

      badge.className = "badge " + row.type;
      badge.textContent = row.type;

      wrapper.appendChild(checkbox);
      wrapper.appendChild(titleWrap);
      wrapper.appendChild(badge);
      container.appendChild(wrapper);
    });
  }

  function collectAccepted() {
    var accepted = {
      walls: [],
      floors: [],
      doors: [],
      windows: []
    };

    Array.prototype.forEach.call(document.querySelectorAll(".element-row input[type='checkbox']"), function (checkbox) {
      if (checkbox.checked && accepted[checkbox.dataset.category]) {
        accepted[checkbox.dataset.category].push(checkbox.dataset.id);
      }
    });

    return accepted;
  }

  function toggleAll() {
    state.allSelected = !state.allSelected;
    Array.prototype.forEach.call(document.querySelectorAll(".element-row input[type='checkbox']"), function (checkbox) {
      checkbox.checked = state.allSelected;
    });
    setText("toggle-all", state.allSelected ? "Clear All" : "Select All");
  }

  function analyzePDF() {
    if (!validateCalibration()) {
      return;
    }

    hideMessage();
    state.analysis = null;
    byId("create-elements").disabled = true;
    setStatus("Analyzing PDF vectors");
    postWebViewMessage({
      type: "analyze",
      calibration: getCalibration(),
      settings: getSettings()
    });
  }

  function createElements() {
    if (!state.analysis) {
      showMessage("Analysis Needed", "Run analysis before creating Revit elements.", true);
      return;
    }

    hideMessage();
    setStatus("Creating Revit elements");
    byId("create-elements").disabled = true;
    postWebViewMessage({
      type: "create",
      accepted: collectAccepted(),
      settings: getSettings()
    });
  }

  function loadData(payload) {
    var image = byId("pdf-preview");
    var options;

    state.payload = payload || {};
    state.analysis = null;

    setText("pdf-name", state.payload.pdfName || "PDF");
    setText("page-label", "Page " + text(state.payload.pageNumber || "-") + " of " + text(state.payload.pageCount || "-"));

    options = state.payload.options || {};
    populateSelect("level-select", options.levels);
    populateSelect("wall-type-select", options.wallTypes);
    populateSelect("floor-type-select", options.floorTypes);
    populateSelect("door-type-select", options.doorTypes);
    populateSelect("window-type-select", options.windowTypes);

    image.onload = resizeOverlay;
    image.src = state.payload.previewUri || "";
    updatePointLabels();
    setMetrics();
    renderElementList();
    setText("toggle-all", "Clear All");
    state.allSelected = true;
  }

  function loadAnalysis(result) {
    var elements = (result && result.elements) || {};
    var warnings = (result && result.warnings) || [];

    state.analysis = result || null;
    setMetrics(elements);
    renderElementList();
    state.allSelected = true;
    setText("toggle-all", "Clear All");
    drawOverlay();
    byId("create-elements").disabled = !allElements().length;
    setStatus("Analysis complete");

    if (warnings.length) {
      showMessage("Analysis Warnings", warnings, false);
    } else {
      showMessage("Analysis Complete", "Review the detected elements before creating Revit geometry.", false);
    }
  }

  function loadCreateResult(result) {
    var created = (result && result.created) || {};
    var lines = [
      "Walls: " + formatNumber(created.walls || 0),
      "Floors: " + formatNumber(created.floors || 0),
      "Doors: " + formatNumber(created.doors || 0),
      "Windows: " + formatNumber(created.windows || 0)
    ];
    var warnings = (result && result.warnings) || [];

    if (warnings.length) {
      lines = lines.concat(warnings);
    }

    setStatus("Creation complete");
    showMessage("Created Revit Elements", lines, false);
    byId("create-elements").disabled = false;
  }

  function showError(message) {
    setStatus("Needs attention");
    showMessage("PDF2Revit Error", message, true);
    byId("create-elements").disabled = !state.analysis;
  }

  function attachEvents() {
    byId("overlay").addEventListener("click", onCanvasClick);
    byId("clear-calibration").addEventListener("click", clearCalibration);
    byId("analyze-pdf").addEventListener("click", analyzePDF);
    byId("create-elements").addEventListener("click", createElements);
    byId("toggle-all").addEventListener("click", toggleAll);
    byId("close-window").addEventListener("click", function () {
      postWebViewMessage({ type: "closeWindow" });
    });
    globalScope.addEventListener("resize", resizeOverlay);
  }

  globalScope.pdf2revit = {
    loadData: loadData,
    loadAnalysis: loadAnalysis,
    loadCreateResult: loadCreateResult,
    showError: showError
  };

  document.addEventListener("DOMContentLoaded", function () {
    attachEvents();
    postWebViewMessage({ type: "appReady" });
  });
})(window);
