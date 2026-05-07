(function attachCalculator(globalScope) {
  "use strict";

  const FT_TO_M = 0.3048;
  const IN_TO_M = 0.0254;
  const FPM_TO_MPS = FT_TO_M / 60;
  const IN_WC_TO_PA = 249.08891;
  const AIR_DENSITY = 1.204;
  const AIR_VISCOSITY = 1.81e-5;
  const GALVANIZED_STEEL_ROUGHNESS_M = 0.00015;
  const RE_LAMINAR_MAX = 2300;
  const RE_TURBULENT_FLOOR = 4000;
  const LENGTH_100_FT_M = 100 * FT_TO_M;

  const revitBridge = {
    available: false,
    duct: null,
    selectedSize: null,
    pendingDuct: null
  };

  function isFiniteNumber(value) {
    if (value === null || value === undefined || value === "") {
      return false;
    }
    return Number.isFinite(Number(value));
  }

  function cleanNumber(value, digits) {
    if (!isFiniteNumber(value)) {
      return "";
    }

    return Number(Number(value).toFixed(digits)).toString();
  }

  function activateTab(tabName) {
    if (typeof document === "undefined") {
      return;
    }

    var tabButton = document.querySelector('.tab-btn[data-tab="' + tabName + '"]');
    if (tabButton) {
      tabButton.click();
    }
  }

  function getRevitStatusBanner() {
    if (typeof document === "undefined") {
      return null;
    }
    return document.querySelector("#revit-status-banner");
  }

  function setRevitStatus(statusOrState, message) {
    var state = typeof statusOrState === "object"
      ? statusOrState
      : { status: statusOrState || "idle", message: message || "" };
    var banner = getRevitStatusBanner();

    if (banner && state.message) {
      updateStatusBanner(banner, state);
    }
  }

  function postRevitMessage(message) {
    if (
      globalScope.chrome &&
      globalScope.chrome.webview &&
      typeof globalScope.chrome.webview.postMessage === "function"
    ) {
      globalScope.chrome.webview.postMessage(JSON.stringify(message));
      return true;
    }

    setRevitStatus("error", "Revit bridge is not available in this browser.");
    return false;
  }

  function formatMaybe(value, formatter) {
    if (!isFiniteNumber(value)) {
      return "-";
    }
    return formatter(Number(value));
  }

  function describeRevitDuct(duct) {
    if (!duct) {
      return "No Revit duct loaded.";
    }

    var sizeText;
    if (duct.shape === "round") {
      sizeText = formatMaybe(duct.diameterIn, function (value) {
        return formatDimension(value) + " round";
      });
    } else {
      sizeText = formatMaybe(duct.widthIn, function (width) {
        return formatDimension(width) + " x " + formatMaybe(duct.heightIn, formatDimension);
      });
    }

    var bits = [
      "Element " + duct.elementId,
      sizeText
    ];

    if (duct.flowCfm !== null && duct.flowCfm !== undefined && isFiniteNumber(duct.flowCfm)) {
      bits.push(formatFlow(Number(duct.flowCfm)) + " CFM");
    }

    if (duct.systemName) {
      bits.push(duct.systemName);
    }

    return bits.join(" | ");
  }

  function renderRevitDuctSummary() {
    var summary = typeof document !== "undefined"
      ? document.querySelector("#revit-duct-summary")
      : null;
    if (summary) {
      summary.textContent = describeRevitDuct(revitBridge.duct);
    }
  }

  function describeSelectedSize(selectedSize) {
    if (!selectedSize) {
      return "Select a generated same-shape size to apply it to the Revit duct.";
    }

    if (selectedSize.shape === "round") {
      return "Selected " + formatDimension(selectedSize.diameterIn) + " round.";
    }

    return "Selected " +
      formatDimension(selectedSize.widthIn) +
      " x " +
      formatDimension(selectedSize.heightIn) +
      ".";
  }

  function updateRevitApplyUi() {
    if (typeof document === "undefined") {
      return;
    }

    var selectionSummary = document.querySelector("#revit-selection-summary");
    var applyButton = document.querySelector("#revit-apply-button");
    var selectedSize = revitBridge.selectedSize;
    var duct = revitBridge.duct;
    var hasSameShapeSelection = Boolean(
      duct &&
      selectedSize &&
      selectedSize.shape === duct.shape
    );

    if (selectionSummary) {
      selectionSummary.textContent = describeSelectedSize(selectedSize);
    }

    if (applyButton) {
      applyButton.disabled = !hasSameShapeSelection;
    }
  }

  function clearRevitSelection() {
    revitBridge.selectedSize = null;
    if (typeof document !== "undefined") {
      document.querySelectorAll(".is-revit-selected").forEach(function (element) {
        element.classList.remove("is-revit-selected");
      });
    }
    updateRevitApplyUi();
  }

  function selectRevitSize(size, element) {
    if (!revitBridge.duct) {
      setRevitStatus("warning", "This size is selected, but no Revit duct is loaded.");
      return;
    }

    clearRevitSelection();

    revitBridge.selectedSize = Object.assign(
      { elementId: revitBridge.duct.elementId },
      size
    );

    if (element) {
      element.classList.add("is-revit-selected");
    }

    updateRevitApplyUi();

    if (size.shape !== revitBridge.duct.shape) {
      setRevitStatus("error", "Shape changes are not supported in this MVP.");
      return;
    }

    setRevitStatus("ready", describeSelectedSize(revitBridge.selectedSize));
  }

  function applySelectedRevitSize() {
    if (!revitBridge.selectedSize) {
      setRevitStatus("warning", "Select a generated duct size first.");
      return;
    }

    if (!revitBridge.duct) {
      setRevitStatus("error", "No Revit duct is loaded.");
      return;
    }

    if (revitBridge.selectedSize.shape !== revitBridge.duct.shape) {
      setRevitStatus("error", "Shape changes are not supported in this MVP.");
      return;
    }

    setRevitStatus("warning", "Sending selected size to Revit...");
    postRevitMessage({
      type: "applyDuctSize",
      payload: revitBridge.selectedSize
    });
  }

  function writeFieldValue(form, selector, value, digits) {
    var field = form ? form.querySelector(selector) : null;
    if (!field) {
      return;
    }
    field.value = cleanNumber(value, digits);
  }

  function loadDuctIntoCalculator(duct) {
    var form = typeof document !== "undefined"
      ? document.querySelector("#calculator-form")
      : null;
    if (!form || !duct) {
      return;
    }

    var shapeSelector = 'input[name="ductType"][value="' + duct.shape + '"]';
    var shapeInput = form.querySelector(shapeSelector);
    if (shapeInput) {
      shapeInput.checked = true;
    }

    writeFieldValue(form, "#diameter", duct.shape === "round" ? duct.diameterIn : null, 3);
    writeFieldValue(form, "#width", duct.shape === "rectangular" ? duct.widthIn : null, 3);
    writeFieldValue(form, "#height", duct.shape === "rectangular" ? duct.heightIn : null, 3);
    writeFieldValue(form, "#flowRate", duct.flowCfm, 1);
    writeFieldValue(form, "#velocity", null, 1);
    writeFieldValue(form, "#pressureDrop", null, 4);

    if (!form.querySelector("#maxAspectRatio").value.trim()) {
      form.querySelector("#maxAspectRatio").value = "4.0";
    }

    toggleShapeFields(form);
    updateOperatingFieldLocks(form);
    form.dispatchEvent(new Event("input", { bubbles: true }));
  }

  function loadDuctIntoGrid(duct) {
    var form = typeof document !== "undefined"
      ? document.querySelector("#grid-form")
      : null;
    if (!form || !duct) {
      return;
    }

    var shapeInput = form.querySelector('input[name="gridDuctType"][value="' + duct.shape + '"]');
    if (shapeInput) {
      shapeInput.checked = true;
      shapeInput.dispatchEvent(new Event("change", { bubbles: true }));
    }

    writeFieldValue(form, "#gridFlowRate", duct.flowCfm, 1);

    if (isFiniteNumber(duct.frictionInWgPer100Ft)) {
      var pressureMetric = form.querySelector('input[name="gridMetric"][value="pressure_drop"]');
      if (pressureMetric) {
        pressureMetric.checked = true;
        pressureMetric.dispatchEvent(new Event("change", { bubbles: true }));
      }
      writeFieldValue(form, "#gridTargetValue", duct.frictionInWgPer100Ft, 4);
    } else if (isFiniteNumber(duct.velocityFpm)) {
      var velocityMetric = form.querySelector('input[name="gridMetric"][value="velocity"]');
      if (velocityMetric) {
        velocityMetric.checked = true;
        velocityMetric.dispatchEvent(new Event("change", { bubbles: true }));
      }
      writeFieldValue(form, "#gridTargetValue", duct.velocityFpm, 1);
    }
  }

  function showRevitPanel() {
    var panel = typeof document !== "undefined"
      ? document.querySelector("#revit-bridge-panel")
      : null;
    if (panel) {
      panel.hidden = false;
    }
  }

  function loadRevitDuct(duct) {
    revitBridge.available = true;
    revitBridge.duct = duct;
    revitBridge.selectedSize = null;

    if (typeof document === "undefined" || document.readyState === "loading") {
      revitBridge.pendingDuct = duct;
      return;
    }

    showRevitPanel();
    renderRevitDuctSummary();
    loadDuctIntoCalculator(duct);
    loadDuctIntoGrid(duct);
    clearRevitSelection();
    activateTab("ductulator");
    setRevitStatus("ready", "Loaded Revit duct: " + describeRevitDuct(duct));
  }

  function handleResizeResult(result) {
    if (result && result.duct) {
      loadRevitDuct(result.duct);
    }

    if (result) {
      setRevitStatus(result.status || "idle", result.message || "");
    }
  }

  function flushPendingRevitDuct() {
    if (revitBridge.pendingDuct) {
      var pendingDuct = revitBridge.pendingDuct;
      revitBridge.pendingDuct = null;
      loadRevitDuct(pendingDuct);
    }
  }

  function notifyRevitAppReady() {
    postRevitMessage({ type: "appReady" });
  }

  function parseOptionalPositive(rawValue) {
    if (rawValue === null || rawValue === undefined) {
      return { provided: false, valid: true, value: null };
    }

    const trimmed = String(rawValue).trim();
    if (!trimmed) {
      return { provided: false, valid: true, value: null };
    }

    const value = Number(trimmed);
    if (!Number.isFinite(value) || value <= 0) {
      return { provided: true, valid: false, value: null };
    }

    return { provided: true, valid: true, value };
  }

  function parseOptionalAspectRatio(rawValue) {
    const parsed = parseOptionalPositive(rawValue);
    if (parsed.provided && parsed.valid && parsed.value < 1) {
      return { provided: true, valid: false, value: null };
    }
    return parsed;
  }

  function roundAreaSqFt(diameterIn) {
    const diameterFt = diameterIn / 12;
    return (Math.PI * diameterFt * diameterFt) / 4;
  }

  function rectangularAreaSqFt(widthIn, heightIn) {
    return (widthIn * heightIn) / 144;
  }

  function hydraulicDiameterRectangularIn(widthIn, heightIn) {
    return (2 * widthIn * heightIn) / (widthIn + heightIn);
  }

  function rectangularAspectRatio(widthIn, heightIn) {
    return Math.max(widthIn, heightIn) / Math.min(widthIn, heightIn);
  }

  function buildGeometry(inputs) {
    if (inputs.ductType === "round") {
      if (!inputs.diameter.valid) {
        return { error: "Diameter must be greater than zero." };
      }

      if (!inputs.diameter.provided) {
        return { error: "Enter a round-duct diameter." };
      }

      const hydraulicDiameterIn = inputs.diameter.value;
      return {
        areaSqFt: roundAreaSqFt(inputs.diameter.value),
        hydraulicDiameterIn,
        hydraulicDiameterM: hydraulicDiameterIn * IN_TO_M
      };
    }

    if (!inputs.width.valid || !inputs.height.valid) {
      return { error: "Width and height must both be greater than zero." };
    }

    if (!inputs.width.provided || !inputs.height.provided) {
      return { error: "Enter width and height for the rectangular duct." };
    }

    if (inputs.maxAspectRatio.provided && !inputs.maxAspectRatio.valid) {
      return { error: "Maximum aspect ratio must be 1.00 or greater." };
    }

    const aspectRatio = rectangularAspectRatio(inputs.width.value, inputs.height.value);
    if (
      inputs.maxAspectRatio.provided &&
      aspectRatio > inputs.maxAspectRatio.value + 1e-9
    ) {
      return {
        error: `Rectangular aspect ratio ${aspectRatio.toFixed(2)} exceeds the maximum of ${inputs.maxAspectRatio.value.toFixed(2)}.`
      };
    }

    const hydraulicDiameterIn = hydraulicDiameterRectangularIn(
      inputs.width.value,
      inputs.height.value
    );

    return {
      areaSqFt: rectangularAreaSqFt(inputs.width.value, inputs.height.value),
      hydraulicDiameterIn,
      hydraulicDiameterM: hydraulicDiameterIn * IN_TO_M
    };
  }

  function previewGeometry(inputs) {
    if (inputs.ductType === "round") {
      if (inputs.diameter.provided && inputs.diameter.valid) {
        return buildGeometry(inputs);
      }

      return null;
    }

    if (
      inputs.width.provided &&
      inputs.height.provided &&
      inputs.width.valid &&
      inputs.height.valid
    ) {
      return buildGeometry(inputs);
    }

    return null;
  }

  function reynoldsNumber(velocityMps, hydraulicDiameterM) {
    return (AIR_DENSITY * velocityMps * hydraulicDiameterM) / AIR_VISCOSITY;
  }

  function flowRegime(reynolds) {
    if (reynolds < RE_LAMINAR_MAX) {
      return "laminar";
    }
    if (reynolds < RE_TURBULENT_FLOOR) {
      return "transitional";
    }
    return "turbulent";
  }

  function colebrookFrictionFactor(reynolds, hydraulicDiameterM) {
    const relativeRoughness = GALVANIZED_STEEL_ROUGHNESS_M / hydraulicDiameterM;
    let friction = 1 / Math.pow(
      -1.8 * Math.log10(Math.pow(relativeRoughness / 3.7, 1.11) + 6.9 / reynolds),
      2
    );

    for (let index = 0; index < 25; index += 1) {
      const inverseRoot = -2 * Math.log10(
        (relativeRoughness / 3.7) + (2.51 / (reynolds * Math.sqrt(friction)))
      );
      const nextFriction = 1 / (inverseRoot * inverseRoot);

      if (Math.abs(nextFriction - friction) / Math.max(nextFriction, 1e-9) < 1e-8) {
        return nextFriction;
      }

      friction = nextFriction;
    }

    return friction;
  }

  function frictionFactor(reynolds, hydraulicDiameterM) {
    if (reynolds < RE_LAMINAR_MAX) {
      return 64 / reynolds;
    }

    return colebrookFrictionFactor(reynolds, hydraulicDiameterM);
  }

  function pressureDropFromVelocityMps(velocityMps, hydraulicDiameterM) {
    const reynolds = reynoldsNumber(velocityMps, hydraulicDiameterM);
    const friction = frictionFactor(reynolds, hydraulicDiameterM);
    const regime = flowRegime(reynolds);
    const deltaP =
      friction *
      (LENGTH_100_FT_M / hydraulicDiameterM) *
      ((AIR_DENSITY * velocityMps * velocityMps) / 2);

    return {
      pressureDropPa: deltaP,
      pressureDropInWc: deltaP / IN_WC_TO_PA,
      reynolds,
      friction,
      regime
    };
  }

  function computeFromVelocity(velocityFpm, geometry, sourceLabel) {
    const velocityMps = velocityFpm * FPM_TO_MPS;
    const pressure = pressureDropFromVelocityMps(velocityMps, geometry.hydraulicDiameterM);
    const flowRateCfm = velocityFpm * geometry.areaSqFt;

    return {
      pressureDrop: pressure.pressureDropInWc,
      velocity: velocityFpm,
      flowRate: flowRateCfm,
      reynolds: pressure.reynolds,
      frictionFactor: pressure.friction,
      flowRegime: pressure.regime,
      hydraulicDiameterIn: geometry.hydraulicDiameterIn,
      areaSqFt: geometry.areaSqFt,
      solvePath: sourceLabel
    };
  }

  function computeFromFlowRate(flowRateCfm, geometry) {
    const velocityFpm = flowRateCfm / geometry.areaSqFt;
    return computeFromVelocity(
      velocityFpm,
      geometry,
      "Flow rate -> velocity -> pressure drop"
    );
  }

  function solveVelocityFromPressureDrop(pressureDropInWc, geometry) {
    const targetPa = pressureDropInWc * IN_WC_TO_PA;

    let low = 0.01;
    let high = 1;

    while (
      pressureDropFromVelocityMps(high, geometry.hydraulicDiameterM).pressureDropPa < targetPa &&
      high < 150
    ) {
      high *= 2;
    }

    if (high >= 150) {
      return { error: "Pressure-drop input is outside the solver range." };
    }

    for (let index = 0; index < 70; index += 1) {
      const mid = (low + high) / 2;
      const current = pressureDropFromVelocityMps(mid, geometry.hydraulicDiameterM).pressureDropPa;

      if (current < targetPa) {
        low = mid;
      } else {
        high = mid;
      }
    }

    return { value: ((low + high) / 2) / FPM_TO_MPS };
  }

  function computeFromPressureDrop(pressureDropInWc, geometry) {
    const velocity = solveVelocityFromPressureDrop(pressureDropInWc, geometry);
    if (velocity.error) {
      return { error: velocity.error };
    }

    return computeFromVelocity(
      velocity.value,
      geometry,
      "Pressure drop -> velocity -> flow rate"
    );
  }

  function nearlyEqual(a, b) {
    const scale = Math.max(Math.abs(a), Math.abs(b), 1);
    return Math.abs(a - b) / scale <= 0.02;
  }

  function buildConsistencyMessages(inputs, result) {
    const messages = [];

    if (inputs.velocity.provided && !nearlyEqual(inputs.velocity.value, result.velocity)) {
      messages.push(
        "Velocity input does not match the selected duct area or the primary solve path."
      );
    }

    if (inputs.flowRate.provided && !nearlyEqual(inputs.flowRate.value, result.flowRate)) {
      messages.push(
        "Flow-rate input does not match the selected duct area or the primary solve path."
      );
    }

    if (
      inputs.pressureDrop.provided &&
      !nearlyEqual(inputs.pressureDrop.value, result.pressureDrop)
    ) {
      messages.push(
        "Pressure-drop input does not match the Darcy-Weisbach result for the computed velocity."
      );
    }

    return messages;
  }

  function solve(inputs) {
    if (!inputs.pressureDrop.valid || !inputs.velocity.valid || !inputs.flowRate.valid) {
      return {
        status: "error",
        message: "Operating values must be greater than zero when provided."
      };
    }

    const providedCount = [
      inputs.pressureDrop.provided,
      inputs.velocity.provided,
      inputs.flowRate.provided
    ].filter(Boolean).length;

    if (!providedCount) {
      return {
        status: "idle",
        message:
          "Enter geometry and one operating value to solve the remaining fields.",
        geometry: previewGeometry(inputs)
      };
    }

    const geometry = buildGeometry(inputs);
    if (geometry.error) {
      return { status: "error", message: geometry.error };
    }

    let result;
    if (inputs.velocity.provided) {
      result = computeFromVelocity(
        inputs.velocity.value,
        geometry,
        "Velocity -> flow rate -> pressure drop"
      );
    } else if (inputs.flowRate.provided) {
      result = computeFromFlowRate(inputs.flowRate.value, geometry);
    } else {
      result = computeFromPressureDrop(inputs.pressureDrop.value, geometry);
    }

    if (result.error) {
      return { status: "error", message: result.error };
    }

    const consistencyMessages = buildConsistencyMessages(inputs, result);
    const warnings = consistencyMessages.slice();

    if (result.flowRegime === "transitional") {
      warnings.push(
        "Calculated Reynolds number is in the transition zone between laminar and turbulent flow; the calculator uses the turbulent Colebrook branch in this range."
      );
    }

    if (warnings.length) {
      return {
        status: "warning",
        message: warnings.join(" "),
        geometry,
        result
      };
    }

    return {
      status: "ready",
      message: "Solved with Darcy-Weisbach using the current geometry and operating inputs.",
      geometry,
      result
    };
  }

  // ── Solution Grid computation ────────────────────────────────────
  const DUCT_SIZE_MIN_IN = 4;
  const DUCT_SIZE_STEP_IN = 2;

  function ceilDuctSize(valueIn) {
    return Math.max(
      DUCT_SIZE_MIN_IN,
      Math.ceil(valueIn / DUCT_SIZE_STEP_IN) * DUCT_SIZE_STEP_IN
    );
  }

  function rectangularGridStep(sizeIn) {
    const bandIndex = Math.max(0, Math.floor(sizeIn / 100));
    return DUCT_SIZE_STEP_IN * Math.pow(2, bandIndex);
  }

  function generateRectangularGridSizes(maxSizeIn) {
    const maxDimensionIn = ceilDuctSize(maxSizeIn);
    const sizes = [DUCT_SIZE_MIN_IN];

    while (sizes[sizes.length - 1] < maxDimensionIn) {
      const currentSize = sizes[sizes.length - 1];
      let nextSize = currentSize + rectangularGridStep(currentSize);
      const nextBoundary = (Math.floor(currentSize / 100) + 1) * 100;
      if (currentSize < nextBoundary && nextBoundary < nextSize) {
        nextSize = nextBoundary;
      }
      sizes.push(Math.min(nextSize, maxDimensionIn));
    }

    return sizes;
  }

  function generateRoundSolutionGrid(flowRateCfm, primaryMetric, maxDuctSizeIn, minValue, maxValue) {
    if (flowRateCfm <= 0) {
      return { status: "error", message: "Flow rate must be greater than zero." };
    }
    if (primaryMetric !== "velocity" && primaryMetric !== "pressure_drop") {
      return { status: "error", message: "Select either velocity or pressure drop." };
    }
    if (maxDuctSizeIn <= 0) {
      return { status: "error", message: "Maximum duct size must be greater than zero." };
    }
    if (minValue > maxValue) {
      return { status: "error", message: "Minimum bound cannot be greater than maximum bound." };
    }

    const maxDiameterIn = ceilDuctSize(maxDuctSizeIn);
    const rows = [];

    for (let diameterIn = DUCT_SIZE_MIN_IN; diameterIn <= maxDiameterIn; diameterIn += DUCT_SIZE_STEP_IN) {
      const geometry = {
        areaSqFt: roundAreaSqFt(diameterIn),
        hydraulicDiameterIn: diameterIn,
        hydraulicDiameterM: diameterIn * IN_TO_M
      };
      const result = computeFromFlowRate(flowRateCfm, geometry);
      const metricValue = primaryMetric === "velocity" ? result.velocity : result.pressureDrop;

      if (metricValue < minValue || metricValue > maxValue) {
        continue;
      }

      rows.push({
        diameterIn,
        areaSqFt: geometry.areaSqFt,
        velocity: result.velocity,
        pressureDrop: result.pressureDrop,
        reynolds: result.reynolds,
        frictionFactor: result.frictionFactor
      });
    }

    if (!rows.length) {
      const metricName = primaryMetric === "velocity" ? "velocity" : "pressure-drop";
      return {
        status: "error",
        message: `No round duct sizes up to the selected maximum matched the current airflow within the selected ${metricName} bounds.`
      };
    }

    rows.sort((a, b) => a.diameterIn - b.diameterIn);

    const metricLabel = primaryMetric === "velocity" ? "velocity" : "pressure drop per 100'";
    const metricUnits = primaryMetric === "velocity" ? "fpm" : "in. w.c. / 100'";
    const firstDiameter = rows[0].diameterIn;
    const lastDiameter = rows[rows.length - 1].diameterIn;

    return {
      status: "ready",
      message: `Generated ${rows.length} round options from ${firstDiameter} to ${lastDiameter} inches in ${DUCT_SIZE_STEP_IN}-inch increments, showing ${metricLabel} at each diameter, with bounds limited to ${minValue}\u2013${maxValue} ${metricUnits} and a maximum size of ${maxDiameterIn} inches.`,
      rows
    };
  }

  function generateRectangularSolutionGrid(flowRateCfm, primaryMetric, maxAspectRatio, maxDuctSizeIn, minValue, maxValue) {
    if (flowRateCfm <= 0) {
      return { status: "error", message: "Flow rate must be greater than zero." };
    }
    if (primaryMetric !== "velocity" && primaryMetric !== "pressure_drop") {
      return { status: "error", message: "Select either velocity or pressure drop." };
    }
    if (maxDuctSizeIn <= 0) {
      return { status: "error", message: "Maximum duct size must be greater than zero." };
    }
    if (minValue > maxValue) {
      return { status: "error", message: "Minimum bound cannot be greater than maximum bound." };
    }
    if (maxAspectRatio < 1) {
      return { status: "error", message: "Maximum aspect ratio must be 1.00 or greater." };
    }

    const rectangularSizes = generateRectangularGridSizes(maxDuctSizeIn);
    const maxDimensionIn = rectangularSizes[rectangularSizes.length - 1];
    const widthsSet = new Set();
    const heightsSet = new Set();
    const cells = [];

    for (const heightIn of rectangularSizes) {
      const minWidthIn = heightIn / maxAspectRatio;
      const maxWidthIn = Math.min(maxDimensionIn, heightIn * maxAspectRatio);

      for (const widthIn of rectangularSizes) {
        if (widthIn < minWidthIn) {
          continue;
        }
        if (widthIn > maxWidthIn) {
          break;
        }

        const hydraulicDiamIn = hydraulicDiameterRectangularIn(widthIn, heightIn);
        const geometry = {
          areaSqFt: rectangularAreaSqFt(widthIn, heightIn),
          hydraulicDiameterIn: hydraulicDiamIn,
          hydraulicDiameterM: hydraulicDiamIn * IN_TO_M
        };
        const result = computeFromFlowRate(flowRateCfm, geometry);
        const metricValue = primaryMetric === "velocity" ? result.velocity : result.pressureDrop;

        if (metricValue < minValue || metricValue > maxValue) {
          continue;
        }

        widthsSet.add(widthIn);
        heightsSet.add(heightIn);
        cells.push({ widthIn, heightIn, velocity: result.velocity, pressureDrop: result.pressureDrop });
      }
    }

    if (!cells.length) {
      const metricName = primaryMetric === "velocity" ? "velocity" : "pressure-drop";
      return {
        status: "error",
        message: `No rectangular duct sizes matched the current airflow, aspect ratio, and selected ${metricName} bounds.`
      };
    }

    const metricLabel = primaryMetric === "velocity" ? "velocity" : "pressure drop per 100'";
    const metricUnits = primaryMetric === "velocity" ? "fpm" : "in. w.c. / 100'";

    return {
      status: "ready",
      message: `Generated a rectangular solution matrix with ${cells.length} valid cells. Heights run down the left, widths run across the top, cells show ${metricLabel}, and only values between ${minValue} and ${maxValue} ${metricUnits} are shown.`,
      widths: [...widthsSet].sort((a, b) => a - b),
      heights: [...heightsSet].sort((a, b) => a - b),
      cells
    };
  }

  // ── Heat map ─────────────────────────────────────────────────────
  // Blends from white (#ffffff) toward dark green (#0f6b3a) by amount 0-1
  function heatColor(value, target) {
    if (target === null || target <= 0 || !Number.isFinite(target)) {
      return null;
    }
    var distance = Math.abs(value - target) / target;
    var amount = Math.max(0, 1 - distance * 4);   // fades to white beyond 25% away
    if (amount <= 0) { return null; }
    var r = Math.round(255 + (15  - 255) * amount);
    var g = Math.round(255 + (107 - 255) * amount);
    var b = Math.round(255 + (58  - 255) * amount);
    var luminance = 0.299 * r + 0.587 * g + 0.114 * b;
    var textColor = luminance < 140 ? "#ffffff" : "#111111";
    return { bg: "rgb(" + r + "," + g + "," + b + ")", text: textColor };
  }

  // ── Solution Grid rendering ──────────────────────────────────────
  function renderRoundGrid(state, container, primaryMetric, target) {
    container.innerHTML = "";
    if (state.status !== "ready" || !state.rows || !state.rows.length) {
      return;
    }

    var wrapper = document.createElement("div");
    wrapper.className = "table-scroll";

    var table = document.createElement("table");
    table.className = "solution-table";
    table.id = "ductulator-table";

    table.innerHTML = "<thead><tr>" +
      "<th>Diameter (in)</th>" +
      "<th>Area (ft\u00b2)</th>" +
      "<th>Velocity (fpm)</th>" +
      "<th>Pressure Drop (in.\u00a0w.c.\u00a0/\u00a0100\u2019)</th>" +
      "<th>Reynolds</th>" +
      "<th>Friction Factor</th>" +
      "</tr></thead>";

    var tbody = document.createElement("tbody");
    state.rows.forEach(function (row) {
      var tr = document.createElement("tr");
      var heatValue = primaryMetric === "velocity" ? row.velocity : row.pressureDrop;
      var heat = heatColor(heatValue, target);

      tr.className = "solution-row";
      tr.tabIndex = 0;
      tr.setAttribute("role", "button");
      tr.setAttribute("aria-label", "Select " + row.diameterIn + " inch round duct");
      tr.addEventListener("click", function () {
        selectRevitSize({ shape: "round", diameterIn: row.diameterIn }, tr);
      });
      tr.addEventListener("keydown", function (event) {
        if (event.key === "Enter" || event.key === " ") {
          event.preventDefault();
          selectRevitSize({ shape: "round", diameterIn: row.diameterIn }, tr);
        }
      });

      ["diameterIn", "areaSqFt", "velocity", "pressureDrop", "reynolds", "frictionFactor"].forEach(function (key, i) {
        var td = document.createElement("td");
        td.className = "num";
        if (i === 0) { td.textContent = row.diameterIn; }
        else if (i === 1) { td.textContent = formatArea(row.areaSqFt); }
        else if (i === 2) {
          td.textContent = formatVelocity(row.velocity);
          if (primaryMetric === "velocity" && heat) {
            td.style.background = heat.bg;
            td.style.color = heat.text;
            td.style.fontWeight = "700";
          }
        }
        else if (i === 3) {
          td.textContent = formatPressure(row.pressureDrop);
          if (primaryMetric === "pressure_drop" && heat) {
            td.style.background = heat.bg;
            td.style.color = heat.text;
            td.style.fontWeight = "700";
          }
        }
        else if (i === 4) { td.textContent = formatValue(row.reynolds, { maximumFractionDigits: 0 }); }
        else { td.textContent = formatFriction(row.frictionFactor); }
        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    wrapper.appendChild(table);
    container.appendChild(wrapper);
  }

  function renderRectangularGrid(state, container, primaryMetric, target) {
    container.innerHTML = "";
    if (state.status !== "ready" || !state.cells || !state.cells.length) {
      return;
    }

    var { widths, heights, cells } = state;
    var cellMap = new Map();
    cells.forEach(function (cell) {
      cellMap.set(cell.widthIn + "," + cell.heightIn, cell);
    });

    var wrapper = document.createElement("div");
    wrapper.className = "table-scroll";

    var table = document.createElement("table");
    table.className = "solution-table rect-matrix";
    table.id = "ductulator-table";

    var thead = document.createElement("thead");
    var headerRow = document.createElement("tr");
    var cornerTh = document.createElement("th");
    cornerTh.textContent = "H \u2193 \\ W \u2192 (in)";
    cornerTh.className = "corner-cell";
    headerRow.appendChild(cornerTh);
    widths.forEach(function (w) {
      var th = document.createElement("th");
      th.textContent = w;
      headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    var tbody = document.createElement("tbody");
    heights.forEach(function (h) {
      var tr = document.createElement("tr");
      var rowTh = document.createElement("th");
      rowTh.textContent = h;
      tr.appendChild(rowTh);

      widths.forEach(function (w) {
        var td = document.createElement("td");
        var cell = cellMap.get(w + "," + h);
        if (cell) {
          var value = primaryMetric === "velocity" ? cell.velocity : cell.pressureDrop;
          td.textContent = primaryMetric === "velocity" ? formatVelocity(value) : formatPressure(value);
          var heat = heatColor(value, target);
          if (heat) {
            td.style.background = heat.bg;
            td.style.color = heat.text;
            td.style.fontWeight = "700";
          }
          td.className = "num solution-cell";
          td.tabIndex = 0;
          td.setAttribute("role", "button");
          td.setAttribute("aria-label", "Select " + w + " by " + h + " inch rectangular duct");
          td.addEventListener("click", function () {
            selectRevitSize({ shape: "rectangular", widthIn: w, heightIn: h }, td);
          });
          td.addEventListener("keydown", function (event) {
            if (event.key === "Enter" || event.key === " ") {
              event.preventDefault();
              selectRevitSize({ shape: "rectangular", widthIn: w, heightIn: h }, td);
            }
          });
        } else {
          td.textContent = "\u2014";
          td.className = "no-value";
        }
        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    wrapper.appendChild(table);
    container.appendChild(wrapper);
  }

  // ── Export ───────────────────────────────────────────────────────
  var _lastGridState = null;
  var _lastGridMetric = null;

  function exportCSV() {
    if (!_lastGridState || _lastGridState.status !== "ready") { return; }
    var lines = [];
    if (_lastGridState.rows) {
      lines.push("Diameter (in),Area (ft2),Velocity (fpm),Pressure Drop (in.wc/100'),Reynolds,Friction Factor");
      _lastGridState.rows.forEach(function (row) {
        lines.push([
          row.diameterIn,
          row.areaSqFt.toFixed(4),
          row.velocity.toFixed(1),
          row.pressureDrop.toFixed(4),
          Math.round(row.reynolds),
          row.frictionFactor.toFixed(4)
        ].join(","));
      });
    } else if (_lastGridState.cells) {
      var widths = _lastGridState.widths;
      var heights = _lastGridState.heights;
      var cellMap = new Map();
      _lastGridState.cells.forEach(function (c) { cellMap.set(c.widthIn + "," + c.heightIn, c); });
      var unit = _lastGridMetric === "velocity" ? "fpm" : "in.wc/100'";
      lines.push("Height (in)," + widths.map(function (w) { return "W=" + w + " (" + unit + ")"; }).join(","));
      heights.forEach(function (h) {
        var row = [h];
        widths.forEach(function (w) {
          var cell = cellMap.get(w + "," + h);
          if (cell) {
            var val = _lastGridMetric === "velocity" ? cell.velocity : cell.pressureDrop;
            row.push(val.toFixed(_lastGridMetric === "velocity" ? 1 : 4));
          } else { row.push(""); }
        });
        lines.push(row.join(","));
      });
    }
    var blob = new Blob([lines.join("\n")], { type: "text/csv" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = "ductulator.csv";
    a.click();
    URL.revokeObjectURL(url);
  }

  function exportImage() {
    var table = document.getElementById("ductulator-table");
    if (!table || typeof html2canvas === "undefined") { return; }
    var wrapper = table.closest(".table-scroll") || table;
    html2canvas(wrapper, { backgroundColor: "#ffffff", scale: 2 }).then(function (canvas) {
      var a = document.createElement("a");
      a.download = "ductulator.png";
      a.href = canvas.toDataURL("image/png");
      a.click();
    });
  }

  function formatValue(value, options) {
    if (!Number.isFinite(value)) {
      return "-";
    }

    return new Intl.NumberFormat("en-US", options).format(value);
  }

  function formatPressure(value) {
    const options =
      Math.abs(value) >= 1
        ? { maximumFractionDigits: 3, minimumFractionDigits: 3 }
        : { maximumFractionDigits: 4, minimumFractionDigits: 4 };
    return formatValue(value, options);
  }

  function formatVelocity(value) {
    return formatValue(value, { maximumFractionDigits: 1, minimumFractionDigits: 1 });
  }

  function formatFlow(value) {
    return formatValue(value, { maximumFractionDigits: 1, minimumFractionDigits: 1 });
  }

  function formatDimension(value) {
    return `${formatValue(value, {
      maximumFractionDigits: 2,
      minimumFractionDigits: 2
    })} in`;
  }

  function formatArea(value) {
    return `${formatValue(value, {
      maximumFractionDigits: 3,
      minimumFractionDigits: 3
    })} ft2`;
  }

  function formatFriction(value) {
    return formatValue(value, { maximumFractionDigits: 4, minimumFractionDigits: 4 });
  }

  function readInputs(form) {
    const ductType = form.querySelector('input[name="ductType"]:checked').value;

    return {
      ductType,
      diameter: parseOptionalPositive(form.querySelector("#diameter").value),
      width: parseOptionalPositive(form.querySelector("#width").value),
      height: parseOptionalPositive(form.querySelector("#height").value),
      maxAspectRatio: parseOptionalAspectRatio(form.querySelector("#maxAspectRatio").value),
      pressureDrop: parseOptionalPositive(form.querySelector("#pressureDrop").value),
      velocity: parseOptionalPositive(form.querySelector("#velocity").value),
      flowRate: parseOptionalPositive(form.querySelector("#flowRate").value)
    };
  }

  function updateStatusBanner(element, state) {
    element.textContent = state.message;
    element.className = "status-banner";

    if (state.status === "ready") {
      element.classList.add("is-ready");
    } else if (state.status === "warning") {
      element.classList.add("is-warning");
    } else if (state.status === "error") {
      element.classList.add("is-error");
    }
  }

  function updateResults(state, outputs) {
    if (!state.result) {
      outputs.pressure.textContent = "-";
      outputs.velocity.textContent = "-";
      outputs.flow.textContent = "-";
      outputs.reynolds.textContent = "-";
      outputs.hydraulic.textContent = state.geometry
        ? formatDimension(state.geometry.hydraulicDiameterIn)
        : "-";
      outputs.area.textContent = state.geometry ? formatArea(state.geometry.areaSqFt) : "-";
      outputs.friction.textContent = "-";
      outputs.solvePath.textContent = "-";
      return;
    }

    outputs.pressure.textContent = formatPressure(state.result.pressureDrop);
    outputs.velocity.textContent = formatVelocity(state.result.velocity);
    outputs.flow.textContent = formatFlow(state.result.flowRate);
    outputs.reynolds.textContent = formatValue(state.result.reynolds, {
      maximumFractionDigits: 0
    });
    outputs.hydraulic.textContent = formatDimension(state.result.hydraulicDiameterIn);
    outputs.area.textContent = formatArea(state.result.areaSqFt);
    outputs.friction.textContent = formatFriction(state.result.frictionFactor);
    outputs.solvePath.textContent = state.result.solvePath;
  }

  function toggleShapeFields(form) {
    const selectedType = form.querySelector('input[name="ductType"]:checked').value;
    const groups = form.querySelectorAll("[data-shape]");

    groups.forEach((group) => {
      const shouldShow = group.dataset.shape === selectedType;
      group.classList.toggle("is-hidden", !shouldShow);
    });
  }

  function updateOperatingFieldLocks(form) {
    const ductType = form.querySelector('input[name="ductType"]:checked').value;
    const operatingInputs = [
      form.querySelector("#pressureDrop"),
      form.querySelector("#velocity"),
      form.querySelector("#flowRate")
    ];
    const filledCount = operatingInputs.filter((input) => input.value.trim() !== "").length;
    const diameter = parseOptionalPositive(form.querySelector("#diameter").value);
    const width = parseOptionalPositive(form.querySelector("#width").value);
    const height = parseOptionalPositive(form.querySelector("#height").value);
    const geometryReady = ductType === "round"
      ? diameter.provided && diameter.valid
      : width.provided && width.valid && height.provided && height.valid;
    const shouldLockUnusedOperatingFields = geometryReady && filledCount >= 1;

    operatingInputs.forEach((input) => {
      const shouldDisable = shouldLockUnusedOperatingFields && input.value.trim() === "";
      const field = input.closest(".field");
      const group = input.closest(".input-group");

      input.disabled = shouldDisable;
      if (field) {
        field.classList.toggle("field--disabled", shouldDisable);
      }
      if (group) {
        group.classList.toggle("input-group--disabled", shouldDisable);
      }
    });
  }

  function clearCalculatorValuesForShapeChange(form) {
    [
      "#diameter",
      "#width",
      "#height",
      "#pressureDrop",
      "#velocity",
      "#flowRate"
    ].forEach((selector) => {
      form.querySelector(selector).value = "";
    });

    if (!form.querySelector("#maxAspectRatio").value.trim()) {
      form.querySelector("#maxAspectRatio").value = "4.0";
    }
  }

  function loadExample(form) {
    form.querySelector('input[name="ductType"][value="rectangular"]').checked = true;
    form.querySelector("#diameter").value = "";
    form.querySelector("#width").value = "24";
    form.querySelector("#height").value = "12";
    form.querySelector("#maxAspectRatio").value = "4.0";
    form.querySelector("#pressureDrop").value = "";
    form.querySelector("#velocity").value = "900";
    form.querySelector("#flowRate").value = "";
    toggleShapeFields(form);
  }

  function clearForm(form) {
    form.reset();
    form.querySelector('input[name="ductType"][value="rectangular"]').checked = true;
    toggleShapeFields(form);
  }

  function initApp() {
    const form = document.querySelector("#calculator-form");
    if (!form) {
      return;
    }

    const outputs = {
      pressure: document.querySelector("#pressure-result"),
      velocity: document.querySelector("#velocity-result"),
      flow: document.querySelector("#flow-result"),
      reynolds: document.querySelector("#reynolds-result"),
      hydraulic: document.querySelector("#hydraulic-result"),
      area: document.querySelector("#area-result"),
      friction: document.querySelector("#friction-result"),
      solvePath: document.querySelector("#solve-path-result")
    };

    const statusBanner = document.querySelector("#status-banner");
    const render = function render() {
      updateOperatingFieldLocks(form);
      const state = solve(readInputs(form));
      updateStatusBanner(statusBanner, state);
      updateResults(state, outputs);
    };

    toggleShapeFields(form);
    render();

    form.addEventListener("input", render);
    form.addEventListener("change", function handleChange(event) {
      if (event.target.name === "ductType") {
        clearCalculatorValuesForShapeChange(form);
        toggleShapeFields(form);
      }
      render();
    });

    const exampleButton = document.querySelector("#example-button");
    if (exampleButton) {
      exampleButton.addEventListener("click", function handleExample() {
        loadExample(form);
        render();
      });
    }

    document
      .querySelector("#clear-button")
      .addEventListener("click", function handleClear() {
        clearForm(form);
        render();
      });
  }

  function initGridApp() {
    var form = document.querySelector("#grid-form");
    if (!form) { return; }

    var statusBanner    = document.querySelector("#grid-status-banner");
    var resultContainer = document.querySelector("#grid-result-container");
    var aspectRatioGroup = document.querySelector("#gridAspectRatioGroup");
    var minLabel = document.querySelector("#gridMinLabel");
    var maxLabel = document.querySelector("#gridMaxLabel");
    var minUnit  = document.querySelector("#gridMinUnit");
    var maxUnit  = document.querySelector("#gridMaxUnit");
    var minInput = document.querySelector("#gridMinValue");
    var maxInput = document.querySelector("#gridMaxValue");

    function getMetric() {
      return form.querySelector('input[name="gridMetric"]:checked').value;
    }

    function getDuctType() {
      return form.querySelector('input[name="gridDuctType"]:checked').value;
    }

    var exportBar = document.querySelector("#grid-export-bar");
    var targetLabel = document.querySelector("#gridTargetLabel");
    var targetUnit  = document.querySelector("#gridTargetUnit");

    function getMetric() {
      return form.querySelector('input[name="gridMetric"]:checked').value;
    }

    function getDuctType() {
      return form.querySelector('input[name="gridDuctType"]:checked').value;
    }

    function updateLabels() {
      var metric = getMetric();
      var ductType = getDuctType();
      var isVelocity = metric === "velocity";
      var label = isVelocity ? "velocity" : "pressure drop";
      var unit  = isVelocity ? "fpm" : "in.\u00a0w.c.";

      minLabel.textContent    = "Min " + label;
      maxLabel.textContent    = "Max " + label;
      targetLabel.textContent = "Target " + label;
      minUnit.textContent     = unit;
      maxUnit.textContent     = unit;
      targetUnit.textContent  = unit;

      aspectRatioGroup.style.display = ductType === "rectangular" ? "" : "none";
    }

    function runGrid() {
      var flowRate    = parseOptionalPositive(form.querySelector("#gridFlowRate").value);
      var maxDuctSize = parseOptionalPositive(form.querySelector("#gridMaxDuctSize").value);
      var minValue    = parseOptionalPositive(form.querySelector("#gridMinValue").value);
      var maxValue    = parseOptionalPositive(form.querySelector("#gridMaxValue").value);
      var targetValue = parseOptionalPositive(form.querySelector("#gridTargetValue").value);
      var metric      = getMetric();
      var ductType    = getDuctType();
      var target      = (targetValue.provided && targetValue.valid) ? targetValue.value : null;

      if (!flowRate.provided || !flowRate.valid) {
        updateStatusBanner(statusBanner, { status: "error", message: "Enter a valid flow rate." });
        resultContainer.innerHTML = "";
        exportBar.style.display = "none";
        return;
      }
      if (!maxDuctSize.provided || !maxDuctSize.valid) {
        updateStatusBanner(statusBanner, { status: "error", message: "Enter a valid maximum duct size." });
        resultContainer.innerHTML = "";
        exportBar.style.display = "none";
        return;
      }
      if (!minValue.provided || !minValue.valid) {
        updateStatusBanner(statusBanner, { status: "error", message: "Enter a valid minimum bound." });
        resultContainer.innerHTML = "";
        exportBar.style.display = "none";
        return;
      }
      if (!maxValue.provided || !maxValue.valid) {
        updateStatusBanner(statusBanner, { status: "error", message: "Enter a valid maximum bound." });
        resultContainer.innerHTML = "";
        exportBar.style.display = "none";
        return;
      }

      var state;

      if (ductType === "round") {
        state = generateRoundSolutionGrid(flowRate.value, metric, maxDuctSize.value, minValue.value, maxValue.value);
        updateStatusBanner(statusBanner, state);
        renderRoundGrid(state, resultContainer, metric, target);
      } else {
        var maxAspectRatio = parseOptionalPositive(form.querySelector("#gridMaxAspectRatio").value);
        if (!maxAspectRatio.provided || !maxAspectRatio.valid) {
          updateStatusBanner(statusBanner, { status: "error", message: "Enter a valid maximum aspect ratio." });
          resultContainer.innerHTML = "";
          exportBar.style.display = "none";
          return;
        }
        state = generateRectangularSolutionGrid(flowRate.value, metric, maxAspectRatio.value, maxDuctSize.value, minValue.value, maxValue.value);
        updateStatusBanner(statusBanner, state);
        renderRectangularGrid(state, resultContainer, metric, target);
      }

      _lastGridState  = state;
      _lastGridMetric = metric;
      exportBar.style.display = (state.status === "ready") ? "" : "none";
      clearRevitSelection();
    }

    form.addEventListener("change", function (event) {
      if (event.target.name === "gridMetric" || event.target.name === "gridDuctType") {
        updateLabels();
      }
    });

    document.querySelector("#grid-generate-button").addEventListener("click", runGrid);

    document.querySelector("#grid-clear-button").addEventListener("click", function () {
      form.reset();
      form.querySelector('input[name="gridDuctType"][value="rectangular"]').checked = true;
      form.querySelector('input[name="gridMetric"][value="pressure_drop"]').checked = true;
      updateLabels();
      updateStatusBanner(statusBanner, {
        status: "idle",
        message: "Enter a flow rate and bounds above, then click Generate Grid."
      });
      resultContainer.innerHTML = "";
      exportBar.style.display = "none";
      _lastGridState = null;
      clearRevitSelection();
    });

    document.querySelector("#export-csv-button").addEventListener("click", exportCSV);

    var exportImageButton = document.querySelector("#export-image-button");
    if (typeof html2canvas === "undefined") {
      exportImageButton.style.display = "none";
    } else {
      exportImageButton.addEventListener("click", exportImage);
    }

    var revitApplyButton = document.querySelector("#revit-apply-button");
    if (revitApplyButton) {
      revitApplyButton.addEventListener("click", applySelectedRevitSize);
    }

    updateLabels();
    updateRevitApplyUi();
  }

  const api = {
    buildGeometry,
    computeFromFlowRate,
    computeFromPressureDrop,
    computeFromVelocity,
    frictionFactor,
    generateRectangularSolutionGrid,
    generateRoundSolutionGrid,
    hydraulicDiameterRectangularIn,
    pressureDropFromVelocityMps,
    reynoldsNumber,
    solve,
    solveVelocityFromPressureDrop
  };

  if (globalScope) {
    globalScope.ffeRevit = {
      loadDuct: loadRevitDuct,
      setStatus: setRevitStatus,
      handleResizeResult
    };
  }

  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  }

  if (globalScope && typeof globalScope.addEventListener === "function") {
    globalScope.addEventListener("DOMContentLoaded", function () {
      initApp();
      initGridApp();
      flushPendingRevitDuct();
      notifyRevitAppReady();
    });
  }
})(typeof window !== "undefined" ? window : globalThis);
