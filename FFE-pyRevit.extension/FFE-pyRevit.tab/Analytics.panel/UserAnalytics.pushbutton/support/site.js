(function attachAnalyticsApp(globalScope) {
  "use strict";

  var state = {
    payload: null,
    entries: [],
    filteredEntries: [],
    selectedContributionYear: null,
    resizeTimer: null
  };

  var STATUS_TITLES = {
    ready: "Log Loaded",
    missingFolder: "Logs Folder Missing",
    missingFile: "User Log Missing",
    emptyLog: "No Usage Data",
    invalidJson: "Invalid Log File",
    invalidSchema: "Unexpected Log Format",
    readError: "Log Read Error"
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

  function setText(target, value) {
    var element = typeof target === "string" ? byId(target) : target;
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

  function pad2(value) {
    value = Number(value);
    return value < 10 ? "0" + value : String(value);
  }

  function formatNumber(value) {
    var number = Number(value || 0);
    try {
      return number.toLocaleString();
    } catch (ignore) {
      return String(number);
    }
  }

  function parseLocalDate(value) {
    var source = text(value).trim();
    var match;
    var date;

    if (!source) {
      return null;
    }

    match = source.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?/);
    if (match) {
      date = new Date(
        Number(match[1]),
        Number(match[2]) - 1,
        Number(match[3]),
        Number(match[4] || 0),
        Number(match[5] || 0),
        Number(match[6] || 0)
      );
      if (!isNaN(date.getTime())) {
        return date;
      }
    }

    date = new Date(source.replace(" ", "T"));
    if (!isNaN(date.getTime())) {
      return date;
    }

    return null;
  }

  function dateKey(date) {
    if (!date) {
      return "";
    }
    return date.getFullYear() + "-" + pad2(date.getMonth() + 1) + "-" + pad2(date.getDate());
  }

  function localDayNumber(date) {
    return Math.floor(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()) / 86400000);
  }

  function calendarDayDiff(startDate, endDate) {
    return localDayNumber(endDate) - localDayNumber(startDate);
  }

  function monthKey(date) {
    if (!date) {
      return "";
    }
    return date.getFullYear() + "-" + pad2(date.getMonth() + 1);
  }

  function monthLabel(date) {
    var names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    return names[date.getMonth()] + " " + date.getFullYear();
  }

  function shortMonthLabel(date) {
    var names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    return names[date.getMonth()];
  }

  function formatDateTime(date, fallback) {
    if (!date) {
      return fallback || "-";
    }

    try {
      return date.toLocaleString(undefined, {
        year: "numeric",
        month: "short",
        day: "2-digit",
        hour: "numeric",
        minute: "2-digit"
      });
    } catch (ignore) {
      return dateKey(date);
    }
  }

  function formatShortDate(date) {
    if (!date) {
      return "-";
    }

    try {
      return date.toLocaleDateString(undefined, {
        year: "numeric",
        month: "short",
        day: "2-digit"
      });
    } catch (ignore) {
      return dateKey(date);
    }
  }

  function shortPath(path) {
    var value = text(path);
    var parts;
    if (!value) {
      return "";
    }
    parts = value.split(/[\\/]/);
    return parts[parts.length - 1] || value;
  }

  function docLabel(entry) {
    return entry.docTitle || shortPath(entry.docPath) || "(Untitled)";
  }

  function docKey(entry) {
    return entry.docTitle || entry.docPath || "";
  }

  function normalizeEntries(entries) {
    var normalized = [];

    if (!entries || !entries.length) {
      return normalized;
    }

    entries.forEach(function (entry, index) {
      var parsedDate = parseLocalDate(entry.datetime);
      normalized.push({
        index: Number(entry.index === undefined ? index : entry.index),
        datetime: text(entry.datetime),
        date: parsedDate,
        time: parsedDate ? parsedDate.getTime() : 0,
        username: text(entry.username),
        action: text(entry.action) || "(Unknown)",
        status: text(entry.status),
        docTitle: text(entry.doc_title),
        docPath: text(entry.doc_path),
        revitVersion: text(entry.revit_version_number),
        revitBuild: text(entry.revit_build),
        familyName: text(entry.family_name),
        familyOrigin: text(entry.family_origin),
        familyPath: text(entry.family_path)
      });
    });

    return normalized;
  }

  function getBaseDate() {
    var generatedDate = state.payload ? parseLocalDate(state.payload.generatedAt) : null;
    var latestDate = null;

    if (generatedDate) {
      return generatedDate;
    }

    state.entries.forEach(function (entry) {
      if (entry.date && (!latestDate || entry.time > latestDate.getTime())) {
        latestDate = entry.date;
      }
    });

    return latestDate || new Date();
  }

  function passesDateRange(entry, rangeValue) {
    var baseDate;
    var cutoff;
    var endDate;
    var days;

    if (rangeValue === "all") {
      return true;
    }

    if (!entry.date) {
      return false;
    }

    baseDate = getBaseDate();

    if (rangeValue === "currentYear") {
      return entry.date.getFullYear() === baseDate.getFullYear();
    }

    days = Number(rangeValue);
    if (!days) {
      return true;
    }

    cutoff = new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate());
    cutoff.setDate(cutoff.getDate() - days + 1);
    endDate = new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate() + 1);
    return entry.date >= cutoff && entry.date < endDate;
  }

  function filterEntries(entries) {
    var rangeValue = byId("range-filter") ? byId("range-filter").value : "all";
    var actionValue = byId("action-filter") ? byId("action-filter").value : "all";
    var query = byId("document-search") ? byId("document-search").value.toLowerCase().trim() : "";

    return entries.filter(function (entry) {
      var docSearchText;

      if (!passesDateRange(entry, rangeValue)) {
        return false;
      }

      if (actionValue !== "all" && entry.action !== actionValue) {
        return false;
      }

      if (query) {
        docSearchText = (entry.docTitle + " " + entry.docPath).toLowerCase();
        if (docSearchText.indexOf(query) === -1) {
          return false;
        }
      }

      return true;
    });
  }

  function countBy(entries, getLabel) {
    var counts = {};
    var labels = {};
    var rows = [];

    entries.forEach(function (entry) {
      var label = text(getLabel(entry)).trim();
      var key;

      if (!label) {
        return;
      }

      key = label.toLowerCase();
      if (!counts[key]) {
        counts[key] = 0;
        labels[key] = label;
      }
      counts[key] += 1;
    });

    Object.keys(counts).forEach(function (key) {
      rows.push({
        label: labels[key],
        count: counts[key]
      });
    });

    rows.sort(function (a, b) {
      if (b.count !== a.count) {
        return b.count - a.count;
      }
      return a.label.localeCompare(b.label);
    });

    return rows;
  }

  function populateActionFilter() {
    var select = byId("action-filter");
    var previousValue = select ? select.value : "all";
    var rows;
    var hasPreviousValue = previousValue === "all";

    if (!select) {
      return;
    }

    rows = countBy(state.entries, function (entry) {
      return entry.action;
    });

    clearElement(select);
    select.appendChild(new Option("All actions", "all"));

    rows.forEach(function (row) {
      var option = new Option(row.label + " (" + formatNumber(row.count) + ")", row.label);
      if (row.label === previousValue) {
        hasPreviousValue = true;
      }
      select.appendChild(option);
    });

    select.value = hasPreviousValue ? previousValue : "all";
  }

  function renderPayloadMeta() {
    var payload = state.payload || {};
    var generatedDate = parseLocalDate(payload.generatedAt);
    var status = payload.status || "ready";
    var message = payload.message || "";
    var banner = byId("state-banner");
    var isError = status === "invalidJson" || status === "invalidSchema" || status === "readError";

    setText(document.querySelector("[data-username]"), payload.username || "Unknown user");
    setText(document.querySelector("[data-generated-at]"), generatedDate ? "Updated " + formatDateTime(generatedDate) : "Waiting for log data");
    setText("log-path", payload.logPath || "Log path unavailable");

    if (!banner) {
      return;
    }

    if (status !== "ready" || !state.entries.length) {
      banner.hidden = false;
      banner.setAttribute("data-tone", isError ? "error" : "warning");
      setText("state-title", STATUS_TITLES[status] || "Log Status");
      setText("state-message", message || "No usage entries are available.");
    } else {
      banner.hidden = true;
    }
  }

  function renderMetrics(entries) {
    var docs = {};
    var days = {};
    var datedEntries = entries.filter(function (entry) {
      return Boolean(entry.date);
    });
    var syncCount = 0;
    var firstEntry = null;
    var lastEntry = null;

    entries.forEach(function (entry) {
      var key = docKey(entry);

      if (entry.action.toLowerCase() === "sync") {
        syncCount += 1;
      }

      if (key) {
        docs[key.toLowerCase()] = true;
      }

      if (entry.date) {
        days[dateKey(entry.date)] = true;
      }
    });

    datedEntries.sort(function (a, b) {
      return a.time - b.time;
    });

    firstEntry = datedEntries[0] || null;
    lastEntry = datedEntries[datedEntries.length - 1] || null;

    setText("metric-total", formatNumber(entries.length));
    setText("metric-syncs", formatNumber(syncCount));
    setText("metric-documents", formatNumber(Object.keys(docs).length));
    setText("metric-active-days", formatNumber(Object.keys(days).length));
    setText("metric-first", firstEntry ? formatDateTime(firstEntry.date, firstEntry.datetime) : "-");
    setText("metric-last", lastEntry ? formatDateTime(lastEntry.date, lastEntry.datetime) : "-");
  }

  function makeEmptyState(message) {
    var empty = document.createElement("div");
    empty.className = "empty-state";
    empty.textContent = message;
    return empty;
  }

  function renderBarChart(containerId, rows, maxRows, emptyMessage) {
    var container = byId(containerId);
    var visibleRows;
    var maxCount;

    if (!container) {
      return;
    }

    clearElement(container);

    if (!rows.length) {
      container.appendChild(makeEmptyState(emptyMessage));
      return;
    }

    visibleRows = rows.slice(0, maxRows);
    maxCount = visibleRows[0].count || 1;

    visibleRows.forEach(function (row) {
      var item = document.createElement("div");
      var label = document.createElement("div");
      var count = document.createElement("div");
      var track = document.createElement("div");
      var fill = document.createElement("div");

      item.className = "bar-row";
      label.className = "bar-label";
      count.className = "bar-count";
      track.className = "bar-track";
      fill.className = "bar-fill";

      label.title = row.label;
      label.textContent = row.label;
      count.textContent = formatNumber(row.count);
      fill.style.width = Math.max(2, (row.count / maxCount) * 100) + "%";

      track.appendChild(fill);
      item.appendChild(label);
      item.appendChild(count);
      item.appendChild(track);
      container.appendChild(item);
    });
  }

  function buildActivitySeries(entries) {
    var datedEntries = entries.filter(function (entry) {
      return Boolean(entry.date);
    }).sort(function (a, b) {
      return a.time - b.time;
    });
    var first;
    var last;
    var daySpan;
    var useMonths;
    var counts = {};
    var series = [];
    var cursor;
    var end;

    if (!datedEntries.length) {
      return series;
    }

    first = datedEntries[0].date;
    last = datedEntries[datedEntries.length - 1].date;
    daySpan = Math.max(1, Math.round((last.getTime() - first.getTime()) / 86400000) + 1);
    useMonths = daySpan > 120;

    datedEntries.forEach(function (entry) {
      var key = useMonths ? monthKey(entry.date) : dateKey(entry.date);
      counts[key] = (counts[key] || 0) + 1;
    });

    if (useMonths) {
      cursor = new Date(first.getFullYear(), first.getMonth(), 1);
      end = new Date(last.getFullYear(), last.getMonth(), 1);
      while (cursor <= end) {
        series.push({
          key: monthKey(cursor),
          label: monthLabel(cursor),
          count: counts[monthKey(cursor)] || 0
        });
        cursor.setMonth(cursor.getMonth() + 1);
      }
    } else {
      cursor = new Date(first.getFullYear(), first.getMonth(), first.getDate());
      end = new Date(last.getFullYear(), last.getMonth(), last.getDate());
      while (cursor <= end) {
        series.push({
          key: dateKey(cursor),
          label: pad2(cursor.getMonth() + 1) + "/" + pad2(cursor.getDate()),
          count: counts[dateKey(cursor)] || 0
        });
        cursor.setDate(cursor.getDate() + 1);
      }
    }

    return series;
  }

  function setupCanvas(canvas) {
    var dpr = globalScope.devicePixelRatio || 1;
    var width = Math.max(320, canvas.clientWidth || canvas.width || 900);
    var height = Math.max(220, canvas.clientHeight || canvas.height || 260);
    var context;

    canvas.width = Math.floor(width * dpr);
    canvas.height = Math.floor(height * dpr);
    context = canvas.getContext("2d");
    context.setTransform(dpr, 0, 0, dpr, 0, 0);
    context.clearRect(0, 0, width, height);

    return {
      context: context,
      width: width,
      height: height
    };
  }

  function drawNoCanvasData(context, width, height, message) {
    context.fillStyle = "#667386";
    context.font = "700 14px Segoe UI, Arial, sans-serif";
    context.textAlign = "center";
    context.textBaseline = "middle";
    context.fillText(message, width / 2, height / 2);
  }

  function drawActivityChart(entries) {
    var canvas = byId("activity-chart");
    var setup;
    var context;
    var width;
    var height;
    var series;
    var maxCount;
    var padding;
    var chartWidth;
    var chartHeight;
    var baseline;
    var drawBars;

    if (!canvas || !canvas.getContext) {
      return;
    }

    setup = setupCanvas(canvas);
    context = setup.context;
    width = setup.width;
    height = setup.height;
    series = buildActivitySeries(entries);

    if (!series.length) {
      drawNoCanvasData(context, width, height, "No activity in current filter");
      return;
    }

    maxCount = series.reduce(function (maxValue, item) {
      return Math.max(maxValue, item.count);
    }, 1);

    padding = {
      top: 28,
      right: 24,
      bottom: 42,
      left: 52
    };
    chartWidth = Math.max(1, width - padding.left - padding.right);
    chartHeight = Math.max(1, height - padding.top - padding.bottom);
    baseline = padding.top + chartHeight;
    drawBars = series.length <= 65;

    context.strokeStyle = "#dce3ea";
    context.lineWidth = 1;
    context.fillStyle = "#667386";
    context.font = "700 11px Segoe UI, Arial, sans-serif";
    context.textAlign = "right";
    context.textBaseline = "middle";

    [0, 0.25, 0.5, 0.75, 1].forEach(function (ratio) {
      var y = baseline - chartHeight * ratio;
      var label = Math.round(maxCount * ratio);
      context.beginPath();
      context.moveTo(padding.left, y);
      context.lineTo(width - padding.right, y);
      context.stroke();
      context.fillText(label, padding.left - 8, y);
    });

    context.strokeStyle = "#c7d2de";
    context.beginPath();
    context.moveTo(padding.left, padding.top);
    context.lineTo(padding.left, baseline);
    context.lineTo(width - padding.right, baseline);
    context.stroke();

    if (drawBars) {
      series.forEach(function (item, index) {
        var slot = chartWidth / series.length;
        var barWidth = Math.max(3, Math.min(24, slot * 0.68));
        var x = padding.left + slot * index + (slot - barWidth) / 2;
        var barHeight = (item.count / maxCount) * chartHeight;
        var y = baseline - barHeight;

        context.fillStyle = index % 3 === 0 ? "#13857d" : index % 3 === 1 ? "#17395c" : "#b88717";
        context.fillRect(x, y, barWidth, barHeight || 1);
      });
    } else {
      context.strokeStyle = "#13857d";
      context.lineWidth = 3;
      context.beginPath();
      series.forEach(function (item, index) {
        var x = padding.left + (chartWidth * index) / Math.max(1, series.length - 1);
        var y = baseline - (item.count / maxCount) * chartHeight;
        if (index === 0) {
          context.moveTo(x, y);
        } else {
          context.lineTo(x, y);
        }
      });
      context.stroke();
    }

    context.fillStyle = "#667386";
    context.font = "700 11px Segoe UI, Arial, sans-serif";
    context.textBaseline = "top";
    context.textAlign = "left";
    context.fillText(series[0].label, padding.left, baseline + 14);
    context.textAlign = "right";
    context.fillText(series[series.length - 1].label, width - padding.right, baseline + 14);
  }

  function renderCharts(entries) {
    var actionRows = countBy(entries, function (entry) {
      return entry.action;
    });
    var docRows = countBy(entries, function (entry) {
      return docLabel(entry);
    });
    var familyEntries = entries.filter(function (entry) {
      return entry.action.toLowerCase() === "family-loaded" || entry.familyOrigin || entry.familyName;
    });
    var originRows = countBy(familyEntries, function (entry) {
      return entry.familyOrigin || "Unspecified";
    });

    setText("activity-subtitle", formatNumber(entries.length) + " entries");
    setText("actions-subtitle", formatNumber(actionRows.length) + " types");
    setText("docs-subtitle", formatNumber(docRows.length) + " documents");
    setText("origins-subtitle", formatNumber(familyEntries.length) + " loads");

    drawActivityChart(entries);
    renderBarChart("actions-chart", actionRows, 9, "No actions in current filter");
    renderBarChart("docs-chart", docRows, 9, "No documents in current filter");
    renderBarChart("origins-chart", originRows, 9, "No family load origins in current filter");
  }

  function getContributionYears(entries) {
    var years = {};
    var result = [];

    entries.forEach(function (entry) {
      if (entry.date) {
        years[entry.date.getFullYear()] = true;
      }
    });

    Object.keys(years).forEach(function (year) {
      result.push(Number(year));
    });

    result.sort(function (a, b) {
      return b - a;
    });

    return result;
  }

  function getDefaultContributionYear(years) {
    var generatedDate = state.payload ? parseLocalDate(state.payload.generatedAt) : null;
    var generatedYear = generatedDate ? generatedDate.getFullYear() : null;

    if (!years.length) {
      return null;
    }

    if (state.selectedContributionYear && years.indexOf(state.selectedContributionYear) !== -1) {
      return state.selectedContributionYear;
    }

    if (generatedYear && years.indexOf(generatedYear) !== -1) {
      return generatedYear;
    }

    return years[0];
  }

  function contributionText(count) {
    return formatNumber(count) + " contribution" + (count === 1 ? "" : "s");
  }

  function getContributionLevel(count, maxCount) {
    var ratio;

    if (!count) {
      return 0;
    }

    if (maxCount <= 4) {
      return Math.min(4, count);
    }

    ratio = count / maxCount;
    if (ratio <= 0.25) {
      return 1;
    }
    if (ratio <= 0.5) {
      return 2;
    }
    if (ratio <= 0.75) {
      return 3;
    }
    return 4;
  }

  function getContributionDateWindow(year) {
    var start = new Date(year, 0, 1);
    var end = new Date(year, 11, 31);

    start.setDate(start.getDate() - start.getDay());
    end.setDate(end.getDate() + (6 - end.getDay()));

    return {
      start: start,
      end: end,
      weeks: Math.floor(calendarDayDiff(start, end) / 7) + 1
    };
  }

  function buildContributionData(entries, year) {
    var counts = {};
    var total = 0;
    var maxCount = 0;
    var firstActivity = null;
    var lastActivity = null;

    entries.forEach(function (entry) {
      var key;

      if (!entry.date || entry.date.getFullYear() !== year) {
        return;
      }

      key = dateKey(entry.date);
      counts[key] = (counts[key] || 0) + 1;
      total += 1;
      maxCount = Math.max(maxCount, counts[key]);

      if (!firstActivity || entry.time < firstActivity.getTime()) {
        firstActivity = entry.date;
      }
      if (!lastActivity || entry.time > lastActivity.getTime()) {
        lastActivity = entry.date;
      }
    });

    return {
      counts: counts,
      total: total,
      maxCount: maxCount,
      firstActivity: firstActivity,
      lastActivity: lastActivity
    };
  }

  function renderContributionYears(years, selectedYear) {
    var container = byId("contribution-years");

    if (!container) {
      return;
    }

    clearElement(container);

    if (!years.length) {
      var empty = document.createElement("div");
      empty.className = "empty-state";
      empty.textContent = "No years";
      container.appendChild(empty);
      return;
    }

    years.forEach(function (year) {
      var button = document.createElement("button");
      button.type = "button";
      button.className = "contribution-year-button";
      if (year === selectedYear) {
        button.className += " is-active";
      }
      button.textContent = year;
      button.addEventListener("click", function () {
        state.selectedContributionYear = year;
        renderContributionGraph();
      });
      container.appendChild(button);
    });
  }

  function renderContributionMonthLabels(grid, year, dateWindow) {
    var months = document.createElement("div");
    var monthIndex;

    months.className = "contribution-months";
    months.style.gridTemplateColumns = "repeat(" + dateWindow.weeks + ", 12px)";

    for (monthIndex = 0; monthIndex < 12; monthIndex += 1) {
      var monthStart = new Date(year, monthIndex, 1);
      var nextMonthStart = monthIndex === 11 ? new Date(year + 1, 0, 1) : new Date(year, monthIndex + 1, 1);
      var startWeek = Math.floor(calendarDayDiff(dateWindow.start, monthStart) / 7) + 1;
      var endWeek = Math.floor(calendarDayDiff(dateWindow.start, nextMonthStart) / 7) + 1;
      var span = Math.max(1, Math.min(dateWindow.weeks + 1, endWeek) - startWeek);
      var label = document.createElement("span");

      label.className = "contribution-month-label";
      label.style.gridColumn = startWeek + " / span " + span;
      label.textContent = shortMonthLabel(monthStart);
      months.appendChild(label);
    }

    grid.appendChild(months);
  }

  function renderContributionCells(grid, year, dateWindow, contributionData) {
    var body = document.createElement("div");
    var weekdays = document.createElement("div");
    var cells = document.createElement("div");
    var weekdayLabels = ["", "Mon", "", "Wed", "", "Fri", ""];
    var cursor = new Date(dateWindow.start.getFullYear(), dateWindow.start.getMonth(), dateWindow.start.getDate());

    body.className = "contribution-body";
    weekdays.className = "contribution-weekdays";
    cells.className = "contribution-cells";
    cells.style.gridTemplateColumns = "repeat(" + dateWindow.weeks + ", 12px)";

    weekdayLabels.forEach(function (labelText) {
      var label = document.createElement("span");
      label.textContent = labelText;
      weekdays.appendChild(label);
    });

    while (cursor <= dateWindow.end) {
      var key = dateKey(cursor);
      var count = contributionData.counts[key] || 0;
      var week = Math.floor(calendarDayDiff(dateWindow.start, cursor) / 7) + 1;
      var day = cursor.getDay() + 1;
      var square = document.createElement("span");
      var inSelectedYear = cursor.getFullYear() === year;
      var level = inSelectedYear ? getContributionLevel(count, contributionData.maxCount) : 0;

      square.className = "contribution-square";
      square.setAttribute("data-level", level);
      square.style.gridColumn = week;
      square.style.gridRow = day;
      square.title = inSelectedYear
        ? contributionText(count) + " on " + formatShortDate(cursor)
        : "";
      square.setAttribute("aria-label", square.title || "Outside selected year");
      cells.appendChild(square);
      cursor.setDate(cursor.getDate() + 1);
    }

    body.appendChild(weekdays);
    body.appendChild(cells);
    grid.appendChild(body);
  }

  function renderContributionGraph() {
    var grid = byId("contribution-grid");
    var years = getContributionYears(state.entries);
    var selectedYear = getDefaultContributionYear(years);
    var contributionData;
    var dateWindow;
    var rangeText;

    state.selectedContributionYear = selectedYear;
    renderContributionYears(years, selectedYear);

    if (!grid) {
      return;
    }

    clearElement(grid);

    if (!selectedYear) {
      setText("contribution-summary", "0 contributions");
      setText("contribution-range", "No activity yet");
      grid.appendChild(makeEmptyState("No contribution data"));
      return;
    }

    contributionData = buildContributionData(state.entries, selectedYear);
    dateWindow = getContributionDateWindow(selectedYear);

    setText(
      "contribution-summary",
      contributionText(contributionData.total) + " in " + selectedYear
    );

    if (contributionData.firstActivity && contributionData.lastActivity) {
      rangeText = formatShortDate(contributionData.firstActivity) +
        " - " +
        formatShortDate(contributionData.lastActivity);
    } else {
      rangeText = "No contributions in " + selectedYear;
    }
    setText("contribution-range", rangeText);

    renderContributionMonthLabels(grid, selectedYear, dateWindow);
    renderContributionCells(grid, selectedYear, dateWindow, contributionData);
  }

  function statusTone(status) {
    var value = text(status).toLowerCase();
    if (!value) {
      return "";
    }
    if (value.indexOf("success") !== -1 || value.indexOf("ready") !== -1) {
      return "success";
    }
    if (value.indexOf("warn") !== -1 || value.indexOf("cancel") !== -1 || value.indexOf("skip") !== -1) {
      return "warning";
    }
    if (value.indexOf("error") !== -1 || value.indexOf("fail") !== -1) {
      return "error";
    }
    return "";
  }

  function appendCell(row, className) {
    var cell = document.createElement("td");
    if (className) {
      cell.className = className;
    }
    row.appendChild(cell);
    return cell;
  }

  function appendMainMuted(cell, mainText, mutedText) {
    var main = document.createElement("div");
    var muted = document.createElement("div");

    main.className = "cell-main";
    main.textContent = mainText || "-";
    main.title = mainText || "";
    cell.appendChild(main);

    if (mutedText) {
      muted.className = "cell-muted";
      muted.textContent = mutedText;
      muted.title = mutedText;
      cell.appendChild(muted);
    }
  }

  function detailText(entry) {
    var details = [];

    if (entry.familyName) {
      details.push(entry.familyName);
    }
    if (entry.familyOrigin) {
      details.push(entry.familyOrigin);
    }
    if (entry.revitVersion) {
      details.push("Revit " + entry.revitVersion);
    }
    if (entry.revitBuild) {
      details.push(entry.revitBuild);
    }

    return details.join(" | ");
  }

  function renderRecentTable(entries) {
    var body = byId("recent-table-body");
    var sorted = entries.slice().sort(function (a, b) {
      if (b.time !== a.time) {
        return b.time - a.time;
      }
      return b.index - a.index;
    });
    var visibleRows = sorted.slice(0, 80);

    setText("recent-subtitle", formatNumber(entries.length) + " entries");

    if (!body) {
      return;
    }

    clearElement(body);

    if (!visibleRows.length) {
      var emptyRow = document.createElement("tr");
      var emptyCell = document.createElement("td");
      emptyCell.colSpan = 5;
      emptyCell.className = "empty-cell";
      emptyCell.textContent = "No activity in current filter.";
      emptyRow.appendChild(emptyCell);
      body.appendChild(emptyRow);
      return;
    }

    visibleRows.forEach(function (entry) {
      var row = document.createElement("tr");
      var dateCell = appendCell(row);
      var actionCell = appendCell(row);
      var docCell = appendCell(row);
      var statusCell = appendCell(row);
      var detailCell = appendCell(row);
      var pill = document.createElement("span");
      var tone = statusTone(entry.status);

      appendMainMuted(dateCell, entry.date ? formatDateTime(entry.date, entry.datetime) : (entry.datetime || "-"), "");
      appendMainMuted(actionCell, entry.action, "");
      appendMainMuted(docCell, docLabel(entry), entry.docPath);

      pill.className = "status-pill";
      if (tone) {
        pill.setAttribute("data-status", tone);
      }
      pill.textContent = entry.status || "-";
      statusCell.appendChild(pill);

      appendMainMuted(detailCell, detailText(entry) || "-", entry.familyPath);
      body.appendChild(row);
    });
  }

  function renderAll() {
    var filtered = filterEntries(state.entries);
    state.filteredEntries = filtered;
    renderPayloadMeta();
    renderMetrics(filtered);
    renderCharts(filtered);
    renderRecentTable(filtered);
  }

  function loadData(payload) {
    var refreshButton = byId("refresh-data");

    state.payload = payload || {};
    state.entries = normalizeEntries(state.payload.entries || []);

    if (refreshButton) {
      refreshButton.classList.remove("is-loading");
      refreshButton.textContent = "Refresh";
    }

    populateActionFilter();
    renderContributionGraph();
    renderAll();
  }

  function refreshData() {
    var button = byId("refresh-data");
    if (button) {
      button.classList.add("is-loading");
      button.textContent = "Refreshing...";
    }

    if (!postWebViewMessage({ type: "refreshData" }) && button) {
      button.classList.remove("is-loading");
      button.textContent = "Refresh";
    }
  }

  function closeWindow() {
    if (postWebViewMessage({ type: "closeWindow" })) {
      return;
    }

    try {
      globalScope.close();
    } catch (ignore) {
      // Browser fallback only.
    }
  }

  function onResize() {
    if (state.resizeTimer) {
      globalScope.clearTimeout(state.resizeTimer);
    }
    state.resizeTimer = globalScope.setTimeout(function () {
      drawActivityChart(state.filteredEntries || []);
    }, 120);
  }

  function init() {
    var rangeFilter = byId("range-filter");
    var actionFilter = byId("action-filter");
    var documentSearch = byId("document-search");
    var refreshButton = byId("refresh-data");
    var closeButton = byId("close-window");

    if (rangeFilter) {
      rangeFilter.addEventListener("change", renderAll);
    }
    if (actionFilter) {
      actionFilter.addEventListener("change", renderAll);
    }
    if (documentSearch) {
      documentSearch.addEventListener("input", renderAll);
    }
    if (refreshButton) {
      refreshButton.addEventListener("click", refreshData);
    }
    if (closeButton) {
      closeButton.addEventListener("click", closeWindow);
    }

    globalScope.addEventListener("resize", onResize);
    postWebViewMessage({ type: "appReady" });
  }

  globalScope.ffeAnalytics = {
    loadData: loadData
  };

  if (typeof document !== "undefined") {
    document.addEventListener("DOMContentLoaded", init);
  }
}(typeof window !== "undefined" ? window : globalThis));
