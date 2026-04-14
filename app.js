const DATA_URLS = {
  monthly:
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vRxPyhikHbKtGOyKP5lKkH5oZZHpzGOgEMnxVV1_ObCzrT_R1N-VBbzd9DzOaxPICgsJyxx24crOMWG/pub?gid=1529654031&single=true&output=csv",
  aht:
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vRxPyhikHbKtGOyKP5lKkH5oZZHpzGOgEMnxVV1_ObCzrT_R1N-VBbzd9DzOaxPICgsJyxx24crOMWG/pub?gid=386528999&single=true&output=csv",
  crosswalk:
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vRxPyhikHbKtGOyKP5lKkH5oZZHpzGOgEMnxVV1_ObCzrT_R1N-VBbzd9DzOaxPICgsJyxx24crOMWG/pub?gid=1559798172&single=true&output=csv",
};

const SCORE_LABELS = {
  1: "Poor",
  2: "Below Expectations",
  3: "Meets Expectations",
  4: "Above Expectations",
  5: "Excellent",
};

const SUMMARY_CARD_ICONS = {
  overall: "/icons/overall.svg",
  attendance: "/icons/attendance.svg",
  qa: "/icons/qa.svg",
  aht: "/icons/aht.svg",
};

const MONTH_ORDER = [
  "january",
  "february",
  "march",
  "april",
  "may",
  "june",
  "july",
  "august",
  "september",
  "october",
  "november",
  "december",
];

const state = {
  initialized: false,
  activeTab: "all",
  monthlyRows: [],
  ahtRows: [],
  aliasLookup: new Map(),
  qaWeeklyAvailable: false,
  options: {
    years: [],
    monthsByYear: new Map(),
    agents: [],
    latestYear: "",
    latestMonth: "",
    minAhtDate: "",
    maxAhtDate: "",
  },
  filters: {
    year: "",
    month: "",
    week: "all",
    agent: "all",
    ahtStart: "",
    ahtEnd: "",
  },
  charts: {
    primary: null,
    secondary: null,
  },
  tableSort: {
    key: "agent",
    direction: "asc",
  },
  currentDataset: null,
  mobileLegendOpen: false,
};

const elements = {
  dataStatusText: document.querySelector("#dataStatusText"),
  sourceNoteText: document.querySelector("#sourceNoteText"),
  latestPeriodText: document.querySelector("#latestPeriodText"),
  ahtRangeText: document.querySelector("#ahtRangeText"),
  mobileTopbarContext: document.querySelector("#mobileTopbarContext"),
  mobileHomeButtons: [...document.querySelectorAll("[data-mobile-target]")],
  mobileBottomNavButtons: [...document.querySelectorAll(".mobile-bottom-nav-button[data-mobile-action]")],
  legendPanel: document.querySelector("#legendPanel"),
  mobileLegendPanel: document.querySelector("#mobileLegendPanel"),
  mobileLegendContent: document.querySelector("#mobileLegendContent"),
  mobileLegendClose: document.querySelector("#mobileLegendClose"),
  chartsSection: document.querySelector("#chartsSection"),
  compactInsightsGrid: document.querySelector(".compact-insights-grid"),
  filtersTitle: document.querySelector("#filtersTitle"),
  filtersGrid: document.querySelector("#filtersGrid"),
  resetFilters: document.querySelector("#resetFilters"),
  summaryGrid: document.querySelector("#summaryGrid"),
  managerInsights: document.querySelector("#managerInsights"),
  singleAgentOverview: document.querySelector("#singleAgentOverview"),
  primaryChartTitle: document.querySelector("#primaryChartTitle"),
  primaryChartSubcopy: document.querySelector("#primaryChartSubcopy, #primaryChartSubnote"),
  secondaryChartTitle: document.querySelector("#secondaryChartTitle"),
  secondaryChartSubcopy: document.querySelector("#secondaryChartSubnote, #secondaryChartSubcopy"),
  primaryChart: document.querySelector("#primaryChart"),
  secondaryChart: document.querySelector("#secondaryChart"),
  secondaryChartCard: document.querySelector("#secondaryChart")?.closest(".chart-card"),
  topBottomCard: document.querySelector("#topBottomChart")?.closest(".chart-card"),
  topBottomTitle: document.querySelector("#topBottomTitle"),
  topBottomChart: document.querySelector("#topBottomChart"),
  ahtComponentsCard: document.querySelector("#ahtComponentsCard"),
  ahtComponentsTitle: document.querySelector("#ahtComponentsTitle"),
  ahtComponentsSubcopy: document.querySelector("#ahtComponentsSubcopy"),
  ahtComponentsChart: document.querySelector("#ahtComponentsChart"),
  callsVolumeCard: document.querySelector("#callsVolumeCard"),
  callsVolumeTitle: document.querySelector("#callsVolumeTitle"),
  callsVolumeChart: document.querySelector("#callsVolumeChart"),
  compositionCard: document.querySelector("#compositionCard"),
  compositionTitle: document.querySelector("#compositionTitle"),
  compositionChart: document.querySelector("#compositionChart"),
  trendChartTitle: document.querySelector("#trendChartTitle"),
  trendChartSubcopy: document.querySelector("#trendChartSubcopy"),
  trendScopeSummary: document.querySelector("#trendScopeSummary"),
  trendChart: document.querySelector("#trendChart"),
  trendSection: document.querySelector("#trendSection"),
  tableTitle: document.querySelector("#tableTitle"),
  tableSubcopy: document.querySelector("#tableSubcopy"),
  primaryScopeSummary: document.querySelector("#primaryScopeSummary"),
  secondaryScopeSummary: document.querySelector("#secondaryScopeSummary"),
  filtersScopeSummary: document.querySelector("#filtersScopeSummary"),
  tableScopeSummary: document.querySelector("#tableScopeSummary"),
  resultsCount: document.querySelector("#resultsCount"),
  exportTableCsv: document.querySelector("#exportTableCsv"),
  tableHeadRow: document.querySelector("#tableHeadRow"),
  tableBody: document.querySelector("#tableBody"),
  tabButtons: [...document.querySelectorAll("[data-tab]")],
  mobileLogoutButton: document.querySelector("#googleLogoutMobile"),
};

let removeTrendTooltipDismiss = null;

function normalizeHeader(value) {
  return String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function normalizeName(value) {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

function normalizeAgentKey(value) {
  return normalizeName(value).toLowerCase();
}

function csvToRows(text) {
  const rows = [];
  let current = "";
  let row = [];
  let inQuotes = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const nextChar = text[index + 1];

    if (char === '"') {
      if (inQuotes && nextChar === '"') {
        current += '"';
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (char === "," && !inQuotes) {
      row.push(current);
      current = "";
      continue;
    }

    if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && nextChar === "\n") {
        index += 1;
      }
      row.push(current);
      if (row.some((cell) => String(cell).trim() !== "")) {
        rows.push(row);
      }
      row = [];
      current = "";
      continue;
    }

    current += char;
  }

  if (current || row.length) {
    row.push(current);
    if (row.some((cell) => String(cell).trim() !== "")) {
      rows.push(row);
    }
  }

  return rows;
}

function mapCsvRows(text) {
  const [headerRow = [], ...bodyRows] = csvToRows(text);
  const headers = headerRow.map(normalizeHeader);
  return bodyRows.map((row) =>
    headers.reduce((record, header, index) => {
      record[header] = row[index] ?? "";
      return record;
    }, {})
  );
}

async function fetchCsvRows(url) {
  const response = await fetch(url, { cache: "no-store" });
  if (!response.ok) {
    throw new Error(`Unable to load ${url}`);
  }
  return mapCsvRows(await response.text());
}

function safeNumber(value) {
  if (value === null || value === undefined) return null;
  const trimmed = String(value).trim();
  if (!trimmed) return null;
  const numeric = Number(trimmed.replace(/[,%]/g, ""));
  return Number.isFinite(numeric) ? numeric : null;
}

function durationToSeconds(value) {
  const trimmed = String(value ?? "").trim();
  if (!trimmed) return 0;
  const parts = trimmed.split(":").map(Number);
  if (parts.some((part) => Number.isNaN(part))) return 0;
  if (parts.length === 3) return parts[0] * 3600 + parts[1] * 60 + parts[2];
  if (parts.length === 2) return parts[0] * 60 + parts[1];
  return parts[0];
}

function formatDurationFromSeconds(totalSeconds) {
  if (totalSeconds === null || totalSeconds === undefined || Number.isNaN(totalSeconds)) return "No entry";
  const normalized = Math.max(0, Math.round(totalSeconds));
  const hours = Math.floor(normalized / 3600);
  const minutes = Math.floor((normalized % 3600) / 60);
  const seconds = normalized % 60;
  if (hours > 0) {
    return `${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")}:${String(seconds).padStart(2, "0")}`;
  }
  return `${String(minutes).padStart(2, "0")}:${String(seconds).padStart(2, "0")}`;
}

function monthIndex(monthName) {
  return MONTH_ORDER.indexOf(String(monthName ?? "").trim().toLowerCase());
}

function formatMonthLabel(year, month) {
  return year && month ? `${month} ${year}` : "--";
}

function parseSourceDate(value, fallbackYear) {
  const trimmed = String(value ?? "").trim();
  if (!trimmed) return null;
  const parsed = new Date(`${trimmed}, ${fallbackYear}`);
  if (Number.isNaN(parsed.getTime())) return null;
  return new Date(Date.UTC(parsed.getFullYear(), parsed.getMonth(), parsed.getDate()));
}

function toInputDate(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  return date.toISOString().slice(0, 10);
}

function formatDate(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "N/A";
  return date.toLocaleDateString("en-US", {
    month: "short",
    day: "numeric",
    year: "numeric",
    timeZone: "UTC",
  });
}

function formatShortDate(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "N/A";
  return date.toLocaleDateString("en-US", {
    month: "short",
    day: "numeric",
    timeZone: "UTC",
  });
}

function formatDateRange(start, end) {
  if (!start && !end) return "N/A";
  if (start && end) return `${formatDate(start)} to ${formatDate(end)}`;
  return formatDate(start || end);
}

function formatPercent(value, digits = 2) {
  if (value === null || value === undefined || Number.isNaN(value)) return "N/A";
  return `${Number(value).toFixed(digits)}%`;
}

function formatNumber(value) {
  if (value === null || value === undefined || Number.isNaN(value)) return "N/A";
  return Number(value).toFixed(2);
}

function setText(element, value) {
  if (!element) return;
  element.textContent = value;
}

function calculateAttendanceScore(rawValue) {
  const trimmed = String(rawValue ?? "").trim();
  if (!trimmed) return 4;
  const percent = safeNumber(rawValue);
  if (percent === null) return null;
  if (percent >= 100) return 5;
  if (percent >= 95) return 3;
  if (percent >= 90) return 2;
  return 1;
}

function calculateQaScore(rawValue) {
  const percent = safeNumber(rawValue);
  if (percent === null) return null;
  if (percent >= 100) return 5;
  if (percent >= 99) return 4;
  if (percent >= 98) return 3;
  if (percent >= 95) return 2;
  return 1;
}

function calculateAhtScore(seconds) {
  if (seconds === null || seconds === undefined || Number.isNaN(seconds)) return 4;
  if (seconds < 43) return 5;
  if (seconds <= 181) return 3;
  if (seconds <= 252) return 2;
  return 1;
}

function equivalentFromScore(score) {
  if (score === null || score === undefined || Number.isNaN(score)) return "N/A";
  if (score >= 4.5) return SCORE_LABELS[5];
  if (score >= 3.5) return SCORE_LABELS[4];
  if (score >= 2.5) return SCORE_LABELS[3];
  if (score >= 1.5) return SCORE_LABELS[2];
  return SCORE_LABELS[1];
}

function average(values) {
  const clean = values.filter((value) => typeof value === "number" && !Number.isNaN(value));
  if (!clean.length) return null;
  return clean.reduce((sum, value) => sum + value, 0) / clean.length;
}

function previousMonthContext(year, month) {
  const currentMonthIndex = monthIndex(month);
  if (!year || currentMonthIndex < 0) return null;
  const previousDate = new Date(Date.UTC(Number(year), currentMonthIndex, 1));
  previousDate.setUTCMonth(previousDate.getUTCMonth() - 1);
  return {
    year: String(previousDate.getUTCFullYear()),
    month: MONTH_ORDER[previousDate.getUTCMonth()]?.replace(/^\w/, (char) => char.toUpperCase()) || "",
  };
}

function getMonthlyRowsForPeriod(year, month, agent = state.filters.agent) {
  return state.monthlyRows
    .filter((row) => row.year === year)
    .filter((row) => row.month === month)
    .filter((row) => agent === "all" || row.agent === agent);
}

function getMonthlyAhtRowsForPeriod(year, month, agent = state.filters.agent) {
  const monthValue = monthIndex(month);
  if (!year || monthValue < 0) return [];

  return state.ahtRows
    .filter((row) => row.date instanceof Date && !Number.isNaN(row.date.getTime()))
    .filter((row) => row.date.getUTCFullYear() === Number(year))
    .filter((row) => row.date.getUTCMonth() === monthValue)
    .filter((row) => agent === "all" || row.agent === agent);
}

function getPreviousAhtRangeRows() {
  const start = state.filters.ahtStart ? new Date(`${state.filters.ahtStart}T00:00:00Z`) : null;
  const end = state.filters.ahtEnd ? new Date(`${state.filters.ahtEnd}T00:00:00Z`) : null;
  if (!start || !end) return [];

  const daySpan = Math.max(1, Math.round((end.getTime() - start.getTime()) / 86400000) + 1);
  const previousEnd = new Date(start);
  previousEnd.setUTCDate(previousEnd.getUTCDate() - 1);
  const previousStart = new Date(previousEnd);
  previousStart.setUTCDate(previousEnd.getUTCDate() - (daySpan - 1));

  return state.ahtRows
    .filter((row) => row.date instanceof Date && !Number.isNaN(row.date.getTime()))
    .filter((row) => row.date >= previousStart && row.date <= new Date(`${toInputDate(previousEnd)}T23:59:59Z`))
    .filter((row) => state.filters.agent === "all" || row.agent === state.filters.agent);
}

function formatTrendDelta(current, previous, mode = "number") {
  if ([current, previous].some((value) => value === null || value === undefined || Number.isNaN(value))) return "";
  const delta = current - previous;
  const sign = delta > 0 ? "+" : delta < 0 ? "-" : "";
  const absolute = Math.abs(delta);
  if (mode === "duration") return `${sign}${formatDurationFromSeconds(absolute)}`;
  if (mode === "count") return `${sign}${Math.round(absolute)}`;
  return `${sign}${absolute.toFixed(2)}`;
}

function buildTrendChip(current, previous, options = {}) {
  if ([current, previous].some((value) => value === null || value === undefined || Number.isNaN(value))) return null;
  const {
    higherIsBetter = true,
    mode = "number",
    scopeLabel = "previous period",
  } = options;
  const delta = current - previous;
  const isNeutral = Math.abs(delta) < (mode === "count" ? 1 : 0.005);
  const improved = higherIsBetter ? delta > 0 : delta < 0;
  const tone = isNeutral ? "neutral" : improved ? "up" : "down";
  const arrow = isNeutral ? "Flat" : delta > 0 ? "Up" : "Down";
  const deltaLabel = formatTrendDelta(current, previous, mode);
  return {
    tone,
    text: isNeutral ? "Flat vs prev" : `${arrow} ${deltaLabel}`,
    tooltip: `${deltaLabel || "No change"} compared with ${scopeLabel}`,
  };
}

function previewAgentNames(names, limit = 2) {
  const clean = [...new Set((names || []).filter(Boolean))];
  if (!clean.length) return "No agents in current scope";
  if (clean.length <= limit) return clean.join(", ");
  return `${clean.slice(0, limit).join(", ")} +${clean.length - limit} more`;
}

function isNeedsAttention(scores) {
  const numericScores = (scores || []).filter((score) => typeof score === "number" && !Number.isNaN(score));
  const poorCount = numericScores.filter((score) => score === 1).length;
  const lowCount = numericScores.filter((score) => score <= 2).length;
  return poorCount >= 1 || lowCount >= 2;
}

function formatInsightCount(count, singular = "Agent", plural = "Agents") {
  return `${count} ${count === 1 ? singular : plural}`;
}

function managerInsightScopeLabel() {
  if (state.activeTab === "aht") return "Selected Range";
  if (state.activeTab === "qa" && state.filters.week !== "all") return formatWeekLabel(state.filters.week);
  return "Current Scope";
}

function renderManagerInsights(dataset) {
  if (!elements.managerInsights) return;

  const shouldShow = !isSingleAgentView(dataset) && dataset.rows.length > 0;
  elements.managerInsights.hidden = !shouldShow;
  if (!shouldShow) {
    elements.managerInsights.innerHTML = "";
    return;
  }

  let cards = [];

  if (state.activeTab === "all") {
    const averages = [
      { label: "Attendance", value: average(dataset.rows.map((row) => row.attendanceScore)), tone: "attendance" },
      { label: "QA", value: average(dataset.rows.map((row) => row.qaScore)), tone: "qa" },
      { label: "AHT", value: average(dataset.rows.map((row) => row.ahtScore)), tone: "aht" },
    ].filter((item) => typeof item.value === "number" && !Number.isNaN(item.value));

    const strongest = [...averages].sort((a, b) => b.value - a.value)[0] || null;
    const weakest = [...averages].sort((a, b) => a.value - b.value)[0] || null;
    const attentionRows = dataset.rows.filter((row) => isNeedsAttention([row.attendanceScore, row.qaScore, row.ahtScore]));
    const bestPerformer = [...dataset.rows]
      .filter((row) => typeof row.overallScore === "number" && !Number.isNaN(row.overallScore))
      .sort((a, b) => b.overallScore - a.overallScore)[0] || null;

    cards = [
      strongest && {
        eyebrow: managerInsightScopeLabel(),
        badge: "Highest Avg",
        title: "Strongest KPI",
        value: strongest.label,
        note: `Team avg ${formatNumber(strongest.value)}`,
        meta: "Highest average score this month",
        tone: strongest.tone,
      },
      weakest && {
        eyebrow: managerInsightScopeLabel(),
        badge: "Lowest Avg",
        title: "Weakest KPI",
        value: weakest.label,
        note: `Team avg ${formatNumber(weakest.value)}`,
        meta: "Lowest average score this month",
        tone: weakest.tone,
      },
      {
        eyebrow: managerInsightScopeLabel(),
        badge: "Priority",
        title: "Needs Attention",
        value: formatInsightCount(attentionRows.length),
        note: previewAgentNames(attentionRows.map((row) => row.agent)),
        meta: "At least one Poor score or two KPI scores at 2 or below",
        tone: "alert",
      },
      bestPerformer && {
        eyebrow: managerInsightScopeLabel(),
        badge: "Leader",
        title: "Best Performer",
        value: bestPerformer.agent,
        note: `Overall ${formatNumber(bestPerformer.overallScore)} • ${bestPerformer.equivalent || "N/A"}`,
        meta: "Highest overall score in the current month",
        tone: "overall",
      },
    ].filter(Boolean);
  } else if (state.activeTab === "attendance") {
    const attendanceRows = dataset.rows.map((row) => ({
      ...row,
      attendancePercent: safeNumber(row.attendancePercentDisplay),
    }));
    const bestAttendance = [...attendanceRows]
      .filter((row) => typeof row.attendancePercent === "number" && !Number.isNaN(row.attendancePercent))
      .sort((a, b) => b.attendancePercent - a.attendancePercent)[0] || null;
    const attentionRows = attendanceRows.filter((row) => isNeedsAttention([row.attendanceScore]));
    const perfectRows = attendanceRows.filter((row) => typeof row.attendancePercent === "number" && row.attendancePercent >= 100);
    const teamAverage = average(attendanceRows.map((row) => row.attendancePercent));

    cards = [
      bestAttendance && {
        eyebrow: managerInsightScopeLabel(),
        badge: "Leader",
        title: "Best Attendance",
        value: bestAttendance.agent,
        note: bestAttendance.attendancePercentDisplay,
        meta: "Highest attendance percentage this month",
        tone: "attendance",
      },
      {
        eyebrow: managerInsightScopeLabel(),
        badge: "Priority",
        title: "Needs Attention",
        value: formatInsightCount(attentionRows.length),
        note: previewAgentNames(attentionRows.map((row) => row.agent)),
        meta: "Attendance score is 2 or below",
        tone: "alert",
      },
      {
        eyebrow: managerInsightScopeLabel(),
        badge: "Coverage",
        title: "Perfect Attendance",
        value: formatInsightCount(perfectRows.length),
        note: previewAgentNames(perfectRows.map((row) => row.agent)),
        meta: "Reached 100% attendance this month",
        tone: "attendance",
      },
      {
        eyebrow: managerInsightScopeLabel(),
        badge: "Average",
        title: "Team Average",
        value: formatPercent(teamAverage),
        note: `Score ${formatNumber(average(attendanceRows.map((row) => row.attendanceScore)))}`,
        meta: "Average attendance for the current month",
        tone: "overall",
      },
    ];
  } else if (state.activeTab === "qa") {
    const qaRows = dataset.rows.map((row) => ({
      ...row,
      qaAverage: safeNumber(row.qaAverageDisplay),
    }));
    const bestQa = [...qaRows]
      .filter((row) => typeof row.qaAverage === "number" && !Number.isNaN(row.qaAverage))
      .sort((a, b) => b.qaAverage - a.qaAverage)[0] || null;
    const attentionRows = qaRows.filter((row) => isNeedsAttention([row.qaScore]));
    const excellentRows = qaRows.filter((row) => typeof row.qaScore === "number" && row.qaScore >= 4);
    const teamAverage = average(qaRows.map((row) => row.qaAverage));

    cards = [
      bestQa && {
        eyebrow: managerInsightScopeLabel(),
        badge: "Leader",
        title: "Best QA",
        value: bestQa.agent,
        note: bestQa.qaAverageDisplay,
        meta: "Highest QA average in the selected scope",
        tone: "qa",
      },
      {
        eyebrow: managerInsightScopeLabel(),
        badge: "Priority",
        title: "Needs Attention",
        value: formatInsightCount(attentionRows.length),
        note: previewAgentNames(attentionRows.map((row) => row.agent)),
        meta: "QA score is 2 or below",
        tone: "alert",
      },
      {
        eyebrow: managerInsightScopeLabel(),
        badge: "High Scores",
        title: "Excellent QA",
        value: formatInsightCount(excellentRows.length),
        note: previewAgentNames(excellentRows.map((row) => row.agent)),
        meta: "Scored Above Expectations or Excellent",
        tone: "qa",
      },
      {
        eyebrow: managerInsightScopeLabel(),
        badge: "Average",
        title: "Team Average",
        value: formatPercent(teamAverage),
        note: `Score ${formatNumber(average(qaRows.map((row) => row.qaScore)))}`,
        meta: "Average QA for the selected scope",
        tone: "overall",
      },
    ];
  } else if (state.activeTab === "aht") {
    const ahtRows = dataset.rows.map((row) => ({
      ...row,
      totalCalls: (row.inboundCalls ?? 0) + (row.outboundCalls ?? 0),
      ahtSeconds: durationToSeconds(row.ahtDisplay),
    }));
    const bestAht = [...ahtRows]
      .filter((row) => typeof row.ahtScore === "number" && !Number.isNaN(row.ahtScore))
      .sort((a, b) => (b.ahtScore - a.ahtScore) || (a.ahtSeconds - b.ahtSeconds))[0] || null;
    const attentionRows = ahtRows.filter((row) => isNeedsAttention([row.ahtScore]));
    const volumeLeader = [...ahtRows].sort((a, b) => b.totalCalls - a.totalCalls)[0] || null;
    const teamAverage = average(ahtRows.map((row) => row.ahtSeconds));

    cards = [
      bestAht && {
        eyebrow: managerInsightScopeLabel(),
        badge: "Leader",
        title: "Best AHT",
        value: bestAht.agent,
        note: `${bestAht.ahtDisplay} • Score ${formatNumber(bestAht.ahtScore)}`,
        meta: "Best AHT result in the selected range",
        tone: "aht",
      },
      {
        eyebrow: managerInsightScopeLabel(),
        badge: "Priority",
        title: "Needs Attention",
        value: formatInsightCount(attentionRows.length),
        note: previewAgentNames(attentionRows.map((row) => row.agent)),
        meta: "AHT score is 2 or below",
        tone: "alert",
      },
      volumeLeader && {
        eyebrow: managerInsightScopeLabel(),
        badge: "Volume",
        title: "Call Volume Leader",
        value: volumeLeader.agent,
        note: `${volumeLeader.totalCalls} total calls`,
        meta: "Highest handled volume in the selected range",
        tone: "overall",
      },
      {
        eyebrow: managerInsightScopeLabel(),
        badge: "Average",
        title: "Team Average",
        value: formatDurationFromSeconds(teamAverage),
        note: `Score ${formatNumber(average(ahtRows.map((row) => row.ahtScore)))}`,
        meta: "Average AHT in the selected range",
        tone: "aht",
      },
    ].filter(Boolean);
  }

  elements.managerInsights.innerHTML = cards
    .map(
      (card) => `
        <article class="chart-card card insight-card insight-card-${card.tone}">
          <div class="insight-card-head">
            <p class="eyebrow">${escapeHtml(card.eyebrow || "Current Scope")}</p>
            <span class="insight-card-badge">${escapeHtml(card.badge || card.title)}</span>
          </div>
          <h3>${escapeHtml(card.title)}</h3>
          <strong>${escapeHtml(card.value)}</strong>
          <p>${escapeHtml(card.note)}</p>
          <span class="insight-card-meta">${escapeHtml(card.meta)}</span>
        </article>
      `
    )
    .join("");
}

function scoreBadge(score) {
  const numeric = typeof score === "number" && !Number.isNaN(score) ? score : null;
  const rounded = numeric === null ? 4 : Math.max(1, Math.min(5, Math.round(numeric)));
  const label = numeric === null ? "N/A" : numeric.toFixed(2);
  return `<span class="score-pill score-${rounded}">${label}</span>`;
}

function statusBadge(value) {
  const label = String(value ?? "").trim();
  if (!label) return "N/A";
  const tone = label.toLowerCase().replace(/[^a-z0-9]+/g, "-");
  return `<span class="status-pill status-${tone}">${label}</span>`;
}

function csvEscape(value) {
  const text = String(value ?? "");
  if (/[",\n\r]/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`;
  }
  return text;
}

function exportCellValue(column, row) {
  if (column.type === "score") {
    const numeric = typeof row[column.key] === "number" && !Number.isNaN(row[column.key]) ? row[column.key] : null;
    return numeric === null ? "N/A" : numeric.toFixed(2);
  }

  return row[column.key] ?? "N/A";
}

function getTableSortValue(column, row) {
  const rawValue = row?.[column.key];
  if (column.type === "score") {
    return typeof rawValue === "number" && !Number.isNaN(rawValue) ? rawValue : null;
  }

  if (column.key === "agent" || column.key === "equivalent") {
    return String(rawValue ?? "").trim().toLowerCase();
  }

  if (column.key === "attendancePercentDisplay" || column.key === "qaAverageDisplay" || column.key === "totalAverageDisplay") {
    return safeNumber(rawValue);
  }

  if (["inboundCalls", "outboundCalls"].includes(column.key)) {
    return safeNumber(rawValue);
  }

  if (["inboundMinutes", "outboundMinutes", "talkTime", "holdTime", "ahtDisplay"].includes(column.key)) {
    return durationToSeconds(rawValue);
  }

  if (column.key === "dateRange") {
    const startValue = String(rawValue ?? "").split(" to ")[0]?.trim();
    const parsed = startValue ? new Date(startValue) : null;
    return parsed instanceof Date && !Number.isNaN(parsed.getTime()) ? parsed.getTime() : String(rawValue ?? "").trim().toLowerCase();
  }

  if (typeof rawValue === "number") {
    return rawValue;
  }

  const numeric = safeNumber(rawValue);
  if (numeric !== null) return numeric;

  return String(rawValue ?? "").trim().toLowerCase();
}

function compareTableValues(left, right) {
  const leftMissing = left === null || left === undefined || left === "";
  const rightMissing = right === null || right === undefined || right === "";
  if (leftMissing && rightMissing) return 0;
  if (leftMissing) return 1;
  if (rightMissing) return -1;

  if (typeof left === "number" && typeof right === "number") {
    return left - right;
  }

  return String(left).localeCompare(String(right), undefined, { sensitivity: "base", numeric: true });
}

function sortDatasetRows(dataset) {
  const sortColumn = dataset.columns.find((column) => column.key === state.tableSort.key) || dataset.columns[0];
  if (!sortColumn) return [...dataset.rows];

  const directionFactor = state.tableSort.direction === "desc" ? -1 : 1;
  return [...dataset.rows].sort((leftRow, rightRow) => {
    const comparison = compareTableValues(
      getTableSortValue(sortColumn, leftRow),
      getTableSortValue(sortColumn, rightRow)
    );
    if (comparison !== 0) return comparison * directionFactor;
    return compareTableValues(String(leftRow.agent ?? "").toLowerCase(), String(rightRow.agent ?? "").toLowerCase());
  });
}

function buildExportFilename() {
  const tabLabel =
    state.activeTab === "attendance"
      ? "Attendance"
      : state.activeTab === "qa"
        ? "Quality-Assurance"
        : state.activeTab === "aht"
          ? "AHT"
          : "All-KPIs";

  const sanitizePart = (value) =>
    String(value ?? "")
      .trim()
      .replace(/[\\/:*?"<>|]+/g, "")
      .replace(/\s+/g, "-")
      .replace(/-+/g, "-");

  const monthLabel = state.filters.year && state.filters.month
    ? `${state.filters.month}-${state.filters.year}`
    : "Current-Scope";
  const agentLabel = state.filters.agent === "all" ? "All-Agents" : sanitizePart(state.filters.agent);

  const scopeLabel =
    state.activeTab === "aht"
      ? sanitizePart(
          formatDateRange(
            state.filters.ahtStart ? new Date(`${state.filters.ahtStart}T00:00:00Z`) : null,
            state.filters.ahtEnd ? new Date(`${state.filters.ahtEnd}T00:00:00Z`) : null
          )
        )
      : state.activeTab === "qa"
        ? sanitizePart(`${monthLabel}-${formatWeekLabel(state.filters.week)}`)
        : sanitizePart(monthLabel);

  const now = new Date();
  const exportedAt = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;

  return `Banyan-Treatment-Centers_${tabLabel}_${scopeLabel}_${agentLabel}_${exportedAt}.csv`;
}

function buildExportMetadataRows(dataset) {
  const rows = [
    ["Dashboard", "Banyan Treatment Centers KPI Dashboard"],
    ["Tab", dataset.title],
  ];

  if (state.activeTab === "aht") {
    rows.push(
      [
        "Date Range",
        formatDateRange(
          state.filters.ahtStart ? new Date(`${state.filters.ahtStart}T00:00:00Z`) : null,
          state.filters.ahtEnd ? new Date(`${state.filters.ahtEnd}T00:00:00Z`) : null
        ),
      ],
      ["Agent", state.filters.agent === "all" ? "All Agents" : state.filters.agent]
    );
  } else {
    rows.push(
      ["Year", state.filters.year || "N/A"],
      ["Month", state.filters.month || "N/A"]
    );

    if (state.activeTab === "qa") {
      rows.push([
        "Week",
        state.filters.week === "all" ? "All Weeks" : state.filters.week.replace("week", "Week "),
      ]);
    }

    rows.push(["Agent", state.filters.agent === "all" ? "All Agents" : state.filters.agent]);

    if (state.activeTab === "all") {
      rows.push([
        "AHT Date Range",
        formatDateRange(
          state.filters.ahtStart ? new Date(`${state.filters.ahtStart}T00:00:00Z`) : null,
          state.filters.ahtEnd ? new Date(`${state.filters.ahtEnd}T00:00:00Z`) : null
        ),
      ]);
    }
  }

  rows.push(["Rows Exported", String(dataset.rows.length)]);
  return rows;
}

function exportCurrentTable() {
  const dataset = state.currentDataset || getDataset();
  if (!dataset?.rows?.length || !dataset?.columns?.length) return;

  const metadataRows = buildExportMetadataRows(dataset).map((row) => row.map(csvEscape).join(","));
  const headerRow = dataset.columns.map((column) => csvEscape(column.label)).join(",");
  const bodyRows = dataset.rows.map((row) =>
    dataset.columns.map((column) => csvEscape(exportCellValue(column, row))).join(",")
  );
  const csv = [...metadataRows, "", headerRow, ...bodyRows].join("\r\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = buildExportFilename();
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function setMobileNavActive(action) {
  elements.mobileBottomNavButtons.forEach((button) => {
    button.classList.toggle("is-active", button.dataset.mobileAction === action);
  });
}

function applyMobileLegendState() {
  elements.mobileLegendPanel?.classList.toggle("is-open", state.mobileLegendOpen);
  elements.mobileLegendPanel?.setAttribute("aria-hidden", String(!state.mobileLegendOpen));
  document.body?.classList.toggle("has-mobile-legend-open", state.mobileLegendOpen);
}

function initializeMobileLegend() {
  if (!elements.legendPanel || !elements.mobileLegendContent) return;
  const legendItems = [...elements.legendPanel.querySelectorAll(".legend-kpi")];
  elements.mobileLegendContent.innerHTML = legendItems.map((item) => item.outerHTML).join("");
}

function scrollToSection(targetId, action) {
  state.mobileLegendOpen = false;
  applyMobileLegendState();
  const target = document.getElementById(targetId);
  if (!target) return;
  target.scrollIntoView({ behavior: "smooth", block: "start" });
  setMobileNavActive(action);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function richTooltipData(payload) {
  return escapeHtml(
    JSON.stringify({
      title: payload?.title || "",
      lines: Array.isArray(payload?.lines)
        ? payload.lines
            .filter(Boolean)
            .map((line) => (typeof line === "string" ? { text: line } : { text: line?.text || "", color: line?.color || "" }))
            .filter((line) => line.text)
        : [],
      legend: Array.isArray(payload?.legend) ? payload.legend.filter((item) => item?.label && item?.color) : [],
    })
  );
}

let richTooltipElement = null;
let activeRichTooltipTarget = null;

function ensureRichTooltipElement() {
  if (richTooltipElement) return richTooltipElement;
  richTooltipElement = document.createElement("div");
  richTooltipElement.className = "rich-tooltip";
  richTooltipElement.hidden = true;
  document.body.appendChild(richTooltipElement);
  return richTooltipElement;
}

function renderRichTooltip(payload) {
  const tooltip = ensureRichTooltipElement();
  const title = payload?.title ? `<strong class="rich-tooltip-title">${escapeHtml(payload.title)}</strong>` : "";
  const lines = (payload?.lines || [])
    .map((line) => {
      const text = typeof line === "string" ? line : line?.text || "";
      const color = typeof line === "string" ? "" : line?.color || "";
      return `
        <span class="rich-tooltip-line${color ? " rich-tooltip-line-with-swatch" : ""}">
          ${color ? `<span class="rich-tooltip-swatch" style="background:${escapeHtml(color)}"></span>` : ""}
          <span>${escapeHtml(text)}</span>
        </span>
      `;
    })
    .join("");
  const hasInlineSwatches = (payload?.lines || []).some((line) => typeof line !== "string" && line?.color);
  const legend = (payload?.legend || []).length && !hasInlineSwatches
    ? `
        <div class="rich-tooltip-legend">
          ${payload.legend
            .map(
              (item) => `
                <span class="rich-tooltip-legend-item">
                  <span class="rich-tooltip-swatch" style="background:${escapeHtml(item.color)}"></span>
                  <span>${escapeHtml(item.label)}</span>
                </span>
              `
            )
            .join("")}
        </div>
      `
    : "";

  tooltip.innerHTML = `
    ${title}
    <div class="rich-tooltip-lines">${lines}</div>
    ${legend}
  `;
  return tooltip;
}

function positionRichTooltip(target, event = null) {
  const tooltip = ensureRichTooltipElement();
  const rect = target.getBoundingClientRect();
  const pointerX = event?.clientX ?? rect.left + rect.width / 2;
  const baseTop = rect.top - 12;

  tooltip.style.left = "0px";
  tooltip.style.top = "0px";
  tooltip.hidden = false;

  const tooltipRect = tooltip.getBoundingClientRect();
  const margin = 12;
  let left = pointerX - tooltipRect.width / 2;
  let top = baseTop - tooltipRect.height;

  if (left < margin) left = margin;
  if (left + tooltipRect.width > window.innerWidth - margin) {
    left = window.innerWidth - tooltipRect.width - margin;
  }

  if (top < margin) {
    top = rect.bottom + 12;
  }

  tooltip.style.left = `${Math.round(left)}px`;
  tooltip.style.top = `${Math.round(top)}px`;
}

function hideRichTooltip() {
  if (!richTooltipElement) return;
  richTooltipElement.hidden = true;
  richTooltipElement.innerHTML = "";
  activeRichTooltipTarget = null;
}

function showRichTooltip(target, event = null) {
  if (!target?.dataset?.richTooltip) return;
  try {
    const payload = JSON.parse(target.dataset.richTooltip);
    renderRichTooltip(payload);
    positionRichTooltip(target, event);
    activeRichTooltipTarget = target;
  } catch {
    hideRichTooltip();
  }
}

function bindRichTooltipEvents() {
  const supportsHover =
    typeof window !== "undefined" && typeof window.matchMedia === "function"
      ? window.matchMedia("(hover: hover) and (pointer: fine)").matches
      : true;
  const resolveTarget = (eventTarget) => (eventTarget instanceof Element ? eventTarget.closest("[data-rich-tooltip]") : null);

  document.addEventListener("mouseover", (event) => {
    if (!supportsHover) return;
    const target = resolveTarget(event.target);
    if (!target) return;
    showRichTooltip(target, event);
  });

  document.addEventListener("mousemove", (event) => {
    if (!supportsHover || !activeRichTooltipTarget) return;
    if (!resolveTarget(event.target)) return;
    positionRichTooltip(activeRichTooltipTarget, event);
  });

  document.addEventListener("mouseout", (event) => {
    if (!supportsHover) return;
    const target = resolveTarget(event.target);
    if (!target || target !== activeRichTooltipTarget) return;
    if (event.relatedTarget instanceof Node && target.contains(event.relatedTarget)) return;
    hideRichTooltip();
  });

  document.addEventListener("focusin", (event) => {
    const target = resolveTarget(event.target);
    if (!target) return;
    showRichTooltip(target);
  });

  document.addEventListener("focusout", (event) => {
    const target = resolveTarget(event.target);
    if (!target || target !== activeRichTooltipTarget) return;
    hideRichTooltip();
  });

  document.addEventListener("pointerdown", (event) => {
    const target = resolveTarget(event.target);
    if (target) {
      showRichTooltip(target, event);
      return;
    }
    hideRichTooltip();
  }, true);

  window.addEventListener("scroll", () => {
    if (activeRichTooltipTarget) hideRichTooltip();
  }, { passive: true });

  window.addEventListener("resize", () => {
    if (activeRichTooltipTarget) hideRichTooltip();
  });
}

function monthlyRowToModel(row) {
  const year = String(row.year ?? "").trim();
  const month = String(row.month ?? "").trim();
  const agent = normalizeName(row.agent);

  const attendancePercent = safeNumber(row.attendance);
  const qaPercent = safeNumber(row.ttl_avg);
  const ahtSeconds = row.aht ? durationToSeconds(row.aht) : null;

  const attendanceScore = calculateAttendanceScore(row.attendance);
  const qaScore = calculateQaScore(row.ttl_avg);
  const ahtScore = calculateAhtScore(ahtSeconds);
  const overallScore =
    attendanceScore === null || qaScore === null || ahtScore === null
      ? null
      : attendanceScore * 0.5 + qaScore * 0.25 + ahtScore * 0.25;

  return {
    year,
    month,
    monthIndex: monthIndex(month),
    agent,
    attendancePercent,
    attendanceScore,
    qaPercent,
    qaScore,
    ahtSeconds,
    ahtScore,
    overallScore,
    qaWeek1: safeNumber(row.qa_wk1_avg),
    qaWeek2: safeNumber(row.qa_wk2_avg),
    qaWeek3: safeNumber(row.qa_wk3_avg),
    qaWeek4: safeNumber(row.qa_wk4_avg),
    qaTotalAverage: safeNumber(row.ttl_avg),
  };
}

function createAliasLookup(rows) {
  const lookup = new Map();

  rows.forEach((row) => {
    const id1 = normalizeName(row.id1);
    const id2 = normalizeName(row.id2);
    if (id1) {
      lookup.set(normalizeAgentKey(id1), id1);
    }
    if (id2) {
      lookup.set(normalizeAgentKey(id2), id1 || id2);
    }
  });

  return lookup;
}

function resolveCanonicalAgentName(rawName) {
  const normalized = normalizeAgentKey(rawName);
  return state.aliasLookup.get(normalized) || normalizeName(rawName);
}

function ahtRowToModel(row, fallbackYear) {
  const rawAgentName = row.user_name || row.user;
  const agent = resolveCanonicalAgentName(rawAgentName);
  const date = parseSourceDate(row.date, fallbackYear);
  const inboundCalls = safeNumber(row.inbound_calls) || 0;
  const outboundCalls = safeNumber(row.outbound_calls) || 0;
  const talkSeconds = durationToSeconds(row.talk_time_h_mm_ss || row.talk_time);
  const holdSeconds = durationToSeconds(row.hold_time_h_mm_ss || row.hold_time);
  const totalCalls = inboundCalls + outboundCalls;
  const ahtSeconds = totalCalls ? (talkSeconds + holdSeconds) / totalCalls : null;

  return {
    agent,
    matchedAgent: state.aliasLookup.has(normalizeAgentKey(rawAgentName)),
    date,
    inboundCalls,
    inboundSeconds: durationToSeconds(row.inbound_minutes_h_mm_ss || row.inbound_minutes),
    outboundCalls,
    outboundSeconds: durationToSeconds(row.outbound_minutes_h_mm_ss || row.outbound_minutes),
    talkSeconds,
    holdSeconds,
    ahtSeconds,
  };
}

function buildOptions() {
  const years = [...new Set(state.monthlyRows.map((row) => row.year).filter(Boolean))].sort((a, b) => Number(a) - Number(b));
  const monthsByYear = new Map();

  years.forEach((year) => {
    const months = [...new Set(state.monthlyRows.filter((row) => row.year === year).map((row) => row.month))]
      .filter(Boolean)
      .sort((a, b) => monthIndex(a) - monthIndex(b));
    monthsByYear.set(year, months);
  });

  const latestRow = [...state.monthlyRows].sort((a, b) => {
    if (a.year !== b.year) return Number(a.year) - Number(b.year);
    return a.monthIndex - b.monthIndex;
  }).at(-1);

  const dates = state.ahtRows.map((row) => row.date).filter(Boolean).sort((a, b) => a - b);

  state.options = {
    years,
    monthsByYear,
    agents: [...new Set(state.monthlyRows.map((row) => row.agent).concat(state.ahtRows.map((row) => row.agent)).filter(Boolean))].sort(),
    latestYear: latestRow?.year || "",
    latestMonth: latestRow?.month || "",
    minAhtDate: toInputDate(dates[0]),
    maxAhtDate: toInputDate(dates.at(-1)),
  };
}

function resetFilters() {
  state.filters.year = state.options.latestYear;
  state.filters.month = state.options.latestMonth;
  state.filters.week = "all";
  state.filters.agent = "all";
  state.filters.ahtEnd = state.options.maxAhtDate;
  if (state.options.maxAhtDate) {
    const endDate = new Date(`${state.options.maxAhtDate}T00:00:00Z`);
    const startDate = new Date(Date.UTC(endDate.getUTCFullYear(), endDate.getUTCMonth(), 1));
    const boundedStart = state.options.minAhtDate
      ? new Date(Math.max(startDate.getTime(), new Date(`${state.options.minAhtDate}T00:00:00Z`).getTime()))
      : startDate;
    state.filters.ahtStart = toInputDate(boundedStart);
  } else {
    state.filters.ahtStart = state.options.minAhtDate;
  }
}

function createSelectField(key, label, value, options) {
  const wrapper = document.createElement("label");
  wrapper.className = "filter-field";

  const labelElement = document.createElement("span");
  labelElement.className = "filter-title";
  labelElement.textContent = label;

  const select = document.createElement("select");
  select.dataset.filterKey = key;

  options.forEach((option) => {
    const optionElement = document.createElement("option");
    optionElement.value = option.value;
    optionElement.textContent = option.label;
    optionElement.selected = option.value === value;
    select.appendChild(optionElement);
  });

  wrapper.append(labelElement, select);
  return wrapper;
}

function createInputField(key, label, value, min, max) {
  const wrapper = document.createElement("label");
  wrapper.className = "filter-field";

  const labelElement = document.createElement("span");
  labelElement.className = "filter-title";
  labelElement.textContent = label;

  const input = document.createElement("input");
  input.type = "date";
  input.dataset.filterKey = key;
  input.value = value;
  if (min) input.min = min;
  if (max) input.max = max;

  wrapper.append(labelElement, input);
  return wrapper;
}

function createAgentSearchField(value, options) {
  const wrapper = document.createElement("label");
  wrapper.className = "filter-field";

  const labelElement = document.createElement("span");
  labelElement.className = "filter-title";
  labelElement.textContent = "Agent";

  const input = document.createElement("input");
  input.type = "text";
  input.dataset.filterKey = "agent";
  input.className = "agent-search-input";
  input.setAttribute("list", "agentOptionsList");
  input.setAttribute("autocomplete", "off");
  input.setAttribute("placeholder", value === "all" ? "All Agents" : "Search or choose agent");
  input.value = value === "all" ? "" : value;

  const dataList = document.createElement("datalist");
  dataList.id = "agentOptionsList";
  options.forEach((option) => {
    const optionElement = document.createElement("option");
    optionElement.value = option.label;
    dataList.appendChild(optionElement);
  });

  wrapper.append(labelElement, input, dataList);
  return wrapper;
}

function renderFilters() {
  const months = state.options.monthsByYear.get(state.filters.year) || [];
  const agentOptions = [{ value: "all", label: "All Agents" }].concat(
    state.options.agents.map((agent) => ({ value: agent, label: agent }))
  );

  const fields = [];
  if (state.activeTab === "aht") {
    fields.push(createInputField("ahtStart", "Start Date", state.filters.ahtStart, state.options.minAhtDate, state.filters.ahtEnd));
    fields.push(createInputField("ahtEnd", "End Date", state.filters.ahtEnd, state.filters.ahtStart, state.options.maxAhtDate));
    fields.push(createAgentSearchField(state.filters.agent, agentOptions));
  } else {
    fields.push(
      createSelectField("year", "Year", state.filters.year, state.options.years.map((year) => ({ value: year, label: year }))),
      createSelectField("month", "Month", state.filters.month, months.map((month) => ({ value: month, label: month })))
    );
    if (state.activeTab === "qa") {
      const weekOptions = state.qaWeeklyAvailable
        ? [
            { value: "all", label: "All Weeks" },
            { value: "week1", label: "Week 1" },
            { value: "week2", label: "Week 2" },
            { value: "week3", label: "Week 3" },
            { value: "week4", label: "Week 4" },
          ]
        : [{ value: "all", label: "All Weeks (monthly total only)" }];
      fields.push(createSelectField("week", "Week", state.filters.week, weekOptions));
    }
    fields.push(createAgentSearchField(state.filters.agent, agentOptions));
  }

  elements.filtersGrid.innerHTML = "";
  fields.forEach((field) => elements.filtersGrid.appendChild(field));
}

function getFilteredMonthlyRows() {
  return state.monthlyRows
    .filter((row) => row.year === state.filters.year)
    .filter((row) => row.month === state.filters.month)
    .filter((row) => state.filters.agent === "all" || row.agent === state.filters.agent);
}

function getFilteredAhtRows() {
  const start = state.filters.ahtStart ? new Date(`${state.filters.ahtStart}T00:00:00Z`) : null;
  const end = state.filters.ahtEnd ? new Date(`${state.filters.ahtEnd}T23:59:59Z`) : null;

  return state.ahtRows
    .filter((row) => row.date)
    .filter((row) => !start || row.date >= start)
    .filter((row) => !end || row.date <= end)
      .filter((row) => state.filters.agent === "all" || row.agent === state.filters.agent);
}

function getMonthlyAhtRows() {
  const monthValue = monthIndex(state.filters.month);
  if (!state.filters.year || monthValue < 0) return [];

  return state.ahtRows
    .filter((row) => row.date instanceof Date && !Number.isNaN(row.date.getTime()))
    .filter((row) => row.date.getUTCFullYear() === Number(state.filters.year))
    .filter((row) => row.date.getUTCMonth() === monthValue)
    .filter((row) => state.filters.agent === "all" || row.agent === state.filters.agent);
}

function qaValueForWeek(row, week) {
  if (week === "week1") return row.qaWeek1;
  if (week === "week2") return row.qaWeek2;
  if (week === "week3") return row.qaWeek3;
  if (week === "week4") return row.qaWeek4;
  return row.qaTotalAverage ?? row.qaPercent;
}

function aggregateAhtRows(rows) {
  const grouped = new Map();
  rows.forEach((row) => {
    const entry = grouped.get(row.agent) || {
      agent: row.agent,
      startDate: row.date,
      endDate: row.date,
      inboundCalls: 0,
      outboundCalls: 0,
      inboundSeconds: 0,
      outboundSeconds: 0,
      talkSeconds: 0,
      holdSeconds: 0,
    };

    entry.startDate = !entry.startDate || row.date < entry.startDate ? row.date : entry.startDate;
    entry.endDate = !entry.endDate || row.date > entry.endDate ? row.date : entry.endDate;
    entry.inboundCalls += row.inboundCalls;
    entry.outboundCalls += row.outboundCalls;
    entry.inboundSeconds += row.inboundSeconds;
    entry.outboundSeconds += row.outboundSeconds;
    entry.talkSeconds += row.talkSeconds;
    entry.holdSeconds += row.holdSeconds;

    grouped.set(row.agent, entry);
  });

  return [...grouped.values()].sort((a, b) => a.agent.localeCompare(b.agent)).map((entry) => {
    const totalCalls = entry.inboundCalls + entry.outboundCalls;
    const ahtSeconds = totalCalls ? (entry.talkSeconds + entry.holdSeconds) / totalCalls : null;
    const ahtScore = calculateAhtScore(ahtSeconds);

    return {
      agent: entry.agent,
      dateRange: formatDateRange(entry.startDate, entry.endDate),
      inboundCalls: entry.inboundCalls,
      inboundSeconds: entry.inboundSeconds,
      inboundMinutes: formatDurationFromSeconds(entry.inboundSeconds),
      outboundCalls: entry.outboundCalls,
      outboundSeconds: entry.outboundSeconds,
      outboundMinutes: formatDurationFromSeconds(entry.outboundSeconds),
      talkSeconds: entry.talkSeconds,
      talkTime: formatDurationFromSeconds(entry.talkSeconds),
      holdSeconds: entry.holdSeconds,
      holdTime: formatDurationFromSeconds(entry.holdSeconds),
      ahtDisplay: formatDurationFromSeconds(ahtSeconds),
      ahtScore,
      equivalent: SCORE_LABELS[ahtScore] || "N/A",
    };
  });
}

function buildAllRows(monthlyRows, aggregatedAhtRows) {
  const ahtByAgent = new Map(aggregatedAhtRows.map((row) => [row.agent, row]));

  return monthlyRows.sort((a, b) => a.agent.localeCompare(b.agent)).map((row) => {
    const ahtScore = ahtByAgent.get(row.agent)?.ahtScore ?? row.ahtScore;
    const overallScore =
      row.attendanceScore === null || row.qaScore === null || ahtScore === null
        ? null
        : row.attendanceScore * 0.5 + row.qaScore * 0.25 + ahtScore * 0.25;

    return {
      agent: row.agent,
      attendanceScore: row.attendanceScore,
      qaScore: row.qaScore,
      ahtScore,
      overallScore,
      equivalent: equivalentFromScore(overallScore),
    };
  });
}

function getDataset() {
  const monthlyRows = getFilteredMonthlyRows();
  const ahtRows = aggregateAhtRows(getFilteredAhtRows());
  const monthlyAhtRows = aggregateAhtRows(getMonthlyAhtRows());

  if (state.activeTab === "attendance") {
    const singleAgent = state.filters.agent !== "all";
    return {
      rows: monthlyRows.map((row) => ({
        agent: row.agent,
        attendancePercentDisplay: formatPercent(row.attendancePercent),
        attendanceScore: row.attendanceScore,
        equivalent: SCORE_LABELS[row.attendanceScore] || "N/A",
      })),
      columns: [
        { key: "agent", label: "Agent" },
        { key: "attendancePercentDisplay", label: "Attendance %" },
        { key: "attendanceScore", label: "Attendance Score", type: "score" },
        { key: "equivalent", label: "Equivalent" },
      ],
      title: singleAgent ? `${state.filters.agent} Attendance Scorecard` : "Attendance Scorecards",
      subcopy: singleAgent
        ? `Monthly attendance results for ${state.filters.agent} in ${formatMonthLabel(state.filters.year, state.filters.month)}.`
        : `Monthly attendance results for ${formatMonthLabel(state.filters.year, state.filters.month)}.`,
      chartTitle: singleAgent ? "Selected Agent Attendance Snapshot" : "Attendance by Agent",
      chartSubcopy: singleAgent ? `Monthly attendance profile for ${state.filters.agent}.` : "Monthly attendance percentage per visible agent.",
      distributionTitle: singleAgent ? "Attendance Snapshot" : "Attendance Score Bands",
      distributionSubcopy: singleAgent ? `Focused attendance view for ${state.filters.agent}.` : "How the current attendance rows are distributed by score.",
    };
  }

  if (state.activeTab === "qa") {
    const singleAgent = state.filters.agent !== "all";
    const rows = monthlyRows.map((row) => {
      const qaValue = qaValueForWeek(row, state.filters.week);
      const qaScore = calculateQaScore(qaValue);
      return {
        agent: row.agent,
        qaAverageDisplay: formatPercent(qaValue),
        totalAverageDisplay: formatPercent(row.qaTotalAverage ?? row.qaPercent),
        qaScore,
        equivalent: SCORE_LABELS[qaScore] || "N/A",
      };
    });

    const columns = [
      { key: "agent", label: "Agent" },
      { key: "qaAverageDisplay", label: "QA AVG" },
    ];
    if (state.filters.week === "all") {
      columns.push({ key: "totalAverageDisplay", label: "TTL AVG" });
    }
    columns.push({ key: "qaScore", label: "QA Score", type: "score" });
    columns.push({ key: "equivalent", label: "Equivalent" });

    return {
      rows,
      columns,
      title: singleAgent ? `${state.filters.agent} QA Scorecard` : "Quality Assurance Scorecards",
      subcopy: state.qaWeeklyAvailable
        ? singleAgent
          ? `${state.filters.week === "all" ? "All-week QA totals" : state.filters.week.replace("week", "Week ")} for ${state.filters.agent} in ${formatMonthLabel(state.filters.year, state.filters.month)}.`
          : `${state.filters.week === "all" ? "All-week QA totals" : state.filters.week.replace("week", "Week ")} for ${formatMonthLabel(state.filters.year, state.filters.month)}.`
        : `Weekly QA columns are not present in the published source, so this view is using the monthly QA total for ${formatMonthLabel(state.filters.year, state.filters.month)}.`,
      chartTitle: singleAgent ? "Selected Agent QA Snapshot" : "QA Average by Agent",
      chartSubcopy: singleAgent ? `QA profile for ${state.filters.agent}.` : "Selected QA average per visible agent.",
      distributionTitle: singleAgent ? "QA Snapshot" : "QA Score Bands",
      distributionSubcopy: singleAgent ? `Focused QA view for ${state.filters.agent}.` : "How the current QA rows are distributed by score.",
    };
  }

  if (state.activeTab === "aht") {
    const singleAgent = state.filters.agent !== "all";
    return {
      rows: ahtRows,
      columns: [
        { key: "agent", label: "Agent" },
        { key: "dateRange", label: "Date Range" },
        { key: "inboundCalls", label: "Inbound Calls" },
        { key: "inboundMinutes", label: "Inbound Minutes" },
        { key: "outboundCalls", label: "Outbound Calls" },
        { key: "outboundMinutes", label: "Outbound Minutes" },
        { key: "talkTime", label: "Talk Time" },
        { key: "holdTime", label: "Hold Time" },
        { key: "ahtDisplay", label: "AHT" },
        { key: "ahtScore", label: "AHT Score", type: "score" },
        { key: "equivalent", label: "Equivalent" },
      ],
      title: singleAgent ? `${state.filters.agent} AHT Scorecard` : "AHT Scorecards",
      subcopy: singleAgent
        ? `AHT results for ${state.filters.agent} from ${formatDateRange(
          state.filters.ahtStart ? new Date(`${state.filters.ahtStart}T00:00:00Z`) : null,
          state.filters.ahtEnd ? new Date(`${state.filters.ahtEnd}T00:00:00Z`) : null
        )}.`
        : `AHT results for ${formatDateRange(
          state.filters.ahtStart ? new Date(`${state.filters.ahtStart}T00:00:00Z`) : null,
          state.filters.ahtEnd ? new Date(`${state.filters.ahtEnd}T00:00:00Z`) : null
        )}.`,
      chartTitle: singleAgent ? "Selected Agent AHT Snapshot" : "AHT by Agent",
      chartSubcopy: singleAgent ? `AHT profile for ${state.filters.agent} in the selected date range.` : "Average handle time per visible agent in the selected date range.",
      distributionTitle: singleAgent ? "AHT Snapshot" : "AHT Score Bands",
      distributionSubcopy: singleAgent ? `Focused AHT view for ${state.filters.agent}.` : "How the current AHT rows are distributed by score.",
    };
  }

  const allRows = buildAllRows(monthlyRows, monthlyAhtRows);
  const singleAgentAll = state.filters.agent !== "all";

  return {
    rows: allRows,
    columns: [
      { key: "agent", label: "Agent" },
      { key: "attendanceScore", label: "Attendance Score", type: "score" },
      { key: "qaScore", label: "QA Score", type: "score" },
      { key: "ahtScore", label: "AHT Score", type: "score" },
      { key: "overallScore", label: "Overall Score", type: "score" },
      { key: "equivalent", label: "Equivalent" },
    ],
    title: singleAgentAll ? `${state.filters.agent} KPI Scorecard` : "All KPI Scorecards",
    subcopy: singleAgentAll
      ? `Weighted monthly scorecard for ${state.filters.agent} in ${formatMonthLabel(state.filters.year, state.filters.month)} using Attendance, QA, and AHT.`
      : `Weighted monthly scorecards for ${formatMonthLabel(state.filters.year, state.filters.month)} using Attendance, QA, and AHT.`,
    chartTitle: singleAgentAll ? "Selected Agent KPI Snapshot" : "Overall Score by Agent",
    chartSubcopy: singleAgentAll ? `Overall and component scores for ${state.filters.agent}.` : "Weighted score per visible agent.",
    distributionTitle: singleAgentAll ? "Performance Snapshot" : "Overall Score Bands",
    distributionSubcopy: singleAgentAll ? `Focused view for ${state.filters.agent}.` : "How the current overall rows are distributed by score.",
  };
}

function renderSummaryCards(dataset) {
  const cards = [];
  const previousMonth = previousMonthContext(state.filters.year, state.filters.month);

  if (state.activeTab === "attendance") {
    if (isSingleAgentView(dataset) && dataset.rows[0]) {
      const row = dataset.rows[0];
      const previousRow = previousMonth ? getMonthlyRowsForPeriod(previousMonth.year, previousMonth.month, state.filters.agent)[0] : null;
      cards.push({ label: "Attendance %", value: row.attendancePercentDisplay, note: "Monthly attendance percentage", tone: "attendance", trend: buildTrendChip(safeNumber(row.attendancePercentDisplay), previousRow?.attendancePercent, { mode: "number", scopeLabel: `${previousMonth?.month || "previous"} ${previousMonth?.year || "period"}` }) });
      cards.push({ label: "Attendance Score", value: formatNumber(row.attendanceScore), note: "Monthly attendance score", tone: "attendance", trend: buildTrendChip(row.attendanceScore, previousRow?.attendanceScore, { mode: "number", scopeLabel: `${previousMonth?.month || "previous"} ${previousMonth?.year || "period"}` }) });
      cards.push({ label: "Equivalent", value: row.equivalent, note: "Attendance performance label", tone: "attendance", trend: buildTrendChip(row.attendanceScore, previousRow?.attendanceScore, { mode: "number", scopeLabel: `${previousMonth?.month || "previous"} ${previousMonth?.year || "period"}` }) });
    } else {
      const previousRows = previousMonth ? getMonthlyRowsForPeriod(previousMonth.year, previousMonth.month) : [];
      cards.push({ label: "Average Attendance %", value: formatPercent(average(getFilteredMonthlyRows().map((row) => row.attendancePercent))), note: "Monthly attendance average", tone: "attendance", trend: buildTrendChip(average(getFilteredMonthlyRows().map((row) => row.attendancePercent)), average(previousRows.map((row) => row.attendancePercent)), { mode: "number", scopeLabel: `${previousMonth?.month || "previous"} ${previousMonth?.year || "period"}` }) });
      cards.push({ label: "Average Attendance Score", value: formatNumber(average(dataset.rows.map((row) => row.attendanceScore))), note: "Team attendance score", tone: "attendance", trend: buildTrendChip(average(dataset.rows.map((row) => row.attendanceScore)), average(previousRows.map((row) => row.attendanceScore)), { mode: "number", scopeLabel: `${previousMonth?.month || "previous"} ${previousMonth?.year || "period"}` }) });
      cards.push({ label: "Visible Agents", value: String(dataset.rows.length), note: "Rows in the current table", tone: "attendance" });
    }
  } else if (state.activeTab === "qa") {
    if (isSingleAgentView(dataset) && dataset.rows[0]) {
      const row = dataset.rows[0];
      const previousRow = previousMonth ? getMonthlyRowsForPeriod(previousMonth.year, previousMonth.month, state.filters.agent)[0] : null;
      const previousQaValue = previousRow ? qaValueForWeek(previousRow, state.filters.week) : null;
      cards.push({ label: state.filters.week === "all" ? "QA AVG" : formatWeekLabel(state.filters.week), value: row.qaAverageDisplay, note: "Selected QA average", tone: "qa", trend: buildTrendChip(safeNumber(row.qaAverageDisplay), previousQaValue, { mode: "number", scopeLabel: `${previousMonth?.month || "previous"} ${previousMonth?.year || "period"}` }) });
      cards.push({ label: "QA Score", value: formatNumber(row.qaScore), note: "QA score for the selected scope", tone: "qa", trend: buildTrendChip(row.qaScore, previousRow ? calculateQaScore(previousQaValue) : null, { mode: "number", scopeLabel: `${previousMonth?.month || "previous"} ${previousMonth?.year || "period"}` }) });
      cards.push({ label: "Equivalent", value: row.equivalent, note: "QA performance label", tone: "qa", trend: buildTrendChip(row.qaScore, previousRow ? calculateQaScore(previousQaValue) : null, { mode: "number", scopeLabel: `${previousMonth?.month || "previous"} ${previousMonth?.year || "period"}` }) });
    } else {
      const qaValues = getFilteredMonthlyRows().map((row) => qaValueForWeek(row, state.filters.week));
      const previousRows = previousMonth ? getMonthlyRowsForPeriod(previousMonth.year, previousMonth.month) : [];
      const previousQaValues = previousRows.map((row) => qaValueForWeek(row, state.filters.week));
      cards.push({ label: "Average QA %", value: formatPercent(average(qaValues)), note: "QA average for the selected scope", tone: "qa", trend: buildTrendChip(average(qaValues), average(previousQaValues), { mode: "number", scopeLabel: `${previousMonth?.month || "previous"} ${previousMonth?.year || "period"}` }) });
      cards.push({ label: "Average QA Score", value: formatNumber(average(dataset.rows.map((row) => row.qaScore))), note: "Team QA score", tone: "qa", trend: buildTrendChip(average(dataset.rows.map((row) => row.qaScore)), average(previousQaValues.map((value) => calculateQaScore(value))), { mode: "number", scopeLabel: `${previousMonth?.month || "previous"} ${previousMonth?.year || "period"}` }) });
      cards.push({ label: "Visible Agents", value: String(dataset.rows.length), note: "Rows in the current table", tone: "qa" });
    }
  } else if (state.activeTab === "aht") {
    if (isSingleAgentView(dataset) && dataset.rows[0]) {
      const row = dataset.rows[0];
      const previousRow = aggregateAhtRows(getPreviousAhtRangeRows())[0] || null;
      cards.push({ label: "AHT", value: row.ahtDisplay, note: "Average handle time in range", tone: "aht", trend: buildTrendChip(durationToSeconds(row.ahtDisplay), previousRow ? durationToSeconds(previousRow.ahtDisplay) : null, { higherIsBetter: false, mode: "duration", scopeLabel: "the previous matching range" }) });
      cards.push({ label: "AHT Score", value: formatNumber(row.ahtScore), note: "AHT score for the selected scope", tone: "aht", trend: buildTrendChip(row.ahtScore, previousRow?.ahtScore, { mode: "number", scopeLabel: "the previous matching range" }) });
      cards.push({ label: "Equivalent", value: row.equivalent, note: "AHT performance label", tone: "aht", trend: buildTrendChip(row.ahtScore, previousRow?.ahtScore, { mode: "number", scopeLabel: "the previous matching range" }) });
      cards.push({ label: "Total Calls", value: String((row.inboundCalls ?? 0) + (row.outboundCalls ?? 0)), note: "Calls in the selected date range", tone: "aht" });
    } else {
      const ahtSeconds = dataset.rows.map((row) => durationToSeconds(row.ahtDisplay));
      const previousRows = aggregateAhtRows(getPreviousAhtRangeRows());
      cards.push({ label: "Average AHT", value: formatDurationFromSeconds(average(ahtSeconds)), note: "Average handle time in range", tone: "aht", trend: buildTrendChip(average(ahtSeconds), average(previousRows.map((row) => durationToSeconds(row.ahtDisplay))), { higherIsBetter: false, mode: "duration", scopeLabel: "the previous matching range" }) });
      cards.push({ label: "Average AHT Score", value: formatNumber(average(dataset.rows.map((row) => row.ahtScore))), note: "Team AHT score", tone: "aht", trend: buildTrendChip(average(dataset.rows.map((row) => row.ahtScore)), average(previousRows.map((row) => row.ahtScore)), { mode: "number", scopeLabel: "the previous matching range" }) });
      cards.push({ label: "Visible Agents", value: String(dataset.rows.length), note: "Rows in the current table", tone: "aht" });
    }
  } else if (isSingleAgentAllView(dataset) && dataset.rows[0]) {
    const row = dataset.rows[0];
    const previousRows = previousMonth ? getMonthlyRowsForPeriod(previousMonth.year, previousMonth.month, state.filters.agent) : [];
    const previousMonthlyAhtRows = previousMonth ? aggregateAhtRows(getMonthlyAhtRowsForPeriod(previousMonth.year, previousMonth.month, state.filters.agent)) : [];
    const previousAllRow = buildAllRows(previousRows, previousMonthlyAhtRows)[0] || null;
    cards.push({ label: "Overall Score", value: formatNumber(row.overallScore), note: row.equivalent || "Selected agent overall score", tone: "overall", trend: buildTrendChip(row.overallScore, previousAllRow?.overallScore, { mode: "number", scopeLabel: `${previousMonth?.month || "previous"} ${previousMonth?.year || "period"}` }) });
    cards.push({ label: "Attendance Score", value: formatNumber(row.attendanceScore), note: "Monthly attendance score", tone: "attendance", trend: buildTrendChip(row.attendanceScore, previousAllRow?.attendanceScore, { mode: "number", scopeLabel: `${previousMonth?.month || "previous"} ${previousMonth?.year || "period"}` }) });
    cards.push({ label: "QA Score", value: formatNumber(row.qaScore), note: "Monthly quality score", tone: "qa", trend: buildTrendChip(row.qaScore, previousAllRow?.qaScore, { mode: "number", scopeLabel: `${previousMonth?.month || "previous"} ${previousMonth?.year || "period"}` }) });
    cards.push({ label: "AHT Score", value: formatNumber(row.ahtScore), note: "Monthly AHT score", tone: "aht", trend: buildTrendChip(row.ahtScore, previousAllRow?.ahtScore, { mode: "number", scopeLabel: `${previousMonth?.month || "previous"} ${previousMonth?.year || "period"}` }) });
  } else {
    const previousRows = previousMonth ? getMonthlyRowsForPeriod(previousMonth.year, previousMonth.month) : [];
    const previousMonthlyAhtRows = previousMonth ? aggregateAhtRows(getMonthlyAhtRowsForPeriod(previousMonth.year, previousMonth.month)) : [];
    const previousAllRows = buildAllRows(previousRows, previousMonthlyAhtRows);
    cards.push({ label: "Average Overall Score", value: formatNumber(average(dataset.rows.map((row) => row.overallScore))), note: "Weighted score across visible agents", tone: "overall", trend: buildTrendChip(average(dataset.rows.map((row) => row.overallScore)), average(previousAllRows.map((row) => row.overallScore)), { mode: "number", scopeLabel: `${previousMonth?.month || "previous"} ${previousMonth?.year || "period"}` }) });
    cards.push({ label: "Average Attendance Score", value: formatNumber(average(dataset.rows.map((row) => row.attendanceScore))), note: "Monthly attendance score", tone: "attendance", trend: buildTrendChip(average(dataset.rows.map((row) => row.attendanceScore)), average(previousAllRows.map((row) => row.attendanceScore)), { mode: "number", scopeLabel: `${previousMonth?.month || "previous"} ${previousMonth?.year || "period"}` }) });
    cards.push({ label: "Average QA Score", value: formatNumber(average(dataset.rows.map((row) => row.qaScore))), note: "Monthly quality score", tone: "qa", trend: buildTrendChip(average(dataset.rows.map((row) => row.qaScore)), average(previousAllRows.map((row) => row.qaScore)), { mode: "number", scopeLabel: `${previousMonth?.month || "previous"} ${previousMonth?.year || "period"}` }) });
    cards.push({ label: "Average AHT Score", value: formatNumber(average(dataset.rows.map((row) => row.ahtScore))), note: "Monthly AHT score", tone: "aht", trend: buildTrendChip(average(dataset.rows.map((row) => row.ahtScore)), average(previousAllRows.map((row) => row.ahtScore)), { mode: "number", scopeLabel: `${previousMonth?.month || "previous"} ${previousMonth?.year || "period"}` }) });
  }

  elements.summaryGrid.innerHTML = "";
  cards.forEach(({ label, value, note, tone, trend }) => {
    const iconPath = SUMMARY_CARD_ICONS[tone] || SUMMARY_CARD_ICONS.overall;
    const card = document.createElement("article");
    card.className = `summary-card card card-${tone}`;
    card.dataset.richTooltip = richTooltipData({ title: label, lines: [note] });
    card.innerHTML = `
      <div class="card-topline">
        <span class="summary-card-icon" aria-hidden="true">
          <img src="${iconPath}" alt="" />
        </span>
        <span>${label}</span>
      </div>
      <div class="summary-card-value-row">
        <strong>${value}</strong>
        ${trend ? `<span class="trend-chip trend-chip-${trend.tone}" data-rich-tooltip="${richTooltipData({ title: trend.text, lines: [trend.tooltip] })}">${escapeHtml(trend.text)}</span>` : ""}
      </div>
      <p>${note}</p>
    `;
    elements.summaryGrid.appendChild(card);
  });
}

function buildDistribution(rows, scoreKey) {
  return [1, 2, 3, 4, 5].map((score) => ({
    score,
    count: rows.filter((row) => Math.round(row[scoreKey] ?? 0) === score).length,
    agents: rows
      .filter((row) => Math.round(row[scoreKey] ?? 0) === score)
      .map((row) => row.agent),
  }));
}

function distributionTooltip(entry) {
  return `${entry.score} - ${SCORE_LABELS[entry.score]}: ${entry.agents.length ? entry.agents.join(", ") : "No agents"}`;
}

function renderDistributionDetail(panel, entry, color) {
  if (!panel) return;

  panel.innerHTML = `
    <div class="donut-detail-title">
      <span class="compact-legend-swatch" style="background:${color}"></span>
      <span>${entry.score} - ${escapeHtml(SCORE_LABELS[entry.score])}</span>
    </div>
    <p class="donut-detail-summary">${entry.count} ${entry.count === 1 ? "agent" : "agents"} in this band</p>
    ${
      entry.agents.length
        ? `<div class="donut-agents">
            ${entry.agents
              .map(
                (agent) => `
                  <span class="donut-agent-pill">${escapeHtml(agent)}</span>
                `
              )
              .join("")}
          </div>`
        : '<p class="donut-empty-note">No visible agents in this score band.</p>'
    }
  `;
}

function getActiveMetricKey() {
  if (state.activeTab === "attendance") return "attendanceScore";
  if (state.activeTab === "qa") return "qaScore";
  if (state.activeTab === "aht") return "ahtScore";
  return "overallScore";
}

function isSingleAgentView(dataset = null) {
  if (state.filters.agent === "all") return false;
  if (!dataset) return true;
  return dataset.rows.length <= 1;
}

function isSingleAgentAllView(dataset = null) {
  return state.activeTab === "all" && isSingleAgentView(dataset);
}

function strongestAndWeakestKpi(row) {
  const metrics = [
    { label: "Attendance", value: row.attendanceScore ?? null },
    { label: "QA", value: row.qaScore ?? null },
    { label: "AHT", value: row.ahtScore ?? null },
  ].filter((metric) => typeof metric.value === "number" && !Number.isNaN(metric.value));

  if (!metrics.length) {
    return { strongest: null, weakest: null };
  }

  const strongest = [...metrics].sort((a, b) => b.value - a.value)[0];
  const weakest = [...metrics].sort((a, b) => a.value - b.value)[0];
  return { strongest, weakest };
}

function formatWeekLabel(week) {
  if (week === "all") return "All Weeks";
  return week.replace("week", "Week ");
}

function buildEmptyStateCopy() {
  const agentLabel = state.filters.agent === "all" ? "" : ` for ${state.filters.agent}`;

  if (state.activeTab === "attendance") {
    const scope = `${formatMonthLabel(state.filters.year, state.filters.month)}${agentLabel}`;
    return {
      chart: `No attendance chart data found for ${scope}.`,
      table: `No attendance rows found for ${scope}.`,
    };
  }

  if (state.activeTab === "qa") {
    const weekLabel = state.filters.week === "all" ? "" : ` • ${formatWeekLabel(state.filters.week)}`;
    const scope = `${formatMonthLabel(state.filters.year, state.filters.month)}${weekLabel}${agentLabel}`;
    return {
      chart: `No QA chart data found for ${scope}.`,
      table: `No QA rows found for ${scope}.`,
    };
  }

  if (state.activeTab === "aht") {
    const range = formatDateRange(
      state.filters.ahtStart ? new Date(`${state.filters.ahtStart}T00:00:00Z`) : null,
      state.filters.ahtEnd ? new Date(`${state.filters.ahtEnd}T00:00:00Z`) : null
    );
    const scope = `${range}${agentLabel}`;
    return {
      chart: `No AHT chart data found for ${scope}.`,
      table: `No AHT rows found for ${scope}.`,
    };
  }

  const scope = `${formatMonthLabel(state.filters.year, state.filters.month)}${agentLabel}`;
  return {
    chart: `No KPI chart data found for ${scope}.`,
    table: `No KPI rows found for ${scope}.`,
  };
}

function buildFilterSummaryLine() {
  const parts = [];
  const monthLabel = formatMonthLabel(state.filters.year, state.filters.month);
  const agentLabel = state.filters.agent === "all" ? "All Agents" : state.filters.agent;

  if (state.activeTab === "aht") {
    parts.push(formatDateRange(
      state.filters.ahtStart ? new Date(`${state.filters.ahtStart}T00:00:00Z`) : null,
      state.filters.ahtEnd ? new Date(`${state.filters.ahtEnd}T00:00:00Z`) : null
    ));
    parts.push(agentLabel);
    parts.push("AHT Range");
    return parts.filter(Boolean).join(" • ");
  }

  parts.push(monthLabel);
  if (state.activeTab === "qa") {
    parts.push(formatWeekLabel(state.filters.week));
  }
  parts.push(agentLabel);
  if (state.activeTab === "all") {
    parts.push("Monthly KPI Scope");
  }
  return parts.filter(Boolean).join(" • ");
}

function rowHighlightTone(row) {
  const label = String(row?.equivalent ?? "").trim().toLowerCase();
  if (label === "poor") return "poor";
  if (label === "below expectations") return "below";
  return "";
}

function renderSingleAgentOverview(dataset) {
  if (!elements.singleAgentOverview) return;

  const shouldShow = isSingleAgentView(dataset);
  elements.singleAgentOverview.classList.toggle("single-agent-overview-all", shouldShow && state.activeTab === "all");
  elements.singleAgentOverview.classList.toggle("single-agent-overview-attendance", shouldShow && state.activeTab === "attendance");
  elements.singleAgentOverview.classList.toggle("single-agent-overview-qa", shouldShow && state.activeTab === "qa");
  elements.singleAgentOverview.classList.toggle("single-agent-overview-aht", shouldShow && state.activeTab === "aht");
  elements.singleAgentOverview.hidden = !shouldShow;
  if (!shouldShow) {
    elements.singleAgentOverview.innerHTML = "";
    return;
  }

  const row = dataset.rows[0];
  const monthlyRow = getFilteredMonthlyRows()[0] || null;
  const ahtRow = aggregateAhtRows(getFilteredAhtRows())[0] || null;
  const monthlyAhtRow = aggregateAhtRows(getMonthlyAhtRows())[0] || null;
  if (!row) {
    elements.singleAgentOverview.innerHTML = "";
    return;
  }

  const { strongest, weakest } = strongestAndWeakestKpi(row);
  const qaAverage = monthlyRow ? qaValueForWeek(monthlyRow, "all") : null;
  const totalCalls = ahtRow ? ahtRow.inboundCalls + ahtRow.outboundCalls : 0;
  const monthlyAhtTotalCalls = monthlyAhtRow ? monthlyAhtRow.inboundCalls + monthlyAhtRow.outboundCalls : 0;
  const selectedDateRange = formatDateRange(
    state.filters.ahtStart ? new Date(`${state.filters.ahtStart}T00:00:00Z`) : null,
    state.filters.ahtEnd ? new Date(`${state.filters.ahtEnd}T00:00:00Z`) : null
  );

  if (state.activeTab === "attendance") {
    elements.singleAgentOverview.innerHTML = `
      <article class="chart-card card single-agent-card single-agent-card-wide single-agent-card-attendance-summary">
        <div class="section-heading section-heading-compact">
          <div>
            <p class="eyebrow">Selected Agent</p>
            <h3>Attendance Summary</h3>
            <p class="chart-subnote">Monthly attendance profile for ${escapeHtml(row.agent)}.</p>
          </div>
        </div>
        <div class="single-agent-hero">
          <div class="single-agent-hero-score">
            <span class="single-agent-hero-label">Attendance Score</span>
            ${scoreBadge(row.attendanceScore)}
          </div>
          <div class="single-agent-hero-equivalent">
            <span class="single-agent-hero-label">Equivalent</span>
            ${statusBadge(row.equivalent)}
          </div>
          <div class="single-agent-hero-note">
            <span class="single-agent-hero-label">Attendance %</span>
            <strong>${escapeHtml(row.attendancePercentDisplay || "N/A")}</strong>
          </div>
          <div class="single-agent-hero-note">
            <span class="single-agent-hero-label">Selected Month</span>
            <strong>${escapeHtml(formatMonthLabel(state.filters.year, state.filters.month))}</strong>
          </div>
        </div>
      </article>

      <article class="chart-card card single-agent-card single-agent-card-attendance-scope">
        <div class="section-heading section-heading-compact">
          <div>
            <p class="eyebrow">Monthly Detail</p>
            <h3>Selected Scope</h3>
          </div>
        </div>
        <div class="single-agent-detail-list">
          <div class="single-agent-detail-row">
            <span>Agent</span>
            <strong>${escapeHtml(row.agent)}</strong>
          </div>
          <div class="single-agent-detail-row">
            <span>Month</span>
            <strong>${escapeHtml(formatMonthLabel(state.filters.year, state.filters.month))}</strong>
          </div>
        </div>
      </article>

      <article class="chart-card card single-agent-card single-agent-card-attendance-detail">
        <div class="section-heading section-heading-compact">
          <div>
            <p class="eyebrow">Attendance Detail</p>
            <h3>Performance Snapshot</h3>
          </div>
        </div>
        <div class="single-agent-stat-grid">
          <div class="single-agent-stat">
            <span class="single-agent-stat-label">Attendance %</span>
            <strong>${escapeHtml(row.attendancePercentDisplay || "N/A")}</strong>
          </div>
          <div class="single-agent-stat">
            <span class="single-agent-stat-label">Score</span>
            <strong>${formatNumber(row.attendanceScore)}</strong>
          </div>
        </div>
      </article>
    `;
    return;
  }

  if (state.activeTab === "qa") {
    elements.singleAgentOverview.innerHTML = `
      <article class="chart-card card single-agent-card single-agent-card-wide single-agent-card-qa-summary">
        <div class="section-heading section-heading-compact">
          <div>
            <p class="eyebrow">Selected Agent</p>
            <h3>Quality Summary</h3>
            <p class="chart-subnote">QA profile for ${escapeHtml(row.agent)}.</p>
          </div>
        </div>
        <div class="single-agent-hero">
          <div class="single-agent-hero-score">
            <span class="single-agent-hero-label">QA Score</span>
            ${scoreBadge(row.qaScore)}
          </div>
          <div class="single-agent-hero-equivalent">
            <span class="single-agent-hero-label">Equivalent</span>
            ${statusBadge(row.equivalent)}
          </div>
          <div class="single-agent-hero-note">
            <span class="single-agent-hero-label">${escapeHtml(state.filters.week === "all" ? "QA AVG" : formatWeekLabel(state.filters.week))}</span>
            <strong>${escapeHtml(row.qaAverageDisplay || "N/A")}</strong>
          </div>
          <div class="single-agent-hero-note">
            <span class="single-agent-hero-label">Scope</span>
            <strong>${escapeHtml(`${formatMonthLabel(state.filters.year, state.filters.month)} | ${formatWeekLabel(state.filters.week)}`)}</strong>
          </div>
        </div>
      </article>

      <article class="chart-card card single-agent-card single-agent-card-qa-scope">
        <div class="section-heading section-heading-compact">
          <div>
            <p class="eyebrow">Selected Scope</p>
            <h3>Month and Week</h3>
          </div>
        </div>
        <div class="single-agent-detail-list">
          <div class="single-agent-detail-row">
            <span>Month</span>
            <strong>${escapeHtml(formatMonthLabel(state.filters.year, state.filters.month))}</strong>
          </div>
          <div class="single-agent-detail-row">
            <span>Week</span>
            <strong>${escapeHtml(formatWeekLabel(state.filters.week))}</strong>
          </div>
        </div>
      </article>

      <article class="chart-card card single-agent-card single-agent-card-qa-detail">
        <div class="section-heading section-heading-compact">
          <div>
            <p class="eyebrow">QA Detail</p>
            <h3>Average Snapshot</h3>
          </div>
        </div>
        <div class="single-agent-stat-grid">
          <div class="single-agent-stat">
            <span class="single-agent-stat-label">${escapeHtml(state.filters.week === "all" ? "QA AVG" : formatWeekLabel(state.filters.week))}</span>
            <strong>${escapeHtml(row.qaAverageDisplay || "N/A")}</strong>
          </div>
          <div class="single-agent-stat">
            <span class="single-agent-stat-label">TTL AVG</span>
            <strong>${escapeHtml(row.totalAverageDisplay || "N/A")}</strong>
          </div>
        </div>
      </article>
    `;
    return;
  }

  if (state.activeTab === "aht") {
    elements.singleAgentOverview.innerHTML = `
      <article class="chart-card card single-agent-card single-agent-card-wide single-agent-card-aht-summary">
        <div class="section-heading section-heading-compact">
          <div>
            <p class="eyebrow">Selected Agent</p>
            <h3>AHT Summary</h3>
            <p class="chart-subnote">Date-range AHT profile for ${escapeHtml(row.agent)}.</p>
          </div>
        </div>
        <div class="single-agent-hero">
          <div class="single-agent-hero-score">
            <span class="single-agent-hero-label">AHT Score</span>
            ${scoreBadge(row.ahtScore)}
          </div>
          <div class="single-agent-hero-equivalent">
            <span class="single-agent-hero-label">Equivalent</span>
            ${statusBadge(row.equivalent)}
          </div>
          <div class="single-agent-hero-note">
            <span class="single-agent-hero-label">AHT</span>
            <strong>${escapeHtml(row.ahtDisplay || "N/A")}</strong>
          </div>
          <div class="single-agent-hero-note">
            <span class="single-agent-hero-label">Date Range</span>
            <strong>${escapeHtml(row.dateRange || selectedDateRange)}</strong>
          </div>
        </div>
      </article>

      <article class="chart-card card single-agent-card single-agent-card-aht-volume">
        <div class="section-heading section-heading-compact">
          <div>
            <p class="eyebrow">Call Volume</p>
            <h3>Inbound and Outbound</h3>
          </div>
        </div>
        <div class="single-agent-stat-grid">
          <div class="single-agent-stat">
            <span class="single-agent-stat-label">Inbound Calls</span>
            <strong>${row.inboundCalls ?? 0}</strong>
          </div>
          <div class="single-agent-stat">
            <span class="single-agent-stat-label">Outbound Calls</span>
            <strong>${row.outboundCalls ?? 0}</strong>
          </div>
        </div>
      </article>

      <article class="chart-card card single-agent-card single-agent-card-aht-time">
        <div class="section-heading section-heading-compact">
          <div>
            <p class="eyebrow">Time Detail</p>
            <h3>Talk and Hold</h3>
          </div>
        </div>
        <div class="single-agent-stat-grid">
          <div class="single-agent-stat">
            <span class="single-agent-stat-label">Talk Time</span>
            <strong>${escapeHtml(row.talkTime || "N/A")}</strong>
          </div>
          <div class="single-agent-stat">
            <span class="single-agent-stat-label">Hold Time</span>
            <strong>${escapeHtml(row.holdTime || "N/A")}</strong>
          </div>
        </div>
      </article>

      <article class="chart-card card single-agent-card single-agent-card-wide single-agent-card-aht-detail">
        <div class="section-heading section-heading-compact">
          <div>
            <p class="eyebrow">Range Snapshot</p>
            <h3>AHT Detail</h3>
            <p class="chart-subnote">${escapeHtml(row.dateRange || selectedDateRange)}</p>
          </div>
        </div>
        <div class="single-agent-stat-grid single-agent-stat-grid-wide">
          <div class="single-agent-stat">
            <span class="single-agent-stat-label">Inbound Minutes</span>
            <strong>${escapeHtml(row.inboundMinutes || "N/A")}</strong>
          </div>
          <div class="single-agent-stat">
            <span class="single-agent-stat-label">Outbound Minutes</span>
            <strong>${escapeHtml(row.outboundMinutes || "N/A")}</strong>
          </div>
          <div class="single-agent-stat">
            <span class="single-agent-stat-label">Total Calls</span>
            <strong>${(row.inboundCalls ?? 0) + (row.outboundCalls ?? 0)}</strong>
          </div>
          <div class="single-agent-stat">
            <span class="single-agent-stat-label">AHT</span>
            <strong>${escapeHtml(row.ahtDisplay || "N/A")}</strong>
          </div>
        </div>
      </article>
    `;
    return;
  }

  elements.singleAgentOverview.innerHTML = `
    <article class="chart-card card single-agent-card single-agent-card-wide single-agent-card-all-summary">
      <div class="section-heading section-heading-compact">
        <div>
          <p class="eyebrow">Selected Agent</p>
          <h3>Performance Summary</h3>
          <p class="chart-subnote">Focused profile for ${escapeHtml(row.agent)}.</p>
        </div>
      </div>
      <div class="single-agent-hero">
        <div class="single-agent-hero-score">
          <span class="single-agent-hero-label">Overall Score</span>
          ${scoreBadge(row.overallScore)}
        </div>
        <div class="single-agent-hero-equivalent">
          <span class="single-agent-hero-label">Equivalent</span>
          ${statusBadge(row.equivalent)}
        </div>
        <div class="single-agent-hero-note">
          <span class="single-agent-hero-label">Strongest KPI</span>
          <strong>${strongest ? `${escapeHtml(strongest.label)} (${formatNumber(strongest.value)})` : "N/A"}</strong>
        </div>
        <div class="single-agent-hero-note">
          <span class="single-agent-hero-label">Needs Attention</span>
          <strong>${weakest ? `${escapeHtml(weakest.label)} (${formatNumber(weakest.value)})` : "N/A"}</strong>
        </div>
      </div>
    </article>

    <article class="chart-card card single-agent-card single-agent-card-all-breakdown">
      <div class="section-heading section-heading-compact">
        <div>
          <p class="eyebrow">KPI Detail</p>
          <h3>Score Breakdown</h3>
        </div>
      </div>
      <div class="single-agent-detail-list">
        <div class="single-agent-detail-row">
          <span>Attendance Score</span>
          ${scoreBadge(row.attendanceScore)}
        </div>
        <div class="single-agent-detail-row">
          <span>QA Score</span>
          ${scoreBadge(row.qaScore)}
        </div>
        <div class="single-agent-detail-row">
          <span>AHT Score</span>
          ${scoreBadge(row.ahtScore)}
        </div>
      </div>
    </article>

    <article class="chart-card card single-agent-card single-agent-card-all-monthly">
      <div class="section-heading section-heading-compact">
        <div>
          <p class="eyebrow">Monthly Detail</p>
          <h3>Attendance and QA</h3>
        </div>
      </div>
      <div class="single-agent-stat-grid">
        <div class="single-agent-stat">
          <span class="single-agent-stat-label">Attendance %</span>
          <strong>${formatPercent(monthlyRow?.attendancePercent)}</strong>
        </div>
        <div class="single-agent-stat">
          <span class="single-agent-stat-label">QA Average</span>
          <strong>${formatPercent(qaAverage)}</strong>
        </div>
      </div>
    </article>

    <article class="chart-card card single-agent-card single-agent-card-wide single-agent-card-all-aht">
      <div class="section-heading section-heading-compact">
        <div>
          <p class="eyebrow">AHT Detail</p>
          <h3>Monthly Snapshot</h3>
          <p class="chart-subnote">${escapeHtml(formatMonthLabel(state.filters.year, state.filters.month))}</p>
        </div>
      </div>
      <div class="single-agent-stat-grid single-agent-stat-grid-wide">
        <div class="single-agent-stat">
          <span class="single-agent-stat-label">Total Calls</span>
          <strong>${monthlyAhtTotalCalls}</strong>
        </div>
        <div class="single-agent-stat">
          <span class="single-agent-stat-label">Talk Time</span>
          <strong>${monthlyAhtRow?.talkTime || "N/A"}</strong>
        </div>
        <div class="single-agent-stat">
          <span class="single-agent-stat-label">Hold Time</span>
          <strong>${monthlyAhtRow?.holdTime || "N/A"}</strong>
        </div>
        <div class="single-agent-stat">
          <span class="single-agent-stat-label">AHT</span>
          <strong>${monthlyAhtRow?.ahtDisplay || "N/A"}</strong>
        </div>
      </div>
    </article>
  `;
}

function renderTopBottomChart(dataset) {
  if (!elements.topBottomChart) return;
  const metricKey = getActiveMetricKey();
  const ranked = [...dataset.rows]
    .filter((row) => typeof row[metricKey] === "number" && !Number.isNaN(row[metricKey]))
    .sort((a, b) => b[metricKey] - a[metricKey]);

  const shouldShow = !isSingleAgentView(dataset) && ranked.length > 1;
  if (elements.topBottomCard) elements.topBottomCard.hidden = !shouldShow;
  if (!shouldShow) {
    elements.topBottomChart.innerHTML = "";
    return;
  }

  setText(elements.topBottomTitle, "Top / Bottom Agents");

  const top = ranked.slice(0, 3).map((row) => ({ ...row, direction: "top" }));
  const bottom = ranked.slice(-3).reverse().map((row) => ({ ...row, direction: "bottom" }));

  elements.topBottomChart.innerHTML = `
    <div class="compact-chart-legend">
      <span class="compact-legend-item"><span class="compact-legend-swatch segment-top"></span>Top performers</span>
      <span class="compact-legend-item"><span class="compact-legend-swatch segment-bottom"></span>Bottom performers</span>
    </div>
    <div class="podium-grid">
      <section class="podium-column">
        <header class="podium-column-header">
          <span class="podium-column-kicker">Top</span>
          <span class="podium-column-note">Highest scores</span>
        </header>
        <div class="podium-list">
          ${top
            .map(
              (row, index) => `
                <article class="podium-card podium-card-top" data-rich-tooltip="${richTooltipData({
                  title: row.agent,
                  lines: [{ text: `Score: ${Number(row[metricKey]).toFixed(2)}`, color: "#2f7a4c" }],
                })}">
                  <span class="podium-rank">#${index + 1}</span>
                  <div class="podium-copy">
                    <strong>${row.agent}</strong>
                    <span>${Number(row[metricKey]).toFixed(2)}</span>
                  </div>
                </article>
              `
            )
            .join("")}
        </div>
      </section>
      <section class="podium-column">
        <header class="podium-column-header">
          <span class="podium-column-kicker">Bottom</span>
          <span class="podium-column-note">Needs attention</span>
        </header>
        <div class="podium-list">
          ${bottom
            .map(
              (row, index) => `
                <article class="podium-card podium-card-bottom" data-rich-tooltip="${richTooltipData({
                  title: row.agent,
                  lines: [{ text: `Score: ${Number(row[metricKey]).toFixed(2)}`, color: "#ca5a4f" }],
                })}">
                  <span class="podium-rank">#${index + 1}</span>
                  <div class="podium-copy">
                    <strong>${row.agent}</strong>
                    <span>${Number(row[metricKey]).toFixed(2)}</span>
                  </div>
                </article>
              `
            )
            .join("")}
        </div>
      </section>
    </div>
  `;
}

function renderAhtComponentsChart() {
  if (!elements.ahtComponentsCard || !elements.ahtComponentsChart) return;
  const rows = aggregateAhtRows(getFilteredAhtRows());
  const shouldShow = (state.activeTab === "all" || state.activeTab === "aht") && !isSingleAgentView() && rows.length > 0;
  elements.ahtComponentsCard.hidden = !shouldShow;
  if (!shouldShow) {
    elements.ahtComponentsChart.innerHTML = "";
    return;
  }

  setText(elements.ahtComponentsTitle, "AHT Components");
  setText(
    elements.ahtComponentsSubcopy,
    `Talk and hold totals for ${formatDateRange(
      state.filters.ahtStart ? new Date(`${state.filters.ahtStart}T00:00:00Z`) : null,
      state.filters.ahtEnd ? new Date(`${state.filters.ahtEnd}T00:00:00Z`) : null
    )}.`
  );

  const maxTotal = Math.max(...rows.map((row) => row.talkSeconds + row.holdSeconds), 0);
  elements.ahtComponentsChart.innerHTML = `
    <div class="compact-chart-legend">
      <span class="compact-legend-item"><span class="compact-legend-swatch stacked-segment-qa"></span>Talk Time</span>
      <span class="compact-legend-item"><span class="compact-legend-swatch stacked-segment-aht"></span>Hold Time</span>
    </div>
    <div class="aht-components-chart">
      ${rows
        .map((row) => {
          const talkPercent = maxTotal > 0 ? (row.talkSeconds / maxTotal) * 100 : 0;
          const holdPercent = maxTotal > 0 ? (row.holdSeconds / maxTotal) * 100 : 0;
          return `
            <div class="aht-component-row" data-rich-tooltip="${richTooltipData({
              title: row.agent,
              lines: [
                { text: `Talk: ${row.talkTime}`, color: "#62b6e8" },
                { text: `Hold: ${row.holdTime}`, color: "#4f92a0" },
              ],
            })}">
              <span class="stacked-label">${row.agent}</span>
              <div class="aht-component-bar">
                <div class="aht-component-talk" style="width:${talkPercent}%"></div>
                <div class="aht-component-hold" style="width:${holdPercent}%"></div>
              </div>
            </div>
          `;
        })
        .join("")}
    </div>
  `;
}

function renderCallsVolumeChart() {
  if (!elements.callsVolumeCard || !elements.callsVolumeChart) return;
  const rows = aggregateAhtRows(getFilteredAhtRows());
  const shouldShow = (state.activeTab === "all" || state.activeTab === "aht") && !isSingleAgentView() && rows.length > 0;
  elements.callsVolumeCard.hidden = !shouldShow;
  if (!shouldShow) {
    elements.callsVolumeChart.innerHTML = "";
    return;
  }

  setText(elements.callsVolumeTitle, "Calls Volume Snapshot");

  const maxCalls = Math.max(...rows.map((row) => row.inboundCalls + row.outboundCalls), 0);
  elements.callsVolumeChart.innerHTML = `
    <div class="compact-chart-legend">
      <span class="compact-legend-item"><span class="compact-legend-swatch calls-volume-stem-swatch"></span>Total calls</span>
      <span class="compact-legend-item"><span class="compact-legend-swatch calls-volume-dot-swatch"></span>Inbound share</span>
    </div>
    <div class="calls-volume-chart">
      ${rows
        .map((row) => {
          const totalCalls = row.inboundCalls + row.outboundCalls;
          const totalPercent = maxCalls > 0 ? (totalCalls / maxCalls) * 100 : 0;
          const inboundPercent = totalCalls > 0 ? (row.inboundCalls / totalCalls) * 100 : 0;

          return `
            <div class="calls-volume-row" data-rich-tooltip="${richTooltipData({
              title: row.agent,
              lines: [
                { text: `Inbound: ${row.inboundCalls}`, color: "#4ba8dc" },
                { text: `Outbound: ${row.outboundCalls}`, color: "#9ecf63" },
                { text: `Total: ${totalCalls}`, color: "#dbbf8a" },
              ],
            })}">
              <span class="compact-list-label">${row.agent}</span>
              <div class="calls-volume-lollipop">
                <div class="calls-volume-stem" style="width:${totalPercent}%"></div>
                <div class="calls-volume-dot" style="left:calc(${totalPercent}% - 12px);">
                  <span class="calls-volume-dot-core" style="width:${Math.max(6, inboundPercent * 0.24)}px;height:${Math.max(6, inboundPercent * 0.24)}px;"></span>
                </div>
              </div>
              <span class="compact-list-value">${totalCalls}</span>
            </div>
          `;
        })
        .join("")}
    </div>
  `;
}

function renderCompositionChart(dataset) {
  if (!elements.compositionCard || !elements.compositionChart) return;

  const shouldShow = state.activeTab === "all" && !isSingleAgentView(dataset) && dataset.rows.length > 1;
  elements.compositionCard.hidden = !shouldShow;
  if (!shouldShow) {
    elements.compositionChart.innerHTML = "";
    return;
  }

  setText(elements.compositionTitle, "KPI Composition");
  const rows = dataset.rows;

  elements.compositionChart.innerHTML = `
    <div class="compact-chart-legend">
      <span class="compact-legend-item"><span class="compact-legend-swatch stacked-segment-attendance"></span>Attendance</span>
      <span class="compact-legend-item"><span class="compact-legend-swatch stacked-segment-qa"></span>QA</span>
      <span class="compact-legend-item"><span class="compact-legend-swatch stacked-segment-aht"></span>AHT</span>
    </div>
    <div class="stacked-chart">
      ${rows
        .map((row) => {
          const attendance = ((row.attendanceScore ?? 0) / 5) * 100 * 0.5;
          const qa = ((row.qaScore ?? 0) / 5) * 100 * 0.25;
          const aht = ((row.ahtScore ?? 0) / 5) * 100 * 0.25;
          return `
            <div class="stacked-row" data-rich-tooltip="${richTooltipData({
              title: row.agent,
              lines: [
                { text: `Attendance: ${formatNumber(row.attendanceScore ?? 0)}`, color: "#e0ad4a" },
                { text: `QA: ${formatNumber(row.qaScore ?? 0)}`, color: "#4ba8dc" },
                { text: `AHT: ${formatNumber(row.ahtScore ?? 0)}`, color: "#66a85f" },
              ],
            })}">
              <div class="stacked-bar">
                <div class="stacked-segment-attendance" style="--segment-size:${attendance}%"></div>
                <div class="stacked-segment-qa" style="--segment-size:${qa}%"></div>
                <div class="stacked-segment-aht" style="--segment-size:${aht}%"></div>
              </div>
              <span class="stacked-label">${row.agent}</span>
            </div>
          `;
        })
        .join("")}
    </div>
  `;
}

function getTrendDataset() {
  if (state.activeTab === "aht") {
    const selectedEnd = state.filters.ahtEnd ? new Date(`${state.filters.ahtEnd}T00:00:00Z`) : null;
    const latestAhtDate =
      [...state.ahtRows]
        .map((row) => row.date)
        .filter((date) => date instanceof Date && !Number.isNaN(date.getTime()))
        .sort((a, b) => a - b)
        .at(-1) || null;

    const endDate = selectedEnd || latestAhtDate;
    if (!endDate) return [];

    const startDate = new Date(endDate);
    startDate.setUTCDate(endDate.getUTCDate() - 6);

    const groups = new Map();
    state.ahtRows
      .filter((row) => row.date)
      .filter((row) => row.date >= startDate && row.date <= new Date(`${toInputDate(endDate)}T23:59:59Z`))
      .filter((row) => state.filters.agent === "all" || row.agent === state.filters.agent)
      .forEach((row) => {
        const key = toInputDate(row.date);
        const entry = groups.get(key) || {
          date: row.date,
          inboundCalls: 0,
          outboundCalls: 0,
          talkSeconds: 0,
          holdSeconds: 0,
        };

        entry.inboundCalls += row.inboundCalls;
        entry.outboundCalls += row.outboundCalls;
        entry.talkSeconds += row.talkSeconds;
        entry.holdSeconds += row.holdSeconds;
        groups.set(key, entry);
      });

    return [...groups.entries()]
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([, entry]) => {
        const totalCalls = entry.inboundCalls + entry.outboundCalls;
        const ahtSeconds = totalCalls ? (entry.talkSeconds + entry.holdSeconds) / totalCalls : null;
        return {
          label: formatShortDate(entry.date),
          fullLabel: formatDate(entry.date),
          value: ahtSeconds,
          displayValue: formatDurationFromSeconds(ahtSeconds),
        };
      })
      .filter((entry) => entry.value !== null);
  }

  const groups = new Map();

  state.monthlyRows.forEach((row) => {
    if (state.filters.agent !== "all" && row.agent !== state.filters.agent) return;
    const key = `${row.year}-${String(row.monthIndex + 1).padStart(2, "0")}`;
    const label = `${row.month} ${row.year}`;
    const entry = groups.get(key) || { label, values: [] };

    if (state.activeTab === "attendance") {
      entry.values.push(row.attendanceScore);
    } else if (state.activeTab === "qa") {
      entry.values.push(calculateQaScore(qaValueForWeek(row, state.filters.week)));
    } else if (state.activeTab === "aht") {
      entry.values.push(row.ahtScore);
    } else {
      entry.values.push(row.overallScore);
    }

      groups.set(key, entry);
    });

  return [...groups.entries()]
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([, entry]) => ({
      label: entry.label,
      value: average(entry.values),
    }))
    .filter((entry) => entry.value !== null);
}

function renderTrendChart() {
  if (!elements.trendChart || !elements.trendSection) return;
  if (isSingleAgentView()) {
    elements.trendChart.innerHTML = "";
    elements.trendSection.hidden = true;
    return;
  }
  if (removeTrendTooltipDismiss) {
    removeTrendTooltipDismiss();
    removeTrendTooltipDismiss = null;
  }

  const trendData = getTrendDataset();
  const isAhtTrend = state.activeTab === "aht";
  const isMobileTrend = typeof window !== "undefined" && window.innerWidth <= 760;
  const trendEndDate = state.filters.ahtEnd ? new Date(`${state.filters.ahtEnd}T00:00:00Z`) : null;
  const title =
    state.activeTab === "attendance"
      ? "Attendance Trend"
      : state.activeTab === "qa"
        ? "QA Trend"
        : state.activeTab === "aht"
            ? "AHT Trend"
            : "Overall Score Trend";

  const subcopy = isAhtTrend
    ? state.filters.agent === "all"
      ? `Daily AHT for the last 7 days ending ${formatDate(trendEndDate)}.`
      : `Daily AHT for ${state.filters.agent} over the last 7 days ending ${formatDate(trendEndDate)}.`
    : state.filters.agent === "all"
      ? "Month-over-month average for the visible team scope."
      : `Month-over-month average for ${state.filters.agent}.`;

  setText(elements.trendChartTitle, title);
  setText(elements.trendChartSubcopy, subcopy);

  if (trendData.length < 2) {
      elements.trendChart.innerHTML = "";
      elements.trendSection.hidden = true;
      return;
  }

  elements.trendSection.hidden = false;
  const measuredWidth = Math.round(elements.trendChart.getBoundingClientRect().width || 0);
  const width = isAhtTrend && !isMobileTrend ? Math.max(measuredWidth, 640) : 640;
  const height = 220;
  const padding = 28;
  const values = trendData.map((entry) => entry.value);
  const maxValue = Math.max(...values);
  const minValue = Math.min(...values);
  const range = Math.max(1, maxValue - minValue);
  const stepX = trendData.length > 1 ? (width - padding * 2) / (trendData.length - 1) : 0;

  const points = trendData.map((entry, index) => {
    const x = padding + stepX * index;
    const y = height - padding - ((entry.value - minValue) / range) * (height - padding * 2);
    return { ...entry, x, y };
  });

  const compactAhtTicks =
    isAhtTrend && isMobileTrend
      ? points.filter((point, index) => index === 0 || index === Math.floor((points.length - 1) / 2) || index === points.length - 1)
      : points;

  const linePath = points.map((point, index) => `${index === 0 ? "M" : "L"} ${point.x} ${point.y}`).join(" ");
  const areaPath = `${linePath} L ${points.at(-1).x} ${height - padding} L ${points[0].x} ${height - padding} Z`;

  elements.trendChart.innerHTML = `
    <div class="line-chart">
      <div class="line-chart-stage">
        <svg viewBox="0 0 ${width} ${height}" class="line-chart-svg" aria-hidden="true">
          <line x1="${padding}" y1="${height - padding}" x2="${width - padding}" y2="${height - padding}" class="line-chart-axis"></line>
          <path d="${areaPath}" class="line-chart-area"></path>
            <path d="${linePath}" class="line-chart-path"></path>
            ${points
              .map(
                (point) => `
                  <circle cx="${point.x}" cy="${point.y}" r="5" class="line-chart-point">
                    <title>${point.fullLabel || point.label}: ${point.displayValue || Number(point.value).toFixed(2)}</title>
                  </circle>
                `
              )
              .join("")}
          </svg>
          <div class="line-chart-hotspots">
            ${points
              .map(
                (point, index) => `
                  <button
                    type="button"
                    class="line-chart-hotspot"
                    data-trend-point-index="${index}"
                    aria-label="${point.fullLabel || point.label}: ${point.displayValue || Number(point.value).toFixed(2)}"
                    style="left:${(point.x / width) * 100}%; top:${(point.y / height) * 100}%"
                  ></button>
                `
              )
              .join("")}
          </div>
          <div class="line-chart-tooltip" id="trendTooltip" hidden></div>
        </div>
        <div class="line-chart-labels">
          ${(isAhtTrend && isMobileTrend ? compactAhtTicks : points)
            .map(
              (point) => `
                  <button
                    type="button"
                    class="line-chart-label line-chart-label-button${isAhtTrend && isMobileTrend ? " line-chart-label-compact" : ""}"
                    data-trend-point-index="${points.findIndex((entry) => entry.label === point.label && entry.value === point.value)}"
                  >
                    ${point.label}${isAhtTrend && isMobileTrend ? "" : `<br>${point.displayValue || Number(point.value).toFixed(2)}`}
                  </button>
                `
              )
            .join("")}
        </div>
    </div>
  `;

  const tooltip = elements.trendChart.querySelector("#trendTooltip");
  const stage = elements.trendChart.querySelector(".line-chart-stage");
  const triggers = [...elements.trendChart.querySelectorAll("[data-trend-point-index]")];
  const supportsHover =
    typeof window !== "undefined" && typeof window.matchMedia === "function"
      ? window.matchMedia("(hover: hover) and (pointer: fine)").matches
      : true;
  let pinnedIndex = null;

  const hideTooltip = () => {
    if (!tooltip) return;
    tooltip.hidden = true;
    tooltip.innerHTML = "";
  };

  const updateTooltip = (index, options = {}) => {
    const point = points[index];
    if (!tooltip || !point) return;
    pinnedIndex = options.pinned ? index : null;
    tooltip.hidden = false;
    tooltip.innerHTML = `
      <strong>${point.fullLabel || point.label}</strong>
      <span>${point.displayValue || Number(point.value).toFixed(2)}</span>
    `;
    tooltip.style.left = `${(point.x / width) * 100}%`;
    tooltip.style.top = `${(point.y / height) * 100}%`;
  };

  triggers.forEach((trigger) => {
    const index = Number(trigger.dataset.trendPointIndex);
    trigger.addEventListener("click", (event) => {
      event.stopPropagation();
      if (supportsHover) {
        updateTooltip(index);
        return;
      }

      if (pinnedIndex === index && tooltip && !tooltip.hidden) {
        pinnedIndex = null;
        hideTooltip();
        return;
      }

      updateTooltip(index, { pinned: true });
    });
    trigger.addEventListener("focus", () => updateTooltip(index));
    trigger.addEventListener("mouseenter", () => {
      if (supportsHover) updateTooltip(index);
    });
    trigger.addEventListener("mouseleave", () => {
      if (supportsHover && pinnedIndex === null) hideTooltip();
    });
    trigger.addEventListener("blur", () => {
      if (pinnedIndex === null) hideTooltip();
    });
  });

  stage?.addEventListener("mouseleave", () => {
    if (supportsHover && pinnedIndex === null) hideTooltip();
  });

  const dismissTooltip = (event) => {
    if (!elements.trendChart?.contains(event.target)) {
      pinnedIndex = null;
      hideTooltip();
    }
  };

  document.addEventListener("pointerdown", dismissTooltip, true);
  removeTrendTooltipDismiss = () => {
    document.removeEventListener("pointerdown", dismissTooltip, true);
  };
}

function renderCharts(dataset) {
  if (!elements.primaryChart || !elements.secondaryChart) return;
  const emptyCopy = buildEmptyStateCopy();
  const singleAgentView = isSingleAgentView(dataset);
  const scoreKey =
    state.activeTab === "attendance"
      ? "attendanceScore"
      : state.activeTab === "qa"
        ? "qaScore"
        : state.activeTab === "aht"
          ? "ahtScore"
          : "overallScore";

  setText(elements.primaryChartTitle, dataset.chartTitle);
  setText(elements.primaryChartSubcopy, dataset.chartSubcopy);
  setText(elements.secondaryChartTitle, dataset.distributionTitle);
  setText(elements.secondaryChartSubcopy, dataset.distributionSubcopy);
  if (elements.secondaryChartCard) {
    elements.secondaryChartCard.hidden = singleAgentView;
  }
  elements.chartsSection?.classList.toggle("content-grid-single", singleAgentView);

  const labels = dataset.rows.map((row) => row.agent);
  const values = dataset.rows.map((row) => {
    if (state.activeTab === "attendance") return safeNumber(row.attendancePercentDisplay);
    if (state.activeTab === "qa") return safeNumber(row.qaAverageDisplay);
    if (state.activeTab === "aht") return durationToSeconds(row.ahtDisplay);
    return row.overallScore ?? 0;
  });

  if (!labels.length) {
    elements.primaryChart.innerHTML = `<div class="chart-empty">${escapeHtml(emptyCopy.chart)}</div>`;
    elements.secondaryChart.innerHTML = `<div class="chart-empty">${escapeHtml(emptyCopy.chart)}</div>`;
    return;
  }

  if (isSingleAgentAllView(dataset)) {
    const row = dataset.rows[0];
    const metrics = [
      {
        label: "Overall Score",
        value: row.overallScore,
        displayValue: formatNumber(row.overallScore),
        color: "linear-gradient(90deg, rgba(44, 142, 207, 0.95), rgba(132, 205, 68, 0.95))",
        tooltip: {
          title: row.agent,
          lines: [
            { text: `Overall Score: ${formatNumber(row.overallScore)}`, color: "#4ba8dc" },
            row.equivalent || "N/A",
          ],
        },
      },
      {
        label: "Attendance Score",
        value: row.attendanceScore,
        displayValue: formatNumber(row.attendanceScore),
        color: "linear-gradient(90deg, rgba(224, 173, 74, 0.92), rgba(132, 205, 68, 0.92))",
        tooltip: {
          title: row.agent,
          lines: [{ text: `Attendance Score: ${formatNumber(row.attendanceScore)}`, color: "#e0ad4a" }],
        },
      },
      {
        label: "QA Score",
        value: row.qaScore,
        displayValue: formatNumber(row.qaScore),
        color: "linear-gradient(90deg, rgba(75, 146, 162, 0.95), rgba(62, 168, 220, 0.95))",
        tooltip: {
          title: row.agent,
          lines: [{ text: `QA Score: ${formatNumber(row.qaScore)}`, color: "#4ba8dc" }],
        },
      },
      {
        label: "AHT Score",
        value: row.ahtScore,
        displayValue: formatNumber(row.ahtScore),
        color: "linear-gradient(90deg, rgba(79, 163, 95, 0.94), rgba(136, 206, 68, 0.95))",
        tooltip: {
          title: row.agent,
          lines: [{ text: `AHT Score: ${formatNumber(row.ahtScore)}`, color: "#66a85f" }],
        },
      },
    ];

    elements.primaryChart.innerHTML = `
      <div class="simple-bar-chart">
        ${metrics
          .map((metric) => {
            const percent = Math.max(0, Math.min(100, ((metric.value ?? 0) / 5) * 100));
            return `
              <div class="simple-bar-row" data-rich-tooltip="${richTooltipData(metric.tooltip)}">
                <span class="simple-bar-label">${metric.label}</span>
                <div class="simple-bar-track"><div class="simple-bar-fill" style="width:${percent}%; background:${metric.color};"></div></div>
                <span class="simple-bar-value">${metric.displayValue}</span>
              </div>
            `;
          })
          .join("")}
      </div>
    `;
    elements.secondaryChart.innerHTML = "";
    return;
  }

  const maxValue = Math.max(...values, 0);
  elements.primaryChart.innerHTML = `
    <div class="simple-bar-chart">
      ${labels
        .map((label, index) => {
          const rawValue = values[index];
          const percent = maxValue > 0 ? (rawValue / maxValue) * 100 : 0;
          const displayValue =
            state.activeTab === "attendance" || state.activeTab === "qa"
              ? `${Number(rawValue).toFixed(2)}%`
              : state.activeTab === "aht"
                ? formatDurationFromSeconds(rawValue)
                : Number(rawValue).toFixed(2);

          return `
            <div class="simple-bar-row" data-rich-tooltip="${richTooltipData({
              title: label,
              lines: [displayValue],
            })}">
              <span class="simple-bar-label">${label}</span>
              <div class="simple-bar-track"><div class="simple-bar-fill" style="width:${percent}%"></div></div>
              <span class="simple-bar-value">${displayValue}</span>
            </div>
          `;
        })
        .join("")}
    </div>
  `;

  const distribution = buildDistribution(dataset.rows, scoreKey);
  const total = distribution.reduce((sum, entry) => sum + entry.count, 0);
  const colors = ["#ca5a4f", "#dda34a", "#4b92a2", "#69aa60", "#2f7a4c"];
  let start = 0;
  const segments = distribution.map((entry, index) => {
    const size = total > 0 ? (entry.count / total) * 360 : 0;
    const segment = `${colors[index]} ${start}deg ${start + size}deg`;
    start += size;
    return segment;
  });

  const defaultIndex = Math.max(
    0,
    distribution.findIndex((entry) => entry.count > 0)
  );
  const defaultEntry = distribution[defaultIndex] || distribution[0];

  elements.secondaryChart.innerHTML = `
    <div class="donut-chart">
      <div class="donut-visual" style="background: conic-gradient(${segments.join(", ")});">
        <div class="donut-center">
          <div>
            <strong>${total}</strong>
            <span>Agents</span>
          </div>
        </div>
      </div>
      <div class="donut-side-panel">
        <div class="donut-band-list">
        ${distribution
          .map(
            (entry, index) => `
              <button
                type="button"
                class="donut-band-button${index === defaultIndex ? " is-active" : ""}"
                data-distribution-index="${index}"
                data-rich-tooltip="${richTooltipData({
                  title: `${entry.score} - ${SCORE_LABELS[entry.score]}`,
                  lines: [
                    { text: `${entry.count} ${entry.count === 1 ? "agent" : "agents"}`, color: colors[index] },
                    previewAgentNames(entry.agents, 3),
                  ],
                })}"
              >
                <span class="compact-legend-swatch" style="background:${colors[index]}"></span>
                <span class="donut-band-label">${entry.score} - ${SCORE_LABELS[entry.score]}</span>
                <span class="donut-band-count">${entry.count}</span>
              </button>
            `
          )
          .join("")}
        </div>
        <div class="donut-detail-panel" id="distributionDetailPanel"></div>
      </div>
    </div>
  `;

  const detailPanel = elements.secondaryChart.querySelector("#distributionDetailPanel");
  const bandButtons = [...elements.secondaryChart.querySelectorAll(".donut-band-button")];
  renderDistributionDetail(detailPanel, defaultEntry, colors[defaultIndex]);

  bandButtons.forEach((button) => {
    const updateDetail = () => {
      const index = Number(button.dataset.distributionIndex);
      const entry = distribution[index];
      if (!entry) return;
      bandButtons.forEach((item) => item.classList.toggle("is-active", item === button));
      renderDistributionDetail(detailPanel, entry, colors[index]);
    };

    button.addEventListener("click", updateDetail);
    button.addEventListener("mouseenter", updateDetail);
    button.addEventListener("focus", updateDetail);
  });
}

function applyChartVisibilityCleanup(dataset) {
  const singleAgentView = isSingleAgentView(dataset);
  const compactCards = [
    elements.topBottomCard,
    elements.ahtComponentsCard,
    elements.callsVolumeCard,
    elements.compositionCard,
  ].filter(Boolean);
  const visibleCompactCards = compactCards.filter((card) => !card.hidden);

  if (elements.chartsSection) {
    elements.chartsSection.hidden = singleAgentView || dataset.rows.length === 0;
  }

  if (elements.compactInsightsGrid) {
    elements.compactInsightsGrid.hidden = singleAgentView || visibleCompactCards.length === 0;
    elements.compactInsightsGrid.classList.toggle("compact-insights-grid-duo", visibleCompactCards.length === 2);
    elements.compactInsightsGrid.classList.toggle("compact-insights-grid-solo", visibleCompactCards.length === 1);
  }

  compactCards.forEach((card) => {
    card.classList.toggle("compact-chart-card-expanded", visibleCompactCards.length === 1 && !card.hidden);
  });
}

function renderTable(dataset) {
  if (!elements.tableHeadRow || !elements.tableBody) return;
  const emptyCopy = buildEmptyStateCopy();
  const sortedRows = sortDatasetRows(dataset);
  state.currentDataset = { ...dataset, rows: sortedRows };
  setText(elements.tableTitle, dataset.title);
  setText(elements.tableSubcopy, "Filtered KPI results.");
  setText(elements.resultsCount, `${sortedRows.length} ${sortedRows.length === 1 ? "row" : "rows"}`);
  if (elements.exportTableCsv) {
    elements.exportTableCsv.disabled = !sortedRows.length;
  }
  elements.tableHeadRow.innerHTML = "";
  elements.tableBody.innerHTML = "";

  dataset.columns.forEach((column) => {
    const th = document.createElement("th");
    const isSorted = state.tableSort.key === column.key;
    const button = document.createElement("button");
    button.type = "button";
    button.className = `table-sort-button${isSorted ? " is-active" : ""}`;
    button.dataset.sortKey = column.key;
    button.setAttribute("aria-label", `Sort by ${column.label}`);
    button.innerHTML = `
      <span>${column.label}</span>
      <span class="table-sort-indicator" aria-hidden="true">${isSorted ? (state.tableSort.direction === "asc" ? "▲" : "▼") : "↕"}</span>
    `;
    button.addEventListener("click", () => {
      if (state.tableSort.key === column.key) {
        state.tableSort.direction = state.tableSort.direction === "asc" ? "desc" : "asc";
      } else {
        state.tableSort.key = column.key;
        state.tableSort.direction = column.key === "agent" ? "asc" : "desc";
      }
      render();
    });
    th.appendChild(button);
    elements.tableHeadRow.appendChild(th);
  });

  if (!sortedRows.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="${dataset.columns.length}" class="empty-state">${escapeHtml(emptyCopy.table)}</td>`;
    elements.tableBody.appendChild(tr);
    return;
  }

  sortedRows.forEach((row) => {
    const tr = document.createElement("tr");
    const highlightTone = rowHighlightTone(row);
    if (highlightTone) {
      tr.classList.add(`table-row-${highlightTone}`);
    }
    dataset.columns.forEach((column) => {
      const td = document.createElement("td");
      if (column.type === "score") {
        td.innerHTML = scoreBadge(row[column.key]);
      } else if (column.key === "equivalent") {
        td.innerHTML = statusBadge(row[column.key]);
      } else {
        td.textContent = row[column.key] ?? "N/A";
      }
      tr.appendChild(td);
    });
    elements.tableBody.appendChild(tr);
  });
}

function updateHeaderContext() {
  const label =
    state.activeTab === "attendance"
      ? "Attendance"
      : state.activeTab === "qa"
        ? "Quality Assurance"
        : state.activeTab === "aht"
          ? "AHT"
          : "All KPIs";

  setText(elements.mobileTopbarContext, label);
  setText(elements.latestPeriodText, formatMonthLabel(state.filters.year, state.filters.month));
  setText(elements.ahtRangeText, formatDateRange(
    state.filters.ahtStart ? new Date(`${state.filters.ahtStart}T00:00:00Z`) : null,
    state.filters.ahtEnd ? new Date(`${state.filters.ahtEnd}T00:00:00Z`) : null
  ));
}

function render() {
  const title =
    state.activeTab === "attendance"
      ? "Attendance Filters"
      : state.activeTab === "qa"
        ? "Quality Assurance Filters"
        : state.activeTab === "aht"
        ? "AHT Filters"
        : "All KPI Filters";
  const filterSummary = buildFilterSummaryLine();

  setText(elements.filtersTitle, title);
  setText(elements.primaryScopeSummary, filterSummary);
  setText(elements.secondaryScopeSummary, filterSummary);
  setText(elements.trendScopeSummary, filterSummary);
  setText(elements.filtersScopeSummary, filterSummary);
  setText(elements.tableScopeSummary, filterSummary);
  updateHeaderContext();
  renderFilters();
  const dataset = getDataset();
  renderSummaryCards(dataset);
  renderManagerInsights(dataset);
  renderSingleAgentOverview(dataset);
  renderTrendChart();
  renderCharts(dataset);
  renderTopBottomChart(dataset);
  renderAhtComponentsChart();
  renderCallsVolumeChart();
  renderCompositionChart(dataset);
  applyChartVisibilityCleanup(dataset);
  renderTable(dataset);
}

function setActiveTab(tab) {
  state.activeTab = tab;
  elements.tabButtons.forEach((button) => {
    const isActive = button.dataset.tab === tab;
    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-selected", String(isActive));
  });
  render();
}

function handleFilterChange(key, value) {
  if (key === "agent") {
    const trimmed = String(value ?? "").trim();
    const matchedAgent = state.options.agents.find((agent) => agent.toLowerCase() === trimmed.toLowerCase());
    if (!trimmed || trimmed.toLowerCase() === "all agents") {
      state.filters.agent = "all";
      render();
      return;
    }
    if (!matchedAgent) {
      render();
      return;
    }
    state.filters.agent = matchedAgent;
    render();
    return;
  }

  state.filters[key] = value;

  if (key === "year") {
    const months = state.options.monthsByYear.get(value) || [];
    if (!months.includes(state.filters.month)) {
      state.filters.month = months.at(-1) || "";
    }
  }

  if (key === "ahtStart" && state.filters.ahtEnd && value > state.filters.ahtEnd) {
    state.filters.ahtEnd = value;
  }

  if (key === "ahtEnd" && state.filters.ahtStart && value < state.filters.ahtStart) {
    state.filters.ahtStart = value;
  }

  render();
}

function bindEvents() {
  elements.tabButtons.forEach((button) => {
    button.addEventListener("click", () => setActiveTab(button.dataset.tab));
  });

  elements.filtersGrid.addEventListener("change", (event) => {
    const key = event.target?.dataset?.filterKey;
    if (!key) return;
    handleFilterChange(key, event.target.value);
  });

  elements.resetFilters.addEventListener("click", () => {
    resetFilters();
    render();
  });

  elements.exportTableCsv?.addEventListener("click", () => {
    exportCurrentTable();
  });

  elements.mobileHomeButtons.forEach((button) => {
    button.addEventListener("click", () => {
      const targetId = button.dataset.mobileTarget;
      const action =
        targetId === "summaryGrid"
          ? "summary"
          : targetId === "chartsSection"
            ? "charts"
            : targetId === "tableSection"
              ? "table"
              : targetId === "filtersSection"
                ? "filters"
                : "home";
      scrollToSection(targetId, action);
    });
  });

  elements.mobileBottomNavButtons.forEach((button) => {
    button.addEventListener("click", () => {
      const action = button.dataset.mobileAction;
      if (action === "home") {
        state.mobileLegendOpen = false;
        applyMobileLegendState();
        window.scrollTo({ top: 0, behavior: "smooth" });
        setMobileNavActive("home");
        return;
      }

      if (action === "legend") {
        state.mobileLegendOpen = !state.mobileLegendOpen;
        applyMobileLegendState();
        setMobileNavActive(state.mobileLegendOpen ? "legend" : "home");
        return;
      }

      const targetId =
        action === "summary"
          ? "summaryGrid"
          : action === "charts"
            ? "chartsSection"
            : action === "table"
              ? "tableSection"
              : "filtersSection";

      scrollToSection(targetId, action);
    });
  });

  elements.mobileLegendClose?.addEventListener("click", () => {
    state.mobileLegendOpen = false;
    applyMobileLegendState();
    setMobileNavActive("home");
  });

  elements.mobileLogoutButton?.addEventListener("click", () => {
    document.querySelector("#googleLogout")?.click();
  });

  window.addEventListener("flyland:auth-signed-out", () => {
    state.mobileLegendOpen = false;
    applyMobileLegendState();
    setMobileNavActive("home");
  });
}

function handleLoadError(error) {
  console.error(error);
  setText(elements.dataStatusText, "Unable to load Banyan KPI data.");
  const detail = error?.message ? `Error: ${error.message}` : "Unknown runtime error.";
  setText(elements.sourceNoteText, detail);
  if (elements.primaryChart) {
    elements.primaryChart.innerHTML = `<div class="chart-empty">Unable to load primary chart.<br>${detail}</div>`;
  }
  if (elements.secondaryChart) {
    elements.secondaryChart.innerHTML = `<div class="chart-empty">Unable to load distribution chart.<br>${detail}</div>`;
  }
}

async function loadData() {
  const [monthlyRows, ahtRows, crosswalkRows] = await Promise.all([
    fetchCsvRows(DATA_URLS.monthly),
    fetchCsvRows(DATA_URLS.aht),
    fetchCsvRows(DATA_URLS.crosswalk),
  ]);

  state.aliasLookup = createAliasLookup(crosswalkRows);

  state.monthlyRows = monthlyRows.map(monthlyRowToModel).filter((row) => row.agent && row.year && row.month);
  state.qaWeeklyAvailable = state.monthlyRows.some((row) =>
    [row.qaWeek1, row.qaWeek2, row.qaWeek3, row.qaWeek4, row.qaTotalAverage].some((value) => value !== null)
  );

  const fallbackYear =
    state.monthlyRows
      .map((row) => Number(row.year))
      .filter((value) => Number.isFinite(value))
      .sort((a, b) => a - b)
      .at(-1) || new Date().getFullYear();

  state.ahtRows = ahtRows
    .map((row) => ahtRowToModel(row, fallbackYear))
    .filter((row) => row.agent && row.date && row.matchedAgent);

  buildOptions();
  resetFilters();

  setText(elements.dataStatusText, "Banyan KPI data loaded.");
  setText(elements.sourceNoteText, state.qaWeeklyAvailable
    ? "Monthly and weekly-ready sources are active."
    : "The published monthly CSV currently exposes only overall QA, not QA WK1-WK4 / TTL AVG columns.");

  render();
}

function initialize() {
  if (state.initialized) return;
  state.initialized = true;
  bindEvents();
  bindRichTooltipEvents();
  initializeMobileLegend();

  const start = () => loadData().catch(handleLoadError);

  if (document.body.dataset.requireAuth === "true") {
    window.addEventListener("flyland:auth-granted", start, { once: true });
    if (window.__flylandAuthState?.authorized) start();
  } else {
    start();
  }
}

window.addEventListener("error", (event) => {
  if (!state.initialized) return;
  handleLoadError(event.error || new Error(event.message || "Window error"));
});

window.addEventListener("unhandledrejection", (event) => {
  if (!state.initialized) return;
  handleLoadError(event.reason instanceof Error ? event.reason : new Error(String(event.reason)));
});

initialize();
