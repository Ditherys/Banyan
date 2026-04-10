import { loadAllDatasets } from "./dataLoader.js";
import {
  buildKpiDataset,
  buildRealtimeDataset,
  calculateAHTScore,
  calculateAdmitsScore,
  calculateAttendanceScore,
  calculateOverallScore,
  calculatePerformanceScore,
  calculateQAScore,
  calculateTransferRate,
  calculateTransferScore,
  getBottomPerformers,
  getOverallComposition,
} from "./kpiCalculator.js";
import { getAvailableAgents, getAvailableWeeks, getFilteredRecords, populateSelect } from "./filters.js";
import {
  renderBarChart,
  renderAgentDeltaVarianceChart,
  renderComparisonChart,
  renderContributionChart,
  renderDistributionChart,
  renderLineChart,
  renderRangeComparisonChart,
  renderScoreSpreadChart,
  renderStackedBarChart,
  renderVarianceChart,
} from "./charts.js";
import { exportRowsToCsv, initializeTable, renderTable, renderTableSummary, sortRecords } from "./table.js";

const state = {
  dataset: null,
  realtimeDataset: null,
  filters: { month: "all", week: "all", agent: "all", search: "" },
  aggregateAllRanges: false,
  forceAllWeekSelection: false,
  dashboardFocus: "all",
  sort: { key: "overallScore", direction: "desc" },
  expandedRowKeys: new Set(),
  mobileSidebarOpen: false,
  mobileTableExpanded: true,
  mobileFiltersExpanded: false,
  floatingLegendOpen: false,
  distributionDrilldownOpen: false,
  distributionDrilldownExpanded: false,
  distributionDrilldownDetail: null,
  resizeRaf: null,
  charts: {},
  initialized: false,
  dualDateRange: {
    transfer: { initialized: false, primaryStart: "", primaryEnd: "", compareEnabled: false, compareStart: "", compareEnd: "" },
    aht: { initialized: false, primaryStart: "", primaryEnd: "", compareEnabled: false, compareStart: "", compareEnd: "" },
  },
};

const KPI_DEFINITIONS = [
  {
    key: "transferScore",
    label: "Transfer",
    raw: (summary) => `Transfer rate ${formatPercent(summary.transferRatePercent)}`,
  },
  {
    key: "admitsScore",
    label: "Admits",
    raw: (summary) =>
      `Admits count ${
        summary.admitsCount === null || summary.admitsCount === undefined ? "N/A" : Number(summary.admitsCount).toFixed(2)
      }`,
  },
  {
    key: "ahtScore",
    label: "AHT",
    raw: (summary) => `Average handle time ${formatTimeFromSeconds(summary.ahtSeconds)}`,
  },
  {
    key: "attendanceScore",
    label: "Attendance",
    raw: (summary) => `Attendance ${formatPercent(summary.attendancePercentValue)}`,
  },
  {
    key: "qaScore",
    label: "Quality Assurance",
    raw: (summary) => `QA ${formatPercent(summary.qaPercentValue)}`,
  },
];

const FOCUS_SCORING = {
  all: { metricKey: "overallScore", title: "Overall", distribution: "transfer" },
  performance: { metricKey: "performanceScore", title: "Performance", distribution: "transfer" },
  transfer: { metricKey: "transferScore", title: "Transfer", distribution: "transfer" },
  admits: { metricKey: "admitsScore", title: "Admits", distribution: "admits" },
  aht: { metricKey: "ahtScore", title: "AHT", distribution: "aht" },
  realtime: { metricKey: "performanceScore", title: "Real Time Performance", distribution: "transfer" },
  attendance: { metricKey: "attendanceScore", title: "Attendance", distribution: "attendance" },
  qa: { metricKey: "qaScore", title: "Quality Assurance", distribution: "qa" },
};

const PERFORMANCE_FOCUS_KEYS = ["performance", "realtime", "transfer", "admits", "aht"];
const SINGLE_KPI_SCORE_KEYS = {
  transfer: "transferScore",
  admits: "admitsScore",
  aht: "ahtScore",
  attendance: "attendanceScore",
  qa: "qaScore",
};

const elements = {
  appShell: document.querySelector("#appShell"),
  mobileSidebar: document.querySelector("#mobileSidebarPanel"),
  mobileMenuToggle: document.querySelector("#mobileMenuToggle"),
  mobileTableToggle: document.querySelector("#mobileTableToggle"),
  mobileLegendLauncher: document.querySelector("#mobileLegendLauncher"),
  mobileFiltersLauncher: document.querySelector("#mobileFiltersLauncher"),
  mobileStatusMonth: document.querySelector("#mobileStatusMonth"),
  mobileStatusWeek: document.querySelector("#mobileStatusWeek"),
  mobileStatusScope: document.querySelector("#mobileStatusScope"),
  mobileLastUpdatedStrip: document.querySelector("#mobileLastUpdatedStrip"),
  mobileLastUpdatedText: document.querySelector("#mobileLastUpdatedText"),
  mobileHomeButtons: [...document.querySelectorAll("[data-mobile-target]")],
  dualDateRangeSection: document.querySelector("#dualDateRangeSection"),
  compareRangeHeaderToggle: document.querySelector("#compareRangeHeaderToggle"),
  primaryDateStart: document.querySelector("#primaryDateStart"),
  primaryDateEnd: document.querySelector("#primaryDateEnd"),
  compareDateToggle: document.querySelector("#compareDateToggle"),
  compareDateStartField: document.querySelector("#compareDateStartField"),
  compareDateEndField: document.querySelector("#compareDateEndField"),
  compareDateStart: document.querySelector("#compareDateStart"),
  compareDateEnd: document.querySelector("#compareDateEnd"),
  dualRangeAgentFilter: document.querySelector("#dualRangeAgentFilter"),
  dualRangeReset: document.querySelector("#dualRangeReset"),
  focusButtons: [...document.querySelectorAll("[data-dashboard-focus]")],
  mobileBottomNavButtons: [...document.querySelectorAll(".mobile-bottom-nav-button[data-mobile-action]")],
  summaryCardsList: [...document.querySelectorAll(".summary-card")],
  mobileDistributionButtons: [...document.querySelectorAll(".distribution-mobile-chip[data-distribution-panel]")],
  mobileDistributionItems: [...document.querySelectorAll("[data-distribution-item]")],
  mobileMorePanel: document.querySelector("#mobileMorePanel"),
  mobileMoreClose: document.querySelector("#mobileMoreClose"),
  mobileMoreLegend: document.querySelector("#mobileMoreLegend"),
  mobileMoreFilters: document.querySelector("#mobileMoreFilters"),
  mobileMoreExport: document.querySelector("#mobileMoreExport"),
  mobileMoreReset: document.querySelector("#mobileMoreReset"),
  floatingLegendToggle: document.querySelector("#floatingLegendToggle"),
  floatingLegendPanel: document.querySelector("#floatingLegendPanel"),
  floatingLegendClose: document.querySelector("#floatingLegendClose"),
  floatingLegendContent: document.querySelector("#floatingLegendContent"),
  distributionDrilldownPanel: document.querySelector("#distributionDrilldownPanel"),
  distributionDrilldownClose: document.querySelector("#distributionDrilldownClose"),
  distributionDrilldownTitle: document.querySelector("#distributionDrilldownTitle"),
  distributionDrilldownSubnote: document.querySelector("#distributionDrilldownSubnote"),
  distributionDrilldownList: document.querySelector("#distributionDrilldownList"),
  legendPanel: document.querySelector(".legend-panel"),
  monthFilterField: document.querySelector("#monthFilterField"),
  weekFilterField: document.querySelector("#weekFilterField"),
  agentFilterField: document.querySelector("#agentFilterField"),
  topFiltersTitle: document.querySelector("#topFiltersTitle"),
  topFiltersGrid: document.querySelector(".top-filters-grid"),
  monthFilterLabel: document.querySelector("#monthFilterLabel"),
  weekFilterLabel: document.querySelector("#weekFilterLabel"),
  agentFilterLabel: document.querySelector("#agentFilterLabel"),
  monthFilter: document.querySelector("#monthFilter"),
  weekFilter: document.querySelector("#weekFilter"),
  agentFilter: document.querySelector("#agentFilter"),
  tableSearch: document.querySelector("#tableSearch"),
  resetFilters: document.querySelector("#resetFilters"),
  exportCsvButton: document.querySelector("#exportCsvButton"),
  expandAllRowsButton: document.querySelector("#expandAllRowsButton"),
  collapseAllRowsButton: document.querySelector("#collapseAllRowsButton"),
  filtersCollapseToggle: document.querySelector("#filtersCollapseToggle"),
  filtersContentArea: document.querySelector("#filtersContentArea"),
  tableCollapseToggle: document.querySelector("#tableCollapseToggle"),
  tableContentArea: document.querySelector("#tableContentArea"),
  tableSummaryStrip: document.querySelector("#tableSummaryStrip"),
  tableRowHint: document.querySelector("#tableRowHint"),
  tableHeadRow: document.querySelector("#tableHeadRow"),
  tableBody: document.querySelector("#tableBody"),
  dataStatusText: document.querySelector("#dataStatusText"),
  realtimeLastUpdatedPill: document.querySelector("#realtimeLastUpdatedPill"),
  realtimeLastUpdatedText: document.querySelector("#realtimeLastUpdatedText"),
  trendChart: document.querySelector("#trendChart"),
  weekScoreChart: document.querySelector("#weekScoreChart"),
  contributionChart: document.querySelector("#contributionChart"),
  varianceChart: document.querySelector("#varianceChart"),
  varianceCard: document.querySelector("#varianceChart")?.closest(".chart-card"),
  transferDistributionChart: document.querySelector("#transferDistributionChart"),
  admitsDistributionChart: document.querySelector("#admitsDistributionChart"),
  ahtDistributionChart: document.querySelector("#ahtDistributionChart"),
  attendanceDistributionChart: document.querySelector("#attendanceDistributionChart"),
  qaDistributionChart: document.querySelector("#qaDistributionChart"),
  performerSpreadChart: document.querySelector("#performerSpreadChart"),
  performerSpreadSummary: document.querySelector("#performerSpreadSummary"),
  stackedChart: document.querySelector("#stackedChart"),
  comparisonChart: document.querySelector("#comparisonChart"),
  comparisonTitle: document.querySelector("#comparisonTitle"),
  comparisonSubnote: document.querySelector("#comparisonSubnote"),
  rankingSubnote: document.querySelector("#rankingSubnote"),
  viewInsightTitle: document.querySelector("#viewInsightTitle"),
  viewInsightBody: document.querySelector("#viewInsightBody"),
  bestInsightTitle: document.querySelector("#bestInsightTitle"),
  bestInsightBody: document.querySelector("#bestInsightBody"),
  watchInsightTitle: document.querySelector("#watchInsightTitle"),
  watchInsightBody: document.querySelector("#watchInsightBody"),
  momentumInsightTitle: document.querySelector("#momentumInsightTitle"),
  momentumInsightBody: document.querySelector("#momentumInsightBody"),
  agentFocusName: document.querySelector("#agentFocusName"),
  agentFocusNarrative: document.querySelector("#agentFocusNarrative"),
  agentOverallValue: document.querySelector("#agentOverallValue"),
  agentOverallNarrative: document.querySelector("#agentOverallNarrative"),
  agentStrongestValue: document.querySelector("#agentStrongestValue"),
  agentStrongestNarrative: document.querySelector("#agentStrongestNarrative"),
  agentWeakestValue: document.querySelector("#agentWeakestValue"),
  agentWeakestNarrative: document.querySelector("#agentWeakestNarrative"),
  navButtons: [...document.querySelectorAll(".nav-item[data-target]")],
  teamOnlySections: [...document.querySelectorAll("[data-scope='team-only']")],
  agentOnlySections: [...document.querySelectorAll("[data-scope='agent-only']")],
  summaryCards: [...document.querySelectorAll(".summary-card[data-kpi]")],
  focusGroupedNodes: [...document.querySelectorAll("[data-focus-group]:not(.summary-card)")],
  rawMetricsGrid: document.querySelector("#rawMetricsGrid"),
  overallCardTitle: document.querySelector("#overallCardTitle"),
  contributionCard: document.querySelector("#contributionCard"),
  trendsSection: document.querySelector("#trendsSection"),
  snapshotSection: document.querySelector("#snapshotSection"),
  performersSection: document.querySelector("#performersSection"),
  deepDiveSection: document.querySelector("#deepDiveSection"),
  diagnosticsSection: document.querySelector("#diagnosticsSection"),
  distributionSection: document.querySelector("#distributionSection"),
  agentsSection: document.querySelector("#agentsSection"),
  tableSectionSubnote: document.querySelector("#tableSectionSubnote"),
  tableDebugLine: document.querySelector("#tableDebugLine"),
  trendSectionTitle: document.querySelector("#trendSectionTitle"),
  snapshotSectionTitle: document.querySelector("#snapshotSectionTitle"),
  varianceSectionTitle: document.querySelector("#varianceSectionTitle"),
  rankingSectionTitle: document.querySelector("#rankingSectionTitle"),
  bottomSectionTitle: document.querySelector("#bottomSectionTitle"),
  improvedSectionTitle: document.querySelector("#improvedSectionTitle"),
  breakdownSectionTitle: document.querySelector("#breakdownSectionTitle"),
  distributionSectionTitle: document.querySelector("#distributionSectionTitle"),
  tableSectionTitle: document.querySelector("#tableSectionTitle"),
};

state.mobileMoreOpen = false;
state.mobileDistributionPanel = "transfer";

function isRealtimeFocus() {
  return state.dashboardFocus === "realtime";
}

function usesRealtimeSource() {
  return ["realtime", "transfer", "aht"].includes(state.dashboardFocus);
}

function isPerformanceLikeFocus() {
  return PERFORMANCE_FOCUS_KEYS.includes(state.dashboardFocus);
}

function isSingleKpiFocus() {
  return Object.prototype.hasOwnProperty.call(SINGLE_KPI_SCORE_KEYS, state.dashboardFocus);
}

function shouldEnableRowExpansion() {
  return !isSingleKpiFocus() && state.dashboardFocus !== "all";
}

function getActiveDataset() {
  return usesRealtimeSource() ? state.realtimeDataset : state.dataset;
}

function getFocusScopedRecords(records) {
  if (!Array.isArray(records)) return [];
  if (state.dashboardFocus === "admits") {
    return records.filter((record) => record.sourceFlags?.admits);
  }
  if (state.dashboardFocus === "attendance") {
    return records.filter((record) => record.sourceFlags?.attendance);
  }
  if (state.dashboardFocus === "qa") {
    return records.filter((record) => record.sourceFlags?.qa);
  }
  return records;
}

function getCurrentFocusFilteredRecords(search = state.filters.search) {
  const activeDataset = getActiveDataset();
  const focusRecords = getFocusScopedRecords(activeDataset.records);
  return getFilteredRecords(focusRecords, { ...state.filters, search });
}

function isAllDateRangesControlSelected() {
  return false;
}

function populateWeekSelect(options) {
  if (!elements.weekFilter) return;
  elements.weekFilter.innerHTML = "";

  options.forEach((option) => {
    const optionElement = document.createElement("option");
    optionElement.value = option;
    optionElement.textContent = option;
    elements.weekFilter.appendChild(optionElement);
  });
}

function getAggregateDateRangeLabel(records) {
  if (!records.length) return "All Date Ranges";
  if (state.filters.month !== "all") return `${state.filters.month} (All Date Ranges)`;
  const uniqueMonths = [...new Set(records.map((record) => record.monthLabel).filter(Boolean))];
  if (uniqueMonths.length === 1) return `${uniqueMonths[0]} (All Date Ranges)`;
  return "All Selected Date Ranges";
}

function getLatestDisplayFrom(records, dateKey, displayKey) {
  const latestRecord = [...records]
    .filter((record) => record[dateKey] instanceof Date && !Number.isNaN(record[dateKey].getTime()))
    .sort((left, right) => left[dateKey].getTime() - right[dateKey].getTime())
    .at(-1);
  return latestRecord?.[displayKey] || "N/A";
}

function getCurrentTableRecords(search = state.filters.search) {
  if (isDualRangeFocus()) {
    const dualData = getDualRangeDashboardData();
    if (!dualData) return [];
    if (!search) return dualData.tableRecords;
    const searchTerm = search.trim().toLowerCase();
    return dualData.tableRecords.filter((record) =>
      `${record.agentName} ${record.weekEnding}`.toLowerCase().includes(searchTerm)
    );
  }

  const activeDataset = getActiveDataset();
  const focusRecords = getFocusScopedRecords(activeDataset.records);
  const currentFilters = {
    month: state.filters.month,
    week: state.filters.week,
    agent: state.filters.agent,
    search,
  };

  if (currentFilters.week !== "all") {
    return getFilteredRecords(focusRecords, currentFilters);
  }

  const baseRecords = getFilteredRecords(focusRecords, {
    ...currentFilters,
    week: "all",
    search: "",
  });

  const aggregatedByAgent = new Map();
  baseRecords.forEach((record) => {
    const aggregateKey = `${(record.email || "").toLowerCase()}__${record.agentName}`;
    const bucket = aggregatedByAgent.get(aggregateKey) || [];
    bucket.push(record);
    aggregatedByAgent.set(aggregateKey, bucket);
  });

  const aggregated = [...aggregatedByAgent.values()].map((items) => {
    const base = items[0];
    const aggregateLabel = currentFilters.month !== "all"
      ? `${currentFilters.month} (All Date Ranges)`
      : getAggregateDateRangeLabel(items);
    const firstTimeCaller = items.reduce((sum, item) => sum + (Number(item.firstTimeCaller) || 0), 0);
    const transferCount = items.reduce((sum, item) => sum + (Number(item.transferCount) || 0), 0);
    const inboundCalls = items.reduce((sum, item) => sum + (Number(item.inboundCalls) || 0), 0);
    const inboundMinutesSeconds = items.reduce((sum, item) => sum + (Number(item.inboundMinutesSeconds) || 0), 0);
    const holdTimeSeconds = items.reduce((sum, item) => sum + (Number(item.holdTimeSeconds) || 0), 0);
    const admitsCount = items.reduce((sum, item) => sum + (Number(item.admitsCount) || 0), 0);
    const hoursPresent = items.reduce((sum, item) => sum + (Number(item.hoursPresent) || 0), 0);
    const hoursAbsent = items.reduce((sum, item) => sum + (Number(item.hoursAbsent) || 0), 0);
    const sickLeaveHours = items.reduce((sum, item) => sum + (Number(item.sickLeaveHours) || 0), 0);
    const undertimeHours = items.reduce((sum, item) => sum + (Number(item.undertimeHours) || 0), 0);
    const requiredHours = items.reduce((sum, item) => sum + (Number(item.requiredHours) || 0), 0);
    const attendancePercentValue = requiredHours > 0
      ? (hoursPresent / requiredHours) * 100
      : items.reduce((sum, item) => sum + (Number(item.attendancePercentValue) || 0), 0) / Math.max(items.length, 1);
    const qaValues = items
      .map((item) => item.qaPercentValue)
      .filter((value) => typeof value === "number" && !Number.isNaN(value));
    const qaPercentValue = qaValues.length
      ? qaValues.reduce((sum, value) => sum + value, 0) / qaValues.length
      : null;
    const transferRate = calculateTransferRate(
      transferCount === null || transferCount === undefined ? null : transferCount,
      firstTimeCaller === null || firstTimeCaller === undefined ? null : firstTimeCaller
    );
    const transferScore = calculateTransferScore(transferRate);
    const admitsScore = calculateAdmitsScore(admitsCount || null);
    const ahtSeconds = inboundCalls > 0 ? (inboundMinutesSeconds + holdTimeSeconds) / inboundCalls : null;
    const ahtScore = calculateAHTScore(ahtSeconds);
    const attendanceScore = calculateAttendanceScore(attendancePercentValue);
    const qaScore = calculateQAScore(qaPercentValue);
    const performanceScore = calculatePerformanceScore(transferScore, admitsScore, ahtScore);
    const overallWeights = getOverallComposition(performanceScore, attendanceScore, qaScore);
    const overallScore = calculateOverallScore(performanceScore, attendanceScore, qaScore);

    return {
      ...base,
      key: `aggregate__${aggregateKey}__${currentFilters.month}`,
      monthLabel: base.monthLabel,
      weekEnding: aggregateLabel,
      dateRange: aggregateLabel,
      firstTimeCaller,
      transferCount,
      inboundCalls,
      inboundMinutesSeconds,
      inboundMinutes: formatTimeFromSeconds(inboundMinutesSeconds),
      holdTimeSeconds,
      holdTime: formatTimeFromSeconds(holdTimeSeconds),
      admitsCount,
      hoursPresent,
      hoursAbsent,
      sickLeaveHours,
      undertimeHours,
      requiredHours,
      attendancePercentValue,
      attendancePercentDisplay: formatPercent(attendancePercentValue),
      qaPercentValue,
      qaPercentDisplay: formatPercent(qaPercentValue),
      transferRate,
      transferRatePercent: transferRate === null ? null : transferRate * 100,
      transferRateDisplay: formatPercent(transferRate === null ? null : transferRate * 100),
      transferScore,
      admitsScore,
      ahtSeconds,
      ahtDisplay: formatTimeFromSeconds(ahtSeconds),
      ahtScore,
      attendanceScore,
      qaScore,
      performanceScore,
      overallScore,
      overallIncludesQa: overallWeights.includesQa,
      overallWeights: overallWeights.weights,
      lastUpdatedDisplay: getLatestDisplayFrom(items, "lastUpdatedAt", "lastUpdatedDisplay"),
      admitsLastUpdatedDisplay: getLatestDisplayFrom(items, "admitsLastUpdatedAt", "admitsLastUpdatedDisplay"),
      qaLastUpdatedDisplay: getLatestDisplayFrom(items, "qaLastUpdatedAt", "qaLastUpdatedDisplay"),
    };
  });

  if (!currentFilters.search) return aggregated;
  const searchTerm = currentFilters.search.trim().toLowerCase();
  return aggregated.filter((record) =>
    `${record.agentName} ${record.weekEnding}`.toLowerCase().includes(searchTerm)
  );
}

function setMobileDockActive(action) {
  elements.mobileBottomNavButtons.forEach((button) => {
    button.classList.toggle("is-active", button.dataset.mobileAction === action);
  });
}

function triggerMobileAction(action) {
  const isMobile = typeof window !== "undefined" && window.innerWidth <= 720;
  if (!isMobile) return;

  if (action === "more") {
    state.mobileMoreOpen = !state.mobileMoreOpen;
    applyMobileMoreState();
    setMobileDockActive(action);
    return;
  }

  state.mobileMoreOpen = false;
  applyMobileMoreState();

  state.floatingLegendOpen = false;
  applyFloatingLegendState();

  if (action === "table") {
    state.mobileTableExpanded = true;
    applyMobileUiState();
    document.getElementById("tableSection")?.scrollIntoView({ behavior: "smooth", block: "start" });
    setMobileDockActive(action);
    return;
  }

  if (action === "filters") {
    state.mobileFiltersExpanded = true;
    applyMobileUiState();
    document.getElementById("filtersSection")?.scrollIntoView({ behavior: "smooth", block: "start" });
    setMobileDockActive(action);
    return;
  }

  if (action === "summary") {
    const summaryTarget = state.filters.agent === "all" ? "summaryCards" : "agentFocusSection";
    document.getElementById(summaryTarget)?.scrollIntoView({ behavior: "smooth", block: "start" });
    setMobileDockActive(action);
    return;
  }

  if (action === "agents") {
    const agentsTarget = state.filters.agent === "all" ? "agentsSection" : "agentFocusSection";
    document.getElementById(agentsTarget)?.scrollIntoView({ behavior: "smooth", block: "start" });
    setMobileDockActive(action);
    return;
  }

  window.scrollTo({ top: 0, behavior: "smooth" });
  document.getElementById("dashboardTop")?.scrollIntoView({ behavior: "smooth", block: "start" });
  setMobileDockActive("home");
}

function applyMobileUiState() {
  const isMobile = typeof window !== "undefined" && window.innerWidth <= 720;

  elements.mobileSidebar?.classList.toggle("is-mobile-open", isMobile && state.mobileSidebarOpen);
  elements.mobileMenuToggle?.setAttribute("aria-expanded", String(isMobile && state.mobileSidebarOpen));
  if (elements.mobileMenuToggle) {
    elements.mobileMenuToggle.textContent = isMobile && state.mobileSidebarOpen ? "Close" : "Menu";
  }

  elements.tableContentArea?.classList.toggle("is-collapsed-mobile", isMobile && !state.mobileTableExpanded);
  elements.tableCollapseToggle?.setAttribute("aria-expanded", String(!isMobile || state.mobileTableExpanded));
  elements.mobileTableToggle?.setAttribute("aria-expanded", String(!isMobile || state.mobileTableExpanded));
  if (elements.tableCollapseToggle) {
    elements.tableCollapseToggle.textContent = !isMobile || state.mobileTableExpanded ? "Collapse Table" : "Expand Table";
  }
  if (elements.mobileTableToggle) {
    elements.mobileTableToggle.textContent = !isMobile || state.mobileTableExpanded ? "Hide Table" : "Show Table";
  }

  elements.filtersContentArea?.classList.toggle("is-collapsed-mobile", isMobile && !state.mobileFiltersExpanded);
  elements.filtersCollapseToggle?.setAttribute("aria-expanded", String(!isMobile || state.mobileFiltersExpanded));
  if (elements.filtersCollapseToggle) {
    elements.filtersCollapseToggle.textContent = !isMobile || state.mobileFiltersExpanded ? "Hide Filters" : "Show Filters";
  }

  if (!isMobile) {
    state.mobileMoreOpen = false;
    elements.summaryCardsList.forEach((card) => {
      card.classList.remove("is-mobile-expanded");
      card.setAttribute("aria-expanded", "false");
    });
    setMobileDockActive("home");
  }
}

function applyMobileMoreState() {
  const isMobile = typeof window !== "undefined" && window.innerWidth <= 720;
  elements.mobileMorePanel?.classList.toggle("is-open", isMobile && state.mobileMoreOpen);
  elements.mobileMorePanel?.setAttribute("aria-hidden", String(!(isMobile && state.mobileMoreOpen)));
}

function applyMobileDistributionState() {
  const isMobile = typeof window !== "undefined" && window.innerWidth <= 720;
  elements.mobileDistributionButtons.forEach((button) => {
    button.classList.toggle("is-active", button.dataset.distributionPanel === state.mobileDistributionPanel);
  });
  elements.mobileDistributionItems.forEach((item) => {
    item.classList.toggle(
      "is-mobile-active",
      !isMobile || item.dataset.distributionItem === state.mobileDistributionPanel
    );
  });
}

function bindSummaryCardInteractions() {
  elements.summaryCardsList.forEach((card) => {
    card.setAttribute("tabindex", "0");
    card.setAttribute("role", "button");
    card.setAttribute("aria-expanded", "false");

    const toggleCard = () => {
      const willExpand = !card.classList.contains("is-mobile-expanded");
      elements.summaryCardsList.forEach((item) => {
        item.classList.remove("is-mobile-expanded");
        item.setAttribute("aria-expanded", "false");
      });

      if (willExpand) {
        card.classList.add("is-mobile-expanded");
        card.setAttribute("aria-expanded", "true");
      }
    };

    card.addEventListener("click", toggleCard);
    card.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        toggleCard();
      }
    });
  });
}

function resetDashboardFilters() {
  if (isDualRangeFocus()) {
    state.filters.agent = "all";
    state.filters.search = "";
    const dualState = getDualDateRangeState();
    if (dualState) {
      dualState.initialized = false;
      ensureDualDateRangeDefaults(true);
    }
    state.expandedRowKeys = new Set();
    elements.tableSearch.value = "";
    syncDualDateRangeControls();
    updateDashboard();
    return;
  }

  const activeDataset = getActiveDataset();
  const focusRecords = getFocusScopedRecords(activeDataset.records);
  const monthOptions = [...new Set(focusRecords.map((record) => record.monthLabel))];
  const latestMonth = monthOptions.at(-1) || "all";
  const latestRange = getAvailableWeeks(focusRecords, latestMonth).at(-1) || [...new Set(focusRecords.map((record) => record.weekEnding))].at(-1) || "all";
  state.filters = {
    month: latestMonth,
    week: latestRange,
    agent: "all",
    search: "",
  };
  state.expandedRowKeys = new Set();
  elements.tableSearch.value = "";
  syncFilterOptions();
  updateDashboard();
}

function bindMobileScrollSpy() {
  const sectionMap = [
    { action: "home", title: "Dashboard Home", getElement: () => document.getElementById("dashboardTop") },
    { action: "summary", title: state.filters.agent === "all" ? "KPI Scorecards" : "Agent Focus", getElement: () => document.getElementById(state.filters.agent === "all" ? "summaryCards" : "agentFocusSection") },
    { action: "agents", title: state.filters.agent === "all" ? "Agent Insights" : "Agent Focus", getElement: () => document.getElementById(state.filters.agent === "all" ? "agentsSection" : "agentFocusSection") },
    { action: "table", title: "Team Table", getElement: () => document.getElementById("tableSection") },
  ];

  const updateActiveByScroll = () => {
    const isMobile = typeof window !== "undefined" && window.innerWidth <= 720;
    if (!isMobile || state.mobileMoreOpen) return;

    const threshold = window.innerHeight * 0.28;
    let activeAction = "home";
    let activeTitle = "Dashboard Home";

    sectionMap.forEach(({ action, title, getElement }) => {
      const element = getElement();
      if (!element) return;
      const rect = element.getBoundingClientRect();
      if (rect.top <= threshold) {
        activeAction = action;
        activeTitle = title;
      }
    });

    setMobileDockActive(activeAction);
  };

  window.addEventListener("scroll", updateActiveByScroll, { passive: true });
  updateActiveByScroll();
}

function applyFloatingLegendState() {
  elements.floatingLegendPanel?.classList.toggle("is-open", state.floatingLegendOpen);
  elements.floatingLegendPanel?.setAttribute("aria-hidden", String(!state.floatingLegendOpen));
  elements.floatingLegendToggle?.setAttribute("aria-expanded", String(state.floatingLegendOpen));
}

function applyDistributionDrilldownState() {
  elements.distributionDrilldownPanel?.classList.toggle("is-open", state.distributionDrilldownOpen);
  elements.distributionDrilldownPanel?.setAttribute("aria-hidden", String(!state.distributionDrilldownOpen));
}

function renderDistributionDrilldownContent() {
  if (!elements.distributionDrilldownList || !elements.distributionDrilldownTitle || !elements.distributionDrilldownSubnote) {
    return;
  }

  const detail = state.distributionDrilldownDetail;
  if (!detail) return;

  if (detail.mode === "breakdown") {
    elements.distributionDrilldownTitle.textContent = detail.agentName;
    elements.distributionDrilldownSubnote.textContent = `${detail.weekEnding} | KPI breakdown for the selected agent row.`;
    elements.distributionDrilldownList.innerHTML = detail.items.map((item) => `
      <div class="distribution-drilldown-item">
        ${item.label}: ${item.value}${item.score ? ` | Score ${item.score}` : ""}
      </div>
    `).join("");
    return;
  }

  elements.distributionDrilldownTitle.textContent = `${detail.label} | Score ${detail.score}`;
  elements.distributionDrilldownSubnote.textContent = `${detail.count} agent(s) in this bucket out of ${detail.total} scored agents.`;
  if (!detail.agents.length) {
    elements.distributionDrilldownList.innerHTML =
      '<div class="distribution-drilldown-empty">No agents found in this score bucket.</div>';
    return;
  }

  const previewLimit = 4;
  const visibleAgents = state.distributionDrilldownExpanded ? detail.agents : detail.agents.slice(0, previewLimit);
  const remainingCount = Math.max(detail.agents.length - visibleAgents.length, 0);

  elements.distributionDrilldownList.innerHTML = [
    ...visibleAgents.map((agent) => `<div class="distribution-drilldown-item">${agent}</div>`),
    !state.distributionDrilldownExpanded && remainingCount > 0
      ? `<button class="distribution-drilldown-more" type="button">Show ${remainingCount} more agent(s)</button>`
      : "",
  ].join("");
}

function openDistributionDrilldown(detail) {
  state.distributionDrilldownOpen = true;
  state.distributionDrilldownExpanded = false;
  state.distributionDrilldownDetail = detail;
  renderDistributionDrilldownContent();
  applyDistributionDrilldownState();
}

function openBreakdownDrilldown(record) {
  const focusItems = state.dashboardFocus === "all"
    ? [
      {
        label: "Transfer",
        value: record.transferRateDisplay,
        score: formatScore(record.transferScore),
      },
      {
        label: "Admits",
        value: record.admitsCount === null || record.admitsCount === undefined ? "N/A" : Number(record.admitsCount).toFixed(0),
        score: formatScore(record.admitsScore),
      },
      {
        label: "AHT",
        value: record.ahtDisplay,
        score: formatScore(record.ahtScore),
      },
      {
        label: "Attendance",
        value: record.attendancePercentDisplay,
        score: formatScore(record.attendanceScore),
      },
      {
        label: "Quality Assurance",
        value: record.qaPercentDisplay,
        score: formatScore(record.qaScore),
      },
      {
        label: "Overall",
        value: formatScore(record.overallScore),
        score: null,
      },
    ]
    : state.dashboardFocus === "performance" || isRealtimeFocus()
    ? [
      {
        label: "Transfer",
        value: record.transferRateDisplay,
        score: formatScore(record.transferScore),
      },
      {
        label: "Admits",
        value: record.admitsCount === null || record.admitsCount === undefined ? "N/A" : Number(record.admitsCount).toFixed(0),
        score: formatScore(record.admitsScore),
      },
      {
        label: "AHT",
        value: record.ahtDisplay,
        score: formatScore(record.ahtScore),
      },
      {
        label: "Performance",
        value: formatScore(record.performanceScore),
        score: null,
      },
    ]
    : state.dashboardFocus === "transfer"
    ? [
      {
        label: "Transfer",
        value: record.transferRateDisplay,
        score: formatScore(record.transferScore),
      },
      {
        label: "First-Time Callers",
        value: record.firstTimeCaller === null || record.firstTimeCaller === undefined ? "N/A" : Number(record.firstTimeCaller).toFixed(0),
        score: null,
      },
      {
        label: "Transfer Count",
        value: record.transferCount === null || record.transferCount === undefined ? "N/A" : Number(record.transferCount).toFixed(0),
        score: null,
      },
    ]
    : state.dashboardFocus === "admits"
      ? [
        {
          label: "Admits",
          value: record.admitsCount === null || record.admitsCount === undefined ? "N/A" : Number(record.admitsCount).toFixed(0),
          score: formatScore(record.admitsScore),
        },
      ]
      : state.dashboardFocus === "aht"
        ? [
          {
            label: "AHT",
            value: record.ahtDisplay,
            score: formatScore(record.ahtScore),
          },
          {
            label: "Inbound Calls",
            value: record.inboundCalls === null || record.inboundCalls === undefined ? "N/A" : Number(record.inboundCalls).toFixed(0),
            score: null,
          },
          {
            label: "Inbound Minutes",
            value: record.inboundMinutes || "N/A",
            score: null,
          },
          {
            label: "Hold Time",
            value: record.holdTime || "N/A",
            score: null,
          },
        ]
    : state.dashboardFocus === "attendance"
      ? [
        {
          label: "Attendance",
          value: record.attendancePercentDisplay,
          score: formatScore(record.attendanceScore),
        },
      ]
      : [
        {
          label: "Quality Assurance",
          value: record.qaPercentDisplay,
          score: formatScore(record.qaScore),
        },
      ];

  state.distributionDrilldownOpen = true;
  state.distributionDrilldownExpanded = false;
  state.distributionDrilldownDetail = {
    mode: "breakdown",
    agentName: record.agentName,
    weekEnding: record.weekEnding,
    items: focusItems,
  };
  renderDistributionDrilldownContent();
  applyDistributionDrilldownState();
}

function initializeFloatingLegend() {
  if (!elements.legendPanel || !elements.floatingLegendContent) return;
  elements.floatingLegendContent.innerHTML = elements.legendPanel.innerHTML;
}

function initializeMobileDefaults() {
  const isMobile = typeof window !== "undefined" && window.innerWidth <= 720;
  state.mobileTableExpanded = !isMobile;
  state.mobileFiltersExpanded = !isMobile;
  state.mobileSidebarOpen = false;
}

const cardMap = {
  overall: {
    value: document.querySelector("#overallScoreValue"),
    badge: document.querySelector("#overallCardBadge"),
    raw: document.querySelector("#overallScoreRaw"),
    delta: document.querySelector("#overallScoreDelta"),
  },
  transfer: {
    value: document.querySelector("#transferScoreValue"),
    badge: document.querySelector("#transferCardBadge"),
    raw: document.querySelector("#transferScoreRaw"),
    delta: document.querySelector("#transferScoreDelta"),
  },
  admits: {
    value: document.querySelector("#admitsScoreValue"),
    badge: document.querySelector("#admitsCardBadge"),
    raw: document.querySelector("#admitsScoreRaw"),
    delta: document.querySelector("#admitsScoreDelta"),
  },
  aht: {
    value: document.querySelector("#ahtScoreValue"),
    badge: document.querySelector("#ahtCardBadge"),
    raw: document.querySelector("#ahtScoreRaw"),
    delta: document.querySelector("#ahtScoreDelta"),
  },
  attendance: {
    value: document.querySelector("#attendanceScoreValue"),
    badge: document.querySelector("#attendanceCardBadge"),
    raw: document.querySelector("#attendanceScoreRaw"),
    delta: document.querySelector("#attendanceScoreDelta"),
  },
  qa: {
    value: document.querySelector("#qaScoreValue"),
    badge: document.querySelector("#qaCardBadge"),
    raw: document.querySelector("#qaScoreRaw"),
    delta: document.querySelector("#qaScoreDelta"),
  },
};

function scoreClassName(score) {
  if (score === null || score === undefined || Number.isNaN(score)) return "score-na";
  return `score-${Math.max(1, Math.min(5, Math.round(score)))}`;
}

function formatScore(score) {
  if (score === null || score === undefined || Number.isNaN(score)) return "N/A";
  return Number(score).toFixed(2);
}

function formatPercent(value, digits = 2) {
  if (value === null || value === undefined || Number.isNaN(value)) return "N/A";
  return `${Number(value).toFixed(digits)}%`;
}

function formatTimeFromSeconds(seconds) {
  if (!seconds) return "00:00";
  const mins = Math.floor(seconds / 60);
  const secs = Math.round(seconds % 60);
  return `${String(mins).padStart(2, "0")}:${String(secs).padStart(2, "0")}`;
}

function averageField(records, field) {
  const selector = typeof field === "function"
    ? field
    : (record) => String(field)
      .split(".")
      .reduce((value, key) => value?.[key], record);
  const valid = records
    .map((record) => selector(record))
    .filter((value) => typeof value === "number" && !Number.isNaN(value));
  if (!valid.length) return null;
  return valid.reduce((sum, value) => sum + value, 0) / valid.length;
}

function formatSignedDelta(delta) {
  if (delta === null || delta === undefined || Number.isNaN(delta)) return "N/A";
  const prefix = delta > 0 ? "+" : "";
  return `${prefix}${Number(delta).toFixed(2)}`;
}

function parseInputDate(value) {
  if (!value) return null;
  const parsed = new Date(`${value}T00:00:00`);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function formatInputDate(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatShortRange(dateStart, dateEnd) {
  if (!(dateStart instanceof Date) || !(dateEnd instanceof Date)) return "Custom date range";
  const options = { month: "2-digit", day: "2-digit" };
  return `${dateStart.toLocaleDateString("en-US", options)} - ${dateEnd.toLocaleDateString("en-US", options)}`;
}

// Filters CTM daily rows against the chosen inclusive start/end dates.
function filterByDateRange(data, start, end) {
  if (!(start instanceof Date) || Number.isNaN(start.getTime())) return [];
  if (!(end instanceof Date) || Number.isNaN(end.getTime())) return [];
  const startTime = start.getTime();
  const endTime = end.getTime();
  return data.filter((item) => {
    const dateValue = item?.dateAt instanceof Date ? item.dateAt : null;
    const time = dateValue?.getTime?.() ?? NaN;
    return !Number.isNaN(time) && time >= startTime && time <= endTime;
  });
}

// Recomputes the active KPI from the filtered CTM rows only.
function computeKPI(data, focus) {
  if (!Array.isArray(data) || !data.length) {
    return { score: null, rawText: "No valid records in the selected range.", valueText: "N/A" };
  }

  const perAgentRecords = aggregateDualRangeRecords(data, "Selected Date Range");

  if (focus === "transfer") {
    const transferRatePercent = averageField(perAgentRecords, "transferRatePercent");
    const transferScore = averageField(perAgentRecords, "transferScore");
    return {
      score: transferScore,
      valueText: formatScore(transferScore),
      rawText: `Primary range ${formatShortRange(data[0]?.dateAt, data.at(-1)?.dateAt)} | Mean agent transfer rate ${formatPercent(transferRatePercent)}`,
    };
  }

  if (focus === "aht") {
    const ahtSeconds = averageField(perAgentRecords, "ahtSeconds");
    const ahtScore = averageField(perAgentRecords, "ahtScore");
    return {
      score: ahtScore,
      valueText: formatScore(ahtScore),
      rawText: `Primary range ${formatShortRange(data[0]?.dateAt, data.at(-1)?.dateAt)} | Mean agent AHT ${formatTimeFromSeconds(ahtSeconds)}`,
    };
  }

  return { score: null, rawText: "N/A", valueText: "N/A" };
}

// Keeps delta math isolated so the scorecard render stays simple.
function computeDelta(current, compare) {
  if (
    current === null ||
    current === undefined ||
    Number.isNaN(current) ||
    compare === null ||
    compare === undefined ||
    Number.isNaN(compare)
  ) {
    return null;
  }
  return current - compare;
}

function getDualDateRangeState() {
  return state.dualDateRange[state.dashboardFocus] || null;
}

function isDualRangeFocus() {
  return ["transfer", "aht"].includes(state.dashboardFocus);
}

function getScopedRealtimeDailyRecords() {
  const dailyRecords = state.realtimeDataset?.dailyRecords || [];
  return dailyRecords.filter((record) => state.filters.agent === "all" || record.agentName === state.filters.agent);
}

function getAllRealtimeDailyRecordsForFocus() {
  return state.realtimeDataset?.dailyRecords || [];
}

function ensureDualDateRangeDefaults(force = false) {
  const dualState = getDualDateRangeState();
  if (!dualState) return;
  if (dualState.initialized && !force) return;

  const dailyRecords = getScopedRealtimeDailyRecords();
  const dated = dailyRecords
    .map((record) => record.dateAt)
    .filter((value) => value instanceof Date && !Number.isNaN(value.getTime()))
    .sort((left, right) => left.getTime() - right.getTime());
  const latest = dated.at(-1);
  if (!latest) return;

  dualState.primaryStart = formatInputDate(latest);
  dualState.primaryEnd = formatInputDate(latest);
  dualState.compareEnabled = false;
  dualState.compareStart = "";
  dualState.compareEnd = "";
  dualState.initialized = true;
}

function syncDualDateRangeControls() {
  if (!elements.dualDateRangeSection) return;
  const active = isDualRangeFocus();
  elements.dualDateRangeSection.classList.toggle("is-hidden", !active);
  elements.compareRangeHeaderToggle?.classList.toggle("is-hidden", !active);
  elements.topFiltersGrid?.classList.toggle("is-hidden", active);
  elements.monthFilterField?.classList.toggle("is-hidden", active);
  elements.weekFilterField?.classList.toggle("is-hidden", active);
  elements.agentFilterField?.classList.toggle("is-hidden", active);
  elements.resetFilters?.closest(".top-filter-action")?.classList.toggle("is-hidden", active);
  if (!active) return;

  ensureDualDateRangeDefaults();
  const dualState = getDualDateRangeState();
  if (!dualState) return;

  const agentOptions = [...new Set(getAllRealtimeDailyRecordsForFocus().map((record) => record.agentName))].sort((left, right) =>
    left.localeCompare(right)
  );
  if (elements.dualRangeAgentFilter) {
    populateSelect(elements.dualRangeAgentFilter, agentOptions, "All Agents");
    elements.dualRangeAgentFilter.value = state.filters.agent;
  }

  elements.primaryDateStart.value = dualState.primaryStart || "";
  elements.primaryDateEnd.value = dualState.primaryEnd || "";
  elements.compareDateToggle.checked = Boolean(dualState.compareEnabled);
  elements.compareDateStart.value = dualState.compareStart || "";
  elements.compareDateEnd.value = dualState.compareEnd || "";
  elements.compareDateStartField?.classList.toggle("is-hidden", !dualState.compareEnabled);
  elements.compareDateEndField?.classList.toggle("is-hidden", !dualState.compareEnabled);
}

function getDualDateRangeComparison() {
  const dualState = getDualDateRangeState();
  if (!dualState || !["transfer", "aht"].includes(state.dashboardFocus)) return null;

  ensureDualDateRangeDefaults();
  const scopedDailyRecords = getScopedRealtimeDailyRecords();
  const primaryStart = parseInputDate(dualState.primaryStart);
  const primaryEnd = parseInputDate(dualState.primaryEnd);
  const primaryRecords = filterByDateRange(scopedDailyRecords, primaryStart, primaryEnd)
    .sort((left, right) => (left.dateAt?.getTime() ?? 0) - (right.dateAt?.getTime() ?? 0));
  const currentMetrics = computeKPI(primaryRecords, state.dashboardFocus);

  if (!dualState.compareEnabled) {
    return { currentMetrics, compareMetrics: null, delta: null, compareEnabled: false };
  }

  const compareStart = parseInputDate(dualState.compareStart);
  const compareEnd = parseInputDate(dualState.compareEnd);
  const compareRecords = filterByDateRange(scopedDailyRecords, compareStart, compareEnd)
    .sort((left, right) => (left.dateAt?.getTime() ?? 0) - (right.dateAt?.getTime() ?? 0));
  const compareMetrics = computeKPI(compareRecords, state.dashboardFocus);
  const delta = computeDelta(currentMetrics.score, compareMetrics.score);
  return { currentMetrics, compareMetrics, delta, compareEnabled: true };
}

function getSelectedPrimaryRangeLabel() {
  const dualState = getDualDateRangeState();
  const start = parseInputDate(dualState?.primaryStart);
  const end = parseInputDate(dualState?.primaryEnd);
  if (!(start instanceof Date) || !(end instanceof Date)) return "Selected Date Range";
  return `${start.toLocaleDateString("en-US")} - ${end.toLocaleDateString("en-US")}`;
}

function aggregateDualRangeRecords(records, rangeLabel) {
  const grouped = new Map();

  records.forEach((record) => {
    const key = `${(record.email || "").toLowerCase()}__${record.agentName}`;
    const bucket = grouped.get(key) || [];
    bucket.push(record);
    grouped.set(key, bucket);
  });

  return [...grouped.values()].map((items) => {
    const base = items[0];
    const firstTimeCaller = items.reduce((sum, item) => sum + (Number(item.firstTimeCaller) || 0), 0);
    const transferCount = items.reduce((sum, item) => sum + (Number(item.transferCount) || 0), 0);
    const inboundCalls = items.reduce((sum, item) => sum + (Number(item.inboundCalls) || 0), 0);
    const inboundMinutesSeconds = items.reduce((sum, item) => sum + (Number(item.inboundMinutesSeconds) || 0), 0);
    const holdTimeSeconds = items.reduce((sum, item) => sum + (Number(item.holdTimeSeconds) || 0), 0);
    const transferRate = calculateTransferRate(
      transferCount === null || transferCount === undefined ? null : transferCount,
      firstTimeCaller === null || firstTimeCaller === undefined ? null : firstTimeCaller
    );
    const ahtSeconds = inboundCalls > 0 ? (inboundMinutesSeconds + holdTimeSeconds) / inboundCalls : null;
    const transferScore = calculateTransferScore(transferRate);
    const ahtScore = calculateAHTScore(ahtSeconds);

    return {
      ...base,
      key: `dual__${base.email || base.agentName}__${rangeLabel}`,
      weekEnding: rangeLabel,
      dateRange: rangeLabel,
      monthLabel: base.monthLabel,
      firstTimeCaller,
      transferCount,
      inboundCalls,
      inboundMinutesSeconds,
      inboundMinutes: formatTimeFromSeconds(inboundMinutesSeconds),
      holdTimeSeconds,
      holdTime: formatTimeFromSeconds(holdTimeSeconds),
      transferRate,
      transferRatePercent: transferRate === null ? null : transferRate * 100,
      transferRateDisplay: formatPercent(transferRate === null ? null : transferRate * 100),
      transferScore,
      ahtSeconds,
      ahtDisplay: formatTimeFromSeconds(ahtSeconds),
      ahtScore,
      performanceScore: calculatePerformanceScore(transferScore, null, ahtScore),
      overallScore: calculatePerformanceScore(transferScore, null, ahtScore),
      lastUpdatedDisplay: getLatestLastUpdatedDisplay(items),
      lastUpdatedAt: [...items]
        .map((item) => item.lastUpdatedAt)
        .filter((value) => value instanceof Date && !Number.isNaN(value.getTime()))
        .sort((left, right) => left.getTime() - right.getTime())
        .at(-1) || null,
    };
  });
}

function buildDualRangeSummaryFromRecords(records, rangeLabel) {
  const metrics = getSummaryMetrics(records);
  return {
    weekEnding: rangeLabel,
    weekDate: parseInputDate(getDualDateRangeState()?.primaryEnd),
    monthLabel: records[0]?.monthLabel || "Custom",
    transferScore: metrics?.transferScore ?? null,
    admitsScore: null,
    ahtScore: metrics?.ahtScore ?? null,
    attendanceScore: null,
    qaScore: null,
    performanceScore: metrics?.performanceScore ?? null,
    overallScore: metrics?.overallScore ?? null,
    overallIncludesQa: false,
    overallWeights: { performance: 1, attendance: 0, qa: 0 },
    transferRatePercent: metrics?.transferRatePercent ?? null,
    admitsCount: null,
    firstTimeCaller: metrics?.firstTimeCaller ?? null,
    transferCount: metrics?.transferCount ?? null,
    inboundCalls: metrics?.inboundCalls ?? null,
    inboundMinutesSeconds: metrics?.inboundMinutesSeconds ?? null,
    holdTimeSeconds: metrics?.holdTimeSeconds ?? null,
    ahtSeconds: metrics?.ahtSeconds ?? null,
    attendancePercentValue: null,
    qaPercentValue: null,
    agentCount: metrics?.agentCount ?? 0,
  };
}

function buildDualRangeTrendSeries(records) {
  const grouped = new Map();
  records.forEach((record) => {
    const key = record.date;
    const bucket = grouped.get(key) || [];
    bucket.push(record);
    grouped.set(key, bucket);
  });

  return [...grouped.entries()]
    .map(([dateKey, items]) => {
      const dateAt = items[0]?.dateAt || parseInputDate(dateKey);
      const rangeLabel = dateAt ? dateAt.toLocaleDateString("en-US") : dateKey;
      return {
        weekEnding: rangeLabel,
        weekDate: dateAt,
        monthLabel: items[0]?.monthLabel || "Custom",
        transferScore: averageField(items, "transferScore"),
        admitsScore: null,
        ahtScore: averageField(items, "ahtScore"),
        attendanceScore: null,
        qaScore: null,
        performanceScore: averageField(items, "performanceScore"),
        overallScore: averageField(items, "overallScore"),
        overallIncludesQa: false,
        overallWeights: { performance: 1, attendance: 0, qa: 0 },
        transferRatePercent: averageField(items, "transferRatePercent"),
        admitsCount: null,
        firstTimeCaller: averageField(items, "firstTimeCaller"),
        transferCount: averageField(items, "transferCount"),
        inboundCalls: averageField(items, "inboundCalls"),
        inboundMinutesSeconds: averageField(items, "inboundMinutesSeconds"),
        holdTimeSeconds: averageField(items, "holdTimeSeconds"),
        ahtSeconds: averageField(items, "ahtSeconds"),
        attendancePercentValue: null,
        qaPercentValue: null,
        agentCount: new Set(items.map((item) => item.agentName)).size,
      };
    })
    .sort((left, right) => (left.weekDate?.getTime() ?? 0) - (right.weekDate?.getTime() ?? 0));
}

function getDualRangeDashboardData() {
  const dualState = getDualDateRangeState();
  if (!dualState) return null;

  const primaryStart = parseInputDate(dualState.primaryStart);
  const primaryEnd = parseInputDate(dualState.primaryEnd);
  const primaryLabel = getSelectedPrimaryRangeLabel();
  const scopedDailyRecords = getScopedRealtimeDailyRecords();
  const primaryDailyRecords = filterByDateRange(scopedDailyRecords, primaryStart, primaryEnd)
    .sort((left, right) => (left.dateAt?.getTime() ?? 0) - (right.dateAt?.getTime() ?? 0));
  const dashboardRecords = aggregateDualRangeRecords(primaryDailyRecords, primaryLabel);
  const trendWeeklyAverages = buildDualRangeTrendSeries(primaryDailyRecords);
  const currentSummary = buildDualRangeSummaryFromRecords(dashboardRecords, primaryLabel);

  const weeklyAverages = [];
  let compareRecords = [];
  let compareLabel = "Compare Range";
  if (dualState.compareEnabled) {
    const compareStart = parseInputDate(dualState.compareStart);
    const compareEnd = parseInputDate(dualState.compareEnd);
    compareLabel = compareStart && compareEnd
      ? `${compareStart.toLocaleDateString("en-US")} - ${compareEnd.toLocaleDateString("en-US")}`
      : "Compare Range";
    const compareDailyRecords = filterByDateRange(scopedDailyRecords, compareStart, compareEnd)
      .sort((left, right) => (left.dateAt?.getTime() ?? 0) - (right.dateAt?.getTime() ?? 0));
    compareRecords = aggregateDualRangeRecords(compareDailyRecords, compareLabel);
    const compareSummary = buildDualRangeSummaryFromRecords(compareRecords, compareLabel);
    weeklyAverages.push(compareSummary);
  }
  weeklyAverages.push(currentSummary);

  const compareByAgent = new Map(compareRecords.map((record) => [record.agentName, record]));
  const varianceByAgent = dashboardRecords
    .map((record) => {
      const compareRecord = compareByAgent.get(record.agentName);
      const metricKey = state.dashboardFocus === "transfer" ? "transferScore" : "ahtScore";
      const primaryValue = record?.[metricKey];
      const compareValue = compareRecord?.[metricKey];
      if (
        primaryValue === null ||
        primaryValue === undefined ||
        Number.isNaN(primaryValue) ||
        compareValue === null ||
        compareValue === undefined ||
        Number.isNaN(compareValue)
      ) {
        return null;
      }
      return {
        agentName: record.agentName,
        primaryValue,
        compareValue,
        delta: Number(primaryValue) - Number(compareValue),
      };
    })
    .filter(Boolean);

  return {
    dashboardRecords,
    compareRecords,
    primaryLabel,
    compareLabel,
    varianceByAgent,
    tableRecords: state.filters.search
      ? dashboardRecords.filter((record) => `${record.agentName} ${record.weekEnding}`.toLowerCase().includes(state.filters.search.trim().toLowerCase()))
      : dashboardRecords,
    weeklyAverages,
    trendWeeklyAverages,
  };
}

function getLatestLastUpdatedDisplay(records) {
  const latestRecord = [...records]
    .filter((record) => record.lastUpdatedAt instanceof Date && !Number.isNaN(record.lastUpdatedAt.getTime()))
    .sort((left, right) => left.lastUpdatedAt.getTime() - right.lastUpdatedAt.getTime())
    .at(-1);
  return latestRecord?.lastUpdatedDisplay || "--";
}

function getLatestAdmitsLastUpdatedDisplay(records) {
  const latestRecord = [...records]
    .filter((record) => record.admitsLastUpdatedAt instanceof Date && !Number.isNaN(record.admitsLastUpdatedAt.getTime()))
    .sort((left, right) => left.admitsLastUpdatedAt.getTime() - right.admitsLastUpdatedAt.getTime())
    .at(-1);
  return latestRecord?.admitsLastUpdatedDisplay || "--";
}

function getLatestQaLastUpdatedDisplay(records) {
  const latestRecord = [...records]
    .filter((record) => record.qaLastUpdatedAt instanceof Date && !Number.isNaN(record.qaLastUpdatedAt.getTime()))
    .sort((left, right) => left.qaLastUpdatedAt.getTime() - right.qaLastUpdatedAt.getTime())
    .at(-1);
  return latestRecord?.qaLastUpdatedDisplay || "--";
}

function getComparisonUnitLabel() {
  return "date range";
}

function getLatestRealtimeWeek() {
  const monthFilter = state.filters.month === "all" ? state.realtimeDataset?.monthOptions?.at(-1) || "all" : state.filters.month;
  const availableWeeks = getAvailableWeeks(state.realtimeDataset?.records || [], monthFilter);
  return availableWeeks.at(-1) || "all";
}

function shouldShowRealtimeLastUpdated() {
  return ["transfer", "aht"].includes(state.dashboardFocus) || (isRealtimeFocus() && state.filters.week === getLatestRealtimeWeek());
}

function shouldShowFocusLastUpdated() {
  return shouldShowRealtimeLastUpdated() || ["admits", "attendance", "qa"].includes(state.dashboardFocus);
}

function getFocusLastUpdatedDisplay(records) {
  if (state.dashboardFocus === "admits") {
    return getLatestAdmitsLastUpdatedDisplay(records);
  }
  if (state.dashboardFocus === "attendance") {
    const latestRecord = [...records]
      .filter((record) => record.attendanceLastUpdatedAt instanceof Date && !Number.isNaN(record.attendanceLastUpdatedAt.getTime()))
      .sort((left, right) => left.attendanceLastUpdatedAt.getTime() - right.attendanceLastUpdatedAt.getTime())
      .at(-1);
    return latestRecord?.attendanceLastUpdatedDisplay || "--";
  }
  if (state.dashboardFocus === "qa") {
    return getLatestQaLastUpdatedDisplay(records);
  }
  if (["transfer", "aht", "realtime"].includes(state.dashboardFocus)) {
    return getLatestLastUpdatedDisplay(records);
  }
  return "--";
}

function getRankedKpis(summary) {
  if (!summary) return [];
  return KPI_DEFINITIONS
    .map((kpi) => ({
      ...kpi,
      score: summary[kpi.key],
      rawText: kpi.raw(summary),
    }))
    .filter((item) => typeof item.score === "number" && !Number.isNaN(item.score))
    .sort((left, right) => right.score - left.score || left.label.localeCompare(right.label));
}

function getFocusRankedKpis(summary) {
  if (!summary) return [];
  const scoped = state.dashboardFocus === "all"
    ? KPI_DEFINITIONS
    : state.dashboardFocus === "performance" || state.dashboardFocus === "realtime"
    ? KPI_DEFINITIONS.filter((item) => ["transferScore", "admitsScore", "ahtScore"].includes(item.key))
    : isSingleKpiFocus()
    ? KPI_DEFINITIONS.filter((item) => item.key === SINGLE_KPI_SCORE_KEYS[state.dashboardFocus])
    : KPI_DEFINITIONS.filter((item) => item.key === (state.dashboardFocus === "attendance" ? "attendanceScore" : "qaScore"));

  return scoped
    .map((kpi) => ({
      ...kpi,
      score: summary[kpi.key],
      rawText: kpi.raw(summary),
    }))
    .filter((item) => typeof item.score === "number" && !Number.isNaN(item.score))
    .sort((left, right) => right.score - left.score || left.label.localeCompare(right.label));
}

function getSummaryMetrics(records) {
  if (!records.length) return null;
  return {
    transferScore: averageField(records, "transferScore"),
    admitsScore: averageField(records, "admitsScore"),
    ahtScore: averageField(records, "ahtScore"),
    attendanceScore: averageField(records, "attendanceScore"),
    qaScore: averageField(records, "qaScore"),
    performanceScore: averageField(records, "performanceScore"),
    overallScore: averageField(records, "overallScore"),
    transferRatePercent: averageField(records, "transferRatePercent"),
    admitsCount: averageField(records, "admitsCount"),
    firstTimeCaller: averageField(records, "firstTimeCaller"),
    transferCount: averageField(records, "transferCount"),
    inboundCalls: averageField(records, "inboundCalls"),
    inboundMinutesSeconds: averageField(records, "inboundMinutesSeconds"),
    holdTimeSeconds: averageField(records, "holdTimeSeconds"),
    ahtSeconds: averageField(records, "ahtSeconds"),
    attendancePercentValue: averageField(records, "attendancePercentValue"),
    qaPercentValue: averageField(records, "qaPercentValue"),
    agentCount: new Set(records.map((record) => record.agentName)).size,
    overallIncludesQa: records.some((record) => hasValidNumber(record.qaScore)),
    overallWeights: {
      performance: averageField(records, "overallWeights.performance"),
      attendance: averageField(records, "overallWeights.attendance"),
      qa: averageField(records, "overallWeights.qa"),
    },
  };
}

function getScopedWeeklyAverages(records) {
  const weekMap = new Map();
  records.forEach((record) => {
    const bucket = weekMap.get(record.weekEnding) || [];
    bucket.push(record);
    weekMap.set(record.weekEnding, bucket);
  });

  return [...weekMap.entries()]
    .map(([weekEnding, items]) => ({
      weekEnding,
      weekDate: items[0]?.weekDate,
      transferScore: averageField(items, "transferScore"),
      admitsScore: averageField(items, "admitsScore"),
      ahtScore: averageField(items, "ahtScore"),
      attendanceScore: averageField(items, "attendanceScore"),
      qaScore: averageField(items, "qaScore"),
      performanceScore: averageField(items, "performanceScore"),
      overallScore: averageField(items, "overallScore"),
      transferRatePercent: averageField(items, "transferRatePercent"),
      admitsCount: averageField(items, "admitsCount"),
      firstTimeCaller: averageField(items, "firstTimeCaller"),
      transferCount: averageField(items, "transferCount"),
      inboundCalls: averageField(items, "inboundCalls"),
      inboundMinutesSeconds: averageField(items, "inboundMinutesSeconds"),
      holdTimeSeconds: averageField(items, "holdTimeSeconds"),
      ahtSeconds: averageField(items, "ahtSeconds"),
      attendancePercentValue: averageField(items, "attendancePercentValue"),
      qaPercentValue: averageField(items, "qaPercentValue"),
      overallIncludesQa: items.some((item) => hasValidNumber(item.qaScore)),
      overallWeights: {
        performance: averageField(items, "overallWeights.performance"),
        attendance: averageField(items, "overallWeights.attendance"),
        qa: averageField(items, "overallWeights.qa"),
      },
    }))
    .sort((left, right) => (left.weekDate?.getTime() ?? 0) - (right.weekDate?.getTime() ?? 0));
}

function pickWeekSummary(filteredRecords, weeklyAverages) {
  if (!weeklyAverages.length) {
    return {
      transferScore: 0,
      admitsScore: 0,
      ahtScore: 0,
      attendanceScore: 0,
      qaScore: 0,
      performanceScore: 0,
      overallScore: 0,
      transferRatePercent: null,
      admitsCount: null,
      firstTimeCaller: null,
      transferCount: null,
      inboundCalls: null,
      inboundMinutesSeconds: null,
      holdTimeSeconds: null,
      ahtSeconds: null,
      attendancePercentValue: null,
      qaPercentValue: null,
      overallIncludesQa: false,
      overallWeights: {
        performance: 0,
        attendance: 0,
        qa: 0,
      },
      weekEnding: "No Data",
    };
  }

  if (state.filters.week !== "all") {
    return weeklyAverages.find((item) => item.weekEnding === state.filters.week) || weeklyAverages.at(-1);
  }

  if (filteredRecords.length) {
    const latestWeek = [...filteredRecords]
      .sort((left, right) => (left.weekDate?.getTime() ?? 0) - (right.weekDate?.getTime() ?? 0))
      .at(-1)?.weekEnding;
    return weeklyAverages.find((item) => item.weekEnding === latestWeek) || weeklyAverages.at(-1);
  }

  return weeklyAverages.at(-1);
}

function getComparisonWeeklyAverages(weeklyAverages, selectedWeekEnding) {
  if (!weeklyAverages.length) return [];

  const comparisonRecords = getFilteredRecords(getActiveDataset().records, {
    ...state.filters,
    month: "all",
    week: "all",
    search: "",
  });
  const comparisonPool = getScopedWeeklyAverages(comparisonRecords);
  if (!comparisonPool.length) return weeklyAverages;
  if (comparisonPool.length === 1) return comparisonPool;

  const resolvedSelectedWeek =
    selectedWeekEnding ||
    (state.filters.week !== "all" ? state.filters.week : weeklyAverages.at(-1)?.weekEnding) ||
    comparisonPool.at(-1)?.weekEnding;

  const selectedEntry = comparisonPool.find((item) => item.weekEnding === resolvedSelectedWeek);
  if (!selectedEntry?.weekDate) return comparisonPool.slice(-2);

  if (isRealtimeFocus()) {
    const eligibleRanges = comparisonPool.filter((item) => (item.weekDate?.getTime() ?? 0) <= selectedEntry.weekDate.getTime());
    if (!eligibleRanges.length) return [selectedEntry];
    return eligibleRanges.slice(-2);
  }

  const eligibleWeeks = comparisonPool.filter((item) => (item.weekDate?.getTime() ?? 0) <= selectedEntry.weekDate.getTime());
  if (!eligibleWeeks.length) return comparisonPool.slice(-2);
  return eligibleWeeks.slice(-2);
}

function getTrendWeeklyAverages(weeklyAverages, selectedWeekEnding) {
  if (!weeklyAverages.length) return [];

  const trendRecords = getFilteredRecords(getActiveDataset().records, {
    ...state.filters,
    month: "all",
    week: "all",
    search: "",
  });
  const trendPool = getScopedWeeklyAverages(trendRecords);
  if (!trendPool.length) return weeklyAverages;

  const resolvedSelectedWeek =
    selectedWeekEnding ||
    (state.filters.week !== "all" ? state.filters.week : weeklyAverages.at(-1)?.weekEnding) ||
    trendPool.at(-1)?.weekEnding;

  const selectedEntry = trendPool.find((item) => item.weekEnding === resolvedSelectedWeek);
  if (!selectedEntry?.weekDate) return trendPool.slice(-6);

  const eligibleWeeks = trendPool.filter((item) => (item.weekDate?.getTime() ?? 0) <= selectedEntry.weekDate.getTime());
  if (!eligibleWeeks.length) return trendPool.slice(-6);
  return eligibleWeeks.slice(-6);
}

function updateInsights(filteredRecords, weeklyAverages) {
  const selectedSummary = pickWeekSummary(filteredRecords, weeklyAverages);
  const comparisonWeeklyAverages = getComparisonWeeklyAverages(weeklyAverages, selectedSummary.weekEnding);
  const rankedKpis = getFocusRankedKpis(selectedSummary);
  const bestKpi = rankedKpis[0] || null;
  const weakestKpi = rankedKpis.at(-1) || null;
  const totalAgents = new Set(filteredRecords.map((record) => record.agentName)).size;
  const totalWeeks = new Set(filteredRecords.map((record) => record.weekEnding)).size;
  const isAgentView = state.filters.agent !== "all";

  elements.viewInsightTitle.textContent = isAgentView
    ? `${state.filters.agent} in focus`
    : `${totalAgents || 0} agents in current scope`;
  elements.viewInsightBody.textContent = isAgentView
    ? isRealtimeFocus()
      ? `${selectedSummary.weekEnding || "Latest date range"} is selected across ${totalWeeks || 0} available date range(s) for this agent.`
      : `${selectedSummary.weekEnding || "Latest date range"} is selected across ${totalWeeks || 0} available date range(s) for this agent.`
    : isRealtimeFocus()
    ? `${state.filters.week === "all" ? "All available date ranges" : state.filters.week} is showing ${filteredRecords.length} live row(s) for the current scope.`
    : `${state.filters.month === "all" ? "All months" : state.filters.month} with ${state.filters.week === "all" ? "all available date ranges" : state.filters.week} is showing ${filteredRecords.length} score row(s).`;

  elements.bestInsightTitle.textContent = bestKpi
    ? `${bestKpi.label} is leading at ${formatScore(bestKpi.score)}`
    : "No KPI lead available";
  elements.bestInsightBody.textContent = bestKpi
    ? `${bestKpi.rawText}. This is the strongest signal in the currently selected view.`
    : "Select a scope with KPI data to surface the current strength.";

  elements.watchInsightTitle.textContent = weakestKpi
    ? `${weakestKpi.label} needs the most attention`
    : "No watch item available";
  elements.watchInsightBody.textContent = weakestKpi
    ? `${weakestKpi.rawText}. This is the lowest-scoring KPI in the current selection.`
    : "The watchlist will appear once KPI scores are available.";

  const currentWeek = comparisonWeeklyAverages.at(-1) || null;
  const previousWeek = comparisonWeeklyAverages.length > 1 ? comparisonWeeklyAverages[0] : null;
  const momentumKey = state.dashboardFocus === "all"
    ? "overallScore"
    : isPerformanceLikeFocus()
    ? "performanceScore"
    : state.dashboardFocus === "attendance"
      ? "attendanceScore"
      : "qaScore";
  const overallDelta =
    currentWeek && previousWeek ? Number(currentWeek[momentumKey]) - Number(previousWeek[momentumKey]) : null;

  elements.momentumInsightTitle.textContent =
    overallDelta === null || Number.isNaN(overallDelta)
      ? isRealtimeFocus() ? "No previous range for momentum" : "No previous week for momentum"
      : `${FOCUS_SCORING[state.dashboardFocus].title} ${overallDelta >= 0 ? "up" : "down"} ${formatSignedDelta(overallDelta)}`;
  elements.momentumInsightBody.textContent =
    overallDelta === null || Number.isNaN(overallDelta)
      ? isRealtimeFocus()
        ? `${selectedSummary.weekEnding || "Selected date range"} has no earlier date range available for comparison.`
        : `${selectedSummary.weekEnding || "Selected date range"} has no earlier date range available for comparison.`
      : isRealtimeFocus()
      ? `${currentWeek.weekEnding} versus previous date range (${previousWeek.weekEnding}) on ${FOCUS_SCORING[state.dashboardFocus].title.toLowerCase()} score.`
      : `${currentWeek.weekEnding} versus ${previousWeek.weekEnding} on ${FOCUS_SCORING[state.dashboardFocus].title.toLowerCase()} score.`;

  if (elements.mobileStatusMonth && elements.mobileStatusWeek && elements.mobileStatusScope) {
    elements.mobileStatusMonth.textContent = isRealtimeFocus()
      ? "Real-Time Performance"
      : state.filters.month === "all"
      ? "All Months"
      : state.filters.month;
    elements.mobileStatusWeek.textContent = state.filters.week === "all" ? selectedSummary.weekEnding || "Latest Range" : state.filters.week;
    elements.mobileStatusScope.textContent = state.filters.agent === "all" ? `${totalAgents || 0} Agents` : state.filters.agent;
  }

  if (elements.mobileLegendLauncher) {
    const legendTitle = elements.mobileLegendLauncher.querySelector("strong");
    const legendCopy = elements.mobileLegendLauncher.querySelector("p");
    if (legendTitle && legendCopy) {
      legendTitle.textContent = "KPI Legend";
      legendCopy.textContent = "Keep score definitions one tap away.";
    }
  }

  if (elements.mobileFiltersLauncher) {
    const filtersTitle = elements.mobileFiltersLauncher.querySelector("strong");
    const filtersCopy = elements.mobileFiltersLauncher.querySelector("p");
    if (filtersTitle && filtersCopy) {
      if (state.filters.agent === "all") {
        filtersTitle.textContent = "Filters";
        filtersCopy.textContent = "Open month, date range, and agent filters.";
      } else {
        filtersTitle.textContent = "Agent Scope";
        filtersCopy.textContent = "Change the selected month, date range, or agent focus.";
      }
    }
  }
}

function updateAgentFocus(filteredRecords, weeklyAverages) {
  if (state.filters.agent === "all") {
    elements.agentFocusName.textContent = "Selected agent";
    elements.agentFocusNarrative.textContent = "Pick one agent to show an individualized coaching snapshot.";
    elements.agentOverallValue.textContent = "--";
    elements.agentOverallNarrative.textContent = "Latest weighted score summary.";
    elements.agentStrongestValue.textContent = "--";
    elements.agentStrongestNarrative.textContent = "Best-performing KPI in the latest date range.";
    elements.agentWeakestValue.textContent = "--";
    elements.agentWeakestNarrative.textContent = "Most urgent KPI to coach in the latest date range.";
    return;
  }

  const selectedSummary = pickWeekSummary(filteredRecords, weeklyAverages);
  const rankedKpis = getFocusRankedKpis(selectedSummary);
  const bestKpi = rankedKpis[0] || null;
  const weakestKpi = rankedKpis.at(-1) || null;
  const comparisonWeeklyAverages = getComparisonWeeklyAverages(weeklyAverages, selectedSummary.weekEnding);
  const previousWeek = comparisonWeeklyAverages.length > 1 ? comparisonWeeklyAverages[0] : null;
  const currentWeek = comparisonWeeklyAverages.at(-1) || null;
  const overallDelta =
    currentWeek && previousWeek ? Number(currentWeek.overallScore) - Number(previousWeek.overallScore) : null;

  elements.agentFocusName.textContent = state.filters.agent;
  elements.agentFocusNarrative.textContent = `${weeklyAverages.length} tracked date range(s) in the current month filter. Use this as a quick coaching snapshot before reviewing the charts below.`;
  if (isRealtimeFocus()) {
    elements.agentFocusNarrative.textContent = `${weeklyAverages.length} tracked date range(s) in the real-time feed. Use this as a quick coaching snapshot before reviewing the charts below.`;
  }
  elements.agentOverallValue.textContent = formatScore(selectedSummary.overallScore);
  elements.agentOverallNarrative.textContent =
    overallDelta === null || Number.isNaN(overallDelta)
      ? isRealtimeFocus()
        ? `Latest available date range: ${selectedSummary.weekEnding || "No Data"}.`
        : `Latest available date range: ${selectedSummary.weekEnding || "No Data"}.`
      : isRealtimeFocus()
      ? `${currentWeek.weekEnding} moved ${formatSignedDelta(overallDelta)} versus previous date range (${previousWeek.weekEnding}).`
      : `${currentWeek.weekEnding} moved ${formatSignedDelta(overallDelta)} versus ${previousWeek.weekEnding}.`;
  elements.agentStrongestValue.textContent = bestKpi ? `${bestKpi.label} ${formatScore(bestKpi.score)}` : "N/A";
  elements.agentStrongestNarrative.textContent = bestKpi
    ? `${bestKpi.rawText}. This is the best KPI to reinforce.`
    : "No strongest KPI available for the selected view.";
  elements.agentWeakestValue.textContent = weakestKpi ? `${weakestKpi.label} ${formatScore(weakestKpi.score)}` : "N/A";
  elements.agentWeakestNarrative.textContent = weakestKpi
    ? `${weakestKpi.rawText}. This is the KPI to prioritize in coaching.`
    : "No weakest KPI available for the selected view.";
}

function getMostImprovedAgents(records, comparisonWeeklyAverages, limit = 5, metricKey = "overallScore") {
  if (comparisonWeeklyAverages.length < 2) return [];

  const previousWeek = comparisonWeeklyAverages[0]?.weekEnding;
  const currentWeek = comparisonWeeklyAverages.at(-1)?.weekEnding;
  if (!previousWeek || !currentWeek) return [];

  const previousByAgent = new Map();
  const currentByAgent = new Map();

  records.forEach((record) => {
    if (record.weekEnding === previousWeek) previousByAgent.set(record.agentName, record);
    if (record.weekEnding === currentWeek) currentByAgent.set(record.agentName, record);
  });

  return [...currentByAgent.entries()]
    .map(([agentName, currentRecord]) => {
      const previousRecord = previousByAgent.get(agentName);
      if (!previousRecord) return null;
      if (
        typeof currentRecord[metricKey] !== "number" ||
        Number.isNaN(currentRecord[metricKey]) ||
        typeof previousRecord[metricKey] !== "number" ||
        Number.isNaN(previousRecord[metricKey])
      ) {
        return null;
      }

      return {
        agentName,
        delta: currentRecord[metricKey] - previousRecord[metricKey],
        currentOverallScore: currentRecord[metricKey],
        previousOverallScore: previousRecord[metricKey],
      };
    })
    .filter((record) => record && record.delta > 0)
    .sort((left, right) => right.delta - left.delta)
    .slice(0, limit);
}

function buildDeltaMarkup(current, previous) {
  const comparisonUnit = getComparisonUnitLabel();
  if (previous === null || previous === undefined) {
    return `<span class="delta-flat">Flat</span> No previous ${comparisonUnit} available`;
  }
  const delta = current - previous;
  if (delta > 0) return `<span class="delta-up">Up</span> ${Math.abs(delta).toFixed(2)} vs previous ${comparisonUnit}`;
  if (delta < 0) return `<span class="delta-down">Down</span> ${Math.abs(delta).toFixed(2)} vs previous ${comparisonUnit}`;
  return `<span class="delta-flat">Flat</span> 0.00 vs previous ${comparisonUnit}`;
}

function buildTextDelta(current, previous, formatter) {
  const comparisonUnit = getComparisonUnitLabel();
  if (current === null || current === undefined) return `Vs previous ${comparisonUnit} unavailable`;
  if (previous === null || previous === undefined) return `Vs previous ${comparisonUnit} unavailable`;
  return `${formatter(current)} vs ${formatter(previous)} previous ${comparisonUnit}`;
}

function buildCurrentPreviousDetail(label, currentValue, previousValue, formatter) {
  const currentText = currentValue === null || currentValue === undefined ? "N/A" : formatter(currentValue);
  const previousText = previousValue === null || previousValue === undefined ? "N/A" : formatter(previousValue);
  return `${label}: ${currentText} | Previous: ${previousText}`;
}

function hasValidNumber(value) {
  return typeof value === "number" && !Number.isNaN(value);
}

function describeOverallFormula(summary) {
  if (!summary) return "Overall score pending";
  if (summary.overallIncludesQa) {
    return "Performance + Attendance + QA";
  }
  return "Performance + Attendance only | QA pending";
}

function getRawMetricCards(metrics, currentSummary, previousSummary) {
  if (!metrics) return [];

  if (state.dashboardFocus === "all") {
    return [
      ["Average Transfer Rate", formatPercent(metrics.transferRatePercent), buildTextDelta(currentSummary?.transferRatePercent, previousSummary?.transferRatePercent, (value) => formatPercent(value))],
      ["Average Admit Count", metrics.admitsCount === null || metrics.admitsCount === undefined ? "N/A" : metrics.admitsCount.toFixed(2), buildTextDelta(currentSummary?.admitsCount, previousSummary?.admitsCount, (value) => Number(value).toFixed(2))],
      ["Average AHT", formatTimeFromSeconds(metrics.ahtSeconds), buildTextDelta(currentSummary?.ahtSeconds, previousSummary?.ahtSeconds, (value) => formatTimeFromSeconds(value))],
      ["Average Attendance", formatPercent(metrics.attendancePercentValue), buildTextDelta(currentSummary?.attendancePercentValue, previousSummary?.attendancePercentValue, (value) => formatPercent(value))],
      ["Average Quality Assurance", formatPercent(metrics.qaPercentValue), buildTextDelta(currentSummary?.qaPercentValue, previousSummary?.qaPercentValue, (value) => formatPercent(value))],
      ["Agents in Scope", String(metrics.agentCount), buildTextDelta(currentSummary?.agentCount, previousSummary?.agentCount, (value) => `${Math.round(value)} agents`)],
    ];
  }

  if (isPerformanceLikeFocus()) {
    if (state.dashboardFocus === "transfer") {
      return [
        ["Average Transfer Rate", formatPercent(metrics.transferRatePercent), buildTextDelta(currentSummary?.transferRatePercent, previousSummary?.transferRatePercent, (value) => formatPercent(value))],
        ["Transfer Score", formatScore(metrics.transferScore), buildTextDelta(currentSummary?.transferScore, previousSummary?.transferScore, (value) => formatScore(value))],
        ["First-Time Callers", metrics.firstTimeCaller === null || metrics.firstTimeCaller === undefined ? "N/A" : metrics.firstTimeCaller.toFixed(2), buildTextDelta(currentSummary?.firstTimeCaller, previousSummary?.firstTimeCaller, (value) => Number(value).toFixed(2))],
        ["Transfer Count", metrics.transferCount === null || metrics.transferCount === undefined ? "N/A" : metrics.transferCount.toFixed(2), buildTextDelta(currentSummary?.transferCount, previousSummary?.transferCount, (value) => Number(value).toFixed(2))],
        ["Agents in Scope", String(metrics.agentCount), buildTextDelta(currentSummary?.agentCount, previousSummary?.agentCount, (value) => `${Math.round(value)} agents`)],
      ];
    }
    if (state.dashboardFocus === "admits") {
      return [
        ["Average Admit Count", metrics.admitsCount === null || metrics.admitsCount === undefined ? "N/A" : metrics.admitsCount.toFixed(2), buildTextDelta(currentSummary?.admitsCount, previousSummary?.admitsCount, (value) => Number(value).toFixed(2))],
        ["Admits Score", formatScore(metrics.admitsScore), buildTextDelta(currentSummary?.admitsScore, previousSummary?.admitsScore, (value) => formatScore(value))],
        ["Agents in Scope", String(metrics.agentCount), buildTextDelta(currentSummary?.agentCount, previousSummary?.agentCount, (value) => `${Math.round(value)} agents`)],
      ];
    }
    if (state.dashboardFocus === "aht") {
      return [
        ["Average AHT", formatTimeFromSeconds(metrics.ahtSeconds), buildTextDelta(currentSummary?.ahtSeconds, previousSummary?.ahtSeconds, (value) => formatTimeFromSeconds(value))],
        ["Inbound Calls", metrics.inboundCalls === null || metrics.inboundCalls === undefined ? "N/A" : metrics.inboundCalls.toFixed(2), buildTextDelta(currentSummary?.inboundCalls, previousSummary?.inboundCalls, (value) => Number(value).toFixed(2))],
        ["Inbound Minutes", formatTimeFromSeconds(metrics.inboundMinutesSeconds), buildTextDelta(currentSummary?.inboundMinutesSeconds, previousSummary?.inboundMinutesSeconds, (value) => formatTimeFromSeconds(value))],
        ["Hold Time", formatTimeFromSeconds(metrics.holdTimeSeconds), buildTextDelta(currentSummary?.holdTimeSeconds, previousSummary?.holdTimeSeconds, (value) => formatTimeFromSeconds(value))],
        ["Agents in Scope", String(metrics.agentCount), buildTextDelta(currentSummary?.agentCount, previousSummary?.agentCount, (value) => `${Math.round(value)} agents`)],
      ];
    }
    return [
      ["Average Transfer Rate", formatPercent(metrics.transferRatePercent), buildTextDelta(currentSummary?.transferRatePercent, previousSummary?.transferRatePercent, (value) => formatPercent(value))],
      ["Average Admit Count", metrics.admitsCount === null || metrics.admitsCount === undefined ? "N/A" : metrics.admitsCount.toFixed(2), buildTextDelta(currentSummary?.admitsCount, previousSummary?.admitsCount, (value) => Number(value).toFixed(2))],
      ["Average AHT", formatTimeFromSeconds(metrics.ahtSeconds), buildTextDelta(currentSummary?.ahtSeconds, previousSummary?.ahtSeconds, (value) => formatTimeFromSeconds(value))],
      ["Average Inbound Calls", metrics.inboundCalls === null || metrics.inboundCalls === undefined ? "N/A" : metrics.inboundCalls.toFixed(2), buildTextDelta(currentSummary?.inboundCalls, previousSummary?.inboundCalls, (value) => Number(value).toFixed(2))],
      ["Agents in Scope", String(metrics.agentCount), buildTextDelta(currentSummary?.agentCount, previousSummary?.agentCount, (value) => `${Math.round(value)} agents`)],
    ];
  }

  if (state.dashboardFocus === "attendance") {
    return [
      ["Average Attendance", formatPercent(metrics.attendancePercentValue), buildTextDelta(currentSummary?.attendancePercentValue, previousSummary?.attendancePercentValue, (value) => formatPercent(value))],
      ["Attendance Score", formatScore(metrics.attendanceScore), buildTextDelta(currentSummary?.attendanceScore, previousSummary?.attendanceScore, (value) => formatScore(value))],
      ["Agents in Scope", String(metrics.agentCount), buildTextDelta(currentSummary?.agentCount, previousSummary?.agentCount, (value) => `${Math.round(value)} agents`)],
    ];
  }

  return [
    ["Average QA", formatPercent(metrics.qaPercentValue), buildTextDelta(currentSummary?.qaPercentValue, previousSummary?.qaPercentValue, (value) => formatPercent(value))],
    ["QA Score", formatScore(metrics.qaScore), buildTextDelta(currentSummary?.qaScore, previousSummary?.qaScore, (value) => formatScore(value))],
    ["Agents in Scope", String(metrics.agentCount), buildTextDelta(currentSummary?.agentCount, previousSummary?.agentCount, (value) => `${Math.round(value)} agents`)],
  ];
}

function describeSpreadGroup(records, metricKey, label, predicate) {
  const matches = records
    .filter((record) => typeof record[metricKey] === "number" && !Number.isNaN(record[metricKey]))
    .filter((record) => predicate(Number(record[metricKey])))
    .sort((left, right) => right[metricKey] - left[metricKey]);
  const names = matches.map((record) => getSpreadDisplayName(record.agentName)).join(", ");
  return {
    label,
    count: matches.length,
    names: names || "No agents in this group",
  };
}

function getSpreadDisplayName(agentName) {
  if (!agentName) return "";
  const parts = String(agentName).split(",");
  if (parts.length < 2) return String(agentName).trim();
  return parts.slice(1).join(",").trim() || String(agentName).trim();
}

function renderPerformerSpreadSummary(records, metricKey) {
  if (!elements.performerSpreadSummary) return;
  const scored = records.filter((record) => typeof record[metricKey] === "number" && !Number.isNaN(record[metricKey]));
  if (!scored.length) {
    elements.performerSpreadSummary.innerHTML = "";
    return;
  }

  if (state.dashboardFocus === "qa") {
    const totalAgents = new Set(filteredRecords.map((record) => record.agentName)).size;
    const totalDates = new Set(filteredRecords.map((record) => record.weekEnding)).size;
    const latestUpdatedDisplay = getLatestQaLastUpdatedDisplay(filteredRecords);
    elements.dataStatusText.textContent =
      state.filters.agent === "all"
        ? `${filteredRecords.length} QA rows across ${totalAgents} agents and ${totalDates} date range(s).${latestUpdatedDisplay !== "--" ? ` Last updated ${latestUpdatedDisplay}.` : ""}`
        : `${filteredRecords.length} QA rows for ${state.filters.agent} across ${totalDates} date range(s).${latestUpdatedDisplay !== "--" ? ` Last updated ${latestUpdatedDisplay}.` : ""}`;
    if (elements.realtimeLastUpdatedText) {
      elements.realtimeLastUpdatedText.textContent = latestUpdatedDisplay;
    }
    if (elements.realtimeLastUpdatedPill) {
      elements.realtimeLastUpdatedPill.hidden = latestUpdatedDisplay === "--";
    }
    return;
  }

  const groups = [
    describeSpreadGroup(scored, metricKey, "Top", (value) => value >= 4),
    describeSpreadGroup(scored, metricKey, "Middle", (value) => value >= 3 && value < 4),
    describeSpreadGroup(scored, metricKey, "Needs Support", (value) => value < 3),
  ];

  elements.performerSpreadSummary.innerHTML = groups.map((group) => `
    <article class="performer-spread-card performer-spread-card-${group.label.toLowerCase().replace(/\s+/g, "-")}">
      <span class="performer-spread-label">${group.label}</span>
      <strong>${group.count}</strong>
      <p>${group.names}</p>
    </article>
  `).join("");
}

function renderRawMetricCards(metrics, currentSummary, previousSummary) {
  const cards = getRawMetricCards(metrics, currentSummary, previousSummary);
  elements.rawMetricsGrid.innerHTML = cards.map(([label, value, delta]) => `
    <article class="raw-card card">
      <span class="raw-label">${label}</span>
      <strong>${value}</strong>
      <p class="raw-delta">${delta}</p>
    </article>
  `).join("");
}

function updateSummaryCards(filteredRecords, weeklyAverages) {
  const metrics = getSummaryMetrics(filteredRecords);

  if (!metrics) {
    if (elements.rawMetricsGrid) {
      elements.rawMetricsGrid.innerHTML = "";
    }
    Object.values(cardMap).forEach((card) => {
      card.value.textContent = "--";
      card.badge.textContent = "-";
      card.badge.className = "score-chip";
      card.raw.textContent = "No data for the selected filters";
      card.delta.innerHTML = "No data for the selected filters";
    });
    return;
  }

  const currentWeekSummary = pickWeekSummary(filteredRecords, weeklyAverages);
  const comparisonWeeklyAverages = getComparisonWeeklyAverages(weeklyAverages, currentWeekSummary.weekEnding);
  const previousWeekSummary = comparisonWeeklyAverages.length > 1 ? comparisonWeeklyAverages[0] : null;
  renderRawMetricCards(metrics, currentWeekSummary, previousWeekSummary);

  const rawText = {
    overall: `${describeOverallFormula(currentWeekSummary)} | Performance ${buildCurrentPreviousDetail("current", currentWeekSummary.performanceScore, previousWeekSummary?.performanceScore, formatScore)} | Attendance ${buildCurrentPreviousDetail("current", currentWeekSummary.attendanceScore, previousWeekSummary?.attendanceScore, formatScore)} | QA ${buildCurrentPreviousDetail("current", currentWeekSummary.qaScore, previousWeekSummary?.qaScore, formatScore)}`,
    transfer: buildCurrentPreviousDetail(
      "Avg transfer rate",
      currentWeekSummary.transferRatePercent,
      previousWeekSummary?.transferRatePercent,
      (value) => formatPercent(value)
    ),
    admits: buildCurrentPreviousDetail(
      "Avg admits count",
      currentWeekSummary.admitsCount,
      previousWeekSummary?.admitsCount,
      (value) => Number(value).toFixed(2)
    ),
    aht: buildCurrentPreviousDetail(
      "Avg handle time",
      currentWeekSummary.ahtSeconds,
      previousWeekSummary?.ahtSeconds,
      (value) => formatTimeFromSeconds(value)
    ),
    attendance: buildCurrentPreviousDetail(
      "Avg attendance",
      currentWeekSummary.attendancePercentValue,
      previousWeekSummary?.attendancePercentValue,
      (value) => formatPercent(value)
    ),
    qa: buildCurrentPreviousDetail(
      "Avg QA",
      currentWeekSummary.qaPercentValue,
      previousWeekSummary?.qaPercentValue,
      (value) => formatPercent(value)
    ),
  };

  if (!hasValidNumber(currentWeekSummary.qaScore)) {
    rawText.qa = "QA data pending | Current: N/A | Previous: N/A";
  }

  const heroIsPerformance = isPerformanceLikeFocus();
  const heroIsAll = state.dashboardFocus === "all";
  elements.overallCardTitle.textContent = heroIsPerformance ? "Performance Score" : "Team Overall Score";

  [
    ["overall", "overallScore"],
    ["transfer", "transferScore"],
    ["admits", "admitsScore"],
    ["aht", "ahtScore"],
    ["attendance", "attendanceScore"],
    ["qa", "qaScore"],
  ].forEach(([cardKey, metricKey]) => {
    const score = metrics[metricKey];
    const card = cardMap[cardKey];
    const effectiveScore = cardKey === "overall" && heroIsPerformance ? metrics.performanceScore : score;
    const effectiveCurrent = cardKey === "overall" && heroIsPerformance ? currentWeekSummary.performanceScore : currentWeekSummary[metricKey];
    const effectivePrevious = cardKey === "overall" && heroIsPerformance ? previousWeekSummary?.performanceScore : previousWeekSummary?.[metricKey];
    card.value.textContent = formatScore(effectiveScore);
    card.badge.textContent =
      effectiveScore === null || effectiveScore === undefined || Number.isNaN(effectiveScore) ? "N/A" : Math.round(effectiveScore);
    card.badge.className = `score-chip ${scoreClassName(effectiveScore)}`;
    card.value.closest(".summary-card")?.setAttribute("data-kpi", cardKey);
    card.raw.textContent = cardKey === "overall" && heroIsPerformance
      ? `Transfer ${buildCurrentPreviousDetail("current", currentWeekSummary.transferScore, previousWeekSummary?.transferScore, formatScore)} | Admits ${buildCurrentPreviousDetail("current", currentWeekSummary.admitsScore, previousWeekSummary?.admitsScore, formatScore)} | AHT ${buildCurrentPreviousDetail("current", currentWeekSummary.ahtScore, previousWeekSummary?.ahtScore, formatScore)}`
      : cardKey === "overall" && heroIsAll
        ? rawText.overall
      : rawText[cardKey];
    card.delta.innerHTML = buildDeltaMarkup(effectiveCurrent, effectivePrevious);
  });

  if (["transfer", "aht"].includes(state.dashboardFocus)) {
    const customComparison = getDualDateRangeComparison();
    const activeCard = cardMap[state.dashboardFocus];
    const activeRangeState = getDualDateRangeState();
    if (customComparison && activeCard && activeRangeState) {
      const currentScore = customComparison.currentMetrics?.score;
      activeCard.value.textContent = customComparison.currentMetrics?.valueText || "N/A";
      activeCard.badge.textContent =
        currentScore === null || currentScore === undefined || Number.isNaN(currentScore) ? "N/A" : Math.round(currentScore);
      activeCard.badge.className = `score-chip ${scoreClassName(currentScore)}`;
      activeCard.raw.textContent = customComparison.currentMetrics?.rawText || "No valid records in the selected range.";

      if (!customComparison.compareEnabled) {
        activeCard.delta.textContent = "";
      } else if (customComparison.delta === null) {
        activeCard.delta.innerHTML = '<span class="delta-flat">Flat</span> No compare data in the selected range';
      } else {
        const direction = customComparison.delta > 0 ? "up" : customComparison.delta < 0 ? "down" : "flat";
        const arrow = direction === "up" ? "↑" : direction === "down" ? "↓" : "→";
        const toneClass = direction === "up" ? "delta-positive" : direction === "down" ? "delta-negative" : "delta-flat";
        activeCard.delta.innerHTML = customComparison.delta > 0
          ? `<span class="delta-up">Up</span> ${Math.abs(customComparison.delta).toFixed(2)} vs compare range`
          : customComparison.delta < 0
          ? `<span class="delta-down">Down</span> ${Math.abs(customComparison.delta).toFixed(2)} vs compare range`
          : '<span class="delta-flat">Flat</span> 0.00 vs compare range';
      }
    }
  }
}

function syncFilterOptions() {
  const activeDataset = getActiveDataset();
  const focusRecords = getFocusScopedRecords(activeDataset.records);
  const realtimeFocus = isRealtimeFocus();

  const monthOptions = [...new Set(focusRecords.map((record) => record.monthLabel))];
  if (state.filters.month !== "all" && !monthOptions.includes(state.filters.month)) {
    state.filters.month = monthOptions.at(-1) || "all";
  }
  populateSelect(elements.monthFilter, monthOptions, "All Months");

  const weekOptions = realtimeFocus
    ? getAvailableWeeks(focusRecords, state.filters.month)
    : getAvailableWeeks(focusRecords, state.filters.month);
  populateWeekSelect(weekOptions);
  if (!weekOptions.includes(state.filters.week)) {
    state.filters.week = weekOptions.at(-1) || "all";
  }
  state.aggregateAllRanges = false;

  const agentOptions = getAvailableAgents(focusRecords, state.filters.month, state.filters.week);
  populateSelect(elements.agentFilter, agentOptions, "All Agents");
  if (state.filters.agent !== "all" && !agentOptions.includes(state.filters.agent)) {
    state.filters.agent = "all";
  }

  elements.monthFilter.value = state.filters.month;
  elements.weekFilter.value = state.filters.week;
  elements.agentFilter.value = state.filters.agent;
}

function updateCharts(filteredRecords, weeklyAverages, scopedRecords, trendWeeklyAverages = weeklyAverages, customConfig = null) {
  const weekSummary = pickWeekSummary(filteredRecords, weeklyAverages);
  const chartRecords =
    state.filters.week === "all"
      ? filteredRecords.filter((record) => record.weekEnding === weekSummary.weekEnding)
      : filteredRecords;

  const comparisonWeeklyAverages = getComparisonWeeklyAverages(weeklyAverages, weekSummary.weekEnding);

  if (elements.comparisonTitle && elements.comparisonSubnote) {
    const selectedWeek = comparisonWeeklyAverages.at(-1)?.weekEnding || (isRealtimeFocus() ? "Selected Date Range" : "Selected Week");
    const previousWeek = comparisonWeeklyAverages.length > 1 ? comparisonWeeklyAverages[0]?.weekEnding : isRealtimeFocus() ? "No previous date range" : `No previous ${getComparisonUnitLabel()}`;
    const qaPending = !hasValidNumber(weekSummary.qaScore);
    if (customConfig?.comparisonMode === "dual-range") {
      elements.comparisonTitle.textContent = `${FOCUS_SCORING[state.dashboardFocus].title} compare vs primary`;
      elements.comparisonSubnote.textContent = customConfig.compareLabel
        ? `${customConfig.compareLabel} versus ${customConfig.primaryLabel}.`
        : `${customConfig.primaryLabel} has no compare range selected.`;
    } else {
      elements.comparisonTitle.textContent = state.dashboardFocus === "all"
        ? `Selected date range vs previous date range`
        : `${FOCUS_SCORING[state.dashboardFocus].title} vs previous date range`;
      elements.comparisonSubnote.textContent =
        comparisonWeeklyAverages.length > 1
          ? isRealtimeFocus()
            ? `${selectedWeek} vs previous date range (${previousWeek}).${qaPending ? " QA is pending and excluded from overall where missing." : ""}`
            : `${selectedWeek} vs ${previousWeek}.${qaPending ? " QA is pending and excluded from overall where missing." : ""}`
          : isRealtimeFocus()
          ? `${selectedWeek} has no previous date range available.${qaPending ? " QA is pending and excluded from overall where missing." : ""}`
          : `${selectedWeek} has no previous ${getComparisonUnitLabel()} available.${qaPending ? " QA is pending and excluded from overall where missing." : ""}`;
    }
  }

  if (elements.rankingSubnote) {
    const qaPending = !hasValidNumber(weekSummary.qaScore);
    elements.rankingSubnote.textContent =
      state.filters.week !== "all"
        ? `${FOCUS_SCORING[state.dashboardFocus].title} spread for ${state.filters.week}. This shows where the team is clustering from top to trailing groups.${qaPending ? " QA pending is excluded from overall where missing." : ""}`
        : `Latest ${FOCUS_SCORING[state.dashboardFocus].title.toLowerCase()} spread within ${state.filters.month === "all" ? "this view" : state.filters.month}. This shows where the team is clustering from top to trailing groups.${qaPending ? " QA pending is excluded from overall where missing." : ""}`;
  }

  state.charts.trend = renderLineChart(elements.trendChart, state.charts.trend, trendWeeklyAverages, state.dashboardFocus);
  state.charts.weekScore = renderBarChart(elements.weekScoreChart, state.charts.weekScore, weekSummary, state.dashboardFocus);
  state.charts.contribution = renderContributionChart(elements.contributionChart, state.charts.contribution, weekSummary);
  state.charts.variance = customConfig?.comparisonMode === "dual-range"
    ? renderAgentDeltaVarianceChart(elements.varianceChart, state.charts.variance, customConfig.varianceByAgent, {
      metricLabel: `${FOCUS_SCORING[state.dashboardFocus].title} delta`,
      primaryLabel: customConfig.primaryLabel,
      compareLabel: customConfig.compareLabel,
      color: state.dashboardFocus === "transfer" ? "#3B82F6" : "#F59E0B",
    })
    : renderVarianceChart(elements.varianceChart, state.charts.variance, comparisonWeeklyAverages, state.dashboardFocus);
  state.charts.transferDistribution = renderDistributionChart(
    elements.transferDistributionChart,
    state.charts.transferDistribution,
    chartRecords,
    "transferScore",
    "Transfer Distribution",
    "#3B82F6"
  );
  state.charts.admitsDistribution = renderDistributionChart(
    elements.admitsDistributionChart,
    state.charts.admitsDistribution,
    chartRecords,
    "admitsScore",
    "Admits Distribution",
    "#10B981"
  );
  state.charts.ahtDistribution = renderDistributionChart(
    elements.ahtDistributionChart,
    state.charts.ahtDistribution,
    chartRecords,
    "ahtScore",
    "AHT Distribution",
    "#F59E0B"
  );
  state.charts.attendanceDistribution = renderDistributionChart(
    elements.attendanceDistributionChart,
    state.charts.attendanceDistribution,
    chartRecords,
    "attendanceScore",
    "Attendance Distribution",
    "#8B5CF6"
  );
  state.charts.qaDistribution = renderDistributionChart(
    elements.qaDistributionChart,
    state.charts.qaDistribution,
    chartRecords,
    "qaScore",
    "QA Distribution",
    "#EC4899"
  );
  state.charts.performerSpread = renderScoreSpreadChart(
    elements.performerSpreadChart,
    state.charts.performerSpread,
    chartRecords,
    {
      scoreKey: FOCUS_SCORING[state.dashboardFocus].metricKey,
      scoreLabel: FOCUS_SCORING[state.dashboardFocus].title,
    }
  );
  renderPerformerSpreadSummary(chartRecords, FOCUS_SCORING[state.dashboardFocus].metricKey);
  state.charts.stacked = renderStackedBarChart(elements.stackedChart, state.charts.stacked, chartRecords, openBreakdownDrilldown, state.dashboardFocus);
  state.charts.comparison = customConfig?.comparisonMode === "dual-range"
    ? renderRangeComparisonChart(
      elements.comparisonChart,
      state.charts.comparison,
      {
        metricLabel: FOCUS_SCORING[state.dashboardFocus].title,
        primaryLabel: customConfig.primaryLabel,
        compareLabel: customConfig.compareLabel || "Compare",
        primaryValue: weeklyAverages.at(-1)?.[FOCUS_SCORING[state.dashboardFocus].metricKey] ?? null,
        compareValue: weeklyAverages.length > 1 ? weeklyAverages[0]?.[FOCUS_SCORING[state.dashboardFocus].metricKey] ?? null : null,
        color: state.dashboardFocus === "transfer" ? "#3B82F6" : "#F59E0B",
      }
    )
    : renderComparisonChart(
      elements.comparisonChart,
      state.charts.comparison,
      comparisonWeeklyAverages,
      state.dashboardFocus
    );
}

function handleRowToggle(rowKey) {
  if (state.expandedRowKeys.has(rowKey)) {
    state.expandedRowKeys.delete(rowKey);
  } else {
    state.expandedRowKeys.add(rowKey);
  }
  updateTable(getCurrentTableRecords());
}

function updateTable(filteredRecords) {
  const tableOptions = {
    showLastUpdated: shouldShowRealtimeLastUpdated(),
    enableRowExpansion: shouldEnableRowExpansion(),
  };
  if (elements.tableDebugLine) {
    elements.tableDebugLine.hidden = true;
    elements.tableDebugLine.textContent = "";
  }
  state.expandedRowKeys = new Set(
    [...state.expandedRowKeys].filter((key) => filteredRecords.some((record) => record.key === key))
  );
  initializeTable(elements.tableHeadRow, handleSortChange, state.sort, state.dashboardFocus, tableOptions);
  renderTableSummary(elements.tableSummaryStrip, filteredRecords, state.dashboardFocus);
  renderTable(
    elements.tableBody,
    sortRecords(filteredRecords, state.sort),
    state.expandedRowKeys,
    handleRowToggle,
    state.dashboardFocus,
    tableOptions
  );
}

function updateStatus(filteredRecords) {
  if (elements.mobileLastUpdatedStrip && elements.mobileLastUpdatedText) {
    const mobileLastUpdatedDisplay = getFocusLastUpdatedDisplay(filteredRecords);
    const showMobileLastUpdated = shouldShowFocusLastUpdated() && mobileLastUpdatedDisplay !== "--";
    elements.mobileLastUpdatedText.textContent = mobileLastUpdatedDisplay;
    elements.mobileLastUpdatedStrip.hidden = !showMobileLastUpdated;
  }

  if (isRealtimeFocus()) {
    const showLastUpdated = shouldShowRealtimeLastUpdated();
    const totalAgents = new Set(filteredRecords.map((record) => record.agentName)).size;
    const totalDates = new Set(filteredRecords.map((record) => record.weekEnding)).size;
    const latestUpdatedDisplay = getLatestLastUpdatedDisplay(filteredRecords);
    elements.dataStatusText.textContent =
      state.filters.agent === "all"
        ? `${filteredRecords.length} live rows across ${totalAgents} agents and ${totalDates} date range(s).${showLastUpdated && latestUpdatedDisplay !== "--" ? ` Last updated ${latestUpdatedDisplay}.` : ""}`
        : `${filteredRecords.length} live rows for ${state.filters.agent} across ${totalDates} date range(s).${showLastUpdated && latestUpdatedDisplay !== "--" ? ` Last updated ${latestUpdatedDisplay}.` : ""}`;
    if (elements.realtimeLastUpdatedText) {
      elements.realtimeLastUpdatedText.textContent = showLastUpdated ? latestUpdatedDisplay : "--";
    }
    if (elements.realtimeLastUpdatedPill) {
      elements.realtimeLastUpdatedPill.hidden = !showLastUpdated;
    }
    return;
  }

  if (state.dashboardFocus === "transfer") {
    const focusLabel = state.dashboardFocus === "transfer" ? "transfer" : "AHT";
    const rowLabel = state.dashboardFocus === "transfer" ? "transfer rows" : "AHT rows";
    const showLastUpdated = shouldShowRealtimeLastUpdated();
    const totalAgents = new Set(filteredRecords.map((record) => record.agentName)).size;
    const totalDates = new Set(filteredRecords.map((record) => record.weekEnding)).size;
    const latestUpdatedDisplay = getLatestLastUpdatedDisplay(filteredRecords);
    elements.dataStatusText.textContent =
      state.filters.agent === "all"
        ? `${filteredRecords.length} ${rowLabel} across ${totalAgents} agents and ${totalDates} date range(s).${showLastUpdated && latestUpdatedDisplay !== "--" ? ` Last updated ${latestUpdatedDisplay}.` : ""}`
        : `${filteredRecords.length} ${focusLabel === "transfer" ? "transfer rows" : "AHT rows"} for ${state.filters.agent} across ${totalDates} date range(s).${showLastUpdated && latestUpdatedDisplay !== "--" ? ` Last updated ${latestUpdatedDisplay}.` : ""}`;
    if (elements.realtimeLastUpdatedText) {
      elements.realtimeLastUpdatedText.textContent = showLastUpdated ? latestUpdatedDisplay : "--";
    }
    if (elements.realtimeLastUpdatedPill) {
      elements.realtimeLastUpdatedPill.hidden = !showLastUpdated;
    }
    return;
  }

  if (state.dashboardFocus === "aht") {
    const showLastUpdated = shouldShowRealtimeLastUpdated();
    const totalAgents = new Set(filteredRecords.map((record) => record.agentName)).size;
    const totalDates = new Set(filteredRecords.map((record) => record.weekEnding)).size;
    const latestUpdatedDisplay = getLatestLastUpdatedDisplay(filteredRecords);
    elements.dataStatusText.textContent =
      state.filters.agent === "all"
        ? `${filteredRecords.length} AHT rows across ${totalAgents} agents and ${totalDates} date range(s).${showLastUpdated && latestUpdatedDisplay !== "--" ? ` Last updated ${latestUpdatedDisplay}.` : ""}`
        : `${filteredRecords.length} AHT rows for ${state.filters.agent} across ${totalDates} date range(s).${showLastUpdated && latestUpdatedDisplay !== "--" ? ` Last updated ${latestUpdatedDisplay}.` : ""}`;
    if (elements.realtimeLastUpdatedText) {
      elements.realtimeLastUpdatedText.textContent = showLastUpdated ? latestUpdatedDisplay : "--";
    }
    if (elements.realtimeLastUpdatedPill) {
      elements.realtimeLastUpdatedPill.hidden = !showLastUpdated;
    }
    return;
  }

  if (state.dashboardFocus === "admits") {
    const totalAgents = new Set(filteredRecords.map((record) => record.agentName)).size;
    const totalDates = new Set(filteredRecords.map((record) => record.weekEnding)).size;
    const latestUpdatedDisplay = getLatestAdmitsLastUpdatedDisplay(filteredRecords);
    elements.dataStatusText.textContent =
      state.filters.agent === "all"
        ? `${filteredRecords.length} admits rows across ${totalAgents} agents and ${totalDates} date range(s).${latestUpdatedDisplay !== "--" ? ` Last updated ${latestUpdatedDisplay}.` : ""}`
        : `${filteredRecords.length} admits rows for ${state.filters.agent} across ${totalDates} date range(s).${latestUpdatedDisplay !== "--" ? ` Last updated ${latestUpdatedDisplay}.` : ""}`;
    if (elements.realtimeLastUpdatedText) {
      elements.realtimeLastUpdatedText.textContent = latestUpdatedDisplay;
    }
    if (elements.realtimeLastUpdatedPill) {
      elements.realtimeLastUpdatedPill.hidden = latestUpdatedDisplay === "--";
    }
    return;
  }

  if (state.dashboardFocus === "attendance") {
    const totalAgents = new Set(filteredRecords.map((record) => record.agentName)).size;
    const totalDates = new Set(filteredRecords.map((record) => record.weekEnding)).size;
    const latestUpdatedDisplay = getFocusLastUpdatedDisplay(filteredRecords);
    elements.dataStatusText.textContent =
      state.filters.agent === "all"
        ? `${filteredRecords.length} attendance rows across ${totalAgents} agents and ${totalDates} date range(s).${latestUpdatedDisplay !== "--" ? ` Last updated ${latestUpdatedDisplay}.` : ""}`
        : `${filteredRecords.length} attendance rows for ${state.filters.agent} across ${totalDates} date range(s).${latestUpdatedDisplay !== "--" ? ` Last updated ${latestUpdatedDisplay}.` : ""}`;
    if (elements.realtimeLastUpdatedText) {
      elements.realtimeLastUpdatedText.textContent = latestUpdatedDisplay;
    }
    if (elements.realtimeLastUpdatedPill) {
      elements.realtimeLastUpdatedPill.hidden = latestUpdatedDisplay === "--";
    }
    return;
  }

  if (state.dashboardFocus === "qa") {
    const totalAgents = new Set(filteredRecords.map((record) => record.agentName)).size;
    const totalDates = new Set(filteredRecords.map((record) => record.weekEnding)).size;
    const latestUpdatedDisplay = getLatestQaLastUpdatedDisplay(filteredRecords);
    elements.dataStatusText.textContent =
      state.filters.agent === "all"
        ? `${filteredRecords.length} QA rows across ${totalAgents} agents and ${totalDates} date range(s).${latestUpdatedDisplay !== "--" ? ` Last updated ${latestUpdatedDisplay}.` : ""}`
        : `${filteredRecords.length} QA rows for ${state.filters.agent} across ${totalDates} date range(s).${latestUpdatedDisplay !== "--" ? ` Last updated ${latestUpdatedDisplay}.` : ""}`;
    if (elements.realtimeLastUpdatedText) {
      elements.realtimeLastUpdatedText.textContent = latestUpdatedDisplay;
    }
    if (elements.realtimeLastUpdatedPill) {
      elements.realtimeLastUpdatedPill.hidden = latestUpdatedDisplay === "--";
    }
    return;
  }

  const totalAgents = new Set(filteredRecords.map((record) => record.agentName)).size;
  const totalWeeks = new Set(filteredRecords.map((record) => record.weekEnding)).size;
  const qaAvailable = filteredRecords.some((record) => hasValidNumber(record.qaScore));
  elements.dataStatusText.textContent =
    state.filters.agent === "all"
      ? `${filteredRecords.length} score rows across ${totalAgents} agents and ${totalWeeks} date ranges.${qaAvailable ? "" : " QA data is pending, so overall rankings are using the available KPIs only."}`
      : `${filteredRecords.length} score rows for ${state.filters.agent} across ${totalWeeks} date range(s).${qaAvailable ? "" : " QA data is pending, so overall scores are using the available KPIs only."}`;
}

function updateLayoutVisibility() {
  const isAgentView = state.filters.agent !== "all";
  const singleKpiFocus = isSingleKpiFocus();
  elements.teamOnlySections.forEach((section) => {
    section.classList.toggle("is-hidden", isAgentView);
  });
  elements.agentOnlySections.forEach((section) => {
    section.classList.toggle("is-hidden", !isAgentView);
  });

  const focus = state.dashboardFocus;
  const showCompactSingleKpiCharts = ["transfer", "admits", "aht", "attendance", "qa"].includes(focus);

  if (elements.distributionSection && elements.snapshotSection && elements.diagnosticsSection) {
    const distributionHost = showCompactSingleKpiCharts ? elements.snapshotSection : elements.diagnosticsSection;
    if (elements.distributionSection.parentElement !== distributionHost) {
      distributionHost.appendChild(elements.distributionSection);
    }
  }
  if (elements.agentsSection) {
    elements.agentsSection.classList.toggle("is-hidden", showCompactSingleKpiCharts);
  }
  if (elements.varianceCard) {
    elements.varianceCard.classList.toggle("is-hidden", showCompactSingleKpiCharts);
  }

  elements.summaryCards.forEach((card) => {
    const group = card.dataset.focusGroup;
    const kpi = card.dataset.kpi;
    const visible =
      !group
        ? focus === "all" || focus === "performance" || focus === "realtime"
        : focus === "all"
        ? true
        : (focus === "performance" || focus === "realtime")
        ? group === "performance"
        : isSingleKpiFocus()
          ? kpi === focus
          : kpi === "overall" || group === focus;
    card.classList.toggle("is-hidden", !visible);
  });

  elements.focusGroupedNodes.forEach((node) => {
    const visible = focus === "all"
      ? true
      : (focus === "performance" || focus === "realtime")
      ? node.dataset.focusGroup === "performance"
      : ["transfer", "admits", "aht"].includes(focus)
        ? node.dataset.distributionItem === focus
      : node.dataset.focusGroup === focus;
    node.classList.toggle("is-hidden", !visible);
  });

  elements.contributionCard?.classList.toggle("is-hidden", PERFORMANCE_FOCUS_KEYS.includes(focus) || singleKpiFocus);

  if (elements.trendsSection) {
    elements.trendsSection.classList.toggle("is-hidden", focus === "attendance" || focus === "qa");
  }
  if (elements.snapshotSection) {
    elements.snapshotSection.classList.toggle("is-hidden", focus === "qa");
  }
  if (elements.performersSection) {
    elements.performersSection.classList.toggle("is-hidden", focus === "qa" || showCompactSingleKpiCharts);
  }
  if (elements.deepDiveSection) {
    elements.deepDiveSection.classList.toggle("is-hidden", focus === "attendance" || focus === "qa");
  }
  if (elements.diagnosticsSection) {
    elements.diagnosticsSection.classList.toggle("is-hidden", focus === "qa" || !["all", "performance", "realtime"].includes(focus));
  }
}

function updateDashboard() {
  if (isDualRangeFocus()) {
    const dualData = getDualRangeDashboardData();
    if (!dualData) {
      updateTable([]);
      updateStatus([]);
      updateLayoutVisibility();
      return;
    }

    state.distributionDrilldownOpen = false;
    applyDistributionDrilldownState();
    updateInsights(dualData.dashboardRecords, dualData.weeklyAverages);
    updateSummaryCards(dualData.dashboardRecords, dualData.weeklyAverages);
    updateAgentFocus(dualData.dashboardRecords, dualData.weeklyAverages);
    updateCharts(
      dualData.dashboardRecords,
      dualData.weeklyAverages,
      dualData.dashboardRecords,
      dualData.trendWeeklyAverages,
      {
        comparisonMode: "dual-range",
        primaryLabel: dualData.primaryLabel,
        compareLabel: dualData.compareLabel,
        varianceByAgent: dualData.varianceByAgent,
      }
    );
    updateTable(dualData.tableRecords);
    updateStatus(dualData.dashboardRecords);
    updateLayoutVisibility();
    return;
  }

  const activeDataset = getActiveDataset();
  const focusRecords = getFocusScopedRecords(activeDataset.records);
  const dashboardRecords = getFilteredRecords(focusRecords, { ...state.filters, search: "" });
  const tableRecords = getCurrentTableRecords();
  const trendRecords = getFilteredRecords(focusRecords, {
    ...state.filters,
    month: usesRealtimeSource() ? "all" : state.filters.month,
    week: "all",
    search: "",
  });
  const weeklyAverages = getScopedWeeklyAverages(trendRecords);
  const selectedSummary = pickWeekSummary(dashboardRecords, weeklyAverages);
  const trendWeeklyAverages = getTrendWeeklyAverages(weeklyAverages, selectedSummary.weekEnding);

  state.distributionDrilldownOpen = false;
  applyDistributionDrilldownState();

  if (state.dashboardFocus === "qa") {
    updateSummaryCards(dashboardRecords, weeklyAverages);
    updateTable(tableRecords);
    updateStatus(dashboardRecords);
    updateLayoutVisibility();
    return;
  }

  updateInsights(dashboardRecords, weeklyAverages);
  updateSummaryCards(dashboardRecords, weeklyAverages);
  updateAgentFocus(dashboardRecords, weeklyAverages);
  updateCharts(dashboardRecords, weeklyAverages, trendRecords, trendWeeklyAverages);
  updateTable(tableRecords);
  updateStatus(dashboardRecords);
  updateLayoutVisibility();
}

function applyDashboardFocus() {
  const focusLabel = FOCUS_SCORING[state.dashboardFocus].title;
  const realtimeFocus = isRealtimeFocus();
  const performanceLike = isPerformanceLikeFocus();
  if (usesRealtimeSource()) {
    state.filters.month = state.realtimeDataset?.monthOptions?.at(-1) || "all";
    state.filters.week = getLatestRealtimeWeek();
  }
  syncFilterOptions();
  if (elements.appShell) {
    elements.appShell.dataset.dashboardFocus = state.dashboardFocus;
  }
  elements.focusButtons.forEach((button) => {
    const active = button.dataset.dashboardFocus === state.dashboardFocus;
    button.classList.toggle("is-active", active);
    button.setAttribute("aria-selected", String(active));
  });
  elements.summaryCardsList.forEach((card) => {
    card.classList.remove("is-mobile-expanded");
    card.setAttribute("aria-expanded", "false");
  });

  if (elements.monthFilterField && elements.weekFilterField && elements.monthFilterLabel && elements.weekFilterLabel) {
    elements.weekFilterLabel.textContent = "Date Range";
    elements.monthFilterLabel.textContent = "Month";
  }
  if (elements.topFiltersTitle) {
    elements.topFiltersTitle.textContent = isDualRangeFocus()
      ? ""
      : "Filter by month, date range, and agent";
    elements.topFiltersTitle.classList.toggle("is-hidden", isDualRangeFocus());
  }
  if (elements.realtimeLastUpdatedPill && elements.realtimeLastUpdatedText) {
    elements.realtimeLastUpdatedPill.hidden = !shouldShowFocusLastUpdated();
    const activeDataset = getActiveDataset();
    elements.realtimeLastUpdatedText.textContent = state.dashboardFocus === "admits"
      ? getLatestAdmitsLastUpdatedDisplay(getFilteredRecords(activeDataset.records, { ...state.filters, search: "" }))
      : state.dashboardFocus === "qa"
      ? getLatestQaLastUpdatedDisplay(getFilteredRecords(activeDataset.records, { ...state.filters, search: "" }))
      : usesRealtimeSource()
      ? shouldShowRealtimeLastUpdated()
        ? activeDataset?.lastUpdatedDisplay || "N/A"
        : "--"
      : "--";
  }

  if (elements.trendSectionTitle) {
    elements.trendSectionTitle.textContent = state.dashboardFocus === "all"
      ? "KPI trends"
      : realtimeFocus
      ? "Real-time performance trends"
      : state.dashboardFocus === "performance"
      ? "Performance trends"
      : `${focusLabel} trends`;
  }
  if (elements.snapshotSectionTitle) {
    elements.snapshotSectionTitle.textContent = realtimeFocus ? "Real-time performance snapshot" : state.dashboardFocus === "all" ? "KPI score snapshot" : `${focusLabel} score snapshot`;
  }
  if (elements.varianceSectionTitle) {
    elements.varianceSectionTitle.textContent = realtimeFocus ? "Previous date variance" : state.dashboardFocus === "all" ? "KPI variance" : `${focusLabel} variance`;
  }
  if (elements.rankingSectionTitle) {
    elements.rankingSectionTitle.textContent = realtimeFocus ? "Real-time performance spread" : state.dashboardFocus === "all" ? "Overall score spread" : `${focusLabel} score spread`;
  }
  if (elements.bottomSectionTitle) {
    elements.bottomSectionTitle.textContent = state.dashboardFocus === "all" ? "Bottom 5 agents" : `Bottom 5 ${focusLabel.toLowerCase()} agents`;
  }
  if (elements.improvedSectionTitle) {
    elements.improvedSectionTitle.textContent = state.dashboardFocus === "all" ? "Most improved agents" : `Most improved ${focusLabel.toLowerCase()} agents`;
  }
  if (elements.breakdownSectionTitle) {
    elements.breakdownSectionTitle.textContent = state.dashboardFocus === "all"
      ? "KPI breakdown per agent"
      : state.dashboardFocus === "performance" || realtimeFocus
      ? "Performance KPI breakdown per agent"
      : `${focusLabel} breakdown per agent`;
  }
  if (elements.distributionSectionTitle) {
    elements.distributionSectionTitle.textContent = state.dashboardFocus === "all"
      ? "Agent score distribution by KPI"
      : state.dashboardFocus === "performance" || realtimeFocus
      ? "Performance KPI score distribution"
      : `${focusLabel} score distribution`;
  }
  if (elements.tableSectionTitle) {
    elements.tableSectionTitle.textContent = state.dashboardFocus === "all"
      ? "All agents KPI score table"
      : realtimeFocus
      ? "Real-time performance table"
      : state.dashboardFocus === "performance"
      ? "Performance KPI score table"
      : `${focusLabel} score table`;
  }
  if (elements.tableSectionSubnote) {
    elements.tableSectionSubnote.textContent = state.dashboardFocus === "all"
      ? "Balanced date-range view across performance, attendance, and quality assurance."
      : realtimeFocus
      ? "Live CTM performance rows by date range, compared against the previous available range."
      : state.dashboardFocus === "performance"
      ? "Focused on transfer, admits, and AHT rows for the selected scope."
      : ["transfer", "admits", "aht"].includes(state.dashboardFocus)
      ? `Focused on ${focusLabel.toLowerCase()} rows for the selected scope.`
      : state.dashboardFocus === "attendance"
        ? "Focused on attendance rows and attendance score movement for the selected scope."
        : "Focused on quality assurance rows, QA scoring, and source updates for the selected scope.";
  }
  if (elements.tableRowHint) {
    elements.tableRowHint.textContent = !shouldEnableRowExpansion()
      ? ""
      : state.dashboardFocus === "performance" || realtimeFocus
      ? "Click any row to view grouped Transfer, Admits, and AHT details."
      : ["transfer", "admits", "aht"].includes(state.dashboardFocus)
      ? `Click any row to open the ${focusLabel.toLowerCase()} debugging details.`
      : state.dashboardFocus === "attendance"
        ? "Click any row to open the attendance details."
        : "Click any row to open the quality assurance details.";
    elements.tableRowHint.classList.toggle("is-hidden", !shouldEnableRowExpansion());
  }
  if (elements.expandAllRowsButton && elements.collapseAllRowsButton) {
    const expandable = shouldEnableRowExpansion();
    elements.expandAllRowsButton.classList.toggle("is-hidden", !expandable);
    elements.collapseAllRowsButton.classList.toggle("is-hidden", !expandable);
  }
  if (elements.tableSearch) {
    elements.tableSearch.placeholder = state.dashboardFocus === "all"
      ? "Search agent, date range, or KPI row"
      : realtimeFocus
      ? "Search agent, date range, or real-time row"
      : performanceLike
      ? "Search agent, date range, or performance row"
      : state.dashboardFocus === "attendance"
        ? "Search agent, date range, or attendance row"
        : "Search agent, date range, or quality row";
  }

  state.mobileDistributionPanel = FOCUS_SCORING[state.dashboardFocus].distribution;
  applyMobileDistributionState();
  syncDualDateRangeControls();
  updateDashboard();
}

function handleSortChange(key) {
  if (state.sort.key === key) {
    state.sort.direction = state.sort.direction === "asc" ? "desc" : "asc";
  } else {
    state.sort.key = key;
    state.sort.direction = key === "agentName" || key === "weekEnding" ? "asc" : "desc";
  }
  updateTable(getCurrentTableRecords());
}

function bindNavigation() {
  elements.navButtons.forEach((button) => {
    button.addEventListener("click", () => {
      elements.navButtons.forEach((item) => item.classList.remove("is-active"));
      button.classList.add("is-active");
      document.getElementById(button.dataset.target)?.scrollIntoView({ behavior: "smooth", block: "start" });
      state.mobileSidebarOpen = false;
      applyMobileUiState();
    });
  });
}

function applyWeekFilterSelection(selectElement) {
  state.filters.week = selectElement.value;
  state.aggregateAllRanges = false;
  state.forceAllWeekSelection = false;
  syncFilterOptions();
  requestAnimationFrame(() => {
    updateDashboard();
  });
}

function queueWeekFilterSync() {
  requestAnimationFrame(() => {
    if (elements.weekFilter) {
      applyWeekFilterSelection(elements.weekFilter);
    }
  });
}

function bindViewportEvents() {
  window.addEventListener("resize", () => {
    if (state.resizeRaf) {
      cancelAnimationFrame(state.resizeRaf);
    }
    state.resizeRaf = requestAnimationFrame(() => {
      state.resizeRaf = null;
      if (state.dataset) {
        if (window.innerWidth > 720) {
          state.mobileTableExpanded = true;
          state.mobileFiltersExpanded = true;
          state.mobileSidebarOpen = false;
          state.floatingLegendOpen = false;
          state.mobileMoreOpen = false;
        }
        applyMobileUiState();
        applyMobileMoreState();
        applyMobileDistributionState();
        applyFloatingLegendState();
        updateDashboard();
      }
    });
  });
}

function bindEvents() {
  elements.mobileMenuToggle?.addEventListener("click", () => {
    state.mobileSidebarOpen = !state.mobileSidebarOpen;
    applyMobileUiState();
  });

  elements.mobileHomeButtons.forEach((button) => {
    button.addEventListener("click", () => {
      let targetId = button.dataset.mobileTarget;
      if (state.filters.agent !== "all" && (targetId === "summaryCards" || targetId === "agentsSection")) {
        targetId = "agentFocusSection";
      }
      if (targetId === "tableSection") {
        state.mobileTableExpanded = true;
        applyMobileUiState();
        setMobileDockActive("table");
      } else if (targetId === "summaryCards" || targetId === "agentFocusSection") {
        setMobileDockActive("summary");
      } else if (targetId === "agentsSection") {
        setMobileDockActive("agents");
      } else {
        setMobileDockActive("home");
      }
      document.getElementById(targetId)?.scrollIntoView({ behavior: "smooth", block: "start" });
    });
  });

  elements.mobileLegendLauncher?.addEventListener("click", () => {
    state.mobileMoreOpen = false;
    applyMobileMoreState();
    state.floatingLegendOpen = true;
    applyFloatingLegendState();
    setMobileDockActive("more");
  });

  elements.mobileFiltersLauncher?.addEventListener("click", () => {
    state.mobileMoreOpen = false;
    applyMobileMoreState();
    state.mobileFiltersExpanded = true;
    applyMobileUiState();
    setMobileDockActive("more");
    document.getElementById("filtersSection")?.scrollIntoView({ behavior: "smooth", block: "start" });
  });

  elements.mobileBottomNavButtons.forEach((button) => {
    button.addEventListener("click", () => {
      triggerMobileAction(button.dataset.mobileAction);
    });
  });

  elements.mobileDistributionButtons.forEach((button) => {
    button.addEventListener("click", () => {
      state.mobileDistributionPanel = button.dataset.distributionPanel;
      applyMobileDistributionState();
    });
  });

  elements.floatingLegendToggle?.addEventListener("click", () => {
    state.floatingLegendOpen = !state.floatingLegendOpen;
    applyFloatingLegendState();
    if (state.floatingLegendOpen) {
      setMobileDockActive("more");
    }
  });

  elements.floatingLegendClose?.addEventListener("click", () => {
    state.floatingLegendOpen = false;
    applyFloatingLegendState();
    setMobileDockActive("home");
  });

  elements.mobileMoreClose?.addEventListener("click", () => {
    state.mobileMoreOpen = false;
    applyMobileMoreState();
    setMobileDockActive("home");
  });

  elements.mobileMoreLegend?.addEventListener("click", () => {
    state.mobileMoreOpen = false;
    applyMobileMoreState();
    state.floatingLegendOpen = true;
    applyFloatingLegendState();
    setMobileDockActive("more");
  });

  elements.mobileMoreFilters?.addEventListener("click", () => {
    state.mobileMoreOpen = false;
    applyMobileMoreState();
    state.mobileFiltersExpanded = true;
    applyMobileUiState();
    document.getElementById("filtersSection")?.scrollIntoView({ behavior: "smooth", block: "start" });
    setMobileDockActive("more");
  });

  elements.mobileMoreExport?.addEventListener("click", () => {
    const filteredRecords = sortRecords(getCurrentTableRecords(), state.sort);
    exportRowsToCsv(filteredRecords, state.dashboardFocus, { showLastUpdated: shouldShowRealtimeLastUpdated() });
    state.mobileMoreOpen = false;
    applyMobileMoreState();
    setMobileDockActive("more");
  });

  elements.mobileMoreReset?.addEventListener("click", () => {
    resetDashboardFilters();
    state.mobileMoreOpen = false;
    applyMobileMoreState();
    setMobileDockActive("home");
  });

  elements.distributionDrilldownClose?.addEventListener("click", () => {
    state.distributionDrilldownOpen = false;
    state.distributionDrilldownExpanded = false;
    state.distributionDrilldownDetail = null;
    applyDistributionDrilldownState();
  });

  elements.distributionDrilldownList?.addEventListener("click", (event) => {
    const trigger = event.target.closest(".distribution-drilldown-more");
    if (!trigger) return;
    state.distributionDrilldownExpanded = true;
    renderDistributionDrilldownContent();
  });

  elements.tableCollapseToggle?.addEventListener("click", () => {
    state.mobileTableExpanded = !state.mobileTableExpanded;
    applyMobileUiState();
  });

  elements.filtersCollapseToggle?.addEventListener("click", () => {
    state.mobileFiltersExpanded = !state.mobileFiltersExpanded;
    applyMobileUiState();
  });

  elements.mobileTableToggle?.addEventListener("click", () => {
    state.mobileTableExpanded = !state.mobileTableExpanded;
    applyMobileUiState();
    if (state.mobileTableExpanded) {
      setMobileDockActive("table");
    }
    document.getElementById("tableSection")?.scrollIntoView({ behavior: "smooth", block: "start" });
  });

  elements.monthFilter.addEventListener("change", (event) => {
    state.filters.month = event.target.value;
    state.forceAllWeekSelection = false;
    syncFilterOptions();
    updateDashboard();
  });

  elements.weekFilter.addEventListener("input", (event) => {
    applyWeekFilterSelection(event.target);
  });

  elements.weekFilter.addEventListener("change", (event) => {
    applyWeekFilterSelection(event.target);
  });

  elements.agentFilter.addEventListener("change", (event) => {
    state.filters.agent = event.target.value;
    if (["transfer", "aht"].includes(state.dashboardFocus)) {
      const dualState = getDualDateRangeState();
      if (dualState) {
        dualState.initialized = false;
      }
      syncDualDateRangeControls();
    }
    updateDashboard();
  });

  elements.dualRangeAgentFilter?.addEventListener("change", (event) => {
    state.filters.agent = event.target.value;
    const dualState = getDualDateRangeState();
    if (dualState) {
      dualState.initialized = false;
      ensureDualDateRangeDefaults(true);
    }
    syncDualDateRangeControls();
    updateDashboard();
  });

  const handleDualDateRangeChange = () => {
    const dualState = getDualDateRangeState();
    if (!dualState) return;
    dualState.primaryStart = elements.primaryDateStart?.value || "";
    dualState.primaryEnd = elements.primaryDateEnd?.value || "";
    dualState.compareEnabled = Boolean(elements.compareDateToggle?.checked);
    dualState.compareStart = elements.compareDateStart?.value || "";
    dualState.compareEnd = elements.compareDateEnd?.value || "";
    syncDualDateRangeControls();
    updateDashboard();
  };

  elements.primaryDateStart?.addEventListener("change", handleDualDateRangeChange);
  elements.primaryDateEnd?.addEventListener("change", handleDualDateRangeChange);
  elements.compareDateToggle?.addEventListener("change", handleDualDateRangeChange);
  elements.compareDateStart?.addEventListener("change", handleDualDateRangeChange);
  elements.compareDateEnd?.addEventListener("change", handleDualDateRangeChange);
  elements.dualRangeReset?.addEventListener("click", () => {
    resetDashboardFilters();
  });

  elements.tableSearch.addEventListener("input", (event) => {
    state.filters.search = event.target.value;
    updateDashboard();
  });

  elements.expandAllRowsButton?.addEventListener("click", () => {
    const filteredRecords = sortRecords(getCurrentTableRecords(), state.sort);
    state.expandedRowKeys = new Set(filteredRecords.map((record) => record.key));
    updateTable(filteredRecords);
  });

  elements.collapseAllRowsButton?.addEventListener("click", () => {
    state.expandedRowKeys = new Set();
    updateTable(getCurrentTableRecords());
  });

  elements.resetFilters.addEventListener("click", () => {
    resetDashboardFilters();
  });

  elements.focusButtons.forEach((button) => {
    button.addEventListener("click", () => {
      state.dashboardFocus = button.dataset.dashboardFocus || "all";
      state.sort = {
        key: FOCUS_SCORING[state.dashboardFocus].metricKey,
        direction: "desc",
      };
      applyDashboardFocus();
    });
  });

  elements.exportCsvButton.addEventListener("click", () => {
    const filteredRecords = sortRecords(getCurrentTableRecords(), state.sort);
    exportRowsToCsv(filteredRecords, state.dashboardFocus, { showLastUpdated: shouldShowRealtimeLastUpdated() });
  });
}

async function init() {
  if (state.initialized) {
    return;
  }

  state.initialized = true;

  try {
    elements.dataStatusText.textContent = "Loading Google Sheets data and calculating KPI scores.";
    const rawDatasets = await loadAllDatasets();
    state.dataset = buildKpiDataset(rawDatasets);
    state.realtimeDataset = buildRealtimeDataset(rawDatasets);
    state.filters.month = state.dataset.monthOptions.at(-1) || "all";
    state.filters.week = state.dataset.weekOptions.at(-1) || "all";

    initializeMobileDefaults();
    syncFilterOptions();
    initializeFloatingLegend();
    bindNavigation();
    bindSummaryCardInteractions();
    bindEvents();
    bindViewportEvents();
    bindMobileScrollSpy();
    applyMobileUiState();
    applyMobileMoreState();
    applyMobileDistributionState();
    applyFloatingLegendState();
    applyDistributionDrilldownState();
    applyDashboardFocus();
  } catch (error) {
    state.initialized = false;
    console.error(error);
    elements.dataStatusText.textContent = "The dashboard could not load the CSV sources. Check the published sheet permissions and try again.";
    elements.tableBody.innerHTML = '<tr><td colspan="14"><div class="status-message">Unable to load dashboard data.</div></td></tr>';
  }
}

const authRequired = document.body?.dataset.requireAuth === "true";

if (!authRequired) {
  init();
} else {
  if (window.__flylandAuthState?.authorized) {
    init();
  }

  window.addEventListener("flyland:auth-granted", () => {
    init();
  });
}
