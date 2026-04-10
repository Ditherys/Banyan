const COLUMN_SETS = {
  all: [
    { key: "agentName", label: "Agent Name", type: "text" },
    { key: "weekEnding", label: "Date Range", type: "text" },
    { key: "transfer", label: "Transfer", type: "kpi" },
    { key: "admits", label: "Admits", type: "kpi" },
    { key: "aht", label: "AHT", type: "kpi" },
    { key: "attendance", label: "Attendance", type: "kpi" },
    { key: "qa", label: "Quality Assurance", type: "kpi" },
    { key: "overallScore", label: "Overall", type: "number2" },
  ],
  performance: [
    { key: "agentName", label: "Agent Name", type: "text" },
    { key: "weekEnding", label: "Date Range", type: "text" },
    { key: "transfer", label: "Transfer", type: "kpi" },
    { key: "admits", label: "Admits", type: "kpi" },
    { key: "aht", label: "AHT", type: "kpi" },
    { key: "performanceScore", label: "Performance", type: "number2" },
  ],
  transfer: [
    { key: "agentName", label: "Agent Name", type: "text" },
    { key: "weekEnding", label: "Date Range", type: "text" },
    { key: "firstTimeCaller", label: "First Time Caller", type: "number0" },
    { key: "transferCount", label: "Transfer Count", type: "number0" },
    { key: "transfer", label: "Transfer", type: "kpi" },
    { key: "lastUpdatedDisplay", label: "Last Updated", type: "text" },
  ],
  admits: [
    { key: "agentName", label: "Agent Name", type: "text" },
    { key: "weekEnding", label: "Date Range", type: "text" },
    { key: "admitsCount", label: "Admits Count", type: "number0" },
    { key: "admitsScore", label: "Admits Score", type: "scorepill" },
    { key: "admitsLastUpdatedDisplay", label: "Last Updated", type: "text" },
  ],
  aht: [
    { key: "agentName", label: "Agent Name", type: "text" },
    { key: "weekEnding", label: "Date Range", type: "text" },
    { key: "inboundCalls", label: "Inbound Calls", type: "number0" },
    { key: "inboundMinutes", label: "Inbound Minutes", type: "text" },
    { key: "holdTime", label: "Hold Time", type: "text" },
    { key: "aht", label: "AHT", type: "kpi" },
    { key: "lastUpdatedDisplay", label: "Last Updated", type: "text" },
  ],
  realtime: [
    { key: "agentName", label: "Agent Name", type: "text" },
    { key: "weekEnding", label: "Date Range", type: "text" },
    { key: "transfer", label: "Transfer", type: "kpi" },
    { key: "admits", label: "Admits", type: "kpi" },
    { key: "aht", label: "AHT", type: "kpi" },
    { key: "performanceScore", label: "Performance", type: "number2" },
    { key: "lastUpdatedDisplay", label: "Last Updated", type: "text" },
  ],
  attendance: [
    { key: "agentName", label: "Agent Name", type: "text" },
    { key: "weekEnding", label: "Date Range", type: "text" },
    { key: "hoursPresent", label: "Hrs Present", type: "number2" },
    { key: "hoursAbsent", label: "Hrs Absent", type: "number2" },
    { key: "sickLeaveHours", label: "SL Hrs", type: "number2" },
    { key: "undertimeHours", label: "UT", type: "number2" },
    { key: "requiredHours", label: "Req Hrs", type: "number2" },
    { key: "attendancePercentDisplay", label: "Att %", type: "text" },
    { key: "attendanceScore", label: "KPI Score", type: "scorepill" },
    { key: "attendanceLastUpdatedDisplay", label: "Last Updated", type: "text" },
  ],
  qa: [
    { key: "agentName", label: "Agent Name", type: "text" },
    { key: "weekEnding", label: "Date Range", type: "text" },
    { key: "qaPercentDisplay", label: "Quality Assurance", type: "text" },
    { key: "qaScore", label: "QA Score", type: "scorepill" },
    { key: "qaLastUpdatedDisplay", label: "Last Updated", type: "text" },
  ],
};

function getColumns(focus = "performance", options = {}) {
  const columns = [...(COLUMN_SETS[focus] || COLUMN_SETS.all)];
  if ((focus === "realtime" || focus === "transfer" || focus === "aht") && options.showLastUpdated === false) {
    return columns.filter((column) => column.key !== "lastUpdatedDisplay");
  }
  return columns;
}

function getSummaryMetricConfig(focus = "performance") {
  if (focus === "all") {
    return {
      key: "overallScore",
      label: "Overall",
      description: "Mean overall score across shown rows.",
    };
  }
  if (focus === "attendance") {
    return {
      key: "attendanceScore",
      label: "Attendance",
      description: "Mean attendance score across shown rows.",
    };
  }
  if (focus === "transfer") {
    return {
      key: "transferScore",
      label: "Transfer",
      description: "Mean transfer score across shown rows.",
    };
  }
  if (focus === "admits") {
    return {
      key: "admitsScore",
      label: "Admits",
      description: "Mean admits score across shown rows.",
    };
  }
  if (focus === "aht") {
    return {
      key: "ahtScore",
      label: "AHT",
      description: "Mean AHT score across shown rows.",
    };
  }
  if (focus === "qa") {
    return {
      key: "qaScore",
      label: "Quality Assurance",
      description: "Mean QA score across shown rows.",
    };
  }
  if (focus === "realtime") {
    return {
      key: "performanceScore",
      label: "Performance",
      description: "Mean real-time performance score across shown rows.",
    };
  }
  return {
    key: "performanceScore",
    label: "Performance",
    description: "Mean performance score across shown rows.",
  };
}

function getColumnSortLabel(column, sortState) {
  if (sortState.key !== column.key) return column.label;
  return `${column.label} ${sortState.direction === "asc" ? "^" : "v"}`;
}

function scoreText(value) {
  return value === null || value === undefined || Number.isNaN(value) ? "N/A" : Number(value).toFixed(2);
}

function getRowState(record, focus = "performance") {
  const focusKey = focus === "all"
    ? "overallScore"
    : focus === "transfer"
      ? "transferScore"
      : focus === "admits"
        ? "admitsScore"
        : focus === "aht"
          ? "ahtScore"
          : focus === "attendance"
            ? "attendanceScore"
            : focus === "qa"
              ? "qaScore"
              : "performanceScore";
  const focusValue = record[focusKey];
  const candidateScores = focus === "all"
    ? [record.transferScore, record.admitsScore, record.ahtScore, record.attendanceScore, record.qaScore]
    : focus === "performance" || focus === "realtime"
      ? [record.transferScore, record.admitsScore, record.ahtScore]
      : focus === "transfer"
        ? [record.transferScore]
        : focus === "admits"
          ? [record.admitsScore]
          : focus === "aht"
            ? [record.ahtScore]
            : focus === "attendance"
              ? [record.attendanceScore]
              : [record.qaScore];
  const weakestScore = candidateScores
    .filter((value) => typeof value === "number" && !Number.isNaN(value))
    .reduce((lowest, value) => (lowest === null ? value : Math.min(lowest, value)), null);

  if (typeof focusValue === "number" && focusValue >= 4.25) return "top";
  if (weakestScore !== null && weakestScore <= 2) return "watch";
  return "steady";
}

function getRowStateLabel(rowState) {
  if (rowState === "top") return "Top Performer";
  if (rowState === "watch") return "Needs Attention";
  return "Stable";
}

function getDetailKpiRows(record, focus = "performance") {
  if (focus === "all") {
    return [
      ["Transfer", `${record.transferRateDisplay} | Score ${scoreText(record.transferScore)}`],
      ["Admits", `${record.admitsCount === null || record.admitsCount === undefined ? "N/A" : Number(record.admitsCount).toFixed(0)} | Score ${scoreText(record.admitsScore)}`],
      ["AHT", `${record.ahtDisplay} | Score ${scoreText(record.ahtScore)}`],
      ["Attendance", `${record.attendancePercentDisplay} | Score ${scoreText(record.attendanceScore)}`],
      ["Quality Assurance", `${record.qaPercentDisplay} | Score ${scoreText(record.qaScore)}`],
      ["Overall", scoreText(record.overallScore)],
    ];
  }

  if (focus === "attendance") {
    return [
      ["Total Hours Present", record.hoursPresent === null || record.hoursPresent === undefined ? "N/A" : Number(record.hoursPresent).toFixed(2)],
      ["Total Hours Absent", record.hoursAbsent === null || record.hoursAbsent === undefined ? "N/A" : Number(record.hoursAbsent).toFixed(2)],
      ["Total Sick Leave Hours", record.sickLeaveHours === null || record.sickLeaveHours === undefined ? "N/A" : Number(record.sickLeaveHours).toFixed(2)],
      ["Undertime", record.undertimeHours === null || record.undertimeHours === undefined ? "N/A" : Number(record.undertimeHours).toFixed(2)],
      ["Required Hours", record.requiredHours === null || record.requiredHours === undefined ? "N/A" : Number(record.requiredHours).toFixed(2)],
      ["Attendance Percentage", record.attendancePercentDisplay],
      ["KPI Score", scoreText(record.attendanceScore)],
    ];
  }

  if (focus === "qa") {
    return [
      ["Quality Assurance", record.qaPercentDisplay],
      ["QA Score", scoreText(record.qaScore)],
      ["Last Updated", record.qaLastUpdatedDisplay || "N/A"],
    ];
  }

  if (focus === "transfer") {
    return [
      ["Transfer", `${record.transferRateDisplay} | Score ${scoreText(record.transferScore)}`],
      ["First-Time Callers", record.firstTimeCaller === null || record.firstTimeCaller === undefined ? "N/A" : Number(record.firstTimeCaller).toFixed(0)],
      ["Transfer Count", record.transferCount === null || record.transferCount === undefined ? "N/A" : Number(record.transferCount).toFixed(0)],
      ["Last Updated", record.lastUpdatedDisplay || "N/A"],
    ];
  }

  if (focus === "admits") {
    return [
      ["Admits", `${record.admitsCount === null || record.admitsCount === undefined ? "N/A" : Number(record.admitsCount).toFixed(0)} | Score ${scoreText(record.admitsScore)}`],
      ["Last Updated", record.admitsLastUpdatedDisplay || "N/A"],
    ];
  }

  if (focus === "aht") {
    return [
      ["AHT", `${record.ahtDisplay} | Score ${scoreText(record.ahtScore)}`],
      ["Inbound Calls", record.inboundCalls === null || record.inboundCalls === undefined ? "N/A" : Number(record.inboundCalls).toFixed(0)],
      ["Inbound Minutes", record.inboundMinutes || "N/A"],
      ["Hold Time", record.holdTime || "N/A"],
      ["Last Updated", record.lastUpdatedDisplay || "N/A"],
    ];
  }

  return [
    ["Transfer", `${record.transferRateDisplay} | Score ${scoreText(record.transferScore)}`],
    ["First-Time Callers", record.firstTimeCaller === null || record.firstTimeCaller === undefined ? "N/A" : Number(record.firstTimeCaller).toFixed(0)],
    ["Transfer Count", record.transferCount === null || record.transferCount === undefined ? "N/A" : Number(record.transferCount).toFixed(0)],
    ["Admits", `${record.admitsCount === null || record.admitsCount === undefined ? "N/A" : Number(record.admitsCount).toFixed(0)} | Score ${scoreText(record.admitsScore)}`],
    ["AHT", `${record.ahtDisplay} | Score ${scoreText(record.ahtScore)}`],
    ["Inbound Calls", record.inboundCalls === null || record.inboundCalls === undefined ? "N/A" : Number(record.inboundCalls).toFixed(0)],
    ["Inbound Minutes", record.inboundMinutes || "N/A"],
    ["Hold Time", record.holdTime || "N/A"],
    ["Overall", scoreText(record.overallScore)],
  ];
}

function renderPerformanceDetailGroups(record) {
  return `
    <div class="table-detail-groups table-detail-groups-performance">
      <article class="table-detail-group table-detail-group-transfer">
        <div class="table-detail-group-header">
          <span class="table-detail-group-label">Transfer</span>
          <span class="score-pill ${record.transferScore === null || record.transferScore === undefined || Number.isNaN(record.transferScore) ? "score-na" : `score-${Math.round(record.transferScore)}`}">${record.transferScore === null || record.transferScore === undefined || Number.isNaN(record.transferScore) ? "N/A" : Math.round(record.transferScore)}</span>
        </div>
        <strong>${record.transferRateDisplay}</strong>
        <p>First-Time Callers: ${record.firstTimeCaller === null || record.firstTimeCaller === undefined ? "N/A" : Number(record.firstTimeCaller).toFixed(0)}</p>
        <p>Transfer Count: ${record.transferCount === null || record.transferCount === undefined ? "N/A" : Number(record.transferCount).toFixed(0)}</p>
      </article>
      <article class="table-detail-group table-detail-group-admits">
        <div class="table-detail-group-header">
          <span class="table-detail-group-label">Admits</span>
          <span class="score-pill ${record.admitsScore === null || record.admitsScore === undefined || Number.isNaN(record.admitsScore) ? "score-na" : `score-${Math.round(record.admitsScore)}`}">${record.admitsScore === null || record.admitsScore === undefined || Number.isNaN(record.admitsScore) ? "N/A" : Math.round(record.admitsScore)}</span>
        </div>
        <strong>${record.admitsCount === null || record.admitsCount === undefined ? "N/A" : Number(record.admitsCount).toFixed(0)}</strong>
        <p>${record.admitsCount === null || record.admitsCount === undefined ? "Admits is not available in this source yet." : "Admit Count for this selected row."}</p>
      </article>
      <article class="table-detail-group table-detail-group-aht">
        <div class="table-detail-group-header">
          <span class="table-detail-group-label">AHT</span>
          <span class="score-pill ${record.ahtScore === null || record.ahtScore === undefined || Number.isNaN(record.ahtScore) ? "score-na" : `score-${Math.round(record.ahtScore)}`}">${record.ahtScore === null || record.ahtScore === undefined || Number.isNaN(record.ahtScore) ? "N/A" : Math.round(record.ahtScore)}</span>
        </div>
        <strong>${record.ahtDisplay}</strong>
        <p>Inbound Calls: ${record.inboundCalls === null || record.inboundCalls === undefined ? "N/A" : Number(record.inboundCalls).toFixed(0)}</p>
        <p>Inbound Minutes: ${record.inboundMinutes || "N/A"}</p>
        <p>Hold Time: ${record.holdTime || "N/A"}</p>
        ${record.lastUpdatedDisplay ? `<p>Last Updated: ${record.lastUpdatedDisplay}</p>` : ""}
      </article>
    </div>
  `;
}

function renderExpandedRow(record, focus = "performance") {
  const columns = getColumns(focus);
  const detailItems = focus === "performance" || focus === "realtime"
    ? renderPerformanceDetailGroups(record)
    : getDetailKpiRows(record, focus)
      .map(([label, value]) => `<div class="table-detail-item"><span>${label}</span><strong>${value}</strong></div>`)
      .join("");
  const focusHeadline = focus === "all"
    ? `Overall ${scoreText(record.overallScore)}`
    : focus === "performance" || focus === "realtime"
      ? `Performance ${scoreText(record.performanceScore)}`
      : focus === "transfer"
        ? `Transfer ${scoreText(record.transferScore)}`
        : focus === "admits"
          ? `Admits ${scoreText(record.admitsScore)}`
          : focus === "aht"
            ? `AHT ${scoreText(record.ahtScore)}`
            : focus === "attendance"
              ? `Attendance ${scoreText(record.attendanceScore)}`
              : `Quality Assurance ${scoreText(record.qaScore)}`;

  return `
    <tr class="table-detail-row">
      <td colspan="${columns.length}">
        <div class="table-detail-panel">
          <div class="table-detail-topline">
            <span class="table-row-state table-row-state-${getRowState(record, focus)}">${getRowStateLabel(getRowState(record, focus))}</span>
            <strong>${record.agentName} | Date Range ${record.weekEnding}</strong>
            <span>${focusHeadline}</span>
          </div>
          <div class="${focus === "performance" || focus === "realtime" ? "table-detail-shell" : "table-detail-grid"}">${detailItems}</div>
        </div>
      </td>
    </tr>
  `;
}

function renderKpiCell(record, key) {
  const map = {
    transfer: [record.transferRateDisplay, record.transferScore],
    admits: [record.admitsCount === null || record.admitsCount === undefined ? "N/A" : Number(record.admitsCount).toFixed(0), record.admitsScore],
    aht: [record.ahtDisplay, record.ahtScore],
    attendance: [record.attendancePercentDisplay, record.attendanceScore],
    qa: [record.qaPercentDisplay, record.qaScore],
  };

  const [value, score] = map[key];
  const scoreClass =
    score === null || score === undefined || Number.isNaN(score) ? "score-na" : `score-${Math.round(score)}`;
  const scoreDisplay = score === null || score === undefined || Number.isNaN(score) ? "N/A" : Number(score).toFixed(0);
  return `<div class="kpi-inline"><span>${value}</span><span class="score-pill ${scoreClass}">${scoreDisplay}</span></div>`;
}

function formatValue(column, record) {
  const attendanceAlertKeys = new Set(["hoursAbsent", "sickLeaveHours", "undertimeHours"]);
  if (column.type === "kpi") {
    return renderKpiCell(record, column.key);
  }
  if (column.type === "scorepill") {
    const value = record[column.key];
    const scoreClass =
      value === null || value === undefined || Number.isNaN(value) ? "score-na" : `score-${Math.round(value)}`;
    const scoreDisplay = value === null || value === undefined || Number.isNaN(value) ? "N/A" : Number(value).toFixed(0);
    return `<span class="score-pill ${scoreClass}">${scoreDisplay}</span>`;
  }
  if (column.type === "number0") {
    const value = record[column.key];
    return value === null || value === undefined || Number.isNaN(value) ? "N/A" : Number(value).toFixed(0);
  }
  if (column.type === "number2") {
    const value = record[column.key];
    if (value === null || value === undefined || Number.isNaN(value)) return "N/A";
    const formatted = Number(value).toFixed(2);
    if (attendanceAlertKeys.has(column.key) && Number(value) > 0) {
      return `<span class="table-cell-alert">${formatted}</span>`;
    }
    return formatted;
  }
  return record[column.key] ?? "";
}

function buildTableCell(column, record, options = {}) {
  const enableRowExpansion = options.enableRowExpansion !== false;
  if (column.key === "agentName") {
    return `<td data-label="${column.label}"><div class="table-row-trigger${enableRowExpansion ? "" : " is-static"}"><strong>${record.agentName}</strong></div></td>`;
  }
  return `<td data-label="${column.label}">${formatValue(column, record)}</td>`;
}

export function initializeTable(headRow, onSort, sortState, focus = "performance", options = {}) {
  const columns = getColumns(focus, options);
  headRow.innerHTML = "";
  columns.forEach((column) => {
    const th = document.createElement("th");
    th.scope = "col";
    th.textContent = getColumnSortLabel(column, sortState);
    th.addEventListener("click", () => onSort(column.key));
    headRow.appendChild(th);
  });
}

export function renderTableSummary(summaryElement, records, focus = "performance") {
  if (!summaryElement) return;
  if (!records.length) {
    summaryElement.innerHTML = "";
    return;
  }

  const metric = getSummaryMetricConfig(focus);
  const uniqueAgents = [...new Set(records.map((record) => record.agentName))];
  const singleAgentView = uniqueAgents.length === 1;
  const validScores = records
    .map((record) => record[metric.key])
    .filter((value) => typeof value === "number" && !Number.isNaN(value));
  const averageScore = validScores.length
    ? validScores.reduce((sum, value) => sum + value, 0) / validScores.length
    : null;
  const latestRecord = [...records]
    .sort((left, right) => (left.weekDate?.getTime() ?? 0) - (right.weekDate?.getTime() ?? 0))
    .at(-1);
  const topRecord = [...records]
    .filter((record) => typeof record[metric.key] === "number" && !Number.isNaN(record[metric.key]))
    .sort((left, right) => right[metric.key] - left[metric.key])[0];
  const lowRecord = [...records]
    .filter((record) => typeof record[metric.key] === "number" && !Number.isNaN(record[metric.key]))
    .sort((left, right) => left[metric.key] - right[metric.key])[0];

  if (singleAgentView) {
    summaryElement.innerHTML = `
      <div class="table-summary-card">
        <span class="table-summary-label">Agent In View</span>
        <strong>${uniqueAgents[0]}</strong>
        <p>Current filtered agent scope.</p>
      </div>
      <div class="table-summary-card">
        <span class="table-summary-label">Date Ranges Shown</span>
        <strong>${records.length}</strong>
        <p>Date-range rows available for this agent.</p>
      </div>
      <div class="table-summary-card">
        <span class="table-summary-label">Average ${metric.label}</span>
        <strong>${scoreText(averageScore)}</strong>
        <p>${metric.description}</p>
      </div>
      <div class="table-summary-card">
        <span class="table-summary-label">Latest ${metric.label}</span>
        <strong>${latestRecord ? `${latestRecord.weekEnding} | ${scoreText(latestRecord[metric.key])}` : "N/A"}</strong>
        <p>Most recent visible date range for this agent.</p>
      </div>
    `;
    return;
  }

  summaryElement.innerHTML = `
    <div class="table-summary-card">
      <span class="table-summary-label">Rows Shown</span>
      <strong>${records.length}</strong>
      <p>Current filtered records in view.</p>
    </div>
    <div class="table-summary-card">
      <span class="table-summary-label">Average ${metric.label}</span>
      <strong>${scoreText(averageScore)}</strong>
      <p>${metric.description}</p>
    </div>
    <div class="table-summary-card">
      <span class="table-summary-label">Highest ${metric.label}</span>
      <strong>${topRecord ? `${topRecord.agentName} ${scoreText(topRecord[metric.key])}` : "N/A"}</strong>
      <p>Best row in the current focus selection.</p>
    </div>
    <div class="table-summary-card">
      <span class="table-summary-label">Lowest ${metric.label}</span>
      <strong>${lowRecord ? `${lowRecord.agentName} ${scoreText(lowRecord[metric.key])}` : "N/A"}</strong>
      <p>Most urgent row to review for this focus.</p>
    </div>
  `;
}

export function renderTable(bodyElement, records, expandedRowKeys, onRowSelect, focus = "performance", options = {}) {
  const columns = getColumns(focus, options);
  const enableRowExpansion = options.enableRowExpansion !== false;
  bodyElement.innerHTML = "";
  if (!records.length) {
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.colSpan = columns.length;
    cell.innerHTML = '<div class="status-message">No rows match the current filters.</div>';
    row.appendChild(cell);
    bodyElement.appendChild(row);
    return;
  }

  records.forEach((record) => {
    const rowState = getRowState(record, focus);
    const row = document.createElement("tr");
    const isExpanded = enableRowExpansion && expandedRowKeys instanceof Set ? expandedRowKeys.has(record.key) : false;
    row.className = `table-data-row table-data-row-${rowState}${isExpanded ? " is-selected" : ""}${enableRowExpansion ? "" : " is-static"}`;
    if (enableRowExpansion) {
      row.tabIndex = 0;
      row.setAttribute("role", "button");
      row.setAttribute("aria-expanded", isExpanded ? "true" : "false");
      row.addEventListener("click", () => onRowSelect(record.key));
      row.addEventListener("keydown", (event) => {
        if (event.key === "Enter" || event.key === " ") {
          event.preventDefault();
          onRowSelect(record.key);
        }
      });
    }
    row.innerHTML = columns.map((column) => buildTableCell(column, record, options)).join("");
    bodyElement.appendChild(row);

    if (enableRowExpansion && isExpanded) {
      const detailRow = document.createElement("tbody");
      detailRow.innerHTML = renderExpandedRow(record, focus);
      bodyElement.appendChild(detailRow.firstElementChild);
    }
  });
}

function getSortValue(record, key) {
  const sortMap = {
    transfer: record.transferScore,
    admits: record.admitsScore,
    aht: record.ahtScore,
    attendance: record.attendanceScore,
    qa: record.qaScore,
  };
  return sortMap[key] ?? record[key];
}

export function sortRecords(records, sortState) {
  const direction = sortState.direction === "asc" ? 1 : -1;
  return [...records].sort((left, right) => {
    const leftValue = getSortValue(left, sortState.key);
    const rightValue = getSortValue(right, sortState.key);

    if (typeof leftValue === "number" && typeof rightValue === "number") {
      return (leftValue - rightValue) * direction;
    }

    return String(leftValue).localeCompare(String(rightValue)) * direction;
  });
}

export function exportRowsToCsv(records, focus = "performance", options = {}) {
  const columns = getColumns(focus, options);
  const csvColumns = columns.flatMap((column) => {
    if (column.type === "kpi") {
      return [
        { key: `${column.key}Value`, label: column.label, source: column.key, kind: "kpiValue" },
        { key: `${column.key}Score`, label: `${column.label} Score`, source: column.key, kind: "kpiScore" },
      ];
    }
    return [{ ...column, source: column.key, kind: column.type || "text" }];
  });

  const headerRow = csvColumns.map((column) => column.label).join(",");
  const lines = records.map((record) => {
    const values = csvColumns.map((column) => {
      if (column.kind === "kpiValue") {
        if (column.source === "transfer") return record.transferRateDisplay || "N/A";
        if (column.source === "admits") return record.admitsCount === null || record.admitsCount === undefined ? "N/A" : Number(record.admitsCount).toFixed(0);
        if (column.source === "aht") return record.ahtDisplay || "N/A";
        if (column.source === "attendance") return record.attendancePercentDisplay || "N/A";
        if (column.source === "qa") return record.qaPercentDisplay || "N/A";
      }
      if (column.kind === "kpiScore") {
        const scoreKeyMap = {
          transfer: "transferScore",
          admits: "admitsScore",
          aht: "ahtScore",
          attendance: "attendanceScore",
          qa: "qaScore",
        };
        const scoreValue = record[scoreKeyMap[column.source]];
        return scoreValue === null || scoreValue === undefined || Number.isNaN(scoreValue) ? "N/A" : Number(scoreValue).toFixed(0);
      }
      if (column.kind === "number0") return record[column.key] === null || record[column.key] === undefined ? "N/A" : Number(record[column.key]).toFixed(0);
      if (column.kind === "number2") return record[column.key] === null || record[column.key] === undefined ? "N/A" : Number(record[column.key]).toFixed(2);
      return record[column.key] ?? "";
    });
    return values.map((value) => `"${String(value).replace(/"/g, '""')}"`).join(",");
  });

  const csv = [headerRow, ...lines].join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `flyland-${focus}-dashboard.csv`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}
