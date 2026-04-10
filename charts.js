function scoreColor(score) {
  if (score <= 1.5) return "#d83b3b";
  if (score <= 2.5) return "#f28c28";
  if (score <= 3.5) return "#e8b100";
  if (score <= 4.5) return "#74c365";
  return "#1f9d55";
}

const CHART_COLORS = {
  transfer: "#3B82F6",
  admits: "#10B981",
  aht: "#F59E0B",
  attendance: "#8B5CF6",
  qa: "#EC4899",
  overall: "#1F2937",
  ranking: "#1F2937",
};

const KPI_TOOLTIP_MAP = [
  ["Transfer", "transferScore", "transferRateDisplay"],
  ["Admits", "admitsScore", "admitsCount"],
  ["AHT", "ahtScore", "ahtDisplay"],
  ["Attendance", "attendanceScore", "attendancePercentDisplay"],
  ["Quality Assurance", "qaScore", "qaPercentDisplay"],
];

const FOCUS_METRICS = {
  all: [
    { label: "Transfer", scoreKey: "transferScore", rawType: "transferRatePercent", color: CHART_COLORS.transfer },
    { label: "Admits", scoreKey: "admitsScore", rawType: "admitsCount", color: CHART_COLORS.admits },
    { label: "AHT", scoreKey: "ahtScore", rawType: "ahtSeconds", color: CHART_COLORS.aht },
    { label: "Attendance", scoreKey: "attendanceScore", rawType: "attendancePercentValue", color: CHART_COLORS.attendance },
    { label: "Quality Assurance", scoreKey: "qaScore", rawType: "qaPercentValue", color: CHART_COLORS.qa },
    { label: "Overall", scoreKey: "overallScore", rawType: "overallScore", color: CHART_COLORS.overall },
  ],
  performance: [
    { label: "Transfer", scoreKey: "transferScore", rawType: "transferRatePercent", color: CHART_COLORS.transfer },
    { label: "Admits", scoreKey: "admitsScore", rawType: "admitsCount", color: CHART_COLORS.admits },
    { label: "AHT", scoreKey: "ahtScore", rawType: "ahtSeconds", color: CHART_COLORS.aht },
    { label: "Performance", scoreKey: "performanceScore", rawType: "performanceScore", color: CHART_COLORS.overall },
  ],
  transfer: [
    { label: "Transfer", scoreKey: "transferScore", rawType: "transferRatePercent", color: CHART_COLORS.transfer },
  ],
  admits: [
    { label: "Admits", scoreKey: "admitsScore", rawType: "admitsCount", color: CHART_COLORS.admits },
  ],
  aht: [
    { label: "AHT", scoreKey: "ahtScore", rawType: "ahtSeconds", color: CHART_COLORS.aht },
  ],
  realtime: [
    { label: "Transfer", scoreKey: "transferScore", rawType: "transferRatePercent", color: CHART_COLORS.transfer },
    { label: "Admits", scoreKey: "admitsScore", rawType: "admitsCount", color: CHART_COLORS.admits },
    { label: "AHT", scoreKey: "ahtScore", rawType: "ahtSeconds", color: CHART_COLORS.aht },
    { label: "Performance", scoreKey: "performanceScore", rawType: "performanceScore", color: CHART_COLORS.overall },
  ],
  attendance: [
    { label: "Attendance", scoreKey: "attendanceScore", rawType: "attendancePercentValue", color: CHART_COLORS.attendance },
  ],
  qa: [
    { label: "Quality Assurance", scoreKey: "qaScore", rawType: "qaPercentValue", color: CHART_COLORS.qa },
  ],
};

const valueLabelPlugin = {
  id: "valueLabelPlugin",
  afterDatasetsDraw(chart, _args, pluginOptions) {
    if (!pluginOptions?.enabled) return;

    const { ctx } = chart;
    const formatter = pluginOptions.formatter || ((value) => String(value));
    const insideFormatter = pluginOptions.insideFormatter || null;
    const textColor = pluginOptions.color || "#10213d";
    const insideTextColor = pluginOptions.insideColor || "#ffffff";
    const fontSize = pluginOptions.fontSize || 11;
    const fontWeight = pluginOptions.fontWeight || "700";
    const offset = pluginOptions.offset ?? 8;
    const inside = pluginOptions.inside ?? false;

    ctx.save();
    ctx.fillStyle = textColor;
    ctx.font = `${fontWeight} ${fontSize}px Manrope, sans-serif`;
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";

    chart.data.datasets.forEach((dataset, datasetIndex) => {
      const meta = chart.getDatasetMeta(datasetIndex);
      if (meta.hidden) return;

      meta.data.forEach((element, index) => {
        const rawValue = dataset.data[index];
        if (rawValue === null || rawValue === undefined || Number.isNaN(rawValue)) return;

        const parsed = meta.controller.getParsed(index);
        let numericValue = Number(rawValue);
        if (typeof parsed === "object" && parsed !== null) {
          if (chart.config.type === "bar" && chart.options.indexAxis === "y") {
            numericValue = Number(parsed.x);
          } else if (chart.config.type === "bar") {
            numericValue = Number(parsed.y);
          } else {
            numericValue = Number(parsed.y ?? parsed.x ?? rawValue);
          }
        }
        if (Number.isNaN(numericValue)) return;

        const label = formatter(numericValue, {
          chart,
          dataset,
          datasetIndex,
          dataIndex: index,
        });
        if (!label) return;

        const position = element.tooltipPosition();
        let x = position.x;
        let y = position.y - offset;

        if (chart.config.type === "bar") {
          if (chart.options.indexAxis === "y") {
            x = inside ? position.x - offset : position.x + offset;
            ctx.textAlign = inside ? "right" : "left";
            y = position.y;
          } else if (inside) {
            y = position.y + 10;
          }
        }

        if (chart.chartArea) {
          const topLimit = chart.chartArea.top + fontSize;
          const bottomLimit = chart.chartArea.bottom - fontSize;
          y = Math.min(Math.max(y, topLimit), bottomLimit);
        }

        if (insideFormatter && chart.config.type === "bar" && !inside) {
          const insideLabel = insideFormatter(numericValue, {
            chart,
            dataset,
            datasetIndex,
            dataIndex: index,
          });
          if (insideLabel) {
            const insidePosition = element.tooltipPosition();
            const insideX = insidePosition.x;
            const insideY = Math.min(
              Math.max(insidePosition.y + 10, chart.chartArea?.top + fontSize || insidePosition.y),
              chart.chartArea?.bottom - fontSize || insidePosition.y
            );
            ctx.fillStyle = insideTextColor;
            ctx.fillText(insideLabel, insideX, insideY);
            ctx.fillStyle = textColor;
          }
        }

        ctx.fillText(label, x, y);
      });
    });

    ctx.restore();
  },
};

const strictChartAreaTooltipPlugin = {
  id: "strictChartAreaTooltipPlugin",
  beforeEvent(chart, args, pluginOptions) {
    if (!pluginOptions?.enabled) return;

    const event = args.event;
    if (!event || (event.type !== "mousemove" && event.type !== "click")) return;

    const area = chart.chartArea;
    if (!area) return;

    const outsideChartArea =
      event.x < area.left || event.x > area.right || event.y < area.top || event.y > area.bottom;

    if (outsideChartArea) {
      chart.tooltip?.setActiveElements([], { x: 0, y: 0 });
      chart.setActiveElements([]);
      args.changed = true;
    }
  },
};

const exactTouchTooltipPlugin = {
  id: "exactTouchTooltipPlugin",
  afterEvent(chart, args, pluginOptions) {
    if (!pluginOptions?.enabled) return;

    const event = args.event;
    if (!event || !["click", "touchstart", "touchmove", "mousemove"].includes(event.type)) return;

    const matches = [];
    if (pluginOptions.strategy === "horizontal-bar-row") {
      const meta = chart.getSortedVisibleDatasetMetas()[0];
      const area = chart.chartArea;
      meta?.data?.forEach((element, index) => {
        const props = element.getProps(["y", "height"], true);
        const left = area?.left ?? 0;
        const right = area?.right ?? 0;
        const top = props.y - props.height / 2 - 6;
        const bottom = props.y + props.height / 2 + 6;

        if (event.y >= top && event.y <= bottom && event.x >= left && event.x <= right) {
          matches.push({
            datasetIndex: meta.index,
            index,
          });
        }
      });
    } else if (pluginOptions.strategy === "stacked-horizontal-row-all") {
      const metas = chart.getSortedVisibleDatasetMetas();
      const leadMeta = metas[0];
      const area = chart.chartArea;
      leadMeta?.data?.forEach((element, index) => {
        const props = element.getProps(["y", "height"], true);
        const left = area?.left ?? 0;
        const right = area?.right ?? 0;
        const top = props.y - props.height / 2 - 8;
        const bottom = props.y + props.height / 2 + 8;

        if (event.y >= top && event.y <= bottom && event.x >= left && event.x <= right) {
          metas.forEach((meta) => {
            matches.push({
              datasetIndex: meta.index,
              index,
            });
          });
        }
      });
    } else {
      chart.getSortedVisibleDatasetMetas().forEach((meta) => {
        meta.data.forEach((element, index) => {
          if (typeof element.inRange === "function" && element.inRange(event.x, event.y, true)) {
            matches.push({
              datasetIndex: meta.index,
              index,
            });
          }
        });
      });
    }

    if (!matches.length) {
      if (pluginOptions.clearOnMiss) {
        chart.setActiveElements([]);
        chart.tooltip?.setActiveElements([], { x: 0, y: 0 });
        args.changed = true;
      }
      return;
    }

    const [match] = matches;
    const element = chart.getDatasetMeta(match.datasetIndex)?.data?.[match.index];
    if (!element) return;

    const props = element.getProps(["x", "y", "base"], true);
    const position = pluginOptions.strategy === "horizontal-bar-row" || pluginOptions.strategy === "stacked-horizontal-row-all"
      ? {
        x: (Number(props.x) + Number(props.base)) / 2,
        y: Number(props.y),
      }
      : element.tooltipPosition();
    chart.setActiveElements(matches);
    chart.tooltip?.setActiveElements(matches, position);
    args.changed = true;
  },
};

Chart.register(valueLabelPlugin, strictChartAreaTooltipPlugin, exactTouchTooltipPlugin);

function destroyIfExists(chart) {
  if (chart?.canvas?.dataset?.externalTooltipId) {
    const tooltip = document.getElementById(chart.canvas.dataset.externalTooltipId);
    tooltip?.remove();
    delete chart.canvas.dataset.externalTooltipId;
  }
  if (chart) chart.destroy();
}

function getOrCreateExternalTooltip(canvas) {
  if (!canvas?.id) return null;
  const tooltipId = `${canvas.id}-external-tooltip`;
  let tooltip = document.getElementById(tooltipId);

  if (!tooltip) {
    tooltip = document.createElement("div");
    tooltip.id = tooltipId;
    tooltip.className = "chart-external-tooltip";
    document.body.appendChild(tooltip);
  }

  canvas.dataset.externalTooltipId = tooltipId;
  return tooltip;
}

function renderTooltipSection(lines, className = "") {
  if (!lines?.length) return "";
  return `<div class="${className}">${lines.map((line) => `<div>${line}</div>`).join("")}</div>`;
}

function externalDistributionTooltip(context) {
  const { chart, tooltip } = context;
  const tooltipEl = getOrCreateExternalTooltip(chart.canvas);
  if (!tooltipEl) return;

  if (tooltip.opacity === 0) {
    tooltipEl.style.opacity = "0";
    tooltipEl.style.pointerEvents = "none";
    return;
  }

  const titleLines = tooltip.title || [];
  const bodyLines = (tooltip.body || []).flatMap((bodyItem) => bodyItem.lines || []);
  const afterBodyLines = tooltip.afterBody || [];
  const footerLines = tooltip.footer || [];

  tooltipEl.innerHTML = [
    renderTooltipSection(titleLines, "chart-external-tooltip-title"),
    renderTooltipSection(bodyLines, "chart-external-tooltip-body"),
    renderTooltipSection(afterBodyLines, "chart-external-tooltip-body chart-external-tooltip-agents"),
    renderTooltipSection(footerLines, "chart-external-tooltip-footer"),
  ].join("");

  const canvasRect = chart.canvas.getBoundingClientRect();
  const tooltipRect = tooltipEl.getBoundingClientRect();
  const viewportWidth = window.innerWidth;
  const viewportHeight = window.innerHeight;
  const gap = 14;

  let left = canvasRect.left + window.scrollX + tooltip.caretX - tooltipRect.width / 2;
  let top = canvasRect.top + window.scrollY + tooltip.caretY - tooltipRect.height - gap;

  if (left + tooltipRect.width > window.scrollX + viewportWidth - 12) {
    left = window.scrollX + viewportWidth - tooltipRect.width - 12;
  }
  if (left < window.scrollX + 12) {
    left = window.scrollX + 12;
  }

  if (top < window.scrollY + 12) {
    top = canvasRect.top + window.scrollY + tooltip.caretY + gap;
  }
  if (top + tooltipRect.height > window.scrollY + viewportHeight - 12) {
    top = Math.max(window.scrollY + 12, window.scrollY + viewportHeight - tooltipRect.height - 12);
  }

  tooltipEl.style.opacity = "1";
  tooltipEl.style.pointerEvents = "none";
  tooltipEl.style.left = `${left}px`;
  tooltipEl.style.top = `${top}px`;
}

function externalStackedTooltip(context) {
  const { chart, tooltip } = context;
  const tooltipEl = getOrCreateExternalTooltip(chart.canvas);
  if (!tooltipEl) return;

  if (tooltip.opacity === 0) {
    tooltipEl.style.opacity = "0";
    tooltipEl.style.pointerEvents = "none";
    return;
  }

  const title = tooltip.title?.[0] || "";
  const uniqueDataPoints = (tooltip.dataPoints || []).map((dataPoint, index) => ({
    dataPoint,
    bodyLines: tooltip.body?.[index]?.lines || [],
  })).filter(({ dataPoint }, index, items) => {
    const datasetLabel = dataPoint?.dataset?.label || "";
    const dataIndex = dataPoint?.dataIndex ?? -1;
    return index === items.findIndex((item) => {
      const itemLabel = item?.dataPoint?.dataset?.label || "";
      const itemIndex = item?.dataPoint?.dataIndex ?? -1;
      return itemLabel === datasetLabel && itemIndex === dataIndex;
    });
  });

  const tooltipRows = uniqueDataPoints.map(({ dataPoint, bodyLines }) => {
    const color = dataPoint?.dataset?.backgroundColor || "#10213d";
    const kpiLabel = dataPoint?.dataset?.label || "";
    return `
      <div class="chart-external-tooltip-row">
        <div class="chart-external-tooltip-kpi">
          <span class="chart-external-tooltip-swatch" style="background:${color}"></span>
          <strong>${kpiLabel}</strong>
        </div>
        <div class="chart-external-tooltip-body">
          ${bodyLines.map((line) => `<div>${line}</div>`).join("")}
        </div>
      </div>
    `;
  }).join("");
  const footer = tooltip.footer?.[0] || "";

  tooltipEl.innerHTML = `
    <div class="chart-external-tooltip-title">${title}</div>
    <div class="chart-external-tooltip-rows">${tooltipRows}</div>
    ${footer ? `<div class="chart-external-tooltip-footer">${footer}</div>` : ""}
  `;

  const canvasRect = chart.canvas.getBoundingClientRect();
  const tooltipRect = tooltipEl.getBoundingClientRect();
  const viewportWidth = window.innerWidth;
  const viewportHeight = window.innerHeight;
  const gap = 14;

  let left = canvasRect.left + window.scrollX + tooltip.caretX - tooltipRect.width / 2;
  let top = canvasRect.top + window.scrollY + tooltip.caretY - tooltipRect.height - gap;

  if (left + tooltipRect.width > window.scrollX + viewportWidth - 12) {
    left = window.scrollX + viewportWidth - tooltipRect.width - 12;
  }
  if (left < window.scrollX + 12) {
    left = window.scrollX + 12;
  }

  if (top < window.scrollY + 12) {
    top = canvasRect.top + window.scrollY + tooltip.caretY + gap;
  }
  if (top + tooltipRect.height > window.scrollY + viewportHeight - 12) {
    top = Math.max(window.scrollY + 12, window.scrollY + viewportHeight - tooltipRect.height - 12);
  }

  tooltipEl.style.opacity = "1";
  tooltipEl.style.pointerEvents = "none";
  tooltipEl.style.left = `${left}px`;
  tooltipEl.style.top = `${top}px`;
}

function buildBaseOptions() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    interaction: {
      mode: "index",
      intersect: false,
    },
    plugins: {
      legend: {
        labels: {
          usePointStyle: true,
          boxWidth: 10,
          boxHeight: 10,
        },
      },
      tooltip: {
        backgroundColor: "rgba(16, 33, 61, 0.96)",
        titleFont: {
          size: 16,
          weight: "700",
        },
        bodyFont: {
          size: 15,
        },
        footerFont: {
          size: 14,
          weight: "700",
        },
        padding: 16,
        cornerRadius: 12,
        callbacks: {
          label(context) {
            const value = typeof context.parsed === "object" ? context.parsed.y ?? context.parsed.x : context.parsed;
            return `${context.dataset.label}: ${Number(value).toFixed(2)}`;
          },
        },
      },
    },
    scales: {
      y: {
        beginAtZero: true,
        max: 5,
        grid: {
          color: "rgba(16, 33, 61, 0.08)",
        },
        ticks: {
          stepSize: 1,
        },
      },
      x: {
        grid: {
          display: false,
        },
      },
    },
  };
}

function formatPercent(value) {
  if (value === null || value === undefined || Number.isNaN(value)) return "N/A";
  return `${Number(value).toFixed(2)}%`;
}

function formatTime(seconds) {
  if (seconds === null || seconds === undefined || Number.isNaN(seconds)) return "N/A";
  const mins = Math.floor(seconds / 60);
  const secs = Math.round(seconds % 60);
  return `${String(mins).padStart(2, "0")}:${String(secs).padStart(2, "0")}`;
}

function hexToRgba(hex, alpha) {
  const cleaned = hex.replace("#", "");
  const value = cleaned.length === 3
    ? cleaned.split("").map((char) => char + char).join("")
    : cleaned;
  const red = Number.parseInt(value.slice(0, 2), 16);
  const green = Number.parseInt(value.slice(2, 4), 16);
  const blue = Number.parseInt(value.slice(4, 6), 16);
  return `rgba(${red}, ${green}, ${blue}, ${alpha})`;
}

function isMobileViewport() {
  return typeof window !== "undefined" && window.innerWidth <= 720;
}

function getMobileKpiLabel(label) {
  const labelMap = {
    Transfer: "TR",
    Admits: "AD",
    AHT: "AHT",
    Attendance: "ATT",
    "Quality Assurance": "QA",
    QA: "QA",
    Overall: "OVR",
  };
  return isMobileViewport() ? (labelMap[label] || label) : label;
}

function formatShortDateRangeLabel(label) {
  if (!label) return "";
  const [start = "", end = ""] = String(label).split(" - ").map((part) => part.trim());
  const shorten = (value) => {
    const pieces = value.split("/");
    if (pieces.length < 2) return value;
    return `${pieces[0].padStart(2, "0")}/${pieces[1].padStart(2, "0")}`;
  };
  if (!end) return shorten(start);
  return `${shorten(start)} - ${shorten(end)}`;
}

function getChartDisplayName(agentName) {
  if (!agentName) return "";
  const parts = String(agentName).split(",");
  if (parts.length < 2) return String(agentName).trim();
  return parts.slice(1).join(",").trim() || String(agentName).trim();
}

function formatSnapshotLabel(label, weekSummary) {
  const rawMap = {
    Transfer: formatPercent(weekSummary.transferRatePercent),
    Admits: weekSummary.admitsCount === null || weekSummary.admitsCount === undefined ? "N/A" : Number(weekSummary.admitsCount).toFixed(2),
    AHT: formatTime(weekSummary.ahtSeconds),
    Attendance: formatPercent(weekSummary.attendancePercentValue),
    "Quality Assurance": formatPercent(weekSummary.qaPercentValue),
    QA: formatPercent(weekSummary.qaPercentValue),
    Overall: weekSummary.overallScore === null || weekSummary.overallScore === undefined ? "N/A" : Number(weekSummary.overallScore).toFixed(2),
    Performance: weekSummary.performanceScore === null || weekSummary.performanceScore === undefined ? "N/A" : Number(weekSummary.performanceScore).toFixed(2),
  };

  return rawMap[label] || "";
}

function formatComparisonRaw(label, weekSummary) {
  if (!weekSummary) return "Average: N/A";
  const includesQa = Boolean(weekSummary.overallIncludesQa);

  const rawMap = {
    Transfer: `Avg rate: ${formatPercent(weekSummary.transferRatePercent)}`,
    Admits: `Avg count: ${
      weekSummary.admitsCount === null || weekSummary.admitsCount === undefined
        ? "N/A"
        : Number(weekSummary.admitsCount).toFixed(2)
    }`,
    AHT: `Avg time: ${formatTime(weekSummary.ahtSeconds)}`,
    Attendance: `Avg attendance: ${formatPercent(weekSummary.attendancePercentValue)}`,
    "Quality Assurance": `Avg QA: ${formatPercent(weekSummary.qaPercentValue)}`,
    QA: `Avg QA: ${formatPercent(weekSummary.qaPercentValue)}`,
    Overall: `${includesQa ? "Weighted overall" : "Weighted overall | QA pending"}: ${
      weekSummary.overallScore === null || weekSummary.overallScore === undefined
        ? "N/A"
        : Number(weekSummary.overallScore).toFixed(2)
    }`,
    Performance: `Performance score: ${
      weekSummary.performanceScore === null || weekSummary.performanceScore === undefined
        ? "N/A"
        : Number(weekSummary.performanceScore).toFixed(2)
    }`,
  };

  return rawMap[label] || "Average: N/A";
}

function isValidMetric(value) {
  return typeof value === "number" && !Number.isNaN(value);
}

function getOverallBreakdown(summary) {
  if (!summary) {
    return {
      includesQa: false,
      weights: { performance: 0, attendance: 0, qa: 0 },
    };
  }

  return {
    includesQa: Boolean(summary.overallIncludesQa),
    weights: {
      performance: summary.overallWeights?.performance ?? 0,
      attendance: summary.overallWeights?.attendance ?? 0,
      qa: summary.overallWeights?.qa ?? 0,
    },
  };
}

function getWeakestKpi(record) {
  return KPI_TOOLTIP_MAP
    .map(([label, scoreKey, rawKey]) => ({
      label,
      score: record?.[scoreKey],
      rawValue: record?.[rawKey],
    }))
    .filter((item) => isValidMetric(item.score))
    .sort((left, right) => left.score - right.score || left.label.localeCompare(right.label))[0] || null;
}

function getFocusMetrics(focus = "performance") {
  return FOCUS_METRICS[focus] || FOCUS_METRICS.performance;
}

function formatMetricRaw(metric, item) {
  switch (metric?.rawType) {
    case "transferRatePercent":
      return `Avg rate: ${formatPercent(item?.transferRatePercent)}`;
    case "admitsCount":
      return `Avg count: ${item?.admitsCount === null || item?.admitsCount === undefined ? "N/A" : Number(item.admitsCount).toFixed(2)}`;
    case "ahtSeconds":
      return `Avg time: ${formatTime(item?.ahtSeconds)}`;
    case "attendancePercentValue":
      return `Avg attendance: ${formatPercent(item?.attendancePercentValue)}`;
    case "qaPercentValue":
      return `Avg QA: ${formatPercent(item?.qaPercentValue)}`;
    case "overallScore":
      return `Overall score: ${item?.overallScore === null || item?.overallScore === undefined ? "N/A" : Number(item.overallScore).toFixed(2)}`;
    case "performanceScore":
      return `Performance score: ${item?.performanceScore === null || item?.performanceScore === undefined ? "N/A" : Number(item.performanceScore).toFixed(2)}`;
    default:
      return "Average: N/A";
  }
}

export function renderLineChart(canvas, chart, weeklyAverages, focus = "performance") {
  destroyIfExists(chart);
  const isMobile = isMobileViewport();
  const trendPoints = isMobile ? weeklyAverages.slice(-4) : weeklyAverages;
  const metrics = getFocusMetrics(focus);
  return new Chart(canvas, {
    type: "line",
    data: {
      labels: trendPoints.map((item) => isMobile ? formatShortDateRangeLabel(item.weekEnding) : item.weekEnding),
      datasets: metrics.map((metric) => ({
        label: metric.label,
        data: trendPoints.map((item) => item[metric.scoreKey]),
        borderColor: metric.color,
        backgroundColor: metric.color,
        borderWidth: metric.label === "Performance" ? 3 : 2,
        tension: 0.35,
      })),
    },
    options: {
      ...buildBaseOptions(),
      plugins: {
        ...buildBaseOptions().plugins,
          tooltip: {
            ...buildBaseOptions().plugins.tooltip,
            callbacks: {
              title(items) {
                const item = trendPoints[items[0]?.dataIndex ?? -1];
                return item?.weekEnding || items[0]?.label || "";
              },
              label(context) {
                const item = trendPoints[context.dataIndex];
                const score = Number(context.parsed.y).toFixed(2);
                const metric = metrics.find((entry) => entry.label === context.dataset.label);
                return [`Score: ${score}`, formatMetricRaw(metric, item)];
              },
            },
        },
      },
    },
  });
}

export function renderBarChart(canvas, chart, weekSummary, focus = "performance") {
  destroyIfExists(chart);
  const metrics = getFocusMetrics(focus);
  const chartLabels = metrics.map((metric) => metric.label);
  return new Chart(canvas, {
    type: "bar",
    data: {
      labels: chartLabels.map((label) => getMobileKpiLabel(label)),
      datasets: [
        {
          label: "Average Score",
          data: metrics.map((metric) => weekSummary[metric.scoreKey]),
          backgroundColor: metrics.map((metric) => metric.color),
          borderRadius: 10,
        },
      ],
    },
    options: {
      ...buildBaseOptions(),
      plugins: {
        ...buildBaseOptions().plugins,
        valueLabelPlugin: {
          enabled: true,
          color: "#10213d",
          fontSize: 11,
          insideFormatter(value) {
            return Number(value).toFixed(2);
          },
          insideColor: "#ffffff",
          formatter(_value, context) {
            const label = chartLabels[context.dataIndex];
            return formatSnapshotLabel(label, weekSummary);
          },
        },
        tooltip: {
          ...buildBaseOptions().plugins.tooltip,
          callbacks: {
            label(context) {
              const label = chartLabels[context.dataIndex];
              const rawMap = {
                Transfer: `Avg rate: ${formatPercent(weekSummary.transferRatePercent)}`,
                Admits: `Avg count: ${weekSummary.admitsCount === null || weekSummary.admitsCount === undefined ? "N/A" : Number(weekSummary.admitsCount).toFixed(2)}`,
                AHT: `Avg time: ${formatTime(weekSummary.ahtSeconds)}`,
                Attendance: `Avg attendance: ${formatPercent(weekSummary.attendancePercentValue)}`,
                "Quality Assurance": `Avg QA: ${formatPercent(weekSummary.qaPercentValue)}`,
                Overall: `Weighted score`,
              };
              return [`Score: ${Number(context.parsed.y).toFixed(2)}`, rawMap[label]];
            },
          },
        },
      },
    },
  });
}

export function renderContributionChart(canvas, chart, weekSummary) {
  destroyIfExists(chart);
  const overallBreakdown = getOverallBreakdown(weekSummary);
  const performanceAvailable = isValidMetric(weekSummary.performanceScore);
  const attendanceAvailable = isValidMetric(weekSummary.attendanceScore);
  const qaAvailable = isValidMetric(weekSummary.qaScore);
  const contributionConfig = [
    {
      key: "performance",
      label: performanceAvailable
        ? `Performance (${Math.round(overallBreakdown.weights.performance * 100)}%)`
        : "Performance (Missing)",
      value: performanceAvailable ? weekSummary.performanceScore * overallBreakdown.weights.performance : 0,
      score: weekSummary.performanceScore,
      color: performanceAvailable ? CHART_COLORS.overall : "#d5ddea",
      detail: performanceAvailable
        ? `Performance score: ${Number(weekSummary.performanceScore).toFixed(2)}`
        : "Transfer, Admits, or AHT data is incomplete",
    },
    {
      key: "attendance",
      label: attendanceAvailable
        ? `Attendance (${Math.round(overallBreakdown.weights.attendance * 100)}%)`
        : "Attendance (Missing)",
      value: attendanceAvailable ? weekSummary.attendanceScore * overallBreakdown.weights.attendance : 0,
      score: weekSummary.attendanceScore,
      color: attendanceAvailable ? CHART_COLORS.attendance : "#d5ddea",
      detail: attendanceAvailable
        ? `Attendance score: ${Number(weekSummary.attendanceScore).toFixed(2)}`
        : "Attendance data is unavailable",
    },
    {
      key: "qa",
      label: qaAvailable ? `QA (${Math.round(overallBreakdown.weights.qa * 100)}%)` : "QA (Pending)",
      value: qaAvailable ? weekSummary.qaScore * overallBreakdown.weights.qa : 0,
      score: weekSummary.qaScore,
      color: qaAvailable ? CHART_COLORS.qa : "#d5ddea",
      detail: qaAvailable
        ? `QA score: ${Number(weekSummary.qaScore).toFixed(2)}`
        : "QA data is pending and excluded from overall where missing",
    },
  ];
  return new Chart(canvas, {
    type: "doughnut",
    data: {
      labels: contributionConfig.map((item) => item.label),
      datasets: [
        {
          data: contributionConfig.map((item) => item.value),
          backgroundColor: contributionConfig.map((item) => item.color),
          borderColor: "#ffffff",
          borderWidth: 2,
          hoverOffset: 8,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: "bottom",
        },
        tooltip: {
          ...buildBaseOptions().plugins.tooltip,
          callbacks: {
            title(items) {
              const index = items[0]?.dataIndex ?? -1;
              return kpiLabels[index] || "";
            },
            label(context) {
              const item = contributionConfig[context.dataIndex];
              if (!item) return "";
              if (!isValidMetric(item.score)) {
                return ["Contribution unavailable", item.detail];
              }
              return [
                `Contribution: ${Number(item.value).toFixed(2)}`,
                item.detail,
              ];
            },
          },
        },
      },
      cutout: "68%",
    },
  });
}

export function renderVarianceChart(canvas, chart, weeklyAverages, focus = "performance") {
  destroyIfExists(chart);
  const isMobile = isMobileViewport();
  const currentWeek = weeklyAverages.at(-1) || null;
  const previousWeek = weeklyAverages.length > 1 ? weeklyAverages[0] : null;
  const metrics = getFocusMetrics(focus);
  const kpiLabels = metrics.map((metric) => metric.label);
  const kpiKeys = metrics.map((metric) => metric.scoreKey);
  const colors = metrics.map((metric) => metric.color);

  const varianceValues = kpiKeys.map((key) => {
    if (
      !currentWeek ||
      currentWeek[key] === null ||
      currentWeek[key] === undefined ||
      Number.isNaN(currentWeek[key]) ||
      !previousWeek ||
      previousWeek[key] === null ||
      previousWeek[key] === undefined ||
      Number.isNaN(previousWeek[key])
    ) {
      return 0;
    }
    return Number(currentWeek[key]) - Number(previousWeek[key]);
  });

  return new Chart(canvas, {
    type: "bar",
    data: {
      labels: kpiLabels.map((label) => getMobileKpiLabel(label)),
      datasets: [
        {
          label: "Variance",
          data: varianceValues,
          backgroundColor: varianceValues.map((value, index) =>
            value >= 0 ? colors[index] : hexToRgba(colors[index], 0.28)
          ),
          borderColor: colors,
          borderWidth: 1.2,
          borderRadius: 10,
          maxBarThickness: 34,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: {
        mode: "nearest",
        axis: "x",
        intersect: false,
      },
      plugins: {
        legend: {
          display: false,
        },
        strictChartAreaTooltipPlugin: {
          enabled: true,
        },
        exactTouchTooltipPlugin: {
          enabled: isMobile,
        },
        valueLabelPlugin: {
          enabled: true,
          color: "#10213d",
          fontSize: isMobile ? 10 : 11,
          fontWeight: "800",
          offset: isMobile ? 6 : 10,
          formatter(value) {
            const prefix = value > 0 ? "+" : "";
            return `${prefix}${Number(value).toFixed(2)}`;
          },
        },
        tooltip: {
          ...buildBaseOptions().plugins.tooltip,
          callbacks: {
            title(items) {
              return items[0]?.label || "";
            },
            label(context) {
              const currentValue = currentWeek?.[kpiKeys[context.dataIndex]];
              const previousValue = previousWeek?.[kpiKeys[context.dataIndex]];
              const variance = Number(context.parsed.y);
              const prefix = variance > 0 ? "+" : "";
              return [
                `Variance: ${prefix}${variance.toFixed(2)}`,
                `Selected range: ${currentValue === null || currentValue === undefined ? "N/A" : Number(currentValue).toFixed(2)}`,
                `Previous range: ${previousValue === null || previousValue === undefined ? "N/A" : Number(previousValue).toFixed(2)}`,
              ];
            },
          },
        },
      },
      scales: {
        y: {
          beginAtZero: false,
          grid: {
            color: "rgba(16, 33, 61, 0.08)",
          },
          ticks: {
            callback(value) {
              const prefix = Number(value) > 0 ? "+" : "";
              return `${prefix}${Number(value).toFixed(1)}`;
            },
          },
        },
        x: {
          grid: {
            display: false,
          },
          ticks: {
            font: {
              size: isMobile ? 10 : 12,
            },
          },
        },
      },
      layout: {
        padding: {
          top: isMobile ? 14 : 4,
          right: 6,
        },
      },
    },
  });
}

export function renderRankingList(container, records, options = {}) {
  if (!container) return;
  const scoreKey = options.scoreKey || "overallScore";
  const scoreLabel = options.scoreLabel || "Overall";
  const ranked = [...records]
    .filter((record) => typeof record[scoreKey] === "number" && !Number.isNaN(record[scoreKey]))
    .sort((left, right) => right[scoreKey] - left[scoreKey])
    .slice(0, 5);
  if (!ranked.length) {
    container.innerHTML = '<div class="status-message">No ranked rows are available for the current selection.</div>';
    return;
  }

  container.innerHTML = ranked.map((record, index) => {
    const weakestKpi = getWeakestKpi(record);
    const rankClass = index === 0 ? "is-rank-1" : index === 1 ? "is-rank-2" : index === 2 ? "is-rank-3" : "";
    const width = Math.max(8, Math.min(100, (Number(record[scoreKey]) / 5) * 100));
    const secondaryMeta = scoreKey === "performanceScore"
      ? (weakestKpi ? `Weakest: ${weakestKpi.label} (${Number(weakestKpi.score).toFixed(0)})` : "Weakest: N/A")
      : `Overall ${isValidMetric(record?.overallScore) ? Number(record.overallScore).toFixed(2) : "N/A"}`;
    return `
      <article class="ranking-list-item ${rankClass}" role="listitem">
        <div class="ranking-list-rank">#${index + 1}</div>
        <div class="ranking-list-main">
          <div class="ranking-list-topline">
            <strong>${getChartDisplayName(record.agentName)}</strong>
            <span>${Number(record[scoreKey]).toFixed(2)}</span>
          </div>
          <div class="ranking-list-bar-track" aria-hidden="true">
            <span class="ranking-list-bar-fill" style="width: ${width}%"></span>
          </div>
          <div class="ranking-list-meta">
            <span>${scoreLabel} ${isValidMetric(record?.[scoreKey]) ? Number(record[scoreKey]).toFixed(2) : "N/A"}</span>
            <span>${secondaryMeta}</span>
          </div>
        </div>
      </article>
    `;
  }).join("");
}

export function renderStackedBarChart(canvas, chart, records, onRowSelect, focus = "performance") {
  destroyIfExists(chart);
  const isMobile = isMobileViewport();
  const metrics = getFocusMetrics(focus).filter((metric) => !["Performance", "Overall"].includes(metric.label));
  return new Chart(canvas, {
    type: "bar",
    data: {
      labels: records.map((record) => getChartDisplayName(record.agentName)),
      datasets: metrics.map((metric) => ({
        label: metric.label,
        data: records.map((record) => record[metric.scoreKey]),
        backgroundColor: metric.color,
        borderRadius: 8,
        maxBarThickness: 18,
        categoryPercentage: isMobile ? 0.9 : 0.72,
        barPercentage: isMobile ? 0.98 : 0.82,
      })),
    },
    options: {
      ...buildBaseOptions(),
      indexAxis: isMobile ? "y" : "x",
      interaction: {
        mode: isMobile ? "nearest" : "index",
        axis: isMobile ? "y" : "x",
        intersect: isMobile,
      },
      plugins: {
        ...buildBaseOptions().plugins,
        strictChartAreaTooltipPlugin: {
          enabled: true,
        },
        exactTouchTooltipPlugin: {
          enabled: isMobile,
          strategy: isMobile ? "stacked-horizontal-row-all" : undefined,
          clearOnMiss: isMobile,
        },
        tooltip: {
          ...buildBaseOptions().plugins.tooltip,
          enabled: false,
          external: isMobile ? undefined : externalStackedTooltip,
          callbacks: {
            title(items) {
              return items[0]?.label || "";
            },
            label(context) {
              const record = records[context.dataIndex];
              const parsedValue = Number(context.parsed?.[isMobile ? "x" : "y"]);
              const rawMap = {
                Transfer: `Avg rate: ${formatPercent(record?.transferRatePercent)}`,
                Admits: `Avg count: ${record?.admitsCount === null || record?.admitsCount === undefined ? "N/A" : Number(record.admitsCount).toFixed(0)}`,
                AHT: `Avg time: ${formatTime(record?.ahtSeconds)}`,
                Attendance: `Avg attendance: ${formatPercent(record?.attendancePercentValue)}`,
                "Quality Assurance": `Avg QA: ${formatPercent(record?.qaPercentValue)}`,
              };
              return [`Score: ${Number.isFinite(parsedValue) ? parsedValue.toFixed(2) : "N/A"}`, rawMap[context.dataset.label]];
            },
            footer(items) {
              const total = items.reduce((sum, item) => sum + Number(item.parsed?.[isMobile ? "x" : "y"] || 0), 0);
              return `Combined KPI total: ${total.toFixed(2)}`;
            },
          },
        },
      },
      scales: {
        x: {
          stacked: true,
          beginAtZero: true,
          offset: true,
          grid: { display: false },
          ticks: {
            font: {
              size: isMobile ? 10 : 12,
            },
          },
        },
        y: {
          stacked: true,
          beginAtZero: true,
          grid: { color: "rgba(16, 33, 61, 0.08)" },
          ticks: {
            font: {
              size: isMobile ? 10 : 11,
            },
            callback(value, index) {
              const label = this.getLabelForValue(value);
              if (!isMobile) return label;
              return label.length > 22 ? `${label.slice(0, 22)}...` : label;
            },
          },
        },
      },
      layout: {
        padding: isMobile
          ? { left: 4, right: 10, top: 0, bottom: 0 }
          : { left: 8, right: 24, top: 0, bottom: 0 },
      },
      onClick(event, _active, chartInstance) {
        if (!isMobile || typeof onRowSelect !== "function") return;
        const indexMatches = chartInstance.getElementsAtEventForMode(
          event,
          "index",
          {
            axis: "y",
            intersect: false,
          },
          false
        );
        const rowIndex = indexMatches?.[0]?.index ?? -1;

        if (rowIndex < 0) return;
        const record = records[rowIndex];
        if (!record) return;
        onRowSelect(record);
      },
    },
  });
}

export function renderComparisonChart(canvas, chart, weeklyAverages, focus = "performance") {
  destroyIfExists(chart);
  const currentWeek = weeklyAverages.at(-1) || null;
  const previousWeek = weeklyAverages.length > 1 ? weeklyAverages[0] : null;
  const metrics = getFocusMetrics(focus);
  const kpiLabels = metrics.map((metric) => metric.label);
  const kpiKeys = metrics.map((metric) => metric.scoreKey);
  const colors = metrics.map((metric) => metric.color);

  return new Chart(canvas, {
    type: "bar",
    data: {
      labels: kpiLabels.map((label) => getMobileKpiLabel(label)),
      datasets: [
        {
          label: previousWeek?.weekEnding || "Previous Range",
          data: kpiKeys.map((key) => previousWeek?.[key] ?? null),
          backgroundColor: colors.map((color) => hexToRgba(color, 0.3)),
          borderColor: colors.map((color) => hexToRgba(color, 0.55)),
          borderWidth: 1.2,
          borderRadius: 12,
          categoryPercentage: 0.62,
          barPercentage: 0.92,
          maxBarThickness: 30,
        },
        {
          label: currentWeek?.weekEnding || "Selected Range",
          data: kpiKeys.map((key) => currentWeek?.[key] ?? null),
          backgroundColor: colors,
          borderColor: colors,
          borderWidth: 1.2,
          borderRadius: 12,
          categoryPercentage: 0.62,
          barPercentage: 0.92,
          maxBarThickness: 30,
        },
      ],
    },
    options: {
      ...buildBaseOptions(),
      plugins: {
        ...buildBaseOptions().plugins,
        valueLabelPlugin: {
          enabled: true,
          color: "#10213d",
          fontSize: 11,
          fontWeight: "800",
          offset: 10,
          formatter(value, context) {
            if (context.datasetIndex !== 1) return "";
            const previousValue = previousWeek?.[kpiKeys[context.dataIndex]];
            if (previousValue === null || previousValue === undefined || Number.isNaN(previousValue)) {
              return Number(value).toFixed(2);
            }
            const delta = Number(value) - Number(previousValue);
            const deltaPrefix = delta > 0 ? "+" : "";
            return `${deltaPrefix}${delta.toFixed(2)}`;
          },
        },
        tooltip: {
          ...buildBaseOptions().plugins.tooltip,
          callbacks: {
            title(items) {
              const index = items[0]?.dataIndex ?? -1;
              return kpiLabels[index] || "";
            },
            label(context) {
              const weekSummary = context.datasetIndex === 0 ? previousWeek : currentWeek;
              return [
                `${context.dataset.label}: ${Number(context.parsed.y).toFixed(2)}`,
                formatComparisonRaw(kpiLabels[context.dataIndex], weekSummary),
              ];
            },
            footer(items) {
              const currentItem = items.find((item) => item.datasetIndex === 1);
              const previousItem = items.find((item) => item.datasetIndex === 0);
              if (!currentItem || !previousItem) return "";
              const delta = Number(currentItem.parsed.y) - Number(previousItem.parsed.y);
              const prefix = delta > 0 ? "+" : "";
              return `Delta: ${prefix}${delta.toFixed(2)}`;
            },
          },
        },
      },
      scales: {
        y: {
          beginAtZero: true,
          max: 5,
          grid: {
            color: "rgba(16, 33, 61, 0.08)",
          },
          ticks: {
            stepSize: 1,
          },
        },
        x: {
          offset: true,
          stacked: false,
          grid: {
            display: false,
          },
        },
      },
    },
  });
}

function buildDistributionData(records, field) {
  const counts = [1, 2, 3, 4, 5].map((score) => ({
    score,
    count: records.filter((record) => Math.round(record[field]) === score).length,
    agents: records
      .filter((record) => Math.round(record[field]) === score)
      .map((record) => getChartDisplayName(record.agentName))
      .sort((left, right) => left.localeCompare(right)),
  }));
  const total = counts.reduce((sum, item) => sum + item.count, 0);
  return { counts, total };
}

function buildDistributionAgentTooltipLines(agents) {
  if (!agents.length) return ["Agents: none"];
  const lines = ["Agents:"];
  agents.forEach((agent) => {
    lines.push(`- ${agent}`);
  });
  return lines;
}

export function renderDistributionChart(canvas, chart, records, field, label, color) {
  destroyIfExists(chart);
  const { counts, total } = buildDistributionData(
    records.filter((record) => typeof record[field] === "number" && !Number.isNaN(record[field])),
    field
  );
  const maxCount = Math.max(...counts.map((item) => item.count), 0);
  const suggestedMax = Math.max(1, maxCount + Math.max(1, Math.ceil(maxCount * 0.18)));

  return new Chart(canvas, {
    type: "bar",
    data: {
      labels: counts.map((item) => `Score ${item.score}`),
      datasets: [
        {
          label,
          data: counts.map((item) => item.count),
          backgroundColor: counts.map((item) => (item.count ? color : "rgba(127, 143, 179, 0.18)")),
          borderColor: color,
          borderWidth: 1.2,
          borderRadius: 10,
          maxBarThickness: 28,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: {
        mode: "index",
        intersect: false,
      },
      plugins: {
        legend: {
          display: false,
        },
        valueLabelPlugin: {
          enabled: true,
          color: "#10213d",
          fontSize: 13,
          fontWeight: "800",
          offset: 12,
          formatter(value) {
            return String(Math.round(value));
          },
        },
        tooltip: {
          enabled: false,
          external: externalDistributionTooltip,
          callbacks: {
            title(items) {
              return items[0]?.label || "";
            },
            label(context) {
              const count = Number(context.parsed.y || 0);
              const percent = total ? ((count / total) * 100).toFixed(1) : "0.0";
              return [`Agents: ${count}`, `Share: ${percent}%`];
            },
            afterBody(items) {
              const index = items[0]?.dataIndex ?? -1;
              if (index < 0) return [];
              return buildDistributionAgentTooltipLines(counts[index]?.agents || []);
            },
            afterLabel() {
              return [];
            },
            footer() {
              return `Scored agents: ${total}`;
            },
          },
        },
      },
      scales: {
        x: {
          offset: true,
          grid: {
            display: false,
          },
          ticks: {
            color: "#6e7d98",
            font: {
              size: 11,
              weight: "700",
            },
          },
        },
        y: {
          beginAtZero: true,
          suggestedMax,
          grid: {
            display: false,
          },
          ticks: {
            display: false,
          },
          border: {
            display: false,
          },
        },
      },
    },
  });
}

export function renderAgentDeltaVarianceChart(canvas, chart, deltas, config = {}) {
  destroyIfExists(chart);
  const isMobile = isMobileViewport();
  const color = config.color || CHART_COLORS.overall;
  const sorted = [...deltas]
    .filter((item) => typeof item.delta === "number" && !Number.isNaN(item.delta))
    .sort((left, right) => right.delta - left.delta);

  return new Chart(canvas, {
    type: "bar",
    data: {
      labels: sorted.map((item) => getChartDisplayName(item.agentName)),
      datasets: [
        {
          label: config.metricLabel || "Variance",
          data: sorted.map((item) => item.delta),
          backgroundColor: sorted.map((item) => item.delta >= 0 ? color : hexToRgba(color, 0.28)),
          borderColor: sorted.map(() => color),
          borderWidth: 1.2,
          borderRadius: 10,
          maxBarThickness: 30,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        valueLabelPlugin: {
          enabled: true,
          color: "#10213d",
          fontSize: isMobile ? 10 : 11,
          fontWeight: "800",
          offset: isMobile ? 6 : 10,
          formatter(value) {
            const prefix = value > 0 ? "+" : "";
            return `${prefix}${Number(value).toFixed(2)}`;
          },
        },
        tooltip: {
          ...buildBaseOptions().plugins.tooltip,
          callbacks: {
            title(items) {
              return items[0]?.label || "";
            },
            label(context) {
              const item = sorted[context.dataIndex];
              const prefix = item.delta > 0 ? "+" : "";
              return [
                `${config.metricLabel || "Variance"}: ${prefix}${Number(item.delta).toFixed(2)}`,
                `${config.primaryLabel || "Primary"}: ${item.primaryValue === null || item.primaryValue === undefined ? "N/A" : Number(item.primaryValue).toFixed(2)}`,
                `${config.compareLabel || "Compare"}: ${item.compareValue === null || item.compareValue === undefined ? "N/A" : Number(item.compareValue).toFixed(2)}`,
              ];
            },
          },
        },
      },
      scales: {
        y: {
          beginAtZero: false,
          grid: { color: "rgba(16, 33, 61, 0.08)" },
          ticks: {
            callback(value) {
              const prefix = Number(value) > 0 ? "+" : "";
              return `${prefix}${Number(value).toFixed(1)}`;
            },
          },
        },
        x: {
          grid: { display: false },
          ticks: {
            font: { size: isMobile ? 10 : 12 },
            callback(value) {
              const label = this.getLabelForValue(value);
              if (!isMobile) return label;
              return label.length > 12 ? `${label.slice(0, 12)}...` : label;
            },
          },
        },
      },
      layout: {
        padding: {
          top: isMobile ? 14 : 4,
          right: 6,
        },
      },
    },
  });
}

export function renderRangeComparisonChart(canvas, chart, config) {
  destroyIfExists(chart);
  const isMobile = isMobileViewport();
  const color = config.color || CHART_COLORS.overall;
  const labels = [config.compareLabel || "Compare", config.primaryLabel || "Primary"];
  const values = [config.compareValue ?? null, config.primaryValue ?? null];

  return new Chart(canvas, {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: config.metricLabel || "Score",
          data: values,
          backgroundColor: [hexToRgba(color, 0.3), color],
          borderColor: [hexToRgba(color, 0.55), color],
          borderWidth: 1.2,
          borderRadius: 12,
          categoryPercentage: 0.62,
          barPercentage: 0.92,
          maxBarThickness: 42,
        },
      ],
    },
    options: {
      ...buildBaseOptions(),
      plugins: {
        ...buildBaseOptions().plugins,
        legend: { display: false },
        valueLabelPlugin: {
          enabled: true,
          color: "#10213d",
          fontSize: 11,
          fontWeight: "800",
          offset: 10,
          formatter(value, context) {
            if (context.dataIndex !== 1 || values[0] === null || values[0] === undefined) {
              return Number(value).toFixed(2);
            }
            const delta = Number(values[1]) - Number(values[0]);
            const prefix = delta > 0 ? "+" : "";
            return `${prefix}${delta.toFixed(2)}`;
          },
        },
        tooltip: {
          ...buildBaseOptions().plugins.tooltip,
          callbacks: {
            title(items) {
              return items[0]?.label || "";
            },
            label(context) {
              return `${config.metricLabel || "Score"}: ${Number(context.parsed.y).toFixed(2)}`;
            },
            footer() {
              if (values[0] === null || values[0] === undefined || values[1] === null || values[1] === undefined) {
                return "";
              }
              const delta = Number(values[1]) - Number(values[0]);
              const prefix = delta > 0 ? "+" : "";
              return `Delta: ${prefix}${delta.toFixed(2)}`;
            },
          },
        },
      },
      scales: {
        y: {
          beginAtZero: true,
          max: 5,
          grid: { color: "rgba(16, 33, 61, 0.08)" },
          ticks: { stepSize: 1 },
        },
        x: {
          grid: { display: false },
          ticks: {
            font: { size: isMobile ? 10 : 12 },
            callback(value) {
              const label = this.getLabelForValue(value);
              if (!isMobile) return label;
              return label.length > 14 ? `${label.slice(0, 14)}...` : label;
            },
          },
        },
      },
    },
  });
}

export function renderScoreSpreadChart(canvas, chart, records, options = {}) {
  destroyIfExists(chart);
  const isMobile = isMobileViewport();
  const scoreKey = options.scoreKey || "overallScore";
  const scoreLabel = options.scoreLabel || "Overall";
  const scored = records.filter((record) => typeof record[scoreKey] === "number" && !Number.isNaN(record[scoreKey]));
  const bands = [
    { label: "Below 2.0", color: "#d83b3b", predicate: (value) => value < 2, agents: [] },
    { label: "2.0 to 2.99", color: "#f28c28", predicate: (value) => value >= 2 && value < 3, agents: [] },
    { label: "3.0 to 3.99", color: "#e8b100", predicate: (value) => value >= 3 && value < 4, agents: [] },
    { label: "4.0 to 4.49", color: "#74c365", predicate: (value) => value >= 4 && value < 4.5, agents: [] },
    { label: "4.5 to 5.0", color: "#1f9d55", predicate: (value) => value >= 4.5, agents: [] },
  ];

  scored.forEach((record) => {
    const value = Number(record[scoreKey]);
    const band = bands.find((item) => item.predicate(value));
    if (band) {
      band.agents.push(getChartDisplayName(record.agentName));
    }
  });

  const maxCount = Math.max(...bands.map((item) => item.agents.length), 0);
  const suggestedMax = Math.max(1, maxCount + Math.max(1, Math.ceil(maxCount * 0.18)));

  return new Chart(canvas, {
    type: "bar",
    data: {
      labels: bands.map((item) => item.label),
      datasets: [
        {
          label: `${scoreLabel} Spread`,
          data: bands.map((item) => item.agents.length),
          backgroundColor: bands.map((item) => item.agents.length ? item.color : "rgba(127, 143, 179, 0.18)"),
          borderColor: bands.map((item) => item.color),
          borderWidth: 1.2,
          borderRadius: 10,
          maxBarThickness: 44,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: {
        mode: "index",
        intersect: false,
      },
      plugins: {
        legend: {
          display: false,
        },
        valueLabelPlugin: {
          enabled: true,
          color: "#10213d",
          fontSize: isMobile ? 14 : 15,
          fontWeight: "800",
          offset: 12,
          formatter(value) {
            return String(Math.round(value));
          },
        },
        tooltip: {
          enabled: false,
          external: externalDistributionTooltip,
          callbacks: {
            title(items) {
              return items[0]?.label || "";
            },
            label(context) {
              const count = Number(context.parsed.y || 0);
              const percent = scored.length ? ((count / scored.length) * 100).toFixed(1) : "0.0";
              return [`Agents: ${count}`, `Share: ${percent}%`];
            },
            afterBody(items) {
              const index = items[0]?.dataIndex ?? -1;
              if (index < 0) return [];
              return buildDistributionAgentTooltipLines((bands[index]?.agents || []).sort((left, right) => left.localeCompare(right)));
            },
            afterLabel() {
              return [];
            },
            footer() {
              return `Scored agents: ${scored.length}`;
            },
          },
        },
      },
      scales: {
        x: {
          offset: true,
          grid: {
            display: false,
          },
          ticks: {
            maxRotation: 0,
            minRotation: 0,
            padding: isMobile ? 10 : 6,
            color: "#6e7d98",
            font: {
              size: isMobile ? 11 : 12,
              weight: "700",
            },
            callback(_value, index) {
              if (!isMobile) return bands[index]?.label || "";
              const shortLabels = ["<2.0", "2-2.9", "3-3.9", "4-4.4", "4.5+"];
              return shortLabels[index] || bands[index]?.label || "";
            },
          },
        },
        y: {
          beginAtZero: true,
          suggestedMax,
          grid: {
            display: false,
          },
          ticks: {
            display: false,
          },
          border: {
            display: false,
          },
        },
      },
      layout: {
        padding: isMobile ? { left: 6, right: 6, top: 2, bottom: 18 } : { left: 8, right: 8, top: 4, bottom: 8 },
      },
    },
  });
}

export function renderTopBottomChart(canvas, chart, records, label, scoreKey = "overallScore") {
  destroyIfExists(chart);
  return new Chart(canvas, {
    type: "bar",
    data: {
      labels: records.map((record) => getChartDisplayName(record.agentName)),
      datasets: [
        {
          label,
          data: records.map((record) => record[scoreKey]),
          backgroundColor: records.map((record) => scoreColor(record[scoreKey])),
          borderRadius: 10,
        },
      ],
    },
    options: {
      ...buildBaseOptions(),
      interaction: {
        mode: "nearest",
        axis: "xy",
        intersect: true,
      },
      plugins: {
        ...buildBaseOptions().plugins,
        legend: {
          display: false,
        },
        strictChartAreaTooltipPlugin: {
          enabled: true,
        },
        valueLabelPlugin: {
          enabled: true,
          color: "#10213d",
          fontSize: 11,
          formatter(value) {
            return Number(value).toFixed(2);
          },
        },
        tooltip: {
          ...buildBaseOptions().plugins.tooltip,
          mode: "nearest",
          axis: "xy",
          intersect: true,
        },
      },
    },
  });
}

export function renderImprovementChart(canvas, chart, records, metricLabel = "Improvement") {
  destroyIfExists(chart);
  return new Chart(canvas, {
    type: "bar",
    data: {
      labels: records.map((record) => getChartDisplayName(record.agentName)),
      datasets: [
        {
          label: metricLabel,
          data: records.map((record) => record.delta),
          backgroundColor: records.map((record) =>
            record.delta >= 0 ? CHART_COLORS.admits : hexToRgba(CHART_COLORS.admits, 0.26)
          ),
          borderColor: CHART_COLORS.admits,
          borderWidth: 1.2,
          borderRadius: 10,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: {
        mode: "nearest",
        axis: "xy",
        intersect: true,
      },
      plugins: {
        legend: {
          display: false,
        },
        strictChartAreaTooltipPlugin: {
          enabled: true,
        },
        valueLabelPlugin: {
          enabled: true,
          color: "#10213d",
          fontSize: 11,
          fontWeight: "800",
          formatter(value) {
            const prefix = value > 0 ? "+" : "";
            return `${prefix}${Number(value).toFixed(2)}`;
          },
        },
        tooltip: {
          ...buildBaseOptions().plugins.tooltip,
          mode: "nearest",
          axis: "xy",
          intersect: true,
          callbacks: {
            title(items) {
              return items[0]?.label || "";
            },
            label(context) {
              const record = records[context.dataIndex];
              const delta = Number(context.parsed.y);
              const prefix = delta > 0 ? "+" : "";
              return [
                `${metricLabel}: ${prefix}${delta.toFixed(2)}`,
                `Current score: ${Number(record.currentOverallScore).toFixed(2)}`,
                `Previous score: ${Number(record.previousOverallScore).toFixed(2)}`,
              ];
            },
          },
        },
      },
      scales: {
        y: {
          beginAtZero: true,
          grid: {
            color: "rgba(16, 33, 61, 0.08)",
          },
        },
          x: {
            grid: {
              display: false,
            },
            ticks: {
              maxRotation: 0,
              minRotation: 0,
              autoSkip: false,
              font: {
                size: isMobile ? 10 : 12,
              },
            },
          },
        },
      },
    });
}
