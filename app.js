const STATUS_DEFS = [
  { code: "00", label: "SQL", color: "#8d99ae", score: 0 },
  { code: "01", label: "アポ獲得", color: "#3a86ff", score: 1 },
  { code: "02", label: "初回訪問完了", color: "#06d6a0", score: 2 },
  { code: "03", label: "商談中", color: "#8338ec", score: 3 },
  { code: "04", label: "見積提示", color: "#ff006e", score: 4 },
  { code: "5", label: "口頭内示", color: "#fb5607", score: 5 },
  { code: "6", label: "受注", color: "#2a9d8f", score: 6 },
  { code: "7", label: "請求済み", color: "#1d3557", score: 7 },
  { code: "80", label: "ペンディング", color: "#ffbe0b", score: 2.5 },
  { code: "99", label: "失注", color: "#e63946", score: -1 }
];

const STATUS_BY_CODE = new Map(STATUS_DEFS.map((s) => [s.code, s]));

const MONTHS = {
  Jan: 0,
  Feb: 1,
  Mar: 2,
  Apr: 3,
  May: 4,
  Jun: 5,
  Jul: 6,
  Aug: 7,
  Sep: 8,
  Oct: 9,
  Nov: 10,
  Dec: 11
};

const charts = {
  timeline: null
};

let projectTracking = {};
let projectChartInstance = null;
let allProjects = [];
let currentWeekColumns = [];
let modalDeals = [];
let modalSort = { key: "no", direction: "asc" };

function compareMaybeNumeric(a, b) {
  const numA = Number(a);
  const numB = Number(b);
  const aIsNum = Number.isFinite(numA);
  const bIsNum = Number.isFinite(numB);

  if (aIsNum && bIsNum) return numA - numB;
  return String(a).localeCompare(String(b), undefined, { numeric: true, sensitivity: "base" });
}

function getReadableTextColor(hexColor) {
  const hex = String(hexColor || "").replace("#", "").trim();
  if (hex.length !== 6) return "#ffffff";
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
  return luminance > 0.62 ? "#1f2937" : "#ffffff";
}

function setCanvasWidth(canvasElement, dataPointCount) {
  try {
    if (!canvasElement || !dataPointCount || dataPointCount < 1) return;
    
    const container = canvasElement.parentElement;
    if (!container) return;
    
    // Get container width
    let containerWidth = container.clientWidth;
    if (containerWidth <= 0) {
      containerWidth = Math.min(window.innerWidth - 80, 1168); // Account for padding
    }
    
    // Calculate needed width: ~60px per data point
    const minWidthPerPoint = 60;
    const calculatedWidth = dataPointCount * minWidthPerPoint;
    
    // Set canvas width to expand for scrolling if needed, but respect container bounds
    if (calculatedWidth > containerWidth) {
      // Canvas needs to expand - set explicit width for scrolling
      canvasElement.style.width = calculatedWidth + "px";
    } else {
      // Canvas fits in container - use container width
      canvasElement.style.width = "100%";
    }
  } catch (err) {
    console.error("Error in setCanvasWidth:", err);
  }
}

function hexToRgba(hex, alpha) {
  const clean = String(hex || "").replace("#", "").trim();
  if (clean.length !== 6) return `rgba(148, 163, 184, ${alpha})`;
  const r = parseInt(clean.slice(0, 2), 16);
  const g = parseInt(clean.slice(2, 4), 16);
  const b = parseInt(clean.slice(4, 6), 16);
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

function buildStatusBreakdown(deals) {
  const counts = Object.fromEntries(STATUS_DEFS.map((s) => [s.code, 0]));
  deals.forEach((deal) => {
    if (counts[deal.statusCode] !== undefined) counts[deal.statusCode] += 1;
  });

  return STATUS_DEFS
    .filter((s) => counts[s.code] > 0)
    .map((s) => `${s.code}-${s.label}: ${counts[s.code]}`)
    .join(" | ");
}

function sortModalDeals(deals, sortKey, direction) {
  const sorted = [...deals].sort((a, b) => {
    if (sortKey === "contract") {
      const aValue = parseContractValue(a.contract);
      const bValue = parseContractValue(b.contract);
      if (aValue === null && bValue === null) return 0;
      if (aValue === null) return 1;
      if (bValue === null) return -1;
      return aValue - bValue;
    }

    if (sortKey === "status") {
      return (a.statusScore ?? -99) - (b.statusScore ?? -99);
    }

    return compareMaybeNumeric(a[sortKey] ?? "", b[sortKey] ?? "");
  });

  if (direction === "desc") sorted.reverse();
  return sorted;
}

function updateModalSortHeader() {
  const headers = document.querySelectorAll(".deal-table thead th.sortable");
  headers.forEach((th) => {
    th.classList.remove("sorted-asc", "sorted-desc");
    if (th.dataset.sort === modalSort.key) {
      th.classList.add(modalSort.direction === "asc" ? "sorted-asc" : "sorted-desc");
    }
  });
}

function renderModalDeals() {
  const tbody = document.getElementById("modalTableBody");
  const modalCount = document.getElementById("modalCount");
  if (!tbody) return;

  const sortedDeals = sortModalDeals(modalDeals, modalSort.key, modalSort.direction);
  tbody.innerHTML = sortedDeals
    .map(
      (deal) => {
        const statusDef = STATUS_BY_CODE.get(deal.statusCode);
        const statusColor = statusDef?.color || "#94a3b8";
        const statusBg = hexToRgba(statusColor, 0.14);
        const statusBorder = hexToRgba(statusColor, 0.4);
        return `
        <tr>
          <td class="td-no">${deal.no}</td>
          <td>${deal.client}</td>
          <td>${deal.project}</td>
          <td>
            <span class="status-tag" style="background:${statusBg};border-color:${statusBorder};color:${statusColor};">
              ${deal.statusCode} - ${deal.statusLabel}
            </span>
          </td>
          <td class="td-contract">${deal.contract}</td>
        </tr>
      `;
      }
    )
    .join("");

  if (modalCount) {
    modalCount.textContent = `${sortedDeals.length} deals • Sorted by ${modalSort.key} (${modalSort.direction})`;
  }

  updateModalSortHeader();
}

function setupModalSorting() {
  const headers = document.querySelectorAll(".deal-table thead th.sortable");
  headers.forEach((th) => {
    th.addEventListener("click", () => {
      const sortKey = th.dataset.sort;
      if (!sortKey) return;

      if (modalSort.key === sortKey) {
        modalSort.direction = modalSort.direction === "asc" ? "desc" : "asc";
      } else {
        modalSort.key = sortKey;
        modalSort.direction = "asc";
      }

      renderModalDeals();
    });
  });
}

function setupSearchableDropdown() {
  const selected = document.getElementById("dropdownSelected");
  const menu = document.getElementById("dropdownMenu");
  const search = document.getElementById("dropdownSearch");
  const options = document.getElementById("dropdownOptions");

  if (!selected || !menu || !search || !options) return;

  // Populate dropdown options
  options.innerHTML = allProjects
    .map(
      (proj) => `
    <div class="dropdown-option" data-project="${proj}">
      ${proj}
    </div>
  `
    )
    .join("");

  // Toggle menu on selected click
  selected.addEventListener("click", (e) => {
    e.stopPropagation();
    menu.classList.toggle("open");
    if (menu.classList.contains("open")) {
      search.focus();
      search.value = "";
      updateVisibleOptions();
    }
  });

  // Filter options on search input
  search.addEventListener("input", () => {
    updateVisibleOptions();
  });

  // Handle option click
  options.addEventListener("click", (e) => {
    const optionEl = e.target.closest(".dropdown-option");
    if (optionEl && !optionEl.classList.contains("hidden")) {
      const projectName = optionEl.dataset.project;
      selectProject(projectName);
    }
  });

  // Close menu on click outside
  document.addEventListener("click", (e) => {
    if (!e.target.closest(".searchable-dropdown")) {
      menu.classList.remove("open");
    }
  });

  function updateVisibleOptions() {
    const query = search.value.toLowerCase().trim();
    document.querySelectorAll(".dropdown-option").forEach((opt) => {
      const projectName = opt.dataset.project;
      const shouldShow = projectName.toLowerCase().includes(query);
      opt.classList.toggle("hidden", !shouldShow);
    });
  }

  function selectProject(projectName) {
    selected.textContent = projectName;
    menu.classList.remove("open");
    buildProjectChart(currentWeekColumns, projectName);
  }
}

// Defer DOM queries until elements are available
let kpiGrid;
let statusLegend;
let statusText;
let csvInput;
let reloadBtn;

function initializeDOMElements() {
  kpiGrid = document.getElementById("kpiGrid");
  statusLegend = document.getElementById("statusLegend");
  statusText = document.getElementById("statusText");
  csvInput = document.getElementById("csvInput");
  reloadBtn = document.getElementById("reloadBtn");
  
  if (!kpiGrid || !statusLegend || !statusText || !csvInput || !reloadBtn) {
    console.error("Some required DOM elements not found");
  }
}

function cleanText(value) {
  return String(value ?? "")
    .replace(/\u200b|\ufeff/g, "")
    .trim();
}

function normalizeStatus(raw) {
  const cleaned = cleanText(raw);
  if (!cleaned) return null;

  const match = cleaned.match(/\d+/);
  if (!match) return null;

  const num = Number.parseInt(match[0], 10);
  if (Number.isNaN(num)) return null;

  if (num === 0) return "00";
  if (num >= 1 && num <= 4) return String(num).padStart(2, "0");
  return String(num);
}

function parseWeekLabel(label) {
  const cleaned = cleanText(label);
  
  // Try format: dd-MMM (e.g., 13-Apr)
  let m = cleaned.match(/^(\d{1,2})-([A-Za-z]{3})$/);
  if (m) {
    const day = Number.parseInt(m[1], 10);
    const monthIndex = MONTHS[m[2]];
    if (Number.isInteger(day) && monthIndex !== undefined) {
      return new Date(2026, monthIndex, day).getTime();
    }
  }
  
  // Try format: M/D or MM/DD (e.g., 4/13, 3/30)
  m = cleaned.match(/^(\d{1,2})\/(\d{1,2})$/);
  if (m) {
    const month = Number.parseInt(m[1], 10);
    const day = Number.parseInt(m[2], 10);
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return new Date(2026, month - 1, day).getTime();
    }
  }
  
  return null;
}

function formatPercent(value) {
  return `${(value * 100).toFixed(1)}%`;
}

function parseContractValue(raw) {
  const cleaned = cleanText(raw)
    .replace(/[\u200b\ufeff]/g, "")
    .replace(/,/g, "")
    .replace(/\/月/g, "")
    .replace(/円/g, "")
    .replace(/¥/g, "")
    .replace(/TBD/gi, "")
    .trim();
  const num = parseFloat(cleaned);
  return isNaN(num) || num <= 0 ? null : num;
}

function formatYen(value) {
  if (value === null || value === undefined) return null;
  if (value >= 1e8) return `¥${(value / 1e8).toFixed(1)}億`;
  if (value >= 1e4) return `¥${Math.round(value / 1e4)}万`;
  return `¥${value.toLocaleString()}`;
}

function weekColumnsFrom(headers) {
  return headers
    .map((h) => ({ raw: h, cleaned: cleanText(h), ts: parseWeekLabel(h) }))
    .filter((h) => h.ts !== null)
    .sort((a, b) => a.ts - b.ts)
    .map((h) => h.raw);
}

function aggregateByWeek(rows, weekColumns, headerMap) {
  dealMapping = {};
  
  return weekColumns.map((week) => {
    const counts = Object.fromEntries(STATUS_DEFS.map((s) => [s.code, 0]));
    const deals = Object.fromEntries(STATUS_DEFS.map((s) => [s.code, []]));
    let scored = 0;
    let scoredCount = 0;

    let weeklyValue = 0;
    let valueCount = 0;

    rows.forEach((row) => {
      const code = normalizeStatus(row[week]);
      if (!code || !STATUS_BY_CODE.has(code)) return;

      counts[code] += 1;
      scored += STATUS_BY_CODE.get(code).score;
      scoredCount += 1;

      const contractRaw = cleanText(row[headerMap.contract]);
      const contractNum = parseContractValue(contractRaw);
      if (contractNum !== null) {
        weeklyValue += contractNum;
        valueCount += 1;
      }

      const dealInfo = {
        no: cleanText(row[headerMap.no]) || "N/A",
        client: cleanText(row[headerMap.client]) || "N/A",
        project: cleanText(row[headerMap.project]) || "N/A",
        contract: contractRaw || "N/A",
        statusCode: code,
        statusLabel: STATUS_BY_CODE.get(code)?.label || "Unknown",
        statusScore: STATUS_BY_CODE.get(code)?.score ?? -99
      };
      deals[code].push(dealInfo);
    });
    
    const weekLabel = cleanText(week);
    STATUS_DEFS.forEach((status) => {
      dealMapping[`${weekLabel}|${status.code}`] = deals[status.code];
    });

    return {
      week: weekLabel,
      counts,
      avgScore: scoredCount ? scored / scoredCount : 0,
      totalTagged: scoredCount,
      weeklyValue: valueCount > 0 ? weeklyValue : null,
      valueCount
    };
  });
}

function findHeaderColumn(headers, keywords) {
  for (const header of headers) {
    const cleaned = cleanText(header).toLowerCase();
    if (keywords.some((kw) => cleaned.includes(kw.toLowerCase()))) {
      return header;
    }
  }
  return null;
}

function latestStatusForRow(row, weekColumns) {
  for (let i = weekColumns.length - 1; i >= 0; i -= 1) {
    const code = normalizeStatus(row[weekColumns[i]]);
    if (code && STATUS_BY_CODE.has(code)) return code;
  }
  return null;
}

function buildKpis(rows, weeklyData, latestWeekLabel, weekColumns) {
  const latest = weeklyData[weeklyData.length - 1];
  const latestCounts = latest ? latest.counts : {};

  const won = latestCounts["6"] || 0;
  const lost = latestCounts["99"] || 0;
  const decided = won + lost;
  const winRate = decided ? won / decided : 0;

  const withAnyStatus = rows.filter((row) => !!latestStatusForRow(row, weekColumns)).length;
  const pending = latestCounts["80"] || 0;

  // Total known contract value across all deals (latest week)
  const latestValue = latest ? latest.weeklyValue : null;
  const latestValueCount = latest ? latest.valueCount : 0;

  return [
    {
      title: "Total Projects",
      value: String(withAnyStatus),
      note: "Has at least 1 status in the weekly series"
    },
    {
      title: "Latest Week",
      value: latestWeekLabel || "N/A",
      note: `Tagged records: ${latest ? latest.totalTagged : 0}`
    },
    {
      title: "Pipeline Value",
      value: latestValue !== null ? formatYen(latestValue) : "N/A",
      note: `${latestValueCount} deals with known contract value`
    },
    {
      title: "Pending",
      value: String(pending),
        note: "Projects currently at status 80"
    }
  ];
}

function renderKpis(cards) {
  kpiGrid.innerHTML = cards
    .map(
      (card) => `
        <article class="kpi-card">
          <p class="kpi-title">${card.title}</p>
          <p class="kpi-value">${card.value}</p>
          <p class="kpi-note">${card.note}</p>
        </article>
      `
    )
    .join("");
}

function renderLegend() {
  const legendHtml = STATUS_DEFS.map(
    (s) => `
      <div class="legend-item">
        <span class="legend-dot" style="background:${s.color}"></span>
        <div class="legend-text">${s.code}: ${s.label}</div>
      </div>
    `
  ).join("");

  if (statusLegend) {
    statusLegend.innerHTML = legendHtml;
  }

  document.querySelectorAll(".status-legend-inline").forEach((el) => {
    el.innerHTML = legendHtml;
  });
}

function destroyCharts() {
  Object.values(charts).forEach((chart) => {
    if (chart) chart.destroy();
  });
  if (projectChartInstance) {
    projectChartInstance.destroy();
    projectChartInstance = null;
  }
}

Chart.register(ChartDataLabels);

// ─── Sales KPI Tracker ─────────────────────────────────────────────────────
const KPI_DEFS = [
  { key: "apo",   label: "アポ獲得数",   codes: ["01"],     target: 2, color: "#3a86ff", failColor: "#ef4444" },
  { key: "visit", label: "顧客訪問数",   codes: ["02"],     target: 3, color: "#8b5cf6" },
  { key: "won",   label: "案件獲得件数", codes: ["6", "7"], target: 1, color: "#f97316", lineColor: "#ea580c", failColor: "#b91c1c" },
];

let kpiChartInstance = null;

function calcKpiWeekly(weeklyData) {
  return weeklyData.map((w) => {
    const result = { week: w.week };
    KPI_DEFS.forEach((kpi) => {
      const actual = kpi.codes.reduce((sum, code) => sum + (w.counts[code] || 0), 0);
      result[kpi.key] = { actual, target: kpi.target, rate: kpi.target > 0 ? actual / kpi.target : null };
    });
    return result;
  });
}

function buildKpiChart(weeklyData) {
  if (kpiChartInstance) kpiChartInstance.destroy();

  const kpiWeekly = calcKpiWeekly(weeklyData);
  const labels = kpiWeekly.map((w) => w.week);
  const weeklyStatus = kpiWeekly.map((week) => {
    const metCount = KPI_DEFS.reduce((count, kpi) => {
      return count + (week[kpi.key].actual >= kpi.target ? 1 : 0);
    }, 0);
    return {
      passed: metCount === KPI_DEFS.length,
      metCount
    };
  });

  const datasets = [
    {
      type: "bar",
      label: "週次KPI判定",
      data: weeklyStatus.map(() => 1),
      backgroundColor: weeklyStatus.map((s) => (s.passed ? "#86efac" : "#ef4444")),
      hoverBackgroundColor: weeklyStatus.map((s) => (s.passed ? "#4ade80" : "#dc2626")),
      borderRadius: 4,
      borderSkipped: false,
      borderWidth: 0,
      categoryPercentage: 0.9,
      barPercentage: 0.9,
      maxBarThickness: 24,
      datalabels: { display: false }
    }
  ];

  const kpiCanvas = document.getElementById("kpiChart");
  setCanvasWidth(kpiCanvas, labels.length);

  kpiChartInstance = new Chart(kpiCanvas, {
    data: { labels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          display: true,
          position: "bottom",
          labels: { font: { size: 11 }, padding: 12, boxWidth: 14 }
        },
        tooltip: {
          callbacks: {
            title: (items) => {
              const week = items?.[0]?.label || "";
              return `Week ${week}`;
            },
            label: (item) => {
              const status = weeklyStatus[item.dataIndex];
              return `総合判定: ${status.passed ? "達成" : "未達"} (${status.metCount}/${KPI_DEFS.length})`;
            },
            afterLabel: (item) => {
              const week = kpiWeekly[item.dataIndex];
              return KPI_DEFS.map((kpi) => {
                const actual = week[kpi.key].actual;
                const pass = actual >= kpi.target ? "達成" : "未達";
                return `${kpi.label}: ${actual}/${kpi.target} ${pass}`;
              });
            }
          }
        }
      },
      scales: {
        x: {
          ticks: { maxRotation: 0, autoSkip: true, maxTicksLimit: 16 },
          grid: { display: false }
        },
        y: {
          beginAtZero: true,
          max: 1,
          ticks: { stepSize: 1, callback: () => "" },
          grid: { display: false },
          border: { display: false }
        }
      }
    }
  });

  renderKpiSummary(kpiWeekly);
  renderKpiProgressTable(kpiWeekly);
}

function rateClass(rate) {
  if (rate === null) return "";
  if (rate >= 1) return "rate-good";
  if (rate >= 0.7) return "rate-warn";
  return "rate-bad";
}

function renderKpiSummary(kpiWeekly) {
  const summary = document.getElementById("kpiTrackerSummary");
  if (!summary) return;
  summary.innerHTML = KPI_DEFS.map((kpi) => {
    let cumActual = 0, cumTarget = 0;
    kpiWeekly.forEach((w) => { cumActual += w[kpi.key].actual; cumTarget += kpi.target; });
    const cumRate = cumTarget > 0 ? cumActual / cumTarget : null;
    const pct = cumRate !== null ? Math.round(cumRate * 100) : null;
    const cls = rateClass(cumRate);
    return `
      <div class="kpi-summary-card">
        <div class="kpi-name">${kpi.label}</div>
        <div class="kpi-values">${cumActual} <span>/ ${cumTarget} 計画</span></div>
        <div class="kpi-rate ${cls}">${pct !== null ? `累積 ${pct}%` : "N/A"}</div>
      </div>`;
  }).join("");
}

function renderKpiProgressTable(kpiWeekly) {
  const table = document.getElementById("kpiProgressTable");
  if (!table) return;
  const weekCols = kpiWeekly.map((w) => `<th>${w.week}</th>`).join("");
  const rows = KPI_DEFS.map((kpi) => {
    let cumActual = 0, cumTarget = 0;
    const targetCells = kpiWeekly.map(() => `<td>${kpi.target}</td>`).join("");
    const actualCells = kpiWeekly.map((w) => `<td>${w[kpi.key].actual}</td>`).join("");
    const weekRateCells = kpiWeekly.map((w) => {
      const r = w[kpi.key].rate;
      const pct = r !== null ? Math.round(r * 100) : null;
      return `<td class="${rateClass(r)}">${pct !== null ? pct + "%" : "-"}</td>`;
    }).join("");
    const cumRateCells = kpiWeekly.map((w) => {
      cumActual += w[kpi.key].actual;
      cumTarget += kpi.target;
      const r = cumTarget > 0 ? cumActual / cumTarget : null;
      const pct = r !== null ? Math.round(r * 100) : null;
      return `<td class="${rateClass(r)}">${pct !== null ? pct + "%" : "-"}</td>`;
    }).join("");
    return `
      <tr class="section-header"><td colspan="${kpiWeekly.length + 1}">${kpi.label}</td></tr>
      <tr><td>計画</td>${targetCells}</tr>
      <tr><td>実績</td>${actualCells}</tr>
      <tr><td>週次進捗率</td>${weekRateCells}</tr>
      <tr><td>累積進捗率</td>${cumRateCells}</tr>`;
  }).join("");
  table.innerHTML = `
    <thead><tr><th>KPI</th>${weekCols}</tr></thead>
    <tbody>${rows}</tbody>`;
}
// ─── End Sales KPI Tracker ──────────────────────────────────────────────────

function buildCharts(weeklyData) {
  destroyCharts();

  const labels = weeklyData.map((w) => w.week);
  const totals = weeklyData.map((w) => w.totalTagged);
  const scoreSeries = weeklyData.map((w) => Number(w.avgScore.toFixed(2)));
  const valueSeries = weeklyData.map((w) => w.weeklyValue);
  const valueCountSeries = weeklyData.map((w) => w.valueCount);

  const timelineCanvas = document.getElementById("timelineChart");
  setCanvasWidth(timelineCanvas, labels.length);

  charts.timeline = new Chart(timelineCanvas, {
    data: {
      labels,
      datasets: [
        ...STATUS_DEFS.map((status) => ({
          type: "bar",
          label: `${status.code} ${status.label}`,
          data: weeklyData.map((w) => w.counts[status.code] || 0),
          backgroundColor: status.color,
          borderRadius: 4,
          borderWidth: 0,
          stack: "status",
          datalabels: {
            display: (ctx) => ctx.dataset.data[ctx.dataIndex] > 0,
            formatter: (value) => `${value}`,
            color: getReadableTextColor(status.color),
            font: { size: 9, weight: "700" },
            align: "center",
            anchor: "center",
            textAlign: "center",
            clip: true
          }
        })),
        {
          type: "line",
          label: "Average Stage Score",
          data: scoreSeries,
          yAxisID: "y1",
          borderColor: "#0b5e57",
          backgroundColor: "#0b5e57",
          pointRadius: 5,
          pointHoverRadius: 7,
          borderWidth: 2,
          tension: 0.28,
          datalabels: {
            display: true,
            align: "top",
            anchor: "end",
            offset: 4,
            color: "#0b5e57",
            font: { size: 11, weight: "700" },
            formatter: (v) => (v === null || v === undefined ? "" : v.toFixed(1)),
            backgroundColor: "rgba(255,255,255,0.82)",
            borderRadius: 4,
            padding: { top: 2, bottom: 2, left: 4, right: 4 }
          }
        },
        {
          type: "line",
          label: "Pipeline Value (¥)",
          data: valueSeries,
          yAxisID: "y2",
          borderColor: "#e76f51",
          backgroundColor: "rgba(231,111,81,0.12)",
          pointRadius: 5,
          pointHoverRadius: 7,
          borderWidth: 2.5,
          borderDash: [6, 3],
          tension: 0.28,
          spanGaps: true,
          fill: false,
          datalabels: {
            display: (ctx) => valueSeries[ctx.dataIndex] !== null,
            align: "bottom",
            anchor: "end",
            offset: 4,
            color: "#c1440e",
            font: { size: 10, weight: "700" },
            formatter: (v) => (v !== null ? formatYen(v) : ""),
            backgroundColor: "rgba(255,255,255,0.85)",
            borderRadius: 4,
            padding: { top: 2, bottom: 2, left: 4, right: 4 }
          }
        },
        {
          type: "line",
          label: "_Total Projects",
          data: totals,
          yAxisID: "y",
          borderColor: "rgba(0,0,0,0)",
          backgroundColor: "rgba(0,0,0,0)",
          pointRadius: 0,
          pointHoverRadius: 0,
          borderWidth: 0,
          showLine: false,
          datalabels: {
            display: true,
            align: "top",
            anchor: "end",
            offset: -2,
            color: "#3c4257",
            font: { size: 12, weight: "700" },
            formatter: (v) => v
          }
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          display: true,
          labels: {
            filter: (item) => item.text === "Average Stage Score" || item.text === "Pipeline Value (¥)",
            usePointStyle: true,
            pointStyle: "line",
            font: { size: 12 },
            padding: 16,
          },
          onHover: (event, legendItem) => {
            const native = event.native;
            if (!native) return;
            const tooltipEl = document.getElementById("legendTooltip");
            const descriptions = {
              "Average Stage Score":
                "📈 Average Stage Score\n\nWeighted average of stage scores across all deals in the week.\n\nFormula:\n  Σ(stage score per deal) ÷ total deals\n\nScale:  Min −1  →  Max 7\n  < 1  🔴 Early stage / low progress\n  1–3  🟡 Pipeline developing\n  3–5  🟢 Advancing well\n  > 5  ✅ Near close / strong pipeline\n\nStage scores:\n  00 SQL = 0\n  01 アポ獲得 = 1\n  02 初回訪問完了 = 2\n  03 商談中 = 3\n  04 見積提示 = 4\n  5 口頭内示 = 5\n  6 受注 = 6\n  7 請求済み = 7\n  80 ペンディング = 2.5\n  99 失注 = −1",
              "Pipeline Value (¥)":
                "💴 Pipeline Value (¥)\n\nTotal contract value (¥) of all deals with an entered amount in that week.",
            };
            tooltipEl.textContent = descriptions[legendItem.text] || legendItem.text;
            tooltipEl.style.display = "block";
            tooltipEl.style.left = native.clientX + 14 + "px";
            tooltipEl.style.top = native.clientY - 10 + "px";
          },
          onLeave: () => {
            const tooltipEl = document.getElementById("legendTooltip");
            if (tooltipEl) tooltipEl.style.display = "none";
          },
        },
        tooltip: {
          backgroundColor: "rgba(0, 0, 0, 0.92)",
          padding: 14,
          titleFont: { size: 13, weight: "bold" },
          bodyFont: { size: 11, family: "monospace" },
          footerFont: { size: 10 },
          borderColor: "#555",
          borderWidth: 1,
          displayColors: true,
          callbacks: {
            title: (items) => {
              if (!items.length) return "";
              const item = items[0];
              return `${labels[item.dataIndex]}`;
            },
            label: (item) => {
              if (item.dataset.label === "Average Stage Score") {
                return `📈 Stage Score: ${item.formattedValue}`;
              }
              if (item.dataset.label === "Pipeline Value (¥)") {
                const v = valueSeries[item.dataIndex];
                const cnt = valueCountSeries[item.dataIndex];
                return v !== null
                  ? `💴 Pipeline Value: ${formatYen(v)} (${cnt} deals)`
                  : `💴 Pipeline Value: N/A`;
              }
              if (item.dataset.label === "_Total Projects") {
                return `👥 Total Projects: ${item.formattedValue}`;
              }
              const status = STATUS_DEFS[item.datasetIndex];
              return `${status.code} - ${status.label}: ${item.formattedValue} deals`;
            },
            footer: (items) => {
              if (!items.length) return "";
              const index = items[0].dataIndex;
              return `👥 Total: ${totals[index]} projects`;
            }
          }
        }
      },
      scales: {
        x: {
          stacked: true,
          ticks: {
            maxRotation: 0,
            autoSkip: true
          }
        },
        y: {
          stacked: true,
          beginAtZero: true,
          ticks: { precision: 0 },
          title: {
            display: true,
            text: "Number of Projects"
          }
        },
        y1: {
          position: "right",
          beginAtZero: false,
          min: -1,
          max: 7,
          grid: { drawOnChartArea: false },
          title: {
            display: true,
            text: "Avg Stage Score"
          }
        },
        y2: {
          position: "right",
          beginAtZero: true,
          grid: { drawOnChartArea: false },
          ticks: {
            callback: (v) => formatYen(v)
          },
          title: {
            display: true,
            text: "Pipeline Value"
          }
        }
      }
    }
  });

  // Click bar → open modal with all deals
  timelineCanvas.onclick = (evt) => {
    const chart = charts.timeline;
    if (!chart) return;

    const points = chart.getElementsAtEventForMode(evt, "index", { intersect: true }, false);
    if (!points.length) return;

    const barPoint = points.find((p) => p.datasetIndex < STATUS_DEFS.length);
    if (!barPoint) return;

    const { index } = barPoint;
    const week = labels[index];

    const allDeals = STATUS_DEFS.flatMap((status) => {
      const key = `${week}|${status.code}`;
      return dealMapping[key] || [];
    });

    if (!allDeals.length) return;
    openDealModal(week, { code: "ALL", label: "All Statuses" }, allDeals);
  };
}

function buildProjectChart(weekColumns, projectName) {
  if (projectChartInstance) projectChartInstance.destroy();

  const statusHistory = projectTracking[projectName];
  if (!statusHistory || !statusHistory.length) return;

  const labels = weekColumns.map((w) => cleanText(w));
  const datasets = [];
  const statusCounts = {};

  statusHistory.forEach((statuses) => {
    statuses.forEach((status, weekIdx) => {
      if (status) {
        const key = `${weekIdx}|${status}`;
        statusCounts[key] = (statusCounts[key] || 0) + 1;
      }
    });
  });

  STATUS_DEFS.forEach((statusDef) => {
    const data = [];
    for (let i = 0; i < weekColumns.length; i++) {
      const key = `${i}|${statusDef.code}`;
      data.push(statusCounts[key] || 0);
    }
    datasets.push({
      label: `${statusDef.code} - ${statusDef.label}`,
      data,
      backgroundColor: statusDef.color,
      borderRadius: 4,
      borderWidth: 0,
      stack: "status"
    });
  });

  const projectCanvas = document.getElementById("projectChart");
  setCanvasWidth(projectCanvas, labels.length);

  projectChartInstance = new Chart(projectCanvas, {
    type: "bar",
    data: { labels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        datalabels: {
          display: (ctx) => ctx.dataset.data[ctx.dataIndex] > 0,
          formatter: (value, ctx) => {
            const status = STATUS_DEFS[ctx.datasetIndex];
            return `${status.code} ${status.label}`;
          },
          color: "#fff",
          font: { size: 11, weight: "bold" },
          anchor: "center",
          align: "center",
          overflow: "hidden",
          clip: true,
        },
        title: {
          display: true,
          text: `Project: ${projectName}`
        },
        tooltip: {
          callbacks: {
            label: (item) => {
              const status = STATUS_DEFS[item.datasetIndex];
              return `${status.code} - ${status.label}: ${item.formattedValue}`;
            }
          }
        }
      },
      scales: {
        x: {
          stacked: true,
          ticks: { maxRotation: 0, autoSkip: true }
        },
        y: {
          stacked: true,
          beginAtZero: true,
          ticks: { precision: 0 },
          title: { display: true, text: "Count" }
        }
      }
    }
  });

  document.getElementById("projectChartWrap").hidden = false;
  document.getElementById("projectHint").hidden = true;
}

function openDealModal(week, status, deals) {
  const modal = document.getElementById("dealModal");
  const eyebrow = document.getElementById("modalWeek");
  const breakdown = buildStatusBreakdown(deals);

  modalDeals = [...deals];
  modalSort = { key: "no", direction: "asc" };
  renderModalDeals();

  const title = document.getElementById("modalTitle");
  title.textContent = status.code === "ALL"
    ? `${week} • All Statuses (${deals.length} deals)`
    : `${week} • ${status.code} - ${status.label} (${deals.length} deals)`;
  if (eyebrow) eyebrow.textContent = `${deals.length} deals • ${breakdown}`;
  modal.hidden = false;
}

function closeModal() {
  document.getElementById("dealModal").hidden = true;
}

function processRows(rows) {
  if (!rows.length) {
    statusText.textContent = "No valid data rows found in CSV.";
    return;
  }

  const headers = Object.keys(rows[0]);
  const weekColumns = weekColumnsFrom(headers);

  if (!weekColumns.length) {
    statusText.textContent = "No week columns found with format dd-MMM or M/D (e.g. 13-Apr or 4/13).";
    return;
  }

  const headerMap = {
    no: findHeaderColumn(headers, ["no", "no."]) || headers[1],
    client: findHeaderColumn(headers, ["client"]) || headers[2],
    project: findHeaderColumn(headers, ["project", "project name"]) || headers[3],
    contract: findHeaderColumn(headers, ["contract", "value", "total"]) || headers[4]
  };

  const firstHeader = cleanText(Object.keys(rows[0])[0]).toLowerCase();
  const usefulRows = rows.filter((row) => {
    // Skip rows that repeat the header (first cell matches first header key)
    const firstCell = cleanText(Object.values(row)[0]).toLowerCase();
    if (firstCell === firstHeader) return false;
    return weekColumns.some((week) => {
      const code = normalizeStatus(row[week]);
      return code && STATUS_BY_CODE.has(code);
    });
  });

  // Build project tracking history
  projectTracking = {};
  usefulRows.forEach((row) => {
    const projectName = cleanText(row[headerMap.project]) || "Unknown";
    if (!projectTracking[projectName]) {
      projectTracking[projectName] = [];
    }
    const statuses = weekColumns.map((week) => {
      const code = normalizeStatus(row[week]);
      return code && STATUS_BY_CODE.has(code) ? code : null;
    });
    projectTracking[projectName].push(statuses);
  });

  const weeklyData = aggregateByWeek(usefulRows, weekColumns, headerMap);
  const latestWeekLabel = weeklyData.length ? weeklyData[weeklyData.length - 1].week : "N/A";

  renderKpis(buildKpis(usefulRows, weeklyData, latestWeekLabel, weekColumns));
  renderLegend();
  buildCharts(weeklyData);
  buildKpiChart(weeklyData);

  // Show dashboard sections after successful load
  document.getElementById("kpiGrid").hidden = false;
  document.querySelector(".timeline-panel").hidden = false;
  document.getElementById("legendFooter").hidden = true;
  document.querySelector(".project-panel").hidden = false;
  document.getElementById("kpiTrackerPanel").hidden = false;

  // Store weekColumns globally for chart building
  currentWeekColumns = weekColumns;

  // Populate project dropdown with searchable options
  allProjects = Object.keys(projectTracking).sort();
  
  // Setup searchable dropdown
  setupSearchableDropdown();

    statusText.textContent = `Loaded ${usefulRows.length} projects from ${weekColumns.length} weeks.`;
}

function parseCsvText(text) {
  Papa.parse(text, {
    header: true,
    skipEmptyLines: "greedy",
    complete: (result) => {
      try {
        processRows(result.data || []);
      } catch (err) {
        console.error("Error processing CSV rows:", err);
        statusText.textContent = "Error processing data: " + err.message;
      }
    },
    error: (err) => {
      console.error("CSV parse error:", err);
      statusText.textContent = "Failed to parse CSV. Please check the file format.";
    }
  });
}

function setupCSVEventListeners() {
  if (!csvInput || !reloadBtn) return;
  
  csvInput.addEventListener("change", (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      const text = event.target?.result;
      if (typeof text === "string") parseCsvText(text);
    };
    reader.readAsText(file);
  });

  reloadBtn.addEventListener("click", () => {
    csvInput.value = "";
    destroyCharts();
    document.getElementById("kpiGrid").hidden = true;
    document.querySelector(".timeline-panel").hidden = true;
    document.getElementById("legendFooter").hidden = true;
    document.querySelector(".project-panel").hidden = true;
    document.getElementById("kpiTrackerPanel").hidden = true;
    statusText.textContent = "Awaiting CSV file...";
  });
}

document.addEventListener("DOMContentLoaded", initializeApp);
window.addEventListener("load", initializeApp);

function initializeApp() {
  // Only run once
  if (window._appInitialized) return;
  window._appInitialized = true;
  
  initializeDOMElements();
  setupCSVEventListeners();
  
  if (statusText) {
    statusText.textContent = "Awaiting CSV file...";
  }
  
  try {
    setupModalSorting();
  } catch (err) {
    console.error("Error in setupModalSorting:", err);
  }
  
  try {
    const modalClose = document.getElementById("modalClose");
    const dealModal = document.getElementById("dealModal");
    if (modalClose) {
      modalClose.addEventListener("click", closeModal);
    }
    if (dealModal) {
      dealModal.addEventListener("click", (e) => {
        if (e.target.id === "dealModal") closeModal();
      });
    }
  } catch (err) {
    console.error("Error setting up modal listeners:", err);
  }
}
