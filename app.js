const STATUS_DEFS = [
  { code: "00", label: "SQL", color: "#8d99ae", score: 0 },
  { code: "01", label: "アポ獲得", color: "#3a86ff", score: 1 },
  { code: "02", label: "初回訪問完了", color: "#4361ee", score: 2 },
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
let dealMapping = {};

const kpiGrid = document.getElementById("kpiGrid");
const statusLegend = document.getElementById("statusLegend");
const statusText = document.getElementById("statusText");
const csvInput = document.getElementById("csvInput");
const reloadBtn = document.getElementById("reloadBtn");

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
        contract: contractRaw || "N/A"
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
      title: "Total Opportunities",
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
      note: "Opportunities currently at status 80"
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
  statusLegend.innerHTML = STATUS_DEFS.map(
    (s) => `
      <div class="legend-item">
        <span class="legend-dot" style="background:${s.color}"></span>
        <div class="legend-text">${s.code}: ${s.label}</div>
      </div>
    `
  ).join("");
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

function buildCharts(weeklyData) {
  destroyCharts();

  const labels = weeklyData.map((w) => w.week);
  const totals = weeklyData.map((w) => w.totalTagged);
  const scoreSeries = weeklyData.map((w) => Number(w.avgScore.toFixed(2)));
  const valueSeries = weeklyData.map((w) => w.weeklyValue);
  const valueCountSeries = weeklyData.map((w) => w.valueCount);

  charts.timeline = new Chart(document.getElementById("timelineChart"), {
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
          datalabels: { display: false }
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
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: "bottom" },
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
              if (item.datasetIndex === STATUS_DEFS.length) {
                return `📈 Stage Score: ${item.formattedValue}`;
              }
              if (item.datasetIndex === STATUS_DEFS.length + 1) {
                const v = valueSeries[item.dataIndex];
                const cnt = valueCountSeries[item.dataIndex];
                return v !== null
                  ? `💴 Pipeline Value: ${formatYen(v)} (${cnt} deals)`
                  : `💴 Pipeline Value: N/A`;
              }
              const status = STATUS_DEFS[item.datasetIndex];
              return `${status.code} - ${status.label}: ${item.formattedValue} deals`;
            },
            afterLabel: (item) => {
              if (item.datasetIndex >= STATUS_DEFS.length) return "";
              
              const week = labels[item.dataIndex];
              const status = STATUS_DEFS[item.datasetIndex];
              const key = `${week}|${status.code}`;
              const deals = dealMapping[key] || [];
              
              if (!deals.length) return "";
              
              let result = "\n" + "─".repeat(35) + "\nDeals:";
              deals.forEach((deal, idx) => {
                if (idx < 5) {
                  result += `\n\n[${idx + 1}] No: ${deal.no}`;
                  result += `\n    Client: ${deal.client}`;
                  result += `\n    Project: ${deal.project}`;
                  result += `\n    Value: ${deal.contract}`;
                }
              });
              if (deals.length > 5) {
                result += `\n\n... và ${deals.length - 5} deal khác`;
                result += "\n💡 Click bar để xem tất cả";
              }
              result += "\n" + "─".repeat(35);
              return result;
            },
            footer: (items) => {
              if (!items.length) return "";
              const index = items[0].dataIndex;
              return `👥 Total: ${totals[index]} opportunities`;
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
            text: "Number of Opportunities"
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
  document.getElementById("timelineChart").addEventListener("click", (evt) => {
    const chart = charts.timeline;
    if (!chart) return;
    const points = chart.getElementsAtEventForMode(evt, "nearest", { intersect: true }, false);
    if (!points.length) return;
    const { datasetIndex, index } = points[0];
    if (datasetIndex >= STATUS_DEFS.length) return;
    const status = STATUS_DEFS[datasetIndex];
    const week = labels[index];
    const key = `${week}|${status.code}`;
    const deals = dealMapping[key] || [];
    if (!deals.length) return;
    openDealModal(week, status, deals);
  });
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

  projectChartInstance = new Chart(document.getElementById("projectChart"), {
    type: "bar",
    data: { labels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: "bottom" },
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
  const tbody = document.getElementById("modalTableBody");
  
  tbody.innerHTML = deals
    .map(
      (deal) => `
        <tr>
          <td>${deal.no}</td>
          <td>${deal.client}</td>
          <td>${deal.project}</td>
          <td>${deal.contract}</td>
        </tr>
      `
    )
    .join("");

  const title = document.getElementById("modalTitle");
  title.textContent = `${week} • ${status.code} - ${status.label} (${deals.length} deals)`;
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

  const usefulRows = rows.filter((row) =>
    weekColumns.some((week) => {
      const code = normalizeStatus(row[week]);
      return code && STATUS_BY_CODE.has(code);
    })
  );

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

  // Show dashboard sections after successful load
  document.getElementById("kpiGrid").hidden = false;
  document.querySelector(".timeline-panel").hidden = false;
  document.querySelector(".legend-panel").hidden = false;
  document.querySelector(".project-panel").hidden = false;

  // Populate project selector
  const projectSelect = document.getElementById("projectSelect");
  projectSelect.innerHTML = '<option value="">-- Choose a project --</option>';
  Object.keys(projectTracking)
    .sort()
    .forEach((proj) => {
      const opt = document.createElement("option");
      opt.value = proj;
      opt.textContent = proj;
      projectSelect.appendChild(opt);
    });

  projectSelect.addEventListener("change", (e) => {
    if (e.target.value) buildProjectChart(weekColumns, e.target.value);
    else {
      document.getElementById("projectChartWrap").hidden = true;
      document.getElementById("projectHint").hidden = false;
    }
  });

  statusText.textContent = `Loaded ${usefulRows.length} opportunities from ${weekColumns.length} weeks.`;
}

function parseCsvText(text) {
  Papa.parse(text, {
    header: true,
    skipEmptyLines: "greedy",
    complete: (result) => {
      processRows(result.data || []);
    },
    error: () => {
      statusText.textContent = "Failed to parse CSV. Please check the file format.";
    }
  });
}

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
  document.querySelector(".legend-panel").hidden = true;
  document.querySelector(".project-panel").hidden = true;
  statusText.textContent = "Awaiting CSV file...";
});

window.addEventListener("load", () => {
  statusText.textContent = "Awaiting CSV file...";
});

document.getElementById("modalClose").addEventListener("click", closeModal);
document.getElementById("dealModal").addEventListener("click", (e) => {
  if (e.target.id === "dealModal") closeModal();
});
