const dropZone = document.getElementById("drop-zone");
const fileInput = document.getElementById("file-input");
const summarySection = document.getElementById("summary");
const filterSection = document.getElementById("filters");
const insightsSection = document.getElementById("insights");
const columnsSection = document.getElementById("columns");
const chartsSection = document.getElementById("charts");
const previewSection = document.getElementById("preview");
const businessChartsSection = document.getElementById("business-charts");
const businessFilters = document.getElementById("business-filters");
const summaryCards = document.getElementById("summary-cards");
const columnsBody = document.getElementById("columns-body");
const previewHead = document.getElementById("preview-head");
const previewBody = document.getElementById("preview-body");
const filterColumn = document.getElementById("filter-column");
const filterQuery = document.getElementById("filter-query");
const clearFilters = document.getElementById("clear-filters");
const filterResult = document.getElementById("filter-result");
const chartColumn = document.getElementById("chart-column");
const chartCanvas = document.getElementById("chart-canvas");
const warehouseFilter = document.getElementById("warehouse-filter");
const billableFilter = document.getElementById("billable-filter");
const warehouseCanvas = document.getElementById("warehouse-canvas");
const billableCanvas = document.getElementById("billable-canvas");
const completenessCanvas = document.getElementById("completeness-canvas");
const typesCanvas = document.getElementById("types-canvas");

let originalRows = [];
let filteredRows = [];
let headers = [];
let businessColumns = {
  warehouse: null,
  billable: null,
};

dropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropZone.classList.add("dragover");
});

dropZone.addEventListener("dragleave", () => {
  dropZone.classList.remove("dragover");
});

dropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropZone.classList.remove("dragover");
  const [file] = event.dataTransfer.files;
  if (file) {
    processFile(file);
  }
});

fileInput.addEventListener("change", (event) => {
  const [file] = event.target.files;
  if (file) {
    processFile(file);
  }
});

filterColumn.addEventListener("change", applyFilters);
filterQuery.addEventListener("input", applyFilters);
warehouseFilter.addEventListener("change", applyFilters);
billableFilter.addEventListener("change", applyFilters);
clearFilters.addEventListener("click", () => {
  filterColumn.value = "__all";
  filterQuery.value = "";
  warehouseFilter.value = "all";
  billableFilter.value = "all";
  applyFilters();
});
chartColumn.addEventListener("change", () => {
  drawQuickChart(filteredRows, chartColumn.value);
});

async function processFile(file) {
  try {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    originalRows = XLSX.utils.sheet_to_json(worksheet, { defval: null });

    if (!originalRows.length) {
      alert("El archivo no contiene filas con datos en la primera hoja.");
      return;
    }

    headers = Object.keys(originalRows[0]);
    filteredRows = [...originalRows];
    businessColumns = detectBusinessColumns(headers);

    renderSummary(file.name, firstSheetName, originalRows.length, headers.length);
    buildSelectors(headers);
    buildBusinessSelectors(originalRows);
    renderColumnAnalysis(filteredRows, headers);
    renderPreview(filteredRows, headers);
    renderInsights(filteredRows, headers);
    drawQuickChart(filteredRows, headers[0]);
    renderBusinessCharts(filteredRows);
    updateFilterMeta();

    [summarySection, insightsSection, filterSection, columnsSection, chartsSection, previewSection].forEach(
      (section) => section.classList.remove("hidden"),
    );

    const hasBusinessCharts = Boolean(businessColumns.warehouse || businessColumns.billable);
    businessChartsSection.classList.toggle("hidden", !hasBusinessCharts);
    businessFilters.classList.toggle("hidden", !hasBusinessCharts);
  } catch (error) {
    console.error(error);
    alert("No pude leer el archivo. Revisa que sea un Excel válido (.xlsx o .xls).");
  }
}

function detectBusinessColumns(currentHeaders) {
  const normalizedPairs = currentHeaders.map((header) => ({
    raw: header,
    normalized: normalizeKey(header),
  }));

  const warehouse =
    normalizedPairs.find(({ normalized }) => normalized.includes("almacen"))?.raw || null;

  const billable =
    normalizedPairs.find(({ normalized }) => normalized.includes("facturable"))?.raw || null;

  return { warehouse, billable };
}

function renderSummary(fileName, sheetName, totalRows, totalColumns) {
  summaryCards.innerHTML = "";
  const cards = [
    { label: "Archivo", value: fileName },
    { label: "Hoja analizada", value: sheetName },
    { label: "Total filas originales", value: totalRows.toLocaleString("es-ES") },
    { label: "Total columnas", value: totalColumns.toLocaleString("es-ES") },
  ];

  cards.forEach(({ label, value }) => {
    const card = document.createElement("article");
    card.className = "card";
    card.innerHTML = `<small>${label}</small><strong>${value}</strong>`;
    summaryCards.appendChild(card);
  });
}

function buildSelectors(currentHeaders) {
  const options = ['<option value="__all">Todas las columnas</option>']
    .concat(currentHeaders.map((header) => `<option value="${escapeHtml(header)}">${escapeHtml(header)}</option>`))
    .join("");

  filterColumn.innerHTML = options;

  chartColumn.innerHTML = currentHeaders
    .map((header) => `<option value="${escapeHtml(header)}">${escapeHtml(header)}</option>`)
    .join("");

  chartColumn.value = currentHeaders[0];
}

function buildBusinessSelectors(rows) {
  if (!businessColumns.warehouse) {
    warehouseFilter.innerHTML = '<option value="all">Sin columna almacén</option>';
    return;
  }

  const warehouses = [...new Set(rows.map((row) => String(row[businessColumns.warehouse] ?? "").trim()).filter(Boolean))]
    .sort((a, b) => a.localeCompare(b, "es"));

  warehouseFilter.innerHTML = ['<option value="all">Todos los almacenes</option>']
    .concat(warehouses.map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`))
    .join("");
}

function applyFilters() {
  const query = filterQuery.value.trim().toLowerCase();
  const column = filterColumn.value;

  filteredRows = originalRows.filter((row) => {
    const textMatches = !query
      ? true
      : column === "__all"
        ? headers.some((header) => String(row[header] ?? "").toLowerCase().includes(query))
        : String(row[column] ?? "").toLowerCase().includes(query);

    if (!textMatches) {
      return false;
    }

    if (businessColumns.warehouse && warehouseFilter.value !== "all") {
      const currentWarehouse = String(row[businessColumns.warehouse] ?? "").trim();
      if (currentWarehouse !== warehouseFilter.value) {
        return false;
      }
    }

    if (businessColumns.billable && billableFilter.value !== "all") {
      const billableValue = parseBillableValue(row[businessColumns.billable]);
      if (billableFilter.value === "yes" && billableValue !== true) {
        return false;
      }
      if (billableFilter.value === "no" && billableValue !== false) {
        return false;
      }
    }

    return true;
  });

  renderColumnAnalysis(filteredRows, headers);
  renderPreview(filteredRows, headers);
  renderInsights(filteredRows, headers);
  drawQuickChart(filteredRows, chartColumn.value);
  renderBusinessCharts(filteredRows);
  updateFilterMeta();
}

function updateFilterMeta() {
  filterResult.textContent = `Mostrando ${filteredRows.length.toLocaleString("es-ES")} de ${originalRows.length.toLocaleString("es-ES")} filas.`;
}

function renderBusinessCharts(rows) {
  if (businessColumns.warehouse) {
    drawWarehouseChart(rows, businessColumns.warehouse);
  }

  if (businessColumns.billable) {
    drawBillableChart(rows, businessColumns.billable);
  }
}

function renderColumnAnalysis(rows, headersList) {
  columnsBody.innerHTML = "";
  const columnAnalysis = computeColumnAnalysis(rows, headersList);

  columnAnalysis.forEach(({ header, type, completeness, highlight }) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td>${escapeHtml(header)}</td>
      <td>${type}</td>
      <td><span class="ok">${completeness.toFixed(1)}%</span></td>
      <td>${escapeHtml(highlight)}</td>
    `;

    columnsBody.appendChild(row);
  });
}

function renderInsights(rows, headersList) {
  const analysis = computeColumnAnalysis(rows, headersList);
  drawCompletenessChart(analysis);
  drawColumnTypesChart(analysis);
}

function computeColumnAnalysis(rows, headersList) {
  return headersList.map((header) => {
    const values = rows.map((row) => row[header]);
    const filledValues = values.filter((value) => value !== null && value !== "");
    const completeness = values.length ? (filledValues.length / values.length) * 100 : 0;

    const numericValues = filledValues
      .map((value) => (typeof value === "number" ? value : Number(value)))
      .filter((value) => !Number.isNaN(value));

    const type =
      !filledValues.length
        ? "Vacía"
        : numericValues.length > filledValues.length * 0.7
          ? "Numérica"
          : "Texto / Mixta";

    let highlight = "-";
    if (type === "Numérica" && numericValues.length) {
      const min = Math.min(...numericValues);
      const max = Math.max(...numericValues);
      const avg =
        numericValues.reduce((acc, value) => acc + value, 0) / numericValues.length;
      highlight = `Min: ${formatNumber(min)} · Max: ${formatNumber(max)} · Prom: ${formatNumber(avg)}`;
    } else if (filledValues.length) {
      const counts = new Map();
      filledValues.forEach((value) => {
        const key = String(value).trim();
        counts.set(key, (counts.get(key) || 0) + 1);
      });
      const [topValue, topCount] = [...counts.entries()].sort((a, b) => b[1] - a[1])[0];
      highlight = `Más frecuente: "${topValue}" (${topCount})`;
    }

    return { header, type, completeness, highlight };
  });
}

function renderPreview(rows, headersList) {
  previewHead.innerHTML = "";
  previewBody.innerHTML = "";

  const headRow = document.createElement("tr");
  headersList.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    headRow.appendChild(th);
  });
  previewHead.appendChild(headRow);

  rows.slice(0, 10).forEach((item) => {
    const tr = document.createElement("tr");
    headersList.forEach((header) => {
      const td = document.createElement("td");
      const value = item[header];
      td.textContent = value === null || value === "" ? "—" : String(value);
      tr.appendChild(td);
    });
    previewBody.appendChild(tr);
  });
}

function drawQuickChart(rows, columnName) {
  const context = chartCanvas.getContext("2d");
  const width = chartCanvas.width;
  const height = chartCanvas.height;
  context.clearRect(0, 0, width, height);

  const values = rows
    .map((row) => row[columnName])
    .filter((value) => value !== null && value !== "")
    .map((value) => String(value).trim());

  if (!values.length) {
    drawEmptyState(context, chartCanvas, "No hay datos suficientes para graficar.");
    return;
  }

  const counts = new Map();
  values.forEach((value) => counts.set(value, (counts.get(value) || 0) + 1));

  const top = [...counts.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 8);

  drawHorizontalBars(context, width, top, `Columna: ${columnName}`);
}

function drawWarehouseChart(rows, columnName) {
  const context = warehouseCanvas.getContext("2d");
  const width = warehouseCanvas.width;

  const values = rows
    .map((row) => row[columnName])
    .filter((value) => value !== null && value !== "")
    .map((value) => String(value).trim());

  if (!values.length) {
    drawEmptyState(context, warehouseCanvas, "Sin datos de almacén para mostrar.");
    return;
  }

  const counts = new Map();
  values.forEach((value) => counts.set(value, (counts.get(value) || 0) + 1));

  const top = [...counts.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 8);

  drawHorizontalBars(context, width, top, `Top almacenes (${columnName})`);
}

function drawBillableChart(rows, columnName) {
  const context = billableCanvas.getContext("2d");
  const width = billableCanvas.width;
  const height = billableCanvas.height;

  context.clearRect(0, 0, width, height);

  let yesCount = 0;
  let noCount = 0;

  rows.forEach((row) => {
    const parsed = parseBillableValue(row[columnName]);
    if (parsed === true) {
      yesCount += 1;
    } else if (parsed === false) {
      noCount += 1;
    }
  });

  const total = yesCount + noCount;
  if (!total) {
    drawEmptyState(context, billableCanvas, "Sin datos facturables para mostrar.");
    return;
  }

  const centerX = width / 2;
  const centerY = height / 2;
  const radius = 95;

  const yesAngle = (yesCount / total) * (Math.PI * 2);
  context.beginPath();
  context.moveTo(centerX, centerY);
  context.fillStyle = "#34d399";
  context.arc(centerX, centerY, radius, 0, yesAngle);
  context.fill();

  context.beginPath();
  context.moveTo(centerX, centerY);
  context.fillStyle = "#f87171";
  context.arc(centerX, centerY, radius, yesAngle, Math.PI * 2);
  context.fill();

  context.fillStyle = "#e2e8f0";
  context.font = "bold 14px sans-serif";
  context.fillText("Facturable", 22, 26);
  context.font = "13px sans-serif";
  context.fillText(`Sí: ${yesCount} (${Math.round((yesCount / total) * 100)}%)`, 22, height - 52);
  context.fillText(`No: ${noCount} (${Math.round((noCount / total) * 100)}%)`, 22, height - 30);
}

function drawCompletenessChart(analysis) {
  const context = completenessCanvas.getContext("2d");
  const rows = analysis
    .map((entry) => [entry.header, Number(entry.completeness.toFixed(1))])
    .sort((a, b) => b[1] - a[1])
    .slice(0, 8);

  if (!rows.length) {
    drawEmptyState(context, completenessCanvas, "Sin columnas para analizar.");
    return;
  }

  drawHorizontalBars(context, completenessCanvas.width, rows, "Completitud (%)");
}

function drawColumnTypesChart(analysis) {
  const context = typesCanvas.getContext("2d");
  const width = typesCanvas.width;
  const height = typesCanvas.height;
  context.clearRect(0, 0, width, height);

  const totals = {
    "Numérica": 0,
    "Texto / Mixta": 0,
    "Vacía": 0,
  };

  analysis.forEach((entry) => {
    totals[entry.type] = (totals[entry.type] || 0) + 1;
  });

  const entries = Object.entries(totals).filter(([, count]) => count > 0);
  const grandTotal = entries.reduce((acc, [, count]) => acc + count, 0);

  if (!grandTotal) {
    drawEmptyState(context, typesCanvas, "Sin tipos detectados.");
    return;
  }

  const colors = {
    "Numérica": "#22d3ee",
    "Texto / Mixta": "#a78bfa",
    "Vacía": "#64748b",
  };

  let startAngle = -Math.PI / 2;
  const centerX = width / 2;
  const centerY = height / 2;
  const radius = 92;

  entries.forEach(([type, count]) => {
    const angle = (count / grandTotal) * Math.PI * 2;
    context.beginPath();
    context.moveTo(centerX, centerY);
    context.fillStyle = colors[type];
    context.arc(centerX, centerY, radius, startAngle, startAngle + angle);
    context.fill();
    startAngle += angle;
  });

  context.fillStyle = "#e2e8f0";
  context.font = "bold 13px sans-serif";
  context.fillText("Tipos detectados", 16, 24);
  context.font = "12px sans-serif";

  entries.forEach(([type, count], index) => {
    const y = height - 64 + index * 18;
    const pct = Math.round((count / grandTotal) * 100);
    context.fillStyle = colors[type];
    context.fillRect(16, y - 9, 10, 10);
    context.fillStyle = "#e2e8f0";
    context.fillText(`${type}: ${count} (${pct}%)`, 32, y);
  });
}

function drawHorizontalBars(context, width, entries, title) {
  context.clearRect(0, 0, width, 320);

  const maxCount = entries[0][1] || 1;
  const barAreaWidth = width - 260;
  const barHeight = 24;
  const gap = 12;

  context.fillStyle = "#e2e8f0";
  context.font = "bold 14px sans-serif";
  context.fillText(title, 20, 24);

  entries.forEach(([label, count], index) => {
    const y = 45 + index * (barHeight + gap);
    const safeLabel = label.length > 24 ? `${label.slice(0, 24)}…` : label;
    const barWidth = Math.max((count / maxCount) * barAreaWidth, 2);

    context.fillStyle = "#94a3b8";
    context.font = "13px sans-serif";
    context.fillText(safeLabel, 20, y + 16);

    context.fillStyle = "#22d3ee";
    context.fillRect(200, y, barWidth, barHeight);

    context.fillStyle = "#e2e8f0";
    context.fillText(String(count), 210 + barWidth, y + 16);
  });
}

function drawEmptyState(context, canvas, message) {
  context.clearRect(0, 0, canvas.width, canvas.height);
  context.fillStyle = "#94a3b8";
  context.font = "16px sans-serif";
  context.fillText(message, 20, 40);
}

function parseBillableValue(value) {
  if (typeof value === "boolean") {
    return value;
  }

  const normalized = normalizeKey(String(value ?? ""));
  if (["si", "yes", "true", "1", "x", "checked", "facturable"].includes(normalized)) {
    return true;
  }

  if (["no", "false", "0", "unchecked", "nofacturable"].includes(normalized)) {
    return false;
  }

  return null;
}

function normalizeKey(value) {
  return String(value)
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "")
    .replace(/[^a-zA-Z0-9]/g, "")
    .toLowerCase();
}

function formatNumber(number) {
  return Number(number).toLocaleString("es-ES", {
    maximumFractionDigits: 2,
  });
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
