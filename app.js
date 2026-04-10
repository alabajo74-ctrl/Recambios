const dropZone = document.getElementById("drop-zone");
const fileInput = document.getElementById("file-input");
const summarySection = document.getElementById("summary");
const filterSection = document.getElementById("filters");
const columnsSection = document.getElementById("columns");
const chartsSection = document.getElementById("charts");
const previewSection = document.getElementById("preview");
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

let originalRows = [];
let filteredRows = [];
let headers = [];

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
clearFilters.addEventListener("click", () => {
  filterColumn.value = "__all";
  filterQuery.value = "";
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

    renderSummary(file.name, firstSheetName, originalRows.length, headers.length);
    buildSelectors(headers);
    renderColumnAnalysis(filteredRows, headers);
    renderPreview(filteredRows, headers);
    drawQuickChart(filteredRows, headers[0]);
    updateFilterMeta();

    [summarySection, filterSection, columnsSection, chartsSection, previewSection].forEach(
      (section) => section.classList.remove("hidden"),
    );
  } catch (error) {
    console.error(error);
    alert("No pude leer el archivo. Revisa que sea un Excel válido (.xlsx o .xls).");
  }
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

function applyFilters() {
  const query = filterQuery.value.trim().toLowerCase();
  const column = filterColumn.value;

  filteredRows = originalRows.filter((row) => {
    if (!query) {
      return true;
    }

    if (column === "__all") {
      return headers.some((header) => String(row[header] ?? "").toLowerCase().includes(query));
    }

    return String(row[column] ?? "").toLowerCase().includes(query);
  });

  renderColumnAnalysis(filteredRows, headers);
  renderPreview(filteredRows, headers);
  drawQuickChart(filteredRows, chartColumn.value);
  updateFilterMeta();
}

function updateFilterMeta() {
  filterResult.textContent = `Mostrando ${filteredRows.length.toLocaleString("es-ES")} de ${originalRows.length.toLocaleString("es-ES")} filas.`;
}

function renderColumnAnalysis(rows, headersList) {
  columnsBody.innerHTML = "";

  headersList.forEach((header) => {
    const values = rows.map((row) => row[header]);
    const filledValues = values.filter((value) => value !== null && value !== "");
    const completeness = values.length ? (filledValues.length / values.length) * 100 : 0;

    const numericValues = filledValues
      .map((value) => (typeof value === "number" ? value : Number(value)))
      .filter((value) => !Number.isNaN(value));

    const type =
      numericValues.length > filledValues.length * 0.7
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
    context.fillStyle = "#94a3b8";
    context.font = "16px sans-serif";
    context.fillText("No hay datos suficientes para graficar.", 20, 40);
    return;
  }

  const counts = new Map();
  values.forEach((value) => counts.set(value, (counts.get(value) || 0) + 1));

  const top = [...counts.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 8);

  const maxCount = top[0][1] || 1;
  const barAreaWidth = width - 260;
  const barHeight = 24;
  const gap = 12;

  context.fillStyle = "#e2e8f0";
  context.font = "bold 14px sans-serif";
  context.fillText(`Columna: ${columnName}`, 20, 24);

  top.forEach(([label, count], index) => {
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
