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
    const rows = XLSX.utils.sheet_to_json(worksheet, { defval: null });

    if (!rows.length) {
      alert("El archivo no contiene filas con datos en la primera hoja.");
      return;
    }

    const headers = Object.keys(rows[0]);
    renderSummary(file.name, firstSheetName, rows.length, headers.length);
    renderColumnAnalysis(rows, headers);
    renderPreview(rows, headers);

    [summarySection, columnsSection, previewSection].forEach((section) =>
      section.classList.remove("hidden"),
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
    { label: "Total de filas", value: totalRows.toLocaleString("es-ES") },
    { label: "Total de columnas", value: totalColumns.toLocaleString("es-ES") },
  ];

  cards.forEach(({ label, value }) => {
    const card = document.createElement("article");
    card.className = "card";
    card.innerHTML = `<small>${label}</small><strong>${value}</strong>`;
    summaryCards.appendChild(card);
  });
}

function renderColumnAnalysis(rows, headers) {
  columnsBody.innerHTML = "";

  headers.forEach((header) => {
    const values = rows.map((row) => row[header]);
    const filledValues = values.filter((value) => value !== null && value !== "");
    const completeness = (filledValues.length / values.length) * 100;

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
      <td>${header}</td>
      <td>${type}</td>
      <td><span class="ok">${completeness.toFixed(1)}%</span></td>
      <td>${highlight}</td>
    `;

    columnsBody.appendChild(row);
  });
}

function renderPreview(rows, headers) {
  previewHead.innerHTML = "";
  previewBody.innerHTML = "";

  const headRow = document.createElement("tr");
  headers.forEach((header) => {
    const th = document.createElement("th");
    th.textContent = header;
    headRow.appendChild(th);
  });
  previewHead.appendChild(headRow);

  rows.slice(0, 10).forEach((item) => {
    const tr = document.createElement("tr");
    headers.forEach((header) => {
      const td = document.createElement("td");
      const value = item[header];
      td.textContent = value === null || value === "" ? "—" : String(value);
      tr.appendChild(td);
    });
    previewBody.appendChild(tr);
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
