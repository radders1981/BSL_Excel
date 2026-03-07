"use strict";

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------
let serverUrl = "";
let currentSchema = null; // { model, dimensions: [], measures: [] }
let filterDimensions = [];

// ---------------------------------------------------------------------------
// DOM helpers
// ---------------------------------------------------------------------------
const $ = (id) => document.getElementById(id);

function setStatus(msg, type = "idle") {
  const el = $("status");
  el.textContent = msg;
  el.className = type;
}

function buildCheckboxList(containerId, items) {
  const container = $(containerId);
  container.innerHTML = "";
  if (!items || items.length === 0) {
    container.innerHTML = '<span style="color:#999;padding:4px;">None available</span>';
    return;
  }
  items.forEach((item) => {
    const label = document.createElement("label");
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.value = item;
    cb.checked = false;
    label.appendChild(cb);
    label.appendChild(document.createTextNode(item));
    container.appendChild(label);
  });
}

function buildFilterRow() {
  const row = document.createElement("div");
  row.className = "filter-row";

  const fieldSel = document.createElement("select");
  fieldSel.className = "filter-field";
  filterDimensions.forEach((d) => {
    const opt = document.createElement("option");
    opt.value = d;
    opt.textContent = d;
    fieldSel.appendChild(opt);
  });

  const opSel = document.createElement("select");
  opSel.className = "filter-op";
  [["eq", "="], ["neq", "!="], ["gt", ">"], ["gte", ">="], ["lt", "<"], ["lte", "<="], ["contains", "contains"]].forEach(
    ([val, text]) => {
      const opt = document.createElement("option");
      opt.value = val;
      opt.textContent = text;
      opSel.appendChild(opt);
    }
  );

  const valInput = document.createElement("input");
  valInput.type = "text";
  valInput.className = "filter-value";
  valInput.placeholder = "value";

  const removeBtn = document.createElement("button");
  removeBtn.type = "button";
  removeBtn.className = "btn-remove-filter";
  removeBtn.textContent = "✕";
  removeBtn.addEventListener("click", () => row.remove());

  row.appendChild(fieldSel);
  row.appendChild(opSel);
  row.appendChild(valInput);
  row.appendChild(removeBtn);

  return row;
}

function getChecked(containerId) {
  return Array.from(
    document.querySelectorAll(`#${containerId} input[type=checkbox]:checked`)
  ).map((cb) => cb.value);
}

// ---------------------------------------------------------------------------
// Connect flow
// ---------------------------------------------------------------------------
async function connect() {
  const rawUrl = $("serverUrl").value.trim().replace(/\/$/, "");
  if (!rawUrl) {
    setStatus("Please enter a server URL.", "error");
    return;
  }
  serverUrl = rawUrl;
  localStorage.setItem("bsl_server_url", serverUrl);
  setStatus("Connecting...", "loading");

  try {
    const res = await fetch(`${serverUrl}/health`);
    if (!res.ok) throw new Error(`Server returned ${res.status}`);
    await res.json();
  } catch (e) {
    setStatus(`Cannot reach server: ${e.message}`, "error");
    return;
  }

  // Load model list
  try {
    const res = await fetch(`${serverUrl}/models`);
    const data = await res.json();
    const select = $("modelSelect");
    select.innerHTML = '<option value="">-- select a model --</option>';
    data.models.forEach((m) => {
      const opt = document.createElement("option");
      opt.value = m;
      opt.textContent = m;
      select.appendChild(opt);
    });
    $("querySection").classList.remove("hidden");
    setStatus("Connected. Select a model.", "ok");
  } catch (e) {
    setStatus(`Failed to load models: ${e.message}`, "error");
  }
}

// ---------------------------------------------------------------------------
// Schema load flow
// ---------------------------------------------------------------------------
async function loadSchema() {
  const modelName = $("modelSelect").value;
  if (!modelName) {
    $("fieldSection").classList.add("hidden");
    return;
  }
  setStatus("Loading schema...", "loading");

  try {
    const res = await fetch(`${serverUrl}/models/${modelName}/schema`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.detail || data.error || `HTTP ${res.status}`);
    currentSchema = data;

    buildCheckboxList("dimensionList", data.dimensions);
    buildCheckboxList("measureList", data.measures);

    // Reset filter builder
    filterDimensions = data.dimensions;
    $("filterList").innerHTML = "";

    $("fieldSection").classList.remove("hidden");
    setStatus(`Schema loaded for '${modelName}'.`, "ok");
  } catch (e) {
    setStatus(`Failed to load schema: ${e.message}`, "error");
  }
}

// ---------------------------------------------------------------------------
// Run query flow
// ---------------------------------------------------------------------------
async function runQuery() {
  const modelName = $("modelSelect").value;
  if (!modelName) {
    setStatus("Select a model first.", "error");
    return;
  }

  const dimensions = getChecked("dimensionList");
  const measures = getChecked("measureList");

  if (dimensions.length === 0 && measures.length === 0) {
    setStatus("Select at least one dimension or measure.", "error");
    return;
  }

  const filters = [];
  document.querySelectorAll("#filterList .filter-row").forEach((row) => {
    const field = row.querySelector(".filter-field").value;
    const op = row.querySelector(".filter-op").value;
    const value = row.querySelector(".filter-value").value.trim();
    if (field && value) {
      filters.push({ dimension: field, op, value });
    }
  });

  const limit = parseInt($("limitInput").value, 10) || 1000;

  const payload = { model: modelName, dimensions, measures, filters, limit };

  setStatus("Running query...", "loading");
  $("runBtn").disabled = true;

  try {
    const res = await fetch(`${serverUrl}/query`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const data = await res.json();
    if (!res.ok) {
      throw new Error(data.detail || data.error || "Query failed");
    }
    await writeToExcel(data.columns, data.rows);
    setStatus(`Done. ${data.rows.length} row(s) written to sheet.`, "ok");
  } catch (e) {
    setStatus(`Error: ${e.message}`, "error");
  } finally {
    $("runBtn").disabled = false;
  }
}

// ---------------------------------------------------------------------------
// Excel write
// ---------------------------------------------------------------------------
async function writeToExcel(columns, rows) {
  await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address,rowIndex,columnIndex");
    await context.sync();

    const totalRows = rows.length + 1; // +1 for header
    const totalCols = columns.length;

    const sheet = range.worksheet;
    const startCell = sheet.getCell(range.rowIndex, range.columnIndex);
    const dataRange = startCell.getResizedRange(totalRows - 1, totalCols - 1);

    // Build 2D array: header row + data rows
    const values = [columns, ...rows];
    dataRange.values = values;

    // Style header row bold
    const headerRange = startCell.getResizedRange(0, totalCols - 1);
    headerRange.format.font.bold = true;
    headerRange.format.fill.color = "#EFF6FC";

    // Auto-fit columns
    dataRange.format.autofitColumns();

    await context.sync();
  });
}

// ---------------------------------------------------------------------------
// Init
// ---------------------------------------------------------------------------
Office.onReady(() => {
  // Restore saved server URL
  const saved = localStorage.getItem("bsl_server_url");
  if (saved) {
    $("serverUrl").value = saved;
  }

  $("connectBtn").addEventListener("click", connect);
  $("modelSelect").addEventListener("change", loadSchema);
  $("runBtn").addEventListener("click", runQuery);
  $("addFilterBtn").addEventListener("click", () => {
    if (filterDimensions.length === 0) return;
    $("filterList").appendChild(buildFilterRow());
  });

  // Allow Enter key in server URL field to connect
  $("serverUrl").addEventListener("keydown", (e) => {
    if (e.key === "Enter") connect();
  });

  setStatus("Enter your server URL and click Connect.", "idle");
});
