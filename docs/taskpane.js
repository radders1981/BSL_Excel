"use strict";

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------
let serverUrl = "";
let currentSchema = null; // { model, dimensions: [{name, type}], measures: [] }
let dimTypes = {};         // { fieldName: "date" | "timestamp" | "integer" | ... }
let filterDimensions = [];
let sortFields = [];       // all dims + measures for the current schema
let _dragSrc = null;       // element being dragged

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

// Build left (checkboxes) + right (ordered selection) dual panel.
function buildDualPanel(availId, selectedId, items) {
  buildCheckboxList(availId, items);
  $(selectedId).innerHTML = "";
  addDragHandlers($(selectedId));

  $(availId).addEventListener("change", (e) => {
    if (e.target.type !== "checkbox") return;
    if (e.target.checked) {
      addToSelected(availId, selectedId, e.target.value);
    } else {
      removeFromSelected(selectedId, e.target.value);
    }
  });
}

function buildSelectedItem(availId, selectedId, fieldName) {
  const item = document.createElement("div");
  item.className = "selected-item";
  item.draggable = true;
  item.dataset.field = fieldName;

  const handle = document.createElement("span");
  handle.className = "drag-handle";
  handle.textContent = "⠿";
  handle.setAttribute("aria-hidden", "true");

  const name = document.createElement("span");
  name.className = "item-name";
  name.textContent = fieldName;
  name.title = fieldName;

  const removeBtn = document.createElement("button");
  removeBtn.type = "button";
  removeBtn.className = "btn-deselect";
  removeBtn.textContent = "✕";
  removeBtn.addEventListener("click", () => {
    const cb = [...document.querySelectorAll(`#${availId} input[type=checkbox]`)]
      .find((c) => c.value === fieldName);
    if (cb) cb.checked = false;
    item.remove();
    _removeSortField(fieldName);
  });

  item.appendChild(handle);
  item.appendChild(name);

  const fieldType = dimTypes[fieldName];
  if (fieldType === "date" || fieldType === "timestamp") {
    const grainSel = document.createElement("select");
    grainSel.className = "grain-select";
    grainSel.title = "Time grain";
    [["", "raw"], ["year", "year"], ["quarter", "qtr"], ["month", "month"], ["date", "date"]].forEach(([val, text]) => {
      const opt = document.createElement("option");
      opt.value = val;
      opt.textContent = text;
      grainSel.appendChild(opt);
    });
    grainSel.addEventListener("change", () => {
      item.dataset.grain = grainSel.value;
      grainSel.classList.toggle("grain-active", grainSel.value !== "");
    });
    // Prevent drag from firing when interacting with the select
    grainSel.addEventListener("mousedown", (e) => e.stopPropagation());
    item.appendChild(grainSel);
  }

  item.appendChild(removeBtn);
  return item;
}

function addToSelected(availId, selectedId, fieldName) {
  const container = $(selectedId);
  const already = [...container.querySelectorAll("[data-field]")].some(
    (el) => el.dataset.field === fieldName
  );
  if (already) return;
  container.appendChild(buildSelectedItem(availId, selectedId, fieldName));
  _addSortOption(fieldName);
}

function removeFromSelected(selectedId, fieldName) {
  const item = [...$(selectedId).querySelectorAll("[data-field]")].find(
    (el) => el.dataset.field === fieldName
  );
  if (item) {
    item.remove();
    _removeSortField(fieldName);
  }
}

function _addSortOption(fieldName) {
  document.querySelectorAll("#sortList .sort-row").forEach((row) => {
    const sel = row.querySelector(".sort-field");
    if (![...sel.options].some((o) => o.value === fieldName)) {
      const opt = document.createElement("option");
      opt.value = fieldName;
      opt.textContent = fieldName;
      sel.appendChild(opt);
    }
  });
}

function _removeSortField(fieldName) {
  document.querySelectorAll("#sortList .sort-row").forEach((row) => {
    const sel = row.querySelector(".sort-field");
    if (sel.value === fieldName) {
      row.remove();
      return;
    }
    const opt = [...sel.options].find((o) => o.value === fieldName);
    if (opt) opt.remove();
  });
}

function getSelected(selectedId) {
  return [...$(selectedId).querySelectorAll("[data-field]")].map((el) => el.dataset.field);
}

function addDragHandlers(container) {
  container.addEventListener("dragstart", (e) => {
    const item = e.target.closest(".selected-item");
    if (!item) return;
    _dragSrc = item;
    e.dataTransfer.effectAllowed = "move";
    requestAnimationFrame(() => item.classList.add("dragging"));
  });

  container.addEventListener("dragover", (e) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = "move";
    const target = e.target.closest(".selected-item");
    if (target && target !== _dragSrc) {
      container.querySelectorAll(".selected-item.drag-over").forEach((l) => l.classList.remove("drag-over"));
      target.classList.add("drag-over");
    }
  });

  container.addEventListener("dragleave", (e) => {
    if (!container.contains(e.relatedTarget)) {
      container.querySelectorAll(".selected-item.drag-over").forEach((l) => l.classList.remove("drag-over"));
    }
  });

  container.addEventListener("drop", (e) => {
    e.preventDefault();
    const target = e.target.closest(".selected-item");
    if (target && _dragSrc && target !== _dragSrc) {
      target.classList.remove("drag-over");
      const rect = target.getBoundingClientRect();
      if (e.clientY < rect.top + rect.height / 2) {
        container.insertBefore(_dragSrc, target);
      } else {
        container.insertBefore(_dragSrc, target.nextSibling);
      }
    }
  });

  container.addEventListener("dragend", () => {
    container.querySelectorAll(".selected-item").forEach((l) => l.classList.remove("dragging", "drag-over"));
    _dragSrc = null;
  });
}

function setAllChecked(availId, selectedId, checked) {
  document.querySelectorAll(`#${availId} input[type=checkbox]`).forEach((cb) => {
    if (cb.checked === checked) return;
    cb.checked = checked;
    if (checked) {
      addToSelected(availId, selectedId, cb.value);
    } else {
      removeFromSelected(selectedId, cb.value);
    }
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

function buildSortRow() {
  const row = document.createElement("div");
  row.className = "sort-row";

  const activeFields = [
    ...getSelected("dimensionSelected"),
    ...getSelected("measureSelected"),
  ];

  const fieldSel = document.createElement("select");
  fieldSel.className = "sort-field";
  activeFields.forEach((f) => {
    const opt = document.createElement("option");
    opt.value = f;
    opt.textContent = f;
    fieldSel.appendChild(opt);
  });

  const dirSel = document.createElement("select");
  dirSel.className = "sort-dir";
  [["asc", "↑ asc"], ["desc", "↓ desc"]].forEach(([val, text]) => {
    const opt = document.createElement("option");
    opt.value = val;
    opt.textContent = text;
    dirSel.appendChild(opt);
  });

  const removeBtn = document.createElement("button");
  removeBtn.type = "button";
  removeBtn.className = "btn-remove-filter";
  removeBtn.textContent = "✕";
  removeBtn.addEventListener("click", () => row.remove());

  row.appendChild(fieldSel);
  row.appendChild(dirSel);
  row.appendChild(removeBtn);

  return row;
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

    // data.dimensions is now [{name, type}, ...] — extract names and build type map
    const dimNames = data.dimensions.map((d) => d.name);
    dimTypes = {};
    data.dimensions.forEach((d) => { dimTypes[d.name] = d.type; });

    buildDualPanel("dimensionList", "dimensionSelected", dimNames);
    buildDualPanel("measureList", "measureSelected", data.measures);

    // Reset filter, sort builders
    filterDimensions = dimNames;
    sortFields = [...dimNames, ...data.measures];
    $("filterList").innerHTML = "";
    $("sortList").innerHTML = "";

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

  const dimensions = getSelected("dimensionSelected");
  const measures = getSelected("measureSelected");

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

  const sort_by = [];
  document.querySelectorAll("#sortList .sort-row").forEach((row) => {
    const field = row.querySelector(".sort-field").value;
    const direction = row.querySelector(".sort-dir").value;
    if (field) sort_by.push({ field, direction });
  });

  const grains = {};
  document.querySelectorAll("#dimensionSelected [data-field]").forEach((el) => {
    if (el.dataset.grain) grains[el.dataset.field] = el.dataset.grain;
  });

  const payload = { model: modelName, dimensions, measures, filters, sort_by, grains, limit };

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

  $("addSortBtn").addEventListener("click", () => {
    const activeFields = [
      ...getSelected("dimensionSelected"),
      ...getSelected("measureSelected"),
    ];
    if (activeFields.length === 0) return;
    $("sortList").appendChild(buildSortRow());
  });

  $("dimAllBtn").addEventListener("click", (e) => { e.preventDefault(); setAllChecked("dimensionList", "dimensionSelected", true); });
  $("dimNoneBtn").addEventListener("click", (e) => { e.preventDefault(); setAllChecked("dimensionList", "dimensionSelected", false); });
  $("measAllBtn").addEventListener("click", (e) => { e.preventDefault(); setAllChecked("measureList", "measureSelected", true); });
  $("measNoneBtn").addEventListener("click", (e) => { e.preventDefault(); setAllChecked("measureList", "measureSelected", false); });

  // Allow Enter key in server URL field to connect
  $("serverUrl").addEventListener("keydown", (e) => {
    if (e.key === "Enter") connect();
  });

  setStatus("Enter your server URL and click Connect.", "idle");
});
