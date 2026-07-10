/* global XLSX, loadPyodide */

(() => {
  "use strict";

  const MODES = {
    categories: {
      templateUrl: "sample_sheets/NewProductCateg.csv",
      lede: "Two-step import — categories first, then products",
      uploadHint: "Drop your supplier file here, or browse to upload",
      loadedMessage: (rows) => `${rows} rows loaded — click Generate`,
      processLabel: "Generate",
      processingLabel: "Generating…",
      successMessage: (count) => `${count} categories ready. Download Excel or CSV.`,
      previewLabel: (count, previewLen) =>
        previewLen < count ? `${previewLen} of ${count} rows` : `${count} rows`,
      downloadSuffix: "_categories",
      headersKey: "category_headers",
    },
    products: {
      templateUrl: "sample_sheets/SampleOdooFormat.xlsx",
      lede: "Two-step import — categories first, then products",
      uploadHint: "Drop your supplier product file here, or browse to upload",
      loadedMessage: (rows) => `${rows} product rows loaded — upload Odoo category export, then Convert`,
      supplierReadyMessage: (rows) => `${rows} product rows loaded — now upload the Odoo category export (File 2)`,
      categoryReadyMessage: "Category export validated — click Convert",
      processLabel: "Convert",
      processingLabel: "Converting…",
      successMessage: (count) => `${count} products ready. Download Product Excel or CSV.`,
      previewLabel: (count, previewLen) =>
        previewLen < count ? `${previewLen} of ${count} rows` : `${count} rows`,
      downloadSuffix: "_odoo",
      headersKey: "odoo_headers",
      excelLabel: "Download Product Excel (.xlsx)",
      csvLabel: "Download Product CSV (.csv)",
    },
  };

  const CATEGORY_REQUIRED = [
    "parent category/category name",
    "category name",
    "external id",
  ];

  const CATEGORY_REQUIRED_LABELS = {
    "parent category/category name": "Parent Category/Category Name",
    "category name": "Category Name",
    "external id": "External ID",
  };

  const els = {
    heroLede: document.getElementById("heroLede"),
    tabProducts: document.getElementById("tabProducts"),
    tabCategories: document.getElementById("tabCategories"),
    uploadHint: document.getElementById("uploadHint"),
    dropZone: document.getElementById("dropZone"),
    fileInput: document.getElementById("fileInput"),
    uploadIdle: document.getElementById("uploadIdle"),
    uploadReady: document.getElementById("uploadReady"),
    fileName: document.getElementById("fileName"),
    fileMeta: document.getElementById("fileMeta"),
    fileBadge: document.getElementById("fileBadge"),
    messages: document.getElementById("messages"),
    actions: document.getElementById("actions"),
    btnProcess: document.getElementById("btnProcess"),
    btnExcel: document.getElementById("btnExcel"),
    btnCsv: document.getElementById("btnCsv"),
    progress: document.getElementById("progress"),
    progressFill: document.getElementById("progressFill"),
    progressPct: document.getElementById("progressPct"),
    previewWrap: document.getElementById("previewWrap"),
    previewCount: document.getElementById("previewCount"),
    previewTable: document.getElementById("previewTable"),
    mappingProducts: document.getElementById("mappingProducts"),
    mappingCategories: document.getElementById("mappingCategories"),
    supplierFileLabel: document.getElementById("supplierFileLabel"),
    categoryUpload: document.getElementById("categoryUpload"),
    categoryFileInput: document.getElementById("categoryFileInput"),
    categoryDropZone: document.getElementById("categoryDropZone"),
    categoryUploadIdle: document.getElementById("categoryUploadIdle"),
    categoryUploadReady: document.getElementById("categoryUploadReady"),
    categoryFileBadge: document.getElementById("categoryFileBadge"),
    categoryFileName: document.getElementById("categoryFileName"),
    categoryFileMeta: document.getElementById("categoryFileMeta"),
  };

  function createEmptyModeState(forProducts = false) {
    const base = { file: null, rows: null, detectedFormat: null, result: null };
    if (!forProducts) return base;
    return {
      ...base,
      categoryFile: null,
      categoryRows: null,
      categoryFormat: null,
      categoryValid: false,
    };
  }

  const state = {
    mode: "categories",
    pyodide: null,
    ready: false,
    byMode: {
      categories: createEmptyModeState(false),
      products: createEmptyModeState(true),
    },
    templateBytesByMode: {},
  };

  function currentMode() {
    return state.byMode[state.mode];
  }

  function modeConfig() {
    return MODES[state.mode];
  }

  function clearMessages() {
    els.messages.innerHTML = "";
  }

  function showMessage(type, text) {
    const div = document.createElement("div");
    div.className = `msg ${type}`;
    div.textContent = text;
    els.messages.appendChild(div);
  }

  function setProgress(pct) {
    const value = Math.max(0, Math.min(100, Math.round(Number(pct) || 0)));
    els.progress.hidden = false;
    els.progressFill.style.width = `${value}%`;
    els.progressPct.textContent = `${value}%`;
    els.progress.setAttribute("aria-valuenow", String(value));
  }

  function hideProgress() {
    els.progress.hidden = true;
    els.progressFill.style.width = "0%";
    els.progressPct.textContent = "0%";
    els.progress.setAttribute("aria-valuenow", "0");
  }

  function tickUI() {
    return new Promise((resolve) => {
      requestAnimationFrame(() => setTimeout(resolve, 0));
    });
  }

  function normalizeHeader(value) {
    return String(value ?? "")
      .trim()
      .toLowerCase()
      .replace(/\s*\/\s*/g, "/")
      .replace(/\s+/g, " ");
  }

  function validateCategoryRows(rows) {
    if (!rows?.length) {
      throw new Error("The Odoo category export file is empty.");
    }

    let headerIdx = -1;
    for (let i = 0; i < Math.min(rows.length, 20); i++) {
      const norms = new Set(rows[i].map(normalizeHeader).filter(Boolean));
      if (norms.has("category name") && norms.has("external id")) {
        headerIdx = i;
        break;
      }
    }

    if (headerIdx < 0) {
      throw new Error(
        "Could not find required headers. The file must contain: Parent Category/Category Name, Category Name, External ID."
      );
    }

    const norms = new Set(rows[headerIdx].map(normalizeHeader).filter(Boolean));
    const missing = CATEGORY_REQUIRED.filter((key) => !norms.has(key));
    if (missing.length) {
      throw new Error(
        `The uploaded category file is missing required column(s): ${missing
          .map((key) => CATEGORY_REQUIRED_LABELS[key])
          .join(", ")}. The file must contain: Parent Category/Category Name, Category Name, External ID.`
      );
    }

    return true;
  }

  function canProcess(modeName = state.mode) {
    const mode = state.byMode[modeName];
    if (!state.ready || !mode?.rows) return false;
    if (modeName === "products") {
      return Boolean(mode.categoryValid && mode.categoryRows);
    }
    return true;
  }

  function updateDownloadLabels() {
    const cfg = modeConfig();
    if (state.mode === "products") {
      els.btnExcel.textContent = cfg.excelLabel || "Download Product Excel (.xlsx)";
      els.btnCsv.textContent = cfg.csvLabel || "Download Product CSV (.csv)";
    } else {
      els.btnExcel.textContent = "Download Excel (.xlsx)";
      els.btnCsv.textContent = "Download CSV (.csv)";
    }
  }

  function updateProductsPanels() {
    const isProducts = state.mode === "products";
    els.supplierFileLabel.hidden = !isProducts;
    els.categoryUpload.hidden = !isProducts;

    if (!isProducts) return;

    const mode = currentMode();
    els.categoryUploadIdle.hidden = Boolean(mode.categoryFile);
    els.categoryUploadReady.hidden = !mode.categoryFile;

    if (mode.categoryFile) {
      els.categoryFileBadge.textContent = formatBadge(mode.categoryFormat);
      els.categoryFileName.textContent = mode.categoryFile.name;
      els.categoryFileMeta.textContent = `${(mode.categoryFile.size / 1024).toFixed(1)} KB`;
    }
  }

  function productsStatusMessage(mode) {
    const cfg = MODES.products;
    if (mode.result) return cfg.successMessage(mode.result.row_count);
    if (mode.rows && mode.categoryValid) return cfg.categoryReadyMessage;
    if (mode.rows) return cfg.supplierReadyMessage(rowsLength(mode.rows));
    return cfg.uploadHint;
  }

  function refreshProcessButton() {
    els.btnProcess.disabled = !canProcess();
  }
  function applyModeUI() {
    const cfg = modeConfig();
    els.heroLede.textContent = cfg.lede;
    els.uploadHint.textContent = cfg.uploadHint;
    els.btnProcess.textContent = cfg.processLabel;
    els.tabCategories.classList.toggle("active", state.mode === "categories");
    els.tabProducts.classList.toggle("active", state.mode === "products");
    els.tabCategories.setAttribute("aria-selected", state.mode === "categories" ? "true" : "false");
    els.tabProducts.setAttribute("aria-selected", state.mode === "products" ? "true" : "false");
    els.mappingCategories.hidden = state.mode !== "categories";
    els.mappingProducts.hidden = state.mode !== "products";
    updateDownloadLabels();
    updateProductsPanels();
  }

  function initCollapsibles() {
    document.querySelectorAll(".collapse").forEach((section) => {
      const trigger = section.querySelector(".collapse-trigger");
      const body = section.querySelector(".collapse-body");
      if (!trigger || !body) return;

      trigger.addEventListener("click", () => {
        const expanded = trigger.getAttribute("aria-expanded") === "true";
        trigger.setAttribute("aria-expanded", String(!expanded));
        body.hidden = expanded;
      });
    });
  }

  function formatBadge(format) {
    if (!format) return "FILE";
    if (format.includes("xlsx")) return "XLSX";
    if (format.includes("xls")) return "XLS";
    return "CSV";
  }

  function clearAllModes() {
    state.byMode.categories = createEmptyModeState(false);
    state.byMode.products = createEmptyModeState(true);
  }

  function resetWorkflow() {
    clearAllModes();
    els.uploadIdle.hidden = false;
    els.uploadReady.hidden = true;
    els.actions.hidden = true;
    els.previewWrap.hidden = true;
    els.fileInput.value = "";
    if (els.categoryFileInput) els.categoryFileInput.value = "";
    hideProgress();
    clearMessages();
    els.btnExcel.disabled = true;
    els.btnCsv.disabled = true;
    updateProductsPanels();
    refreshProcessButton();
  }

  function restoreModeWorkflow() {
    const cfg = modeConfig();
    const mode = currentMode();
    hideProgress();
    clearMessages();
    els.btnProcess.textContent = cfg.processLabel;
    els.fileInput.value = "";
    if (els.categoryFileInput) els.categoryFileInput.value = "";
    updateProductsPanels();

    if (!mode.file || !mode.rows) {
      els.uploadIdle.hidden = false;
      els.uploadReady.hidden = true;
      els.actions.hidden = true;
      els.previewWrap.hidden = true;
      els.btnExcel.disabled = true;
      els.btnCsv.disabled = true;
      refreshProcessButton();
      return;
    }

    els.uploadIdle.hidden = true;
    els.uploadReady.hidden = false;
    els.actions.hidden = false;
    if (els.fileBadge) els.fileBadge.textContent = formatBadge(mode.detectedFormat);
    els.fileName.textContent = mode.file.name;
    els.fileMeta.textContent = `${(mode.file.size / 1024).toFixed(1)} KB`;

    if (mode.result) {
      renderPreview(mode.result);
      els.btnExcel.disabled = false;
      els.btnCsv.disabled = false;
      showMessage(
        "info",
        state.mode === "products"
          ? MODES.products.successMessage(mode.result.row_count)
          : cfg.successMessage(mode.result.row_count)
      );
    } else {
      els.previewWrap.hidden = true;
      els.btnExcel.disabled = true;
      els.btnCsv.disabled = true;
      const msg =
        state.mode === "products"
          ? productsStatusMessage(mode)
          : cfg.loadedMessage(rowsLength(mode.rows));
      showMessage("info", msg);
    }
    refreshProcessButton();
  }

  const SUPPORTED_FILE_RE = /\.(csv|xlsx|xls)$/i;

  function unsupportedFileMessage(file) {
    const name = file?.name || "This file";
    const dot = name.lastIndexOf(".");
    const ext = dot >= 0 ? name.slice(dot).toLowerCase() : "";
    if (ext) {
      return `Unsupported file type (${ext}). Please upload a .csv, .xlsx, or .xls spreadsheet.`;
    }
    return "Unsupported file type. Please upload a .csv, .xlsx, or .xls spreadsheet.";
  }

  function isSupportedFile(file) {
    return Boolean(file && SUPPORTED_FILE_RE.test(file.name || ""));
  }

  function rejectUnsupportedFile(file) {
    showMessage("error", unsupportedFileMessage(file));
    els.fileInput.value = "";
  }

  function switchMode(mode) {
    if (state.mode === mode) return;
    state.mode = mode;
    applyModeUI();
    restoreModeWorkflow();
  }

  function detectFormat(file) {
    const name = (file.name || "").toLowerCase();
    if (name.endsWith(".csv")) return "CSV";
    if (name.endsWith(".xlsx")) return "Excel (.xlsx)";
    if (name.endsWith(".xls")) return "Excel (.xls)";

    const type = (file.type || "").toLowerCase();
    if (type.includes("csv") || type === "text/plain") return "CSV";
    if (type.includes("spreadsheetml") || type.includes("openxmlformats")) return "Excel (.xlsx)";
    if (type === "application/vnd.ms-excel") return "Excel (.xls)";

    return null;
  }

  function detectFormatFromBytes(bytes, fallbackName) {
    if (bytes.length >= 4) {
      if (bytes[0] === 0x50 && bytes[1] === 0x4b) return "Excel (.xlsx)";
      if (bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0) {
        return "Excel (.xls)";
      }
    }
    const name = (fallbackName || "").toLowerCase();
    if (name.endsWith(".csv")) return "CSV";
    return null;
  }

  function sheetToRows(workbook) {
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: "",
      raw: false,
      blankrows: false,
    });
  }

  async function parseFile(file) {
    const buffer = await file.arrayBuffer();
    const bytes = new Uint8Array(buffer);
    let format = detectFormat(file) || detectFormatFromBytes(bytes, file.name);
    if (!format) {
      throw new Error("This file could not be read as a spreadsheet. Please upload a .csv, .xlsx, or .xls file.");
    }

    let rows;
    if (format === "CSV") {
      const text = new TextDecoder("utf-8").decode(bytes);
      const workbook = XLSX.read(text, { type: "string", raw: false });
      rows = sheetToRows(workbook);
    } else {
      const workbook = XLSX.read(bytes, { type: "array", cellDates: true });
      rows = sheetToRows(workbook);
      format = detectFormatFromBytes(bytes, file.name) || format;
    }

    return { format, rows };
  }

  function setFileUI(file, format) {
    const cfg = modeConfig();
    const mode = currentMode();
    els.uploadIdle.hidden = true;
    els.uploadReady.hidden = false;
    els.actions.hidden = false;
    if (els.fileBadge) els.fileBadge.textContent = formatBadge(format);
    els.fileName.textContent = file.name;
    els.fileMeta.textContent = `${(file.size / 1024).toFixed(1)} KB`;
    els.btnExcel.disabled = true;
    els.btnCsv.disabled = true;
    els.previewWrap.hidden = true;
    hideProgress();
    mode.result = null;
    updateProductsPanels();
    const msg =
      state.mode === "products"
        ? productsStatusMessage(mode)
        : cfg.loadedMessage(rowsLength(mode.rows));
    showMessage("info", msg);
    refreshProcessButton();
  }

  function rowsLength(rows) {
    return Array.isArray(rows) ? rows.length : 0;
  }

  async function handleFile(file) {
    clearMessages();
    if (!file) return;

    if (!isSupportedFile(file)) {
      rejectUnsupportedFile(file);
      return;
    }

    try {
      const { format, rows } = await parseFile(file);
      const mode = currentMode();
      mode.file = file;
      mode.detectedFormat = format;
      mode.rows = rows;
      mode.result = null;
      setFileUI(file, format);
    } catch (err) {
      console.error(err);
      showMessage("error", err.message || String(err));
      els.fileInput.value = "";
    }
  }

  async function handleCategoryFile(file) {
    if (state.mode !== "products") return;
    clearMessages();
    if (!file) return;

    if (!isSupportedFile(file)) {
      showMessage("error", unsupportedFileMessage(file));
      els.categoryFileInput.value = "";
      return;
    }

    const mode = currentMode();
    try {
      const { format, rows } = await parseFile(file);
      validateCategoryRows(rows);
      mode.categoryFile = file;
      mode.categoryRows = rows;
      mode.categoryFormat = format;
      mode.categoryValid = true;
      mode.result = null;
      els.categoryUploadIdle.hidden = true;
      els.categoryUploadReady.hidden = false;
      els.categoryFileBadge.textContent = formatBadge(format);
      els.categoryFileName.textContent = file.name;
      els.categoryFileMeta.textContent = `${(file.size / 1024).toFixed(1)} KB`;
      els.previewWrap.hidden = true;
      els.btnExcel.disabled = true;
      els.btnCsv.disabled = true;
      if (mode.file) {
        els.actions.hidden = false;
        showMessage("info", productsStatusMessage(mode));
      } else {
        showMessage("info", "Category export validated — now upload the supplier product file (File 1).");
      }
      refreshProcessButton();
    } catch (err) {
      console.error(err);
      mode.categoryFile = null;
      mode.categoryRows = null;
      mode.categoryFormat = null;
      mode.categoryValid = false;
      updateProductsPanels();
      showMessage("error", err.message || String(err));
      els.categoryFileInput.value = "";
      refreshProcessButton();
    }
  }

  function renderPreview(result) {
    const cfg = modeConfig();
    const thead = els.previewTable.querySelector("thead");
    const tbody = els.previewTable.querySelector("tbody");
    thead.innerHTML = "";
    tbody.innerHTML = "";

    const headers = result[cfg.headersKey] || [];
    const preview = result.preview || [];

    const hr = document.createElement("tr");
    headers.forEach((h) => {
      const th = document.createElement("th");
      th.textContent = h;
      hr.appendChild(th);
    });
    thead.appendChild(hr);

    preview.forEach((row) => {
      const tr = document.createElement("tr");
      headers.forEach((h) => {
        const td = document.createElement("td");
        const val = row[h];
        td.textContent = val == null || val === "" ? "" : String(val);
        td.title = td.textContent;
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });

    els.previewCount.textContent = cfg.previewLabel(result.row_count, preview.length);
    els.previewWrap.hidden = false;
  }

  function base64ToBlob(base64, mime) {
    const binary = atob(base64);
    const len = binary.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) bytes[i] = binary.charCodeAt(i);
    return new Blob([bytes], { type: mime });
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function downloadBaseName() {
    const cfg = modeConfig();
    const name = currentMode().file?.name || "odoo_export";
    const base = name.replace(/\.(csv|xlsx|xls)$/i, "") + cfg.downloadSuffix;
    const now = new Date();
    const pad = (n) => String(n).padStart(2, "0");
    const stamp =
      `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}` +
      `_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
    return `${base}_${stamp}`;
  }

  async function loadTemplateBytes() {
    if (state.templateBytesByMode[state.mode]) {
      return state.templateBytesByMode[state.mode];
    }
    const cfg = modeConfig();
    const resp = await fetch(cfg.templateUrl);
    if (!resp.ok) throw new Error(`Could not load template (${resp.status})`);
    const bytes = new Uint8Array(await resp.arrayBuffer());
    state.templateBytesByMode[state.mode] = bytes;
    return bytes;
  }

  async function processProducts(pyodide) {
    const mode = currentMode();
    const templateBytes = await loadTemplateBytes();
    pyodide.FS.writeFile("/tmp/template.xlsx", templateBytes);
    pyodide.globals.set("SUPPLIER_ROWS_JSON", JSON.stringify(mode.rows));
    pyodide.globals.set("CATEGORY_ROWS_JSON", JSON.stringify(mode.categoryRows));
    pyodide.globals.set("DETECTED_FORMAT", mode.detectedFormat);

    setProgress(8);
    await tickUI();
    await pyodide.runPythonAsync(`
import json, convert
_rows = json.loads(SUPPLIER_ROWS_JSON)
_category_rows = json.loads(CATEGORY_ROWS_JSON)
_headers, _records = convert.parse_supplier_rows(_rows)
_category_headers, _category_records = convert.parse_odoo_category_export(_category_rows)
`);

    setProgress(20);
    await tickUI();
    await pyodide.runPythonAsync(`
_lookups = convert.build_category_lookups(_category_headers, _category_records)
_odoo_headers, _mapped = convert.map_to_odoo(_headers, _records, _lookups)
`);

    setProgress(40);
    await tickUI();
    await pyodide.runPythonAsync(`
_template = open("/tmp/template.xlsx", "rb").read()
_xlsx_bytes = convert.build_xlsx_bytes(_template, _mapped)
`);

    setProgress(78);
    await tickUI();
    await pyodide.runPythonAsync(`
_csv_text = convert.build_csv_text(_odoo_headers, _mapped)
`);

    setProgress(90);
    await tickUI();
    const resultJson = await pyodide.runPythonAsync(`
import json, base64
json.dumps({
    "ok": True,
    "detected_format": DETECTED_FORMAT,
    "supplier_headers": _headers,
    "row_count": len(_mapped),
    "odoo_headers": _odoo_headers,
    "preview": _mapped[:100],
    "xlsx_base64": base64.b64encode(_xlsx_bytes).decode("ascii"),
    "csv_text": _csv_text,
})
`);
    return JSON.parse(resultJson);
  }

  async function processCategories(pyodide) {
    const mode = currentMode();
    const templateBytes = await loadTemplateBytes();
    pyodide.FS.writeFile("/tmp/category_template.csv", templateBytes);
    pyodide.globals.set("SUPPLIER_ROWS_JSON", JSON.stringify(mode.rows));
    pyodide.globals.set("DETECTED_FORMAT", mode.detectedFormat);

    setProgress(10);
    await tickUI();
    await pyodide.runPythonAsync(`
import json, categories
_rows = json.loads(SUPPLIER_ROWS_JSON)
_headers, _records = categories.parse_category_source_rows(_rows)
`);

    setProgress(35);
    await tickUI();
    await pyodide.runPythonAsync(`
_category_headers, _mapped = categories.build_category_records(_records)
`);

    setProgress(60);
    await tickUI();
    await pyodide.runPythonAsync(`
_xlsx_bytes = categories.build_category_xlsx_bytes(_category_headers, _mapped)
`);

    setProgress(85);
    await tickUI();
    await pyodide.runPythonAsync(`
_csv_text = categories.build_category_csv_text(_category_headers, _mapped)
`);

    setProgress(92);
    await tickUI();
    const resultJson = await pyodide.runPythonAsync(`
import json, base64
json.dumps({
    "ok": True,
    "detected_format": DETECTED_FORMAT,
    "supplier_headers": _headers,
    "row_count": len(_mapped),
    "category_headers": _category_headers,
    "preview": _mapped[:100],
    "xlsx_base64": base64.b64encode(_xlsx_bytes).decode("ascii"),
    "csv_text": _csv_text,
})
`);
    return JSON.parse(resultJson);
  }

  async function processFile() {
    const cfg = modeConfig();
    const mode = currentMode();
    clearMessages();
    if (!state.ready || !mode.rows) {
      showMessage("error", "Wait for the Python engine to finish loading, then upload a file.");
      return;
    }

    if (state.mode === "products") {
      if (!mode.categoryRows || !mode.categoryValid) {
        showMessage(
          "error",
          "Upload a valid Odoo category export file (File 2) before converting. The file must contain: Parent Category/Category Name, Category Name, External ID."
        );
        return;
      }
    }

    els.btnProcess.disabled = true;
    els.btnProcess.textContent = cfg.processingLabel;
    els.btnExcel.disabled = true;
    els.btnCsv.disabled = true;
    setProgress(0);
    await tickUI();

    try {
      setProgress(4);
      await tickUI();
      await loadTemplateBytes();

      const pyodide = state.pyodide;
      const result =
        state.mode === "categories"
          ? await processCategories(pyodide)
          : await processProducts(pyodide);

      if (!result.ok) {
        throw new Error(result.error || "Conversion failed");
      }

      setProgress(100);
      await tickUI();

      mode.result = result;
      renderPreview(result);
      els.btnExcel.disabled = false;
      els.btnCsv.disabled = false;
      showMessage("info", cfg.successMessage(result.row_count));

      setTimeout(() => {
        if (mode.result === result) hideProgress();
      }, 700);
    } catch (err) {
      console.error(err);
      const msg = String(err.message || err);
      const m = msg.match(/ValueError:\s*(.+?)(?:\n|$)/);
      showMessage("error", m ? m[1] : msg);
      els.btnExcel.disabled = true;
      els.btnCsv.disabled = true;
      hideProgress();
    } finally {
      els.btnProcess.textContent = cfg.processLabel;
      refreshProcessButton();
    }
  }

  async function initPyodide() {
    try {
      const pyodide = await loadPyodide({
        indexURL: "https://cdn.jsdelivr.net/pyodide/v0.27.5/full/",
      });
      await pyodide.loadPackage("micropip");
      await pyodide.runPythonAsync(`
import micropip
await micropip.install("openpyxl")
`);
      const [convertCode, categoriesCode] = await Promise.all([
        fetch("static/convert.py?v=" + Date.now()).then((r) => {
          if (!r.ok) throw new Error("Could not load convert.py");
          return r.text();
        }),
        fetch("static/categories.py?v=" + Date.now()).then((r) => {
          if (!r.ok) throw new Error("Could not load categories.py");
          return r.text();
        }),
      ]);
      pyodide.FS.writeFile("/home/pyodide/convert.py", convertCode);
      pyodide.FS.writeFile("/home/pyodide/categories.py", categoriesCode);
      await pyodide.runPythonAsync(`
import sys
sys.path.insert(0, "/home/pyodide")
import convert
import categories
`);
      state.pyodide = pyodide;
      state.ready = true;
      refreshProcessButton();
    } catch (err) {
      console.error(err);
      showMessage(
        "error",
        `Could not start the in-browser Python engine: ${err.message || err}`
      );
    }
  }

  els.tabCategories.addEventListener("click", () => switchMode("categories"));
  els.tabProducts.addEventListener("click", () => switchMode("products"));

  els.fileInput.addEventListener("change", (e) => {
    const file = e.target.files?.[0];
    handleFile(file);
  });

  els.categoryFileInput.addEventListener("change", (e) => {
    const file = e.target.files?.[0];
    handleCategoryFile(file);
  });

  ["dragenter", "dragover"].forEach((evt) => {
    els.categoryDropZone.addEventListener(evt, (e) => {
      if (state.mode !== "products") return;
      e.preventDefault();
      e.stopPropagation();
      els.categoryUpload.classList.add("dragover");
    });
  });
  ["dragleave", "drop"].forEach((evt) => {
    els.categoryDropZone.addEventListener(evt, (e) => {
      if (state.mode !== "products") return;
      e.preventDefault();
      e.stopPropagation();
      els.categoryUpload.classList.remove("dragover");
    });
  });
  els.categoryDropZone.addEventListener("drop", (e) => {
    if (state.mode !== "products") return;
    const file = e.dataTransfer?.files?.[0];
    if (file) handleCategoryFile(file);
  });

  ["dragenter", "dragover"].forEach((evt) => {
    els.dropZone.addEventListener(evt, (e) => {
      e.preventDefault();
      els.dropZone.classList.add("dragover");
    });
  });
  ["dragleave", "drop"].forEach((evt) => {
    els.dropZone.addEventListener(evt, (e) => {
      e.preventDefault();
      els.dropZone.classList.remove("dragover");
    });
  });
  els.dropZone.addEventListener("drop", (e) => {
    const file = e.dataTransfer?.files?.[0];
    if (file) handleFile(file);
  });

  els.btnProcess.addEventListener("click", processFile);

  els.btnExcel.addEventListener("click", () => {
    const result = currentMode().result;
    if (!result?.xlsx_base64) return;
    const blob = base64ToBlob(
      result.xlsx_base64,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    downloadBlob(blob, `${downloadBaseName()}.xlsx`);
  });

  els.btnCsv.addEventListener("click", () => {
    const result = currentMode().result;
    if (!result?.csv_text) return;
    const blob = new Blob([result.csv_text], { type: "text/csv;charset=utf-8" });
    downloadBlob(blob, `${downloadBaseName()}.csv`);
  });

  applyModeUI();
  initCollapsibles();
  initPyodide();
})();
