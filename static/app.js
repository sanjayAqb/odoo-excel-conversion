/* global XLSX, loadPyodide */

(() => {
  "use strict";

  const TEMPLATE_URL = "sample_sheets/SampleOdooFormat.xlsx";

  const els = {
    dropZone: document.getElementById("dropZone"),
    fileInput: document.getElementById("fileInput"),
    uploadIdle: document.getElementById("uploadIdle"),
    uploadReady: document.getElementById("uploadReady"),
    fileName: document.getElementById("fileName"),
    fileMeta: document.getElementById("fileMeta"),
    formatPill: document.getElementById("formatPill"),
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
  };

  const state = {
    pyodide: null,
    ready: false,
    file: null,
    detectedFormat: null,
    rows: null,
    result: null,
    templateBytes: null,
  };

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

  function detectFormat(file) {
    const name = (file.name || "").toLowerCase();
    if (name.endsWith(".csv")) return "CSV";
    if (name.endsWith(".xlsx")) return "Excel (.xlsx)";
    if (name.endsWith(".xls")) return "Excel (.xls)";

    const type = (file.type || "").toLowerCase();
    if (type.includes("csv") || type === "text/plain") return "CSV";
    if (type.includes("spreadsheetml") || type.includes("openxmlformats")) return "Excel (.xlsx)";
    if (type === "application/vnd.ms-excel") return "Excel (.xls)";

    // Fallback: inspect magic bytes later in parse
    return null;
  }

  function detectFormatFromBytes(bytes, fallbackName) {
    if (bytes.length >= 4) {
      // ZIP / xlsx
      if (bytes[0] === 0x50 && bytes[1] === 0x4b) return "Excel (.xlsx)";
      // OLE compound / xls
      if (bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0) {
        return "Excel (.xls)";
      }
    }
    const name = (fallbackName || "").toLowerCase();
    if (name.endsWith(".csv")) return "CSV";
    return "CSV";
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

    let rows;
    if (format === "CSV") {
      const text = new TextDecoder("utf-8").decode(bytes);
      const workbook = XLSX.read(text, { type: "string", raw: false });
      rows = sheetToRows(workbook);
    } else {
      const workbook = XLSX.read(bytes, { type: "array", cellDates: true });
      rows = sheetToRows(workbook);
      // Refine if extension lied
      format = detectFormatFromBytes(bytes, file.name) || format;
    }

    return { format, rows };
  }

  function setFileUI(file, format) {
    els.uploadIdle.hidden = true;
    els.uploadReady.hidden = false;
    els.actions.hidden = false;
    els.formatPill.hidden = false;
    els.formatPill.textContent = `Detected: ${format}`;
    els.fileName.textContent = file.name;
    els.fileMeta.textContent = `${(file.size / 1024).toFixed(1)} KB · ${format}`;
    els.btnProcess.disabled = !state.ready;
    els.btnExcel.disabled = true;
    els.btnCsv.disabled = true;
    els.previewWrap.hidden = true;
    hideProgress();
    state.result = null;
  }

  async function handleFile(file) {
    clearMessages();
    if (!file) return;

    const allowed = /\.(csv|xlsx|xls)$/i.test(file.name);
    if (!allowed) {
      showMessage("error", "Please upload a .csv, .xlsx, or .xls file.");
      return;
    }

    try {
      const { format, rows } = await parseFile(file);
      state.file = file;
      state.detectedFormat = format;
      state.rows = rows;
      setFileUI(file, format);
      showMessage(
        "info",
        `Loaded ${rows.length} rows from ${format} file. Click “Convert to Odoo format” to map columns.`
      );
    } catch (err) {
      console.error(err);
      showMessage("error", `Failed to read file: ${err.message || err}`);
    }
  }

  function renderPreview(result) {
    const thead = els.previewTable.querySelector("thead");
    const tbody = els.previewTable.querySelector("tbody");
    thead.innerHTML = "";
    tbody.innerHTML = "";

    const headers = result.odoo_headers || [];
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

    els.previewCount.textContent =
      result.row_count > preview.length
        ? `Showing ${preview.length} of ${result.row_count} products`
        : `${result.row_count} products`;
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
    const name = state.file?.name || "odoo_products";
    const base = name.replace(/\.(csv|xlsx|xls)$/i, "") + "_odoo";
    const now = new Date();
    const pad = (n) => String(n).padStart(2, "0");
    const stamp =
      `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}` +
      `_${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
    return `${base}_${stamp}`;
  }

  async function processFile() {
    clearMessages();
    if (!state.ready || !state.rows) {
      showMessage("error", "Wait for the Python engine to finish loading, then upload a file.");
      return;
    }

    els.btnProcess.disabled = true;
    els.btnProcess.textContent = "Converting…";
    els.btnExcel.disabled = true;
    els.btnCsv.disabled = true;
    setProgress(0);
    await tickUI();

    try {
      if (!state.templateBytes) {
        setProgress(4);
        await tickUI();
        const resp = await fetch(TEMPLATE_URL);
        if (!resp.ok) throw new Error(`Could not load Odoo template (${resp.status})`);
        state.templateBytes = new Uint8Array(await resp.arrayBuffer());
      }

      const pyodide = state.pyodide;
      pyodide.FS.writeFile("/tmp/template.xlsx", state.templateBytes);
      pyodide.globals.set("SUPPLIER_ROWS_JSON", JSON.stringify(state.rows));
      pyodide.globals.set("DETECTED_FORMAT", state.detectedFormat);

      // Staged conversion so the progress % can paint between steps
      setProgress(8);
      await tickUI();
      await pyodide.runPythonAsync(`
import json, convert
_rows = json.loads(SUPPLIER_ROWS_JSON)
_headers, _records = convert.parse_supplier_rows(_rows)
`);

      setProgress(25);
      await tickUI();
      await pyodide.runPythonAsync(`
_odoo_headers, _mapped = convert.map_to_odoo(_headers, _records)
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

      const result = JSON.parse(resultJson);
      if (!result.ok) {
        throw new Error(result.error || "Conversion failed");
      }

      setProgress(100);
      await tickUI();

      state.result = result;
      renderPreview(result);
      els.btnExcel.disabled = false;
      els.btnCsv.disabled = false;
      showMessage(
        "info",
        `Mapped ${result.row_count} products into the Odoo template. Choose Excel or CSV to download.`
      );

      setTimeout(() => {
        if (state.result === result) hideProgress();
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
      els.btnProcess.disabled = false;
      els.btnProcess.textContent = "Convert to Odoo format";
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
      const pyCode = await fetch("static/convert.py?v=" + Date.now()).then((r) => {
        if (!r.ok) throw new Error("Could not load convert.py");
        return r.text();
      });
      pyodide.FS.writeFile("/home/pyodide/convert.py", pyCode);
      await pyodide.runPythonAsync(`
import sys
sys.path.insert(0, "/home/pyodide")
import convert
`);
      state.pyodide = pyodide;
      state.ready = true;
      if (state.file) els.btnProcess.disabled = false;
    } catch (err) {
      console.error(err);
      showMessage(
        "error",
        `Could not start the in-browser Python engine: ${err.message || err}`
      );
    }
  }

  // Events
  els.fileInput.addEventListener("change", (e) => {
    const file = e.target.files?.[0];
    handleFile(file);
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
    if (!state.result?.xlsx_base64) return;
    const blob = base64ToBlob(
      state.result.xlsx_base64,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    downloadBlob(blob, `${downloadBaseName()}.xlsx`);
  });

  els.btnCsv.addEventListener("click", () => {
    if (!state.result?.csv_text) return;
    const blob = new Blob([state.result.csv_text], { type: "text/csv;charset=utf-8" });
    downloadBlob(blob, `${downloadBaseName()}.csv`);
  });

  initPyodide();
})();
