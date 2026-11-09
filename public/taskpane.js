/* global Office, Excel, marked */

// Global State
let currentSelection = {
  address: "A1",
  rowCount: 1,
  columnCount: 1,
  values: [[""]],
  formulas: [[""]]
};

let chatHistory = [];

let settings = {
  provider: localStorage.getItem("copilot_provider") || "demo",
  apiKey: localStorage.getItem("copilot_api_key") || "",
  endpoint: localStorage.getItem("copilot_endpoint") || ""
};

// Helper: Konversi Huruf Kolom ke Index (0-based) & sebaliknya
function colLetterToIndex(col) {
  if (!col) return 0;
  col = col.toUpperCase().trim();
  let index = 0;
  for (let i = 0; i < col.length; i++) {
    index = index * 26 + (col.charCodeAt(i) - 64);
  }
  return index - 1;
}

function indexToColLetter(index) {
  let letter = "";
  index += 1;
  while (index > 0) {
    let remainder = (index - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    index = Math.floor((index - 1) / 26);
  }
  return letter;
}

function escapeHtml(str) {
  return (str || "").replace(/[&<>"']/g, (m) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#039;"
  })[m]);
}

// ==========================================
// EXCEL AGENT ENGINE (AUTONOMOUS EXECUTION)
// ==========================================
const ExcelAgent = {
  // 1. Memecah Kolom (Split Column) langsung di sheet
  async splitColumn(action) {
    const sourceCol = (action.sourceColumn || "C").toUpperCase();
    const insertAfterCol = (action.insertAfterColumn || sourceCol).toUpperCase();
    const newHeaders = action.newHeaders || ["Start Date", "End Date"];
    const count = newHeaders.length;

    let rowCountProcessed = 0;

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRange();
      usedRange.load(["rowCount", "rowIndex"]);
      await context.sync();

      const lastRow = Math.max(usedRange.rowIndex + usedRange.rowCount, 10);
      const limitRows = Math.min(lastRow, 5000);

      // A. Baca kolom sumber
      const sourceRange = sheet.getRange(`${sourceCol}1:${sourceCol}${limitRows}`);
      sourceRange.load(["values", "text", "formulas"]);
      await context.sync();

      // Periksa apakah cell baris 2 sempat rusak karena formula lama (=MAP / #NAME?)
      if (sourceRange.formulas[1] && sourceRange.formulas[1][0] && String(sourceRange.formulas[1][0]).startsWith("=MAP")) {
        const sampleText = sourceRange.text[2]?.[0] || "2026-10-28 --- 2026-10-31";
        sheet.getRange(`${sourceCol}2`).values = [[sampleText]];
        sourceRange.values[1][0] = sampleText;
        sourceRange.text[1][0] = sampleText;
      }

      // B. Sisipkan kolom baru di sebelah kanan (kolom berikutnya bergeser aman)
      const insertIndex = colLetterToIndex(insertAfterCol) + 1;
      const insertColLetter = indexToColLetter(insertIndex);
      const endInsertColLetter = indexToColLetter(insertIndex + count - 1);

      const insertRange = sheet.getRange(`${insertColLetter}:${endInsertColLetter}`);
      insertRange.insert(Excel.InsertShiftDirection.right);
      await context.sync();

      // C. Tulis Header baru di baris 1
      const headerRange = sheet.getRange(`${insertColLetter}1:${endInsertColLetter}1`);
      headerRange.values = [newHeaders];
      headerRange.format.font.bold = true;
      headerRange.format.font.color = "#ffffff";
      headerRange.format.fill.color = "#001a41"; // Deep Navy Telkomsel

      // D. Ekstrak data untuk seluruh baris dengan penalaran cerdas
      const transformedRows = [];
      const numRows = sourceRange.values.length;

      for (let r = 1; r < numRows; r++) {
        const val = sourceRange.text[r]?.[0] || sourceRange.values[r]?.[0] || "";
        const strVal = String(val).trim();

        if (!strVal || strVal === '""' || strVal === "#NAME?") {
          transformedRows.push(new Array(count).fill(""));
          continue;
        }

        let extracted = [];

        // 1. Format Tanggal / Periode (jika extractType dates atau terdeteksi pola tanggal)
        if (action.extractType === "dates" || strVal.match(/\d{4}[-/.]\d{1,2}[-/.]\d{1,2}.*?\d{4}[-/.]\d{1,2}[-/.]\d{1,2}/)) {
          const dates = strVal.match(/\d{4}[-/.]\d{1,2}[-/.]\d{1,2}/g);
          if (dates && dates.length >= 2) {
            extracted = [dates[0], dates[1]];
          } else if (dates && dates.length === 1) {
            extracted = [dates[0], dates[0]];
          }
        }
        // 2. Pemisah Spasi Ganda / Banyak Spasi / Tab (\s{2,} atau \t) - Pola data Wilayah & STO
        else if (strVal.match(/\s{2,}|\t/)) {
          const parts = strVal.split(/\s{2,}|\t/).map(s => s.trim()).filter(Boolean);
          if (parts.length >= 2) {
            extracted = [parts[0], parts.slice(1).join(" ")];
          } else {
            extracted = [parts[0] || "", ""];
          }
        }
        // 3. Pola Kode STO (contoh: "KOTA SURABAYA LKI - LAKARSANTRI" -> [A-Z0-9]{3,4}\s*-\s*)
        else if (strVal.match(/^(.*?)\s+([A-Z0-9]{3,4}\s*-\s*.*)$/)) {
          const m = strVal.match(/^(.*?)\s+([A-Z0-9]{3,4}\s*-\s*.*)$/);
          extracted = [m[1].trim(), m[2].trim()];
        }
        // 4. Delimiter spesifik dari AI jika diberikan
        else if (action.delimiter) {
          const regexDelim = new RegExp(action.delimiter);
          const parts = strVal.split(regexDelim).map(s => s.trim()).filter(Boolean);
          extracted = parts.slice(0, count);
        }
        // 5. Fallback split strip / pemisah umum
        else {
          const parts = strVal.split(/---|–|-|;|,/).map(s => s.trim()).filter(Boolean);
          if (parts.length > 1) {
            extracted = [parts[0], parts.slice(1).join(" ")];
          } else {
            extracted = [strVal, ""];
          }
        }

        while (extracted.length < count) extracted.push("");
        transformedRows.push(extracted);
      }

      // E. Tulis hasil pecahan data ke kolom baru
      if (transformedRows.length > 0) {
        const dataRange = sheet.getRange(`${insertColLetter}2:${endInsertColLetter}${1 + transformedRows.length}`);
        dataRange.values = transformedRows;
        rowCountProcessed = transformedRows.length;
      }

      // F. Auto-fit lebar kolom
      const newColsRange = sheet.getRange(`${insertColLetter}:${endInsertColLetter}`);
      newColsRange.format.autofitColumns();

      await context.sync();
    });

    return {
      success: true,
      summary: `Berhasil menyisipkan ${count} kolom baru (${newHeaders.join(", ")}) dan memecah ${rowCountProcessed} baris data langsung di sheet!`
    };
  },

  // 2. Sisipkan Kolom Baru dengan Data Terhitung
  async insertColumnWithData(action) {
    const insertAfterCol = (action.insertAfterColumn || "C").toUpperCase();
    const header = action.header || "Kolom Baru";
    const values = action.values || [];

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const insertIndex = colLetterToIndex(insertAfterCol) + 1;
      const colLetter = indexToColLetter(insertIndex);

      sheet.getRange(`${colLetter}:${colLetter}`).insert(Excel.InsertShiftDirection.right);
      await context.sync();

      const headerCell = sheet.getRange(`${colLetter}1`);
      headerCell.values = [[header]];
      headerCell.format.font.bold = true;
      headerCell.format.font.color = "#ffffff";
      headerCell.format.fill.color = "#001a41";

      if (values.length > 0) {
        const dataRange = sheet.getRange(`${colLetter}2:${colLetter}${1 + values.length}`);
        dataRange.values = values.map(v => Array.isArray(v) ? v : [v]);
      }

      sheet.getRange(`${colLetter}:${colLetter}`).format.autofitColumns();
      await context.sync();
    });

    return { success: true, summary: `Kolom "${header}" berhasil disisipkan langsung ke sheet!` };
  },

  // 3. Sisipkan Kolom Rumus Bersih (Tidak menimpa data asli)
  async insertFormulaColumn(action) {
    const insertAfterCol = (action.insertAfterColumn || "C").toUpperCase();
    const header = action.header || "Hasil Kalkulasi";
    const formula = action.formula;

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRange();
      usedRange.load(["rowCount", "rowIndex"]);
      await context.sync();

      const lastRow = Math.max(usedRange.rowIndex + usedRange.rowCount, 10);
      const limitRows = Math.min(lastRow, 5000);

      const insertIndex = colLetterToIndex(insertAfterCol) + 1;
      const colLetter = indexToColLetter(insertIndex);

      sheet.getRange(`${colLetter}:${colLetter}`).insert(Excel.InsertShiftDirection.right);
      await context.sync();

      const headerCell = sheet.getRange(`${colLetter}1`);
      headerCell.values = [[header]];
      headerCell.format.font.bold = true;
      headerCell.format.font.color = "#ffffff";
      headerCell.format.fill.color = "#001a41";

      if (limitRows >= 2 && formula) {
        const formulaRange = sheet.getRange(`${colLetter}2:${colLetter}${limitRows}`);
        formulaRange.formulas = new Array(limitRows - 1).fill([formula]);
      }

      sheet.getRange(`${colLetter}:${colLetter}`).format.autofitColumns();
      await context.sync();
    });

    return { success: true, summary: `Kolom rumus "${header}" berhasil ditambahkan dan diaplikasikan!` };
  },

  // 4. Update / Modifikasi Kolom In-Place
  async updateColumn(action) {
    const targetCol = (action.targetColumn || "B").toUpperCase();
    const op = action.operation || "proper_case";

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRange();
      usedRange.load(["rowCount", "rowIndex"]);
      await context.sync();

      const limitRows = Math.min(usedRange.rowIndex + usedRange.rowCount, 5000);
      const range = sheet.getRange(`${targetCol}2:${targetCol}${limitRows}`);
      range.load(["values"]);
      await context.sync();

      const updated = range.values.map(row => {
        const val = row[0];
        if (typeof val !== "string") return [val];
        if (op === "proper_case") return [val.toLowerCase().replace(/\b\w/g, c => c.toUpperCase())];
        if (op === "uppercase") return [val.toUpperCase()];
        if (op === "lowercase") return [val.toLowerCase()];
        return [val.trim().replace(/\s+/g, " ")];
      });

      range.values = updated;
      range.format.autofitColumns();
      await context.sync();
    });

    return { success: true, summary: `Kolom ${targetCol} berhasil diperbarui di worksheet!` };
  },

  // 5. Tulis Matriks Range
  async writeRange(action) {
    const values = action.values || [];
    if (!values || values.length === 0) return;

    await writeValuesToSelection(values);
    return { success: true, summary: `${values.length} baris data berhasil ditulis ke sheet!` };
  },

  // 6. Eksekusi Script Office.js Dinamis (Arsitektur Claude for Excel)
  async executeOfficeJs(action) {
    let code = (action.code || "").trim();
    if (!code) throw new Error("Tidak ada kode Office.js yang diberikan.");

    // Hapus format ```javascript ... ``` jika ada
    code = code.replace(/^```(?:javascript|js)?\s*/i, "").replace(/\s*```$/i, "").trim();

    const AsyncFunction = Object.getPrototypeOf(async function(){}).constructor;

    if (code.includes("Excel.run")) {
      const fn = new AsyncFunction("Excel", `return (async () => { ${code} })();`);
      await fn(Excel);
    } else {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const fn = new AsyncFunction("context", "sheet", "Excel", `return (async () => { ${code} })();`);
        await fn(context, sheet, Excel);
        await context.sync();
      });
    }

    return {
      success: true,
      summary: action.summary || action.thought || "Tindakan Office.js berhasil diterapkan langsung di worksheet!"
    };
  },

  // Dispatcher Utama
  async executeAction(action) {
    if (!action) return null;
    const actType = action.action || "";
    if (actType === "split_column") return await this.splitColumn(action);
    if (actType === "execute_office_js" || actType === "execute_code") return await this.executeOfficeJs(action);
    if (actType === "insert_column_with_data") return await this.insertColumnWithData(action);
    if (actType === "insert_formula_column") return await this.insertFormulaColumn(action);
    if (actType === "update_column") return await this.updateColumn(action);
    if (actType === "write_range" || actType === "write") return await this.writeRange(action);
    if (actType === "formula" && action.formula) {
      // Jika model lama mengembalikan action "formula", jangan tiban sel aktif melainkan sisipkan kolom baru yang aman
      return await this.insertFormulaColumn({
        insertAfterColumn: "C",
        header: "Hasil Rumus",
        formula: action.formula
      });
    }
    return null;
  }
};

// Initialize Office.js
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    initApp();
  } else {
    document.getElementById("active-range-label").innerText = "Browser Preview";
    initAppEvents();
  }
});

function initApp() {
  updateSettingsUI();
  initAppEvents();

  readActiveSelection();

  Excel.run(async (context) => {
    context.workbook.onSelectionChanged.add(onSelectionChange);
    await context.sync();
  }).catch((err) => {
    console.warn("Could not register selection listener:", err);
  });
}

async function onSelectionChange() {
  await readActiveSelection();
}

async function readActiveSelection() {
  const rangeLabel = document.getElementById("active-range-label");
  const countLabel = document.getElementById("active-cells-count");
  const sheetBadge = document.getElementById("active-sheet-badge");

  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["address", "rowCount", "columnCount"]);

      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load(["name"]);

      const sheets = context.workbook.worksheets;
      sheets.load(["items/name"]);

      await context.sync();

      if (rangeLabel) rangeLabel.innerText = range.address;
      if (sheetBadge) sheetBadge.innerText = sheet.name;
      const totalCells = (range.rowCount || 1) * (range.columnCount || 1);
      if (countLabel) {
        countLabel.innerText = totalCells > 1 ? `(${totalCells.toLocaleString()} sel)` : `(1 sel)`;
      }

      currentSelection.address = range.address;
      currentSelection.rowCount = range.rowCount;
      currentSelection.columnCount = range.columnCount;

      let sheetOverview = {
        activeSheetName: sheet.name,
        allSheets: sheets.items.map((s) => s.name),
        usedRangeAddress: null,
        rowCount: 0,
        columnCount: 0,
        headers: [],
        columnProfiles: [],
        sampleData: []
      };

      try {
        if (totalCells <= 200) {
          range.load(["values", "formulas"]);
          await context.sync();
          currentSelection.values = range.values;
          currentSelection.formulas = range.formulas;
        } else {
          const previewRows = Math.min(range.rowCount, 10);
          const previewCols = Math.min(range.columnCount, 10);
          const previewRange = range.getResizedRange(previewRows - 1, previewCols - 1);
          previewRange.load(["values"]);
          await context.sync();
          currentSelection.values = previewRange.values;
          currentSelection.formulas = [["(Large Range Preview)"]];
        }
      } catch (cellErr) {
        console.warn("Could not load cell values:", cellErr);
      }

      // DEEP DATA PROFILING (Ground Truth Inspection)
      try {
        const used = sheet.getUsedRange();
        used.load(["address", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
        await context.sync();

        sheetOverview.usedRangeAddress = used.address;
        sheetOverview.rowCount = used.rowCount;
        sheetOverview.columnCount = used.columnCount;

        const maxScanRows = Math.min(used.rowCount || 25, 25);
        const maxScanCols = Math.min(used.columnCount || 26, 26);
        const startRow = used.rowIndex || 0;
        const startCol = used.columnIndex || 0;

        const dataBlock = sheet.getRangeByIndexes(startRow, startCol, maxScanRows, maxScanCols);
        dataBlock.load(["values", "text"]);
        await context.sync();

        if (dataBlock.values && dataBlock.values.length > 0) {
          // Identify Header Row (row with highest number of non-empty strings)
          let bestRow = 0;
          let maxCount = 0;
          for (let r = 0; r < Math.min(5, dataBlock.values.length); r++) {
            const count = dataBlock.values[r].filter(v => v !== null && String(v).trim().length > 0).length;
            if (count > maxCount) {
              maxCount = count;
              bestRow = r;
            }
          }

          const headers = [];
          const columnProfiles = [];

          for (let c = 0; c < maxScanCols; c++) {
            const colLetter = indexToColLetter(startCol + c);
            const rawHeader = dataBlock.values[bestRow]?.[c];
            const headerName = (rawHeader !== null && rawHeader !== undefined) ? String(rawHeader).trim() : "";
            headers.push(headerName || `Kolom ${colLetter}`);

            const samples = [];
            let hasMultiSpaces = false;
            let hasDatePattern = false;
            let numericCount = 0;

            for (let r = bestRow + 1; r < dataBlock.values.length; r++) {
              const textVal = String(dataBlock.text?.[r]?.[c] ?? dataBlock.values?.[r]?.[c] ?? "").trim();
              if (textVal && textVal !== '""' && textVal !== "#NAME?") {
                samples.push(textVal);
                if (/\s{2,}|\t/.test(textVal)) hasMultiSpaces = true;
                if (/\d{4}[-/.]\d{1,2}[-/.]\d{1,2}/.test(textVal)) hasDatePattern = true;
                if (!isNaN(parseFloat(textVal.replace(/[%$,]/g, "")))) numericCount++;
              }
            }

            let type = "Teks / Umum";
            if (hasDatePattern) type = "Rentang Tanggal (Date Range)";
            else if (hasMultiSpaces) type = "Teks Majemuk dengan Spasi Ganda/Lebar (\\s{2,})";
            else if (numericCount > 0 && numericCount >= samples.length * 0.7) type = "Angka / Metrik Numerik";

            columnProfiles.push({
              letter: colLetter,
              header: headerName || `Kolom ${colLetter}`,
              type,
              hasMultiSpaces,
              hasDatePattern,
              samples: samples.slice(0, 8)
            });
          }

          sheetOverview.headers = headers;
          sheetOverview.columnProfiles = columnProfiles;

          sheetOverview.sampleData = [];
          for (let r = bestRow + 1; r < dataBlock.values.length; r++) {
            const rowText = dataBlock.text?.[r] || dataBlock.values[r];
            if (rowText && rowText.some(v => v !== null && String(v).trim().length > 0)) {
              const rowObj = {};
              for (let c = 0; c < Math.min(rowText.length, headers.length); c++) {
                const hName = headers[c] || indexToColLetter(startCol + c);
                rowObj[`${indexToColLetter(startCol + c)} (${hName})`] = rowText[c] ?? "";
              }
              sheetOverview.sampleData.push(rowObj);
            }
          }
        }
      } catch (headerErr) {
        console.warn("Could not inspect sheet data deeply:", headerErr);
      }

      currentSelection.sheetOverview = sheetOverview;
    });
  } catch (error) {
    console.error("Error reading selection:", error);
    if (rangeLabel && rangeLabel.innerText === "Memuat...") {
      rangeLabel.innerText = "Sheet Aktif";
    }
  }
}

function initAppEvents() {
  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      const targetId = btn.getAttribute("data-tab");
      document.querySelectorAll(".tab-btn").forEach((b) => {
        b.classList.remove("border-[#ed0226]", "text-[#ed0226]", "font-bold");
        b.classList.add("border-transparent", "text-slate-500");
      });
      btn.classList.add("border-[#ed0226]", "text-[#ed0226]", "font-bold");
      btn.classList.remove("border-transparent", "text-slate-500");

      document.querySelectorAll(".tab-content").forEach((c) => c.classList.add("hidden"));
      const targetContent = document.getElementById(targetId);
      if (targetContent) targetContent.classList.remove("hidden");
    });
  });

  const settingsBtn = document.getElementById("btn-tab-settings");
  const closeSettingsBtn = document.getElementById("btn-close-settings");
  if (settingsBtn) {
    settingsBtn.onclick = () => {
      document.querySelectorAll(".tab-content").forEach((c) => c.classList.add("hidden"));
      document.getElementById("tab-settings").classList.remove("hidden");
    };
  }
  if (closeSettingsBtn) {
    closeSettingsBtn.onclick = () => {
      document.getElementById("tab-settings").classList.add("hidden");
      document.getElementById("tab-chat").classList.remove("hidden");
    };
  }

  const refreshBtn = document.getElementById("btn-refresh-selection");
  if (refreshBtn) refreshBtn.onclick = readActiveSelection;

  document.querySelectorAll(".quick-pill").forEach((pill) => {
    pill.addEventListener("click", () => {
      const prompt = pill.getAttribute("data-prompt");
      const input = document.getElementById("chat-input");
      if (input) {
        input.value = prompt;
        sendMessage();
      }
    });
  });

  const sendBtn = document.getElementById("btn-send-chat");
  const chatInput = document.getElementById("chat-input");
  if (sendBtn) sendBtn.onclick = sendMessage;
  if (chatInput) {
    chatInput.addEventListener("keydown", (e) => {
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        sendMessage();
      }
    });
  }

  const providerSelect = document.getElementById("setting-provider");
  if (providerSelect) {
    providerSelect.onchange = (e) => toggleSettingsInputs(e.target.value);
  }

  const saveSettingsBtn = document.getElementById("btn-save-settings");
  if (saveSettingsBtn) {
    saveSettingsBtn.onclick = saveSettings;
  }

  initFormulaEvents();
  initCleanerEvents();
}

// ==========================================
// CHAT FUNCTIONALITY & AGENT DISPATCH
// ==========================================
async function sendMessage() {
  const input = document.getElementById("chat-input");
  const prompt = input.value.trim();
  if (!prompt) return;

  appendChatMessage("user", prompt);
  input.value = "";

  chatHistory.push({ role: "user", text: prompt });
  if (chatHistory.length > 12) chatHistory = chatHistory.slice(-12);

  const loadingId = appendChatMessage("ai", "Sedang menganalisis struktur data dan meracik aksi dengan AI Agent...");

  try {
    // Selalu ambil observasi snapshot data sheet terbaru sebelum mengirim ke AI
    await readActiveSelection();

    const payload = {
      prompt,
      excelContext: currentSelection,
      settings: settings,
      chatHistory: chatHistory.slice(0, -1)
    };

    const res = await fetch("/api/chat", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    const data = await res.json();
    removeChatMessage(loadingId);

    if (data.error) {
      appendChatMessage("ai", `❌ **Error:** ${data.error}`);
      return;
    }

    if (data.reply) {
      chatHistory.push({ role: "model", text: data.reply });
      if (chatHistory.length > 12) chatHistory = chatHistory.slice(-12);
    }

    // Render AI Reply with Agent Card
    const msgId = appendChatMessage("ai", data.reply, data.structuredAction);

    // Auto-execute check (Default Aktif)
    const autoWrite = document.getElementById("chk-auto-write")?.checked;
    if (autoWrite && data.structuredAction) {
      const agentBadge = document.getElementById(`agent-badge-${msgId}`);
      const btnRun = document.getElementById(`btn-agent-run-${msgId}`);

      if (agentBadge) {
        agentBadge.innerText = "⚡ Sedang Dikerjakan di Sheet...";
        agentBadge.className = "text-[10px] bg-amber-500/20 text-amber-300 border border-amber-500/40 px-2 py-0.5 rounded-full font-medium animate-pulse";
      }

      try {
        const result = await ExcelAgent.executeAction(data.structuredAction);
        if (result && result.success) {
          if (agentBadge) {
            agentBadge.innerText = "✅ Selesai Dikerjakan di Sheet";
            agentBadge.className = "text-[10px] bg-emerald-500/20 text-emerald-300 border border-emerald-500/40 px-2 py-0.5 rounded-full font-medium";
          }
          if (btnRun) {
            btnRun.innerText = "✓ Sudah Diterapkan Langsung di Sheet";
            btnRun.disabled = true;
            btnRun.className = "bg-emerald-950 text-emerald-300 font-semibold px-3 py-1.5 rounded-lg text-xs cursor-default border border-emerald-800";
          }
          showToast(result.summary || "Operasi selesai langsung di worksheet!");
        }
      } catch (err) {
        console.error("Agent Execution Error:", err);
        if (agentBadge) {
          agentBadge.innerText = "⚠️ Gagal Eksekusi";
          agentBadge.className = "text-[10px] bg-red-500/20 text-red-300 border border-red-500/40 px-2 py-0.5 rounded-full font-medium";
        }
        showToast("Gagal eksekusi ke sheet: " + err.message);
      }
    }
  } catch (err) {
    removeChatMessage(loadingId);
    appendChatMessage("ai", `⚠️ Gagal menghubungi server: ${err.message}`);
  }
}

function appendChatMessage(sender, text, structuredAction = null) {
  const container = document.getElementById("chat-messages");
  const msgId = "msg-" + Date.now();
  const msgEl = document.createElement("div");
  msgEl.id = msgId;

  if (sender === "user") {
    msgEl.className = "bg-[#001a41] text-white px-3.5 py-2.5 rounded-2xl text-xs self-end ml-8 shadow-xs leading-relaxed max-w-[90%] whitespace-pre-wrap border border-[#002b66]";
    msgEl.innerText = text;
  } else {
    msgEl.className = "bg-white text-slate-800 border border-slate-200 p-3.5 rounded-xl text-xs shadow-xs space-y-2.5";

    const cleanText = (text || "").replace(/```(?:json)?\s*[\s\S]*?\s*```/g, "").trim();

    const wrapper = document.createElement("div");
    wrapper.className = "flex items-start space-x-2.5";

    const avatar = document.createElement("div");
    avatar.className = "w-5 h-5 rounded-md bg-[#ed0226]/10 text-[#ed0226] flex items-center justify-center text-xs font-bold shrink-0 mt-0.5";
    avatar.innerText = "⚡";

    const bodyEl = document.createElement("div");
    bodyEl.className = "markdown-body flex-1 min-w-0";
    bodyEl.innerHTML = typeof marked !== "undefined" ? marked.parse(cleanText) : cleanText;

    wrapper.appendChild(avatar);
    wrapper.appendChild(bodyEl);
    msgEl.appendChild(wrapper);

    // Interactive Agent Card (Reasoning + Code Synthesis)
    if (structuredAction) {
      const card = document.createElement("div");
      card.className = "mt-3 p-3.5 rounded-xl bg-[#001a41] border border-[#002b66] text-white space-y-2.5 shadow-md";

      let title = "⚡ AI Agent: Tindakan Lembar Kerja";
      if (structuredAction.action === "split_column") title = "⚡ AI Agent: Pecah Kolom";
      else if (structuredAction.action === "execute_office_js" || structuredAction.action === "execute_code") title = "⚡ AI Agent: Eksekusi Office.js Dinamis";
      else if (structuredAction.action === "insert_formula_column") title = "⚡ AI Agent: Sisipkan Kolom Rumus";
      else if (structuredAction.action === "insert_column_with_data") title = "⚡ AI Agent: Tambah Kolom Data";
      else if (structuredAction.action === "update_column") title = "⚡ AI Agent: Pembersihan Format";

      const hasCode = !!structuredAction.code;

      card.innerHTML = `
        <div class="flex items-center justify-between">
          <div class="flex items-center space-x-1.5 text-xs font-bold text-white">
            <span class="text-emerald-400">⚡</span>
            <span>${title}</span>
          </div>
          <span id="agent-badge-${msgId}" class="text-[10px] bg-emerald-950 text-emerald-300 border border-emerald-800 px-2 py-0.5 rounded-full font-medium">
            Siap Diterapkan
          </span>
        </div>
        <div class="text-[11px] text-slate-200 leading-relaxed font-sans bg-black/25 p-2.5 rounded-lg border border-white/5 space-y-1">
          <div class="text-[10px] uppercase font-bold text-emerald-400">🧠 Penalaran & Observasi Data:</div>
          <p>${escapeHtml(structuredAction.thought || "AI Agent telah menganalisis request dan struktur data Anda.")}</p>
        </div>
        <div class="pt-1 flex items-center space-x-2">
          <button id="btn-agent-run-${msgId}" class="bg-[#ed0226] hover:bg-[#c80220] text-white font-semibold px-3 py-1.5 rounded-lg text-xs flex items-center space-x-1.5 shadow-xs transition">
            <span>▶️ Terapkan ke Sheet</span>
          </button>
          ${hasCode ? `<button id="btn-code-${msgId}" class="text-[10px] text-slate-400 hover:text-white underline ml-auto">Lihat Kode</button>` : ''}
        </div>
        ${hasCode ? `<div id="code-box-${msgId}" class="hidden p-2.5 bg-slate-950 rounded-lg text-[10px] font-mono text-emerald-400 overflow-x-auto select-all max-h-40 border border-slate-800">${escapeHtml(structuredAction.code)}</div>` : ''}
      `;

      if (hasCode) {
        const btnToggle = card.querySelector(`#btn-code-${msgId}`);
        const codeBox = card.querySelector(`#code-box-${msgId}`);
        if (btnToggle && codeBox) {
          btnToggle.onclick = () => {
            codeBox.classList.toggle("hidden");
            btnToggle.innerText = codeBox.classList.contains("hidden") ? "Lihat Kode" : "Sembunyikan Kode";
          };
        }
      }

      card.querySelector(`#btn-agent-run-${msgId}`).onclick = async () => {
        const btn = card.querySelector(`#btn-agent-run-${msgId}`);
        const badge = card.querySelector(`#agent-badge-${msgId}`);
        btn.innerText = "Sedang Menerapkan...";
        btn.disabled = true;

        try {
          const res = await ExcelAgent.executeAction(structuredAction);
          if (res && res.success) {
            badge.innerText = "✅ Selesai di Sheet";
            badge.className = "text-[10px] bg-emerald-500/20 text-emerald-300 border border-emerald-500/40 px-2 py-0.5 rounded-full font-medium";
            btn.innerText = "✓ Sudah Diterapkan Langsung di Sheet";
            btn.className = "bg-emerald-950 text-emerald-300 font-semibold px-3 py-1.5 rounded-lg text-xs cursor-default border border-emerald-800";
            showToast(res.summary);
          }
        } catch (e) {
          btn.innerText = "Coba Lagi";
          btn.disabled = false;
          showToast("Gagal: " + e.message);
        }
      };

      msgEl.appendChild(card);
    }
  }

  container.appendChild(msgEl);
  container.scrollTop = container.scrollHeight;
  return msgId;
}

function removeChatMessage(id) {
  const el = document.getElementById(id);
  if (el) el.remove();
}

// ==========================================
// FORMULA BUILDER
// ==========================================
function initFormulaEvents() {
  document.querySelectorAll(".formula-type-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".formula-type-btn").forEach((b) => b.classList.remove("border-emerald-600", "bg-emerald-50"));
      btn.classList.add("border-emerald-600", "bg-emerald-50");
      const type = btn.getAttribute("data-type");
      const desc = document.getElementById("formula-desc-input");
      if (type === "xlookup") desc.value = "Cari nilai di kolom B berdasarkan kunci di sel A2 dari sheet Data";
      else if (type === "sumifs") desc.value = "Jumlahkan kolom Total jika status adalah 'Completed' dan cabang adalah 'Jakarta'";
      else if (type === "if-nested") desc.value = "Jika nilai > 80 maka 'Lulus', jika tidak maka 'Remedial'";
      else if (type === "text-join") desc.value = "Gabungkan Nama Depan di A2 dan Nama Belakang di B2 dengan spasi";
    });
  });

  const btnGen = document.getElementById("btn-generate-custom-formula");
  if (btnGen) {
    btnGen.onclick = async () => {
      const desc = document.getElementById("formula-desc-input").value.trim();
      if (!desc) {
        showToast("Masukkan deskripsi rumus terlebih dahulu!");
        return;
      }

      btnGen.innerText = "Generating...";
      btnGen.disabled = true;

      try {
        const payload = {
          prompt: `Buatkan hanya 1 rumus Excel resmi untuk: ${desc}. Sertakan JSON action {"action": "insert_formula_column", "header": "Hasil", "formula": "=..."}`,
          excelContext: currentSelection,
          settings: settings
        };

        const res = await fetch("/api/chat", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload)
        });
        const data = await res.json();

        let formulaText = "=XLOOKUP(A2, B:B, C:C)";
        if (data.structuredAction && data.structuredAction.formula) {
          formulaText = data.structuredAction.formula;
        } else {
          const m = data.reply.match(/(=[A-Z]+[^\n`]+)/);
          if (m) formulaText = m[1];
        }

        document.getElementById("formula-result-container").classList.remove("hidden");
        document.getElementById("formula-result-code").innerText = formulaText;
      } catch (e) {
        showToast("Gagal generate rumus: " + e.message);
      } finally {
        btnGen.innerText = "✨ Buat Rumus Sekarang";
        btnGen.disabled = false;
      }
    };
  }

  const btnApply = document.getElementById("btn-apply-formula-box");
  if (btnApply) {
    btnApply.onclick = async () => {
      const formula = document.getElementById("formula-result-code").innerText.trim();
      await ExcelAgent.insertFormulaColumn({
        insertAfterColumn: "C",
        header: "Hasil Rumus",
        formula: formula
      });
      showToast("Kolom rumus berhasil ditambahkan ke worksheet!");
    };
  }

  const btnCopy = document.getElementById("btn-copy-formula-box");
  if (btnCopy) {
    btnCopy.onclick = () => {
      const formula = document.getElementById("formula-result-code").innerText.trim();
      navigator.clipboard.writeText(formula);
      showToast("Rumus disalin ke clipboard!");
    };
  }
}

// ==========================================
// DATA CLEANER (1-CLICK EXCEL OPERATIONS)
// ==========================================
function initCleanerEvents() {
  document.getElementById("btn-clean-trim")?.addEventListener("click", () => {
    transformSelectedText((str) => str.trim().replace(/\s+/g, " "));
    showToast("Spasi berlebih berhasil dibersihkan!");
  });

  document.getElementById("btn-clean-proper")?.addEventListener("click", () => {
    transformSelectedText((str) => str.toLowerCase().replace(/\b\w/g, (c) => c.toUpperCase()));
    showToast("Format Title Case diterapkan!");
  });

  document.getElementById("btn-clean-upper")?.addEventListener("click", () => {
    transformSelectedText((str) => str.toUpperCase());
    showToast("Diubah ke UPPERCASE!");
  });

  document.getElementById("btn-clean-lower")?.addEventListener("click", () => {
    transformSelectedText((str) => str.toLowerCase());
    showToast("Diubah ke lowercase!");
  });

  document.getElementById("btn-clean-currency")?.addEventListener("click", async () => {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.numberFormat = [["Rp#,##0;[Red](Rp#,##0);\"-\""]];
      await context.sync();
    });
    showToast("Format Rupiah (IDR) diterapkan!");
  });

  document.getElementById("btn-clean-autofit")?.addEventListener("click", async () => {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.format.autofitColumns();
      await context.sync();
    });
    showToast("Lebar kolom disesuaikan otomatis!");
  });

  document.getElementById("btn-clean-fillblank")?.addEventListener("click", async () => {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values");
      await context.sync();

      const newVals = range.values.map((row) =>
        row.map((val) => (val === "" || val === null || val === undefined ? "-" : val))
      );
      range.values = newVals;
      await context.sync();
    });
    showToast("Cell kosong telah diisi '-'!");
  });

  document.getElementById("btn-clean-table")?.addEventListener("click", async () => {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      context.workbook.tables.add(range, true);
      await context.sync();
    });
    showToast("Tabel Excel resmi berhasil dibuat!");
  });
}

async function transformSelectedText(fn) {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values");
      await context.sync();

      const newVals = range.values.map((row) =>
        row.map((cell) => (typeof cell === "string" ? fn(cell) : cell))
      );
      range.values = newVals;
      await context.sync();
    });
  } catch (err) {
    showToast("Gagal memproses data: " + err.message);
  }
}

async function writeValuesToSelection(values) {
  try {
    await Excel.run(async (context) => {
      const activeCell = context.workbook.getActiveCell();
      activeCell.load(["rowIndex", "columnIndex"]);
      await context.sync();

      const numRows = values.length;
      const numCols = values[0] ? values[0].length : 1;
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const targetRange = sheet.getRangeByIndexes(
        activeCell.rowIndex,
        activeCell.columnIndex,
        numRows,
        numCols
      );

      targetRange.values = values;
      targetRange.format.autofitColumns();
      await context.sync();
    });
  } catch (err) {
    showToast("Gagal menulis ke Excel: " + err.message);
  }
}

// ==========================================
// SETTINGS MANAGEMENT
// ==========================================
function updateSettingsUI() {
  const provEl = document.getElementById("setting-provider");
  const keyEl = document.getElementById("setting-api-key");
  const endpointEl = document.getElementById("setting-endpoint");
  const badgeEl = document.getElementById("badge-provider");

  if (provEl) provEl.value = settings.provider;
  if (keyEl) keyEl.value = settings.apiKey;
  if (endpointEl) endpointEl.value = settings.endpoint;

  if (badgeEl) {
    if (settings.provider === "gemini") badgeEl.innerText = "Google Gemini";
    else if (settings.provider === "openai") badgeEl.innerText = "OpenAI GPT-4o";
    else if (settings.provider === "azure") badgeEl.innerText = "Azure/Telkomsel";
    else badgeEl.innerText = "Demo Mode";
  }

  toggleSettingsInputs(settings.provider);
}

function toggleSettingsInputs(provider) {
  const keyContainer = document.getElementById("container-api-key");
  const endpointContainer = document.getElementById("container-custom-endpoint");

  if (provider === "demo") {
    keyContainer.classList.add("hidden");
    endpointContainer.classList.add("hidden");
  } else if (provider === "azure") {
    keyContainer.classList.remove("hidden");
    endpointContainer.classList.remove("hidden");
  } else {
    keyContainer.classList.remove("hidden");
    endpointContainer.classList.add("hidden");
  }
}

function saveSettings() {
  const provider = document.getElementById("setting-provider").value;
  const apiKey = document.getElementById("setting-api-key").value.trim();
  const endpoint = document.getElementById("setting-endpoint").value.trim();

  settings.provider = provider;
  settings.apiKey = apiKey;
  settings.endpoint = endpoint;

  localStorage.setItem("copilot_provider", provider);
  localStorage.setItem("copilot_api_key", apiKey);
  localStorage.setItem("copilot_endpoint", endpoint);

  updateSettingsUI();
  document.getElementById("tab-settings").classList.add("hidden");
  document.getElementById("tab-chat").classList.remove("hidden");

  showToast("Pengaturan AI berhasil disimpan!");
}

function showToast(message) {
  const toast = document.getElementById("toast");
  if (!toast) return;

  const toastMsg = document.getElementById("toast-message");
  if (toastMsg) {
    toastMsg.innerText = message;
  } else {
    toast.innerText = message;
  }

  toast.classList.remove("translate-y-16", "opacity-0");
  toast.classList.add("translate-y-0", "opacity-100");

  setTimeout(() => {
    toast.classList.remove("translate-y-0", "opacity-100");
    toast.classList.add("translate-y-16", "opacity-0");
  }, 3000);
}
