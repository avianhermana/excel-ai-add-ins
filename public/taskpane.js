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

// Initialize Office.js
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    initApp();
  } else {
    // Running in standalone browser preview
    document.getElementById("active-range-label").innerText = "Browser Preview";
    initAppEvents();
  }
});

function initApp() {
  updateSettingsUI();
  initAppEvents();

  // Initial read of active selection
  readActiveSelection();

  // Register live listener for Excel selection change
  Excel.run(async (context) => {
    context.workbook.onSelectionChanged.add(onSelectionChange);
    await context.sync();
  }).catch((err) => {
    console.warn("Could not register selection listener:", err);
  });
}

// Event handler when selection changes in Excel
async function onSelectionChange() {
  await readActiveSelection();
}

// Read active selected range and rich worksheet metadata safely & instantly
async function readActiveSelection() {
  const rangeLabel = document.getElementById("active-range-label");
  const countLabel = document.getElementById("active-cells-count");
  const sheetBadge = document.getElementById("active-sheet-badge");

  try {
    await Excel.run(async (context) => {
      // 1. Fase Cepat: Muat address & nama sheet langsung (< 5ms)
      const range = context.workbook.getSelectedRange();
      range.load(["address", "rowCount", "columnCount"]);

      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load(["name"]);

      const sheets = context.workbook.worksheets;
      sheets.load(["items/name"]);

      await context.sync();

      // Langsung update tampilan UI agar TIDAK STUCK di "Memuat..."
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
        sampleData: []
      };

      // 2. Fase Sel: Ambil nilai sel aman (hindari crash jika user blok seluruh kolom)
      try {
        if (totalCells <= 200) {
          range.load(["values", "formulas"]);
          await context.sync();
          currentSelection.values = range.values;
          currentSelection.formulas = range.formulas;
        } else {
          // Range besar (misal 1 kolom penuh 1 juta baris): ambil 10 baris pertama saja
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

      // 3. Fase Header & Struktur Data: Baca baris 1-3 tanpa scan seluruh 200rb baris
      try {
        const topRange = sheet.getRange("A1:AZ3");
        topRange.load(["values"]);
        await context.sync();

        if (topRange.values && topRange.values.length > 0) {
          let bestRow = 0;
          let maxCount = 0;
          for (let r = 0; r < topRange.values.length; r++) {
            const count = topRange.values[r].filter(v => v !== null && String(v).trim().length > 0).length;
            if (count > maxCount) {
              maxCount = count;
              bestRow = r;
            }
          }

          sheetOverview.headers = topRange.values[bestRow]
            .map(v => String(v ?? "").trim())
            .filter(v => v.length > 0);

          if (bestRow + 1 < topRange.values.length) {
            sheetOverview.sampleData = topRange.values.slice(bestRow + 1);
          }
        }
      } catch (headerErr) {
        console.warn("Could not read headers safely:", headerErr);
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

// UI Setup & Tab Switching
function initAppEvents() {
  // Tabs
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

  // Settings Panel Open/Close
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

  // Refresh Selection Button
  const refreshBtn = document.getElementById("btn-refresh-selection");
  if (refreshBtn) refreshBtn.onclick = readActiveSelection;

  // Quick Prompt Pills
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

  // Chat Send
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

  // Settings Provider change
  const providerSelect = document.getElementById("setting-provider");
  if (providerSelect) {
    providerSelect.onchange = (e) => toggleSettingsInputs(e.target.value);
  }

  // Save Settings
  const saveSettingsBtn = document.getElementById("btn-save-settings");
  if (saveSettingsBtn) {
    saveSettingsBtn.onclick = saveSettings;
  }

  // Formula Builder Events
  initFormulaEvents();

  // Data Cleaner Events
  initCleanerEvents();
}

// ==========================================
// CHAT FUNCTIONALITY
// ==========================================
async function sendMessage() {
  const input = document.getElementById("chat-input");
  const prompt = input.value.trim();
  if (!prompt) return;

  // Add User Message
  appendChatMessage("user", prompt);
  input.value = "";

  // Add to chat history
  chatHistory.push({ role: "user", text: prompt });
  if (chatHistory.length > 12) chatHistory = chatHistory.slice(-12);

  // Loading state
  const loadingId = appendChatMessage("ai", "Sedang memproses instruksi dengan AI...");

  try {
    const payload = {
      prompt,
      excelContext: currentSelection,
      settings: settings,
      chatHistory: chatHistory.slice(0, -1) // kirim riwayat sebelumnya
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

    // Simpan balasan model ke riwayat obrolan
    if (data.reply) {
      chatHistory.push({ role: "model", text: data.reply });
      if (chatHistory.length > 12) chatHistory = chatHistory.slice(-12);
    }

    // Render AI Reply
    appendChatMessage("ai", data.reply, data.structuredAction);

    // Auto-write check
    const autoWrite = document.getElementById("chk-auto-write")?.checked;
    if (autoWrite && data.structuredAction) {
      if (data.structuredAction.action === "write" && data.structuredAction.values) {
        await writeValuesToSelection(data.structuredAction.values);
        showToast("✅ Data berhasil ditulis otomatis ke sheet!");
      } else if (data.structuredAction.action === "formula" && data.structuredAction.formula) {
        await applyFormulaToSelection(data.structuredAction.formula);
        showToast("✅ Rumus berhasil diaplikasikan ke cell aktif!");
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

    // Clean any residual raw json block from text
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

    // Interactive Action Cards (Telkomsel style)
    if (structuredAction) {
      if (structuredAction.action === "formula" && structuredAction.formula) {
        const formulaCard = document.createElement("div");
        formulaCard.className = "mt-3 p-3 rounded-lg bg-red-50/70 border border-red-200 space-y-2";
        
        formulaCard.innerHTML = `
          <div class="flex items-center justify-between text-[11px] font-semibold text-[#b91c1c]">
            <span class="flex items-center space-x-1"><span>⚡</span> <span>Rekomendasi Rumus Excel</span></span>
            <span class="text-[10px] bg-red-100 text-[#b91c1c] px-1.5 py-0.5 rounded font-mono font-bold">Formula</span>
          </div>
          <div class="p-2 bg-white rounded border border-red-200 font-mono text-[11px] text-slate-900 select-all overflow-x-auto">${escapeHtml(structuredAction.formula)}</div>
          <div class="flex items-center space-x-2 pt-1">
            <button class="btn-card-apply flex-1 bg-[#ed0226] hover:bg-[#c80220] text-white py-1.5 px-3 rounded text-[11px] font-semibold transition flex items-center justify-center space-x-1 shadow-xs">
              <span>⚡ Terapkan ke Sel Aktif</span>
            </button>
            <button class="btn-card-copy bg-white hover:bg-slate-50 text-slate-700 border border-slate-300 py-1.5 px-3 rounded text-[11px] font-medium transition">
              <span>📋 Salin</span>
            </button>
          </div>
        `;

        formulaCard.querySelector(".btn-card-apply").onclick = async () => {
          await applyFormulaToSelection(structuredAction.formula);
          showToast("✅ Rumus berhasil dimasukkan ke sel aktif!");
        };

        formulaCard.querySelector(".btn-card-copy").onclick = () => {
          navigator.clipboard.writeText(structuredAction.formula);
          showToast("📋 Rumus disalin ke clipboard!");
        };

        msgEl.appendChild(formulaCard);
      } else if (structuredAction.action === "write" && structuredAction.values) {
        const writeCard = document.createElement("div");
        writeCard.className = "mt-3 p-3 rounded-lg bg-slate-50 border border-slate-200 space-y-2";
        
        writeCard.innerHTML = `
          <div class="flex items-center justify-between text-[11px] font-semibold text-[#001a41]">
            <span class="flex items-center space-x-1"><span>📊</span> <span>Data Hasil Komputasi</span></span>
            <span class="text-[10px] bg-slate-200 text-[#001a41] px-1.5 py-0.5 rounded font-mono font-bold">Write</span>
          </div>
          <p class="text-[11px] text-slate-600">Terdapat ${structuredAction.values.length} baris data yang siap dituliskan ke worksheet.</p>
          <button class="btn-card-write w-full bg-[#001a41] hover:bg-[#002b66] text-white py-1.5 px-3 rounded text-[11px] font-semibold transition flex items-center justify-center space-x-1 shadow-xs">
            <span>📥 Tulis Data ke Sheet</span>
          </button>
        `;

        writeCard.querySelector(".btn-card-write").onclick = async () => {
          await writeValuesToSelection(structuredAction.values);
          showToast("✅ Data berhasil ditulis ke sheet!");
        };

        msgEl.appendChild(writeCard);
      }
    }
  }

  container.appendChild(msgEl);
  container.scrollTop = container.scrollHeight;
  return msgId;
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
      else if (type === "if") desc.value = "Jika nilai > 80 maka 'Lulus', jika tidak maka 'Remedial'";
      else if (type === "text") desc.value = "Gabungkan Nama Depan di A2 dan Nama Belakang di B2 dengan spasi";
    });
  });

  const btnGen = document.getElementById("btn-generate-formula");
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
          prompt: `Buatkan hanya 1 rumus Excel resmi untuk: ${desc}. Sertakan JSON action {"action": "formula", "formula": "=..."}`,
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
          // Extract regex
          const m = data.reply.match(/(=[A-Z]+[^\n`]+)/);
          if (m) formulaText = m[1];
        }

        document.getElementById("formula-result-container").classList.remove("hidden");
        document.getElementById("generated-formula-text").innerText = formulaText;
      } catch (e) {
        showToast("Gagal generate rumus: " + e.message);
      } finally {
        btnGen.innerText = "⚡ Generate Rumus Excel";
        btnGen.disabled = false;
      }
    };
  }

  const btnApply = document.getElementById("btn-apply-formula");
  if (btnApply) {
    btnApply.onclick = async () => {
      const formula = document.getElementById("generated-formula-text").innerText.trim();
      await applyFormulaToSelection(formula);
      showToast("Rumus berhasil dimasukkan ke cell aktif!");
    };
  }

  const btnCopy = document.getElementById("btn-copy-formula");
  if (btnCopy) {
    btnCopy.onclick = () => {
      const formula = document.getElementById("generated-formula-text").innerText.trim();
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

  document.getElementById("btn-format-idr")?.addEventListener("click", async () => {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.numberFormat = [["Rp#,##0;[Red](Rp#,##0);\"-\""]];
      await context.sync();
    });
    showToast("Format Rupiah (IDR) diterapkan!");
  });

  document.getElementById("btn-format-autofit")?.addEventListener("click", async () => {
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

  document.getElementById("btn-format-table")?.addEventListener("click", async () => {
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

// ==========================================
// EXCEL OFFICE.JS EXECUTION HELPERS
// ==========================================
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

async function applyFormulaToSelection(formula) {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.formulas = [[formula]];
      await context.sync();
    });
  } catch (err) {
    showToast("Gagal menerapkan rumus: " + err.message);
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
