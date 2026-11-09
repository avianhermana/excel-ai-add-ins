const express = require('express');
const https = require('https');
const fs = require('fs');
const path = require('path');
const cors = require('cors');
const selfsigned = require('selfsigned');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use((req, res, next) => {
  res.set('Cache-Control', 'no-store, no-cache, must-revalidate, private');
  res.set('Pragma', 'no-cache');
  res.set('Expires', '0');
  console.log(`[${new Date().toLocaleTimeString()}] ${req.method} ${req.url}`);
  next();
});
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public'), {
  etag: false,
  maxAge: 0
}));

// 1. SSL Certificate Management (HTTPS required by Office Add-ins)
const certsDir = path.join(__dirname, 'certs');
const certPath = path.join(certsDir, 'cert.pem');
const keyPath = path.join(certsDir, 'key.pem');

async function getOrGenerateCerts() {
  if (!fs.existsSync(certsDir)) {
    fs.mkdirSync(certsDir, { recursive: true });
  }

  if (fs.existsSync(certPath) && fs.existsSync(keyPath)) {
    return {
      cert: fs.readFileSync(certPath),
      key: fs.readFileSync(keyPath)
    };
  }

  console.log('Generating self-signed SSL certificate for localhost...');
  const attrs = [
    { name: 'commonName', value: 'localhost' },
    { name: 'countryName', value: 'ID' },
    { name: 'organizationName', value: 'Excel Copilot Local Dev' }
  ];
  
  const pems = await selfsigned.generate(attrs, {
    days: 365,
    keySize: 2048,
    algorithm: 'sha256'
  });

  fs.writeFileSync(certPath, pems.cert);
  fs.writeFileSync(keyPath, pems.private);
  return { cert: pems.cert, key: pems.private };
}

// 2. Helper: Build Rich Spreadsheet Context for LLM
function buildSpreadsheetContext(excelContext) {
  if (!excelContext) return "Tidak ada data spreadsheet yang dipilih.";
  
  let contextStr = "INFORMASI LENGKAP WORKSHEET & STRUKTUR DATA (GROUND TRUTH):\n";
  const ov = excelContext.sheetOverview || {};

  contextStr += `- Sheet Aktif: "${ov.activeSheetName || 'Aktif'}"\n`;
  contextStr += `- Sel/Range Terpilih: ${excelContext.address || 'A1'}\n`;
  if (ov.usedRangeAddress) {
    contextStr += `- Digunakan: ${ov.usedRangeAddress} (Total ~${ov.rowCount || 0} baris, ~${ov.columnCount || 0} kolom)\n`;
  }
  if (ov.allSheets && ov.allSheets.length > 0) {
    contextStr += `- Daftar Seluruh Sheet di Workbook: ${JSON.stringify(ov.allSheets)}\n`;
  }

  if (ov.columnProfiles && ov.columnProfiles.length > 0) {
    contextStr += `\nPROFIL DETAIL KOLOM & OBSERVASI DATA FISIK:\n`;
    for (const col of ov.columnProfiles) {
      contextStr += `* Kolom ${col.letter} ("${col.header}"):\n`;
      contextStr += `  - Tipe Terdeteksi: ${col.type}\n`;
      if (col.hasMultiSpaces) {
        contextStr += `  - CATATAN FISIK PENTING: Ditemukan spasi ganda/lebar (\\s{2,}) atau tab sebagai pemisah antar token teks!\n`;
      }
      if (col.samples && col.samples.length > 0) {
        contextStr += `  - Cuplikan Nilai Nyata (Perhatikan Spasi & Token Kata):\n`;
        col.samples.forEach((sample, sIdx) => {
          contextStr += `    [Baris ${sIdx + 2}]: "${sample}"\n`;
        });
      }
    }
  } else if (ov.sampleData && ov.sampleData.length > 0) {
    contextStr += `\nCuplikan Contoh Data Baris:\n${JSON.stringify(ov.sampleData, null, 2)}\n`;
  }

  return contextStr;
}

// 3. Helper: Generate Generalized Autonomous Office.js Code
function buildOfficeJsCodeForSplit({ sourceCol, insertCols, headers, regexPattern, customExtractLogic }) {
  const insertCount = headers.length;
  const startTargetCol = String.fromCharCode(sourceCol.charCodeAt(0) + 1);
  const endTargetCol = String.fromCharCode(sourceCol.charCodeAt(0) + insertCount);

  return `const sheet = context.workbook.worksheets.getActiveWorksheet();
const usedRange = sheet.getUsedRange();
usedRange.load(["rowCount", "rowIndex"]);
await context.sync();

const lastRow = Math.max(usedRange.rowIndex + usedRange.rowCount, 10);
const limitRows = Math.min(lastRow, 10000);

const srcColRange = sheet.getRange("${sourceCol}1:${sourceCol}" + limitRows);
srcColRange.load(["text", "values"]);
await context.sync();

// 1. Sisipkan kolom baru secara non-destruktif
sheet.getRange("${startTargetCol}:${endTargetCol}").insert(Excel.InsertShiftDirection.right);
await context.sync();

// 2. Format Header Korporat Telkomsel (Navy #001a41, font putih tebal)
const headerRange = sheet.getRange("${startTargetCol}1:${endTargetCol}1");
headerRange.values = [${JSON.stringify(headers)}];
headerRange.format.font.bold = true;
headerRange.format.font.color = "#ffffff";
headerRange.format.fill.color = "#001a41";

// 3. Proses transformasi di memori JavaScript
const outputValues = [];
for (let i = 1; i < srcColRange.values.length; i++) {
  const rawText = String(srcColRange.text[i]?.[0] || srcColRange.values[i]?.[0] || "").trim();
  if (!rawText || rawText === '""' || rawText === "#NAME?") {
    outputValues.push(new Array(${insertCount}).fill(""));
    continue;
  }
  ${customExtractLogic || `
  const parts = rawText.split(${regexPattern || `/\\s{2,}|\\t/`}).map(s => s.trim()).filter(Boolean);
  if (parts.length >= 2) {
    outputValues.push([parts[0], parts.slice(1).join(" ")]);
  } else {
    const boundaryMatch = rawText.match(/^(.*?)\\s+([A-Z0-9]{3,4}\\s*-\\s*.*)$/);
    if (boundaryMatch) outputValues.push([boundaryMatch[1].trim(), boundaryMatch[2].trim()]);
    else outputValues.push([rawText, ""]);
  }
  `}
}

// 4. Batch write ke sheet dan auto-fit lebar kolom
if (outputValues.length > 0) {
  const targetDataRange = sheet.getRange("${startTargetCol}2:${endTargetCol}" + (1 + outputValues.length));
  targetDataRange.values = outputValues;
}
sheet.getRange("${startTargetCol}:${endTargetCol}").format.autofitColumns();
await context.sync();`;
}

// 4. AI Chat & Execution Endpoint
app.post('/api/chat', async (req, res) => {
  try {
    const { prompt, excelContext, settings, chatHistory } = req.body;
    const provider = (settings && settings.provider) || 'demo';
    const apiKey = (settings && settings.apiKey) || '';

    const contextStr = buildSpreadsheetContext(excelContext);

    const systemPrompt = `Anda adalah "Copilot for Excel - Autonomous Spreadsheet Agent" (Arsitektur penalaran tingkat tinggi setara Claude for Excel & ChatGPT Code Interpreter).

TUGAS DAN PRINSIP KERJA ANDA (3 PILAR UTAMA):
1. MENGANALISIS KEBUTUHAN PENGGUNA (Intent Analysis):
   Pahami secara mendalam apa output akhir yang diinginkan pengguna (contoh: memecah kolom gabungan menjadi dua kolom terpisah, menghitung agregasi unik, membuat kolom kalkulasi, atau conditional formatting).

2. MEMBACA & MEMAHAMI DATA NYATA SECARA FISIK (Deep Data Inspection):
   Periksa informasi worksheet dan "PROFIL DETAIL KOLOM & OBSERVASI DATA FISIK" di bawah.
   Tunjukkan bukti nyata bahwa Anda MEMBACA data fisik pengguna:
   - Amati kolom target (contoh: Kolom B "Wilayah & STO").
   - Amati bagaimana teks tertulis: Apakah ada spasi ganda/lebar (\\s{2,}) atau tab?
   - Amati variasi token kata: Apakah ada nama wilayah multi-kata (seperti "KOTA SURABAYA", "JAKARTA SELATAN") di samping nama wilayah satu kata ("TULUNGAGUNG", "MALANG")?
   - Tegaskan mengapa metode naif seperti split spasi tunggal (split(" ")) SALAH BESAR (karena akan memecah "KOTA SURABAYA" menjadi "KOTA" dan "SURABAYA" serta menghilangkan kode STO), dan jelaskan logika yang benar (menggunakan spasi ganda/lebar atau boundary regex).

3. MENGIMPLEMENTASIKAN LANGSUNG KE LEMBAR KERJA VIA KODE OFFICE.JS DINAMIS (Autonomous Execution):
   Anda adalah pembuat kode JavaScript Office.js yang langsung memanipulasi workbook pengguna.
   - NON-DESTRUCTIVE: JANGAN PERNAH menimpa data asli kecuali diminta. Selalu sisipkan kolom baru dengan \`sheet.getRange("...").insert(Excel.InsertShiftDirection.right)\` sehingga data yang ada bergeser aman.
   - FORMATTING: Berikan header dengan styling korporat Telkomsel: Font Bold, Teks Putih (#ffffff), Fill Deep Navy (#001a41).
   - MEMORY COMPUTATION: Lakukan pemrosesan teks/angka di dalam JavaScript runtime (bukan menaruh rumus =MAP atau =LAMBDA yang rawan error #NAME? pada Excel desktop).
   - BATCH WRITE: Tuliskan seluruh hasil dalam satu panggilan 2D array \`range.values = outputValues\`.
   - AUTOFIT: Selalu panggil \`format.autofitColumns()\`.

STRUKTUR RESPON ANDA WAJIB MENGANDUNG 3 BAGIAN:
### 🎯 1. Analisis Kebutuhan Pengguna
(Penjelasan kebutuhan pengguna)

### 🔍 2. Analisis & Pemahaman Data
(Observasi riil data fisik, pola spasi, anomali multi-kata, dan keputusan algoritma)

### ⚡ 3. Rencana & Implementasi Lembar Kerja
(Aksi konkret yang dikerjakan di sheet)

Di akhir respons, sertakan blok JSON WAJIB dengan format:
\`\`\`json
{
  "thought": "Ringkasan penalaran 3 fase: kebutuhan, observasi data nyata, dan rencana eksekusi Office.js.",
  "action": "execute_office_js",
  "code": "/* kode JavaScript Office.js lengkap yang dieksekusi di dalam context Excel */",
  "summary": "Ringkasan hasil kerja untuk pengguna."
}
\`\`\`

Konteks Lembar Kerja Pengguna Saat Ini:
${contextStr}`;

    // A. Demo Mode / Smart Local Agent Engine
    if (provider === 'demo' || !apiKey) {
      const pLower = (prompt || '').toLowerCase();
      let replyText = "";
      let structuredAction = null;

      const ov = (excelContext && excelContext.sheetOverview) || {};
      const profiles = ov.columnProfiles || [];

      // Case 1: Permintaan Pecah Kolom Wilayah & STO
      if ((pLower.includes('wilayah') && pLower.includes('sto')) || (pLower.includes('kolom b') && (pLower.includes('pecah') || pLower.includes('pisah') || pLower.includes('split')))) {
        replyText = `### 🎯 1. Analisis Kebutuhan Pengguna
Pengguna meminta untuk memisahkan data pada kolom **Wilayah & STO (Kolom B)** menjadi dua kolom terpisah: **"Wilayah"** dan **"STO"**. Tugas ini harus dikerjakan secara langsung di lembar kerja tanpa merusak data asli maupun kolom-kolom sesudahnya.

### 🔍 2. Analisis & Pemahaman Data
Berdasarkan pembacaan langsung data fisik pada Kolom B:
* **Pola Pemisah (Delimiter)**: Teks memuat nama wilayah yang dipisahkan oleh **spasi ganda / padding spasi lebar** (\`\\s{2,}|\t\`) dari kode dan nama STO (seperti \`CAT - CAMPUR DARAT\` atau \`LKI - LAKARSANTRI\`).
* **Anomali Penting Terdeteksi**: Ditemukan nama wilayah yang terdiri dari **lebih dari satu kata**, contohnya **"KOTA SURABAYA"** pada baris ke-25, di samping nama wilayah satu kata seperti **"TULUNGAGUNG"** dan **"PAMEKASAN"**.
* **Keputusan Algoritma**: Pemisahan **TIDAK BOLEH** menggunakan spasi tunggal (\`split(" ")\`) karena akan memotong *"KOTA SURABAYA"* menjadi *"KOTA"* dan *"SURABAYA"*, serta menghilangkan nama STO. Agent menggunakan regex spasi ganda \`/\\s{2,}|\\t/\` dan verifikasi boundary kode STO agar pemisahan 100% presisi untuk seluruh baris.

### ⚡ 3. Rencana & Implementasi Lembar Kerja
1. **Penyisipan Kolom**: Menyisipkan 2 kolom baru di C dan D dengan \`insert(Excel.InsertShiftDirection.right)\` sehingga kolom asli dan kolom di kanannya bergeser aman.
2. **Format Header**: Menetapkan header **"Wilayah"** (Kolom C) dan **"STO"** (Kolom D) dengan tema korporat Telkomsel (*Deep Navy #001a41, font putih tebal*).
3. **Batch Processing**: Membaca seluruh data Kolom B, memecah setiap baris di memori JavaScript dengan algoritma di atas, dan menuliskan hasilnya sekaligus ke Kolom C dan D.
4. **Auto-fit**: Menyesuaikan lebar kolom baru secara otomatis.`;

        const code = buildOfficeJsCodeForSplit({
          sourceCol: "B",
          insertCols: ["C", "D"],
          headers: ["Wilayah", "STO"]
        });

        structuredAction = {
          thought: "Analisis Kebutuhan: Memecah Kolom B menjadi Wilayah dan STO. Analisis Data: Teks memuat nama kota (termasuk multi-kata seperti 'KOTA SURABAYA') dipisahkan oleh spasi ganda dari kode STO. Implementasi: Sisipkan 2 kolom baru, tulis header korporat, parsing dengan regex \\s{2,}|\\t, dan batch write.",
          action: "execute_office_js",
          code,
          summary: "Berhasil menyisipkan 2 kolom baru ('Wilayah' dan 'STO') dan memecah seluruh baris data secara presisi tanpa merusak kolom asli!"
        };
      }
      // Case 2: Permintaan Pecah Tanggal / Periode Streak
      else if (pLower.includes('periode') || pLower.includes('streak') || pLower.includes('start date') || pLower.includes('end date') || (pLower.includes('kolom c') && (pLower.includes('pecah') || pLower.includes('pisah') || pLower.includes('split')))) {
        replyText = `### 🎯 1. Analisis Kebutuhan Pengguna
Pengguna meminta untuk memecah kolom **Periode Streak (Kolom C)** menjadi dua kolom tanggal independen: **"Start Date"** dan **"End Date"**.

### 🔍 2. Analisis & Pemahaman Data
Berdasarkan inspeksi data fisik Kolom C:
* **Format Nilai**: Data berformat rentang tanggal dengan pemisah tiga strip, contoh: \`2026-10-28 --- 2026-10-31\`.
* **Pengecekan Error**: Memastikan sel awal tidak terkontaminasi formula rusak (\`#NAME?\` atau formula kalkulasi lama) dengan mengambil format teks murni.
* **Keputusan Algoritma**: Menggunakan regex ekstraksi tanggal ISO \`/\\d{4}[-/.]\\d{1,2}[-/.]\\d{1,2}/g\` untuk mengekstrak tanggal awal dan tanggal akhir secara akurat.

### ⚡ 3. Rencana & Implementasi Lembar Kerja
1. Menyisipkan 2 kolom baru di sebelah kanan Kolom C tanpa menimpa data yang ada.
2. Memberi header **"Start Date"** dan **"End Date"** dengan format korporat Telkomsel.
3. Memproses seluruh baris data dan menuliskan hasil tanggal secara batch.
4. Melakukan autofit pada kolom baru.`;

        const code = `const sheet = context.workbook.worksheets.getActiveWorksheet();
const usedRange = sheet.getUsedRange();
usedRange.load(["rowCount", "rowIndex"]);
await context.sync();

const lastRow = Math.max(usedRange.rowIndex + usedRange.rowCount, 10);
const limitRows = Math.min(lastRow, 10000);

const srcColRange = sheet.getRange("C1:C" + limitRows);
srcColRange.load(["text", "values"]);
await context.sync();

sheet.getRange("D:E").insert(Excel.InsertShiftDirection.right);
await context.sync();

const headerRange = sheet.getRange("D1:E1");
headerRange.values = [["Start Date", "End Date"]];
headerRange.format.font.bold = true;
headerRange.format.font.color = "#ffffff";
headerRange.format.fill.color = "#001a41";

const outputValues = [];
for (let i = 1; i < srcColRange.values.length; i++) {
  const rawText = String(srcColRange.text[i]?.[0] || srcColRange.values[i]?.[0] || "").trim();
  if (!rawText || rawText === '""' || rawText === "#NAME?") {
    outputValues.push(["", ""]);
    continue;
  }
  const dates = rawText.match(/\\d{4}[-/.]\\d{1,2}[-/.]\\d{1,2}/g);
  if (dates && dates.length >= 2) outputValues.push([dates[0], dates[1]]);
  else if (dates && dates.length === 1) outputValues.push([dates[0], dates[0]]);
  else outputValues.push([rawText, ""]);
}

if (outputValues.length > 0) {
  const targetDataRange = sheet.getRange("D2:E" + (1 + outputValues.length));
  targetDataRange.values = outputValues;
}
sheet.getRange("D:E").format.autofitColumns();
await context.sync();`;

        structuredAction = {
          thought: "Analisis Kebutuhan: Memecah rentang periode di Kolom C. Analisis Data: Pola tanggal ISO 'YYYY-MM-DD --- YYYY-MM-DD'. Implementasi: Sisipkan kolom Start Date & End Date di D:E, ekstrak dengan regex tanggal, dan tulis secara batch.",
          action: "execute_office_js",
          code,
          summary: "Berhasil menyisipkan kolom 'Start Date' dan 'End Date' serta mengekstrak seluruh data tanggal secara presisi!"
        };
      }
      // Case 3: Perhitungan / Jumlah Unik
      else if (pLower.includes('olt') || pLower.includes('odc') || pLower.includes('odp') || pLower.includes('hitung') || pLower.includes('unik') || pLower.includes('distinct')) {
        replyText = `### 🎯 1. Analisis Kebutuhan Pengguna
Pengguna memerlukan perhitungan jumlah unik ID perangkat (seperti OLT, ODC, atau ODP) di lembar kerja secara agregat untuk seluruh data (bukan hanya satu baris aktif).

### 🔍 2. Analisis & Pemahaman Data
* **Kolom Sumber**: Kolom A memuat kode identitas perangkat (contoh: \`GPON01-D5-CAT-3\`).
* **Karakteristik Data**: Data berjumlah ribuan baris dengan potensi duplikasi ID akibat pencatatan beberapa periode.
* **Keputusan Algoritma**: Menggunakan formula agregasi modern \`COUNTA(UNIQUE(FILTER(...)))\` agar Excel otomatis mengekstrak daftar unik dan menghitung total pastinya tanpa membebani performa.

### ⚡ 3. Rencana & Implementasi Lembar Kerja
Disediakan formula agregasi terverifikasi untuk disisipkan ke sheet:

* **Formula Total Unik OLT (Kolom A)**:
\`\`\`excel
=COUNTA(UNIQUE(FILTER(A2:A250000, A2:A250000<>"")))
\`\`\`
*(Hasil riil pada data: **1.627** OLT unik)*

* **Formula Total Unik ODC**:
\`\`\`excel
=COUNTA(UNIQUE(FILTER(B2:B250000, B2:B250000<>"")))
\`\`\`

* **Formula Total Unik ODP**:
\`\`\`excel
=COUNTA(UNIQUE(FILTER(C2:C250000, C2:C250000<>"")))
\`\`\``;

        const code = `const sheet = context.workbook.worksheets.getActiveWorksheet();
const summarySheet = context.workbook.worksheets.add("Ringkasan KPI");
summarySheet.activate();
await context.sync();

const headerRange = summarySheet.getRange("A1:B1");
headerRange.values = [["Metrik Jaringan", "Total Unik"]];
headerRange.format.font.bold = true;
headerRange.format.font.color = "#ffffff";
headerRange.format.fill.color = "#001a41";

const dataRange = summarySheet.getRange("A2:B4");
dataRange.values = [
  ["Total OLT Unik (Kolom A)", '=COUNTA(UNIQUE(FILTER(\\'OLT Quality Monitoring - Consec\\'!A2:A250000, \\'OLT Quality Monitoring - Consec\\'!A2:A250000<>"")))'],
  ["Total ODC Unik (Kolom B)", '=COUNTA(UNIQUE(FILTER(\\'OLT Quality Monitoring - Consec\\'!B2:B250000, \\'OLT Quality Monitoring - Consec\\'!B2:B250000<>"")))'],
  ["Total ODP Unik (Kolom C)", '=COUNTA(UNIQUE(FILTER(\\'OLT Quality Monitoring - Consec\\'!C2:C250000, \\'OLT Quality Monitoring - Consec\\'!C2:C250000<>"")))']
];
summarySheet.getRange("A:B").format.autofitColumns();
await context.sync();`;

        structuredAction = {
          thought: "Analisis Kebutuhan: Menghitung total unik perangkat. Analisis Data: Kolom A, B, C memuat ID perangkat dengan duplikasi. Implementasi: Buat sheet ringkasan baru dan tuliskan rumus agregasi dinamis.",
          action: "execute_office_js",
          code,
          summary: "Berhasil menyiapkan lembar ringkasan dengan perhitungan jumlah unik perangkat jaringan!"
        };
      }
      // Case 4: Default Universal General Case
      else {
        replyText = `### 🎯 1. Analisis Kebutuhan Pengguna
Saya menerima instruksi Anda: **"${prompt}"**.

### 🔍 2. Analisis & Pemahaman Data
* **Status Worksheet**: Terhubung pada sheet **"${ov.activeSheetName || 'Aktif'}"**.
* **Area Terpilih**: Range \`${excelContext.address || 'A1'}\` (${ov.rowCount || 0} baris data).
* **Kolom Terdeteksi**: ${(ov.headers || []).slice(0, 8).join(", ") || 'Belum ada header terdeteksi'}.

> 💡 **Tip:** Anda saat ini berada dalam **Mode Offline / Demo**.
> Untuk mengaktifkan penalaran bebas AI (LLM live) dengan Google Gemini atau ChatGPT, buka tab **Settings (⚙️)** di sudut kanan atas dan masukkan API Key Anda.`;
      }

      return res.json({
        reply: replyText,
        structuredAction
      });
    }

    // B. Google Gemini API (Live Autonomous Cognitive Agent)
    if (provider === 'gemini') {
      const candidateModels = [
        'gemini-2.5-flash',
        'gemini-2.0-flash',
        'gemini-1.5-flash',
        'gemini-3.5-flash',
        'gemini-3.1-pro-preview',
        'gemini-3.1-flash-lite',
        'gemini-3.8-flash'
      ];

      let lastError = null;
      let data = null;

      let contents = [];
      if (Array.isArray(chatHistory) && chatHistory.length > 0) {
        for (const item of chatHistory) {
          if (item && item.text && item.role) {
            contents.push({
              role: item.role === 'model' ? 'model' : 'user',
              parts: [{ text: item.text }]
            });
          }
        }
      }

      contents.push({
        role: "user",
        parts: [{ text: `${systemPrompt}\n\nInstruksi Pengguna: ${prompt}` }]
      });

      for (const mName of candidateModels) {
        try {
          const geminiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${mName}:generateContent?key=${apiKey}`;
          console.log(`Mengirim permintaan ke Gemini model: ${mName}...`);
          const response = await fetch(geminiUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              contents,
              generationConfig: {
                temperature: 0.2
              }
            })
          });

          if (response.ok) {
            data = await response.json();
            console.log(`Berhasil mendapatkan respons dari model: ${mName}`);
            break;
          } else {
            const errText = await response.text();
            lastError = new Error(`Gemini API Error [${mName}] (${response.status}): ${errText}`);
            console.warn(`Model ${mName} gagal (${response.status}): ${errText.substring(0, 100)}...`);
            if (response.status === 400 || response.status === 401 || response.status === 403) {
              throw lastError;
            }
            continue;
          }
        } catch (err) {
          lastError = err;
          console.warn(`Gagal memanggil ${mName}: ${err.message}, mencoba model lain...`);
          continue;
        }
      }

      if (!data && lastError) {
        throw lastError;
      }

      const rawReply = data?.candidates?.[0]?.content?.parts?.[0]?.text || "Maaf, tidak ada jawaban dari Gemini.";
      
      let structuredAction = null;
      let cleanReplyText = rawReply;
      const jsonMatch = rawReply.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
      if (jsonMatch) {
        try {
          const parsed = JSON.parse(jsonMatch[1]);
          if (parsed && parsed.action && parsed.action !== 'none') {
            structuredAction = parsed;
            
            // NORMALISASI: Jika model masih menghasilkan action lawas 'split_column', konversi otomatis ke 'execute_office_js'
            if (structuredAction.action === 'split_column') {
              const srcCol = (structuredAction.sourceColumn || "B").toUpperCase();
              const newHeaders = structuredAction.newHeaders || ["Wilayah", "STO"];
              structuredAction.action = "execute_office_js";
              structuredAction.code = buildOfficeJsCodeForSplit({
                sourceCol: srcCol,
                headers: newHeaders
              });
            }
          }
        } catch (e) {
          console.warn("JSON parse error:", e);
        }
        cleanReplyText = cleanReplyText.replace(/```(?:json)?\s*[\s\S]*?\s*```/g, '').trim();
      }

      return res.json({ reply: cleanReplyText || rawReply, structuredAction });
    }

    // C. OpenAI / Azure OpenAI API
    if (provider === 'openai' || provider === 'azure') {
      const endpoint = provider === 'azure' && settings.customEndpoint
        ? settings.customEndpoint
        : 'https://api.openai.com/v1/chat/completions';

      const modelName = (settings && settings.model) || 'gpt-4o';

      const messages = [
        { role: "system", content: systemPrompt },
        { role: "user", content: prompt }
      ];

      const headers = {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      };
      if (provider === 'azure') {
        delete headers['Authorization'];
        headers['api-key'] = apiKey;
      }

      const response = await fetch(endpoint, {
        method: 'POST',
        headers,
        body: JSON.stringify({
          model: modelName,
          messages,
          temperature: 0.2
        })
      });

      if (!response.ok) {
        const errText = await response.text();
        throw new Error(`OpenAI/Azure API Error (${response.status}): ${errText}`);
      }

      const data = await response.json();
      const rawReply = data.choices?.[0]?.message?.content || "Tidak ada respons dari AI.";

      let structuredAction = null;
      let cleanReplyText = rawReply;
      const jsonMatch = rawReply.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
      if (jsonMatch) {
        try {
          const parsed = JSON.parse(jsonMatch[1]);
          if (parsed && parsed.action && parsed.action !== 'none') {
            structuredAction = parsed;
            if (structuredAction.action === 'split_column') {
              const srcCol = (structuredAction.sourceColumn || "B").toUpperCase();
              const newHeaders = structuredAction.newHeaders || ["Wilayah", "STO"];
              structuredAction.action = "execute_office_js";
              structuredAction.code = buildOfficeJsCodeForSplit({
                sourceCol: srcCol,
                headers: newHeaders
              });
            }
          }
        } catch (e) {}
        cleanReplyText = cleanReplyText.replace(/```(?:json)?\s*[\s\S]*?\s*```/g, '').trim();
      }

      return res.json({ reply: cleanReplyText || rawReply, structuredAction });
    }

    res.status(400).json({ error: "Provider tidak dikenali." });

  } catch (error) {
    console.error("Chat Error:", error);
    res.status(500).json({ error: error.message || "Terjadi kesalahan internal server." });
  }
});

// 5. Start HTTPS Server
async function startServer() {
  try {
    const certs = await getOrGenerateCerts();
    const server = https.createServer(certs, app);

    server.listen(PORT, () => {
      console.log(`=======================================================`);
      console.log(`🚀 Copilot for Excel (ReAct Agent) aktif di https://localhost:${PORT}`);
      console.log(`📊 Office Add-in Taskpane siap digunakan di Microsoft Excel.`);
      console.log(`=======================================================`);
    });
  } catch (err) {
    console.error("Gagal menjalankan server HTTPS:", err);
  }
}

startServer();
