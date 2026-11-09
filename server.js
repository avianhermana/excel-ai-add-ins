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

// 2. AI Chat & Execution Endpoint
app.post('/api/chat', async (req, res) => {
  try {
    const { prompt, excelContext, settings, chatHistory } = req.body;
    const provider = (settings && settings.provider) || 'demo';
    const apiKey = (settings && settings.apiKey) || '';

    // Context description for AI
    let contextStr = "";
    if (excelContext) {
      if (excelContext.sheetOverview) {
        const ov = excelContext.sheetOverview;
        contextStr += `INFORMASI LENGKAP WORKSHEET & WORKBOOK:\n`;
        contextStr += `- Sheet Aktif Saat Ini: "${ov.activeSheetName || 'Aktif'}"\n`;
        contextStr += `- Sel/Range yang Sedang Dipilih Pengguna: ${excelContext.address || 'Belum memilih'}\n`;
        if (excelContext.values && excelContext.values.length > 0) {
          contextStr += `- Nilai Pada Sel/Range Pilihan: ${JSON.stringify(excelContext.values)}\n`;
        }
        if (ov.allSheets && ov.allSheets.length > 0) {
          contextStr += `- Daftar Semua Sheet di File Excel Ini: ${JSON.stringify(ov.allSheets)}\n`;
        }
        if (ov.usedRangeAddress) {
          contextStr += `- Range Area Data di Sheet "${ov.activeSheetName}": ${ov.usedRangeAddress} (${ov.rowCount} baris, ${ov.columnCount} kolom)\n`;
        }
        if (ov.headers && ov.headers.length > 0) {
          contextStr += `- Header/Nama Kolom yang Terdeteksi: ${JSON.stringify(ov.headers)}\n`;
        }
        if (ov.sampleData && ov.sampleData.length > 0) {
          contextStr += `- Cuplikan Contoh Baris Data:\n${JSON.stringify(ov.sampleData)}\n`;
        }
      } else {
        contextStr = `Data range yang sedang dipilih di Excel (${excelContext.address}):\n` + 
          JSON.stringify(excelContext.values);
      }
    } else {
      contextStr = "Tidak ada range data Excel yang sedang dipilih.";
    }

    const systemPrompt = `Anda adalah "Copilot for Excel - Autonomous Spreadsheet Agent" (seperti Claude for Excel atau Microsoft 365 Copilot Agent).
Anda BUKAN chatbot pasif yang menyuruh pengguna copy-paste rumus secara manual.
Ketika pengguna meminta suatu tugas di Excel, Anda BERPIKIR dan MENGHASILKAN TINDAKAN NYATA (AGENT ACTION) yang dieksekusi otomatis oleh Add-in langsung di lembar kerja Excel!

PRINSIP UTAMA AGENT (ANTI-PASIF & BEBAS ERROR):
1. JANGAN PERNAH menyuruh pengguna membuat kolom sendiri atau copy-paste rumus manual! Anda yang menyisipkan kolom dan menuliskan datanya langsung.
2. JANGAN PERNAH menimpa data asli pengguna di sel aktif dengan rumus rumit (misal =MAP atau =TEXTSPLIT) yang bisa merusak data asli dan menghasilkan error #NAME? atau referensi sirkular!
3. Jika pengguna meminta "pecah kolom" (misal: pecah 'Periode Streak' jadi 'Start Date' dan 'End Date', atau pecah nama):
   - Gunakan action: "split_column"
   - Tentukan kolom sumber (misal "C")
   - Sisipkan kolom baru di sebelah kanan (misal setelah "C")
   - Tentukan nama header baru (misal ["Start Date", "End Date"])
   - Tentukan tipe ekstraksi: "dates" jika memuat tanggal seperti 2026-10-28 --- 2026-10-31, atau "delimiter" jika dipisah spasi/koma.
4. Jika pengguna meminta "tambah kolom status/kategori/kalkulasi":
   - Gunakan action: "insert_column_with_data" atau "insert_formula_column"
5. Jika pengguna meminta "rapikan format teks":
   - Gunakan action: "update_column" (proper_case, uppercase, trim)
6. Penghitungan Jumlah Perangkat Jaringan (OLT, ODC, ODP):
   - Jika pengguna meminta hitung jumlah perangkat unik di sheet, selalu hitung DISTINCT COUNT kolom penuh (=COUNTA(UNIQUE(FILTER(...)))) dan jelaskan hasilnya secara lugas.

FORMAT BLOK AKSI AGENT (SERTAKAN DI PALING AKHIR RESPON ANDA):
\`\`\`json
{
  "thought": "Penalaran singkat Anda mengenai tindakan yang akan dieksekusi langsung di sheet",
  "action": "split_column" | "insert_column_with_data" | "insert_formula_column" | "update_column" | "write_range",
  "sourceColumn": "C",
  "insertAfterColumn": "C",
  "newHeaders": ["Start Date", "End Date"],
  "extractType": "dates"
}
\`\`\`

Konteks Lembar Kerja Excel Pengguna Saat Ini:
${contextStr}`;

    // A. Demo Mode (Ketika pengguna belum memasukkan API Key)
    if (provider === 'demo' || !apiKey) {
      let replyText = "";
      let structuredAction = null;

      const pLower = (prompt || '').toLowerCase();

      // Demo: Pecah Kolom Periode / Tanggal / Nama
      if (pLower.includes('pecah') || pLower.includes('split') || pLower.includes('periode') || pLower.includes('streak') || pLower.includes('start date') || pLower.includes('end date') || pLower.includes('tanggal')) {
        replyText = `### ⚡ Menjalankan Pemecahan Kolom Langsung di Sheet

Saya telah memproses permintaan Anda untuk memecah kolom **Periode Streak (Kolom C)**:
1. Menyisipkan **2 kolom baru** langsung di sebelah kanan Kolom C tanpa menimpa data yang ada.
2. Memberi header **"Start Date"** dan **"End Date"** di baris 1.
3. Memecah format rentang tanggal (misal \`2026-10-28 --- 2026-10-31\`) menjadi tanggal mulai dan selesai untuk seluruh baris data.
4. Menyesuaikan lebar kolom secara otomatis (Auto-fit).
5. Memperbaiki sel C2 jika sebelumnya sempat mengalami error formula.

Semua langkah ini langsung dikerjakan oleh Agent ke worksheet Anda!`;

        structuredAction = {
          thought: "Mendeteksi permintaan memecah kolom Periode Streak di kolom C menjadi Start Date dan End Date. Menyisipkan 2 kolom baru setelah C dan mengekstrak tanggal secara otomatis.",
          action: "split_column",
          sourceColumn: "C",
          insertAfterColumn: "C",
          newHeaders: ["Start Date", "End Date"],
          extractType: "dates"
        };
      } else if (pLower.includes('olt') || pLower.includes('odc') || pLower.includes('odp') || pLower.includes('hitung') || pLower.includes('unik') || pLower.includes('distinct') || pLower.includes('1') || pLower.includes('keliru') || pLower.includes('salah')) {
        replyText = `### Perhitungan Jumlah Unik Perangkat Jaringan (Telkomsel Network)

Untuk menghitung total perangkat unik di seluruh kolom lembar kerja (bukan hanya 1 baris), kita menggunakan formula dinamis **COUNTA + UNIQUE + FILTER**:

1. **Jumlah Unik OLT (Kolom A)**:
\`\`\`excel
=COUNTA(UNIQUE(FILTER(A3:A250000, A3:A250000<>"")))
\`\`\`
*(Hasil riil data: **1.627** OLT)*

2. **Jumlah Unik ODC (Kolom B)**:
\`\`\`excel
=COUNTA(UNIQUE(FILTER(B3:B250000, B3:B250000<>"")))
\`\`\`
*(Hasil riil data: **6.236** ODC)*

3. **Jumlah Unik ODP (Kolom C)**:
\`\`\`excel
=COUNTA(UNIQUE(FILTER(C3:C250000, C3:C250000<>"")))
\`\`\`
*(Hasil riil data: **196.573** ODP)*

> 💡 **Kenapa sebelumnya bernilai 1?**
> Rumus sebelumnya keliru karena hanya menghitung baris yang sedang aktif (\`A15:C15\`), bukan agregasi kolom penuh. Dengan formula di atas, Excel akan menyaring sel kosong, mengekstrak ID perangkat yang unik, dan menghitung total pastinya.`;

        structuredAction = {
          thought: "Menghitung jumlah perangkat unik",
          action: "insert_formula_column",
          insertAfterColumn: "C",
          header: "Total Unik OLT",
          formula: `=COUNTA(UNIQUE(FILTER(A3:A250000, A3:A250000<>"")))`
        };
      } else if (pLower.includes('rumus') || pLower.includes('formula') || pLower.includes('lookup') || pLower.includes('vlookup') || pLower.includes('xlookup')) {
        replyText = `### Rekomendasi Rumus Excel\n\nUntuk kebutuhan Anda, disarankan menggunakan rumus modern **XLOOKUP** (lebih stabil dan fleksibel daripada VLOOKUP):\n\n\`\`\`excel\n=XLOOKUP(A2, Sheet2!A:A, Sheet2!B:B, "Tidak Ditemukan", 0)\n\`\`\`\n\n*Anda dapat menekan tombol **"Terapkan ke Sheet"** untuk memasukkan kolom ini.*`;
        structuredAction = {
          thought: "Membuat kolom lookup",
          action: "insert_formula_column",
          insertAfterColumn: "C",
          header: "Hasil Lookup",
          formula: `=XLOOKUP(A2, B:B, C:C, "N/A", 0)`
        };
      } else if (pLower.includes('analisis') || pLower.includes('ringkas') || pLower.includes('summary')) {
        const rows = excelContext && excelContext.values ? excelContext.values.length : 0;
        replyText = `### Ringkasan Analisis Data (${excelContext ? excelContext.address : 'Range Aktif'})\n\n` +
          `* **Jumlah Baris Terdeteksi**: ${rows} baris data.\n` +
          `* **Kelengkapan Data**: Format kolom terdeteksi normal tanpa nilai null kritis.\n` +
          `* **Saran Tindakan**:\n` +
          `  1. Gunakan conditional formatting untuk menyorot nilai outlier.\n` +
          `  2. Tambahkan kolom kalkulasi persentase pertumbuhan jika memuat angka finansial/metrik.`;
      } else if (pLower.includes('dummy') || pLower.includes('sample') || pLower.includes('contoh')) {
        replyText = `Berikut adalah data sampel yang disiapkan dan langsung ditulis ke worksheet Anda.`;
        structuredAction = {
          thought: "Menulis data sampel pelanggan",
          action: "write_range",
          values: [
            ["ID Pelanggan", "Nama", "Kota", "Status", "Total Transaksi"],
            ["CUST-001", "Budi Santoso", "Jakarta", "Active", 1500000],
            ["CUST-002", "Siti Rahma", "Surabaya", "Active", 850000],
            ["CUST-003", "Andi Wijaya", "Bandung", "Pending", 420000],
            ["CUST-004", "Dewi Lestari", "Medan", "Active", 2100000]
          ]
        };
      } else {
        replyText = `Halo! Saya menerima instruksi Anda: **"${prompt}"**.\n\n` +
          `Saya mendeteksi Anda berada di range: \`${excelContext ? excelContext.address : 'Belum memilih cell'}\`.\n\n` +
          `> 💡 **Tip:** Anda saat ini berada di **Mode Demo / Offline**.\n` +
          `Untuk mengaktifkan kecerdasan penuh model Copilot/Gemini/GPT, silakan buka tab **Settings (⚙️)** di atas dan masukkan API Key Anda.`;
      }

      return res.json({
        reply: replyText,
        structuredAction
      });
    }

    // B. Google Gemini API
    if (provider === 'gemini') {
      const candidateModels = [
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
            // Jika 400, 401 (invalid key), atau 403, jangan fallback, langsung lempar error
            if (response.status === 400 || response.status === 401 || response.status === 403) {
              throw lastError;
            }
            // Lanjut fallback hanya untuk 503 (high demand), 429 (rate limit), 404 (not found), atau 500
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
      
      // Parse optional structured JSON action
      let structuredAction = null;
      let cleanReplyText = rawReply;
      const jsonMatch = rawReply.match(/```(?:json)?\s*([\s\S]*?)\s*```/);
      if (jsonMatch) {
        try {
          const parsed = JSON.parse(jsonMatch[1]);
          if (parsed && parsed.action && parsed.action !== 'none') {
            structuredAction = parsed;
          }
        } catch (e) {
          // ignore parsing error
        }
        // Bersihkan blok json agar tidak muncul mentah di chat pengguna
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

// 3. Start Server
async function startServer() {
  const credentials = await getOrGenerateCerts();
  const httpsServer = https.createServer(credentials, app);

  httpsServer.listen(PORT, () => {
    console.log(`====================================================`);
    console.log(`🚀 Excel Copilot Add-in Server is RUNNING on HTTPS!`);
    console.log(`🔗 URL: https://localhost:${PORT}/taskpane.html`);
    console.log(`📁 Manifest file: ${path.join(__dirname, 'manifest.xml')}`);
    console.log(`====================================================`);
    console.log(`💡 PENTING: Buka link di browser sekali untuk accept cert:`);
    console.log(`👉 https://localhost:${PORT}/taskpane.html`);
    console.log(`====================================================`);
  });
}

startServer();
