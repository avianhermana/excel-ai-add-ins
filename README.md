# ⚡ Excel AI Add-ins

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Platform: Microsoft Excel](https://img.shields.io/badge/Excel-Desktop%20%7C%20Web-107C41.svg?logo=microsoftexcel&logoColor=white)](https://office.com)
[![Node.js](https://img.shields.io/badge/Node.js-18%2B-339933.svg?logo=node.js&logoColor=white)](https://nodejs.org)

An intelligent, enterprise-grade AI Taskpane Add-in for **Microsoft Excel**. Operates directly inside your Excel workbook (both Excel Desktop and Excel for the Web) to analyze data, generate advanced formulas, clean datasets with 1-click tools, and write computation results directly into worksheet cells with autonomous agent capabilities.

---

## 🌟 Key Features

* **⚡ Real-Time Selection Tracking**: Instantly detects and reflects the active cell or range selection (`Sheet1!A1:D20`) directly in the sidepanel.
* **🧠 High Context Awareness**: Comprehends spreadsheet layouts, headers, and telecommunications/network hierarchies (OLT, ODC, ODP, FTTH) without freezing on large datasets (up to 250,000+ rows).
* **💬 Two-Way Interactive Chat**:
  * Ask natural language questions about your active sheet.
  * Receive recommendations, insights, and formula breakdowns.
  * **Interactive Action Buttons**: Directly apply formulas or write data matrices into active cells with a single click.
* **📐 Smart Formula Builder**:
  * Generates modern dynamic array formulas (`XLOOKUP`, `INDEX/MATCH`, `SUMIFS`, `COUNTA(UNIQUE(FILTER(...)))`, `TEXTJOIN`, `IF/IFS`).
  * One-click insertion into the active Excel cell.
* **✨ 1-Click Data Cleaner**:
  * ✂️ **Trim Extra Spaces**: Cleans leading, trailing, and duplicate spaces.
  * 🔠 **Case Converter**: Title Case (Proper), UPPERCASE, or lowercase.
  * 💰 **Currency Formatting**: Formats numbers into standard currency (`Rp #,##0`).
  * ↔️ **Auto-Fit Columns**: Automatically optimizes column widths for selected data.
  * ⭕ **Fill Blank Cells**: Fills empty cells with `-` or default values.
  * 📋 **Convert to Excel Table**: Formats selected range into an official Excel Table with headers.
* **🤖 Multi-Model AI Flexibility**:
  * **Demo / Offline Mode**: Works out of the box without any API Key!
  * **Google Gemini API**: Recommended (Gemini 3.5 Flash / 3.8 Flash for ultra-fast, accurate responses).
  * **Anthropic Claude**: Support for Claude 3.5 Sonnet / Claude 3.7.
  * **OpenAI**: Supports GPT-4o and GPT-4o mini.
  * **Azure OpenAI / Enterprise Gateway**: Connect to internal enterprise endpoints.

---

## 📦 Project Structure

```text
excel-ai-add-ins/
│
├── 1-Install-Addin.bat        # [1-Click Setup] Windows automated installer
├── 2-Jalankan-Server.bat      # Starts local HTTPS server & opens dashboard
├── uninstall.bat              # Unregisters add-in from Excel catalog
├── manifest.xml               # Official Office Add-in manifest configuration
├── server.js                  # Local HTTPS backend (Express + SSL + Multi-LLM API)
├── generate-cert.js           # SSL self-signed certificate generator
├── package.json               # Node.js project manifest & scripts
├── PANDUAN_INSTALASI.md       # Comprehensive installation guide (Indonesian)
├── README.md                  # Project documentation
│
└── public/                    # Web Taskpane front-end
    ├── taskpane.html          # Taskpane UI layout & styling
    ├── taskpane.js            # Office.js automation & AI connection
    └── assets/                # Add-in icons (16, 32, 64, 80px)
```

---

## 🚀 Quick Start & Installation

### Prerequisites
* Windows 10 / Windows 11 (64-bit).
* **Microsoft Excel** (Desktop 2019/2021/365, or Excel on the Web).
* **Node.js** (v18 or higher recommended).

### Automated Installation (Windows 1-Click)

1. Clone or download this repository:
   ```bash
   git clone https://github.com/avianhermana/excel-ai-add-ins.git
   cd excel-ai-add-ins
   ```
2. Double-click **`1-Install-Addin.bat`**. The installer will automatically:
   * Verify your Node.js installation (or offer to install it).
   * Install npm dependencies (`npm install`).
   * Generate a self-signed SSL certificate for `localhost`.
   * Register the certificate in the Windows Trusted Root store.
   * Configure Excel WebView2 and register the manifest folder into Excel's `Trusted Catalogs`.
   * Create a **Mulai Excel AI Add-ins** shortcut on your Desktop.

3. Double-click **`2-Jalankan-Server.bat`** (or the Desktop shortcut) to start the local HTTPS server.

4. Open **Microsoft Excel Desktop**:
   * Navigate to the **Insert** tab ➔ click **My Add-ins**.
   * Switch to the **SHARED FOLDER** tab.
   * Select **Excel AI Add-ins** ➔ click **Add**.
   * The AI Add-in sidepanel will open immediately!

---

### Running on Excel for the Web (Office 365)

1. Ensure the local server is running (`2-Jalankan-Server.bat`).
2. Open any spreadsheet on [Office.com / Excel](https://www.office.com).
3. Click the **Insert** tab ➔ **Add-ins**.
4. In the upper-right corner of the dialog, click **Manage My Add-ins** ➔ **Upload My Add-in**.
5. Browse and select `manifest.xml` from this repository.
6. Click **Upload**.

---

## ⚙️ AI Configuration

Click the **Settings (⚙️)** gear icon in the top right of the sidepanel to choose your preferred AI backend:

1. **Google Gemini (Recommended)**:
   * Select **Google Gemini API**.
   * Get a free API Key from [Google AI Studio](https://aistudio.google.com/).
   * Paste your key into the API Key field and save.
2. **OpenAI / Azure OpenAI / Claude**:
   * Select **OpenAI** or **Azure OpenAI / Enterprise Gateway**.
   * Enter your API Key (and custom Endpoint URL for Azure).
   * All API keys are securely stored solely in your local browser (`localStorage`) and never shared externally.

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).
