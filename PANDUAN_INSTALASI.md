# 📑 Panduan Instalasi & Penggunaan: Telkomsel Copilot for Excel

> **Asisten AI Cerdas Berkinerja Tinggi untuk Pengolahan Spreadsheet Jaringan Telekomunikasi & Analitik Data Telkomsel**

---

## ⚡ Ringkasan Cepat (3 Langkah Mudah)

1. **Jalankan `1-Install-Addin.bat`** (cukup sekali di awal).
2. **Jalankan `2-Jalankan-Server.bat`** (atau klik shortcut *Mulai Telkomsel Copilot* di Desktop).
3. **Buka Excel**:
   * Masuk menu **Insert** (*Sisipkan*) ➔ **My Add-ins** (*Add-in Saya*).
   * Pilih tab **SHARED FOLDER** (*Folder Bersama*).
   * Klik **Copilot for Excel** ➔ **Add** (*Tambah*).

---

## 📋 Persyaratan Sistem
* **Sistem Operasi**: Windows 10 / Windows 11 (64-bit).
* **Microsoft Excel**:
  * Excel Desktop (Microsoft 365 Enterprise Telkomsel / Office 2021).
  * Atau Excel Web (Office 365 melalui browser).
* **Node.js**: Versi 18 LTS atau lebih baru.
  *(Jika belum terpasang, installer `1-Install-Addin.bat` akan memandu instalasi otomatis).*

---

## 🛠️ Panduan Instalasi Detail

### Opsi A: Penggunaan di Excel Desktop (Windows)
1. Clone repository ini menggunakan Git (`git clone https://github.com/avianhermana/excel-ai-copilot.git`) atau ekstrak file ke folder kerja Anda.
2. Klik 2x file **`1-Install-Addin.bat`**. Installer otomatis:
   * Memeriksa apakah Node.js sudah ada.
   * Mendaftarkan sertifikat SSL localhost ke Windows Trust Store agar tidak diblokir WebView2.
   * Mendaftarkan folder manifest ke katalog resmi Excel Desktop (*Trusted Catalogs*).
   * Membuat shortcut **Mulai Telkomsel Copilot** di Desktop.
3. Klik 2x shortcut **Mulai Telkomsel Copilot** di Desktop (atau jalankan `2-Jalankan-Server.bat`).
   * *Biarkan jendela terminal hitam ini tetap terbuka selama menggunakan Excel.*
4. Buka Microsoft Excel di komputer Anda:
   * Buka workbook data yang ingin Anda olah.
   * Klik tab **Insert** ➔ klik panah kecil di **My Add-ins**.
   * Pilih tab **SHARED FOLDER** ➔ pilih **Copilot for Excel** ➔ klik **Add**.
5. Panel Telkomsel Copilot akan langsung terbuka di sebelah kanan lembar kerja Anda.

---

### Opsi B: Penggunaan di Excel Web (Office 365 / SharePoint Telkomsel)
Jika Anda membuka Excel melalui browser (`365tsel-my.sharepoint.com`):
1. Pastikan server lokal sudah berjalan dengan menjalankan `2-Jalankan-Server.bat`.
2. Di lembar kerja Excel Web:
   * Klik tab **Insert** (Sisipkan) ➔ **Add-ins** (Add-in).
   * Klik tautan **Manage My Add-ins** (*Kelola Add-in Saya*) di kanan atas dialog.
   * Pilih **Upload My Add-in** (*Unggah Add-in Saya*).
   * Klik tombol **Browse** dan pilih file `manifest.xml` dari folder distribusi ini.
   * Klik **Upload**.
3. Ikon **Copilot AI** akan langsung muncul di Ribbon menu atas.

---

## 🔑 Pengaturan Model AI (Google Gemini / Azure OpenAI)

Copilot for Excel sudah dilengkapi dengan **Mode Demo/Offline** bawaan. Namun untuk mengaktifkan kecerdasan penuh tingkat tinggi:

1. Klik tombol roda gigi **Settings (⚙️)** di pojok kanan atas header panel Copilot.
2. Pilih provider AI:
   * **Google Gemini API** *(Sangat direkomendasikan - Model Gemini 3.5 Flash cepat dan akurat)*.
   * **Azure OpenAI / Enterprise Telkomsel Gateway**.
   * **OpenAI (ChatGPT GPT-4o)**.
3. Masukkan **API Key** Anda:
   * Cara mendapatkan API Key Gemini gratis:
     1. Buka [https://aistudio.google.com/](https://aistudio.google.com/).
     2. Login dengan akun Google Anda.
     3. Klik **Get API key** ➔ **Create API key**.
     4. Salin kunci (berawalan `AIzaSy...` atau `AQ...`) dan tempelkan ke kolom API Key di panel Copilot.
4. Klik tombol **💾 Simpan Pengaturan**.

> 🔒 **Keamanan & Privasi**: API Key disimpan hanya di browser lokal Anda (`localStorage`) dan tidak pernah dikirim ke pihak luar mana pun.

---

## 💡 Contoh Penggunaan Praktis untuk Data Jaringan Telkomsel

Copilot memahami hirarki dan terminologi jaringan telekomunikasi (FTTH, OLT, ODC, ODP, feeder, core, status provisioning):

### 1. Menghitung Jumlah Perangkat Jaringan (Distinct Count Kolom Penuh)
Ketik santai di kolom chat:
> *"Tolong hitung jumlah unik perangkat OLT, ODC, dan ODP di sheet ini"*

*Copilot akan otomatis menyusun rumus Excel 365 dinamis:*
```excel
=COUNTA(UNIQUE(FILTER(A3:A250000, A3:A250000<>"")))
```
*Mengabaikan baris kosong, mengambil ID unik, dan menghitung total pastinya tanpa salah hitung 1 baris.*

### 2. Pencocokan Data Antar-Sheet (VLOOKUP / XLOOKUP Cerdas)
> *"Bang tolong kolom yang gw pilih isi datanya dengan 'FBB_ODP' jika odp_name di kolom BC ditemukan di sheet fbb_odp kolom A. Kalau tidak ada, ganti dengan strip '-'"*

### 3. Pembersihan Format Otomatis
* Buka tab **Clean (✨)** di navigasi atas:
  * **Hapus Spasi Ekstra**: Merapikan teks dengan fungsi `TRIM`.
  * **Proper Case / UPPERCASE**: Menyeragamkan penulisan nama perangkat atau kota.
  * **Isi Sel Kosong**: Mengganti sel kosong (blank) dengan tanda `-`.
  * **Auto-fit Kolom**: Menyesuaikan lebar kolom secara otomatis.

---

## ❓ Troubleshooting (Tanya Jawab Kendala Umum)

**Q: Panel menampilkan pesan "Add-in Error: The content is blocked because it isn't signed by a valid security certificate"?**  
**A:** Jalankan ulang file `1-Install-Addin.bat` (Run as Administrator jika diperlukan), lalu restart Microsoft Excel Anda.

**Q: Panel hanya menampilkan layar putih atau tidak bisa terhubung?**  
**A:** Pastikan jendela terminal `2-Jalankan-Server.bat` sedang berjalan di latar belakang (port 3000 HTTPS).

**Q: Bagaimana cara memperbarui tampilan taskpane jika ada perubahan?**  
**A:** Cukup klik kanan di area kosong mana saja di panel Copilot ➔ pilih **Reload** (atau tekan tombol **F5** pada keyboard).

---

*Dikembangkan untuk Tim Data & Operasi Jaringan Telkomsel* ⚡
