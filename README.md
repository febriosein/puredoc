# PureDoc PDF 📄✨
> **Konverter Word ke PDF Tanpa Kompresi Gambar & Tata Letak 100% Identik**  
> Sesuai Spesifikasi PRD PureDoc PDF v1.0 (MVP)

---

## 1. Masalah yang Diselesaikan

Ketika mengonversi file Word (`.docx`) ke PDF, mayoritas *converter* online atau ekspor bawaan menyebabkan:
1. **Gambar buram / pecah:** Word dan converter default melakukan *downsampling* (misal ke 150–220 ppi) dan kompresi ulang JPEG berkualitas rendah.
2. **Layout bergeser:** Spasi baris, posisi tabel, *text wrapping*, nomor halaman, dan pergantian font berubah karena mesin render pihak ketiga tidak identik dengan Microsoft Word.

**PureDoc PDF** memecahkan masalah ini dengan dua pilar utama:
- **Layout Identik 100%:** Menggunakan otomasi asli mesin Microsoft Word pada Windows (`Word.Application` COM) sehingga perataan teks, margin, tabel, persamaan (*equation*), dan font dirender persis seperti di Word.
- **Pass-through Gambar Asli (Zero Loss):** Mengekstrak data gambar asli bit-for-bit dari arsip `.docx` (`word/media/`), memetakan posisinya, dan menginjeksi kembali resolusi aslinya ke dalam PDF XObjects tanpa re-encoding lossy dan tanpa downsampling.

---

## 2. Fitur MVP (Fase 1)

- 📤 **Drag & Drop Upload:** Unggah file `.docx` hingga 100 MB dengan validasi *magic bytes*.
- 🔍 **Analisis Pra-Konversi (FR-4, FR-5, FR-6):**
  - Deteksi jumlah gambar, resolusi terendah & tertinggi (piksel), dan ukuran file.
  - Pengecekan ketersediaan font dokumen terhadap font sistem Windows (`C:\Windows\Fonts` dan registry).
  - Peringatan jika ada font yang tidak terpasang beserta rekomendasi font pengganti yang metriknya setara.
  - Deteksi elemen kompleks (tabel, equation, shape, textbox).
- ⚙️ **Konversi High-Fidelity (FR-7 s.d FR-14):**
  - Ekspor Word dengan flag kualitas cetak maksimal (`wdExportOptimizeForPrint`, *heading bookmarks*, *doc structure tags*).
  - Restorasi gambar asli beresolusi penuh secara lossless.
- 📊 **Verifikasi Otomatis & Laporan Kualitas (FR-15 s.d FR-17, Bab 13):**
  - Membandingkan jumlah halaman Word vs PDF.
  - Membandingkan dimensi piksel dan *checksum* SHA-256 gambar sebelum dan sesudah konversi.
  - Memverifikasi *embedded fonts* pada PDF.
  - Status Laporan: **Lolos** (Hijau), **Lolos dengan Peringatan** (Kuning), atau **Perlu Dicek** (Merah).
- 👁️ **Pratinjau & Unduh (FR-18, FR-19):**
  - Pratinjau interaktif PDF di browser.
  - Tombol unduh file PDF hasil konversi.
  - Tombol unduh Laporan Kualitas ringkas (JSON).
- 🔒 **Privasi & Keamanan (FR-20):**
  - Tombol **"Hapus Sekarang"** untuk menghapus file secara instan dari server setelah selesai.
  - Mekanisme pembersihan otomatis di latar belakang (*background TTL*) setelah 1 jam.

---

## 3. Struktur Direktori

```text
wordtopdf/
├── app/
│   ├── __init__.py
│   ├── config.py           # Konfigurasi direktori, batas 100MB, TTL 1 jam
│   ├── font_detector.py    # Scanner font sistem Windows & fallback mapping
│   ├── docx_analyzer.py    # Ekstraksi media docx, analisis XML font & elemen
│   ├── word_converter.py   # Wrapper Word COM automation (ExportAsFixedFormat)
│   ├── pdf_restorer.py     # Injeksi kembali resolusi gambar asli ke PDF
│   ├── verifier.py         # Generator Laporan Kualitas Bab 13 PRD
│   ├── cleanup.py          # Auto-cleanup TTL & hapus permanen manual
│   └── main.py             # FastAPI backend & routing API
├── static/
│   ├── index.html          # Web UI responsif & modern (Bahasa Indonesia)
│   ├── css/
│   │   └── style.css       # Styling bersih & modern
│   └── js/
│       └── app.js          # Logika frontend, progress stepper, preview, download
├── storage/
│   ├── uploads/            # File docx yang diunggah
│   ├── outputs/            # File PDF hasil konversi
│   └── temp/               # File sementara
├── tests/
│   └── test_puredoc.py     # End-to-end unit & integration test
├── requirements.txt
├── run.py                  # Script peluncur satu klik
└── README.md
```

---

## 4. Cara Menjalankan

### Persyaratan:
- Windows OS dengan **Microsoft Word** terpasang (Office 2016, 2019, 2021, atau Office 365).
- Python 3.10+ (sudah teruji di Python 3.14).

### Menjalankan Server:
Cukup jalankan perintah berikut dari folder proyek:

```powershell
python run.py
```

Server akan otomatis berjalan di **`http://127.0.0.1:8000`** dan membuka browser secara otomatis.

---

## 5. Ringkasan Endpoint API

| Method | Endpoint | Deskripsi |
| --- | --- | --- |
| `GET` | `/` | Antarmuka Web PureDoc PDF |
| `GET` | `/api/health` | Status server, Word engine, dan font sistem |
| `POST` | `/api/upload` | Mengunggah `.docx` dan menghasilkan analisis pra-konversi |
| `POST` | `/api/jobs/{id}/convert` | Menjalankan konversi Word, restorasi gambar, dan verifikasi |
| `GET` | `/api/jobs/{id}/pdf` | Mengunduh atau menampilkan pratinjau PDF |
| `GET` | `/api/jobs/{id}/report` | Mengambil laporan kualitas lengkap (JSON) |
| `DELETE` | `/api/jobs/{id}` | Menghapus permanen file dokumen dan PDF dari server (FR-20) |
