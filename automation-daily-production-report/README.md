# 🚀 Automation Daily Report

Aplikasi PyQt6 untuk otomatisasi pengambilan data dari beberapa file Excel (IMM/TCM/BEK/GPK/JBG/TIS/SUM),
mengisi Draft output, dan menyiapkan file untuk Power BI Dashboard.

## ✨ Fitur

- GUI input file (Browse)
- Otomatis mapping cell Excel → Draft Output
- ROM Stock otomatis (khusus IMM ada rule West/East di kolom AK)
- Summary → Output (FC Plan + Formula %)
- Force recalculation via Excel (win32com) ✅ Windows

## 📁 Struktur Folder

- `app/gui` : GUI & UI
- `app/logic` : proses utama excel mapping
- `app/config` : konfigurasi mapping (`inputan.json`)
- `app/assets` : icon & background
- `app/style` : stylesheet

## ⚙️ Requirement

- Python 3.8+
- Windows (karena `pywin32 / win32com` untuk recalculation Excel)

## 🛠 Install

```bash
pip install -r requirements.txt
```

## Menjalankan

Masuk folder app/ lalu jalankan:

```bash
run.bat
```

Atau via python:

```bash
python app/main.py
```

## ⚠️ Catatan Konfigurasi

File `inputan.json` berisi **contoh konfigurasi**.

User perlu menyesuaikan **path file Excel** sesuai environment masing-masing
(seperti lokasi file di komputer lokal).
