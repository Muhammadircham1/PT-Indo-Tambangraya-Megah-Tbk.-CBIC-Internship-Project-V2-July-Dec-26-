# 📌 Project System Automasi SSO to Summary

**Project System Automasi SSO to Summary** adalah sebuah sistem otomatisasi yang memindahkan data dari hasil output SSO ke dalam file Summary. Sistem ini dirancang untuk mempermudah proses penyusunan serta pengelolaan database Summary pada tim CBIC, sehingga alur kerja menjadi lebih efisien, konsisten, dan terstruktur.

---

## ✨ Fitur Utama
- ✅ Mengolah data dari file Excel (`data/`)
- ✅ GUI sederhana untuk interaksi pengguna (`gui/`)
- ✅ Struktur project modular dengan folder `app/`, `logic/`, dan `style/`
- ✅ Dukungan untuk Windows (`run.bat`) dan Linux/Mac (`setup.sh`)

---
## 📂 Struktur Folder

```
project-name/
│── app/ # Core aplikasi
│── assets/ # Gambar, ikon, atau asset lainnya
│── config/ # File konfigurasi JSON / settings
│── gui/ # Modul GUI (interface)
│── logic/ # Modul logika bisnis / processing
│── style/ # File style / tema UI
│── main.py # Entry point aplikasi
│── run.bat # Script untuk menjalankan aplikasi di Windows
│
├── data/ # Dataset / file input
│ ├── Sample_Input.xlsx
│
├── .gitignore # File gitignore
├── README.md # Dokumentasi
└── requirements.txt # Dependensi Python
```

## 🧠 Memahami Alur Logika Sistem (Quick Logic Learn)

Bagi user atau developer baru, berikut adalah gambaran ringkas bagaimana sistem ini bekerja:
1. **Copy Sheet**: Sistem mengambil sheet `Loading` dari file SSO Output dan menyalinnya ke dalam file Summary tujuan.
2. **Backup Data Penting**: Sebelum memproses, sistem mencadangkan (*backup*) baris-baris berstatus "Plan", nilai *Demurrage rate*, *FPG*, kualitas, serta *color fill* pada sheet `ITM Summary`.
3. **Pemrosesan per Bulan**: Sistem membaca data dari sheet `Loading`, mencocokkan kolom berdasarkan nama *header* (*column mapping*), lalu memperbarui, menimpa, atau menambahkan baris baru ke dalam blok bulan yang sesuai di `ITM Summary`.
4. **Restore & Kalkulasi Ulang**: Sistem mengembalikan data "Plan" yang di-backup, merapikan nomor urut, dan menyuntikkan ulang formula-formula (seperti Mahakam/BoCT).
5. **Penyimpanan**: Pembersihan sheet sementara, lalu file Summary langsung disimpan dengan perubahan terbaru.

### 🛠️ Panduan Menyesuaikan Kode (Adjusting / Customizing)

Jika ada perubahan format pada Excel atau *rules* dari user, berikut adalah file-file kunci yang perlu Anda periksa:

*   **Mengubah Pencocokan Kolom (Mapping Kolom)**: `app/config/column_mapping.py`
    *   *Kapan digunakan?* Jika ada penambahan kolom baru, perubahan nama *header* di file SSO Output, atau letak *header* di file Summary bergeser. Anda cukup mengganti mappingnya di dalam file ini.
*   **Mengubah Formula Excel**: `app/logic/main_logic.py` dan `app/logic/helpers.py`
    *   *Kapan digunakan?* Jika ada perubahan *logic* rumus perhitungan Excel (misal: perhitungan Laytime, rate BoCT/Mahakam). Formula umum ada di variabel *dictionary* `formulas` pada `main_logic.py`. Untuk injeksi formula khusus (*BoCT* & *Mahakam*), cari fungsi terkait di dalam `helpers.py`.
*   **Mengubah Logika Pemrosesan Per Baris (Update/Insert)**: `app/logic/data_handler.py`
    *   *Kapan digunakan?* File ini memuat fungsi `process_data_per_month` yang merupakan *jantung* dari proses per baris. Buka file ini jika Anda ingin mengubah kriteria *update* baris lama, menambah *rules* insert baris baru, atau mengubah logika pengecekan status ("Plan", "In Progress", "Complete").
*   **Mengubah Cara Backup & Restore Data**: Folder `app/logic/` (File berawalan `backup_restore_...` / `..._backup.py`)
    *   *Kapan digunakan?* Jika Anda ingin menambahkan atau menghilangkan kolom data manual yang harus di-backup saat *script* berjalan. (Contoh: `fpg_backup.py`, `dem_rate_backup.py`, `backup_restore_quality.py`).

---

## ⚠️ Persiapan Sebelum Menjalankan Otomatisasi

Lakukan langkah-langkah wajib ini **sebelum** Anda menjalankan sistem agar terhindar dari *error*:

1. **Tutup File Excel**: Pastikan file Summary dan file Output SSO **TIDAK SEDANG DIBUKA** di aplikasi Microsoft Excel. Jika file sedang terbuka, sistem akan gagal memproses dan menghasilkan *PermissionError*.
2. **Siapkan File**: Pastikan Anda mengetahui lokasi path file Excel **Output SSO** dan **Summary (Master)** yang akan diproses.
3. **Periksa Nama Sheet**: 
   - Pada file **Summary**, WAJIB terdapat sheet dengan nama tepat: `ITM Summary`.
   - Pada file **Output SSO**, WAJIB terdapat sheet dengan nama tepat: `Loading`.
4. **Struktur Header**: Pastikan baris header (judul kolom) pada data Excel tidak diubah-ubah, karena sistem mengandalkan nama header untuk mencocokkan data.

---

## ⚙️ Instalasi & Menjalankan

### 1. Install Python
Pastikan Python **3.8 atau lebih baru** sudah terinstall.  
Download di: [Python.org](https://www.python.org/downloads/)

Cek apakah Python sudah terinstall dengan:
```bash
python --version

```

## ⚙️ Clone Repository

1. Clone repository ini:
```
git clone https://github.com/ITM-CBIC-Team/SystemAutomation_SSOtoSummary.git
```

## Install Requirement yang dibutuhkan

```
pip install -r requirements.txt
```

## Jalankan run.bat
```
run.bat
```