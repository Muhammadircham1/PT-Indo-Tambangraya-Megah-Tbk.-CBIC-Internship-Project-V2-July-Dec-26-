# 📌 Project System Automasi Weekly Report

**Project System Automasi Weekly Report** adalah sebuah sistem otomatisasi yang melakukan perpindahan data dari file *summary* ke file *draft*.  
File *draft* tersebut nantinya digunakan sebagai sumber data di **Power BI** untuk membuat dashboard laporan mingguan.

---

## ✨ Fitur Utama
- ✅ Mengolah data dari file Excel (`data/`)
- ✅ GUI sederhana untuk interaksi pengguna (`gui/`)
- ✅ Otomatisasi laporan & integrasi dengan Power BI (`powerBI/`)
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
│── setup.sh # Script setup & run di Linux/Mac
│
├── data/ # Dataset / file input
│ ├── Draft_3rdparty.xlsx
│ └── Draft_weeklyReport.xlsx
│
├── powerBI/ # Integrasi / laporan Power BI
│
├── .gitignore # File gitignore
├── README.md # Dokumentasi
└── requirements.txt # Dependensi Python
```

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
git clone https://github.com/ITM-CBIC-Team/SystemAutomation_WeeklyReport.git
```

## Install Requirement yang dibutuhkan

```
pip install -r requirements.txt
```

## Jalankan run.bat
```
run.bat
```