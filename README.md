# ⭐ Italikin v1.0.0

**Ubah format kata tertentu di dokumen Word — tanpa merusak layout ✨**

Italikin adalah aplikasi desktop Python untuk memformat kata-kata target secara otomatis di dokumen Microsoft Word (.docx). Aplikasi ini mengubah format (italic, bold, underline, highlight, warna) pada kata tertentu tanpa merusak layout, gambar, atau formatting lain di dokumen.

---

## 🎯 Fitur Utama

- **Deteksi Kata Target** — Case-insensitive di seluruh dokumen (termasuk tabel & tabel bersarang)
- **Aman untuk Formatting** — Manipulasi di level run sehingga format asli (font-size, bold, warna, layout) tidak hilang
- **Casing Asli Dipertahankan** — Hanya properti format yang diubah, teks tetap apa adanya
- **Pencarian Presisi** — Kata utuh saja (whole word) & case-sensitive
- **Format Output Fleksibel** — Italic / Bold / Underline / Highlight / Warna kustom
- **Dual Engine** — python-docx (cepat) atau MS Word COM (cover header/footer)
- **Drag & Drop** — Seret file .docx/.txt langsung ke jendela aplikasi
- **GUI Modern** — CustomTkinter dengan Dark Mode
- **Menu Bar** — File, Tools, Bantuan
- **Bebas Asset Eksternal** — Semua ikon memakai unicode standar

---

## 📸 Tampilan Aplikasi

```
┌────────────────────────────────────────────────────────────────────┐
│  📁 File    🛠️ Tools    ❓ Bantuan                                │
├────────────────────────────────────────────────────────────────────┤
│  ⭐ ITALIKIN                                        v1.0.0       │
│  Ubah format kata tertentu di dokumen Word          by Italikin   │
│  🖱️ Tips: seret file .docx / .txt langsung ke jendela ini.       │
├────────────────────────────────────────────────────────────────────┤
│  📄 File Dokumen Word (.docx)                                    │
│  [______________________________] [📁 Browse…]                   │
├────────────────────────────────────────────────────────────────────┤
│  🔤 Kata Target                                                  │
│  [______________________________] [📂 Muat .txt…]               │
├────────────────────────────────────────────────────────────────────┤
│  ⚙️ Pengaturan Proses                                            │
│  ☑ Kata utuh saja  ☐ Case-sensitive                             │
│  🎨 Format: [Italic|Bold|Underline|Highlight|Warna]             │
│  🛠️ Engine: python-docx (cepat)                                 │
├────────────────────────────────────────────────────────────────────┤
│         [ ✨  PROSES DOKUMEN  ✨ ]                               │
├────────────────────────────────────────────────────────────────────┤
│  📋 Log Proses                                                   │
│  [09:15:02] 🚀 Memulai proses untuk: dokumen.docx               │
│  [09:15:03] ✅ SELESAI — 42 kata berhasil diformat              │
│  [09:15:03] 💾 Hasil disimpan ke: dokumen_italik.docx           │
└────────────────────────────────────────────────────────────────────┘
```

---

## 🚀 Instalasi

### Persyaratan

- Python 3.8+
- Windows (untuk engine MS Word COM)

### Install Dependencies

```bash
# Wajib
pip install -r requirements.txt

# Atau install manual
pip install customtkinter python-docx

# Opsional - Drag & Drop
pip install tkinterdnd2

# Opsional - Engine MS Word COM (Windows + MS Word terinstall)
pip install pywin32
```

### Jalankan Aplikasi

```bash
# Dari source code
python italikin.py

# Atau install sebagai package
pip install -e .
italikin
```

---

## 📦 Build .exe (Executable)

### Persyaratan Build

- Python 3.8+
- Windows
- PyInstaller

### Build Script

```bash
# Jalankan script build (Windows)
build.bat

# Atau manual
pip install pyinstaller
pyinstaller italikin.spec --clean --noconfirm
```

### Output

```
dist/
└── Italikin/
    └── Italikin.exe    # File executable
```

### Distribusi

File `.exe` bisa didistribusikan tanpa perlu install Python. Cukup:
1. Build .exe dengan script di atas
2. Copy folder `dist/Italikin` ke komputer client
3. Jalankan `Italikin.exe`

---

## 📖 Cara Penggunaan

### 1. Pilih Dokumen

- Klik **📁 Browse…** atau seret file .docx ke jendela aplikasi
- Atau ketik path file langsung di kolom input

### 2. Masukkan Kata Target

**Cara 1: Ketik langsung**
```
PT Maju Jaya, Sumber Daya Manusia, penting
```

**Cara 2: Muat dari file .txt**
- Buat file .txt berisi kata target (dipisahkan koma atau baris baru)
- Klik **📂 Muat .txt…** atau seret file .txt ke jendela

Contoh isi file .txt:
```
PT Maju Jaya
Sumber Daya Manusia
penting, kritis
```

### 3. Atur Pengaturan

| Pengaturan | Opsi | Keterangan |
|------------|------|------------|
| **Kata utuh saja** | ON/OFF | "PT" tidak mengenai "PTRA" |
| **Case-sensitive** | ON/OFF | Huruf besar/kecil harus persis |
| **Format** | Italic/Bold/Underline/Highlight/Warna | Format output yang diinginkan |
| **Warna** | Merah/Biru/Hijau/Oranye/Ungu | Hanya aktif jika format = Warna |
| **Engine** | python-docx / MS Word COM | COM butuh MS Word + pywin32 |

### 4. Proses Dokumen

Klik tombol **✨ PROSES DOKUMEN ✨** dan tunggu hingga selesai.

### 5. Hasil

- File output disimpan dengan suffix `_italik`
- Contoh: `dokumen.docx` → `dokumen_italik.docx`
- File asli **tidak akan diubah**

---

## 🛠️ Engine

### python-docx (Default)

- **Kelebihan**: Cepat, tidak butuh MS Word
- **Kekurangan**: Hanya memproses badan dokumen + tabel
- **Cocok untuk**: Dokumen standar tanpa header/footer penting

### MS Word COM

- **Kelebihan**: Mengcover header, footer, footnote, endnote, textbox
- **Kekurangan**: Lebih lambat, butuh MS Word terinstall + pywin32
- **Cocok untuk**: Dokumen resmi dengan header/footer yang perlu diformat

---

## 📁 Struktur File

```
.
├── italikin.py          # Aplikasi utama (GUI + Core)
├── pyproject.toml       # Konfigurasi Python packaging
├── setup.py             # Backward compatibility
├── requirements.txt     # Dependencies
├── italikin.spec        # PyInstaller spec file
├── build.bat            # Build script untuk .exe
├── file_version_info.txt # Version info untuk .exe
├── README.md            # Dokumentasi project
├── CONTEXT.md           # Context untuk AI agent
├── _t.docx              # Contoh dokumen sumber
├── _t_italik.docx       # Contoh output hasil proses
└── __pycache__/         # Cache Python
```

---

## 🔧 Troubleshooting

### File terkunci / PermissionError

**Penyebab**: File sedang dibuka di Microsoft Word

**Solusi**:
1. Tutup file tersebut di Microsoft Word
2. Klik **✨ PROSES DOKUMEN ✨** lagi

### MS Word COM tidak tersedia

**Penyebab**: pywin32 belum terinstall atau MS Word tidak terinstall

**Solusi**:
```bash
pip install pywin32
```

Pastikan Microsoft Word terinstall di komputer Anda.

### Drag & Drop tidak berfungsi

**Penyebab**: tkinterdnd2 belum terinstall

**Solusi**:
```bash
pip install tkinterdnd2
```

Atau gunakan tombol Browse untuk memilih file.

### Aplikasi freeze / tidak merespons

**Penyebab**: Dokumen sangat besar atau banyak kata target

**Catatan**: Aplikasi menggunakan background thread, sehingga GUI tetap responsif. Jika masih freeze, coba kurangi jumlah kata target.

---

## 💡 Tips

1. **Gunakan kata utuh saja** untuk pencarian lebih presisi
2. **Muat dari .txt** jika kata target banyak (lebih dari 10)
3. **Tutup MS Word** sebelum memproses untuk menghindari error
4. **Gunakan engine COM** jika dokumen memiliki header/footer penting
5. **Periksa log** untuk melihat detail proses
6. **Gunakan menu Bantuan** untuk panduan lengkap

---

## 📄 Lisensi

MIT License

---

## 👥 Kontribusi

Contributions diterima! Silakan buka issue atau pull request.

---

## 🔗 Links

- [Python](https://python.org)
- [python-docx](https://python-docx.readthedocs.io/)
- [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter)
- [pywin32](https://github.com/mhammond/pywin32)
- [PyInstaller](https://pyinstaller.org/)
