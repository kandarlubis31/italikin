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
- **File Copy** — Original 100% aman, file di-copy ke workspace
- **Export Dialog** — Pilih lokasi & nama file output
- **AI Assistant** — Generate list kata target dengan ChatGPT/Gemini
- **Keyboard Shortcuts** — 8 shortcuts untuk produktivitas
- **Recent Files** — Track 5 file terakhir
- **Save/Load Settings** — Pengaturan tersimpan otomatis
- **Drag & Drop** — Seret file .docx/.txt langsung ke jendela aplikasi
- **GUI Modern** — CustomTkinter dengan Dark Mode

---

## 📸 Tampilan Aplikasi

```
┌────────────────────────────────────────────────────────────────────┐
│  📁 File    🛠️ Tools    ❓ Bantuan                                │
├────────────────────────────────────────────────────────────────────┤
│  ⭐ ITALIKIN                                        v1.0.0       │
│  Format kata di dokumen Word tanpa merusak layout                 │
├────────────────────────────────────────────────────────────────────┤
│  📄 File Dokumen Word (.docx)                                    │
│  [______________________________] [📁 Pilih Dokumen]             │
├────────────────────────────────────────────────────────────────────┤
│  🔤 Kata Target                                                  │
│  [______________________________] [📋 Pilih Kata]                │
│  📊 3 kata target: PT Maju Jaya, Perjanjian Kerja, Klausul      │
├────────────────────────────────────────────────────────────────────┤
│  ⚙️ Pengaturan Proses                                            │
│  ☑ Kata utuh saja  ☐ Case-sensitive                             │
│  🎨 Format: [Italic|Bold|Underline|Highlight|Warna]             │
│  🛠️ Engine: [MS Word COM (Recommended) ▼]                       │
├────────────────────────────────────────────────────────────────────┤
│         [ ✨  PROSES DOKUMEN  ✨ ]                               │
├────────────────────────────────────────────────────────────────────┤
│  📋 Log Proses                                                   │
│  [09:15:02] 🚀 Memulai proses untuk: dokumen.docx               │
│  [09:15:02] 📋 Memproses COPY (original tetap aman)              │
│  [09:15:03] ✅ SELESAI — 42 kata berhasil diformat              │
│  [09:15:03] 💾 Hasil: dokumen_italik.docx                       │
└────────────────────────────────────────────────────────────────────┘
```

---

## 🔬 Precision & Testing

### Word Boundary Matching

Italikin menggunakan regex word boundary untuk pencarian yang presisi:

```
whole_word=ON (Default):
┌─────────────────────────────────────────────────────────────────┐
│  ✅ "PT adalah"     → MATCH     (PT di awal kata)              │
│  ✅ "di PT."         → MATCH     (PT sebelum titik)             │
│  ✅ "(PT)"           → MATCH     (PT dalam kurung)              │
│  ❌ "PTP tidak"     → NO MATCH  (PTP bukan PT)                 │
│  ❌ "PTXYZ juga"    → NO MATCH  (PTXYZ bukan PT)               │
│  ❌ "aPTb"          → NO MATCH  (PT di tengah huruf)            │
│  ❌ "manajemenrisiko" → NO MATCH (tanpa spasi)                  │
└─────────────────────────────────────────────────────────────────┘

whole_word=OFF:
┌─────────────────────────────────────────────────────────────────┐
│  ✅ "PTMaju"        → MATCH     (substring ditemukan)           │
│  ✅ "PT adalah"     → MATCH                                    │
│  ✅ "manajemenrisiko" → MATCH   (substring ditemukan)           │
└─────────────────────────────────────────────────────────────────┘
```

### Case Sensitivity

```
case_sensitive=OFF (Default):
┌─────────────────────────────────────────────────────────────────┐
│  "Pt Maju" matches:                                            │
│    ✅ "Pt Maju adalah"   (exact)                                │
│    ✅ "pt maju adalah"   (lowercase)                            │
│    ✅ "PT MAJU"          (uppercase)                            │
│    ✅ "pT mAjU"          (mixed case)                           │
└─────────────────────────────────────────────────────────────────┘

case_sensitive=ON:
┌─────────────────────────────────────────────────────────────────┐
│  "Pt Maju" matches:                                            │
│    ✅ "Pt Maju adalah"   (exact)                                │
│    ❌ "pt maju adalah"   (lowercase - no match)                 │
│    ❌ "PT MAJU"          (uppercase - no match)                 │
│    ❌ "pT mAjU"          (mixed case - no match)                │
└─────────────────────────────────────────────────────────────────┘
```

### Test Results (45/45 Passed ✅)

```
┌─────────────────────────────────────────────────────────────────┐
│                    PRECISION TESTS                              │
├─────────────────────────────────────────────────────────────────┤
│  [PASS] Word Boundary        10/10                              │
│  [PASS] Substring             7/7                               │
│  [PASS] Case Sensitivity      6/6                               │
│  [PASS] Multiple Words        3/3                               │
├─────────────────────────────────────────────────────────────────┤
│                    HUMAN ERROR TESTS                            │
├─────────────────────────────────────────────────────────────────┤
│  [PASS] Empty Input           6/6                               │
│  [PASS] Invalid File          3/3                               │
│  [PASS] Special Characters    8/8                               │
│  [PASS] Large Word List       1/1                               │
│  [PASS] Format Preservation   1/1                               │
│  [PASS] Concurrent            1/1                               │
├─────────────────────────────────────────────────────────────────┤
│  TOTAL: 45 passed, 0 failed                     ✅ ALL PASSED  │
└─────────────────────────────────────────────────────────────────┘
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
    ├── Italikin.exe    # File executable (5.0 MB)
    └── _internal/      # Dependencies
```

### Distribusi

File `.exe` bisa didistribusikan tanpa perlu install Python:
1. Build .exe dengan script di atas
2. Copy folder `dist/Italikin` ke komputer client
3. Jalankan `Italikin.exe`

---

## 📖 Cara Penggunaan

### 1. Pilih Dokumen

- Klik **📁 Pilih Dokumen** atau seret file .docx ke jendela aplikasi
- File akan di-copy ke workspace (original tetap aman)

### 2. Masukkan Kata Target

**Cara 1: Ketik langsung**
```
PT Maju Jaya, Sumber Daya Manusia, penting
```

**Cara 2: Muat dari file .txt**
- Buat file .txt berisi kata target (dipisahkan koma atau baris baru)
- Klik **📋 Pilih Kata** atau seret file .txt ke jendela

### 3. Atur Pengaturan

| Pengaturan | Default | Keterangan |
|------------|---------|------------|
| **Kata utuh saja** | ON | "PT" tidak mengenai "PTP" |
| **Case-sensitive** | OFF | "pt" match "PT" |
| **Format** | Italic | Pilihan: Italic/Bold/Underline/Highlight/Warna |
| **Engine** | MS Word COM | Recommended untuk hasil terbaik |

### 4. Proses & Export

1. Klik **✨ PROSES DOKUMEN ✨**
2. Tunggu hingga selesai
3. Pilih lokasi export di dialog yang muncul
4. Klik **💾 Export** atau **📁 Buka Folder**

### 5. Hasil

```
Original: D:\Kerja\Kontrak.docx (tidak berubah)
Output:   D:\Kerja\Kontrak_italik.docx (file baru)
```

---

## 🤖 AI Workflow

Gunakan AI untuk generate list kata target secara otomatis:

```
┌─────────────────────────────────────────────────────────────────┐
│  STEP 1: Upload Dokumen ke AI                                  │
│  ┌─────────────────────────────────────────────────────────┐   │
│  │  Upload .docx/.pdf ke ChatGPT / Gemini / Claude         │   │
│  └─────────────────────────────────────────────────────────┘   │
│                          ↓                                      │
│  STEP 2: Minta AI Generate List                               │
│  ┌─────────────────────────────────────────────────────────┐   │
│  │  "Buatkan daftar istilah hukum penting dari dokumen     │   │
│  │   ini. Format: kata1, kata2, kata3 (dipisah koma)"     │   │
│  └─────────────────────────────────────────────────────────┘   │
│                          ↓                                      │
│  STEP 3: Copy ke Italikin                                      │
│  ┌─────────────────────────────────────────────────────────┐   │
│  │  Copy output AI → Paste ke kolom "Kata Target"          │   │
│  └─────────────────────────────────────────────────────────┘   │
│                          ↓                                      │
│  STEP 4: Proses Dokumen                                        │
│  ┌─────────────────────────────────────────────────────────┐   │
│  │  Klik PROSES → Export → Selesai!                        │   │
│  └─────────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────────┘
```

**Akses AI Assistant**: Menu 🛠️ Tools → 🤖 AI Assistant

---

## 📁 Struktur File

```
.
├── italikin.py              # Aplikasi utama (GUI + Core)
├── italikin.spec            # PyInstaller spec file
├── pyproject.toml           # Konfigurasi Python packaging
├── setup.py                 # Backward compatibility
├── requirements.txt         # Dependencies
├── build.bat                # Build script untuk .exe
├── file_version_info.txt    # Version info untuk .exe
├── README.md                # Dokumentasi ini
├── CONTEXT.md               # Context untuk AI agent
├── .gitignore               # Git ignore rules
├── .github/
│   └── workflows/
│       └── build.yml        # GitHub Actions auto-build
└── dist/
    └── Italikin/
        └── Italikin.exe     # Built executable (5.0 MB)
```

---

## 🛠️ Engine Comparison

| Fitur | python-docx | MS Word COM |
|-------|-------------|-------------|
| **Kecepatan** | ⚡ Cepat | 🐢 Lambat |
| **Tanpa MS Word** | ✅ Ya | ❌ Tidak |
| **Badan Dokumen** | ✅ Ya | ✅ Ya |
| **Tabel** | ✅ Ya | ✅ Ya |
| **Header/Footer** | ❌ Tidak | ✅ Ya |
| **Footnote/Endnote** | ❌ Tidak | ✅ Ya |
| **Textbox** | ❌ Tidak | ✅ Ya |
| **Rekomendasi** | Basic use | ✅ **Recommended** |

---

## 🔧 Troubleshooting

### File terkunci / PermissionError

**Penyebab**: File sedang dibuka di Microsoft Word

**Solusi**: Tutup file tersebut di Microsoft Word, lalu coba lagi.

### MS Word COM tidak tersedia

**Solusi**:
```bash
pip install pywin32
```

Pastikan Microsoft Word terinstall di komputer Anda.

### Drag & Drop tidak berfungsi

**Solusi**:
```bash
pip install tkinterdnd2
```

Atau gunakan tombol Browse untuk memilih file.

---

## ⌨️ Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl+O` | Buka dokumen |
| `Ctrl+Shift+O` | Muat daftar kata |
| `Ctrl+Enter` | Proses dokumen |
| `Ctrl+E` | Export |
| `Ctrl+L` | Clear log |
| `Ctrl+S` | Simpan settings |
| `Ctrl+Shift+S` | Muat settings |
| `F1` | Help |

---

## 📄 Lisensi

MIT License

---

## 🔗 Links

- [GitHub Repository](https://github.com/kandarlubis31/italikin)
- [Releases](https://github.com/kandarlubis31/italikin/releases)
- [Python](https://python.org)
- [python-docx](https://python-docx.readthedocs.io/)
- [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter)
- [pywin32](https://github.com/mhammond/pywin32)
- [PyInstaller](https://pyinstaller.org/)

---

Made with ❤️ by Italikin Team
