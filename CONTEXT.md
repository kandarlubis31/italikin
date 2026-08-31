# Italikin — Project Context (v1.0.0)

## Overview

**Italikin** adalah aplikasi desktop Python untuk memformat kata-kata target secara otomatis di dokumen Microsoft Word (.docx). Aplikasi ini mengubah format (italic, bold, underline, highlight, warna) pada kata tertentu tanpa merusak layout, gambar, atau formatting lain di dokumen.

## Purpose

Memudahkan pengguna untuk menandai/memformat istilah penting di dokumen Word secara massal — misalnya untuk keperluan legal review, editing naskah, pembuatan indeks, atau penekanan istilah teknis.

## Architecture

### Single-file Application (`italikin.py`)

Aplikasi ini dibangun dalam satu file Python (~1722 baris) dengan struktur:

```
┌─────────────────────────────────────────────────────────┐
│  CONSTANTS & CONFIG                                     │
│  - APP_NAME, APP_VERSION, OUTPUT_SUFFIX                 │
│  - COLOR_CHOICES, FORMAT_MODES, ENGINE_*                │
├─────────────────────────────────────────────────────────┤
│  CORE LOGIC (Document Processing)                       │
│  - parse_words() / load_words_from_txt()                │
│  - build_pattern() — regex builder                      │
│  - make_formatter() — format applicator                 │
│  - collect_paragraphs() — paragraph gatherer            │
│  - format_paragraph() — run-level splitting             │
│  - process_document() — python-docx engine              │
│  - process_document_com() — MS Word COM engine          │
├─────────────────────────────────────────────────────────┤
│  GUI LAYER (CustomTkinter)                              │
│  - ItalikinApp(BaseWindow)                              │
│  - Menu Bar (File, Tools, Bantuan)                      │
│  - Dark Mode, Drag & Drop (optional)                    │
│  - Keyboard shortcuts (8 shortcuts)                     │
│  - Recent files, Save/Load settings                     │
│  - AI Assistant dialog                                  │
│  - Export panel (integrated)                            │
│  - Background threading for processing                  │
└─────────────────────────────────────────────────────────┘
```

### Processing Engines

| Engine | Library | Kecepatan | Coverage | Rekomendasi |
|--------|---------|-----------|----------|-------------|
| `MS Word COM` | pywin32 | Lambat | Header, footer, footnote, endnote, textbox | ✅ Default |
| `python-docx` | python-docx | Cepat | Badan dokumen + tabel | ⚠️ Basic |

**MS Word COM** direkomendasikan karena mengakses MS Word asli, menghasilkan format yang lebih rapi dan tidak merusak text.

### File Handling: Copy-First Pattern

```
Original File → COPY ke workspace → Proses COPY → Export ke lokasi pilihan
     ↓                ↓                  ↓                ↓
  (aman)        (italikin_workspace)   (output)     (user pilih)
```

- **Original**: Tidak pernah disentuh
- **Workspace**: Folder `italikin_workspace/` di samping executable
- **Output**: Disimpan di workspace, lalu user export ke lokasi pilihan

## Key Technical Concepts

### Run-Level Formatting

Inti dari Italikin adalah pemrosesan di level **run** (bukan paragraph). Setiap paragraf dipecah menjadi run-run kecil, lalu regex dijalankan terhadap teks gabungan. Saat match ditemukan:

1. Run dipecah (`_split_run_at`) agar match berada di awal/akhir run
2. Run yang berisi kata target diformat
3. Casing asli kata **dipertahankan** (hanya properti format yang diubah)

Algoritma ini memastikan:
- Font-size, bold, warna, layout tidak bergeser
- Tabel dan gambar tidak terganggu
- Format asli run lain tidak berubah

### Word Boundary Matching

```
whole_word=True:
  ✅ "PT adalah"    → MATCH (PT di awal kata)
  ✅ "di PT."        → MATCH (PT di akhir kata)
  ❌ "PTP tidak"    → NO MATCH (PTP bukan PT)
  ❌ "PTXYZ juga"   → NO MATCH (PTXYZ bukan PT)
  ❌ "aPTb"         → NO MATCH (PT di tengah huruf)

whole_word=False:
  ✅ "PTMaju"       → MATCH (substring)
  ✅ "PT adalah"    → MATCH
```

### Case Sensitivity

```
case_sensitive=True:
  "Pt Maju" matches "Pt Maju" ✅
  "Pt Maju" matches "pt maju" ❌
  "Pt Maju" matches "PT MAJU" ❌

case_sensitive=False:
  "Pt Maju" matches "pt maju" ✅
  "Pt Maju" matches "PT MAJU" ✅
```

### Dual Engine Pattern

Kedua engine menghasilkan output ke file baru (`*_italik.docx`) — file asli tidak pernah dimodifikasi.

### Threading

GUI berjalan di main thread, processing di background thread via `threading.Thread`. Callback `progress_cb` dikirim ke UI via `self.after(0, ...)` untuk update progress bar.

## File Structure

```
.
├── italikin.py              # Aplikasi utama (GUI + Core, 1722 lines)
├── italikin.spec            # PyInstaller spec file
├── pyproject.toml           # Package configuration
├── setup.py                 # Backward compatibility
├── requirements.txt         # Dependencies
├── build.bat                # Build script for .exe
├── file_version_info.txt    # Version info for .exe
├── README.md                # Full documentation
├── CONTEXT.md               # This file
├── .gitignore               # Git ignore rules
├── .github/
│   └── workflows/
│       └── build.yml        # GitHub Actions auto-build
└── dist/
    └── Italikin/
        └── Italikin.exe     # Built executable (5.0 MB)
```

## Dependencies

### Required
- `customtkinter>=5.2.0` — GUI framework (Dark Mode)
- `python-docx>=1.0.0` — Manipulasi dokumen .docx

### Optional
- `tkinterdnd2>=0.3.0` — Drag & drop file ke jendela aplikasi
- `pywin32>=306` — Engine MS Word COM (Windows only)

### Build
- `pyinstaller>=6.0` — Build .exe executable

## Configuration

### User Settings (auto-saved)
- `whole_word` — Kata utuh saja (default: ON)
- `case_sensitive` — Case-sensitive (default: OFF)
- `format` — Italic/Bold/Underline/Highlight/Warna
- `color` — Warna kustom (Merah/Biru/Hijau/Oranye/Ungu)
- `engine` — MS Word COM / python-docx
- `recent_files` — 5 file terakhir

### Keyboard Shortcuts
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

## Naming Conventions

- **Output file**: Menambahkan suffix `_italik` sebelum ekstensi (misal: `dokumen_italik.docx`)
- **Variabel**: snake_case
- **Kelas**: PascalCase (`ItalikinApp`)
- **Konstanta**: UPPER_SNAKE_CASE

## Language

Seluruh UI dan komentar menggunakan **Bahasa Indonesia**.

## State Management

State utama disimpan sebagai instance variables di `ItalikinApp`:
- `self.original_path` — path file asli (tidak disentuh)
- `self.working_copy_path` — path copy di workspace
- `self.working_dir` — working directory untuk proses
- `self.output_path` — path output terakhir
- `self.is_processing` — flag apakah sedang memproses
- `self.recent_files` — 5 file terakhir
- `self.settings` — settings tersimpan

## Testing

### Test Results (v1.0.0)

**Precision Tests:**
- Word Boundary: 10/10 passed ✅
- Substring: 7/7 passed ✅
- Case Sensitivity: 6/6 passed ✅
- Multiple Words: 3/3 passed ✅

**Human Error Tests:**
- Empty Input: 6/6 passed ✅
- Invalid File: 3/3 passed ✅
- Special Characters: 8/8 passed ✅
- Large Word List: 1/1 passed ✅
- Format Preservation: 1/1 passed ✅
- Concurrent Processing: 1/1 passed ✅

**Total: 45/45 tests passed**

### Test Coverage

| Category | Tests | Status |
|----------|-------|--------|
| Word Boundary | 10 | ✅ |
| Substring | 7 | ✅ |
| Case Sensitivity | 6 | ✅ |
| Multiple Words | 3 | ✅ |
| Empty Input | 6 | ✅ |
| Invalid File | 3 | ✅ |
| Special Characters | 8 | ✅ |
| Large Word List | 1 | ✅ |
| Format Preservation | 1 | ✅ |
| Concurrent | 1 | ✅ |

## Known Limitations

1. Tidak mendukung format `.doc` lama (hanya `.docx`)
2. File yang sedang dibuka di MS Word akan terkunci (PermissionError)
3. Engine COM hanya berjalan di Windows dengan MS Word
4. Drag & drop membutuhkan `tkinterdnd2` (opsional)
5. Tidak ada undo — output selalu file baru
6. Word boundary: dash `-` dianggap non-word char (standard regex behavior)

## AI Integration

### AI Assistant Feature
- Template copy untuk prompt AI
- Quick link ke ChatGPT, Gemini, Claude
- Workflow: Upload doc → AI generate list → Copy paste ke Italikin

### Workflow
```
1. Upload dokumen ke AI (ChatGPT/Gemini/Claude)
2. Copy template prompt dari Italikin
3. AI generate list kata (comma-separated)
4. Copy paste ke kolom Kata Target
5. Proses dokumen!
```

## Release Info

- **Version**: v1.0.0
- **Release Date**: 2026-08-27
- **GitHub**: https://github.com/kandarlubis31/italikin
- **Download**: https://github.com/kandarlubis31/italikin/releases/tag/v1.0.0
- **Executable Size**: 5.0 MB
