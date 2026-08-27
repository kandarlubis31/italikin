# Italikin — Project Context

## Overview

**Italikin** adalah aplikasi desktop Python untuk memformat kata-kata target secara otomatis di dokumen Microsoft Word (.docx). Aplikasi ini mengubah format (italic, bold, underline, highlight, warna) pada kata tertentu tanpa merusak layout, gambar, atau formatting lain di dokumen.

## Purpose

Memudahkan pengguna untuk menandai/memformat istilah penting di dokumen Word secara massal — misalnya untuk keperluan legal review, editing naskah, pembuatan indeks, atau penekanan istilah teknis.

## Architecture

### Single-file Application (`italikin.py`)

Aplikasi ini dibangun dalam satu file Python (~600 baris) dengan struktur:

```
┌─────────────────────────────────────────────┐
│  CONSTANTS & CONFIG                         │
│  - APP_NAME, APP_VERSION, OUTPUT_SUFFIX     │
│  - COLOR_CHOICES, FORMAT_MODES              │
├─────────────────────────────────────────────┤
│  CORE LOGIC (Document Processing)           │
│  - parse_words() / load_words_from_txt()    │
│  - build_pattern() — regex builder          │
│  - make_formatter() — format applicator     │
│  - collect_paragraphs() — paragraph gatherer│
│  - format_paragraph() — run-level splitting │
│  - process_document() — python-docx engine  │
│  - process_document_com() — MS Word COM     │
├─────────────────────────────────────────────┤
│  GUI LAYER (CustomTkinter)                  │
│  - ItalikinApp(BaseWindow)                  │
│  - Dark Mode, Drag & Drop (optional)        │
│  - Background threading for processing      │
└─────────────────────────────────────────────┘
```

### Processing Engines

| Engine | Library | Kecepatan | Coverage |
|--------|---------|-----------|----------|
| `python-docx` | python-docx | Cepat | Badan dokumen + tabel |
| `MS Word COM` | pywin32 | Lambat | + Header, footer, footnote, endnote, textbox |

Engine COM hanya tersedia di Windows dengan MS Word terinstall.

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

### Dual Engine Pattern

Kedua engine menghasilkan output ke file baru (`*_italik.docx`) — file asli tidak pernah dimodifikasi.

### Threading

GUI berjalan di main thread, processing di background thread via `threading.Thread`. Callback `progress_cb` dikirim ke UI via `self.after(0, ...)` untuk update progress bar.

## File Structure

```
.
├── italikin.py          # Aplikasi utama (GUI + Core)
├── _t.docx              # Contoh dokumen sumber (test)
├── _t_italik.docx       # Contoh output hasil proses
├── CONTEXT.md           # Dokumen ini
└── __pycache__/         # Cache Python
```

## Dependencies

### Required
- `customtkinter` — GUI framework (Dark Mode)
- `python-docx` — Manipulasi dokumen .docx

### Optional
- `tkinterdnd2` — Drag & drop file ke jendela aplikasi
- `pywin32` — Engine MS Word COM (Windows only)

## Configuration

Tidak ada file konfigurasi terpisah. Semua pengaturan diatur melalui GUI:
- **Format**: Italic / Bold / Underline / Highlight / Warna (Merah, Biru, Hijau, Oranye, Ungu)
- **Pencarian**: Kata utuh saja (whole word) / Case-sensitive
- **Engine**: python-docx / MS Word COM

## Naming Conventions

- **Output file**: Menambahkan suffix `_italik` sebelum ekstensi (misal: `dokumen_italik.docx`)
- **Variabel**: snake_case
- **Kelas**: PascalCase (`ItalikinApp`)
- **Konstanta**: UPPER_SNAKE_CASE

## Language

Seluruh UI dan komentar menggunakan **Bahasa Indonesia**.

## State Management

State utama disimpan sebagai instance variables di `ItalikinApp`:
- `self.docx_path` — path dokumen yang dipilih
- `self.is_processing` — flag apakah sedang memproses (mencegah double-click)
- UI widgets: `entry_docx`, `entry_words`, `sw_whole_word`, `sw_case_sensitive`, `seg_format`, `menu_color`, `menu_engine`

## Testing

Belum ada test files atau test framework yang terpasang. Untuk testing manual:
1. Siapkan dokumen `.docx` test
2. Jalankan `python italikin.py`
3. Pilih dokumen, masukkan kata target, proses
4. Bandingkan output dengan file asli

## Known Limitations

1. Tidak mendukung format `.doc` lama (hanya `.docx`)
2. File yang sedang dibuka di MS Word akan terkunci (PermissionError)
3. Engine COM hanya berjalan di Windows dengan MS Word
4. Drag & drop membutuhkan `tkinterdnd2` (opsional)
5. Tidak ada undo — output selalu file baru
