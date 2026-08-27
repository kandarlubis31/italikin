# -*- coding: utf-8 -*-
"""
==============================================================================
 ITALIKIN — Pembuat Format Otomatis untuk Dokumen MS Word (.docx)
==============================================================================
 Fitur:
   - Deteksi kata target secara CASE-INSENSITIVE di seluruh dokumen
     (termasuk di dalam tabel dan tabel bersarang).
   - Manipulasi dilakukan pada level RUN (paragraph.runs) sehingga format
     asli (font-size, bold, warna, tabel, gambar, layout) TIDAK hilang
     dan TIDAK bergeser.
   - Casing asli kata dipertahankan, hanya properti format yang diubah.
   - OPSI PENCARIAN PRESISI:
       * "Kata utuh saja"  -> "PT" tidak lagi mengenai "PTRA" (whole word)
       * "Case-sensitive"  -> huruf besar/kecil harus persis sama
   - FORMAT OUTPUT FLEKSIBEL: Italic / Bold / Underline / Highlight /
     Warna kustom (Merah, Biru, Hijau, Oranye, Ungu).
   - DUAL ENGINE:
       * python-docx  -> cepat, tanpa MS Word, proses badan dokumen + tabel.
       * MS Word COM  -> otomatis mengcover header, footer, footnote,
         endnote, dan textbox (butuh MS Word ter-install + pywin32).
   - DRAG & DROP: seret file .docx / .txt langsung ke jendela aplikasi
     (opsional — aktif otomatis bila package 'tkinterdnd2' terpasang).
   - GUI modern berbasis CustomTkinter (Dark Mode default).
   - Bebas asset eksternal: semua ikon memakai unicode standar.

 Dependensi wajib:
   pip install customtkinter python-docx
 Dependensi opsional:
   pip install tkinterdnd2   (drag & drop)
   pip install pywin32       (engine MS Word COM)
==============================================================================
"""

import copy
import os
import re
import shutil
import sys
import tempfile
import threading
from datetime import datetime
from pathlib import Path

import customtkinter as ctk
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import RGBColor
from docx.text.run import Run
from tkinter import filedialog, messagebox

# --- Drag & Drop opsional (tidak wajib; aplikasi tetap jalan tanpanya) ------
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

# --- Engine MS Word COM opsional (butuh pywin32 + MS Word ter-install) -----
try:
    import pythoncom
    import win32com.client as _win32com
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False


# ============================================================================
# KONSTANTA APLIKASI
# ============================================================================

APP_NAME = "Italikin"
APP_VERSION = "1.0.0"
OUTPUT_SUFFIX = "_italik"

# Konstanta enum MS Word (wd*) untuk engine COM
WD_STORY_MAIN = 6            # wdMainTextStory
WD_FIND_CONTINUE = 1         # wdFindContinue
WD_REPLACE_ALL = 2           # wdReplaceAll
WD_UNDERLINE_SINGLE = 1      # wdUnderlineSingle
WD_FORMAT_XML_DOCUMENT = 12  # wdFormatXMLDocument (.docx)

# Pilihan label engine di GUI
ENGINE_COM = "MS Word COM (Recommended)"
ENGINE_DOCX = "python-docx (Basic - dapat merusak text)"

# Pilihan warna kustom untuk format output "Warna"
COLOR_CHOICES = {
    "Merah":   "FF0000",
    "Biru":    "0000FF",
    "Hijau":   "008000",
    "Oranye":  "FFA500",
    "Ungu":    "800080",
}

# Pemetaan label tombol segmented -> kode mode internal
FORMAT_MODES = ["Italic", "Bold", "Underline", "Highlight", "Warna"]

# URL & Info
APP_URL = "https://github.com/yourusername/italikin"
APP_AUTHOR = "Italikin Team"
APP_DESCRIPTION = "Aplikasi untuk memformat kata tertentu di dokumen Word secara otomatis"


# ============================================================================
# BAGIAN CORE : LOGIKA PENCARIAN & MANIPULASI RUN DI DOKUMEN WORD
# ============================================================================

def parse_words(raw_text: str) -> list:
    """
    Mengubah string mentah (dipisahkan koma / baris baru) menjadi daftar kata
    bersih tanpa duplikat dan tanpa elemen kosong.
    """
    if not raw_text:
        return []
    candidates = re.split(r"[,\n\r]+", raw_text)
    seen = set()
    words = []
    for w in candidates:
        w = w.strip()
        if w and w.lower() not in seen:
            seen.add(w.lower())
            words.append(w)
    return words


def load_words_from_txt(txt_path: str) -> list:
    """
    Membaca file .txt berisi daftar kata (dipisahkan koma atau baris baru)
    lalu mengembalikan daftar kata bersih.
    """
    with open(txt_path, "r", encoding="utf-8-sig") as f:
        content = f.read()
    return parse_words(content)


def build_pattern(words: list, whole_word: bool = False,
                  case_sensitive: bool = False):
    """
    Membangun objek Regex gabungan dari seluruh kata target.

    - Kata yang lebih panjang didahulukan agar substring tidak 'memakan'
      kata yang lebih spesifik (mis. 'PT' vs 'PT XYZ').
    - whole_word=True   : bungkus dengan lookaround (?<!\\w)...(?!\\w)
                          sehingga hanya match jika berdiri sebagai kata
                          utuh ('PT' TIDAK mengenai 'PTRA').
    - case_sensitive=False : gunakan flag IGNORECASE (default).
    """
    if not words:
        return None
    escaped = [re.escape(w.strip()) for w in sorted(words, key=len, reverse=True)]
    body = "|".join(escaped)
    if whole_word:
        body = "(?<!\\w)(?:" + body + ")(?!\\w)"
    flags = 0 if case_sensitive else re.IGNORECASE
    return re.compile("(" + body + ")", flags)


def make_formatter(mode: str, color_hex: str = None):
    """
    Mengembalikan fungsi apply(run) yang menerapkan format pilihan pengguna
    pada SATU run hasil pemecahan — tanpa menyentuh formatting lain.
    """
    def apply(run: Run):
        if mode == "italic":
            run.italic = True
        elif mode == "bold":
            run.bold = True
        elif mode == "underline":
            run.underline = True
        elif mode == "highlight":
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        elif mode == "color":
            run.font.color.rgb = RGBColor.from_string(color_hex or "FF0000")
        else:  # fallback aman
            run.italic = True
    return apply


def _iter_cell_paragraphs(cell) -> list:
    """
    Mengambil semua paragraf dari sebuah cell tabel,
    termasuk paragraf di tabel bersarang (nested table) di dalamnya.
    """
    result = list(cell.paragraphs)
    for inner_table in cell.tables:
        result.extend(_iter_table_paragraphs(inner_table))
    return result


def _iter_table_paragraphs(table) -> list:
    """Iterasi semua paragraf di seluruh cell milik sebuah tabel."""
    result = []
    for row in table.rows:
        for cell in row.cells:
            result.extend(_iter_cell_paragraphs(cell))
    return result


def collect_paragraphs(doc) -> list:
    """
    Mengumpulkan SEMUA paragraf dokumen: badan dokumen utama,
    tabel tingkat atas, dan tabel bersarang di dalamnya.
    """
    paragraphs = list(doc.paragraphs)
    for table in doc.tables:
        paragraphs.extend(_iter_table_paragraphs(table))
    return paragraphs


def _locate_run(runs: list, char_index: int):
    """
    Menemukan indeks run dan offset di dalam run tersebut
    untuk sebuah posisi karakter absolut pada teks gabungan paragraf.
    Mengembalikan tuple (run_index, offset_dalam_run).
    """
    position = 0
    for i, run in enumerate(runs):
        length = len(run.text)
        if char_index < position + length:
            return i, char_index - position
        position += length
    # Fallback: posisi tepat di ujung akhir paragraf
    return len(runs) - 1, len(runs[-1].text)


def _clone_run_after(run: Run) -> Run:
    """
    Menduplikasi elemen XML sebuah run (beserta SELURUH formatting-nya:
    font, size, bold, underline, warna, highlight, dll) lalu menempatkan
    salinan itu tepat SETELAH run aslinya.
    """
    new_xml_element = copy.deepcopy(run._r)
    run._r.addnext(new_xml_element)
    return Run(new_xml_element, run._parent)


def _split_run_at(run: Run, offset: int) -> Run:
    """
    Memecah satu run menjadi dua run pada posisi karakter 'offset'.
    Formatting kedua bagian identik karena di-copy via deepcopy XML.
    Run asli menyimpan bagian KIRI, run baru (return value) menyimpan
    bagian KANAN.
    """
    text = run.text
    left_part = text[:offset]
    right_part = text[offset:]
    run.text = left_part
    new_run = _clone_run_after(run)
    new_run.text = right_part
    return new_run


def format_paragraph(paragraph, pattern, formatter) -> int:
    """
    Mencari semua kemunculan kata target pada SATU paragraf lalu
    menerapkan format pada masing-masing kemunculan dengan teknik
    pemecahan run.

    Algoritma:
      1. Gabungkan teks seluruh run menjadi satu string penuh.
      2. Cari match regex pada string penuh.
      3. Pastikan batas AWAL match jatuh tepat di awal sebuah run
         (pecah run bila perlu).
      4. Pastikan batas AKHIR match jatuh tepat di akhir sebuah run
         (pecah run bila perlu).
      5. Run yang sekarang berisi PERSIS kata target di-format.
      6. Lanjut mencari dari posisi setelah match (hindari infinite loop).

    Mengembalikan jumlah kata yang berhasil diformat.
    Casing asli kata otomatis dipertahankan karena kita hanya memindah
    karakter yang sudah ada ke run baru, bukan menulis ulang teksnya.
    """
    count = 0
    search_from = 0

    while True:
        runs = paragraph.runs
        if not runs:
            break

        full_text = "".join(r.text for r in runs)
        match = pattern.search(full_text, search_from)
        if match is None:
            break

        start, end = match.span()

        # ---- Langkah 3: pastikan START match berada di awal sebuah run ----
        idx, off = _locate_run(runs, start)
        if off > 0:
            _split_run_at(runs[idx], off)
            # Setelah split, match kini mulai TEPAT di awal run berikutnya.
            idx += 1
            runs = paragraph.runs  # refresh daftar run setelah struktur berubah

        # ---- Langkah 4: pastikan END match berada di akhir sebuah run ----
        idx_end, off_end = _locate_run(runs, end)
        if off_end > 0:
            _split_run_at(runs[idx_end], off_end)
            runs = paragraph.runs  # refresh lagi

        # ---- Langkah 5: run ke-'idx' sekarang berisi persis kata target ----
        target_run = runs[idx]
        formatter(target_run)
        count += 1

        # ---- Langkah 6: lanjut pencarian setelah akhir match ----
        search_from = end

    return count


def process_document(input_path: str, words: list,
                     format_mode: str = "italic", color_hex: str = None,
                     whole_word: bool = False, case_sensitive: bool = False,
                     progress_cb=None) -> tuple:
    """
    Fungsi utama core: membuka dokumen, memformat semua kata target,
    lalu menyimpan hasilnya sebagai file baru '<nama>_italik.docx'.

    Parameter:
        input_path     : path file .docx sumber
        words          : daftar kata target
        format_mode    : 'italic' | 'bold' | 'underline' | 'highlight' | 'color'
        color_hex      : kode warna hex (mis. 'FF0000') bila format_mode='color'
        whole_word     : True = hanya match kata utuh
        case_sensitive : True = huruf besar/kecil harus persis
        progress_cb    : fungsi callback (current, total, message) opsional

    Return:
        tuple (output_path, jumlah_kata_diubah, total_paragraf_diperiksa)

    Melempar PermissionError bila file sedang dikunci/dibuka MS Word.
    """
    if progress_cb is None:
        progress_cb = lambda cur, tot, msg: None

    # --- Buka dokumen (bisa gagal jika file terkunci oleh MS Word) ---
    try:
        doc = Document(input_path)
    except PermissionError:
        raise PermissionError(
            "File sedang dibuka / dikunci oleh Microsoft Word.\n"
            "Tutup dokumen tersebut terlebih dahulu, lalu coba lagi.\n\n"
            f"File: {os.path.basename(input_path)}"
        )

    pattern = build_pattern(words, whole_word=whole_word,
                            case_sensitive=case_sensitive)
    if pattern is None:
        raise ValueError("Tidak ada kata target yang valid untuk diproses.")

    formatter = make_formatter(format_mode, color_hex)

    paragraphs = collect_paragraphs(doc)
    total_paragraphs = len(paragraphs)

    changed_count = 0
    for i, paragraph in enumerate(paragraphs, start=1):
        changed_count += format_paragraph(paragraph, pattern, formatter)
        if i % 25 == 0 or i == total_paragraphs:
            progress_cb(i, total_paragraphs,
                        f"Memeriksa paragraf {i}/{total_paragraphs}...")

    # --- Simpan hasil ke file BARU agar file asli tidak tersentuh ---
    base, ext = os.path.splitext(input_path)
    output_path = f"{base}{OUTPUT_SUFFIX}{ext or '.docx'}"
    try:
        doc.save(output_path)
    except PermissionError:
        raise PermissionError(
            "Gagal menyimpan hasil: file tujuan sedang dibuka / dikunci "
            "oleh Microsoft Word.\nTutup dokumen tersebut lalu coba lagi.\n\n"
            f"File: {os.path.basename(output_path)}"
        )

    return output_path, changed_count, total_paragraphs


def _hex_to_word_color(color_hex: str) -> int:
    """
    Konversi warna hex 'RRGGBB' menjadi nilai integer BGR yang dipakai
    oleh properti Font.Color milik MS Word COM.
    """
    color_hex = (color_hex or "FF0000").lstrip("#")
    r = int(color_hex[0:2], 16)
    g = int(color_hex[2:4], 16)
    b = int(color_hex[4:6], 16)
    return b * 65536 + g * 256 + r


def process_document_com(input_path: str, words: list,
                         format_mode: str = "italic", color_hex: str = None,
                         whole_word: bool = False, case_sensitive: bool = False,
                         progress_cb=None) -> tuple:
    """
    Engine alternatif via MS Word COM Automation (butuh pywin32 + MS Word).

    Keunggulan: otomatis mengcover header, footer, footnote, endnote,
    dan textbox — tidak hanya badan dokumen & tabel.
    Kekurangan: lebih lambat (1 pass Find per kata) dan butuh Word ter-install.

    Return:
        tuple (output_path, jumlah_istilah_yang_minimal_1x_replace, total_istilah)

    Melempar PermissionError bila file terkunci, RuntimeError bila Word
    tidak dapat dibuka. Proses Word SELALU di-cleanup di blok finally.
    """
    if not COM_AVAILABLE:
        raise RuntimeError(
            "Engine MS Word COM membutuhkan package 'pywin32'.\n"
            "Install dulu dengan: pip install pywin32"
        )
    if progress_cb is None:
        progress_cb = lambda cur, tot, msg: None

    # Frasa panjang diproses duluan agar tidak 'dimakan' kata pendek
    sorted_words = sorted(set(words), key=len, reverse=True)
    total_terms = len(sorted_words)

    base, ext = os.path.splitext(input_path)
    output_path = f"{base}{OUTPUT_SUFFIX}{ext or '.docx'}"

    abs_input = os.path.abspath(input_path)
    abs_output = os.path.abspath(output_path)

    word = None
    doc = None
    replaced_terms = 0

    # COM wajib di-initialize per-thread (proses berjalan di worker thread)
    pythoncom.CoInitialize()
    try:
        try:
            word = _win32com.Dispatch("Word.Application")
        except Exception as exc:
            raise RuntimeError(
                f"MS Word tidak dapat dibuka di komputer ini.\n"
                f"Pastikan Microsoft Word ter-install.\n\nDetail: {exc}"
            )

        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone — jangan munculkan dialog apapun

        try:
            doc = word.Documents.Open(abs_input)
        except Exception as exc:
            raise PermissionError(
                "File tidak dapat dibuka oleh MS Word.\n"
                "Kemungkinan besar dokumen sedang terbuka/dikunci.\n\n"
                f"File: {os.path.basename(abs_input)}\nDetail: {exc}"
            )

        for i, term in enumerate(sorted_words, start=1):
            try:
                word.Selection.HomeKey(Unit=WD_STORY_MAIN)  # kembali ke awal dokumen

                find = word.Selection.Find
                find.ClearFormatting()
                find.Replacement.ClearFormatting()

                find.Text = term
                find.Replacement.Text = term
                find.Forward = True
                find.Wrap = WD_FIND_CONTINUE
                find.Format = True
                find.MatchCase = bool(case_sensitive)
                find.MatchWholeWord = bool(whole_word)
                find.MatchWildcards = False
                find.MatchSoundsLike = False
                find.MatchAllWordForms = False

                # Terapkan format pada hasil replacement sesuai pilihan user
                if format_mode == "italic":
                    find.Replacement.Font.Italic = True
                elif format_mode == "bold":
                    find.Replacement.Font.Bold = True
                elif format_mode == "underline":
                    find.Replacement.Font.Underline = WD_UNDERLINE_SINGLE
                elif format_mode == "highlight":
                    find.Replacement.Highlight = True
                elif format_mode == "color":
                    find.Replacement.Font.Color = _hex_to_word_color(color_hex)

                if find.Execute(Replace=WD_REPLACE_ALL):
                    replaced_terms += 1
            except Exception as exc:
                # Satu kata gagal jangan hentikan keseluruhan proses;
                # catat lalu lanjut ke istilah berikutnya.
                progress_cb(i, total_terms,
                            f"⚠️ Gagal memproses '{term}': {exc}")
                continue

            progress_cb(i, total_terms,
                        f"[MS Word] Istilah {i}/{total_terms}: {term}")

        try:
            doc.SaveAs2(abs_output, FileFormat=WD_FORMAT_XML_DOCUMENT)
        except Exception as exc:
            raise PermissionError(
                "Gagal menyimpan hasil lewat MS Word.\n"
                "Tutup dokumen tujuan bila sedang terbuka, lalu coba lagi.\n\n"
                f"File: {os.path.basename(abs_output)}\nDetail: {exc}"
            )

    finally:
        # CLEANUP AMAN: pastikan proses WINWORD.EXE tidak jadi zombie.
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    return output_path, replaced_terms, total_terms


# ============================================================================
# BAGIAN GUI : ANTARMUKA CUSTOMTKINTER (DARK MODE)
# ============================================================================

# Base window mendukung Drag & Drop bila tkinterdnd2 tersedia,
# jika tidak, pakai CTk biasa (fitur DnD dimatikan secara anggun).
BaseWindow = TkinterDnD.Tk if DND_AVAILABLE else ctk.CTk


class ItalikinApp(BaseWindow):

    def __init__(self):
        super().__init__()

        ctk.set_appearance_mode("Dark")          # Dark Mode default
        ctk.set_default_color_theme("blue")

        self.title(f"⭐ {APP_NAME} v{APP_VERSION} — Format Kata Otomatis (.docx)")
        self.geometry("800x780")
        self.minsize(720, 700)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # File management
        self.original_path = ""        # Path file asli (tidak disentuh)
        self.working_copy_path = ""     # Path copy di working directory
        self.working_dir = ""           # Working directory untuk proses
        self.output_path = ""           # Path output terakhir
        self.is_processing = False
        self.recent_files = []           # File terakhir dibuka
        self.settings = {}              # Settings tersimpan

        # Buat working directory
        self._setup_working_directory()
        
        self._build_ui()
        self._register_drag_and_drop()
        self._setup_keyboard_shortcuts()
        
        # Load settings saat startup
        self.after(100, self._load_settings)

    # ------------------------------------------------------- SETTINGS --
    def _save_settings(self):
        """Simpan settings ke file."""
        settings = {
            "whole_word": bool(self.sw_whole_word.get()),
            "case_sensitive": bool(self.sw_case_sensitive.get()),
            "format": self.seg_format.get(),
            "color": self.menu_color.get(),
            "engine": self.menu_engine.get(),
            "recent_files": self.recent_files,
        }
        try:
            settings_path = os.path.join(self.working_dir, "settings.json")
            import json
            with open(settings_path, "w") as f:
                json.dump(settings, f, indent=2)
            self.log("💾 Settings tersimpan.")
        except Exception as e:
            self.log(f"⚠️ Gagal simpan settings: {e}")
    
    def _load_settings(self):
        """Muat settings dari file."""
        try:
            settings_path = os.path.join(self.working_dir, "settings.json")
            if not os.path.exists(settings_path):
                return
            import json
            with open(settings_path, "r") as f:
                settings = json.load(f)
            
            # Apply settings
            if settings.get("whole_word"):
                self.sw_whole_word.select()
            else:
                self.sw_whole_word.deselect()
            
            if settings.get("case_sensitive"):
                self.sw_case_sensitive.select()
            else:
                self.sw_case_sensitive.deselect()
            
            if settings.get("format"):
                self.seg_format.set(settings["format"])
                self._on_format_change(settings["format"])
            
            if settings.get("color"):
                self.menu_color.set(settings["color"])
            
            if settings.get("engine"):
                self.menu_engine.set(settings["engine"])
            
            if settings.get("recent_files"):
                self.recent_files = settings["recent_files"]
            
            self.log("📂 Settings dimuat.")
        except Exception as e:
            self.log(f"⚠️ Gagal muat settings: {e}")
    
    # ---------------------------------------------------- KEYBOARD SHORTCUTS --
    def _setup_keyboard_shortcuts(self):
        """Setup keyboard shortcuts untuk aplikasi."""
        # Ctrl+O: Buka dokumen
        self.bind('<Control-o>', lambda e: self.browse_docx())
        self.bind('<Control-O>', lambda e: self.browse_docx())
        
        # Ctrl+Shift+O: Muat daftar kata
        self.bind('<Control-Shift-O>', lambda e: self.browse_wordlist())
        self.bind('<Control-Shift-o>', lambda e: self.browse_wordlist())
        
        # Ctrl+Enter: Proses dokumen
        self.bind('<Control-Return>', lambda e: self.start_processing())
        
        # Ctrl+E: Export
        self.bind('<Control-e>', lambda e: self._export_output(self.output_path) if self.output_path else None)
        self.bind('<Control-E>', lambda e: self._export_output(self.output_path) if self.output_path else None)
        
        # Ctrl+L: Clear log
        self.bind('<Control-l>', lambda e: self._clear_log())
        self.bind('<Control-L>', lambda e: self._clear_log())
        
        # Ctrl+S: Save settings
        self.bind('<Control-s>', lambda e: self._save_settings())
        self.bind('<Control-S>', lambda e: self._save_settings())
        
        # Ctrl+Shift+S: Load settings
        self.bind('<Control-Shift-S>', lambda e: self._load_settings())
        
        # F1: Help
        self.bind('<F1>', lambda e: self._show_help())
        
        # Escape: Cancel/Tutup dialog
        self.bind('<Escape>', lambda e: None)  # Placeholder

    # ------------------------------------------------------ FILE MANAGEMENT --
    def _setup_working_directory(self):
        """Setup working directory untuk file copy."""
        # Gunakan folder 'italikin_workspace' di samping executable/script
        if getattr(sys, 'frozen', False):
            # Running as .exe
            base_dir = os.path.dirname(sys.executable)
        else:
            # Running as script
            base_dir = os.path.dirname(os.path.abspath(__file__))
        
        self.working_dir = os.path.join(base_dir, "italikin_workspace")
        os.makedirs(self.working_dir, exist_ok=True)
        
        # Cleanup workspace lama saat startup
        self._cleanup_workspace()

    def _cleanup_workspace(self):
        """Hapus file copy lama di working directory."""
        try:
            for file in Path(self.working_dir).glob("*.docx"):
                file.unlink(missing_ok=True)
        except Exception:
            pass

    def _copy_file_to_workspace(self, source_path: str) -> str:
        """
        Copy file .docx ke working directory.
        Original TIDAK akan disentuh.
        
        Returns: path ke file copy di workspace
        """
        if not os.path.exists(source_path):
            raise FileNotFoundError(f"File tidak ditemukan: {source_path}")
        
        # Buat nama unuk untuk copy
        filename = os.path.basename(source_path)
        name, ext = os.path.splitext(filename)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        copy_name = f"{name}_{timestamp}{ext}"
        copy_path = os.path.join(self.working_dir, copy_name)
        
        # Copy file
        shutil.copy2(source_path, copy_path)
        
        self.log(f"📋 File di-copy ke workspace:")
        self.log(f"   Original: {source_path}")
        self.log(f"   Copy: {copy_path}")
        
        return copy_path

    def _show_export_dialog(self, output_path: str) -> str:
        """
        Tampilkan dialog untuk pilih lokasi & nama file output.
        
        Returns: path output pilihan user, atau None jika dibatalkan
        """
        # Dapatkan nama default
        default_name = os.path.basename(output_path)
        default_dir = os.path.dirname(output_path)
        
        # Tampilkan Save As dialog
        save_path = filedialog.asksaveasfilename(
            title="Simpan Hasil Formatting",
            initialdir=default_dir,
            initialfile=default_name,
            defaultextension=".docx",
            filetypes=[
                ("Dokumen Word", "*.docx"),
                ("Semua File", "*.*")
            ]
        )
        
        return save_path if save_path else None

    # ---------------------------------------------------------- MENU BAR --
    def _build_menu_bar(self):
        """Membangun menu bar untuk aplikasi."""
        self.menu_bar = ctk.CTkMenuBar(self)

        # Menu File
        file_menu = self.menu_bar.add_cascade("📁 File")
        file_menu.add_command(label="📂 Buka Dokumen", command=self.browse_docx, accelerator="Ctrl+O")
        file_menu.add_command(label="📋 Muat Daftar Kata", command=self.browse_wordlist, accelerator="Ctrl+Shift+O")
        file_menu.add_separator()
        
        # Recent files submenu
        self.recent_menu = file_menu.add_cascade("🕐 File Terakhir")
        self._update_recent_menu()
        
        file_menu.add_separator()
        file_menu.add_command(label="❌ Keluar", command=self._on_close, accelerator="Alt+F4")

        # Menu Tools
        tools_menu = self.menu_bar.add_cascade("🛠️ Tools")
        tools_menu.add_command(label="🤖 AI Assistant", command=self._show_ai_assistant)
        tools_menu.add_command(label="🕐 File Terakhir", command=self._show_recent_files)
        tools_menu.add_separator()
        tools_menu.add_command(label="💾 Simpan Settings", command=self._save_settings, accelerator="Ctrl+S")
        tools_menu.add_command(label="📂 Muat Settings", command=self._load_settings, accelerator="Ctrl+Shift+S")
        tools_menu.add_separator()
        tools_menu.add_command(label="🗑️ Bersihkan Log", command=self._clear_log, accelerator="Ctrl+L")
        tools_menu.add_command(label="📁 Buka Folder Working", command=self._open_output_folder)

        # Menu Help
        help_menu = self.menu_bar.add_cascade("❓ Bantuan")
        help_menu.add_command(label="📖 Panduan Penggunaan", command=self._show_help)
        help_menu.add_command(label="ℹ️ Tentang Italikin", command=self._show_about)

    def _clear_log(self):
        """Bersihkan area log."""
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.configure(state="disabled")
        self.log("📋 Log dibersihkan.")
    
    def _add_recent_file(self, filepath: str):
        """Tambah file ke recent files."""
        if filepath in self.recent_files:
            self.recent_files.remove(filepath)
        self.recent_files.insert(0, filepath)
        self.recent_files = self.recent_files[:5]  # Max 5 recent
        self._update_recent_menu()
    
    def _update_recent_menu(self):
        """Update menu recent files."""
        # Hapus placeholder jika ada (akan di-rebuild)
        pass  # Menu akan di-update saat dibuka
    
    def _show_recent_files(self):
        """Tampilkan dialog recent files."""
        if not self.recent_files:
            messagebox.showinfo("Info", "Belum ada file terakhir.")
            return
        
        dialog = ctk.CTkToplevel(self)
        dialog.title("File Terakhir")
        dialog.geometry("400x300")
        dialog.transient(self)
        dialog.grab_set()
        
        ctk.CTkLabel(
            dialog, text="🕐 File Terakhir Dibuka",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(pady=(15, 10))
        
        for i, filepath in enumerate(self.recent_files):
            frame = ctk.CTkFrame(dialog)
            frame.pack(fill="x", padx=15, pady=3)
            
            ctk.CTkLabel(
                frame, text=os.path.basename(filepath),
                font=ctk.CTkFont(size=12),
                anchor="w"
            ).pack(side="left", padx=10, pady=8)
            
            def open_file(path=filepath):
                self._load_document(path)
                dialog.destroy()
            
            ctk.CTkButton(
                frame, text="📂 Buka", width=70, height=28,
                command=open_file
            ).pack(side="right", padx=10, pady=5)
        
        ctk.CTkButton(
            dialog, text="✕ Tutup", width=80,
            command=dialog.destroy
        ).pack(pady=(10, 15))

    def _open_output_folder(self):
        """Buka folder working directory."""
        if os.path.exists(self.working_dir):
            try:
                os.startfile(self.working_dir)  # Windows
            except AttributeError:
                import subprocess
                opener = "open" if sys.platform == "darwin" else "xdg-open"
                subprocess.Popen([opener, self.working_dir])
        else:
            messagebox.showinfo("Info", "Working directory belum dibuat.")

    def _show_help(self):
        """Tampilkan dialog bantuan."""
        help_text = f"""📖 PANDUAN PENGGUNAAN {APP_NAME}
{'='*50}

🎯 TUJUAN:
Memformat kata tertentu di dokumen Word (.docx)
secara otomatis tanpa merusak layout.

📝 CARA PENGGUNAAN:
1. Klik 'Browse' atau seret file .docx ke jendela
2. Ketik kata target (dipisah koma) atau muat dari .txt
3. Atur pengaturan sesuai kebutuhan
4. Klik 'PROSES DOKUMEN'
5. File output tersimpan dengan suffix '_italik'

🤖 WORKFLOW DENGAN AI:
1. Upload dokumen ke ChatGPT/Gemini/Claude
2. Minta AI buat list kata penting (comma-separated)
3. Copy paste list ke kolom Kata Target
4. Proses dokumen!

⚙️ PENGATURAN:
• Kata utuh saja: Hanya match kata lengkap
• Case-sensitive: Huruf besar/kecil harus persis
• Format: Italic/Bold/Underline/Highlight/Warna
• Engine: MS Word COM (recommended)

📁 TENTANG FILE:
• Italikin HANYA menyimpan PATH file, bukan copy file
• File asli tidak akan dipindah/dihapus
• Output disimpan di folder yang sama dengan file asli
• Nama output: namafile_italik.docx

💡 TIPS:
• Tutup MS Word sebelum memproses
• Gunakan .txt untuk banyak kata target
• Cek log untuk detail proses
• Gunakan AI Assistant (menu Tools) untuk bantuan

📧 Bantuan: {APP_URL}
"""
        messagebox.showinfo("Panduan Penggunaan", help_text)

    def _show_about(self):
        """Tampilkan dialog About."""
        about_text = f"""⭐ {APP_NAME} v{APP_VERSION}
{'='*40}

{APP_DESCRIPTION}

👤 Author: {APP_AUTHOR}
🌐 Website: {APP_URL}

📦 Fitur:
• Format kata target otomatis
• MS Word COM engine (recommended)
• Drag & Drop support
• Dark Mode GUI
• AI Assistant untuk generate kata target

🔧 Dependencies:
• customtkinter
• python-docx
• pywin32 (MS Word COM engine)
• tkinterdnd2 (optional, untuk Drag & Drop)

📄 Lisensi: MIT License

Made with ❤️ by {APP_AUTHOR}
"""
        messagebox.showinfo("Tentang Italikin", about_text)

    def _show_ai_assistant(self):
        """Tampilkan dialog AI Assistant untuk guide user."""
        ai_text = f"""🤖 AI ASSISTANT — Cara Generate Kata Target
{'='*55}

📋 WORKFLOW:
1. Upload dokumen Word/PDF ke AI (ChatGPT, Gemini, Claude, dll)
2. Minta AI membuat list kata yang perlu diformat
3. Copy paste list ke Italikin
4. Proses dokumen!

{'─'*55}

💡 CONTOH PROMPT UNTUK AI:

"Tolong buatkan daftar istilah hukum penting dari dokumen ini.
Format: kata1, kata2, kata3 (dipisah koma)"

ATAU:

"Extract key terms from this document that should be
highlighted/italicized. Return as comma-separated list."

{'─'*55}

📝 CONTOH OUTPUT DARI AI:

PT Maju Jaya, Perjanjian Kerja, Klausul Penalti,
Force Majeure, Ganti Rugi, Wanprestasi

{'─'*55}

⚙️ SETELAH COPY PASTE:
1. Paste list ke kolom 'Kata Target' di Italikin
2. Atau simpan sebagai file .txt lalu muat
3. Pilih format (Italic/Bold/Highlight)
4. Klik PROSES DOKUMEN

{'─'*55}

💡 TIPS:
• Minta AI format sebagai comma-separated list
• Bisa request case-sensitive atau whole word
• Semakin spesifik prompt, semakin akurat hasilnya
• Review hasil AI sebelum diproses

🌐 AI YANG BISA DIPAKAI:
• ChatGPT (chat.openai.com)
• Gemini (gemini.google.com)
• Claude (claude.ai)
• Copilot (copilot.microsoft.com)
"""
        # Buat dialog custom dengan scrollable text
        dialog = ctk.CTkToplevel(self)
        dialog.title("🤖 AI Assistant")
        dialog.geometry("600x700")
        dialog.transient(self)
        dialog.grab_set()

        # Header
        header = ctk.CTkFrame(dialog, fg_color=("gray85", "gray17"))
        header.pack(fill="x", padx=10, pady=10)

        ctk.CTkLabel(
            header, text="🤖 AI ASSISTANT",
            font=ctk.CTkFont(size=20, weight="bold")
        ).pack(pady=(10, 5))

        ctk.CTkLabel(
            header,
            text="Generate list kata target dengan bantuan AI",
            font=ctk.CTkFont(size=12),
            text_color=("gray40", "gray60"),
        ).pack(pady=(0, 10))

        # Scrollable text area
        text_frame = ctk.CTkFrame(dialog)
        text_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        text_box = ctk.CTkTextbox(
            text_frame, wrap="word",
            font=ctk.CTkFont(size=12, family="Consolas"),
            state="normal"
        )
        text_box.pack(fill="both", expand=True, padx=5, pady=5)
        text_box.insert("1.0", ai_text)
        text_box.configure(state="disabled")  # Read-only

        # Buttons
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=(0, 10))

        def copy_template():
            template = "PT Maju Jaya, Perjanjian Kerja, Klausul Penalti, Force Majeure, Ganti Rugi"
            dialog.clipboard_clear()
            dialog.clipboard_append(template)
            messagebox.showinfo("Tersalin!", "Template sudah disalin ke clipboard!\nPaste ke kolom Kata Target.")

        def open_chatgpt():
            import webbrowser
            webbrowser.open("https://chat.openai.com")

        def open_gemini():
            import webbrowser
            webbrowser.open("https://gemini.google.com")

        ctk.CTkButton(
            btn_frame, text="📋 Copy Template Kata",
            command=copy_template, width=150,
            fg_color="#28a745", hover_color="#218838"
        ).pack(side="left", padx=(0, 5))

        ctk.CTkButton(
            btn_frame, text="🌐 Buka ChatGPT",
            command=open_chatgpt, width=120,
            fg_color="#10a37f", hover_color="#0d8c6c"
        ).pack(side="left", padx=(0, 5))

        ctk.CTkButton(
            btn_frame, text="🌐 Buka Gemini",
            command=open_gemini, width=120,
            fg_color="#4285f4", hover_color="#3367d6"
        ).pack(side="left", padx=(0, 5))

        ctk.CTkButton(
            btn_frame, text="✕ Tutup",
            command=dialog.destroy, width=80,
            fg_color="#dc3545", hover_color="#c82333"
        ).pack(side="right")

    # ------------------------------------------------------------------ UI --
    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(5, weight=1)

        # ================= MENU BAR =================
        self._build_menu_bar()

        # ================= HEADER =================
        header = ctk.CTkFrame(self, corner_radius=12, fg_color=("gray85", "gray17"))
        header.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 8))
        header.grid_columnconfigure(0, weight=1)
        header.grid_columnconfigure(1, weight=0)

        # Left side: Title (compact)
        title_frame = ctk.CTkFrame(header, fg_color="transparent")
        title_frame.grid(row=0, column=0, rowspan=2, sticky="w", padx=14, pady=(12, 12))

        ctk.CTkLabel(
            title_frame, text="⭐ ITALIKIN",
            font=ctk.CTkFont(size=24, weight="bold")
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            title_frame,
            text="Format kata di dokumen Word tanpa merusak layout",
            font=ctk.CTkFont(size=12),
            text_color=("gray40", "gray60"),
        ).grid(row=1, column=0, sticky="w")

        # Right side: Version info
        version_frame = ctk.CTkFrame(header, fg_color="transparent")
        version_frame.grid(row=0, column=1, rowspan=2, sticky="e", padx=14, pady=(12, 12))

        ctk.CTkLabel(
            version_frame, text=f"v{APP_VERSION}",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=("gray50", "gray50"),
        ).grid(row=0, column=0, sticky="e")

        # ========== PANEL 1 : INPUT FILE DOKUMEN ==========
        frame_doc = ctk.CTkFrame(self, corner_radius=10)
        frame_doc.grid(row=1, column=0, sticky="ew", padx=16, pady=6)
        frame_doc.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            frame_doc, text="📄  File Dokumen Word (.docx)",
            font=ctk.CTkFont(size=14, weight="bold"), anchor="w",
        ).grid(row=0, column=0, columnspan=2, sticky="ew", padx=14, pady=(10, 4))

        self.entry_docx = ctk.CTkEntry(
            frame_doc, placeholder_text="Pilih / seret file .docx ke sini …",
            font=ctk.CTkFont(size=13), height=36,
        )
        self.entry_docx.grid(row=1, column=0, sticky="ew", padx=(14, 8), pady=(0, 12))

        ctk.CTkButton(
            frame_doc, text="📁 Pilih Dokumen", width=120, height=36,
            command=self.browse_docx,
        ).grid(row=1, column=1, sticky="e", padx=(0, 14), pady=(0, 12))

        # ========== PANEL 2 : INPUT KATA TARGET ==========
        frame_words = ctk.CTkFrame(self, corner_radius=10)
        frame_words.grid(row=2, column=0, sticky="ew", padx=16, pady=6)
        frame_words.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            frame_words,
            text="🔤  Kata Target  (ketik dipisah koma, ATAU muat dari file .txt)",
            font=ctk.CTkFont(size=14, weight="bold"), anchor="w",
        ).grid(row=0, column=0, columnspan=2, sticky="ew", padx=14, pady=(10, 4))

        self.entry_words = ctk.CTkEntry(
            frame_words,
            placeholder_text="contoh: penting, PT Maju Jaya, Sumber Daya Manusia",
            font=ctk.CTkFont(size=13), height=36,
        )
        self.entry_words.grid(row=1, column=0, sticky="ew", padx=(14, 8), pady=(0, 4))
        
        # Word count preview
        self.lbl_word_count = ctk.CTkLabel(
            frame_words,
            text="",
            font=ctk.CTkFont(size=11),
            text_color=("gray45", "gray55"),
            anchor="w"
        )
        self.lbl_word_count.grid(row=2, column=0, columnspan=2, sticky="w", padx=14, pady=(0, 8))
        
        # Bind event untuk update word count
        self.entry_words.bind('<KeyRelease>', self._update_word_count)

        ctk.CTkButton(
            frame_words, text="📋 Pilih Kata", width=120, height=36,
            command=self.browse_wordlist,
        ).grid(row=1, column=1, sticky="e", padx=(0, 14), pady=(0, 4))

        # ========== PANEL 3 : OPSI PENCARIAN & FORMAT OUTPUT ==========
        frame_opt = ctk.CTkFrame(self, corner_radius=10)
        frame_opt.grid(row=3, column=0, sticky="ew", padx=16, pady=6)
        frame_opt.grid_columnconfigure(0, weight=1)
        frame_opt.grid_columnconfigure(1, weight=0)

        ctk.CTkLabel(
            frame_opt, text="⚙️  Pengaturan Proses",
            font=ctk.CTkFont(size=14, weight="bold"), anchor="w",
        ).grid(row=0, column=0, columnspan=2, sticky="ew", padx=14, pady=(10, 6))

        # --- Kolom kiri: opsi presisi pencarian ---
        frame_precision = ctk.CTkFrame(frame_opt, fg_color="transparent")
        frame_precision.grid(row=1, column=0, sticky="w", padx=(14, 8), pady=(0, 12))

        self.sw_whole_word = ctk.CTkSwitch(
            frame_precision, text="Kata utuh saja",
            font=ctk.CTkFont(size=13),
        )
        self.sw_whole_word.grid(row=0, column=0, sticky="w", padx=(0, 18))
        self.sw_whole_word.select()  # default ON agar pencarian lebih presisi

        self.sw_case_sensitive = ctk.CTkSwitch(
            frame_precision, text="Case-sensitive",
            font=ctk.CTkFont(size=13),
        )
        self.sw_case_sensitive.grid(row=0, column=1, sticky="w")

        # --- Kolom kanan: pilihan format output ---
        frame_format = ctk.CTkFrame(frame_opt, fg_color="transparent")
        frame_format.grid(row=1, column=1, sticky="e", padx=(8, 14), pady=(0, 12))

        ctk.CTkLabel(
            frame_format, text="🎨 Format:",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).pack(side="left", padx=(0, 8))

        self.seg_format = ctk.CTkSegmentedButton(
            frame_format, values=FORMAT_MODES,
            font=ctk.CTkFont(size=12),
            command=self._on_format_change,
        )
        self.seg_format.set("Italic")
        self.seg_format.pack(side="left")

        self.menu_color = ctk.CTkOptionMenu(
            frame_format, values=list(COLOR_CHOICES.keys()),
            width=100, font=ctk.CTkFont(size=12), state="disabled",
        )
        self.menu_color.set("Merah")
        self.menu_color.pack(side="left", padx=(8, 0))

        # --- Baris kedua: pilihan engine pemroses ---
        frame_engine = ctk.CTkFrame(frame_opt, fg_color="transparent")
        frame_engine.grid(row=2, column=0, columnspan=2, sticky="ew",
                          padx=14, pady=(0, 12))

        ctk.CTkLabel(
            frame_engine, text="🛠️  Engine:",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).pack(side="left", padx=(0, 8))

        engine_values = [ENGINE_COM]
        if COM_AVAILABLE:
            engine_values.append(ENGINE_DOCX)
        else:
            engine_values.append("python-docx (install pywin32 untuk COM)")

        self.menu_engine = ctk.CTkOptionMenu(
            frame_engine, values=engine_values,
            font=ctk.CTkFont(size=12), width=260,
        )
        self.menu_engine.set(ENGINE_COM)
        self.menu_engine.pack(side="left")

        if not COM_AVAILABLE:
            self.menu_engine.configure(state="disabled")

        engine_note = ("✅ Direkomendasikan — akses MS Word asli, hasil lebih rapi"
                       if COM_AVAILABLE else
                       "⚠️ Install 'pywin32' + MS Word untuk engine terbaik")
        ctk.CTkLabel(
            frame_engine, text=f"ℹ️ {engine_note}",
            font=ctk.CTkFont(size=11), text_color=("gray45", "gray55"),
        ).pack(side="left", padx=(10, 0))

        # ========== TOMBOL PROSES BESAR ==========
        self.btn_process = ctk.CTkButton(
            self, text="✨  PROSES DOKUMEN  ✨",
            font=ctk.CTkFont(size=18, weight="bold"),
            height=56, corner_radius=12,
            fg_color="#1F6AA5", hover_color="#144870",
            command=self.start_processing,
        )
        self.btn_process.grid(row=4, column=0, sticky="ew", padx=16, pady=10)

        # ========== PANEL LOG / HASIL ==========
        frame_log = ctk.CTkFrame(self, corner_radius=10)
        frame_log.grid(row=5, column=0, sticky="nsew", padx=16, pady=(6, 4))
        frame_log.grid_rowconfigure(1, weight=1)
        frame_log.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            frame_log, text="📋  Log Proses",
            font=ctk.CTkFont(size=14, weight="bold"), anchor="w",
        ).grid(row=0, column=0, sticky="ew", padx=14, pady=(10, 2))

        self.txt_log = ctk.CTkTextbox(frame_log, wrap="word", font=ctk.CTkFont(size=13))
        self.txt_log.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 6))
        self.txt_log.configure(state="disabled")

        # ========== STATUS BAR + PROGRESS BAR ==========
        frame_status = ctk.CTkFrame(self, fg_color="transparent")
        frame_status.grid(row=6, column=0, sticky="ew", padx=16, pady=(2, 14))
        frame_status.grid_columnconfigure(0, weight=1)

        self.progress = ctk.CTkProgressBar(frame_status, height=10)
        self.progress.set(0)
        self.progress.grid(row=0, column=0, sticky="ew", pady=(0, 4))

        self.lbl_status = ctk.CTkLabel(
            frame_status, text="✔ Siap. Silakan pilih dokumen & masukkan kata target.",
            font=ctk.CTkFont(size=12), anchor="w", text_color=("gray35", "gray65"),
        )
        self.lbl_status.grid(row=1, column=0, sticky="ew")

    # -------------------------------------------------- DRAG & DROP (opsi) --
    def _register_drag_and_drop(self):
        if not DND_AVAILABLE:
            return
        try:
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self._on_drop)
        except Exception as exc:
            # Jangan sampai DnD yang gagal membuat aplikasi mati total
            print(f"[Italikin] Drag & drop tidak dapat diaktifkan: {exc}")

    @staticmethod
    def _parse_dropped_paths(data: str) -> list:
        """Parsing string path dari event drop (mendukung path ber-spasi)."""
        raw_tokens = re.findall(r"\{[^}]*\}|\S+", data)
        return [token.strip("{}") for token in raw_tokens]

    def _on_drop(self, event):
        if self.is_processing:
            return
        for path in self._parse_dropped_paths(event.data):
            lower = path.lower()
            if lower.endswith(".docx"):
                self._load_document(path)
            elif lower.endswith(".txt"):
                self._load_wordlist_file(path)
            else:
                self.log(f"⚠️ File dilewati (bukan .docx/.txt): {os.path.basename(path)}")

    # ----------------------------------------------------------- ACTIONS --
    def log(self, message: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", f"[{timestamp}] {message}\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def set_status(self, message: str, progress_value=None):
        self.lbl_status.configure(text=message)
        if progress_value is not None:
            self.progress.set(progress_value)

    def browse_docx(self):
        try:
            path = filedialog.askopenfilename(
                title="Pilih Dokumen Word",
                filetypes=[("Dokumen Word", "*.docx"), ("Semua File", "*.*")],
            )
            if path:
                self._load_document(path)
        except Exception as exc:
            messagebox.showerror("Error", f"Gagal membuka dialog file:\n{exc}")

    def _load_document(self, path: str):
        """Load dokumen dan copy ke workspace untuk keamanan."""
        try:
            # Simpan path original
            self.original_path = path
            
            # Copy ke workspace (original tetap aman)
            self.working_copy_path = self._copy_file_to_workspace(path)
            
            # Tambah ke recent files
            self._add_recent_file(path)
            
            # Update UI
            self.entry_docx.delete(0, "end")
            self.entry_docx.insert(0, path)
            
            self.log(f"📄 Dokumen dipilih: {path}")
            self.log(f"🔒 Original file TIDAK akan diubah")
            self.set_status("✔ Dokumen siap diproses (copy di workspace).", 0)
        except Exception as exc:
            messagebox.showerror("Error", f"Gagal memuat dokumen:\n{exc}")

    def browse_wordlist(self):
        try:
            path = filedialog.askopenfilename(
                title="Pilih File Daftar Kata (.txt)",
                filetypes=[("Text File", "*.txt"), ("Semua File", "*.*")],
            )
            if path:
                self._load_wordlist_file(path)
        except Exception as exc:
            messagebox.showerror("Error", f"Gagal membaca file .txt:\n{exc}")

    def _load_wordlist_file(self, txt_path: str):
        """Loader bersama untuk tombol 'Muat .txt' dan drag & drop file .txt."""
        try:
            words = load_words_from_txt(txt_path)
            if not words:
                messagebox.showwarning(
                    "Daftar Kosong",
                    "File .txt tidak berisi kata yang valid.\n"
                    "Pastikan kata dipisahkan koma atau baris baru."
                )
                return
            self.entry_words.delete(0, "end")
            self.entry_words.insert(0, ", ".join(words))
            self.log(f"📂 {len(words)} kata dimuat dari: {txt_path}")
            self.set_status(f"✔ {len(words)} kata target dimuat dari file.", 0)
        except Exception as exc:
            messagebox.showerror("Error", f"Gagal membaca file .txt:\n{exc}")

    def _on_format_change(self, value: str):
        """Aktifkan pemilih warna hanya saat format 'Warna' dipilih."""
        if value == "Warna":
            self.menu_color.configure(state="normal")
        else:
            self.menu_color.configure(state="disabled")
    
    def _update_word_count(self, event=None):
        """Update word count preview saat user mengetik."""
        raw = self.entry_words.get().strip()
        if not raw:
            self.lbl_word_count.configure(text="")
            return
        words = parse_words(raw)
        if words:
            preview = ", ".join(words[:5])
            if len(words) > 5:
                preview += f"... (+{len(words)-5} lagi)"
            self.lbl_word_count.configure(
                text=f"📊 {len(words)} kata target: {preview}"
            )
        else:
            self.lbl_word_count.configure(text="")

    def _resolve_format_settings(self) -> tuple:
        """Baca UI dan kembalikan tuple (format_mode, color_hex)."""
        selected = self.seg_format.get()
        mapping = {
            "Italic": ("italic", None),
            "Bold": ("bold", None),
            "Underline": ("underline", None),
            "Highlight": ("highlight", None),
            "Warna": ("color", COLOR_CHOICES[self.menu_color.get()]),
        }
        return mapping.get(selected, ("italic", None))

    def _resolve_engine(self) -> str:
        """Baca pilihan engine dari UI. COM adalah default jika tersedia."""
        selected = self.menu_engine.get()
        if COM_AVAILABLE and selected == ENGINE_COM:
            return "com"
        # Jika COM tidak tersedia atau user pilih docx, gunakan docx
        return "docx"

    def start_processing(self):
        if self.is_processing:
            return

        # ---------- Validasi input ----------
        if not self.working_copy_path or not os.path.isfile(self.working_copy_path):
            messagebox.showwarning("Belum Lengkap", "Silakan pilih file dokumen .docx terlebih dahulu.")
            return

        words = parse_words(self.entry_words.get())
        if not words:
            messagebox.showwarning(
                "Kata Target Kosong",
                "Masukkan minimal satu kata target\n"
                "(ketik dipisah koma ATAU muat dari file .txt)."
            )
            return

        format_mode, color_hex = self._resolve_format_settings()
        whole_word = bool(self.sw_whole_word.get())
        case_sensitive = bool(self.sw_case_sensitive.get())
        engine = self._resolve_engine()

        # ---------- Mulai proses di thread terpisah ----------
        self.is_processing = True
        self.btn_process.configure(state="disabled", text="⏳  SEDANG MEMPROSES…")
        self.set_status("⏳ Memproses dokumen…", 0)
        self.log("=" * 52)
        self.log(f"🚀 Memulai proses untuk: {os.path.basename(self.original_path)}")
        self.log(f"📋 Memproses COPY (original tetap aman): {os.path.basename(self.working_copy_path)}")
        self.log(f"🎯 Kata target ({len(words)}): {', '.join(words[:15])}"
                 + ("…" if len(words) > 15 else ""))
        self.log(f"⚙️ Format={format_mode}"
                 + (f" ({self.menu_color.get()})" if format_mode == "color" else "")
                 + f" | KataUtuh={'YA' if whole_word else 'TIDAK'}"
                 + f" | CaseSensitive={'YA' if case_sensitive else 'TIDAK'}")
        self.log(f"🛠️ Engine: {'MS Word COM' if engine == 'com' else 'python-docx'}")

        worker = threading.Thread(
            target=self._processing_worker,
            args=(self.working_copy_path, words, format_mode, color_hex,
                  whole_word, case_sensitive, engine),
            daemon=True,
        )
        worker.start()

    def _processing_worker(self, docx_path: str, words: list,
                           format_mode: str, color_hex: str,
                           whole_word: bool, case_sensitive: bool,
                           engine: str):
        """Jalankan core processing di background thread agar GUI tidak freeze."""
        try:
            def progress_cb(current, total, message):
                self.after(0, lambda: self.set_status(message, current / max(total, 1)))

            if engine == "com":
                output_path, changed, total_paras = process_document_com(
                    docx_path, words,
                    format_mode=format_mode, color_hex=color_hex,
                    whole_word=whole_word, case_sensitive=case_sensitive,
                    progress_cb=progress_cb,
                )
            else:
                output_path, changed, total_paras = process_document(
                    docx_path, words,
                    format_mode=format_mode, color_hex=color_hex,
                    whole_word=whole_word, case_sensitive=case_sensitive,
                    progress_cb=progress_cb,
                )
            self.after(0, lambda: self._on_success(output_path, changed, total_paras))
        except PermissionError as exc:
            self.after(0, lambda err=exc: self._on_file_locked(err))
        except Exception as exc:
            self.after(0, lambda err=exc: self._on_failure(exc))

    def _on_success(self, output_path: str, changed: int, total_paras: int):
        self.is_processing = False
        self.btn_process.configure(state="normal", text="✨  PROSES DOKUMEN  ✨")
        self.progress.set(1.0)
        self.output_path = output_path  # Simpan untuk export
        self.set_status(f"✅ Selesai! {changed} kata berhasil diformat.", 1.0)
        self.log(f"✅ SELESAI — {changed} kata berhasil diformat "
                 f"dari {total_paras} paragraf yang diperiksa.")
        self.log(f"💾 Hasil: {output_path}")
        self.log("=" * 52)
        
        # Tampilkan panel export (bukan dialog)
        self._show_export_panel(output_path, changed)

    def _show_export_panel(self, output_path: str, changed: int):
        """Tampilkan panel export integrated (bukan dialog)."""
        # Buat top-level dialog yang simple
        dialog = ctk.CTkToplevel(self)
        dialog.title("Export Hasil")
        dialog.geometry("500x280")
        dialog.transient(self)
        dialog.grab_set()
        dialog.resizable(False, False)
        
        # Center dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (500 // 2)
        y = (dialog.winfo_screenheight() // 2) - (280 // 2)
        dialog.geometry(f"500x280+{x}+{y}")
        
        # Main frame
        main_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Success header
        ctk.CTkLabel(
            main_frame, text="✅ Proses Berhasil!",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="#28a745"
        ).pack(pady=(0, 10))
        
        # Stats
        ctk.CTkLabel(
            main_frame,
            text=f"📊 {changed} kata berhasil diformat",
            font=ctk.CTkFont(size=14)
        ).pack(pady=(0, 5))
        
        # Output path info
        ctk.CTkLabel(
            main_frame,
            text=f"📄 Output: {os.path.basename(output_path)}",
            font=ctk.CTkFont(size=12),
            text_color="gray60"
        ).pack(pady=(0, 15))
        
        # Export section
        export_frame = ctk.CTkFrame(main_frame)
        export_frame.pack(fill="x", pady=(0, 15))
        
        ctk.CTkLabel(
            export_frame, text="📤 Export ke:",
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(padx=10, pady=(10, 5), anchor="w")
        
        path_frame = ctk.CTkFrame(export_frame, fg_color="transparent")
        path_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        self.export_path_entry = ctk.CTkEntry(
            path_frame,
            font=ctk.CTkFont(size=11),
            height=30
        )
        self.export_path_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.export_path_entry.insert(0, output_path)
        
        def browse_export():
            path = filedialog.asksaveasfilename(
                title="Export Hasil Formatting",
                initialdir=os.path.dirname(self.original_path),
                initialfile=os.path.basename(output_path),
                defaultextension=".docx",
                filetypes=[("Dokumen Word", "*.docx"), ("Semua File", "*.*")]
            )
            if path:
                self.export_path_entry.delete(0, "end")
                self.export_path_entry.insert(0, path)
        
        ctk.CTkButton(
            path_frame, text="📁 Browse", width=80, height=30,
            command=browse_export
        ).pack(side="right")
        
        # Buttons
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x")
        
        def do_export():
            target = self.export_path_entry.get().strip()
            if not target:
                messagebox.showwarning("Warning", "Pilih lokasi export!")
                return
            try:
                shutil.copy2(output_path, target)
                self.log(f"📤 Di-export ke: {target}")
                messagebox.showinfo("Berhasil", f"File di-export ke:\n{target}")
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Gagal export:\n{e}")
        
        def open_folder():
            try:
                os.startfile(os.path.dirname(output_path))
            except Exception:
                pass
            dialog.destroy()
        
        def close_dialog():
            dialog.destroy()
        
        ctk.CTkButton(
            btn_frame, text="💾 Export", width=100, height=35,
            fg_color="#28a745", hover_color="#218838",
            command=do_export
        ).pack(side="left", padx=(0, 5))
        
        ctk.CTkButton(
            btn_frame, text="📁 Buka Folder", width=110, height=35,
            fg_color="#17a2b8", hover_color="#138496",
            command=open_folder
        ).pack(side="left", padx=(0, 5))
        
        ctk.CTkButton(
            btn_frame, text="✕ Tutup", width=80, height=35,
            fg_color="#6c757d", hover_color="#5a6268",
            command=close_dialog
        ).pack(side="right")

    def _export_output(self, temp_output_path: str):
        """Export file output ke lokasi pilihan user."""
        try:
            # Dapatkan nama default untuk export
            original_name = os.path.basename(self.original_path)
            name, ext = os.path.splitext(original_name)
            default_export_name = f"{name}{OUTPUT_SUFFIX}{ext}"
            
            # Tampilkan Save As dialog
            export_path = filedialog.asksaveasfilename(
                title="Export Hasil Formatting",
                initialdir=os.path.dirname(self.original_path),
                initialfile=default_export_name,
                defaultextension=".docx",
                filetypes=[
                    ("Dokumen Word", "*.docx"),
                    ("Semua File", "*.*")
                ]
            )
            
            if export_path:
                # Copy hasil ke lokasi pilihan user
                shutil.copy2(temp_output_path, export_path)
                self.log(f"📤 File berhasil di-export ke:")
                self.log(f"   {export_path}")
                self.log(f"🔒 Original file tetap aman: {self.original_path}")
                
                messagebox.showinfo(
                    "Export Berhasil ✅",
                    f"File berhasil di-export ke:\n{export_path}\n\n"
                    f"Original file tidak berubah:\n{self.original_path}"
                )
                
                # Tawarkan buka folder
                if messagebox.askyesno("Buka Folder?", "Buka folder lokasi export?"):
                    try:
                        os.startfile(os.path.dirname(export_path))
                    except Exception:
                        pass
            else:
                self.log("ℹ️ Export dibatalkan oleh user.")
        except Exception as exc:
            messagebox.showerror("Error Export", f"Gagal export file:\n{exc}")

    def _on_file_locked(self, err: PermissionError):
        self.is_processing = False
        self.btn_process.configure(state="normal", text="✨  PROSES DOKUMEN  ✨")
        self.set_status("⛔ File terkunci!", 0)
        self.log(f"⛔ GAGAL: {err}")
        messagebox.showerror(
            "⛔ File Sedang Dikunci",
            str(err)
        )

    def _on_failure(self, err: Exception):
        self.is_processing = False
        self.btn_process.configure(state="normal", text="✨  PROSES DOKUMEN  ✨")
        self.set_status("⛔ Terjadi error!", 0)
        self.log(f"⛔ ERROR: {err}")
        messagebox.showerror(
            "Terjadi Kesalahan",
            f"Proses gagal dijalankan:\n\n{err}\n\n"
            f"Pastikan file benar-benar berekstensi .docx\n"
            f"(bukan .doc lama) dan tidak rusak."
        )

    def _on_close(self):
        if self.is_processing:
            confirm = messagebox.askyesno(
                "Sedang Memproses",
                "Proses masih berjalan. Yakin ingin keluar?"
            )
            if not confirm:
                return
        
        # Simpan settings sebelum keluar
        try:
            self._save_settings()
        except Exception:
            pass
        
        self.destroy()


# ============================================================================
# ENTRY POINT
# ============================================================================

def main():
    """Entry point untuk aplikasi Italikin."""
    app = ItalikinApp()
    app.mainloop()


if __name__ == "__main__":
    main()
