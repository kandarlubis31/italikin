# -*- coding: utf-8 -*-
"""
COMPREHENSIVE TEST SCRIPT FOR ITALIKIN
Tests complex edge cases that might damage text
"""

import os
import sys
import copy
import tempfile
import shutil
from pathlib import Path

# Add current directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Import functions from italikin
from italikin import (
    parse_words,
    build_pattern,
    make_formatter,
    collect_paragraphs,
    format_paragraph,
    process_document,
    _locate_run,
    _split_run_at,
    _clone_run_after
)


class TestResult:
    def __init__(self):
        self.passed = 0
        self.failed = 0
        self.errors = []
        
    def check(self, name, condition, details=""):
        if condition:
            self.passed += 1
            print(f"  [PASS] {name}")
        else:
            self.failed += 1
            self.errors.append(f"{name}: {details}")
            print(f"  [FAIL] {name}")
            if details:
                print(f"         {details}")
    
    def summary(self):
        total = self.passed + self.failed
        print(f"\n{'='*60}")
        print(f"RESULTS: {self.passed}/{total} passed, {self.failed} failed")
        if self.errors:
            print(f"\nFailed tests:")
            for err in self.errors:
                print(f"  - {err}")
        print(f"{'='*60}\n")
        return self.failed == 0


def create_test_document(test_name: str, content_config: dict) -> str:
    """Create a test document with specific content."""
    doc = Document()
    
    # Title
    doc.add_heading(f'Test: {test_name}', level=1)
    
    # Regular paragraphs
    if 'paragraphs' in content_config:
        for para_config in content_config['paragraphs']:
            if isinstance(para_config, str):
                doc.add_paragraph(para_config)
            elif isinstance(para_config, dict):
                p = doc.add_paragraph()
                if 'runs' in para_config:
                    for run_config in para_config['runs']:
                        run = p.add_run(run_config['text'])
                        if run_config.get('bold'):
                            run.bold = True
                        if run_config.get('italic'):
                            run.italic = True
                        if run_config.get('color'):
                            run.font.color.rgb = RGBColor.from_string(run_config['color'])
    
    # Tables
    if 'tables' in content_config:
        for table_config in content_config['tables']:
            rows = table_config.get('rows', 2)
            cols = table_config.get('cols', 2)
            table = doc.add_table(rows=rows, cols=cols)
            
            if 'data' in table_config:
                for i, row_data in enumerate(table_config['data']):
                    if i < rows:
                        for j, cell_text in enumerate(row_data):
                            if j < cols:
                                table.cell(i, j).text = cell_text
    
    # Nested tables
    if 'nested_tables' in content_config:
        for nested_config in content_config['nested_tables']:
            # Create outer table
            outer_table = doc.add_table(rows=2, cols=2)
            outer_table.cell(0, 0).text = "Outer Cell 1"
            outer_table.cell(0, 1).text = "Outer Cell 2"
            
            # Add nested table in cell
            cell = outer_table.cell(1, 0)
            cell.text = ""  # Clear cell
            nested_table = cell.add_table(rows=2, cols=2)
            nested_table.cell(0, 0).text = nested_config.get('nested_text', 'Nested')
            nested_table.cell(0, 1).text = "Nested 2"
            nested_table.cell(1, 0).text = "Nested 3"
            nested_table.cell(1, 1).text = "Nested 4"
            
            outer_table.cell(1, 1).text = nested_config.get('outer_text', 'Outer with target')
    
    # Save to temp file
    temp_dir = tempfile.mkdtemp()
    filepath = os.path.join(temp_dir, f"{test_name}.docx")
    doc.save(filepath)
    
    return filepath, temp_dir


def test_basic_formatting():
    """Test basic formatting functionality."""
    print("\n" + "="*60)
    print("TEST 1: BASIC FORMATTING")
    print("="*60)
    
    result = TestResult()
    
    # Create test document
    content = {
        'paragraphs': [
            "PT Maju Jaya adalah perusahaan besar.",
            "PT XYZ juga merupakan perusahaan.",
            "Kata penting harus diitalic.",
        ]
    }
    
    filepath, temp_dir = create_test_document("basic_formatting", content)
    
    try:
        # Test parse_words
        words = parse_words("PT Maju Jaya, penting")
        result.check("Parse words", len(words) == 2, f"Expected 2, got {len(words)}")
        
        # Test build_pattern
        pattern = build_pattern(words, whole_word=True)
        result.check("Build pattern", pattern is not None, "Pattern is None")
        
        # Test format_paragraph on first paragraph (contains "PT Maju Jaya")
        doc = Document(filepath)
        
        # Find the paragraph with "PT Maju Jaya" (skip title)
        paragraph = None
        for p in doc.paragraphs:
            if "PT Maju Jaya" in p.text:
                paragraph = p
                break
        
        if paragraph is None:
            result.check("Paragraph found", False, "Could not find 'PT Maju Jaya' paragraph")
        else:
            formatter = make_formatter("italic")
            count = format_paragraph(paragraph, pattern, formatter)
            result.check("Format paragraph", count > 0, f"Count: {count}")
            
            # Verify text is preserved
            original_text = "PT Maju Jaya adalah perusahaan besar."
            formatted_text = paragraph.text
            result.check("Text preserved", original_text == formatted_text, 
                         f"Original: {original_text}\nFormatted: {formatted_text}")
            
            # Verify formatting applied
            has_italic = any(run.italic for run in paragraph.runs)
            result.check("Italic applied", has_italic, "No italic found in any run")
        
    finally:
        shutil.rmtree(temp_dir)
    
    return result


def test_nested_tables():
    """Test with nested tables."""
    print("\n" + "="*60)
    print("TEST 2: NESTED TABLES")
    print("="*60)
    
    result = TestResult()
    
    content = {
        'nested_tables': [
            {
                'nested_text': 'PT Nested Target',
                'outer_text': 'PT Outer Target here'
            }
        ]
    }
    
    filepath, temp_dir = create_test_document("nested_tables", content)
    
    try:
        doc = Document(filepath)
        paragraphs = collect_paragraphs(doc)
        
        # Find paragraphs with our content (skip title)
        content_paragraphs = [p for p in paragraphs if "PT" in p.text]
        
        result.check("Collect paragraphs from nested", len(content_paragraphs) > 0,
                     f"Found {len(content_paragraphs)} content paragraphs")
        
        # Check if nested table paragraphs are included
        has_nested = any("PT Nested Target" in p.text for p in content_paragraphs)
        result.check("Nested table content found", has_nested,
                     "Nested table content not found in collected paragraphs")
        
        # Test formatting
        words = ["PT"]
        pattern = build_pattern(words, whole_word=True)
        formatter = make_formatter("bold")
        
        total_formatted = 0
        for para in paragraphs:
            total_formatted += format_paragraph(para, pattern, formatter)
        
        result.check("Format nested content", total_formatted > 0,
                     f"Formatted {total_formatted} occurrences")
        
    finally:
        shutil.rmtree(temp_dir)
    
    return result


def test_word_boundary():
    """Test word boundary matching."""
    print("\n" + "="*60)
    print("TEST 3: WORD BOUNDARY MATCHING")
    print("="*60)
    
    result = TestResult()
    
    content = {
        'paragraphs': [
            "PT adalah singkatan dari Perseroan Terbatas.",
            "PTP tidak boleh terkena.",
            "PTXYZ juga tidak boleh.",
            "PT Maju Jaya boleh.",
        ]
    }
    
    filepath, temp_dir = create_test_document("word_boundary", content)
    
    try:
        doc = Document(filepath)
        
        # Test with whole_word=True
        words = ["PT"]
        pattern_whole = build_pattern(words, whole_word=True)
        formatter = make_formatter("italic")
        
        # Find paragraph with "PT adalah" (skip title)
        paragraph = None
        for p in doc.paragraphs:
            if "PT adalah" in p.text:
                paragraph = p
                break
        
        if paragraph:
            count_whole = format_paragraph(paragraph, pattern_whole, formatter)
        else:
            count_whole = 0
            result.check("Paragraph found", False, "Could not find 'PT adalah' paragraph")
        
        # PT at start should match
        result.check("Whole word: PT at start matches", count_whole == 1,
                     f"Expected 1, got {count_whole}")
        
        # Test that PTP, PTXYZ don't match
        doc2 = Document(filepath)
        
        # Find PTP paragraph
        paragraph2 = None
        for p in doc2.paragraphs:
            if "PTP" in p.text:
                paragraph2 = p
                break
        
        if paragraph2:
            count_ptp = format_paragraph(paragraph2, pattern_whole, formatter)
            result.check("Whole word: PTP doesn't match", count_ptp == 0,
                         f"Expected 0, got {count_ptp}")
        else:
            result.check("PTP paragraph found", False, "Could not find 'PTP' paragraph")
        
        # Find PTXYZ paragraph
        paragraph3 = None
        for p in doc2.paragraphs:
            if "PTXYZ" in p.text:
                paragraph3 = p
                break
        
        if paragraph3:
            count_ptxyz = format_paragraph(paragraph3, pattern_whole, formatter)
            result.check("Whole word: PTXYZ doesn't match", count_ptxyz == 0,
                         f"Expected 0, got {count_ptxyz}")
        else:
            result.check("PTXYZ paragraph found", False, "Could not find 'PTXYZ' paragraph")
        
    finally:
        shutil.rmtree(temp_dir)
    
    return result


def test_case_sensitivity():
    """Test case sensitivity."""
    print("\n" + "="*60)
    print("TEST 4: CASE SENSITIVITY")
    print("="*60)
    
    result = TestResult()
    
    content = {
        'paragraphs': [
            "PT Maju Jaya adalah perusahaan.",
            "pt maju jaya juga ada.",
            "Pt MaJu JaYa ada juga.",
        ]
    }
    
    filepath, temp_dir = create_test_document("case_sensitivity", content)
    
    try:
        # Test case-insensitive (default)
        doc = Document(filepath)
        words = ["PT Maju Jaya"]
        pattern_insensitive = build_pattern(words, whole_word=True, case_sensitive=False)
        formatter = make_formatter("bold")
        
        total_insensitive = 0
        for para in doc.paragraphs:
            if "PT" in para.text or "pt" in para.text.lower():
                total_insensitive += format_paragraph(para, pattern_insensitive, formatter)
        
        # Should match all 3 variations (case-insensitive)
        result.check("Case-insensitive: matches all variations", total_insensitive >= 3,
                     f"Expected >= 3, got {total_insensitive}")
        
        # Test case-sensitive
        doc2 = Document(filepath)
        pattern_sensitive = build_pattern(words, whole_word=True, case_sensitive=True)
        
        total_sensitive = 0
        for para in doc2.paragraphs:
            if "PT Maju Jaya" in para.text:  # Only exact case
                total_sensitive += format_paragraph(para, pattern_sensitive, formatter)
        
        result.check("Case-sensitive: matches only exact case", total_sensitive == 1,
                     f"Expected 1, got {total_sensitive}")
        
    finally:
        shutil.rmtree(temp_dir)
    
    return result


def test_format_preservation():
    """Test that original formatting is preserved."""
    print("\n" + "="*60)
    print("TEST 5: FORMAT PRESERVATION")
    print("="*60)
    
    result = TestResult()
    
    content = {
        'paragraphs': [
            {
                'runs': [
                    {'text': 'Normal '},
                    {'text': 'Bold Text ', 'bold': True},
                    {'text': 'and '},
                    {'text': 'Italic Text', 'italic': True},
                    {'text': ' target here.'},
                ]
            }
        ]
    }
    
    filepath, temp_dir = create_test_document("format_preservation", content)
    
    try:
        doc = Document(filepath)
        # Find the paragraph with our content (skip title)
        paragraph = None
        for p in doc.paragraphs:
            if "Normal Bold Text" in p.text:
                paragraph = p
                break
        
        if paragraph is None:
            result.check("Paragraph found", False, "Could not find test paragraph")
        else:
            # Record original formatting
            original_bold = paragraph.runs[1].bold if len(paragraph.runs) > 1 else None
            original_italic = paragraph.runs[3].italic if len(paragraph.runs) > 3 else None
            
            # Apply italic to "target"
            words = ["target"]
            pattern = build_pattern(words, whole_word=True)
            formatter = make_formatter("italic")
            
            format_paragraph(paragraph, pattern, formatter)
            
            # Check that original formatting is preserved
            # Find the run that contains "target" (should be new run after split)
            target_found = False
            for run in paragraph.runs:
                if "target" in run.text.lower():
                    target_found = True
                    # The target run should have italic
                    result.check("Target has italic", run.italic == True,
                               f"italic={run.italic}")
            
            result.check("Target run found", target_found, "Target text not found")
            
            # Check that other runs still have their original formatting
            # Note: After split, we need to find the original runs
            full_text = "".join(r.text for r in paragraph.runs)
            result.check("Full text preserved", "Normal Bold Text and Italic Text target here." == full_text,
                         f"Text: {full_text}")
        
    finally:
        shutil.rmtree(temp_dir)
    
    return result


def test_special_characters():
    """Test with special characters and unicode."""
    print("\n" + "="*60)
    print("TEST 6: SPECIAL CHARACTERS & UNICODE")
    print("="*60)
    
    result = TestResult()
    
    content = {
        'paragraphs': [
            "Kata dengan tanda baca: PT, Maju Jaya!",
            "Kata dengan angka: PT123 tidak boleh.",
            "Kata dengan simbol: PT@Maju tidak boleh.",
            "Unicode: café, naïve, résumé target ada.",
            "Emoji: ⭐ PT Maju Jaya ⭐ harus diitalic.",
        ]
    }
    
    filepath, temp_dir = create_test_document("special_chars", content)
    
    try:
        doc = Document(filepath)
        
        # Test various special characters
        test_cases = [
            (["PT"], True, "PT with comma"),
            (["PT123"], True, "PT123 matches itself"),
            (["café"], True, "Unicode cafe"),
            (["naïve"], True, "Unicode naive"),
            (["⭐"], True, "Emoji star"),
        ]
        
        for words, should_match, desc in test_cases:
            doc_copy = Document(filepath)  # Fresh copy
            pattern = build_pattern(words, whole_word=True)
            formatter = make_formatter("underline")
            
            total = 0
            for para in doc_copy.paragraphs:
                total += format_paragraph(para, pattern, formatter)
            
            if should_match:
                result.check(f"Special char: {desc}", total > 0,
                           f"Expected match, got {total}")
            else:
                result.check(f"Special char: {desc}", total == 0,
                           f"Expected no match, got {total}")
        
    finally:
        shutil.rmtree(temp_dir)
    
    return result


def test_overlapping_patterns():
    """Test with overlapping keyword patterns."""
    print("\n" + "="*60)
    print("TEST 7: OVERLAPPING PATTERNS")
    print("="*60)
    
    result = TestResult()
    
    content = {
        'paragraphs': [
            "PT adalah singkatan.",
            "PT Maju Jaya adalah perusahaan.",
            "Maju Jaya Group ada.",
            "PT Maju Jaya Group besar.",
        ]
    }
    
    filepath, temp_dir = create_test_document("overlapping", content)
    
    try:
        doc = Document(filepath)
        
        # Test with multiple overlapping keywords
        words = ["PT", "PT Maju Jaya", "Maju Jaya"]
        pattern = build_pattern(words, whole_word=True)
        formatter = make_formatter("italic")
        
        total = 0
        for para in doc.paragraphs:
            if any(w in para.text for w in ["PT", "Maju Jaya"]):
                count = format_paragraph(para, pattern, formatter)
                total += count
        
        # Should match all occurrences (longer patterns first to avoid substring issues)
        result.check("Overlapping patterns: total matches", total >= 4,
                     f"Expected >= 4, got {total}")
        
        # Verify longer patterns match first
        doc2 = Document(filepath)
        paragraph = None
        for p in doc2.paragraphs:
            if "PT Maju Jaya Group" in p.text:
                paragraph = p
                break
        
        if paragraph:
            pattern2 = build_pattern(["PT", "Maju Jaya"], whole_word=True)
            count = format_paragraph(paragraph, pattern2, formatter)
            
            # Both "PT" and "Maju Jaya" should match in this paragraph
            result.check("Overlapping: both patterns match", count >= 2,
                         f"Expected >= 2, got {count}")
        else:
            result.check("PT Maju Jaya paragraph found", False, "Could not find test paragraph")
        
    finally:
        shutil.rmtree(temp_dir)
    
    return result


def test_large_document():
    """Test with large document."""
    print("\n" + "="*60)
    print("TEST 8: LARGE DOCUMENT PERFORMANCE")
    print("="*60)
    
    result = TestResult()
    
    # Create large document
    paragraphs_content = []
    for i in range(200):
        if i % 10 == 0:
            paragraphs_content.append(f"PT Target {i} adalah penting.")
        else:
            paragraphs_content.append(f"Paragraf biasa nomor {i} tanpa target.")
    
    content = {'paragraphs': paragraphs_content}
    filepath, temp_dir = create_test_document("large_doc", content)
    
    try:
        doc = Document(filepath)
        paragraphs = collect_paragraphs(doc)
        
        # Account for title paragraph
        content_paras = [p for p in paragraphs if "Paragraf" in p.text or "PT Target" in p.text]
        result.check("Large doc: paragraphs collected", len(content_paras) == 200,
                     f"Expected 200, got {len(content_paras)}")
        
        # Test formatting
        words = ["PT", "penting"]
        pattern = build_pattern(words, whole_word=True)
        formatter = make_formatter("bold")
        
        total = 0
        for para in paragraphs:
            total += format_paragraph(para, pattern, formatter)
        
        result.check("Large doc: formatting completed", total > 0,
                     f"Formatted {total} occurrences")
        
        # Verify no text damage
        doc2 = Document(filepath)
        all_text = "".join(p.text for p in doc2.paragraphs)
        result.check("Large doc: text intact", "Paragraf biasa" in all_text,
                     "Text appears damaged")
        
    finally:
        shutil.rmtree(temp_dir)
    
    return result


def test_edge_cases():
    """Test edge cases."""
    print("\n" + "="*60)
    print("TEST 9: EDGE CASES")
    print("="*60)
    
    result = TestResult()
    
    # Test empty document
    content = {'paragraphs': []}
    filepath, temp_dir = create_test_document("empty", content)
    
    try:
        doc = Document(filepath)
        paragraphs = collect_paragraphs(doc)
        # Account for title paragraph
        content_paras = [p for p in paragraphs if p.text.strip() and "Test:" not in p.text]
        result.check("Empty document: no content paragraphs", len(content_paras) == 0,
                     f"Expected 0, got {len(content_paras)}")
    finally:
        shutil.rmtree(temp_dir)
    
    # Test document with only whitespace
    content = {'paragraphs': ["   ", "\t", "\n"]}
    filepath, temp_dir = create_test_document("whitespace", content)
    
    try:
        doc = Document(filepath)
        words = ["target"]
        pattern = build_pattern(words, whole_word=True)
        formatter = make_formatter("italic")
        
        total = 0
        for para in doc.paragraphs:
            total += format_paragraph(para, pattern, formatter)
        
        result.check("Whitespace only: no matches", total == 0,
                     f"Expected 0, got {total}")
    finally:
        shutil.rmtree(temp_dir)
    
    # Test very long word
    long_word = "A" * 100
    content = {'paragraphs': [f"Kata panjang {long_word} ada di sini."]}
    filepath, temp_dir = create_test_document("long_word", content)
    
    try:
        doc = Document(filepath)
        words = [long_word]
        pattern = build_pattern(words, whole_word=True)
        formatter = make_formatter("bold")
        
        total = 0
        for para in doc.paragraphs:
            total += format_paragraph(para, pattern, formatter)
        
        result.check("Long word: matches correctly", total == 1,
                     f"Expected 1, got {total}")
    finally:
        shutil.rmtree(temp_dir)
    
    return result


def test_run_splitting():
    """Test run splitting algorithm."""
    print("\n" + "="*60)
    print("TEST 10: RUN SPLITTING ALGORITHM")
    print("="*60)
    
    result = TestResult()
    
    # Create document with known run structure
    doc = Document()
    p = doc.add_paragraph()
    run1 = p.add_run("Hello ")
    run2 = p.add_run("World target here")
    run3 = p.add_run(" end")
    
    # Test _locate_run
    idx, offset = _locate_run(p.runs, 0)
    result.check("Locate run: start of first run", idx == 0 and offset == 0,
                 f"idx={idx}, offset={offset}")
    
    idx, offset = _locate_run(p.runs, 6)
    result.check("Locate run: start of second run", idx == 1 and offset == 0,
                 f"idx={idx}, offset={offset}")
    
    idx, offset = _locate_run(p.runs, 12)
    result.check("Locate run: middle of second run", idx == 1 and offset == 6,
                 f"idx={idx}, offset={offset}")
    
    # Test _split_run_at
    doc2 = Document()
    p2 = doc2.add_paragraph()
    run = p2.add_run("Hello World")
    
    new_run = _split_run_at(run, 5)
    result.check("Split run: original run has left part", run.text == "Hello",
                 f"Got: {run.text}")
    result.check("Split run: new run has right part", new_run.text == " World",
                 f"Got: {new_run.text}")
    result.check("Split run: total text preserved", 
                 run.text + new_run.text == "Hello World",
                 f"Got: {run.text + new_run.text}")
    
    return result


def run_all_tests():
    """Run all tests."""
    print("\n" + "="*60)
    print("ITALIKIN COMPREHENSIVE TEST SUITE")
    print("="*60)
    
    all_results = []
    
    tests = [
        ("Basic Formatting", test_basic_formatting),
        ("Nested Tables", test_nested_tables),
        ("Word Boundary", test_word_boundary),
        ("Case Sensitivity", test_case_sensitivity),
        ("Format Preservation", test_format_preservation),
        ("Special Characters", test_special_characters),
        ("Overlapping Patterns", test_overlapping_patterns),
        ("Large Document", test_large_document),
        ("Edge Cases", test_edge_cases),
        ("Run Splitting", test_run_splitting),
    ]
    
    for test_name, test_func in tests:
        try:
            result = test_func()
            all_results.append((test_name, result))
        except Exception as e:
            print(f"\n[CRASH] {test_name}: {e}")
            all_results.append((test_name, None))
    
    # Summary
    print("\n" + "="*60)
    print("FINAL SUMMARY")
    print("="*60)
    
    total_passed = 0
    total_failed = 0
    
    for test_name, result in all_results:
        if result is None:
            status = "[CRASH]"
            total_failed += 1
        elif result.failed == 0:
            status = "[PASS]"
            total_passed += result.passed
        else:
            status = f"[FAIL] ({result.failed} failures)"
            total_passed += result.passed
            total_failed += result.failed
        print(f"  {status}: {test_name}")
    
    print(f"\n{'='*60}")
    print(f"TOTAL: {total_passed} passed, {total_failed} failed")
    print(f"{'='*60}\n")
    
    return total_failed == 0


if __name__ == "__main__":
    success = run_all_tests()
    sys.exit(0 if success else 1)
