#!/usr/bin/env python3
"""
Multi-Bank PDF Statement Parser
Converts HDFC/SBI/IOB bank statement PDFs to CSV/Excel using coordinate-based layout parsing.
Author: Antigravity AI
"""

import os
import sys
import re
import csv
import argparse
from datetime import datetime

# Try to import fitz (PyMuPDF) or pdfplumber
HAS_FITZ = False
try:
    import fitz
    HAS_FITZ = True
except ImportError:
    pass

HAS_PDFPLUMBER = False
try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    pass

if not HAS_FITZ and not HAS_PDFPLUMBER:
    print("Error: Neither 'pymupdf' nor 'pdfplumber' is installed. Please install at least one of them (e.g. 'pip install pymupdf').")
    sys.exit(1)

# Try to import openpyxl for Excel output
try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


class BankStatementParser:
    """Unified parser for HDFC and SBI bank statement PDFs with auto-detection."""

    def __init__(self, pdf_path, verbose=False, bank=None):
        self.pdf_path = pdf_path
        self.verbose = verbose
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")
        self.bank = (bank or self.detect_bank(pdf_path)).lower()

    # ── Bank auto-detection ──────────────────────────────────────────────

    @staticmethod
    def detect_bank(pdf_path):
        """Read page 1 text and detect the bank."""
        try:
            if HAS_FITZ:
                doc = fitz.open(pdf_path)
                text = doc[0].get_text()
                doc.close()
            else:
                with pdfplumber.open(pdf_path) as pdf:
                    text = pdf.pages[0].extract_text() or ""
            text_lower = text.lower()
            if "state bank of india" in text_lower:
                return "sbi"
            if "indian overseas bank" in text_lower or "ioba" in text_lower:
                return "iob"
            if "hdfc bank" in text_lower:
                return "hdfc"
            return "hdfc"  # default
        except Exception:
            return "hdfc"

    # ── Shared helpers ──────────────────────────────────────────────────

    def clean_amount(self, amount_str):
        """Cleans formatting from amount strings and converts to float/string representation."""
        if not amount_str:
            return None
        cleaned = amount_str.replace(",", "").strip()
        if not cleaned:
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None

    def format_amount(self, val):
        if val is None:
            return ""
        return f"{val:.2f}"

    # ── Top-level dispatch ───────────────────────────────────────────────

    def parse(self):
        if self.bank == "sbi":
            print(f"Detected bank: SBI")
            if HAS_FITZ:
                try:
                    return self.parse_sbi_fitz()
                except Exception as e:
                    print(f"PyMuPDF parsing failed: {e}. Falling back to pdfplumber...")
            return self.parse_sbi_pdfplumber()
        elif self.bank == "iob":
            print(f"Detected bank: Indian Overseas Bank (IOB)")
            if HAS_FITZ:
                try:
                    return self.parse_iob_fitz()
                except Exception as e:
                    print(f"PyMuPDF parsing failed: {e}. Falling back to pdfplumber...")
            return self.parse_iob_pdfplumber()
        else:
            print(f"Detected bank: HDFC")
            if HAS_FITZ:
                try:
                    return self.parse_hdfc_fitz()
                except Exception as e:
                    print(f"PyMuPDF parsing failed: {e}. Falling back to pdfplumber...")
            return self.parse_hdfc_pdfplumber()

    # ═══════════════════════════════════════════════════════════════════
    #  HDFC Parsing (existing logic, unchanged)
    # ═══════════════════════════════════════════════════════════════════

    def parse_hdfc_fitz(self):
        print(f"Opening PDF (PyMuPDF): {self.pdf_path}")
        transactions = []
        current_tx = None
        date_pattern = re.compile(r'^\d{2}/\d{2}/\d{2,4}$')

        col_definitions = [
            ('date', 30, 65),
            ('narration', 65, 280),
            ('chq_ref', 280, 350),
            ('val_date', 350, 400),
            ('withdrawal', 400, 480),
            ('deposit', 480, 560),
            ('balance', 560, 630),
        ]

        doc = fitz.open(self.pdf_path)
        total_pages = len(doc)
        print(f"Total pages to parse: {total_pages}")

        for page_idx in range(total_pages):
            page = doc[page_idx]
            page_num = page_idx + 1

            if self.verbose or page_num % 50 == 0 or page_num == total_pages:
                print(f"Processing page {page_num}/{total_pages}...")

            words = page.get_text("words")
            table_words = [w for w in words if 225 <= w[1] <= 780]

            grouped_lines = {}
            for w in table_words:
                y0_val = w[1]
                found = False
                for gk in grouped_lines.keys():
                    if abs(gk - y0_val) < 2.0:
                        grouped_lines[gk].append(w)
                        found = True
                        break
                if not found:
                    grouped_lines[y0_val] = [w]

            for gk in sorted(grouped_lines.keys()):
                line_words = sorted(grouped_lines[gk], key=lambda w: w[0])
                cols = {name: [] for name, _, _ in col_definitions}
                for w in line_words:
                    x0 = w[0]
                    text = w[4]
                    for name, start, end in col_definitions:
                        if start <= x0 < end:
                            cols[name].append(text)
                            break
                ld = {name: " ".join(words_list).strip() for name, words_list in cols.items()}

                if ld['date'] == 'Date' or ld['narration'] == 'Narration':
                    continue
                if not any(ld.values()):
                    continue

                is_new_tx = date_pattern.match(ld['date'])

                if is_new_tx:
                    if current_tx:
                        transactions.append(current_tx)
                    current_tx = {
                        'date': ld['date'],
                        'narration': ld['narration'],
                        'chq_ref': ld['chq_ref'],
                        'val_date': ld['val_date'],
                        'withdrawal': self.clean_amount(ld['withdrawal']),
                        'deposit': self.clean_amount(ld['deposit']),
                        'balance': self.clean_amount(ld['balance']),
                        'page': page_num,
                    }
                else:
                    if current_tx:
                        continuation_parts = []
                        for col_name in ['date', 'narration', 'chq_ref', 'val_date', 'withdrawal', 'deposit', 'balance']:
                            val = ld[col_name]
                            if val:
                                continuation_parts.append(val)
                        continuation_text = " ".join(continuation_parts).strip()
                        if continuation_text:
                            current_tx['narration'] += " " + continuation_text

        if current_tx:
            transactions.append(current_tx)

        print(f"Extraction complete. Found {len(transactions)} raw transactions.")
        return transactions

    def parse_hdfc_pdfplumber(self):
        print(f"Opening PDF (pdfplumber): {self.pdf_path}")
        transactions = []
        current_tx = None
        date_pattern = re.compile(r'^\d{2}/\d{2}/\d{2,4}$')

        col_definitions = [
            ('date', 30, 65),
            ('narration', 65, 280),
            ('chq_ref', 280, 350),
            ('val_date', 350, 400),
            ('withdrawal', 400, 480),
            ('deposit', 480, 560),
            ('balance', 560, 630),
        ]

        with pdfplumber.open(self.pdf_path) as pdf:
            total_pages = len(pdf.pages)
            print(f"Total pages to parse: {total_pages}")

            for page_idx in range(total_pages):
                page = pdf.pages[page_idx]
                page_num = page_idx + 1

                if self.verbose or page_num % 10 == 0 or page_num == total_pages:
                    print(f"Processing page {page_num}/{total_pages}...")

                words = page.extract_words()
                table_words = [w for w in words if 225 <= w['top'] <= 780]

                grouped_lines = {}
                for w in table_words:
                    top_val = w['top']
                    found = False
                    for gk in grouped_lines.keys():
                        if abs(gk - top_val) < 2.0:
                            grouped_lines[gk].append(w)
                            found = True
                            break
                    if not found:
                        grouped_lines[top_val] = [w]

                for gk in sorted(grouped_lines.keys()):
                    line_words = sorted(grouped_lines[gk], key=lambda w: w['x0'])
                    cols = {name: [] for name, _, _ in col_definitions}
                    for w in line_words:
                        x0 = w['x0']
                        text = w['text']
                        for name, start, end in col_definitions:
                            if start <= x0 < end:
                                cols[name].append(text)
                                break
                    ld = {name: " ".join(words_list).strip() for name, words_list in cols.items()}

                    if ld['date'] == 'Date' or ld['narration'] == 'Narration':
                        continue
                    if not any(ld.values()):
                        continue

                    is_new_tx = date_pattern.match(ld['date'])

                    if is_new_tx:
                        if current_tx:
                            transactions.append(current_tx)
                        current_tx = {
                            'date': ld['date'],
                            'narration': ld['narration'],
                            'chq_ref': ld['chq_ref'],
                            'val_date': ld['val_date'],
                            'withdrawal': self.clean_amount(ld['withdrawal']),
                            'deposit': self.clean_amount(ld['deposit']),
                            'balance': self.clean_amount(ld['balance']),
                            'page': page_num,
                        }
                    else:
                        if current_tx:
                            continuation_parts = []
                            for col_name in ['date', 'narration', 'chq_ref', 'val_date', 'withdrawal', 'deposit', 'balance']:
                                val = ld[col_name]
                                if val:
                                    continuation_parts.append(val)
                            continuation_text = " ".join(continuation_parts).strip()
                            if continuation_text:
                                current_tx['narration'] += " " + continuation_text

            if current_tx:
                transactions.append(current_tx)

        print(f"Extraction complete. Found {len(transactions)} raw transactions.")
        return transactions

    # ═══════════════════════════════════════════════════════════════════
    #  SBI Parsing
    # ═══════════════════════════════════════════════════════════════════

    def parse_sbi_fitz(self):
        print(f"Opening PDF (PyMuPDF): {self.pdf_path}")
        transactions = []
        current_tx = None
        date_pattern = re.compile(r'^\d{2}/\d{2}/\d{4}$')

        doc = fitz.open(self.pdf_path)
        total_pages = len(doc)
        print(f"Total pages to parse (SBI): {total_pages}")

        for page_idx in range(total_pages):
            page = doc[page_idx]
            page_num = page_idx + 1

            if self.verbose or page_num % 10 == 0 or page_num == total_pages:
                print(f"Processing page {page_num}/{total_pages}...")

            words = page.get_text("words")
            # SBI table region: y 40-730 (excludes page headers & footers)
            table_words = [w for w in words if 40 <= w[1] <= 730]

            # Group by y-coordinate
            grouped_lines = {}
            for w in table_words:
                y0_val = w[1]
                found = False
                for gk in grouped_lines.keys():
                    if abs(gk - y0_val) < 2.0:
                        grouped_lines[gk].append(w)
                        found = True
                        break
                if not found:
                    grouped_lines[y0_val] = [w]

            sorted_ys = sorted(grouped_lines.keys())

            # Pass 1: find all date-line y values (transaction headers)
            date_ys = []
            for gk in sorted_ys:
                line_words = sorted(grouped_lines[gk], key=lambda w: w[0])
                date_text = " ".join(w[4] for w in line_words if 20 <= w[0] < 78).strip()
                if date_pattern.match(date_text):
                    date_ys.append(gk)

            if not date_ys:
                continue  # skip summary / header pages

            # Pass 2: chunk the page into transaction blocks.
            # Each block = [date_y - 12, next_date_y - 9)  so the
            # transaction-type label (always ~7 pt above the date line)
            # is captured with its transaction, but the next transaction's
            # type label is excluded.
            for i, dy in enumerate(date_ys):
                if current_tx:
                    transactions.append(current_tx)
                    current_tx = None

                block_end = date_ys[i + 1] - 9 if i + 1 < len(date_ys) else (sorted_ys[-1] + 1)
                block_ys = [gk for gk in sorted_ys if dy - 12 <= gk < block_end]

                narration_parts = []
                date_text = ""
                val_date_text = ""
                wd_text = None
                dp_text = None
                bal_text = None

                for gk in block_ys:
                    line_words = sorted(grouped_lines[gk], key=lambda w: w[0])
                    if gk == dy:
                        # Data line – extract dates and amounts
                        date_text = " ".join(w[4] for w in line_words if 20 <= w[0] < 78).strip()
                        val_date_text = " ".join(w[4] for w in line_words if 78 <= w[0] < 130).strip()
                        wd_raw = [w[4] for w in line_words if 300 <= w[0] < 400 and w[4] != '-']
                        dp_raw = [w[4] for w in line_words if 400 <= w[0] < 480 and w[4] != '-']
                        bal_raw = [w[4] for w in line_words if 480 <= w[0] < 560 and w[4] != '-']
                        wd_text = wd_raw[0] if wd_raw else None
                        dp_text = dp_raw[0] if dp_raw else None
                        bal_text = bal_raw[0] if bal_raw else None
                    # Narration: everything in the group from x >= 130
                    n_parts = [w[4] for w in line_words if w[0] >= 130]
                    if n_parts:
                        narration_parts.extend(n_parts)

                current_tx = {
                    'date': date_text,
                    'narration': " ".join(narration_parts).strip(),
                    'chq_ref': '',
                    'val_date': val_date_text if val_date_text else date_text,
                    'withdrawal': self.clean_amount(wd_text),
                    'deposit': self.clean_amount(dp_text),
                    'balance': self.clean_amount(bal_text),
                    'page': page_num,
                }

        if current_tx:
            transactions.append(current_tx)

        print(f"Extraction complete. Found {len(transactions)} raw transactions.")
        return transactions

    def parse_sbi_pdfplumber(self):
        print(f"Opening PDF (pdfplumber): {self.pdf_path}")
        transactions = []
        current_tx = None
        date_pattern = re.compile(r'^\d{2}/\d{2}/\d{4}$')

        with pdfplumber.open(self.pdf_path) as pdf:
            total_pages = len(pdf.pages)
            print(f"Total pages to parse (SBI): {total_pages}")

            for page_idx in range(total_pages):
                page = pdf.pages[page_idx]
                page_num = page_idx + 1

                if self.verbose or page_num % 10 == 0 or page_num == total_pages:
                    print(f"Processing page {page_num}/{total_pages}...")

                words = page.extract_words()
                table_words = [w for w in words if 40 <= w['top'] <= 730]

                grouped_lines = {}
                for w in table_words:
                    top_val = w['top']
                    found = False
                    for gk in grouped_lines.keys():
                        if abs(gk - top_val) < 2.0:
                            grouped_lines[gk].append(w)
                            found = True
                            break
                    if not found:
                        grouped_lines[top_val] = [w]

                sorted_ys = sorted(grouped_lines.keys())

                date_ys = []
                for gk in sorted_ys:
                    line_words = sorted(grouped_lines[gk], key=lambda w: w['x0'])
                    date_text = " ".join(w['text'] for w in line_words if 20 <= w['x0'] < 78).strip()
                    if date_pattern.match(date_text):
                        date_ys.append(gk)

                if not date_ys:
                    continue

                for i, dy in enumerate(date_ys):
                    if current_tx:
                        transactions.append(current_tx)
                        current_tx = None

                    block_end = date_ys[i + 1] - 9 if i + 1 < len(date_ys) else (sorted_ys[-1] + 1)
                    block_ys = [gk for gk in sorted_ys if dy - 12 <= gk < block_end]

                    narration_parts = []
                    date_text = ""
                    val_date_text = ""
                    wd_text = None
                    dp_text = None
                    bal_text = None

                    for gk in block_ys:
                        line_words = sorted(grouped_lines[gk], key=lambda w: w['x0'])
                        if gk == dy:
                            date_text = " ".join(w['text'] for w in line_words if 20 <= w['x0'] < 78).strip()
                            val_date_text = " ".join(w['text'] for w in line_words if 78 <= w['x0'] < 130).strip()
                            wd_raw = [w['text'] for w in line_words if 300 <= w['x0'] < 400 and w['text'] != '-']
                            dp_raw = [w['text'] for w in line_words if 400 <= w['x0'] < 480 and w['text'] != '-']
                            bal_raw = [w['text'] for w in line_words if 480 <= w['x0'] < 560 and w['text'] != '-']
                            wd_text = wd_raw[0] if wd_raw else None
                            dp_text = dp_raw[0] if dp_raw else None
                            bal_text = bal_raw[0] if bal_raw else None
                        n_parts = [w['text'] for w in line_words if w['x0'] >= 130]
                        if n_parts:
                            narration_parts.extend(n_parts)

                    current_tx = {
                        'date': date_text,
                        'narration': " ".join(narration_parts).strip(),
                        'chq_ref': '',
                        'val_date': val_date_text if val_date_text else date_text,
                        'withdrawal': self.clean_amount(wd_text),
                        'deposit': self.clean_amount(dp_text),
                        'balance': self.clean_amount(bal_text),
                        'page': page_num,
                    }

        if current_tx:
            transactions.append(current_tx)

        print(f"Extraction complete. Found {len(transactions)} raw transactions.")
        return transactions

    # ═══════════════════════════════════════════════════════════════════
    #  IOB Parsing (Indian Overseas Bank)
    #  Reverse chronological, 2-line blocks, DD-Mon-YY dates
    # ═══════════════════════════════════════════════════════════════════

    def parse_iob_fitz(self):
        print(f"Opening PDF (PyMuPDF): {self.pdf_path}")
        transactions = []
        date_pattern = re.compile(r'^\d{2}-[A-Z][a-z]{2}-\d{2}$')

        doc = fitz.open(self.pdf_path)
        total_pages = len(doc)
        print(f"Total pages to parse (IOB): {total_pages}")

        for page_idx in range(total_pages):
            page = doc[page_idx]
            page_num = page_idx + 1

            if self.verbose or page_num % 10 == 0 or page_num == total_pages:
                print(f"Processing page {page_num}/{total_pages}...")

            words = page.get_text("words")
            table_words = [w for w in words if 30 <= w[1] <= 790]
            table_words = [w for w in table_words if not w[4].startswith('Page ')]
            if not table_words:
                continue

            grouped_lines = {}
            for w in table_words:
                y0_val = w[1]
                found = False
                for gk in grouped_lines.keys():
                    if abs(gk - y0_val) < 2.0:
                        grouped_lines[gk].append(w)
                        found = True
                        break
                if not found:
                    grouped_lines[y0_val] = [w]

            sorted_ys = sorted(grouped_lines.keys())

            # Find detail lines (have S-prefixed ref at 278 ≤ x < 344)
            ref_pat = re.compile(r'^S\d+')
            detail_ys = []
            for gk in sorted_ys:
                line_words = sorted(grouped_lines[gk], key=lambda w: w[0])
                if any(278 <= w[0] < 344 and ref_pat.match(w[4]) for w in line_words):
                    detail_ys.append(gk)

            if not detail_ys:
                continue

            # Pre-scan: map each line to its date
            line_date = {}
            line_vd = {}
            for gk in sorted_ys:
                line_words = sorted(grouped_lines[gk], key=lambda w: w[0])
                for w in line_words:
                    x, text = w[0], w[4]
                    if 40 <= x < 90:
                        m = re.match(r'^\((\d{2}-[A-Z][a-z]{2}-\d{2})\)$', text)
                        if m:
                            line_vd[gk] = m.group(1)
                            continue
                        if date_pattern.match(text) and gk not in line_date:
                            line_date[gk] = text

            # For each detail line, pair with its preceding date
            for dk in detail_ys:
                dk_line = sorted(grouped_lines[dk], key=lambda w: w[0])

                # Extract ref from detail line
                ref_parts = [w[4] for w in dk_line if 278 <= w[0] < 344 and ref_pat.match(w[4])]

                # Collect all amount-like tokens at x >= 398, sorted by x
                # Use x-ranges: Debit(398-456) Credit(456-510) Balance(510+)
                amts_sorted = sorted(
                    [w for w in dk_line if w[0] >= 398 and w[4] not in ('-', '(', ')')],
                    key=lambda w: w[0]
                )
                wd_val = None
                dp_val = None
                bal_val = None
                for w in amts_sorted:
                    x = w[0]
                    if x < 456:
                        wd_val = w[4]
                    elif x < 510:
                        dp_val = w[4]
                    else:
                        bal_val = w[4]

                # Scan backwards from dk to find nearest date line
                date_text = ""
                val_date_text = ""
                for gk in reversed(sorted_ys):
                    if gk > dk:
                        continue
                    if gk in line_date:
                        date_text = line_date[gk]
                        # Collect narration between this date and detail
                        narration_lines = []
                        for ck in sorted_ys:
                            if ck <= gk or ck >= dk:
                                continue
                            if ck in detail_ys:
                                continue
                            cline = sorted(grouped_lines[ck], key=lambda w: w[0])
                            for w in cline:
                                if 112 <= w[0] < 278:
                                    narration_lines.append(w[4])
                        # Also collect from the date line itself
                        dline = sorted(grouped_lines[gk], key=lambda w: w[0])
                        for w in dline:
                            if 112 <= w[0] < 278:
                                narration_lines.append(w[4])
                        for vk in sorted_ys:
                            if vk <= gk or vk >= dk:
                                continue
                            if vk in line_vd:
                                val_date_text = line_vd[vk]
                        break

                if date_text:
                    transactions.append({
                        'date': date_text,
                        'narration': " ".join(narration_lines).strip(),
                        'chq_ref': " ".join(ref_parts).strip(),
                        'val_date': val_date_text if val_date_text else date_text,
                        'withdrawal': self.clean_amount(wd_val),
                        'deposit': self.clean_amount(dp_val),
                        'balance': self.clean_amount(bal_val),
                        'page': page_num,
                    })
        # IOB: each page is newest-first. Pages go 1 (newest) → 22 (oldest).
        # Process pages oldest-first (22→1), reversing each to chronological.
        page_groups = {}
        for tx in transactions:
            page_groups.setdefault(tx['page'], []).append(tx)
        result = []
        for p in sorted(page_groups.keys(), reverse=True):
            result.extend(reversed(page_groups[p]))
        print(f"Extraction complete. Found {len(result)} raw transactions.")
        return result

    def parse_iob_pdfplumber(self):
        print(f"Opening PDF (pdfplumber): {self.pdf_path}")
        transactions = []
        date_pattern = re.compile(r'^\d{2}-[A-Z][a-z]{2}-\d{2}$')
        ref_pat = re.compile(r'^S\d+')

        with pdfplumber.open(self.pdf_path) as pdf:
            total_pages = len(pdf.pages)
            print(f"Total pages to parse (IOB): {total_pages}")

            for page_idx in range(total_pages):
                page = pdf.pages[page_idx]
                page_num = page_idx + 1

                if self.verbose or page_num % 10 == 0 or page_num == total_pages:
                    print(f"Processing page {page_num}/{total_pages}...")

                words = page.extract_words()
                table_words = [w for w in words if 30 <= w['top'] <= 790]
                table_words = [w for w in table_words if not w['text'].startswith('Page ')]
                if not table_words:
                    continue

                grouped_lines = {}
                for w in table_words:
                    top_val = w['top']
                    found = False
                    for gk in grouped_lines.keys():
                        if abs(gk - top_val) < 2.0:
                            grouped_lines[gk].append(w)
                            found = True
                            break
                    if not found:
                        grouped_lines[top_val] = [w]

                sorted_ys = sorted(grouped_lines.keys())

                # Find detail lines
                detail_ys = []
                for gk in sorted_ys:
                    line_words = sorted(grouped_lines[gk], key=lambda w: w['x0'])
                    if any(278 <= w['x0'] < 344 and ref_pat.match(w['text']) for w in line_words):
                        detail_ys.append(gk)

                if not detail_ys:
                    continue

                # Pre-scan dates
                line_date = {}
                line_vd = {}
                for gk in sorted_ys:
                    line_words = sorted(grouped_lines[gk], key=lambda w: w['x0'])
                    for w in line_words:
                        x, text = w['x0'], w['text']
                        if 40 <= x < 90:
                            m = re.match(r'^\((\d{2}-[A-Z][a-z]{2}-\d{2})\)$', text)
                            if m:
                                line_vd[gk] = m.group(1)
                                continue
                            if date_pattern.match(text) and gk not in line_date:
                                line_date[gk] = text

                for dk in detail_ys:
                    dk_line = sorted(grouped_lines[dk], key=lambda w: w['x0'])
                    ref_parts = [w['text'] for w in dk_line if 278 <= w['x0'] < 344 and ref_pat.match(w['text'])]

                    amts_sorted = sorted(
                        [w for w in dk_line if w['x0'] >= 398 and w['text'] not in ('-', '(', ')')],
                        key=lambda w: w['x0']
                    )
                    wd_val = dp_val = bal_val = None
                    for w in amts_sorted:
                        x = w['x0']
                        if x < 456:
                            wd_val = w['text']
                        elif x < 510:
                            dp_val = w['text']
                        else:
                            bal_val = w['text']

                    date_text = ""
                    val_date_text = ""
                    for gk in reversed(sorted_ys):
                        if gk > dk:
                            continue
                        if gk in line_date:
                            date_text = line_date[gk]
                            narration_lines = []
                            for ck in sorted_ys:
                                if ck <= gk or ck >= dk:
                                    continue
                                if ck in detail_ys:
                                    continue
                                cline = sorted(grouped_lines[ck], key=lambda w: w['x0'])
                                for w in cline:
                                    if 112 <= w['x0'] < 278:
                                        narration_lines.append(w['text'])
                            dline = sorted(grouped_lines[gk], key=lambda w: w['x0'])
                            for w in dline:
                                if 112 <= w['x0'] < 278:
                                    narration_lines.append(w['text'])
                            for vk in sorted_ys:
                                if vk <= gk or vk >= dk:
                                    continue
                                if vk in line_vd:
                                    val_date_text = line_vd[vk]
                            break

                    if date_text:
                        transactions.append({
                            'date': date_text,
                            'narration': " ".join(narration_lines).strip(),
                            'chq_ref': " ".join(ref_parts).strip(),
                            'val_date': val_date_text if val_date_text else date_text,
                            'withdrawal': self.clean_amount(wd_val),
                            'deposit': self.clean_amount(dp_val),
                            'balance': self.clean_amount(bal_val),
                            'page': page_num,
                        })

        # IOB: per-page newest-first -> oldest-first chronological
        page_groups = {}
        for tx in transactions:
            page_groups.setdefault(tx['page'], []).append(tx)
        result = []
        for p in sorted(page_groups.keys(), reverse=True):
            result.extend(reversed(page_groups[p]))
        print(f"Extraction complete. Found {len(result)} raw transactions.")
        return result

    # ═══════════════════════════════════════════════════════════════════
    #  Post-processing & output (shared by both banks)
    # ═══════════════════════════════════════════════════════════════════

    def categorize(self, narration):
        """Categorizes a transaction based on narration keywords."""
        text = narration.lower()
        if any(w in text for w in ['salary', 'payroll', 'salary credit', 'betterplace']):
            return 'Salary'
        elif any(w in text for w in ['foreign', 'usd', 'eur', 'inw', 'remittance', 'forex']):
            return 'Foreign Exchange'
        elif any(w in text for w in ['upi-']):
            return 'UPI'
        elif any(w in text for w in ['imps-']):
            return 'IMPS'
        elif any(w in text for w in ['neft-']):
            return 'NEFT'
        elif any(w in text for w in ['rtgs-']):
            return 'RTGS'
        elif any(w in text for w in ['atm-', 'atm wdl']):
            return 'ATM Withdrawal'
        elif any(w in text for w in ['card-', 'pos-', 'pos wdl']):
            return 'Card Payment'
        elif any(w in text for w in ['chq-', 'cheque', 'clg-']):
            return 'Cheque'
        elif any(w in text for w in ['interest credit', 'int.coll', 'interest paid']):
            return 'Interest Income'
        elif any(w in text for w in ['refund', 'cashback']):
            return 'Refund/Cashback'
        elif any(w in text for w in ['charge', 'fee', 'gst-', 'tax', 'annual maint']):
            return 'Bank Charges'
        elif any(w in text for w in ['sweep', 'autosweep', 'mod ']):
            return 'Sweep/MOD'
        elif any(w in text for w in ['mutual fund', 'zerodha', 'groww', 'indmoney', 'icici direct']):
            return 'Investment'
        elif any(w in text for w in ['loan', 'emi-']):
            return 'Loan EMI'
        elif any(w in text for w in ['insurance', 'lic ']):
            return 'Insurance'
        else:
            return 'Other Transfer/Spending'

    def validate_and_process(self, transactions):
        """Validates math equations and adds categorizations."""
        processed = []
        validation_warnings = 0

        print("Validating transaction math integrity...")
        for i, tx in enumerate(transactions):
            tx['category'] = self.categorize(tx['narration'])

            if i > 0:
                prev_balance = processed[i - 1]['balance']
                curr_balance = tx['balance']
                w_amt = tx['withdrawal'] or 0.0
                d_amt = tx['deposit'] or 0.0
                expected_balance = prev_balance - w_amt + d_amt

                if abs(expected_balance - curr_balance) > 0.02:
                    validation_warnings += 1
                    if validation_warnings <= 5:
                        print(f"Warning: Math discrepancy at index {i} (Page {tx['page']}):")
                        print(f"  Prev Balance:     {prev_balance}")
                        print(f"  Withdrawal (-):   {w_amt}")
                        print(f"  Deposit (+):      {d_amt}")
                        print(f"  Expected Balance: {expected_balance:.2f}")
                        print(f"  Parsed Balance:   {curr_balance}")
                        print(f"  Narration:        {tx['narration'][:60]}...")

            processed.append(tx)

        if validation_warnings > 0:
            print(f"Math Validation: WARNING: Found {validation_warnings} discrepancies out of {len(transactions)} transactions.")
        else:
            print("Math Validation: SUCCESS: All transaction balances are mathematically consistent!")

        return processed

    def save_csv(self, transactions, output_path):
        print(f"Saving to CSV: {output_path}")
        with open(output_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Date', 'Narration', 'Chq/Ref No', 'Value Date', 'Withdrawal', 'Deposit', 'Balance', 'Category', 'Page'])
            for tx in transactions:
                writer.writerow([
                    tx['date'],
                    tx['narration'],
                    tx['chq_ref'],
                    tx['val_date'],
                    self.format_amount(tx['withdrawal']),
                    self.format_amount(tx['deposit']),
                    self.format_amount(tx['balance']),
                    tx['category'],
                    tx['page'],
                ])
        print(f"CSV file saved successfully: {os.path.abspath(output_path)}")

    def save_excel(self, transactions, output_path):
        if not HAS_OPENPYXL:
            print("Excel export skipped: openpyxl is not installed.")
            return False

        print(f"Saving to Excel: {output_path}")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Transactions"

        headers = ['Date', 'Narration', 'Chq/Ref No', 'Value Date', 'Withdrawal', 'Deposit', 'Balance', 'Category', 'Page']
        ws.append(headers)

        for tx in transactions:
            ws.append([
                tx['date'],
                tx['narration'],
                tx['chq_ref'],
                tx['val_date'],
                tx['withdrawal'] if tx['withdrawal'] is not None else "",
                tx['deposit'] if tx['deposit'] is not None else "",
                tx['balance'] if tx['balance'] is not None else "",
                tx['category'],
                tx['page'],
            ])

        from openpyxl.styles import Font, Alignment, PatternFill
        from openpyxl.utils import get_column_letter

        header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

        for row in range(2, len(transactions) + 2):
            for col in [5, 6, 7]:
                cell = ws.cell(row=row, column=col)
                if cell.value != "":
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal="right")
            ws.cell(row=row, column=1).alignment = Alignment(horizontal="center")
            ws.cell(row=row, column=4).alignment = Alignment(horizontal="center")

        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                if cell.column == 2:
                    max_len = min(max(max_len, len(val_str)), 60)
                else:
                    max_len = max(max_len, len(val_str))
            ws.column_dimensions[col_letter].width = max(max_len + 3, 10)

        wb.save(output_path)
        print(f"Excel file saved successfully: {os.path.abspath(output_path)}")
        return True


def main():
    parser = argparse.ArgumentParser(description="Convert Bank PDF Statement (HDFC/SBI/IOB) to CSV/Excel")
    parser.add_argument("pdf_path", help="Path to Bank Statement PDF file")
    parser.add_argument("--csv", help="Output path for CSV file (default: pdf_name_parsed.csv)")
    parser.add_argument("--xlsx", help="Output path for Excel file (default: pdf_name_parsed.xlsx)")
    parser.add_argument("--bank", choices=['hdfc', 'sbi', 'iob'], help="Bank type (auto-detected by default)")
    parser.add_argument("--verbose", action="store_true", help="Enable verbose log printing")

    args = parser.parse_args()

    pdf_base = os.path.splitext(args.pdf_path)[0]
    csv_path = args.csv or f"{pdf_base}_parsed.csv"
    xlsx_path = args.xlsx or f"{pdf_base}_parsed.xlsx"

    try:
        start_time = datetime.now()
        parser_inst = BankStatementParser(args.pdf_path, verbose=args.verbose, bank=args.bank)

        raw_txs = parser_inst.parse()
        processed_txs = parser_inst.validate_and_process(raw_txs)
        parser_inst.save_csv(processed_txs, csv_path)

        if HAS_OPENPYXL:
            parser_inst.save_excel(processed_txs, xlsx_path)
        else:
            print("Note: Install 'openpyxl' (pip install openpyxl) to enable formatting Excel .xlsx output.")

        end_time = datetime.now()
        elapsed = (end_time - start_time).total_seconds()

        total_wd = sum(tx['withdrawal'] or 0.0 for tx in processed_txs)
        total_dp = sum(tx['deposit'] or 0.0 for tx in processed_txs)

        print("\n" + "=" * 40)
        print("Parsing Summary")
        print("=" * 40)
        print(f"Processed:          {len(processed_txs)} transactions")
        print(f"Total Withdrawals:  INR {total_wd:,.2f}")
        print(f"Total Deposits:     INR {total_dp:,.2f}")
        print(f"Net Change:         INR {(total_dp - total_wd):,.2f}")
        if processed_txs:
            first_bal = processed_txs[0]['balance']
            first_dep = processed_txs[0]['deposit'] or 0
            first_wd = processed_txs[0]['withdrawal'] or 0
            print(f"Starting Balance:   INR {first_bal - first_dep + first_wd:,.2f}")
            print(f"Ending Balance:     INR {processed_txs[-1]['balance']:,.2f}")
        print(f"Time Elapsed:       {elapsed:.1f} seconds")
        print("=" * 40)

    except Exception as e:
        print(f"Error during parsing: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
