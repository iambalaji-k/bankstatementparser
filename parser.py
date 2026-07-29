#!/usr/bin/env python3
"""
Multi-Bank PDF Statement Parser
Converts HDFC/SBI/IOB/Canara/Indian Bank statement PDFs to CSV/Excel using coordinate-based layout parsing.
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
    """Unified parser for HDFC, SBI, IOB, and Canara bank statement PDFs with auto-detection."""

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
            if "canara bank" in text_lower or "e-pass sheet" in text_lower:
                return "canara"
            if "indian bank" in text_lower or "indianbank" in text_lower or "idib" in text_lower:
                return "indianbank"
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
        elif self.bank == "canara":
            print(f"Detected bank: Canara Bank")
            if HAS_FITZ:
                try:
                    return self.parse_canara_fitz()
                except Exception as e:
                    print(f"PyMuPDF parsing failed: {e}. Falling back to pdfplumber...")
            return self.parse_canara_pdfplumber()
        elif self.bank == "indianbank":
            print(f"Detected bank: Indian Bank")
            if HAS_FITZ:
                try:
                    return self.parse_indianbank_fitz()
                except Exception as e:
                    print(f"PyMuPDF parsing failed: {e}. Falling back to pdfplumber...")
            return self.parse_indianbank_pdfplumber()
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
            min_y = 180 if page_idx == 0 else 40
            table_words = [w for w in words if min_y <= w[1] <= 730]

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

            # Filter out summary/footer lines
            valid_ys = []
            for gk in sorted_ys:
                line_text = " ".join(w[4] for w in sorted(grouped_lines[gk], key=lambda w: w[0]))
                if any(kw in line_text for kw in [
                    'Statement Summary', 'Total:', 'Closing Balance',
                    'Page ', 'This is a computer'
                ]):
                    break
                valid_ys.append(gk)

            date_ys = []
            for gk in valid_ys:
                line_words = sorted(grouped_lines[gk], key=lambda w: w[0])
                dt_text = " ".join(w[4] for w in line_words if 20 <= w[0] < 78).strip()
                if date_pattern.match(dt_text):
                    date_ys.append((gk, dt_text))

            if not date_ys:
                continue

            for i, (dy, dt_str) in enumerate(date_ys):
                block_end = date_ys[i + 1][0] - 10 if i + 1 < len(date_ys) else (valid_ys[-1] + 1)
                block_ys = [gk for gk in valid_ys if (dy - 10) <= gk < block_end]

                narration_parts = []
                val_date_text = dt_str
                chq_ref_parts = []
                debit_val = None
                credit_val = None
                balance_val = None

                for gk in block_ys:
                    line_words = sorted(grouped_lines[gk], key=lambda w: w[0])

                    if gk == dy:
                        vd_text = " ".join(w[4] for w in line_words if 78 <= w[0] < 130).strip()
                        if date_pattern.match(vd_text):
                            val_date_text = vd_text

                    n_words = [w[4] for w in line_words if 130 <= w[0] < 290]
                    if n_words:
                        narration_parts.append(" ".join(n_words))

                    ref_words = [w[4] for w in line_words if 290 <= w[0] < 335 and w[4] != '-']
                    if ref_words:
                        chq_ref_parts.extend(ref_words)

                    for w in line_words:
                        x = w[0]
                        val_str = w[4].replace(',', '')
                        if val_str == '-':
                            continue
                        try:
                            fval = float(val_str)
                        except ValueError:
                            continue
                        if 335 <= x < 410:
                            debit_val = fval
                        elif 410 <= x < 485:
                            credit_val = fval
                        elif 485 <= x < 580:
                            balance_val = fval

                transactions.append({
                    'date': dt_str,
                    'narration': " ".join(narration_parts).strip(),
                    'chq_ref': " ".join(chq_ref_parts).strip(),
                    'val_date': val_date_text,
                    'withdrawal': debit_val,
                    'deposit': credit_val,
                    'balance': balance_val,
                    'page': page_num,
                })

        doc.close()
        print(f"Extraction complete. Found {len(transactions)} raw transactions.")
        return transactions

    def parse_sbi_pdfplumber(self):
        print(f"Opening PDF (pdfplumber): {self.pdf_path}")
        transactions = []
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
                min_y = 180 if page_idx == 0 else 40
                table_words = [w for w in words if min_y <= w['top'] <= 730]

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

                valid_ys = []
                for gk in sorted_ys:
                    line_text = " ".join(w['text'] for w in sorted(grouped_lines[gk], key=lambda w: w['x0']))
                    if any(kw in line_text for kw in [
                        'Statement Summary', 'Total:', 'Closing Balance',
                        'Page ', 'This is a computer'
                    ]):
                        break
                    valid_ys.append(gk)

                date_ys = []
                for gk in valid_ys:
                    line_words = sorted(grouped_lines[gk], key=lambda w: w['x0'])
                    dt_text = " ".join(w['text'] for w in line_words if 20 <= w['x0'] < 78).strip()
                    if date_pattern.match(dt_text):
                        date_ys.append((gk, dt_text))

                if not date_ys:
                    continue

                for i, (dy, dt_str) in enumerate(date_ys):
                    block_end = date_ys[i + 1][0] - 10 if i + 1 < len(date_ys) else (valid_ys[-1] + 1)
                    block_ys = [gk for gk in valid_ys if (dy - 10) <= gk < block_end]

                    narration_parts = []
                    val_date_text = dt_str
                    chq_ref_parts = []
                    debit_val = None
                    credit_val = None
                    balance_val = None

                    for gk in block_ys:
                        line_words = sorted(grouped_lines[gk], key=lambda w: w['x0'])

                        if gk == dy:
                            vd_text = " ".join(w['text'] for w in line_words if 78 <= w['x0'] < 130).strip()
                            if date_pattern.match(vd_text):
                                val_date_text = vd_text

                        n_words = [w['text'] for w in line_words if 130 <= w['x0'] < 290]
                        if n_words:
                            narration_parts.append(" ".join(n_words))

                        ref_words = [w['text'] for w in line_words if 290 <= w['x0'] < 335 and w['text'] != '-']
                        if ref_words:
                            chq_ref_parts.extend(ref_words)

                        for w in line_words:
                            x = w['x0']
                            val_str = w['text'].replace(',', '')
                            if val_str == '-':
                                continue
                            try:
                                fval = float(val_str)
                            except ValueError:
                                continue
                            if 335 <= x < 410:
                                debit_val = fval
                            elif 410 <= x < 485:
                                credit_val = fval
                            elif 485 <= x < 580:
                                balance_val = fval

                    transactions.append({
                        'date': dt_str,
                        'narration': " ".join(narration_parts).strip(),
                        'chq_ref': " ".join(chq_ref_parts).strip(),
                        'val_date': val_date_text,
                        'withdrawal': debit_val,
                        'deposit': credit_val,
                        'balance': balance_val,
                        'page': page_num,
                    })

        print(f"Extraction complete. Found {len(transactions)} raw transactions.")
        return transactions

    # ═══════════════════════════════════════════════════════════════════
    #  IOB Parsing (Indian Overseas Bank)
    #  Reverse chronological, 2-line blocks, DD-Mon-YY dates
    # ═══════════════════════════════════════════════════════════════════

    def parse_iob_fitz(self):
        print(f"Opening PDF (PyMuPDF): {self.pdf_path}")
        raw_txs = []
        date_pattern = re.compile(r'^\d{2}-[A-Z][a-z]{2}-\d{2}$')
        vdate_pattern = re.compile(r'^\((\d{2}-[A-Z][a-z]{2}-\d{2})\)$')

        doc = fitz.open(self.pdf_path)
        total_pages = len(doc)
        print(f"Total pages to parse (IOB): {total_pages}")

        for page_idx in range(total_pages):
            page = doc[page_idx]
            page_num = page_idx + 1

            if self.verbose or page_num % 10 == 0 or page_num == total_pages:
                print(f"Processing page {page_num}/{total_pages}...")

            words = page.get_text("words")
            min_y = 260 if page_idx == 0 else 30
            table_words = [w for w in words if min_y <= w[1] <= 790 and not w[4].startswith('Page ')]

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

            # Filter out header and summary/footer lines
            valid_ys = []
            for gk in sorted_ys:
                line_text = " ".join(w[4] for w in sorted(grouped_lines[gk], key=lambda w: w[0]))
                if any(kw in line_text for kw in [
                    'STATEMENT OF THE ACCOUNT', 'CUSTOMER DETAILS',
                    'Particulars', 'Ref No.', 'Debit(Rs)',
                    'Effective available balance', 'computer generated statement', 'Page '
                ]):
                    continue
                valid_ys.append(gk)

            # Identify transaction lines via Balance column (x >= 510)
            bal_ys = []
            for gk in valid_ys:
                line_words = sorted(grouped_lines[gk], key=lambda w: w[0])
                for w in line_words:
                    if w[0] >= 510 and w[4] not in ('-', '(', ')'):
                        try:
                            fval = float(w[4].replace(',', ''))
                            bal_ys.append((gk, fval))
                            break
                        except ValueError:
                            pass

            for i, (by, bal_val) in enumerate(bal_ys):
                next_by = bal_ys[i + 1][0] if i + 1 < len(bal_ys) else None
                block_start = by - 6
                block_end = (next_by - 6) if next_by is not None else (by + 15)
                block_ys = [gk for gk in valid_ys if block_start <= gk < block_end]

                date_text = ""
                val_date_text = ""
                narration_parts = []
                ref_parts = []
                debit_val = None
                credit_val = None

                for gk in block_ys:
                    line_words = sorted(grouped_lines[gk], key=lambda w: w[0])

                    for w in line_words:
                        if 40 <= w[0] < 90:
                            vm = vdate_pattern.match(w[4])
                            if vm:
                                val_date_text = vm.group(1)
                            elif date_pattern.match(w[4]) and not date_text:
                                date_text = w[4]

                    n_words = [w[4] for w in line_words if 90 <= w[0] < 275]
                    if n_words:
                        narration_parts.append(" ".join(n_words))

                    r_words = [w[4] for w in line_words if 275 <= w[0] < 340 and w[4] != '-']
                    if r_words:
                        ref_parts.extend(r_words)

                    for w in line_words:
                        x = w[0]
                        val_str = w[4].replace(',', '')
                        if val_str in ('-', '(', ')'):
                            continue
                        try:
                            fval = float(val_str)
                        except ValueError:
                            continue
                        if 395 <= x < 455:
                            debit_val = fval
                        elif 455 <= x < 510:
                            credit_val = fval

                raw_txs.append({
                    'date': date_text,
                    'narration': " ".join(narration_parts).strip(),
                    'chq_ref': " ".join(ref_parts).strip(),
                    'val_date': val_date_text if val_date_text else date_text,
                    'withdrawal': debit_val,
                    'deposit': credit_val,
                    'balance': bal_val,
                    'page': page_num,
                })

        doc.close()

        # Restore chronological order (pages N -> 1, items reversed)
        page_groups = {}
        for tx in raw_txs:
            page_groups.setdefault(tx['page'], []).append(tx)
        result = []
        for p in sorted(page_groups.keys(), reverse=True):
            result.extend(reversed(page_groups[p]))

        print(f"Extraction complete. Found {len(result)} raw transactions.")
        return result

    def parse_iob_pdfplumber(self):
        print(f"Opening PDF (pdfplumber): {self.pdf_path}")
        raw_txs = []
        date_pattern = re.compile(r'^\d{2}-[A-Z][a-z]{2}-\d{2}$')
        vdate_pattern = re.compile(r'^\((\d{2}-[A-Z][a-z]{2}-\d{2})\)$')

        with pdfplumber.open(self.pdf_path) as pdf:
            total_pages = len(pdf.pages)
            print(f"Total pages to parse (IOB): {total_pages}")

            for page_idx in range(total_pages):
                page = pdf.pages[page_idx]
                page_num = page_idx + 1

                if self.verbose or page_num % 10 == 0 or page_num == total_pages:
                    print(f"Processing page {page_num}/{total_pages}...")

                words = page.extract_words()
                min_y = 260 if page_idx == 0 else 30
                table_words = [w for w in words if min_y <= w['top'] <= 790 and not w['text'].startswith('Page ')]

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

                valid_ys = []
                for gk in sorted_ys:
                    line_text = " ".join(w['text'] for w in sorted(grouped_lines[gk], key=lambda w: w['x0']))
                    if any(kw in line_text for kw in [
                        'STATEMENT OF THE ACCOUNT', 'CUSTOMER DETAILS',
                        'Particulars', 'Ref No.', 'Debit(Rs)',
                        'Effective available balance', 'computer generated statement', 'Page '
                    ]):
                        continue
                    valid_ys.append(gk)

                bal_ys = []
                for gk in valid_ys:
                    line_words = sorted(grouped_lines[gk], key=lambda w: w['x0'])
                    for w in line_words:
                        if w['x0'] >= 510 and w['text'] not in ('-', '(', ')'):
                            try:
                                fval = float(w['text'].replace(',', ''))
                                bal_ys.append((gk, fval))
                                break
                            except ValueError:
                                pass

                for i, (by, bal_val) in enumerate(bal_ys):
                    next_by = bal_ys[i + 1][0] if i + 1 < len(bal_ys) else None
                    block_start = by - 6
                    block_end = (next_by - 6) if next_by is not None else (by + 15)
                    block_ys = [gk for gk in valid_ys if block_start <= gk < block_end]

                    date_text = ""
                    val_date_text = ""
                    narration_parts = []
                    ref_parts = []
                    debit_val = None
                    credit_val = None

                    for gk in block_ys:
                        line_words = sorted(grouped_lines[gk], key=lambda w: w['x0'])

                        for w in line_words:
                            if 40 <= w['x0'] < 90:
                                vm = vdate_pattern.match(w['text'])
                                if vm:
                                    val_date_text = vm.group(1)
                                elif date_pattern.match(w['text']) and not date_text:
                                    date_text = w['text']

                        n_words = [w['text'] for w in line_words if 90 <= w['x0'] < 275]
                        if n_words:
                            narration_parts.append(" ".join(n_words))

                        ref_words = [w['text'] for w in line_words if 275 <= w['x0'] < 340 and w['text'] != '-']
                        if ref_words:
                            ref_parts.extend(ref_words)

                        for w in line_words:
                            x = w['x0']
                            val_str = w['text'].replace(',', '')
                            if val_str in ('-', '(', ')'):
                                continue
                            try:
                                fval = float(val_str)
                            except ValueError:
                                continue
                            if 395 <= x < 455:
                                debit_val = fval
                            elif 455 <= x < 510:
                                credit_val = fval

                    raw_txs.append({
                        'date': date_text,
                        'narration': " ".join(narration_parts).strip(),
                        'chq_ref': " ".join(ref_parts).strip(),
                        'val_date': val_date_text if val_date_text else date_text,
                        'withdrawal': debit_val,
                        'deposit': credit_val,
                        'balance': bal_val,
                        'page': page_num,
                    })

        # Restore chronological order
        page_groups = {}
        for tx in raw_txs:
            page_groups.setdefault(tx['page'], []).append(tx)
        result = []
        for p in sorted(page_groups.keys(), reverse=True):
            result.extend(reversed(page_groups[p]))

        print(f"Extraction complete. Found {len(result)} raw transactions.")
        return result

    # ═══════════════════════════════════════════════════════════════════
    #  Canara Bank Parsing (e-Pass Sheet)
    #  Chronological order, DD-Mon-YY dates, block-based layout
    # ═══════════════════════════════════════════════════════════════════

    def parse_canara_fitz(self):
        print(f"Opening PDF (PyMuPDF): {self.pdf_path}")
        transactions = []
        date_pattern = re.compile(r'^\d{2}-[A-Z][a-z]{2}-\d{2}$')

        doc = fitz.open(self.pdf_path)
        total_pages = len(doc)
        print(f"Total pages to parse (Canara): {total_pages}")

        for page_idx in range(total_pages):
            page = doc[page_idx]
            page_num = page_idx + 1

            if self.verbose or page_num % 10 == 0 or page_num == total_pages:
                print(f"Processing page {page_num}/{total_pages}...")

            words = page.get_text("words")
            min_y = 240 if page_idx == 0 else 10
            table_words = [w for w in words if w[1] >= min_y]

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

            # Filter out summary/footer lines
            valid_ys = []
            for gk in sorted_ys:
                line_text = " ".join(w[4] for w in sorted(grouped_lines[gk], key=lambda w: w[0]))
                if any(kw in line_text for kw in [
                    'Total:', 'Opening Balance', 'Closing Balance',
                    'Dr Count', 'Cr Count', 'UNLESS THE',
                    'COMPUTER OUTPUT', 'End of Statement'
                ]):
                    break
                valid_ys.append(gk)

            date_ys = []
            for gk in valid_ys:
                line_words = sorted(grouped_lines[gk], key=lambda w: w[0])
                date_text = " ".join(w[4] for w in line_words if 30 <= w[0] <= 90).strip()
                if date_pattern.match(date_text):
                    date_ys.append((gk, date_text))

            if not date_ys:
                continue

            for i, (dy, date_text) in enumerate(date_ys):
                block_end = date_ys[i + 1][0] if i + 1 < len(date_ys) else (valid_ys[-1] + 1)
                block_ys = [gk for gk in valid_ys if dy <= gk < block_end]

                narration_parts = []
                debit_val = None
                credit_val = None
                balance_val = None

                for gk in block_ys:
                    line_words = sorted(grouped_lines[gk], key=lambda w: w[0])

                    n_words = [w[4] for w in line_words if 90 <= w[0] < 310]
                    if n_words:
                        narration_parts.append(" ".join(n_words))

                    for w in line_words:
                        x = w[0]
                        val_str = w[4].replace(',', '')
                        try:
                            fval = float(val_str)
                        except ValueError:
                            continue
                        if 310 <= x < 424:
                            debit_val = fval
                        elif 424 <= x < 515:
                            credit_val = fval
                        elif x >= 515:
                            balance_val = fval

                transactions.append({
                    'date': date_text,
                    'narration': " ".join(narration_parts).strip(),
                    'chq_ref': '',
                    'val_date': date_text,
                    'withdrawal': debit_val,
                    'deposit': credit_val,
                    'balance': balance_val,
                    'page': page_num,
                })

        doc.close()
        print(f"Extraction complete. Found {len(transactions)} raw transactions.")
        return transactions

    def parse_canara_pdfplumber(self):
        print(f"Opening PDF (pdfplumber): {self.pdf_path}")
        transactions = []
        date_pattern = re.compile(r'^\d{2}-[A-Z][a-z]{2}-\d{2}$')

        with pdfplumber.open(self.pdf_path) as pdf:
            total_pages = len(pdf.pages)
            print(f"Total pages to parse (Canara): {total_pages}")

            for page_idx in range(total_pages):
                page = pdf.pages[page_idx]
                page_num = page_idx + 1

                if self.verbose or page_num % 10 == 0 or page_num == total_pages:
                    print(f"Processing page {page_num}/{total_pages}...")

                words = page.extract_words()
                # Filter to table region
                table_words = [w for w in words if w['top'] >= 230]

                # Group by y-coordinate
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

                # Find date lines
                date_ys = []
                for gk in sorted_ys:
                    line_words = sorted(grouped_lines[gk], key=lambda w: w['x0'])
                    # Date is at x 30-80
                    date_text = " ".join(w['text'] for w in line_words if 30 <= w['x0'] <= 80).strip()
                    if date_pattern.match(date_text):
                        date_ys.append(gk)

                if not date_ys:
                    continue  # summary/footer page

                # For each date, collect narration and amounts
                for i, dy in enumerate(date_ys):
                    block_end = date_ys[i + 1] if i + 1 < len(date_ys) else (sorted_ys[-1] + 1)

                    # Collect all lines between this date and the next
                    block_ys = [gk for gk in sorted_ys if dy <= gk < block_end]

                    # Date text from the date line
                    date_line = sorted(grouped_lines[dy], key=lambda w: w['x0'])
                    date_text = " ".join(w['text'] for w in date_line if 30 <= w['x0'] <= 80).strip()

                    narration_parts = []
                    debit_val = None
                    credit_val = None
                    balance_val = None

                    for gk in block_ys:
                        line_words = sorted(grouped_lines[gk], key=lambda w: w['x0'])
                        line_text = " ".join(w['text'] for w in line_words).strip()

                        # Skip header rows
                        if line_text in ('Txn Date', 'Txn Description', 'Debit', 'Credit', 'Balance'):
                            continue
                        if line_text.startswith('OPENING_BALANCE') or \
                           any(kw in line_text for kw in ['Total:', 'Opening Balance', 'Closing Balance',
                                                           'Dr Count', 'Cr Count', 'Dr amount', 'Cr amount',
                                                           'UNLESS THE', 'COMPUTER OUTPUT', 'End of Statement']):
                            continue

                        # Check if this is an amount line (has values at x >= 310)
                        amt_words = [w for w in line_words if w['x0'] >= 310]
                        if amt_words:
                            for w in amt_words:
                                x = w['x0']
                                try:
                                    fval = float(w['text'].replace(',', ''))
                                except ValueError:
                                    continue
                                if 320 <= x < 424:
                                    debit_val = fval
                                elif 424 <= x < 515:
                                    credit_val = fval
                                elif x >= 515:
                                    balance_val = fval
                        else:
                            # Narration line
                            n_words = [w['text'] for w in line_words if 100 <= w['x0'] < 160]
                            if n_words:
                                narration_parts.append(" ".join(n_words).strip())

                    if date_text:
                        transactions.append({
                            'date': date_text,
                            'narration': " ".join(narration_parts).strip(),
                            'chq_ref': '',
                            'val_date': date_text,
                            'withdrawal': debit_val,
                            'deposit': credit_val,
                            'balance': balance_val,
                            'page': page_num,
                        })

        print(f"Extraction complete. Found {len(transactions)} raw transactions.")
        return transactions

    # ═══════════════════════════════════════════════════════════════════
    #  Indian Bank Parsing
    #  Chronological order, DD-Mon-YYYY dates, INR-prefixed amounts
    # ═══════════════════════════════════════════════════════════════════

    def parse_indianbank_fitz(self):
        print(f"Opening PDF (PyMuPDF): {self.pdf_path}")
        transactions = []
        # Indian Bank dates: "02 May 2026" as 3 separate words
        date_pattern = re.compile(r'^\d{2}$')
        month_pattern = re.compile(r'^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)$')
        year_pattern = re.compile(r'^\d{4}$')

        doc = fitz.open(self.pdf_path)
        total_pages = len(doc)
        print(f"Total pages to parse (Indian Bank): {total_pages}")

        for page_idx in range(total_pages):
            page = doc[page_idx]
            page_num = page_idx + 1

            if self.verbose or page_num % 10 == 0 or page_num == total_pages:
                print(f"Processing page {page_num}/{total_pages}...")

            words = page.get_text("words")
            # Indian Bank table region: y 65-780 (excludes page headers & footers)
            table_words = [w for w in words if 65 <= w[1] <= 780]

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

            # Find date lines (3 consecutive words: DD MMM YYYY)
            date_ys = []
            for gk in sorted_ys:
                line_words = sorted(grouped_lines[gk], key=lambda w: w[0])
                # Check for date pattern at x=71.5-123.5
                date_words = [w for w in line_words if 71 <= w[0] <= 124]
                if len(date_words) >= 3:
                    day_text = date_words[0][4]
                    month_text = date_words[1][4]
                    year_text = date_words[2][4]
                    if (date_pattern.match(day_text) and 
                        month_pattern.match(month_text) and 
                        year_pattern.match(year_text)):
                        date_ys.append((gk, f"{day_text} {month_text} {year_text}"))

            if not date_ys:
                continue  # skip summary/header pages

            # Process each transaction
            for i, (dy, date_text) in enumerate(date_ys):
                # Determine block boundaries
                block_end = date_ys[i + 1][0] if i + 1 < len(date_ys) else (sorted_ys[-1] + 1)
                block_ys = [gk for gk in sorted_ys if dy <= gk < block_end]

                # Collect all words in this transaction block
                narration_parts = []
                debit_val = None
                credit_val = None
                balance_val = None

                for gk in block_ys:
                    line_words = sorted(grouped_lines[gk], key=lambda w: w[0])

                    # Extract narration (x=145-260)
                    n_words = [w[4] for w in line_words if 145 <= w[0] <= 260]
                    if n_words:
                        narration_parts.append(" ".join(n_words))

                    # Check for debit indicator at x=306 (means credit is empty)
                    debit_indicator = [w for w in line_words if 305 <= w[0] <= 308 and w[4] == '-']
                    # Check for credit indicator at x=401 (means debit is empty)
                    credit_indicator = [w for w in line_words if 400 <= w[0] <= 403 and w[4] == '-']

                    # Extract debit amount (x=283-330) - only if no debit indicator
                    if not debit_indicator:
                        debit_words = [w for w in line_words if 283 <= w[0] <= 330 and w[4] != 'INR']
                        if debit_words:
                            try:
                                debit_val = float(debit_words[0][4].replace(',', ''))
                            except ValueError:
                                pass

                    # Extract credit amount (x=378-425) - only if no credit indicator
                    if not credit_indicator:
                        credit_words = [w for w in line_words if 378 <= w[0] <= 425 and w[4] != 'INR']
                        if credit_words:
                            try:
                                credit_val = float(credit_words[0][4].replace(',', ''))
                            except ValueError:
                                pass

                    # Extract balance (x=470-525)
                    balance_words = [w for w in line_words if 470 <= w[0] <= 525 and w[4] != 'INR']
                    if balance_words:
                        try:
                            balance_val = float(balance_words[0][4].replace(',', ''))
                        except ValueError:
                            pass

                if date_text and balance_val is not None:
                    transactions.append({
                        'date': date_text,
                        'narration': " ".join(narration_parts).strip(),
                        'chq_ref': '',
                        'val_date': date_text,
                        'withdrawal': debit_val,
                        'deposit': credit_val,
                        'balance': balance_val,
                        'page': page_num,
                    })

        doc.close()
        print(f"Extraction complete. Found {len(transactions)} raw transactions.")
        return transactions

    def parse_indianbank_pdfplumber(self):
        print(f"Opening PDF (pdfplumber): {self.pdf_path}")
        transactions = []
        date_pattern = re.compile(r'^\d{2}$')
        month_pattern = re.compile(r'^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)$')
        year_pattern = re.compile(r'^\d{4}$')

        with pdfplumber.open(self.pdf_path) as pdf:
            total_pages = len(pdf.pages)
            print(f"Total pages to parse (Indian Bank): {total_pages}")

            for page_idx in range(total_pages):
                page = pdf.pages[page_idx]
                page_num = page_idx + 1

                if self.verbose or page_num % 10 == 0 or page_num == total_pages:
                    print(f"Processing page {page_num}/{total_pages}...")

                words = page.extract_words()
                # Indian Bank table region: y 65-780
                table_words = [w for w in words if 65 <= w['top'] <= 780]

                # Group by y-coordinate
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

                # Find date lines
                date_ys = []
                for gk in sorted_ys:
                    line_words = sorted(grouped_lines[gk], key=lambda w: w['x0'])
                    date_words = [w for w in line_words if 71 <= w['x0'] <= 124]
                    if len(date_words) >= 3:
                        day_text = date_words[0]['text']
                        month_text = date_words[1]['text']
                        year_text = date_words[2]['text']
                        if (date_pattern.match(day_text) and 
                            month_pattern.match(month_text) and 
                            year_pattern.match(year_text)):
                            date_ys.append((gk, f"{day_text} {month_text} {year_text}"))

                if not date_ys:
                    continue

                # Process each transaction
                for i, (dy, date_text) in enumerate(date_ys):
                    block_end = date_ys[i + 1][0] if i + 1 < len(date_ys) else (sorted_ys[-1] + 1)
                    block_ys = [gk for gk in sorted_ys if dy <= gk < block_end]

                    narration_parts = []
                    debit_val = None
                    credit_val = None
                    balance_val = None

                    for gk in block_ys:
                        line_words = sorted(grouped_lines[gk], key=lambda w: w['x0'])

                        # Extract narration (x=145-260)
                        n_words = [w['text'] for w in line_words if 145 <= w['x0'] <= 260]
                        if n_words:
                            narration_parts.append(" ".join(n_words))

                        # Check for debit/credit indicators
                        debit_indicator = [w for w in line_words if 305 <= w['x0'] <= 308 and w['text'] == '-']
                        credit_indicator = [w for w in line_words if 400 <= w['x0'] <= 403 and w['text'] == '-']

                        # Extract debit amount (x=283-330)
                        if not debit_indicator:
                            debit_words = [w for w in line_words if 283 <= w['x0'] <= 330 and w['text'] != 'INR']
                            if debit_words:
                                try:
                                    debit_val = float(debit_words[0]['text'].replace(',', ''))
                                except ValueError:
                                    pass

                        # Extract credit amount (x=378-425)
                        if not credit_indicator:
                            credit_words = [w for w in line_words if 378 <= w['x0'] <= 425 and w['text'] != 'INR']
                            if credit_words:
                                try:
                                    credit_val = float(credit_words[0]['text'].replace(',', ''))
                                except ValueError:
                                    pass

                        # Extract balance (x=470-525)
                        balance_words = [w for w in line_words if 470 <= w['x0'] <= 525 and w['text'] != 'INR']
                        if balance_words:
                            try:
                                balance_val = float(balance_words[0]['text'].replace(',', ''))
                            except ValueError:
                                pass

                    if date_text and balance_val is not None:
                        transactions.append({
                            'date': date_text,
                            'narration': " ".join(narration_parts).strip(),
                            'chq_ref': '',
                            'val_date': date_text,
                            'withdrawal': debit_val,
                            'deposit': credit_val,
                            'balance': balance_val,
                            'page': page_num,
                        })

        print(f"Extraction complete. Found {len(transactions)} raw transactions.")
        return transactions

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
    parser = argparse.ArgumentParser(description="Convert Bank PDF Statement (HDFC/SBI/IOB/Canara/Indian Bank) to CSV/Excel")
    parser.add_argument("pdf_path", help="Path to Bank Statement PDF file")
    parser.add_argument("--csv", help="Output path for CSV file (default: pdf_name_parsed.csv)")
    parser.add_argument("--xlsx", help="Output path for Excel file (default: pdf_name_parsed.xlsx)")
    parser.add_argument("--bank", choices=['hdfc', 'sbi', 'iob', 'canara', 'indianbank'], help="Bank type (auto-detected by default)")
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
