#!/usr/bin/env python3
"""
HDFC Bank PDF Statement Parser
Converts HDFC bank statement PDF to CSV/Excel using coordinate-based layout parsing.
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


class HDFCParser:
    def __init__(self, pdf_path, verbose=False):
        self.pdf_path = pdf_path
        self.verbose = verbose
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")

    def clean_amount(self, amount_str):
        """Cleans formatting from amount strings and converts to float/string representation."""
        if not amount_str:
            return None
        # Remove commas, spaces
        cleaned = amount_str.replace(",", "").strip()
        if not cleaned:
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None

    def format_amount(self, val):
        """Formats a float to 2 decimal places, or returns empty string if None."""
        if val is None:
            return ""
        return f"{val:.2f}"

    def parse(self):
        """Parses the bank statement PDF using PyMuPDF (fast) or pdfplumber (fallback)."""
        if HAS_FITZ:
            try:
                return self.parse_fitz()
            except Exception as e:
                print(f"PyMuPDF parsing failed: {e}. Falling back to pdfplumber...")
        return self.parse_pdfplumber()

    def parse_fitz(self):
        """Parses using PyMuPDF (fitz) for maximum speed."""
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
            ('balance', 560, 630)
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
            
            # Filter table region
            table_words = [w for w in words if 225 <= w[1] <= 780]
            
            # Group by vertical coordinate y0
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
            
            # Process lines
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
                        'page': page_num
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

    def parse_pdfplumber(self):
        """Parses the bank statement PDF page by page using pdfplumber (fallback)."""
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
            ('balance', 560, 630)
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
                        
                        mapped = False
                        for name, start, end in col_definitions:
                            if start <= x0 < end:
                                cols[name].append(text)
                                mapped = True
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
                            'page': page_num
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

    def categorize(self, narration):
        """Intelligently categorizes a transaction based on narration keywords."""
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
            # Categorization
            tx['category'] = self.categorize(tx['narration'])
            
            # Math validation
            if i > 0:
                prev_balance = processed[i - 1]['balance']
                curr_balance = tx['balance']
                w_amt = tx['withdrawal'] or 0.0
                d_amt = tx['deposit'] or 0.0
                
                # Expected balance
                expected_balance = prev_balance - w_amt + d_amt
                
                # Check with floating point tolerance
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
        """Saves transactions to CSV file."""
        print(f"Saving to CSV: {output_path}")
        with open(output_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            # Write header
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
                    tx['page']
                ])
        print(f"CSV file saved successfully: {os.path.abspath(output_path)}")

    def save_excel(self, transactions, output_path):
        """Saves transactions to Excel workbook if openpyxl is installed."""
        if not HAS_OPENPYXL:
            print("Excel export skipped: openpyxl is not installed.")
            return False
            
        print(f"Saving to Excel: {output_path}")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Transactions"
        
        # Style/Headers
        headers = ['Date', 'Narration', 'Chq/Ref No', 'Value Date', 'Withdrawal', 'Deposit', 'Balance', 'Category', 'Page']
        ws.append(headers)
        
        # Write data
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
                tx['page']
            ])
            
        # Add basic cell formatting (alignment & number formatting)
        from openpyxl.styles import Font, Alignment, PatternFill
        from openpyxl.utils import get_column_letter
        
        # Format Header row
        header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")
            
        # Format columns
        for row in range(2, len(transactions) + 2):
            # Withdrawal (col 5)
            cell_w = ws.cell(row=row, column=5)
            if cell_w.value != "":
                cell_w.number_format = '#,##0.00'
                cell_w.alignment = Alignment(horizontal="right")
                
            # Deposit (col 6)
            cell_d = ws.cell(row=row, column=6)
            if cell_d.value != "":
                cell_d.number_format = '#,##0.00'
                cell_d.alignment = Alignment(horizontal="right")
                
            # Balance (col 7)
            cell_b = ws.cell(row=row, column=7)
            if cell_b.value != "":
                cell_b.number_format = '#,##0.00'
                cell_b.alignment = Alignment(horizontal="right")
                
            # Dates (cols 1, 4)
            ws.cell(row=row, column=1).alignment = Alignment(horizontal="center")
            ws.cell(row=row, column=4).alignment = Alignment(horizontal="center")
            
        # Auto-adjust column widths
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                if cell.column == 2:  # Limit narration width to keep it neat
                    max_len = min(max(max_len, len(val_str)), 60)
                else:
                    max_len = max(max_len, len(val_str))
            ws.column_dimensions[col_letter].width = max(max_len + 3, 10)
            
        wb.save(output_path)
        print(f"Excel file saved successfully: {os.path.abspath(output_path)}")
        return True


def main():
    parser = argparse.ArgumentParser(description="Convert HDFC Bank PDF Statement to CSV/Excel")
    parser.add_argument("pdf_path", help="Path to HDFC Bank Statement PDF file")
    parser.add_argument("--csv", help="Output path for CSV file (default: pdf_name.csv)")
    parser.add_argument("--xlsx", help="Output path for Excel file (default: pdf_name.xlsx)")
    parser.add_argument("--verbose", action="store_true", help="Enable verbose log printing")
    
    args = parser.parse_args()
    
    # Auto-resolve output paths if not specified
    pdf_base = os.path.splitext(args.pdf_path)[0]
    csv_path = args.csv or f"{pdf_base}_parsed.csv"
    xlsx_path = args.xlsx or f"{pdf_base}_parsed.xlsx"
    
    try:
        start_time = datetime.now()
        parser_inst = HDFCParser(args.pdf_path, verbose=args.verbose)
        
        # Parse PDF
        raw_txs = parser_inst.parse()
        
        # Process and Validate
        processed_txs = parser_inst.validate_and_process(raw_txs)
        
        # Save to CSV
        parser_inst.save_csv(processed_txs, csv_path)
        
        # Save to Excel
        if HAS_OPENPYXL:
            parser_inst.save_excel(processed_txs, xlsx_path)
        else:
            print("Note: Install 'openpyxl' (pip install openpyxl) to enable formatting Excel .xlsx output.")
            
        end_time = datetime.now()
        elapsed = (end_time - start_time).total_seconds()
        
        # Print summary
        total_wd = sum(tx['withdrawal'] or 0.0 for tx in processed_txs)
        total_dp = sum(tx['deposit'] or 0.0 for tx in processed_txs)
        print("\n" + "="*40)
        print("Parsing Summary")
        print("="*40)
        print(f"Processed:          {len(processed_txs)} transactions")
        print(f"Total Withdrawals:  INR {total_wd:,.2f}")
        print(f"Total Deposits:     INR {total_dp:,.2f}")
        print(f"Net Change:         INR {(total_dp - total_wd):,.2f}")
        if processed_txs:
            print(f"Starting Balance:   INR {processed_txs[0]['balance'] - (processed_txs[0]['deposit'] or 0) + (processed_txs[0]['withdrawal'] or 0):,.2f}")
            print(f"Ending Balance:     INR {processed_txs[-1]['balance']:,.2f}")
        print(f"Time Elapsed:       {elapsed:.1f} seconds")
        print("="*40)
        
    except Exception as e:
        print(f"Error during parsing: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
