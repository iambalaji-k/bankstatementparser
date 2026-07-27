# Bank Statement PDF Parser (HDFC / SBI / IOB)

A high-performance parser to convert **HDFC**, **SBI**, and **Indian Overseas Bank** PDF statements into clean, structured CSV and Excel (.xlsx) formats. Handles multi-line narrations, page-boundary overflows, reverse-chronological ordering, and mixed withdrawal/deposit alignments with **100% mathematical consistency**.

Available in both **Rust** (recommended) and **Python** versions.

---

## Features

- **Coordinate-Based Extraction**: Uses precise bounding boxes (`x` coordinates) to align mixed withdrawals and deposits, preventing column shifting.
- **Narrations Stitching**: Merges multi-line narrations sequentially, including those split across page margins.
- **Mathematical Integrity Validation**: Confirms that:
  $$\text{Previous Balance} - \text{Withdrawal} + \text{Deposit} = \text{Closing Balance}$$
  for every transaction, flagging any discrepancy.
- **Auto-Categorization**: Dynamically classifies transactions into groups (Salary, UPI, IMPS, NEFT, Card Payments, ATM, Investments, etc.).
- **Dual Formats**: Generates clean CSV and beautifully formatted Excel (.xlsx) files (bold header, right-aligned currency columns, auto-width adjustments).

---

## Installation (Rust - Recommended)

### Prerequisites
- [Rust](https://www.rust-lang.org/tools/install) (1.70+)

### Build from source
```bash
cargo build --release
```

The binary will be at `target/release/hdfc-parser.exe` (Windows) or `target/release/hdfc-parser` (Linux/macOS).

---

## Installation (Python)

1. Clone or download this repository.
2. Initialize virtual environment and activate it:
   ```bash
   python -m venv venv
   # On Windows:
   venv\Scripts\activate
   ```
3. Install dependencies:
   ```bash
   pip install pymupdf openpyxl pdfplumber
   ```

---

## Usage

### Rust version
```bash
./hdfc-parser "statement.pdf"
```

### Python version
```bash
python parser.py "statement.pdf"
```

**Auto-detection:** The parser reads page 1 and detects the bank automatically. Pass `--bank hdfc|sbi|iob` to override.

| Option | Description |
|--------|-------------|
| `--csv PATH` | Specify custom output path for CSV |
| `--xlsx PATH` | Specify custom output path for Excel |
| `--bank hdfc|sbi|iob` | Manually specify bank (auto-detected by default) |
| `--verbose` | Enable detailed processing logs |

---

## Output Files

The script generates two files alongside your statement:
1. `statement_parsed.csv` — Raw structured CSV
2. `statement_parsed.xlsx` — Formatted Excel workbook

---

## Performance

The Rust version leverages native compilation for maximum performance. The Python version uses PyMuPDF's C-bindings for fast extraction.

| Metric | Python (PyMuPDF) | Rust |
|--------|------------------|------|
| 147-page statement | ~0.5s | <0.2s |
| Dependencies | pymupdf, openpyxl, pdfplumber | lopdf, rust_xlsxwriter |

---

## Architecture

### Rust Implementation
```
src/
├── main.rs          # CLI entry point with clap argument parsing
├── parser.rs        # PDF parsing with lopdf (coordinate-based extraction)
├── transaction.rs   # Transaction struct and categorization logic
└── output.rs        # CSV and Excel export
```

### Python Implementation
- `parser.py` — Single-file implementation with PyMuPDF/pdfplumber fallback

---

## Transaction Categories

| Category | Keywords |
|----------|----------|
| Salary | salary, payroll, betterplace |
| Foreign Exchange | foreign, usd, eur, inw, remittance, forex |
| UPI | upi- |
| IMPS | imps- |
| NEFT | neft- |
| RTGS | rtgs- |
| ATM Withdrawal | atm-, atm wdl |
| Card Payment | card-, pos-, pos wdl |
| Cheque | chq-, cheque, clg- |
| Interest Income | interest credit, int.coll, interest paid |
| Refund/Cashback | refund, cashback |
| Bank Charges | charge, fee, gst-, tax, annual maint |
| Sweep/MOD | sweep, autosweep, mod |
| Investment | mutual fund, zerodha, groww, indmoney, icici direct |
| Loan EMI | loan, emi- |
| Insurance | insurance, lic |

---

## License

This project is open source and available for personal and commercial use.