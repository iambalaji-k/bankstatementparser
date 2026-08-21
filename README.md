# Multi-Bank Statement PDF Parser

Convert bank statement PDFs into clean, structured **CSV** and **Excel (.xlsx)** files using coordinate-based layout parsing.

Supports statements from five Indian banks:

| Bank | Profile Key | Date Format | Notes |
| :--- | :---: | :--- | :--- |
| HDFC Bank | `hdfc` | `DD/MM/YY(YY)` | Default fallback profile |
| State Bank of India | `sbi` | `DD/MM/YYYY` | Separate value-date column |
| Indian Overseas Bank | `iob` | `DD-Mon-YY` | Reverse-chronological output, auto re-sorted |
| Canara Bank | `canara` | `DD-Mon-YY` | No chq/ref or value-date columns |
| Indian Bank | `indianbank` | `DD Mon YYYY` | Day/month/year split across three words |

## How It Works

The parser extracts every word from the PDF along with its x/y coordinates (via **PyMuPDF**, falling back to **pdfplumber**), then:

1. Filters out page headers/footers using per-bank y-coordinate cutoffs.
2. Groups words into horizontal lines by y-tolerance.
3. Drops footer/header keyword lines (statement summaries, column titles).
4. Segments lines into transactions between date anchors.
5. Appends narration text that overflows across page breaks to the previous transaction.
6. Validates every row against the accounting equation:

   ```
   previous balance - withdrawal + deposit == balance   (tolerance 0.02)
   ```

7. Auto-categorizes each transaction (UPI, IMPS, NEFT, RTGS, ATM, Salary, Loan EMI, etc.) and writes CSV + styled XLSX output.

Because bank-specific behavior is entirely data-driven by declarative `BankProfile` definitions (`BANK_PROFILES` in `parser.py`), adding a new bank requires no changes to core parsing code.

## Installation

Requires Python 3.8+.

```bash
pip install -r requirements.txt
```

> **Note:** `pdf_unlock.py` additionally requires `pypdf`, which is *not* in `requirements.txt`. Install it separately if you need to unlock encrypted PDFs:
>
> ```bash
> pip install pypdf
> ```

## Usage

### 1. Parse a statement

```bash
python parser.py pdf/HDFC.pdf
```

Auto-detects the bank from page-1 header text and writes `HDFC_parsed.csv` / `HDFC_parsed.xlsx` beside the input file.

```bash
# Override auto-detection
python parser.py stmt.pdf --bank iob

# Custom output paths
python parser.py Canara.pdf --csv out/july.csv --xlsx out/july.xlsx

# Per-page progress logs
python parser.py pdf/SBI.pdf --verbose
```

| Argument | Required | Default | Description |
| :--- | :---: | :--- | :--- |
| `pdf_path` | Yes | — | Path to the statement PDF |
| `--bank` | No | auto-detected | One of `hdfc`, `sbi`, `iob`, `canara`, `indianbank` |
| `--csv` | No | `<pdf>_parsed.csv` | Custom CSV output path |
| `--xlsx` | No | `<pdf>_parsed.xlsx` | Custom Excel output path |
| `--verbose` | No | off | Log progress for every page |

### 2. Unlock an encrypted PDF (optional)

Bank statements are often password-protected. Strip the password first, then parse the unlocked copy:

```bash
python pdf_unlock.py encrypted.pdf          # prompts for password
python pdf_unlock.py encrypted.pdf -p secret123
python pdf_unlock.py encrypted.pdf -o plain.pdf

python parser.py encrypted_unlocked.pdf
```

### Output Columns

Both formats share the same schema:

`Date, Narration, Chq/Ref No, Value Date, Withdrawal, Deposit, Balance, Category, Page`

## Adding a New Bank

Three touchpoints in `parser.py`, none inside the parsing engine:

1. Add a `BankProfile` entry to `BANK_PROFILES` (column x-bounds in PDF points, date regex, y-cutoffs, footer/header keywords).
2. Add a match rule to `detect_bank()` (page-1 text matching).
3. Add the key to argparse `--bank` choices.

## Troubleshooting

- **Math Validation warnings** usually mean wrong `col_bounds` or a missed footer keyword — not bad arithmetic. Check the "Detected bank:" line first; unrecognized PDFs silently fall back to the HDFC profile (warning only), which produces mass discrepancies on other banks.
- **Encrypted PDF errors**: run `pdf_unlock.py` before parsing.
- **Excel export skipped**: install `openpyxl`.

## Verification

No test suite — the regression check is running every sample PDF in `pdf/` and confirming:

1. Transaction counts match baseline: HDFC 1601, SBI 167, IOB 693, Canara 521, Indian Bank 30.
2. Output ends with `Math Validation: SUCCESS` (zero discrepancies).

Validate across all five samples before considering parser changes done — a fix that works on one bank often breaks another.

## License

Released under the [Unlicense](LICENSE) — free and unencumbered software dedicated to the public domain. Copy, modify, publish, use, compile, sell, or distribute it for any purpose, commercial or non-commercial, with no attribution required.

### Disclaimer

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

This tool performs automated extraction of financial data and may produce parsing errors, math discrepancies, or incomplete output. **Always verify parsed results against the original PDF before relying on them for accounting, tax, legal, or any other financial purposes.** Use at your own risk.
