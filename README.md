# 🏦 Multi-Bank Statement PDF Parser

**Turn messy bank statement PDFs into clean, analysis-ready spreadsheets — in seconds.**

A high-performance, profile-driven engine that converts **HDFC**, **SBI**, **Indian Overseas Bank (IOB)**, **Canara Bank**, and **Indian Bank** statement PDFs into structured **CSV** and beautifully formatted **Excel (.xlsx)** files — with built-in mathematical proof that every single row was extracted correctly.

No manual copy-pasting. No OCR guesswork. No broken columns. Just point it at a PDF and get perfect tabular data.

![Python](https://img.shields.io/badge/python-3.8%2B-blue) ![License](https://img.shields.io/badge/license-Unlicense-green) ![Banks](https://img.shields.io/badge/banks-5%20supported-orange)

---

## ✨ Why You'll Love It

- 🎯 **Coordinate-Based Precision** — Instead of fragile text scraping, every word is extracted with its exact x/y position on the page. Withdrawals, deposits, and balances land in the right column *every time*, even with multi-line narrations.
- 🧮 **Mathematical Integrity Guarantee** — Every parsed row is verified against the accounting equation `previous balance − withdrawal + deposit = balance`. If the math doesn't add up, you'll know immediately — parsing errors can't hide.
- 🤖 **Automatic Bank Detection** — The parser reads page 1 and figures out which bank's layout it's dealing with. Zero configuration required.
- 🔀 **Cross-Page Continuation** — Narrations that wrap across page boundaries are automatically stitched back onto the right transaction.
- 🔄 **Reverse-Chronological Handling** — IOB statements list transactions newest-first; the parser detects this and re-sorts everything chronologically for you.
- 🏷️ **Smart Auto-Categorization** — Every transaction is classified into categories like UPI, IMPS, NEFT, RTGS, ATM, Salary, Loan EMI, Investments, and more — ready for pivot tables and spending analysis.
- 📊 **Gorgeous Excel Output** — Styled workbooks with navy headers, right-aligned currency columns, `#,##0.00` number formatting, and auto-sized columns. Open it and send it.
- ⚡ **Blazing Fast** — Powered by PyMuPDF's low-level C bindings, hundreds of pages parse in under a second.

---

## 🏛️ Supported Banks

| Bank | Profile Key | Date Format | Layout Quirks Handled |
| :--- | :---: | :---: | :--- |
| **HDFC Bank** | `hdfc` | `DD/MM/YY(YY)` | Separate value-date column, summary footer suppression |
| **State Bank of India** | `sbi` | `DD/MM/YYYY` | Computer-generated disclaimer footers, per-page headers |
| **Indian Overseas Bank** | `iob` | `DD-Mon-YY` | Balance-anchored blocks, value dates in parentheses, reverse-chronological order |
| **Canara Bank** | `canara` | `DD-Mon-YY` | Taller pages (`max_y=840`), no chq/ref or value-date columns |
| **Indian Bank** | `indianbank` | `DD Mon YYYY` | Dates split across three words, dash indicators for empty amount cells, `INR` currency prefixes |

> 🚧 **More banks are on the way!** New bank profiles will be added in future releases — see the [Roadmap](#🗺️-roadmap).

---

## 📦 Installation

Requires **Python 3.8+**.

```bash
git clone https://github.com/iambalaji-k/bankstatementparser.git
cd bankstatementparser
pip install -r requirements.txt
```

> **🔐 Encrypted statements?** `pdf_unlock.py` needs an extra dependency that isn't in `requirements.txt`:
>
> ```bash
> pip install pypdf
> ```

---

## 🚀 Usage

### Parse a statement

```bash
python parser.py pdf/HDFC.pdf
```

That's it. The bank is auto-detected from page-1 header text, and `HDFC_parsed.csv` + `HDFC_parsed.xlsx` appear beside the input file.

### Unlock a password-protected PDF first

Bank statements are usually password-protected (often your date of birth or PAN). Strip the password, then parse:

```bash
python pdf_unlock.py encrypted.pdf              # prompts for password securely
python pdf_unlock.py encrypted.pdf -p secret123 # or pass it inline
python pdf_unlock.py encrypted.pdf -o plain.pdf # custom output path

python parser.py encrypted_unlocked.pdf
```

### Command-Line Reference

| Argument | Required | Default | Description |
| :--- | :---: | :--- | :--- |
| `pdf_path` | ✅ Yes | — | Path to the statement PDF |
| `--bank` | No | auto-detected | Force a profile: `hdfc`, `sbi`, `iob`, `canara`, `indianbank` |
| `--csv` | No | `<pdf>_parsed.csv` | Custom CSV output path |
| `--xlsx` | No | `<pdf>_parsed.xlsx` | Custom Excel output path |
| `--verbose` | No | off | Log progress for **every** page instead of every 10th |

### Examples

```bash
python parser.py stmt.pdf --bank iob                              # override auto-detection
python parser.py Canara.pdf --csv out/july.csv --xlsx out/july.xlsx  # custom destinations
python parser.py pdf/SBI.pdf --verbose                            # watch every page fly by
```

---

## ⚙️ How It Works

The parser treats each PDF as a coordinate system, not a document. Here's the pipeline:

```
PDF ─▶ Extract words + x/y coordinates (PyMuPDF ▶ pdfplumber fallback)
   ─▶ Filter header/footer zones using per-bank y-cutoffs
   ─▶ Group words into horizontal lines by y-tolerance
   ─▶ Drop column-header rows & footer summary blocks (keyword matching)
   ─▶ Segment lines into transactions between date anchors
   ─▶ Append cross-page narration overflow to the previous transaction
   ─▶ Validate math: prev_balance − withdrawal + deposit == balance
   ─▶ Auto-categorize → write styled CSV + XLSX
```

### Three Anchor Strategies

Different banks lay out transactions differently, so the engine supports three ways of finding where each transaction starts:

1. **`single`** *(HDFC, SBI, Canara)* — A line whose leftmost column matches the bank's date regex starts a new transaction. Everything until the next date belongs to it.
2. **`split_3`** *(Indian Bank)* — Dates arrive as three separate words (`15` `Aug` `2025`) in the date column; all three must match before a transaction anchor is recognized.
3. **`iob`** *(IOB)* — Dates are unreliable, so transactions are anchored to **balance values** instead, with configurable y-offsets defining each block. Reverse-chronological pages are re-sorted at the end.

---

## 🧩 Under the Hood: `BankProfile`

All bank-specific behavior lives in declarative `BankProfile` dataclasses inside `BANK_PROFILES` — the core parsing engine never changes. A profile encodes:

```python
'hdfc': BankProfile(
    name='hdfc',
    display_name='HDFC Bank',
    date_pattern=re.compile(r'^\d{2}/\d{2}/\d{2,4}$'),  # what a date looks like
    date_x_range=(30.0, 65.0),                          # where dates live horizontally
    col_bounds={...},                                   # x-bounds (in points) per column
    page1_min_y=225.0,                                  # skip account info above the table
    pageN_min_y=40.0,                                   # later pages start higher
    max_y=780.0,                                        # stop before the footer
    footer_keywords=['STATEMENT SUMMARY', 'TOTAL'],     # halt parsing here
    header_keywords=['Date', 'Narration', 'Chq/Ref'],   # skip repeated column titles
)
```

| Field | Purpose |
| :--- | :--- |
| `col_bounds` | x-coordinate ranges (PDF points) for `date`, `narration`, `chq_ref`, `val_date`, `withdrawal`, `deposit`, `balance` — `None` if the bank lacks a column |
| `page1_min_y` / `pageN_min_y` / `max_y` | Vertical safe zone containing only table rows |
| `date_type` | Anchor strategy: `single`, `split_3`, or `iob` |
| `footer_keywords` / `header_keywords` | Word-boundary-matched lines to suppress (footers **stop** the page; headers are skipped) |
| `chronological` | If `False`, pages and rows are re-sorted oldest-first at the end |
| `line_group_tolerance` | Max y-distance between words considered part of one line |
| `debit_dash_x_range` / `credit_dash_x_range` | Exact positions of `-` placeholders in amount columns (Indian Bank) |
| `currency_prefix` | Currency token to ignore while scanning amounts (e.g. `INR`) |

---

## ➕ Adding a New Bank

Three touchpoints — **zero changes to core parsing logic**:

1. **Add a `BankProfile`** entry to `BANK_PROFILES` in `parser.py`
2. **Add a match rule** to `detect_bank()` (page-1 text matching)
3. **Register the key** in argparse `--bank` choices

Measure the column x-positions from the PDF (any inspector tool works), fill in the bounds, and you're done.

---

## 📤 Output Format

Both CSV and Excel share the same schema:

| Column | Content |
| :--- | :--- |
| `Date` | Transaction date |
| `Narration` | Full description, including multi-line and cross-page overflow |
| `Chq/Ref No` | Cheque or reference number (blank where the bank omits it) |
| `Value Date` | Value/effective date |
| `Withdrawal` / `Deposit` | Amounts as decimals; blank when not applicable |
| `Balance` | Running balance after the transaction |
| `Category` | Auto-assigned spending category |
| `Page` | Source page number in the PDF — perfect for auditing back to the original |

The Excel workbook gets navy header styling, centered date columns, right-aligned `#,##0.00` currency formatting, and auto-fitted column widths.

---

## 🏷️ Auto-Categorization Rules

| Category | Triggered By |
| :--- | :--- |
| Salary | `salary`, `payroll`, `betterplace` |
| Foreign Exchange | `foreign`, `usd`, `eur`, `inw`, `remittance`, `forex` |
| UPI / IMPS / NEFT / RTGS | `upi-`, `imps-`, `neft-`, `rtgs-` prefixes |
| ATM Withdrawal | `atm-`, `atm wdl` |
| Card Payment | `card-`, `pos-`, `pos wdl` |
| Cheque | `chq-`, `cheque`, `clg-` |
| Interest Income | `interest credit`, `int.coll`, `interest paid` |
| Refund/Cashback | `refund`, `cashback` |
| Bank Charges | `charge`, `fee`, `gst-`, `tax`, `annual maint` |
| Sweep/MOD | `sweep`, `autosweep`, `mod ` |
| Investment | `mutual fund`, `zerodha`, `groww`, `indmoney`, `icici direct` |
| Loan EMI | `loan`, `emi-` |
| Insurance | `insurance`, `lic ` |
| Other Transfer/Spending | fallback |

---

## ✅ Accuracy & Validation

The math validator is the correctness gate. After extraction, every row is checked against:

```
previous balance − withdrawal + deposit ≈ balance   (tolerance ±0.02)
```

- **Zero discrepancies** → `Math Validation: SUCCESS`
- Any mismatch prints the row, both balances, and its source page so you can inspect the original instantly

Regression baselines across bundled sample statements:

| Statement | Transactions |
| :--- | :---: |
| HDFC | 1,601 |
| SBI | 167 |
| IOB | 693 |
| Canara | 521 |
| Indian Bank | 30 |

---

## 🩺 Troubleshooting

| Symptom | Likely Cause & Fix |
| :--- | :--- |
| Mass math discrepancies | Wrong profile applied — check the `Detected bank:` line first. Unrecognized PDFs fall back to HDFC with only a warning; pass `--bank` explicitly |
| Summary rows leaking into output | Missing footer keyword in the profile |
| Transactions merged/split wrongly | Adjust `line_group_tolerance` or the profile's date pattern/x-range |
| `ImportError` on unlock | Install the extra dependency: `pip install pypdf` |
| Excel file missing | Install openpyxl: `pip install openpyxl` |
| Password prompt loops | Run `python pdf_unlock.py <file>` before parsing |

---

## 🗺️ Roadmap

- [x] HDFC Bank
- [x] State Bank of India
- [x] Indian Overseas Bank
- [x] Canara Bank
- [x] Indian Bank
- [ ] **More banks coming soon!** New bank profiles will be added in future releases — stay tuned 🚀

Want a specific bank supported? [Open an issue](https://github.com/iambalaji-k/bankstatementparser/issues) with a redacted sample statement.

---

## 🤝 Contributing

Contributions are welcome! Adding a new bank is the most valuable contribution — see [Adding a New Bank](#➕-adding-a-new-bank). Please validate against all five existing samples before submitting parser changes: fixes that help one bank often break another.

---

## 📄 License

Released under the [Unlicense](LICENSE) — free and unencumbered software dedicated to the public domain. Copy, modify, publish, use, compile, sell, or distribute it for any purpose, commercial or non-commercial, with no attribution required.

### ⚠️ Disclaimer

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

This tool performs automated extraction of financial data and may produce parsing errors, math discrepancies, or incomplete output. **Always verify parsed results against the original PDF before relying on them for accounting, tax, legal, or any other financial purposes.** Use at your own risk.
