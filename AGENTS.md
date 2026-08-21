# AGENTS.md

## What this is

Two standalone Python CLIs, no package/test/CI infrastructure:

- `parser.py` — converts bank statement PDFs (HDFC, SBI, IOB, Canara, Indian Bank) to CSV/XLSX via coordinate-based extraction.
- `pdf_unlock.py` — strips password protection from PDFs; run this before parsing encrypted statements.

## Commands

```bash
pip install -r requirements.txt

python parser.py pdf/HDFC.pdf              # auto-detects bank; writes HDFC_parsed.csv/.xlsx beside the input
python parser.py stmt.pdf --bank iob       # override auto-detection
python parser.py pdf/SBI.pdf --verbose     # per-page progress logs

python pdf_unlock.py encrypted.pdf         # prompts for password; writes <name>_unlocked.pdf
```

Gotcha: `pdf_unlock.py` imports `pypdf`, which is **not** in `requirements.txt` — install it separately or the script fails with ImportError.

## Verification (no test suite exists)

The regression check is running every sample PDF in `pdf/` and confirming:

1. Transaction count matches baseline: HDFC 1601, SBI 167, IOB 693, Canara 521, Indianbank 30.
2. Output ends with `Math Validation: SUCCESS` (0 discrepancies).

Validate across **all five** samples before considering parser changes done — a fix that works on one bank often breaks another. Sample PDFs and `*_parsed.*` outputs are gitignored; they exist only in this working copy.

## Architecture notes

- All parsing logic is in `parser.py`. Bank-specific behavior is entirely data-driven by `BankProfile` dataclasses in `BANK_PROFILES`: column x-bounds in PDF points, date regexes, y-cutoffs (`page1_min_y` / `pageN_min_y` / `max_y`), footer/header keywords.
- Adding a bank touches **three** places: a `BANK_PROFILES` entry, `detect_bank()` (page-1 text matching), and argparse `--bank` choices. Core parsing code should not need changes.
- Three anchor strategies via `date_type`: `"single"` (date regex in date column), `"split_3"` (Indian Bank: day/month/year as three separate words), `"iob"` (balance-value anchored blocks; IOB statements are reverse-chronological and get re-sorted at the end).
- Pipeline: extract words with coordinates (PyMuPDF preferred, pdfplumber fallback) → y-filter header/footer → group words into lines by y-tolerance → drop footer/header keyword lines → segment transactions between date anchors → append cross-page narration overflow to the previous transaction.
- Math validation (`prev_balance - withdrawal + deposit == balance`, tolerance 0.02) is the correctness gate. Mass discrepancies usually mean wrong `col_bounds` or a missed footer keyword — not bad arithmetic.
- Unrecognized page-1 text falls back to the **HDFC profile** with only a warning; wrong-profile parses show up as math discrepancies, so check the "Detected bank" line first when debugging.

## Conventions

- Commit messages: no Co-authored-by or other author trailers.
- Static code review means reading only — don't execute scripts or trigger builds unless explicitly asked to validate.
- When asked a question or told "don't edit yet", investigate, report findings, and wait for approval before editing files.
