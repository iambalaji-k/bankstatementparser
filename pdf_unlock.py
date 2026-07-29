#!/usr/bin/env python3

import argparse
import getpass
import sys
from pathlib import Path

from pypdf import PdfReader, PdfWriter


def remove_password(input_pdf, output_pdf, password=None):
    reader = PdfReader(input_pdf)

    if reader.is_encrypted:
        if password is None:
            password = getpass.getpass("PDF Password: ")

        result = reader.decrypt(password)

        if result == 0:
            print("❌ Incorrect password.")
            sys.exit(1)

    writer = PdfWriter()

    for page in reader.pages:
        writer.add_page(page)

    with open(output_pdf, "wb") as f:
        writer.write(f)

    print(f"✅ Password removed.")
    print(f"Saved as: {output_pdf}")


def main():
    parser = argparse.ArgumentParser(
        description="Remove password protection from a PDF."
    )

    parser.add_argument(
        "input",
        help="Input encrypted PDF"
    )

    parser.add_argument(
        "-o",
        "--output",
        help="Output PDF (default: <input>_unlocked.pdf)"
    )

    parser.add_argument(
        "-p",
        "--password",
        help="PDF password (if omitted, you'll be prompted)"
    )

    args = parser.parse_args()

    input_pdf = Path(args.input)

    if not input_pdf.exists():
        print(f"❌ File not found: {input_pdf}")
        sys.exit(1)

    if args.output:
        output_pdf = args.output
    else:
        output_pdf = input_pdf.with_stem(input_pdf.stem + "_unlocked")

    remove_password(input_pdf, output_pdf, args.password)


if __name__ == "__main__":
    main()