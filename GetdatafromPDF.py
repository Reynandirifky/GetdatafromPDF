!pip -q install pypdf pandas openpyxl
import re
from pathlib import Path
from datetime import datetime

import pandas as pd
from pypdf import PdfReader


def extract_text_from_pdf(pdf_path: str) -> str:
    reader = PdfReader(pdf_path)
    return "\n".join(page.extract_text() or "" for page in reader.pages)


def extract_ici_data(pdf_path: str) -> dict:
    pdf_file = Path(pdf_path)
    if not pdf_file.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_file}")

    text = extract_text_from_pdf(str(pdf_file))

    date_match = re.search(
        r"(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)\s+(\d{1,2}\s+[A-Za-z]+\s+\d{4})",
        text,
    )
    if not date_match:
        raise ValueError("Could not find report date in the PDF.")

    report_date = datetime.strptime(date_match.group(2), "%d %B %Y").strftime("%Y-%m-%d")

    patterns = {
        "ICI 1": r"ICI 1 .*?\)\s+([0-9]+\.[0-9]+)",
        "ICI 2": r"ICI 2 .*?\)\s+([0-9]+\.[0-9]+)",
        "ICI 3": r"ICI 3 .*?\)\s+([0-9]+\.[0-9]+)",
        "ICI4": r"ICI 4 .*?\)\s+([0-9]+\.[0-9]+)",
        "ICI 5": r"ICI 5 .*?\)\s+([0-9]+\.[0-9]+)",
    }

    row = {"Date": report_date}

    for col, pattern in patterns.items():
        match = re.search(pattern, text, re.DOTALL)
        if not match:
            raise ValueError(f"Could not find value for {col}")
        row[col] = float(match.group(1))

    return row


def save_to_excel(row: dict, output_excel: str) -> None:
    columns = ["Date", "ICI 1", "ICI 2", "ICI 3", "ICI4", "ICI 5"]
    new_df = pd.DataFrame([row], columns=columns)

    output_file = Path(output_excel)
    if output_file.exists():
        existing_df = pd.read_excel(output_file)
        combined_df = pd.concat([existing_df, new_df], ignore_index=True)
        combined_df = combined_df.drop_duplicates(subset=["Date"], keep="last")
        combined_df = combined_df.sort_values("Date")
    else:
        combined_df = new_df

    combined_df.to_excel(output_file, index=False)


# ===== NOTEBOOK USAGE =====
input_pdf = "/content/Argus_Coalindo Indonesian Coal Index Report.pdf"
output_excel = "ici_prices.xlsx"

row = extract_ici_data(input_pdf)
save_to_excel(row, output_excel)

print(row)
print(f"Saved to {output_excel}")
