import unicodedata
import re
from typing import List
import fitz
import pandas as pd


def remove_non_ascii(text):
    cleaned_text = re.sub(r'[^\x00-\x7F]', '', text)
    cleaned_text = re.sub(r'[\u263a-\U0001f645]', '', cleaned_text)
    cleaned_text = cleaned_text.replace('/', '')
    return cleaned_text.strip()


def extract_tables_from_pdf(pdf_path):
    tables_data = []
    with fitz.open(pdf_path) as pdf:
        page = pdf[0]
        tables = page.find_tables()

        for table in tables:
            # Extract table data
            table_data = table.extract()
            cleaned_table_data = [[remove_non_ascii(cell) if cell else None for cell in row] for row in table_data]
            # Convert table data to DataFrame
            df = pd.DataFrame(cleaned_table_data[1:], columns=cleaned_table_data[0])
            # # Append DataFrame to list
            tables_data.append(df)
    return tables_data

pdf_path = "GeM-Bidding-5927594.pdf"
pdf_tables = extract_tables_from_pdf(pdf_path)
for idx, table_df in enumerate(pdf_tables):
    print(f"Table {idx+1}:")
    print(table_df)
    print("\n")
