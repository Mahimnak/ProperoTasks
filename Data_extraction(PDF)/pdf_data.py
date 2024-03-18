import unicodedata
import re
import fitz
import openpyxl
import pandas as pd

class pdf_data:
    def __init__(self,pdf_path) -> None:
        self.pdf_path = pdf_path

    def remove_non_ascii(self, text)->str:
        cleaned_text = re.sub(r'[^\x00-\x7F]', '', text)
        cleaned_text = re.sub(r'[\u263a-\U0001f645]', '', cleaned_text)
        cleaned_text = cleaned_text.replace('/', '')
        return cleaned_text.strip()

    def extract_tables_from_pdf(self):
        tables_data = []
        with fitz.open(self.pdf_path) as pdf:
            page = pdf[0]
            tables = page.find_tables()

            for table in tables:
                # Extract table data
                table_data = table.extract()
                cleaned_table_data = [[self.remove_non_ascii(cell) if cell else None for cell in row] for row in table_data]
                # Convert table data to DataFrame
                df = pd.DataFrame(cleaned_table_data[1:], columns=cleaned_table_data[0])
                # # Append DataFrame to list
                tables_data.append(df)
        return tables_data

    def save_tables_to_excel(self, tables, worksheet_name):
        workbook = openpyxl.load_workbook("tables_from_pdf.xlsx")
        worksheet = workbook.get_sheet_by_name(worksheet_name)
        for idx, table_df in enumerate(tables):
            for r_idx, (_, row) in enumerate(table_df.iterrows(), start=1):
                for c_idx, value in enumerate(row, start=1):
                    value = str(value)
                    # Replace any characters that cannot be used in Excel worksheets
                    value = ''.join(char if unicodedata.category(char)[0] != 'C' else '' for char in value)
                    worksheet.cell(row=r_idx, column=c_idx, value=value)
        workbook.save("tables_from_pdf.xlsx")

    def extract_hyperlinks_from_pdf(self)->str:
        hyperlinks = []
        with fitz.open(self.pdf_path) as pdf:
            page = pdf[2]
            links = page.get_links()
            for link in links:
                hyperlink = {
                    "page_number": 3,
                    "url": link.get('uri', ''),
                    "rect": link.get('rect', [])
                }
                hyperlinks.append(hyperlink)
        tech_specs = hyperlinks[1]["url"]
        return tech_specs
    
def main():
    pdf_path = "GeM-Bidding-5927594.pdf"
    # excel_file = "tables_from_pdf.xlsx"
    # save_tables_to_excel(pdf_tables, excel_file)
    # print(f"Tables saved to {excel_file}")

if __name__ == "__main__":
    main()

