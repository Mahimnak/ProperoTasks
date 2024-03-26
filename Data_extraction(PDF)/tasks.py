from robocorp.tasks import task
from pdf_data import pdf_data
import os

@task
def main_task():
    """
    Scan PDF, extract tables containing bid details and Technical Specifications and save them as Excel files.
    """
    
    try:
        # Enter the pdf path here or use a file picker dialog to select it
        pdf_path = "GeM-Bidding-5927594.pdf"
        pdf = pdf_data(pdf_path)
        
        # Extract bid details tables
        try:
            tables = pdf.extract_tables_from_pdf()
            pdf.save_tables_to_excel(tables, "bid_details")
        except Exception as e:
            print(f"Error extracting bid details tables: {e}")
        # Extract technical specifications hyperlink
        try:
            link = pdf.extract_hyperlinks_from_pdf()
            pdf.add_link_excel(link)
        except Exception as e:
            print(f"Error extracting technical specifications hyperlink: {e}")
            return  # Exit the task if unable to extract hyperlink
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
