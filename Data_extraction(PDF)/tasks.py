from robocorp.tasks import task
from pdf_data import pdf_data
import os

@task
def main_task():
    """
    Scan PDF, extract tables containing bid details and Technical Specifications and save them as Excel files.
    """
    
    try:
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
        except Exception as e:
            print(f"Error extracting technical specifications hyperlink: {e}")
            return  # Exit the task if unable to extract hyperlink
        
        # Download PDF from the hyperlink
        try:
            pdf_filename = os.path.basename(pdf_path)
            save_path = os.path.join(os.path.dirname(__file__), pdf_filename)  # Save in the same directory as the program
            pdf.download_pdf(link, save_path)
        except Exception as e:
            print(f"Error downloading PDF: {e}")
            return  # Exit the task if unable to download PDF
        
        # Extract tables from downloaded PDF (technical specifications)
        try:
            pdf = pdf_data(save_path)
            table = pdf.extract_tables_from_pdf()
            pdf.save_tables_to_excel(table, "technical_specification")
        except Exception as e:
            print(f"Error extracting tables from downloaded PDF: {e}")
        
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
