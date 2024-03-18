from robocorp.tasks import task
from pdf_data import pdf_data

@task
def main_task():
    """
    Scan PDF, extract tables containing bid details and Technical Specifications and save them as Excel files.
    """
    pdf_path = "GeM-Bidding-5927594.pdf"
    pdf = pdf_data(pdf_path)
    
    tables = pdf.extract_tables_from_pdf()
    pdf.save_tables_to_excel(tables,"bid_details")
    # technical_specification
    link = pdf.extract_hyperlinks_from_pdf()

    pdf  = pdf_data(link)
    tables = pdf.extract_tables_from_pdf()
    pdf.save_tables_to_excel(tables,"technical_specification")
    