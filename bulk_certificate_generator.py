import os
import pandas as pd
from pptx import Presentation
import comtypes.client
from datetime import datetime

def replace_placeholders(slide, replacements):
    """Replace placeholders in the slide's text with values from replacements."""
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    for placeholder, value in replacements.items():
                        if placeholder in run.text:
                            run.text = run.text.replace(placeholder, value)

def convert_ppt_to_pdf(input_pptx_path, output_pdf_path):
    """Convert a PPTX file to PDF format."""
    try:
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = 1

        # Check if the file exists
        if not os.path.isfile(input_pptx_path):
            raise FileNotFoundError(f"The file {input_pptx_path} does not exist.")

        presentation = powerpoint.Presentations.Open(input_pptx_path)
        presentation.SaveAs(output_pdf_path, FileFormat=32)  # 32 for PDF format
        presentation.Close()
        print(f"Converted {input_pptx_path} to {output_pdf_path}")
        return True  # Indicate success
    except Exception as e:
        print(f"Failed to convert {input_pptx_path} to PDF: {e}")
        return False  # Indicate failure
    finally:
        powerpoint.Quit()

def process_ppt_and_generate_pdfs(template_ppt_path, csv_path, output_dir):
    """Generate PPTX and PDF files with placeholders replaced from CSV data."""
    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)

    # Read the CSV data
    data = pd.read_csv(csv_path)

    # Process each record in the CSV
    for index, row in data.iterrows():
        replacements = {f"{{{{{col}}}}}": str(row[col]) for col in data.columns}
        
        # Create a new Presentation object for each record
        prs = Presentation(template_ppt_path)
        
        # Create a copy of the presentation for each record
        timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
        output_pptx_path = os.path.join(output_dir, f"{row['name']}_{timestamp}_output.pptx")
        output_pdf_path = os.path.join(output_dir, f"{row['name']}_{timestamp}_output.pdf")
        
        # Apply the replacements
        for slide in prs.slides:
            replace_placeholders(slide, replacements)

        # Save the updated PPTX
        prs.save(output_pptx_path)
        print(f"Generated PPTX: {output_pptx_path}")

        # Convert PPTX to PDF
        if convert_ppt_to_pdf(output_pptx_path, output_pdf_path):
            # Remove the PPTX file after successful PDF conversion
            os.remove(output_pptx_path)
            print(f"Deleted PPTX: {output_pptx_path}")
        else:
            print(f"Failed to delete PPTX: {output_pptx_path} due to conversion failure.")

# Paths for the template, CSV, and output directory
template_ppt_path = r'C:\Users\krish\Projects\python test\Astro Quest Certificates.pptx'
csv_path = r'C:\Users\krish\Projects\python test\Results_evaluation - Participants (5).csv'
output_dir = r'C:\Users\krish\Projects\python test\output_cert'

# Process and generate the PPTs and PDFs
process_ppt_and_generate_pdfs(template_ppt_path, csv_path, output_dir)
