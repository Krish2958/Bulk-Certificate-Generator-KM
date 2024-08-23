### README.md

# Bulk Certificate Generator

This Python-based tool automates the generation of bulk certificates by replacing placeholders in a PowerPoint (PPTX) template with data from a CSV file and converting the modified presentations into PDF format.

## Features

- **Placeholder Replacement**: Automatically replaces placeholders in the template with data from a CSV file.
- **PPTX to PDF Conversion**: Converts the modified PPTX files to PDFs.
- **Automated Workflow**: Processes multiple records from a CSV, generating certificates for each entry and saving them as PDFs.
- **Error Handling**: Handles file existence checks and errors during PPTX to PDF conversion.

## Prerequisites

- Python 3.x
- Required Python Libraries:
  - `pandas`
  - `python-pptx`
  - `comtypes`
- Microsoft PowerPoint installed (for PPTX to PDF conversion using `comtypes`)

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/bulk-certificate-generator.git
   cd bulk-certificate-generator
   ```

2. Install the required Python libraries:
   ```bash
   pip install pandas python-pptx comtypes
   ```

## Usage

1. **Prepare Your Template**: Design your certificate in Microsoft PowerPoint, and add placeholders in the format `{{placeholder_name}}`. For example, use `{{name}}` where you want the recipient's name to appear.

2. **Prepare Your Data**: Create a CSV file where each column corresponds to a placeholder in your template. For example, the CSV might look like this:

   ```csv
   name,course,completion_date
   John Doe,Python Programming,2024-08-01
   Jane Smith,Web Development,2024-08-02
   ```

3. **Run the Script**: Update the paths for the template PPTX, CSV file, and output directory in the script. Then run the script:

   ```bash
   python bulk_certificate_generator.py
   ```

4. **View the Output**: The generated PDFs will be saved in the specified output directory. The intermediate PPTX files will be deleted after successful conversion to PDFs.

## Code Explanation

- `replace_placeholders(slide, replacements)`: Replaces the placeholders in each slide's text with the corresponding values from the CSV data.
- `convert_ppt_to_pdf(input_pptx_path, output_pdf_path)`: Converts the generated PPTX file to PDF using Microsoft PowerPoint.
- `process_ppt_and_generate_pdfs(template_ppt_path, csv_path, output_dir)`: Reads the CSV file, processes each record, generates a new PPTX file with the placeholders replaced, and then converts it to PDF.

## Example

Suppose you have a template file named `Openhack.pptx` with placeholders like `{{name}}`, `{{course}}`, and `{{completion_date}}`. If you have a `data.csv` file as shown in the "Usage" section, running the script will generate certificates for each person listed in the CSV and save them as PDFs in the output directory.

## Future Goals

### Creating a Web Platform to Automate Bulk Certificate Creation

- **Web Interface**: Develop a web-based platform that allows users to upload their PPTX templates and CSV files.
- **Automation**: The platform will automatically handle the entire process of generating and converting certificates in the cloud, removing the need for local PowerPoint installations.
- **Cloud Storage Integration**: Implement integration with cloud storage services like Google Drive or Dropbox for saving the generated certificates.
- **User Authentication**: Add user accounts and sessions so that users can track their generated certificates.
- **REST API**: Create a REST API that allows programmatic access to the certificate generation functionality.
- **Styling Options**: Provide an option for users to customize the styling of the replaced text, ensuring it matches the template's format.

## Contact

For any queries or issues, please contact [krishm.km17@gmail.com].

---

