### Updated README.md

# Bulk Certificate Generator

This Python-based tool automates the generation of bulk certificates by replacing placeholders in a PowerPoint (PPTX) template with data from a CSV file and converting the modified presentations into PDF format.

## Features

- **Placeholder Replacement**: Automatically replaces placeholders in the template with data from a CSV file.
- **PPTX to PDF Conversion**: Converts the modified PPTX files to PDFs.
- **Automated Workflow**: Processes multiple records from a CSV, generating certificates for each entry and saving them as PDFs.
- **Error Handling**: Handles file existence checks and errors during PPTX to PDF conversion.
- **Simple UI**: A user-friendly interface to select the template, CSV file, output directory, and generate certificates with just one click.

## Prerequisites

- Windows OS (Installer works on Windows systems)
- Microsoft PowerPoint installed (for PPTX to PDF conversion)
  
## Installation

### Download & Install the App

1. **Download the Installer**:
   - Head to the `dist` folder of the project (or the latest release) and download the `BulkCertificateGenerator.exe` file.

2. **Run the Installer**:
   - Locate the downloaded `.exe` file and double-click it to begin the installation process.
   - Follow the on-screen instructions to install the Bulk Certificate Generator tool.

## Usage

1. **Open the Application**:
   - After installation, launch the `Bulk Certificate Generator` from your desktop or start menu.

2. **Select Your Template**:
   - Use the "Select PPT Template" button to upload your PowerPoint template file with placeholders (e.g., `{{name}}`, `{{course}}`, etc.).

3. **Select Your Data**:
   - Click the "Select CSV File" button to choose the CSV file containing the data. The CSV should have columns matching the placeholders in your template (e.g., `name`, `course`, `completion_date`).

4. **Choose Output Directory**:
   - Select the folder where you want to save the generated PDF certificates using the "Select Output Directory" button.

5. **Generate Certificates**:
   - Once all the fields are selected, click the "Generate Certificates" button. The app will process the data, replace the placeholders in your template, and generate PDFs for each row in the CSV file.

6. **View the Output**:
   - The generated PDF certificates will be saved in the output directory you specified. Each certificate will be named based on the data from the CSV file (e.g., `JohnDoe_output.pdf`).

## Example

Suppose you have a PowerPoint template named `Openhack.pptx` with placeholders like `{{name}}`, `{{course}}`, and `{{completion_date}}`. If you have a CSV file (e.g., `data.csv`) structured like:

```csv
name,course,completion_date
John Doe,Python Programming,2024-08-01
Jane Smith,Web Development,2024-08-02
```

By selecting this template and CSV in the app, and choosing an output folder, the app will generate a PDF certificate for each entry in the CSV, replacing placeholders like `{{name}}` with "John Doe", `{{course}}` with "Python Programming", etc.

## Future Goals

### Web Platform Automation

- **Web Interface**: Develop a web-based platform that allows users to upload PPTX templates and CSV files.
- **Cloud Automation**: Automate the entire certificate generation and conversion process in the cloud, removing the need for local PowerPoint installations.
- **Cloud Storage**: Integrate cloud storage services like Google Drive or Dropbox for saving certificates.
- **User Authentication**: Add user authentication and sessions to allow tracking of generated certificates.
- **API Integration**: Provide a REST API for programmatic access to certificate generation features.
- **Custom Styling**: Allow users to customize the style of the replaced text, ensuring it matches the template's formatting.

## Contact

For any queries or issues, please contact [krishm.km17@gmail.com].

---

This updated README now reflects the steps to download, install, and use the executable version of the Bulk Certificate Generator, providing users with an intuitive interface to generate certificates effortlessly.