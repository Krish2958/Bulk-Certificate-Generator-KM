# Bulk Certificate Generator

This Python-based tool automates bulk certificate generation by replacing placeholders in a **PDF template** with data from a CSV file.

## Features

- **PDF Template Workflow**: Uses PDF templates directly (no PPTX conversion).
- **Platform Agnostic**: Works on Windows, macOS, Linux, and cloud/container environments.
- **Placeholder Replacement**: Replaces text placeholders like `{{name}}` from CSV columns.
- **Form-Field Support**: Fills PDF form fields (AcroForm widgets) using matching CSV columns.
- **Automated Workflow**: Processes multiple CSV records and generates one PDF certificate per row.
- **Simple UI**: Tkinter desktop UI for selecting template, CSV, and output folder.

## Prerequisites

- Python 3.10+

## Installation

```bash
pip install -r requirements.txt
```

## Usage

1. Run the app:
   ```bash
   python bulk_certificate_generator.py
   ```
2. Select a **PDF template** containing placeholders such as `{{name}}`, `{{course}}`, etc.
3. Select a CSV file with matching column names.
4. Select output directory.
5. Click **Generate Certificates**.

## Template Notes

- Text placeholders: put tokens like `{{name}}` in PDF text layers.
- Form fields: use form field names matching either `name` or `{{name}}`.
- CSV format remains unchanged.

## Example CSV

```csv
name,course,completion_date
John Doe,Python Programming,2024-08-01
Jane Smith,Web Development,2024-08-02
```

## Cloud / Container Deployment

### Docker

```bash
docker build -t bulk-certificate-generator .
docker run --rm -it bulk-certificate-generator
```

For serverless or web-service deployments, reuse the same PDF generation functions from `bulk_certificate_generator.py`.
