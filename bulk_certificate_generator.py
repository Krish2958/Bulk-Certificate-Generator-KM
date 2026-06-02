import os
from datetime import datetime

import fitz
import pandas as pd
try:
    from tkinter import Tk, filedialog, messagebox, Label, Button, Frame
except ImportError:  # pragma: no cover - depends on runtime image
    Tk = filedialog = messagebox = Label = Button = Frame = None

# Global variables to store selected paths
template_pdf_path = ""
csv_path = ""
output_dir = ""


def sanitize_filename(value):
    """Create a filesystem-safe filename fragment."""
    value = str(value).strip() or "certificate"
    safe = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in value)
    return safe[:100]


def normalize_placeholder_key(key):
    """Normalize placeholder/form-field names for matching."""
    key = str(key).strip()
    if key.startswith("{{") and key.endswith("}}"):
        key = key[2:-2]
    return key.strip().lower()


def replace_text_placeholders(page, replacements):
    """Replace text placeholders on a page using redaction + insertion."""
    for placeholder, value in replacements.items():
        matches = page.search_for(placeholder)
        for rect in matches:
            page.add_redact_annot(rect, fill=(1, 1, 1))
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
            fontsize = max(6, min(36, rect.height * 0.8))
            page.insert_text(
                fitz.Point(rect.x0, rect.y1 - 1),
                value,
                fontsize=fontsize,
                fontname="helv",
                color=(0, 0, 0),
            )


def replace_form_fields(pdf_doc, row):
    """Replace PDF form fields if present."""
    normalized_row = {normalize_placeholder_key(col): str(row[col]) for col in row.index}

    for page in pdf_doc:
        widgets = page.widgets()
        if not widgets:
            continue

        for widget in widgets:
            field_name = widget.field_name or ""
            normalized_field = normalize_placeholder_key(field_name)

            if field_name in row:
                value = str(row[field_name])
            elif normalized_field in normalized_row:
                value = normalized_row[normalized_field]
            else:
                continue

            widget.field_value = value
            widget.update()


def process_pdf_and_generate_pdfs(template_pdf_path, csv_path, output_dir):
    """Generate PDF certificates with placeholders replaced from CSV data."""
    if not os.path.isfile(template_pdf_path):
        raise FileNotFoundError("Selected PDF template does not exist.")
    if not os.path.isfile(csv_path):
        raise FileNotFoundError("Selected CSV file does not exist.")
    if not os.path.isdir(output_dir):
        raise NotADirectoryError("Selected output directory does not exist.")

    data = pd.read_csv(csv_path)

    for index, row in data.iterrows():
        replacements = {f"{{{{{col}}}}}": str(row[col]) for col in data.columns}

        pdf_doc = fitz.open(template_pdf_path)
        replace_form_fields(pdf_doc, row)

        for page in pdf_doc:
            replace_text_placeholders(page, replacements)

        timestamp = datetime.now().strftime("%Y%m%d%H%M%S%f")
        row_name = row.get("name", row.iloc[0] if len(row) else f"certificate_{index + 1}")
        filename = sanitize_filename(row_name)
        output_pdf_path = os.path.join(output_dir, f"{filename}_{timestamp}_output.pdf")

        pdf_doc.save(output_pdf_path)
        pdf_doc.close()
        print(f"Generated PDF: {output_pdf_path}")


# UI to select files and directories
def select_pdf_template():
    global template_pdf_path
    template_pdf_path = filedialog.askopenfilename(
        title="Select PDF Template",
        filetypes=[("PDF files", "*.pdf")],
    )
    if template_pdf_path:
        template_label.config(text=f"Selected: {template_pdf_path}")


def select_csv_file():
    global csv_path
    csv_path = filedialog.askopenfilename(
        title="Select CSV File",
        filetypes=[("CSV files", "*.csv")],
    )
    if csv_path:
        csv_label.config(text=f"Selected: {csv_path}")


def select_output_dir():
    global output_dir
    output_dir = filedialog.askdirectory(title="Select Output Directory")
    if output_dir:
        output_label.config(text=f"Selected: {output_dir}")


def run_process():
    try:
        if not template_pdf_path or not csv_path or not output_dir:
            raise ValueError("All inputs (template, CSV, output directory) must be selected.")
        process_pdf_and_generate_pdfs(template_pdf_path, csv_path, output_dir)
        messagebox.showinfo("Success", "PDF generation completed!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


def build_ui():
    global template_label, csv_label, output_label
    if Tk is None:
        raise RuntimeError("Tkinter is not available in this environment.")

    root = Tk()
    root.title("Bulk Certificate Generator")
    root.geometry("500x400")
    root.configure(bg="#f0f0f0")

    frame = Frame(root, bg="#ffffff", padx=20, pady=20, relief="groove", borderwidth=2)
    frame.pack(padx=20, pady=20, fill="both", expand=True)

    title_label = Label(frame, text="Bulk Certificate Generator", font=("Arial", 16, "bold"), bg="#ffffff")
    title_label.pack(pady=10)

    template_label = Label(frame, text="No PDF template selected", bg="#ffffff", fg="#333", wraplength=400)
    template_label.pack(pady=5)

    template_button = Button(
        frame,
        text="Select PDF Template",
        command=select_pdf_template,
        bg="#4CAF50",
        fg="white",
        font=("Arial", 12),
        width=25,
    )
    template_button.pack(pady=5)

    csv_label = Label(frame, text="No CSV file selected", bg="#ffffff", fg="#333", wraplength=400)
    csv_label.pack(pady=5)

    csv_button = Button(frame, text="Select CSV File", command=select_csv_file, bg="#4CAF50", fg="white", font=("Arial", 12), width=25)
    csv_button.pack(pady=5)

    output_label = Label(frame, text="No output directory selected", bg="#ffffff", fg="#333", wraplength=400)
    output_label.pack(pady=5)

    output_button = Button(frame, text="Select Output Directory", command=select_output_dir, bg="#4CAF50", fg="white", font=("Arial", 12), width=25)
    output_button.pack(pady=5)

    run_button = Button(frame, text="Generate Certificates", command=run_process, bg="#2196F3", fg="white", font=("Arial", 14, "bold"), width=30, height=2)
    run_button.pack(pady=5)

    return root


def main():
    root = build_ui()
    root.mainloop()


if __name__ == "__main__":
    main()
