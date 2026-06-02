import os
import tempfile
import unittest

import fitz
import pandas as pd

from bulk_certificate_generator import process_pdf_and_generate_pdfs


class TestPDFGeneration(unittest.TestCase):
    def test_text_placeholder_replacement(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            template_path = os.path.join(temp_dir, "template.pdf")
            csv_path = os.path.join(temp_dir, "data.csv")
            output_dir = os.path.join(temp_dir, "out")
            os.makedirs(output_dir)

            doc = fitz.open()
            page = doc.new_page()
            page.insert_text((72, 100), "Certificate for {{name}}")
            doc.save(template_path)
            doc.close()

            pd.DataFrame([{"name": "Alice"}]).to_csv(csv_path, index=False)

            process_pdf_and_generate_pdfs(template_path, csv_path, output_dir)

            generated = [f for f in os.listdir(output_dir) if f.endswith("_output.pdf")]
            self.assertEqual(len(generated), 1)

            result = fitz.open(os.path.join(output_dir, generated[0]))
            text = result[0].get_text()
            result.close()

            self.assertIn("Alice", text)
            self.assertNotIn("{{name}}", text)

    def test_form_field_replacement(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            template_path = os.path.join(temp_dir, "template_form.pdf")
            csv_path = os.path.join(temp_dir, "data.csv")
            output_dir = os.path.join(temp_dir, "out")
            os.makedirs(output_dir)

            doc = fitz.open()
            page = doc.new_page()
            widget = fitz.Widget()
            widget.field_name = "name"
            widget.field_type = fitz.PDF_WIDGET_TYPE_TEXT
            widget.rect = fitz.Rect(72, 72, 300, 96)
            widget.field_value = "{{name}}"
            page.add_widget(widget)
            doc.save(template_path)
            doc.close()

            pd.DataFrame([{"name": "Bob"}]).to_csv(csv_path, index=False)

            process_pdf_and_generate_pdfs(template_path, csv_path, output_dir)

            generated = [f for f in os.listdir(output_dir) if f.endswith("_output.pdf")]
            self.assertEqual(len(generated), 1)

            result = fitz.open(os.path.join(output_dir, generated[0]))
            widgets = list(result[0].widgets())
            self.assertEqual(len(widgets), 1)
            self.assertEqual(widgets[0].field_value, "Bob")
            result.close()


if __name__ == "__main__":
    unittest.main()
