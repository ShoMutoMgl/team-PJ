import os

import pandas as pd
from docx import Document


def generate_contracts():
    excel_path = "data/contract_data.xlsx"
    template_path = "templates/contract_template.docx"
    output_dir = "output"

    # Check if files exist
    if not os.path.exists(excel_path):
        print(f"Error: Excel file not found at {excel_path}")
        return
    if not os.path.exists(template_path):
        print(f"Error: Template file not found at {template_path}")
        return

    # Load data
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    print(f"Loaded {len(df)} records from Excel.")

    # Process each row
    for index, row in df.iterrows():
        try:
            # Load template for each iteration to start fresh
            doc = Document(template_path)

            # Dictionary of placeholders and values
            replacements = {
                "{{property_name}}": str(row["property_name"]),
                "{{address}}": str(row["address"]),
                "{{amount}}": f"{row['amount']:,}",  # Format with commas
            }

            # Replace text in paragraphs
            for paragraph in doc.paragraphs:
                for key, value in replacements.items():
                    if key in paragraph.text:
                        # Simple text replacement (might lose formatting if
                        # not careful, but good for prototype)
                        # For more complex replacement, we'd iterate
                        # through runs.
                        paragraph.text = paragraph.text.replace(key, value)

            # Save the document
            output_filename = f"Contract_{row['property_name']}.docx"
            output_path = os.path.join(output_dir, output_filename)
            doc.save(output_path)
            print(f"Generated: {output_path}")

        except Exception as e:
            print(f"Error processing row {index}: {e}")


if __name__ == "__main__":
    generate_contracts()
