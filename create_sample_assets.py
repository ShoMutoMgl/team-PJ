import os

import pandas as pd
from docx import Document


def create_sample_data():
    # Ensure directories exist
    os.makedirs("data", exist_ok=True)
    os.makedirs("templates", exist_ok=True)
    os.makedirs("output", exist_ok=True)

    # 1. Create Sample Excel Data
    data = {
        "property_name": [
            "グランドハイツ東京",
            "サニーサイド横浜",
            "リバーサイド大阪",
        ],
        "address": [
            "東京都千代田区1-1-1",
            "神奈川県横浜市西区2-2-2",
            "大阪府大阪市北区3-3-3",
        ],
        "amount": [100000, 150000, 80000],
    }
    df = pd.DataFrame(data)
    excel_path = "data/contract_data.xlsx"
    df.to_excel(excel_path, index=False)
    print(f"Created Excel data: {excel_path}")

    # 2. Create Sample Word Template
    doc = Document()
    doc.add_heading("賃貸借契約書", 0)

    doc.add_paragraph(
        "貸主（以下「甲」という）と借主（以下「乙」という）は、以下の通り賃貸借契約を締結する。"
    )

    doc.add_heading("第1条（物件の表示）", level=1)
    p = doc.add_paragraph()
    p.add_run("物件名: ").bold = True
    p.add_run("{{property_name}}")

    p = doc.add_paragraph()
    p.add_run("住所: ").bold = True
    p.add_run("{{address}}")

    doc.add_heading("第2条（賃料）", level=1)
    p = doc.add_paragraph()
    p.add_run("賃料: 金 ").bold = True
    p.add_run("{{amount}}")
    p.add_run(" 円")

    template_path = "templates/contract_template.docx"
    doc.save(template_path)
    print(f"Created Word template: {template_path}")


if __name__ == "__main__":
    try:
        create_sample_data()
        print("Sample assets created successfully.")
    except Exception as e:
        print(f"Error creating sample assets: {e}")
