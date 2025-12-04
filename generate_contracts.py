import os

import pandas as pd
from docx import Document

import utils  # ログ出力とエラーハンドリング用


def check_files_exist(excel_path, template_path):
    """
    必要なファイルが存在するか確認します。
    """
    if not os.path.exists(excel_path):
        utils.log_error(f"Excelファイルが見つかりません: {excel_path}")
        print(f"エラー: Excelファイルが見つかりません: {excel_path}")
        return False

    if not os.path.exists(template_path):
        utils.log_error(
            f"テンプレートファイルが見つかりません: {template_path}"
        )
        print(f"エラー: テンプレートファイルが見つかりません: {template_path}")
        return False

    return True


def validate_dataframe(df):
    """
    データフレームに必要な列が存在するか確認します。
    """
    required_columns = ["property_name", "address", "amount"]
    missing_columns = [
        col for col in required_columns if col not in df.columns
    ]

    if missing_columns:
        utils.log_error(
            f"必要な列が見つかりません: {', '.join(missing_columns)}"
        )
        print(
            f"エラー: 必要な列が見つかりません: {', '.join(missing_columns)}"
        )
        return False

    return True


def process_single_contract(index, row, template_path, output_dir):
    """
    1件の契約書生成処理を行います。

    Returns:
        bool: 成功した場合は True
    """
    try:
        # テンプレートを毎回新しく読み込む
        doc = Document(template_path)

        # プレースホルダーと置換値の辞書
        replacements = {
            "{{property_name}}": str(row["property_name"]),
            "{{address}}": str(row["address"]),
            "{{amount}}": f"{row['amount']:,}",
        }

        # 段落内のテキストを置換
        for paragraph in doc.paragraphs:
            for key, value in replacements.items():
                if key in paragraph.text:
                    paragraph.text = paragraph.text.replace(key, value)

        # ドキュメントを保存
        output_filename = f"Contract_{row['property_name']}.docx"
        output_path = os.path.join(output_dir, output_filename)

        try:
            doc.save(output_path)
            print(f"成功: 契約書を生成しました: {output_path}")
            return True
        except Exception as e:
            utils.log_error(f"ファイル保存に失敗: {output_path} - {e}")
            return False

    except Exception as e:
        utils.log_error(f"行 {index} の処理中にエラーが発生: {e}")
        return False


def generate_contracts():
    """
    Excelデータを読み込み、Wordテンプレートを使用して
    複数の契約書ファイルを自動生成します。
    """
    utils.log_start("generate_contracts")

    excel_path = "data/contract_data.xlsx"
    template_path = "templates/contract_template.docx"
    output_dir = "output"

    # ファイル存在確認
    if not check_files_exist(excel_path, template_path):
        utils.log_end("generate_contracts")
        return

    # データの読み込み
    try:
        df = pd.read_excel(excel_path)
        print(f"成功: {len(df)}件のレコードを読み込みました。")
    except Exception as e:
        utils.log_error(f"Excelファイルの読み込みに失敗: {excel_path}")
        utils.handle_error(e)

    # バリデーション
    if not validate_dataframe(df):
        utils.log_end("generate_contracts")
        return

    # 各行を処理
    success_count = 0
    error_count = 0

    for index, row in df.iterrows():
        if process_single_contract(index, row, template_path, output_dir):
            success_count += 1
        else:
            error_count += 1

    print("\n--- 処理完了 ---")
    print(f"成功: {success_count}件")
    print(f"失敗: {error_count}件")

    utils.log_end("generate_contracts")


if __name__ == "__main__":
    generate_contracts()
