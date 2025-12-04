import os

import pandas as pd
from docx import Document

import utils  # ログ出力とエラーハンドリング用


def generate_contracts():
    """
    Excelデータを読み込み、Wordテンプレートを使用して
    複数の契約書ファイルを自動生成します。
    """
    # 処理開始ログ
    utils.log_start("generate_contracts")

    # ファイルパスの定義
    excel_path = "data/contract_data.xlsx"
    template_path = "templates/contract_template.docx"
    output_dir = "output"

    # --- セキュリティ: ファイルの存在確認 ---
    # 初心者向け解説: ファイルが存在しない場合は処理を中断します
    if not os.path.exists(excel_path):
        utils.log_error(f"Excelファイルが見つかりません: {excel_path}")
        print(f"エラー: Excelファイルが見つかりません: {excel_path}")
        utils.log_end("generate_contracts")
        return

    if not os.path.exists(template_path):
        utils.log_error(
            f"テンプレートファイルが見つかりません: {template_path}"
        )
        print(f"エラー: テンプレートファイルが見つかりません: {template_path}")
        utils.log_end("generate_contracts")
        return

    # --- データの読み込み ---
    # 外部ファイルの読み込みはI/O操作なので、エラーハンドリングを行います
    try:
        df = pd.read_excel(excel_path)
        print(f"成功: {len(df)}件のレコードを読み込みました。")
    except Exception as e:
        utils.log_error(f"Excelファイルの読み込みに失敗: {excel_path}")
        utils.handle_error(e)

    # --- データのバリデーション ---
    # 必要な列の存在確認
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
        utils.log_end("generate_contracts")
        return

    # --- 各行を処理して契約書を生成 ---
    success_count = 0
    error_count = 0

    for index, row in df.iterrows():
        try:
            # テンプレートを毎回新しく読み込む（前の変更を引き継がないため）
            doc = Document(template_path)

            # プレースホルダーと置換値の辞書
            # 初心者向け解説: {{property_name}} などのプレースホルダーを実際の値に置換します
            replacements = {
                "{{property_name}}": str(row["property_name"]),
                "{{address}}": str(row["address"]),
                "{{amount}}": f"{row['amount']:,}",  # カンマ区切りでフォーマット
            }

            # 段落内のテキストを置換
            for paragraph in doc.paragraphs:
                for key, value in replacements.items():
                    if key in paragraph.text:
                        # シンプルなテキスト置換
                        # 注意: 複雑なフォーマットの場合は run を使った置換が必要
                        paragraph.text = paragraph.text.replace(key, value)

            # ドキュメントを保存
            output_filename = f"Contract_{row['property_name']}.docx"
            output_path = os.path.join(output_dir, output_filename)

            # 外部ファイルへの書き込みはI/O操作なので、エラーハンドリング
            try:
                doc.save(output_path)
                print(f"成功: 契約書を生成しました: {output_path}")
                success_count += 1
            except Exception as e:
                utils.log_error(f"ファイル保存に失敗: {output_path} - {e}")
                error_count += 1

        except Exception as e:
            utils.log_error(f"行 {index} の処理中にエラーが発生: {e}")
            error_count += 1

    # 処理結果のサマリー表示
    print("\n--- 処理完了 ---")
    print(f"成功: {success_count}件")
    print(f"失敗: {error_count}件")

    # 処理終了ログ
    utils.log_end("generate_contracts")


if __name__ == "__main__":
    generate_contracts()
