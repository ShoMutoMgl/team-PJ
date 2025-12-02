import logging
import sys
import traceback

# ロギングの設定
# INFOレベル以上のログを出力するように設定します。
# フォーマットは [日時] レベル: メッセージ とします。
logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] %(levelname)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)


def log_start(func_name):
    """
    処理の開始をログ出力します（INFOレベル）。
    システムの起動や大きな処理の節目に使用してください。

    Args:
        func_name (str): 開始する処理や関数の名前
    """
    logging.info(f"START: {func_name} を開始します。")


def log_end(func_name):
    """
    処理の終了をログ出力します（INFOレベル）。

    Args:
        func_name (str): 終了する処理や関数の名前
    """
    logging.info(f"END:   {func_name} を終了しました。")


def handle_error(e):
    """
    エラー発生時に簡潔な情報を出力して終了します。

    Args:
        e (Exception): 発生した例外オブジェクト
    """
    # エラーが発生したファイル名と行番号を特定します
    # tb_frame.f_code.co_filename でファイル名、tb_lineno で行番号が取得できます
    tb = traceback.extract_tb(e.__traceback__)[-1]
    filename = tb.filename
    lineno = tb.lineno
    error_type = type(e).__name__
    error_message = str(e)

    # エラー情報を出力
    # 初心者向け解説: エラーの原因を特定するために必要な情報を表示します。
    print("\n[エラーが発生しました]")
    print(f"ファイル: {filename}")
    print(f"行番号: {lineno}")
    print(f"種類: {error_type}")
    print(f"メッセージ: {error_message}")

    # システムを異常終了させます
    # exit(1) は「何か問題があって終了した」ことをOSに伝えます（0なら正常終了）。
    sys.exit(1)
