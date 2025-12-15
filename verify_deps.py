import os

from bs4 import BeautifulSoup
from dotenv import load_dotenv

import utils


def main():
    utils.log_start("Dependency Verification")

    # 1. Verify BeautifulSoup and lxml
    print("1. BeautifulSoup4 & lxml のテスト中...")
    try:
        html_doc = "<html><body><p id='test'>Hello, World!</p></body></html>"
        soup = BeautifulSoup(html_doc, "lxml")
        text = soup.find("p", id="test").text
        if text == "Hello, World!":
            print("   [OK] BeautifulSoup & lxml は正常に動作しています。")
        else:
            print("   [FAIL] BeautifulSoup の解析結果が期待と異なります。")
    except Exception as e:
        print(f"   [ERROR] BeautifulSoup/lxml test failed: {e}")

    # 2. Verify python-dotenv
    print("2. python-dotenv のテスト中...")
    # Create a dummy .env file
    with open(".env.test", "w") as f:
        f.write("TEST_VAR=Success")

    try:
        load_dotenv(".env.test")
        val = os.getenv("TEST_VAR")
        if val == "Success":
            print("   [OK] python-dotenv は正常に動作しています。")
        else:
            print(f"   [FAIL] 環境変数が読み込めませんでした。取得値: {val}")
    except Exception as e:
        print(f"   [ERROR] python-dotenv test failed: {e}")
    finally:
        if os.path.exists(".env.test"):
            os.remove(".env.test")

    utils.log_end("Dependency Verification")


if __name__ == "__main__":
    main()
