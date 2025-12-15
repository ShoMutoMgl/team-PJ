import time

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

import utils


def main():
    utils.log_start("Selenium Demo")
    print("Seleniumの動作確認を開始します...")

    try:
        # Chromeドライバのセットアップ（初回はダウンロードが発生します）
        service = ChromeService(ChromeDriverManager().install())

        # ブラウザの起動オプション
        options = webdriver.ChromeOptions()
        # options.add_argument("--headless") # 画面を出したくない場合は有効化

        print("ブラウザを起動しています...")
        driver = webdriver.Chrome(service=service, options=options)

        # Python.orgにアクセス
        url = "https://www.python.org"
        print(f"URLにアクセス中: {url}")
        driver.get(url)

        # タイトルを取得して表示
        title = driver.title
        print(f"ページタイトル: {title}")

        # 少し待機（動作確認のため）
        time.sleep(3)

        # ブラウザを閉じる
        driver.quit()
        print("ブラウザを正常に閉じました。")
        utils.log_end("Selenium Demo")

    except Exception as e:
        utils.handle_error(e)


if __name__ == "__main__":
    main()
