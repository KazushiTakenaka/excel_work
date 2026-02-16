# PythonでGoogleスプレッドシートを操作しよう

このガイドでは、`gspread` ライブラリを使ってPythonからGoogleスプレッドシートを読み書きする基本的な方法を解説します。

---

## 準備：ファイル構成

プロジェクトフォルダ（例: `excel_automation`）の中に、以下のファイルがある状態を想定しています。

- `sheet_operation.py` （これから作成するPythonファイル）
- `~/.gemini/credentials.json` （[02_setup_service_account.md](./02_setup_service_account.md) で配置した共通キー）

---

## 1. 認証とスプレッドシートを開く

まず、Googleへの「認証」を行い、操作したいスプレッドシートを開くコードを書きます。

### コード例 (`sheet_operation.py`)

```python
import gspread
import os
from oauth2client.service_account import ServiceAccountCredentials

# 1. 認証の設定
# 認証に使われるスコープ（権限の範囲）を指定します
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

# 共通の場所(.gemini)にある認証情報ファイルを読み込みます
# os.path.expanduser('~') はユーザーのホームディレクトリ(C:\Users\User名)を意味します
json_path = os.path.expanduser('~/.gemini/credentials.json')

if not os.path.exists(json_path):
    print(f"エラー: 認証ファイルが見つかりません: {json_path}")
    exit()

creds = ServiceAccountCredentials.from_json_keyfile_name(json_path, scope)

# クライアントを作成して認証を通します
client = gspread.authorize(creds)

# 2. スプレッドシートを開く
# スプレッドシートの名前（タイトル）で開く場合
# ※ 注意: サービスアカウントに共有したスプレッドシートの正確な名前を指定してください
sheet_name = 'テストシート'  # ここを実際のスプレッドシート名に変更！
spreadsheet = client.open(sheet_name)

# シート（タブ）を選択する（0番目＝一番左のシート）
worksheet = spreadsheet.get_worksheet(0)

print(f"「{sheet_name}」を開きました！")
```

---

## 2. データの読み込み

シートからデータを取得する方法です。

```python
# 全データを辞書のリストとして取得（1行目をヘッダーとして扱う）
# 例: [{'名前': '田中', '年齢': 25}, {'名前': '佐藤', '年齢': 30}]
records = worksheet.get_all_records()
print("--- 全データ ---")
print(records)

# セルの値を指定して取得 (行, 列) ※1から始まります
# 例: A1セルの値を取得（1行目, 1列目）
cell_value = worksheet.cell(1, 1).value
print(f"A1セルの値: {cell_value}")

# 特定の行を丸ごと取得（例: 2行目）
row_values = worksheet.row_values(2)
print(f"2行目のデータ: {row_values}")
```

---

## 3. データの書き込み・更新

シートにデータを書き込んだり、修正する方法です。

```python
# 特定のセルを更新する (行, 列, 新しい値)
# 例: B2セル（2行目, 2列目）を「更新完了」書き換え
worksheet.update_cell(2, 2, '更新完了')
print("B2セルを更新しました")

# 新しい行を一番下に追加する
# リスト形式でデータを渡します
new_row = ['鈴木', 28, 'エンジニア']
worksheet.append_row(new_row)
print("新しい行を追加しました")
```

---

## 全体のサンプルコードまとめ

以下をコピペして `sheet_test.py` などの名前で保存し、実行してみてください。
※ `credentials.json` が同じ場所にあること、スプレッドシート名が正しいことを確認してください。

```python
import gspread
import os
import pprint
from oauth2client.service_account import ServiceAccountCredentials

# --- 設定 ---
# 共通の場所にあるキーファイルを指定
KEY_FILE = os.path.expanduser('~/.gemini/credentials.json')
SHEET_NAME = 'テストシート' # 共有したスプレッドシートの名前に変えてください

def main():
    # 認証ファイルの確認
    if not os.path.exists(KEY_FILE):
        print(f"エラー: 認証ファイルが見つかりません。")
        print(f"配置場所: {KEY_FILE}")
        return

    # 認証
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(KEY_FILE, scope)
    client = gspread.authorize(creds)

    # シートを開く
    try:
        spreadsheet = client.open(SHEET_NAME)
        worksheet = spreadsheet.get_worksheet(0)
    except Exception as e:
        print(f"エラー: スプレッドシート「{SHEET_NAME}」が見つかりません。")
        print("サービスアカウントへの共有設定と、ファイル名を確認してください。")
        return

    # 読み込みテスト
    print("現在のデータ:")
    data = worksheet.get_all_records()
    pprint.pprint(data)

    # 書き込みテスト（行の追加）
    print("行を追加します...")
    worksheet.append_row(['Pythonで追加', 'テストデータ', 12345])
    
    print("完了しました！スプレッドシートを確認してください。")

if __name__ == '__main__':
    main()
```
