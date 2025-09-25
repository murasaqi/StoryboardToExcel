# ストーリーボード Excel 生成ツール

絵コンテ画像のフォルダを 1 つの Excel ブックにまとめるためのコマンドラインツールです。サブフォルダごとにワークシートを作成し、画像をセルへ埋め込み、フレーム番号を付与して確認しやすい形に整えます。

## 必要環境

- Python 3.10 以上（3.13 で動作確認）
- Python パッケージ: `Pillow`, `XlsxWriter`
  - インストール例: `python -m pip install pillow xlsxwriter`

## 同梱ファイル

- `generate_storyboard_workbook.py` – 主要な CLI スクリプト。オプションで細かな設定が可能です。
- `generate_storyboard_workbook.bat` – Windows 用ラッパー。バッチ経由で同じ引数を Python スクリプトへ渡します。

両ファイルは同じディレクトリに置いてください。

## 基本的な使い方

1. ルートフォルダの直下に、各演出（絵コンテ）単位のサブフォルダを用意します。
2. コマンドプロンプトまたは PowerShell で本ディレクトリを開きます。
3. 以下のいずれかを実行します（引数は共通です）。

```powershell
# Python で直接実行
python generate_storyboard_workbook.py --root "C:\path\to\storyboards"

# バッチファイル経由
generate_storyboard_workbook.bat --root "C:\path\to\storyboards"
```

`--output` を省略した場合は、`storyboard_collection_YYYYMMDD_HHMMSS.xlsx` が自動生成されます。

## 主なオプション

| オプション | 説明 | 既定値 |
| --- | --- | --- |
| `--root PATH` | 絵コンテフォルダのルート | カレントディレクトリ |
| `--output PATH` | 出力する Excel ファイルのパス | 自動命名 |
| `--max-width N` | 画像の最大幅 (px) | 800 |
| `--max-height N` | 画像の最大高さ (px) | 600 |
| `--row-padding N` | 行の高さに加える余白 (px) | 10 |
| `--col-padding N` | 画像列に加える余白 (px) | 20 |
| `--frame-start N` | フレーム番号の開始値 | 1 |
| `--exclude NAME ...` | 無視するサブフォルダ名（大小文字無視・複数指定可） | `Backup`, `Bakcup` |
| `--password STR` | シート保護パスワード | `lock` |
| `--no-protect` | シート保護を無効化 | 無効 |

画像は B 列に埋め込み、A 列にフレーム番号、C 列にファイル名を出力します。

## 備考

- 既定ではシート保護が有効です。`--password` で変更、`--no-protect` で無効化できます。
- 対応画像形式: `.png`, `.jpg`, `.jpeg`, `.bmp`, `.gif`, `.tif`, `.tiff`。
- 画像が存在しないフォルダは、その旨をワークシート上に表示します。
- 英語版のドキュメントは `README.md` に記載しています。

## トラブルシューティング

- `ModuleNotFoundError` が表示された場合は `pip install pillow xlsxwriter` で依存パッケージを導入してください。
- 文字化けする場合は、実行前に `chcp 65001` で UTF-8 に切り替えると改善します。
