# Storyboard to Excel Generator

Tools for consolidating storyboard image folders into a single Excel workbook, with one worksheet per storyboard, images embedded in cells, and frame numbering for quick reviews.

## Requirements

- Python 3.10+ (確認済み: 3.13)
- Python packages: `Pillow`, `XlsxWriter`
  - インストール例: `python -m pip install pillow xlsxwriter`

## Files

- `generate_storyboard_workbook.py` – メインスクリプト。CLI から実行して設定を細かく指定できます。
- `generate_storyboard_workbook.bat` – Windows 用ラッパー。バッチから同じオプションを渡せます。

両ファイルは同一ディレクトリに配置してください。

## Basic Usage

1. コマ画像が入ったフォルダ階層を用意します。（各サブフォルダが 1 作品）
2. コマンドプロンプトまたは PowerShell を開き、本ディレクトリへ移動します。
3. 次のいずれかを実行します。

```powershell
# Python 直接
python generate_storyboard_workbook.py --root "C:\path\to\storyboard"

# バッチ経由（同じ引数を使用可能）
generate_storyboard_workbook.bat --root "C:\path\to\storyboard"
```

出力ファイル名を省略した場合、`storyboard_collection_YYYYMMDD_HHMMSS.xlsx` が生成されます。

## Key Options

| オプション | 説明 | 既定値 |
| --- | --- | --- |
| `--root PATH` | 絵コンテフォルダのルート | カレントディレクトリ |
| `--output PATH` | 出力 Excel パス | 自動命名 |
| `--max-width N` | 画像最大幅 (px) | 800 |
| `--max-height N` | 画像最大高さ (px) | 600 |
| `--row-padding N` | 行高さに加算する余白 (px) | 10 |
| `--col-padding N` | 列幅に加算する余白 (px) | 20 |
| `--frame-start N` | フレーム番号開始値 | 1 |
| `--exclude NAME ...` | 除外するサブフォルダ名（大文字小文字無視、複数可） | `Backup`, `Bakcup` |
| `--password STR` | シート保護パスワード | `lock` |
| `--no-protect` | シート保護を無効化 | 指定なし |

画像はリサイズ後にセルへ埋め込まれ、A 列にフレーム番号、B 列に画像、C 列にファイル名を配置します。

## Notes

- シート保護は既定で有効（パスワード `lock`）。変更したい場合は `--password` または `--no-protect` を利用してください。
- 画像ファイルは `.png`, `.jpg`, `.jpeg`, `.bmp`, `.gif`, `.tif`, `.tiff` をサポートします。
- フォルダ内に画像が無い場合、その旨をシートに記載します。

## Troubleshooting

- `ModuleNotFoundError` が出た場合は依存パッケージが未インストールです。`pip install pillow xlsxwriter` を実行してください。
- 文字化けする場合は端末のコードページを UTF-8 (`chcp 65001`) に切り替えると改善します。

