# Storyboard to Excel Generator

Command-line tools for consolidating storyboard image folders into an Excel workbook. Each subfolder becomes a worksheet, images are resized and embedded into cells, and frames are numbered for easy review.

## Requirements

- Python 3.10 or newer (verified with 3.13)
- Python packages: `Pillow`, `XlsxWriter`
  - Install via `python -m pip install pillow xlsxwriter`

## Files

- `generate_storyboard_workbook.py` – primary CLI script with configurable options.
- `generate_storyboard_workbook.bat` – Windows helper that forwards arguments to the Python script.

Keep both files in the same directory.

## Quick Start

1. Prepare a root folder containing storyboard subfolders. Each subfolder should contain the frames for one storyboard.
2. Open Command Prompt or PowerShell in this repository directory.
3. Run either command below (arguments are interchangeable):

```powershell
# Direct Python execution
python generate_storyboard_workbook.py --root "C:\path\to\storyboards"

# Windows batch wrapper
generate_storyboard_workbook.bat --root "C:\path\to\storyboards"
```

If `--output` is omitted, the script creates `storyboard_collection_YYYYMMDD_HHMMSS.xlsx` in the root folder.

## Key Options

| Option | Description | Default |
| --- | --- | --- |
| `--root PATH` | Root folder containing storyboard subdirectories | Current directory |
| `--output PATH` | Output Excel path | Auto-generated name |
| `--max-width N` | Maximum image width (px) | 800 |
| `--max-height N` | Maximum image height (px) | 600 |
| `--row-padding N` | Extra pixels added to each row | 10 |
| `--col-padding N` | Extra pixels added to the image column | 20 |
| `--frame-start N` | Starting frame number | 1 |
| `--exclude NAME ...` | Subfolders to skip (case-insensitive, multiple allowed) | `Backup`, `Bakcup` |
| `--password STR` | Worksheet protection password | `lock` |
| `--no-protect` | Disable worksheet protection | Off |

Images are embedded in column B, frame numbers appear in column A, and filenames are written in column C.

## Notes

- Worksheet protection is enabled by default. Change the password with `--password` or disable protection with `--no-protect`.
- Supported image formats: `.png`, `.jpg`, `.jpeg`, `.bmp`, `.gif`, `.tif`, `.tiff`.
- Empty folders are noted directly on the worksheet.
- A Japanese translation of this documentation is available in `README_ja.md`.

## Troubleshooting

- `ModuleNotFoundError`: install the required packages with `pip install pillow xlsxwriter`.
- Garbled characters in the console: switch to UTF-8 with `chcp 65001` before running the script.
