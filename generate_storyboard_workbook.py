#!/usr/bin/env python3
"""
Generate an Excel workbook that consolidates storyboard images into per-folder sheets.

Usage:
    python generate_storyboard_workbook.py --root . --output storyboard.xlsx
"""

import argparse
import math
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Iterable, List, Sequence, Tuple

from PIL import Image
import xlsxwriter


def natural_key(path: Path) -> List[object]:
    """Sort key that keeps numeric suffixes in order."""
    parts = re.split(r"(\d+)", path.name)
    key: List[object] = []
    for part in parts:
        if part.isdigit():
            key.append(int(part))
        else:
            key.append(part.lower())
    return key


def sanitize_sheet_name(name: str, existing: Sequence[str]) -> str:
    sanitized = re.sub(r"[\\/*?:\[\]]", "_", name).strip() or "Sheet"
    if len(sanitized) > 31:
        sanitized = sanitized[:31]
    base = sanitized
    counter = 1
    while sanitized in existing:
        suffix = f"_{counter}"
        if len(base) + len(suffix) > 31:
            sanitized = f"{base[:31 - len(suffix)]}{suffix}"
        else:
            sanitized = f"{base}{suffix}"
        counter += 1
    return sanitized


def iter_target_directories(root: Path, excludes: Iterable[str]) -> List[Path]:
    normalized = {name.lower() for name in excludes}
    dirs = [p for p in root.iterdir() if p.is_dir() and p.name.lower() not in normalized]
    return sorted(dirs, key=lambda d: d.name)


def compute_scale(width: int, height: int, max_width: int, max_height: int) -> float:
    scale = 1.0
    if max_width and width > max_width:
        scale = min(scale, max_width / width)
    if max_height and height > max_height:
        scale = min(scale, max_height / height)
    return min(scale, 1.0)


def embed_images_for_folder(
    worksheet,
    files: Sequence[Path],
    frame_start: int,
    max_width: int,
    max_height: int,
    row_padding: int,
    col_padding: int,
    text_format,
    number_format,
):
    if not files:
        worksheet.set_column(0, 0, 12)
        worksheet.set_column_pixels(1, 1, 240)
        worksheet.set_column(2, 2, 45)
        worksheet.write(1, 1, "画像が見つかりませんでした", text_format)
        return 0

    precomputed: List[Tuple[Path, int, int, float]] = []
    max_scaled_width = 0

    for image_path in files:
        with Image.open(image_path) as img:
            width_px, height_px = img.size
        scale = compute_scale(width_px, height_px, max_width, max_height)
        scaled_width = int(round(width_px * scale))
        scaled_height = int(round(height_px * scale))
        if scale == 1.0:
            scaled_width = width_px
            scaled_height = height_px
        max_scaled_width = max(max_scaled_width, scaled_width)
        precomputed.append((image_path, scaled_width, scaled_height, scale))

    worksheet.set_column(0, 0, 12)
    worksheet.set_column_pixels(1, 1, max_scaled_width + col_padding)
    worksheet.set_column(2, 2, 45)

    current_row = 1
    frame_number = frame_start

    for image_path, scaled_width, scaled_height, scale in precomputed:
        worksheet.set_row_pixels(current_row, scaled_height + row_padding)
        worksheet.write_number(current_row, 0, frame_number, number_format)
        options = {
            "x_scale": scale,
            "y_scale": scale,
            "locked": True,
        }
        worksheet.embed_image(current_row, 1, str(image_path), options)
        worksheet.write(current_row, 2, image_path.name, text_format)
        frame_number += 1
        current_row += 1

    return len(files)


def generate_workbook(
    root: Path,
    output: Path,
    exclude_names: Iterable[str],
    max_width: int,
    max_height: int,
    row_padding: int,
    col_padding: int,
    frame_start: int,
    protect_password: str,
) -> Path:
    target_dirs = iter_target_directories(root, exclude_names)
    workbook = xlsxwriter.Workbook(output)

    header_center = workbook.add_format({"bold": True, "align": "center", "valign": "vcenter"})
    text_format = workbook.add_format({"valign": "top"})
    number_format = workbook.add_format({"align": "center", "valign": "vcenter"})

    sheet_names: List[str] = []
    summary: List[Tuple[str, int]] = []

    if not target_dirs:
        workbook.close()
        raise ValueError("対象フォルダが見つかりませんでした")

    for folder in target_dirs:
        files = sorted(
            [p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff"}],
            key=natural_key,
        )
        sheet_title = sanitize_sheet_name(folder.name, sheet_names)
        sheet_names.append(sheet_title)

        worksheet = workbook.add_worksheet(sheet_title)
        worksheet.write(0, 0, "フレーム数", header_center)
        worksheet.write(0, 2, "ファイル名", header_center)

        count = embed_images_for_folder(
            worksheet,
            files,
            frame_start=frame_start,
            max_width=max_width,
            max_height=max_height,
            row_padding=row_padding,
            col_padding=col_padding,
            text_format=text_format,
            number_format=number_format,
        )

        if protect_password:
            worksheet.protect(
                protect_password,
                {
                    "format_cells": False,
                    "format_columns": False,
                    "format_rows": False,
                    "insert_columns": False,
                    "insert_rows": False,
                    "delete_columns": False,
                    "delete_rows": False,
                    "objects": False,
                    "scenarios": False,
                    "sort": False,
                    "autofilter": False,
                    "pivot_tables": False,
                },
            )

        summary.append((folder.name, count))

    workbook.close()

    for folder_name, count in summary:
        print(f"{folder_name}: {count}枚")

    print(f"Workbook saved to: {output}")
    return output


def parse_args(argv: Sequence[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Storyboard image collector")
    parser.add_argument("--root", type=Path, default=Path.cwd(), help="ルートフォルダ (default: current directory)")
    parser.add_argument("--output", type=Path, help="出力Excelパス (default: 自動生成)")
    parser.add_argument("--max-width", type=int, default=800, help="画像最大幅(px)")
    parser.add_argument("--max-height", type=int, default=600, help="画像最大高さ(px)")
    parser.add_argument("--row-padding", type=int, default=10, help="行に追加する余白(px)")
    parser.add_argument("--col-padding", type=int, default=20, help="列に追加する余白(px)")
    parser.add_argument("--frame-start", type=int, default=1, help="フレーム番号の開始値")
    parser.add_argument(
        "--exclude",
        nargs="*",
        default=["backup", "bakcup"],
        help="除外するフォルダ名 (case-insensitive)",
    )
    parser.add_argument("--password", default="lock", help="シート保護パスワード (空文字で保護なし)")
    parser.add_argument("--no-protect", action="store_true", help="シート保護を行わない")
    return parser.parse_args(argv)


def main(argv: Sequence[str]) -> int:
    args = parse_args(argv)
    root = args.root.resolve()
    if not root.is_dir():
        print(f"ルートフォルダが見つかりません: {root}", file=sys.stderr)
        return 1

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output = args.output
    if output is None:
        output = root / f"storyboard_collection_{timestamp}.xlsx"
    else:
        output = output.resolve()

    password = "" if args.no_protect else args.password

    try:
        generate_workbook(
            root=root,
            output=output,
            exclude_names=args.exclude,
            max_width=args.max_width,
            max_height=args.max_height,
            row_padding=args.row_padding,
            col_padding=args.col_padding,
            frame_start=args.frame_start,
            protect_password=password,
        )
    except Exception as exc:  # noqa: BLE001
        print(f"エラーが発生しました: {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
