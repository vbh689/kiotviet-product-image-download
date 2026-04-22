from __future__ import annotations

import argparse
import mimetypes
import sys
from pathlib import Path
from typing import Iterable
from urllib.error import HTTPError, URLError
from urllib.parse import urlparse
from urllib.request import urlopen

from openpyxl import load_workbook


DEFAULT_EXCEL_PATH = "SP.xlsx"
DEFAULT_OUTPUT_DIR = "downloaded_images"
CODE_HEADER = "Mã hàng"
IMAGE_HEADER = "Hình ảnh (url1,url2...)"
REQUEST_TIMEOUT_SECONDS = 30


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Download product images from an Excel file."
    )
    parser.add_argument(
        "excel_path",
        nargs="?",
        default=DEFAULT_EXCEL_PATH,
        help=f"Path to the Excel file. Defaults to {DEFAULT_EXCEL_PATH}.",
    )
    parser.add_argument(
        "output_dir",
        nargs="?",
        default=DEFAULT_OUTPUT_DIR,
        help=f"Directory to save images. Defaults to {DEFAULT_OUTPUT_DIR}.",
    )
    return parser.parse_args()


def wait_for_exit(message: str = "Press ENTER to exit.") -> None:
    try:
        input(message)
    except EOFError:
        pass


def fatal(message: str) -> int:
    print(f"Error: {message}", file=sys.stderr)
    print()
    wait_for_exit()
    return 1


def sanitize_filename(value: str) -> str:
    cleaned = "".join("_" if char in '<>:"/\\|?*' else char for char in value.strip())
    cleaned = cleaned.rstrip(". ")
    return cleaned or "unknown"


def split_image_urls(value: object) -> list[str]:
    if value is None:
        return []
    return [part.strip() for part in str(value).split(",") if part and part.strip()]


def load_rows(excel_path: Path) -> tuple[Iterable[tuple[object, ...]], int, int]:
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    try:
        workbook = load_workbook(excel_path, read_only=True, data_only=True)
    except Exception as exc:  # pragma: no cover - openpyxl error types vary
        raise RuntimeError(f"Unable to read Excel file '{excel_path}': {exc}") from exc

    if not workbook.sheetnames:
        raise ValueError(f"Workbook '{excel_path}' does not contain any worksheets.")

    worksheet = workbook[workbook.sheetnames[0]]
    rows = worksheet.iter_rows(values_only=True)

    try:
        header_row = list(next(rows))
    except StopIteration as exc:
        raise ValueError(f"Workbook '{excel_path}' is empty.") from exc

    missing_headers = [
        header for header in (CODE_HEADER, IMAGE_HEADER) if header not in header_row
    ]
    if missing_headers:
        joined = ", ".join(missing_headers)
        raise ValueError(f"Missing required column(s): {joined}")

    return rows, header_row.index(CODE_HEADER), header_row.index(IMAGE_HEADER)


def guess_extension(url: str, content_type: str | None) -> str:
    path = urlparse(url).path
    suffix = Path(path).suffix
    if suffix:
        return suffix.lower()

    if content_type:
        guessed = mimetypes.guess_extension(content_type.split(";", 1)[0].strip())
        if guessed:
            return guessed.lower()

    return ".jpg"


def choose_target_path(output_dir: Path, base_name: str, extension: str) -> Path:
    candidate = output_dir / f"{base_name}{extension}"
    counter = 1
    while candidate.exists():
        candidate = output_dir / f"{base_name}_{counter}{extension}"
        counter += 1
    return candidate


def build_base_name(product_code: str, image_index: int, total_images: int) -> str:
    safe_code = sanitize_filename(product_code)
    if total_images > 1:
        return f"{safe_code}_{image_index}"
    return safe_code


def download_file(url: str, destination: Path) -> None:
    with urlopen(url, timeout=REQUEST_TIMEOUT_SECONDS) as response:
        extension = guess_extension(url, response.headers.get_content_type())
        target_path = choose_target_path(destination.parent, destination.stem, extension)
        with target_path.open("wb") as output_file:
            output_file.write(response.read())


def prompt_with_default(label: str, default_value: str) -> str:
    prompt = f"{label} [{default_value}]: " if default_value else f"{label}: "
    try:
        entered_value = input(prompt).strip()
    except EOFError:
        return default_value
    return entered_value or default_value


def confirm_paths(excel_path: Path, output_dir: Path) -> tuple[Path, Path]:
    print("Product image downloader")
    print()

    selected_excel_path = excel_path
    while True:
        selected_excel_path = Path(
            prompt_with_default("Input Excel file", str(selected_excel_path))
        )

        if selected_excel_path.exists():
            break

        print(f"Warning: input file not found: {selected_excel_path}")
        print("Please press ENTER to keep the same name or type a different file path.")
        print()

    selected_output_dir = Path(
        prompt_with_default("Output directory", str(output_dir))
    )

    if selected_output_dir.exists():
        print(f"Output directory ready: {selected_output_dir.resolve()}")
    else:
        selected_output_dir.mkdir(parents=True, exist_ok=True)
        print(f"Created output directory: {selected_output_dir.resolve()}")

    print()
    wait_for_exit("Press ENTER to download.")
    return selected_excel_path, selected_output_dir


def main() -> int:
    args = parse_args()
    excel_path = Path(args.excel_path)
    output_dir = Path(args.output_dir)

    excel_path, output_dir = confirm_paths(excel_path, output_dir)

    try:
        rows, code_index, image_index = load_rows(excel_path)
    except (FileNotFoundError, RuntimeError, ValueError) as exc:
        return fatal(str(exc))

    output_dir.mkdir(parents=True, exist_ok=True)

    processed_rows = 0
    success_count = 0
    failure_count = 0

    for row_number, row in enumerate(rows, start=2):
        processed_rows += 1

        product_code_raw = row[code_index] if code_index < len(row) else None
        image_urls_raw = row[image_index] if image_index < len(row) else None

        product_code = "" if product_code_raw is None else str(product_code_raw).strip()
        image_urls = split_image_urls(image_urls_raw)

        if not product_code:
            print(f"Warning: row {row_number} skipped because '{CODE_HEADER}' is empty.")
            continue

        if not image_urls:
            print(
                f"Warning: row {row_number} ({product_code}) skipped because "
                f"'{IMAGE_HEADER}' is empty."
            )
            continue

        for image_number, url in enumerate(image_urls, start=1):
            base_name = build_base_name(product_code, image_number, len(image_urls))
            destination = output_dir / base_name
            try:
                download_file(url, destination)
                success_count += 1
                print(f"Downloaded: {product_code} <- {url}")
            except (HTTPError, URLError, TimeoutError, ValueError, OSError) as exc:
                failure_count += 1
                print(f"Failed: {product_code} <- {url} ({exc})")

    print()
    print("Download summary")
    print(f"Rows processed: {processed_rows}")
    print(f"Images downloaded: {success_count}")
    print(f"Images failed: {failure_count}")
    print(f"Output directory: {output_dir.resolve()}")
    print()
    wait_for_exit()

    return 0


if __name__ == "__main__":
    sys.exit(main())
