from __future__ import annotations

import argparse
import csv
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


HEADER_ALIASES = {
    "title": {"title", "name", "product name", "название", "наименование"},
    "slug": {"slug", "url slug", "чпу"},
    "sku": {"sku", "артикул", "код"},
    "regular_price": {"regular_price", "regular price", "price", "цена", "стоимость"},
    "short_description": {
        "short_description",
        "short description",
        "excerpt",
        "краткое описание",
    },
    "description": {"description", "full_description", "полное описание", "описание"},
    "categories": {"categories", "category", "категории", "категория"},
    "tags": {"tags", "tag", "теги", "тег"},
    "download_file": {
        "download_file",
        "file",
        "download",
        "digital file",
        "digital file (zip)",
        "файл",
        "path",
        "путь к файлу",
    },
    "product_folder": {
        "product_folder",
        "folder",
        "folder_path",
        "product files",
        "папка",
        "папка товара",
        "путь к папке",
        "путь к файлам",
        "ссылка на папку",
    },
    "images": {"images", "image", "photos", "listing images", "картинки", "изображения", "фото"},
    "download_name": {
        "download_name",
        "file_name",
        "имя файла",
        "название файла",
    },
    "download_url": {"download_url", "url", "url файла", "file url"},
    "seo_title": {"seo_title", "meta title", "title seo", "seo title", "seo-заголовок"},
    "seo_description": {
        "seo_description",
        "meta description",
        "description seo",
        "seo description",
        "seo-описание",
    },
    "focus_keyword": {
        "focus_keyword",
        "focus keyphrase",
        "seo keyword",
        "ключевая фраза",
        "ключевое слово",
    },
    "published": {"published", "status", "опубликовано"},
}


OUTPUT_COLUMNS = [
    "Type",
    "Published",
    "Name",
    "Slug",
    "SKU",
    "Regular price",
    "Short description",
    "Description",
    "Categories",
    "Tags",
    "Images",
    "Downloadable",
    "Virtual",
    "Download 1 name",
    "Download 1 URL",
    "Meta: _yoast_wpseo_title",
    "Meta: _yoast_wpseo_metadesc",
    "Meta: _yoast_wpseo_focuskw",
]


MANIFEST_COLUMNS = [
    "Name",
    "SKU",
    "Kind",
    "Source file",
    "Resolved path",
    "Public URL",
    "Exists",
    "Notes",
]


IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".gif"}
ARCHIVE_EXTENSIONS = {".zip"}


@dataclass(frozen=True)
class ConversionResult:
    product_count: int
    missing_downloads: int
    status_workbook_path: Path | None = None
    status_saved_to_source: bool = False

    def __iter__(self):
        yield self.product_count
        yield self.missing_downloads


def normalize_header(value: Any) -> str:
    return " ".join(str(value or "").strip().lower().replace("_", " ").split())


def slugify(value: str) -> str:
    text = str(value).strip().lower()
    result = []
    previous_dash = False
    for char in text:
        if char.isalnum():
            result.append(char)
            previous_dash = False
        elif not previous_dash:
            result.append("-")
            previous_dash = True
    slug = "".join(result).strip("-")
    return slug or "product"


def split_list(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    parts: list[str] = []
    normalized = text.replace("\r", "\n").replace("|", "\n").replace(";", "\n").replace(",", "\n")
    for chunk in normalized.split("\n"):
        chunk = chunk.strip()
        if chunk:
            parts.append(chunk)
    return " | ".join(parts)


def split_raw_list(value: Any) -> list[str]:
    text = split_list(value)
    if not text:
        return []
    return [item.strip() for item in text.split("|") if item.strip()]


def as_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_price(value: Any) -> str:
    text = as_text(value)
    if not text:
        return ""
    return text.replace("$", "").replace(",", ".").strip()


def as_bool_flag(value: Any, default: str = "1") -> str:
    text = as_text(value).lower()
    if not text:
        return default
    if text in {"0", "no", "false", "off", "нет", "не"}:
        return "0"
    if text in {"1", "yes", "true", "on", "да"}:
        return "1"
    return default


def is_empty_row(values: list[Any]) -> bool:
    return not any(cell not in (None, "") for cell in values)


def published_flag(publication_status: str) -> str:
    return "0" if publication_status == "draft" else "1"


def build_header_map(headers: list[str]) -> dict[str, int]:
    normalized = {normalize_header(header): index for index, header in enumerate(headers)}
    result: dict[str, int] = {}
    for canonical, aliases in HEADER_ALIASES.items():
        for alias in aliases:
            alias_key = normalize_header(alias)
            if alias_key in normalized:
                result[canonical] = normalized[alias_key]
                break
    return result


def get_cell(row: list[Any], header_map: dict[str, int], key: str) -> Any:
    index = header_map.get(key)
    if index is None or index >= len(row):
        return None
    return row[index]


def resolve_local_path(raw_value: Any, source_path: Path | None) -> Path | None:
    text = as_text(raw_value)
    if not text:
        return None

    raw_path = Path(text)
    if not raw_path.is_absolute() and source_path is not None:
        return (source_path.parent / raw_path).resolve()
    if not raw_path.is_absolute():
        return raw_path.resolve()
    return raw_path


def build_public_url(raw_value: Any, resolved_path: Path, base_url: str | None) -> str:
    text = as_text(raw_value)
    if not base_url:
        return ""
    url_part = text.replace("\\", "/").lstrip("./")
    if not url_part or Path(text).is_absolute():
        url_part = resolved_path.name
    return f"{base_url.rstrip('/')}/{url_part}"


def resolve_asset_path(raw_value: Any, workbook_path: Path, product_folder: Path | None = None) -> Path | None:
    text = as_text(raw_value)
    if not text:
        return None
    raw_path = Path(text)
    if raw_path.is_absolute():
        return raw_path
    if product_folder is not None:
        return (product_folder / raw_path).resolve()
    return (workbook_path.parent / raw_path).resolve()


def resolve_download_reference(
    raw_value: Any,
    source_path: Path | None,
    base_url: str | None,
) -> tuple[str, str, bool]:
    raw_path = resolve_local_path(raw_value, source_path)
    if raw_path is None:
        return "", "", False

    exists = raw_path.exists()
    if base_url:
        return build_public_url(raw_value, raw_path, base_url), str(raw_path), exists
    return "", str(raw_path), exists


def find_product_assets(folder: Path) -> tuple[list[Path], Path | None]:
    if not folder.exists() or not folder.is_dir():
        return [], None
    files = [path for path in folder.iterdir() if path.is_file()]
    images = sorted(
        [path for path in files if path.suffix.lower() in IMAGE_EXTENSIONS],
        key=lambda path: path.name.lower(),
    )
    archives = sorted(
        [path for path in files if path.suffix.lower() in ARCHIVE_EXTENSIONS],
        key=lambda path: path.name.lower(),
    )
    return images, archives[0] if archives else None


def manifest_row(
    title: str,
    sku: str,
    kind: str,
    source_file: str,
    resolved_path: str,
    public_url: str,
    exists: bool,
    notes: str = "",
) -> dict[str, str]:
    return {
        "Name": title,
        "SKU": sku,
        "Kind": kind,
        "Source file": source_file,
        "Resolved path": resolved_path,
        "Public URL": public_url,
        "Exists": "yes" if exists else "no",
        "Notes": notes if notes else ("" if exists else "File not found"),
    }


def convert_workbook(
    input_path: Path,
    output_path: Path,
    manifest_path: Path,
    sheet_name: str | None = None,
    download_base_url: str | None = None,
    publication_status: str = "publish",
    update_source_status: bool = False,
) -> ConversionResult:
    try:
        workbook = load_workbook(input_path, data_only=True)
    except PermissionError as error:
        raise PermissionError(
            f"Excel-файл недоступен для чтения: {input_path}. Закройте файл в Excel и попробуйте снова."
        ) from error
    sheet = workbook[sheet_name] if sheet_name else workbook.active
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        raise ValueError("Workbook is empty")

    header_map = build_header_map([as_text(item) for item in rows[0]])
    if "title" not in header_map:
        raise ValueError("Missing required column: title")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    manifest_path.parent.mkdir(parents=True, exist_ok=True)

    product_count = 0
    missing_downloads = 0

    with output_path.open("w", newline="", encoding="utf-8-sig") as products_file, manifest_path.open(
        "w", newline="", encoding="utf-8-sig"
    ) as manifest_file:
        products_writer = csv.DictWriter(products_file, fieldnames=OUTPUT_COLUMNS)
        manifest_writer = csv.DictWriter(manifest_file, fieldnames=MANIFEST_COLUMNS)
        products_writer.writeheader()
        manifest_writer.writeheader()

        for row in rows[1:]:
            if is_empty_row(list(row)):
                continue

            title = as_text(get_cell(row, header_map, "title"))
            if not title:
                continue

            sku = as_text(get_cell(row, header_map, "sku"))
            slug = as_text(get_cell(row, header_map, "slug")) or slugify(title)
            regular_price = normalize_price(get_cell(row, header_map, "regular_price"))
            short_description = as_text(get_cell(row, header_map, "short_description"))
            description = as_text(get_cell(row, header_map, "description"))
            categories = split_list(get_cell(row, header_map, "categories"))
            tags = split_list(get_cell(row, header_map, "tags"))
            seo_title = as_text(get_cell(row, header_map, "seo_title"))
            seo_description = as_text(get_cell(row, header_map, "seo_description"))
            focus_keyword = as_text(get_cell(row, header_map, "focus_keyword"))
            download_name = as_text(get_cell(row, header_map, "download_name")) or title
            download_url = as_text(get_cell(row, header_map, "download_url"))
            explicit_download_file = get_cell(row, header_map, "download_file")
            product_folder_text = as_text(get_cell(row, header_map, "product_folder"))
            product_folder = resolve_local_path(product_folder_text, input_path)
            explicit_images = split_raw_list(get_cell(row, header_map, "images"))
            image_urls: list[str] = []

            resolved_path = ""
            exists = False
            if product_folder_text:
                folder_exists = bool(product_folder and product_folder.exists() and product_folder.is_dir())
                manifest_writer.writerow(
                    manifest_row(
                        title=title,
                        sku=sku,
                        kind="folder",
                        source_file=product_folder_text,
                        resolved_path=str(product_folder) if product_folder else "",
                        public_url="",
                        exists=folder_exists,
                    )
                )
                if not folder_exists:
                    missing_downloads += 1

                images, archive = find_product_assets(product_folder) if product_folder else ([], None)
                for image in ([] if explicit_images else images):
                    image_url = build_public_url(image.name, image, download_base_url)
                    if image_url:
                        image_urls.append(image_url)
                    manifest_writer.writerow(
                        manifest_row(
                            title=title,
                            sku=sku,
                            kind="image",
                            source_file=image.name,
                            resolved_path=str(image),
                            public_url=image_url,
                            exists=image.exists(),
                        )
                    )

                if archive and not download_url and not as_text(explicit_download_file):
                    download_name = archive.stem
                    download_url = build_public_url(archive.name, archive, download_base_url)
                    resolved_path = str(archive)
                    exists = archive.exists()
                    manifest_writer.writerow(
                        manifest_row(
                            title=title,
                            sku=sku,
                            kind="download",
                            source_file=archive.name,
                            resolved_path=resolved_path,
                            public_url=download_url,
                            exists=exists,
                        )
                    )
                elif folder_exists and not archive and not download_url:
                    missing_downloads += 1
                    manifest_writer.writerow(
                        manifest_row(
                            title=title,
                            sku=sku,
                            kind="download",
                            source_file="",
                            resolved_path=str(product_folder),
                            public_url="",
                            exists=False,
                            notes="ZIP archive not found",
                        )
                    )

            if explicit_images:
                for image_item in explicit_images:
                    image_path = resolve_asset_path(image_item, input_path, product_folder)
                    if image_path:
                        image_url = build_public_url(image_item, image_path, download_base_url)
                        if image_url:
                            image_urls.append(image_url)
                        manifest_writer.writerow(
                            manifest_row(
                                title=title,
                                sku=sku,
                                kind="image",
                                source_file=image_item,
                                resolved_path=str(image_path),
                                public_url=image_url,
                                exists=image_path.exists(),
                            )
                        )
                        if not image_path.exists():
                            missing_downloads += 1

            if not download_url:
                download_file_path = resolve_asset_path(explicit_download_file, input_path, product_folder)
                if download_file_path:
                    download_url = build_public_url(as_text(explicit_download_file), download_file_path, download_base_url)
                    resolved_path = str(download_file_path)
                    exists = download_file_path.exists()
                else:
                    download_url, resolved_path, exists = resolve_download_reference(
                        explicit_download_file,
                        input_path,
                        download_base_url,
                    )
                if resolved_path:
                    manifest_writer.writerow(
                        manifest_row(
                            title=title,
                            sku=sku,
                            kind="download",
                            source_file=as_text(explicit_download_file),
                            resolved_path=resolved_path,
                            public_url=download_url,
                            exists=exists,
                        )
                    )
                    if not exists:
                        missing_downloads += 1

            products_writer.writerow(
                {
                    "Type": "simple",
                    "Published": published_flag(publication_status),
                    "Name": title,
                    "Slug": slug,
                    "SKU": sku,
                    "Regular price": regular_price,
                    "Short description": short_description,
                    "Description": description,
                    "Categories": categories,
                    "Tags": tags,
                    "Images": ", ".join(image_urls),
                    "Downloadable": "yes",
                    "Virtual": "yes",
                    "Download 1 name": download_name,
                    "Download 1 URL": download_url,
                    "Meta: _yoast_wpseo_title": seo_title,
                    "Meta: _yoast_wpseo_metadesc": seo_description,
                    "Meta: _yoast_wpseo_focuskw": focus_keyword,
                }
            )
            product_count += 1

    status_workbook_path = None
    status_saved_to_source = False
    if update_source_status:
        status_workbook_path = update_workbook_status(
            input_path=input_path,
            sheet_name=sheet_name,
            status_value="CSV_READY",
            publication_status=publication_status,
            fallback_path=output_path.with_name(f"{input_path.stem}_with_status.xlsx"),
        )
        status_saved_to_source = status_workbook_path == input_path

    return ConversionResult(
        product_count=product_count,
        missing_downloads=missing_downloads,
        status_workbook_path=status_workbook_path,
        status_saved_to_source=status_saved_to_source,
    )


def find_or_create_column(sheet: Any, header: str, after_column: int | None = None) -> int:
    normalized_header = normalize_header(header)
    for column in range(1, sheet.max_column + 1):
        if normalize_header(sheet.cell(1, column).value) == normalized_header:
            return column
    if after_column is not None:
        sheet.insert_cols(after_column + 1)
        sheet.cell(1, after_column + 1).value = header
        return after_column + 1
    column = sheet.max_column + 1
    sheet.cell(1, column).value = header
    return column


def update_workbook_status(
    input_path: Path,
    sheet_name: str | None,
    status_value: str,
    publication_status: str,
    fallback_path: Path | None = None,
) -> Path:
    workbook = load_workbook(input_path)
    sheet = workbook[sheet_name] if sheet_name else workbook.active
    headers = [as_text(sheet.cell(1, column).value) for column in range(1, sheet.max_column + 1)]
    header_map = build_header_map(headers)
    status_column = header_map.get("published")
    if status_column is None:
        status_column = sheet.max_column
        sheet.cell(1, status_column + 1).value = "status"
        status_column += 1
    else:
        status_column += 1

    date_column = find_or_create_column(sheet, "status_date", after_column=status_column)
    mode_column = find_or_create_column(sheet, "publish_mode")
    now_text = datetime.now().strftime("%Y-%m-%d %H:%M")
    for row_number in range(2, sheet.max_row + 1):
        row_values = [sheet.cell(row_number, column).value for column in range(1, sheet.max_column + 1)]
        if is_empty_row(row_values):
            continue
        sheet.cell(row_number, status_column).value = status_value
        sheet.cell(row_number, date_column).value = now_text
        sheet.cell(row_number, mode_column).value = publication_status
    try:
        workbook.save(input_path)
        return input_path
    except PermissionError:
        if fallback_path is None:
            raise
        fallback_path.parent.mkdir(parents=True, exist_ok=True)
        workbook.save(fallback_path)
        return fallback_path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Convert an Excel workbook with digital products into WooCommerce CSV files."
    )
    parser.add_argument("--input", required=True, type=Path, help="Path to the source .xlsx workbook")
    parser.add_argument(
        "--output",
        required=True,
        type=Path,
        help="Path to the output WooCommerce CSV",
    )
    parser.add_argument(
        "--manifest",
        type=Path,
        help="Path to the download manifest CSV. Defaults to <output>/downloads_manifest.csv",
    )
    parser.add_argument("--sheet", default=None, help="Worksheet name. Defaults to the active sheet.")
    parser.add_argument(
        "--download-base-url",
        default=None,
        help="Optional public base URL for downloadable files.",
    )
    parser.add_argument(
        "--publication-status",
        choices=["draft", "publish"],
        default="publish",
        help="Use draft for unpublished WooCommerce import rows, publish for ready listings.",
    )
    parser.add_argument(
        "--update-source-status",
        action="store_true",
        help="Write CSV_READY, status_date, and publish_mode back into the source workbook.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if not args.input.exists():
        parser.error(f"Input file not found: {args.input}")

    manifest_path = args.manifest or args.output.with_name("downloads_manifest.csv")
    products, missing_downloads = convert_workbook(
        input_path=args.input,
        output_path=args.output,
        manifest_path=manifest_path,
        sheet_name=args.sheet,
        download_base_url=args.download_base_url,
        publication_status=args.publication_status,
        update_source_status=args.update_source_status,
    )

    print(f"Processed {products} products")
    print(f"Products CSV: {args.output}")
    print(f"Downloads manifest: {manifest_path}")
    if missing_downloads:
        print(f"Missing download files: {missing_downloads}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
