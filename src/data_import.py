"""Module for downloading Statistics Poland tables."""

from datetime import date, timedelta
from pathlib import Path
from shutil import move
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup as bs

RAW_DIR = Path("data/raw")
BIN_DIR = Path("data/bin")
MONTH_NAMES = {
    1: "january",
    2: "february",
    3: "march",
    4: "april",
    5: "may",
    6: "june",
    7: "july",
    8: "august",
    9: "september",
    10: "october",
    11: "november",
    12: "december",
}

BULLETIN_URL = (
    "https://bazydanych.new.stat.gov.pl/"
    "en/biuletyn-statystyczny-dlugie-szeregi"
)
BUSINESS_SITUATION_URL = (
    "https://publikacje.new.stat.gov.pl/en/publications-portal/"
    "business-tendency-2000-"
)
CONSUMER_SITUATION_URL = (
    "https://publikacje.new.stat.gov.pl/en/publications-portal/"
    "consumer-tendency-"
)


def _archive_destination(file_path: Path, bin_dir: Path) -> Path:
    suffix = "".join(file_path.suffixes)
    stem = file_path.name[:-len(suffix)] if suffix else file_path.name
    destination = bin_dir / file_path.name
    counter = 1
    while destination.exists():
        destination = bin_dir / f"{stem}_{counter}{suffix}"
        counter += 1
    return destination


def _move_raw_files_to_bin(
    raw_dir: Path = RAW_DIR,
    bin_dir: Path = BIN_DIR,
) -> None:
    if not raw_dir.exists():
        return

    raw_files = [path for path in raw_dir.iterdir() if path.is_file()]
    if not raw_files:
        return

    bin_dir.mkdir(parents=True, exist_ok=True)
    for file_path in raw_files:
        move(str(file_path), _archive_destination(file_path, bin_dir))


def _fetch_page(url: str) -> bs:
    response = requests.get(url, timeout=30)
    response.raise_for_status()
    return bs(response.text, "html.parser")


def _list_downloads(
    soup: bs,
    limit: int | None = None,
) -> list[tuple[str, str]]:
    files_list = soup.select_one("ul.files-list")
    if files_list is None:
        raise ValueError("Could not find the bulletin file list.")

    file_items = files_list.select("li")
    if limit is not None:
        file_items = file_items[-limit:]

    downloads = []
    for file_item in file_items:
        filename_element = file_item.select_one("span.filename")
        url_element = file_item.select_one("span.file-url")
        if filename_element is None or url_element is None:
            continue
        downloads.append(
            (
                filename_element.get_text(strip=True),
                url_element.get_text(strip=True),
            )
        )
    return downloads


def _download_files(
    page_url: str,
    downloads: list[tuple[str, str]],
    extension: str,
    output_dir: Path = RAW_DIR,
) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    for filename, file_url in downloads:
        response = requests.get(urljoin(page_url, file_url), timeout=60)
        response.raise_for_status()
        (output_dir / f"{filename}{extension}").write_bytes(response.content)


def _import_bulletin(
    url: str = BULLETIN_URL,
    output_dir: Path = RAW_DIR,
) -> None:
    soup = _fetch_page(url)
    _download_files(url, _list_downloads(soup), ".xlsx", output_dir)


def _monthly_page_url(
    base_url: str,
    publication_date: date,
    repeat_year: bool,
) -> str:
    month = MONTH_NAMES[publication_date.month]
    year = publication_date.year
    if repeat_year:
        ending = f"{year}-{month}-{year}"
    else:
        ending = f"{month}-{year}"
    return f"{base_url}{ending}"


def _fetch_latest_monthly_page(
    base_url: str,
    repeat_year: bool,
) -> tuple[str, bs]:
    publication_date = date.today()
    page_url = _monthly_page_url(base_url, publication_date, repeat_year)
    try:
        soup = _fetch_page(page_url)
    except requests.exceptions.HTTPError:
        previous_month = date(
            publication_date.year,
            publication_date.month,
            1,
        ) - timedelta(days=1)
        page_url = _monthly_page_url(base_url, previous_month, repeat_year)
        soup = _fetch_page(page_url)
    return page_url, soup


def _import_monthly_publication(
    base_url: str,
    extension: str,
    repeat_year: bool,
    output_dir: Path = RAW_DIR,
) -> None:
    page_url, soup = _fetch_latest_monthly_page(base_url, repeat_year)
    _download_files(
        page_url,
        _list_downloads(soup, limit=1),
        extension,
        output_dir,
    )


def importer() -> None:
    """Import data from Statistics Poland."""
    _move_raw_files_to_bin()
    _import_bulletin()
    _import_monthly_publication(
        BUSINESS_SITUATION_URL,
        ".zip",
        repeat_year=True,
    )
    _import_monthly_publication(
        CONSUMER_SITUATION_URL,
        ".xlsx",
        repeat_year=False,
    )
    print("Data import completed successfully.")


if __name__ == "__main__":
    importer()
