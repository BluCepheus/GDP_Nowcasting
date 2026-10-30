"""Loaders for monthly business-tendency workbooks."""

import re
from pathlib import Path

import pandas as pd


MANUFACTURING_LABELS = {
    "Dobra inwestycyjne Investment goods": "Investment goods",
    "Dobra konsumpcyjne Consumer goods": "Consumer goods",
    "Dobra pośrednie Intermediate goods": "Intermediate goods",
    "Energia Energy": "Energy",
    "Nietrwałe dobra konsumpcyjne Non-durable consumer goods": (
        "Non-durable consumer goods"
    ),
    (
        "Produkcja podstawowych substancji farmaceutycznych oraz leków i "
        "pozostałych wyrobów farmaceutycznych Manufacture of basic "
        "pharmaceutical products and pharmaceutical preparations"
    ): "Manufacture of basic pharmaceutical products and pharmaceutical preparations",
    "Trwałe dobra konsumpcyjne Durable consumer goods": "Durable consumer goods",
    "Żywność i napoje Food and beverages": "Food and beverages",
}


def _replace_missing_values(df: pd.DataFrame) -> pd.DataFrame:
    return df.mask(df.eq("x"), pd.NA).mask(df.eq("."), pd.NA)


def _load_construction_file(file_path: Path) -> list[pd.DataFrame]:
    frames = []
    with pd.ExcelFile(file_path) as workbook:
        for sheet_name in workbook.sheet_names:
            if not sheet_name.isdigit():
                continue
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            variable = df.iloc[3, 0]
            df = df.iloc[6:].copy()
            df.columns = [
                column if pd.notna(column) else f"Unnamed: {position}"
                for position, column in enumerate(df.iloc[0])
            ]
            df = df.iloc[1:].set_index(df.columns[0])
            df.index = pd.to_datetime(
                df.index.str.replace("M", "-", regex=False) + "-01",
                format="%Y-%m-%d",
            )
            df.columns = [
                f"Construction - {variable} - {column}"
                for column in df.columns
            ]
            frames.append(_replace_missing_values(df))
    return frames


def _load_manufacturing_file(file_path: Path) -> list[pd.DataFrame]:
    frames = []
    with pd.ExcelFile(file_path) as workbook:
        for sheet_name in workbook.sheet_names:
            if not sheet_name.isdigit():
                continue
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            variable = df.iloc[3, 0]
            df = df.iloc[4:, 1:].copy()
            df.columns = [
                column if pd.notna(column) else f"Unnamed: {position}"
                for position, column in enumerate(df.iloc[0])
            ]
            df = df.iloc[1:].copy()
            label_column = df.columns[0]
            df[label_column] = (
                df[label_column]
                .astype("string")
                .str.strip()
                .replace(MANUFACTURING_LABELS)
                .str.replace(r"^.*(?:\t+| {2,})", "", regex=True)
                .str.strip()
            )
            df = df.set_index(label_column).T
            df.index = pd.to_datetime(
                df.index.str.replace("M", "-", regex=False) + "-01",
                format="%Y-%m-%d",
            )
            df.columns = [
                f"Manufacturing - {variable} - {column}"
                for column in df.columns
            ]
            frames.append(_replace_missing_values(df))
    return frames


def _load_services_file(file_path: Path) -> list[pd.DataFrame]:
    frames = []
    with pd.ExcelFile(file_path) as workbook:
        for sheet_name in workbook.sheet_names:
            if not sheet_name.isdigit():
                continue
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            variable = df.iloc[3, 0]
            df = df.iloc[4:, 1:].copy()
            df.columns = [
                column if pd.notna(column) else f"Unnamed: {position}"
                for position, column in enumerate(df.iloc[0])
            ]
            df = df.iloc[1:].copy()
            label_column = df.columns[0]
            df[label_column] = (
                df[label_column]
                .astype("string")
                .str.strip()
                .str.replace(r"[\r\n]+", " ", regex=True)
                .str.replace(r"^.*(?:\t+| {2,})", "", regex=True)
                .str.strip()
            )
            df = df.set_index(label_column).T
            df.index = pd.to_datetime(
                df.index.str.replace("M", "-", regex=False) + "-01",
                format="%Y-%m-%d",
            )
            df.columns = [
                f"Services - {variable} - {column}"
                for column in df.columns
            ]
            frames.append(_replace_missing_values(df))
    return frames


def _clean_services_header(value: object) -> str:
    lines = (
        [line.strip() for line in value.splitlines() if line.strip()]
        if isinstance(value, str)
        else []
    )
    if not lines:
        raise ValueError(f"Empty header value: {value!r}.")
    if len(lines) > 1:
        label = lines[-1]
    else:
        label = re.sub(r"^.*(?:\t+| {2,})", "", lines[0])
    return re.sub(r"\s+", " ", label).strip()


def _parse_services_size_dates(
    date_labels: pd.Series,
    sheet_name: str,
) -> list[pd.Timestamp]:
    dates = list(
        pd.to_datetime(
            date_labels.astype("string").str.replace(
                "M", "-", regex=False
            )
            + "-01",
            format="%Y-%m-%d",
        )
    )
    duplicate_dates = pd.Index(dates).duplicated(keep=False)
    for position in range(1, len(dates) - 1):
        if (
            duplicate_dates[position]
            and dates[position] > dates[position + 1]
        ):
            inferred_date = dates[position - 1] + pd.offsets.MonthBegin(1)
            if (
                dates[position + 1]
                == inferred_date + pd.offsets.MonthBegin(1)
                and inferred_date not in dates
            ):
                dates[position] = inferred_date
    if pd.Index(dates).has_duplicates:
        raise ValueError(
            f"Duplicate dates in services size-class sheet {sheet_name!r}."
        )
    return dates


def _load_services_size_file(file_path: Path) -> list[pd.DataFrame]:
    frames = []
    with pd.ExcelFile(file_path) as workbook:
        for sheet_name in workbook.sheet_names:
            if not sheet_name.isdigit():
                continue
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            variable = df.iloc[3, 0]
            last_column = max(
                (
                    position
                    for position in range(1, df.shape[1])
                    if df.iloc[:, position].notna().any()
                ),
                default=0,
            )
            columns = []
            category = None
            for position in range(1, last_column + 1):
                category_cell = df.iloc[6, position]
                if pd.notna(category_cell) and str(category_cell).strip():
                    category = _clean_services_header(category_cell)
                    columns.append(f"Services - {variable} - {category}")
                    continue

                subcategory_cell = df.iloc[7, position]
                if (
                    category is None
                    or pd.isna(subcategory_cell)
                    or not str(subcategory_cell).strip()
                ):
                    raise ValueError(
                        f"Missing services size-class header in sheet {sheet_name!r}."
                    )
                subcategory = _clean_services_header(subcategory_cell)
                columns.append(
                    f"Services - {variable} - {category} - {subcategory}"
                )

            data = df.iloc[17:, 1:last_column + 1].copy()
            data.columns = columns
            data.index = _parse_services_size_dates(
                df.iloc[17:, 0],
                sheet_name,
            )
            frames.append(_replace_missing_values(data))
    return frames


def _load_trade_file(file_path: Path) -> list[pd.DataFrame]:
    frames = []
    with pd.ExcelFile(file_path) as workbook:
        for sheet_name in workbook.sheet_names:
            if not sheet_name.isdigit():
                continue
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            last_column = max(
                (
                    position
                    for position in range(1, df.shape[1])
                    if df.iloc[:, position].notna().any()
                ),
                default=0,
            )
            sections = df.iloc[4, 1:last_column + 1].ffill()
            columns = []
            for position, (section, label) in enumerate(
                zip(sections, df.iloc[6, 1:last_column + 1]),
                start=1,
            ):
                section_match = re.search(r"\(([^()]*)\)\s*$", str(section))
                if section_match is None or pd.isna(label):
                    raise ValueError(
                        f"Missing trade header in sheet {sheet_name!r}, "
                        f"column {position + 1}."
                    )
                section_name = section_match.group(1).strip()
                label_name = _clean_services_header(label)
                columns.append(f"{section_name} - {label_name}")

            data = df.iloc[7:, 1:last_column + 1].copy()
            data.columns = columns
            data.index = pd.to_datetime(
                df.iloc[7:, 0].astype("string").str.replace(
                    "M", "-", regex=False
                )
                + "-01",
                format="%Y-%m-%d",
            )
            frames.append(_replace_missing_values(data))
    return frames


def _load_business_tendency_file(file_path: Path) -> list[pd.DataFrame]:
    frames = []
    with pd.ExcelFile(file_path) as workbook:
        for sheet_name in workbook.sheet_names:
            if re.fullmatch(r"Tabl\. \d+", sheet_name) is None:
                continue
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            variable = str(df.iloc[1, 0]).replace("\t", "").strip()
            last_column = max(
                (
                    position
                    for position in range(1, df.shape[1])
                    if df.iloc[:, position].notna().any()
                ),
                default=0,
            )
            groups = df.iloc[3, 1:last_column + 1].ffill()
            columns = []
            for position, (group, label) in enumerate(
                zip(groups, df.iloc[4, 1:last_column + 1]),
                start=1,
            ):
                if pd.isna(group) or pd.isna(label):
                    raise ValueError(
                        f"Missing business-tendency header in sheet "
                        f"{sheet_name!r}, column {position + 1}."
                    )
                group_name = _clean_services_header(group)
                label_name = _clean_services_header(label)
                columns.append(
                    f"{variable} - {group_name} - {label_name}"
                )

            data = df.iloc[5:, 1:last_column + 1].copy()
            data.columns = columns
            data.index = pd.to_datetime(
                df.iloc[5:, 0].astype("string").str.replace(
                    "M", "-", regex=False
                )
                + "-01",
                format="%Y-%m-%d",
            )
            frames.append(_replace_missing_values(data))
    return frames
