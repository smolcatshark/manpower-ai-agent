import hashlib
import io
import json
import re
from typing import Any

import pandas as pd
import pdfplumber
import streamlit as st


st.set_page_config(
    page_title="Manpower AI Agent",
    page_icon="🏗️",
    layout="wide",
)

REQUIRED_TRADES = ("AC", "EL", "FS", "PD")
APP_VERSION = "0.4.2 — Location sorting"
PARSER_MODE_OPTIONS = {
    "自動偵測": "auto",
    "Location＋Manpower數字欄": "numeric_table",
    "工人姓名表（每列1人）": "worker_table",
    "工作描述中的X人／姓名數量": "description",
    "只使用報告總人數": "total_only",
}


# =========================================================
# General helper functions
# =========================================================


def text_to_list(value: str) -> list[str]:
    return [
        item.strip()
        for item in value.replace("，", ",").split(",")
        if item.strip()
    ]


def normalize_floor(value: str) -> str:
    original = value.strip().upper()
    compact = re.sub(r"\s+", "", original)

    if re.fullmatch(r"G/?F", compact):
        return "GF"

    basement_match = re.fullmatch(r"B(\d+)/?F?", compact)
    if basement_match:
        return f"B{int(basement_match.group(1))}"

    floor_match = re.fullmatch(r"(\d+)/?F?", compact)
    if floor_match:
        return f"{int(floor_match.group(1))}F"

    return original


def floor_sort_key(value: str) -> tuple[int, str]:
    match = re.fullmatch(r"(\d+)F", value)
    if match:
        return int(match.group(1)), value
    return 9999, value


def calculate_file_hash(file_bytes: bytes) -> str:
    return hashlib.sha256(file_bytes).hexdigest()


def detect_trade_from_filename(filename: str) -> str | None:
    filename_upper = filename.upper()
    trade_patterns = {
        "AC": [r"(^|[^A-Z0-9])AC([^A-Z0-9]|$)", r"MVAC"],
        "EL": [r"(^|[^A-Z0-9])EL([^A-Z0-9]|$)", r"ELECTRICAL"],
        "FS": [r"(^|[^A-Z0-9])FS([^A-Z0-9]|$)", r"FIRE[\s_-]*SERVICE"],
        "PD": [r"(^|[^A-Z0-9])PD([^A-Z0-9]|$)", r"PLUMBING", r"DRAINAGE"],
    }

    for trade, patterns in trade_patterns.items():
        if any(re.search(pattern, filename_upper) for pattern in patterns):
            return trade

    return None


def format_file_size(size_bytes: int) -> str:
    if size_bytes < 1024:
        return f"{size_bytes} B"
    if size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    return f"{size_bytes / (1024 * 1024):.1f} MB"


def clean_cell(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def location_sort_key(location: str) -> tuple[int, int, int, str]:
    """Sort locations in site order rather than alphabetically.

    Order:
    1. T1, T2, T3... with floors from highest to lowest
    2. Podium with numbered floors from highest to lowest, then GF
    3. Basement from B1 downward to B2, B3, B4...
    4. Special and uncertain locations
    """
    text = clean_cell(location).upper()

    tower_floor_match = re.fullmatch(
        r"T\s*(\d+)\s*/\s*(\d+)F",
        text,
    )
    if tower_floor_match:
        tower_number = int(tower_floor_match.group(1))
        floor_number = int(tower_floor_match.group(2))
        return (0, tower_number, -floor_number, text)

    podium_floor_match = re.fullmatch(
        r"PODIUM\s*/\s*(GF|\d+F)",
        text,
    )
    if podium_floor_match:
        floor = podium_floor_match.group(1)
        floor_rank = (
            0
            if floor == "GF"
            else -int(re.search(r"\d+", floor).group())
        )
        return (1, 0, floor_rank, text)

    basement_floor_match = re.fullmatch(
        r"BASEMENT\s*/\s*B(\d+)",
        text,
    )
    if basement_floor_match:
        basement_number = int(basement_floor_match.group(1))
        return (2, 0, basement_number, text)

    tower_unspecified_match = re.fullmatch(
        r"T\s*(\d+)\s*/\s*FLOOR\s*U",
        text,
    )
    if tower_unspecified_match:
        return (
            4,
            int(tower_unspecified_match.group(1)),
            0,
            text,
        )

    if text == "PODIUM / FLOOR U":
        return (5, 0, 0, text)

    if text == "BASEMENT / FLOOR U":
        return (6, 0, 0, text)

    cross_tower_match = re.match(
        r"CROSS-FLOOR\s*/\s*T\s*(\d+)",
        text,
    )
    if cross_tower_match:
        return (
            7,
            int(cross_tower_match.group(1)),
            0,
            text,
        )

    if text.startswith("CROSS-FLOOR"):
        return (8, 0, 0, text)

    if text.startswith("DISTRIBUTION U"):
        return (9, 0, 0, text)

    if text.startswith("UNSPECIFIED TOWER"):
        return (10, 0, 0, text)

    if text == "UNSPECIFIED":
        return (11, 0, 0, text)

    # Roof and other configured special locations come after the main zones.
    return (3, 0, 0, text)


def build_location_summary(
    detail_df: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Merge manpower by date, standard location and trade."""
    empty_columns = [
        "日期",
        "標準位置",
        "AC",
        "EL",
        "FS",
        "PD",
        "合計",
    ]

    if detail_df.empty:
        empty_df = pd.DataFrame(columns=empty_columns)
        return empty_df, empty_df.copy()

    working_df = detail_df.copy()

    required_columns = {
        "日期": "",
        "工種": "",
        "標準位置": "Unspecified",
        "人數": 0,
        "位置狀態": "需人工確認位置",
    }

    for column_name, default_value in required_columns.items():
        if column_name not in working_df.columns:
            working_df[column_name] = default_value

    working_df["日期"] = (
        working_df["日期"]
        .fillna("")
        .astype(str)
        .str.strip()
    )
    working_df["工種"] = (
        working_df["工種"]
        .fillna("")
        .astype(str)
        .str.upper()
        .str.strip()
    )
    working_df["標準位置"] = (
        working_df["標準位置"]
        .fillna("Unspecified")
        .astype(str)
        .str.strip()
        .replace("", "Unspecified")
    )
    working_df["位置狀態"] = (
        working_df["位置狀態"]
        .fillna("需人工確認位置")
        .astype(str)
        .str.strip()
    )
    working_df["人數"] = pd.to_numeric(
        working_df["人數"],
        errors="coerce",
    ).fillna(0)

    working_df = working_df[
        working_df["人數"] > 0
    ].copy()

    def create_summary(source_df: pd.DataFrame) -> pd.DataFrame:
        if source_df.empty:
            return pd.DataFrame(columns=empty_columns)

        grouped_df = (
            source_df.groupby(
                ["日期", "標準位置", "工種"],
                dropna=False,
                as_index=False,
            )["人數"]
            .sum()
        )

        pivot_df = grouped_df.pivot_table(
            index=["日期", "標準位置"],
            columns="工種",
            values="人數",
            aggfunc="sum",
            fill_value=0,
        ).reset_index()

        pivot_df.columns.name = None

        for trade in REQUIRED_TRADES:
            if trade not in pivot_df.columns:
                pivot_df[trade] = 0

        for trade in REQUIRED_TRADES:
            pivot_df[trade] = (
                pd.to_numeric(
                    pivot_df[trade],
                    errors="coerce",
                )
                .fillna(0)
                .astype(int)
            )

        pivot_df["合計"] = pivot_df[
            list(REQUIRED_TRADES)
        ].sum(axis=1)

        pivot_df = pivot_df[
            [
                "日期",
                "標準位置",
                *REQUIRED_TRADES,
                "合計",
            ]
        ]

        ordered_index = sorted(
            pivot_df.index,
            key=lambda index: (
                str(pivot_df.at[index, "日期"]),
                location_sort_key(
                    str(pivot_df.at[index, "標準位置"])
                ),
            ),
        )

        return pivot_df.loc[
            ordered_index
        ].reset_index(drop=True)

    confirmed_mask = (
        working_df["位置狀態"] == "已解析"
    )

    confirmed_summary = create_summary(
        working_df[confirmed_mask]
    )
    review_summary = create_summary(
        working_df[~confirmed_mask]
    )

    return confirmed_summary, review_summary


# =========================================================
# PDF and EL parsing functions
# =========================================================


def extract_pdf_text(file_bytes: bytes) -> str:
    page_texts: list[str] = []

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            text = page.extract_text(x_tolerance=2, y_tolerance=3) or ""
            page_texts.append(f"--- Page {page_number} ---\n{text}")

    return "\n\n".join(page_texts)


def find_first(pattern: str, text: str, flags: int = 0) -> str | None:
    match = re.search(pattern, text, flags)
    return match.group(1).strip() if match else None


def extract_header_data(text: str) -> dict[str, Any]:
    date = find_first(r"Date:\s*(\d{4}-\d{2}-\d{2})", text, re.I)
    trade = find_first(r"Trade:\s*([^\n]+?)(?:\s+Ref\.|\n)", text, re.I)
    reference = find_first(r"Ref\.\s*No\.:?\s*([^\s\n]+)", text, re.I)
    worker_text = find_first(r"\bWorker\s+(\d+)\b", text, re.I)
    total_text = find_first(
        r"Today\s+Total\s+Manpower\s*\*?\s*(\d+)\b",
        text,
        re.I,
    )

    worker = int(worker_text) if worker_text is not None else None
    total = int(total_text) if total_text is not None else None

    return {
        "date": date,
        "trade": trade,
        "reference": reference,
        "worker": worker,
        "today_total": total,
        "management_staff": (
            total - worker
            if total is not None and worker is not None
            else None
        ),
    }


def extract_site_work_table_rows(file_bytes: bytes) -> list[dict[str, str]]:
    extracted_rows: list[dict[str, str]] = []

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables() or []

            for table in tables:
                if not table:
                    continue

                cleaned_table = [
                    [clean_cell(cell) for cell in row]
                    for row in table
                    if row
                ]

                header_index = None
                description_index = None
                location_index = None
                item_index = None

                for row_index, row in enumerate(cleaned_table):
                    upper_row = [cell.upper() for cell in row]

                    description_candidates = [
                        index
                        for index, cell in enumerate(upper_row)
                        if "DESCRIPTION" in cell
                    ]
                    location_candidates = [
                        index
                        for index, cell in enumerate(upper_row)
                        if "LOCATION" in cell
                    ]

                    if description_candidates and location_candidates:
                        header_index = row_index
                        description_index = description_candidates[0]
                        location_index = location_candidates[0]
                        item_candidates = [
                            index
                            for index, cell in enumerate(upper_row)
                            if cell == "ITEM"
                        ]
                        item_index = item_candidates[0] if item_candidates else 0
                        break

                if header_index is None:
                    continue

                for row in cleaned_table[header_index + 1 :]:
                    required_length = max(
                        item_index or 0,
                        description_index or 0,
                        location_index or 0,
                    ) + 1

                    if len(row) < required_length:
                        row += [""] * (required_length - len(row))

                    item = row[item_index or 0]
                    description = row[description_index or 0]
                    location = row[location_index or 0]

                    if not description and not location:
                        continue

                    if item and not re.fullmatch(r"1\.\d+", item):
                        continue

                    extracted_rows.append(
                        {
                            "item": item,
                            "description": description,
                            "location": location,
                        }
                    )

    return extracted_rows


def extract_fallback_detail_rows(text: str) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    seen_lines: set[str] = set()

    for line in text.splitlines():
        cleaned = clean_cell(line)

        if not cleaned or cleaned in seen_lines:
            continue

        has_explicit_people = bool(re.search(r"\d+\s*人", cleaned))
        has_name_list = bool(
            re.search(r"[：:]\s*[\u3400-\u9fff]{2,4}(?:\s*[,，、]\s*[\u3400-\u9fff]{2,4})+", cleaned)
        )

        if not has_explicit_people and not has_name_list:
            continue

        if "MANPOWER" in cleaned.upper() or "TOTAL" in cleaned.upper():
            continue

        seen_lines.add(cleaned)
        rows.append(
            {
                "item": "",
                "description": cleaned,
                "location": "",
            }
        )

    return rows


def parse_description_manpower(description: str) -> tuple[int | None, str]:
    explicit_counts = [
        int(value)
        for value in re.findall(r"(\d+)\s*人", description)
    ]

    if explicit_counts:
        return sum(explicit_counts), "報告列明人數"

    parts = re.split(r"[：:]", description, maxsplit=1)
    if len(parts) != 2:
        return None, "無法計算"

    possible_names = re.split(r"[,，、]", parts[1])
    normalized_names = [
        re.sub(r"\s+", "", name.strip())
        for name in possible_names
    ]
    valid_names = [
        name
        for name in normalized_names
        if re.fullmatch(r"[\u3400-\u9fff]{2,4}", name)
    ]

    if valid_names:
        return len(valid_names), "按姓名數量計算"

    return None, "無法計算"


def extract_towers(text: str) -> list[str]:
    normalized = text.upper().replace("Ｔ", "T")
    tower_numbers: set[int] = set()

    for group in re.findall(
        r"((?:\d+\s*[,，/&、]\s*)+\d+)\s*座",
        normalized,
    ):
        for number in re.findall(r"\d+", group):
            tower_numbers.add(int(number))

    for group in re.findall(
        r"T\s*(\d+(?:\s*[,/&、]\s*\d+)+)",
        normalized,
    ):
        for number in re.findall(r"\d+", group):
            tower_numbers.add(int(number))

    for number in re.findall(r"(?<!\d)(\d+)\s*座", normalized):
        tower_numbers.add(int(number))

    for number in re.findall(r"\bT\s*(\d+)\b", normalized):
        tower_numbers.add(int(number))

    return [f"T{number}" for number in sorted(tower_numbers)]


def normalize_report_location(value: str) -> str:
    """Normalise floor notation without collapsing multi-floor references."""
    text = clean_cell(value).upper()

    if not text:
        return ""

    text = (
        text.replace("地下", "GF")
        .replace("樓", "F")
        .replace("層", "F")
        .replace("－", "-")
        .replace("–", "-")
        .replace("—", "-")
        .replace("，", ",")
        .replace("、", ",")
        .replace("＆", "&")
    )

    # Convert G/F even when it appears inside a list such as G/F, 1F.
    text = re.sub(r"(?<![A-Z0-9])G\s*/\s*F(?![A-Z0-9])", "GF", text)

    # Convert B4/F and 6/F forms.
    text = re.sub(r"B\s*(\d+)\s*/\s*F", r"B\1", text)
    text = re.sub(r"(?<![A-Z])(?<!\d)(\d+)\s*/\s*F", r"\1F", text)

    text = re.sub(r"\s+", "", text)

    if text in {"N/A", "NA", "NIL", "NONE"}:
        return "NA"

    basement = re.fullmatch(r"B(\d+)F?", text)
    if basement:
        return f"B{int(basement.group(1))}"

    numbered_floor = re.fullmatch(r"(\d+)F", text)
    if numbered_floor:
        return f"{int(numbered_floor.group(1))}F"

    if text in {"G", "G/F", "GF"}:
        return "GF"

    return text


def extract_floor_tokens(text: str) -> list[str]:
    """Extract every floor token, including compressed lists such as 5,8F."""
    normalized = normalize_report_location(text)

    if normalized in {"", "NA"}:
        return []

    tokens: list[str] = []

    def add_floor(floor: str) -> None:
        if floor not in tokens:
            tokens.append(floor)

    if re.search(r"(?<![A-Z0-9])GF(?![A-Z0-9])", normalized):
        add_floor("GF")

    for number in re.findall(r"B(\d+)", normalized):
        add_floor(f"B{int(number)}")

    # Compressed forms where the final F applies to the whole list:
    # 5,8F / 2&3F / 2/3F. Parse these first to preserve source order.
    compressed_groups = re.findall(
        r"(?<![A-Z0-9])((?:\d+[,/&+]\s*)+\d+)F(?![A-Z0-9])",
        normalized,
    )
    for group in compressed_groups:
        for number in re.findall(r"\d+", group):
            add_floor(f"{int(number)}F")

    # Explicit forms: 2F, 6F, 25F.
    for number in re.findall(
        r"(?<![A-Z0-9])(\d+)F(?![A-Z0-9])",
        normalized,
    ):
        add_floor(f"{int(number)}F")

    return tokens


def build_location_result(
    description: str,
    location_raw: str,
    config: dict[str, Any],
) -> tuple[str, str, str, str]:
    canonical_location = normalize_report_location(location_raw)
    combined_text = f"{description} {location_raw}"
    towers = extract_towers(combined_text)
    floors = extract_floor_tokens(location_raw)

    if not floors:
        floors = extract_floor_tokens(description)

    tower_text = ", ".join(towers) if towers else ""
    floor_text = ", ".join(floors) if floors else ""

    if canonical_location == "NA" or not floors:
        return "Unspecified", tower_text, floor_text, "需人工確認位置"

    is_cross_floor = (
        len(floors) > 1
        or any(separator in canonical_location for separator in ["-", "+", "至", "&"])
    )

    if is_cross_floor:
        readable_floors = " + ".join(floors)
        if len(towers) == 1:
            review_location = (
                f"Cross-floor / {towers[0]} / {readable_floors}"
            )
        elif len(towers) > 1:
            review_location = (
                f"Distribution U / {', '.join(towers)} / "
                f"{readable_floors}"
            )
        else:
            review_location = f"Cross-floor / {readable_floors}"

        return (
            review_location,
            tower_text,
            floor_text,
            "跨樓層／多位置，保留於人工確認區",
        )

    floor = floors[0]
    podium_floors = set(config.get("podium_floors", []))
    basement_floors = set(config.get("basement_floors", []))
    tower_floors = set(config.get("tower_floors", []))
    special_locations = set(config.get("special_locations", []))
    merge_podium = bool(config.get("merge_podium_tower_references", True))

    if floor in podium_floors and merge_podium:
        return f"Podium / {floor}", tower_text, floor_text, "已解析"

    if floor in basement_floors or floor.startswith("B"):
        return f"Basement / {floor}", tower_text, floor_text, "已解析"

    if floor in podium_floors:
        if towers:
            return f"{', '.join(towers)} / {floor}", tower_text, floor_text, "已解析"
        return f"Podium / {floor}", tower_text, floor_text, "已解析"

    if floor in tower_floors:
        if towers:
            return f"{', '.join(towers)} / {floor}", tower_text, floor_text, "已解析"
        return f"Unspecified Tower / {floor}", tower_text, floor_text, "需人工確認樓座"

    if floor in special_locations:
        return floor, tower_text, floor_text, "已解析"

    fallback = canonical_location or clean_cell(combined_text)
    return fallback, tower_text, floor_text, "需人工確認位置"


def extract_worker_table_rows(file_bytes: bytes) -> list[dict[str, str]]:
    """Extract rows from a Worker table containing Item/No., Name and Location."""
    worker_rows: list[dict[str, str]] = []

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables() or []

            for table in tables:
                if not table:
                    continue

                cleaned_table = [
                    [clean_cell(cell) for cell in row]
                    for row in table
                    if row
                ]

                header_index = None
                item_index = None
                name_index = None
                location_index = None

                for row_index, row in enumerate(cleaned_table):
                    upper_row = [cell.upper() for cell in row]

                    name_candidates = [
                        index
                        for index, cell in enumerate(upper_row)
                        if cell == "NAME" or cell.endswith(" NAME")
                    ]
                    location_candidates = [
                        index
                        for index, cell in enumerate(upper_row)
                        if "LOCATION" in cell
                    ]

                    if not name_candidates or not location_candidates:
                        continue

                    header_index = row_index
                    name_index = name_candidates[0]
                    location_index = location_candidates[0]

                    item_candidates = [
                        index
                        for index, cell in enumerate(upper_row)
                        if cell in {"ITEM", "NO", "NO.", "NUMBER"}
                    ]
                    item_index = item_candidates[0] if item_candidates else 0
                    break

                if header_index is None:
                    continue

                for row in cleaned_table[header_index + 1 :]:
                    required_length = max(
                        item_index or 0,
                        name_index or 0,
                        location_index or 0,
                    ) + 1

                    if len(row) < required_length:
                        row += [""] * (required_length - len(row))

                    item = row[item_index or 0]
                    name = row[name_index or 0]
                    location = row[location_index or 0]

                    if not item and not name and not location:
                        continue

                    joined = " ".join(row).upper()
                    if "SUBMITTED BY" in joined or "SITE PHOTOS" in joined:
                        break

                    if name.upper() in {"NAME", "NIL", "NONE"}:
                        continue

                    # Worker rows normally use a plain integer item number.
                    if item and not re.fullmatch(r"\d+", item):
                        continue

                    if not name:
                        continue

                    worker_rows.append(
                        {
                            "item": item,
                            "name": name,
                            "location": location,
                        }
                    )

    return worker_rows


def build_worker_location_result(
    location_raw: str,
    config: dict[str, Any],
) -> tuple[str, str, str, str]:
    """Standardise worker-table locations without inventing a floor split."""
    original = clean_cell(location_raw)
    upper = original.upper().replace("＆", "&")

    if not original or upper in {"N/A", "NA", "NIL", "NONE", "ALL SITE", "SITE"}:
        return "Unspecified", "", "", "需人工確認位置"

    # Exact tower/floor or exact floor references can use the normal parser.
    floors = extract_floor_tokens(original)
    towers = extract_towers(original)

    if floors:
        return build_location_result(original, original, config)

    has_basement = "BASEMENT" in upper
    has_podium = "PODIUM" in upper
    tower_word_only = upper.strip() in {"TOWER", "TOWERS"}

    components: list[str] = []
    if has_basement:
        components.append("Basement")
    if has_podium:
        components.append("Podium")
    components.extend(towers)

    # Remove duplicates but keep a stable readable order.
    unique_components: list[str] = []
    for component in components:
        if component not in unique_components:
            unique_components.append(component)

    tower_text = ", ".join(towers)

    if tower_word_only:
        return "Unspecified Tower / Floor U", "", "", "需人工確認樓座及樓層"

    if len(unique_components) > 1:
        label = " + ".join(unique_components)
        return (
            f"Distribution U / {label}",
            tower_text,
            "",
            "位置分布未指定，不作假設分配",
        )

    if unique_components == ["Basement"]:
        return (
            "Basement / Floor U",
            "",
            "",
            "需人工確認樓層",
        )

    if unique_components == ["Podium"]:
        return (
            "Podium / Floor U",
            "",
            "",
            "需人工確認樓層",
        )

    if len(towers) == 1 and len(unique_components) == 1:
        return (
            f"{towers[0]} / Floor U",
            towers[0],
            "",
            "需人工確認樓層",
        )

    return original, tower_text, "", "需人工確認位置"


def analyse_worker_table_pdf(
    file_data: dict[str, Any],
    config: dict[str, Any],
    trade: str,
) -> dict[str, Any]:
    """Analyse FS or PD reports using the Worker attendance table."""
    raw_text = extract_pdf_text(file_data["bytes"])
    header = extract_header_data(raw_text)
    worker_rows = extract_worker_table_rows(file_data["bytes"])

    detail_rows: list[dict[str, Any]] = []

    for row in worker_rows:
        name = clean_cell(row.get("name", ""))
        location_raw = clean_cell(row.get("location", ""))
        standard_location, towers, floors, status = build_worker_location_result(
            location_raw,
            config,
        )

        detail_rows.append(
            {
                "日期": header["date"] or "",
                "工種": trade,
                "文件": file_data["name"],
                "Item": row.get("item", ""),
                "工作描述": "Worker attendance",
                "工人姓名": name,
                "原始位置": location_raw,
                "樓座": towers,
                "樓層": floors,
                "標準位置": standard_location,
                "人數": 1,
                "人數計算方法": "Worker表每列1人",
                "位置狀態": status,
            }
        )

    extracted_worker_total = len(detail_rows)
    worker = header["worker"]

    if worker is None:
        reconciliation = "未能讀取 Worker"
    elif extracted_worker_total == worker:
        reconciliation = "一致"
    else:
        reconciliation = f"不一致，相差 {worker - extracted_worker_total:+d}"

    names = [row["工人姓名"] for row in detail_rows if row["工人姓名"]]
    duplicate_names = sorted(
        {
            name
            for name in names
            if names.count(name) > 1
        }
    )

    summary = {
        "文件": file_data["name"],
        "日期": header["date"] or "未讀取",
        "工種": trade,
        "Ref. No.": header["reference"] or "未讀取",
        "Today Total Manpower": header["today_total"],
        "Worker": worker,
        "管理及技術人員": header["management_staff"],
        "工作明細人數合計": extracted_worker_total,
        "Worker核對": reconciliation,
        "明細提取方式": "Worker出勤表",
        "明細列數": len(detail_rows),
        "重複姓名": ", ".join(duplicate_names) if duplicate_names else "沒有",
    }

    return {
        "summary": summary,
        "details": detail_rows,
        "raw_text": raw_text,
    }


def analyse_el_pdf(
    file_data: dict[str, Any],
    config: dict[str, Any],
) -> dict[str, Any]:
    raw_text = extract_pdf_text(file_data["bytes"])
    header = extract_header_data(raw_text)
    site_rows = extract_site_work_table_rows(file_data["bytes"])

    extraction_method = "PDF表格"
    if not site_rows:
        site_rows = extract_fallback_detail_rows(raw_text)
        extraction_method = "文字後備規則"

    detail_rows: list[dict[str, Any]] = []

    for row in site_rows:
        description = clean_cell(row.get("description", ""))
        location_raw = clean_cell(row.get("location", ""))

        manpower, manpower_method = parse_description_manpower(description)
        standard_location, towers, floors, status = build_location_result(
            description,
            location_raw,
            config,
        )

        if manpower is None:
            status = "需人工確認人數"

        detail_rows.append(
            {
                "日期": header["date"] or "",
                "工種": "EL",
                "文件": file_data["name"],
                "Item": row.get("item", ""),
                "工作描述": description,
                "工人姓名": "",
                "原始位置": location_raw,
                "樓座": towers,
                "樓層": floors,
                "標準位置": standard_location,
                "人數": manpower,
                "人數計算方法": manpower_method,
                "位置狀態": status,
            }
        )

    extracted_worker_total = sum(
        int(row["人數"])
        for row in detail_rows
        if row["人數"] is not None
    )

    worker = header["worker"]
    if worker is None:
        reconciliation = "未能讀取 Worker"
    elif extracted_worker_total == worker:
        reconciliation = "一致"
    else:
        reconciliation = f"不一致，相差 {worker - extracted_worker_total:+d}"

    summary = {
        "文件": file_data["name"],
        "日期": header["date"] or "未讀取",
        "工種": header["trade"] or "EL",
        "Ref. No.": header["reference"] or "未讀取",
        "Today Total Manpower": header["today_total"],
        "Worker": worker,
        "管理及技術人員": header["management_staff"],
        "工作明細人數合計": extracted_worker_total,
        "Worker核對": reconciliation,
        "明細提取方式": extraction_method,
        "明細列數": len(detail_rows),
        "重複姓名": "不適用",
    }

    return {
        "summary": summary,
        "details": detail_rows,
        "raw_text": raw_text,
    }


# =========================================================
# Flexible multi-format manpower parsing
# =========================================================


def parse_integer_cell(value: Any) -> int | None:
    """Read a positive whole-number manpower value from one table cell."""
    text = clean_cell(value)
    if not text:
        return None

    match = re.fullmatch(r"\s*(\d+)\s*(?:人|PERSONS?|WORKERS?|PAX)?\s*", text, re.I)
    if not match:
        return None

    return int(match.group(1))


def extract_numeric_manpower_table_rows(
    file_bytes: bytes,
) -> list[dict[str, Any]]:
    """Extract rows from tables containing Location and Manpower-like columns."""
    extracted_rows: list[dict[str, Any]] = []

    manpower_keywords = (
        "MANPOWER",
        "NO. OF WORKER",
        "NO OF WORKER",
        "WORKER COUNT",
        "WORKERS",
        "PERSONS",
        "HEADCOUNT",
        "QTY",
        "QUANTITY",
    )
    location_keywords = (
        "LOCATION",
        "WORK LOCATION",
        "AREA",
        "ZONE",
        "FLOOR",
    )
    description_keywords = (
        "DESCRIPTION",
        "ACTIVITY",
        "WORK DESCRIPTION",
        "WORK ITEM",
    )

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables() or []:
                if not table:
                    continue

                cleaned_table = [
                    [clean_cell(cell) for cell in row]
                    for row in table
                    if row
                ]

                header_index = None
                manpower_index = None
                location_index = None
                description_index = None
                item_index = None

                for row_index, row in enumerate(cleaned_table):
                    upper_row = [cell.upper() for cell in row]

                    manpower_candidates = [
                        index
                        for index, cell in enumerate(upper_row)
                        if any(keyword in cell for keyword in manpower_keywords)
                    ]
                    location_candidates = [
                        index
                        for index, cell in enumerate(upper_row)
                        if any(keyword in cell for keyword in location_keywords)
                    ]

                    # A Location-like column is required so the general manpower
                    # summary table (Position / Persons) is not mistaken for
                    # location distribution data.
                    if not manpower_candidates or not location_candidates:
                        continue

                    header_index = row_index
                    manpower_index = manpower_candidates[0]
                    location_index = location_candidates[0]

                    description_candidates = [
                        index
                        for index, cell in enumerate(upper_row)
                        if any(keyword in cell for keyword in description_keywords)
                    ]
                    description_index = (
                        description_candidates[0]
                        if description_candidates
                        else None
                    )

                    item_candidates = [
                        index
                        for index, cell in enumerate(upper_row)
                        if cell in {"ITEM", "NO", "NO.", "NUMBER"}
                    ]
                    item_index = item_candidates[0] if item_candidates else None
                    break

                if header_index is None:
                    continue

                for row in cleaned_table[header_index + 1 :]:
                    required_indices = [
                        index
                        for index in (
                            manpower_index,
                            location_index,
                            description_index,
                            item_index,
                        )
                        if index is not None
                    ]
                    required_length = max(required_indices) + 1
                    if len(row) < required_length:
                        row += [""] * (required_length - len(row))

                    manpower = parse_integer_cell(row[manpower_index or 0])
                    if manpower is None or manpower <= 0:
                        continue

                    location = row[location_index or 0]
                    description = (
                        row[description_index]
                        if description_index is not None
                        else ""
                    )
                    item = row[item_index] if item_index is not None else ""

                    joined = " ".join(row).upper()
                    if any(word in joined for word in ("TOTAL", "SUBTOTAL", "GRAND TOTAL")):
                        continue

                    if not location and not description:
                        continue

                    extracted_rows.append(
                        {
                            "item": item,
                            "description": description,
                            "location": location,
                            "manpower": manpower,
                        }
                    )

    return extracted_rows


def build_standard_detail_row(
    *,
    header: dict[str, Any],
    file_data: dict[str, Any],
    config: dict[str, Any],
    trade: str,
    item: str,
    description: str,
    worker_name: str,
    location_raw: str,
    manpower: int | None,
    manpower_method: str,
    worker_location_mode: bool = False,
) -> dict[str, Any]:
    """Convert any extraction method into the common detail schema."""
    if worker_location_mode:
        standard_location, towers, floors, status = build_worker_location_result(
            location_raw,
            config,
        )
    else:
        standard_location, towers, floors, status = build_location_result(
            description,
            location_raw,
            config,
        )

    if manpower is None:
        status = "需人工確認人數"

    return {
        "日期": header["date"] or "",
        "工種": trade,
        "文件": file_data["name"],
        "Item": item,
        "工作描述": description,
        "工人姓名": worker_name,
        "原始位置": location_raw,
        "樓座": towers,
        "樓層": floors,
        "標準位置": standard_location,
        "人數": manpower,
        "人數計算方法": manpower_method,
        "位置狀態": status,
    }


def make_analysis_summary(
    *,
    header: dict[str, Any],
    file_data: dict[str, Any],
    trade: str,
    detail_rows: list[dict[str, Any]],
    extraction_method: str,
    duplicate_names: list[str] | None = None,
    reconciliation_override: str | None = None,
) -> dict[str, Any]:
    extracted_total = sum(
        int(row["人數"])
        for row in detail_rows
        if row.get("人數") is not None
    )
    worker = header["worker"]

    if reconciliation_override is not None:
        reconciliation = reconciliation_override
    elif worker is None:
        reconciliation = "未能讀取 Worker"
    elif extracted_total == worker:
        reconciliation = "一致"
    else:
        reconciliation = f"不一致，相差 {worker - extracted_total:+d}"

    return {
        "文件": file_data["name"],
        "日期": header["date"] or "未讀取",
        "工種": trade,
        "Ref. No.": header["reference"] or "未讀取",
        "Today Total Manpower": header["today_total"],
        "Worker": worker,
        "管理及技術人員": header["management_staff"],
        "工作明細人數合計": extracted_total,
        "Worker核對": reconciliation,
        "明細提取方式": extraction_method,
        "明細列數": len(detail_rows),
        "重複姓名": (
            ", ".join(duplicate_names)
            if duplicate_names
            else "沒有"
        ),
    }


def analyse_numeric_table_pdf(
    file_data: dict[str, Any],
    config: dict[str, Any],
    trade: str,
) -> dict[str, Any]:
    raw_text = extract_pdf_text(file_data["bytes"])
    header = extract_header_data(raw_text)
    numeric_rows = extract_numeric_manpower_table_rows(file_data["bytes"])

    detail_rows = [
        build_standard_detail_row(
            header=header,
            file_data=file_data,
            config=config,
            trade=trade,
            item=clean_cell(row.get("item", "")),
            description=clean_cell(row.get("description", "")),
            worker_name="",
            location_raw=clean_cell(row.get("location", "")),
            manpower=row.get("manpower"),
            manpower_method="Location＋Manpower數字欄",
        )
        for row in numeric_rows
    ]

    return {
        "summary": make_analysis_summary(
            header=header,
            file_data=file_data,
            trade=trade,
            detail_rows=detail_rows,
            extraction_method="Location＋Manpower數字欄",
        ),
        "details": detail_rows,
        "raw_text": raw_text,
    }


def analyse_description_pdf(
    file_data: dict[str, Any],
    config: dict[str, Any],
    trade: str,
) -> dict[str, Any]:
    raw_text = extract_pdf_text(file_data["bytes"])
    header = extract_header_data(raw_text)
    site_rows = extract_site_work_table_rows(file_data["bytes"])
    extraction_source = "PDF表格"

    if not site_rows:
        site_rows = extract_fallback_detail_rows(raw_text)
        extraction_source = "文字後備規則"

    detail_rows: list[dict[str, Any]] = []
    for row in site_rows:
        description = clean_cell(row.get("description", ""))
        location_raw = clean_cell(row.get("location", ""))
        manpower, method = parse_description_manpower(description)

        # Only keep rows where a count was actually found. This prevents
        # ordinary work-description tables from being treated as manpower
        # distribution tables.
        if manpower is None:
            continue

        detail_rows.append(
            build_standard_detail_row(
                header=header,
                file_data=file_data,
                config=config,
                trade=trade,
                item=clean_cell(row.get("item", "")),
                description=description,
                worker_name="",
                location_raw=location_raw,
                manpower=manpower,
                manpower_method=method,
            )
        )

    return {
        "summary": make_analysis_summary(
            header=header,
            file_data=file_data,
            trade=trade,
            detail_rows=detail_rows,
            extraction_method=f"{extraction_source}／描述人數",
        ),
        "details": detail_rows,
        "raw_text": raw_text,
    }


def analyse_total_only_pdf(
    file_data: dict[str, Any],
    config: dict[str, Any],
    trade: str,
    note: str = "",
) -> dict[str, Any]:
    raw_text = extract_pdf_text(file_data["bytes"])
    header = extract_header_data(raw_text)

    manpower = header["worker"]
    method = "報告Worker總數"
    if manpower is None:
        manpower = header["today_total"]
        method = "Today Total Manpower（可能包括管理人員）"

    detail_rows: list[dict[str, Any]] = []
    if manpower is not None:
        detail_rows.append(
            {
                "日期": header["date"] or "",
                "工種": trade,
                "文件": file_data["name"],
                "Item": "",
                "工作描述": "報告只提供總人數" + (f"；{note}" if note else ""),
                "工人姓名": "",
                "原始位置": "Unspecified",
                "樓座": "",
                "樓層": "",
                "標準位置": "Unspecified",
                "人數": manpower,
                "人數計算方法": method,
                "位置狀態": "需人工分配位置",
            }
        )

    return {
        "summary": make_analysis_summary(
            header=header,
            file_data=file_data,
            trade=trade,
            detail_rows=detail_rows,
            extraction_method="只使用報告總人數",
            reconciliation_override="位置未分配",
        ),
        "details": detail_rows,
        "raw_text": raw_text,
    }


def analyse_report_pdf(
    file_data: dict[str, Any],
    config: dict[str, Any],
    trade: str,
    parser_mode: str,
) -> dict[str, Any]:
    """Select the requested parser or auto-detect a suitable format."""
    if parser_mode == "numeric_table":
        result = analyse_numeric_table_pdf(file_data, config, trade)
        if result["details"]:
            return result
        return analyse_total_only_pdf(
            file_data,
            config,
            trade,
            "找不到Location＋Manpower數字表",
        )

    if parser_mode == "worker_table":
        result = analyse_worker_table_pdf(file_data, config, trade)
        if result["details"]:
            return result
        return analyse_total_only_pdf(
            file_data,
            config,
            trade,
            "找不到工人姓名表",
        )

    if parser_mode == "description":
        result = analyse_description_pdf(file_data, config, trade)
        if result["details"]:
            return result
        return analyse_total_only_pdf(
            file_data,
            config,
            trade,
            "工作描述中找不到可計算人數",
        )

    if parser_mode == "total_only":
        return analyse_total_only_pdf(file_data, config, trade)

    # Automatic detection. Prefer explicit location/manpower tables. For FS
    # and PD, a worker attendance table is normally more reliable than any
    # incidental number appearing in the work-description section.
    numeric_result = analyse_numeric_table_pdf(file_data, config, trade)
    if numeric_result["details"]:
        numeric_result["summary"]["明細提取方式"] = "自動偵測：Location＋Manpower數字欄"
        return numeric_result

    if trade in {"FS", "PD"}:
        worker_result = analyse_worker_table_pdf(file_data, config, trade)
        if worker_result["details"]:
            worker_result["summary"]["明細提取方式"] = "自動偵測：工人姓名表"
            return worker_result

        description_result = analyse_description_pdf(file_data, config, trade)
        if description_result["details"]:
            description_result["summary"]["明細提取方式"] = "自動偵測：工作描述人數"
            return description_result
    else:
        description_result = analyse_description_pdf(file_data, config, trade)
        if description_result["details"]:
            description_result["summary"]["明細提取方式"] = "自動偵測：工作描述人數"
            return description_result

        worker_result = analyse_worker_table_pdf(file_data, config, trade)
        if worker_result["details"]:
            worker_result["summary"]["明細提取方式"] = "自動偵測：工人姓名表"
            return worker_result

    return analyse_total_only_pdf(
        file_data,
        config,
        trade,
        "自動偵測不到位置人數明細",
    )


# =========================================================
# Page title
# =========================================================

st.title("🏗️ 工地人力日報智能體（Python版）")
st.write(
    "建立工程及樓層配置，上傳 AC、EL、FS 和 PD 日報。"
    "目前以純 Python 分析 AC、EL、FS 和 PD PDF，不需要任何 AI API。"
    "系統可自動偵測姓名表、Manpower數字欄、描述中的人數，或只保留總人數。"
)
st.divider()


# =========================================================
# Section 1: Project configuration
# =========================================================

st.subheader("1. 工程項目設定")

with st.form("project_configuration_form"):
    project_name = st.text_input(
        "工程名稱",
        placeholder="例如：土瓜灣鴻福街及銀漢街發展項目",
    )
    tower_input = st.text_input(
        "樓座名稱",
        value="T1, T2, T3",
        help="用逗號分隔，例如：T1, T2, T3",
    )

    col1, col2 = st.columns(2)

    with col1:
        basement_input = st.text_input(
            "Basement 樓層",
            value="B4, B3, B2, B1",
        )
        podium_input = st.text_input(
            "Podium 樓層",
            value="GF, 1F, 2F, 3F",
        )
        special_floor_input = st.text_input(
            "特殊樓層或區域",
            value="Roof",
        )

    with col2:
        tower_start_floor = st.number_input(
            "Tower 最低樓層編號",
            min_value=1,
            max_value=200,
            value=5,
            step=1,
        )
        tower_end_floor = st.number_input(
            "Tower 最高樓層編號",
            min_value=1,
            max_value=200,
            value=29,
            step=1,
        )
        excluded_tower_floor_input = st.text_input(
            "不存在的 Tower 樓層",
            value="13F, 14F, 24F",
        )

    merge_podium_towers = st.checkbox(
        "Podium 樓層忽略 Tower 編號並合併",
        value=True,
    )

    submitted = st.form_submit_button(
        "儲存工程設定",
        type="primary",
        width="stretch",
    )

if submitted:
    if not project_name.strip():
        st.error("請輸入工程名稱。")
    elif tower_end_floor < tower_start_floor:
        st.error("Tower 最高樓層不能低於最低樓層。")
    else:
        towers = [tower.upper() for tower in text_to_list(tower_input)]
        basement_floors = [
            normalize_floor(item) for item in text_to_list(basement_input)
        ]
        podium_floors = [
            normalize_floor(item) for item in text_to_list(podium_input)
        ]
        special_locations = [
            normalize_floor(item) for item in text_to_list(special_floor_input)
        ]
        excluded_tower_floors = {
            normalize_floor(item)
            for item in text_to_list(excluded_tower_floor_input)
        }
        tower_floors = [
            f"{floor_number}F"
            for floor_number in range(
                int(tower_start_floor),
                int(tower_end_floor) + 1,
            )
            if f"{floor_number}F" not in excluded_tower_floors
        ]

        st.session_state["project_config"] = {
            "project_name": project_name.strip(),
            "towers": towers,
            "basement_floors": basement_floors,
            "podium_floors": podium_floors,
            "tower_start_floor": int(tower_start_floor),
            "tower_end_floor": int(tower_end_floor),
            "excluded_tower_floors": sorted(
                excluded_tower_floors,
                key=floor_sort_key,
            ),
            "tower_floors": tower_floors,
            "special_locations": special_locations,
            "merge_podium_tower_references": merge_podium_towers,
        }
        st.session_state.pop("analysis_results", None)
        st.success("工程設定已儲存。")


# =========================================================
# Section 2: Saved configuration
# =========================================================

if "project_config" in st.session_state:
    st.divider()
    st.subheader("2. 已儲存的工程設定")
    config = st.session_state["project_config"]

    col1, col2, col3 = st.columns(3)

    with col1:
        st.write(f"**工程名稱：** {config['project_name']}")
        st.write("**樓座：** " + (", ".join(config["towers"]) or "未設定"))

    with col2:
        st.write(
            "**Basement：** "
            + (", ".join(config["basement_floors"]) or "未設定")
        )
        st.write(
            "**Podium：** "
            + (", ".join(config["podium_floors"]) or "未設定")
        )
        st.write(
            "**特殊位置：** "
            + (", ".join(config["special_locations"]) or "未設定")
        )

    with col3:
        st.write(
            f"**Tower 範圍：** {config['tower_start_floor']}F–"
            f"{config['tower_end_floor']}F"
        )
        st.write(
            "**不存在樓層：** "
            + (", ".join(config["excluded_tower_floors"]) or "沒有")
        )
        st.write(f"**有效 Tower 樓層：** {len(config['tower_floors'])} 層")

    config_json = json.dumps(config, ensure_ascii=False, indent=2)
    st.download_button(
        "下載工程設定 JSON",
        data=config_json,
        file_name="project_config.json",
        mime="application/json",
        width="stretch",
    )


# =========================================================
# Section 3: Report upload
# =========================================================

st.divider()
st.subheader("3. 上傳工地日報")

project_is_ready = "project_config" in st.session_state
if not project_is_ready:
    st.warning("請先完成並儲存工程設定，再上傳日報。")

uploaded_files = st.file_uploader(
    "選擇 AC、EL、FS、PD 日報",
    type=["pdf", "png", "jpg", "jpeg", "xlsx", "xls"],
    accept_multiple_files=True,
    disabled=not project_is_ready,
)


# =========================================================
# Section 4: File checking
# =========================================================

if uploaded_files:
    st.divider()
    st.subheader("4. 文件檢查結果")

    seen_hashes: dict[str, str] = {}
    file_records: list[dict[str, Any]] = []
    unique_files: list[dict[str, Any]] = []
    duplicate_files: list[dict[str, Any]] = []

    for uploaded_file in uploaded_files:
        file_bytes = uploaded_file.getvalue()
        file_hash = calculate_file_hash(file_bytes)
        detected_trade = detect_trade_from_filename(uploaded_file.name)

        if file_hash in seen_hashes:
            original_filename = seen_hashes[file_hash]
            file_records.append(
                {
                    "文件名稱": uploaded_file.name,
                    "工種": detected_trade or "未識別",
                    "大小": format_file_size(len(file_bytes)),
                    "狀態": "重複，已忽略",
                    "備註": f"與 {original_filename} 完全相同",
                }
            )
            duplicate_files.append(
                {
                    "name": uploaded_file.name,
                    "original_name": original_filename,
                }
            )
        else:
            seen_hashes[file_hash] = uploaded_file.name
            file_records.append(
                {
                    "文件名稱": uploaded_file.name,
                    "工種": detected_trade or "未識別",
                    "大小": format_file_size(len(file_bytes)),
                    "狀態": "保留",
                    "備註": "",
                }
            )
            unique_files.append(
                {
                    "name": uploaded_file.name,
                    "bytes": file_bytes,
                    "hash": file_hash,
                    "trade": detected_trade,
                    "size": len(file_bytes),
                    "type": uploaded_file.type,
                }
            )

    st.session_state["unique_uploaded_files"] = unique_files

    metric1, metric2, metric3 = st.columns(3)
    metric1.metric("上傳文件", len(uploaded_files))
    metric2.metric("有效文件", len(unique_files))
    metric3.metric("重複文件", len(duplicate_files))

    st.dataframe(file_records, width="stretch", hide_index=True)

    if duplicate_files:
        for duplicate in duplicate_files:
            st.warning(
                f"{duplicate['name']} 與 {duplicate['original_name']} 完全相同，"
                "第二份已忽略。"
            )
    else:
        st.success("沒有發現完全重複的文件。")

    detected_trades = {
        file_data["trade"]
        for file_data in unique_files
        if file_data["trade"] is not None
    }
    missing_trades = set(REQUIRED_TRADES) - detected_trades

    trade_columns = st.columns(4)
    for column, trade in zip(trade_columns, REQUIRED_TRADES):
        with column:
            if trade in detected_trades:
                st.success(f"{trade}：已上傳")
            else:
                st.error(f"{trade}：缺少")

    if missing_trades:
        st.warning("目前缺少以下工種報告：" + ", ".join(sorted(missing_trades)))
    else:
        st.success("AC、EL、FS 和 PD 四個工種的報告均已上傳。")

    st.markdown("#### 準備分析的有效文件")
    for index, file_data in enumerate(unique_files, start=1):
        st.write(
            f"{index}. {file_data['name']} — "
            f"{file_data['trade'] or '未識別'} — "
            f"{format_file_size(file_data['size'])}"
        )
else:
    st.info("尚未上傳日報。請選擇 PDF、圖片或 Excel 文件。")


# =========================================================
# Section 5: Flexible Python report analysis
# =========================================================

st.divider()
st.subheader("5. AC／EL／FS／PD 彈性 Python 分析")
st.caption(f"目前版本：{APP_VERSION}")

unique_uploaded_files = st.session_state.get("unique_uploaded_files", [])
supported_pdf_files = [
    file_data
    for file_data in unique_uploaded_files
    if file_data.get("trade") in REQUIRED_TRADES
    and file_data["name"].lower().endswith(".pdf")
]

if not project_is_ready:
    st.warning("請先儲存工程設定。")
elif not supported_pdf_files:
    st.info("請先上傳至少一份 AC、EL、FS 或 PD PDF 日報。")
else:
    st.write(
        "準備分析："
        + ", ".join(file_data["name"] for file_data in supported_pdf_files)
    )

    st.markdown("### 各工種人數提取方式")
    st.caption(
        "建議保持『自動偵測』。其他工程沒有工人姓名表時，"
        "系統會改找Location＋Manpower數字欄、描述中的X人，"
        "再不成功才只保留總人數並要求人工分配位置。"
    )

    uploaded_trades = [
        trade
        for trade in REQUIRED_TRADES
        if any(file_data.get("trade") == trade for file_data in supported_pdf_files)
    ]

    parser_modes: dict[str, str] = {}
    parser_columns = st.columns(max(1, len(uploaded_trades)))
    option_labels = list(PARSER_MODE_OPTIONS.keys())

    for column, trade in zip(parser_columns, uploaded_trades):
        with column:
            selected_label = st.selectbox(
                f"{trade} 提取方式",
                options=option_labels,
                index=0,
                key=f"parser_mode_{trade}",
            )
            parser_modes[trade] = PARSER_MODE_OPTIONS[selected_label]

    if st.button(
        "使用 Python 分析全部 PDF",
        type="primary",
        width="stretch",
    ):
        results: list[dict[str, Any]] = []

        with st.spinner("正在辨認報告格式、提取位置及核對人數..."):
            for file_data in supported_pdf_files:
                trade = file_data.get("trade") or ""

                try:
                    result = analyse_report_pdf(
                        file_data,
                        st.session_state["project_config"],
                        trade,
                        parser_modes.get(trade, "auto"),
                    )
                    results.append(result)

                except Exception as error:
                    results.append(
                        {
                            "summary": {
                                "文件": file_data["name"],
                                "日期": "分析失敗",
                                "工種": trade or "未識別",
                                "Ref. No.": "",
                                "Today Total Manpower": None,
                                "Worker": None,
                                "管理及技術人員": None,
                                "工作明細人數合計": None,
                                "Worker核對": f"錯誤：{error}",
                                "明細提取方式": "失敗",
                                "明細列數": 0,
                                "重複姓名": "未檢查",
                            },
                            "details": [],
                            "raw_text": "",
                        }
                    )

        st.session_state["analysis_results"] = results
        st.success("AC／EL／FS／PD PDF 分析完成。")


analysis_results = st.session_state.get("analysis_results", [])

if analysis_results:
    st.markdown("### 分析摘要")
    summary_rows = [result["summary"] for result in analysis_results]
    summary_df = pd.DataFrame(summary_rows)
    st.dataframe(summary_df, width="stretch", hide_index=True)

    for summary in summary_rows:
        reconciliation = summary.get("Worker核對", "")
        if reconciliation == "一致":
            st.success(f"{summary['文件']}：明細合計與 Worker 一致。")
        elif reconciliation == "位置未分配":
            st.warning(
                f"{summary['文件']}：只取得總人數，位置需要人工分配。"
            )
        else:
            st.warning(f"{summary['文件']}：{reconciliation}")

        duplicate_names = summary.get("重複姓名", "沒有")
        if duplicate_names not in {"沒有", "不適用", "未檢查"}:
            st.warning(
                f"{summary['文件']}：Worker 表出現重複姓名：{duplicate_names}"
            )

    all_detail_rows = [
        row
        for result in analysis_results
        for row in result["details"]
    ]

    st.markdown("### 工作位置及人數明細（可人工修改）")

    if all_detail_rows:
        detail_df = pd.DataFrame(all_detail_rows)

        preferred_columns = [
            "日期",
            "工種",
            "文件",
            "Item",
            "工作描述",
            "工人姓名",
            "原始位置",
            "樓座",
            "樓層",
            "標準位置",
            "人數",
            "人數計算方法",
            "位置狀態",
        ]
        detail_df = detail_df.reindex(columns=preferred_columns)

        edited_detail_df = st.data_editor(
            detail_df,
            width="stretch",
            hide_index=True,
            num_rows="dynamic",
            disabled=[
                "日期",
                "工種",
                "文件",
                "Item",
                "工作描述",
                "工人姓名",
                "原始位置",
                "樓座",
                "樓層",
                "人數計算方法",
            ],
            column_config={
                "人數": st.column_config.NumberColumn(
                    "人數",
                    min_value=0,
                    step=1,
                )
            },
            key="combined_detail_editor_v040",
        )

        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                "下載分析摘要 CSV",
                data=summary_df.to_csv(index=False).encode("utf-8-sig"),
                file_name="report_analysis_summary.csv",
                mime="text/csv",
                width="stretch",
            )
        with col2:
            st.download_button(
                "下載位置明細 CSV",
                data=edited_detail_df.to_csv(index=False).encode("utf-8-sig"),
                file_name="location_details.csv",
                mime="text/csv",
                width="stretch",
            )

        st.session_state["edited_detail_records"] = (
            edited_detail_df.to_dict("records")
        )

        # =================================================
        # Section 6: Merge same-date and same-location data
        # =================================================

        st.divider()
        st.subheader("6. 跨工種位置合併")

        confirmed_summary_df, review_summary_df = (
            build_location_summary(edited_detail_df)
        )

        if not confirmed_summary_df.empty:
            total_confirmed = int(confirmed_summary_df["合計"].sum())

            metric1, metric2 = st.columns(2)
            metric1.metric("已合併標準位置", len(confirmed_summary_df))
            metric2.metric("已解析位置人數合計", total_confirmed)

            st.markdown("### 已解析位置摘要")
            st.caption(
                "合併條件：同一日期＋同一標準位置。"
                "排序：T1、T2、T3依次排列，每座由高層至低層；"
                "Podium由高層至GF；Basement由B1至B2、B3、B4。"
            )
            st.dataframe(
                confirmed_summary_df,
                width="stretch",
                hide_index=True,
            )

            st.download_button(
                "下載跨工種合併摘要 CSV",
                data=confirmed_summary_df.to_csv(index=False).encode("utf-8-sig"),
                file_name="merged_trade_location_summary.csv",
                mime="text/csv",
                width="stretch",
            )
        else:
            st.warning("目前沒有可直接合併的已解析位置。")

        st.markdown("### Cross-floor／Distribution U／未指定位置")
        st.caption(
            "只有總數、只寫樓座、跨多個區域或無法確定樓層的資料會保留在這裡。"
            "系統不會擅自平均分配。"
        )

        if not review_summary_df.empty:
            st.dataframe(
                review_summary_df,
                width="stretch",
                hide_index=True,
            )

            st.download_button(
                "下載待確認位置摘要 CSV",
                data=review_summary_df.to_csv(index=False).encode("utf-8-sig"),
                file_name="cross_floor_distribution_u.csv",
                mime="text/csv",
                width="stretch",
            )
        else:
            st.success("沒有 Cross-floor、Distribution U 或未指定位置。")

    else:
        st.error("未能從 PDF 提取任何可計算明細。")

    with st.expander("查看 PDF 原始提取文字（除錯用）", expanded=False):
        for result in analysis_results:
            st.markdown(f"#### {result['summary']['文件']}")
            st.text(result["raw_text"] or "沒有提取到文字。")