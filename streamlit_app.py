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
    valid_names = [
        name.strip()
        for name in possible_names
        if re.fullmatch(r"[\u3400-\u9fff]{2,4}", name.strip())
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
    text = clean_cell(value).upper()

    if not text:
        return ""

    text = text.replace("地下", "GF")
    text = text.replace("樓", "F")
    text = text.replace("－", "-").replace("–", "-")
    text = re.sub(r"B(\d+)\s*/\s*F", r"B\1", text)
    text = re.sub(r"(?<!B)(\d+)\s*/\s*F", r"\1F", text)
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
    normalized = normalize_report_location(text)

    if normalized in {"", "NA"}:
        return []

    tokens: list[str] = []

    if "GF" in normalized:
        tokens.append("GF")

    for number in re.findall(r"B(\d+)", normalized):
        floor = f"B{int(number)}"
        if floor not in tokens:
            tokens.append(floor)

    for number in re.findall(r"(?<![A-Z0-9])(\d+)F", normalized):
        floor = f"{int(number)}F"
        if floor not in tokens:
            tokens.append(floor)

    return tokens


def build_location_result(
    description: str,
    location_raw: str,
    config: dict[str, Any],
) -> tuple[str, str, str, str]:
    canonical_location = normalize_report_location(location_raw)
    combined_text = f"{description} {location_raw}"
    towers = extract_towers(description)
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
        return (
            f"Cross-floor / {canonical_location}",
            tower_text,
            floor_text,
            "保留為跨樓層，不作假設分配",
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
    }

    return {
        "summary": summary,
        "details": detail_rows,
        "raw_text": raw_text,
    }


# =========================================================
# Page title
# =========================================================

st.title("🏗️ 工地人力日報智能體（Python版）")
st.write(
    "建立工程及樓層配置，上傳 AC、EL、FS 和 PD 日報。"
    "目前先以純 Python 分析 EL PDF，不需要任何 AI API。"
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
        st.session_state.pop("el_analysis_results", None)
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
# Section 5: Python EL analysis
# =========================================================

st.divider()
st.subheader("5. EL PDF Python 分析")

unique_uploaded_files = st.session_state.get("unique_uploaded_files", [])
el_pdf_files = [
    file_data
    for file_data in unique_uploaded_files
    if file_data.get("trade") == "EL"
    and file_data["name"].lower().endswith(".pdf")
]

if not project_is_ready:
    st.warning("請先儲存工程設定。")
elif not el_pdf_files:
    st.info("請先上傳至少一份檔名包含 EL 的 PDF 日報。")
else:
    st.write("準備分析：" + ", ".join(file_data["name"] for file_data in el_pdf_files))

    if st.button(
        "使用 Python 分析 EL PDF",
        type="primary",
        width="stretch",
    ):
        results: list[dict[str, Any]] = []

        with st.spinner("正在讀取 PDF 表格及核對人數..."):
            for file_data in el_pdf_files:
                try:
                    results.append(
                        analyse_el_pdf(
                            file_data,
                            st.session_state["project_config"],
                        )
                    )
                except Exception as error:
                    results.append(
                        {
                            "summary": {
                                "文件": file_data["name"],
                                "日期": "分析失敗",
                                "工種": "EL",
                                "Ref. No.": "",
                                "Today Total Manpower": None,
                                "Worker": None,
                                "管理及技術人員": None,
                                "工作明細人數合計": None,
                                "Worker核對": f"錯誤：{error}",
                                "明細提取方式": "失敗",
                                "明細列數": 0,
                            },
                            "details": [],
                            "raw_text": "",
                        }
                    )

        st.session_state["el_analysis_results"] = results
        st.success("EL PDF 分析完成。")


analysis_results = st.session_state.get("el_analysis_results", [])

if analysis_results:
    st.markdown("### 分析摘要")
    summary_rows = [result["summary"] for result in analysis_results]
    summary_df = pd.DataFrame(summary_rows)
    st.dataframe(summary_df, width="stretch", hide_index=True)

    for summary in summary_rows:
        if summary["Worker核對"] == "一致":
            st.success(f"{summary['文件']}：工作明細合計與 Worker 一致。")
        else:
            st.warning(f"{summary['文件']}：{summary['Worker核對']}")

    all_detail_rows = [
        row
        for result in analysis_results
        for row in result["details"]
    ]

    st.markdown("### 工作位置及人數明細（可人工修改）")

    if all_detail_rows:
        detail_df = pd.DataFrame(all_detail_rows)
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
            key="el_detail_editor",
        )

        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                "下載 EL 分析摘要 CSV",
                data=summary_df.to_csv(index=False).encode("utf-8-sig"),
                file_name="el_analysis_summary.csv",
                mime="text/csv",
                width="stretch",
            )
        with col2:
            st.download_button(
                "下載 EL 位置明細 CSV",
                data=edited_detail_df.to_csv(index=False).encode("utf-8-sig"),
                file_name="el_location_details.csv",
                mime="text/csv",
                width="stretch",
            )
    else:
        st.error("未能從 PDF 提取 Site Work 明細表。")

    with st.expander("查看 PDF 原始提取文字（除錯用）", expanded=False):
        for result in analysis_results:
            st.markdown(f"#### {result['summary']['文件']}")
            st.text(result["raw_text"] or "沒有提取到文字。")