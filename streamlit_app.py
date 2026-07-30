import hashlib
import json
import re
from typing import Any

import streamlit as st


st.set_page_config(
    page_title="Manpower AI Agent",
    page_icon="🏗️",
    layout="wide",
)


# =========================================================
# Helper functions
# =========================================================

def text_to_list(value: str) -> list[str]:
    """Convert comma-separated text into a clean list."""
    return [
        item.strip()
        for item in value.replace("，", ",").split(",")
        if item.strip()
    ]


def normalize_floor(value: str) -> str:
    """Normalize floor names such as 5/F, 5f and 5 F into 5F."""
    cleaned = value.strip().upper()
    cleaned = cleaned.replace("/", "")
    cleaned = cleaned.replace(" ", "")

    match = re.fullmatch(r"(\d+)F?", cleaned)

    if match:
        return f"{int(match.group(1))}F"

    return cleaned


def floor_sort_key(value: str) -> tuple[int, str]:
    """Sort numbered floors numerically."""
    match = re.fullmatch(r"(\d+)F", value)

    if match:
        return int(match.group(1)), value

    return 9999, value


def calculate_file_hash(file_bytes: bytes) -> str:
    """Return SHA-256 hash for duplicate-file detection."""
    return hashlib.sha256(file_bytes).hexdigest()


def detect_trade_from_filename(filename: str) -> str | None:
    """
    Detect AC, EL, FS or PD from the filename.

    Examples:
    DA-AC-369.pdf -> AC
    DA-EL-351.pdf -> EL
    FS-361.pdf -> FS
    PD-368.pdf -> PD
    MVAC report.pdf -> AC
    """
    filename_upper = filename.upper()

    trade_patterns = {
        "AC": [
            r"(^|[^A-Z0-9])AC([^A-Z0-9]|$)",
            r"MVAC",
        ],
        "EL": [
            r"(^|[^A-Z0-9])EL([^A-Z0-9]|$)",
            r"ELECTRICAL",
        ],
        "FS": [
            r"(^|[^A-Z0-9])FS([^A-Z0-9]|$)",
            r"FIRE[\s_-]*SERVICE",
        ],
        "PD": [
            r"(^|[^A-Z0-9])PD([^A-Z0-9]|$)",
            r"PLUMBING",
            r"DRAINAGE",
        ],
    }

    for trade, patterns in trade_patterns.items():
        for pattern in patterns:
            if re.search(pattern, filename_upper):
                return trade

    return None


def format_file_size(size_bytes: int) -> str:
    """Convert file size into a readable value."""
    if size_bytes < 1024:
        return f"{size_bytes} B"

    if size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"

    return f"{size_bytes / (1024 * 1024):.1f} MB"


# =========================================================
# Page title
# =========================================================

st.title("🏗️ 工地人力日報 AI 智能體")

st.write(
    """
    建立工程項目及樓層配置，上傳 AC、EL、FS 和 PD 日報。
    系統將檢查重複文件、識別缺少的工種，並為下一階段的
    AI 日報分析準備資料。
    """
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
            help="用逗號分隔，例如：B4, B3, B2, B1",
        )

        podium_input = st.text_input(
            "Podium 樓層",
            value="GF, 1F, 2F, 3F",
            help="每個工程可以自行設定，不會固定為GF至3F。",
        )

        special_floor_input = st.text_input(
            "特殊樓層或區域",
            value="Roof",
            help="例如：Roof, M/F, UG/F, Clubhouse",
        )

    with col2:
        tower_start_floor = st.number_input(
            "Tower 最低樓層編號",
            min_value=1,
            max_value=200,
            value=4,
            step=1,
        )

        tower_end_floor = st.number_input(
            "Tower 最高樓層編號",
            min_value=1,
            max_value=200,
            value=30,
            step=1,
        )

        excluded_tower_floor_input = st.text_input(
            "不存在的 Tower 樓層",
            value="4F, 13F, 14F, 24F",
            help="用逗號分隔，例如：4F, 13F, 14F, 24F",
        )

    merge_podium_towers = st.checkbox(
        "Podium 樓層忽略 Tower 編號並合併",
        value=True,
        help="例如 T1 GF 和 T3 GF 合併為 Podium / GF。",
    )

    submitted = st.form_submit_button(
        "儲存工程設定",
        type="primary",
        use_container_width=True,
    )


if submitted:
    if not project_name.strip():
        st.error("請輸入工程名稱。")

    elif tower_end_floor < tower_start_floor:
        st.error("Tower最高樓層不能低於最低樓層。")

    else:
        towers = [
            tower.upper()
            for tower in text_to_list(tower_input)
        ]

        basement_floors = [
            normalize_floor(item)
            for item in text_to_list(basement_input)
        ]

        podium_floors = [
            normalize_floor(item)
            for item in text_to_list(podium_input)
        ]

        special_locations = [
            normalize_floor(item)
            for item in text_to_list(special_floor_input)
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

        excluded_tower_floors_sorted = sorted(
            excluded_tower_floors,
            key=floor_sort_key,
        )

        project_config = {
            "project_name": project_name.strip(),
            "towers": towers,
            "basement_floors": basement_floors,
            "podium_floors": podium_floors,
            "tower_start_floor": int(tower_start_floor),
            "tower_end_floor": int(tower_end_floor),
            "excluded_tower_floors": excluded_tower_floors_sorted,
            "tower_floors": tower_floors,
            "special_locations": special_locations,
            "merge_podium_tower_references": merge_podium_towers,
        }

        st.session_state["project_config"] = project_config
        st.success("工程設定已儲存。")


# =========================================================
# Section 2: Saved project configuration
# =========================================================

if "project_config" in st.session_state:
    st.divider()
    st.subheader("2. 已儲存的工程設定")

    config = st.session_state["project_config"]

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("#### 基本資料")

        st.write(
            f"**工程名稱：** {config['project_name']}"
        )

        st.write(
            "**樓座：** "
            + (
                ", ".join(config["towers"])
                if config["towers"]
                else "未設定"
            )
        )

    with col2:
        st.markdown("#### 公共樓層")

        st.write(
            "**Basement：** "
            + (
                ", ".join(config["basement_floors"])
                if config["basement_floors"]
                else "未設定"
            )
        )

        st.write(
            "**Podium：** "
            + (
                ", ".join(config["podium_floors"])
                if config["podium_floors"]
                else "未設定"
            )
        )

        st.write(
            "**特殊位置：** "
            + (
                ", ".join(config["special_locations"])
                if config["special_locations"]
                else "未設定"
            )
        )

    with col3:
        st.markdown("#### Tower 樓層")

        st.write(
            f"**樓層範圍：** "
            f"{config['tower_start_floor']}F–"
            f"{config['tower_end_floor']}F"
        )

        st.write(
            "**不存在樓層：** "
            + (
                ", ".join(config["excluded_tower_floors"])
                if config["excluded_tower_floors"]
                else "沒有"
            )
        )

        st.write(
            f"**有效 Tower 樓層數量：** "
            f"{len(config['tower_floors'])}"
        )

    merge_text = (
        "是"
        if config["merge_podium_tower_references"]
        else "否"
    )

    st.info(
        f"Podium 是否忽略並合併 Tower 編號：{merge_text}"
    )

    with st.expander(
        "查看全部有效 Tower 樓層",
        expanded=False,
    ):
        st.write(", ".join(config["tower_floors"]))

    config_json = json.dumps(
        config,
        ensure_ascii=False,
        indent=2,
    )

    st.download_button(
        label="下載工程設定 JSON",
        data=config_json,
        file_name="project_config.json",
        mime="application/json",
        use_container_width=True,
    )


# =========================================================
# Section 3: Report upload
# =========================================================

st.divider()
st.subheader("3. 上傳工地日報")

if "project_config" not in st.session_state:
    st.warning(
        "請先完成並儲存工程設定，再上傳日報。"
    )

uploaded_files = st.file_uploader(
    "選擇 AC、EL、FS、PD 日報",
    type=[
        "pdf",
        "png",
        "jpg",
        "jpeg",
        "xlsx",
        "xls",
    ],
    accept_multiple_files=True,
    help=(
        "可以一次選擇多個文件。"
        "系統會使用SHA-256檢查完全重複的文件。"
    ),
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
        detected_trade = detect_trade_from_filename(
            uploaded_file.name
        )

        if file_hash in seen_hashes:
            original_filename = seen_hashes[file_hash]

            record = {
                "文件名稱": uploaded_file.name,
                "工種": detected_trade or "未識別",
                "大小": format_file_size(len(file_bytes)),
                "狀態": "重複，已忽略",
                "備註": f"與 {original_filename} 完全相同",
            }

            duplicate_files.append(
                {
                    "name": uploaded_file.name,
                    "original_name": original_filename,
                    "hash": file_hash,
                    "trade": detected_trade,
                }
            )

        else:
            seen_hashes[file_hash] = uploaded_file.name

            record = {
                "文件名稱": uploaded_file.name,
                "工種": detected_trade or "未識別",
                "大小": format_file_size(len(file_bytes)),
                "狀態": "保留",
                "備註": "",
            }

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

        file_records.append(record)

    st.session_state["unique_uploaded_files"] = unique_files

    unique_count = len(unique_files)
    duplicate_count = len(duplicate_files)

    detected_trades = {
        file_data["trade"]
        for file_data in unique_files
        if file_data["trade"] is not None
    }

    unknown_files = [
        file_data["name"]
        for file_data in unique_files
        if file_data["trade"] is None
    ]

    required_trades = {"AC", "EL", "FS", "PD"}
    missing_trades = required_trades - detected_trades

    metric1, metric2, metric3 = st.columns(3)

    with metric1:
        st.metric(
            "上傳文件",
            len(uploaded_files),
        )

    with metric2:
        st.metric(
            "有效文件",
            unique_count,
        )

    with metric3:
        st.metric(
            "重複文件",
            duplicate_count,
        )

    st.markdown("#### 文件清單")

    st.dataframe(
        file_records,
        use_container_width=True,
        hide_index=True,
    )

    if duplicate_files:
        for duplicate in duplicate_files:
            st.warning(
                f"重複文件：{duplicate['name']} "
                f"與 {duplicate['original_name']} 完全相同，"
                "第二份已忽略。"
            )
    else:
        st.success("沒有發現完全重複的文件。")

    st.markdown("#### 工種檢查")

    trade_columns = st.columns(4)

    for column, trade in zip(
        trade_columns,
        ["AC", "EL", "FS", "PD"],
    ):
        with column:
            if trade in detected_trades:
                st.success(f"{trade}：已上傳")
            else:
                st.error(f"{trade}：缺少")

    if missing_trades:
        missing_text = ", ".join(
            sorted(missing_trades)
        )

        st.warning(
            f"目前缺少以下工種報告：{missing_text}"
        )
    else:
        st.success(
            "AC、EL、FS和PD四個工種的報告均已上傳。"
        )

    if unknown_files:
        st.warning(
            "以下文件未能從檔名判斷工種："
            + ", ".join(unknown_files)
        )

        st.info(
            "下一階段會加入人工選擇工種，"
            "以及讓AI根據報告內容判斷工種。"
        )

    st.markdown("#### 準備分析的有效文件")

    for index, file_data in enumerate(
        unique_files,
        start=1,
    ):
        trade_text = file_data["trade"] or "未識別"

        st.write(
            f"{index}. {file_data['name']} "
            f"— {trade_text} "
            f"— {format_file_size(file_data['size'])}"
        )

    st.success(
        f"共有 {unique_count} 份有效文件已準備進入AI分析。"
    )

else:
    st.info(
        "尚未上傳日報。請選擇PDF、圖片或Excel文件。"
    )