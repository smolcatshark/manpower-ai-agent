import json
import re

import streamlit as st


st.set_page_config(
    page_title="Manpower AI Agent",
    page_icon="🏗️",
    layout="wide",
)


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
    """Sort numbered floors numerically and other labels alphabetically."""
    match = re.fullmatch(r"(\d+)F", value)

    if match:
        return int(match.group(1)), value

    return 9999, value


st.title("🏗️ 工地人力日報 AI 智能體")

st.write(
    """
    建立工程項目及樓層配置。系統日後會按照每個工程的設定，
    分析 AC、EL、FS 和 PD 日報。
    """
)

st.divider()

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
            help="每個工程可以自行設定，不會固定為 GF 至 3F。",
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
        st.error("Tower 最高樓層不能低於最低樓層。")

    else:
        towers = text_to_list(tower_input)

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


if "project_config" in st.session_state:
    st.divider()
    st.subheader("2. 已儲存的工程設定")

    config = st.session_state["project_config"]

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("#### 基本資料")
        st.write(f"**工程名稱：** {config['project_name']}")
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

    st.info(f"Podium 是否忽略並合併 Tower 編號：{merge_text}")

    with st.expander("查看全部有效 Tower 樓層", expanded=True):
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