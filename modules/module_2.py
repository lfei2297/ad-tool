import streamlit as st
import pandas as pd
import gc
from collections import defaultdict
from utils import expand_material_versions, read_uploaded_excel, create_zip_package


def _sku_key(row_dict):
    sku = str(row_dict.get("真实SKU", "")).strip()
    if sku and sku.lower() != "nan":
        return sku
    sku = str(row_dict.get("虚拟SKU", "")).strip()
    if sku and sku.lower() != "nan":
        return sku
    return ""


def build_module2_tables(valid_rows, repeat_1=1, repeat_2=1):
    """
    先按 SKU+国家 汇总素材条数，再按「国家 + 素材数」合并进同一张子表。
    例如两个美国 SKU 都是 5 条素材 → 同一份 美国_素材数_5。
    """
    all_expanded = []
    per_sku = defaultdict(list)
    per_sku_count = defaultdict(int)

    for row_dict in valid_rows:
        vs = expand_material_versions(row_dict)
        m_len = len(vs)
        row_rows = []
        for v in vs:
            nr = row_dict.copy()
            nr["广告素材版本名称"] = v
            for _ in range(repeat_1):
                row_rows.append(nr.copy())
        tmp = pd.DataFrame(row_rows)
        if repeat_2 > 1 and not tmp.empty:
            tmp = pd.concat([tmp] * repeat_2, ignore_index=True)
        if tmp.empty:
            continue
        all_expanded.append(tmp)
        sku = _sku_key(row_dict)
        country = str(row_dict.get("国家", "")).strip()
        if country.lower() == "nan":
            country = ""
        per_sku[(sku, country)].append(tmp)
        per_sku_count[(sku, country)] += m_len

    if not all_expanded:
        return {}

    tasks = {"总表": pd.concat(all_expanded, ignore_index=True)}
    by_file = defaultdict(list)
    for key, dfs in per_sku.items():
        _, country = key
        total_m = per_sku_count[key]
        c_tag = f"{country}_" if country else ""
        filename = f"{c_tag}素材数_{total_m}"
        by_file[filename].extend(dfs)

    for filename, dfs in by_file.items():
        tasks[filename] = pd.concat(dfs, ignore_index=True)
    return tasks


def run(params):
    st.subheader("🌍 模块二：同SKU+国家聚合拆分")
    st.caption("同一国家、素材条数相同的 SKU 会合并进同一张子表（例如两个美国 SKU 都是 5 条 → 美国_素材数_5）。")

    up = st.file_uploader("📂 上传原始素材表 (.xlsx)", type=["xlsx"], key="m2_up")

    if "m2_last_up" not in st.session_state or st.session_state["m2_last_up"] != (up.name if up else None):
        st.session_state["m2_last_up"] = up.name if up else None
        st.session_state.pop("m2_zip_bytes", None)
        st.session_state.pop("m2_table_names", None)

    if up:
        st.write("")
        if st.button("🚀 开始处理", key="m2_btn"):
            df_raw = read_uploaded_excel(up.getvalue())
            raw_dicts = df_raw.to_dict("records")
            valid_rows = [
                row for row in raw_dicts
                if not any("此行为说明" in str(v) or "可不填" in str(v) for v in row.values())
            ]
            if not valid_rows:
                st.warning("⚠️ 上传的表格中未检测到有效的业务数据，请检查后再试。")
                return

            tasks = build_module2_tables(valid_rows, repeat_1=1, repeat_2=1)
            if not tasks:
                st.warning("⚠️ 未生成有效数据。")
                return

            zip_bytes = create_zip_package(tasks, params)
            st.session_state["m2_zip_bytes"] = zip_bytes
            st.session_state["m2_table_names"] = {n: len(df) for n, df in tasks.items()}
            del df_raw, valid_rows, tasks
            gc.collect()

        if "m2_zip_bytes" in st.session_state:
            names = st.session_state.get("m2_table_names") or {}
            st.success(f"🎉 聚合处理完成！共 **{len(names)}** 个文件（含总表）。")
            st.download_button(
                "📦 下载聚合结果包 (ZIP)",
                data=st.session_state["m2_zip_bytes"],
                file_name=f"{params.get('prefix', '项目_')}聚合结果.zip",
                mime="application/zip",
                key="dl_m2_zip",
            )
            with st.expander("查看各文件行数", expanded=False):
                for n, cnt in names.items():
                    st.caption(f"• **{n}**：{cnt} 行")
