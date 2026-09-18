import streamlit as st
import pandas as pd
import gc
from utils import write_excel_final, read_uploaded_excel, safe_int, expand_material_versions

DEFAULT_VERSION = "默认版本"


def build_material_pool(row_dict, provided_count):
    """按提供素材版本数量展开；数量空/0 表示没有真实素材，由调用方补默认版本。"""
    if provided_count <= 0:
        return []
    expand_row = dict(row_dict)
    expand_row["广告素材数量"] = provided_count
    return expand_material_versions(expand_row)


def _pick_material(materials, idx):
    if 0 <= idx < len(materials):
        return materials[idx]
    return DEFAULT_VERSION


def iter_campaign_slots(m_groups, n_ads, series_count):
    """生成导出行顺序：(系列0起, 广告组0起, 广告0起)。

    例：2 个系列、1:2:3 → 共 12 行。前 6 行系列1，后 6 行系列2。
    每个系列内按广告组交叉：组1、组2、组1、组2… 直到每组都有 N 条广告。
    """
    for series in range(int(series_count)):
        for ad in range(int(n_ads)):
            for group in range(int(m_groups)):
                yield series, group, ad


def material_index_for_slot(logic_mode, series, group, ad, m_groups, n_ads):
    """系列间接着消耗素材，不从头数。"""
    if logic_mode == "组间测素材":
        return series * m_groups + group
    return series * (m_groups * n_ads) + ad * m_groups + group


def assigned_version(logic_mode, materials, series, group, ad, m_groups, n_ads):
    idx = material_index_for_slot(logic_mode, series, group, ad, m_groups, n_ads)
    return _pick_material(materials, idx)


def build_series_matrices(logic_mode, materials, m_groups, n_ads, series_count):
    """按系列依次消耗素材，系列之间不重复；不够的坑位用默认版本。

    组间测素材：每个系列消耗 M 条（一组一条，组内 N 个广告共用）。
    组内测素材：每个系列消耗 M×N 条（每个广告位一条）。
    返回 list[matrix]，matrix[组][广告] 为版本名。
    """
    matrices = [
        [[DEFAULT_VERSION for _ in range(n_ads)] for _ in range(m_groups)]
        for _ in range(int(series_count))
    ]
    for series, group, ad in iter_campaign_slots(m_groups, n_ads, series_count):
        matrices[series][group][ad] = assigned_version(
            logic_mode, materials, series, group, ad, m_groups, n_ads
        )
    return matrices


def flatten_series_rows(matrices, m_groups, n_ads):
    """与导出一致：先 Ad1 的所有组，再 Ad2 的所有组。"""
    names = []
    series_count = len(matrices)
    for series, group, ad in iter_campaign_slots(m_groups, n_ads, series_count):
        names.append(matrices[series][group][ad])
    return names


def run(params):
    st.subheader("🎯 模块四：补齐默认版本 (全结构适配版)")
    
    with st.expander("💡 点击查看：模块四运行逻辑说明"):
        st.markdown("""
        - **行排列顺序**：先写完一个系列，再写下一系列。系列内部按广告组交叉：组1、组2、组1、组2… 直到每组都凑满 N 条广告。
        - **示例（2 个系列、结构 1:2:3）**：共 12 行。前 6 行系列1（组1、组2、组1、组2、组1、组2），后 6 行系列2，同样交叉。
        - **系列间不同素材**：多个系列接着往下用素材，不会每个系列都从第 1 条重数。
        - **自动补齐**：当现有素材不足以填满全部系列的 $1:M:N$ 结构时，系统将自动填充“默认版本”。
        """)
    
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        M_groups = st.number_input("网页设定广告组数 (M)", min_value=1, value=1)
    with col_b:
        N_ads = st.number_input("网页设定广告数 (N)", min_value=1, value=3)
    with col_c:
        logic_mode = st.radio("🎨 素材分配逻辑", ["组间测素材", "组内测素材"])

    up = st.file_uploader("📂 上传模块四专用模板", type=["xlsx"], key="m4_up")
    sig = (up.name if up else None, int(M_groups), int(N_ads), logic_mode)
    if st.session_state.get("m4_sig") != sig:
        st.session_state["m4_sig"] = sig
        st.session_state.pop("m4_xlsx", None)
        st.session_state.pop("m4_warnings", None)

    if up and st.button("🚀 开始校验并生成", key="m4_btn"):
        df_raw = read_uploaded_excel(up.getvalue())
        
        valid_rows = [
            r for r in df_raw.to_dict('records') 
            if not any('此行为说明' in str(v) or '可不填' in str(v) for v in r.values())
        ]
        
        all_results = []
        warning_logs = []

        for idx, row_dict in enumerate(valid_rows):
            excel_line_no = idx + 3
            
            provided_count = safe_int(row_dict.get("提供素材版本数量"))
            series_count = safe_int(row_dict.get("导品系列数"), default=1)
            template_padding = safe_int(row_dict.get("补充默认版本数"))
            materials = build_material_pool(row_dict, provided_count)
            m_groups = int(M_groups)
            n_ads = int(N_ads)

            final_series_rows = []
            for series, group, ad in iter_campaign_slots(m_groups, n_ads, series_count):
                new_row = row_dict.copy()
                v_name = assigned_version(
                    logic_mode, materials, series, group, ad, m_groups, n_ads
                )
                new_row["广告素材版本名称"] = v_name
                new_row["备注"] = "系统自动补齐" if v_name == DEFAULT_VERSION else ""
                final_series_rows.append(new_row)

            actual_padding = sum(1 for r in final_series_rows if r["广告素材版本名称"] == "默认版本")
            if template_padding != 0 and template_padding != actual_padding:
                warning_logs.append(f"❓ Excel第 {excel_line_no} 行: 模板补充数({template_padding}) vs 实际({actual_padding})")

            all_results.append(pd.DataFrame(final_series_rows))

        if all_results:
            final_df = pd.concat(all_results, ignore_index=True)
            cols_to_hide = ["提供素材版本数量", "广告组数量", "导品系列数", "补充默认版本数"]
            final_df = final_df.drop(columns=[c for c in cols_to_hide if c in final_df.columns], errors='ignore')
            
            # 🌟 1. 提取当前所有列名
            all_cols = list(final_df.columns)

            # 🌟 2. 重新校准【系列标注】的位置（强行插到“出价/竞价”正右侧）
            if "系列标注" in all_cols:
                all_cols.remove("系列标注")  # 先剥离默认排在末尾的系列标注
                
                if "出价/竞价" in all_cols:
                    idx = all_cols.index("出价/竞价") + 1
                    all_cols.insert(idx, "系列标注")  # 精准插到“出价/竞价”后面
                else:
                    # 兜底：若无出价列，插在第 3 列（像素ID 后面）
                    insert_pos = min(3, len(all_cols))
                    all_cols.insert(insert_pos, "系列标注")

            # 🌟 3. 强制把【注意事项】和【备注】推到表格最右侧末尾
            for tail_col in ["注意事项", "备注"]:
                if tail_col in all_cols:
                    all_cols.remove(tail_col)
                    all_cols.append(tail_col)

            # 🌟 4. 应用重新排列后的列名顺序
            final_df = final_df.reindex(columns=all_cols)
            
            st.session_state["m4_xlsx"] = write_excel_final(final_df, "结构补齐结果", params)
            st.session_state["m4_warnings"] = warning_logs
            del df_raw, valid_rows, all_results, final_df
            gc.collect()

    if "m4_xlsx" in st.session_state:
        warns = st.session_state.get("m4_warnings") or []
        if warns:
            with st.expander("📝 逻辑校验报告"):
                for log in warns:
                    st.warning(log)
        st.success("✅ 生成完毕！")
        file_prefix = params.get("prefix", "项目_")
        st.download_button(
            f"💾 下载：{file_prefix}结果",
            data=st.session_state["m4_xlsx"],
            file_name=f"{file_prefix}补齐.xlsx",
            key="dl_m4",
        )