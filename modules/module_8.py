import streamlit as st
import pandas as pd
import io
import zipfile
import gc
from collections import defaultdict
from utils import safe_int

# ─────────────────────────────────────────────
# 1. 动态表头与导出写入引擎
# ─────────────────────────────────────────────

def is_landing_page_url(val_series):
    """
    智能检测 G 列填入的是否为网址/链接：
    若包含 http, www, .com, /products/ 或 '/' 斜杠特征，则判定为链接；
    否则判定为着陆页版本名称。
    """
    sample_vals = [
        str(v).strip() for v in val_series.dropna() 
        if str(v).strip() and str(v).strip().lower() != "nan"
    ]
    if not sample_vals:
        return True  # 默认兜底为链接

    link_indicators = [".com", ".cn", ".top", ".shop", ".net", ".org", ".site", "http://", "https://", "www.", "/products/", "/funnel/", "/"]
    link_count = 0
    for v in sample_vals:
        v_lower = v.lower()
        if any(ind in v_lower for ind in link_indicators):
            link_count += 1

    return (link_count / len(sample_vals)) >= 0.5


def write_landing_page_excel(df_data, params):
    """
    导出 Excel，严格对齐模板结构并根据 G 列内容动态调整 G1/H1 表头：
    - 若 G 列为链接：G1=着陆页链接, H1=广告素材版本
    - 若 G 列为版本：G1=着陆页版本名称, H1=广告素材版本名称
    """
    df_out = df_data.copy()
    
    # 查找 G 列实际对应的字段名（兼容多种历史叫法）
    g_col_candidates = [c for c in df_out.columns if "着陆页" in c]
    g_actual_col = g_col_candidates[0] if g_col_candidates else "着陆页链接"
    
    # 查找 H 列实际对应的字段名
    h_col_candidates = [c for c in df_out.columns if "广告素材" in c or "素材版本" in c]
    h_actual_col = h_col_candidates[0] if h_col_candidates else "广告素材版本"

    # 判断是否为 URL 链接
    is_link = is_landing_page_url(df_out[g_actual_col])

    if is_link:
        g1_header = "着陆页链接"
        g2_hint = "填写完整的商品链接"
        h1_header = "广告素材版本"
        h2_hint = "广告素材库中的具体版本名称"
    else:
        g1_header = "着陆页版本名称"
        g2_hint = "着陆页库中的具体版本名称"
        h1_header = "广告素材版本名称"
        h2_hint = "广告素材库中的具体版本名称"

    target_columns = [
        "广告账号ID", "主页ID", "像素ID", "真实SKU", "虚拟SKU", 
        "国家", g1_header, h1_header, "出价/竞价", "系列标注", "Unnamed: 10"
    ]

    hints = {
        "广告账号ID": "",
        "主页ID": "可不填，不填则使用资产管理中的默认主页",
        "像素ID": "可不填，不填则使用资产管理中的默认像素",
        "真实SKU": "",
        "虚拟SKU": "",
        "国家": "美国/英国/德国/法国/西班牙",
        g1_header: g2_hint,
        h1_header: h2_hint,
        "出价/竞价": "如需指定“真实/虚拟SKU”与“出价/竞价”的关系，请填写，最多2位小数，可不填，不填则全不填，填了则全填",
        "系列标注": "可不填",
        "Unnamed: 10": "此行为说明，勿删除，请从第三行开始填写"
    }

    # 将原有数据列对齐到当前确定的动态表头上
    df_out[g1_header] = df_out[g_actual_col]
    df_out[h1_header] = df_out[h_actual_col]

    df_out = df_out.fillna("")
    for col in target_columns:
        if col not in df_out.columns:
            df_out[col] = ""
            
    df_out = df_out[target_columns]
    hint_row = pd.DataFrame([hints])
    
    final_export_df = pd.concat([hint_row, df_out], ignore_index=True)

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        final_export_df.to_excel(writer, index=False, sheet_name="Sheet1")
        workbook = writer.book
        worksheet = writer.sheets["Sheet1"]
        
        # 自动筛选与格式
        worksheet.autofilter(0, 0, 0, len(target_columns) - 1)
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#E0E0E0', 'border': 1})
        hint_fmt = workbook.add_format({'bg_color': '#FFFFCC', 'font_color': 'red', 'bold': True})
        
        worksheet.set_row(1, 22, hint_fmt)
        
        for i, col in enumerate(target_columns):
            col_name_str = "" if "Unnamed" in str(col) else str(col)
            worksheet.write(0, i, col_name_str, header_fmt)
            w = max(len(col_name_str.encode('gbk', errors='ignore')) + 4, 18)
            if col in [g1_header, h1_header]:
                w = 35
            elif col == "Unnamed: 10":
                w = 32
            worksheet.set_column(i, i, w)

    return out.getvalue()


# ─────────────────────────────────────────────
# 2. 核心匹配与拆表算法
# ─────────────────────────────────────────────

def get_sku_key(row):
    """有效 SKU 标识：优先虚拟SKU，否则真实SKU"""
    v = str(row.get("虚拟SKU", "")).strip()
    r = str(row.get("真实SKU", "")).strip()
    if v and v.lower() != "nan":
        return v
    if r and r.lower() != "nan":
        return r
    return ""


def run_module_8_matching(acc_file_bytes, n_group=50, mode="模式一"):
    # 1. 读取 SKU 表并按时间倒序去重
    df_sku_raw = pd.read_excel(io.BytesIO(acc_file_bytes), sheet_name="SKU表", dtype=str)
    df_sku_raw.columns = [str(c).strip() for c in df_sku_raw.columns]
    
    mat_col = [c for c in df_sku_raw.columns if "广告素材" in c or "素材版本" in c]
    mat_col_name = mat_col[0] if mat_col else "广告素材版本"
    
    df_sku_raw["SKU_KEY"] = df_sku_raw.apply(get_sku_key, axis=1)
    df_sku_raw["广告素材版本_CLEAN"] = df_sku_raw[mat_col_name].fillna("").astype(str).str.strip()
    
    if "创建时间" in df_sku_raw.columns:
        df_sku_raw["创建时间_dt"] = pd.to_datetime(df_sku_raw["创建时间"], errors="coerce")
        df_sku_raw = df_sku_raw.sort_values(by="创建时间_dt", ascending=False)
    
    # [版本 + SKU] 唯一去重，保留最新素材
    df_sku_dedup = df_sku_raw.drop_duplicates(subset=["SKU_KEY", "广告素材版本_CLEAN"], keep="first")
    
    sku_materials_map = defaultdict(list)
    for _, r in df_sku_dedup.iterrows():
        k = r["SKU_KEY"]
        mat = r["广告素材版本_CLEAN"]
        if k and mat:
            sku_materials_map[k].append(mat)

    # 2. 读取账号表并过滤说明行
    df_acc_raw = pd.read_excel(io.BytesIO(acc_file_bytes), sheet_name="账号表", dtype=str)
    df_acc_raw.columns = [str(c).strip() for c in df_acc_raw.columns]
    
    if len(df_acc_raw) > 0 and any("可不填" in str(v) for v in df_acc_raw.iloc[0].values):
        df_acc_data = df_acc_raw.iloc[1:].copy()
    else:
        df_acc_data = df_acc_raw.copy()
        
    df_acc_data = df_acc_data.dropna(how="all", axis=0)

    # 3. 按账号归集 SKU（1账号1SKU拆分机制）
    account_groups = defaultdict(list)
    for _, r in df_acc_data.iterrows():
        acc_id = str(r.get("广告账号ID", "")).strip()
        if not acc_id or acc_id.lower() == "nan":
            continue
        account_groups[acc_id].append(r.to_dict())

    if not account_groups:
        return {}

    max_skus_per_acc = max(len(items) for items in account_groups.values())
    table_buckets = [[] for _ in range(max_skus_per_acc)]
    for acc_id, items in account_groups.items():
        for k, item in enumerate(items):
            table_buckets[k].append(item)

    # 4. 匹配素材并生成子表
    result_tables = {}
    for t_idx, bucket in enumerate(table_buckets):
        matched_rows = []
        for acc_item in bucket:
            sku_k = get_sku_key(acc_item)
            mats = sku_materials_map.get(sku_k, [])
            
            if mode == "模式一":
                target_mats = mats[:n_group] if mats else [""]
            else:
                target_mats = mats if mats else [""]

            for mat in target_mats:
                new_row = acc_item.copy()
                new_row["广告素材版本"] = mat
                matched_rows.append(new_row)
        
        if matched_rows:
            result_tables[f"表{t_idx + 1}"] = pd.DataFrame(matched_rows)

    return result_tables


# ─────────────────────────────────────────────
# 3. 页面渲染与下载
# ─────────────────────────────────────────────

def run(params):
    st.subheader("🔗 模块八：着陆页导入与素材智能匹配")

    with st.expander("💡 点击查看：运行逻辑说明", expanded=False):
        st.markdown("""
        - **智能识别表头**：系统会自动识别 G 列内容。若为商品链接，表头保持 `着陆页链接` 与 `广告素材版本`；若为版本名称（如优化组版本），自动切换为 `着陆页版本名称` 与 `广告素材版本名称`。
        - **数据预处理**：SKU表按 `[版本 + SKU]` 联合唯一去重，且按 `创建时间` 倒序排序（最新的优先匹配）。
        - **分表规则（1账号1SKU）**：同一张表下只允许一个账号对应一个SKU；如果同一账号有多个不同SKU，系统会自动拆分为多个子表格导出。
        - **模式一 (单组截断)**：每个账号-SKU组合仅取前 $n$ 个最新素材版本，超出部分舍弃。
        - **模式二 (全量素材)**：提取该 SKU 对应的全部素材版本，有多少取多少。
        """)

    st.markdown("##### ⚙️ 1. 匹配规则与参数配置")
    c_left, c_right = st.columns([1, 1], gap="large")

    with c_left:
        run_mode = st.radio(
            "选择素材提取模式",
            ["模式一：单组截断 (只要一组数据，最多n个)", "模式二：全量提取 (取全部素材，有多少是多少)"],
            help="模式一截取前 n 个素材；模式二提取全量素材"
        )
        mode_key = "模式一" if "模式一" in run_mode else "模式二"

        n_group = st.number_input(
            "🔢 每组素材数量上限 (n)", 
            min_value=1, 
            max_value=500, 
            value=50, 
            step=1,
            help="模式一下用于控制每个账号-SKU匹配的最大素材数量（默认50）"
        )

    with c_right:
        up = st.file_uploader("上传包含【账号表】和【SKU表】的需求文件 (.xlsx)", type=["xlsx"], key="m8_up")

    # 清除旧缓存
    if "m8_last_up" not in st.session_state or st.session_state["m8_last_up"] != (up.name if up else None):
        st.session_state["m8_last_up"] = up.name if up else None
        st.session_state.pop("m8_download_bytes", None)
        st.session_state.pop("m8_table_counts", None)
        st.session_state.pop("m8_is_zip", None)

    if up:
        st.write("")
        st.markdown("##### 🚀 2. 执行匹配与结果导出")
        btn_col1, btn_col2 = st.columns([1, 1], gap="large")

        with btn_col1:
            start_btn = st.button("🚀 开始匹配并拆表", key="m8_btn", use_container_width=True)

        if start_btn:
            file_bytes = up.getvalue()
            with st.spinner("正在智能拆表并匹配最新素材..."):
                try:
                    result_tables = run_module_8_matching(file_bytes, n_group=int(n_group), mode=mode_key)
                except Exception as e:
                    st.error(f"⚠️ 处理失败，请检查文件格式是否正确：{e}")
                    return

                if not result_tables:
                    st.warning("⚠️ 未匹配到有效数据，请检查账号表与SKU表内容！")
                    return

                file_prefix = params.get("prefix", "项目_")
                table_counts_info = {tbl: len(df) for tbl, df in result_tables.items()}

                # 智能打包：单表直接出 .xlsx，多表打包为 .zip
                if len(result_tables) == 1:
                    single_tbl = list(result_tables.values())[0]
                    st.session_state["m8_download_bytes"] = write_landing_page_excel(single_tbl, params)
                    st.session_state["m8_is_zip"] = False
                else:
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                        for tbl_name, df_tbl in result_tables.items():
                            excel_bytes = write_landing_page_excel(df_tbl, params)
                            file_name = f"{file_prefix}着陆页导入_{tbl_name}.xlsx"
                            zf.writestr(file_name, excel_bytes)
                    st.session_state["m8_download_bytes"] = zip_buffer.getvalue()
                    st.session_state["m8_is_zip"] = True

                st.session_state["m8_table_counts"] = table_counts_info

            possible_vars = ["df_sku_raw", "df_acc_raw", "result_tables"]
            for var in possible_vars:
                if var in locals():
                    del locals()[var]
            gc.collect()

        if "m8_download_bytes" in st.session_state:
            file_prefix = params.get("prefix", "项目_")
            table_counts = st.session_state.get("m8_table_counts", {})
            total_rows = sum(table_counts.values())
            is_zip = st.session_state.get("m8_is_zip", False)

            with btn_col2:
                if is_zip:
                    st.download_button(
                        f"💾 下载拆表结果压缩包 (共 {len(table_counts)} 个表 / {total_rows} 行)",
                        data=st.session_state["m8_download_bytes"],
                        file_name=f"{file_prefix}着陆页导入_拆表匹配结果.zip",
                        mime="application/zip",
                        use_container_width=True,
                        key="dl_m8_btn",
                    )
                else:
                    st.download_button(
                        f"💾 下载匹配结果 Excel ({total_rows} 行)",
                        data=st.session_state["m8_download_bytes"],
                        file_name=f"{file_prefix}着陆页导入_表1.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="dl_m8_btn",
                    )

            st.write("")
            st.success(f"🎉 匹配完成！已成功按「1账号1SKU」拆分为 **{len(table_counts)}** 个表格。")
            with st.expander("📊 查看各拆分子表的数据行数明细", expanded=True):
                for tbl_name, cnt in table_counts.items():
                    st.caption(f"• **{tbl_name}**：共包含 **{cnt}** 条广告素材导入记录")