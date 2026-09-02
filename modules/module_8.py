import streamlit as st
import pandas as pd
import io
import zipfile
import gc
from collections import defaultdict
from utils import safe_int

# ─────────────────────────────────────────────
# 1. 动态表头与自适应导出引擎
# ─────────────────────────────────────────────

def is_landing_page_url(val_series):
    sample_vals = [
        str(v).strip() for v in val_series.dropna() 
        if str(v).strip() and str(v).strip().lower() != "nan"
    ]
    if not sample_vals:
        return True

    link_indicators = [".com", ".cn", ".top", ".shop", ".net", ".org", ".site", "http://", "https://", "www.", "/products/", "/funnel/", "/"]
    link_count = 0
    for v in sample_vals:
        v_lower = v.lower()
        if any(ind in v_lower for ind in link_indicators):
            link_count += 1

    return (link_count / len(sample_vals)) >= 0.5


def write_landing_page_excel(df_data, raw_hint_dict=None):
    df_out = df_data.copy().fillna("")

    # 1. 识别 G 列与 H 列
    g_col_candidates = [c for c in df_out.columns if "着陆页" in c]
    g_actual_col = g_col_candidates[0] if g_col_candidates else "着陆页链接"
    
    is_link = is_landing_page_url(df_out[g_actual_col])

    # G 列动态切换，H 列统一锁定为【广告素材版本名称】
    if is_link:
        g1_header = "着陆页链接"
        g2_hint = "填写完整的商品链接"
    else:
        g1_header = "着陆页版本名称"
        g2_hint = "着陆页库中的具体版本名称"

    h1_header = "广告素材版本名称"
    h2_hint = "广告素材库中的具体版本名称"

    # 更名 G 列（如果有差异）
    if g_actual_col != g1_header:
        df_out = df_out.rename(columns={g_actual_col: g1_header})

    # 🌟 核心防错：彻底删除历史旧的【广告素材版本】列，防止与【广告素材版本名称】重名冲突
    if "广告素材版本" in df_out.columns and "广告素材版本名称" in df_out.columns:
        df_out = df_out.drop(columns=["广告素材版本"])
    elif "广告素材版本" in df_out.columns:
        df_out = df_out.rename(columns={"广告素材版本": h1_header})

    # 🌟 列名强制去重（仅保留第一次出现的同名列）
    df_out = df_out.loc[:, ~df_out.columns.duplicated()].copy()

    # 2. 动态调整列顺序：确保【广告素材ID】紧跟在【广告素材版本名称】正后方
    current_cols = list(df_out.columns)
    if "广告素材ID" in current_cols:
        current_cols.remove("广告素材ID")
        target_pos_col = h1_header if h1_header in current_cols else current_cols[min(7, len(current_cols)-1)]
        idx = current_cols.index(target_pos_col) + 1
        current_cols.insert(idx, "广告素材ID")

    # 确保 Unnamed 说明列排在最后
    unnamed_cols = [c for c in current_cols if "Unnamed" in str(c)]
    for uc in unnamed_cols:
        current_cols.remove(uc)
        current_cols.append(uc)

    df_out = df_out[current_cols]

    # 3. 构造提示行 (Hints)
    default_hints = {
        "广告账号ID": "",
        "主页ID": "可不填，不填则使用资产管理中的默认主页",
        "像素ID": "可不填，不填则使用资产管理中的默认像素",
        "真实SKU": "",
        "虚拟SKU": "",
        "国家": "美国/英国/德国/法国/西班牙",
        "着陆页链接": "填写完整的商品链接",
        "着陆页版本名称": "着陆页库中的具体版本名称",
        "广告素材版本名称": "广告素材库中的具体版本名称",
        "广告素材ID": "可不填，对应广告素材库中的唯一ID",
        "出价/竞价": "如需指定“真实/虚拟SKU”与“出价/竞价”的关系，请填写，最多2位小数，可不填，不填则全不填，填了则全填",
        "系列标注": "可不填",
        "Unnamed: 10": "此行为说明，勿删除，请从第三行开始填写"
    }

    final_hints = {}
    for col in current_cols:
        if raw_hint_dict and col in raw_hint_dict and str(raw_hint_dict[col]).strip() not in ["", "nan", "None"]:
            final_hints[col] = raw_hint_dict[col]
        else:
            final_hints[col] = default_hints.get(col, "可不填" if "Unnamed" not in str(col) else "此行为说明，勿删除，请从第三行开始填写")

    hint_row = pd.DataFrame([final_hints])
    
    # 保证 hint_row 与 df_out 列完全一致且无重复
    hint_row = hint_row.reindex(columns=df_out.columns)
    final_export_df = pd.concat([hint_row, df_out], ignore_index=True)

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        final_export_df.to_excel(writer, index=False, sheet_name="Sheet1")
        workbook = writer.book
        worksheet = writer.sheets["Sheet1"]
        
        worksheet.autofilter(0, 0, 0, len(current_cols) - 1)
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#E0E0E0', 'border': 1})
        hint_fmt = workbook.add_format({'bg_color': '#FFFFCC', 'font_color': 'red', 'bold': True})
        
        worksheet.set_row(1, 22, hint_fmt)
        
        for i, col in enumerate(current_cols):
            col_name_str = "" if "Unnamed" in str(col) else str(col)
            worksheet.write(0, i, col_name_str, header_fmt)
            w = max(len(col_name_str.encode('gbk', errors='ignore')) + 4, 18)
            if "着陆页" in col or "广告素材" in col:
                w = 32
            elif "Unnamed" in str(col):
                w = 32
            worksheet.set_column(i, i, w)

    return out.getvalue()


# ─────────────────────────────────────────────
# 2. 核心匹配与拆表算法
# ─────────────────────────────────────────────

def get_sku_key(row):
    v = str(row.get("虚拟SKU", "")).strip()
    r = str(row.get("真实SKU", "")).strip()
    if v and v.lower() != "nan":
        return v
    if r and r.lower() != "nan":
        return r
    return ""


def run_module_8_matching(acc_file_bytes, n_group=50, mode="模式一"):
    df_sku_raw = pd.read_excel(io.BytesIO(acc_file_bytes), sheet_name="SKU表", dtype=str)
    df_sku_raw.columns = [str(c).strip() for c in df_sku_raw.columns]
    
    mat_ver_col = [c for c in df_sku_raw.columns if "广告素材版本" in c or "素材版本" in c]
    mat_ver_name = mat_ver_col[0] if mat_ver_col else "广告素材版本"

    mat_id_col = [c for c in df_sku_raw.columns if "广告素材ID" in c or "素材ID" in c]
    mat_id_name = mat_id_col[0] if mat_id_col else "广告素材ID"
    
    df_sku_raw["SKU_KEY"] = df_sku_raw.apply(get_sku_key, axis=1)
    df_sku_raw["广告素材版本_CLEAN"] = df_sku_raw[mat_ver_name].fillna("").astype(str).str.strip()
    df_sku_raw["广告素材ID_CLEAN"] = df_sku_raw[mat_id_name].fillna("").astype(str).str.strip() if mat_id_name in df_sku_raw.columns else ""
    
    if "创建时间" in df_sku_raw.columns:
        df_sku_raw["创建时间_dt"] = pd.to_datetime(df_sku_raw["创建时间"], errors="coerce")
        df_sku_raw = df_sku_raw.sort_values(by="创建时间_dt", ascending=False)

    sku_materials_map = defaultdict(list)
    for _, r in df_sku_raw.iterrows():
        k = r["SKU_KEY"]
        mat_v = r["广告素材版本_CLEAN"]
        mat_i = r["广告素材ID_CLEAN"]
        if k and mat_v:
            sku_materials_map[k].append((mat_v, mat_i))

    df_acc_raw = pd.read_excel(io.BytesIO(acc_file_bytes), sheet_name="账号表", dtype=str)
    df_acc_raw.columns = [str(c).strip() for c in df_acc_raw.columns]
    
    raw_hint_dict = {}
    if len(df_acc_raw) > 0 and any("可不填" in str(v) for v in df_acc_raw.iloc[0].values):
        raw_hint_dict = df_acc_raw.iloc[0].to_dict()
        df_acc_data = df_acc_raw.iloc[1:].copy()
    else:
        df_acc_data = df_acc_raw.copy()
        
    df_acc_data = df_acc_data.dropna(how="all", axis=0)

    account_groups = defaultdict(list)
    for _, r in df_acc_data.iterrows():
        acc_id = str(r.get("广告账号ID", "")).strip()
        if not acc_id or acc_id.lower() == "nan":
            continue
        account_groups[acc_id].append(r.to_dict())

    if not account_groups:
        return {}, raw_hint_dict

    max_skus_per_acc = max(len(items) for items in account_groups.values())
    table_buckets = [[] for _ in range(max_skus_per_acc)]
    for acc_id, items in account_groups.items():
        for k, item in enumerate(items):
            table_buckets[k].append(item)

    result_tables = {}
    for t_idx, bucket in enumerate(table_buckets):
        matched_rows = []
        for acc_item in bucket:
            sku_k = get_sku_key(acc_item)
            mats = sku_materials_map.get(sku_k, [])
            
            if mode == "模式一":
                target_mats = mats[:n_group] if mats else [("", "")]
            else:
                target_mats = mats if mats else [("", "")]

            for mat_v, mat_i in target_mats:
                new_row = acc_item.copy()
                # 🌟 清理可能存在的旧字段名，避免字典中重复
                new_row.pop("广告素材版本", None)
                new_row["广告素材版本名称"] = mat_v
                new_row["广告素材ID"] = mat_i
                matched_rows.append(new_row)
        
        if matched_rows:
            result_tables[f"表{t_idx + 1}"] = pd.DataFrame(matched_rows)

    return result_tables, raw_hint_dict


# ─────────────────────────────────────────────
# 3. Streamlit 页面与操作交互
# ─────────────────────────────────────────────

def run(params):
    st.subheader("🔗 模块八：着陆页导入与素材智能匹配")

    with st.expander("💡 点击查看：运行逻辑说明", expanded=False):
        st.markdown("""
        - **表头规范**：H 列统一生成为 `广告素材版本名称`；G 列若为商品链接则为 `着陆页链接`，若为版本名称则切换为 `着陆页版本名称`。
        - **免改代码列自适应**：支持用户在账号表中随意增删自定义列，系统自动 100% 原样继承导出。
        - **时间排序与素材ID注入**：SKU表取消去重，按 `创建时间` 倒序排列（最新在前），并在素材版本后自动关联输出 `广告素材ID`。
        - **分表规则（1账号1SKU）**：同一张表下只允许一个账号对应一个SKU；多 SKU 账号自动拆分为多个子表格。
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
                    result_tables, raw_hint_dict = run_module_8_matching(file_bytes, n_group=int(n_group), mode=mode_key)
                except Exception as e:
                    st.error(f"⚠️ 处理失败，请检查文件格式是否正确：{e}")
                    return

                if not result_tables:
                    st.warning("⚠️ 未匹配到有效数据，请检查账号表与SKU表内容！")
                    return

                file_prefix = params.get("prefix", "项目_")
                table_counts_info = {tbl: len(df) for tbl, df in result_tables.items()}

                if len(result_tables) == 1:
                    single_tbl = list(result_tables.values())[0]
                    st.session_state["m8_download_bytes"] = write_landing_page_excel(single_tbl, raw_hint_dict)
                    st.session_state["m8_is_zip"] = False
                else:
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                        for tbl_name, df_tbl in result_tables.items():
                            excel_bytes = write_landing_page_excel(df_tbl, raw_hint_dict)
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