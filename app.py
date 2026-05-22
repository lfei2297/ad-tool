import streamlit as st
import pandas as pd
import re
from collections import defaultdict
import io
import zipfile

# ==========================
# 页面配置
# ==========================
st.set_page_config(page_title="广告素材生成工具 v9", page_icon="🚀", layout="wide")
st.title("🚀 广告素材批量生成工具 (V9 纯净版)")

# ==========================
# 左侧边栏
# ==========================
st.sidebar.header("🎯 核心功能选择")
PROCESS_MODE = st.sidebar.radio(
    "🔄 请选择工作模式：",
    [
        "模块一：基础独立拆分", 
        "模块二：同SKU+国家聚合拆分",
        "模块三：智能分组 (30个/组 & SKU去重)"
    ]
)

st.sidebar.markdown("---")
st.sidebar.header("⚙️ 细节参数设置")
FILE_PREFIX = st.sidebar.text_input("✏️ 自定义结果文件前缀", value="项目A_")
REPEAT_FIRST = st.sidebar.number_input("第一次重复次数", min_value=1, value=1)
REPEAT_SECOND = st.sidebar.number_input("第二次重复次数", min_value=1, value=1)
ENABLE_COLOR = st.sidebar.checkbox("开启 SKU 颜色标记", value=True)
FAST_MODE = st.sidebar.checkbox("开启极速模式", value=False)

# ==========================
# 核心逻辑函数
# ==========================
def natural_sort_key(s):
    return [int(t) if t.isdigit() else t.lower() for t in re.split('([0-9]+)', str(s))]

def expand_material_versions(row):
    base_name = str(row.get("广告素材版本名称", "")).strip()
    lp_name = str(row.get("着陆页版本名称", "")).strip()
    try:
        material_count = int(row["广告素材数量"]) if row["广告素材数量"] else 1
    except:
        material_count = 1
    material_select = str(row.get("素材选取", "") or row.get("素材选取 (X-Y)", "")).strip()
    lp_match = re.search(r'-(\d+)$', lp_name)
    lp_version = lp_match.group(1) if lp_match else ""

    match_two_digits = re.search(r'-(\d{2})$', base_name)
    match_hyphenated = re.search(r'-(\d+)-(\d+)$', base_name)
    sel_match = re.search(r'(\d+)-(\d+)', material_select)
    has_select = True if sel_match else False
    s_m, e_m = (int(sel_match.group(1)), int(sel_match.group(2))) if has_select else (0,0)

    if match_two_digits:
        suffix = match_two_digits.group(1)
        group_num, start_material = suffix[0], int(suffix[1])
        prefix = base_name[:match_two_digits.start()]
        new_prefix = f"{prefix}-{group_num}-{lp_version}-" if lp_version else f"{prefix}-{group_num}-"
        start = s_m if has_select else start_material
        end = e_m if has_select else (start_material + material_count - 1)
        return [f"{new_prefix}{i}" for i in range(start, end + 1)]
    elif match_hyphenated:
        group_part, start_material = match_hyphenated.group(1), int(match_hyphenated.group(2))
        prefix = base_name[:match_hyphenated.start()]
        new_prefix = f"{prefix}-{group_part}-"
        start = s_m if has_select else start_material
        end = e_m if has_select else (start_material + material_count - 1)
        return [f"{new_prefix}{i}" for i in range(start, end + 1)]
    else:
        match_single = re.search(r'-(\d+)$', base_name)
        prefix = base_name[:match_single.start()] + "-" if match_single else base_name + "-"
        start_material = int(match_single.group(1)) if match_single else 1
        start = s_m if has_select else start_material
        end = e_m if has_select else (start_material + material_count - 1)
        return [f"{prefix}{i}" for i in range(start, end + 1)]

def smart_grouping(df_to_group, group_size=30):
    # 过滤掉 SKU 为空或包含“总计/空白/Unnamed”的数据，防止表 2 中的干扰
    df_clean = df_to_group[df_to_group["真实SKU"].astype(str).str.strip() != ""].copy()
    df_clean = df_clean[~df_clean["真实SKU"].astype(str).str.contains("总计|空白|Unnamed", na=False)]
    
    records = df_clean.to_dict('records')
    # 预排序增加离散度
    records = sorted(records, key=lambda x: (str(x.get("真实SKU","")), str(x.get("着陆页版本名称",""))))
    
    buckets = []
    for rec in records:
        sku = str(rec.get("真实SKU", "")).strip() or str(rec.get("虚拟SKU", "")).strip()
        placed = False
        for bucket in buckets:
            if len(bucket) >= group_size: continue
            existing_skus = { (str(r.get("真实SKU","")).strip() or str(r.get("虚拟SKU","")).strip()) for r in bucket }
            if sku not in existing_skus:
                bucket.append(rec)
                placed = True
                break
        if not placed:
            buckets.append([rec])
    return buckets

def format_excel_worksheet(writer, final_df, sheet_name):
    workbook, worksheet = writer.book, writer.sheets[sheet_name]
    for i, col in enumerate(final_df.columns):
        max_len = max(final_df[col].astype(str).map(len).max(), len(col)) + 2
        worksheet.set_column(i, i, max_len)
    if ENABLE_COLOR:
        color_map = {}
        colors = ["#FFF2CC", "#E2EFDA", "#DDEBF7", "#F8CBAD", "#E4DFEC", "#D9D9D9"]
        c_idx = 0
        for row_idx in range(len(final_df)):
            sku = final_df.iloc[row_idx].get("真实SKU") or final_df.iloc[row_idx].get("虚拟SKU")
            if sku and sku not in color_map:
                color_map[sku] = colors[c_idx % len(colors)]
                c_idx += 1
            if sku in color_map:
                fmt = workbook.add_format({"bg_color": color_map[sku]})
                worksheet.set_row(row_idx + 1, None, fmt)

# ==========================
# 主界面
# ==========================
uploaded_file = st.file_uploader("📂 上传素材表格 (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    if st.button("🚀 执行处理", type="primary"):
        with st.spinner("正在清洗并处理数据..."):
            try:
                # 读取时只读取有标题的行，跳过可能的顶部干扰
                input_df = pd.read_excel(uploaded_file, dtype=str).fillna("")
                
                # 步骤一：通用的素材展开逻辑
                # 自动判定：如果存在“广告素材数量”列且有数字，则认为需要展开
                is_raw = "广告素材数量" in input_df.columns and input_df["广告素材数量"].astype(str).str.isdigit().any()
                
                if is_raw:
                    expanded_list = []
                    # 清洗原始数据中的无效行
                    input_df = input_df[input_df["真实SKU"].astype(str).str.strip() != ""].copy()
                    input_df = input_df[~input_df["真实SKU"].astype(str).str.contains("总计|空白", na=False)]
                    
                    for _, row in input_df.iterrows():
                        versions = expand_material_versions(row)
                        for v in versions:
                            new_row = row.copy()
                            new_row["广告素材版本名称"] = v
                            for _ in range(REPEAT_FIRST): expanded_list.append(new_row.copy())
                    total_df = pd.DataFrame(expanded_list)
                    if REPEAT_SECOND > 1: total_df = pd.concat([total_df] * REPEAT_SECOND, ignore_index=True)
                else:
                    total_df = input_df

                output_tasks = {}

                if "模块一" in PROCESS_MODE:
                    output_tasks["总表"] = total_df
                    if is_raw:
                        file_groups = defaultdict(list)
                        for _, row in input_df.iterrows():
                            v_list = expand_material_versions(row)
                            m_len = len(v_list)
                            row_expanded = []
                            for v in v_list:
                                nr = row.copy(); nr["广告素材版本名称"] = v
                                for _ in range(REPEAT_FIRST): row_expanded.append(nr.copy())
                            row_df = pd.DataFrame(row_expanded)
                            if REPEAT_SECOND > 1: row_df = pd.concat([row_df] * REPEAT_SECOND, ignore_index=True)
                            file_groups[f"素材数_{m_len}"].append(row_df)
                        for k, v in file_groups.items(): output_tasks[k] = pd.concat(v, ignore_index=True)

                elif "模块二" in PROCESS_MODE:
                    # 聚合逻辑...
                    group_totals = defaultdict(int)
                    group_dfs = defaultdict(list)
                    for _, row in input_df.iterrows():
                        v_list = expand_material_versions(row); m_len = len(v_list)
                        sku = str(row.get("真实SKU","")).strip() or str(row.get("虚拟SKU","")).strip()
                        country = str(row.get("国家","")).strip()
                        group_totals[(sku, country)] += m_len
                        row_expanded = []
                        for v in v_list:
                            nr = row.copy(); nr["广告素材版本名称"] = v
                            for _ in range(REPEAT_FIRST): row_expanded.append(nr.copy())
                        row_df = pd.DataFrame(row_expanded)
                        if REPEAT_SECOND > 1: row_df = pd.concat([row_df] * REPEAT_SECOND, ignore_index=True)
                        group_dfs[(sku, country)].append(row_df)
                    
                    for (sku, country), dfs in group_dfs.items():
                        total_m = group_totals[(sku, country)]
                        c_prefix = f"{country}_" if country else ""
                        output_tasks[f"{c_prefix}素材数_{total_m}"] = pd.concat(dfs, ignore_index=True)
                    output_tasks["总表"] = total_df

                elif "模块三" in PROCESS_MODE:
                    buckets = smart_grouping(total_df, 30)
                    m3_buffer = io.BytesIO()
                    with pd.ExcelWriter(m3_buffer, engine='xlsxwriter') as writer:
                        for idx, bucket in enumerate(buckets):
                            b_df = pd.DataFrame(bucket)
                            s_name = f"分组_{idx+1}({len(b_df)}条)"
                            b_df.to_excel(writer, index=False, sheet_name=s_name)
                            if not FAST_MODE: format_excel_worksheet(writer, b_df, s_name)
                        total_df.to_excel(writer, index=False, sheet_name="生成后的总表")
                        if not FAST_MODE: format_excel_worksheet(writer, total_df, "生成后的总表")
                    
                    st.success("🎉 模块三：智能分组总表生成成功！已剔除空白行。")
                    st.download_button(
                        label="📥 下载智能分组总表.xlsx",
                        data=m3_buffer.getvalue(),
                        file_name=f"{FILE_PREFIX}智能分组总表.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.stop()

                # 打包 ZIP
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for name, final_df in output_tasks.items():
                        excel_io = io.BytesIO()
                        with pd.ExcelWriter(excel_io, engine="xlsxwriter") as writer:
                            sn = "总表" if name == "总表" else "Data"
                            final_df.to_excel(writer, index=False, sheet_name=sn)
                            if not FAST_MODE: format_excel_worksheet(writer, final_df, sn)
                        excel_io.seek(0)
                        zip_file.writestr(f"{FILE_PREFIX}{name}.xlsx", excel_io.read())
                
                st.download_button(label="📦 下载结果 ZIP 包", data=zip_buffer.getvalue(), file_name="处理结果.zip")

            except Exception as e:
                st.error(f"处理失败: {e}")