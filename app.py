import streamlit as st
import pandas as pd
import re
from collections import defaultdict
import io
import zipfile

# ==========================
# 页面配置
# ==========================
st.set_page_config(page_title="广告素材生成工具", page_icon="🚀", layout="wide")
st.title("🚀 广告素材批量生成与拆分工具")

# ==========================
# 左侧边栏：参数设置
# ==========================
st.sidebar.header("⚙️ 参数设置")
REPEAT_FIRST = st.sidebar.number_input("第一次重复次数", min_value=1, value=1)
REPEAT_SECOND = st.sidebar.number_input("第二次重复次数", min_value=1, value=1)
ENABLE_COLOR = st.sidebar.checkbox("开启 SKU 颜色标记", value=True)
FAST_MODE = st.sidebar.checkbox("开启极速模式 (直接输出纯数据)", value=False)

# ==========================
# 工具函数 (复用你优化后的核心逻辑)
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

    has_select = False
    start_select, end_select = 0, 0
    if material_select:
        sel_match = re.search(r'(\d+)-(\d+)', material_select)
        if sel_match:
            start_select, end_select = int(sel_match.group(1)), int(sel_match.group(2))
            has_select = True

    match_two_digits = re.search(r'-(\d{2})$', base_name)
    match_hyphenated = re.search(r'-(\d+)-(\d+)$', base_name)

    if match_two_digits:
        suffix = match_two_digits.group(1)
        group_num, start_material = suffix[0], int(suffix[1])
        prefix = base_name[:match_two_digits.start()]
        new_prefix = f"{prefix}-{group_num}-{lp_version}-" if lp_version else f"{prefix}-{group_num}-"
        s_m = start_select if has_select else start_material
        e_m = end_select if has_select else (start_material + material_count - 1)
        return [f"{new_prefix}{i}" for i in range(s_m, e_m + 1)]

    elif match_hyphenated:
        group_part, start_material = match_hyphenated.group(1), int(match_hyphenated.group(2))
        prefix = base_name[:match_hyphenated.start()]
        new_prefix = f"{prefix}-{group_part}-"
        s_m = start_select if has_select else start_material
        e_m = end_select if has_select else (start_material + material_count - 1)
        return [f"{new_prefix}{i}" for i in range(s_m, e_m + 1)]

    else:
        match_single = re.search(r'-(\d+)$', base_name)
        if match_single:
            start_material = int(match_single.group(1))
            prefix = base_name[:match_single.start()] + "-"
        else:
            start_material = 1
            prefix = base_name + "-"
        s_m = start_select if has_select else start_material
        e_m = end_select if has_select else (start_material + material_count - 1)
        return [f"{prefix}{i}" for i in range(s_m, e_m + 1)]

# ==========================
# 主界面逻辑
# ==========================
uploaded_file = st.file_uploader("📂 请上传基础素材表格 (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    if st.button("⚡ 开始处理与生成", type="primary"):
        with st.spinner("正在拼命处理中，请稍候..."):
            try:
                # 1. 读取数据
                df = pd.read_excel(uploaded_file, dtype=str).fillna("")
                
                # 2. 处理数据
                file_groups = defaultdict(list)
                all_dfs = []

                for _, row in df.iterrows():
                    versions = expand_material_versions(row)
                    if not versions: continue
                    material_len = len(versions)
                    new_rows = []
                    for v in versions:
                        new_row = row.copy()
                        new_row["广告素材版本名称"] = v
                        for _ in range(REPEAT_FIRST): new_rows.append(new_row.copy())
                    
                    temp_df = pd.DataFrame(new_rows)
                    temp_df = temp_df.sort_values(by="广告素材版本名称", key=lambda x: x.map(natural_sort_key))
                    if REPEAT_SECOND > 1: temp_df = pd.concat([temp_df] * REPEAT_SECOND, ignore_index=True)
                    
                    file_groups[material_len].append(temp_df)
                    all_dfs.append(temp_df)

                # 3. 准备输出任务
                output_tasks = {}
                if all_dfs: output_tasks["总表"] = pd.concat(all_dfs, ignore_index=True)
                for material_len, dfs in file_groups.items():
                    output_tasks[f"素材数_{material_len}"] = pd.concat(dfs, ignore_index=True)

                # 4. 生成 Excel 并打包成 ZIP (全在内存中进行，无需写入本地硬盘)
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    
                    for file_name, final_df in output_tasks.items():
                        excel_buffer = io.BytesIO()
                        
                        if FAST_MODE:
                            final_df.to_excel(excel_buffer, index=False)
                        else:
                            with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
                                sheet_name = "总表" if file_name == "总表" else "Data"
                                final_df.to_excel(writer, index=False, sheet_name=sheet_name)
                                workbook, worksheet = writer.book, writer.sheets[sheet_name]

                                # 自动列宽
                                for i, col in enumerate(final_df.columns):
                                    max_len = max(final_df[col].astype(str).map(len).max(), len(col)) + 2
                                    worksheet.set_column(i, i, max_len)

                                # 颜色标记
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
                        
                        excel_buffer.seek(0)
                        zip_file.writestr(f"{file_name}.xlsx", excel_buffer.read())
                
                zip_buffer.seek(0)
                
                st.success("🎉 全部生成完成！已将所有文件打包，请点击下方按钮下载。")
                
                # 提供下载按钮
                st.download_button(
                    label="📦 一键下载全部结果 (ZIP压缩包)",
                    data=zip_buffer,
                    file_name="广告素材处理结果.zip",
                    mime="application/zip"
                )

            except Exception as e:
                st.error(f"处理过程中出现错误: {e}")