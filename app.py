import streamlit as st
import pandas as pd
import re
from collections import defaultdict
import io
import zipfile

# ==========================
# 页面配置
# ==========================
st.set_page_config(page_title="广告素材生成工具 v5", page_icon="🚀", layout="wide")
st.title("🚀 广告素材批量生成与拆分工具 (双模块版)")

# ==========================
# 左侧边栏：参数与模块设置
# ==========================
st.sidebar.header("🎯 核心功能选择")

PROCESS_MODE = st.sidebar.radio(
    "🔄 请选择工作模式：",
    [
        "模块一：基础独立拆分 (原版)", 
        "模块二：同SKU+国家聚合拆分 (新版)"
    ],
    help="模块一：每行独立拆分。模块二：按国家和SKU合并计算总素材数拆分，且文件名附带国家。"
)

st.sidebar.markdown("---")
st.sidebar.header("⚙️ 细节参数设置")
FILE_PREFIX = st.sidebar.text_input("✏️ 自定义结果文件前缀", value="项目A_", help="生成的文件名将以该文本开头")
REPEAT_FIRST = st.sidebar.number_input("第一次重复次数", min_value=1, value=1)
REPEAT_SECOND = st.sidebar.number_input("第二次重复次数", min_value=1, value=1)
ENABLE_COLOR = st.sidebar.checkbox("开启 SKU 颜色标记", value=True)
FAST_MODE = st.sidebar.checkbox("开启极速模式 (直接输出纯数据)", value=False)

# ==========================
# 工具函数 (核心解析逻辑)
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

def get_template_buffer():
    template_data = {
        "广告账号ID": ["", "", ""],
        "主页ID": ["", "", ""],
        "像素ID": ["", "", ""],
        "真实SKU": ["SKU1", "SKU1", "SKU1"],
        "虚拟SKU": ["", "", ""],
        "国家": ["美国", "德国", "美国"],
        "着陆页版本名称": ["着陆页版本1", "着陆页版本2", "着陆页版本3"],
        "广告素材版本名称": ["优化组-11", "优化组-22", "优化组-33"],
        "广告素材数量": [2, 3, 4],
        "素材选取 (X-Y)": ["", "", ""]
    }
    t_df = pd.DataFrame(template_data)
    t_buffer = io.BytesIO()
    t_df.to_excel(t_buffer, index=False, engine='openpyxl')
    t_buffer.seek(0)
    return t_buffer

# ==========================
# 界面布局
# ==========================
st.markdown("### 📥 1. 规范数据格式")
st.download_button(
    label="点击下载：标准输入表格模板.xlsx",
    data=get_template_buffer(),
    file_name="标准输入表格模板.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.markdown("---")
st.markdown("### 📂 2. 上传数据并生成")
uploaded_file = st.file_uploader("请上传填入好数据的表格 (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    btn_text = "⚡ 开始执行：模块一 (独立拆分)" if "模块一" in PROCESS_MODE else "⚡ 开始执行：模块二 (聚合拆分)"
    
    if st.button(btn_text, type="primary"):
        with st.spinner("正在拼命处理中..."):
            try:
                df = pd.read_excel(uploaded_file, dtype=str).fillna("")
                file_groups = defaultdict(list)
                all_dfs = []
                
                # ==========================================
                # 模块一逻辑：独立拆分
                # ==========================================
                if "模块一" in PROCESS_MODE:
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
                        
                        # 模块一：直接按素材数命名
                        file_groups[f"素材数_{material_len}"].append(temp_df)
                        all_dfs.append(temp_df)

                # ==========================================
                # 模块二逻辑：同国家+同SKU聚合拆分 (带国家备注)
                # ==========================================
                else:
                    processed_records = []
                    group_material_totals = defaultdict(int)

                    # 第一阶段：统计聚合数据
                    for _, row in df.iterrows():
                        versions = expand_material_versions(row)
                        if not versions: continue
                        
                        material_len = len(versions)
                        sku = str(row.get("真实SKU", "")).strip() or str(row.get("虚拟SKU", "")).strip()
                        country = str(row.get("国家", "")).strip()
                        group_key = (sku, country)
                        
                        group_material_totals[group_key] += material_len
                        
                        new_rows = []
                        for v in versions:
                            new_row = row.copy()
                            new_row["广告素材版本名称"] = v
                            for _ in range(REPEAT_FIRST): new_rows.append(new_row.copy())
                        
                        temp_df = pd.DataFrame(new_rows)
                        temp_df = temp_df.sort_values(by="广告素材版本名称", key=lambda x: x.map(natural_sort_key))
                        if REPEAT_SECOND > 1: temp_df = pd.concat([temp_df] * REPEAT_SECOND, ignore_index=True)
                        
                        processed_records.append({"df": temp_df, "group_key": group_key})

                    # 第二阶段：根据总数和国家分发
                    for record in processed_records:
                        group_key = record["group_key"]
                        country = group_key[1]  # 提取国家名称
                        total_material_len = group_material_totals[group_key]
                        
                        # ✨ 核心修改点：将国家加入到字典键值（即未来的文件名）中
                        country_str = f"{country}_" if country else ""
                        file_key = f"{country_str}素材数_{total_material_len}"
                        
                        file_groups[file_key].append(record["df"])
                        all_dfs.append(record["df"])

                # ==========================================
                # 统一打包与输出装配
                # ==========================================
                output_tasks = {}
                
                if all_dfs:
                    total_df = pd.concat(all_dfs, ignore_index=True)
                    if "模块二" in PROCESS_MODE:
                        total_df = total_df.sort_values(by=["真实SKU", "虚拟SKU", "国家"], kind="stable")
                    output_tasks["总表"] = total_df
                    
                for task_name, dfs in file_groups.items():
                    sub_df = pd.concat(dfs, ignore_index=True)
                    if "模块二" in PROCESS_MODE:
                        sub_df = sub_df.sort_values(by=["真实SKU", "虚拟SKU", "国家"], kind="stable")
                    output_tasks[task_name] = sub_df

                # 生成 ZIP
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
                        
                        excel_buffer.seek(0)
                        
                        # 最终文件名拼装逻辑：用户自定义前缀 + (国家 +) 素材数
                        prefix_clean = str(FILE_PREFIX).strip()
                        final_filename = f"{prefix_clean}{file_name}.xlsx"
                        zip_file.writestr(final_filename, excel_buffer.read())
                
                zip_buffer.seek(0)
                st.success("🎉 处理成功！请点击下方按钮下载打包好的数据。")
                
                st.download_button(
                    label="📦 点击一键下载结果 (ZIP压缩包)",
                    data=zip_buffer,
                    file_name=f"{prefix_clean}广告素材结果.zip" if prefix_clean else "广告素材结果.zip",
                    mime="application/zip"
                )

            except Exception as e:
                st.error(f"处理过程中出现错误: {e}")