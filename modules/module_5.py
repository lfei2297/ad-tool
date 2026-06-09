import streamlit as st
import pandas as pd
import re
import sys
import os
import io
import zipfile  # ✨ 新增：用于内存中打包压缩包

# 路径补丁
current_dir = os.path.dirname(os.path.abspath(__file__))
root_path = os.path.abspath(os.path.join(current_dir, '..'))
if root_path not in sys.path:
    sys.path.insert(0, root_path)

from utils import write_excel_final 

def run(params):
    st.subheader("🚀 模块五：多维素材循环填充与打包拆分中心")
    
    # --- 1. 模式选择与提示信息 ---
    run_mode = st.radio(
        "⚙️ 请选择运行模式",
        ["模式A：固定15行填充 (原有需求)", "模式B：20遍重复并物理切行打包 (ZIP压缩版)"],
        help="模式A固定输出15行大表；模式B生成20遍数据后，按物理行数切分并打包进一个压缩包中下载。"
    )

    if run_mode == "模式A：固定15行填充 (原有需求)":
        st.info("""
        **📌 模式A 运行逻辑：**
        - **15行强制填充**：每一行原始数据固定生成 15 行结果。
        - **循环规则**：不足 15 行时循环重播版本号；超过 15 时自动截断。
        - **备注逻辑**：仅在素材数 > 15 发生截断时显示提醒。
        """)
    else:
        st.info("""
        **📌 模式B 运行逻辑：**
        - **20遍绝对重播**：每个原始行严格铺满 20 个素材周期（A-1, A-2...），汇聚成一份大总表。
        - **物理行拆分与压缩**：大总表生成后，直接按行号切片分流：
          - 1~5 行 ➡️ 新表1.xlsx
          - 6~10 行 ➡️ 新表2.xlsx
          - 11~15 行 ➡️ 新表3.xlsx
          - 16~20 行 ➡️ 新表4.xlsx
          - **全部 5 个 Excel 会自动打成一个 ZIP 压缩包，一键下载！**
        """)

    up = st.file_uploader("📂 上传模块五专用模板", type=["xlsx"], key="m5_up")
    
    if up and st.button("🚀 开始生成", key="m5_btn"):
        df_raw = pd.read_excel(up, dtype=str).fillna("")
        
        display_cols = ["广告账号ID", "主页ID", "像素ID", "真实SKU", "虚拟SKU", "国家", "着陆页版本名称", "广告素材版本名称", "备注"]
        file_prefix = params.get("prefix", "项目_")

        # =====================================================================
        # 核心逻辑一：原有的 15 行强制填充 (模式 A)
        # =====================================================================
        if run_mode == "模式A：固定15行填充 (原有需求)":
            all_results = []
            for idx, row in df_raw.iterrows():
                def safe_int(val, default=0):
                    try:
                        s = str(val).strip()
                        return int(float(s)) if s else default
                    except: return default

                provided_count = safe_int(row.get("广告素材数量", 0))
                base_name = str(row.get("广告素材版本名称", "素材"))
                clean_name = re.sub(r'-\d+$', '', base_name)
                
                if provided_count <= 1:
                    material_pool = [f"{clean_name}-1"]
                else:
                    material_pool = [f"{clean_name}-{i}" for i in range(1, provided_count + 1)]
                
                final_material_list = []
                special_note = ""
                if provided_count > 15:
                    special_note = "素材数超标，仅保留前15个版本"

                if len(material_pool) == 1:
                    final_material_list = material_pool * 15
                else:
                    while len(final_material_list) < 15:
                        for m in material_pool:
                            if len(final_material_list) < 15:
                                final_material_list.append(m)
                            else: break
                
                expanded_rows = []
                for v_name in final_material_list:
                    new_row = row.copy()
                    new_row["广告素材版本名称"] = v_name
                    new_row["备注"] = special_note
                    expanded_rows.append(new_row)
                    
                all_results.append(pd.DataFrame(expanded_rows))

            if all_results:
                final_df = pd.concat(all_results, ignore_index=True)
                final_df = final_df[[c for c in display_cols if c in final_df.columns]]
                
                st.success(f"✅ 模式A 处理完成！")
                xlsx_data = write_excel_final(final_df, "15行填充结果", params, color_by="广告账号ID")
                st.download_button(f"💾 下载模式A结果", data=xlsx_data, file_name=f"{file_prefix}15行强制填充.xlsx")

        # =====================================================================
        # 核心逻辑二：20遍大总表物理切行并打包 ZIP (模式 B)
        # =====================================================================
        else:
            total_expanded_rows = []
            
            # 第一步：先构建全局大总表
            for idx, row in df_raw.iterrows():
                def safe_int(val, default=0):
                    try:
                        s = str(val).strip()
                        return int(float(s)) if s else default
                    except: return default

                provided_count = safe_int(row.get("广告素材数量", 0))
                base_name = str(row.get("广告素材版本名称", "素材"))
                clean_name = re.sub(r'-\d+$', '', base_name)
                
                if provided_count <= 1:
                    material_pool = [f"{clean_name}-1"]
                else:
                    material_pool = [f"{clean_name}-{i}" for i in range(1, provided_count + 1)]
                
                # 严格扩充 20 遍
                for p in range(1, 21):
                    for m in material_pool:
                        new_row = row.copy()
                        new_row["广告素材版本名称"] = m
                        new_row["备注"] = f"第{p}遍生成"
                        total_expanded_rows.append(new_row)

            if total_expanded_rows:
                master_df = pd.DataFrame(total_expanded_rows)
                master_df = master_df[[c for c in display_cols if c in master_df.columns]]
                
                # 创建一个内存中的字节流，用来盛放 ZIP 压缩包
                zip_buffer = io.BytesIO()
                
                # 开始写压缩文件
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    
                    # 1. 压入大总表
                    xlsx_master = write_excel_final(master_df, "大总表结果", params, color_by="广告账号ID")
                    zip_file.writestr(f"{file_prefix}20遍完整大总表.xlsx", xlsx_master)
                    
                    # 2. 按物理行切片并压入 4 个分表
                    slice_ranges = [
                        (1, 0, 5, "1-5行"),
                        (2, 5, 10, "6-10行"),
                        (3, 10, 15, "11-15行"),
                        (4, 15, 20, "16-20行")
                    ]
                    
                    for b_id, start_idx, end_idx, range_name in slice_ranges:
                        sliced_df = master_df.iloc[start_idx:end_idx].copy()
                        if not sliced_df.empty:
                            sliced_df["备注"] = f"物理切分 {range_name}"
                            xlsx_slice = write_excel_final(sliced_df, f"新表{b_id}", params, color_by="广告账号ID")
                            # 将分表塞进压缩包
                            zip_file.writestr(f"{file_prefix}新表{b_id}_{range_name}.xlsx", xlsx_slice)
                
                # 指针回零，准备让用户下载
                zip_buffer.seek(0)
                
                st.success("📦 模式B 总表与4个分表已成功打包！")
                st.download_button(
                    label="🎁 💾 一键下载全部文件 (ZIP 压缩包)",
                    data=zip_buffer,
                    file_name=f"{file_prefix}模式B_20遍全套打包.zip",
                    mime="application/zip",
                    key="m5_zip_btn"
                )