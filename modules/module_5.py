import streamlit as st
import pandas as pd
import re
import sys
import os
import io
import zipfile

# 路径补丁
current_dir = os.path.dirname(os.path.abspath(__file__))
root_path = os.path.abspath(os.path.join(current_dir, '..'))
if root_path not in sys.path:
    sys.path.insert(0, root_path)

from utils import write_excel_final 

def run(params):
    st.subheader("🚀 模块五：通用素材循环填充与智能分流中心")
    
    # --- 1. 全自定义 UI 参数配置 ---
    st.markdown("#### ⚙️ 核心参数配置")
    col_x, col_y = st.columns(2)
    with col_x:
        target_total_rows = st.number_input(
            "📊 目标生成总行数", 
            min_value=1, 
            value=15, 
            help="每行原始数据最终要被扩充撑满的总行数。例如输入 15 或 20。"
        )
    with col_y:
        run_mode = st.radio(
            "🎨 导出与拆分模式",
            ["模式一：仅生成独立总表", "模式二：生成总表并按参数拆分成 N 张新表"],
            help="模式一直接导出一个大表；模式二会按照你设定的规模自动切分并打包成 ZIP 压缩包。"
        )

    # 如果是模式二，动态展示拆分参数框
    split_size = 5
    if run_mode == "模式二：生成总表并按参数拆分成 N 张新表":
        st.markdown("---")
        st.markdown("#### ✂️ 拆分段设置")
        split_size = st.number_input(
            "📐 拆分颗粒度 (每 N 行切一刀)", 
            min_value=1, 
            max_value=int(target_total_rows), 
            value=5, 
            help="例如总行数 20，颗粒度填 5，则会自动切分成 1-5行, 6-10行, 11-15行, 16-20行共 4 张子表。"
        )

    st.markdown("---")
    up = st.file_uploader("📂 上传模块五专用模板", type=["xlsx"], key="m5_up")
    
    if up and st.button("🚀 开始自动化构建", key="m5_btn"):
        df_raw = pd.read_excel(up, dtype=str).fillna("")
        
        display_cols = ["广告账号ID", "主页ID", "像素ID", "真实SKU", "虚拟SKU", "国家", "着陆页版本名称", "广告素材版本名称", "备注"]
        file_prefix = params.get("prefix", "项目_")
        
        total_expanded_rows = []

        # --- 2. 核心算法：通用无限循环填充机制 ---
        for idx, row in df_raw.iterrows():
            def safe_int(val, default=0):
                try:
                    s = str(val).strip()
                    return int(float(s)) if s else default
                except: return default

            provided_count = safe_int(row.get("广告素材数量", 0))
            base_name = str(row.get("广告素材版本名称", "素材"))
            clean_name = re.sub(r'-\d+$', '', base_name)
            
            # 构建素材基础池
            if provided_count <= 1:
                material_pool = [f"{clean_name}-1"]
            else:
                material_pool = [f"{clean_name}-{i}" for i in range(1, provided_count + 1)]
            
            final_material_list = []
            special_note = ""
            
            # 【截断备注逻辑】仅在用户定义的总行数少于素材原本数量时触发
            if provided_count > target_total_rows:
                special_note = f"素材数超标，仅保留前{target_total_rows}个版本"

            if len(material_pool) == 1:
                # 只有一个素材，直接平铺填满自定义行数
                final_material_list = material_pool * int(target_total_rows)
            else:
                # 多个素材，无限循环直到达到自定义的目标总行数
                while len(final_material_list) < target_total_rows:
                    for m in material_pool:
                        if len(final_material_list) < target_total_rows:
                            final_material_list.append(m)
                        else:
                            break
            
            # 装配行数据
            for v_name in final_material_list:
                new_row = row.copy()
                new_row["广告素材版本名称"] = v_name
                new_row["备注"] = special_note
                total_expanded_rows.append(new_row)

        if total_expanded_rows:
            master_df = pd.DataFrame(total_expanded_rows)
            master_df = master_df[[c for c in display_cols if c in master_df.columns]]

            # =====================================================================
            # 分流执行：模式一（直出总大表）
            # =====================================================================
            if run_mode == "模式一：仅生成独立总表":
                st.success(f"✅ 总表构建成功！已按照自定义目标【{target_total_rows}行】平铺填充。")
                xlsx_data = write_excel_final(master_df, "填充总表结果", params, color_by="广告账号ID")
                st.download_button(
                    f"💾 下载：{target_total_rows}行大总表", 
                    data=xlsx_data, 
                    file_name=f"{file_prefix}{target_total_rows}行大总表.xlsx"
                )

            # =====================================================================
            # 分流执行：模式二（智能物理切片并压缩打包）
            # =====================================================================
            else:
                # 建立内存 ZIP 压缩包
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    
                    # 1. 压入完整的大总表
                    xlsx_master = write_excel_final(master_df, "大总表结果", params, color_by="广告账号ID")
                    zip_file.writestr(f"{file_prefix}{target_total_rows}行完整总表.xlsx", xlsx_master)
                    
                    # 2. 动态数学切片算法
                    total_len = len(master_df)
                    step = int(split_size)
                    table_counter = 1
                    
                    # 使用 range 步长进行横向动态切豆腐块
                    for start_idx in range(0, total_len, step):
                        end_idx = min(start_idx + step, total_len)
                        sliced_df = master_df.iloc[start_idx:end_idx].copy()
                        
                        if not sliced_df.empty:
                            range_text = f"{start_idx + 1}-{end_idx}行"
                            sliced_df["备注"] = f"动态分流 {range_text}"
                            
                            xlsx_slice = write_excel_final(
                                sliced_df, 
                                f"新表{table_counter}", 
                                params, 
                                color_by="广告账号ID"
                            )
                            
                            # 将动态生成的子表塞进压缩包
                            zip_file.writestr(
                                f"{file_prefix}新表{table_counter}_{range_text}.xlsx", 
                                xlsx_slice
                            )
                            table_counter += 1
                
                zip_buffer.seek(0)
                
                st.success(f"📦 动态分流完成！总计生成 1 张大总表与 {table_counter - 1} 张自定义切片子表。")
                st.download_button(
                    label="🎁 💾 一键下载全套打包件 (ZIP)",
                    data=zip_buffer,
                    file_name=f"{file_prefix}填充{target_total_rows}_切片{split_size}打包.zip",
                    mime="application/zip",
                    key="m5_custom_zip_btn"
                )