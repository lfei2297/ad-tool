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
    st.subheader("🚀 模块五：通用素材循环填充与分流")
    
    # --- 1. 核心参数配置 ---
    st.markdown("#### ⚙️ 核心参数配置")
    
    # 先渲染模式选择框，并为其指定唯一的 key
    run_mode = st.radio(
        "🎨 导出与拆分模式",
        ["模式A：生成独立总表", "模式B：生成总表并动态平摊拆分"],
        key="m5_run_mode",
        help="模式A会把所有品的数据合在一个大表里导出；模式B会把每个品生成的素材完美平摊到你指定的几张表里，并打包成 ZIP 下载。"
    )

    # ✨ 核心创新点：根据选择的模式，动态给出行数的初始默认值
    default_rows = 15 if "模式A" in run_mode else 20

    # 渲染总行数输入框，value 绑定动态算好的默认值
    target_total_rows = st.number_input(
        "📊 每个品生成总行数", 
        min_value=1, 
        value=default_rows, # 👈 动态初始值绑定在这里
        help="每个独立的产品（行）最终要被扩充撑满的行数。模式A默认15，模式B默认20。"
    )

    sub_table_count = 4
    if "模式B" in run_mode:
        st.markdown("---")
        st.markdown("#### ✂️ 品内拆分段设置")
        sub_table_count = st.number_input(
            "📐 拆分表格数量", 
            min_value=1, 
            max_value=int(target_total_rows), 
            value=4, 
            help="例如希望将素材均匀分流到几张新表里，此处直接填数字即可。"
        )
        
        # 实时平摊预览算法
        base_step = int(target_total_rows // sub_table_count)
        rem = int(target_total_rows % sub_table_count)
        if rem == 0:
            st.caption(f"💡 实时换算：每个品生成的 {target_total_rows} 行将完美平分，每张子表分得 **{base_step}** 行。")
        else:
            st.caption(f"💡 实时换算：前 **{rem}** 张子表每品分得 **{base_step + 1}** 行，后 **{int(sub_table_count) - rem}** 张子表每品分得 **{base_step}** 行。")

    st.markdown("---")
    up = st.file_uploader("📂 上传标准用模板", type=["xlsx"], key="m5_up")
    
    if up and st.button("🚀 开始自动化构建", key="m5_btn"):
        df_raw = pd.read_excel(up, dtype=str).fillna("")
        
        display_cols = ["广告账号ID", "主页ID", "像素ID", "真实SKU", "虚拟SKU", "国家", "着陆页版本名称", "广告素材版本名称", "备注"]
        file_prefix = params.get("prefix", "项目_")
        
        total_expanded_rows = []
        buckets = {i: [] for i in range(1, int(sub_table_count) + 1)}
        
        # 建立动态平摊的基础步长
        base_step = int(target_total_rows // sub_table_count)
        rem = int(target_total_rows % sub_table_count)
        sku_split_lens = [base_step + (1 if i < rem else 0) for i in range(int(sub_table_count))]

        # --- 2. 核心算法：品内切片与循环重播机制 ---
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
            
            this_sku_materials = []
            special_note = ""
            if provided_count > target_total_rows:
                special_note = f"素材数超标，仅保留前{target_total_rows}个版本"

            if len(material_pool) == 1:
                this_sku_materials = material_pool * int(target_total_rows)
            else:
                while len(this_sku_materials) < target_total_rows:
                    for m in material_pool:
                        if len(this_sku_materials) < target_total_rows:
                            this_sku_materials.append(m)
                        else: break
            
            # --- 3. 按平摊长度，进行精准指针分流 ---
            cursor = 0
            for b_idx, length in enumerate(sku_split_lens):
                bucket_id = b_idx + 1
                for _ in range(length):
                    if cursor < len(this_sku_materials):
                        v_name = this_sku_materials[cursor]
                        new_row = row.copy()
                        new_row["广告素材版本名称"] = v_name
                        new_row["备注"] = special_note
                        
                        if "模式B" in run_mode:
                            buckets[bucket_id].append(new_row)
                        cursor += 1

            # 构建总大表的数据
            for v_name in this_sku_materials:
                new_row = row.copy()
                new_row["广告素材版本名称"] = v_name
                new_row["备注"] = special_note
                total_expanded_rows.append(new_row)

        if total_expanded_rows:
            master_df = pd.DataFrame(total_expanded_rows)
            master_df = master_df[[c for c in display_cols if c in master_df.columns]]

            # =====================================================================
            # 分流执行：模式A（直出总大表）
            # =====================================================================
            if "模式A" in run_mode:
                st.success(f"✅ 模式A构建成功！已按照每个品【{target_total_rows}行】平铺填充。")
                xlsx_data = write_excel_final(master_df, "填充总表结果", params, color_by="广告账号ID")
                st.download_button(
                    f"💾 下载：{target_total_rows}行大总表", 
                    data=xlsx_data, 
                    file_name=f"{file_prefix}{target_total_rows}行大总表.xlsx"
                )

            # =====================================================================
            # 分流执行：模式B（智能平摊打包压缩）
            # =====================================================================
            else:
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    
                    # 1. 压入完整大总表
                    xlsx_master = write_excel_final(master_df, "大总表结果", params, color_by="广告账号ID")
                    zip_file.writestr(f"{file_prefix}{target_total_rows}行完整总表.xlsx", xlsx_master)
                    
                    # 2. 压入各子表
                    accumulated_rows = 0
                    for b_id in sorted(buckets.keys()):
                        sliced_rows = buckets[b_id]
                        sliced_df = pd.DataFrame(sliced_rows)
                        sliced_df = sliced_df[[c for c in display_cols if c in sliced_df.columns]]
                        
                        current_len = sku_split_lens[b_id - 1]
                        start_row_num = accumulated_rows + 1
                        end_row_num = accumulated_rows + current_len
                        accumulated_rows += current_len
                        
                        range_text = f"{start_row_num}-{end_row_num}行"
                        sliced_df["备注"] = f"各品内部第 {range_text}"
                        
                        xlsx_slice = write_excel_final(
                            sliced_df, 
                            f"新表{b_id}", 
                            params, 
                            color_by="广告账号ID"
                        )
                        
                        zip_file.writestr(
                            f"{file_prefix}新表{b_id}_各品{range_text}.xlsx", 
                            xlsx_slice
                        )
                
                zip_buffer.seek(0)
                
                st.success(f"📦 模式B品内平摊成功！总计生成 1 张大总表与 {len(buckets)} 张拆分子表。")
                st.download_button(
                    label="🎁 💾 一键下载全套平摊打包件 (ZIP)",
                    data=zip_buffer,
                    file_name=f"{file_prefix}品内平摊_{target_total_rows}行分{sub_table_count}表.zip",
                    mime="application/zip",
                    key="m5_custom_zip_btn"
                )
