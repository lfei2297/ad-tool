import streamlit as st
import pandas as pd
import re
import sys
import os

# 路径补丁
current_dir = os.path.dirname(os.path.abspath(__file__))
root_path = os.path.abspath(os.path.join(current_dir, '..'))
if root_path not in sys.path:
    sys.path.insert(0, root_path)

from utils import write_excel_final # 只引用导出函数

def run(params):
    st.subheader("🎯 模块四：补齐默认版本 (物理展开 + 自动递增)")
    
    st.markdown("#### 📐 系列结构设置 (1:M:N)")
    col_a, col_b = st.columns(2)
    with col_a:
        M_groups = st.number_input("广告组数量 (M)", min_value=1, value=2)
    with col_b:
        N_ads = st.number_input("组内广告数 (N)", min_value=1, value=2)

    up = st.file_uploader("📂 上传模块四专用模板", type=["xlsx"], key="m4_up")
    
    if up and st.button("🚀 开始生成", key="m4_btn"):
        df_raw = pd.read_excel(up, dtype=str).fillna("")
        all_results = []

        for _, row in df_raw.iterrows():
            try:
                provided_count = int(row.get("提供素材版本数量", 0))
                series_count = int(row.get("导品系列数", 1))
            except:
                st.error("请检查模板中的数字列是否正确")
                continue
            
            total_required_rows = M_groups * N_ads * series_count
            
            # --- ✨ 核心修正：手动生成递增序列 ---
            base_name = str(row.get("广告素材版本名称", ""))
            
            # 逻辑：如果名字里有 -1，去掉它作为前缀；如果没有，直接当缀
            # 目标是生成类似 "名字-1", "名字-2"
            clean_name = re.sub(r'-\d+$', '', base_name) 
            
            existing_versions = []
            for i in range(1, provided_count + 1):
                v_name = f"{clean_name}-{i}"
                existing_versions.append(v_name)
            
            # --- 物理展开 ---
            expanded_rows = []
            for v_name in existing_versions:
                for _ in range(N_ads): # 每个版本展 N 次
                    new_row = row.copy()
                    new_row["广告素材版本名称"] = v_name
                    expanded_rows.append(new_row)
            
            # --- 结构截断或补齐 ---
            if len(expanded_rows) >= total_required_rows:
                final_rows = expanded_rows[:total_required_rows]
            else:
                final_rows = expanded_rows
                gap = total_required_rows - len(expanded_rows)
                for _ in range(gap):
                    d_row = row.copy()
                    d_row["广告素材版本名称"] = "默认版本"
                    d_row["备注"] = "系统补齐默认版本"
                    final_rows.append(d_row)

            all_results.append(pd.DataFrame(final_rows))

        if all_results:
            final_df = pd.concat(all_results, ignore_index=True)
            cols_to_keep = ["广告账号ID", "主页ID", "像素ID", "真实SKU", "虚拟SKU", "国家", "着陆页版本名称", "广告素材版本名称", "备注"]
            final_df = final_df[[c for c in cols_to_keep if c in final_df.columns]]
            
            st.success(f"处理完成！已生成递增版本并对齐 {total_required_rows} 个坑位。")
            
            file_prefix = params.get("prefix", "项目A_")
            xlsx_data = write_excel_final(final_df, "结构补齐结果", params)
            
            st.download_button(
                "💾 下载结果", 
                data=xlsx_data, 
                file_name=f"{file_prefix}结构补齐.xlsx"
            )