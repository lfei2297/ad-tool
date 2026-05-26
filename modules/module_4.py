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

from utils import write_excel_final 

def run(params):
    st.subheader("🎯 模块四：补齐默认版本 (1:M:N 结构优化版)")
    
    st.info("""
    **💡 逻辑说明：**
    - **多组模式 (M > 1)**：组间测素材。每组下的 $N$ 条广告素材相同，消耗 1 个素材版本。
    - **单组模式 (M = 1)**：组内测素材。组内 $N$ 条广告素材各不相同，每个广告位消耗 1 个素材版本。
    """)

    # --- 1. 结构参数输入 ---
    st.markdown("#### 📐 系列结构设置 (1:M:N)")
    col_a, col_b = st.columns(2)
    with col_a:
        M_groups = st.number_input("广告组数量 (M)", min_value=1, value=1)
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
            
            # --- 2. 计算总坑位 ---
            total_required_rows = M_groups * N_ads * series_count
            
            # --- 3. 生成递增版本列表 ---
            base_name = str(row.get("广告素材版本名称", ""))
            clean_name = re.sub(r'-\d+$', '', base_name) # 去掉结尾的 -1
            
            existing_versions = [f"{clean_name}-{i}" for i in range(1, provided_count + 1)]
            
            expanded_rows = []

            # --- 4. 分类讨论核心逻辑 ---
            if M_groups > 1:
                # 【多组模式】：组间不同，组内相同
                # 每个版本展开 N 遍
                for v_name in existing_versions:
                    for _ in range(N_ads):
                        new_row = row.copy()
                        new_row["广告素材版本名称"] = v_name
                        new_row["备注"] = f"多组模式: 组内{N_ads}个素材相同"
                        expanded_rows.append(new_row)
            else:
                # 【单组模式】：组内就不同（M=1）
                # 每个版本只放 1 遍，直接排下去
                for v_name in existing_versions:
                    new_row = row.copy()
                    new_row["广告素材版本名称"] = v_name
                    new_row["备注"] = "单组模式: 组内测不同素材"
                    expanded_rows.append(new_row)

            # --- 5. 结构截断或补齐 ---
            if len(expanded_rows) >= total_required_rows:
                # 如果展开出来的素材比坑位多，按坑位截断
                final_rows = expanded_rows[:total_required_rows]
            else:
                # 如果素材不够填满坑位，补齐“默认版本”
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
            
            st.success(f"处理完成！当前模式：{'多组' if M_groups > 1 else '单组'}，已对齐 {total_required_rows} 个坑位。")
            
            file_prefix = params.get("prefix", "项目A_")
            xlsx_data = write_excel_final(final_df, "结构补齐结果", params)
            
            st.download_button("💾 下载结果", data=xlsx_data, file_name=f"{file_prefix}结构补齐.xlsx")