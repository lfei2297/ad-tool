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

from utils import expand_material_versions, write_excel_final

def run(params):
    st.subheader("🎯 模块四：补齐默认版本 (物理展开 + 结构对齐)")
    
    st.info("""
    **💡 运行逻辑：**
    1. **第一步**：按照模块一逻辑，将每个素材版本物理展开 $N$ 次（组内广告数）。
    2. **第二步**：计算总坑位需求 = $M(组) \times N(广告) \times S(系列)$。
    3. **第三步**：如果展开后的行数不足总坑位，用『默认版本』补齐缺口。
    """)

    # --- 1. 结构参数输入 ---
    st.markdown("#### 📐 系列结构设置 (1:M:N)")
    col_a, col_b = st.columns(2)
    with col_a:
        M_groups = st.number_input("广告组数量 (M)", min_value=1, value=2, help="每个系列下的广告组数")
    with col_b:
        N_ads = st.number_input("组内广告数 (N)", min_value=1, value=2, help="每个广告组下的广告位数量")

    up = st.file_uploader("📂 上传模块四专用模板", type=["xlsx"], key="m4_up")
    
    if up and st.button("🚀 开始生成", key="m4_btn"):
        df_raw = pd.read_excel(up, dtype=str).fillna("")
        all_results = []

        for _, row in df_raw.iterrows():
            try:
                provided_count = int(row.get("提供素材版本数量", 0))
                series_count = int(row.get("导品系列数", 1))
            except:
                st.error("请检查模板中的数字列是否为有效数字")
                continue
            
            # --- 核心计算 ---
            # 总共需要的行数 (总坑位)
            total_required_rows = M_groups * N_ads * series_count
            
            # --- 第一步：物理展开已有的素材 ---
            lp_match = re.search(r'-(\d+)$', str(row.get("着陆页版本名称", "")))
            lp_v = lp_match.group(1) if lp_match else ""
            
            # 生成版本列表 [A-1, A-2, A-3]
            existing_versions = expand_material_versions(row, lp_v)
            
            expanded_rows = []
            for v_name in existing_versions:
                # 每个素材版本按 N 展开（满足组内广告数）
                for _ in range(N_ads):
                    new_row = row.copy()
                    new_row["广告素材版本名称"] = v_name
                    expanded_rows.append(new_row)
            
            # --- 第二步：截断或补齐 ---
            current_count = len(expanded_rows)
            
            if current_count >= total_required_rows:
                # 如果素材多了，按照坑位截断
                final_rows = expanded_rows[:total_required_rows]
            else:
                # 如果素材少了，补齐缺口
                final_rows = expanded_rows
                gap = total_required_rows - current_count
                for _ in range(gap):
                    d_row = row.copy()
                    d_row["广告素材版本名称"] = "默认版本"
                    d_row["备注"] = "系统补齐默认版本"
                    final_rows.append(d_row)

            all_results.append(pd.DataFrame(final_rows))

        if all_results:
            final_df = pd.concat(all_results, ignore_index=True)
            
            # 过滤展示列
            cols_to_keep = ["广告账号ID", "主页ID", "像素ID", "真实SKU", "虚拟SKU", "国家", "着陆页版本名称", "广告素材版本名称", "备注"]
            final_df = final_df[[c for c in cols_to_keep if c in final_df.columns]]
            
            st.success(f"处理完成！目标坑位: {total_required_rows}，已自动对齐。")
            
            file_prefix = params.get("prefix", "项目A_")
            xlsx_data = write_excel_final(final_df, "结构补齐结果", params)
            
            st.download_button(
                "💾 下载结果", 
                data=xlsx_data, 
                file_name=f"{file_prefix}结构补齐.xlsx"
            )