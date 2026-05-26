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
    st.subheader("🎯 模块四：补齐默认版本 (1:M:N 结构优化版)")
    
    st.info("""
    **💡 逻辑说明：**
    - **多组模式 (M > 1)**：组间测素材。每组消耗 1 个素材，组内 $N$ 个广告素材相同。
    - **单组模式 (M = 1)**：组内测素材。每个广告位消耗 1 个素材。
    - **自动补齐**：系统将根据 $S \\times M$ (若M>1) 或 $S \\times N$ (若M=1) 自动补足『默认版本』。
    """)

    # --- 1. 结构参数输入 ---
    st.markdown("#### 📐 系列结构设置 (1:M:N)")
    col_m, col_n = st.columns(2)
    with col_a:
        M_groups = st.number_input("广告组数量 (M)", min_value=1, value=1, help="每个系列下的广告组数")
    with col_b:
        N_ads = st.number_input("组内广告数 (N)", min_value=1, value=2, help="每个广告组下的广告位元数")

    up = st.file_uploader("📂 上传模块四专用模板", type=["xlsx"], key="m4_up")
    
    if up and st.button("🚀 生成补齐结构", key="m4_btn"):
        df_raw = pd.read_excel(up, dtype=str).fillna("")
        all_results = []

        for _, row in df_raw.iterrows():
            # 获取基础数据
            provided_count = int(row.get("提供素材版本数量", 0))
            series_count = int(row.get("导品系列数", 1))
            
            # --- 2. 算法核心：分类讨论计算总坑位 ---
            if M_groups > 1:
                # 模式 A：以组为单位消耗素材
                total_slots = series_count * M_groups
                gap = max(0, total_slots - provided_count)
                final_padding_rows = gap * N_ads
                multiplier = N_ads  # 每个素材需要物理展开 N 遍（组内相同）
                mode_desc = f"多组模式(M={M_groups}): 组间测试"
            else:
                # 模式 B：以广告位为单位消耗素材
                total_slots = series_count * N_ads
                gap = max(0, total_slots - provided_count)
                final_padding_rows = gap
                multiplier = 1 # 每个素材只占一位
                mode_desc = f"单组模式(M=1): 组内测试"

            # --- 3. 处理现有素材 ---
            lp_match = re.search(r'-(\d+)$', str(row.get("着陆页版本名称", "")))
            lp_v = lp_match.group(1) if lp_match else ""
            
            # 解析出 3-1, 3-2 等版本名
            existing_versions = expand_material_versions(row, lp_v)
            
            expanded_rows = []
            # 限制已有素材：不能超过总坑位
            limited_versions = existing_versions[:total_slots]
            
            for v_name in limited_versions:
                for _ in range(multiplier):
                    new_row = row.copy()
                    new_row["广告素材版本名称"] = v_name
                    new_row["备注"] = f"{mode_desc} | 组间素材去重，组内素材相同" if M_groups > 1 else "单组组内素材测试"
                    expanded_rows.append(new_row)
            
            # --- 4. 执行补齐 (Final Padding) ---
            for _ in range(final_padding_rows):
                d_row = row.copy()
                d_row["广告素材版本名称"] = "默认版本"
                d_row["备注"] = "系统补齐默认版本"
                expanded_rows.append(d_row)

            all_results.append(pd.DataFrame(expanded_rows))

        if all_results:
            final_df = pd.concat(all_results, ignore_index=True)
            # 强制删除模板中的业务参数列，保持结果纯净
            cols_to_keep = ["广告账号ID", "主页ID", "像素ID", "真实SKU", "虚拟SKU", "国家", "着陆页版本名称", "广告素材版本名称", "备注"]
            final_df = final_df[[c for c in cols_to_keep if c in final_df.columns]]
            
            st.success(f"结构补齐完成！共生成 {len(final_df)} 行数据。")
            
            # 借用 params 传递样式开关
            params = {"prefix": row.get("FILE_PREFIX", "M4_"), "enable_color": True, "fast_mode": False}
            xlsx_data = write_excel_final(final_df, "结构补齐结果", params)
            
            st.download_button("💾 下载补齐后的结果", data=xlsx_data, file_name="模块四_补齐结果.xlsx")