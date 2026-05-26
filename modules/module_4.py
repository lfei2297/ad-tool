import streamlit as st
import pandas as pd
import re
import sys
import os

# 路径补丁：确保能引用到根目录的 utils
current_dir = os.path.dirname(os.path.abspath(__file__))
root_path = os.path.abspath(os.path.join(current_dir, '..'))
if root_path not in sys.path:
    sys.path.insert(0, root_path)

from utils import expand_material_versions, write_excel_final

def run(params):
    st.subheader("🎯 模块四：补齐默认版本 (1:N:N 结构)")
    
    st.info("""
    **功能逻辑：**
    1. 按素材版本号展开明细。
    2. 按『广告组数量 × 每个组的广告数』生成结构。
    3. 若素材不足，自动补齐『默认版本』。
    """)

    up = st.file_uploader("📂 上传原始素材表 (.xlsx)", type=["xlsx"], key="m4_up")
    
    if up and st.button("🚀 开始生成结构", key="m4_btn"):
        df_raw = pd.read_excel(up, dtype=str).fillna("")
        all_series_dfs = []

        for _, row in df_raw.iterrows():
            # --- 1. 获取业务参数 ---
            try:
                provided_count = int(row.get("提供素材版本数量", 1))
                group_count = int(row.get("广告组数量", 1))
                series_count = int(row.get("导品系列数", 1))
                default_fill_count = int(row.get("补充默认版本数", 0))
                # 结构中的“每个组的广告数”由页面左侧的“第一次重复次数”定义
                ads_per_group = params['repeat_1'] 
            except Exception as e:
                st.error(f"行数据格式错误，请检查数值列: {e}")
                continue

            # --- 2. 解析现有素材版本 ---
            lp_match = re.search(r'-(\d+)$', str(row.get("着陆页版本名称", "")))
            lp_v = lp_match.group(1) if lp_match else ""
            
            # 使用 utils 的逻辑展开 1-N
            existing_versions = expand_material_versions(row, lp_v)
            
            # --- 3. 构建该行数据的结构 ---
            # 总共需要的广告行数 = 广告组数 * 每组广告数 * 系列数
            total_needed_rows = group_count * ads_per_group * series_count
            
            expanded_rows = []
            
            # 先填充已有的素材版本
            for v_name in existing_versions:
                # 每个素材版本根据 repeat_1 逻辑进行物理展开
                for _ in range(ads_per_group):
                    new_row = row.copy()
                    new_row["广告素材版本名称"] = v_name
                    expanded_rows.append(new_row)
            
            # --- 4. 补齐默认版本 ---
            # 如果现有行数不足，或者明确要求补充默认版本
            current_len = len(expanded_rows)
            if default_fill_count > 0:
                base_material_name = str(row.get("广告素材版本名称", "素材"))
                # 简单处理：在原名称后加 "_Default"
                default_name = re.sub(r'-\d+-\d+$', '-Default', base_material_name)
                if default_name == base_material_name: default_name += "-Default"
                
                for _ in range(default_fill_count * ads_per_group):
                    d_row = row.copy()
                    d_row["广告素材版本名称"] = default_name
                    # 标记一下这是补充的
                    d_row["备注"] = "系统补齐默认版本"
                    expanded_rows.append(d_row)

            # 转换为 DataFrame
            res_df = pd.DataFrame(expanded_rows)
            
            # 处理第二次重复（系列级复制）
            if params['repeat_2'] > 1:
                res_df = pd.concat([res_df] * params['repeat_2'], ignore_index=True)
            
            all_series_dfs.append(res_df)

        if all_series_dfs:
            final_df = pd.concat(all_series_dfs, ignore_index=True)
            st.success(f"结构生成成功！总行数：{len(final_df)}")
            
            # 导出
            xlsx_data = write_excel_final(final_df, "结构补齐结果", params)
            st.download_button(
                "💾 导出补齐后的结果",
                data=xlsx_data,
                file_name=f"{params['prefix']}结构补齐.xlsx"
            )