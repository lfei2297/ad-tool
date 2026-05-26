import streamlit as st
import pandas as pd
import re

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
            # 1. 读取模板特有的参数
            provided_count = int(row.get("提供素材版本数量", 1))
            group_count = int(row.get("广告组数量", 1))
            series_count = int(row.get("导品系列数", 1))
            default_fill_count = int(row.get("补充默认版本数", 0))
            
            # 这里的 repeat_1 对应 1:N:N 结构中最后一个 N（每个组的广告数）
            ads_per_group = params['repeat_1'] 

            # 2. 生成 1-N 的素材版本（例如 3-1, 3-2, 3-3）
            # 我们直接根据“提供素材版本数量”构造列表
            base_name = str(row.get("广告素材版本名称", ""))
            # 提取前缀，比如从 "xxx-3-1" 提取 "xxx-3-"
            prefix_match = re.search(r'^(.*-)\d+$', base_name)
            name_prefix = prefix_match.group(1) if prefix_match else (base_name + "-")
            
            existing_versions = [f"{name_prefix}{i}" for i in range(1, provided_count + 1)]
            
            # 3. 按照结构展开
            expanded_rows = []
            
            # 先放正常素材
            for v_name in existing_versions:
                for _ in range(ads_per_group):
                    new_row = row.copy()
                    new_row["广告素材版本名称"] = v_name
                    expanded_rows.append(new_row)
            
            # 4. 补齐默认版本 (假设默认版本命名为 xxx-Default)
            default_name = name_prefix + "Default"
            for _ in range(default_fill_count * ads_per_group):
                d_row = row.copy()
                d_row["广告素材版本名称"] = default_name
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