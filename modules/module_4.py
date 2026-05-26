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
            # 获取业务参数
            try:
                provided_count = int(row.get("提供素材版本数量", 1))
                default_fill_count = int(row.get("补充默认版本数", 0))
                ads_per_group = params['repeat_1'] # 对应 1:N:N 结构中的最后广告数
            except:
                st.error("请确保数值列（提供数量、补充数量等）为数字格式")
                continue

            # --- 1. 生成正常素材版本 (例如 3-1, 3-2, 3-3) ---
            base_name = str(row.get("广告素材版本名称", ""))
            # 提取前缀（如从 xxx-3-1 提取出 xxx-3-）
            prefix_match = re.search(r'^(.*-)\d+$', base_name)
            name_prefix = prefix_match.group(1) if prefix_match else (base_name + "-")
            
            existing_versions = [f"{name_prefix}{i}" for i in range(1, provided_count + 1)]
            
            expanded_rows = []
            
            # 填充正常素材
            for v_name in existing_versions:
                for _ in range(ads_per_group):
                    new_row = row.copy()
                    new_row["广告素材版本名称"] = v_name
                    expanded_rows.append(new_row)
            
            # --- 2. 补齐默认版本 ---
            # ✨ 核心修正：补齐行的素材名称直接命名为“默认版本”
            for _ in range(default_fill_count * ads_per_group):
                d_row = row.copy()
                d_row["广告素材版本名称"] = "默认版本" 
                d_row["备注"] = "系统补齐默认版本"
                expanded_rows.append(d_row)

            # 转换为 DataFrame
            res_df = pd.DataFrame(expanded_rows)
            
            # 处理重复次数 (repeat_2 对应系列级复制)
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