import streamlit as st
import pandas as pd
import gc
from utils import write_excel_final, read_uploaded_excel, cycle_repeat

def run(params):
    st.subheader("📦 模块六：多行组合循环填充与分流")
    
    with st.expander("💡 点击查看：模块六运行逻辑说明"):
        st.markdown("""
        - **标准模板完美兼容**：上传标准素材模板后，清洗无关列，保留原生模板说明行。
        - **组合识别机制**：通过“广告账号ID”列智能划分组合。
        - **资产向下填充**：在循环前，自动将首行的各项资产（含出价/竞价）向下填满空缺。
        - **整体循环重播**：将填充完整后的组合作为整体进行循环复制至目标行数。
        """)
    
    st.markdown("#### ⚙️ 核心参数配置")
    target_total_rows = st.number_input("📊 每个组合生成目标总行数", min_value=1, value=30)
    up = st.file_uploader("📂 上传原始模板表格", type=["xlsx"], key="m6_up")
    
    if up and st.button("🚀 开始组合循环构建", key="m6_btn"):
        df_raw = read_uploaded_excel(up.getvalue())
        
        is_hint = df_raw.astype(str).apply(lambda x: x.str.contains('此行为说明|可不填', na=False)).any(axis=1)
        hint_df = df_raw[is_hint].head(1) 
        
        # 仅临时给当前模块传参
        local_params = params.copy()
        local_params['dynamic_hint'] = hint_df 
        
        df_clean = df_raw[~is_hint].reset_index(drop=True)
        if df_clean.empty:
            st.warning("⚠️ 上传的表格中未检测到有效的业务数据，请检查后再试。")
            return
            
        groups = []
        current_group = []
        
        for row_dict in df_clean.to_dict('records'):
            acc_id = str(row_dict.get("广告账号ID", "")).strip()
            if acc_id != "" and current_group:
                groups.append(current_group)
                current_group = []
            current_group.append(row_dict)
            
        if current_group:
            groups.append(current_group)
            
        final_rows = []
        for group in groups:
            group_df = pd.DataFrame(group)
            
            fill_cols = ["广告账号ID", "主页ID", "像素ID", "真实SKU", "虚拟SKU", "国家", "着陆页版本名称", "出价/竞价"]
            for col in fill_cols:
                if col in group_df.columns:
                    group_df[col] = group_df[col].replace("", None).ffill().fillna("")
            
            filled_group = group_df.to_dict('records')
            expanded_group = cycle_repeat(filled_group, target_total_rows)
            final_rows.extend(expanded_group)
            
        if final_rows:
            final_df = pd.DataFrame(final_rows)
            st.success(f"✅ 模块六处理完成！成功识别出 {len(groups)} 个产品组合并填充至 {target_total_rows} 行！")
            
            file_prefix = params.get("prefix", "项目_")
            # 传入隔离的 local_params 生成结果
            xlsx_data = write_excel_final(final_df, "组合填充结果", local_params)
            
            st.download_button(
                f"💾 下载：{target_total_rows}行组合总表", 
                data=xlsx_data, 
                file_name=f"{file_prefix}组合填充_{target_total_rows}行.xlsx"
            )
            
            del df_raw, df_clean, final_rows, final_df, groups
            gc.collect()