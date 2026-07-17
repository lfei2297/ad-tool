import streamlit as st
import pandas as pd
import sys
import os

# --- 路径补丁 ---
current_dir = os.path.dirname(os.path.abspath(__file__))
root_path = os.path.abspath(os.path.join(current_dir, '..'))
if root_path not in sys.path:
    sys.path.insert(0, root_path)

from utils import write_excel_final 

def run(params):
    st.subheader("📦 模块六：多行组合循环填充与分流")
    
    with st.expander("💡 点击查看：模块六运行逻辑说明"):
        st.markdown("""
        - **标准模板完美兼容**：上传标准素材模板后，系统会自动清洗无关列，并100%保留原生模板说明行。
        - **组合识别机制**：系统通过“广告账号ID”列智能划分组合。只要出现新值，即代表新组合开始。
        - **资产向下填充**：在循环前，会自动将首行的各项资产（含出价/竞价）向下填满该组的空缺单元格。
        - **整体循环重播**：将填充完整后的组合作为整体进行循环复制至目标行数。
        """)
    
    # --- 1. 参数配置 ---
    st.markdown("#### ⚙️ 核心参数配置")
    target_total_rows = st.number_input("📊 每个组合生成目标总行数", min_value=1, value=30)
    up = st.file_uploader("📂 上传原始模板表格", type=["xlsx"], key="m6_up")
    
    if up and st.button("🚀 开始组合循环构建", key="m6_btn"):
        df_raw = pd.read_excel(up, dtype=str).fillna("")
        
        # ==========================================
        # ✨ 进门杀毒 + 动态提取说明行引擎
        # ==========================================
        is_hint = df_raw.astype(str).apply(lambda x: x.str.contains('此行为说明|可不填', na=False)).any(axis=1)
        hint_df = df_raw[is_hint].head(1) 
        
        # 将原汁原味的说明行塞进 params，让 utils.py 在最后一步拼回！
        params['dynamic_hint'] = hint_df 
        
        # 剥离说明行，提取干净的业务数据
        df_clean = df_raw[~is_hint].reset_index(drop=True)
        
        if df_clean.empty:
            st.warning("⚠️ 上传的表格中未检测到有效的业务数据，请检查后再试。")
            return
            
        # --- 2. 按账号ID对多行进行“组合切片” ---
        groups = []
        current_group = []
        
        for idx, row in df_clean.iterrows():
            acc_id = str(row.get("广告账号ID", "")).strip()
            if acc_id != "" and current_group:
                groups.append(current_group)
                current_group = []
            current_group.append(row.to_dict())
            
        if current_group:
            groups.append(current_group)
            
        # --- 3. 组合内向下填充 + 整体重播 ---
        final_rows = []
        for g_idx, group in enumerate(groups):
            group_df = pd.DataFrame(group)
            
            # ✨ 优化：已将“出价/竞价”加入自动填满列表
            fill_cols = ["广告账号ID", "主页ID", "像素ID", "真实SKU", "虚拟SKU", "国家", "着陆页版本名称", "出价/竞价"]
            for col in fill_cols:
                if col in group_df.columns:
                    group_df[col] = group_df[col].replace("", None).ffill().fillna("")
            
            filled_group = group_df.to_dict('records')
            expanded_group = []
            
            while len(expanded_group) < target_total_rows:
                for row_dict in filled_group:
                    if len(expanded_group) < target_total_rows:
                        expanded_group.append(row_dict.copy())
                    else:
                        break
            final_rows.extend(expanded_group)
            
        if final_rows:
            final_df = pd.DataFrame(final_rows)
            st.success(f"✅ 模块六处理完成！成功识别出 {len(groups)} 个产品组合并填充至 {target_total_rows} 行！")
            
            file_prefix = params.get("prefix", "项目_")
            xlsx_data = write_excel_final(final_df, "组合填充结果", params)
            
            st.download_button(
                f"💾 下载：{target_total_rows}行组合总表", 
                data=xlsx_data, 
                file_name=f"{file_prefix}组合填充_{target_total_rows}行.xlsx"
            )