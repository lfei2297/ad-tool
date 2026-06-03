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
    st.subheader("🚀 模块五：15行强制循环填充 (按账号配色版)")
    
    st.info("""
    **⚙️ 运行逻辑：**
    - **15行强制填充**：每行原始数据固定生成 15 行结果。
    - **素材循环**：素材不足 15 时从 -1 开始重播；超过 15 时截断。
    - **🎨 颜色区分**：导出文件将按照 **“广告账号ID”** 进行交替着色，方便区分不同账号。
    """)

    up = st.file_uploader("📂 上传标准模板", type=["xlsx"], key="m5_up")
    
    if up and st.button("🚀 开始生成", key="m5_btn"):
        df_raw = pd.read_excel(up, dtype=str).fillna("")
        all_results = []

        for idx, row in df_raw.iterrows():
            def safe_int(val, default=0):
                try:
                    s = str(val).strip()
                    return int(float(s)) if s else default
                except: return default

            provided_count = safe_int(row.get("广告素材数量", 0))
            
            # --- 1. 素材版本号准备 ---
            base_name = str(row.get("广告素材版本名称", "素材"))
            clean_name = re.sub(r'-\d+$', '', base_name)
            
            # 基础素材池
            if provided_count <= 1:
                # 只有1个素材的情况
                material_pool = [f"{clean_name}-1"]
            else:
                # 多个素材的情况
                material_pool = [f"{clean_name}-{i}" for i in range(1, provided_count + 1)]
            
            # --- 2. 构造 15 行的素材序列 ---
            final_material_list = []
            
            if len(material_pool) == 1:
                # 需求①：素材数为1，生成15个相同的
                final_material_list = material_pool * 15
            else:
                # 需求②：素材数>1，循环填充
                while len(final_material_list) < 15:
                    # 将素材池加入列表
                    for m in material_pool:
                        if len(final_material_list) < 15:
                            final_material_list.append(m)
                        else:
                            break
            
            # --- 3. 构建结果行 ---
            expanded_rows = []
            for v_name in final_material_list:
                new_row = row.copy()
                new_row["广告素材版本名称"] = v_name
                new_row["备注"] = "15行强制填充"
                expanded_rows.append(new_row)
                
            all_results.append(pd.DataFrame(expanded_rows))

        if all_results:
            final_df = pd.concat(all_results, ignore_index=True)
            # 保持导出列与之前模块一致
            display_cols = ["广告账号ID", "主页ID", "像素ID", "真实SKU", "虚拟SKU", "国家", "着陆页版本名称", "广告素材版本名称", "备注"]
            final_df = final_df[[c for c in display_cols if c in final_df.columns]]
            
            st.success(f"✅ 处理完成！已按 15 行规格展开。")
            
            # 获取侧边栏前缀
            file_prefix = params.get("prefix", "项目_")
            
            # ✨ 核心修改点：指定 color_by 参数为 "广告账号ID"
            # 如果你的 utils.py 中的 write_excel_final 支持这个参数，请务必传入
            xlsx_data = write_excel_final(
                final_df, 
                "15行填充结果", 
                params, 
                color_by="广告账号ID" # 👈 修改配色基准为账号
            )
            
            st.download_button(
                f"💾 下载：{file_prefix}模块五结果", 
                data=xlsx_data, 
                file_name=f"{file_prefix}15行填充.xlsx"
            )