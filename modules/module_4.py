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
    st.subheader("🎯 模块四：补齐默认版本 (全结构适配版)")
    
    st.info("""
    **💡 运行逻辑说明：**
    - **行排列顺序**：严格遵循“横向交叉”逻辑。系统会先生成所有组的第一个广告（Ad 1），再生成所有组的第二个广告（Ad 2），以此类推。
    - **全结构适配**：无论 $M$ 或 $N$ 是否为 1，系统均能自动对齐坑位，确保导出的表格顺序与分析习惯一致。
    - **自动补齐**：当现有素材不足以填满 $1:M:N$ 结构时，系统将自动填充“默认版本”。
    """)
    
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        M_groups = st.number_input("网页设定广告组数 (M)", min_value=1, value=1)
    with col_b:
        N_ads = st.number_input("网页设定广告数 (N)", min_value=1, value=3)
    with col_c:
        logic_mode = st.radio(
            "🎨 素材分配逻辑", 
            ["组间测素材", "组内测素材"],
            help="""
            【组间测素材】：以“组”为单位消耗素材。同组内的 N 个广告位使用同一个素材版本。
            (例如 1:1:3 结构下，结果为 A-1, A-1, A-1)
            
            【组内测素材】：以“广告位”为单位消耗素材。每一行（每个坑位）都会依次消耗不同的素材。
            (例如 1:1:3 结构下，结果为 A-1, A-2, A-3)
            """,
        )

    up = st.file_uploader("📂 上传模块四专用模板", type=["xlsx"], key="m4_up")
    
    if up and st.button("🚀 开始校验并生成", key="m4_btn"):
        df_raw = pd.read_excel(up, dtype=str).fillna("")
        all_results = []
        warning_logs = []

        for idx, row in df_raw.iterrows():
            line_no = idx + 2
            
            def safe_int(val, default=0):
                try:
                    s = str(val).strip()
                    return int(float(s)) if s else default
                except: return default

            provided_count = safe_int(row.get("提供素材版本数量"))
            series_count = safe_int(row.get("导品系列数"), 1)
            template_padding = safe_int(row.get("补充默认版本数"))
            
            # --- 1. 素材序列生成 ---
            base_name = str(row.get("广告素材版本名称", "素材"))
            clean_name = re.sub(r'-\d+$', '', base_name)
            materials = [f"{clean_name}-{i}" for i in range(1, provided_count + 1)]
            
            final_series_rows = []
            material_cursor = 0 # 素材指针

            for s in range(series_count):
                # --- 2. 核心：构建该系列的素材矩阵 [组m][位n] ---
                matrix = [["默认版本" for _ in range(N_ads)] for _ in range(M_groups)]
                
                if logic_mode == "组间测素材":
                    # 逻辑：每个组消耗一个素材，组内N个位相同
                    for m in range(M_groups):
                        v = materials[material_cursor] if material_cursor < len(materials) else "默认版本"
                        for n in range(N_ads):
                            matrix[m][n] = v
                        material_cursor += 1
                else:
                    # 逻辑：每个广告位都消耗一个素材（不论组）
                    # 按照行生成的物理顺序 [n位][m组] 来分配素材，确保顺序不乱
                    for n in range(N_ads):
                        for m in range(M_groups):
                            v = materials[material_cursor] if material_cursor < len(materials) else "默认版本"
                            matrix[m][n] = v
                            material_cursor += 1

                # --- 3. 按照交叉顺序 (n位优先) 填入结果 ---
                for n in range(N_ads):
                    for m in range(M_groups):
                        new_row = row.copy()
                        v_name = matrix[m][n]
                        new_row["广告素材版本名称"] = v_name
                        new_row["备注"] = "系统自动补齐" if v_name == "默认版本" else ""
                        final_series_rows.append(new_row)

            # --- 4. 补齐校验 ---
            actual_padding = sum(1 for r in final_series_rows if r["广告素材版本名称"] == "默认版本")
            if template_padding != 0 and template_padding != actual_padding:
                warning_logs.append(f"❓ 行 {line_no}: 模板补充数({template_padding}) vs 实际({actual_padding})")

            all_results.append(pd.DataFrame(final_series_rows))

        if warning_logs:
            with st.expander("📝 逻辑校验报告"):
                for log in warning_logs: st.warning(log)
        
        if all_results:
            final_df = pd.concat(all_results, ignore_index=True)
            display_cols = ["广告账号ID", "主页ID", "像素ID", "真实SKU", "虚拟SKU", "国家", "着陆页版本名称", "广告素材版本名称", "备注"]
            final_df = final_df[[c for c in display_cols if c in final_df.columns]]
            
            st.success("✅ 生成完毕！")
            file_prefix = params.get("prefix", "项目_")
            xlsx_data = write_excel_final(final_df, "结构补齐结果", params)
            st.download_button(f"💾 下载：{file_prefix}结果", data=xlsx_data, file_name=f"{file_prefix}补齐.xlsx")