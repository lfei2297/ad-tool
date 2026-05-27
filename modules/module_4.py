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
    st.subheader("🎯 模块四：补齐默认版本 ")
    
    st.info("""
    **💡 运行逻辑：**
    - **多组模式 (M > 1)**：组间测素材。每组下的 $N$ 条广告素材相同，消耗 1 个素材版本。
    - **单组模式 (M = 1)**：组内测素材。组内 $N$ 条广告素材各不相同，每个广告位消耗 1 个素材版本。
    - **自动补齐**：系统根据总坑位需求自动填充“默认版本”。
    """)
    
    # --- 1. 结构参数输入 ---
    st.markdown("#### 📐 当前计划系列结构 (1:M:N)")
    col_a, col_b = st.columns(2)
    with col_a:
        M_groups = st.number_input("网页设定广告组数 (M)", min_value=1, value=1)
    with col_b:
        N_ads = st.number_input("网页设定广告数 (N)", min_value=1, value=1)

    up = st.file_uploader("📂 上传模块四专用模板", type=["xlsx"], key="m4_up")
    
    if up and st.button("🚀 开始校验并生成", key="m4_btn"):
        # 保持空值为 ""，不填充 "0"
        df_raw = pd.read_excel(up, dtype=str).fillna("")
        all_results = []
        warning_logs = []

        for idx, row in df_raw.iterrows():
            line_no = idx + 2
            
            # 安全数字转换
            def safe_int(val, default=0):
                try:
                    s = str(val).strip()
                    return int(float(s)) if s else default
                except:
                    return default

            provided_count = safe_int(row.get("提供素材版本数量"))
            series_count = safe_int(row.get("导品系列数"), 1)
            template_padding = safe_int(row.get("补充默认版本数"))
            template_m = safe_int(row.get("广告组数量")) # 表格填写的组数
            
            # --- ✨ 关键校验：对比网页 M 与表格 M ---
            if template_m != 0 and template_m != M_groups:
                warning_logs.append(f"⚠️ 行 {line_no}: 表格填写组数({template_m}) 与网页设定({M_groups}) 不符，已按网页设定执行。")

            # --- 2. 核心逻辑：计算坑位与素材分配 ---
            total_slots = M_groups * N_ads * series_count 
            
            base_name = str(row.get("广告素材版本名称", "素材"))
            clean_name = re.sub(r'-\d+$', '', base_name)
            materials = [f"{clean_name}-{i}" for i in range(1, provided_count + 1)]
            
            final_rows_data = []
            material_idx = 0
            
            # 坑位驱动填充
            for s in range(series_count):
                for m in range(M_groups):
                    current_material = "默认版本"
                    is_padding = True
                    
                    if material_idx < len(materials):
                        current_material = materials[material_idx]
                        is_padding = False
                    
                    for n in range(N_ads):
                        new_row = row.copy()
                        new_row["广告素材版本名称"] = current_material
                        new_row["备注"] = "系统自动补齐" if is_padding else ""
                        final_rows_data.append(new_row)
                        
                        if M_groups == 1:
                            material_idx += 1
                            if material_idx < len(materials):
                                current_material = materials[material_idx]
                                is_padding = False
                            else:
                                current_material = "默认版本"
                                is_padding = True
                    
                    if M_groups > 1:
                        material_idx += 1

            # --- 3. 校验：补充数验证 ---
            actual_padding_count = sum(1 for r in final_rows_data if r["广告素材版本名称"] == "默认版本")
            if template_padding != 0 and template_padding != actual_padding_count:
                warning_logs.append(f"❓ 行 {line_no}: 表格填写补充数({template_padding}) 与系统实际补齐行数({actual_padding_count}) 不符。")

            all_results.append(pd.DataFrame(final_rows_data))

        # --- 4. 冲突报告展示 ---
        if warning_logs:
            with st.expander("📝 逻辑校验报告 (发现配置冲突)"):
                for log in warning_logs:
                    st.warning(log)
        
        if all_results:
            final_df = pd.concat(all_results, ignore_index=True)
            display_cols = ["广告账号ID", "主页ID", "像素ID", "真实SKU", "虚拟SKU", "国家", "着陆页版本名称", "广告素材版本名称", "备注"]
            final_df = final_df[[c for c in display_cols if c in final_df.columns]]
            
            st.success(f"✅ 处理完成！")
            file_prefix = params.get("prefix", "项目_")
            xlsx_data = write_excel_final(final_df, "结构补齐结果", params)
            st.download_button(f"💾 下载：{file_prefix}结果", data=xlsx_data, file_name=f"{file_prefix}结构补齐.xlsx")