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
    st.subheader("🎯 模块四：补齐默认版本 (严格交叉审计版)")
    
    st.info("""
    **💡 运行逻辑：**
    - **行顺序**：严格按照“横向交叉”排列。先生成所有组的 Ad 1，再生成所有组的 Ad 2。
    - **组间测素材**：组内 $N$ 个广告素材相同，消耗 1 个素材版本/组。
    - **组内测素材**：组内 $N$ 个广告素材各不相同，消耗 1 个素材版本/广告位。
    """)
    
    # --- 1. 结构参数输入 ---
    st.markdown("#### 📐 当前计划系列结构 (1:M:N)")
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        M_groups = st.number_input("网页设定广告组数 (M)", min_value=1, value=1)
    with col_b:
        N_ads = st.number_input("网页设定广告数 (N)", min_value=1, value=2)
    with col_c:
        logic_mode = st.radio(
            "🎨 素材分配逻辑", 
            ["组间测素材", "组内测素材"],
            help="组间测：组内素材相同；组内测：每个位素材都不同"
        )

    up = st.file_uploader("📂 上传模块四专用模板", type=["xlsx"], key="m4_up")
    
    if up and st.button("🚀 开始校验并生成", key="m4_btn"):
        # 保持空值为 ""，不填充 "0"
        df_raw = pd.read_excel(up, dtype=str).fillna("")
        all_results = []
        warning_logs = []

        for idx, row in df_raw.iterrows():
            line_no = idx + 2
            
            # 安全数字转换函数
            def safe_int(val, default=0):
                try:
                    s = str(val).strip()
                    return int(float(s)) if s else default
                except:
                    return default

            provided_count = safe_int(row.get("提供素材版本数量"))
            series_count = safe_int(row.get("导品系列数"), 1)
            template_padding = safe_int(row.get("补充默认版本数"))
            template_m = safe_int(row.get("广告组数量"))
            
            # --- 逻辑校验 ---
            if template_m != 0 and template_m != M_groups:
                warning_logs.append(f"⚠️ 行 {line_no}: 表格组数({template_m}) 与网页设定({M_groups}) 不符。")

            # --- 2. 素材准备与分配 ---
            base_name = str(row.get("广告素材版本名称", "素材"))
            clean_name = re.sub(r'-\d+$', '', base_name)
            materials = [f"{clean_name}-{i}" for i in range(1, provided_count + 1)]
            
            material_idx = 0
            final_series_rows = []

            for s in range(series_count):
                # 预分配素材矩阵 [组m][位n]
                material_map = [["默认版本" for _ in range(N_ads)] for _ in range(M_groups)]
                
                if logic_mode == "组间测素材":
                    # 组 1 全是 A-1，组 2 全是 A-2
                    for m in range(M_groups):
                        v = materials[material_idx] if material_idx < len(materials) else "默认版本"
                        for n in range(N_ads):
                            material_map[m][n] = v
                        material_idx += 1
                else:
                    # 每个位按顺序消耗 A-1, A-2, A-3...
                    for n in range(N_ads):
                        for m in range(M_groups):
                            v = materials[material_idx] if material_idx < len(materials) else "默认版本"
                            material_map[m][n] = v
                            material_idx += 1

                # --- 3. 按照交叉顺序导出行 (横向优先) ---
                for n in range(N_ads):
                    for m in range(M_groups):
                        new_row = row.copy()
                        v_name = material_map[m][n]
                        new_row["广告素材版本名称"] = v_name
                        new_row["备注"] = "系统自动补齐" if v_name == "默认版本" else ""
                        final_series_rows.append(new_row)

            # --- 4. 统计补齐数用于校验 ---
            actual_padding_count = sum(1 for r in final_series_rows if r["广告素材版本名称"] == "默认版本")
            if template_padding != 0 and template_padding != actual_padding_count:
                warning_logs.append(f"❓ 行 {line_no}: 表格补充数({template_padding}) 与系统实际补齐行数({actual_padding_count}) 不符。")

            all_results.append(pd.DataFrame(final_series_rows))

        # --- 5. 结果展示与下载 ---
        if warning_logs:
            with st.expander("📝 逻辑校验报告 (发现配置冲突)"):
                for log in warning_logs:
                    st.warning(log)
        
        if all_results:
            final_df = pd.concat(all_results, ignore_index=True)
            display_cols = ["广告账号ID", "主页ID", "像素ID", "真实SKU", "虚拟SKU", "国家", "着陆页版本名称", "广告素材版本名称", "备注"]
            final_df = final_df[[c for c in display_cols if c in final_df.columns]]
            
            st.success(f"✅ 生成完毕！已应用{logic_mode}逻辑并完成交叉排列。")
            file_prefix = params.get("prefix", "项目_")
            xlsx_data = write_excel_final(final_df, "结构补齐结果", params)
            st.download_button(
                f"💾 下载：{file_prefix}结构补齐结果", 
                data=xlsx_data, 
                file_name=f"{file_prefix}结构补齐.xlsx"
            )