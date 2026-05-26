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
    st.subheader("🎯 模块四：补齐默认版本 (增强审计版)")
    
    st.info("""
    **💡 逻辑说明：**
    - **多组模式 (M > 1)**：组间测素材。每组下的 $N$ 条广告素材相同，消耗 1 个素材版本。
    - **单组模式 (M = 1)**：组内测素材。组内 $N$ 条广告素材各不相同，每个广告位消耗 1 个素材版本。
    """)
    
    # --- 1. 结构参数输入 (最高优先级) ---
    st.markdown("#### 📐 当前计划结构 (1:M:N)")
    col_a, col_b = st.columns(2)
    with col_a:
        M_groups = st.number_input("网页设定组数 (M)", min_value=1, value=1)
    with col_b:
        N_ads = st.number_input("网页设定广告数 (N)", min_value=1, value=2)

    up = st.file_uploader("📂 上传模块四专用模板", type=["xlsx"], key="m4_up")
    
    if up and st.button("🚀 开始校验并生成", key="m4_btn"):
        df_raw = pd.read_excel(up, dtype=str).fillna("0")
        all_results = []
        warning_logs = [] # 记录冲突

        for idx, row in df_raw.iterrows():
            line_no = idx + 2 # Excel 行号
            try:
                provided_count = int(row.get("提供素材版本数量", 0))
                series_count = int(row.get("导品系列数", 1))
                template_m = int(row.get("广告组数量", 0))
                template_padding = int(row.get("补充默认版本数", 0))
            except:
                st.error(f"❌ 第 {line_no} 行包含非数字字符，请检查数字列。")
                continue
            
            # --- 🔍 逻辑审计与对比判断 ---
            # 1. 检查 M 是否一致
            if template_m != M_groups and template_m != 0:
                warning_logs.append(f"⚠️ 行 {line_no}: 模板组数({template_m}) 与网页设定({M_groups}) 不符，已按网页设定执行。")

            # 2. 计算系统缺口
            total_required = M_groups * N_ads * series_count
            # 计算逻辑：M=1 消耗 N 个素材；M>1 消耗 M 个素材（每个展 N 次）
            actual_material_needed = (series_count * M_groups) if M_groups > 1 else (series_count * N_ads)
            
            system_gap = max(0, actual_material_needed - provided_count)
            # 换算成物理行数
            final_padding_rows = (system_gap * N_ads) if M_groups > 1 else system_gap
            
            # 3. 校验用户填写的“补充默认版本数”是否正确
            if template_padding != system_gap and template_padding != 0:
                warning_logs.append(f"❓ 行 {line_no}: 模板填写补充数({template_padding}) 与系统计算({system_gap}) 不匹配。")

            # --- 3. 生成逻辑 ---
            base_name = str(row.get("广告素材版本名称", ""))
            clean_name = re.sub(r'-\d+$', '', base_name)
            existing_versions = [f"{clean_name}-{i}" for i in range(1, provided_count + 1)]
            
            expanded_rows = []
            # 截取所需素材（不超过坑位）
            limited_versions = existing_versions[:actual_material_needed]
            
            for v_name in limited_versions:
                repeat_times = N_ads if M_groups > 1 else 1
                for _ in range(repeat_times):
                    new_row = row.copy()
                    new_row["广告素材版本名称"] = v_name
                    expanded_rows.append(new_row)
            
            # 补齐缺口
            for _ in range(final_padding_rows):
                d_row = row.copy()
                d_row["广告素材版本名称"] = "默认版本"
                d_row["备注"] = "系统自动补齐"
                expanded_rows.append(d_row)

            all_results.append(pd.DataFrame(expanded_rows))

        # --- 4. 结果反馈 ---
        if warning_logs:
            with st.expander("📝 逻辑校验报告 (发现冲突)"):
                for log in warning_logs: st.warning(log)
        
        if all_results:
            final_df = pd.concat(all_results, ignore_index=True)
            st.success(f"✅ 生成完毕！总行数：{len(final_df)}")
            
            file_prefix = params.get("prefix", "结果_")
            xlsx_data = write_excel_final(final_df, "结构补齐结果", params)
            st.download_button("💾 下载校验后的结果", data=xlsx_data, file_name=f"{file_prefix}补齐.xlsx")