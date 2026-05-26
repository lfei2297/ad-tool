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

from utils import write_excel_final 

def run(params):
    st.subheader("🎯 模块四：补齐默认版本 (1:M:N 结构增强审计版)")
    
    st.info("""
    **💡 运行逻辑：**
    - **多组模式 (M > 1)**：组间测素材。每组下的 $N$ 条广告素材相同，消耗 1 个素材版本。
    - **单组模式 (M = 1)**：组内测素材。组内 $N$ 条广告素材各不相同，每个广告位消耗 1 个素材版本。
    - **自动补齐**：系统根据总坑位需求自动填充“默认版本”。
    """)

    # --- 1. 结构参数输入 (最高优先级) ---
    st.markdown("#### 📐 当前计划结构 (1:M:N)")
    col_a, col_b = st.columns(2)
    with col_a:
        M_groups = st.number_input("网页设定组数 (M)", min_value=1, value=1, help="每个系列下的广告组数")
    with col_b:
        N_ads = st.number_input("网页设定广告数 (N)", min_value=1, value=2, help="每个广告组下的广告位数量")

    up = st.file_uploader("📂 上传模块四专用模板", type=["xlsx"], key="m4_up")
    
    if up and st.button("🚀 开始校验并生成", key="m4_btn"):
        try:
            df_raw = pd.read_excel(up, dtype=str).fillna("0")
        except Exception as e:
            st.error(f"读取文件失败: {e}")
            return

        all_results = []
        warning_logs = [] # 用于存储审计发现的冲突

        for idx, row in df_raw.iterrows():
            line_no = idx + 2 # 对应 Excel 中的行号
            try:
                provided_count = int(float(row.get("提供素材版本数量", 0)))
                series_count = int(float(row.get("导品系列数", 1)))
                template_m = int(float(row.get("广告组数量", 0)))
                template_padding = int(float(row.get("补充默认版本数", 0)))
            except:
                st.error(f"❌ 第 {line_no} 行包含非法字符，请确保数字列均为纯数字。")
                continue
            
            # --- 2. 逻辑审计与对比判断 ---
            # (1) 校验广告组数 M
            if template_m != M_groups and template_m != 0:
                warning_logs.append(f"⚠️ 行 {line_no}: 模板组数({template_m}) 与网页设定({M_groups}) 不符，已按网页设定执行。")

            # (2) 计算理论物理总坑位 (行数)
            total_required_rows = M_groups * N_ads * series_count
            
            # (3) 计算现有素材能覆盖的行数
            # M>1 时每素材占 N 行；M=1 时每素材占 1 行
            rows_per_material = N_ads if M_groups > 1 else 1
            existing_rows_covered = provided_count * rows_per_material
            
            # (4) 系统计算出的最终需要补齐的物理行数
            system_padding_rows = max(0, total_required_rows - existing_rows_covered)
            
            # (5) 校验补充数：对比模板填写的数字与系统计算的物理行数
            if template_padding != system_padding_rows and template_padding != 0:
                warning_logs.append(f"❓ 行 {line_no}: 模板填写补充数({template_padding}) 与系统计算的物理补齐行数({system_padding_rows}) 不匹配。")

            # --- 3. 生成递增版本逻辑 ---
            base_name = str(row.get("广告素材版本名称", "素材"))
            clean_name = re.sub(r'-\d+$', '', base_name) # 去除可能存在的旧后缀
            
            # 生成所有可用的递增版本 [A-1, A-2, A-3...]
            all_potential_versions = [f"{clean_name}-{i}" for i in range(1, provided_count + 1)]
            
            expanded_rows = []
            
            # 确定本次循环需要消耗多少个素材版本（不超坑位）
            max_materials_needed = (M_groups * series_count) if M_groups > 1 else (N_ads * series_count)
            limited_versions = all_potential_versions[:max_materials_needed]
            
            for v_name in limited_versions:
                repeat_times = N_ads if M_groups > 1 else 1
                for _ in range(repeat_times):
                    new_row = row.copy()
                    new_row["广告素材版本名称"] = v_name
                    new_row["备注"] = f"{'组间去重' if M_groups > 1 else '组内测试'} | M={M_groups}, N={N_ads}"
                    expanded_rows.append(new_row)
            
            # 执行补齐逻辑
            for _ in range(system_padding_rows):
                d_row = row.copy()
                d_row["广告素材版本名称"] = "默认版本"
                d_row["备注"] = "系统自动补齐"
                expanded_rows.append(d_row)

            all_results.append(pd.DataFrame(expanded_rows))

        # --- 4. 报告与输出 ---
        if warning_logs:
            with st.expander("📝 逻辑校验报告 (发现配置冲突)"):
                for log in warning_logs:
                    st.warning(log)
        
        if all_results:
            final_df = pd.concat(all_results, ignore_index=True)
            
            # 仅保留核心展示列，剔除模板业务参数列
            display_cols = ["广告账号ID", "主页ID", "像素ID", "真实SKU", "虚拟SKU", "国家", "着陆页版本名称", "广告素材版本名称", "备注"]
            final_df = final_df[[c for c in display_cols if c in final_df.columns]]
            
            st.success(f"✅ 处理完成！总计生成 {len(final_df)} 行数据。")
            
            # 获取侧边栏前缀
            file_prefix = params.get("prefix", "项目A_")
            
            # 导出 Excel
            xlsx_data = write_excel_final(final_df, "结构补齐结果", params)
            st.download_button(
                label=f"💾 下载：{file_prefix}结构补齐结果",
                data=xlsx_data,
                file_name=f"{file_prefix}结构补齐.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )