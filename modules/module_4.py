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
    st.subheader("🎯 模块四：补齐默认版本 (1:M:N 坑位驱动版)")
    
    # --- 1. 结构参数输入 ---
    st.markdown("#### 📐 当前计划结构 (1:M:N)")
    col_a, col_b = st.columns(2)
    with col_a:
        M_groups = st.number_input("网页设定组数 (M)", min_value=1, value=1)
    with col_b:
        N_ads = st.number_input("网页设定广告数 (N)", min_value=1, value=2)

    up = st.file_uploader("📂 上传模块四专用模板", type=["xlsx"], key="m4_up")
    
    if up and st.button("🚀 开始生成", key="m4_btn"):
        df_raw = pd.read_excel(up, dtype=str).fillna("0")
        all_results = []
        warning_logs = []

        for idx, row in df_raw.iterrows():
            line_no = idx + 2
            try:
                provided_count = int(float(row.get("提供素材版本数量", 0)))
                series_count = int(float(row.get("导品系列数", 1)))
                template_padding = int(float(row.get("补充默认版本数", 0)))
            except:
                continue
            
            # --- 2. 核心逻辑：计算坑位与素材分配 ---
            total_slots = M_groups * N_ads * series_count # 总物理行数
            
            # 生成素材队列 (如 [A-1, A-2, A-3])
            base_name = str(row.get("广告素材版本名称", "素材"))
            clean_name = re.sub(r'-\d+$', '', base_name)
            materials = [f"{clean_name}-{i}" for i in range(1, provided_count + 1)]
            
            final_rows_data = []
            material_idx = 0
            
            # --- ✨ 坑位驱动填充算法 ---
            # 外部循环：系列 -> 广告组 -> 广告位
            for s in range(series_count):
                for m in range(M_groups):
                    # 【关键点】在 M > 1 模式下，每一组消耗 1 个素材
                    # 在 M = 1 模式下，每一个广告位消耗 1 个素材
                    current_material = "默认版本"
                    is_padding = True
                    
                    # 尝试获取可用素材
                    if material_idx < len(materials):
                        current_material = materials[material_idx]
                        is_padding = False
                    
                    for n in range(N_ads):
                        new_row = row.copy()
                        new_row["广告素材版本名称"] = current_material
                        new_row["备注"] = "系统自动补齐" if is_padding else f"系列{s+1}-组{m+1}"
                        final_rows_data.append(new_row)
                        
                        # 如果是 M=1 模式，每填一个广告位，素材索引就往后推一个
                        if M_groups == 1:
                            material_idx += 1
                            # 重新检查下一个位置是否有素材
                            if material_idx < len(materials):
                                current_material = materials[material_idx]
                                is_padding = False
                            else:
                                current_material = "默认版本"
                                is_padding = True
                    
                    # 如果是 M > 1 模式，填完一组后，素材索引才往后推一个
                    if M_groups > 1:
                        material_idx += 1

            # --- 3. 审计校验 ---
            actual_padding_count = sum(1 for r in final_rows_data if r["广告素材版本名称"] == "默认版本")
            if template_padding != 0 and template_padding != actual_padding_count:
                warning_logs.append(f"❓ 行 {line_no}: 模板填写的补充数({template_padding}) 与系统实际补齐行数({actual_padding_count}) 不符。")

            all_results.append(pd.DataFrame(final_rows_data))

        # --- 4. 输出 ---
        if warning_logs:
            with st.expander("📝 逻辑校验报告"):
                for log in warning_logs: st.warning(log)
        
        if all_results:
            final_df = pd.concat(all_results, ignore_index=True)
            display_cols = ["广告账号ID", "主页ID", "像素ID", "真实SKU", "虚拟SKU", "国家", "着陆页版本名称", "广告素材版本名称", "备注"]
            final_df = final_df[[c for c in display_cols if c in final_df.columns]]
            
            st.success(f"✅ 生成完毕！总行数：{len(final_df)}")
            file_prefix = params.get("prefix", "项目_")
            xlsx_data = write_excel_final(final_df, "结构补齐结果", params)
            st.download_button(f"💾 下载：{file_prefix}结果", data=xlsx_data, file_name=f"{file_prefix}结构补齐.xlsx")