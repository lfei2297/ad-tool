import streamlit as st
import pandas as pd
import gc
from utils import write_excel_final, read_uploaded_excel, safe_int

def run(params):
    st.subheader("🎯 模块四：补齐默认版本 (全结构适配版)")
    
    with st.expander("💡 点击查看：模块四运行逻辑说明"):
        st.markdown("""
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
        logic_mode = st.radio("🎨 素材分配逻辑", ["组间测素材", "组内测素材"])

    up = st.file_uploader("📂 上传模块四专用模板", type=["xlsx"], key="m4_up")
    
    if up and st.button("🚀 开始校验并生成", key="m4_btn"):
        df_raw = read_uploaded_excel(up.getvalue())
        
        valid_rows = [
            r for r in df_raw.to_dict('records') 
            if not any('此行为说明' in str(v) or '可不填' in str(v) for v in r.values())
        ]
        
        all_results = []
        warning_logs = []

        for idx, row_dict in enumerate(valid_rows):
            excel_line_no = idx + 3
            
            provided_count = safe_int(row_dict.get("提供素材版本数量"))
            series_count = safe_int(row_dict.get("导品系列数"), default=1)
            template_padding = safe_int(row_dict.get("补充默认版本数"))
            
            base_name = str(row_dict.get("广告素材版本名称", "素材")).strip()
            clean_name = base_name.rsplit('-', 1)[0] if '-' in base_name and base_name.rsplit('-', 1)[1].isdigit() else base_name
            materials = [f"{clean_name}-{i}" for i in range(1, provided_count + 1)]
            
            final_series_rows = []

            for s in range(series_count):
                matrix = [["默认版本" for _ in range(N_ads)] for _ in range(M_groups)]
                
                # 🌟 使用计算映射替代 mat_idx 状态变量
                if logic_mode == "组间测素材":
                    for m in range(M_groups):
                        v = materials[m] if m < len(materials) else "默认版本"
                        for n in range(N_ads):
                            matrix[m][n] = v
                else:
                    for n in range(N_ads):
                        for m in range(M_groups):
                            mat_i = n * M_groups + m
                            v = materials[mat_i] if mat_i < len(materials) else "默认版本"
                            matrix[m][n] = v

                for n in range(N_ads):
                    for m in range(M_groups):
                        new_row = row_dict.copy() 
                        v_name = matrix[m][n]
                        new_row["广告素材版本名称"] = v_name
                        new_row["备注"] = "系统自动补齐" if v_name == "默认版本" else ""
                        final_series_rows.append(new_row)

            actual_padding = sum(1 for r in final_series_rows if r["广告素材版本名称"] == "默认版本")
            if template_padding != 0 and template_padding != actual_padding:
                warning_logs.append(f"❓ Excel第 {excel_line_no} 行: 模板补充数({template_padding}) vs 实际({actual_padding})")

            all_results.append(pd.DataFrame(final_series_rows))

        if warning_logs:
            with st.expander("📝 逻辑校验报告"):
                for log in warning_logs: st.warning(log)
        
        if all_results:
            final_df = pd.concat(all_results, ignore_index=True)
            cols_to_hide = ["提供素材版本数量", "广告组数量", "导品系列数", "补充默认版本数"]
            final_df = final_df.drop(columns=[c for c in cols_to_hide if c in final_df.columns], errors='ignore')
            
            st.success("✅ 生成完毕！")
            file_prefix = params.get("prefix", "项目_")
            xlsx_data = write_excel_final(final_df, "结构补齐结果", params)
            st.download_button(f"💾 下载：{file_prefix}结果", data=xlsx_data, file_name=f"{file_prefix}补齐.xlsx")
            
            del df_raw, valid_rows, all_results, final_df
            gc.collect()