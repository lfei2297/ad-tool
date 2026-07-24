import streamlit as st
import pandas as pd
import gc
from utils import write_excel_final, read_uploaded_excel, cycle_repeat, safe_int

def run(params):
    # 🌟 局部 CSS 注入：精准拉紧控面板各控件之间的行间距
    st.markdown("""
        <style>
            /* 压缩输入框、单选框等组件的上下外边距 */
            .stRadio, .stNumberInput, .stFileUploader, .stButton {
                margin-bottom: -0.5rem !important;
            }
            /* 压缩分割线上下间距 */
            hr {
                margin-top: 0.8rem !important;
                margin-bottom: 0.8rem !important;
            }
        </style>
    """, unsafe_allow_html=True)

    st.subheader("🚀 模块五：通用素材循环填充与分流")
    
    st.markdown("#### ⚙️ 核心参数配置")
    run_mode = st.radio("🎨 导出与拆分模式", ["模式A：生成独立总表", "模式B：生成总表并动态平摊拆分"])

    default_rows = 15 if "模式A" in run_mode else 20
    target_total_rows = st.number_input("📊 每个品生成总行数", min_value=1, value=default_rows)

    sub_table_count = 4
    if "模式B" in run_mode:
        st.markdown("---")
        st.markdown("#### ✂️ 品内拆分段设置")
        sub_table_count = st.number_input("📐 拆分表格数量", min_value=1, max_value=int(target_total_rows), value=4)
        
        base_step = int(target_total_rows // sub_table_count)
        rem = int(target_total_rows % sub_table_count)
        if rem == 0:
            st.caption(f"💡 实时换算：每个品生成的 {target_total_rows} 行将完美平分，每张子表分得 **{base_step}** 行。")
        else:
            st.caption(f"💡 实时换算：前 **{rem}** 张子表每品分得 **{base_step + 1}** 行，后 **{int(sub_table_count) - rem}** 张子表每品分得 **{base_step}** 行。")

    # 去掉原有的硬换行/分割线，改用微小空隙
    st.write("") 
    up = st.file_uploader("📂 上传标准用模板", type=["xlsx"], key="m5_up")
    
    if up and st.button("🚀 开始自动化构建", key="m5_btn"):
        df_raw = read_uploaded_excel(up.getvalue())
        valid_rows = [r for r in df_raw.to_dict('records') if not any('此行为说明' in str(v) or '可不填' in str(v) for v in r.values())]
        
        file_prefix = params.get("prefix", "项目_")
        total_expanded_rows = []
        buckets = {i: [] for i in range(1, int(sub_table_count) + 1)}
        
        base_step = int(target_total_rows // sub_table_count)
        rem = int(target_total_rows % sub_table_count)
        sku_split_lens = [base_step + (1 if i < rem else 0) for i in range(int(sub_table_count))]

        for row_dict in valid_rows:
            provided_count = safe_int(row_dict.get("广告素材数量", 0))
            base_name = str(row_dict.get("广告素材版本名称", "素材")).strip()
            
            clean_name = base_name.rsplit('-', 1)[0] if '-' in base_name and base_name.rsplit('-', 1)[1].isdigit() else base_name
            
            if provided_count <= 1:
                material_pool = [f"{clean_name}-1"]
            else:
                material_pool = [f"{clean_name}-{i}" for i in range(1, provided_count + 1)]
            
            special_note = f"素材数超标，仅保留前{target_total_rows}个版本" if provided_count > target_total_rows else ""
            this_sku_materials = cycle_repeat(material_pool, int(target_total_rows))
            
            if "模式B" in run_mode:
                materials_iter = iter(this_sku_materials)
                for b_idx, length in enumerate(sku_split_lens):
                    bucket_id = b_idx + 1
                    for _ in range(length):
                        v_name = next(materials_iter)
                        new_row = row_dict.copy() 
                        new_row["广告素材版本名称"] = v_name
                        new_row["备注"] = special_note
                        buckets[bucket_id].append(new_row)

            for v_name in this_sku_materials:
                new_row = row_dict.copy() 
                new_row["广告素材版本名称"] = v_name
                new_row["备注"] = special_note
                total_expanded_rows.append(new_row)

        if total_expanded_rows:
            master_df = pd.DataFrame(total_expanded_rows)

            if "模式A" in run_mode:
                st.success(f"✅ 模式A构建成功！已按照每个品【{target_total_rows}行】平铺填充。")
                xlsx_data = write_excel_final(master_df, "填充总表结果", params, color_by="广告账号ID")
                st.download_button(
                    f"💾 下载：{target_total_rows}行大总表", 
                    data=xlsx_data, 
                    file_name=f"{file_prefix}{target_total_rows}行大总表.xlsx"
                )
            else:
                import io
                import zipfile
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_STORED) as zip_file:
                    xlsx_master = write_excel_final(master_df, "大总表结果", params, color_by="广告账号ID")
                    zip_file.writestr(f"{file_prefix}{target_total_rows}行完整总表.xlsx", xlsx_master)
                    
                    accumulated_rows = 0
                    for b_id in sorted(buckets.keys()):
                        sliced_rows = buckets[b_id]
                        sliced_df = pd.DataFrame(sliced_rows)
                        
                        current_len = sku_split_lens[b_id - 1]
                        start_row_num = accumulated_rows + 1
                        end_row_num = accumulated_rows + current_len
                        accumulated_rows += current_len
                        
                        range_text = f"{start_row_num}-{end_row_num}行"
                        sliced_df["备注"] = f"各品内部第 {range_text}"
                        
                        xlsx_slice = write_excel_final(sliced_df, f"新表{b_id}", params, color_by="广告账号ID")
                        zip_file.writestr(f"{file_prefix}新表{b_id}_各品{range_text}.xlsx", xlsx_slice)
                
                zip_buffer.seek(0)
                st.success(f"📦 模式B品内平摊成功！总计生成 1 张大总表与 {len(buckets)} 张拆分子表。")
                st.download_button(
                    label="🎁 💾 一键下载全套平摊打包件 (ZIP)",
                    data=zip_buffer,
                    file_name=f"{file_prefix}品内平摊_{target_total_rows}行分{sub_table_count}表.zip",
                    mime="application/zip",
                    key="m5_custom_zip_btn"
                )
                
        del df_raw, valid_rows, total_expanded_rows, master_df, buckets
        gc.collect()