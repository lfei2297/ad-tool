import streamlit as st
import pandas as pd
import gc
from collections import defaultdict
from utils import expand_material_versions, read_uploaded_excel, write_excel_final, create_zip_package

def run(params):
    st.subheader("🛠️ 模块一：基础独立拆分")
    
    col_setting1, col_setting2, col_mode = st.columns([1, 1, 1.5])
    
    with col_setting1:
        repeat_1 = st.number_input("第一次重复次数", min_value=1, value=1, help="单条素材基础展开次数，可以看作广告组数")
    with col_setting2:
        repeat_2 = st.number_input("第二次重复次数", min_value=1, value=1, help="整体复制次数，即原表中这一行品需要导入的系列数")
    with col_mode:
        st.write(""); st.write("")
        only_total = st.checkbox("⚡ 仅生成完整总表", value=True, help="取消勾选后，才会额外打包下载子表 ZIP 包")

    params['repeat_1'] = repeat_1
    params['repeat_2'] = repeat_2

    up = st.file_uploader("📂 上传原始素材表 (.xlsx)", type=["xlsx"], key="m1_up")
    
    # 🌟 切换/重置文件时清空缓存
    if "m1_last_up" not in st.session_state or st.session_state["m1_last_up"] != (up.name if up else None):
        st.session_state["m1_last_up"] = up.name if up else None
        st.session_state.pop("m1_excel_bytes", None)
        st.session_state.pop("m1_zip_bytes", None)

    if up:
        st.write("")
        if st.button("🚀 开始处理", key="m1_start_btn"):
            with st.spinner("🚀 数据处理与文件打包中，请稍候..."):
                df_raw = read_uploaded_excel(up.getvalue())
                
                raw_dicts = df_raw.to_dict('records')
                valid_rows = [
                    row for row in raw_dicts 
                    if not any('此行为说明' in str(v) or '可不填' in str(v) for v in row.values())
                ]
                
                all_expanded_dicts = []
                file_groups = defaultdict(list)

                # 遍历字典生成展开数据
                for row_dict in valid_rows:
                    vs = expand_material_versions(row_dict)
                    
                    row_rows = []
                    for v in vs:
                        nr = row_dict.copy()
                        nr["广告素材版本名称"] = v
                        for _ in range(repeat_1): 
                            row_rows.append(nr.copy())
                    
                    if repeat_2 > 1:
                        row_rows = row_rows * repeat_2
                    
                    all_expanded_dicts.extend(row_rows)
                    
                    # 🌟 1. 按素材数量分类保存
                    if not only_total:
                        file_groups[f"素材数_{len(vs)}"].extend(row_rows)

                if all_expanded_dicts:
                    df_total = pd.DataFrame(all_expanded_dicts)
                    
                    local_params = params.copy()
                    local_params['fast_mode'] = False  
                    local_params['enable_color'] = True
                    
                    # 生成总表数据
                    total_excel_data = write_excel_final(df_total, "Data", local_params)
                    st.session_state["m1_excel_bytes"] = total_excel_data
                    
                    # 🌟 2. 构建多文件打包任务 Tasks
                    if not only_total:
                        tasks = {"总表": df_total}
                        
                        # (A) 加入按素材数拆分的子表
                        for k, v in file_groups.items():
                            tasks[k] = pd.DataFrame(v)
                            
                        # (B) 加入按账号ID拆分的子表 (恢复多文件完整性)
                        if "广告账号ID" in df_total.columns:
                            acc_groups = df_total.groupby("广告账号ID")
                            for acc_id, acc_df in acc_groups:
                                acc_str = str(acc_id).strip()
                                if acc_str and acc_str.lower() != "nan":
                                    tasks[f"账号_{acc_str}"] = acc_df
                                    
                        # 生成包含所有多文件的 ZIP 压缩包
                        zip_data = create_zip_package(tasks, local_params, total_excel_data)
                        st.session_state["m1_zip_bytes"] = zip_data
                    else:
                        st.session_state.pop("m1_zip_bytes", None)

                    del df_total, all_expanded_dicts, file_groups
                    gc.collect()

        # 🌟 3. 持久化渲染下载区域（点击任意按钮不会刷新丢失）
        if "m1_excel_bytes" in st.session_state:
            st.success("✨ 处理成功！", icon="✅")

            if only_total or "m1_zip_bytes" not in st.session_state:
                col_btn, _ = st.columns([1, 1])
                with col_btn:
                    st.download_button(
                        label="📊 💾 下载：完整总表 (Excel)", 
                        data=st.session_state["m1_excel_bytes"], 
                        file_name=f"{params.get('prefix', '项目_')}总表.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="dl_m1_excel"
                    )
            else:
                dl_col1, dl_col2 = st.columns(2)
                with dl_col1:
                    st.download_button(
                        label="📊 💾 下载：完整总表 (Excel)", 
                        data=st.session_state["m1_excel_bytes"], 
                        file_name=f"{params.get('prefix', '项目_')}总表.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="dl_m1_excel_multi"
                    )
                with dl_col2:
                    st.download_button(
                        label="📦 💾 下载全套：分类包 (ZIP)", 
                        data=st.session_state["m1_zip_bytes"], 
                        file_name=f"{params.get('prefix', '项目_')}结果.zip",
                        mime="application/zip",
                        use_container_width=True,
                        key="dl_m1_zip"
                    )