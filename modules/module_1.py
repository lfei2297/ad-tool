import streamlit as st
import pandas as pd
import gc
from collections import defaultdict
from utils import expand_material_versions, read_uploaded_excel, write_excel_final, create_zip_package

def run(params):
    st.subheader("🛠️ 模块一：基础独立拆分")
    st.caption("第一次重复 = 广告组数，第二次重复 = 系列数。默认只出总表；取消勾选会再打一份按素材数 / 账号拆开的 ZIP。")

    c_up, c_r1, c_r2, c_opt = st.columns([1.55, 0.7, 0.7, 1.15], gap="medium", vertical_alignment="bottom")
    with c_up:
        up = st.file_uploader("上传原始素材表 (.xlsx)", type=["xlsx"], key="m1_up")
    with c_r1:
        repeat_1 = st.number_input(
            "第一次重复次数",
            min_value=1,
            value=1,
            step=1,
            help="单条素材基础展开次数，可以看作广告组数",
        )
    with c_r2:
        repeat_2 = st.number_input(
            "第二次重复次数",
            min_value=1,
            value=1,
            step=1,
            help="整体复制次数，即原表中这一行品需要导入的系列数",
        )
    with c_opt:
        only_total = st.checkbox(
            "仅生成完整总表",
            value=True,
            help="取消勾选后，才会额外打包下载子表 ZIP 包",
        )

    params["repeat_1"] = repeat_1
    params["repeat_2"] = repeat_2

    m1_sig = (
        up.name if up else None,
        int(repeat_1),
        int(repeat_2),
        bool(only_total),
        params.get("prefix"),
    )
    if st.session_state.get("m1_sig") != m1_sig:
        st.session_state["m1_sig"] = m1_sig
        st.session_state.pop("m1_excel_bytes", None)
        st.session_state.pop("m1_zip_bytes", None)
        st.session_state.pop("m1_row_count", None)

    if only_total:
        c_go, c_dl = st.columns(2, gap="medium")
        c_zip = None
    else:
        c_go, c_dl, c_zip = st.columns(3, gap="medium")

    with c_go:
        start_btn = st.button(
            "🚀 开始处理",
            key="m1_start_btn",
            use_container_width=True,
            disabled=not bool(up),
        )

    empty_result = False
    if start_btn and up:
        with st.spinner("数据处理与文件打包中，请稍候..."):
            df_raw = read_uploaded_excel(up.getvalue())

            raw_dicts = df_raw.to_dict("records")
            valid_rows = [
                row for row in raw_dicts
                if not any("此行为说明" in str(v) or "可不填" in str(v) for v in row.values())
            ]

            all_expanded_dicts = []
            file_groups = defaultdict(list)

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

                if not only_total:
                    file_groups[f"素材数_{len(vs)}"].extend(row_rows)

            if all_expanded_dicts:
                df_total = pd.DataFrame(all_expanded_dicts)

                local_params = params.copy()
                local_params["fast_mode"] = False
                local_params["enable_color"] = True

                total_excel_data = write_excel_final(df_total, "Data", local_params)
                st.session_state["m1_excel_bytes"] = total_excel_data
                st.session_state["m1_row_count"] = len(df_total)

                if not only_total:
                    tasks = {"总表": df_total}
                    for k, v in file_groups.items():
                        tasks[k] = pd.DataFrame(v)
                    if "广告账号ID" in df_total.columns:
                        acc_groups = df_total.groupby("广告账号ID")
                        for acc_id, acc_df in acc_groups:
                            acc_str = str(acc_id).strip()
                            if acc_str and acc_str.lower() != "nan":
                                tasks[f"账号_{acc_str}"] = acc_df
                    zip_data = create_zip_package(tasks, local_params, total_excel_data)
                    st.session_state["m1_zip_bytes"] = zip_data
                else:
                    st.session_state.pop("m1_zip_bytes", None)

                del df_total, all_expanded_dicts, file_groups
                gc.collect()
            else:
                empty_result = True
                st.session_state.pop("m1_excel_bytes", None)
                st.session_state.pop("m1_zip_bytes", None)
                st.session_state.pop("m1_row_count", None)

    prefix = params.get("prefix", "项目_")
    has_excel = "m1_excel_bytes" in st.session_state
    row_count = st.session_state.get("m1_row_count", 0)
    excel_label = f"💾 下载完整总表（{row_count} 行）" if has_excel else "💾 下载完整总表"
    with c_dl:
        st.download_button(
            label=excel_label,
            data=st.session_state.get("m1_excel_bytes") or b"",
            file_name=f"{prefix}总表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            disabled=not has_excel,
            key="dl_m1_excel",
        )

    has_zip = "m1_zip_bytes" in st.session_state
    if c_zip is not None:
        with c_zip:
            st.download_button(
                label="📦 下载分类包 (ZIP)",
                data=st.session_state.get("m1_zip_bytes") or b"",
                file_name=f"{prefix}结果.zip",
                mime="application/zip",
                use_container_width=True,
                disabled=not has_zip,
                key="dl_m1_zip",
            )

    if empty_result:
        st.warning("未检测到有效业务数据，请确认表格从第 3 行开始填写。")
    elif has_excel:
        st.success(f"处理成功，已生成 {row_count} 行。")