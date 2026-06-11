import streamlit as st
import pandas as pd
import zipfile
import io
import re
from collections import defaultdict
from utils import expand_material_versions, write_excel_final

def run(params):
    st.subheader("🛠️ 模块一：基础独立拆分")
    
    # --- 🔄 素材重复设置 ---
    st.markdown("#### 🔄 素材重复设置")
    col1, col2 = st.columns(2)
    with col1:
        repeat_1 = st.number_input("第一次重复次数", min_value=1, value=1, help="单条素材基础展开次数")
    with col2:
        repeat_2 = st.number_input("第二次重复次数", min_value=1, value=1, help="总表整体复制次数")
    
    # 更新局部参数
    params['repeat_1'] = repeat_1
    params['repeat_2'] = repeat_2

    up = st.file_uploader("📂 上传原始素材表 (.xlsx)", type=["xlsx"], key="m1_up")
    
    if up and st.button("🚀 开始处理"):
        df_raw = pd.read_excel(up, dtype=str).fillna("")
        tasks = {}
        all_exp = []
        file_groups = defaultdict(list)

        for _, row in df_raw.iterrows():
            lp_match = re.search(r'-(\d+)$', str(row.get("着陆页版本名称", "")))
            lp_v = lp_match.group(1) if lp_match else ""
            vs = expand_material_versions(row, lp_v)
            
            row_rows = []
            for v in vs:
                nr = row.copy()
                nr["广告素材版本名称"] = v
                for _ in range(repeat_1): row_rows.append(nr.copy())
            
            tmp = pd.DataFrame(row_rows)
            if repeat_2 > 1:
                tmp = pd.concat([tmp] * repeat_2, ignore_index=True)
            
            all_exp.append(tmp)
            file_groups[f"素材数_{len(vs)}"].append(tmp)

        # 汇聚总表与子表数据
        df_total = pd.concat(all_exp, ignore_index=True)
        tasks["总表"] = df_total
        for k, v in file_groups.items():
            tasks[k] = pd.concat(v, ignore_index=True)

        # --- ✨ 核心新增：单独为“总表”生成一份独立 Excel 数据 ---
        total_excel_data = write_excel_final(df_total, "Data", params)

        # 构建全套压缩包数据
        zip_b = io.BytesIO()
        with zipfile.ZipFile(zip_b, "w") as zf:
            for n, d in tasks.items():
                # 复用已有逻辑压入压缩包
                if n == "总表":
                    zf.writestr(f"{params['prefix']}{n}.xlsx", total_excel_data)
                else:
                    zf.writestr(f"{params['prefix']}{n}.xlsx", write_excel_final(d, "Data", params))
        
        st.success("✨ 数据处理成功！请选择下方合适的方式下载：")
        st.markdown("---")

        # --- 📐 按钮横向排版：独立总表与压缩包并存 ---
        dl_col1, dl_col2 = st.columns(2)
        
        with dl_col1:
            # 新增按钮：单独直出下载大总表
            st.download_button(
                label="📊 💾 单独下载：完整大总表 (Excel)", 
                data=total_excel_data, 
                file_name=f"{params['prefix']}总表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="m1_single_total_btn"
            )
            
        with dl_col2:
            # 原有按钮：下载全套压缩包
            st.download_button(
                label="📦 💾 下载全套：结果文件包 (ZIP)", 
                data=zip_b.getvalue(), 
                file_name=f"{params['prefix']}结果.zip",
                mime="application/zip",
                key="m1_zip_package_btn"
            )