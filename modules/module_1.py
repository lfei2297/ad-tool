import streamlit as st
import pandas as pd
import zipfile
import io
import re
from collections import defaultdict
from utils import expand_material_versions, write_excel_final

def run(params):
    st.subheader("🛠️ 模块一：基础独立拆分")
    
    # ✨ 挪到这里的重复设置
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

        tasks["总表"] = pd.concat(all_exp, ignore_index=True)
        for k, v in file_groups.items():
            tasks[k] = pd.concat(v, ignore_index=True)

        zip_b = io.BytesIO()
        with zipfile.ZipFile(zip_b, "w") as zf:
            for n, d in tasks.items():
                zf.writestr(f"{params['prefix']}{n}.xlsx", write_excel_final(d, "Data", params))
        
        st.download_button("📦 下载处理结果包", data=zip_b.getvalue(), file_name=f"{params['prefix']}结果.zip")