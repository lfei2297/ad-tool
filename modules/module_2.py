import streamlit as st
import pandas as pd
import gc
from collections import defaultdict
from utils import expand_material_versions, read_uploaded_excel, create_zip_package

def run(params):
    params['repeat_1'] = 1
    params['repeat_2'] = 1
    st.subheader("🌍 模块二：同SKU+国家聚合拆分")
    up = st.file_uploader("📂 上传原始素材表 (.xlsx)", type=["xlsx"], key="m2_up")
    
    if up and st.button("🚀 开始处理", key="m2_btn"):
        df_raw = read_uploaded_excel(up.getvalue())
        
        raw_dicts = df_raw.to_dict('records')
        valid_rows = [
            row for row in raw_dicts 
            if not any('此行为说明' in str(v) or '可不填' in str(v) for v in row.values())
        ]

        if not valid_rows:
            st.warning("⚠️ 上传的表格中未检测到有效的业务数据，请检查后再试。")
            return

        tasks = {}
        all_expanded = []
        m2_agg = defaultdict(list)
        m2_counts = defaultdict(int)

        for row_dict in valid_rows:
            vs = expand_material_versions(row_dict)
            m_len = len(vs)
            
            row_rows = []
            for v in vs:
                nr = row_dict.copy()
                nr["广告素材版本名称"] = v
                for _ in range(params['repeat_1']): row_rows.append(nr.copy())
            
            tmp = pd.DataFrame(row_rows)
            if params.get('repeat_2', 1) > 1:
                tmp = pd.concat([tmp] * params['repeat_2'], ignore_index=True)
            
            all_expanded.append(tmp)
            
            sku = str(row_dict.get("真实SKU","")).strip()
            if sku == "" or sku.lower() == "nan":
                sku = str(row_dict.get("虚拟SKU","")).strip()
            
            country = str(row_dict.get("国家","")).strip()
            
            m2_agg[(sku, country)].append(tmp)
            m2_counts[(sku, country)] += m_len

        tasks["总表"] = pd.concat(all_expanded, ignore_index=True)
        
        for (sku, country), dfs in m2_agg.items():
            total_m = m2_counts[(sku, country)]
            c_tag = f"{country}_" if country else ""
            filename = f"{c_tag}素材数_{total_m}"
            tasks[filename] = pd.concat(dfs, ignore_index=True)

        zip_bytes = create_zip_package(tasks, params)
        
        st.success("🎉 聚合处理完成！")
        st.download_button("📦 下载聚合结果包 (ZIP)", data=zip_bytes, file_name=f"{params['prefix']}聚合结果.zip")
        
        # 🌟 安全的列表式清理写法，绝无 NameError 风险
        possible_vars = ['df_raw', 'valid_rows', 'all_expanded', 'tasks', 'm2_agg', 'tmp']
        for var in possible_vars:
            if var in locals(): del locals()[var]
        gc.collect()