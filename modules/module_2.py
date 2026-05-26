import streamlit as st
import pandas as pd
import re
from collections import defaultdict
import io
import zipfile
from utils import expand_material_versions, write_excel_final

def run(params):
    st.subheader("🌍 模块二：同SKU+国家聚合拆分")
    up = st.file_uploader("📂 上传原始素材表 (.xlsx)", type=["xlsx"], key="m2_up")
    
    if up and st.button("🚀 开始处理", key="m2_btn"):
        df_raw = pd.read_excel(up, dtype=str).fillna("")
        tasks = {}
        all_expanded = []
        
        # 聚合容器
        m2_agg = defaultdict(list)
        m2_counts = defaultdict(int)

        for _, row in df_raw.iterrows():
            # 1. 解析
            lp_match = re.search(r'-(\d+)$', str(row.get("着陆页版本名称", "")))
            lp_v = lp_match.group(1) if lp_match else ""
            vs = expand_material_versions(row, lp_v)
            m_len = len(vs)
            
            # 2. 展开
            row_rows = []
            for v in vs:
                nr = row.copy()
                nr["广告素材版本名称"] = v
                for _ in range(params['repeat_1']): row_rows.append(nr.copy())
            
            tmp = pd.DataFrame(row_rows)
            if params.get('repeat_2', 1) > 1:
                tmp = pd.concat([tmp] * params['repeat_2'], ignore_index=True)
            
            all_expanded.append(tmp)
            
            # 3. 模块二聚合标识：优先取真实SKU，其次取虚拟SKU
            sku = str(row.get("真实SKU","")).strip()
            if sku == "" or sku.lower() == "nan":
                sku = str(row.get("虚拟SKU","")).strip()
            
            country = str(row.get("国家","")).strip()
            
            # 聚合：同一 (SKU, 国家) 的数据存入列表
            m2_agg[(sku, country)].append(tmp)
            m2_counts[(sku, country)] += m_len

        # 总表 (不强制排序)
        tasks["总表"] = pd.concat(all_expanded, ignore_index=True)
        
        # 聚合拆分表
        for (sku, country), dfs in m2_agg.items():
            total_m = m2_counts[(sku, country)]
            c_tag = f"{country}_" if country else ""
            # 文件名体现：国家_素材数_总数
            filename = f"{c_tag}素材数_{total_m}"
            tasks[filename] = pd.concat(dfs, ignore_index=True)

        # 打包 ZIP
        zip_b = io.BytesIO()
        with zipfile.ZipFile(zip_b, "w") as zf:
            for n, d in tasks.items():
                zf.writestr(f"{params['prefix']}{n}.xlsx", write_excel_final(d, "Data", params))
        
        st.success("🎉 聚合处理完成！")
        st.download_button("📦 下载聚合结果包 (ZIP)", data=zip_b.getvalue(), file_name=f"{params['prefix']}聚合结果.zip")