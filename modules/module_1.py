import streamlit as st
import pandas as pd
import re
from collections import defaultdict
import io
import zipfile
import sys
import os

# ✅ 工业级写法：强制将项目根目录加入 Python 搜索路径
# 获取当前文件的父目录的父目录（即根目录）
root_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if root_path not in sys.path:
    sys.path.append(root_path)

# 现在可以安全地从根目录导入 utils 了
from utils import expand_material_versions, write_excel_final

def run(params):
    st.subheader("🛠️ 模块一：基础独立拆分")
    up = st.file_uploader("📂 上传原始素材表 (.xlsx)", type=["xlsx"], key="m1_up")
    
    if up and st.button("🚀 开始处理", key="m1_btn"):
        df_raw = pd.read_excel(up, dtype=str).fillna("")
        tasks = {}
        all_expanded = []
        file_groups = defaultdict(list)
        
        for _, row in df_raw.iterrows():
            # 获取着陆页版本号用于拼接
            lp_match = re.search(r'-(\d+)$', str(row.get("着陆页版本名称", "")))
            lp_v = lp_match.group(1) if lp_match else ""
            
            # 展开素材版本
            vs = expand_material_versions(row, lp_v)
            m_len = len(vs)
            
            row_rows = []
            for v in vs:
                nr = row.copy()
                nr["广告素材版本名称"] = v
                for _ in range(params['repeat_1']): 
                    row_rows.append(nr.copy())
            
            tmp = pd.DataFrame(row_rows)
            # 处理第二次重复
            if params.get('repeat_2', 1) > 1: 
                tmp = pd.concat([tmp] * params['repeat_2'], ignore_index=True)
            
            all_expanded.append(tmp)
            # 模块一特有：按单行素材数量归类
            file_groups[f"素材数_{m_len}"].append(tmp)

        # 生成总表
        tasks["总表"] = pd.concat(all_expanded, ignore_index=True)
        # 生成拆分表
        for k, v in file_groups.items(): 
            tasks[k] = pd.concat(v, ignore_index=True)

        # 打包 ZIP
        zip_b = io.BytesIO()
        with zipfile.ZipFile(zip_b, "w") as zf:
            for n, d in tasks.items():
                zf.writestr(f"{params['prefix']}{n}.xlsx", write_excel_final(d, "Data", params))
        
        st.success("🎉 处理完成！")
        st.download_button("📦 下载处理结果包 (ZIP)", data=zip_b.getvalue(), file_name=f"{params['prefix']}结果.zip")