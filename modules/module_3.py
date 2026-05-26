import streamlit as st
import pandas as pd
import re
import sys
import os

# ✅ 强制定位补丁：将项目的根目录加入搜索路径
# os.path.dirname(__file__) 获取当前 modules 文件夹路径
# '..' 代表上一级目录（即根目录）
current_dir = os.path.dirname(os.path.abspath(__file__))
root_path = os.path.abspath(os.path.join(current_dir, '..'))

if root_path not in sys.path:
    sys.path.insert(0, root_path)

# 现在可以安全地从根目录导入 utils 了
from utils import expand_material_versions, write_excel_final

def run(params):
    st.subheader("🛠️ 模块三：智能分组处理中心")
    col_a, _ = st.columns([1, 3])
    with col_a:
        group_size = st.number_input("📦 分组规模", min_value=1, value=30)

    st.markdown("#### **步骤一：生成展开总表**")
    f1 = st.file_uploader("----请上传原始表格", type=["xlsx"], key="m3f1")
    if f1:
        raw = pd.read_excel(f1, dtype=str).fillna("")
        st.download_button("💾 导出展开总表", data=process_step1(raw, params), file_name=f"{params['prefix']}展开总表.xlsx")

    st.markdown("---")
    st.markdown("#### **步骤二：生成智能分组总表**")
    f2 = st.file_uploader("----请输入处理后的总表 (或上传原始表一键生成)", type=["xlsx"], key="m3f2")
    if f2:
        df_in = pd.read_excel(f2, dtype=str).fillna("")
        res_df = smart_logic(df_in, group_size, params)
        st.download_button("💾 导出智能分组总表", data=write_excel_final(res_df, "分组结果", params, True), file_name=f"{params['prefix']}智能分组总表.xlsx")

def smart_logic(df, size, params):
    # 步骤一：如果还是原始表，先展开
    if "广告素材数量" in df.columns:
        expanded_rows = []
        for _, row in df.iterrows():
            lp_match = re.search(r'-(\d+)$', str(row.get("着陆页版本名称","")))
            lp_v = lp_match.group(1) if lp_match else ""
            for v in expand_material_versions(row, lp_v):
                nr = row.copy(); nr["广告素材版本名称"] = v
                for _ in range(params['repeat_1']): expanded_rows.append(nr.copy())
        df = pd.DataFrame(expanded_rows)

    # 步骤二：清洗与联合去重分组
    df_clean = df[~df["真实SKU"].astype(str).str.contains("总计|空白", na=False)].copy()
    # 联合标识：优先取真实，为空取虚拟
    df_clean["ident"] = df_clean["真实SKU"].str.strip().replace("", None).fillna(df_clean["虚拟SKU"].str.strip())
    # 联合排序
    df_clean = df_clean.sort_values(by=["真实SKU", "虚拟SKU", "着陆页版本名称"], kind="stable")
    
    records = df_clean.to_dict('records')
    buckets = []
    for rec in records:
        sku_id = rec["ident"]
        if not sku_id: continue
        placed = False
        for bucket in buckets:
            if len(bucket) < size and sku_id not in [r["ident"] for r in bucket]:
                bucket.append(rec); placed = True; break
        if not placed: buckets.append([rec])
        
    final = []
    for i, b in enumerate(buckets):
        for r in b:
            r["分组"] = f"分组{i+1}"
            r["分组数量"] = f"{i+1}_{len(b)}"
            r["备注"] = "满足要求" if len(b) == size else "不满足"
            final.append(r)
    return pd.DataFrame(final)

def process_step1(raw, params):
    # 复用展开逻辑... (省略同上)
    return write_excel_final(expanded_df, "展开总表", params)