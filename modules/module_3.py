import streamlit as st
import pandas as pd
import re
import sys
import os

# --- 第一步：路径补丁 ---
current_dir = os.path.dirname(os.path.abspath(__file__))
root_path = os.path.abspath(os.path.join(current_dir, '..'))
if root_path not in sys.path:
    sys.path.insert(0, root_path)

# 导入工具函数
from utils import expand_material_versions, write_excel_final

def run(params):
    st.subheader("🛠️ 模块三：智能分组处理中心")
    
    # UI 交互
    col_a, col_b = st.columns(2)
    with col_a:
        group_size = st.number_input("📦 分组规模", min_value=1, value=30)
    with col_b:
        repeat_1 = st.number_input("🔄 素材展开次数", min_value=1, value=1, help="同组内对应广告素材版本重复次数")
        params['repeat_1'] = repeat_1
        params['repeat_2'] = 1

    st.markdown("#### **步骤一：生成展开总表**")
    f1 = st.file_uploader("----请上传原始表格", type=["xlsx"], key="m3f1")
    if f1:
        raw = pd.read_excel(f1, dtype=str).fillna("")
        step1_data = process_step1(raw, params)
        st.download_button(
            "💾 导出展开总表", 
            data=write_excel_final(step1_data, "展开总表", params), 
            file_name=f"{params.get('prefix', '项目_')}展开总表.xlsx"
        )

    st.markdown("---")
    st.markdown("#### **步骤二：生成智能分组总表**")
    f2 = st.file_uploader("----请输入处理后的总表 (或上传原始表一键生成)", type=["xlsx"], key="m3f2")
    if f2:
        df_in = pd.read_excel(f2, dtype=str).fillna("")
        res_df = smart_logic(df_in, group_size, params)
        st.download_button(
            "💾 导出智能分组总表", 
            data=write_excel_final(res_df, "分组结果", params, is_m3=True), 
            file_name=f"{params.get('prefix', '项目_')}智能分组总表.xlsx"
        )

def process_step1(df, params):
    """专门负责将原始表按素材数量展开的逻辑"""
    expanded_rows = []
    for _, row in df.iterrows():
        lp_match = re.search(r'-(\d+)$', str(row.get("着陆页版本名称", "")))
        lp_v = lp_match.group(1) if lp_match else ""
        versions = expand_material_versions(row, lp_v)
        for v in versions:
            nr = row.copy()
            nr["广告素材版本名称"] = v
            for _ in range(params.get('repeat_1', 1)):
                expanded_rows.append(nr.copy())
    return pd.DataFrame(expanded_rows)

def smart_logic(df, size, params):
    # 如果上传的是原始表（含有广告素材数量列），先执行展开
    if "广告素材数量" in df.columns:
        df = process_step1(df, params)

    # 数据清洗：剔除无关行
    df_clean = df[~df["真实SKU"].astype(str).str.contains("总计|空白", na=False)].copy()
    
    # 联合标识：优先取真实 SKU，为空则取虚拟 SKU
    df_clean["ident"] = df_clean["真实SKU"].str.strip().replace("", None).fillna(df_clean["虚拟SKU"].str.strip())
    
    # 稳定排序
    df_clean = df_clean.sort_values(by=["真实SKU", "虚拟SKU", "着陆页版本名称"], kind="stable")
    
    records = df_clean.to_dict('records')
    
    # ✨ 重构桶结构：解耦内存引用，并引入高速集合 Set 查重机制
    buckets = []
    
    for rec in records:
        sku_id = rec["ident"]
        if not sku_id: continue
        placed = False
        
        for bucket in buckets:
            # 桶未满 且 该 SKU 还没在这个桶里
            if len(bucket['data']) < size and sku_id not in bucket['skus']:
                # ✨ 核心修复：追加副本 rec.copy()，切断指针
                bucket['data'].append(rec.copy())
                bucket['skus'].add(sku_id)
                placed = True
                break
                
        if not placed:
            # ✨ 核心修复：新建桶时也使用 rec.copy() 保持独立
            buckets.append({
                'data': [rec.copy()],
                'skus': {sku_id}
            })
        
    final = []
    for i, bucket in enumerate(buckets):
        g_id = i + 1
        g_len = len(bucket['data'])
        status_text = "满足要求" if g_len == size else "不满足"
        
        for r in bucket['data']:
            # 此时各个 r 已经在内存中相互隔离，打标签再也不会相互覆盖了！
            r["分组"] = f"分组{g_id}"
            r["分组数量"] = f"{g_id}_{g_len}"
            r["备注"] = status_text
            final.append(r)
            
    res_df = pd.DataFrame(final)
    
    # 导出前顺手把多余的 ident 列删掉，还运营一份干净的表格
    if "ident" in res_df.columns:
        res_df = res_df.drop(columns=["ident"])
        
    return res_df