import streamlit as st
import pandas as pd
import gc
from utils import expand_material_versions, write_excel_final, read_uploaded_excel

def run(params):
    st.subheader("🛠️ 模块三：智能分组处理中心")
    
    col_a, col_b = st.columns(2)
    with col_a:
        group_size = st.number_input("📦 分组规模", min_value=1, value=30)
    with col_b:
        repeat_1 = st.number_input("🔄 素材展开次数", min_value=1, value=1)
        params['repeat_1'] = repeat_1
        params['repeat_2'] = 1

    # ----------------------------------------------------
    # 步骤一：生成展开总表
    # ----------------------------------------------------
    st.markdown("#### **步骤一：生成展开总表**")
    f1 = st.file_uploader("----请上传原始表格", type=["xlsx"], key="m3f1")
    if f1:
        raw = read_uploaded_excel(f1.getvalue())
        # 过滤说明行
        valid_rows = [r for r in raw.to_dict('records') if not any('此行为说明' in str(v) or '可不填' in str(v) for v in r.values())]
        step1_data = process_step1(valid_rows, params)
        
        st.download_button(
            "💾 导出展开总表", 
            data=write_excel_final(step1_data, "展开总表", params), 
            file_name=f"{params.get('prefix', '项目_')}展开总表.xlsx",
            use_container_width=True
        )

    st.markdown("---")
    
    # ----------------------------------------------------
    # 步骤二：生成智能分组总表
    # ----------------------------------------------------
    st.markdown("#### **步骤二：生成智能分组总表**")
    f2 = st.file_uploader("----请输入处理后的总表 (或上传原始表一键生成)", type=["xlsx"], key="m3f2")
    if f2:
        df_in = read_uploaded_excel(f2.getvalue())
        valid_in = [r for r in df_in.to_dict('records') if not any('此行为说明' in str(v) or '可不填' in str(v) for v in r.values())]
        
        # 核心逻辑运算
        res_df = smart_logic(pd.DataFrame(valid_in), group_size, params)
        
        st.download_button(
            "💾 导出智能分组总表", 
            data=write_excel_final(res_df, "分组结果", params, is_m3=True), 
            file_name=f"{params.get('prefix', '项目_')}智能分组总表.xlsx",
            use_container_width=True
        )

def process_step1(records, params):
    """高效处理素材版本展开与重复展开"""
    expanded_rows = []
    repeat_val = params.get('repeat_1', 1)
    
    for row_dict in records:
        versions = expand_material_versions(row_dict)
        for v in versions:
            nr = row_dict.copy()
            nr["广告素材版本名称"] = v
            if repeat_val == 1:
                expanded_rows.append(nr)
            else:
                # 按照用户设定的 repeat_1 复制对应次数
                expanded_rows.extend([nr.copy() for _ in range(repeat_val)])
                
    return pd.DataFrame(expanded_rows)

def smart_logic(df, size, params):
    """100% 还原原有功能 + 极速性能优化版"""
    if df.empty:
        return df

    # 1. 完全恢复原有逻辑：只要包含“广告素材数量”或设置了重复展开 (repeat_1 > 1)，均强制重新计算展开
    if "广告素材数量" in df.columns or params.get('repeat_1', 1) > 1:
        df = process_step1(df.to_dict('records'), params)

    # 安全检查列是否存在
    if "真实SKU" not in df.columns: df["真实SKU"] = ""
    if "虚拟SKU" not in df.columns: df["虚拟SKU"] = ""

    # 2. 过滤无用汇总行/空白行
    df_clean = df[~df["真实SKU"].astype(str).str.contains("总计|空白", na=False)].copy()
    
    # 3. 计算唯一性标识 (ident)
    r_skus = df_clean["真实SKU"].astype(str).str.strip().replace(["nan", "None", ""], None)
    v_skus = df_clean["虚拟SKU"].astype(str).str.strip().replace(["nan", "None", ""], None)
    df_clean["ident"] = r_skus.fillna(v_skus)
    
    # 4. 安定排序
    sort_cols = [c for c in ["真实SKU", "虚拟SKU", "着陆页版本名称"] if c in df_clean.columns]
    if sort_cols:
        df_clean = df_clean.sort_values(by=sort_cols, kind="stable")
    
    records = df_clean.to_dict('records')
    buckets = []
    bucket_sku_sets = []
    
    # 5. 快速桶分组计算
    for rec in records:
        sku_id = rec.get("ident")
        if not sku_id: continue
        placed = False
        
        # 遍历已建立的桶
        for bucket, b_set in zip(buckets, bucket_sku_sets):
            if len(bucket) < size and sku_id not in b_set:
                bucket.append(rec)
                b_set.add(sku_id)
                placed = True
                break
                
        if not placed:
            buckets.append([rec])
            bucket_sku_sets.append({sku_id})
        
    # 6. 结果集快速构建
    final = []
    for i, b in enumerate(buckets):
        group_name = f"分组{i+1}"
        group_count = f"{i+1}_{len(b)}"
        group_note = "满足要求" if len(b) == size else "不满足"
        
        for r in b:
            # 清理中间临时列
            r.pop("ident", None)
            
            # 填入分组字段
            r["分组"] = group_name
            r["分组数量"] = group_count
            r["备注"] = group_note
            final.append(r)
            
    return pd.DataFrame(final)