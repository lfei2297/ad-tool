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
    
    # UI 交互参数设置
    col_a, col_b = st.columns(2)
    with col_a:
        group_size = st.number_input("📦 分组规模", min_value=1, value=30)
    with col_b:
        repeat_1 = st.number_input("🔄 素材展开次数", min_value=1, value=1, help="同一广告素材版本重复次数")
        # 实时更新全局/局部参数包
        params['repeat_1'] = repeat_1
        params['repeat_2'] = 1  # 模块三强制整体复制为1次

    # =========================================================================
    # 步骤一：智能生成展开总表
    # =========================================================================
    st.markdown("#### **步骤一：生成展开总表**")
    f1 = st.file_uploader("----请上传原始表格", type=["xlsx"], key="m3f1")
    if f1 and st.button("🚀 生成步骤一总表", key="m3_btn_s1"):
        raw = pd.read_excel(f1, dtype=str).fillna("")
        step1_df = process_step1(raw, params)
        
        st.success("✅ 步骤一总表展开成功（已执行素材重复）")
        xlsx_data = write_excel_final(step1_df, "展开总表", params)
        st.download_button(
            "💾 下载：步骤一展开总表", 
            data=xlsx_data, 
            file_name=f"{params['prefix']}步骤一_展开总表.xlsx"
        )

    st.markdown("---")
    
    # =========================================================================
    # 步骤二：直接用表格生成分组表（修复重复漏算 Bug）
    # =========================================================================
    st.markdown("#### **步骤二：生成智能分组总表**")
    f2 = st.file_uploader("----请上传需要分组的表格（支持原始表或步骤一总表）", type=["xlsx"], key="m3f2")
    if f2 and st.button("🚀 生成步骤二分组表", key="m3_btn_s2"):
        raw_f2 = pd.read_excel(f2, dtype=str).fillna("")
        
        # ✨ 核心修复点：这里显式调用智能分配算法，并将完整的 params 传入
        # 底层 smart_logic 会自动判断：如果是原始表，会自动进行素材展开并应用 repeat_1 的重复次数
        res_df = smart_logic(raw_f2, group_size, params)
        
        if not res_df.empty:
            st.success("✅ 步骤二智能分组成功（已应用同组SKU冲突避让机制）")
            # 导出模块三的分组结果，传 True 激活按“分组”列切色
            data = write_excel_final(res_df, "分组结果", params, is_m3=True)
            st.download_button(
                "💾 下载：步骤二智能分组表", 
                data=data, 
                file_name=f"{params['prefix']}步骤二_智能分组表.xlsx"
            )

# =========================================================================
# ⚙️ 底层支持函数
# =========================================================================

def process_step1(df, params):
    """提取原始表中的多版本素材，并根据 repeat_1 数量进行克隆扩展"""
    expanded_rows = []
    for _, row in df.iterrows():
        lp_match = re.search(r'-(\d+)$', str(row.get("着陆页版本名称", "")))
        lp_v = lp_match.group(1) if lp_match else ""
        vs = expand_material_versions(row, lp_v)
        
        for v in vs:
            nr = row.copy()
            nr["广告素材版本名称"] = v
            # 根据用户在前端输入的“素材展开次数”进行行克隆
            for _ in range(params.get('repeat_1', 1)):
                expanded_rows.append(nr.copy())
    return pd.DataFrame(expanded_rows)

def smart_logic(df, size, params):
    """智能分组核心资产避让算法"""
    # ✨ 核心修复点：如果用户在步骤二直接上传了包含“广告素材数量”的原始表，
    # 必须先执行展开和“素材重复克隆”流程，确保行数扩充。
    if "广告素材数量" in df.columns:
        df = process_step1(df, params)

    # 数据清洗：剔除表格底部可能遗留的“总计”或“空白”杂行
    df_clean = df[~df["真实SKU"].astype(str).str.contains("总计|空白", na=False)].copy()
    
    # 联合冲突标识：优先监控真实SKU，若为空则降级监控虚拟SKU
    df_clean["ident"] = df_clean["真实SKU"].str.strip().replace("", None).fillna(df_clean["虚拟SKU"].str.strip())
    
    # 执行稳定排序，保证大货货源顺序不被打乱
    df_clean = df_clean.sort_values(by=["真实SKU", "虚拟SKU", "着陆页版本名称"], kind="stable")
    
    records = df_clean.to_dict('records')
    buckets = []
    
    # 动态散列切片：严格确保同一个 SKU（或虚拟SKU）不会在同一个 Bucket 中相遇
    for rec in records:
        sku_id = rec["ident"]
        if not sku_id: continue
        placed = False
        
        for bucket in buckets:
            # 规则：当前桶没装满（如没满30条） 且 当前桶里没有出现过这个 SKU
            if len(bucket['data']) < size and sku_id not in bucket['skus']:
                bucket['data'].append(rec)
                bucket['skus'].add(sku_id)
                placed = True
                break
                
        # 触发出错或无坑位可蹲，创建新桶接收
        if not placed:
            buckets.append({'data': [rec], 'skus': {sku_id}})
            
    # 从桶中提取数据，并追加“分组数量”、“备注”、“分组”等可视化指标
    final_rows = []
    for i, bucket in enumerate(buckets):
        g_id = i + 1
        g_len = len(bucket['data'])
        status_text = "满足要求" if g_len == size else "不满足"
        
        for rec in bucket['data']:
            nr = rec.copy()
            nr["分组数量"] = f"{g_id}_{g_len}"
            nr["备注"] = status_text
            nr["分组"] = f"分组{g_id}"
            final_rows.append(nr)
            
    res_df = pd.DataFrame(final_rows)
    # 清理掉用于算法比对的辅助列
    if "ident" in res_df.columns:
        res_df = res_df.drop(columns=["ident"])
    return res_df