import streamlit as st
import pandas as pd
import re
import sys
import os
import io

# 路径补丁
current_dir = os.path.dirname(os.path.abspath(__file__))
root_path = os.path.abspath(os.path.join(current_dir, '..'))
if root_path not in sys.path:
    sys.path.insert(0, root_path)

from utils import write_excel_final 

def run(params):
    st.subheader("🚀 模块五：多维素材循环填充与拆分中心")
    
    # --- 1. 模式选择与提示信息 ---
    run_mode = st.radio(
        "⚙️ 请选择运行模式",
        ["模式A：固定15行填充 (原有需求)", "模式B：20遍重复并拆分4表 (新需求)"],
        help="模式A固定输出15行大表；模式B将重复生成20遍素材并自动分流打包成4个Excel。"
    )

    if run_mode == "模式A：固定15行填充 (原有需求)":
        st.info("""
        **📌 模式A 运行逻辑：**
        - **15行强制填充**：每一行原始数据固定生成 15 行结果。
        - **循环规则**：不足 15 行时循环重播版本号；超过 15 时自动截断。
        - **备注逻辑**：仅在素材数 > 15 发生截断时显示提醒。
        """)
    else:
        st.info("""
        **📌 模式B 运行逻辑：**
        - **20遍绝对重播**：素材会严格按照 `A-1, A-2...` 的顺序，周而复始地重复平铺 **20个周期**。
        - **多表自动拆分**：
          - 周期 1~5 ➡️ 结果表1.xlsx
          - 周期 6~10 ➡️ 结果表2.xlsx
          - 周期 11~15 ➡️ 结果表3.xlsx
          - 周期 16~20 ➡️ 结果表4.xlsx
        """)

    up = st.file_uploader("📂 上传模块五专用模板", type=["xlsx"], key="m5_up")
    
    if up and st.button("🚀 开始生成", key="m5_btn"):
        df_raw = pd.read_excel(up, dtype=str).fillna("")
        
        display_cols = ["广告账号ID", "主页ID", "像素ID", "真实SKU", "虚拟SKU", "国家", "着陆页版本名称", "广告素材版本名称", "备注"]
        file_prefix = params.get("prefix", "项目_")

        # =====================================================================
        # 核心逻辑一：原有的 15 行强制填充 (模式 A)
        # =====================================================================
        if run_mode == "模式A：固定15行填充 (原有需求)":
            all_results = []
            for idx, row in df_raw.iterrows():
                def safe_int(val, default=0):
                    try:
                        s = str(val).strip()
                        return int(float(s)) if s else default
                    except: return default

                provided_count = safe_int(row.get("提供素材版本数量", 0))
                base_name = str(row.get("广告素材版本名称", "素材"))
                clean_name = re.sub(r'-\d+$', '', base_name)
                
                if provided_count <= 1:
                    material_pool = [f"{clean_name}-1"]
                else:
                    material_pool = [f"{clean_name}-{i}" for i in range(1, provided_count + 1)]
                
                final_material_list = []
                special_note = ""
                if provided_count > 15:
                    special_note = "素材数超标，仅保留前15个版本"

                if len(material_pool) == 1:
                    final_material_list = material_pool * 15
                else:
                    while len(final_material_list) < 15:
                        for m in material_pool:
                            if len(final_material_list) < 15:
                                final_material_list.append(m)
                            else: break
                
                expanded_rows = []
                for v_name in final_material_list:
                    new_row = row.copy()
                    new_row["广告素材版本名称"] = v_name
                    new_row["备注"] = special_note
                    expanded_rows.append(new_row)
                    
                all_results.append(pd.DataFrame(expanded_rows))

            if all_results:
                final_df = pd.concat(all_results, ignore_index=True)
                final_df = final_df[[c for c in display_cols if c in final_df.columns]]
                
                st.success(f"✅ 模式A 处理完成！")
                xlsx_data = write_excel_final(final_df, "15行填充结果", params, color_by="广告账号ID")
                st.download_button(f"💾 下载模式A结果", data=xlsx_data, file_name=f"{file_prefix}15行强制填充.xlsx")

        # =====================================================================
        # 核心逻辑二：20遍循环填充并拆分4表 (模式 B)
        # =====================================================================
        else:
            # 定义4个表格的容器
            table_buckets = {1: [], 2: [], 3: [], 4: []}
            
            for idx, row in df_raw.iterrows():
                def safe_int(val, default=0):
                    try:
                        s = str(val).strip()
                        return int(float(s)) if s else default
                    except: return default

                provided_count = safe_int(row.get("提供素材版本数量", 0))
                base_name = str(row.get("广告素材版本名称", "素材"))
                clean_name = re.sub(r'-\d+$', '', base_name)
                
                # 建立完整的素材版本清单 [A-1, A-2, ..., A-N]
                if provided_count <= 1:
                    material_pool = [f"{clean_name}-1"]
                else:
                    material_pool = [f"{clean_name}-{i}" for i in range(1, provided_count + 1)]
                
                # 为20遍循环分流
                # p 代表第几遍（1 到 20）
                for p in range(1, 21):
                    # 判断当前这一遍应该分流到哪张表
                    if 1 <= p <= 5:
                        bucket_id = 1
                    elif 6 <= p <= 10:
                        bucket_id = 2
                    elif 11 <= p <= 15:
                        bucket_id = 3
                    else:
                        bucket_id = 4
                    
                    # 每一遍，都要把整套素材铺一遍
                    for m in material_pool:
                        new_row = row.copy()
                        new_row["广告素材版本名称"] = m
                        new_row["备注"] = f"第{p}遍生成"
                        table_buckets[bucket_id].append(new_row)

            # 分别导出这 4 张表
            st.success("✅ 模式B 生成完毕！请点击下方按钮分别下载：")
            
            # 循环渲染 4 个下载按钮
            for b_id, rows in table_buckets.items():
                if rows:
                    b_df = pd.DataFrame(rows)
                    b_df = b_df[[c for c in display_cols if c in b_df.columns]]
                    
                    # 定义分表文件名后缀
                    range_name = {1: "1-5遍", 2: "6-10遍", 3: "11-15遍", 4: "16-20遍"}[b_id]
                    
                    xlsx_data = write_excel_final(
                        b_df, 
                        f"新表{b_id}({range_name})", 
                        params, 
                        color_by="广告账号ID"
                    )
                    
                    st.download_button(
                        label=f"💾 下载：新表 {b_id} ({range_name}结果)",
                        data=xlsx_data,
                        file_name=f"{file_prefix}新表{b_id}_{range_name}.xlsx",
                        key=f"m5_b_btn_{b_id}"
                    )