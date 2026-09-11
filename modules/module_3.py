import streamlit as st
import pandas as pd
import gc
from utils import expand_material_versions, write_excel_final, read_uploaded_excel

def run(params):
    st.subheader("🛠️ 模块三：智能分组处理中心")
    st.caption("步骤一展开素材并可复制次数；步骤二按 SKU 去重分组，不再乘展开次数。")

    prefix = params.get("prefix", "项目_")
    params["repeat_2"] = 1

    st.markdown("**1. 步骤一：生成展开总表**")
    c_up1, c_rep = st.columns(2, gap="medium", vertical_alignment="bottom")
    with c_up1:
        f1 = st.file_uploader("上传原始表格", type=["xlsx"], key="m3f1")
    with c_rep:
        repeat_1 = st.number_input(
            "素材展开次数",
            min_value=1,
            value=1,
            help="每条素材版本再复制几次，只作用于步骤一。",
        )
    params["repeat_1"] = repeat_1

    sig1 = (f1.name if f1 else None, int(repeat_1))
    if st.session_state.get("m3_sig1") != sig1:
        st.session_state["m3_sig1"] = sig1
        st.session_state.pop("m3_step1_bytes", None)
        st.session_state.pop("m3_step1_rows", None)

    c_go1, c_dl1 = st.columns(2, gap="medium")
    with c_go1:
        start1 = st.button(
            "🚀 生成展开总表",
            key="m3_step1_btn",
            use_container_width=True,
            disabled=not bool(f1),
        )

    step1_empty = False
    if start1 and f1:
        raw = read_uploaded_excel(f1.getvalue())
        valid_rows = [
            r for r in raw.to_dict("records")
            if not any("此行为说明" in str(v) or "可不填" in str(v) for v in r.values())
        ]
        step1_data = process_step1(valid_rows, params)
        if step1_data.empty:
            step1_empty = True
            st.session_state.pop("m3_step1_bytes", None)
        else:
            st.session_state["m3_step1_bytes"] = write_excel_final(step1_data, "展开总表", params)
            st.session_state["m3_step1_rows"] = len(step1_data)
        gc.collect()

    has_s1 = "m3_step1_bytes" in st.session_state
    with c_dl1:
        st.download_button(
            f"💾 导出展开总表（{st.session_state.get('m3_step1_rows', 0)} 行）" if has_s1 else "💾 导出展开总表",
            data=st.session_state.get("m3_step1_bytes") or b"",
            file_name=f"{prefix}展开总表.xlsx",
            use_container_width=True,
            disabled=not has_s1,
            key="dl_m3_step1",
        )
    if step1_empty:
        st.warning("⚠️ 未检测到有效业务数据。")
    elif has_s1:
        st.success(f"展开总表已生成（{st.session_state.get('m3_step1_rows', 0)} 行）。")

    st.markdown("**2. 步骤二：生成智能分组总表**")
    c_up2, c_size = st.columns(2, gap="medium", vertical_alignment="bottom")
    with c_up2:
        f2 = st.file_uploader(
            "上传展开总表或原始表",
            type=["xlsx"],
            key="m3f2",
            help="上传步骤一的展开总表只分组。直接传原始表时只展开素材版本（次数=1）再分组。",
        )
    with c_size:
        group_size = st.number_input("分组规模", min_value=1, value=30, help="每组最多容纳多少个不重复 SKU。")

    sig2 = (f2.name if f2 else None, int(group_size))
    if st.session_state.get("m3_sig2") != sig2:
        st.session_state["m3_sig2"] = sig2
        st.session_state.pop("m3_step2_bytes", None)
        st.session_state.pop("m3_step2_rows", None)
        st.session_state.pop("m3_step2_groups", None)

    c_go2, c_dl2 = st.columns(2, gap="medium")
    with c_go2:
        start2 = st.button(
            "🚀 生成智能分组",
            key="m3_step2_btn",
            use_container_width=True,
            disabled=not bool(f2),
        )

    step2_empty = False
    if start2 and f2:
        df_in = read_uploaded_excel(f2.getvalue())
        valid_in = [
            r for r in df_in.to_dict("records")
            if not any("此行为说明" in str(v) or "可不填" in str(v) for v in r.values())
        ]
        res_df = smart_logic(pd.DataFrame(valid_in), group_size, params)
        if res_df.empty:
            step2_empty = True
            st.session_state.pop("m3_step2_bytes", None)
        else:
            st.session_state["m3_step2_bytes"] = write_excel_final(res_df, "分组结果", params, is_m3=True)
            st.session_state["m3_step2_rows"] = len(res_df)
            st.session_state["m3_step2_groups"] = res_df["分组"].nunique() if "分组" in res_df.columns else 0
        gc.collect()

    has_s2 = "m3_step2_bytes" in st.session_state
    s2_rows = st.session_state.get("m3_step2_rows", 0)
    s2_groups = st.session_state.get("m3_step2_groups", 0)
    with c_dl2:
        st.download_button(
            f"💾 导出分组总表（{s2_rows} 行 / {s2_groups} 组）" if has_s2 else "💾 导出分组总表",
            data=st.session_state.get("m3_step2_bytes") or b"",
            file_name=f"{prefix}智能分组总表.xlsx",
            use_container_width=True,
            disabled=not has_s2,
            key="dl_m3_step2",
        )
    if step2_empty:
        st.warning("⚠️ 未生成分组结果，请检查真实/虚拟 SKU。")
    elif has_s2:
        st.success(f"智能分组已生成（{s2_rows} 行 / {s2_groups} 组）。")

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

_SOURCE_EXPAND_COLS = ("广告素材数量", "素材选取 (X-Y)", "素材选取")


def _looks_like_source_template(df):
    """展开总表会删掉这些列；还在说明是原始表，步骤二只需展开版本、不再乘次数。"""
    return any(c in df.columns for c in _SOURCE_EXPAND_COLS)


def smart_logic(df, size, params):
    """步骤二只分组。素材展开次数仅步骤一使用。"""
    if df.empty:
        return df

    if _looks_like_source_template(df):
        df = process_step1(df.to_dict("records"), {"repeat_1": 1})

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