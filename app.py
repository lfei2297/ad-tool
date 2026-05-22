import streamlit as st
import pandas as pd
import re
from collections import defaultdict
import io
import zipfile

# ==========================
# 页面配置
# ==========================
st.set_page_config(page_title="广告素材生成工具 v14", page_icon="🚀", layout="wide")
st.title("🚀 广告素材批量生成工具")

# ==========================
# 左侧边栏
# ==========================
st.sidebar.header("🎯 核心功能选择")
PROCESS_MODE = st.sidebar.radio(
    "🔄 请选择工作模式：",
    ["模块一：基础独立拆分", "模块二：同SKU+国家聚合拆分", "模块三：智能分组 (SKU去重)"]
)

st.sidebar.markdown("---")
st.sidebar.header("⚙️ 全局参数设置")
FILE_PREFIX = st.sidebar.text_input("✏️ 自定义结果文件前缀", value="项目A_")
REPEAT_FIRST = st.sidebar.number_input("广告组数", min_value=1, value=1)
REPEAT_SECOND = st.sidebar.number_input("系列数", min_value=1, value=1)
ENABLE_COLOR = st.sidebar.checkbox("开启颜色标记", value=True)
FAST_MODE = st.sidebar.checkbox("开启极速模式 (跳过样式渲染)", value=False)

# ==========================
# 核心解析函数
# ==========================
def natural_sort_key(s):
    return [int(t) if t.isdigit() else t.lower() for t in re.split('([0-9]+)', str(s))]

def expand_material_versions(row):
    base_name = str(row.get("广告素材版本名称", "")).strip()
    lp_name = str(row.get("着陆页版本名称", "")).strip()
    try:
        material_count = int(row["广告素材数量"]) if row["广告素材数量"] else 1
    except:
        material_count = 1
    material_select = str(row.get("素材选取", "") or row.get("素材选取 (X-Y)", "")).strip()
    lp_match = re.search(r'-(\d+)$', lp_name)
    lp_version = lp_match.group(1) if lp_match else ""
    sel_match = re.search(r'(\d+)-(\d+)', material_select)
    s_m, e_m = (int(sel_match.group(1)), int(sel_match.group(2))) if sel_match else (0,0)

    m_two = re.search(r'-(\d{2})$', base_name)
    m_hyphen = re.search(r'-(\d+)-(\d+)$', base_name)

    if m_two:
        suffix = m_two.group(1)
        g_n, s_mat = suffix[0], int(suffix[1])
        new_prefix = f"{base_name[:m_two.start()]}-{g_n}-{lp_version}-" if lp_version else f"{base_name[:m_two.start()]}-{g_n}-"
        start, end = (s_m, e_m) if sel_match else (s_mat, s_mat + material_count - 1)
        return [f"{new_prefix}{i}" for i in range(start, end + 1)]
    elif m_hyphen:
        g_p, s_mat = m_hyphen.group(1), int(m_hyphen.group(2))
        new_prefix = f"{base_name[:m_hyphen.start()]}-{g_p}-"
        start, end = (s_m, e_m) if sel_match else (s_mat, s_mat + material_count - 1)
        return [f"{new_prefix}{i}" for i in range(start, end + 1)]
    else:
        m_s = re.search(r'-(\d+)$', base_name)
        start_mat = int(m_s.group(1)) if m_s else 1
        prefix = base_name[:m_s.start()] + "-" if m_s else base_name + "-"
        start, end = (s_m, e_m) if sel_match else (start_mat, start_mat + material_count - 1)
        return [f"{prefix}{i}" for i in range(start, end + 1)]

def smart_grouping_logic(df_to_group, group_size):
    # 过滤无效行
    df_clean = df_to_group[df_to_group["真实SKU"].astype(str).str.strip() != ""].copy()
    df_clean = df_clean[~df_clean["真实SKU"].astype(str).str.contains("总计|空白", na=False)]
    
    records = df_clean.to_dict('records')
    # 预排序增加离散度稳定性
    records = sorted(records, key=lambda x: (str(x.get("真实SKU","")), str(x.get("着陆页版本名称",""))))
    
    buckets = []
    for rec in records:
        sku = (str(rec.get("真实SKU", "")).strip() or str(rec.get("虚拟SKU", "")).strip())
        placed = False
        for bucket in buckets:
            if len(bucket) >= group_size: continue
            if sku not in { (str(r.get("真实SKU","")).strip() or str(r.get("虚拟SKU","")).strip()) for r in bucket }:
                bucket.append(rec); placed = True; break
        if not placed: buckets.append([rec])
    
    final_rows = []
    for idx, bucket in enumerate(buckets):
        count = len(bucket)
        for row in bucket:
            row["分组"] = f"分组{idx+1}"
            row["分组数量"] = f"{idx+1}_{count}"
            row["备注"] = "满足要求" if count == group_size else "不满足"
            final_rows.append(row)
    return pd.DataFrame(final_rows)

def write_excel_clean(df, sheet_name, is_m3=False):
    cols_to_del = ["广告素材数量", "素材选取 (X-Y)", "素材选取"]
    df_out = df.drop(columns=[c for c in cols_to_del if c in df.columns])
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_out.to_excel(writer, index=False, sheet_name=sheet_name)
        if not FAST_MODE:
            workbook, worksheet = writer.book, writer.sheets[sheet_name]
            for i, col in enumerate(df_out.columns):
                worksheet.set_column(i, i, max(len(str(col)), 15))
            if ENABLE_COLOR:
                colors = ["#FFF2CC", "#E2EFDA", "#DDEBF7", "#F8CBAD", "#E4DFEC", "#D9D9D9", "#EBF1DE"]
                color_map = {}
                color_col = "分组" if (is_m3 and "分组" in df_out.columns) else "_ck"
                if not is_m3: df_out["_ck"] = df_out["真实SKU"].astype(str) + df_out["虚拟SKU"].astype(str)
                keys = df_out[color_col].unique()
                for i, k in enumerate(keys): color_map[k] = colors[i % len(colors)]
                for r in range(len(df_out)):
                    fmt = workbook.add_format({"bg_color": color_map[df_out.iloc[r][color_col]]})
                    worksheet.set_row(r+1, None, fmt)
    return output.getvalue()

# ==========================
# 业务逻辑界面
# ==========================
if "模块三" in PROCESS_MODE:
    st.subheader("🛠️ 模块三：智能分组处理中心")
    col_a, col_b = st.columns([1, 3])
    with col_a:
        GROUP_SIZE = st.number_input("📦 分组规模 (每组数量)", min_value=1, value=30)

    st.markdown("#### **步骤一：生成展开总表**")
    f1 = st.file_uploader("----请上传原始表格", type=["xlsx"], key="m3s1")
    if f1:
        raw = pd.read_excel(f1, dtype=str).fillna("")
        exp = []
        for _, row in raw.iterrows():
            for v in expand_material_versions(row):
                nr = row.copy(); nr["广告素材版本名称"] = v
                for _ in range(REPEAT_FIRST): exp.append(nr.copy())
        st1_df = pd.DataFrame(exp)
        if REPEAT_SECOND > 1: st1_df = pd.concat([st1_df] * REPEAT_SECOND, ignore_index=True)
        st.download_button("💾 导出展开后的总表", data=write_excel_clean(st1_df, "展开总表"), file_name=f"{FILE_PREFIX}展开总表.xlsx")

    st.markdown("---")
    st.markdown("#### **步骤二：生成智能分组总表**")
    f2 = st.file_uploader("----请输入数据处理后的总表 (或直接上传原始表一键生成)", type=["xlsx"], key="m3s2")
    if f2:
        df_in = pd.read_excel(f2, dtype=str).fillna("")
        # 判定是否需要执行步骤一展开
        if "广告素材数量" in df_in.columns:
            exp = []
            for _, row in df_in.iterrows():
                for v in expand_material_versions(row):
                    nr = row.copy(); nr["广告素材版本名称"] = v
                    for _ in range(REPEAT_FIRST): exp.append(nr.copy())
            df_g = pd.DataFrame(exp)
        else:
            df_g = df_in
        res = smart_grouping_logic(df_g, GROUP_SIZE)
        st.download_button("💾 导出智能分组总表", data=write_excel_clean(res, "分组结果", True), file_name=f"{FILE_PREFIX}智能分组总表.xlsx")

else:
    up = st.file_uploader("📂 上传原始素材表 (.xlsx)", type=["xlsx"])
    if up and st.button("🚀 开始处理"):
        df_raw = pd.read_excel(up, dtype=str).fillna("")
        tasks = {}
        m2_agg = defaultdict(list)
        m2_counts = defaultdict(int)
        m1_groups = defaultdict(list)
        all_exp = []

        for _, row in df_raw.iterrows():
            vs = expand_material_versions(row)
            m_len = len(vs)
            row_rows = []
            for v in vs:
                nr = row.copy(); nr["广告素材版本名称"] = v
                for _ in range(REPEAT_FIRST): row_rows.append(nr.copy())
            tmp = pd.DataFrame(row_rows)
            if REPEAT_SECOND > 1: tmp = pd.concat([tmp] * REPEAT_SECOND, ignore_index=True)
            all_exp.append(tmp)
            
            if "模块一" in PROCESS_MODE:
                m1_groups[f"素材数_{m_len}"].append(tmp)
            elif "模块二" in PROCESS_MODE:
                sku = str(row.get("真实SKU","")).strip() or str(row.get("虚拟SKU","")).strip()
                country = str(row.get("国家","")).strip()
                m2_agg[(sku, country)].append(tmp); m2_counts[(sku, country)] += m_len

        # ✨ 模块二已根据您的要求取消强制排序逻辑
        total_df = pd.concat(all_exp, ignore_index=True)
        tasks["总表"] = total_df

        if "模块一" in PROCESS_MODE:
            for k, v in m1_groups.items(): tasks[k] = pd.concat(v, ignore_index=True)
        else:
            for (sku, country), dfs in m2_agg.items():
                tag = f"{country}_" if country else ""
                tasks[f"{tag}素材数_{m2_counts[(sku, country)]}"] = pd.concat(dfs, ignore_index=True)

        zip_b = io.BytesIO()
        with zipfile.ZipFile(zip_b, "w") as zf:
            for n, d in tasks.items():
                zf.writestr(f"{FILE_PREFIX}{n}.xlsx", write_excel_clean(d, "Data"))
        st.download_button("📦 下载处理结果包", data=zip_b.getvalue(), file_name=f"{FILE_PREFIX}结果.zip")
