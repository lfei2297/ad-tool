import streamlit as st
import pandas as pd
import io
import zipfile
import gc
from collections import defaultdict

# ─────────────────────────────────────────────
# 1. 动态表头与自适应导出引擎
# ─────────────────────────────────────────────

def is_landing_page_url(val_series):
    """根据单元格内容判断 G 列是链接还是版本名称。空值或说明行不当成链接。"""
    skip_tokens = ("可不填", "此行为说明", "填写完整", "落地页", "着陆页库")
    sample_vals = []
    for v in val_series.dropna():
        s = str(v).strip()
        if not s or s.lower() in ("nan", "none"):
            continue
        if any(tok in s for tok in skip_tokens):
            continue
        sample_vals.append(s)
    if not sample_vals:
        return False

    link_indicators = (
        "http://", "https://", "www.", ".com", ".cn", ".top", ".shop",
        ".net", ".org", ".site", "/products/", "/funnel/",
    )
    link_count = 0
    for v in sample_vals:
        v_lower = v.lower()
        if any(ind in v_lower for ind in link_indicators):
            link_count += 1
    return (link_count / len(sample_vals)) >= 0.5


def write_landing_page_excel(df_data, raw_hint_dict=None, keep_material_col="广告素材ID"):
    df_out = df_data.copy().fillna("")
    drop_tool_cols = [c for c in ("SKU_KEY", "SKU组") if c in df_out.columns]
    if drop_tool_cols:
        df_out = df_out.drop(columns=drop_tool_cols)

    keep_material_col = str(keep_material_col).strip()
    if keep_material_col not in ("广告素材ID", "广告素材版本名称"):
        keep_material_col = "广告素材ID"

    # 1. 识别 G 列与 H 列
    g_col_candidates = [c for c in df_out.columns if "着陆页" in c]
    g_actual_col = g_col_candidates[0] if g_col_candidates else "着陆页链接"
    
    is_link = is_landing_page_url(df_out[g_actual_col])

    # G 列动态切换；H 列为广告素材版本名称，其后紧跟广告素材ID（两列都导出）
    if is_link:
        g1_header = "着陆页链接"
        g2_hint = "填写完整的商品链接"
    else:
        g1_header = "着陆页版本名称"
        g2_hint = "着陆页库中的具体版本名称"

    h1_header = "广告素材版本名称"

    # 更名 G 列（如果有差异）
    if g_actual_col != g1_header:
        df_out = df_out.rename(columns={g_actual_col: g1_header})

    # 🌟 核心防错：彻底删除历史旧的【广告素材版本】列，防止与【广告素材版本名称】重名冲突
    if "广告素材版本" in df_out.columns and "广告素材版本名称" in df_out.columns:
        df_out = df_out.drop(columns=["广告素材版本"])
    elif "广告素材版本" in df_out.columns:
        df_out = df_out.rename(columns={"广告素材版本": h1_header})

    # 两列都必须出现在结果里；表格导入内容二选一，未选中的那一列留空
    if h1_header not in df_out.columns:
        df_out[h1_header] = ""
    if "广告素材ID" not in df_out.columns:
        df_out["广告素材ID"] = ""
    other_material_col = "广告素材版本名称" if keep_material_col == "广告素材ID" else "广告素材ID"
    df_out[other_material_col] = ""

    # 🌟 列名强制去重（仅保留第一次出现的同名列）
    df_out = df_out.loc[:, ~df_out.columns.duplicated()].copy()

    # 2. 动态调整列顺序：广告素材版本名称在前，广告素材ID 紧随其后
    current_cols = list(df_out.columns)
    standard_skeleton = [
        "广告账号ID", "主页ID", "像素ID", "真实SKU", "虚拟SKU",
        "国家", g1_header, h1_header, "广告素材ID", "出价/竞价", "系列标注"
    ]
    
    # 重新排列：把存在的标准列按顺序排在前面，其余自定义扩展列按原样追加在后面
    ordered_cols = [c for c in standard_skeleton if c in current_cols]
    extra_cols = [c for c in current_cols if c not in standard_skeleton and "Unnamed" not in str(c)]
    unnamed_cols = [c for c in current_cols if "Unnamed" in str(c)]
    
    # 最终组合：标准列 + 用户自定义扩展列 + 说明列（放在最右侧）
    current_cols = ordered_cols + extra_cols + unnamed_cols
    df_out = df_out[current_cols]

    # 3. 构造提示行 (Hints)
    default_hints = {
        "广告账号ID": "",
        "主页ID": "可不填，不填则使用资产管理中的默认主页",
        "像素ID": "可不填，不填则使用资产管理中的默认像素",
        "真实SKU": "",
        "虚拟SKU": "",
        "国家": "美国/英国/德国/法国/西班牙",
        "着陆页链接": "填写完整的一条商品链接，用于广告投放的落地页链接",
        "着陆页版本名称": "着陆页库中的具体版本名称",
        "广告素材版本名称": "产品库中的具体广告素材版本名称，会取该版本下最新的一条广告素材；广告素材版本名称和广告素材ID这两列二选一，均可不填",
        "广告素材ID": "产品素材库中具体的广告素材ID；广告素材版本名称和广告素材ID这两列二选一，均可不填",
        "出价/竞价": "如需指定“真实/虚拟SKU”与“出价/竞价”的关系，请填写，最多2位小数，可不填，不填则全不填，填了则全填",
        "系列标注": "可不填",
        "Unnamed: 10": "此行为说明，勿删除，请从第三行开始填写"
    }

    final_hints = {}
    for col in current_cols:
        if raw_hint_dict and col in raw_hint_dict and str(raw_hint_dict[col]).strip() not in ["", "nan", "None"]:
            final_hints[col] = raw_hint_dict[col]
        else:
            final_hints[col] = default_hints.get(col, "可不填" if "Unnamed" not in str(col) else "此行为说明，勿删除，请从第三行开始填写")

    hint_row = pd.DataFrame([final_hints])
    
    # 保证 hint_row 与 df_out 列完全一致且无重复
    hint_row = hint_row.reindex(columns=df_out.columns)
    final_export_df = pd.concat([hint_row, df_out], ignore_index=True)

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        final_export_df.to_excel(writer, index=False, sheet_name="Sheet1")
        workbook = writer.book
        worksheet = writer.sheets["Sheet1"]
        
        worksheet.autofilter(0, 0, 0, len(current_cols) - 1)
        header_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#E0E0E0', 'text_wrap': True, 'valign': 'vcenter', 'font_name': '宋体', 'font_size': 11,
        })
        hint_fmt = workbook.add_format({
            'bg_color': '#FFFFCC', 'font_color': 'red', 'bold': True, 'text_wrap': True, 'valign': 'top', 'font_name': '宋体', 'font_size': 11,
        })
        
        worksheet.set_row(0, 22, header_fmt)
        worksheet.set_row(1, 56, hint_fmt)

        export_widths = {
            "广告账号ID": 16,
            "主页ID": 16,
            "像素ID": 16,
            "真实SKU": 12,
            "虚拟SKU": 12,
            "国家": 12,
            "着陆页链接": 16,
            "着陆页版本名称": 16,
            "广告素材版本名称": 22,
            "广告素材ID": 22,
            "出价/竞价": 16,
            "系列标注": 16,
            "SKU组": 16,
        }
        for i, col in enumerate(current_cols):
            col_name_str = "" if "Unnamed" in str(col) else str(col)
            worksheet.write(0, i, col_name_str, header_fmt)
            if "Unnamed" in str(col):
                w = 16
            else:
                w = export_widths.get(col_name_str, 16)
            worksheet.set_column(i, i, w)

    return out.getvalue()


# ─────────────────────────────────────────────
# 2. 核心匹配与拆表算法
# ─────────────────────────────────────────────

def get_sku_key(row):
    v = str(row.get("虚拟SKU", "")).strip()
    r = str(row.get("真实SKU", "")).strip()
    if v and v.lower() != "nan":
        return v
    if r and r.lower() != "nan":
        return r
    return ""


_DROP_EXPAND_COLS = ("SKU_KEY", "广告素材版本", "SKU组")


def _expand_sku_rows(acc_item, mats):
    base = acc_item.to_dict() if hasattr(acc_item, "to_dict") else dict(acc_item)
    for col in _DROP_EXPAND_COLS:
        base.pop(col, None)
    rows = []
    for mat_v, mat_i in mats:
        row = base.copy()
        row["广告素材版本名称"] = mat_v
        row["广告素材ID"] = mat_i
        rows.append(row)
    return rows


def _order_matched_columns(df):
    preferred = [
        "广告账号ID", "主页ID", "像素ID", "真实SKU", "虚拟SKU",
        "国家", "着陆页链接", "着陆页版本名称", "广告素材版本名称", "广告素材ID",
        "出价/竞价", "系列标注",
    ]
    cols = list(df.columns)
    ordered = [c for c in preferred if c in cols]
    extra = [c for c in cols if c not in ordered]
    return df[ordered + extra]


def _unique_preserve(seq):
    seen = set()
    out = []
    for x in seq:
        if x in seen:
            continue
        seen.add(x)
        out.append(x)
    return out


def _filled(val):
    s = str(val).strip() if val is not None else ""
    return bool(s) and s.lower() not in ("nan", "none")


def _pick_richer_materials(acc_item, sku_materials_by_any):
    """真实 SKU / 虚拟 SKU 两边都查，用素材条数更多的那份。"""
    v = str(acc_item.get("虚拟SKU", "")).strip()
    r = str(acc_item.get("真实SKU", "")).strip()
    mats_v = list(sku_materials_by_any.get(v, [])) if _filled(v) else []
    mats_r = list(sku_materials_by_any.get(r, [])) if _filled(r) else []
    if len(mats_r) > len(mats_v):
        chosen = mats_r
    elif mats_v:
        chosen = mats_v
    else:
        chosen = mats_r
    return chosen


def _norm_sku_group(val):
    """SKU组填写 1/2/3；Excel 可能写成 1.0，统一成整数文本。"""
    s = str(val).strip() if val is not None else ""
    if not _filled(s):
        return ""
    try:
        f = float(s)
        if f == int(f):
            return str(int(f))
    except (TypeError, ValueError):
        pass
    return s


def _split_acc_sku_groups(acc_items):
    """同一账号下按【SKU组】拆成多组；都为空则整号一组。不使用系列标注。"""
    if not any(_norm_sku_group(it.get("SKU组")) for it in acc_items):
        return [list(acc_items)]

    order = []
    buckets = {}
    for it in acc_items:
        g = _norm_sku_group(it.get("SKU组"))
        if not g:
            g = "__default__"
        if g not in buckets:
            buckets[g] = []
            order.append(g)
        buckets[g].append(it)
    return [buckets[g] for g in order]


def _round_robin_stream(mat_lists):
    """先平铺每个 SKU 第 1 条，再按 SKU 顺序轮流追加后续素材。"""
    n = len(mat_lists)
    if n == 0:
        return []
    ptr = [0] * n
    stream = []
    while True:
        progressed = False
        for i in range(n):
            mats = mat_lists[i]
            if ptr[i] < len(mats):
                stream.append((i, mats[ptr[i]]))
                ptr[i] += 1
                progressed = True
        if not progressed:
            break
    return stream


def _expand_stream_rows(sku_items, stream):
    rows = []
    for sku_i, mat in stream:
        rows.extend(_expand_sku_rows(sku_items[sku_i], [mat]))
    return rows


def _build_multiproduct_series(group_items, sku_materials_by_any, mode, n_group, per_sku_limit=None):
    """
    模式四：单组截断，每组最多 1 个系列、合计 n 条。
    模式五：全量提取，整组素材放在同一张表；超过 n 条由中台按同账号续创建下一系列。
    单个 SKU 超过 per_sku_limit 时只取最新前若干条。
    """
    sku_items = list(group_items)
    mat_lists = [_pick_richer_materials(it, sku_materials_by_any) for it in sku_items]
    if per_sku_limit and int(per_sku_limit) > 0:
        cap = int(per_sku_limit)
        mat_lists = [m[:cap] for m in mat_lists]
    n_sku = len(sku_items)
    if n_sku == 0:
        return []

    if mode == "模式四":
        if n_sku > n_group:
            stream = []
            for i in range(n_sku):
                mats = mat_lists[i]
                if not mats:
                    continue
                stream.append((i, mats[0]))
                if len(stream) >= n_group:
                    break
        else:
            stream = _round_robin_stream(mat_lists)[:n_group]
        return [_expand_stream_rows(sku_items, stream)] if stream else []

    stream = _round_robin_stream(mat_lists)
    if not stream:
        return []
    return [_expand_stream_rows(sku_items, stream)]


def _sku_group_warnings(account_groups):
    warns = []
    for acc_id, items in account_groups.items():
        filled = sum(1 for it in items if _norm_sku_group(it.get("SKU组")))
        empty = len(items) - filled
        if filled and empty:
            warns.append(f"账号 {acc_id} 有 {empty} 行未填 SKU组，已单独成一组")
    return warns


def _note_unmatched(unmatched, acc_id, sku, reason):
    unmatched.append({"账号": str(acc_id), "SKU": str(sku or ""), "原因": reason})


def _build_match_report(mode, acc_data, indexed_mats, result_tables, unmatched, warnings):
    acc_ids = []
    if acc_data is not None and len(acc_data) and "广告账号ID" in acc_data.columns:
        acc_ids = _unique_preserve([
            str(a).strip()
            for a in acc_data["广告账号ID"].tolist()
            if str(a).strip() and str(a).strip().lower() != "nan"
        ])
    return {
        "mode": mode,
        "account_count": len(acc_ids),
        "account_rows": 0 if acc_data is None else int(len(acc_data)),
        "sku_material_rows": int(indexed_mats),
        "output_tables": len(result_tables),
        "output_rows": int(sum(len(df) for df in result_tables.values())),
        "unmatched": unmatched or [],
        "warnings": warnings or [],
    }


def _mode45_table_name(mode, group_idx, series_idx, n_groups, n_series_this_group, ads_cnt, n_group):
    prefix = "多品截断" if mode == "模式四" else "多品全量"
    parts = [prefix]
    if n_groups > 1:
        parts.append(f"第{group_idx + 1}组")
    if ads_cnt > n_group and mode == "模式五":
        parts.append(f"广告数_{n_group}_可多系列")
    elif ads_cnt >= n_group:
        parts.append(f"广告数_{n_group}")
    else:
        parts.append(f"广告数_{ads_cnt}")
    return "【" + "_".join(parts) + "】"


def plan_mode3_packing(cnt_to_skus):
    """
    同素材数 SKU 作为一张底座表；每张表最多再挂载 1 个素材数更小的 SKU。
    用 DP 选择哪些素材数组整组挂到更大底座上，使最终表数最少。
    返回 [{base_count, base_skus, tail_sku, tail_count}, ...]
    """
    if not cnt_to_skus:
        return []

    items = sorted(cnt_to_skus.items(), key=lambda x: x[0], reverse=True)
    n = len(items)
    sizes = [len(skus) for _, skus in items]
    max_slots = n
    inf = 10 ** 9
    dp = [[inf] * (max_slots + 1) for _ in range(n + 1)]
    choice = [[1] * (max_slots + 1) for _ in range(n)]
    for slots in range(max_slots + 1):
        dp[n][slots] = 0

    for i in range(n - 1, -1, -1):
        sz = sizes[i]
        for slots in range(max_slots + 1):
            keep_slots = min(max_slots, slots + 1)
            best = 1 + dp[i + 1][keep_slots]
            ch = 1
            if sz <= slots:
                cand = dp[i + 1][slots - sz]
                if cand <= best:
                    best = cand
                    ch = 2
            dp[i][slots] = best
            choice[i][slots] = ch

    keep_flags = []
    slots = 0
    for i in range(n):
        sz = sizes[i]
        if choice[i][slots] == 2 and sz <= slots:
            keep_flags.append(False)
            slots -= sz
        else:
            keep_flags.append(True)
            slots = min(max_slots, slots + 1)

    kept_counts = [items[i][0] for i in range(n) if keep_flags[i]]
    dissolved = []
    for i in range(n):
        if keep_flags[i]:
            continue
        c, skus = items[i]
        for sku in skus:
            dissolved.append((c, sku))

    assignment = {}
    occupied = set()
    unassigned = []
    for tail_cnt, tail_sku in dissolved:
        candidates = [b for b in kept_counts if b > tail_cnt and b not in occupied]
        if not candidates:
            unassigned.append((tail_cnt, tail_sku))
            continue
        chosen = min(candidates)
        assignment[chosen] = (tail_sku, tail_cnt)
        occupied.add(chosen)

    leftover = defaultdict(list)
    for tail_cnt, tail_sku in unassigned:
        leftover[tail_cnt].append(tail_sku)

    plans = []
    for c in kept_counts:
        tail = assignment.get(c)
        plans.append({
            "base_count": c,
            "base_skus": list(cnt_to_skus[c]),
            "tail_sku": tail[0] if tail else None,
            "tail_count": tail[1] if tail else 0,
        })
    for c, skus in sorted(leftover.items(), reverse=True):
        plans.append({
            "base_count": c,
            "base_skus": list(skus),
            "tail_sku": None,
            "tail_count": 0,
        })
    return plans


def run_module_8_matching(acc_file_bytes, n_group=50, mode="模式一", custom_thresh=None, per_sku_limit=None):
    df_sku_raw = pd.read_excel(io.BytesIO(acc_file_bytes), sheet_name="SKU表", dtype=str)
    df_sku_raw.columns = [str(c).strip() for c in df_sku_raw.columns]
    
    mat_ver_col = [c for c in df_sku_raw.columns if "广告素材版本" in c or "素材版本" in c]
    mat_ver_name = mat_ver_col[0] if mat_ver_col else "广告素材版本"

    mat_id_col = [c for c in df_sku_raw.columns if "广告素材ID" in c or "素材ID" in c]
    mat_id_name = mat_id_col[0] if mat_id_col else "广告素材ID"
    
    df_sku_raw["SKU_KEY"] = df_sku_raw.apply(get_sku_key, axis=1)
    df_sku_raw["广告素材版本_CLEAN"] = df_sku_raw[mat_ver_name].fillna("").astype(str).str.strip()
    df_sku_raw["广告素材ID_CLEAN"] = df_sku_raw[mat_id_name].fillna("").astype(str).str.strip() if mat_id_name in df_sku_raw.columns else ""
    
    if "创建时间" in df_sku_raw.columns:
        df_sku_raw["创建时间_dt"] = pd.to_datetime(df_sku_raw["创建时间"], errors="coerce")
        df_sku_raw = df_sku_raw.sort_values(by="创建时间_dt", ascending=False)

    sku_materials_map = defaultdict(list)
    sku_materials_by_any = defaultdict(list)
    indexed_mats = 0
    for _, r in df_sku_raw.iterrows():
        k = r["SKU_KEY"]
        mat_v = r["广告素材版本_CLEAN"]
        mat_i = r["广告素材ID_CLEAN"]
        if not _filled(mat_v):
            mat_v = ""
        if not _filled(mat_i):
            mat_i = ""
        # 版本名称 / 素材ID 二选一即可，不能只认版本名称
        if not k or (not mat_v and not mat_i):
            continue
        indexed_mats += 1
        sku_materials_map[k].append((mat_v, mat_i))
        keys = []
        for col in ("虚拟SKU", "真实SKU"):
            kk = str(r.get(col, "")).strip()
            if _filled(kk) and kk not in keys:
                keys.append(kk)
        if not keys:
            keys = [k]
        for kk in keys:
            sku_materials_by_any[kk].append((mat_v, mat_i))

    df_acc_raw = pd.read_excel(io.BytesIO(acc_file_bytes), sheet_name="账号表", dtype=str)
    df_acc_raw.columns = [str(c).strip() for c in df_acc_raw.columns]
    
    raw_hint_dict = {}
    if len(df_acc_raw) > 0 and any("可不填" in str(v) for v in df_acc_raw.iloc[0].values):
        raw_hint_dict = df_acc_raw.iloc[0].to_dict()
        df_acc_data = df_acc_raw.iloc[1:].copy()
    else:
        df_acc_data = df_acc_raw.copy()
        
    df_acc_data = df_acc_data.dropna(how="all", axis=0)
    df_acc_data["SKU_KEY"] = df_acc_data.apply(get_sku_key, axis=1)

    result_tables = {}
    unmatched = []
    warnings = []

    def _report():
        return _build_match_report(mode, df_acc_data, indexed_mats, result_tables, unmatched, warnings)

    # ─────────────────────────────────────────────
    # 模式三：同素材数底座 + 每账号最多挂 1 个更小品
    # ─────────────────────────────────────────────
    if mode == "模式三":
        unique_accs = _unique_preserve([
            str(a).strip()
            for a in df_acc_data["广告账号ID"].tolist()
            if str(a).strip() and str(a).strip().lower() != "nan"
        ])
        if not unique_accs:
            return {}, raw_hint_dict, _report()

        sku_rep = {}
        for _, r in df_acc_data.iterrows():
            k = r["SKU_KEY"]
            if k and k not in sku_rep:
                sku_rep[k] = r.to_dict()

        sku_mats_dict = {}
        sku_effective_mats = {}
        for k, item in sku_rep.items():
            mats = _pick_richer_materials(item, sku_materials_by_any)[:n_group]
            sku_mats_dict[k] = mats
            sku_effective_mats[k] = len(mats)
            if not mats:
                for acc_id in unique_accs:
                    hit = df_acc_data[(df_acc_data["广告账号ID"].astype(str) == acc_id) & (df_acc_data["SKU_KEY"] == k)]
                    if len(hit):
                        _note_unmatched(unmatched, acc_id, k, "SKU表中无素材版本/素材ID")

        acc_skus = [k for k in sku_rep if sku_effective_mats.get(k, 0) > 0]
        cnt_to_skus = defaultdict(list)
        for sku in acc_skus:
            cnt_to_skus[sku_effective_mats[sku]].append(sku)

        if custom_thresh is not None and custom_thresh > 0:
            plans = _plan_mode3_by_threshold(cnt_to_skus, n_group, custom_thresh)
        else:
            plans = plan_mode3_packing(cnt_to_skus)

        acc_groups = {}
        for acc_id, grp in df_acc_data.groupby("广告账号ID", sort=False):
            acc_groups[str(acc_id).strip()] = grp

        for plan in plans:
            bc = plan["base_count"]
            base_skus = set(plan["base_skus"])
            tail_sku = plan["tail_sku"]
            tail_cnt = plan["tail_count"]
            matched_rows = []
            for acc_id in unique_accs:
                acc_df = acc_groups.get(acc_id)
                if acc_df is None or acc_df.empty:
                    continue
                base_items = acc_df[acc_df["SKU_KEY"].isin(base_skus)]
                for _, acc_item in base_items.iterrows():
                    sku_k = acc_item["SKU_KEY"]
                    mats = sku_mats_dict.get(sku_k) or []
                    if not mats:
                        continue
                    matched_rows.extend(_expand_sku_rows(acc_item, mats))
                if tail_sku:
                    tail_items = acc_df[acc_df["SKU_KEY"] == tail_sku]
                    for _, acc_item in tail_items.iterrows():
                        mats = sku_mats_dict.get(tail_sku) or []
                        if not mats:
                            continue
                        matched_rows.extend(_expand_sku_rows(acc_item, mats))

            if tail_sku:
                tbl_name = f"【中台设广告数_{bc}_允许品不足】_含底座{len(base_skus)}品+尾随1品({tail_cnt}素材)"
            else:
                tbl_name = f"【中台设广告数_{bc}】_含底座{len(base_skus)}品"
            if matched_rows:
                result_tables[tbl_name] = _order_matched_columns(pd.DataFrame(matched_rows))

        return result_tables, raw_hint_dict, _report()

    # ─────────────────────────────────────────────
    # 模式四 / 模式五：多品单系列（均分素材）
    # ─────────────────────────────────────────────
    if mode in ("模式四", "模式五"):
        account_groups = defaultdict(list)
        acc_order = []
        for _, r in df_acc_data.iterrows():
            acc_id = str(r.get("广告账号ID", "")).strip()
            if not acc_id or acc_id.lower() == "nan":
                continue
            if acc_id not in account_groups:
                acc_order.append(acc_id)
            account_groups[acc_id].append(r.to_dict())

        if not account_groups:
            return {}, raw_hint_dict, _report()

        warnings.extend(_sku_group_warnings(account_groups))
        buckets = defaultdict(list)
        series_len = defaultdict(list)
        max_groups = 0
        max_series_by_group = defaultdict(int)

        for acc_id in acc_order:
            groups = _split_acc_sku_groups(account_groups[acc_id])
            max_groups = max(max_groups, len(groups))
            for g_idx, group_items in enumerate(groups):
                usable = []
                for it in group_items:
                    mats = _pick_richer_materials(it, sku_materials_by_any)
                    if not mats:
                        _note_unmatched(unmatched, acc_id, get_sku_key(it), "SKU表中无素材版本/素材ID")
                    else:
                        usable.append(it)
                series_list = _build_multiproduct_series(
                    usable, sku_materials_by_any, mode, int(n_group),
                    per_sku_limit=per_sku_limit,
                )
                max_series_by_group[g_idx] = max(max_series_by_group[g_idx], len(series_list))
                for s_idx, rows in enumerate(series_list):
                    buckets[(g_idx, s_idx)].extend(rows)
                    series_len[(g_idx, s_idx)].append(len(rows))

        for (g_idx, s_idx) in sorted(buckets.keys()):
            rows = buckets[(g_idx, s_idx)]
            if not rows:
                continue
            ads_cnt = max(series_len[(g_idx, s_idx)]) if series_len[(g_idx, s_idx)] else 0
            tbl_name = _mode45_table_name(
                mode,
                g_idx,
                s_idx,
                max_groups,
                max_series_by_group.get(g_idx, 1),
                ads_cnt,
                int(n_group),
            )
            result_tables[tbl_name] = _order_matched_columns(pd.DataFrame(rows))

        return result_tables, raw_hint_dict, _report()

    # ─────────────────────────────────────────────
    # 模式一 / 模式二：传统 1账号1SKU 拆表
    # ─────────────────────────────────────────────
    account_groups = defaultdict(list)
    for _, r in df_acc_data.iterrows():
        acc_id = str(r.get("广告账号ID", "")).strip()
        if not acc_id or acc_id.lower() == "nan":
            continue
        account_groups[acc_id].append(r.to_dict())

    if not account_groups:
        return {}, raw_hint_dict, _report()

    max_skus_per_acc = max(len(items) for items in account_groups.values())
    table_buckets = [[] for _ in range(max_skus_per_acc)]
    for acc_id, items in account_groups.items():
        for k, item in enumerate(items):
            table_buckets[k].append(item)

    for t_idx, bucket in enumerate(table_buckets):
        matched_rows = []
        for acc_item in bucket:
            sku_k = get_sku_key(acc_item)
            acc_id = str(acc_item.get("广告账号ID", "")).strip()
            mats = _pick_richer_materials(acc_item, sku_materials_by_any)
            if mode == "模式一":
                mats = mats[:n_group]
            if not mats:
                _note_unmatched(unmatched, acc_id, sku_k, "SKU表中无素材版本/素材ID")
                continue
            matched_rows.extend(_expand_sku_rows(acc_item, mats))

        if matched_rows:
            result_tables[f"表{t_idx + 1}"] = _order_matched_columns(pd.DataFrame(matched_rows))

    return result_tables, raw_hint_dict, _report()


def _plan_mode3_by_threshold(cnt_to_skus, n_group, effective_thresh):
    """兼容旧门槛：同素材数品数 >= 门槛（或已达 n 上限）作为底座，其余整 SKU 挂到更大底座。"""
    base_counts = []
    tail_skus = []
    for c in sorted(cnt_to_skus.keys(), reverse=True):
        skus = cnt_to_skus[c]
        if len(skus) >= effective_thresh or c == n_group:
            base_counts.append(c)
        else:
            tail_skus.extend([(sku, c) for sku in skus])

    tail_skus.sort(key=lambda x: x[1], reverse=True)
    unassigned_bases = list(base_counts)
    matched_pairs = {}
    unmatched_tails = []
    for tail_sku, tail_cnt in tail_skus:
        valid_bases = [b for b in unassigned_bases if b > tail_cnt]
        if valid_bases:
            chosen_base = min(valid_bases)
            matched_pairs[chosen_base] = (tail_sku, tail_cnt)
            unassigned_bases.remove(chosen_base)
        else:
            unmatched_tails.append((tail_sku, tail_cnt))

    plans = []
    for bc in base_counts:
        tail = matched_pairs.get(bc)
        plans.append({
            "base_count": bc,
            "base_skus": list(cnt_to_skus[bc]),
            "tail_sku": tail[0] if tail else None,
            "tail_count": tail[1] if tail else 0,
        })
    leftover = defaultdict(list)
    for rem_sku, rem_cnt in unmatched_tails:
        leftover[rem_cnt].append(rem_sku)
    for rem_cnt, skus in sorted(leftover.items(), reverse=True):
        plans.append({
            "base_count": rem_cnt,
            "base_skus": list(skus),
            "tail_sku": None,
            "tail_count": 0,
        })
    return plans


# ─────────────────────────────────────────────
# 3. Streamlit 页面与操作交互
# ─────────────────────────────────────────────

def _render_match_report(rep):
    if not rep:
        return
    st.markdown("##### 📋 匹配报告")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("账号数", rep.get("account_count", 0))
    c2.metric("账号表行数", rep.get("account_rows", 0))
    c3.metric("SKU表素材条数", rep.get("sku_material_rows", 0))
    c4.metric("导出数据行", rep.get("output_rows", 0))
    for w in rep.get("warnings") or []:
        st.warning(w)
    unmatched = rep.get("unmatched") or []
    if unmatched:
        st.warning(f"有 **{len(unmatched)}** 条账号-SKU 未匹配到素材，未写入结果表。")
        st.dataframe(pd.DataFrame(unmatched), use_container_width=True, hide_index=True)
    elif not rep.get("sku_material_rows"):
        st.warning("SKU 表没有可用的素材版本或素材ID（两列至少填一列）。")


def run(params):
    st.subheader("🔗 模块八：着陆页导入与素材智能匹配")

    with st.expander("💡 点击查看：运行逻辑与模式说明", expanded=False):
        st.markdown("""
**模块八：着陆页导入与素材智能匹配**

- **适用场景**：着陆页链接（或着陆页版本名称）批量生成，并按 SKU 从素材库提取最新广告素材。适用于运营已配置好广告账号与对应着陆页的 SKU 清单。
  - **单品多素材**（模式一 / 二 / 三）：一个广告系列里只有 1 个 SKU，满足中台「单张表格中 1 个账号只允许绑定 1 个 SKU」。
  - **多品多素材**（模式四 / 五）：一个系列里有多个 SKU、共用一组素材；用账号表 **SKU组** 把同一账号下的品拆成多个系列。

- **模式选择与如何使用**

  **1. 准备与填写数据**（专用文件，两个工作表）
  - **账号表**：广告账号 ID、主页 ID、像素 ID、真实/虚拟 SKU、国家、着陆页链接（或版本名称）、出价/竞价、系列标注、**SKU组** 等。同一账号可多行绑定不同 SKU。
  - **SKU表**：真实/虚拟 SKU、SKU中文名称、国家、广告素材版本、广告素材 ID、创建时间等。
  - 系统具备跨列容错（真实/虚拟 SKU 填错列也能匹配）；同一 SKU 同时有真实、虚拟两份素材时，取**条数更多**的那份。
  - **SKU组**（紧挨系列标注后）：仅模式四 / 五使用，填 `1`、`2`、`3`… 同一账号下相同数字进入**同一个系列**；整列留空 = 该账号全部 SKU 为一组。不要用「系列标注」来分组。`SKU组` 只用于拆组，**导出时不带出**。

  **2. 配置参数与提取模式**（先选场景，再选提取方式）
  - **每组素材上限 n**：默认 50，可改。模式一 / 三是单 SKU 截断上限；模式四是该组合计条数上限；模式五是导品时的「广告数」，超了不拆文件。
  - **单个SKU上限**：仅多品显示，默认 20。每个 SKU 只取最新的前若干条。
  - **导入填入列**：结果里「广告素材版本名称」「广告素材ID」两列都保留，**内容只填选中的那一列**，另一列留空，导品时不必再手动删列。

  **单品多素材**
  - **模式一 · 单组截断**：每个账号-SKU 只取按创建时间倒序的最新前 n 条，超过 n 的旧素材舍弃。一张表里同一账号只能有 1 个 SKU。
  - **模式二 · 全量提取**：该 SKU 有多少素材导多少。一张表里同一账号只能有 1 个 SKU。
  - **模式三 · 同素材数合表**：仍是「一账号一品一系列」。相同素材条数的 SKU 放进同一张**文件**，少出几张表；每张表每个账号最多再挂 **1 个素材更少的 SKU**。导品时中台广告数 = 文件名里的素材数，允许品不足。

  **多品多素材**
  - **模式四 · 单组截断**：同一 `SKU组` 的多个 SKU 打进**同一个系列**，素材轮流均分，合计最多 n 条。SKU 个数大于 n 时，按表中顺序取前 n 个 SKU、每个 1 条最新素材。
  - **模式五 · 全量提取**：同一 `SKU组` 的素材全部放在**同一张表**（单个 SKU 仍受「单个SKU上限」约束）。同一账号超过 n 条时**不拆文件**；导品设广告数 = n 并勾选继续创建，中台会在同账号下开下一系列。
  - 一个账号有多组：`SKU组` 填 1、2、3… **一组 = 一个系列 / 一个文件**。组与组不会混进同一系列，也不走模式三那种「按素材数拼表」。

  **3. 核心处理与分表原则**
  - **动态感知 G 列**：若为产品链接（含 `.com`、`http`、`/products/` 等），表头保持「着陆页链接」；若为版本名称（如 优化组版本-xxx、默认版本），表头自动改为「着陆页版本名称」，并同步第 2 行红字提示。
  - **最新素材优先**：SKU 表按创建时间降序，优先分配最新产出。
  - **单品拆表（模式一 / 二）**：同一张结果表内只允许同一个账号对应同一个 SKU。同一账号关联多个 SKU 时，表 1 分各账号的第 1 个 SKU，表 2 分第 2 个……
  - **模式三拼文件**：按素材条数打包，不是按 SKU 组；中台仍按「一账号一品」建系列。
  - **多品拆文件（模式四 / 五）**：按账号的 `SKU组` 对齐输出（第 1 组、第 2 组…）。换账号即新系列，同一文件里可以有多个账号。

  **4. 一键匹配与标准导出**
  - 点击「开始匹配并拆表」：单表直接 `.xlsx`，多子表自动打 `.zip`。
  - 每个表格按《着陆页链接导入模板》规范生成（第 1 行标准表头、第 2 行格式说明与红字提示，第 3 行起为业务数据），可直接后台批量导入。
  - 下方匹配报告会列出账号数、素材行数、未匹配 SKU 及 `SKU组` 填写警告。

  **5. 注意事项**
  - 导出品时请用界面「导入填入列」选定填 **广告素材ID** 或 **广告素材版本名称**，不要两列都填；未选中的那一列会留空，不必再手动清数据。
  - 账号表 G 列如果填的是着陆页库里的版本名称（不是商品链接），导出时表头会自动从「着陆页链接」改成「着陆页版本名称」，不用自己改表头。单元格内容原样带出。
  - 模式三导品：广告数看文件名「中台设广告数_xx」，允许该账号在这张表里品不足。
  - 模式五导品：广告数填 n（默认 50）；同一账号素材超过 n 时，在中台继续创建下一系列即可，不必拆多个 Excel。
  - 模式四 / 五请填写 **SKU组**；同一账号有的行填了、有的行空，空行会单独成一组并在匹配报告里提示。
        """)

    st.markdown("##### ⚙️ 1. 匹配规则与参数配置")
    c_scene, c_mode, c_keep = st.columns(3, gap="medium")
    with c_scene:
        scene = st.radio(
            "使用场景",
            ["单品多素材", "多品多素材"],
            help="单品：一个广告系列里只有 1 个 SKU。多品：一个系列里有多个 SKU，共用一组素材。",
        )
    with c_mode:
        if scene == "单品多素材":
            run_mode = st.radio(
                "提取方式",
                [
                    "模式一：单组截断",
                    "模式二：全量提取",
                    "模式三：同素材数合表",
                ],
                index=0,
                help="一张表里同一账号只有 1 个 SKU。把素材条数相同的品放进同一文件，少拆几张表；每账号还可再挂 1 个素材更少的品。",
            )
        else:
            run_mode = st.radio(
                "提取方式",
                [
                    "模式四：单组截断",
                    "模式五：全量提取",
                ],
                index=0,
                help="同一 SKU组 的多个 SKU 进入同一个系列并均分素材。模式四合计最多 n 条；模式五全量放同一张表，超 n 导品时同账号续创建下一系列。多组请在账号表「SKU组」填 1、2、3。",
            )
    with c_keep:
        keep_material_label = st.radio(
            "导入填入列",
            ["广告素材ID（推荐）", "广告素材版本名称"],
            index=0,
            help="两列都导出，内容只填选中的那一列。",
        )
        keep_material_col = "广告素材ID" if "广告素材ID" in keep_material_label else "广告素材版本名称"

    if "模式一" in run_mode:
        mode_key = "模式一"
    elif "模式二" in run_mode:
        mode_key = "模式二"
    elif "模式三" in run_mode:
        mode_key = "模式三"
    elif "模式四" in run_mode:
        mode_key = "模式四"
    else:
        mode_key = "模式五"

    n_help = (
        "模式四：该组合计最多 n 条；模式五：导品广告数，超 n 同账号续创建下一系列"
        if scene == "多品多素材"
        else "模式一：每个 SKU 最多 n 条；模式三：先按 n 截断，再把相同条数的品合成一张表"
    )
    if scene == "多品多素材":
        c_up, c_n, c_cap = st.columns([1.4, 0.8, 0.8], gap="medium")
    else:
        c_up, c_n = st.columns(2, gap="medium")
        c_cap = None
    with c_up:
        up = st.file_uploader(
            "上传需求文件（账号表 + SKU表）",
            type=["xlsx"],
            key="m8_up",
            help="Excel 需包含「账号表」「SKU表」两个 sheet。",
        )
    with c_n:
        n_group = st.number_input(
            "每组素材数量上限 (n)",
            min_value=1,
            max_value=500,
            value=50,
            step=1,
            help=n_help,
        )
    per_sku_limit = None
    if c_cap is not None:
        with c_cap:
            per_sku_limit = st.number_input(
                "单个SKU上限",
                min_value=1,
                max_value=500,
                value=20,
                step=1,
                help="每个 SKU 最多导入这么多条素材（取最新）。超过则截断。默认 20。",
                key="m8_per_sku",
            )
        per_sku_limit = int(per_sku_limit)

    cache_sig = (up.name if up else None, keep_material_col, mode_key, int(n_group), per_sku_limit)
    if st.session_state.get("m8_last_sig") != cache_sig:
        st.session_state["m8_last_sig"] = cache_sig
        st.session_state.pop("m8_download_bytes", None)
        st.session_state.pop("m8_table_counts", None)
        st.session_state.pop("m8_is_zip", None)
        st.session_state.pop("m8_single_filename", None)
        st.session_state.pop("m8_keep_col", None)
        st.session_state.pop("m8_report", None)

    c_go, c_dl = st.columns(2, gap="medium")
    with c_go:
        start_btn = st.button(
            "🚀 开始匹配并拆表",
            key="m8_btn",
            use_container_width=True,
            disabled=not bool(up),
        )

    match_error = None
    empty_match = False
    if start_btn and up:
        file_bytes = up.getvalue()
        with st.spinner("正在匹配素材并按规则拆表..."):
            try:
                result_tables, raw_hint_dict, match_report = run_module_8_matching(
                    file_bytes,
                    n_group=int(n_group),
                    mode=mode_key,
                    per_sku_limit=int(per_sku_limit) if per_sku_limit else None,
                )
            except Exception as e:
                match_error = str(e)
                result_tables, match_report = None, None

            if match_error:
                pass
            elif not result_tables:
                st.session_state["m8_report"] = match_report
                empty_match = True
            else:
                file_prefix = params.get("prefix", "项目_")
                table_counts_info = {tbl: len(df) for tbl, df in result_tables.items()}

                if len(result_tables) == 1:
                    single_tbl = list(result_tables.values())[0]
                    first_tbl_name = list(result_tables.keys())[0]
                    st.session_state["m8_download_bytes"] = write_landing_page_excel(
                        single_tbl, raw_hint_dict, keep_material_col=keep_material_col
                    )
                    st.session_state["m8_is_zip"] = False
                    st.session_state["m8_single_filename"] = f"{file_prefix}着陆页导入_{first_tbl_name}.xlsx"
                else:
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                        for tbl_name, df_tbl in result_tables.items():
                            excel_bytes = write_landing_page_excel(
                                df_tbl, raw_hint_dict, keep_material_col=keep_material_col
                            )
                            file_name = f"{file_prefix}着陆页导入_{tbl_name}.xlsx"
                            zf.writestr(file_name, excel_bytes)
                    st.session_state["m8_download_bytes"] = zip_buffer.getvalue()
                    st.session_state["m8_is_zip"] = True

                st.session_state["m8_table_counts"] = table_counts_info
                st.session_state["m8_keep_col"] = keep_material_col
                st.session_state["m8_report"] = match_report
        gc.collect()

    file_prefix = params.get("prefix", "项目_")
    table_counts = st.session_state.get("m8_table_counts", {})
    total_rows = sum(table_counts.values())
    is_zip = st.session_state.get("m8_is_zip", False)
    has_dl = "m8_download_bytes" in st.session_state
    if any("中台设广告数" in n for n in table_counts):
        zip_name = f"{file_prefix}着陆页导入_底座挂载结果.zip"
    elif any("多品" in n for n in table_counts):
        zip_name = f"{file_prefix}着陆页导入_多品单系列结果.zip"
    else:
        zip_name = f"{file_prefix}着陆页导入_拆表匹配结果.zip"
    single_name = st.session_state.get("m8_single_filename", f"{file_prefix}着陆页导入_表1.xlsx")
    if has_dl:
        dl_label = (
            f"💾 导出结果 ({len(table_counts)} 表 / {total_rows} 行)"
            if is_zip
            else f"💾 导出结果 ({total_rows} 行)"
        )
        dl_name = zip_name if is_zip else single_name
        dl_mime = "application/zip" if is_zip else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        dl_data = st.session_state["m8_download_bytes"]
    else:
        dl_label = "💾 导出结果"
        dl_name = "result.xlsx"
        dl_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        dl_data = b""

    with c_dl:
        st.download_button(
            dl_label,
            data=dl_data,
            file_name=dl_name,
            mime=dl_mime,
            use_container_width=True,
            disabled=not has_dl,
            key="dl_m8_btn",
        )

    if match_error:
        st.error(f"⚠️ 处理失败，请检查文件格式是否正确：{match_error}")
        return
    if empty_match:
        st.warning("⚠️ 未匹配到有效导出行，请查看下方匹配报告。")
        _render_match_report(st.session_state.get("m8_report"))
        return

    if has_dl:
        if any("中台设广告数" in n for n in table_counts):
            st.success(
                f"🎉 处理完成！模式三共生成 **{len(table_counts)}** 个表格"
                f"（同素材数一表，每账号最多挂 1 个更小品；中台设广告数=文件名中的素材数，允许品不足）。"
            )
        elif any("多品" in n for n in table_counts):
            st.success(
                f"🎉 处理完成！多品单系列共生成 **{len(table_counts)}** 个表格"
                f"（同一账号多 SKU 进同一系列，素材轮流均分）。"
            )
        else:
            st.success(f"🎉 匹配完成！已成功按「1账号1SKU」拆分为 **{len(table_counts)}** 个表格。")
        keep_col_used = st.session_state.get("m8_keep_col", "广告素材ID")
        other_col = "广告素材版本名称" if keep_col_used == "广告素材ID" else "广告素材ID"
        st.caption(f"导出两列都在：已将匹配结果写入 **{keep_col_used}**，**{other_col}** 留空。")
        _render_match_report(st.session_state.get("m8_report"))
        with st.expander("📊 查看各子表详情与中台建任务指引", expanded=False):
            for tbl_name, cnt in table_counts.items():
                st.caption(f"• **{tbl_name}**：共包含 **{cnt}** 条广告素材导入记录")