import streamlit as st
import pandas as pd
import gc
from collections import defaultdict
from utils import write_excel_final, read_uploaded_excel, safe_int


# ─────────────────────────────────────────────
# 核心算法（纯函数，可单测、可被 UI 调用）
# ─────────────────────────────────────────────

def _sku_id(item):
    """有效 SKU 标识：优先虚拟SKU，否则真实SKU。"""
    v_code = str(item.get("虚拟SKU", "")).strip()
    r_code = str(item.get("真实SKU", "")).strip()
    if v_code and v_code.lower() != "nan":
        return v_code
    if r_code and r_code.lower() != "nan":
        return r_code
    return ""


def make_combo_key(item):
    """组合唯一键：(SKU, 着陆页版本, 广告素材版本)。"""
    sku = _sku_id(item)
    lp = str(item.get("着陆页版本名称", "")).strip()
    mat = str(item.get("广告素材版本名称", "")).strip()
    return (sku, lp, mat)


def build_depth_matrix(pool):
    """
    将资源池按 SKU 分组，保持首次出现顺序。
    返回: sku_group[sku] = [item1, item2, ...], ordered_skus
    """
    sku_group = defaultdict(list)
    ordered_skus = []
    for item in pool:
        sku = _sku_id(item)
        if not sku:
            continue
        if sku not in sku_group:
            ordered_skus.append(sku)
        sku_group[sku].append(item)
    return sku_group, ordered_skus


def _pick_item(item, selected_items, acc_used_combos, acc_used_skus, globally_consumed_combos, sku):
    """选中一条素材，同步更新账号组合锁 / SKU 集合 / 全局组合锁。"""
    selected_items.append(item.copy())
    combo_k = make_combo_key(item)
    acc_used_combos.add(combo_k)
    acc_used_skus.add(sku)
    globally_consumed_combos.add(combo_k)


def _flatten_pool_items(sku_group, ordered_skus):
    """按 SKU 顺序展开全部素材变体，供阶段 3 顺延复用。"""
    flat = []
    for sku in ordered_skus:
        for item in sku_group[sku]:
            flat.append(item)
    return flat


def allocate_main_category(
    raw_pool,
    target_count,
    globally_consumed_combos,
    strict_unique_sku=True,
    phase3_pointer=0,
):
    """
    主品类阶段 1~3。

    strict_unique_sku=True  (模式一)：账号内 SKU 绝对不重复
    strict_unique_sku=False (模式二)：允许同 SKU 不同素材，禁止同 SKU 同素材

    阶段 1：全局未用的各 SKU 第 1 变体（广度 + 顺延）
    阶段 2：全局未用变体
    阶段 3：全库第一轮耗尽后，按 phase3_pointer **跨账号顺延**复用

    返回 (selected_items, acc_used_combos, acc_used_skus, phase3_pointer)
    """
    selected_items = []
    acc_used_skus = set()
    acc_used_combos = set()

    sku_group, ordered_skus = build_depth_matrix(raw_pool)
    if not ordered_skus:
        return selected_items, acc_used_combos, acc_used_skus, phase3_pointer

    max_depth = max(len(items) for items in sku_group.values())

    def _sku_blocked(sku):
        return strict_unique_sku and sku in acc_used_skus

    # ── 阶段 1：各 SKU 第 1 变体 + 全局未用 + 账号内 SKU 互斥 ──
    for sku in ordered_skus:
        if len(selected_items) >= target_count:
            break
        if sku in acc_used_skus:
            continue
        items = sku_group[sku]
        if not items:
            continue
        item = items[0]
        combo_k = make_combo_key(item)
        if combo_k in globally_consumed_combos or combo_k in acc_used_combos:
            continue
        _pick_item(item, selected_items, acc_used_combos, acc_used_skus, globally_consumed_combos, sku)

    # ── 阶段 2：全局未用变体 ──
    if len(selected_items) < target_count:
        for d in range(max_depth):
            if len(selected_items) >= target_count:
                break
            for sku in ordered_skus:
                if len(selected_items) >= target_count:
                    break
                if _sku_blocked(sku):
                    continue
                items = sku_group[sku]
                if d >= len(items):
                    continue
                item = items[d]
                combo_k = make_combo_key(item)
                if combo_k in globally_consumed_combos or combo_k in acc_used_combos:
                    continue
                _pick_item(item, selected_items, acc_used_combos, acc_used_skus, globally_consumed_combos, sku)

    # ── 阶段 3：跨账号顺延复用（指针接续，不从 0 重开）──
    if len(selected_items) < target_count:
        flat = _flatten_pool_items(sku_group, ordered_skus)
        n = len(flat)
        if n > 0:
            start = phase3_pointer % n
            last_pick_next = start
            picked_any = False
            for step in range(n):
                if len(selected_items) >= target_count:
                    break
                idx = (start + step) % n
                item = flat[idx]
                sku = _sku_id(item)
                if _sku_blocked(sku):
                    continue
                combo_k = make_combo_key(item)
                if combo_k in acc_used_combos:
                    continue
                _pick_item(item, selected_items, acc_used_combos, acc_used_skus, globally_consumed_combos, sku)
                last_pick_next = (idx + 1) % n
                picked_any = True
            if picked_any:
                phase3_pointer = last_pick_next

    return selected_items, acc_used_combos, acc_used_skus, phase3_pointer


def allocate_fallback_round_robin(
    fb_raw_pool,
    target_count,
    selected_items,
    acc_used_combos,
    acc_used_skus,
    globally_consumed_combos,
    global_fallback_pointer,
    strict_unique_sku=False,
):
    """
    模式二 阶段 4：后备品类环形指针均摊补齐。
    """
    fb_sku_group, fb_ordered_skus = build_depth_matrix(fb_raw_pool)
    num_fb = len(fb_ordered_skus)
    if num_fb == 0:
        return selected_items, acc_used_combos, acc_used_skus, global_fallback_pointer

    def _try_pick_from_sku(sku, mode):
        if strict_unique_sku and sku in acc_used_skus:
            return False
        items = fb_sku_group[sku]
        if not items:
            return False

        if mode == "first_free":
            item = items[0]
            combo_k = make_combo_key(item)
            if combo_k in globally_consumed_combos or combo_k in acc_used_combos:
                return False
            _pick_item(item, selected_items, acc_used_combos, acc_used_skus, globally_consumed_combos, sku)
            return True

        if mode == "any_free":
            for item in items:
                combo_k = make_combo_key(item)
                if combo_k in globally_consumed_combos or combo_k in acc_used_combos:
                    continue
                _pick_item(item, selected_items, acc_used_combos, acc_used_skus, globally_consumed_combos, sku)
                return True
            return False

        for item in items:
            combo_k = make_combo_key(item)
            if combo_k in acc_used_combos:
                continue
            _pick_item(item, selected_items, acc_used_combos, acc_used_skus, globally_consumed_combos, sku)
            return True
        return False

    def _round_robin(mode):
        nonlocal global_fallback_pointer
        total_slots = sum(len(v) for v in fb_sku_group.values()) or num_fb
        no_progress_streak = 0
        while len(selected_items) < target_count and no_progress_streak < max(num_fb, total_slots):
            sku = fb_ordered_skus[global_fallback_pointer % num_fb]
            found = _try_pick_from_sku(sku, mode)
            global_fallback_pointer = (global_fallback_pointer + 1) % num_fb
            if found:
                no_progress_streak = 0
            else:
                no_progress_streak += 1

    _round_robin("first_free")
    if len(selected_items) < target_count:
        _round_robin("any_free")
    if len(selected_items) < target_count:
        _round_robin("reuse")

    return selected_items, acc_used_combos, acc_used_skus, global_fallback_pointer


def run_allocation(acc_rows, sku_rows, run_mode, fallback_category=""):
    """
    对全部账号执行匹配
    """
    category_sku_map = defaultdict(list)
    for r in sku_rows:
        cat = str(r.get("商品分类", "")).strip()
        if cat:
            category_sku_map[cat].append(r)

    final_rows = []
    warning_logs = []
    gap_summary_list = []

    globally_consumed_combos = set()
    global_fallback_pointer = 0
    phase3_pointers = defaultdict(int)
    is_mode_two = "模式二" in run_mode
    strict_unique_sku = not is_mode_two

    for idx, acc_dict in enumerate(acc_rows):
        excel_line = idx + 2
        acc_id = str(acc_dict.get("账号ID", "")).strip()
        category = str(acc_dict.get("品类", "")).strip()
        target_sku_count = safe_int(acc_dict.get("需要组合的SKU数量"), default=1)
        zhuye = str(acc_dict.get("主页ID", "")).strip()
        xiangsu = str(acc_dict.get("像素ID", "")).strip()
        xilie = str(acc_dict.get("系列标注", "")).strip()

        if not category:
            warning_logs.append(f"⚠️ 行 {excel_line} [账号:{acc_id}] 未填写品类，跳过。")
            continue

        raw_pool = category_sku_map.get(category, [])
        
        # 🌟 修改核心：即便主品类没有找到 SKU，在模式二下也不直接 continue！
        if raw_pool:
            selected_items, acc_used_combos, acc_used_skus, phase3_pointers[category] = (
                allocate_main_category(
                    raw_pool,
                    target_sku_count,
                    globally_consumed_combos,
                    strict_unique_sku=strict_unique_sku,
                    phase3_pointer=phase3_pointers[category],
                )
            )
        else:
            selected_items = []
            acc_used_combos = set()
            acc_used_skus = set()
            warning_logs.append(f"❓ 行 {excel_line} [账号:{acc_id}] 主品类【{category}】未找到对应SKU。")

        # 🌟 模式二：当已拿到数量不够（包括主品类选出 0 个）且设置了后备品类，自动用后备补齐
        if is_mode_two and len(selected_items) < target_sku_count and fallback_category:
            fb_raw_pool = category_sku_map.get(fallback_category, [])
            if not fb_raw_pool:
                warning_logs.append(
                    f"⚠️ 行 {excel_line} [账号:{acc_id}] 后备品类【{fallback_category}】未找到对应SKU，无法补齐。"
                )
            else:
                selected_items, acc_used_combos, acc_used_skus, global_fallback_pointer = (
                    allocate_fallback_round_robin(
                        fb_raw_pool,
                        target_sku_count,
                        selected_items,
                        acc_used_combos,
                        acc_used_skus,
                        globally_consumed_combos,
                        global_fallback_pointer,
                        strict_unique_sku=False,
                    )
                )
        elif is_mode_two and len(selected_items) < target_sku_count and not fallback_category:
            warning_logs.append(
                f"ℹ️ 行 {excel_line} [账号:{acc_id}] 主品类不足且未设置后备补齐品类，按缺口打住。"
            )

        actual_count = len(selected_items)
        gap = target_sku_count - actual_count

        if gap == 0:
            note_text = "完全匹配"
        else:
            note_text = f"⚠️ 缺{gap}个 ({actual_count}/{target_sku_count})"
            gap_summary_list.append({
                "广告账号ID": acc_id,
                "品类": category,
                "目标数量": target_sku_count,
                "实际获取": actual_count,
                "缺口数量": gap,
            })

        for item in selected_items:
            item["广告账号ID"] = acc_id
            item["主页ID"] = zhuye
            item["像素ID"] = xiangsu
            item["系列标注"] = xilie
            item["备注"] = note_text
            final_rows.append(item)

    return final_rows, warning_logs, gap_summary_list


# ─────────────────────────────────────────────
# Streamlit UI
# ─────────────────────────────────────────────

_LOGIC_DOC = """
按账号顺序分素材，账号间顺延、不从同一起点重复拿。

- **模式一**：只用主品类，不够就停；账号内同一 SKU 只出一次  
- **模式二**：主品类用完可用后备品类补齐；账号内允许同 SKU 不同素材，禁止同 SKU 同素材  
"""


def run(params):
    st.subheader("📦 模块七：账号品类匹配与智能SKU组合")

    with st.expander("💡 点击查看：运行逻辑说明", expanded=False):
        st.markdown(_LOGIC_DOC)

    st.markdown("##### ⚙️ 1. 匹配模式与文件上传")
    c_left, c_right = st.columns([1, 1], gap="large")

    with c_left:
        run_mode = st.radio(
            "选择补齐与填充策略",
            [
                "模式一：有多少是多少 (广度优先，不补齐)",
                "模式二：不足补齐 (模式一逻辑 + 后备品类均摊)",
            ],
            help="模式一：账号内SKU绝对不重复、宁缺毋滥；模式二：允许同SKU不同素材，主品类用干后后备环形补齐（全局消耗）",
        )

        fallback_category = ""
        if "模式二" in run_mode:
            fallback_category = st.text_input(
                "➕ 后备补齐品类 (选填)",
                value="",
                placeholder="例如：家庭建材 / 女鞋 (主品类素材耗尽时触发)",
                help="主品类彻底无货后，按环形指针从该品类均摊补齐；已全局消耗的素材不会重复优先分配。",
            ).strip()

    with c_right:
        up = st.file_uploader("上传模块七专用模板 (.xlsx)", type=["xlsx"], key="m7_up")

    if "m7_last_up" not in st.session_state or st.session_state["m7_last_up"] != (up.name if up else None):
        st.session_state["m7_last_up"] = up.name if up else None
        st.session_state.pop("m7_excel_bytes", None)
        st.session_state.pop("m7_total_rows", None)
        st.session_state.pop("m7_gap_summary", None)
        st.session_state.pop("m7_warning_logs", None)

    if up:
        st.write("")
        st.markdown("##### 🚀 2. 执行匹配与结果导出")
        btn_col1, btn_col2 = st.columns([1, 1], gap="large")

        with btn_col1:
            start_btn = st.button("🚀 开始匹配并生成", key="m7_btn", use_container_width=True)

        if start_btn:
            file_bytes = up.getvalue()

            try:
                df_acc_raw = read_uploaded_excel(file_bytes, sheet_name="账号表")
                df_sku_raw = read_uploaded_excel(file_bytes, sheet_name="SKU表")
            except Exception:
                st.error("⚠️ 读取失败：请确保 Excel 中包含【账号表】和【SKU表】！")
                return

            acc_rows = [r for r in df_acc_raw.to_dict("records") if any(str(v).strip() for v in r.values())]
            sku_rows = [r for r in df_sku_raw.to_dict("records") if any(str(v).strip() for v in r.values())]

            if not acc_rows or not sku_rows:
                st.warning("⚠️ 检测到账号表或SKU表中无有效数据！")
                return

            original_sku_columns = list(df_sku_raw.columns)

            final_rows, warning_logs, gap_summary_list = run_allocation(
                acc_rows, sku_rows, run_mode, fallback_category
            )

            if final_rows:
                final_df = pd.DataFrame(final_rows)

                preferred_order = ["广告账号ID", "主页ID", "像素ID"]
                remaining_cols = [
                    c for c in original_sku_columns
                    if c not in preferred_order and c in final_df.columns
                ]

                all_cols = preferred_order + remaining_cols
                
                if "系列标注" in final_df.columns:
                    if "系列标注" in all_cols:
                        all_cols.remove("系列标注")
                    
                    if "出价/竞价" in all_cols:
                        idx = all_cols.index("出价/竞价") + 1
                        all_cols.insert(idx, "系列标注")
                    else:
                        all_cols.insert(len(preferred_order), "系列标注")

                for tail_col in ["注意事项", "备注"]:
                    if tail_col in all_cols:
                        all_cols.remove(tail_col)
                        all_cols.append(tail_col)

                final_cols = [c for c in all_cols if c in final_df.columns]
                other_cols = [c for c in final_df.columns if c not in final_cols]
                final_df = final_df.reindex(columns=final_cols + other_cols)

                xlsx_data = write_excel_final(final_df, "品类匹配结果", params, color_by="广告账号ID")
                st.session_state["m7_excel_bytes"] = xlsx_data
                st.session_state["m7_total_rows"] = len(final_rows)
                st.session_state["m7_gap_summary"] = gap_summary_list
                st.session_state["m7_warning_logs"] = warning_logs
            else:
                st.session_state.pop("m7_excel_bytes", None)
                st.session_state["m7_total_rows"] = 0
                st.session_state["m7_gap_summary"] = gap_summary_list
                st.session_state["m7_warning_logs"] = warning_logs
                if warning_logs:
                    for w in warning_logs:
                        st.warning(w)
                else:
                    st.warning("⚠️ 未生成任何匹配结果，请检查账号表与SKU表数据。")

            possible_vars = ["df_acc_raw", "df_sku_raw", "acc_rows", "sku_rows", "final_rows", "final_df"]
            for var in possible_vars:
                if var in locals():
                    del locals()[var]
            gc.collect()

        if "m7_excel_bytes" in st.session_state:
            file_prefix = params.get("prefix", "项目_")

            with btn_col2:
                st.download_button(
                    f"💾 下载：{file_prefix}结果 ({st.session_state['m7_total_rows']}行)",
                    data=st.session_state["m7_excel_bytes"],
                    file_name=f"{file_prefix}模块七组合结果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="dl_m7_btn",
                )

            st.write("")
            warning_logs = st.session_state.get("m7_warning_logs", [])
            if warning_logs:
                with st.expander(f"📋 运行提示（{len(warning_logs)} 条）", expanded=False):
                    for w in warning_logs:
                        st.caption(w)

            gap_list = st.session_state.get("m7_gap_summary", [])
            if gap_list:
                st.warning(
                    f"⚠️ 检测到 **{len(gap_list)}** 个账号存在素材缺口，详情请在导出表格的【备注】列核对！"
                )
                with st.expander("🔍 点击展开查看缺口明细概要"):
                    for gap_item in gap_list:
                        st.caption(
                            f"• **{gap_item['广告账号ID']}** [{gap_item['品类']}]："
                            f"目标 {gap_item['目标数量']} ➔ 实际 {gap_item['实际获取']} "
                            f"(缺 {gap_item['缺口数量']})"
                        )
            else:
                st.success("🎉 所有账号素材均 100% 完全匹配！")