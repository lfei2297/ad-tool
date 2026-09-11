"""
模块七分配算法测试用例

核心硬约束（按业务确认）：
  1. 模式一：账号内 SKU 绝对不重复（同一 SKU 最多 1 条）
  2. 模式二：账号内禁止同 SKU 同素材，允许同 SKU 不同素材
  3. 全局组合消耗（他号用过的 SKU+着陆页+素材，下号不再优先拿到同一条）
  4. 模式二后备：环形指针 + 全局消耗（账号1 拿 ABC 素材1 → 账号2 拿 AB 素材2 或 DEF）
"""
import sys
import os
from collections import defaultdict

ROOT = os.path.dirname(os.path.abspath(__file__))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from modules.module_7 import (
    run_allocation,
    make_combo_key,
    allocate_fallback_round_robin,
)


def make_sku(sku, mat, category, lp="LP1", real_sku=""):
    return {
        "真实SKU": real_sku,
        "虚拟SKU": sku,
        "国家": "美国",
        "着陆页版本名称": lp,
        "广告素材版本名称": mat,
        "商品分类": category,
        "出价/竞价": "",
    }


def make_acc(acc_id, category, need, home="H1", pixel="P1"):
    return {
        "账号ID": acc_id,
        "品类": category,
        "需要组合的SKU数量": need,
        "主页ID": home,
        "像素ID": pixel,
    }


def summarize(final_rows):
    order = []
    buckets = defaultdict(list)
    notes = {}
    for r in final_rows:
        aid = r["广告账号ID"]
        if aid not in notes:
            order.append(aid)
            notes[aid] = r.get("备注", "")
        buckets[aid].append((
            str(r.get("虚拟SKU") or r.get("真实SKU") or "").strip(),
            str(r.get("广告素材版本名称", "")).strip(),
            str(r.get("商品分类", "")).strip(),
        ))
    return [(aid, buckets[aid], notes[aid]) for aid in order]


def assert_unique_skus(items, acc_label):
    """模式一：SKU 绝对不重复。"""
    skus = [x[0] for x in items]
    assert len(skus) == len(set(skus)), f"{acc_label} 出现重复 SKU: {skus}"


def assert_unique_combos(items, acc_label):
    """模式二：同 SKU 同素材不可重复；同 SKU 不同素材允许。"""
    keys = [(x[0], x[1]) for x in items]
    assert len(keys) == len(set(keys)), f"{acc_label} 出现重复组合(SKU+素材): {keys}"


def print_result(title, final_rows, gaps, warnings):
    print(f"\n{'=' * 60}")
    print(f"  {title}")
    print(f"{'=' * 60}")
    for aid, items, note in summarize(final_rows):
        print(f"\n▶ {aid}  共 {len(items)} 行  | 备注: {note}")
        for i, (sku, mat, cat) in enumerate(items, 1):
            print(f"  [{i:02d}] {cat} | {sku} | {mat}")
    if gaps:
        print("\n缺口摘要:")
        for g in gaps:
            print(f"  - {g['广告账号ID']}: 目标{g['目标数量']} 实际{g['实际获取']} 缺{g['缺口数量']}")
    if warnings:
        print("\n警告:")
        for w in warnings:
            print(f"  {w}")


def _sku_from(item):
    return str(item.get("虚拟SKU") or item.get("真实SKU") or "").strip()


# ─────────────────────────────────────────────
# 用例 1：模式一 阶段1 — 广度优先 + 顺延
# ─────────────────────────────────────────────

def test_mode1_phase1_breadth_and_sequential():
    acc_rows = [
        make_acc("账号1", "女装", 3),
        make_acc("账号2", "女装", 2),
    ]
    sku_rows = [
        make_sku("SKU1", "素材1", "女装"),
        make_sku("SKU1", "素材2", "女装"),
        make_sku("SKU2", "素材1", "女装"),
        make_sku("SKU3", "素材1", "女装"),
        make_sku("SKU4", "素材1", "女装"),
        make_sku("SKU5", "素材1", "女装"),
    ]

    final, warns, gaps = run_allocation(acc_rows, sku_rows, "模式一")
    print_result("用例1：模式一 阶段1 广度+顺延", final, gaps, warns)

    by_acc = {aid: items for aid, items, _ in summarize(final)}
    assert [x[0] for x in by_acc["账号1"]] == ["SKU1", "SKU2", "SKU3"]
    assert [x[1] for x in by_acc["账号1"]] == ["素材1", "素材1", "素材1"]
    assert [x[0] for x in by_acc["账号2"]] == ["SKU4", "SKU5"]
    assert_unique_skus(by_acc["账号1"], "账号1")
    assert_unique_skus(by_acc["账号2"], "账号2")
    assert not gaps
    print("✅ 用例1 通过")


# ─────────────────────────────────────────────
# 用例 2：模式一 阶段2 — 他号拿走素材1后，本号拿素材2（仍不重复 SKU）
# ─────────────────────────────────────────────

def test_mode1_phase2_variant_of_unowned_sku():
    """
    账号1 拿走 SKU1 素材1、SKU2 素材1
    账号2 要 2 个 → 应拿 SKU1 素材2 + SKU3 素材1（账号内不重复 SKU）
    """
    acc_rows = [
        make_acc("账号1", "女装", 2),
        make_acc("账号2", "女装", 2),
    ]
    sku_rows = [
        make_sku("SKU1", "素材1", "女装"),
        make_sku("SKU1", "素材2", "女装"),
        make_sku("SKU2", "素材1", "女装"),
        make_sku("SKU3", "素材1", "女装"),
    ]

    final, warns, gaps = run_allocation(acc_rows, sku_rows, "模式一")
    print_result("用例2：模式一 阶段2 未用变体（他号已拿素材1）", final, gaps, warns)

    s = {aid: items for aid, items, _ in summarize(final)}
    assert [x[0] for x in s["账号1"]] == ["SKU1", "SKU2"]
    assert ("SKU1", "素材2") in [(x[0], x[1]) for x in s["账号2"]]
    assert ("SKU3", "素材1") in [(x[0], x[1]) for x in s["账号2"]]
    assert_unique_skus(s["账号2"], "账号2")
    assert not gaps
    print("✅ 用例2 通过")


# ─────────────────────────────────────────────
# 用例 3：模式一 阶段3 + 缺口 — 账号内 SKU 上限
# ─────────────────────────────────────────────

def test_mode1_phase3_reuse_and_gap_unique_sku():
    """
    女装仅 3 个独立 SKU（SKU1 有 2 素材也不增加账号内名额）
    账号1 要 3 → 拿满 3 个不同 SKU
    账号2 要 5 → 阶段3 最多再拿 3 个不同 SKU，缺 2
    """
    acc_rows = [
        make_acc("账号1", "女装", 3),
        make_acc("账号2", "女装", 5),
    ]
    sku_rows = [
        make_sku("SKU1", "素材1", "女装"),
        make_sku("SKU1", "素材2", "女装"),
        make_sku("SKU2", "素材1", "女装"),
        make_sku("SKU3", "素材1", "女装"),
    ]

    final, warns, gaps = run_allocation(acc_rows, sku_rows, "模式一")
    print_result("用例3：模式一 账号内SKU上限 + 缺口", final, gaps, warns)

    s = {aid: (items, note) for aid, items, note in summarize(final)}
    assert len(s["账号1"][0]) == 3 and s["账号1"][1] == "完全匹配"
    assert_unique_skus(s["账号1"][0], "账号1")

    assert len(s["账号2"][0]) == 3
    assert "缺2个" in s["账号2"][1]
    assert "(3/5)" in s["账号2"][1]
    assert_unique_skus(s["账号2"][0], "账号2")
    print("✅ 用例3 通过")


# ─────────────────────────────────────────────
# 用例 4：模式一绝不跨品类
# ─────────────────────────────────────────────

def test_mode1_never_cross_category():
    acc_rows = [make_acc("账号1", "女装", 10)]
    sku_rows = [
        make_sku("SKU1", "素材1", "女装"),
        make_sku("SKU2", "素材1", "女装"),
        make_sku("建材A", "素材1", "家庭建材"),
        make_sku("建材B", "素材1", "家庭建材"),
    ]

    final, warns, gaps = run_allocation(
        acc_rows, sku_rows, "模式一", fallback_category="家庭建材"
    )
    print_result("用例4：模式一不跨品类", final, gaps, warns)

    items, note = summarize(final)[0][1], summarize(final)[0][2]
    assert len(items) == 2
    assert all(cat == "女装" for _, _, cat in items)
    assert "缺8个" in note
    print("✅ 用例4 通过")


# ─────────────────────────────────────────────
# 用例 5：模式二后备 — 全局消耗后下一账号拿素材2
# ─────────────────────────────────────────────

def test_mode2_fallback_global_consume_next_mat2():
    """
    女装 2 组合；后备 建材A/B/C 各 2 素材
    账号1 要 5 → 女装2 + 后备 A/B/C 素材1
    账号2 要 4 → 女装阶段3 复用2个不同SKU + 后备 A/B 素材2（不是再拿素材1）
    """
    acc_rows = [
        make_acc("账号1", "女装", 5),
        make_acc("账号2", "女装", 4),
    ]
    sku_rows = [
        make_sku("SKU1", "素材1", "女装"),
        make_sku("SKU2", "素材1", "女装"),
    ]
    for name in ["建材A", "建材B", "建材C"]:
        sku_rows.append(make_sku(name, "素材1", "家庭建材"))
        sku_rows.append(make_sku(name, "素材2", "家庭建材"))

    final, warns, gaps = run_allocation(
        acc_rows, sku_rows, "模式二", fallback_category="家庭建材"
    )
    print_result("用例5：模式二 后备全局消耗→下一号素材2", final, gaps, warns)

    s = {aid: (items, note) for aid, items, note in summarize(final)}

    assert len(s["账号1"][0]) == 5 and s["账号1"][1] == "完全匹配"
    assert_unique_combos(s["账号1"][0], "账号1")
    fb1 = [(sku, mat) for sku, mat, cat in s["账号1"][0] if cat == "家庭建材"]
    assert fb1 == [("建材A", "素材1"), ("建材B", "素材1"), ("建材C", "素材1")]

    assert len(s["账号2"][0]) == 4 and s["账号2"][1] == "完全匹配"
    assert_unique_combos(s["账号2"][0], "账号2")
    main2 = [(sku, mat) for sku, mat, cat in s["账号2"][0] if cat == "女装"]
    fb2 = [(sku, mat) for sku, mat, cat in s["账号2"][0] if cat == "家庭建材"]
    assert len(main2) == 2
    # 指针接续 + 全局消耗：A/B/C 素材1 已用 → 取 A/B 素材2
    assert fb2 == [("建材A", "素材2"), ("建材B", "素材2")]
    assert not gaps
    print("✅ 用例5 通过")


# ─────────────────────────────────────────────
# 用例 6：环形指针单元 + 跨账号接续（打印结果）
# ─────────────────────────────────────────────

def test_fallback_pointer_unit():
    """
    模式二后备：
      - 单账号可 A1→B1→C1→A2（同 SKU 不同素材允许）
      - 跨账号全局消耗：账号1 拿 A1/B1/C1 后，账号2 拿 A2/B2
    """
    fb_pool = []
    for name in ["A", "B", "C"]:
        fb_pool.append(make_sku(name, "素材1", "家庭建材"))
        fb_pool.append(make_sku(name, "素材2", "家庭建材"))

    print(f"\n{'=' * 60}")
    print("  用例6：环形指针 + 全局消耗 + 模式二允许同SKU不同素材")
    print(f"{'=' * 60}")

    # 单账号要 4：A1 B1 C1 A2
    selected0, used0, skus0, ptr0 = [], set(), set(), 0
    g0 = set()
    selected0, used0, skus0, ptr0 = allocate_fallback_round_robin(
        fb_pool, 4, selected0, used0, skus0, g0, ptr0, strict_unique_sku=False
    )
    got0 = [(_sku_from(i), i["广告素材版本名称"]) for i in selected0]
    print(f"\n▶ 单账号要4: {got0}  下一指针={ptr0}")
    assert got0 == [("A", "素材1"), ("B", "素材1"), ("C", "素材1"), ("A", "素材2")]

    # 跨账号：账号1 要 3 → A1 B1 C1；账号2 要 2 → A2 B2
    selected1, used1, skus1, ptr = [], set(), set(), 0
    global_combos = set()
    selected1, used1, skus1, ptr = allocate_fallback_round_robin(
        fb_pool, 3, selected1, used1, skus1, global_combos, ptr, strict_unique_sku=False
    )
    got1 = [(_sku_from(i), i["广告素材版本名称"]) for i in selected1]
    print(f"▶ 账号1 要3: {got1}  下一指针={ptr}")
    assert got1 == [("A", "素材1"), ("B", "素材1"), ("C", "素材1")]

    selected2, used2, skus2 = [], set(), set()
    selected2, used2, skus2, ptr2 = allocate_fallback_round_robin(
        fb_pool, 2, selected2, used2, skus2, global_combos, ptr, strict_unique_sku=False
    )
    got2 = [(_sku_from(i), i["广告素材版本名称"]) for i in selected2]
    print(f"▶ 账号2 要2: {got2}  下一指针={ptr2}")
    assert got2 == [("A", "素材2"), ("B", "素材2")]
    print("✅ 用例6 通过")


# ─────────────────────────────────────────────
# 用例 7：综合文档场景 — 账号内永不重复 SKU
# ─────────────────────────────────────────────

def test_comprehensive_doc_scenario():
    """
    女装：SKU1 有2素材，SKU2~11 各1 → 11 独立 SKU / 12 组合
    账号1 要3 / 账号2 要10 / 账号3 要13

    模式一（SKU 绝对不重复）：
      账号1: SKU1/2/3 素材1
      账号2: 10 个不同 SKU
      账号3: 最多 11 → 缺 2

    模式二（允许同 SKU 不同素材）+ 家庭建材：
      账号3: 主品类可拿满 12 组合 + 1 后备 → 13 完全匹配
    """
    acc_rows = [
        make_acc("账号1", "女装", 3),
        make_acc("账号2", "女装", 10),
        make_acc("账号3", "女装", 13),
    ]
    sku_rows = [
        make_sku("SKU1", "素材1", "女装"),
        make_sku("SKU1", "素材2", "女装"),
    ]
    for i in range(2, 12):
        sku_rows.append(make_sku(f"SKU{i}", "素材1", "女装"))
    for i in range(1, 4):
        sku_rows.append(make_sku(f"建材SKU_{i}", "素材1", "家庭建材"))
        sku_rows.append(make_sku(f"建材SKU_{i}", "素材2", "家庭建材"))

    # —— 模式一 ——
    final1, w1, g1 = run_allocation(acc_rows, sku_rows, "模式一")
    print_result("用例7a：综合-模式一（SKU绝对不重复）", final1, g1, w1)

    s1 = {aid: (items, note) for aid, items, note in summarize(final1)}
    assert len(s1["账号1"][0]) == 3
    assert [x[0] for x in s1["账号1"][0]] == ["SKU1", "SKU2", "SKU3"]
    assert s1["账号1"][1] == "完全匹配"
    assert_unique_skus(s1["账号1"][0], "账号1")

    assert len(s1["账号2"][0]) == 10
    assert s1["账号2"][1] == "完全匹配"
    assert_unique_skus(s1["账号2"][0], "账号2")
    assert [x[0] for x in s1["账号2"][0]].count("SKU1") == 1

    assert len(s1["账号3"][0]) == 11
    assert "缺2个" in s1["账号3"][1]
    assert "(11/13)" in s1["账号3"][1]
    assert_unique_skus(s1["账号3"][0], "账号3")

    # —— 模式二 ——
    final2, w2, g2 = run_allocation(
        acc_rows, sku_rows, "模式二", fallback_category="家庭建材"
    )
    print_result("用例7b：综合-模式二（允许同SKU不同素材）", final2, g2, w2)

    s2 = {aid: (items, note) for aid, items, note in summarize(final2)}
    assert len(s2["账号1"][0]) == 3 and s2["账号1"][1] == "完全匹配"
    assert len(s2["账号2"][0]) == 10 and s2["账号2"][1] == "完全匹配"
    assert_unique_combos(s2["账号2"][0], "账号2-模式二")
    assert len(s2["账号3"][0]) == 13 and s2["账号3"][1] == "完全匹配"
    assert_unique_combos(s2["账号3"][0], "账号3-模式二")
    main3 = [x for x in s2["账号3"][0] if x[2] == "女装"]
    fb3 = [x for x in s2["账号3"][0] if x[2] == "家庭建材"]
    # 模式二主品类可拿满 12 组合（含 SKU1 两个不同素材）+ 1 后备
    assert len(main3) == 12
    assert len(fb3) == 1
    assert not g2
    print("✅ 用例7 通过")


# ─────────────────────────────────────────────
# 用例 8：阶段1 不重 SKU
# ─────────────────────────────────────────────

def test_phase1_no_duplicate_sku_in_first_wave():
    acc_rows = [make_acc("账号1", "女装", 5)]
    sku_rows = []
    for i in range(1, 6):
        sku_rows.append(make_sku(f"SKU{i}", "素材1", "女装"))
        sku_rows.append(make_sku(f"SKU{i}", "素材2", "女装"))
        sku_rows.append(make_sku(f"SKU{i}", "素材3", "女装"))

    final, _, _ = run_allocation(acc_rows, sku_rows, "模式一")
    items = summarize(final)[0][1]
    skus = [x[0] for x in items]
    mats = [x[1] for x in items]
    assert skus == ["SKU1", "SKU2", "SKU3", "SKU4", "SKU5"]
    assert mats == ["素材1"] * 5
    assert_unique_skus(items, "账号1")
    print("✅ 用例8 通过")


# ─────────────────────────────────────────────
# 用例 9：主品类已消耗 ABC 素材1 → 后备拿 DEF + ABC 素材2
# ─────────────────────────────────────────────

def test_fallback_prefers_unused_sku_then_mat2():
    """
    账号1 主品类=家庭建材，要 3 → 拿走 建材A/B/C 素材1
    账号2 主品类=女装(只有1个SKU) 要 5，后备=家庭建材
      → 女装1 + 后备：优先未用 SKU D/E/F 素材1，以及 A/B/C 的素材2
    """
    acc_rows = [
        make_acc("账号1", "家庭建材", 3),
        make_acc("账号2", "女装", 5),
    ]
    sku_rows = [make_sku("女装SKU1", "素材1", "女装")]
    for name in ["建材A", "建材B", "建材C", "建材D", "建材E", "建材F"]:
        sku_rows.append(make_sku(name, "素材1", "家庭建材"))
        sku_rows.append(make_sku(name, "素材2", "家庭建材"))

    final, warns, gaps = run_allocation(
        acc_rows, sku_rows, "模式二", fallback_category="家庭建材"
    )
    print_result("用例9：主品类消耗ABC后后备取DEF+ABC素材2", final, gaps, warns)

    s = {aid: (items, note) for aid, items, note in summarize(final)}
    a1 = s["账号1"][0]
    assert [(x[0], x[1]) for x in a1] == [
        ("建材A", "素材1"), ("建材B", "素材1"), ("建材C", "素材1")
    ]

    a2 = s["账号2"][0]
    assert_unique_combos(a2, "账号2")
    assert len(a2) == 5 and s["账号2"][1] == "完全匹配"
    assert a2[0][0] == "女装SKU1"

    fb2 = [(sku, mat) for sku, mat, cat in a2 if cat == "家庭建材"]
    assert len(fb2) == 4
    # 优先未用独立 SKU 的素材1（DEF），再补 ABC 的素材2
    assert fb2 == [
        ("建材D", "素材1"),
        ("建材E", "素材1"),
        ("建材F", "素材1"),
        ("建材A", "素材2"),
    ]
    assert ("建材A", "素材1") not in fb2
    assert ("建材B", "素材1") not in fb2
    assert ("建材C", "素材1") not in fb2
    print("✅ 用例9 通过")


def test_phase3_sequential_no_restart():
    """
    阶段3 跨账号顺延：全库第一轮被账号A用尽后，
    账号B、账号C 第二轮不应从同一起点重拿导致列表相同。
    """
    acc_rows = [
        make_acc("账号A", "女装", 5),  # 拿走全部 5 组合
        make_acc("账号B", "女装", 2),
        make_acc("账号C", "女装", 2),
    ]
    sku_rows = [make_sku(f"SKU{i}", "素材1", "女装") for i in range(1, 6)]

    final, _, _ = run_allocation(acc_rows, sku_rows, "模式一")
    print_result("用例11：阶段3 跨账号顺延不重开", final, [], [])

    s = {aid: items for aid, items, _ in summarize(final)}
    assert len(s["账号A"]) == 5
    assert len(s["账号B"]) == 2
    assert len(s["账号C"]) == 2
    set_b = {(x[0], x[1]) for x in s["账号B"]}
    set_c = {(x[0], x[1]) for x in s["账号C"]}
    assert set_b.isdisjoint(set_c), f"账号B/C 不应相同: B={set_b} C={set_c}"
    # 顺延：B 拿 SKU1/2，C 拿 SKU3/4
    assert [x[0] for x in s["账号B"]] == ["SKU1", "SKU2"]
    assert [x[0] for x in s["账号C"]] == ["SKU3", "SKU4"]
    print("✅ 用例11 通过")


def test_mode_diff_same_sku_different_material():
    """
    同一池：SKU1 素材1/2，SKU2 素材1；账号要 3
    模式一：最多 2（SKU 不重复）→ 缺1
    模式二：可拿 SKU1素材1 + SKU2素材1 + SKU1素材2 → 完全匹配
    """
    acc_rows = [make_acc("账号1", "女装", 3)]
    sku_rows = [
        make_sku("SKU1", "素材1", "女装"),
        make_sku("SKU1", "素材2", "女装"),
        make_sku("SKU2", "素材1", "女装"),
    ]

    f1, _, g1 = run_allocation(acc_rows, sku_rows, "模式一")
    print_result("用例10a：模式一 禁止同SKU不同素材", f1, g1, [])
    items1 = summarize(f1)[0][1]
    assert len(items1) == 2
    assert_unique_skus(items1, "模式一账号1")
    assert "缺1个" in summarize(f1)[0][2]

    f2, _, g2 = run_allocation(acc_rows, sku_rows, "模式二")
    print_result("用例10b：模式二 允许同SKU不同素材", f2, g2, [])
    items2 = summarize(f2)[0][1]
    assert len(items2) == 3
    assert_unique_combos(items2, "模式二账号1")
    pairs = [(x[0], x[1]) for x in items2]
    assert ("SKU1", "素材1") in pairs
    assert ("SKU1", "素材2") in pairs
    assert ("SKU2", "素材1") in pairs
    assert not g2
    print("✅ 用例10 通过")


if __name__ == "__main__":
    tests = [
        test_mode1_phase1_breadth_and_sequential,
        test_mode1_phase2_variant_of_unowned_sku,
        test_mode1_phase3_reuse_and_gap_unique_sku,
        test_mode1_never_cross_category,
        test_mode2_fallback_global_consume_next_mat2,
        test_fallback_pointer_unit,
        test_comprehensive_doc_scenario,
        test_phase1_no_duplicate_sku_in_first_wave,
        test_fallback_prefers_unused_sku_then_mat2,
        test_mode_diff_same_sku_different_material,
        test_phase3_sequential_no_restart,
    ]
    failed = []
    for t in tests:
        try:
            t()
        except AssertionError as e:
            print(f"❌ {t.__name__} 失败: {e}")
            failed.append(t.__name__)
        except Exception as e:
            print(f"💥 {t.__name__} 异常: {type(e).__name__}: {e}")
            failed.append(t.__name__)

    print("\n" + "=" * 60)
    if not failed:
        print(f"🎉 全部 {len(tests)} 个用例通过")
    else:
        print(f"❌ {len(failed)}/{len(tests)} 失败: {failed}")
        sys.exit(1)
