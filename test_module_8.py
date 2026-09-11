# -*- coding: utf-8 -*-
"""模块八模式三：同素材数合表校验。"""
import io
from collections import defaultdict
from pathlib import Path

import pandas as pd

from modules.module_8 import (
    plan_mode3_packing,
    run_module_8_matching,
    get_sku_key,
    write_landing_page_excel,
)


def test_plan_absorbs_singleton_into_larger_base():
    cnt_to_skus = {
        10: ["A", "B"],
        5: ["C"],
        3: ["D"],
    }
    plans = plan_mode3_packing(cnt_to_skus)
    assert len(plans) == 2
    by_base = {p["base_count"]: p for p in plans}
    assert 10 in by_base
    assert by_base[10]["tail_sku"] == "C"
    assert by_base[10]["tail_count"] == 5
    assert 3 in by_base
    assert by_base[3]["tail_sku"] is None


def test_plan_prefers_dissolving_small_groups():
    cnt_to_skus = {c: [f"{c}_{i}" for i in range(n)] for c, n in {
        10: 2, 9: 1, 8: 1,
    }.items()}
    plans = plan_mode3_packing(cnt_to_skus)
    assert len(plans) == 2
    all_skus = {sku for skus in cnt_to_skus.values() for sku in skus}
    covered = set()
    for p in plans:
        covered.update(p["base_skus"])
        if p["tail_sku"]:
            covered.add(p["tail_sku"])
            assert p["tail_count"] < p["base_count"]
    assert covered == all_skus


def _sku_counts_in_table(df, sku_cnt):
    per_acc = defaultdict(lambda: defaultdict(int))
    for _, r in df.iterrows():
        acc = str(r.get("广告账号ID", "")).strip()
        sku = get_sku_key(r)
        if acc and sku:
            per_acc[acc][sku] += 1
    return per_acc, sku_cnt


def test_real_file_mode3_min_tables():
    path = Path(r"D:\Downloads\2026-09\模块八新增功能点0908.xlsx")
    assert path.exists(), f"missing {path}"
    data = path.read_bytes()
    tables, _, _ = run_module_8_matching(data, n_group=50, mode="模式三")
    assert tables, "mode3 produced no tables"
    print(f"mode3 tables: {len(tables)}")
    for name in tables:
        print(" ", name, "rows", len(tables[name]))
    assert len(tables) == 33, f"expected 33 tables, got {len(tables)}"

    import pandas as pd
    import io
    df_sku = pd.read_excel(io.BytesIO(data), sheet_name="SKU表", dtype=str)
    df_sku.columns = [str(c).strip() for c in df_sku.columns]
    df_sku["SKU_KEY"] = df_sku.apply(get_sku_key, axis=1)
    df_sku["mat"] = df_sku["广告素材版本"].fillna("").astype(str).str.strip()
    df_sku = df_sku[(df_sku["SKU_KEY"] != "") & (df_sku["mat"] != "")]
    sku_cnt_raw = df_sku.groupby("SKU_KEY")["mat"].count().to_dict()
    sku_cnt = {k: min(v, 50) for k, v in sku_cnt_raw.items()}

    df_acc = pd.read_excel(io.BytesIO(data), sheet_name="账号表", dtype=str)
    df_acc.columns = [str(c).strip() for c in df_acc.columns]
    if any("可不填" in str(v) for v in df_acc.iloc[0].values):
        df_acc = df_acc.iloc[1:]
    df_acc = df_acc.dropna(how="all")
    df_acc["SKU_KEY"] = df_acc.apply(get_sku_key, axis=1)
    expected_pairs = set(zip(df_acc["广告账号ID"].astype(str), df_acc["SKU_KEY"]))

    covered_pairs = set()
    for name, df in tables.items():
        assert "SKU_KEY" not in df.columns, name
        assert "广告素材版本名称" in df.columns, name
        assert "广告素材ID" in df.columns, name
        cols = list(df.columns)
        assert cols.index("广告素材版本名称") == cols.index("广告素材ID") - 1, cols

        per_acc = defaultdict(set)
        for _, r in df.iterrows():
            acc = str(r.get("广告账号ID", "")).strip()
            sku = get_sku_key(r)
            if not acc or not sku:
                continue
            per_acc[acc].add(sku)
            covered_pairs.add((acc, sku))

        for acc, skus in per_acc.items():
            counts = [sku_cnt.get(s, 1) for s in skus]
            max_c = max(counts)
            n_smaller = sum(1 for c in counts if c < max_c)
            n_equal = sum(1 for c in counts if c == max_c)
            n_larger = sum(1 for c in counts if c > max_c)
            assert n_larger == 0, (name, acc, counts)
            assert n_smaller <= 1, f"{name} acc {acc} has {n_smaller} smaller SKUs: {list(zip(skus, counts))}"
            assert n_equal >= 1, (name, acc, counts)

            filled = df[df["广告账号ID"].astype(str) == acc]
            for sku in skus:
                sku_rows = filled[filled.apply(get_sku_key, axis=1) == sku]
                expect = sku_cnt.get(sku, 1)
                assert len(sku_rows) == expect, (name, acc, sku, len(sku_rows), expect)
                assert sku_rows["广告素材版本名称"].astype(str).str.strip().ne("").all()
                assert sku_rows["广告素材ID"].astype(str).str.strip().ne("").all()

    assert covered_pairs == expected_pairs, (
        f"missing {len(expected_pairs - covered_pairs)} extra {len(covered_pairs - expected_pairs)}"
    )


def test_g_col_version_name_not_renamed_to_link():
    df = pd.DataFrame([{
        "广告账号ID": "1",
        "虚拟SKU": "ABC",
        "国家": "美国",
        "着陆页链接": "优化组版本-SPF-1",
        "广告素材版本名称": "优化组版本-1",
        "广告素材ID": "ABC-US-1",
    }])
    xls = write_landing_page_excel(df, keep_material_col="广告素材ID")
    out = pd.read_excel(io.BytesIO(xls), dtype=str)
    assert "着陆页版本名称" in out.columns
    assert "着陆页链接" not in out.columns
    assert str(out.iloc[1]["着陆页版本名称"]).strip() == "优化组版本-SPF-1"


def test_g_col_url_stays_link():
    df = pd.DataFrame([{
        "广告账号ID": "1",
        "虚拟SKU": "ABC",
        "国家": "美国",
        "着陆页链接": "https://example.com/products/x-1",
        "广告素材版本名称": "优化组版本-1",
        "广告素材ID": "ABC-US-1",
    }])
    xls = write_landing_page_excel(df, keep_material_col="广告素材ID")
    out = pd.read_excel(io.BytesIO(xls), dtype=str)
    assert "着陆页链接" in out.columns
    assert "着陆页版本名称" not in out.columns


def test_export_keeps_both_material_columns_fills_one():
    df = pd.DataFrame([{
        "广告账号ID": "1",
        "主页ID": "",
        "像素ID": "",
        "真实SKU": "",
        "虚拟SKU": "ABC",
        "国家": "美国",
        "着陆页链接": "https://example.com/p/1",
        "广告素材版本名称": "优化组版本-1",
        "广告素材ID": "ABC-US-1",
        "出价/竞价": "",
        "系列标注": "",
    }])
    id_xls = write_landing_page_excel(df, keep_material_col="广告素材ID")
    id_df = pd.read_excel(io.BytesIO(id_xls), dtype=str)
    assert "广告素材ID" in id_df.columns
    assert "广告素材版本名称" in id_df.columns
    assert list(id_df.columns).index("广告素材ID") == list(id_df.columns).index("广告素材版本名称") + 1
    data_id = id_df.iloc[1]
    assert str(data_id["广告素材ID"]).strip() == "ABC-US-1"
    assert str(data_id["广告素材版本名称"]).strip() in ("", "nan", "None")

    ver_xls = write_landing_page_excel(df, keep_material_col="广告素材版本名称")
    ver_df = pd.read_excel(io.BytesIO(ver_xls), dtype=str)
    assert "广告素材版本名称" in ver_df.columns
    assert "广告素材ID" in ver_df.columns
    data_ver = ver_df.iloc[1]
    assert str(data_ver["广告素材版本名称"]).strip() == "优化组版本-1"
    assert str(data_ver["广告素材ID"]).strip() in ("", "nan", "None")


def _m8_bytes(acc_rows, sku_rows):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        pd.DataFrame(acc_rows).to_excel(w, index=False, sheet_name="账号表")
        pd.DataFrame(sku_rows).to_excel(w, index=False, sheet_name="SKU表")
    return out.getvalue()


def _sku_rows(sku, n, prefix="V"):
    rows = []
    for i in range(n):
        rows.append({
            "真实SKU": "",
            "虚拟SKU": sku,
            "SKU中文名称": sku,
            "国家": "美国",
            "广告素材版本": f"{prefix}-{n - i}",
            "广告素材ID": f"{sku}-{n - i}",
            "创建时间": f"2026-09-0{min(i + 1, 9)} 10:00:00",
        })
    return rows


def test_mode45_per_sku_cap():
    acc = [{"广告账号ID": "A1", "虚拟SKU": "S1", "国家": "美国"}]
    sku = _sku_rows("S1", 30)
    tables, _, _ = run_module_8_matching(
        _m8_bytes(acc, sku), n_group=50, mode="模式五", per_sku_limit=20
    )
    df = next(iter(tables.values()))
    assert len(df) == 20
    tables4, _, _ = run_module_8_matching(
        _m8_bytes(acc, sku), n_group=50, mode="模式四", per_sku_limit=20
    )
    assert len(next(iter(tables4.values()))) == 20


def test_mode4_truncate_even_split():
    acc = [{"广告账号ID": "A1", "虚拟SKU": s, "国家": "美国", "着陆页链接": "https://x.com/p"} for s in ("S1", "S2", "S3")]
    sku = _sku_rows("S1", 30) + _sku_rows("S2", 30) + _sku_rows("S3", 30)
    tables, _, _ = run_module_8_matching(_m8_bytes(acc, sku), n_group=50, mode="模式四")
    assert len(tables) == 1
    df = next(iter(tables.values()))
    assert len(df) == 50
    counts = df.groupby(df.apply(get_sku_key, axis=1)).size().to_dict()
    assert set(counts) == {"S1", "S2", "S3"}
    assert max(counts.values()) - min(counts.values()) <= 1
    order = df.apply(get_sku_key, axis=1).tolist()
    assert order[:3] == ["S1", "S2", "S3"]


def test_mode4_sku_over_n_takes_first_n_one_each():
    skus = [f"S{i:03d}" for i in range(60)]
    acc = [{"广告账号ID": "A1", "虚拟SKU": s, "国家": "美国"} for s in skus]
    sku = []
    for s in skus:
        sku.extend(_sku_rows(s, 5))
    tables, _, _ = run_module_8_matching(_m8_bytes(acc, sku), n_group=50, mode="模式四")
    df = next(iter(tables.values()))
    assert len(df) == 50
    got = df.apply(get_sku_key, axis=1).tolist()
    assert got == skus[:50]
    assert df.groupby(df.apply(get_sku_key, axis=1)).size().max() == 1


def test_mode5_full_stays_one_file_when_over_n():
    acc = [{"广告账号ID": "A1", "虚拟SKU": s, "国家": "美国"} for s in ("S1", "S2", "S3")]
    sku = _sku_rows("S1", 40) + _sku_rows("S2", 40) + _sku_rows("S3", 40)
    tables, _, _ = run_module_8_matching(_m8_bytes(acc, sku), n_group=50, mode="模式五")
    assert len(tables) == 1
    name = next(iter(tables))
    assert "可多系列" in name
    df = next(iter(tables.values()))
    assert len(df) == 120
    assert df.apply(get_sku_key, axis=1).tolist()[:3] == ["S1", "S2", "S3"]


def test_mode4_two_groups_two_series():
    acc = (
        [{"广告账号ID": "A1", "虚拟SKU": "G1A", "SKU组": "1", "国家": "美国"}]
        + [{"广告账号ID": "A1", "虚拟SKU": "G1B", "SKU组": "1", "国家": "美国"}]
        + [{"广告账号ID": "A1", "虚拟SKU": "G2A", "SKU组": "2", "国家": "美国"}]
        + [{"广告账号ID": "A1", "虚拟SKU": "G2B", "SKU组": "2", "国家": "美国"}]
    )
    sku = _sku_rows("G1A", 10) + _sku_rows("G1B", 10) + _sku_rows("G2A", 8) + _sku_rows("G2B", 8)
    tables, _, _ = run_module_8_matching(_m8_bytes(acc, sku), n_group=50, mode="模式四")
    assert len(tables) == 2
    names = list(tables)
    assert "第1组" in names[0] and "第2组" in names[1]
    g1 = list(tables.values())[0].apply(get_sku_key, axis=1)
    g2 = list(tables.values())[1].apply(get_sku_key, axis=1)
    assert set(g1) == {"G1A", "G1B"}
    assert set(g2) == {"G2A", "G2B"}
    assert len(list(tables.values())[0]) == 20
    assert len(list(tables.values())[1]) == 16


def test_mode5_two_groups_full_multi_series():
    acc = (
        [{"广告账号ID": "A1", "虚拟SKU": s, "SKU组": "1", "国家": "美国"} for s in ("G1A", "G1B")]
        + [{"广告账号ID": "A1", "虚拟SKU": s, "SKU组": "2", "国家": "美国"} for s in ("G2A", "G2B")]
    )
    sku = _sku_rows("G1A", 40) + _sku_rows("G1B", 40) + _sku_rows("G2A", 10) + _sku_rows("G2B", 10)
    tables, _, _ = run_module_8_matching(_m8_bytes(acc, sku), n_group=50, mode="模式五")
    names = list(tables)
    assert len(tables) == 2
    assert any("第1组" in n and "可多系列" in n for n in names)
    assert any("第2组" in n for n in names)
    g1 = next(df for n, df in tables.items() if "第1组" in n)
    g2 = next(df for n, df in tables.items() if "第2组" in n)
    assert len(g1) == 80
    assert len(g2) == 20


def test_series_label_does_not_split_groups():
    acc = (
        [{"广告账号ID": "A1", "虚拟SKU": "S1", "系列标注": "夏", "国家": "美国"}]
        + [{"广告账号ID": "A1", "虚拟SKU": "S2", "系列标注": "冬", "国家": "美国"}]
    )
    sku = _sku_rows("S1", 3) + _sku_rows("S2", 3)
    tables, _, _ = run_module_8_matching(_m8_bytes(acc, sku), n_group=50, mode="模式四")
    assert len(tables) == 1
    df = next(iter(tables.values()))
    assert set(df.apply(get_sku_key, axis=1)) == {"S1", "S2"}
    assert "SKU组" not in df.columns


def test_sku_group_float_and_dropped_on_export():
    acc = (
        [{"广告账号ID": "A1", "虚拟SKU": "G1A", "SKU组": "1.0", "国家": "美国", "着陆页链接": "https://x.com/p"}]
        + [{"广告账号ID": "A1", "虚拟SKU": "G1B", "SKU组": "1", "国家": "美国", "着陆页链接": "https://x.com/p"}]
        + [{"广告账号ID": "A1", "虚拟SKU": "G2A", "SKU组": "2", "国家": "美国", "着陆页链接": "https://x.com/p"}]
    )
    sku = _sku_rows("G1A", 2) + _sku_rows("G1B", 2) + _sku_rows("G2A", 2)
    tables, _, _ = run_module_8_matching(_m8_bytes(acc, sku), n_group=50, mode="模式四")
    assert len(tables) == 2
    g1 = list(tables.values())[0]
    assert set(g1.apply(get_sku_key, axis=1)) == {"G1A", "G1B"}
    xls = write_landing_page_excel(g1, keep_material_col="广告素材ID")
    out = pd.read_excel(io.BytesIO(xls), dtype=str)
    assert "SKU组" not in out.columns


def test_match_by_material_id_when_version_empty():
    acc = [{"广告账号ID": "A1", "虚拟SKU": "IN1", "国家": "美国", "着陆页链接": "https://x.com/p"}]
    sku = [{
        "真实SKU": "",
        "虚拟SKU": "IN1",
        "SKU中文名称": "",
        "国家": "美国",
        "广告素材版本": "",
        "广告素材ID": f"IN1-US-{i}",
        "创建时间": "",
    } for i in range(3)]
    tables, _, _ = run_module_8_matching(_m8_bytes(acc, sku), n_group=50, mode="模式四")
    df = next(iter(tables.values()))
    assert len(df) == 3
    assert df["广告素材ID"].tolist() == ["IN1-US-0", "IN1-US-1", "IN1-US-2"]


def test_pick_richer_real_over_virtual():
    acc = [{"广告账号ID": "A1", "真实SKU": "RR", "虚拟SKU": "VV", "国家": "美国"}]
    sku = _sku_rows("VV", 2) + [
        {
            "真实SKU": "RR",
            "虚拟SKU": "",
            "SKU中文名称": "RR",
            "国家": "美国",
            "广告素材版本": f"R-{i}",
            "广告素材ID": f"RR-{i}",
            "创建时间": f"2026-09-0{i} 12:00:00",
        }
        for i in range(1, 6)
    ]
    tables, _, _ = run_module_8_matching(_m8_bytes(acc, sku), n_group=50, mode="模式五")
    df = next(iter(tables.values()))
    assert len(df) == 5
    assert df["广告素材ID"].tolist() == ["RR-5", "RR-4", "RR-3", "RR-2", "RR-1"]


def test_unmatched_sku_skipped_and_reported():
    acc = [
        {"广告账号ID": "A1", "虚拟SKU": "HAS", "国家": "美国"},
        {"广告账号ID": "A1", "虚拟SKU": "MISS", "国家": "美国"},
    ]
    sku = _sku_rows("HAS", 3)
    tables, _, report = run_module_8_matching(_m8_bytes(acc, sku), n_group=50, mode="模式四")
    df = next(iter(tables.values()))
    assert set(df.apply(get_sku_key, axis=1)) == {"HAS"}
    assert len(df) == 3
    assert report["sku_material_rows"] == 3
    assert report["output_rows"] == 3
    skus = {u["SKU"] for u in report["unmatched"]}
    assert "MISS" in skus


def test_mode1_uses_richer_real_sku():
    acc = [{"广告账号ID": "A1", "真实SKU": "RR", "虚拟SKU": "VV", "国家": "美国"}]
    sku = _sku_rows("VV", 2) + [
        {
            "真实SKU": "RR",
            "虚拟SKU": "",
            "SKU中文名称": "RR",
            "国家": "美国",
            "广告素材版本": f"R-{i}",
            "广告素材ID": f"RR-{i}",
            "创建时间": f"2026-09-0{i} 12:00:00",
        }
        for i in range(1, 6)
    ]
    tables, _, _ = run_module_8_matching(_m8_bytes(acc, sku), n_group=50, mode="模式一")
    df = next(iter(tables.values()))
    assert len(df) == 5
    assert df["广告素材ID"].tolist() == ["RR-5", "RR-4", "RR-3", "RR-2", "RR-1"]


def test_mixed_sku_group_warning():
    acc = [
        {"广告账号ID": "A1", "虚拟SKU": "G1A", "SKU组": "1", "国家": "美国"},
        {"广告账号ID": "A1", "虚拟SKU": "G1B", "SKU组": "", "国家": "美国"},
    ]
    sku = _sku_rows("G1A", 2) + _sku_rows("G1B", 2)
    tables, _, report = run_module_8_matching(_m8_bytes(acc, sku), n_group=50, mode="模式四")
    assert len(tables) == 2
    assert report["warnings"]
    assert "未填 SKU组" in report["warnings"][0]


def test_mode1_one_sku_per_account():
    path = Path(r"D:\Downloads\2026-09\模块八新增功能点0908.xlsx")
    data = path.read_bytes()
    tables, _, _ = run_module_8_matching(data, n_group=50, mode="模式一")
    assert len(tables) == 288
    sample = next(iter(tables.values()))
    per_acc = sample.groupby("广告账号ID")
    for acc, grp in per_acc:
        skus = set(grp.apply(get_sku_key, axis=1))
        assert len(skus) == 1, (acc, skus)


if __name__ == "__main__":
    test_plan_absorbs_singleton_into_larger_base()
    print("ok tiny absorb")
    test_plan_prefers_dissolving_small_groups()
    print("ok dissolve")
    test_real_file_mode3_min_tables()
    print("ok real file mode3")
    test_g_col_version_name_not_renamed_to_link()
    print("ok g version name")
    test_g_col_url_stays_link()
    print("ok g url")
    test_export_keeps_both_material_columns_fills_one()
    print("ok export keep-both-fill-one")
    test_mode45_per_sku_cap()
    print("ok per-sku cap")
    test_mode4_truncate_even_split()
    print("ok mode4 even")
    test_mode4_sku_over_n_takes_first_n_one_each()
    print("ok mode4 sku>n")
    test_mode5_full_stays_one_file_when_over_n()
    print("ok mode5 one file")
    test_mode4_two_groups_two_series()
    print("ok mode4 two groups")
    test_mode5_two_groups_full_multi_series()
    print("ok mode5 two groups")
    test_series_label_does_not_split_groups()
    print("ok 系列标注 not group")
    test_sku_group_float_and_dropped_on_export()
    print("ok SKU组 export")
    test_match_by_material_id_when_version_empty()
    print("ok id-only match")
    test_pick_richer_real_over_virtual()
    print("ok richer sku")
    test_unmatched_sku_skipped_and_reported()
    print("ok unmatched report")
    test_mode1_uses_richer_real_sku()
    print("ok mode1 richer")
    test_mixed_sku_group_warning()
    print("ok sku group warn")
    test_mode1_one_sku_per_account()
    print("ok mode1")
    print("ALL PASSED")
