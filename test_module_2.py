# -*- coding: utf-8 -*-
from modules.module_2 import build_module2_tables


def _row(sku, country, ver, count):
    return {
        "广告账号ID": "A1",
        "真实SKU": sku,
        "虚拟SKU": "",
        "国家": country,
        "广告素材版本名称": ver,
        "广告素材数量": str(count),
    }


def test_same_country_same_count_merge_into_one_file():
    rows = [
        _row("SKU1", "美国", "素材-1", 5),
        _row("SKU2", "美国", "素材-1", 5),
    ]
    tasks = build_module2_tables(rows)
    assert "总表" in tasks
    assert "美国_素材数_5" in tasks
    sub = tasks["美国_素材数_5"]
    skus = set(sub["真实SKU"].tolist())
    assert skus == {"SKU1", "SKU2"}
    assert len(sub) == 10
    assert len(tasks["总表"]) == 10


def test_different_count_stay_separate():
    rows = [
        _row("SKU1", "美国", "素材-1", 5),
        _row("SKU2", "美国", "素材-1", 3),
        _row("SKU3", "英国", "素材-1", 5),
    ]
    tasks = build_module2_tables(rows)
    assert "美国_素材数_5" in tasks
    assert "美国_素材数_3" in tasks
    assert "英国_素材数_5" in tasks
    assert set(tasks["美国_素材数_5"]["真实SKU"]) == {"SKU1"}
    assert set(tasks["英国_素材数_5"]["真实SKU"]) == {"SKU3"}


if __name__ == "__main__":
    test_same_country_same_count_merge_into_one_file()
    test_different_count_stay_separate()
    print("ok")
