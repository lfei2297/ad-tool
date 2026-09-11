# -*- coding: utf-8 -*-
from modules.module_3 import process_step1, smart_logic


def test_step1_applies_repeat():
    rows = [
        {
            "真实SKU": "S1",
            "广告素材版本名称": "A-1",
            "广告素材数量": "2",
        }
    ]
    out = process_step1(rows, {"repeat_1": 3})
    assert len(out) == 6


def test_step2_expanded_table_does_not_repeat():
    df_src = process_step1(
        [{"真实SKU": "S1", "广告素材版本名称": "A-1", "广告素材数量": "2"}],
        {"repeat_1": 3},
    )
    # 展开总表导出时会丢掉「广告素材数量」
    df_src = df_src.drop(columns=["广告素材数量"], errors="ignore")
    out = smart_logic(df_src, size=30, params={"repeat_1": 3})
    assert len(out) == 6


def test_step2_raw_table_expands_versions_once():
    import pandas as pd

    df = pd.DataFrame(
        [{"真实SKU": "S1", "广告素材版本名称": "A-1", "广告素材数量": "2"}]
    )
    out = smart_logic(df, size=30, params={"repeat_1": 3})
    assert len(out) == 2
    assert set(out["广告素材版本名称"]) == {"A-1", "A-2"}


if __name__ == "__main__":
    test_step1_applies_repeat()
    print("ok step1 repeat")
    test_step2_expanded_table_does_not_repeat()
    print("ok step2 no re-repeat")
    test_step2_raw_table_expands_versions_once()
    print("ok step2 raw expand once")
