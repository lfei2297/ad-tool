# -*- coding: utf-8 -*-
from modules.module_4 import (
    DEFAULT_VERSION,
    build_material_pool,
    build_series_matrices,
    flatten_series_rows,
    iter_campaign_slots,
)


def _names(logic_mode, materials, m, n, series):
    matrices = build_series_matrices(logic_mode, materials, m, n, series)
    return flatten_series_rows(matrices, m, n)


def test_within_group_two_series_use_different_materials():
    mats = [f"v{i}" for i in range(1, 14)]
    got = _names("组内测素材", mats, 1, 3, 2)
    assert got == ["v1", "v2", "v3", "v4", "v5", "v6"]


def test_across_group_two_series_use_different_materials():
    mats = [f"v{i}" for i in range(1, 14)]
    got = _names("组间测素材", mats, 1, 3, 2)
    assert got == ["v1", "v1", "v1", "v2", "v2", "v2"]


def test_within_group_pads_default_when_short():
    mats = [f"v{i}" for i in range(1, 6)]
    got = _names("组内测素材", mats, 1, 3, 2)
    assert got == ["v1", "v2", "v3", "v4", "v5", DEFAULT_VERSION]


def test_across_group_two_groups_consume_one_each_then_next_series():
    mats = [f"v{i}" for i in range(1, 5)]
    got = _names("组间测素材", mats, 2, 3, 2)
    # Ad1 G1/G2, Ad2 G1/G2, Ad3 G1/G2, then series 2
    assert got == [
        "v1", "v2", "v1", "v2", "v1", "v2",
        "v3", "v4", "v3", "v4", "v3", "v4",
    ]


def test_row_order_two_series_structure_1_2_3():
    slots = list(iter_campaign_slots(2, 3, 2))
    assert len(slots) == 12
    assert [s for s, _, _ in slots[:6]] == [0] * 6
    assert [s for s, _, _ in slots[6:]] == [1] * 6
    assert [g for _, g, _ in slots[:6]] == [0, 1, 0, 1, 0, 1]
    assert [a for _, _, a in slots[:6]] == [0, 0, 1, 1, 2, 2]
    assert [g for _, g, _ in slots[6:]] == [0, 1, 0, 1, 0, 1]
    assert [a for _, _, a in slots[6:]] == [0, 0, 1, 1, 2, 2]


def test_within_group_1_2_3_two_series_is_twelve_distinct():
    mats = [f"v{i}" for i in range(1, 13)]
    got = _names("组内测素材", mats, 2, 3, 2)
    assert got == [f"v{i}" for i in range(1, 13)]


def test_user_template_same_name_within_group_is_1_to_6():
    row = {
        "着陆页版本名称": "优化组版本-OPDY-RX-7",
        "广告素材版本名称": "优化组版本-OPDY-RX-7",
    }
    pool = build_material_pool(row, 13)
    got = _names("组内测素材", pool, 1, 3, 2)
    assert got == [f"优化组版本-OPDY-RX-7-{i}" for i in range(1, 7)]
