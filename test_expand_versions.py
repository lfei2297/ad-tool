# -*- coding: utf-8 -*-
from utils import expand_material_versions


def _expand(**kwargs):
    return expand_material_versions(kwargs)


def test_classic_simple_suffix_increments():
    assert _expand(广告素材版本名称="A-1", 广告素材数量=2) == ["A-1", "A-2"]
    assert _expand(广告素材版本名称="优化组版本-1", 广告素材数量=3) == [
        "优化组版本-1",
        "优化组版本-2",
        "优化组版本-3",
    ]
    assert _expand(广告素材版本名称="素材-1", 广告素材数量=5)[:2] == ["素材-1", "素材-2"]


def test_name_without_numeric_tail_appends():
    assert _expand(广告素材版本名称="默认版本", 广告素材数量=3) == [
        "默认版本-1",
        "默认版本-2",
        "默认版本-3",
    ]
    assert _expand(广告素材版本名称="优化组版本-OPDY-RX", 广告素材数量=2) == [
        "优化组版本-OPDY-RX-1",
        "优化组版本-OPDY-RX-2",
    ]


def test_count_one_keeps_original():
    assert _expand(广告素材版本名称="优化组版本-OPDY-RX-7", 广告素材数量=1) == [
        "优化组版本-OPDY-RX-7"
    ]


def test_landing_page_name_equal_to_material_appends_counter():
    # 着陆页名末尾的 7 是产品版本，不是素材序号
    got = _expand(
        着陆页版本名称="优化组版本-OPDY-RX-7",
        广告素材版本名称="优化组版本-OPDY-RX-7",
        广告素材数量=10,
    )
    assert got == [f"优化组版本-OPDY-RX-7-{i}" for i in range(1, 11)]


def test_material_already_has_counter_after_landing_page():
    got = _expand(
        着陆页版本名称="优化组版本-OPDY-RX-7",
        广告素材版本名称="优化组版本-OPDY-RX-7-1",
        广告素材数量=10,
    )
    assert got == [f"优化组版本-OPDY-RX-7-{i}" for i in range(1, 11)]


def test_opdy_rx_2_same_as_workbook1_last_row():
    got = _expand(
        着陆页版本名称="优化组版本-OPDY-RX-2",
        广告素材版本名称="优化组版本-OPDY-RX-2",
        广告素材数量=10,
    )
    assert got[0] == "优化组版本-OPDY-RX-2-1"
    assert got[-1] == "优化组版本-OPDY-RX-2-10"
    assert "优化组版本-OPDY-RX-3" not in got


def test_landing_page_ending_with_one_still_appends():
    got = _expand(
        着陆页版本名称="优化组版本-OPDY-RX-1",
        广告素材版本名称="优化组版本-OPDY-RX-1",
        广告素材数量=2,
    )
    assert got == ["优化组版本-OPDY-RX-1-1", "优化组版本-OPDY-RX-1-2"]


def test_non_digit_tail_unchanged():
    got = _expand(
        着陆页版本名称="优化组版本-OPDY-RX-U260828",
        广告素材版本名称="优化组版本-OPDY-RX-U260828",
        广告素材数量=3,
    )
    assert got == [
        "优化组版本-OPDY-RX-U260828-1",
        "优化组版本-OPDY-RX-U260828-2",
        "优化组版本-OPDY-RX-U260828-3",
    ]


def test_selection_range_classic():
    assert _expand(广告素材版本名称="优化组版本-1", **{"素材选取 (X-Y)": "2-4"}) == [
        "优化组版本-2",
        "优化组版本-3",
        "优化组版本-4",
    ]


def test_selection_range_on_landing_page_name():
    got = _expand(
        着陆页版本名称="优化组版本-OPDY-RX-7",
        广告素材版本名称="优化组版本-OPDY-RX-7",
        **{"素材选取 (X-Y)": "1-3"},
    )
    assert got == [
        "优化组版本-OPDY-RX-7-1",
        "优化组版本-OPDY-RX-7-2",
        "优化组版本-OPDY-RX-7-3",
    ]


def test_url_landing_page_does_not_change_classic_suffix():
    got = _expand(
        着陆页链接="https://example.com/products/x-7",
        广告素材版本名称="优化组版本-1",
        广告素材数量=2,
    )
    assert got == ["优化组版本-1", "优化组版本-2"]
