import pandas as pd

from cross_table_core import build_cross_check_results, build_cross_lookup_results


def test_numeric_targets_still_use_tolerance_and_diff():
    df_a = pd.DataFrame({"工号": ["001", "002"], "金额": [100, 100]})
    df_b = pd.DataFrame({"工号": ["001", "002"], "金额": [100.004, 100.02]})

    summary, diff, full, error = build_cross_check_results(
        df_a,
        "工号",
        [(1, "金额")],
        "sum",
        [("B", df_b, "工号", [(1, "金额")], "sum")],
        0.01,
        True,
    )

    assert error is None
    assert summary is not None
    by_key = full.set_index("键值")
    assert by_key.loc["001", "核对状态"] == "一致"
    assert by_key.loc["002", "核对状态"] == "金额不一致"
    assert by_key.loc["002", "目标数据1-差额(A-B)"] == -0.02
    assert len(diff) == 1


def test_text_targets_compare_text_instead_of_collapsing_to_zero():
    df_a = pd.DataFrame({"工号": ["001", "002"], "状态": ["通过", "失败"]})
    df_b = pd.DataFrame({"工号": ["001", "002"], "状态": ["通过", "通过"]})

    summary, diff, full, error = build_cross_check_results(
        df_a,
        "工号",
        [(1, "状态")],
        "sum",
        [("B", df_b, "工号", [(1, "状态")], "sum")],
        0.01,
        True,
    )

    assert error is None
    by_key = full.set_index("键值")
    assert by_key.loc["001", "核对状态"] == "一致"
    assert by_key.loc["002", "核对状态"] == "值不一致"
    assert by_key.loc["002", "不一致目标"] == "目标数据1"
    assert len(diff) == 1
    assert ("B", "值不一致", 1) in list(summary.itertuples(index=False, name=None))


def test_lookup_text_target_fills_back_original_text_value():
    df_a = pd.DataFrame({"工号": ["001", "002"], "状态": ["待定", "待定"]})
    df_b = pd.DataFrame({"工号": ["001", "002"], "状态": ["通过", "失败"]})

    out, summary, error = build_cross_lookup_results(
        df_a,
        "工号",
        [(1, "状态")],
        [("B", df_b, "工号", [(1, "状态")], "sum")],
        True,
    )

    assert error is None
    assert summary is not None
    by_id = out.set_index("工号")
    assert by_id.loc["001", "状态"] == "通过"
    assert by_id.loc["002", "状态"] == "失败"
    assert by_id.loc["001", "B_来源数据值"] == "通过"


def test_lookup_numeric_reference_stays_numeric_when_a_original_is_text_placeholder():
    df_a = pd.DataFrame({"工号": ["001"], "金额": ["待定"]})
    df_b = pd.DataFrame({"工号": ["001"], "金额": [12.345]})

    out, _, error = build_cross_lookup_results(
        df_a,
        "工号",
        [(1, "金额")],
        [("B", df_b, "工号", [(1, "金额")], "sum")],
        True,
    )

    assert error is None
    assert out.loc[0, "金额"] == 12.34
