from __future__ import annotations

from collections.abc import Iterable

import pandas as pd


def norm_id(v):
    if pd.isna(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def to_num(v):
    if pd.isna(v):
        return 0.0
    if isinstance(v, str):
        s = v.strip().replace(",", "")
        if not s:
            return 0.0
        try:
            return float(s)
        except ValueError:
            return 0.0
    return float(v)


def text_value(v):
    if pd.isna(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def list_vals(s: pd.Series) -> str:
    return " | ".join(text_value(v) for v in s.dropna())


def _is_blank(v) -> bool:
    if pd.isna(v):
        return True
    return isinstance(v, str) and not v.strip()


def _can_parse_number(v) -> bool:
    if _is_blank(v):
        return True
    if isinstance(v, str):
        v = v.strip().replace(",", "")
    try:
        float(v)
    except (TypeError, ValueError):
        return False
    return True


def _iter_nonblank_values(series_list: Iterable[pd.Series]):
    for series in series_list:
        for value in series.dropna():
            if not _is_blank(value):
                yield value


def infer_target_kind(*series_list: pd.Series) -> str:
    values = list(_iter_nonblank_values(series_list))
    if not values:
        return "number"
    return "number" if all(_can_parse_number(v) for v in values) else "text"


def infer_lookup_target_kind(base_series: pd.Series, ref_series: pd.Series) -> str:
    ref_values = list(_iter_nonblank_values([ref_series]))
    if ref_values:
        return "number" if all(_can_parse_number(v) for v in ref_values) else "text"
    return infer_target_kind(base_series, ref_series)


def _text_agg_mode(mode: str) -> str:
    return mode if mode in {"first", "last", "max", "min"} else "first"


def validate_targets(kA: str, a_pairs: list[tuple[int, str]], ref_label: str,
                     kR: str, r_pairs: list[tuple[int, str]], lookup: bool = False):
    target_map = {target_no: v for target_no, v in a_pairs}
    if not kA or not kR or not a_pairs or not r_pairs:
        return None, "请确认关联相同字段和目标数据都已选择。"
    if len(target_map) != len(a_pairs):
        return None, "A 表目标数据编号不能重复。"
    ref_numbers = [target_no for target_no, _ in r_pairs]
    if len(set(ref_numbers)) != len(ref_numbers):
        return None, f"{ref_label} 表目标数据编号不能重复。"
    missing_numbers = [target_no for target_no in ref_numbers if target_no not in target_map]
    if missing_numbers:
        return None, f"{ref_label} 表包含 A 表不存在的目标数据编号：{missing_numbers}"
    if lookup and len({v for _, v in a_pairs}) != len(a_pairs):
        return None, "查询模式下 A 表目标数据不能重复，否则回填列会冲突。"
    return [(target_no, target_map[target_no], val) for target_no, val in r_pairs], None


def validate_reference_allocations(refs: list[tuple[str, pd.DataFrame, str, list[tuple[int, str]], str]]):
    owners: dict[int, list[str]] = {}
    for ref_label, _, _, r_pairs, _ in refs:
        for target_no, _ in r_pairs:
            owners.setdefault(target_no, []).append(ref_label)
    duplicates = {target_no: labels for target_no, labels in owners.items() if len(labels) > 1}
    if duplicates:
        detail = "；".join(f"目标数据{target_no}: {','.join(labels)}" for target_no, labels in duplicates.items())
        return f"多个参考表不能占用同一个目标数据编号：{detail}"
    return None


def _target_mismatches(row, specs: list[dict], tol: float) -> list[dict]:
    mismatches = []
    for spec in specs:
        if spec["kind"] == "number":
            if abs(row[spec["diff_col"]]) > tol:
                mismatches.append(spec)
        elif not row[spec["equal_col"]]:
            mismatches.append(spec)
    return mismatches


def _mismatch_status(mismatches: list[dict]) -> str:
    if not mismatches:
        return "一致"
    if all(spec["kind"] == "number" for spec in mismatches):
        return "金额不一致"
    return "值不一致"


def build_cross_check_results(dfA: pd.DataFrame, kA: str, a_pairs: list[tuple[int, str]], aggA: str,
                              refs: list[tuple[str, pd.DataFrame, str, list[tuple[int, str]], str]],
                              tol: float, norm: bool):
    allocation_error = validate_reference_allocations(refs)
    if allocation_error:
        return None, None, None, allocation_error

    A = dfA[dfA[kA].notna()].copy()
    A["__k"] = A[kA].map(norm_id) if norm else A[kA]
    full_parts = []
    summary_rows = [("全部", "参考表数量", len(refs)), ("全部", "A表行数", len(A))]

    for ref_label, dfR, kR, r_pairs, aggR in refs:
        target_pairs, error = validate_targets(kA, a_pairs, ref_label, kR, r_pairs)
        if error:
            return None, None, None, error
        R = dfR[dfR[kR].notna()].copy()
        R["__k"] = R[kR].map(norm_id) if norm else R[kR]
        agg_a = {}
        agg_r = {}
        specs = []
        target_labels = {}
        for target_no, vA, vR in target_pairs:
            kind = infer_target_kind(A[vA], R[vR])
            a_src = f"__vA_{target_no}"
            r_src = f"__vR_{target_no}"
            a_col = f"目标数据{target_no}-A值"
            r_col = f"目标数据{target_no}-{ref_label}值"
            diff_col = f"目标数据{target_no}-差额(A-{ref_label})"
            equal_col = f"目标数据{target_no}-是否一致"
            if kind == "number":
                A[a_src] = A[vA].map(to_num)
                R[r_src] = R[vR].map(to_num)
                agg_a[a_col] = (a_src, aggA)
                agg_r[r_col] = (r_src, aggR)
            else:
                A[a_src] = A[vA].map(text_value)
                R[r_src] = R[vR].map(text_value)
                agg_a[a_col] = (a_src, _text_agg_mode(aggA))
                agg_r[r_col] = (r_src, _text_agg_mode(aggR))
            spec = {
                "target_no": target_no,
                "kind": kind,
                "a_col": a_col,
                "r_col": r_col,
                "diff_col": diff_col,
                "equal_col": equal_col,
                "vA": vA,
                "vR": vR,
            }
            specs.append(spec)
            target_labels[target_no] = f"目标数据{target_no}"

        Ag = A.groupby("__k", as_index=False).agg(**agg_a)
        Rg = R.groupby("__k", as_index=False).agg(**agg_r)
        merged = Ag.merge(Rg, on="__k", how="outer", indicator=True)
        for spec in specs:
            if spec["kind"] == "number":
                merged[spec["a_col"]] = merged[spec["a_col"]].fillna(0).round(2)
                merged[spec["r_col"]] = merged[spec["r_col"]].fillna(0).round(2)
                merged[spec["diff_col"]] = (merged[spec["a_col"]] - merged[spec["r_col"]]).round(2)
            else:
                merged[spec["a_col"]] = merged[spec["a_col"]].fillna("").map(text_value)
                merged[spec["r_col"]] = merged[spec["r_col"]].fillna("").map(text_value)
                merged[spec["diff_col"]] = pd.NA
                merged[spec["equal_col"]] = merged[spec["a_col"]] == merged[spec["r_col"]]

        def cls(row):
            if row["_merge"] == "left_only":
                return "仅A有"
            if row["_merge"] == "right_only":
                return f"仅{ref_label}有(A遗漏)"
            return _mismatch_status(_target_mismatches(row, specs, tol))

        merged["核对状态"] = merged.apply(cls, axis=1)
        merged["不一致目标"] = merged.apply(
            lambda row: "" if row["核对状态"] in {"一致", "仅A有", f"仅{ref_label}有(A遗漏)"} else "、".join(
                target_labels[spec["target_no"]] for spec in _target_mismatches(row, specs, tol)
            ),
            axis=1,
        )
        merged = merged.drop(columns=["_merge"]).rename(columns={"__k": "键值"})
        merged.insert(0, "参考表", ref_label)
        full_parts.append(merged)

        cnt = merged["核对状态"].value_counts().to_dict()
        summary_rows.extend([
            (ref_label, "目标数据组数", len(target_pairs)),
            (ref_label, "合集", len(merged)),
            (ref_label, "一致", cnt.get("一致", 0)),
            (ref_label, "金额不一致", cnt.get("金额不一致", 0)),
            (ref_label, "值不一致", cnt.get("值不一致", 0)),
            (ref_label, "仅A有", cnt.get("仅A有", 0)),
            (ref_label, f"仅{ref_label}有(A遗漏)", cnt.get(f"仅{ref_label}有(A遗漏)", 0)),
            (ref_label, f"A表({aggA})行数", len(A)),
            (ref_label, f"{ref_label}表({aggR})行数", len(R)),
        ])
        for spec in specs:
            summary_rows.extend([
                (ref_label, f"目标数据{spec['target_no']} A列", spec["vA"]),
                (ref_label, f"目标数据{spec['target_no']} {ref_label}列", spec["vR"]),
                (ref_label, f"目标数据{spec['target_no']} 类型", "数字" if spec["kind"] == "number" else "文字"),
            ])
            if spec["kind"] == "number":
                summary_rows.extend([
                    (ref_label, f"目标数据{spec['target_no']} A合计", round(merged[spec["a_col"]].sum(), 2)),
                    (ref_label, f"目标数据{spec['target_no']} {ref_label}合计", round(merged[spec["r_col"]].sum(), 2)),
                    (ref_label, f"目标数据{spec['target_no']} 差额合计", round(merged[spec["diff_col"]].sum(), 2)),
                ])

    full = pd.concat(full_parts, ignore_index=True) if full_parts else pd.DataFrame()
    diff = full[full["核对状态"] != "一致"].reset_index(drop=True)
    summary = pd.DataFrame(summary_rows, columns=["参考表", "指标", "值"])
    return summary, diff, full, None


def build_cross_lookup_results(dfA: pd.DataFrame, kA: str, a_pairs: list[tuple[int, str]],
                               refs: list[tuple[str, pd.DataFrame, str, list[tuple[int, str]], str]],
                               norm: bool):
    allocation_error = validate_reference_allocations(refs)
    if allocation_error:
        return None, None, allocation_error

    A = dfA.copy()
    A["__k"] = A[kA].map(norm_id) if norm else A[kA]
    valsA = [v for _, v in a_pairs]
    if len(set(valsA)) != len(valsA):
        return None, None, "查询模式下 A 表目标数据不能重复，否则回填列会冲突。"

    merged = A.copy()
    pair_meta_by_target: dict[int, list[tuple[str, str, str, str, str]]] = {}
    source_cols = []
    record_cols = []
    ref_match_rows = []

    for ref_label, dfR, kR, r_pairs, aggR in refs:
        target_pairs, error = validate_targets(kA, a_pairs, ref_label, kR, r_pairs, lookup=True)
        if error:
            return None, None, error
        R = dfR[dfR[kR].notna()].copy()
        R["__k"] = R[kR].map(norm_id) if norm else R[kR]
        record_col = f"__{ref_label}来源记录数"
        agg_spec = {record_col: ("__k", "count")}
        for target_no, vA, vR in target_pairs:
            kind = infer_lookup_target_kind(A[vA], R[vR])
            value_col = f"__{ref_label}查询值{target_no}"
            source_col = f"__{ref_label}来源数据值{target_no}"
            if kind == "number":
                R[value_col] = R[vR].map(to_num)
                agg_spec[value_col] = (value_col, aggR)
            else:
                R[value_col] = R[vR].map(text_value)
                agg_spec[value_col] = (value_col, _text_agg_mode(aggR))
            agg_spec[source_col] = (vR, list_vals)
            pair_meta_by_target.setdefault(target_no, []).append((ref_label, vR, value_col, source_col, kind))
            source_cols.append((ref_label, target_no, source_col))
        Rg = R.groupby("__k", as_index=False).agg(**agg_spec)
        merged = merged.merge(Rg, on="__k", how="left")
        record_cols.append((ref_label, record_col))
        ref_match_rows.append((ref_label, aggR, kR, record_col))

    def matched_refs(row):
        names = [
            ref_label for ref_label, record_col in record_cols
            if pd.notna(row[record_col]) and row[record_col] > 0
        ]
        return "、".join(names)

    merged["匹配参考表"] = merged.apply(matched_refs, axis=1)
    merged["匹配状态"] = merged["匹配参考表"].apply(lambda x: "未匹配" if not x else "已匹配")

    orig_cols = []
    target_kinds = {}
    for target_no, vA in a_pairs:
        orig_col = f"{vA}_原值" if len(a_pairs) == 1 else f"{vA}_原值(目标数据{target_no})"
        merged[orig_col] = merged[vA]
        fill_values = None
        kinds = []
        for _, _, value_col, _, kind in pair_meta_by_target.get(target_no, []):
            kinds.append(kind)
            fill_values = merged[value_col] if fill_values is None else fill_values.combine_first(merged[value_col])
        if fill_values is not None:
            merged[vA] = fill_values.where(fill_values.notna(), merged[orig_col])
        target_kinds[target_no] = "text" if "text" in kinds else "number"
        orig_cols.append(orig_col)

    source_out_cols = []
    rename_cols = {}
    for ref_label, target_no, source_col in source_cols:
        out_source_col = f"{ref_label}_来源数据值" if len(a_pairs) == 1 else f"{ref_label}_来源数据值(目标数据{target_no})"
        source_out_cols.append(source_col)
        rename_cols[source_col] = out_source_col
    record_out_cols = []
    for ref_label, record_col in record_cols:
        record_out_cols.append(record_col)
        rename_cols[record_col] = f"{ref_label}_来源记录数"
    out_cols = (
        [c for c in A.columns if c != "__k"]
        + orig_cols
        + source_out_cols
        + record_out_cols
        + ["匹配参考表", "匹配状态"]
    )
    out = merged[out_cols].rename(columns=rename_cols)
    for target_no, vA in a_pairs:
        if target_kinds.get(target_no) == "number" and pd.api.types.is_numeric_dtype(out[vA]):
            out[vA] = out[vA].round(2)

    matched = int((out["匹配状态"] == "已匹配").sum())
    total = len(out)
    summary_rows = [
        ("全部", "参考表数量", len(refs)),
        ("全部", "A表目标数据组数", len(a_pairs)),
        ("全部", "A表总行数", total),
        ("全部", "已匹配(任一参考表有数据)", matched),
        ("全部", "未匹配(所有参考表无数据)", total - matched),
    ]
    for ref_label, aggR, kR, record_col in ref_match_rows:
        ref_matched = int((merged[record_col].fillna(0) > 0).sum())
        summary_rows.extend([
            (ref_label, f"{ref_label}聚合方式", aggR),
            (ref_label, f"关联字段(A→{ref_label})", f"{kA} ↔ {kR}"),
            (ref_label, f"{ref_label}匹配行数", ref_matched),
        ])
    for (target_no, vA), orig_col in zip(a_pairs, orig_cols):
        summary_rows.extend([
            ("全部", f"目标数据{target_no} A列", vA),
            ("全部", f"目标数据{target_no} 类型", "数字" if target_kinds.get(target_no) == "number" else "文字"),
        ])
        if target_kinds.get(target_no) == "number":
            before = pd.to_numeric(out[orig_col], errors="coerce").fillna(0).sum()
            after = pd.to_numeric(out[vA], errors="coerce").fillna(0).sum()
            summary_rows.extend([
                ("全部", f"目标数据{target_no} A原合计", round(before, 2)),
                ("全部", f"目标数据{target_no} 回填后合计", round(after, 2)),
                ("全部", f"目标数据{target_no} 变化量", round(after - before, 2)),
            ])
    summary = pd.DataFrame(summary_rows, columns=["参考表", "指标", "值"])
    return out, summary, None
