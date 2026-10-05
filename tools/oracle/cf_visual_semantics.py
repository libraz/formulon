"""Documented-semantics match payloads for DataBar and IconSet rules.

Excel does not expose these shapes through DisplayFormat, so the oracle
generator computes them from the Microsoft VBA object-model semantics.
Pure functions over the case JSON; no Excel or openpyxl needed.
"""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Tuple

# OOXML icon-set tag <-> Formulon `IconSetName` enum string. The OOXML
# attribute (`<iconSet iconSet="3Arrows">`) is what openpyxl emits and
# what Excel reads; the second column is the name the Formulon CF
# evaluator emits in `IconRender::set_name`. Mirrors src/cf/cf_types.h
# and src/io/cf_reader.cpp.
_ICON_SET_OOXML_TO_FORMULON: Dict[str, str] = {
    "3Arrows": "Three_Arrows",
    "3ArrowsGray": "Three_ArrowsGray",
    "3Flags": "Three_Flags",
    "3TrafficLights1": "Three_TrafficLights1",
    "3TrafficLights2": "Three_TrafficLights2",
    "3Signs": "Three_Signs",
    "3Symbols": "Three_Symbols",
    "3Symbols2": "Three_Symbols2",
    "4Arrows": "Four_Arrows",
    "4ArrowsGray": "Four_ArrowsGray",
    "4RedToBlack": "Four_RedToBlack",
    "4Rating": "Four_Rating",
    "4TrafficLights": "Four_TrafficLights",
    "5Arrows": "Five_Arrows",
    "5ArrowsGray": "Five_ArrowsGray",
    "5Rating": "Five_Rating",
    "5Quarters": "Five_Quarters",
}

_ICON_SET_FORMULON_TO_OOXML: Dict[str, str] = {v: k for k, v in _ICON_SET_OOXML_TO_FORMULON.items()}

# Number of buckets for an icon set, keyed by the OOXML tag's leading
# digit. Used to bound the per-cell icon index after walking thresholds.
_ICON_SET_BUCKET_COUNT: Dict[str, int] = {tag: int(tag[0]) for tag in _ICON_SET_OOXML_TO_FORMULON}


def _cell_in_block(row: int, col: int, sqref: List[Dict[str, int]]) -> bool:
    for r in sqref:
        if r["first_row"] <= row <= r["last_row"] and r["first_col"] <= col <= r["last_col"]:
            return True
    return False


def _numeric_population(case: Dict[str, Any], sqref: List[Dict[str, int]]) -> List[float]:
    """Returns the sorted ascending list of numeric cell values inside
    `sqref` for the given `case`. Mirrors the C++ evaluator's
    `collect_numeric_values` helper: booleans, text, and errors are
    skipped; only `kind == "number"` cells contribute. Cells outside
    `sqref` are ignored; cells in `sqref` that are absent from the
    `sheet` array are treated as blank (and skipped).
    """
    values: List[float] = []
    for entry in case["sheet"]:
        if entry.get("kind") != "number":
            continue
        if not _cell_in_block(int(entry["row"]), int(entry["col"]), sqref):
            continue
        values.append(float(entry["value"]))
    values.sort()
    return values


def _resolve_threshold_value(
    threshold_type: str,
    threshold_value: Optional[Any],
    population: List[float],
) -> Optional[float]:
    """Resolves a single CFVO to its numeric threshold. Mirrors
    `cf_evaluator.cpp::resolve_cfvo` for the threshold types the
    generator currently emits (Number / Percent / Percentile / Min /
    Max). Returns `None` when the population is empty and the
    threshold needs it (Min / Max / Percent / Percentile).
    """
    if threshold_type == "Number":
        return float(threshold_value) if threshold_value is not None else None
    if not population:
        return None
    pop_min = population[0]
    pop_max = population[-1]
    if threshold_type == "Min":
        return pop_min
    if threshold_type == "Max":
        return pop_max
    if threshold_type == "Percent":
        if threshold_value is None:
            return None
        pct = float(threshold_value)
        return pop_min + (pct / 100.0) * (pop_max - pop_min)
    if threshold_type == "Percentile":
        if threshold_value is None:
            return None
        pct = float(threshold_value) / 100.0
        count = len(population)
        if count == 1:
            return population[0]
        position = pct * (count - 1)
        lower_index = int(position)
        if lower_index + 1 >= count:
            return population[-1]
        fraction = position - lower_index
        return population[lower_index] + fraction * (population[lower_index + 1] - population[lower_index])
    return None


def _data_bar_match_for_cell(
    cell_value: Optional[float],
    desc: Dict[str, Any],
    population: List[float],
) -> Optional[Dict[str, Any]]:
    """Computes the DataBar match payload for a numeric cell value
    using the documented Microsoft VBA semantics:
    https://learn.microsoft.com/en-us/office/vba/api/excel.databar
    Bar length is the linear fraction `(value - min) / (max - min)`,
    clamped to [0, 1], scaled into [min_length_pct, max_length_pct].
    Negative values are not handled by these smoke cases (see
    README.md "Not yet captured"); `bar_axis_position_pct` is fixed at
    0 and `is_negative` at false.
    """
    if cell_value is None:
        return None
    db = desc["data_bar"]
    threshold_min = _resolve_threshold_value(db["min_type"], db["min_value"], population)
    threshold_max = _resolve_threshold_value(db["max_type"], db["max_value"], population)
    if threshold_min is None or threshold_max is None:
        return None
    if threshold_max == threshold_min:
        return None
    min_len = float(db["min_length_pct"])
    max_len = float(db["max_length_pct"])
    fraction = (cell_value - threshold_min) / (threshold_max - threshold_min)
    fraction = max(0.0, min(1.0, fraction))
    bar_length = min_len + fraction * (max_len - min_len)
    fill = db["fill"]
    return {
        "kind": "DataBar",
        "priority": desc["priority"],
        "bar_length_pct": bar_length,
        "bar_axis_position_pct": 0.0,
        "is_negative": False,
        "fill": {
            "r": int(fill["r"]),
            "g": int(fill["g"]),
            "b": int(fill["b"]),
            "a": int(fill.get("a", 255)),
        },
    }


def _icon_set_match_for_cell(
    cell_value: Optional[float],
    desc: Dict[str, Any],
    population: List[float],
) -> Optional[Dict[str, Any]]:
    """Computes the IconSet match payload for a numeric cell value
    using the documented Microsoft VBA semantics:
    https://learn.microsoft.com/en-us/office/vba/api/excel.iconsetcondition
    Walks thresholds in ascending order, bumping the bucket index for
    each one the cell value crosses (`>=` if `gte`; `>` otherwise).
    Applies `reverse` last by mirroring the index across the bucket
    range.
    """
    if cell_value is None:
        return None
    spec = desc["icon_set"]
    resolved: List[Optional[float]] = []
    for t in spec["thresholds"]:
        resolved.append(_resolve_threshold_value(t["type"], t["value"], population))
    if any(value is None for value in resolved):
        return None
    icon_index = 0
    for i, threshold in enumerate(resolved):
        gte = bool(spec["thresholds"][i]["gte"])
        if (cell_value >= threshold) if gte else (cell_value > threshold):
            icon_index = i + 1
    bucket_count = _ICON_SET_BUCKET_COUNT.get(spec["name_ooxml"], len(resolved) + 1)
    if spec["reverse"]:
        icon_index = bucket_count - 1 - icon_index
    return {
        "kind": "IconSet",
        "priority": desc["priority"],
        "icon_set_name": spec["name_formulon"],
        "icon_index": icon_index,
    }


def _cell_value_for(case: Dict[str, Any], row: int, col: int) -> Optional[float]:
    """Returns the numeric value of `(row, col)` in `case['sheet']`, or
    `None` when the cell is blank, non-numeric, or absent.
    """
    for entry in case["sheet"]:
        if int(entry["row"]) == row and int(entry["col"]) == col:
            if entry.get("kind") == "number":
                return float(entry["value"])
            return None
    return None


def _compute_visual_matches(
    case: Dict[str, Any], descriptors: List[Dict[str, Any]]
) -> Dict[Tuple[int, int], List[Dict[str, Any]]]:
    """Computes per-cell DataBar / IconSet matches for every descriptor
    that carries the corresponding payload. Returns a row/col-keyed map
    of match-dict lists; cells with no visual rule applied are absent.

    Excel does not expose these match shapes through DisplayFormat, so
    the generator computes them from documented semantics. Every cell
    inside the rule's sqref produces a match (visual rules apply to
    every cell in their range — there is no boolean "fired or not"
    decision the way DifferentialFormat-bearing rules have).
    """
    out: Dict[Tuple[int, int], List[Dict[str, Any]]] = {}
    rng = case["range"]
    r1 = int(rng["first_row"])
    c1 = int(rng["first_col"])
    r2 = int(rng["last_row"])
    c2 = int(rng["last_col"])
    sorted_descriptors = sorted(descriptors, key=lambda d: d["priority"])
    for desc in sorted_descriptors:
        if desc["type"] not in ("DataBar", "IconSet"):
            continue
        population = _numeric_population(case, desc["sqref"])
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                if not _cell_in_block(r, c, desc["sqref"]):
                    continue
                cell_val = _cell_value_for(case, r, c)
                if desc["type"] == "DataBar":
                    match = _data_bar_match_for_cell(cell_val, desc, population)
                else:
                    match = _icon_set_match_for_cell(cell_val, desc, population)
                if match is None:
                    continue
                out.setdefault((r, c), []).append(match)
    return out
