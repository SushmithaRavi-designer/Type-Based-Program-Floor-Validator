"""
utils/kpi.py
────────────
KPI calculations for program allocation analysis.
"""

import math
from typing import Dict, List, Tuple


def shannon_diversity(area_by_program: Dict[str, float]) -> float:
    total = sum(area_by_program.values())
    if total <= 0:
        return 0.0
    h = 0.0
    for area in area_by_program.values():
        if area > 0:
            p = area / total
            h -= p * math.log(p)
    return round(h, 4)


def mono_functional_check(
    area_by_program: Dict[str, float],
    thresholds: Dict[str, float],
    default_threshold: float = 70.0,
) -> Tuple[bool, str, float, float]:
    total = sum(area_by_program.values())
    if total <= 0 or not area_by_program:
        return False, "None", 0.0, default_threshold
    dominant = max(area_by_program, key=area_by_program.get)
    pct      = area_by_program[dominant] / total * 100
    allowed  = thresholds.get(dominant, default_threshold)
    return pct > allowed, dominant, round(pct, 2), allowed


def check_zone_compatibility(
    zone_data: Dict[str, Dict[str, float]],
    thresholds: Dict[str, float],
    default_threshold: float = 70.0,
) -> List[str]:
    issues = []
    for zone, prog_areas in zone_data.items():
        total = sum(prog_areas.values())
        if total <= 0:
            continue
        for program, area in prog_areas.items():
            pct     = area / total * 100
            allowed = thresholds.get(program, default_threshold)
            if pct > allowed:
                issues.append(
                    f"Zone {zone} has incompatible program allocation: "
                    f"{program} = {pct:.1f}% (limit = {allowed}%)."
                )
    return issues


def vertical_stacking_continuity(
    floor_data: Dict[str, Dict[str, float]],
) -> Dict[str, float]:
    all_floors = list(floor_data.keys())
    n_floors   = len(all_floors)
    if n_floors == 0:
        return {}
    all_programs: set = set()
    for progs in floor_data.values():
        all_programs.update(progs.keys())
    result = {}
    for program in all_programs:
        present = sum(
            1 for f in all_floors
            if program in floor_data[f] and floor_data[f][program] > 0
        )
        result[program] = round(present / n_floors, 3)
    return result


def floor_summary(
    floor: str,
    prog_areas: Dict[str, float],
    thresholds: Dict[str, float],
    default_threshold: float = 70.0,
) -> dict:
    total = sum(prog_areas.values())
    is_mono, dominant, dom_pct, allowed = mono_functional_check(
        prog_areas, thresholds, default_threshold
    )
    return {
        "floor":              floor,
        "total_area":         round(total, 2),
        "num_programs":       len(prog_areas),
        "dominant_program":   dominant,
        "dominant_pct":       dom_pct,
        "allowed_pct":        allowed,
        "is_mono_functional": is_mono,
        "diversity_index":    shannon_diversity(prog_areas),
    }
