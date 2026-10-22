"""把 Sheet4 的模型程序化，并用多轮贪心推导「把候选 Caster 全部练到绊 9」的完整方案。

单轮模型（与 caster_bond.xlsx 的 Sheet4 一致）：
    后排固定区 1.74 = 1 + 全队自带 0.04 + 前排礼装 0.70
    前排再加站位 0.20，所以孔明、山鲁佐德是 1.94
    两人冲绊 15，各自满了之后全队再 +0.20（孔明满：后排 1.94 / 前排 2.14；两人都满：后排 2.14 / 前排 2.34）
    L_i = K_i × 4748 × (固定区 + n5 × 0.05 + Σ 命中_ij × 启用_j × 系数_j)
    M_i = 距绊 9 的剩余缺口
    N3  = ROUNDUP(1 / MAX(L_i / M_i))        由队里毕业最快的那位决定
    P3  = Σ L_i / (3 × 4748 × 2.29)          归一化总收益
    R3  = (命中户数 − 已利用户数) / (命中户数 + 0.01)
    loss = N3 − P3 + R3                      → 越小越好

约束：每轮最多 3 人、礼装槽 2 个（6 件 +20% 与 2 个 +5% 占位共同竞争）、
      SUM(cost) + 80 ≤ 113、SUM(cost) = 21 时额外奖励第 3 件 +5% 礼装。

多轮贪心：每轮在「尚未毕业」的人里重跑单轮模型，执行 N3 场，按 L_i × N3 扣减缺口，
缺口 ≤ 0 的毕业移出，直到全员毕业。两套准则可对比：
    loss        → Sheet4 原定义（尽快让队里最快的人毕业）
    throughput  → 每场推进的有效缺口最大（总场次最短）

用法：
    python caster_plan.py                        # 读 Sheet4，生成 caster_plan.html
    python caster_plan.py --rule throughput      # 只跑某一套准则
    python caster_plan.py --json plan.json       # 附带机器可读输出
"""

from __future__ import annotations

import argparse
import html
import json
import math
import zipfile
from itertools import combinations
from pathlib import Path
import xml.etree.ElementTree as ET

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_MAP = {"m": NS[1:-1]}
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

BASE_BOND = 4748.0                                     # 单场基础绊
BACK_FIXED = 1 + 0.04 + 0.7                            # 1.74：1 + 全队自带 4% + 前排礼装 70%
FRONT_POSITION = 0.2                                   # 站前排额外 20%，只有孔明和山鲁佐德
FRONT_FIXED = BACK_FIXED + FRONT_POSITION              # 1.94
FIXED_BONUS = BACK_FIXED                               # 三人段用后排固定区
IDEAL_BONUS = BACK_FIXED + 0.05 + 0.4                  # P3 归一化用的理想单人系数
COST_FIXED = 80                                        # 队伍其余成员的 cost（含 2 件礼装）
COST_BASE = 56                                         # 前三排固定占用：16×2 + 12×2
COST_CAP = 113                                         # 队伍 cost 上限
EQUIP_COST = 12                                        # 12 cost 的后排礼装：具名 +20%，或泛用 +5%
CHEAP_COST = 9                                         # 9 cost 的泛用后排礼装
CHEAP_BONUS = 0.02                                     # 9 cost 礼装的全队加成
SLOTS = 2                                              # 具名 +20% 礼装最多两件
MAX_BACK_CE = 3                                        # 后排礼装最多三件
NAMED_CE = 6                                           # 前 6 列是具名 +20%
N5_COST = 21                                           # 三个 7 cost 从者的 cost 和
N5_BONUS = 0.05                                        # 12 cost 泛用礼装的全队加成
AP_PER_RUN = 40                                        # 每场 AP
AP_PER_APPLE = 146                                     # 每个金苹果换算 AP
P3_WEIGHT = 1.0                                        # loss 里 P3 的权重（--p3-weight 可覆盖）
R3_WEIGHT = 0.2                                        # loss 里 R3 的权重：户口浪费只作轻微修正
RUNNER_BONUS = 0.2                                     # 孔明或山鲁佐德到绊 15，全队 +20%（含两人本人）
RUNNERS = [                                            # 前排陪跑：不参与三人决策
    {"name": "诸葛孔明", "need": 6_531_154, "hits": [0, 0, 0, 0, 0, 1, 1, 1]},   # 迦勒底之晨 + 两个占位
    {"name": "山鲁佐德", "need": 8_238_356, "hits": [0, 0, 0, 1, 0, 0, 1, 1]},   # 献给幸福的新娘 + 两个占位
]
MAX_ROUNDS = 400
COST_FIRST = False                                     # 同等吞吐下优先上高 cost 从者，把低 cost 留给后面配对

DEFAULT_EQUIP_LABELS = [
    "异星之神", "迦勒底之人", "检查报告", "献给幸福的新娘",
    "至诚的一针", "迦勒底之晨", "占位1", "占位2",
]

RULES = {
    "loss": ("Sheet4 原始准则", "每轮取 loss = N3 − P3 + R3 最小的组合：优先让队里最快的人毕业。"),
    "throughput": ("吞吐优先", "每轮取「每场推进的有效缺口」最大的组合：溢出更少，总场次更短。"),
}


def column_index(ref: str) -> int:
    letters = "".join(ch for ch in ref if ch.isalpha())
    index = 0
    for ch in letters:
        index = index * 26 + (ord(ch) - 64)
    return index - 1


def load_sheet_cells(path: Path, sheet_name: str) -> dict[tuple[int, int], object]:
    """读取指定 sheet 的单元格（共享字符串 + 数字），返回 {(row, col): value}。"""
    with zipfile.ZipFile(path) as archive:
        shared: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            sst = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for si in sst.findall("m:si", NS_MAP):
                shared.append("".join(t.text or "" for t in si.iter(NS + "t")))

        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rel_target = {r.get("Id"): r.get("Target") for r in rels}
        target = None
        for sheet in workbook.findall(".//m:sheet", NS_MAP):
            if sheet.get("name") == sheet_name:
                rid = sheet.get(f"{REL_NS}id")
                raw = rel_target.get(rid, "")
                target = raw[1:] if raw.startswith("/") else f"xl/{raw}"
                break
        if not target:
            raise SystemExit(f"{path} 里找不到 sheet: {sheet_name}")

        root = ET.fromstring(archive.read(target))
        cells: dict[tuple[int, int], object] = {}
        for row in root.findall("m:sheetData/m:row", NS_MAP):
            row_no = int(row.get("r"))
            for cell in row.findall("m:c", NS_MAP):
                kind = cell.get("t")
                value = cell.find("m:v", NS_MAP)
                if value is None:
                    continue
                text = value.text
                if kind == "s":
                    cells[(row_no, column_index(cell.get("r")))] = shared[int(text)]
                else:
                    try:
                        cells[(row_no, column_index(cell.get("r")))] = float(text)
                    except (TypeError, ValueError):
                        cells[(row_no, column_index(cell.get("r")))] = text
        return cells


def load_inputs(path: Path, sheet_name: str) -> dict:
    cells = load_sheet_cells(path, sheet_name)
    coefficients = [float(cells[(5, c)]) for c in range(2, 10)]
    data_rows = sorted(row for (row, col), value in cells.items() if col == 0 and row >= 8 and value)
    people = []
    for row in data_rows:
        people.append(
            {
                "name": str(cells[(row, 0)]),
                "cost": int(cells.get((row, 1), 0)),
                "hits": [int(cells.get((row, c), 0)) for c in range(2, 10)],
                "need": int(cells.get((row, 12), 0)),
            }
        )
    if not people:
        raise SystemExit(f"{sheet_name} 里没有读到候选从者（A8:A32 为空）")
    return {
        "people": people,
        "coefficients": coefficients,
        "labels": DEFAULT_EQUIP_LABELS[: len(coefficients)],
    }


def ce_loadouts() -> list[dict]:
    """从 12 件里选最多 3 件。相同泛用礼装是重复副本，收益相同的组合只留一份。

    6 件 12 cost 户口 +20%，3 件 12 cost 全队 +5%，3 件 9 cost 全队 +2%。
    C(25,3)×C(12,3)=506000 是只上满 3 人 3 礼装时的上限；少上的组合更小。
    """
    pool = [("named", j) for j in range(NAMED_CE)]
    pool += [("five", i) for i in range(3)]
    pool += [("two", i) for i in range(3)]
    seen = {}
    for count in range(MAX_BACK_CE + 1):
        for combo in combinations(range(len(pool)), count):
            named = []
            five = two = 0
            for index in combo:
                kind, value = pool[index]
                if kind == "named":
                    named.append(value)
                elif kind == "five":
                    five += 1
                else:
                    two += 1
            enabled = tuple(1 if j in named else 0 for j in range(NAMED_CE))
            key = (enabled, five, two)
            if key in seen:
                continue
            seen[key] = {
                "enabled": enabled,
                "five": five,
                "two": two,
                "equip_cost": EQUIP_COST * (len(named) + five) + CHEAP_COST * two,
                "flat": N5_BONUS * five + CHEAP_BONUS * two,
            }
    return list(seen.values())


def ce_bonus(hits: list[int], loadout: dict, coef: list[float]) -> float:
    enabled = loadout["enabled"]
    named = sum(hits[j] * enabled[j] * coef[j] for j in range(NAMED_CE))
    return named + loadout["flat"]


def format_equips(item: dict, labels: list[str], coef: list[float]) -> str:
    enabled = item["enabled"]
    parts = [
        f"{labels[j]}(+{coef[j] * 100:g}%)" for j in range(NAMED_CE) if j < len(enabled) and enabled[j]
    ]
    if item.get("five"):
        parts.append(f"+5%×{item['five']}（12 cost）")
    if item.get("two"):
        parts.append(f"+2%×{item['two']}（9 cost）")
    return "、".join(parts) or "—"


def evaluate(chosen: list[int], loadout: dict, people: list[dict], coef: list[float],
             demands: list[int], fixed_bonus: float = FIXED_BONUS) -> dict | None:
    cost_sum = sum(people[i]["cost"] for i in chosen)
    # 从者与礼装共用同一份 cost 额度：前排固定 56，剩下 113-56=57 给「后排从者 + 后排礼装」。
    equip_cost = loadout["equip_cost"]
    budget = COST_CAP - COST_BASE
    if cost_sum + equip_cost > budget:
        return None
    gains, ratios = [], []
    for i in chosen:
        gain = BASE_BOND * (fixed_bonus + ce_bonus(people[i]["hits"], loadout, coef))
        gains.append(gain)
        ratios.append(gain / demands[i] if demands[i] > 0 else 0.0)
    slowest = max(ratios)
    runs = math.ceil(1 / slowest) if slowest else 0
    ideal = fixed_bonus + 0.05 + 0.4
    p3 = sum(gains) / (3 * BASE_BOND * ideal)
    enabled = loadout["enabled"]
    hit20 = sum(people[i]["hits"][j] for i in chosen for j in range(NAMED_CE))
    used20 = sum(people[i]["hits"][j] * enabled[j] for i in chosen for j in range(NAMED_CE))
    r3 = (hit20 - used20) / (hit20 + 0.01) if hit20 else 0.0
    pushed = sum(min(gain * runs, demands[i]) for gain, i in zip(gains, chosen))
    return {
        "loss": runs - P3_WEIGHT * p3 + R3_WEIGHT * r3,
        "runs": runs,
        "p3": p3,
        "r3": r3,
        "throughput": pushed / runs if runs else 0.0,
        "pushed": pushed,
        "cost": cost_sum,
        "equip_count": sum(enabled) + loadout["five"] + loadout["two"],
        "equip_cost": equip_cost,
        "total_cost": COST_BASE + cost_sum + equip_cost,
        "n5": 0,
        "five": loadout["five"],
        "two": loadout["two"],
        "flat": loadout["flat"],
        "gains": gains,
        "enabled": enabled,
        "chosen": chosen,
    }


def equip_combinations(count: int) -> list[list[int]]:
    """枚举 0/1/2 件礼装。

    礼装与从者共用 cost 额度，所以「少挂一件、多上一个人」是真实可选项：
    2 人 + 2 件打不过 3 人 + 1 件时，就该省下那 12 cost。
    """
    combos = [()]
    combos += [(j,) for j in range(count)]
    combos += list(combinations(range(count), SLOTS))
    return [[1 if j in combo else 0 for j in range(count)] for combo in combos]


def max_feasible_size(people: list[dict], alive: list[int]) -> int:
    """当前存活集合在 cost 上限下最多能同时上几个人（按不挂礼装的最宽松情形算）。"""
    for size in range(min(3, len(alive)), 0, -1):
        for chosen in combinations(alive, size):
            if sum(people[i]["cost"] for i in chosen) + COST_FIXED <= COST_CAP:
                return size
    return 0


def solve_round(people: list[dict], alive: list[int], coef: list[float], demands: list[int],
                rule: str, fixed_bonus: float = FIXED_BONUS) -> dict | None:
    """在存活集合里枚举「人数 × 礼装组合」，按 rule 取最优的一步。

    人数从 3 递减到 1：高 cost 的人凑不出 3 人时（SUM(cost) + 80 > 113）自动降档。
    """
    best = None
    best_key = None
    loadouts = ce_loadouts()
    for size in range(min(3, len(alive)), 0, -1):
        for chosen in combinations(alive, size):
            for loadout in loadouts:
                result = evaluate(list(chosen), loadout, people, coef, demands, fixed_bonus)
                if result is None:
                    continue
                if rule == "throughput":
                    if COST_FIRST:
                        key = (-round(result["throughput"], 6), -result["cost"], result["runs"])
                    else:
                        key = (-round(result["throughput"], 6), result["runs"],
                               -round(max(runner_gain(r, result["enabled"], coef, fixed_bonus, result["flat"]) for r in RUNNERS), 6),
                               result["loss"])
                else:
                    key = (round(result["loss"], 9), result["runs"],
                           -round(max(runner_gain(r, result["enabled"], coef, fixed_bonus, result["flat"]) for r in RUNNERS), 6))
                if best_key is None or key < best_key:
                    best_key, best = key, result
    return best


def greedy_plan(people: list[dict], coef: list[float], rule: str,
                runner_needs: list[int] | None = None) -> dict:
    needs = [p["need"] for p in people]
    alive = list(range(len(people)))
    rounds = []
    total_runs = 0
    lefts = [float(r["need"]) for r in RUNNERS] if runner_needs is None else [float(v) for v in runner_needs]
    fixed = FIXED_BONUS + RUNNER_BONUS * sum(1 for v in lefts if v <= 0)
    while alive and len(rounds) < MAX_ROUNDS:
        fit = max_feasible_size(people, alive)
        used_fixed = fixed
        step = solve_round(people, alive, coef, needs, rule, used_fixed)
        if step is None:
            break
        chosen = step["chosen"]
        runs = step["runs"]
        if runs <= 0:
            break
        before = {i: needs[i] for i in chosen}
        leftover = {}
        graduated = []
        for index, i in enumerate(chosen):
            needs[i] -= step["gains"][index] * runs
            leftover[i] = needs[i]
            if needs[i] <= 0:
                graduated.append(i)
        item = {
            "index": len(rounds) + 1,
            "chosen": chosen,
            "size": len(chosen),
            "max_fit": fit,
            "enabled": step["enabled"],
            "cost": step["cost"],
            "equip_count": step["equip_count"],
            "equip_cost": step["equip_cost"],
            "total_cost": step["total_cost"],
            "n5": step["n5"],
            "five": step["five"],
            "two": step["two"],
            "flat": step["flat"],
            "runs": runs,
            "p3": step["p3"],
            "r3": step["r3"],
            "loss": step["loss"],
            "throughput": step["throughput"],
            "gains": step["gains"],
            "before": before,
            "leftover": leftover,
            "graduated": graduated,
            "alive_after": len(alive) - len(graduated),
            "fixed": used_fixed,
        }
        # 前排陪跑同步吃绊，练满后固定区 +20%
        fixed = apply_runners(lefts, runs, item["enabled"], coef, fixed, item)
        rounds.append(item)
        total_runs += runs
        alive = [i for i in alive if needs[i] > 0]
    return {
        "rule": rule,
        "rounds": rounds,
        "total_runs": total_runs,
        "total_ap": total_runs * AP_PER_RUN,
        "apples": total_runs * AP_PER_RUN / AP_PER_APPLE,
        "finished": not alive,
        "remaining": alive,
        "needs": needs,
        "runner_lefts": lefts,
        "runner_rounds": {
            r["name"]: next((it["index"] for it in rounds if r["name"] in (it.get("runner_graduated") or [])), None)
            for r in RUNNERS
        },
    }


def randomized_greedy(people: list[dict], coef: list[float], seed: int, topk: int = 2) -> dict:
    """随机重启版贪心：每步在吞吐前 topk 里随机挑一个，跳出确定性贪心的局部最优。

    与 greedy_plan 结构一致，只是选择带随机性；多次不同 seed 取总场次最小者，
    就是报告里用的最优解。
    """
    import random

    rng = random.Random(seed)
    loadouts = ce_loadouts()
    bonus_table = [
        [ce_bonus(people[i]["hits"], loadout, coef) for loadout in loadouts]
        for i in range(len(people))
    ]
    needs = [p["need"] for p in people]
    alive = list(range(len(people)))
    lefts = [float(r["need"]) for r in RUNNERS]
    fixed = FIXED_BONUS + RUNNER_BONUS * sum(1 for v in lefts if v <= 0)
    rounds = []
    total_runs = 0
    while alive and len(rounds) < MAX_ROUNDS:
        candidates = []
        for size in range(min(3, len(alive)), 0, -1):
            for chosen in combinations(alive, size):
                cost = sum(people[i]["cost"] for i in chosen)
                for k, loadout in enumerate(loadouts):
                    if cost + loadout["equip_cost"] > COST_CAP - COST_BASE:
                        continue
                    gains = [BASE_BOND * (fixed + bonus_table[i][k]) for i in chosen]
                    ratios = [g / needs[i] if needs[i] > 0 else 0.0 for g, i in zip(gains, chosen)]
                    runs = math.ceil(1 / max(ratios)) if max(ratios) else 0
                    if runs <= 0:
                        continue
                    pushed = sum(min(g * runs, needs[i]) for g, i in zip(gains, chosen))
                    candidates.append(
                        ((-round(pushed / runs, 6), runs),
                         {"chosen": list(chosen), "enabled": loadout["enabled"], "gains": gains, "runs": runs,
                          "n5": 0, "five": loadout["five"], "two": loadout["two"], "flat": loadout["flat"],
                          "cost": cost, "equip_count": sum(loadout["enabled"]) + loadout["five"] + loadout["two"],
                          "equip_cost": loadout["equip_cost"],
                          "total_cost": COST_BASE + cost + loadout["equip_cost"]})
                    )
            if candidates:
                break
        if not candidates:
            break
        candidates.sort(key=lambda item: item[0])
        step = rng.choice(candidates[:topk])[1]
        chosen = step["chosen"]
        runs = step["runs"]
        before = {i: needs[i] for i in chosen}
        meta = evaluate(chosen, step, people, coef, needs, fixed) or {}
        leftover, graduated = {}, []
        for index, i in enumerate(chosen):
            needs[i] -= step["gains"][index] * runs
            leftover[i] = needs[i]
            if needs[i] <= 0:
                graduated.append(i)
        item = {
            "index": len(rounds) + 1,
            "chosen": chosen,
            "size": len(chosen),
            "max_fit": max_feasible_size(people, alive),
            "enabled": step["enabled"],
            "cost": step["cost"],
            "equip_count": step["equip_count"],
            "equip_cost": step["equip_cost"],
            "total_cost": step["total_cost"],
            "n5": step["n5"],
            "five": step["five"],
            "two": step["two"],
            "flat": step["flat"],
            "runs": runs,
            "gains": step["gains"],
            "before": before,
            "leftover": leftover,
            "graduated": graduated,
            "alive_after": len(alive) - len(graduated),
            "fixed": fixed,
            "loss": meta.get("loss", 0.0),
            "p3": meta.get("p3", 0.0),
            "r3": meta.get("r3", 0.0),
            "throughput": meta.get("throughput", 0.0),
        }
        fixed = apply_runners(lefts, runs, item["enabled"], coef, fixed, item)
        rounds.append(item)
        total_runs += runs
        alive = [i for i in alive if needs[i] > 0]
    return {
        "rule": "throughput",
        "rounds": rounds,
        "total_runs": total_runs,
        "needs": needs,
        "seed": seed,
        "finished": not alive,
        "remaining": alive,
        "runner_lefts": lefts,
    }


def runner_gain(runner: dict, enabled: list[int], coef: list[float],
                back_fixed: float = BACK_FIXED, flat: float = 0.0) -> float:
    """陪跑者单场收益：当前前排固定区 + 她命中的已启用礼装 + 泛用礼装。

    前排比后排多 0.2。每有一人到绊 15，back_fixed 就加 0.2，前排跟着变成 2.14、2.34。
    """
    bonus = sum(runner["hits"][j] * enabled[j] * coef[j] for j in range(len(enabled))) + flat
    return BASE_BOND * (back_fixed + FRONT_POSITION + bonus)


def apply_runners(lefts: list[float], runs: int, enabled: list[int], coef: list[float],
                  fixed: float, item: dict) -> float:
    """推进陪跑进度；有人练满就抬升全队固定区，返回新的后排固定区。

    这一轮里两个人都按进来时的固定区结算，毕业加成从下一轮才开始。
    """
    for index, runner in enumerate(RUNNERS):
        if lefts[index] > 0:
            lefts[index] -= runs * runner_gain(runner, enabled, coef, fixed, item.get("flat", 0.0))
            if lefts[index] <= 0:
                fixed += RUNNER_BONUS
                item.setdefault("runner_graduated", []).append(runner["name"])
    return fixed


def reorder_for_runner(people: list[dict], coef: list[float], rounds: list[dict],
                       runner_needs: list[int] | None, rule: str) -> list[dict]:
    """索引重排：阵容与缺口都沿用原解，只把陪跑者收益最高的轮次提前。

    每轮用「原解里该轮的输入缺口」重算场次——陪跑毕业之后 fixed 会抬升，
    这些轮次的场次因此缩短；阵容与决策不改变。
    """
    lefts = [float(r["need"]) for r in RUNNERS] if runner_needs is None else [float(v) for v in runner_needs]
    fixed = FIXED_BONUS + RUNNER_BONUS * sum(1 for v in lefts if v <= 0)

    def total_runner_gain(item: dict) -> float:
        return sum(runner_gain(r, item["enabled"], coef, fixed, item.get("flat", 0.0)) for r in RUNNERS)

    order = sorted(rounds, key=lambda item: -(total_runner_gain(item) * item["runs"]))
    out = []
    for item in order:
        demands = [p["need"] for p in people]
        for i, value in item["before"].items():
            demands[i] = value
        result = evaluate(item["chosen"], item, people, coef, demands, fixed)
        if result is None or result["runs"] <= 0:
            continue
        runs = result["runs"]
        entry = dict(item)
        entry.update(
            {
                "index": len(out) + 1,
                "runs": runs,
                "gains": result["gains"],
                "leftover": {
                    i: item["before"][i] - result["gains"][idx] * runs
                    for idx, i in enumerate(item["chosen"])
                },
                "graduated": [
                    i for idx, i in enumerate(item["chosen"])
                    if item["before"][i] - result["gains"][idx] * runs <= 0
                ],
                "loss": result["loss"],
                "p3": result["p3"],
                "r3": result["r3"],
                "throughput": result["throughput"],
                "equip_count": result["equip_count"],
                "total_cost": result["total_cost"],
                "fixed": fixed,
            }
        )
        fixed = apply_runners(lefts, runs, entry["enabled"], coef, fixed, entry)
        out.append(entry)
    return out


def finalize(rounds: list[dict], needs: list[int], rule: str, runner_lefts: list[float] | None = None) -> dict:
    total_runs = sum(item["runs"] for item in rounds)
    alive = [i for i, value in enumerate(needs) if value > 0]
    size_stats: dict[int, dict] = {}
    for item in rounds:
        stat = size_stats.setdefault(item["size"], {"rounds": 0, "runs": 0})
        stat["rounds"] += 1
        stat["runs"] += item["runs"]
    underfilled = [item["index"] for item in rounds if item["size"] < item.get("max_fit", item["size"])]
    return {
        "rule": rule,
        "rounds": rounds,
        "size_stats": size_stats,
        "underfilled": underfilled,
        "total_runs": total_runs,
        "total_ap": total_runs * AP_PER_RUN,
        "apples": math.ceil(total_runs * AP_PER_RUN / AP_PER_APPLE),
        "finished": not alive,
        "remaining": alive,
        "runner_lefts": runner_lefts,
        "runner_rounds": {
            r["name"]: next((it["index"] for it in rounds if r["name"] in (it.get("runner_graduated") or [])), None)
            for r in RUNNERS
        },
    }


def fmt(value: float) -> str:
    return f"{value:,.0f}"


PAGE_CSS = """
  :root { color-scheme: light; }
  * { box-sizing: border-box; }
  body { margin: 0; padding: 28px 20px 60px; font: 14px/1.6 "Segoe UI", "Microsoft YaHei", sans-serif;
         background: #f5f6f8; color: #1f2430; }
  .wrap { max-width: 1080px; margin: 0 auto; }
  h1 { font-size: 20px; margin: 0 0 4px; }
  h2 { font-size: 15px; margin: 0 0 2px; }
  .sub { color: #6b7280; font-size: 12px; margin-bottom: 18px; }
  .sub2 { color: #9ca3af; font-size: 12px; margin: 0 0 12px; }
  table.cmp { width: 100%; border-collapse: collapse; background: #fff; border: 1px solid #e5e7eb;
              border-radius: 10px; overflow: hidden; margin-bottom: 22px; font-size: 13px; }
  table.cmp th, table.cmp td { padding: 9px 12px; text-align: right; border-bottom: 1px solid #f0f1f3; }
  table.cmp th:first-child, table.cmp td:first-child { text-align: left; }
  table.cmp thead th { background: #f9fafb; color: #6b7280; font-weight: 600; font-size: 12px; }
  table.cmp tbody tr:last-child td { border-bottom: none; }
  .cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: 12px; margin-bottom: 18px; }
  .card { background: #fff; border: 1px solid #e5e7eb; border-radius: 10px; padding: 12px 14px; }
  .card .k { font-size: 12px; color: #6b7280; }
  .card .v { font-size: 20px; font-weight: 600; margin-top: 2px; font-variant-numeric: tabular-nums; }
  .card .v.small { font-size: 13px; font-weight: 500; }
  .card .u { font-size: 12px; color: #9ca3af; margin-left: 2px; font-weight: 400; }
  .panel { background: #fff; border: 1px solid #e5e7eb; border-radius: 10px; padding: 16px 16px 8px; margin-bottom: 22px; }
  details.round { border-top: 1px solid #f0f1f3; }
  details.round:first-of-type { border-top: none; }
  summary { display: flex; gap: 14px; align-items: center; padding: 9px 2px; cursor: pointer; list-style: none; }
  summary::-webkit-details-marker { display: none; }
  .rno { font-weight: 600; min-width: 66px; }
  .runs { background: #eef2ff; color: #3730a3; border-radius: 999px; padding: 1px 10px; font-size: 12px; }
  .size { background: #f3f4f6; color: #4b5563; border-radius: 999px; padding: 1px 10px; font-size: 12px; }
  .size.warn { background: #fef3c7; color: #92400e; }
  .loss { color: #6b7280; font-size: 12px; font-variant-numeric: tabular-nums; }
  .meta { margin-left: auto; color: #9ca3af; font-size: 12px; }
  .body { padding: 2px 2px 14px; }
  ul.crew { list-style: none; margin: 0 0 6px; padding: 0; display: grid; gap: 6px; }
  ul.crew li { display: flex; gap: 10px; align-items: baseline; flex-wrap: wrap;
               background: #fafbfc; border: 1px solid #f0f1f3; border-radius: 8px; padding: 6px 10px; }
  .cost, .gain, .need { font-size: 12px; color: #6b7280; font-variant-numeric: tabular-nums; }
  .gain { color: #047857; }
  .done { margin-left: auto; background: #dcfce7; color: #15803d; border-radius: 999px; padding: 0 8px; font-size: 12px; }
  .equip { margin: 0; font-size: 13px; }
  .dim { color: #9ca3af; font-size: 12px; }
"""


def render_section(entry: dict, people: list[dict], coef: list[float], labels: list[str]) -> str:
    plan = entry["plan"]
    title, desc = RULES[entry["rule"]]
    rows = []
    for item in plan["rounds"]:
        crew = "".join(
            f"<li><b>{html.escape(people[i]['name'])}</b>"
            f"<span class=cost>{people[i]['cost']} cost</span>"
            f"<span class=gain>+{fmt(item['gains'][k])}/场</span>"
            f"<span class=need>{fmt(item['before'][i])} → {fmt(max(0.0, item['leftover'][i]))}</span>"
            f"{'<span class=done>毕业</span>' if item['leftover'][i] <= 0 else ''}</li>"
            for k, i in enumerate(item["chosen"])
        )
        equips = html.escape(format_equips(item, labels, coef))
        rows.append(
            f"""
          <details class=round{' open' if item['index'] <= 2 else ''}>
            <summary>
              <span class=rno>第 {item['index']} 轮</span>
              <span class=runs>{item['runs']} 场</span>
              <span class=size{' warn' if item['size'] < item['max_fit'] else ''}>{item['size']} 人{'（本可上 ' + str(item['max_fit']) + ' 人）' if item['size'] < item['max_fit'] else ''}</span>
              <span class=loss>loss {item['loss']:.4f}</span>
              <span class=meta>cost {item['total_cost']}{' · 第 3 件 +5%' if item['n5'] else ''}{' · 陪跑毕业' if item.get('runner_graduated') else ''}</span>
            </summary>
            <div class=body>
              <ul class=crew>{crew}</ul>
              <p class=equip>礼装：{equips}　<span class=dim>P3 {item['p3']:.4f} · R3 {item['r3']:.3f} · 本轮后剩 {item['alive_after']} 人</span></p>
            </div>
          </details>"""
        )
    status = "全部毕业" if plan["finished"] else f"剩 {len(plan['remaining'])} 人"
    stats = plan.get("size_stats", {})
    non_full_runs = stats.get(2, {}).get("runs", 0) + stats.get(1, {}).get("runs", 0)
    non_full_rounds = stats.get(2, {}).get("rounds", 0) + stats.get(1, {}).get("rounds", 0)
    underfilled = len(plan.get("underfilled") or [])
    rounds_by_runner = plan.get("runner_rounds") or {}
    runner_text = "　".join(
        f"{r['name'].split('·')[0]}：{'第 ' + str(rounds_by_runner[r['name']]) + ' 轮' if rounds_by_runner.get(r['name']) else '未毕业'}"
        for r in RUNNERS
    )
    return f"""
      <section class=panel>
        <h2>{html.escape(title)} <span class=dim>（{html.escape(entry['rule'])}）</span></h2>
        <p class=sub2>{html.escape(desc)}
          {'　·　随机重启 ' + str(plan['restarts']) + ' 次，最优 seed ' + str(plan.get('seed')) if plan.get('restarts') else ''}</p>
        <div class=cards>
          <div class=card><div class=k>总轮次</div><div class=v>{len(plan['rounds'])}</div></div>
          <div class=card><div class=k>总场次</div><div class=v>{plan['total_runs']}<span class=u>场</span></div></div>
          <div class=card><div class=k>总 AP</div><div class=v>{plan['total_ap']:,}<span class=u>AP</span></div></div>
          <div class=card><div class=k>金苹果</div><div class=v>{plan['apples']}<span class=u>个</span></div></div>
          <div class=card><div class=k>非满编场次</div><div class=v>{non_full_runs}<span class=u>场 / {non_full_rounds} 轮</span></div></div>
          <div class=card><div class=k>陪跑毕业</div><div class=v class=small>{runner_text}</div></div>
          <div class=card><div class=k>状态</div><div class=v>{status}</div></div>
        </div>
        <p class=sub2>队伍人数分布：{'、'.join(f"{k} 人 {v['rounds']} 轮 / {v['runs']} 场" for k, v in sorted(stats.items(), reverse=True))}
          {f'；其中 {underfilled} 轮是准则选择未上满（非 cost 限制）' if underfilled else '；所有非满编轮均为 cost 上限所致'}</p>
        {''.join(rows)}
      </section>"""


def _coverage_configs(people: list[dict], coef: list[float]) -> list[dict]:
    """枚举合法队伍。同一组人、同一套礼装收益只留一份。"""
    raw = []
    loadouts = ce_loadouts()
    budget = COST_CAP - COST_BASE
    for size in range(1, 4):
        for chosen in combinations(range(len(people)), size):
            cost = sum(people[i]["cost"] for i in chosen)
            for loadout in loadouts:
                if cost + loadout["equip_cost"] > budget:
                    continue
                bonuses = [ce_bonus(people[i]["hits"], loadout, coef) for i in chosen]
                runner_ce = [ce_bonus(runner["hits"], loadout, coef) for runner in RUNNERS]
                raw.append(
                    {
                        "chosen": chosen,
                        "enabled": loadout["enabled"],
                        "cost": cost,
                        "n5": 0,
                        "five": loadout["five"],
                        "two": loadout["two"],
                        "flat": loadout["flat"],
                        "equip_count": sum(loadout["enabled"]) + loadout["five"] + loadout["two"],
                        "equip_cost": loadout["equip_cost"],
                        "bonuses": bonuses,
                        "runner_ce": runner_ce,
                    }
                )
    unique = {}
    for cfg in raw:
        key = (
            cfg["chosen"],
            tuple(round(b, 6) for b in cfg["bonuses"]),
            cfg["n5"],
            tuple(round(v, 4) for v in cfg["runner_ce"]),
        )
        unique[key] = cfg
    return list(unique.values())


def solve_coverage(people: list[dict], coef: list[float]) -> dict:
    """按绊 15 的全队加成切段，选总场次更少的一种。

    孔明缺口更小，且每场两人都在，所以他先到绊 15。山鲁佐德到绊 15 会再给全队 +20%。
    这一档可有可无：只有它能把总场次压下去时才采用。
    """
    import numpy as np
    from scipy.optimize import LinearConstraint, linprog, milp

    configs = _coverage_configs(people, coef)
    # (标题, 后排, 前排, 这档孔明还在冲绊, 这档山鲁佐德还在冲绊)
    after_kong = [
        ("孔明未到绊15", BACK_FIXED, FRONT_FIXED, True, True),
        ("孔明已到绊15", BACK_FIXED + RUNNER_BONUS, FRONT_FIXED + RUNNER_BONUS, False, True),
    ]
    after_both = after_kong + [
        ("两人都到绊15", BACK_FIXED + 2 * RUNNER_BONUS, FRONT_FIXED + 2 * RUNNER_BONUS, False, False),
    ]
    only_kong = _solve_stages(configs, people, after_kong, require_sche=False, np=np, linprog=linprog, milp=milp, LinearConstraint=LinearConstraint)
    both = _solve_stages(configs, people, after_both, require_sche=True, np=np, linprog=linprog, milp=milp, LinearConstraint=LinearConstraint)
    if both["finished"] and both["total_runs"] < only_kong["total_runs"]:
        both["note"] = "山鲁佐德也刷到绊 15，后排固定区再到 2.14，总场次更低，所以采用这一档。"
        return both
    only_kong["note"] = (
        f"山鲁佐德到绊 15 会再给全队 +20%（后排 2.14）。"
        f"把她也刷满要 {both['total_runs']} 场，不比现在少，所以这版不强制她到绊 15。"
        if both["finished"] else
        "山鲁佐德到绊 15 会再给全队 +20%。这一档没有找到更低的整数解，所以不强制她到绊 15。"
    )
    return only_kong


def _solve_stages(configs, people, stages, require_sche, np, linprog, milp, LinearConstraint) -> dict:
    """stages 里每一档各有一套场次变量。孔明的缺口只由 count_kong 的档覆盖。"""
    need = np.array([p["need"] for p in people], dtype=float)
    n_people = len(people)
    n = len(configs)
    phase_count = len(stages)

    def block(cfgs: list[dict], back: float) -> np.ndarray:
        out = np.zeros((n_people, len(cfgs)))
        for col, cfg in enumerate(cfgs):
            for person, bonus in zip(cfg["chosen"], cfg["bonuses"]):
                out[person, col] = BASE_BOND * (back + cfg["n5"] * N5_BONUS + bonus)
        return out

    def runner_row(cfgs: list[dict], runner_index: int, which) -> np.ndarray:
        row = np.zeros(phase_count * len(cfgs))
        width = len(cfgs)
        for phase, stage in enumerate(stages):
            if not which(stage):
                continue
            front = stage[2]
            row[phase * width:(phase + 1) * width] = [
                BASE_BOND * (front + cfg["runner_ce"][runner_index]) for cfg in cfgs
            ]
        return row

    cover = np.concatenate([block(configs, stage[1]) for stage in stages], axis=1)
    system = np.vstack([cover, runner_row(configs, 0, lambda stage: stage[3])])
    lower_list = [need, np.array([RUNNERS[0]["need"]])]
    if require_sche:
        system = np.vstack([system, runner_row(configs, 1, lambda stage: stage[4])])
        lower_list.append(np.array([RUNNERS[1]["need"]]))
    lower = np.concatenate(lower_list)
    relaxed = linprog(
        np.ones(system.shape[1]), A_ub=-system, b_ub=-lower, bounds=(0, None), method="highs",
    )
    if not relaxed.success:
        return _empty_coverage(stages, float("inf"))

    support = set()
    for phase in range(phase_count):
        support |= set(np.flatnonzero(relaxed.x[phase * n:(phase + 1) * n] > 1e-6))
    sub = [configs[i] for i in sorted(support)]
    width = len(sub)
    packed_cover = np.concatenate([block(sub, stage[1]) for stage in stages], axis=1)
    packed = np.vstack([packed_cover, runner_row(sub, 0, lambda stage: stage[3])])
    packed_lower = [need, np.array([RUNNERS[0]["need"]])]
    if require_sche:
        packed = np.vstack([packed, runner_row(sub, 1, lambda stage: stage[4])])
        packed_lower.append(np.array([RUNNERS[1]["need"]]))
    packed_lower = np.concatenate(packed_lower)
    solved = milp(
        c=np.ones(packed.shape[1]),
        constraints=LinearConstraint(packed, packed_lower, np.full(packed.shape[0], np.inf)),
        integrality=np.ones(packed.shape[1], dtype=int),
        bounds=(0, 3000),
        options={"time_limit": 45},
    )
    if solved.x is None:
        return _empty_coverage(stages, float(relaxed.fun))
    runs = np.rint(solved.x).astype(int)
    return _materialize(sub, people, stages, runs, need, float(relaxed.fun), require_sche)


def _empty_coverage(stages, lower_bound: float) -> dict:
    return {
        "total_runs": 10**9,
        "total_ap": 0,
        "apples": 0,
        "lower_bound": lower_bound,
        "finished": False,
        "phases": [],
        "kong_left": float(RUNNERS[0]["need"]),
        "sche_left": float(RUNNERS[1]["need"]),
    }


def _materialize(sub, people, stages, runs, need, lower_bound: float, require_sche: bool) -> dict:
    width = len(sub)
    remain = need.astype(float).copy()
    kong = float(RUNNERS[0]["need"])
    sche = float(RUNNERS[1]["need"])
    phase_plans = []
    for phase, stage in enumerate(stages):
        title, back, front, count_kong, count_sche = stage
        count = runs[phase * width:(phase + 1) * width]
        segments = []
        order = sorted((i for i in range(width) if count[i] > 0), key=lambda i: -int(count[i]))
        for i in order:
            cfg = sub[i]
            step = int(count[i])
            gains = [
                BASE_BOND * (back + cfg["n5"] * N5_BONUS + bonus)
                for bonus in cfg["bonuses"]
            ]
            before = {person: remain[person] for person in cfg["chosen"]}
            graduated = []
            for person, gain in zip(cfg["chosen"], gains):
                remain[person] -= gain * step
                if remain[person] <= 1e-4:
                    remain[person] = 0.0
                    graduated.append(person)
            kong_gain = BASE_BOND * (front + cfg["runner_ce"][0]) * step
            sche_gain = BASE_BOND * (front + cfg["runner_ce"][1]) * step
            if count_kong:
                kong = max(0.0, kong - kong_gain)
            if count_sche:
                sche = max(0.0, sche - sche_gain)
            segments.append(
                {
                    "runs": step,
                    "chosen": list(cfg["chosen"]),
                    "gains": gains,
                    "enabled": list(cfg["enabled"]),
                    "cost": cfg["cost"],
                    "total_cost": COST_BASE + cfg["cost"] + cfg["equip_cost"],
                    "n5": cfg["n5"],
                    "five": cfg["five"],
                    "two": cfg["two"],
                    "before": before,
                    "leftover": {person: remain[person] for person in cfg["chosen"]},
                    "graduated": graduated,
                    "kong_gain": kong_gain,
                    "sche_gain": sche_gain,
                }
            )
        if not segments:
            continue
        phase_plans.append(
            {
                "title": title,
                "fixed": back,
                "front": front,
                "runs": int(count.sum()),
                "segments": segments,
            }
        )
    import numpy as np
    total = int(runs.sum())
    return {
        "total_runs": total,
        "total_ap": total * AP_PER_RUN,
        "apples": math.ceil(total * AP_PER_RUN / AP_PER_APPLE),
        "lower_bound": lower_bound,
        "finished": bool(np.all(remain <= 1e-4) and kong <= 1 and (sche <= 1 or not require_sche)),
        "phases": phase_plans,
        "kong_left": kong,
        "sche_left": sche,
    }


def render_coverage(plan: dict, people: list[dict], coef: list[float], labels: list[str]) -> str:
    blocks = []
    serial = 0
    for phase in plan["phases"]:
        rows = []
        for item in phase["segments"]:
            serial += 1
            crew = "".join(
                f"<li><b>{html.escape(people[i]['name'])}</b>"
                f"<span class=cost>{people[i]['cost']} cost</span>"
                f"<span class=gain>+{fmt(item['gains'][k])}/场</span>"
                f"<span class=need>{fmt(item['before'][i])} → {fmt(max(0.0, item['leftover'][i]))}</span>"
                f"{'<span class=done>毕业</span>' if item['leftover'][i] <= 0 else ''}</li>"
                for k, i in enumerate(item["chosen"])
            )
            equips = html.escape(format_equips(item, labels, coef))
            rows.append(
                f"""
          <details class=round{' open' if serial <= 2 else ''}>
            <summary>
              <span class=rno>第 {serial} 段</span>
              <span class=runs>{item['runs']} 场</span>
              <span class=size>{len(item['chosen'])} 人</span>
              <span class=meta>cost {item['total_cost']}{' · 第 3 件 +5%' if item['n5'] else ''}</span>
            </summary>
            <div class=body>
              <ul class=crew>{crew}</ul>
              <p class=equip>礼装：{equips}　<span class=dim>孔明 +{fmt(item['kong_gain'])}　山鲁佐德 +{fmt(item['sche_gain'])}</span></p>
            </div>
          </details>"""
            )
        blocks.append(
            f"<p class=sub2 style=\"color:#1f2430;font-weight:600;margin-top:14px\">"
            f"{html.escape(phase['title'])}　前排 {phase['front']:.2f}　后排 {phase['fixed']:.2f}　"
            f"{phase['runs']} 场　{len(phase['segments'])} 段</p>"
            + "".join(rows)
        )
    status = "25 人全部盖满" if plan["finished"] else "仍有缺口"
    return f"""
      <section class=panel>
        <h2>全局覆盖 <span class=dim>（按绊 15 的全队加成分段）</span></h2>
        <p class=sub2>前排礼装给全队 +70%，全队自带 +4%，站前排再 +20%。孔明和山鲁佐德因此是 1.94，后排 25 人是 1.74。两人冲的都是绊 15，各自满了之后全队再 +20%。后排从 12 件里任选最多 3 件：6 件 12 cost 户口 +20%，3 件 12 cost 全队 +5%，3 件 9 cost 全队 +2%。一个人的缺口可以拆给好几支队伍，同一段里的先后顺序不影响总数。</p>
        <p class=sub2>{html.escape(plan.get("note") or "")}</p>
        <div class=cards>
          <div class=card><div class=k>总场次</div><div class=v>{plan['total_runs']}<span class=u>场</span></div></div>
          <div class=card><div class=k>总 AP</div><div class=v>{plan['total_ap']:,}<span class=u>AP</span></div></div>
          <div class=card><div class=k>金苹果</div><div class=v>{plan['apples']}<span class=u>个</span></div></div>
          <div class=card><div class=k>分数下界</div><div class=v>{math.ceil(plan['lower_bound'] - 1e-6)}<span class=u>场</span></div></div>
          <div class=card><div class=k>孔明</div><div class=v class=small>{"已到绊15" if plan["kong_left"] <= 1 else "还差 " + fmt(plan["kong_left"])}</div></div>
          <div class=card><div class=k>山鲁佐德</div><div class=v class=small>{"已到绊15" if plan["sche_left"] <= 1 else "还差 " + fmt(plan["sche_left"])}</div></div>
          <div class=card><div class=k>状态</div><div class=v>{status}</div></div>
        </div>
        {''.join(blocks)}
      </section>"""


def render_html(people: list[dict], coef: list[float], labels: list[str], plans: list[dict],
                source: str, coverage: dict | None = None) -> str:
    total_need = sum(p["need"] for p in people)
    compare = "".join(
        f"<tr><td>{html.escape(RULES[entry['rule']][0])}</td>"
        f"<td>{len(entry['plan']['rounds'])}</td>"
        f"<td>{entry['plan']['total_runs']}</td>"
        f"<td>{entry['plan']['total_ap']:,}</td>"
        f"<td>{entry['plan']['apples']:.0f}</td></tr>"
        for entry in plans
    )
    sections = "".join(render_section(entry, people, coef, labels) for entry in plans)
    if coverage:
        segments = sum(len(phase["segments"]) for phase in coverage["phases"])
        compare += (
            f"<tr><td>全局覆盖</td><td>{segments} 段</td>"
            f"<td>{coverage['total_runs']}</td>"
            f"<td>{coverage['total_ap']:,}</td>"
            f"<td>{coverage['apples']}</td></tr>"
        )
        sections += render_coverage(coverage, people, coef, labels)
    return f"""<!DOCTYPE html>
<html lang=zh-CN>
<head>
<meta charset=utf-8>
<title>Caster 绊 9 贪心方案</title>
<style>{PAGE_CSS}</style>
</head>
<body>
<div class=wrap>
  <h1>Caster 绊 9 贪心方案</h1>
  <div class=sub>数据源：{html.escape(source)} → Sheet4（{len(people)} 名候选，总缺口 {total_need:,}）　·　
    单轮约束：≤3 人 + 2 个礼装槽 + cost {COST_CAP}　·　{AP_PER_RUN} AP/场，{AP_PER_APPLE} AP/果<br>
    前排陪跑：{'、'.join(f"{r['name']}（剩余 {r['need']:,}）" for r in RUNNERS)}。前排礼装 +70% 与自带 +4% 给全队，站前排再 +20%；两人各到绊 15 再给全队 +20%</div>
  <table class=cmp>
    <thead><tr><th>准则</th><th>轮次</th><th>总场次</th><th>总 AP</th><th>金苹果</th></tr></thead>
    <tbody>{compare}</tbody>
  </table>
  {sections}
</div>
</body>
</html>
"""


def main() -> None:
    parser = argparse.ArgumentParser(description="把 Sheet4 的模型程序化，贪心推导完整方案")
    parser.add_argument("-i", "--input", type=Path, default=Path("caster_bond.xlsx"))
    parser.add_argument("--sheet", default="Sheet4")
    parser.add_argument("-o", "--output", type=Path, default=Path("caster_plan.html"))
    parser.add_argument("--rule", choices=["loss", "throughput", "both"], default="both")
    parser.add_argument("--runner-need", type=int, default=RUNNERS[0]["need"],
                        help="阿尔托莉雅的剩余绊缺口，0 表示已毕业")
    parser.add_argument("--runner2-need", type=int, default=RUNNERS[1]["need"],
                        help="山鲁佐德的剩余绊缺口，0 表示已毕业")
    parser.add_argument("--p3-weight", type=float, default=None, help="loss 里 P3 的权重（默认 1.0）")
    parser.add_argument("--r3-weight", type=float, default=None, help="loss 里 R3 的权重（默认 0.2）")
    parser.add_argument("--json", type=Path, default=None, help="可选：同时输出机器可读的 JSON")
    parser.add_argument("--reorder", action="store_true",
                        help="额外尝试「按陪跑收益重排轮次」（实测会让总场次变差，仅用于对比）")
    parser.add_argument("--cost-first", action="store_true",
                        help="同等吞吐下优先上高 cost 从者，把 3★ 留给后面的三人编队")
    parser.add_argument("--restarts", type=int, default=40,
                        help="throughput 准则的随机重启次数（0 = 纯确定性贪心）")
    args = parser.parse_args()

    global P3_WEIGHT, R3_WEIGHT, COST_FIRST
    if args.p3_weight is not None:
        P3_WEIGHT = args.p3_weight
    if args.r3_weight is not None:
        R3_WEIGHT = args.r3_weight
    COST_FIRST = args.cost_first

    data = load_inputs(args.input, args.sheet)
    people, coef, labels = data["people"], data["coefficients"], data["labels"]

    rules = ["loss", "throughput"] if args.rule == "both" else [args.rule]
    runner_needs = [args.runner_need, args.runner2_need]
    entries = []
    for rule in rules:
        if rule == "throughput" and args.restarts > 0:
            best = None
            for seed in range(args.restarts):
                candidate = randomized_greedy(people, coef, seed, topk=2)
                if candidate["finished"] and (best is None or candidate["total_runs"] < best["total_runs"]):
                    best = candidate
            if best is not None:
                plan = finalize(best["rounds"], best["needs"], rule, best.get("runner_lefts"))
                plan["restarts"] = args.restarts
                plan["seed"] = best.get("seed")
                entries.append({"rule": rule, "plan": plan})
                continue
        base = greedy_plan(people, coef, rule, runner_needs)
        if args.reorder:
            ordered = reorder_for_runner(people, coef, base["rounds"], runner_needs, rule)
            plan = finalize(ordered, base["needs"], rule, base.get("runner_lefts"))
        else:
            plan = finalize(base["rounds"], base["needs"], rule, base.get("runner_lefts"))
        entries.append({"rule": rule, "plan": plan})

    coverage = solve_coverage(people, coef)
    args.output.write_text(
        render_html(people, coef, labels, entries, f"{args.input.name}/{args.sheet}", coverage),
        encoding="utf-8",
    )
    if args.json:
        args.json.write_text(
            json.dumps({"people": people, "plans": entries}, ensure_ascii=False, indent=2), encoding="utf-8"
        )

    for entry in entries:
        plan = entry["plan"]
        status = "全部毕业" if plan["finished"] else f"剩 {len(plan['remaining'])} 人"
        print(f"[{entry['rule']}] {len(plan['rounds'])} 轮，总场次 {plan['total_runs']}，"
              f"AP {plan['total_ap']:,}，金苹果 {plan['apples']}，{status}")
        sizes = "，".join(
            f"{size} 人 {stat['rounds']} 轮 / {stat['runs']} 场"
            for size, stat in sorted(plan["size_stats"].items(), reverse=True)
        )
        print(f"   队伍人数：{sizes}；未上满 {len(plan['underfilled'])} 轮")
        for item in plan["rounds"]:
            crew = " / ".join(people[i]["name"] for i in item["chosen"])
            equips = format_equips(item, labels, coef)
            done = "、".join(people[i]["name"] for i in item["graduated"])
            print(f"  R{item['index']:>2} {item['runs']:>3} 场 cost {item['cost']:>2} | {crew} | {equips} | 毕业: {done}")
        print()
    print(f"[coverage] {coverage['total_runs']} 场，AP {coverage['total_ap']:,}，金苹果 {coverage['apples']}，"
          f"{'全部盖满' if coverage['finished'] else '仍有缺口'}")
    print(f"已写出 {args.output}")


if __name__ == "__main__":
    main()
