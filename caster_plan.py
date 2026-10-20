"""把 Sheet4 的模型程序化，并用多轮贪心推导「把候选 Caster 全部练到绊 9」的完整方案。

单轮模型（与 caster_bond.xlsx 的 Sheet4 一致）：
    L_i = K_i × 4748 × (1 + 0.04 + 0.4 + 0.4 + n5 × 0.05 + Σ 命中_ij × 启用_j × 系数_j)
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
FIXED_BONUS = 1 + 0.04 + 0.4 + 0.4                     # 固定三人段的收益系数
IDEAL_BONUS = 1 + 0.04 + 0.4 + 0.4 + 0.05 + 0.4        # P3 归一化用的理想单人系数
COST_FIXED = 80                                        # 队伍其余成员的 cost（含 2 件礼装）
COST_BASE = 56                                         # 前三排固定占用：16×2 + 12×2
COST_CAP = 113                                         # 队伍 cost 上限
EQUIP_COST = 12                                        # 每件礼装占的 cost
SLOTS = 2                                              # 可竞争的两个礼装槽
N5_COST = 21                                           # 恰好 3 个 3★ 时才触发第 3 件 +5% 礼装
N5_BONUS = 0.05
AP_PER_RUN = 40                                        # 每场 AP
AP_PER_APPLE = 146                                     # 每个金苹果换算 AP
P3_WEIGHT = 1.0                                        # loss 里 P3 的权重（--p3-weight 可覆盖）
R3_WEIGHT = 0.2                                        # loss 里 R3 的权重：户口浪费只作轻微修正
RUNNER_BONUS = 0.2                                     # 每位陪跑练满后，全队固定区 +0.2
RUNNER_FIXED_BONUS = FIXED_BONUS + 0.24                # 站前排天然多 0.24 的固定区
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


def evaluate(chosen: list[int], enabled: list[int], people: list[dict], coef: list[float],
             demands: list[int], fixed_bonus: float = FIXED_BONUS) -> dict | None:
    cost_sum = sum(people[i]["cost"] for i in chosen)
    n5 = 1 if cost_sum == N5_COST else 0
    # 从者与礼装共用同一份 cost 额度：前三排固定 56，剩下 113-56=57 给「我方从者 + 礼装」，
    # 每少挂一件礼装就腾出 12 cost，所以「3 人 + 1 礼装」常常打得过「2 人 + 2 礼装」
    equip_count = sum(enabled) + n5
    budget = COST_CAP - COST_BASE
    if cost_sum + EQUIP_COST * equip_count > budget:
        return None
    gains, ratios = [], []
    for i in chosen:
        bonus = sum(people[i]["hits"][j] * enabled[j] * coef[j] for j in range(len(coef)))
        gain = BASE_BOND * (fixed_bonus + n5 * N5_BONUS + bonus)
        gains.append(gain)
        ratios.append(gain / demands[i] if demands[i] > 0 else 0.0)
    slowest = max(ratios)
    runs = math.ceil(1 / slowest) if slowest else 0
    ideal = fixed_bonus + 0.05 + 0.4
    p3 = sum(gains) / (3 * BASE_BOND * ideal)
    hit20 = sum(people[i]["hits"][j] for i in chosen for j in range(6))
    used20 = sum(people[i]["hits"][j] * enabled[j] for i in chosen for j in range(6))
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
        "equip_count": equip_count,
        "total_cost": COST_BASE + cost_sum + EQUIP_COST * equip_count,
        "n5": n5,
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
    combos = equip_combinations(len(coef))
    for size in range(min(3, len(alive)), 0, -1):
        for chosen in combinations(alive, size):
            for enabled in combos:
                result = evaluate(list(chosen), enabled, people, coef, demands, fixed_bonus)
                if result is None:
                    continue
                if rule == "throughput":
                    if COST_FIRST:
                        key = (-round(result["throughput"], 6), -result["cost"], result["runs"])
                    else:
                        key = (-round(result["throughput"], 6), result["runs"],
                               -round(max(runner_gain(r, result["enabled"], coef) for r in RUNNERS), 6),
                               result["loss"])
                else:
                    key = (round(result["loss"], 9), result["runs"],
                           -round(max(runner_gain(r, result["enabled"], coef) for r in RUNNERS), 6))
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
            "total_cost": step["total_cost"],
            "n5": step["n5"],
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


def runner_gain(runner: dict, enabled: list[int], coef: list[float]) -> float:
    """陪跑者单场收益：前排固定区 2.08 + 她命中的已启用礼装。"""
    bonus = sum(runner["hits"][j] * enabled[j] * coef[j] for j in range(len(coef)))
    return BASE_BOND * (RUNNER_FIXED_BONUS + bonus)


def apply_runners(lefts: list[float], runs: int, enabled: list[int], coef: list[float],
                  fixed: float, item: dict) -> float:
    """推进陪跑进度；有人练满就抬升固定区，返回新的 fixed。"""
    for index, runner in enumerate(RUNNERS):
        if lefts[index] > 0:
            lefts[index] -= runs * runner_gain(runner, enabled, coef)
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
    def total_runner_gain(enabled: list[int]) -> float:
        return sum(runner_gain(r, enabled, coef) for r in RUNNERS)

    order = sorted(rounds, key=lambda item: -(total_runner_gain(item["enabled"]) * item["runs"]))
    lefts = [float(r["need"]) for r in RUNNERS] if runner_needs is None else [float(v) for v in runner_needs]
    fixed = FIXED_BONUS + RUNNER_BONUS * sum(1 for v in lefts if v <= 0)
    out = []
    for item in order:
        demands = [p["need"] for p in people]
        for i, value in item["before"].items():
            demands[i] = value
        result = evaluate(item["chosen"], item["enabled"], people, coef, demands, fixed)
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
        equips = "、".join(
            f"{html.escape(labels[j])}(+{coef[j] * 100:g}%)" for j in range(len(coef)) if item["enabled"][j]
        ) or "—"
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
        <p class=sub2>{html.escape(desc)}</p>
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


def render_html(people: list[dict], coef: list[float], labels: list[str], plans: list[dict],
                source: str) -> str:
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
    前排陪跑：{'、'.join(f"{r['name']}（剩余 {r['need']:,}）" for r in RUNNERS)}，每人练满后固定区 +{RUNNER_BONUS:.2f}</div>
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
        base = greedy_plan(people, coef, rule, runner_needs)
        if args.reorder:
            ordered = reorder_for_runner(people, coef, base["rounds"], runner_needs, rule)
            plan = finalize(ordered, base["needs"], rule, base.get("runner_lefts"))
        else:
            plan = finalize(base["rounds"], base["needs"], rule, base.get("runner_lefts"))
        entries.append({"rule": rule, "plan": plan})

    args.output.write_text(
        render_html(people, coef, labels, entries, f"{args.input.name}/{args.sheet}"), encoding="utf-8"
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
            equips = "、".join(labels[j] for j in range(len(coef)) if item["enabled"][j]) or "—"
            done = "、".join(people[i]["name"] for i in item["graduated"])
            print(f"  R{item['index']:>2} {item['runs']:>3} 场 cost {item['cost']:>2} | {crew} | {equips} | 毕业: {done}")
        print()
    print(f"已写出 {args.output}")


if __name__ == "__main__":
    main()
