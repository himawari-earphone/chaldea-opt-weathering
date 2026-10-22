"""导出持有 Caster 的羁绊状况到 xlsx。

Sheet 1「Caster 冲 9 绊」  
    每名持有 Caster 的当前绊等级 / 绊点、绊 9 的累计门槛、还差多少点，
    以及达芬奇工房（魔力棱镜交换）里「带条件的 +20%」羁绊加成礼装：
    每件礼装一列，1 = 该从者满足条件可享，0 = 不满足（数值一律取满破）。

Sheet 2「10 绊以上 Caster」  
    绊等级 ≥ 10 的 Caster：各等级绊点阈值表 + 当前绊点 / 距下一级还差 / 本段进度。

用法：
    python export_caster_bond.py
    python export_caster_bond.py -i "ac(29).php" -o caster_bond.xlsx
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import xlsxwriter

import parse_bond as pb

CASTER_CLASS_ID = 5
TARGET_RANK = 9
TARGET_THRESHOLD_INDEX = TARGET_RANK - 1  # bondGrowth[i] = 升到绊 i+1 的累计点
MANUSCRIPT_WING_CE = 9408060  # 手稿之翼〔Caster〕，满破 +20%


def load_effect_tables(data_dir: Path) -> tuple[dict, dict]:
    skills_path = data_dir / "dist" / "baseSkills.json"
    funcs_path = data_dir / "dist" / "baseFunctions.json"
    missing = [p for p in (skills_path, funcs_path) if not p.exists()]
    if missing:
        names = "、".join(str(p) for p in missing)
        raise SystemExit(f"缺少 {names}，先执行：git -C {data_dir} sparse-checkout add \"/dist/baseSkills.json\" \"/dist/baseFunctions.json\"")
    skills = {s["id"]: s for s in pb.load_json(skills_path)}
    funcs = {f["funcId"]: f for f in pb.load_json(funcs_path)}
    return skills, funcs


def collect_conditional_ces(ce: dict, ce_names: dict, skills: dict, funcs: dict) -> list[dict]:
    """收集达芬奇工房（魔力棱镜交换）里带条件限制的羁绊加成礼装，数值取满破。"""
    result = []
    for essence in ce.values():
        if essence.get("flag") != "svtEquipManaExchange":
            continue
        best = None
        for skill_ref in essence.get("skills") or []:
            skill = skills.get(skill_ref["id"])
            if not skill:
                continue
            for function in skill.get("functions") or []:
                meta = funcs.get(function["funcId"]) or {}
                if meta.get("funcType") != "servantFriendshipUp":
                    continue
                traits = [t["id"] for t in meta.get("functvals") or []]
                groups = [[t["id"] for t in group] for group in meta.get("overWriteTvalsList") or []]
                if not traits and not groups:
                    continue
                svals = (function.get("svals") or [{}])[0]
                rate = int(svals.get("RateCount") or 0)
                if best is None or rate > best["rate"]:
                    best = {"ce_id": essence["id"], "rate": rate, "traits": traits, "groups": groups}
        if best:
            best["name"] = pb.localized_name(ce_names, essence.get("name")) or str(essence["id"])
            result.append(best)
    result.sort(key=lambda item: (0 if item["ce_id"] == MANUSCRIPT_WING_CE else 1, item["name"]))
    return result


def matches_condition(traits: set[int], groups: list[list[int]], single: list[int]) -> bool:
    """groups 为 AND 组合的 OR 列表；single 为任一命中即可的特性列表。"""
    if groups:
        return any(all(t in traits for t in group) for group in groups)
    return any(t in traits for t in single)


def trait_ids(servant: dict) -> set[int]:
    return {t.get("id") for t in (servant.get("traits") or [])}


def owned_svt_ids(login: dict) -> set[int]:
    """当前实际持有（userSvt + 灵基保管室）的 id 集合。

    userSvtCollection 是图鉴记录，会包含已卖掉/已消耗的从者，不能当作「持有」。
    """
    replaced = login["cache"]["replaced"]
    owned: set[int] = set()
    for key in ("userSvt", "userSvtStorage"):
        for row in replaced.get(key) or []:
            owned.add(int(row["svtId"]))
    return owned


def build_caster_rows(owned: set[int], collection: dict, servants: dict, svt_names: dict,
                      conditional: list[dict]) -> list[dict]:
    out = []
    for svt_id in owned:
        servant = servants.get(svt_id)
        if not servant or servant.get("classId") != CASTER_CLASS_ID:
            continue
        curve = servant.get("bondGrowth") or []
        if not curve:
            continue
        row = collection.get(svt_id) or {}
        rank = int(row.get("friendshipRank") or 0)
        points = int(row.get("friendship") or 0)
        need = curve[min(TARGET_THRESHOLD_INDEX, len(curve) - 1)]
        hits = {
            item["ce_id"] for item in conditional
            if matches_condition(trait_ids(servant), item["groups"], item["traits"])
        }
        out.append(
            {
                "collectionNo": servant.get("collectionNo"),
                "svtId": svt_id,
                "name": pb.localized_name(svt_names, servant.get("name")),
                "rarity": servant.get("rarity"),
                "cost": servant.get("cost"),
                "rank": rank,
                "points": points,
                "need9": need,
                "remain": max(0, need - points),
                "reached": rank >= TARGET_RANK,
                "hits": hits,
            }
        )
    out.sort(key=lambda item: (-item["remain"], item["collectionNo"] or 0))
    return out


def write_sheet_target(workbook, rows: list[dict], conditional: list[dict], formats: dict) -> None:
    sheet = workbook.add_worksheet(f"Caster冲{TARGET_RANK}绊")
    base_headers = [
        "图鉴号", "svtId", "名称", "星级", "cost", "当前绊", "当前绊点", f"绊{TARGET_RANK}需求",
        f"距绊{TARGET_RANK}还差", f"已达{TARGET_RANK}绊",
    ]
    column = {name: index for index, name in enumerate(base_headers)}
    headers = base_headers + [item["name"] for item in conditional]
    sheet.write_row(0, 0, headers, formats["header"])
    for offset, essence in enumerate(conditional):
        sheet.write_comment(
            0, len(base_headers) + offset,
            f"{essence['name']}：满破 +{essence['rate'] / 10:g}%（该从者满足限定条件时为 1）",
        )

    for index, item in enumerate(rows, start=1):
        sheet.write_number(index, column["图鉴号"], item["collectionNo"] or 0, formats["int"])
        sheet.write_number(index, column["svtId"], item["svtId"], formats["int"])
        sheet.write(index, column["名称"], item["name"] or "", formats["text"])
        sheet.write_number(index, column["星级"], item["rarity"] or 0, formats["star"])
        sheet.write_number(index, column["cost"], item["cost"] or 0, formats["int"])
        sheet.write_number(index, column["当前绊"], item["rank"], formats["int"])
        sheet.write_number(index, column["当前绊点"], item["points"], formats["int"])
        sheet.write_number(index, column[f"绊{TARGET_RANK}需求"], item["need9"], formats["int"])
        sheet.write_number(index, column[f"距绊{TARGET_RANK}还差"], item["remain"], formats["int"])
        sheet.write(index, column[f"已达{TARGET_RANK}绊"], "是" if item["reached"] else "否",
                    formats["ok"] if item["reached"] else formats["text"])
        for offset, essence in enumerate(conditional, start=len(base_headers)):
            sheet.write_number(index, offset, 1 if essence["ce_id"] in item["hits"] else 0, formats["flag"])

    if rows:
        remain_col = column[f"距绊{TARGET_RANK}还差"]
        sheet.conditional_format(1, remain_col, len(rows), remain_col, {"type": "data_bar", "bar_color": "#63C384"})
        if conditional:
            sheet.conditional_format(
                1, len(base_headers), len(rows), len(headers) - 1,
                {"type": "cell", "criteria": "==", "value": 1, "format": formats["hit"]},
            )
        sheet.autofilter(0, 0, len(rows), len(headers) - 1)
    sheet.freeze_panes(1, 3)
    for col, width in enumerate((8, 10, 26, 8, 8, 8, 11, 12, 13, 11)):
        sheet.set_column(col, col, width)
    for col in range(len(base_headers), len(headers)):
        sheet.set_column(col, col, 15)


def write_sheet_high_rank(workbook, rows: list[dict], servants: dict, formats: dict) -> None:
    sheet = workbook.add_worksheet("10绊以上Caster")
    headers = ["图鉴号", "svtId", "名称", "当前绊", "当前绊点", "距下一级还差", "本段进度"]
    thresholds = max((len(servants[item["svtId"]].get("bondGrowth") or []) for item in rows), default=0)
    headers += [f"绊{index}" for index in range(1, thresholds + 1)]
    sheet.write_row(0, 0, headers, formats["header"])

    for index, item in enumerate(rows, start=1):
        servant = servants[item["svtId"]]
        curve = servant.get("bondGrowth") or []
        rank = item["rank"]
        points = item["points"]
        start = 0 if rank <= 0 else curve[rank - 1]
        nxt = curve[rank] if rank < len(curve) else None
        remain = "" if nxt is None else max(0, nxt - points)
        span = (nxt - start) if nxt is not None else 0
        progress = 1.0 if nxt is None else round((points - start) / span, 4) if span else 1.0

        sheet.write_number(index, 0, item["collectionNo"] or 0, formats["int"])
        sheet.write_number(index, 1, item["svtId"], formats["int"])
        sheet.write(index, 2, item["name"] or "", formats["text"])
        sheet.write_number(index, 3, rank, formats["int"])
        sheet.write_number(index, 4, points, formats["int"])
        if remain == "":
            sheet.write(index, 5, "已满", formats["ok"])
        else:
            sheet.write_number(index, 5, remain, formats["int"])
        sheet.write_number(index, 6, progress, formats["percent"])
        for offset, value in enumerate(curve, start=7):
            sheet.write_number(index, offset, value, formats["int"])

    if rows:
        sheet.autofilter(0, 0, len(rows), len(headers) - 1)
        sheet.conditional_format(1, 5, len(rows), 5, {"type": "data_bar", "bar_color": "#F4A65C"})
    sheet.freeze_panes(1, 3)
    for col, width in enumerate((8, 10, 26, 8, 11, 13, 10), start=0):
        sheet.set_column(col, col, width)
    for col in range(7, 7 + thresholds):
        sheet.set_column(col, col, 12)


def main() -> None:
    parser = argparse.ArgumentParser(description="导出持有 Caster 的羁绊状况到 xlsx")
    parser.add_argument("login", type=Path, nargs="?", default=Path("ac(29).php"))
    parser.add_argument("--data", type=Path, default=Path("chaldea-data"))
    parser.add_argument("-o", "--output", type=Path, default=Path("caster_bond.xlsx"))
    args = parser.parse_args()

    login = pb.load_login(args.login)
    replaced = login["cache"]["replaced"]
    collection = {int(row["svtId"]): row for row in replaced.get("userSvtCollection") or []}
    owned = owned_svt_ids(login)
    servants = pb.load_servants(args.data)
    ce = pb.load_craft_essences(args.data)
    svt_names = pb.load_json(args.data / "mappings" / "svt_names.json")
    ce_names = pb.load_json(args.data / "mappings" / "ce_names.json")
    skills, funcs = load_effect_tables(args.data)

    conditional = collect_conditional_ces(ce, ce_names, skills, funcs)
    caster_rows = build_caster_rows(owned, collection, servants, svt_names, conditional)
    high_rank = [item for item in caster_rows if item["rank"] >= 10]
    high_rank.sort(key=lambda item: (-item["rank"], -item["points"]))

    workbook = xlsxwriter.Workbook(str(args.output))
    formats = {
        "header": workbook.add_format(
            {"bold": True, "bg_color": "#4F81BD", "font_color": "white", "border": 1, "align": "center", "valign": "vcenter"}
        ),
        "text": workbook.add_format({"border": 1}),
        "int": workbook.add_format({"border": 1, "num_format": "#,##0"}),
        "percent": workbook.add_format({"border": 1, "num_format": "0.0%"}),
        "ok": workbook.add_format({"border": 1, "font_color": "#2E7D32", "bold": True}),
        "flag": workbook.add_format({"border": 1, "align": "center", "num_format": "0"}),
        "star": workbook.add_format({"border": 1, "align": "center", "num_format": '"★"0'}),
        "hit": workbook.add_format({"bg_color": "#C6EFCE", "font_color": "#1B5E20", "bold": True}),
    }
    write_sheet_target(workbook, caster_rows, conditional, formats)
    write_sheet_high_rank(workbook, high_rank, servants, formats)
    workbook.close()

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    print(f"持有 Caster {len(caster_rows)} 名，其中绊等级 ≥10 的 {len(high_rank)} 名")
    print(f"未达 {TARGET_RANK} 绊 {sum(1 for item in caster_rows if not item['reached'])} 名，已写入 {args.output}")


if __name__ == "__main__":
    main()
