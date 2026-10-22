"""准备 next-is-who 用的持有从者数据。

1. 从登录缓存读取当前持有（userSvt ∪ 灵基保管室），以及图鉴上的绊等级、绊点。
2. 把每名从者关联到羁绊礼装的户口。灵基 / 已持有灵衣的 individuality 会整表替换基础特性，
   户口是否命中按「单个灵基自己的特性集合」判断，不能把不同灵基的特性并成一份再去配。
"""

from __future__ import annotations

import csv
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "caster-grand-order"))

import parse_bond as pb

DATA_DIR = ROOT / "chaldea-data"
LOGIN = ROOT / "caster-grand-order" / "ac(29).php"
OUT_DIR = Path(__file__).resolve().parent

CLASS_NAME = {
    1: "Saber", 2: "Archer", 3: "Lancer", 4: "Rider", 5: "Caster",
    6: "Assassin", 7: "Berserker", 8: "Shielder", 9: "Ruler", 10: "Alterego",
    11: "Avenger", 23: "MoonCancer", 25: "Foreigner", 28: "Pretender",
    33: "Beast", 38: "Beast",
}


def trait_ids(items) -> list[int]:
    return [int(item["id"]) for item in (items or []) if item.get("id") is not None]


def form_traits(servant: dict) -> tuple[list[int], list[dict]]:
    """基础特性，以及每个再临阶段实际生效的特性（有 individuality 时整表替换）。"""
    base = trait_ids(servant.get("traits"))
    overrides = (
        ((servant.get("ascensionAdd") or {}).get("individuality") or {}).get("ascension") or {}
    )
    stages = sorted({int(key) for key in overrides} | {0, 1, 2, 3})
    forms = []
    for stage in stages:
        replaced = overrides.get(str(stage))
        forms.append({
            "form": f"再临{stage}",
            "limitCount": stage,
            "replacesBase": replaced is not None,
            "traits": trait_ids(replaced) if replaced is not None else list(base),
        })
    return base, forms


def owned_costumes(collection_row: dict, servant: dict) -> list[dict]:
    owned = {int(item) for item in (collection_row.get("costumeIds") or [])}
    overrides = (
        ((servant.get("ascensionAdd") or {}).get("individuality") or {}).get("costume") or {}
    )
    forms = []
    for costume_id in sorted(owned):
        replaced = overrides.get(str(costume_id))
        if not replaced:
            continue
        forms.append({
            "form": f"灵衣{costume_id}",
            "costumeId": costume_id,
            "replacesBase": True,
            "traits": trait_ids(replaced),
        })
    return forms


def matches_condition(traits: set[int], groups: list[list[int]], single: list[int]) -> bool:
    """overWriteTvalsList：组内 AND、组间 OR。只有 functvals 时按任一命中。"""
    if groups:
        return any(all(trait in traits for trait in group) for group in groups)
    return any(trait in traits for trait in single)


def collect_bond_ces(essences: dict, ce_names: dict, skills: dict, funcs: dict) -> list[dict]:
    """达芬奇工房里带户口条件的羁绊礼装，数值取满破。"""
    result = []
    for essence in essences.values():
        if essence.get("flag") != "svtEquipManaExchange":
            continue
        best = None
        for skill_ref in essence.get("skills") or []:
            if skill_ref.get("condLimitCount") not in (None, 4):
                continue
            skill = skills.get(skill_ref["id"])
            if not skill:
                continue
            for function in skill.get("functions") or []:
                meta = funcs.get(function["funcId"]) or {}
                if meta.get("funcType") != "servantFriendshipUp":
                    continue
                traits = trait_ids(meta.get("functvals"))
                groups = [trait_ids(group) for group in meta.get("overWriteTvalsList") or []]
                if not traits and not groups:
                    continue
                rates = [
                    int(item.get("RateCount") or 0)
                    for item in (function.get("svals") or [{}])
                ]
                rate = max(rates) if rates else 0
                if best is None or rate > best["rate"]:
                    best = {
                        "ceId": essence["id"],
                        "rate": rate,
                        "traits": traits,
                        "groups": groups,
                    }
        if best:
            best["name"] = pb.localized_name(ce_names, essence.get("name")) or str(essence["id"])
            result.append(best)
    result.sort(key=lambda item: item["name"])
    return result


def write_bond_map(plan: list[dict]) -> Path:
    """把挖取图写成单文件 HTML。数据嵌在页面里，换机器直接打开。"""
    path = OUT_DIR / "bond_map.html"
    html = path.read_text(encoding="utf-8")
    blob = json.dumps(plan, ensure_ascii=False, separators=(",", ":")).replace("<", "\\u003c")
    marker = '<script id="servants-data" type="application/json">'
    start = html.find(marker)
    if start < 0:
        raise SystemExit("bond_map.html 缺少 servants-data")
    start += len(marker)
    end = html.find("</script>", start)
    path.write_text(html[:start] + blob + html[end:], encoding="utf-8")
    return path


def bond_cap(rank: int, exceed: int, curve_len: int) -> int:
    """已开放的羁绊上限。初始上限是 10，friendshipExceedCount 是其后每次开放。

    玛修的绊 5 到绊 10 由剧情开放，不占道具次数，计数和普通从者一样从 10 起算。
    """
    return min(curve_len, max(rank, 10 + exceed))


def saint_quartz(rarity: int, rank: int) -> int:
    """绊等级对应的圣晶石。0 星按 4 星发。绊 1–5 和绊 10 没有石，绊 11 起每级 30。"""
    if rarity == 0:
        rarity = 4
    if rank <= 5 or rank == 10:
        return 0
    if rank >= 11:
        return 30
    if rarity >= 5:
        return 3
    if rarity == 4:
        return 2
    if rarity == 3:
        return 1
    if rarity == 2:
        return 1 if rank >= 8 else 0
    if rarity == 1:
        return 1 if rank == 9 else 0
    return 0


def bond_rewards(servant: dict, items: dict, essences: dict, ce_names: dict,
                 item_names: dict, current_rank: int) -> list[dict]:
    """每个绊等级的奖励。已领取 = 当前绊等级已经到达这一级。"""
    gifts = servant.get("bondGifts") or {}
    ranks = set(int(key) for key in gifts) | set(range(1, len(servant.get("bondGrowth") or []) + 1))
    rows = []
    for rank in sorted(ranks):
        rewards = []
        for gift in gifts.get(str(rank)) or []:
            num = int(gift.get("num") or 0)
            object_id = int(gift["objectId"])
            if gift.get("type") == "servant":
                essence = essences.get(object_id) or {}
                name = pb.localized_name(ce_names, essence.get("name")) or str(object_id)
                rewards.append({
                    "kind": "craftEssence", "objectId": object_id, "name": name, "num": num,
                })
            else:
                item = items.get(object_id) or {}
                name = pb.localized_name(item_names, item.get("name")) or item.get("name") or str(object_id)
                reward = {
                    "kind": item.get("type") or "item",
                    "objectId": object_id,
                    "name": name,
                    "num": num,
                }
                if item.get("type") == "svtCoin":
                    reward["coinSvtId"] = item.get("value")
                rewards.append(reward)
        rows.append({
            "rank": rank,
            "claimed": rank <= current_rank,
            "saintQuartz": saint_quartz(int(servant.get("rarity") or 0), rank),
            "rewards": rewards,
        })
    return rows


def reward_text(rewards: list[dict]) -> str:
    return "、".join(f"{item['name']}×{item['num']}" for item in rewards)


def condition_text(item: dict, trait_names: dict) -> str:
    def label(trait_id: int) -> str:
        return trait_names.get(trait_id, str(trait_id))

    if item["groups"]:
        return " 或 ".join(" 且 ".join(label(trait) for trait in group) for group in item["groups"])
    return " 或 ".join(label(trait) for trait in item["traits"])


def main() -> None:
    login = pb.load_login(LOGIN)
    replaced = login["cache"]["replaced"]
    owned: set[int] = set()
    limit_count: dict[int, int] = {}
    for key in ("userSvt", "userSvtStorage"):
        for row in replaced.get(key) or []:
            svt_id = int(row["svtId"])
            owned.add(svt_id)
            limit_count[svt_id] = max(limit_count.get(svt_id, 0), int(row.get("limitCount") or 0))

    collection = {int(row["svtId"]): row for row in replaced.get("userSvtCollection") or []}
    servants = pb.load_servants(DATA_DIR)
    essences = pb.load_craft_essences(DATA_DIR)
    svt_names = pb.load_json(DATA_DIR / "mappings" / "svt_names.json")
    ce_names = pb.load_json(DATA_DIR / "mappings" / "ce_names.json")
    item_names = pb.load_json(DATA_DIR / "mappings" / "item_names.json")
    items = {row["id"]: row for row in pb.load_json(DATA_DIR / "dist" / "items.json")}
    trait_table = pb.load_json(DATA_DIR / "mappings" / "trait.json")
    trait_names = {
        int(key): (value.get("CN") or value.get("JP") or key)
        for key, value in trait_table.items()
    }
    skills, funcs = (
        {row["id"]: row for row in pb.load_json(DATA_DIR / "dist" / "baseSkills.json")},
        {row["funcId"]: row for row in pb.load_json(DATA_DIR / "dist" / "baseFunctions.json")},
    )
    bond_ces = collect_bond_ces(essences, ce_names, skills, funcs)

    rows = []
    for svt_id in sorted(owned):
        servant = servants.get(svt_id)
        if not servant or not servant.get("bondGrowth"):
            continue
        collected = collection.get(svt_id) or {}
        base, forms = form_traits(servant)
        forms.extend(owned_costumes(collected, servant))
        hits = []
        for essence in bond_ces:
            matched = [
                form["form"] for form in forms
                if matches_condition(set(form["traits"]), essence["groups"], essence["traits"])
            ]
            if matched:
                hits.append({
                    "ceId": essence["ceId"],
                    "name": essence["name"],
                    "rate": essence["rate"],
                    "forms": matched,
                })
        curve = servant.get("bondGrowth") or []
        rank = int(collected.get("friendshipRank") or 0)
        points = int(collected.get("friendship") or 0)
        rewards = bond_rewards(servant, items, essences, ce_names, item_names, rank)
        exceed = int(collected.get("friendshipExceedCount") or 0)
        rows.append({
            "collectionNo": servant.get("collectionNo"),
            "svtId": svt_id,
            "name": pb.localized_name(svt_names, servant.get("name")),
            "classId": servant.get("classId"),
            "className": CLASS_NAME.get(servant.get("classId"), str(servant.get("classId"))),
            "rarity": servant.get("rarity"),
            "cost": servant.get("cost"),
            "bondRank": rank,
            "bondPoints": points,
            "bondExceedCount": exceed,
            "bondCap": bond_cap(rank, exceed, len(curve)),
            "bondGrowth": curve,
            "currentLimit": limit_count.get(svt_id, 0),
            "baseTraits": base,
            "forms": forms,
            "bondCes": hits,
            "bondRewards": rewards,
        })
    rows.sort(key=lambda item: (item["collectionNo"] or 10**9, item["svtId"]))

    payload = {
        "source": LOGIN.name,
        "ownedServants": len(rows),
        "bondCes": [
            {
                "ceId": item["ceId"],
                "name": item["name"],
                "rate": item["rate"],
                "condition": condition_text(item, trait_names),
                "groups": item["groups"],
                "traits": item["traits"],
            }
            for item in bond_ces
        ],
        "servants": rows,
    }
    plan = [
        {
            "id": item["svtId"],
            "name": item["name"],
            "cls": item["className"],
            "rarity": item["rarity"],
            "rank": item["bondRank"],
            "points": item["bondPoints"],
            "cap": item["bondCap"],
            "curve": item["bondGrowth"],
        }
        for item in rows
    ]
    (OUT_DIR / "bond_plan.json").write_text(
        json.dumps(plan, ensure_ascii=False, separators=(",", ":")),
        encoding="utf-8",
    )
    out_json = OUT_DIR / "owned_bond.json"
    out_json.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    out_csv = OUT_DIR / "owned_bond.csv"
    with out_csv.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow([
            "图鉴号", "svtId", "名称", "职阶", "星级", "cost",
            "绊等级", "绊点", "可享礼装", "触发灵基",
        ])
        for item in rows:
            writer.writerow([
                item["collectionNo"], item["svtId"], item["name"], item["className"],
                item["rarity"], item["cost"], item["bondRank"], item["bondPoints"],
                " / ".join(hit["name"] for hit in item["bondCes"]),
                " / ".join(
                    f"{hit['name']}[{','.join(hit['forms'])}]" for hit in item["bondCes"]
                ),
            ])

    max_rank = max(level["rank"] for item in rows for level in item["bondRewards"])
    out_rewards = OUT_DIR / "bond_quartz.csv"
    with out_rewards.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(
            ["图鉴号", "svtId", "名称", "星级", "当前绊", "未领圣晶石"]
            + [f"绊{rank}" for rank in range(1, max_rank + 1)]
        )
        for item in rows:
            by_rank = {level["rank"]: level["saintQuartz"] for level in item["bondRewards"]}
            remaining = sum(
                level["saintQuartz"] for level in item["bondRewards"] if not level["claimed"]
            )
            writer.writerow([
                item["collectionNo"], item["svtId"], item["name"], item["rarity"],
                item["bondRank"], remaining,
                *[by_rank.get(rank, "") for rank in range(1, max_rank + 1)],
            ])

    aesc = next(item for item in rows if item["name"] == "雨之魔女梣")
    print(f"持有从者 {len(rows)} 人，羁绊礼装 {len(bond_ces)} 件")
    remaining = sum(
        level["saintQuartz"] for item in rows for level in item["bondRewards"] if not level["claimed"]
    )
    map_path = write_bond_map(plan)
    print(f"写入 {out_json.name}、{out_csv.name}、{out_rewards.name}、{map_path.name}")
    print(f"未领圣晶石合计 {remaining}")
    opened = [item for item in rows if item["bondCap"] > 10]
    print(f"已开放上限 {len(opened)} 人，玛修上限 {next(item['bondCap'] for item in rows if item['svtId']==800100)}")
    odd = [
        item["name"] for item in rows
        if any(
            reward["num"] != 5
            for level in item["bondRewards"] if level["rank"] == 1
            for reward in level["rewards"] if reward["kind"] == "svtCoin"
        )
    ]
    if odd:
        print("绊1从者币不是 5 的：", "、".join(odd))
    print(f"雨之魔女梣 绊{aesc['bondRank']} {aesc['bondPoints']}")
    print("  " + " ".join(
        f"{level['rank']}:{level['saintQuartz']}" for level in aesc["bondRewards"] if level["saintQuartz"]
    ))
    for rarity in (0, 1, 2):
        sample = next((item for item in rows if item["rarity"] == rarity), None)
        if not sample:
            continue
        quartz = " ".join(
            f"{level['rank']}:{level['saintQuartz']}" for level in sample["bondRewards"] if level["rank"] <= 11
        )
        print(f"{rarity}星 {sample['name']} {quartz}")
    for hit in aesc["bondCes"]:
        text = next(item["condition"] for item in payload["bondCes"] if item["ceId"] == hit["ceId"])
        print(f"  {hit['name']}（{text}）← {', '.join(hit['forms'])}")


if __name__ == "__main__":
    main()
