"""解析国服/台服登录缓存（ac.php）里的当前绊。

解码方式和 Chaldea 文档一致：URL 解码 -> Base64 解码 -> JSON。
只读取 cache.replaced.userSvtCollection 的绊字段，再用 chaldea-data 的
从者 bondGrowth 算出距离下一级还差多少。
"""

from __future__ import annotations

import argparse
import base64
import csv
import json
import sys
from pathlib import Path
from urllib.parse import unquote_to_bytes


def load_login(path: Path) -> dict:
    raw = path.read_bytes().strip()
    if raw[:1] in (b"{", b"["):
        return json.loads(raw)
    decoded = base64.b64decode(unquote_to_bytes(raw))
    return json.loads(decoded)


def load_json(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


def localized_name(table: dict, jp_name: str | None) -> str | None:
    if not jp_name:
        return None
    entry = table.get(jp_name) or {}
    return entry.get("CN") or entry.get("TW") or entry.get("NA") or jp_name


def load_servants(data_dir: Path) -> dict[int, dict]:
    servants: dict[int, dict] = {}
    for path in sorted((data_dir / "dist").glob("servants.*.json")):
        for servant in load_json(path):
            servants[servant["id"]] = servant
    return servants


def load_craft_essences(data_dir: Path) -> dict[int, dict]:
    essences: dict[int, dict] = {}
    for path in sorted((data_dir / "dist").glob("craftEssences.*.json")):
        for essence in load_json(path):
            essences[essence["id"]] = essence
    return essences


def bond_progress(points: int, rank: int, curve: list[int]) -> dict:
    """bondGrowth[i] 是升到绊等级 i+1 所需的累计绊点。"""
    max_rank = len(curve)
    previous = 0 if rank <= 0 else curve[rank - 1]
    if rank >= max_rank:
        return {
            "maxRank": max_rank,
            "nextRank": None,
            "currentThreshold": curve[-1] if curve else 0,
            "nextThreshold": None,
            "pointsIntoLevel": 0,
            "pointsToNext": 0,
            "progress": 1.0,
        }
    next_threshold = curve[rank]
    span = next_threshold - previous
    gained = points - previous
    return {
        "maxRank": max_rank,
        "nextRank": rank + 1,
        "currentThreshold": previous,
        "nextThreshold": next_threshold,
        "pointsIntoLevel": gained,
        "pointsToNext": next_threshold - points,
        "progress": round(gained / span, 4) if span else 1.0,
    }


def parse_bonds(login: dict, data_dir: Path) -> list[dict]:
    rows = login["cache"]["replaced"]["userSvtCollection"]
    servants = load_servants(data_dir)
    essences = load_craft_essences(data_dir)
    svt_names = load_json(data_dir / "mappings" / "svt_names.json")
    ce_names = load_json(data_dir / "mappings" / "ce_names.json")

    parsed = []
    for row in rows:
        servant = servants.get(int(row["svtId"]))
        curve = (servant or {}).get("bondGrowth")
        if not servant or not curve:
            continue

        points = int(row["friendship"])
        rank = int(row["friendshipRank"])
        progress = bond_progress(points, rank, curve)
        bond_equips = []
        for equip_id in servant.get("bondEquips") or []:
            essence = essences.get(equip_id)
            bond_equips.append(
                {
                    "id": equip_id,
                    "name": localized_name(ce_names, (essence or {}).get("name")),
                }
            )

        parsed.append(
            {
                "collectionNo": servant.get("collectionNo"),
                "svtId": servant["id"],
                "name": localized_name(svt_names, servant.get("name")),
                "bondRank": rank,
                "bondPoints": points,
                "bondExceedCount": int(row["friendshipExceedCount"]),
                **progress,
                "bondEquips": bond_equips,
            }
        )

    parsed.sort(key=lambda item: (-item["bondRank"], -item["bondPoints"], item["collectionNo"] or 0))
    return parsed


def write_csv(path: Path, rows: list[dict]) -> None:
    fields = [
        "collectionNo",
        "svtId",
        "name",
        "bondRank",
        "bondPoints",
        "bondExceedCount",
        "maxRank",
        "nextRank",
        "pointsIntoLevel",
        "pointsToNext",
        "progress",
        "currentThreshold",
        "nextThreshold",
        "bondEquipIds",
        "bondEquipNames",
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        for row in rows:
            equips = row["bondEquips"]
            writer.writerow(
                {
                    **{key: row.get(key) for key in fields if key not in {"bondEquipIds", "bondEquipNames"}},
                    "bondEquipIds": "|".join(str(item["id"]) for item in equips),
                    "bondEquipNames": "|".join(item["name"] or "" for item in equips),
                }
            )


def main() -> None:
    parser = argparse.ArgumentParser(description="解析登录缓存中的当前从者绊")
    parser.add_argument("login", type=Path, nargs="?", default=Path("ac(29).php"))
    parser.add_argument("--data", type=Path, default=Path("chaldea-data"), help="chaldea-data 目录")
    parser.add_argument("-o", "--output", type=Path, default=Path("bond.json"))
    parser.add_argument("--csv", type=Path, default=Path("bond.csv"))
    parser.add_argument("--min-rank", type=int, default=0, help="只保留绊等级大于等于该值的从者")
    args = parser.parse_args()

    login = load_login(args.login)
    rows = parse_bonds(login, args.data)
    if args.min_rank:
        rows = [row for row in rows if row["bondRank"] >= args.min_rank]

    payload = {
        "source": str(args.login),
        "servantCount": len(rows),
        "rankCounts": {},
        "servants": rows,
    }
    for row in rows:
        key = str(row["bondRank"])
        payload["rankCounts"][key] = payload["rankCounts"].get(key, 0) + 1

    args.output.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    write_csv(args.csv, rows)

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    print(f"从者 {len(rows)} 名，已写入 {args.output} 和 {args.csv}")
    print("绊等级人数:", " ".join(f"{rank}:{payload['rankCounts'].get(str(rank), 0)}" for rank in range(0, 17)))
    print("绊等级最高的 10 名:")
    for row in rows[:10]:
        next_text = "已满" if row["nextRank"] is None else f"下一级还差 {row['pointsToNext']}"
        print(f"  {row['collectionNo']:>4} {row['name']}  绊{row['bondRank']}  {row['bondPoints']}  {next_text}")


if __name__ == "__main__":
    main()
