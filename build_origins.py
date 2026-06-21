# -*- coding: utf-8 -*-
"""
特産果樹（種類別×都道府県）の「主要産地（市町村名）」を市区町村役場の緯度経度に
ジオコーディングし、市町村→品種の逆引きを作って origins_data.js（window.ORIGINS）を生成する。
産地（市町村）のドットを日本地図に置く explore 用のデータ。

使い方:
  python build_origins.py p001-01-021.xlsx shikuchoson-lng-lat.csv --out origins_data.js

座標CSV（市区町村役場の緯度経度）:
  https://github.com/UG/JapanCityTownHall_lat_long （MIT License）
  列: 都道府県(かな),市区町村(かな),都道府県(漢字),市区町村(漢字),緯度,経度

依存: build_explore.py / fetch_estat.py（同じフォルダ）, pandas, openpyxl
"""
import argparse
import csv
import datetime
import json
import re

from build_explore import build as build_tokusan
from fetch_estat import PREF

GEO_SRC = {"name": "市区町村役場の緯度経度（UG/JapanCityTownHall_lat_long, MIT License）",
           "url": "https://github.com/UG/JapanCityTownHall_lat_long"}


def load_coords(path):
    byp = {}
    with open(path, encoding="utf-8") as f:
        for row in csv.reader(f):
            if len(row) < 6:
                continue
            pref, muni = row[2].strip(), row[3].strip()
            try:
                la, ln = float(row[4]), float(row[5])
            except ValueError:
                continue
            byp.setdefault(pref, []).append((muni, la, ln))
    return byp


def geocode(byp, pref, name):
    name = re.sub(r"等$", "", str(name)).strip()
    if not name or name.startswith("その他"):
        return None
    tries = [name]
    if name.endswith("市"):
        tries += [name[:-1] + "町", name[:-1] + "村"]
    elif name.endswith("町"):
        tries += [name[:-1] + "市", name[:-1] + "村"]
    elif name.endswith("村"):
        tries += [name[:-1] + "市", name[:-1] + "町"]
    lst = byp.get(pref, [])
    for t in tries:
        for mu, la, ln in lst:
            if mu == t:
                return (la, ln)
        cand = [(mu, la, ln) for mu, la, ln in lst if mu.endswith(t)]
        if cand:
            return (cand[0][1], cand[0][2])
        w = [(la, ln) for mu, la, ln in lst if mu.startswith(t) and mu.endswith("区")]
        if w:
            return (round(sum(p[0] for p in w) / len(w), 6), round(sum(p[1] for p in w) / len(w), 6))
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx", help="特産果樹 かんきつ類 都道府県別 Excel")
    ap.add_argument("coords", help="shikuchoson-lng-lat.csv（市区町村役場の緯度経度）")
    ap.add_argument("--out", default="origins_data.js")
    a = ap.parse_args()

    tk = build_tokusan(a.xlsx)
    byp = load_coords(a.coords)

    idx = {}   # (pid, muni) -> list of variety dicts
    for v in tk["varieties"]:
        for pid, rec in v["pref"].items():
            cities = rec.get("cities") or ""
            for c in re.split(r"[、,／/]", cities):
                c = c.strip()
                if not c or c.startswith("その他") or c in ("―", "-"):
                    continue
                idx.setdefault((int(pid), c), []).append(
                    {"short": v["short"], "name": v["name"], "ph": rec.get("harvest")})

    munis = []
    miss = 0
    for (pid, name), vs in idx.items():
        g = geocode(byp, PREF[pid], name)
        if not g:
            miss += 1
            continue
        seen = {}
        for d in vs:                       # 同一品種の重複除去
            seen[d["short"]] = d
        vlist = sorted(seen.values(), key=lambda d: -(d["ph"] or 0))
        munis.append({"name": name, "pref": PREF[pid], "pid": pid,
                      "lat": round(g[0], 6), "lng": round(g[1], 6),
                      "nvar": len(vlist), "varieties": vlist})

    munis.sort(key=lambda m: -m["nvar"])
    data = {
        "period": tk["period"],
        "source_stat": tk["source"] + "／" + tk["table"],
        "source_geo": GEO_SRC,
        "note": "各市町村は、特産果樹（中晩柑）でその品種の『主要産地』に挙げられた市町村。品種の数値は所属県の値で、市町村別の数量ではありません。",
        "count": len(munis),
        "municipalities": munis,
        "generated": datetime.date.today().isoformat(),
    }
    with open(a.out, "w", encoding="utf-8") as f:
        f.write("window.ORIGINS = " + json.dumps(data, ensure_ascii=False) + ";\n")
    print("市町村ドット: {} 件（未ジオコーディング {} 件）".format(len(munis), miss))
    print("品種数が多い産地 上位5:", "、".join("{}({}品種)".format(m["name"], m["nvar"]) for m in munis[:5]))
    print("書き出し:", a.out)


if __name__ == "__main__":
    main()
