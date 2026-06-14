# -*- coding: utf-8 -*-
"""
特産果樹生産動態等調査（種類別栽培状況・かんきつ類・都道府県別）の Excel から、
品種×都道府県の収穫量・栽培面積・出荷量・加工向け・主要産地を読み取り、
産地・品種を深掘りするページ用の `tokusan_data.js`（window.TOKUSAN）を生成する。

使い方:
  python build_explore.py p001-01-021.xlsx --out tokusan_data.js

依存: fetch_estat.py（同じフォルダ）, pandas, openpyxl
※ この表は温州みかんを含まない中晩柑類。温州みかんは作物統計（mikan_data.js）側を参照。
"""
import argparse
import datetime
import json
import re
import sys

import pandas as pd

from fetch_estat import match_pref, PREF


def num(v):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    if f != f:        # NaN
        return None
    return round(f, 1)


def shortname(nm):
    nm = nm.strip().replace("\u3000", " ").strip()
    m = re.search(r"[（(]([一-龠々ヶ]{1,8})[）)]", nm)   # 最初の漢字括弧
    if m:
        return m.group(1)
    return re.sub(r"[ \u3000]+$", "", nm)


def build(path):
    df = pd.read_excel(path, header=None)
    # ヘッダ「品目名」行を特定（列1）→ データはその下
    hdr = None
    for i in range(min(15, len(df))):
        if any((str(x).strip() == "品目名") for x in df.iloc[i].tolist() if x is not None):
            hdr = i
            break
    if hdr is None:
        sys.exit("『品目名』ヘッダが見つかりませんでした。")
    name_col = df[1].ffill()           # 品目名は縦結合 → 前方補完

    varieties = {}
    order = []
    for i in range(hdr + 3, len(df)):
        nm = name_col.iloc[i]
        if pd.isna(nm):
            continue
        nm = str(nm).strip()
        if not nm or "果樹計" in nm:
            continue
        pref = df.iat[i, 2]
        if pd.isna(pref):
            continue
        pid = match_pref(None, str(pref).strip())   # 「計」「○○計」は None → 除外
        if not pid:
            continue
        rec = {"harvest": num(df.iat[i, 4]), "area": num(df.iat[i, 3]),
               "ship": num(df.iat[i, 5]), "proc": num(df.iat[i, 6]),
               "cities": (str(df.iat[i, 7]).strip() if not pd.isna(df.iat[i, 7]) else "")}
        if nm not in varieties:
            varieties[nm] = {}
            order.append(nm)
        varieties[nm][str(pid)] = rec

    out_vars = []
    for nm in order:
        prefs = varieties[nm]
        total = {}
        for k in ("harvest", "area", "ship", "proc"):
            vals = [p[k] for p in prefs.values() if p.get(k) is not None]
            total[k] = round(sum(vals), 1) if vals else None
        out_vars.append({"name": nm, "short": shortname(nm), "total": total, "pref": prefs})

    out_vars.sort(key=lambda v: -(v["total"].get("harvest") or 0))
    grand = {}
    for k in ("harvest", "area", "ship", "proc"):
        grand[k] = round(sum((v["total"].get(k) or 0) for v in out_vars), 1)

    return {
        "period": "2021年",
        "source": "農林水産省「特産果樹生産動態等調査」(e-Stat)",
        "table": "令和3年産 特産果樹生産出荷実績 種類別栽培状況 かんきつ類（都道府県別）",
        "file": __import__("os").path.basename(path),
        "note": "温州みかんを除く中晩柑類。値は2021年産（令和3年産）。各品種の主要産地（市町村）付き。",
        "measures": ["harvest", "area", "ship", "proc"],
        "measure_labels": {"harvest": "収穫量", "area": "栽培面積", "ship": "出荷量", "proc": "加工向け"},
        "units": {"harvest": "t", "area": "ha", "ship": "t", "proc": "t"},
        "grand": grand,
        "varieties": out_vars,
        "generated": datetime.date.today().isoformat(),
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx", help="特産果樹 かんきつ類 都道府県別 Excel（例: p001-01-021.xlsx）")
    ap.add_argument("--out", default="tokusan_data.js")
    a = ap.parse_args()

    data = build(a.xlsx)
    print("品種数: {} ／ 収穫量合計: {} t（2021）".format(len(data["varieties"]), data["grand"]["harvest"]))
    print("収穫量トップ5:", "、".join("{}({:.0f}t)".format(v["short"], v["total"]["harvest"] or 0) for v in data["varieties"][:5]))
    with open(a.out, "w", encoding="utf-8") as f:
        f.write("window.TOKUSAN = " + json.dumps(data, ensure_ascii=False) + ";\n")
    print("書き出し:", a.out)


if __name__ == "__main__":
    main()
