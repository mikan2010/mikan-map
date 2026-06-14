# -*- coding: utf-8 -*-
"""
青果物卸売市場調査（産地別）→ 産地都道府県別の卸売データセット化（mikan_data.js 追記）

この表は地域軸が「<消費市場>_産地計」または「<消費市場>_産地計_<産地>」という
複合値（例: 札幌市_産地計_和歌山）になっている。各 <産地> について
  ・卸売数量 = 全消費市場の合計
  ・卸売価格 = 数量で加重平均
を計算し、産地都道府県別の値にする（対象月=計を使用）。
※主要消費市場で取り扱われた分の集計のため、各産地の全出荷ではなく
  「主要市場での取扱量・平均卸売価格」を表す。

使い方:
  set ESTAT_APP_ID=（あなたのアプリID）
  python import_wholesale.py --stats-id 0002114420 --name "みかん（卸売・産地別）" --emit-js mikan_data.js --append

依存: fetch_estat.py / import_tokusan.py（同じフォルダ）
"""
import argparse
import datetime
import json
import os
import re
import sys

from fetch_estat import fetch_raw, parse, match_pref, survey_year, _text
from import_tokusan import load_existing, write_js  # 追記・書き出しを共用

SOURCE = "農林水産省「青果物卸売市場調査（産地別）」(e-Stat)"
SPLIT = "_産地計_"


def num(v):
    try:
        return float(str(v).replace(",", ""))
    except (TypeError, ValueError):
        return None


def slugify(name):
    return "oroshi_" + re.sub(r"[^0-9A-Za-zぁ-んァ-ヶ一-龠]+", "", name)[:24]


def build(app_id, sid, name, by="origin"):
    # 対象月=計 のみ取得（cat01=計 を想定。違っても全件取得→後段で計を優先）
    class_obj, values, ti = fetch_raw(app_id, sid, {"cdCat01": "1001"})
    p = parse(class_obj, values)
    nm = p["name_maps"]

    # 数量・価格の次元、月次元、複合の地域次元を特定
    qp_dim = None
    for did, codes in nm.items():
        vals = set(codes.values())
        if "数量" in vals and "価格" in vals:
            qp_dim = did
            break
    area_dim = p["area_id"]
    month_dim = None
    for did, codes in nm.items():
        if did in (qp_dim, area_dim):
            continue
        if "計" in codes.values() and any(re.fullmatch(r"\d+", str(v) or "") for v in codes.values()):
            month_dim = did
            break
    if not qp_dim or not area_dim:
        sys.exit("数量・価格／地域（複合）次元を特定できませんでした。--inspect で構造をご確認ください。")

    # セル（地域複合コード）→ {数量, 価格, 単位}
    cells = {}
    unit_q, unit_p = "", "円/kg"
    for v in values:
        if month_dim and str(v.get("@" + month_dim, "")) not in ("", "1001"):
            continue  # 対象月=計 以外は除外（コード1001=計を想定）
        ac = v.get("@" + area_dim)
        qp = nm[qp_dim].get(v.get("@" + qp_dim))
        val = num(v.get("$"))
        if ac is None or qp is None or val is None:
            continue
        d = cells.setdefault(ac, {})
        if qp == "数量":
            d["q"] = val
            unit_q = v.get("@unit", unit_q) or unit_q
        elif qp == "価格":
            d["p"] = val
            unit_p = v.get("@unit", unit_p) or unit_p

    # 集計：by="origin"（産地別）／ by="market"（消費市場別）
    agg = {}
    for ac, d in cells.items():
        cname = nm[area_dim].get(ac, "")
        if by == "market":
            if SPLIT in cname or not cname.endswith("_産地計"):
                continue                       # 産地内訳行は除外、市場合計行のみ採用
            key_name = cname[:-len("_産地計")].strip()    # 「東京都_産地計」→「東京都」
        else:
            if SPLIT not in cname:
                continue                       # 「<市場>_産地計」= 市場小計はスキップ
            key_name = cname.split(SPLIT)[-1].strip()     # 産地
        pid = match_pref(None, key_name)
        if not pid:
            continue
        q = d.get("q")
        pr = d.get("p")
        a = agg.setdefault(pid, {"q": 0.0, "wsum": 0.0, "wq": 0.0, "praw": None})
        if q is not None:
            a["q"] += q
            if pr is not None:
                a["wsum"] += pr * q
                a["wq"] += q
        elif pr is not None:
            a["praw"] = pr                      # 数量欠落でも価格があれば保持

    qty = {str(pid): round(a["q"], 1) for pid, a in agg.items() if a["q"] > 0}
    price = {}
    for pid, a in agg.items():
        if a["wq"] > 0:
            price[str(pid)] = round(a["wsum"] / a["wq"], 1)
        elif a["praw"] is not None:
            price[str(pid)] = round(a["praw"], 1)
    tot_q = round(sum(a["q"] for a in agg.values()), 1)
    tot_wsum = sum(a["wsum"] for a in agg.values())
    tot_wq = sum(a["wq"] for a in agg.values())
    nat_price = round(tot_wsum / tot_wq, 1) if tot_wq else None

    year = (survey_year(ti) or "—")
    year = year + "年" if re.fullmatch(r"\d{4}", year or "") else "—"
    ttitle = (_text(ti.get("STATISTICS_NAME")) + " / " + _text(ti.get("TITLE"))).strip(" /")

    if by == "market":
        mlabels = {"数量": "卸売数量（市場計）", "価格": "卸売価格（市場平均）"}
        note = "各消費市場（中央卸売市場等）でのみかんの卸売価格・数量＝『各都市でいくらで売られているか』。市場名をその所在都道府県に割り当て。"
    else:
        mlabels = {"数量": "卸売数量", "価格": "卸売価格（数量加重平均）"}
        note = "各産地が主要消費市場へ出した分の合計と加重平均価格＝『どこ産か（産地別）』。卸売価格は数量加重平均（取扱量の多い市場ほど強く反映＝実勢に近い）。産地不明・その他は除外。"

    ds = {
        "key": slugify(name), "name": name, "suffix": "",
        "category": "market",
        "measures": ["数量", "価格"],
        "measure_labels": mlabels,
        "units": {"数量": unit_q or "t", "価格": unit_p or "円/kg"},
        "years": [year],
        "data": {"数量": {year: qty}, "価格": {year: price}},
        "national": {"数量": {year: tot_q}, "価格": {year: nat_price}},
        "note": note,
        "source": SOURCE,
        "tables": [{"id": sid, "title": ttitle, "period": year}],
        "stats_ids": [sid],
        "filters": {"対象月": "計", "集計": ("消費市場別" if by == "market" else "産地別")},
        "generated": datetime.date.today().isoformat(),
    }
    return ds, qty, price


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--stats-id", required=True)
    ap.add_argument("--by", choices=["origin", "market"], default="origin",
                    help="origin=産地別（どこ産か・既定）, market=消費市場別（各都市でいくら）")
    ap.add_argument("--name", default=None)
    ap.add_argument("--app-id", default=os.environ.get("ESTAT_APP_ID"))
    ap.add_argument("--emit-js", default=None)
    ap.add_argument("--append", action="store_true")
    ap.add_argument("--out", default=None, help="生データJSONの保存先（任意）")
    a = ap.parse_args()
    if not a.app_id:
        sys.exit("アプリIDを ESTAT_APP_ID 環境変数か --app-id で指定してください。")

    name = a.name or ("みかん（卸売・消費市場別）" if a.by == "market" else "みかん（卸売・産地別）")
    ds, qty, price = build(a.app_id, a.stats_id, name, a.by)
    unit = "市場" if a.by == "market" else "産地県"
    print("{}数: 数量 {} / 価格 {}".format(unit, len(qty), len(price)))
    top = sorted(price.items(), key=lambda kv: -kv[1])[:5]
    from fetch_estat import PREF
    print("卸売価格 上位5:", "、".join("{}={}".format(PREF[int(k)], v) for k, v in top))

    if a.out:
        json.dump(ds, open(a.out, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
        print("保存:", a.out)

    if a.emit_js:
        datasets, meta_block = (load_existing(a.emit_js) if a.append else ([], None))
        datasets = [d for d in datasets if d.get("key") != ds["key"]]
        datasets.append(ds)
        write_js(a.emit_js, datasets, meta_block)
        print("HTML用データ: {} に書き出し（データセット {} 件）".format(a.emit_js, len(datasets)))


if __name__ == "__main__":
    main()
