# -*- coding: utf-8 -*-
"""
家計調査（品目分類）第6表「都市階級・地方・都道府県庁所在市別 1世帯当たり年間の
品目別支出金額・購入数量（二人以上の世帯）」の Excel から、みかんの
購入数量・支出金額を都道府県庁所在市ごとに取り出し、都道府県に割り当てて
mikan_data.js に追記する（consumer 種別＝消費・価格ページ）。

使い方:
  python import_kakei.py fn0605.xls --emit-js mikan_data.js --append
  （品目を変える場合: --item りんご など。シート既定は「果物」）

依存: fetch_estat.py / import_tokusan.py（同じフォルダ）, pandas, xlrd
  ※ xlrd が無ければ:  pip install xlrd --break-system-packages
"""
import argparse
import datetime
import os
import re
import sys

import pandas as pd

from fetch_estat import city_to_pref, PREF
from import_tokusan import load_existing, write_js

SOURCE = "総務省「家計調査（家計収支編・品目分類）」(e-Stat)"


def norm(s):
    return ("" if s is None or (isinstance(s, float) and pd.isna(s)) else str(s)).replace("\u3000", "").replace(" ", "").strip()


def num(v):
    try:
        return float(str(v).replace(",", ""))
    except (TypeError, ValueError):
        return None


def detect_year(df):
    txt = " ".join(str(x) for x in df.values.flatten() if isinstance(x, str))
    m = re.search(r"令和\s*(元|\d+)\s*年", txt)
    if m:
        n = 1 if m.group(1) == "元" else int(m.group(1))
        return str(2018 + n) + "年"
    m = re.search(r"平成\s*(元|\d+)\s*年", txt)
    if m:
        n = 1 if m.group(1) == "元" else int(m.group(1))
        return str(1988 + n) + "年"
    m = re.search(r"(20\d\d)\s*年", txt)
    return (m.group(1) + "年") if m else "—"


def find_item_cols(df, item):
    """品目見出し行から item の 金額/数量 列を特定。"""
    target = norm(item)
    # 品目見出し行（「みかん」「りんご」を含む行）
    head_row = None
    for r in range(min(12, df.shape[0])):
        cells = [norm(df.iat[r, c]) for c in range(df.shape[1])]
        if target in cells and ("りんご" in cells or "オレンジ" in cells or "ぶどう" in cells):
            head_row = r
            break
    if head_row is None:
        sys.exit("品目見出し行を特定できませんでした（{} が見つからない）。".format(item))
    col = next(c for c in range(df.shape[1]) if norm(df.iat[head_row, c]) == target)
    # 金額/数量 の小見出し行を探す
    sub_row = None
    for r in range(head_row, min(head_row + 12, df.shape[0])):
        vals = [norm(df.iat[r, c]) for c in range(df.shape[1])]
        if vals.count("金額") >= 3 and "数量" in vals:
            sub_row = r
            break
    kingaku, suuryo = col, None
    if sub_row is not None:
        # item列が金額、その右隣が数量の想定。確認して調整
        if norm(df.iat[sub_row, col]) == "金額" and col + 1 < df.shape[1] and norm(df.iat[sub_row, col + 1]) == "数量":
            suuryo = col + 1
        else:
            # 念のため item列以降で最初の 金額,数量 ペアを採用
            for c in range(col, df.shape[1] - 1):
                if norm(df.iat[sub_row, c]) == "金額" and norm(df.iat[sub_row, c + 1]) == "数量":
                    kingaku, suuryo = c, c + 1
                    break
    if suuryo is None:
        suuryo = col + 1
    return kingaku, suuryo, sub_row


def city_section_bounds(df):
    """都道府県庁所在市セクションの開始行と、政令市セクション開始行（終端）を返す。"""
    start = end = None
    for r in range(df.shape[0]):
        joined = "".join(norm(df.iat[r, c]) for c in range(min(3, df.shape[1])))
        if start is None and "都道府県庁所在市" in joined:
            start = r + 1
        elif start is not None and "政令指定都市" in joined:
            end = r
            break
    return start, (end if end is not None else df.shape[0])


def build(path, sheet, item, name):
    df = pd.read_excel(path, sheet_name=sheet, engine="xlrd", header=None)
    year = detect_year(df)
    k_col, q_col, _ = find_item_cols(df, item)
    start, end = city_section_bounds(df)
    if start is None:
        sys.exit("『都道府県庁所在市』セクションが見つかりませんでした。")

    qty, exp = {}, {}
    used_cities = []
    for r in range(start, end):
        city = None
        for c in range(min(3, df.shape[1])):     # 都市名は列1付近
            t = norm(df.iat[r, c])
            if t.endswith("市") or t == "東京都区部":
                city = t
                break
        if not city:
            continue
        pid = city_to_pref(city)
        if not pid:
            continue                              # 政令市など対応外は除外
        q = num(df.iat[r, q_col]); e = num(df.iat[r, k_col])
        if q is not None:
            qty[str(pid)] = q
        if e is not None:
            exp[str(pid)] = e
        used_cities.append((city, pid))

    # 全国（national）
    natQ = natE = None
    for r in range(df.shape[0]):
        if any(norm(df.iat[r, c]) == "全国" for c in range(min(3, df.shape[1]))):
            natQ = num(df.iat[r, q_col]); natE = num(df.iat[r, k_col]); break

    ds = {
        "key": "kakei_" + re.sub(r"[^0-9A-Za-zぁ-んァ-ヶ一-龠]+", "", item)[:12],
        "name": name, "suffix": "",
        "category": "consumer",
        "measures": ["buy_qty", "spend"],
        "measure_labels": {"buy_qty": "購入数量（1世帯・年）", "spend": "支出金額（1世帯・年）"},
        "units": {"buy_qty": "g", "spend": "円"},
        "years": [year],
        "data": {"buy_qty": {year: qty}, "spend": {year: exp}},
        "national": {"buy_qty": {year: natQ}, "spend": {year: natE}},
        "note": "【値は都道府県庁所在市のもの（県全体の平均ではない）】二人以上の世帯・1世帯当たり年間。都道府県庁所在市の値を各県に割り当て（政令指定都市・地方ブロック・全国は除外）。家計の購入であり、産地ではなく消費側の指標です。",
        "source": SOURCE,
        "tables": [{"id": os.path.basename(path),
                    "title": "家計調査 品目分類 第6表 都市階級・地方・都道府県庁所在市別 1世帯当たり年間の品目別支出金額・購入数量（二人以上の世帯）／{}シート・{}".format(sheet, item),
                    "period": year}],
        "generated": datetime.date.today().isoformat(),
    }
    return ds, qty, exp, used_cities, year


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xls", help="家計調査 第6表のExcel（例: fn0605.xls）")
    ap.add_argument("--sheet", default="果物")
    ap.add_argument("--item", default="みかん")
    ap.add_argument("--name", default="みかん（家計消費・県庁所在市）")
    ap.add_argument("--emit-js", default=None)
    ap.add_argument("--append", action="store_true")
    a = ap.parse_args()

    ds, qty, exp, cities, year = build(a.xls, a.sheet, a.item, a.name)
    print("対象年: {} ／ 都道府県数: 数量 {} ・ 支出 {}".format(year, len(qty), len(exp)))
    top = sorted(qty.items(), key=lambda kv: -kv[1])[:5]
    print("購入数量 上位5:", "、".join("{}={}g".format(PREF[int(k)], int(v)) for k, v in top))

    if a.emit_js:
        datasets, meta_block = (load_existing(a.emit_js) if a.append else ([], None))
        datasets = [d for d in datasets if not str(d.get("key","")).startswith("kakei_")]
        datasets.append(ds)
        write_js(a.emit_js, datasets, meta_block)
        print("HTML用データ: {} に書き出し（データセット {} 件）".format(a.emit_js, len(datasets)))


if __name__ == "__main__":
    main()
