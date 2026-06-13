#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fetch_estat.py
e-Stat の統計表を取得し、みかん産地マップ用データ(mikan_data.js)に整形する。

できること:
  --search "キーワード"   : 統計表ID(statsDataId)を日本語名で検索して一覧表示
  --stats-id ID [複数可]  : 取得。複数指定で年次をまたいで1データセットに統合（過去に遡れる）
  --pick "収穫量" [複数可] : コードでなく日本語名で絞り込み（既定: 収穫量）。表ごとに解決するので
                            スキーマの違う表（長期累年＋年別表）を混ぜてもOK。
  --emit-js mikan_data.js : HTML(mikan-map.html)が読み込むデータを書き出し（--append で追記）

セットアップ:
  pip install requests
  # PowerShell:  $env:ESTAT_APP_ID="あなたのappId"

ヒント:
  値が混在していると「⚠ 複数の区分が混在しています」と候補が表示されます。
  そこに出た品目名などを --pick に足して、1つに絞ってください（例: --pick みかん）。
"""
import os
import re
import sys
import json
import argparse
import datetime
import unicodedata

try:
    import requests
except ImportError:
    sys.exit("requests が必要です。先に:  pip install requests")

BASE = "https://api.e-stat.go.jp/rest/3.0/app/json"
API_DATA = BASE + "/getStatsData"
API_LIST = BASE + "/getStatsList"
CROP_STATS_CODE = "00500215"

PREF = ["", "北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県",
        "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県",
        "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県", "岐阜県",
        "静岡県", "愛知県", "三重県", "滋賀県", "京都府", "大阪府", "兵庫県",
        "奈良県", "和歌山県", "鳥取県", "島根県", "岡山県", "広島県", "山口県",
        "徳島県", "香川県", "愛媛県", "高知県", "福岡県", "佐賀県", "長崎県",
        "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県"]

PREF_INDEX = {n: i for i, n in enumerate(PREF) if n}


def _norm(s):
    return unicodedata.normalize("NFKC", s or "").replace(" ", "").replace("\u3000", "")


# 完全名・語幹（県/府/都を外した短縮形。北海道はそのまま）の両方を引けるよう用意
PREF_STEMS = {}
for _i, _n in enumerate(PREF):
    if not _n:
        continue
    PREF_STEMS[_n] = _i
    PREF_STEMS.setdefault(re.sub(r"[都府県]$", "", _n), _i)


def _is_pref_code(code):
    return len(code) == 5 and code.endswith("000") and code[:2].isdigit() and 1 <= int(code[:2]) <= 47


def _strip_national_prefix(nm):
    # 「全国_千葉」「全国 千葉」等の頭の全国＋区切りを除去
    return re.sub(r"^全国[ _\u3000・\-]*", "", nm)


def match_pref(code, name):
    """コード(PP000) か 都道府県名（短縮形/装飾/「全国_」接頭辞可）から JIS番号を返す。全国/不一致は None。"""
    if code and _is_pref_code(code):
        return int(code[:2])
    nm = _norm(name)
    if not nm or nm == "全国":
        return None
    cand = _strip_national_prefix(nm)
    if not cand or cand == "全国":
        return None
    if cand in PREF_STEMS:                 # 「千葉」「全国_千葉」「和歌山県」等
        return PREF_STEMS[cand]
    for _i, _pn in enumerate(PREF):
        if _pn and _pn in cand:            # 「30和歌山県」等の装飾付きフル名
            return _i
    return None


def is_national_area(code, name):
    return code == "00000" or _norm(name) == "全国"


def _as_list(x):
    if x is None:
        return []
    return x if isinstance(x, list) else [x]


def _text(x):
    return x.get("$", "") if isinstance(x, dict) else ("" if x is None else str(x))


def _clean_value(s):
    if s is None:
        return None
    t = str(s).replace(",", "").strip()
    try:
        return float(t)
    except ValueError:
        return None


def search_tables(app_id, word, stats_code, limit=100):
    params = {"appId": app_id, "searchWord": word, "limit": limit}
    if stats_code:
        params["statsCode"] = stats_code
    r = requests.get(API_LIST, params=params, timeout=60)
    r.raise_for_status()
    root = r.json().get("GET_STATS_LIST", {})
    res = root.get("RESULT", {})
    if res.get("STATUS") != 0:
        raise RuntimeError("e-Stat error {}: {}".format(res.get("STATUS"), res.get("ERROR_MSG")))
    return _as_list(root.get("DATALIST_INF", {}).get("TABLE_INF"))


def print_search(tables):
    if not tables:
        print("該当する統計表が見つかりませんでした。キーワードを変えて再検索してください。")
        return
    print("見つかった統計表（statsDataId / 名称・表題）:\n")
    for t in tables:
        sid = t.get("@id", "")
        line = "  {}  {}".format(sid, (_text(t.get("STATISTICS_NAME")) + " / " + _text(t.get("TITLE"))).strip(" /"))
        d = _text(t.get("SURVEY_DATE"))
        if d:
            line += "  [{}]".format(d)
        print(line)
    print("\n→ 上の statsDataId を --stats-id に指定（複数指定で年次を束ねられます）。")


def fetch_raw(app_id, stats_id, extra=None, limit=100000):
    params = {"appId": app_id, "statsDataId": stats_id, "metaGetFlg": "Y",
              "cntGetFlg": "N", "sectionHeaderFlg": "1", "lang": "J", "limit": limit}
    if extra:
        params.update(extra)
    values, class_obj, table_inf = [], None, {}
    start = 1
    while True:
        params["startPosition"] = start
        r = requests.get(API_DATA, params=params, timeout=60)
        r.raise_for_status()
        root = r.json().get("GET_STATS_DATA", {})
        res = root.get("RESULT", {})
        if res.get("STATUS") != 0:
            raise RuntimeError("e-Stat error {}: {}".format(res.get("STATUS"), res.get("ERROR_MSG")))
        sd = root["STATISTICAL_DATA"]
        if class_obj is None:
            class_obj = _as_list(sd.get("CLASS_INF", {}).get("CLASS_OBJ"))
            table_inf = sd.get("TABLE_INF") or {}
        values.extend(_as_list(sd.get("DATA_INF", {}).get("VALUE")))
        nxt = sd.get("RESULT_INF", {}).get("NEXT_KEY")
        if nxt:
            start = int(nxt)
        else:
            break
    return class_obj, values, table_inf


def survey_year(table_inf):
    """統計表メタ(SURVEY_DATE/タイトル)から西暦年(文字列)を推定。年次次元が無い表の補完用。"""
    if not table_inf:
        return None
    sd = str(table_inf.get("SURVEY_DATE", ""))
    m = re.search(r"(\d{4})", sd)
    if m:
        return m.group(1)
    title = _text(table_inf.get("TITLE")) + _text(table_inf.get("STATISTICS_NAME"))
    for era, base in [("令和", 2018), ("平成", 1988), ("昭和", 1925)]:
        mm = re.search(era + r"(元|\d+)年", title)
        if mm:
            n = 1 if mm.group(1) == "元" else int(mm.group(1))
            return str(base + n)
    return None


def parse(class_obj, values):
    name_maps = {o["@id"]: {c["@code"]: c["@name"] for c in _as_list(o.get("CLASS"))}
                 for o in class_obj}
    dim_names = {o["@id"]: o.get("@name", o["@id"]) for o in class_obj}

    def find(pred):
        return next((o["@id"] for o in class_obj if pred(o)), None)

    # 都道府県次元は「中身」で判定（表ごとに次元名が違うため最も確実）。
    # PP000形式(01000..47000)のコードが多く並ぶ次元を都道府県とみなす。
    def pref_score(o):
        sc = 0
        for c in _as_list(o.get("CLASS")):
            if match_pref(c.get("@code", ""), c.get("@name", "")) is not None:
                sc += 1
        return sc
    area_id, best = None, 0
    for o in class_obj:
        sc = pref_score(o)
        if sc > best:
            best, area_id = sc, o["@id"]
    if best < 5:  # 中身で見つからなければ id / 名称で
        area_id = find(lambda o: o["@id"] == "area") or \
            find(lambda o: any(k in (o.get("@name") or "") for k in ("都道府県", "都府県", "地域")))
    time_id = find(lambda o: o["@id"] == "time") or \
        find(lambda o: any(k in (o.get("@name") or "") for k in ("年", "時間")))

    records = []
    for v in values:
        rec = {"value": _clean_value(v.get("$")), "unit": v.get("@unit")}
        for o in class_obj:
            did = o["@id"]; code = v.get("@" + did)
            rec[did] = {"code": code, "name": name_maps.get(did, {}).get(code, code)}
        ac = v.get("@" + area_id) if area_id else None
        an = rec.get(area_id, {}).get("name") if area_id else None
        rec["pref_id"] = match_pref(ac, an)
        rec["area_code"] = ac
        rec["is_national"] = is_national_area(ac, an)
        rec["time_name"] = rec.get(time_id, {}).get("name") if time_id else None
        rec["time_code"] = rec.get(time_id, {}).get("code") if time_id else None
        records.append(rec)
    return {"area_id": area_id, "time_id": time_id, "dims": dim_names,
            "name_maps": name_maps, "records": records}


def resolve_picks(picks, name_maps, area_id, time_id):
    """日本語名を {dim_id: code} に解決（完全一致優先・なければ部分一致）。表ごとに呼ぶ。"""
    filters = {}
    for text in picks:
        exact = sub = None
        for did, codes in name_maps.items():
            if did in (area_id, time_id):
                continue
            for code, nm in codes.items():
                nm = nm or ""
                if nm == text and exact is None:
                    exact = (did, code)
                elif sub is None and text in nm:
                    sub = (did, code)
            if exact:
                break
        chosen = exact or sub
        if chosen:
            filters[chosen[0]] = chosen[1]
    return filters


def _passes(rec, filters):
    return all(rec.get(did, {}).get("code") == code for did, code in filters.items()) if filters else True


def main():
    ap = argparse.ArgumentParser(description="e-Stat 統計表 → みかんマップ用データ")
    ap.add_argument("--search")
    ap.add_argument("--stats-code", default=CROP_STATS_CODE)
    ap.add_argument("--stats-id", action="append")
    ap.add_argument("--app-id", default=os.environ.get("ESTAT_APP_ID"))
    ap.add_argument("--cd-cat01"); ap.add_argument("--cd-cat02")
    ap.add_argument("--cd-area"); ap.add_argument("--cd-time")
    ap.add_argument("--filter", action="append", default=[])
    ap.add_argument("--pick", action="append", default=[])
    ap.add_argument("--name", default="取得データ")
    ap.add_argument("--key", default=None)
    ap.add_argument("--measure-label", default="収穫量")
    ap.add_argument("--suffix", default="")  # e-Statの年ラベルは既に「2016年」等を含むため既定は空
    ap.add_argument("--note", default="")
    ap.add_argument("--out", default="estat_pref_map.json")
    ap.add_argument("--emit-js", default=None)
    ap.add_argument("--append", action="store_true")
    a = ap.parse_args()

    if not a.app_id:
        sys.exit("appId を --app-id か 環境変数 ESTAT_APP_ID で指定してください。")
    if a.search:
        print_search(search_tables(a.app_id, a.search, a.stats_code))
        return
    if not a.stats_id:
        sys.exit("--stats-id を1つ以上指定してください（--search で探せます）。")

    extra = {}
    for attr, pname in [("cd_cat01", "cdCat01"), ("cd_cat02", "cdCat02"),
                        ("cd_area", "cdArea"), ("cd_time", "cdTime")]:
        if getattr(a, attr):
            extra[pname] = getattr(a, attr)

    explicit = {}
    for item in a.filter:
        if "=" in item:
            did, code = item.split("=", 1)
            explicit[did.strip()] = code.strip()
    picks = a.pick if a.pick else ([] if a.filter else ["収穫量"])

    combined_series, combined_national, order, unit = {}, {}, {}, ""
    all_records, area_id, time_id = [], None, None
    warn = {}   # dim_name -> set(values)
    zero_tables = []
    tables_meta = []      # 使用した統計表の正式名称など（出典表示用）
    providers = []        # (提供機関, 調査名) の重複なしリスト
    used_filters = {}

    for sid in a.stats_id:
        class_obj, values, table_inf = fetch_raw(a.app_id, sid, extra)
        p = parse(class_obj, values)
        fy = survey_year(table_inf)  # 年次次元が無い表の年を補完
        ti = table_inf or {}
        ttl = (_text(ti.get("STATISTICS_NAME")) + " / " + _text(ti.get("TITLE"))).strip(" /")
        tables_meta.append({"id": sid, "title": ttl, "period": _text(ti.get("SURVEY_DATE"))})
        prov = (_text(ti.get("GOV_ORG")), _text(ti.get("STAT_NAME")))
        if any(prov) and prov not in providers:
            providers.append(prov)
        all_records.extend(p["records"])
        area_id = area_id or p["area_id"]
        time_id = time_id or p["time_id"]
        npref = sum(1 for r in p["records"] if r["pref_id"] is not None)
        if npref == 0:
            sample = []
            if p["area_id"]:
                sample = ["{}={}".format(c, n) for c, n in list(p["name_maps"].get(p["area_id"], {}).items())[:4]]
            zero_tables.append((sid, list(p["dims"].values()), sample))

        # この表に対する絞り込みを解決
        filt = {did: code for did, code in explicit.items() if did in p["name_maps"]}
        filt.update(resolve_picks(picks, p["name_maps"], p["area_id"], p["time_id"]))
        used_filters.update({p["dims"].get(k, k): p["name_maps"].get(k, {}).get(v, v) for k, v in filt.items()})

        # 未指定で値が複数残る次元 → 警告候補
        present = {}
        for r in p["records"]:
            for did in p["name_maps"]:
                if did in (p["area_id"], p["time_id"]) or did in filt:
                    continue
                c = r.get(did, {}).get("code")
                if c is not None:
                    present.setdefault(did, set()).add(c)
        for did, codes in present.items():
            if len(codes) > 1:
                names = {p["name_maps"][did].get(c, c) for c in codes}
                warn.setdefault(p["dims"].get(did, did), set()).update(names)

        # 集計
        for r in p["records"]:
            if not _passes(r, filt):
                continue
            t = r["time_name"] or ((fy + "年") if fy else "—")
            code = r["time_code"] or ((fy + "000000") if fy else None)
            if code:
                order[t] = code
            if not unit and r.get("unit"):
                unit = r["unit"]
            if r["pref_id"]:
                combined_series.setdefault(t, {})[str(r["pref_id"])] = r["value"]
            elif r.get("is_national"):
                combined_national[t] = r["value"]

    years = sorted(combined_series.keys(), key=lambda t: order.get(t, t))
    dataset = {"key": a.key or a.name, "name": a.name, "suffix": a.suffix,
               "measures": ["harvest"], "measure_labels": {"harvest": a.measure_label},
               "units": {"harvest": unit}, "years": years,
               "data": {"harvest": combined_series}, "national": {"harvest": combined_national},
               "note": a.note}

    rec_path = a.out.rsplit(".", 1)[0] + ".records.json"
    with open(rec_path, "w", encoding="utf-8") as f:
        json.dump({"records": all_records}, f, ensure_ascii=False, indent=1)
    if providers:
        source = "／".join("{}「{}」".format(g, st).replace("「」", "") for g, st in providers) + "（政府統計の総合窓口 e-Stat）"
    else:
        source = "政府統計の総合窓口 e-Stat"
    meta = {"stats_ids": a.stats_id, "filters": used_filters, "tables": tables_meta,
            "source": source, "generated": datetime.date.today().isoformat()}
    with open(a.out, "w", encoding="utf-8") as f:
        json.dump({"meta": meta, "series": combined_series}, f, ensure_ascii=False, indent=1)

    print("保存: {}（年次 {} / 明細 {} 行 → {}）".format(a.out, len(years), len(all_records), rec_path))
    if used_filters:
        print("絞り込み:", ", ".join("{}={}".format(k, v) for k, v in used_filters.items()))
    if not area_id:
        print("⚠ 都道府県の区分が見つかりません。都道府県別の表を指定してください。")
    for sid, dnames, sample in zero_tables:
        print("⚠ 表 {} は都道府県を検出できませんでした。次元: {}".format(sid, " / ".join(dnames)))
        if sample:
            print("    地域候補の中身例: " + " | ".join(sample))
    if warn:
        print("⚠ 複数の区分が混在しています。下記を --pick で1つに絞ってください（数値が混ざる原因）:")
        for name, vals in warn.items():
            sample = "／".join(list(vals)[:8])
            print("   - {}: {}".format(name, sample))
    if years:
        print("年次:", years[0], "〜", years[-1], "（{}区分）".format(len(years)))
        latest = years[-1]
        top = sorted(((p, v) for p, v in combined_series[latest].items() if v is not None), key=lambda x: -x[1])[:5]
        print("[{}] 上位5:".format(latest), ", ".join("{}={:,.0f}".format(PREF[int(p)], v) for p, v in top))

    if a.emit_js:
        datasets = []
        if a.append and os.path.exists(a.emit_js):
            m = re.search(r"window\.MIKAN_DATA\s*=\s*(\[.*?\]);", open(a.emit_js, encoding="utf-8").read(), re.S)
            if m:
                try:
                    datasets = json.loads(m.group(1))
                except json.JSONDecodeError:
                    datasets = []
        datasets = [d for d in datasets if d.get("key") != dataset["key"]]
        datasets.append(dataset)
        with open(a.emit_js, "w", encoding="utf-8") as f:
            f.write("// generated by fetch_estat.py ({})\n".format(meta["generated"]))
            f.write("window.MIKAN_META = " + json.dumps(meta, ensure_ascii=False) + ";\n")
            f.write("window.MIKAN_DATA = " + json.dumps(datasets, ensure_ascii=False) + ";\n")
        print("HTML用データ: {} に書き出し（データセット {} 件）".format(a.emit_js, len(datasets)))


if __name__ == "__main__":
    main()
