#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
import_tokusan.py
特産果樹生産動態等調査「2 種類別栽培状況（都道府県）かんきつ類」の公式Excelを
みかんマップ用データ(mikan_data.js)に変換して追記する。
※ スクレイピングはしない。e-Statから手動ダウンロードした公式Excelを読むだけ。

使い方:
  # 収録されている柑橘の種類を一覧表示
  python import_tokusan.py --kinds p001-01-021.xlsx --list-kinds

  # かんきつ類「計」（都道府県別）を追加
  python import_tokusan.py --total p001-01-022.xlsx --emit-js mikan_data.js --append

  # 種類を選んで追加（品目名の一部でOK。シラヌヒ=不知火/デコポン）
  python import_tokusan.py --kinds p001-01-021.xlsx --pick レモン --pick ユズ --pick シラヌヒ ^
         --emit-js mikan_data.js --append

  # 収録種類をすべて追加
  python import_tokusan.py --kinds p001-01-021.xlsx --all-kinds --emit-js mikan_data.js --append

依存: pip install openpyxl
"""
import os
import re
import sys
import json
import argparse
import datetime

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl が必要です:  pip install openpyxl")

# 都道府県判定は fetch_estat.py の関数を再利用（同じフォルダに置いてください）
try:
    from fetch_estat import match_pref, PREF
except Exception:
    PREF = ["", "北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県",
            "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県",
            "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県", "岐阜県",
            "静岡県", "愛知県", "三重県", "滋賀県", "京都府", "大阪府", "兵庫県",
            "奈良県", "和歌山県", "鳥取県", "島根県", "岡山県", "広島県", "山口県",
            "徳島県", "香川県", "愛媛県", "高知県", "福岡県", "佐賀県", "長崎県",
            "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県"]
    _STEM = {}
    for _i, _n in enumerate(PREF):
        if _n:
            _STEM[_n] = _i
            _STEM.setdefault(re.sub(r"[都府県]$", "", _n), _i)

    def match_pref(code, name):
        nm = (name or "").replace(" ", "").replace("\u3000", "")
        if not nm or nm == "全国":
            return None
        if nm in _STEM:
            return _STEM[nm]
        for i, pn in enumerate(PREF):
            if pn and pn in nm:
                return i
        return None

SOURCE = "農林水産省「特産果樹生産動態等調査」（政府統計の総合窓口 e-Stat）"
NATIONAL_LABELS = ("計", "合計", "全国")
# 指標: (キー, 表示名, 列オフセット)。Excelの 栽培面積/収穫量/出荷量 の並びに対応
MEAS = [("harvest", "収穫量"), ("ship", "出荷量"), ("area", "栽培面積")]


def clean_num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "")
    if s == "" or "・" in s or s in ("-", "－", "x", "X", "－"):
        return None  # 「・・・」=秘匿 等は欠測
    try:
        return float(s)
    except ValueError:
        return None


def detect_year(ws):
    for r in range(1, 6):
        for c in range(1, 4):
            t = str(ws.cell(r, c).value or "")
            m = re.search(r"令和(元|\d+)年", t)
            if m:
                n = 1 if m.group(1) == "元" else int(m.group(1))
                return str(2018 + n)
            m = re.search(r"平成(元|\d+)年", t)
            if m:
                n = 1 if m.group(1) == "元" else int(m.group(1))
                return str(1988 + n)
            m = re.search(r"(\d{4})年産", t)
            if m:
                return m.group(1)
    return None


def title_text(ws):
    parts = []
    for r in range(1, 5):
        for c in range(1, 4):
            t = ws.cell(r, c).value
            if t and str(t).strip():
                parts.append(str(t).strip().replace("\u3000", " "))
    return " / ".join(parts)


def make_dataset(key, name, year, by_pref, nat, note, table_title, table_id):
    """by_pref[meas_key] = {pref_id_str: val}, nat[meas_key] = val"""
    measures = [k for k, _ in MEAS if by_pref.get(k)]
    return {
        "key": key, "name": name, "suffix": "",
        "measures": measures or ["harvest"],
        "measure_labels": {k: lab for k, lab in MEAS},
        "units": {"harvest": "t", "ship": "t", "area": "ha"},
        "years": [year],
        "data": {k: {year: by_pref.get(k, {})} for k, _ in MEAS},
        "national": {k: {year: nat.get(k)} for k, _ in MEAS},
        "note": note,
        "source": SOURCE,
        "tables": [{"id": table_id, "title": table_title, "period": year}],
        "generated": datetime.date.today().isoformat(),
    }


def read_total(path):
    ws = openpyxl.load_workbook(path, data_only=True).active
    year = (detect_year(ws) or "—") + "年" if detect_year(ws) else "—"
    by = {k: {} for k, _ in MEAS}
    nat = {}
    for r in range(8, ws.max_row + 1):
        nm = ws.cell(r, 2).value
        if nm is None:
            continue
        nm = str(nm).strip()
        vals = {"area": clean_num(ws.cell(r, 3).value),
                "harvest": clean_num(ws.cell(r, 4).value),
                "ship": clean_num(ws.cell(r, 5).value)}
        if nm in NATIONAL_LABELS:
            nat = vals
            continue
        pid = match_pref(None, nm)
        if pid:
            for k in by:
                if vals[k] is not None:
                    by[k][str(pid)] = vals[k]
    note = "かんきつ類の合計（種類を問わない都道府県別の計）。栽培面積=ha、収穫量・出荷量=t。「・・・」は秘匿。"
    return make_dataset("kanki_total", "かんきつ類（合計）", year, by, nat, note,
                        title_text(ws), os.path.basename(path))


def read_kinds(path):
    """品目名ごとに (name, by_pref, nat) を返す。"""
    ws = openpyxl.load_workbook(path, data_only=True).active
    year = (detect_year(ws) or "—")
    year = (year + "年") if year != "—" else "—"
    blocks = []
    name_parts, by, nat = [], {k: {} for k, _ in MEAS}, {}
    for r in range(8, ws.max_row + 1):
        b = ws.cell(r, 2).value
        c = ws.cell(r, 3).value
        if b is not None and str(b).strip():
            name_parts.append(str(b).strip())
        if c is None or not str(c).strip():
            continue
        c = str(c).strip()
        vals = {"area": clean_num(ws.cell(r, 4).value),
                "harvest": clean_num(ws.cell(r, 5).value),
                "ship": clean_num(ws.cell(r, 6).value)}
        if c in NATIONAL_LABELS:
            nat = vals
            blocks.append((" ".join(name_parts).strip(), by, nat))
            name_parts, by, nat = [], {k: {} for k, _ in MEAS}, {}
            continue
        pid = match_pref(None, c)
        if pid:
            for k in by:
                if vals[k] is not None:
                    by[k][str(pid)] = vals[k]
    return year, blocks, title_text(ws), os.path.basename(path)


def _kana(s):
    # ひらがな→カタカナに寄せて、かなの表記差を吸収
    return "".join(chr(ord(c) + 0x60) if "\u3041" <= c <= "\u3096" else c for c in (s or ""))


def _foldk(s):
    return _kana((s or "").replace(" ", "").replace("\u3000", ""))


def kind_tokens(name):
    """種類名から照合用トークン集合を作る: 先頭トークン(かっこ前)・かっこ内の各語・全体。"""
    toks = {name.strip()}
    lead = re.split(r"[（(]", name, 1)[0].strip()
    if lead:
        toks.add(lead)
    for m in re.findall(r"[（(]([^）)]*)[）)]", name):
        m = m.strip()
        if m:
            toks.add(m)
    return {_foldk(t) for t in toks}


def match_kind(pick, name):
    return _foldk(pick) in kind_tokens(name)


def slugify(name):
    return "kanki_" + re.sub(r"[^0-9A-Za-zぁ-んァ-ヶ一-龠]+", "", name)[:24]


def load_existing(path):
    datasets, meta_block = [], None
    if os.path.exists(path):
        txt = open(path, encoding="utf-8").read()
        m = re.search(r"window\.MIKAN_DATA\s*=\s*(\[.*?\]);", txt, re.S)
        if m:
            try:
                datasets = json.loads(m.group(1))
            except json.JSONDecodeError:
                datasets = []
        mm = re.search(r"window\.MIKAN_META\s*=\s*(\{.*?\});", txt, re.S)
        if mm:
            meta_block = mm.group(1)
    return datasets, meta_block


def write_js(path, datasets, meta_block):
    if not meta_block:
        meta_block = json.dumps({"source": SOURCE,
                                 "generated": datetime.date.today().isoformat()},
                                ensure_ascii=False)
    with open(path, "w", encoding="utf-8") as f:
        f.write("// generated/updated by import_tokusan.py ({})\n".format(datetime.date.today().isoformat()))
        f.write("window.MIKAN_META = " + meta_block + ";\n")
        f.write("window.MIKAN_DATA = " + json.dumps(datasets, ensure_ascii=False) + ";\n")


def main():
    ap = argparse.ArgumentParser(description="特産果樹（かんきつ類）Excel → mikan_data.js")
    ap.add_argument("--total", help="022（かんきつ類計・都道府県別）のxlsx")
    ap.add_argument("--kinds", help="021（種類別×都道府県別）のxlsx")
    ap.add_argument("--list-kinds", action="store_true", help="収録されている種類名を一覧表示")
    ap.add_argument("--pick", action="append", default=[], help="追加する種類（品目名の一部一致。複数可）")
    ap.add_argument("--all-kinds", action="store_true", help="収録種類をすべて追加")
    ap.add_argument("--emit-js", default=None)
    ap.add_argument("--append", action="store_true")
    a = ap.parse_args()

    new_datasets = []

    if a.kinds:
        year, blocks, ttitle, tid = read_kinds(a.kinds)
        if a.list_kinds:
            print("収録されている種類（{}・{}件）:".format(year, len(blocks)))
            for nm, _by, _nat in blocks:
                print("  -", nm)
            print("\n→ --pick に上の名前の一部を指定（例 --pick レモン）。シラヌヒ=不知火/デコポン。")
            return
        want = a.pick
        for nm, by, nat in blocks:
            if a.all_kinds or any(match_kind(p, nm) for p in want):
                key = slugify(nm)
                note = "{}（種類別・都道府県別）。栽培面積=ha、収穫量・出荷量=t。".format(nm)
                new_datasets.append(make_dataset(key, nm, year, by, nat, note, ttitle, tid))

    if a.total:
        new_datasets.append(read_total(a.total))

    if not new_datasets:
        sys.exit("追加対象がありません。--total か（--kinds と --pick/--all-kinds）を指定してください。")

    # 概要表示
    for d in new_datasets:
        y = d["years"][0]
        h = d["data"].get("harvest", {}).get(y, {})
        top = sorted(((p, v) for p, v in h.items() if v is not None), key=lambda x: -x[1])[:5]
        print("追加: {}（{}）収穫量上位5: {}".format(
            d["name"], y, ", ".join("{}={:,.0f}".format(PREF[int(p)], v) for p, v in top) or "（収穫量なし）"))

    if a.emit_js:
        datasets, meta_block = (load_existing(a.emit_js) if a.append else ([], None))
        keys = {d["key"] for d in new_datasets}
        datasets = [d for d in datasets if d.get("key") not in keys] + new_datasets
        write_js(a.emit_js, datasets, meta_block)
        print("書き出し: {}（データセット計 {} 件）".format(a.emit_js, len(datasets)))
    else:
        print("（--emit-js を指定すると mikan_data.js に書き出します）")


if __name__ == "__main__":
    main()
