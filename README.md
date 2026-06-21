# みかんアトラス — 日本の柑橘を、地図と統計で。

（e-Stat 政府統計ベース）

温州みかんと柑橘類の統計を、**都道府県別の塗り分け地図**で見られるサイトです。生産・市場・消費・担い手の各指標を、すべて政府統計の総合窓口 **e-Stat の公式データ**（API またはダウンロードした統計表 Excel）からのみ作成しています（Webスクレイピングは行っていません）。

公開ページ（GitHub Pages）: `https://mikan2010.github.io/mikan-map/`


---

## ページ構成

| ページ | ファイル | 内容 | 色 |
|---|---|---|---|
| 産地マップ | `index.html` | 温州みかん・柑橘の収穫量／出荷量／結果樹面積／単収を年次スライダーで | オレンジ |
| 市場 | `market.html` | 卸売（産地別・消費市場別）と小売価格・家計消費 | オレンジ／青 |
| 担い手（参考） | `labor.html` | 基幹的農業従事者数・平均年齢（**農業全体**の参考値） | 緑 |
| 分析 | `analyst.html` | 単収・全国シェア・集中度、規模×効率、担い手とのかけ合わせ、生産×消費 | — |
| 品種 | `explore.html` | 産地×品種マトリクス、品種カルテ、品種比較（単収など）・品種構成（中晩柑97品種） | オレンジ |
| 市町村マップ | `origins.html` | 主要産地の市町村をドット表示。選ぶとその市町村の品種が出る | オレンジ |

各ページの上部「**出典・データについて**」に、使用した統計表の正式名称・統計表ID（または Excel ファイル名）・対象期間が全件記載されます。表示している数値はデータ自身に出典が埋め込まれています。

---

## データ出典（すべて e-Stat）

| 区分 | 調査・統計表 | 提供 |
|---|---|---|
| 生産 | 農林水産省「作物統計調査（果樹）」 | API |
| 生産（種類別） | 農林水産省「特産果樹生産動態等調査」 | Excel |
| 卸売 | 農林水産省「青果物卸売市場調査（産地別）」 | API |
| 小売価格 | 総務省「小売物価統計調査（動向編）」 | API |
| 家計消費 | 総務省「家計調査（家計収支編・品目分類 第6表）」 | Excel |
| 担い手 | 農林水産省「農業構造動態調査」 | API |

地図のポリゴンは [dataofjapan/land](https://github.com/dataofjapan/land) の都道府県 GeoJSON を使用しています。

---

## データの読み方・注意点

- **県庁所在市ベースの指標**：「みかん小売価格（県庁所在市）」「みかん（家計消費・県庁所在市）」は、**県全体の平均ではなく県庁所在市の値**を各県に割り当てたものです（名称と注記に明記）。
- **担い手は参考値**：基幹的農業従事者数・平均年齢は**農業全体**の数字で、みかん専従ではありません。生産・市場と色味（緑）を変え、ページ冒頭に注記しています。
- **卸売は主要消費市場が対象**：青果物卸売市場調査は主要消費市場での取扱分の集計のため、各産地の全出荷ではありません。「消費市場別」は市場が所在する都道府県のみ色がつきます。
- **数量加重平均**：卸売価格は「数量加重平均」（取扱量の多い市場ほど強く反映＝実勢に近い平均）で算出しています。
- **加工値**：分析ページの単収（収穫量÷面積）、全国シェア、集中度（HHI）、1人あたり生産量などは、このサイトが上記出典から計算した加工値です。計算に使った表は各図の脚注に併記しています。

---

## データの作り方（再現手順）

### 必要環境

- Python 3.9 以上
- 依存ライブラリ：`pip install pandas openpyxl xlrd`
- e-Stat の **アプリケーションID**（[利用登録](https://www.e-stat.go.jp/api/) で取得）を環境変数に設定：

```powershell
# Windows PowerShell の例
$env:ESTAT_APP_ID = "あなたのアプリID"
```

### スクリプト

| スクリプト | 役割 |
|---|---|
| `fetch_estat.py` | e-Stat API（REST v3.0）から統計表を取得し、都道府県名・市名を解決して `mikan_data.js` に追記。`--inspect` で表構造を確認できる |
| `import_tokusan.py` | 特産果樹（種類別×都道府県）の Excel を取り込み。他スクリプトが使う共通の書き出し関数も提供 |
| `import_wholesale.py` | 青果物卸売市場調査（産地別）を取り込み。`--by origin`（産地別／既定）・`--by market`（消費市場別） |
| `import_kakei.py` | 家計調査 第6表の Excel から、みかんの購入数量・支出を県庁所在市→都道府県で取り込み |

すべて `--emit-js mikan_data.js --append` で同じデータファイルに追記します。各データセットの**正確な統計表ID・対象期間は `mikan_data.js` 内（および各ページの出典欄）に記録**されています。代表的なコマンド例：

```powershell
# 卸売（産地別）と消費市場別
python import_wholesale.py --stats-id 0002114420 --emit-js mikan_data.js --append
python import_wholesale.py --stats-id 0002114420 --by market --emit-js mikan_data.js --append

# 小売価格（県庁所在市・年平均、みかん=銘柄01511）
python fetch_estat.py --stats-id 0003420453 --cd-cat02 01511 --measure price=価格 `
  --name "みかん小売価格（県庁所在市）" --key "みかん小売価格" `
  --note "【値は都道府県庁所在市のもの（県全体の平均ではない）】都道府県庁所在市の店頭価格を各県に割り当てた値です。" `
  --emit-js mikan_data.js --append

# 家計消費（県庁所在市）※ fn0605.xls は家計調査 第6表をダウンロードして配置
python import_kakei.py fn0605.xls --emit-js mikan_data.js --append

# 担い手（基幹的農業従事者数・平均年齢）
python fetch_estat.py --stats-id 0002063487 --cd-cat02 1139 --measure people=基幹的農業従事者数 `
  --name "基幹的農業従事者数" --emit-js mikan_data.js --append
```

表の構造を調べたいときは `--inspect`（`getMetaInfo` を使うので大きな表でも即時）：

```powershell
python fetch_estat.py --inspect 0003420453 --grep みかん
```

---

## ファイル構成

```
.
├── index.html            # 産地マップ（都道府県）
├── market.html           # 市場（卸売・小売・家計）
├── labor.html            # 担い手（参考）
├── analyst.html          # 分析
├── explore.html          # 品種（産地×品種マトリクス・比較・構成）
├── origins.html          # 市町村マップ（産地ドット）
├── mikan_data.js         # 生産・市場・消費・担い手データ
├── tokusan_data.js       # 中晩柑97品種×都道府県（explore.html 用）
├── origins_data.js       # 市町村ドット＋品種逆引き（origins.html 用）
├── fetch_estat.py        # e-Stat API 取得
├── import_tokusan.py     # 特産果樹 Excel 取込（共通関数も提供）
├── import_wholesale.py   # 卸売市場調査 取込（産地別／消費市場別）
├── import_kakei.py       # 家計調査 Excel 取込
├── build_explore.py      # 特産果樹 → tokusan_data.js
├── build_origins.py      # 主要産地の市町村をジオコーディング → origins_data.js
├── .gitignore
└── README.md
```

HTML 4ページは外部依存を最小化するため、Leaflet と都道府県 GeoJSON をインライン化しています（フォントのみ Google Fonts を `<link>` で読み込み）。`mikan_data.js` は各 HTML と**同じ階層**に置いてください。

`.gitignore` は中間ファイル（`estat_pref_map*.json`, `*.records.json`）や秘密情報を除外します。**アプリIDは環境変数で扱い、コミットしないでください。**

---

## 公開方法（GitHub Pages）

1. 上記ファイルをリポジトリ直下に置いて push。
2. リポジトリの **Settings → Pages → Build and deployment** で、Source「Deploy from a branch」、Branch「main」「/(root)」を選び Save。
3. 1分ほどで公開。表示が古いときはハードリフレッシュ、または URL 末尾に `?v=2` を付与（キャッシュ約10分）。

---

## ライセンス・クレジット

- **統計データ**：政府統計の総合窓口（e-Stat）。各府省の統計を [政府標準利用規約（第2.0版）](https://www.e-stat.go.jp/terms-of-use) に従い、出典を明記して利用・加工しています。
- **地図データ**：[dataofjapan/land](https://github.com/dataofjapan/land)（都道府県 GeoJSON）。非商用利用・要クレジット。
- **市区町村の座標**（市町村マップのドット）：[UG/JapanCityTownHall_lat_long](https://github.com/UG/JapanCityTownHall_lat_long)（市区町村役場の緯度経度、MIT License）。
- **ライブラリ等**：[Leaflet](https://leafletjs.com/)、[Google Fonts](https://fonts.google.com/)（Zen Kaku Gothic New / Shippori Mincho B1）。
- 本リポジトリのコード（スクリプト・HTML）のライセンスは、必要に応じて追記してください（例：MIT）。

※ 数値はいずれも出典の政府統計を加工して作成したものであり、加工の責任は本リポジトリの作成者にあります。正確な値は各出典の原典をご確認ください。
