# 案内図・位置図 自動作成ツール

設計spec JSON + 数量表xlsm → 案内図.pdf 相当のレイアウトを自動生成する。
`案内図.pdf`（KTF00039案件）を実測してレイアウトを再現している。

2つの実装を用意している。**基本はHTML版を使う**（インストール不要・ダブルクリックで即使える）。

| | HTML版（推奨） | Python版 |
|---|---|---|
| ファイル | `annai_ichizu_tool.html` | `annai_ichizu_generator.py` |
| 事前準備 | 不要（Chromeで開くだけ） | `pip install -r requirements.txt` |
| 実行方法 | ダブルクリック→ブラウザで操作 | コマンドラインで実行 |
| PDF化 | ブラウザの印刷機能で保存 | 自動でPDFファイル出力 |
| 向いている用途 | 都度ちょっと作る・内容を目で見ながら調整したい | まとめて何件も自動処理したい |

---

## HTML版（`annai_ichizu_tool.html`）

### 使い方

1. `annai_ichizu_tool.html` をダブルクリックしてChrome（またはEdge）で開く
2. 「設計spec (*.json)」欄で `B_design_spec_*.json` を選択
   → 物件名・所在地・経度緯度・標高が自動入力される
3. 「数量表 (*.xlsm)」欄で数量表ファイルを選択（任意）
   → 入力シートのB2/B3/B4から基地局番号・基地局名・工事名が自動入力される。
   　未入力（`**`）の場合や数量表が無い場合は、下の欄に直接手入力すればOK
4. 内容を確認・必要なら修正（縮尺・日付・図面番号なども変更可）
5. 「地図を生成」ボタンを押す（インターネット接続が必要。国土地理院タイルを
   ブラウザが直接取得する。数秒〜十数秒かかる）
6. 「印刷 / PDFとして保存」ボタン→印刷ダイアログで以下を確認してPDF保存:
   - 用紙サイズ = **A4**
   - 印刷の向き = **横（ランドスケープ）**
   - 余白 = **なし（None）**
   - 拡大縮小 = **100%（実際のサイズ）** ※「用紙に合わせる」はOFFにする
   - 詳細設定 → **「背景のグラフィック」に必ずチェック**（これが無いと地図・罫線が印刷されない）
   - 送信先を「PDFに保存」にすればそのままPDFファイルになる

すべてブラウザ内（JavaScript）で完結する。ファイルの中身がどこかにアップロードされる
ことはない（地図タイル画像の取得先＝国土地理院サーバーへの通信のみ発生する）。

### うまく地図が表示されないとき

「地図を生成」後にステータス欄にエラーが出る、または地図が灰色のままの場合は、
社内プロキシ/ファイアウォールが `cyberjapandata.gsi.go.jp` への通信をブロックして
いないか確認してください。同様に数量表(xlsm)読み込み機能は `cdnjs.cloudflare.com`
から解析ライブラリ(SheetJS)を読み込むため、そこもブロックされていないか確認。

---

## Python版（`annai_ichizu_generator.py`）

まとめて何件も自動処理したい場合や、コマンドラインでのバッチ処理に組み込みたい場合はこちら。

### セットアップ（初回のみ）

```
pip install -r requirements.txt
```

Pythonが無ければ https://www.python.org/ からインストール（インストール時に
「Add python.exe to PATH」にチェック）。

### 使い方

```
python annai_ichizu_generator.py "<案件フォルダのパス>"
```

案件フォルダ内に **設計spec(*.json) と 数量表(*.xlsm) が1つずつ** 入っている前提で
自動検出する。実行すると同じフォルダに `案内図.pdf` が出力される。

例:
```
python annai_ichizu_generator.py "C:\...\49_案内図作成\raw\2034_..._R004749901"
```

複数ファイルがある/ファイル名が特殊な場合は明示指定:
```
python annai_ichizu_generator.py . --json spec.json --xlsm suryohyo.xlsm --out 案内図.pdf
```

xlsmが未入力（`**`のまま）の場合や、xlsmと違う値を使いたい場合は上書き指定できる:
```
python annai_ichizu_generator.py . --station-no KTF00039 --station-name "ＡＣＢ大阪大阪市北区第１１" --work-name "撤去プロジェクト_廃局(WCP)"
```

レイアウトだけ先に確認したいとき（地図は灰色プレースホルダーになる）:
```
python annai_ichizu_generator.py . --offline
```

その他オプション: `--scale-ichizu`（位置図縮尺分母, 既定2500）、
`--scale-annai`（案内図縮尺分母, 既定25000）、`--dpi`（既定200）、
`--sheet-no`（図面番号, 既定01）。

---

## データの取得元（両バージョン共通）

| 出力項目 | 取得元 |
|---|---|
| 物件名 | json: `responseBody.shared.site.billdingName.propertyNm` |
| 所在地（都道府県+市区町村） | json: `site.address.prefCd` + `cityCd` → 市区町村コード表で名称に変換 |
| 所在地（丁目番地） | json: `site.address.detailAddressNm` |
| 東経・北緯 | json: `site.degree.longitudeDeg` / `latitudeDeg`（10進度→度分秒に自動変換） |
| 標高 | json: `site.billding.groundLevel` |
| 基地局番号 | xlsm「入力シート」**B2** |
| 基地局名 | xlsm「入力シート」**B3** |
| 工事名 | xlsm「入力シート」**B4** |

「概ね」縮尺なので、地図データの実解像度（国土地理院タイルは最大ズーム18）の
都合でぴったり指定通りにはならない場合がある（自動的に一番近いズームを選ぶ）。

## 既知の制限・注意点

- **市区町村コード表**（HTML版は `.html` 内に埋め込み、Python版は
  `jis_city_codes.csv`）は総務省の公開データ（2026年時点）をもとに作成した
  全国1,894市区町村分の一覧。市町村合併等があった場合は更新が必要。対応する
  コードが無い場合は所在地欄に `(pref:xx)` のように出るので、その場合は手動で
  所在地を直してください。
- 地図は国土地理院 地理院タイル（標準地図, https://maps.gsi.go.jp/development/ichiran.html）
  を使用。社内プロキシ/ファイアウォールで `cyberjapandata.gsi.go.jp` への通信が
  ブロックされていると地図が表示されない（エラーメッセージで分かるようにしてある）。
- レイアウト座標は実測値なので、案件によって表題欄の文字がわずかにはみ出す可能性
  がある。HTML版は `annai_ichizu_tool.html` 内の `LAYOUT` 定数、Python版は
  `annai_ichizu_generator.py` 内の `LAYOUT` 辞書（どちらもpt単位・構造は同一）を
  編集すれば微調整できる。

## ファイル構成

```
annai_ichizu_tool.html      HTML版本体（推奨・単体で動作）
annai_ichizu_generator.py   Python版本体
jis_city_codes.csv          市区町村コード→名称 変換表（Python版が使用）
requirements.txt            Python版に必要なパッケージ一覧
README.md                   このファイル
```
