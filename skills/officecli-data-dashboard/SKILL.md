---
name: officecli-data-dashboard
description: "CSV や表形式データから、複数要素で構成される Excel ダッシュボード — 開いた時に表示される Dashboard シート、数式駆動の複数 KPI カード、複数のチャート、スパークライン、条件付き書式 — を構築する際にこのスキルを使用する。トリガーワード: 'dashboard'、'KPI dashboard'、'analytics dashboard'、'executive dashboard'、'metrics dashboard'、'CSV to dashboard'、'data visualization'。出力は単一の .xlsx。officecli-xlsx 上のシーンレイヤー: xlsx のハードルールをすべて継承する。次の場合は呼び出さないこと: 単一の予算トラッカー／1シートの CSV に書式を付けただけのもの（xlsx を使用）、3ステートメント／DCF／LBO 財務モデル（financial-model を使用）、チャートが1つ以下・行数10未満の週次レポート（xlsx を使用）。"
---

# Data Dashboard（officecli-xlsx 上のシーンレイヤー）

ダッシュボードとは「チャート付きのスプレッドシート」ではない。それは一つの構成物 — **ユーザーが最初にたどり着く単一の Dashboard シート**に、数式駆動の KPI カード、セル範囲にリンクしたチャート、スパークライン、意味づけされた条件付き書式が揃っているものだ。それ以外（生データ、集計処理）はすべて、ユーザーが決して開く必要のない上流のインフラである。このスキルはその構成パターンを教える。xlsx エンジンに関するすべて — セル、数式、バッチ JSON、シェルのクォーティング、validate、HTML プレビュー — は `officecli-xlsx` に由来し、ここでは再説明しない。

## セットアップ

`officecli` が未導入の場合:

- **macOS / Linux**: `curl -fsSL https://d.officecli.ai/install.sh | bash`
- **Windows (PowerShell)**: `irm https://d.officecli.ai/install.ps1 | iex`

`officecli --version` で確認する（PATH がまだ反映されていない場合は新しいターミナルを開く）。インストールに失敗した場合は https://github.com/iOfficeAI/OfficeCLI/releases からバイナリをダウンロードする。

## ⚠️ Help-First ルール

**プロパティ名、enum 値、エイリアスが不確かな場合は、推測する前に help を確認すること。**

```bash
officecli help xlsx                          # element list
officecli help xlsx chart                    # full schema for charts
officecli help xlsx sparkline                # sparklines
officecli help xlsx conditionalformatting    # all CF rule types
```

help はインストール済みの CLI バージョンを反映する。本スキルと help の内容が食い違う場合は **help を優先する**。DeferredAddKeys（`combosplit`、`holesize`）は `add` 時のみ有効 — 詳細は Reference を参照。

## メンタルモデルと継承

このスキルは `officecli-xlsx` から**すべての xlsx ハードルールを継承する** — シェルのクォーティング、数式エラーゼロ、視覚的な納品下限、バッチ JSON の形（`{"command":"set"|"add","path":...,"props":{...}}` — キーは `command` であり、`action` ではない）、バッチ JSON のドット付き名前ルール、チャートのデータフィード形式、バッチ／レジデントの制限、`validate` の徹底。まず officecli-xlsx を読み、それらのルールに従うこと。ここでは再説明しない。

**逆ハンドオフ — 次の場合はこのスキルを使わないこと:**

- 依頼が **単一シートの CSV に書式を付けたトラッカー**（Dashboard シートなし、KPI カードなし、チャート1つ以下）である → `officecli-xlsx` に戻る。
- 依頼が青字入力／黒字数式／シート横断ドライバーを伴う **3ステートメント／DCF／LBO 財務モデル**である → `officecli-financial-model` を使用する。
- 依頼が SUMIF 集計1つとチャート1つ、行数10未満の**週次ステータスレポート**である → `officecli-xlsx`。

このスキルが受け付けるのは「ユーザーが最初に開く Dashboard シート、複数の KPI カード、複数のチャート、いくつかの CF／スパークライン」のみである。

## シェルと実行の規律

→ ベースライン（クォーティング、`!` のための heredoc、段階的な実行）は officecli-xlsx の §Shell & Execution Discipline を参照。

ダッシュボード特有の増分は2つ:

- **長いチャートの `add` コマンドは180文字を超える。** 常に末尾に `\` を付けて複数行に分割すること。チャートコマンドを1行に詰め込んではならない。コマンドが長くなるほど、シェルエスケープのバグが潜む可能性は高くなる。
- **複数インスタンスの件数取得には `query --json | jq length` を使い、`raw-get | grep -c` は使わない。** 例: 「チャートはいくつあるか？」に対しては `officecli query "$FILE" chart --json | jq '.data.results | length'`。

## Core Principles（中核原則）

5つの譲れない原則。どれか一つでも破られていれば、それはダッシュボードではなく「たまたまチャートが付いているスプレッドシート」である。

1. **数式駆動の KPI。** Dashboard シート上のすべての KPI 値は数式である — `SUM`、`AVERAGE`、`IFERROR((...-...)/...,0)` など、いずれも Data／Summary シート上のセルを参照する。計算済みの数値をハードコードしてはならない。翌日に元データが変わっても、KPI は開いた時点で更新される。

2. **チャートはセル範囲を参照する。** すべてのチャートシリーズはセル範囲から読み取る: `series1.values="Sheet1!B2:B13"`。インラインの `data="Revenue:100,200,300"` は5分デモ用であり、納品するダッシュボードには使わない。唯一の例外は、Excel が表現できない集計をデータが必要とする場合（稀）— その場合はコメントセルに例外を明記する。

3. **Dashboard ファーストのアーキテクチャ。** KPI ラベルセル、KPI 値セル、チャート、スパークラインはすべて、ユーザーが最初にたどり着く**単一の Dashboard シート**上に置く。生データのインポートや `SUMIFS` の集計は、Dashboard より上流の Data／Summary シートに置く。ユーザーが答えを探すためにタブを切り替える必要は決してあってはならない。

4. **チャートのデータ元は可視セルのみ。** LibreOffice はレンダリング時に、非表示の列や非表示シート内の数式を評価しない。`series1.values` が非表示列の `SUMIFS` を指しているチャートは空白でレンダリングされる。パターン: **可視の** Summary シートに集計し、チャートは Summary のセルを参照し、チャートのデータ元ではないヘルパー列だけを非表示にする。

5. **データ規模に応じた複雑度。** 10行のデータセットに KPI 5個・チャート4個は不釣り合いだ。200行のデータセットに KPI 1個・チャート1個も不釣り合いだ。入力データに合わせて構成をスケールさせる（§Design Ideas の表を参照）。過剰構築は過小構築と同じくらい誤りである。

## Requirements（要件）

`officecli-xlsx` の要件はすべて適用される（→ officecli-xlsx の §Requirements for Outputs を参照）。ダッシュボードでは以下を追加する:

- **Dashboard シートが開いた時のアクティブタブであること。** `activeTab="N"` を設定する**前に** `officecli query "$FILE" sheet` で 0 始まりのシートインデックスを確認すること。インデックスを推測してはならない。
- **`calc.fullCalcOnLoad=true` を設定する。** `officecli set "$FILE" / --prop calc.fullCalcOnLoad=true` で設定する。`<calcPr>` を `raw-set` してはならない — 要素が重複し validate に失敗する。
- **上流を編集するたびに下流の cachedValue を再計算すること。** `fullCalcOnLoad=true` は実行時の再計算のみをスケジュールし、ビルド時の `cachedValue` は更新しない。`set B=100 → set E==B+D → fix B=150` の後、E の数式を再発行する（あるいは close/reopen する）まで E は古いままだ。キャッシュが古いままだと「Net Change = 0」がボードに届いてしまう。
- **すべてのチャートに説明的なタイトルを、すべてのシリーズに名前を付けること。** 凡例に "Series1" が残っているのは未完成の仕事だ。
- **すべての KPI 値セルに数式があること。** 検証方法: `officecli query "$FILE" 'Dashboard!:has(formula)' --json | jq '.data.results | length'` が計画した KPI 件数と一致すること。
- **すべてのデータシートでヘッダー行を塗りつぶすこと。** Data シート、Summary シート、その他の副次データシートは、1行目を塗りつぶす必要がある（例: `fill=1F3864 + font.color=FFFFFF + font.bold=true`）。
- **10行以上の Data シートには CF ルールを1つ以上つけること。** 20行の表に視覚的なスキャン補助が一つもないのは品質上の欠陥である。
- **Dashboard の値列は、想定される cachedValue の最大幅に合わせてサイズを決める — 固定 22 ではない。** 目安（24pt bold ＋ 通貨 numFmt の場合）: `width ≈ ceil((visible_chars + 2) × 1.3)`。`¥1,958,414,250`（通貨記号とカンマ込みで可視文字数14）を保持する KPI には `width ≥ 28` が必要。4桁の KPI であっても最低 `width ≥ 22` は必要。10桁以上の KPI に `22` をハードコードすると `###` がユーザーに届く。
- **スパークラインの行高は20以上であること。** デフォルトの15pt行に置かれたスパークラインは平坦なギザギザにしか見えない — `/Dashboard/row[N] height=22` を設定する（同じ行に24ptの KPI 値セルが並ぶ場合は24）。
- **印刷用の納品物は `_xlnm.Print_Area` を Dashboard に限定し**、Dashboard 以外のシートを非表示にし、`<pageSetup fitToPage/>` を追加すること。この3つすべてがないと、印刷パイプラインは全シートを出力し、Dashboard は2ページ目以降に配置されてしまう。正確なコマンドは §Print-ready delivery を参照。

## クイックスタート

最小構成のダッシュボード: 12ヶ月分の売上 CSV → KPI 4個 ＋ 折れ線チャート1個 ＋ activeTab ＋ fullCalcOnLoad。数値は状況に合わせて調整し、そのままコピペしないこと。フェーズに分割してあるので、どのフェーズが失敗したかが分かりやすい。

**フェーズ1 — Data シート: 作成、インポート、書式設定。**

```bash
FILE=my_dashboard.xlsx
officecli create "$FILE"
officecli import "$FILE" /Sheet1 --file sales.csv --header
officecli set "$FILE" '/Sheet1/col[A]' --prop width=12
officecli set "$FILE" '/Sheet1/col[B]' --prop width=15
officecli set "$FILE" '/Sheet1/B2:B13' --prop numFmt='$#,##0'
officecli set "$FILE" '/Sheet1/A1:B1' --prop fill=1F3864 --prop font.color=FFFFFF --prop font.bold=true
```

**フェーズ2 — Dashboard シート ＋ KPI カード1個。**

```bash
officecli add "$FILE" / --type sheet --prop name=Dashboard
officecli set "$FILE" '/Dashboard/col[A]' --prop width=22
officecli set "$FILE" '/Dashboard/col[B]' --prop width=12
officecli set "$FILE" /Dashboard/A1 --prop value="Total Revenue" --prop font.size=9 --prop font.color=666666 --prop bold=true
officecli set "$FILE" /Dashboard/A2 --prop 'formula==SUM(Sheet1!B2:B13)' --prop numFmt='$#,##0' --prop font.size=24 --prop bold=true --prop font.color=2E7D32
```

**フェーズ3 — スパークライン ＋ チャート。**

```bash
officecli add "$FILE" /Dashboard --type sparkline --prop cell=B2 --prop range='Sheet1!B2:B13' --prop type=line --prop color=4472C4 --prop highPoint=true --prop highMarkerColor=FF0000
officecli add "$FILE" /Dashboard --type chart \
  --prop chartType=line \
  --prop title="Revenue Trend" \
  --prop series1.name="Revenue" \
  --prop series1.values='Sheet1!B2:B13' \
  --prop series1.categories='Sheet1!A2:A13' \
  --prop preset=dashboard --prop axisNumFmt='$#,##0' \
  --prop x=0 --prop y=5 --prop width=10 --prop height=15
```

**フェーズ4 — fullCalcOnLoad → activeTab（最後）→ close → validate。**

```bash
officecli set "$FILE" / --prop calc.fullCalcOnLoad=true

# Dashboard の 0 始まりインデックスを実際のシート一覧から解決する — 決してハードコードしない。
DASH_IDX=$(officecli query "$FILE" sheet --json \
  | jq '[.data.results[].path] | index("/Dashboard")')
officecli raw-set "$FILE" /workbook --xpath "//x:sheets" --action insertbefore \
  --xml "<bookViews xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><workbookView activeTab=\"$DASH_IDX\" /></bookViews>"
officecli close "$FILE"
officecli validate "$FILE"
```

12行の売上 CSV でエンドツーエンド検証済み: `validate` はエラーなしを報告し、Dashboard が最初に開き、`Dashboard/A2.cachedValue` は解決される（テストデータで 2,075,000）。チャートは値がリンクされた状態でレンダリングされる。

## デザインのアイデア

これらは選択肢であり、テンプレートではない。ユーザーのデータと想定読者が選択を決める。

### レイアウトパターン（1つ選び、一貫させる）

**パターン1 — エグゼクティブサマリー**（役員向け資料）: KPI 帯を A1:H4 に、チャートは6行目から積み上げる。
```
┌ KPI1 │ KPI2 │ KPI3 │ KPI4 ┐  rows 1-4
├──────┴──────┴──────┴──────┤
│     Chart 1 (wide)        │  rows 6-18
├───────────────┬───────────┤
│   Chart 2     │  Chart 3  │  rows 20-32
```

**パターン2 — オペレーションコンソール**（ライブ運用向け）: KPI を A:B に縦に並べ、チャートは C:L を埋める。
```
│ KPI1 │                   │
│ KPI2 │    Chart 1        │  rows 1-12
│ KPI3 │                   │
│ KPI4 ├───────────────────┤
│ KPI5 │    Chart 2        │  rows 14-26
```

**パターン3 — スコアカード**（KPI 6個以上、支配的なチャートなし）: 2×3のカードグリッド（ラベル／値／スパークライン）。
```
│ KPI1 │ KPI2 │ KPI3 │  rows 1-4
│ KPI4 │ KPI5 │ KPI6 │  rows 5-8
```

### データ規模による複雑度のスケーリング

| Rows | KPIs | Charts | Sparklines | CF rules | Preset |
|---|---|---|---|---|---|
| < 10 | 1–2 | 1 | skip | 0–1 | `minimal` |
| 10–50 | 2–3 | 2 | only if sequential time-series | 1–2 | `dashboard` |
| 50–200 | 3–5 | 2–3 | only if sequential time-series | 2–3 | `dashboard` |
| 200+ | 3–5 | 3 | only if sequential time-series | 3–4 | `dashboard` |

### チャート種別の選び方

| Data pattern | Chart type | Notes |
|---|---|---|
| Trend over time, one series | `line` | ノイズの多いシリーズで傾向を示すには `trendline=linear` を追加 |
| Trend over time, multiple components | `line` (multi-series) or `columnStacked` | 各構成要素の合計が意味を持つ場合は積み上げ |
| Comparison across categories in time order | `column` | `bar` ではない — 横棒グラフは左から右への時系列の読みを崩す |
| Part-of-whole breakdown | `doughnut` | `pie` より優先: `chartType=pie` には既知の LibreOffice 空白レンダリングの不具合がある |
| Budget vs actual | `combo` with `combosplit=1` | 最初のシリーズを棒グラフ、残りを折れ線に |
| Correlation | `scatter` | X軸は `categories` / `series1.categories` 経由 — `series1.xValues` は未対応 |

### プリセットのオプション

すべてのチャートに `--prop preset=<name>` を付ける。選択肢: `minimal`、`dashboard`、`corporate`、`magazine`、`colorful`、`monochrome`、`dark`。1つを選び、1枚の Dashboard 上のすべてのチャートで一貫させること — プリセットを混在させると意図しない仕上がりに見える。

### 条件付き書式 — 意味づけされた色

CF ルールは4種類。それぞれ `add` 時に `--type <shorthand>` を指定する:

| Intent | `--type` | Typical props |
|---|---|---|
| 大きさを示すバー（売上、支出） | `databar` | `sqref=B2:B13 color=4472C4` — 予測可能なスケーリングのために明示的な `min=0 max=<plausible>` を推奨するが、省略しても有効（データの最小値／最大値がデフォルトになる） |
| ヒートマップ（比率、成長率） | `colorscale` | `sqref=D2:D13 mincolor=FFCDD2 midcolor=FFFFFF maxcolor=C8E6C9` |
| ステータスインジケーター | `iconset` | `sqref=E2:E13 iconset=3Arrows` — 完全な enum は help を参照 |
| カスタムのビジネスルール | `formulacf` | `sqref=B2:B13 'formula=$B2>=100000' fill=C8E6C9 font.color=2E7D32` — `font.bold` も CF に効く |

ダッシュボード内で一貫させるべき意味づけされた色:

- 良好／プラス: fill `C8E6C9`、font `2E7D32`
- 不良／マイナス: fill `FFCDD2`、font `C62828`
- 中立: fill `F5F5F5`、font `666666`

### KPI カードの構造

カードはラベルセル＋値セルで構成される。ラベルは小さいグレー（font.size=9、font.color=666666、bold）、値は大きく太字（font.size=24、bold=true、numFmt、トーンを示す font.color）。カード領域全体に淡い塗りつぶし（例: `F0F4FF`）を1行入れると、結合セルの足場を組まずに「カード」らしい見た目になる。値列の幅は最大の cachedValue に合わせてサイズを決める必要がある — 22未満にはせず、8桁以上の通貨には26〜32が多い（Requirements 参照）。

### タイトル長によるチャート幅の目安

`dashboard` プリセットのデフォルトタイトルフォントでは、チャートのプロットボックス幅（列単位）はタイトル文字列より広く保たないと、タイトルが単語の途中で切れる。目安: `chart.width ≥ ceil(title.length × 0.18)`。35文字のタイトル（"Department: Year-End Headcount vs Attrition Rate"）には `width ≥ 7` が必要だが、安全のため10〜12を使うこと。アンカーを広げられない場合はタイトルを25文字以下に短縮する — 役員向け納品物でタイトルが切れているのは弁解の余地がない。

`officecli get chart[N]` は数値の `width`（例: `width=480pt`）と `anchor`（例: `"A6:K21"`）を `.data.results[0].format` に公開する。列単位の目安が必要な場合は、アンカーの文字（A→K = 10列）から列スパンを導出すればどちらも使える（Gate 2 用）。

### 印刷用の納品（board-pack／investor-send／one-pager）

トリガー: 依頼文に "print" / "一页" / "董事会" / "投资人" が含まれる場合。Dashboard シートに4つの成果物を設定し、Dashboard 以外のシートは非表示にして、印刷パイプラインが1ページのみを出力するようにする。

```bash
# 1. Print_Area scoped to Dashboard (xlnm convention).
officecli add "$FILE" / --type namedrange --prop name=_xlnm.Print_Area --prop scope=Dashboard --prop 'refersTo=Dashboard!$A$1:$H$36'
# 2. fit-to-page on Dashboard.
officecli raw-set "$FILE" /Dashboard --xpath "//x:worksheet" --action prepend --xml '<sheetPr xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><pageSetUpPr fitToPage="1"/></sheetPr>'
# 3. Landscape page setup.
officecli raw-set "$FILE" /Dashboard --xpath "//x:sheetData" --action insertafter --xml '<pageSetup orientation="landscape" paperSize="9" fitToWidth="1" fitToHeight="1"/>'
# 4. Hide non-Dashboard sheets — Print_Area scope alone does NOT stop the print pipeline from emitting every visible sheet.
for S in Sheet1 Summary; do
  officecli raw-set "$FILE" /workbook --xpath "//x:sheet[@name='$S']" --action setattr --xml "state=hidden" || true
done
```

Data／Summary シートに設定された `Print_Area` は削除すること — スコープが競合すると複数ページ出力になる。

## QA（必須 — 納品ゲート）

**問題は存在すると想定せよ。それを見つけるのがあなたの仕事だ。** チャートがレンダリングされたことは、そのチャートに意味があることを意味しない。「validate が通った」は納品ではない。「Dashboard シートがその事業を知る人が作ったように読める」ことが納品である。

### "done" 前の最低限のサイクル

xlsx のベースライン（`view issues`、数式エラークエリ、`validate`、HTML プレビューのスキャン）を継承する: → officecli-xlsx の §QA minimum cycle を参照。

その後、ダッシュボード固有の納品ゲートを実行する。各ゲートは `.data.*` ラッパー付きの **COUNT-then-if** パターンを使う — `&& echo OK || echo FAIL` を連結してはならない。

**ゲート1 — KPI 数式カバレッジ。** 計画したすべての KPI セルに数式があること。`-lt 2` は計画に合わせて調整する（KPI 4個の計画なら `-lt 4`）。

```bash
KPI_FORMULAS=$(officecli query "$FILE" 'Dashboard!:has(formula)' --json | jq '.data.results | length')
[ "$KPI_FORMULAS" -lt 2 ] && { echo "REJECT Gate 1: $KPI_FORMULAS formula cells on Dashboard"; exit 1; }
```

**ゲート2 — チャート数が計画通りで、すべてのチャートにデータとタイトル幅の妥当性があること。**

```bash
CHART_COUNT=$(officecli query "$FILE" chart --json | jq '.data.results | length')
[ "$CHART_COUNT" -lt 1 ] && { echo "REJECT Gate 2: zero charts"; exit 1; }
col_num () { local c=$1 n=0; for ((k=0;k<${#c};k++)); do n=$((n*26+$(printf '%d' "'${c:$k:1}")-64)); done; echo "$n"; }
for i in $(seq 1 "$CHART_COUNT"); do
  JSON=$(officecli get "$FILE" "/Dashboard/chart[$i]" --json)
  SC=$(echo "$JSON" | jq -r '.data.results[0].format.seriesCount // 0')
  TITLE=$(echo "$JSON" | jq -r '.data.results[0].format.title // ""')
  ANCHOR=$(echo "$JSON" | jq -r '.data.results[0].format.anchor // ""')
  [ "$SC" = "0" ] || [ -z "$TITLE" ] && { echo "REJECT Gate 2: chart[$i] seriesCount=$SC title='$TITLE'"; exit 1; }
  [ -z "$ANCHOR" ] && continue
  LCOL=$(echo "${ANCHOR%%:*}" | sed 's/[0-9]*$//'); RCOL=$(echo "${ANCHOR##*:}" | sed 's/[0-9]*$//')
  SPAN=$(( $(col_num "$RCOL") - $(col_num "$LCOL") + 1 ))
  MIN=$(( (${#TITLE} * 18 + 99) / 100 ))
  [ "$SPAN" -lt "$MIN" ] && { echo "REJECT Gate 2: chart[$i] title=${#TITLE} chars needs width ≥ $MIN, anchor spans $SPAN"; exit 1; }
done
```

プリセット `minimal` / `magazine` の狭いタイトルは、0.18 の係数より早く切れることがある — 目視確認すること。

**ゲート3 — チャートのシリーズ名が設定されていること（凡例に "Series1" が残っていないこと）。**

```bash
for i in $(seq 1 "$CHART_COUNT"); do
  BAD=$(officecli get "$FILE" "/Dashboard/chart[$i]" --json | jq '[.data.results[0].children[]? | select(.type == "series") | select((.format.name // "") | test("^Series[0-9]+$"; "i"))] | length')
  [ "$BAD" -gt 0 ] && { echo "REJECT Gate 3: chart[$i] has $BAD auto-named series"; exit 1; }
done
```

**ゲート4 — Data シート（10行以上）に CF ルールがあること。**

```bash
CF_COUNT=$(officecli query "$FILE" conditionalformatting --json | jq '.data.results | length')
[ "$CF_COUNT" -lt 1 ] && { echo "REJECT Gate 4: zero CF rules on 10+ row data sheet"; exit 1; }
```

注: `query conditionalformatting` が正規の要素名である。`query cf` は 0 を返す（エイリアスではない）。

**ゲート5 — activeTab と fullCalcOnLoad が設定されていること。** 実際の Dashboard インデックスと比較する（Dashboard がインデックス0であることは正しい合格とみなす）。

```bash
DASH_IDX=$(officecli query "$FILE" sheet --json | jq '[.data.results[].path] | index("/Dashboard")')
ACTIVE=$(officecli get "$FILE" /workbook --json | jq '.data.results[0].format.activeTab // -1')
FULLCALC=$(officecli get "$FILE" /workbook --json | jq -r '.data.results[0].format["calc.fullCalcOnLoad"] // false')
[ "$ACTIVE" != "$DASH_IDX" ] && { echo "REJECT Gate 5: activeTab=$ACTIVE Dashboard=$DASH_IDX"; exit 1; }
[ "$FULLCALC" != "true" ] && { echo "REJECT Gate 5: calc.fullCalcOnLoad=$FULLCALC — stale caches will ship"; exit 1; }
```

**ゲート6 — プレースホルダーの一斉点検。** レンダリングされた出力にビルド時のトークンが残っていないこと。

```bash
LEAKS=$(officecli view "$FILE" text 2>/dev/null | grep -niE '\{\{|\$fy\$|<TODO>|xxxx|TBD' | wc -l | tr -d ' ')
[ "$LEAKS" -gt 0 ] && { echo "REJECT Gate 6: $LEAKS placeholder tokens"; exit 1; }
```

**ゲート7 — 視覚的な納品下限（xlsx から移植）。** `officecli view "$FILE" html` を実行し、返された HTML パスを Read すること。以下を確認する:

- Dashboard／Data のどのセルにも `###` が出ていないこと（列幅が狭すぎない）。
- KPI ラベル、シートタブ名、チャートタイトルが切り詰められていないこと。
- プレースホルダートークンがテキストとしてレンダリングされていないこと（`$fy$24`、`{var}`、`<TODO>`、`xxxx`）。
- 円グラフ／ドーナツグラフのスライスがそれぞれ異なる塗り色でレンダリングされていること（LibreOffice で潰れて見える場合は、壊れていると断定する前にユーザーの対象ビューアで確認する — → officecli-xlsx の §Known Issues/Renderer caveats を参照）。
- 空のチャートアンカーがないこと — すべてのチャートに可視で妥当なプロットがあること。
- Dashboard シートが最初に開くこと（タブがハイライトされ、アクティブ領域が先頭にスクロールされている）。

`view html` がブロックされた場合（レンダラーの競合、ヘッドレス、ポート使用中）でも、ゲート7は依然として**必須**である — フォールバックのチェックを**すべて**実行すること:

```bash
# a) Token / ### sweep.
officecli view "$FILE" text 2>/dev/null | grep -nE '###|\{\{|<TODO>|\$fy\$|xxxx' && { echo "REJECT Gate 7: tokens or ### present"; exit 1; }
# b) Per-KPI: cachedValue length × coef must fit col width. coef=0.55 fit-to-page, 0.85 otherwise.
for CELL in A2 C2 E2 G2; do
  CV=$(officecli get "$FILE" "/Dashboard/$CELL" --json | jq -r '.data.results[0].format.cachedValue // .data.results[0].text // ""')
  W=$(officecli get "$FILE" "/Dashboard/col[${CELL%%[0-9]*}]" --json | jq -r '.data.results[0].format.width // 0')
  CAP=$(echo "$W * 0.55" | bc -l | awk '{print int($1)}')
  [ "${#CV}" -gt "$CAP" ] && { echo "REJECT Gate 7: $CELL '$CV' (${#CV} chars) > cap $CAP"; exit 1; }
done
# c) Rerun Gate 2 title × 0.18 ≤ anchor span.  d) Log which fallback was used and why.
```

ゲート7を**絶対に**スキップしてはならない — スキップすればユーザーに `###` が届く。

シーンのキーワードに print / 一页 / board / 投资人 / 董事会 が含まれる場合は、ゲート7を構造的な印刷スコープチェックで拡張する:

```bash
if echo "$USER_REQ" | grep -qiE 'print|一页|投资人|董事会|board'; then
  # Every non-Dashboard sheet must be hidden or veryHidden.
  # query sheet --json carries the name in .preview (no .name/.state); visibility lives in
  # each sheet's own get .format.hidden (absent => visible), so iterate + probe per sheet.
  LEAKING=""
  while IFS=$'\t' read -r SPATH SNAME; do
    [ "$SNAME" = "Dashboard" ] && continue
    HIDDEN=$(officecli get "$FILE" "$SPATH" --json | jq -r '.data.results[0].format.hidden // false')
    [ "$HIDDEN" != "true" ] && LEAKING="$LEAKING $SNAME"
  done < <(officecli query "$FILE" 'sheet' --json | jq -r '.data.results[] | [.path, .preview] | @tsv')
  [ -n "$LEAKING" ] && { echo "REJECT Gate 7 print-scope: visible non-Dashboard sheet(s):$LEAKING — hide before delivery"; exit 1; }
  # Dashboard must carry an explicit Print_Area named range.
  PA=$(officecli query "$FILE" 'namedrange[name="_xlnm.Print_Area"]' --json | jq '.data.results | length')
  [ "$PA" -ge 1 ] || { echo "REJECT Gate 7 print-scope: no _xlnm.Print_Area set"; exit 1; }
fi
```

ユーザーは最終的な印刷プレビューのために、対象ビューア（Office / WPS / Numbers）でファイルを開く — 本スキルはエクスポート成果物のレンダリングは行わない。

**ゲート8 — 数式の健全性（cachedValue が本物であり、古い値やエラーではないこと）。** `fullCalcOnLoad=true` は**実行時**の再計算を保証するが、ビルド時の XML 内キャッシュは更新しない — したがって、すべての数式セルは今この時点で、空でなく、ゼロでなく、エラーでもない `cachedValue` を保持していなければならない。

```bash
for CELL in A2 C2 E2 G2; do
  JSON=$(officecli get "$FILE" "/Dashboard/$CELL" --json)
  [ -z "$(echo "$JSON" | jq -r '.data.results[0].format.formula // ""')" ] && continue
  CV=$(echo "$JSON" | jq -r '.data.results[0].format.cachedValue // ""')
  case "$CV" in
    "" | "0" | "#DIV/0!" | "#REF!" | "#N/A" | "#VALUE!" | "#NAME?" | "null")
      echo "REJECT Gate 8: $CELL cachedValue='$CV' — re-issue formula or close+reopen"; exit 1 ;;
  esac
done
```

KPI が本当にゼロである場合（例:「今四半期の退職者数」＝0）は、ループ内でホワイトリストに入れて明記すること — デフォルトの前提は「ゼロは壊れている」である。

何か失敗した場合は、根本原因を修正し、サイクル全体を再実行すること。

### 正直な限界

散布図は `series1.xValues`（未対応）を受け付けない — X 軸は `categories` / `series1.categories` 経由で与える。LibreOffice のチャート色のずれ／円グラフスライスの潰れ／チェックボックスの二重枠はビューアの表示上の癖であり、まず Office / WPS / Numbers で目視確認すること。

## Reference

- **`add` 時の shorthand `--type`:** `chart`、`sparkline`、`databar`、`colorscale`、`iconset`、`formulacf`。CF ルールは `help xlsx conditionalformatting` にマッピングされる。パスの接尾辞は `/Sheet/cf[N]`。
- **完全なスキーマは help にある:** `officecli help xlsx chart` / `sparkline` / `conditionalformatting`。本スキルはそれらを複製しない。
- **DeferredAddKeys（add のみ）:** `combosplit`、`holesize`。D-1 参照。（`preset`、`trendline`、`referenceline`、`axisNumFmt` は現在 `set` でも使える — help に `[add/set]` と表示される。）
- **構築順序:** チャート＋スパークライン＋CF＋tabColors を先に → 高水準の `set` で `calc.fullCalcOnLoad=true` → `raw-set activeTab` は**最後**（すべてのシートが存在した後）。

## 既知の問題と落とし穴

### ダッシュボード固有

| # | Issue | Mitigation |
|---|---|---|
| D-1 | `combosplit` は DeferredAddKey — `add` 時のみ有効。（`preset`、`referenceline`、`trendline`、`axisNumFmt` は現在 `set` でも適用可能 — help に `[add/set]` と表示される。） | `combosplit` は `add` 時に設定すること。後から適用はできない — 削除して再追加する。他の4つは `set` で作成後に適用・変更できる。 |
| D-2 | `referenceline` の形式は `value:color:label:dash`（color が label より先）。`"0:Break-Even:FF0000:dash"` は `Invalid color value` で失敗する。 | 順序は value、color、label、dash。 |
| D-3 | 散布図は `series1.xValues` を受け付けない（未対応）。X軸は `categories` / `series1.categories` 経由で与える。 | `--prop series1.categories="Sheet1!A2:A13"`（または `--prop categories="Sheet1!A2:A13"`） |
| D-4 | `formulacf` は `fill` と `font.color` に加え、`font.bold` / `font.italic` を尊重する（dxf の font に書き込まれ、読み戻し時にも反映される）。 | `fill` / `font.color` / `font.bold` / `font.italic` のいずれかで CF ルールを表現する。 |
| D-5 | Dashboard の列幅はデフォルト 8.43 — 24pt bold の KPI 値は `###` になる | cachedValue の桁数帯でサイズを決める: 4〜6桁 → 22〜24；7〜9桁（百万単位）→ 26〜30；10桁以上（億／billion）→ 32〜36；百億／10桁＋通貨記号＋fit-to-page ランドスケープ → **40〜44**。数式 `ceil((visible_chars+2)*1.3)` は出発点にすぎない — 必ずゲート7のフォールバック b) で検証すること。スパークライン用の列: 12。 |
| D-6 | `raw-set activeTab` は**最後の**変更でなければならない。すべてのシートが揃う前に挿入するとインデックスがずれる。 | すべてのシート／チャート／CF／スパークライン／tabColors を仕上げてから `raw-set` する。 |
| D-7 | `raw-set` 経由の `calc.fullCalcOnLoad` は `<calcPr>` の重複を生み validate が失敗する | `officecli set "$FILE" / --prop calc.fullCalcOnLoad=true` を使う。 |
| D-8 | LibreOffice はレンダリング時に非表示列の数式を評価しない → 非表示セルを参照するチャートは空白でレンダリングされる | 可視の Summary シートに集計し、チャートは Summary から読み取る。チャートのデータ元でない列だけを非表示にする。 |
| D-9 | `chartType=pie` は LibreOffice で空白レンダリングになる | 部分-全体の内訳には安全な代替として `doughnut` を使う。 |
| D-10 | 日付条件を伴う `SUMIFS` / `AVERAGEIFS` は、条件が文字列だと静かに失敗する | `DATE()` または `DATEVALUE()` でラップする: `=SUMIFS(B2:B13,A2:A13,DATE(2025,1,5))`。 |
| D-11 | Summary シートのパーセンテージ数式は `numFmt` なしでは生の小数（0.098）として表示される | 数式と同じ `set` 呼び出しで `numFmt="0.0%"` を設定する。 |
| D-12 | `import --header` はフリーズ枠と AutoFilter を設定するが、列幅は設定しない。 | `col[]` に幅を設定する。`col[]` パスへの `numFmt` は現在、列レベルのスタイル（`<col s=...>`、スキーマ有効、読み戻し時に `numberformat` として表示）を適用する。これは列内の空セルを書式付ける。独自スタイルを持つセルにはセル範囲ごとの `numFmt` が別途必要。 |
| D-13 | スパークラインの `highpoint` は bool（ハイライトの on/off）であり、色ではない。`--prop highpoint=FF0000` は `Invalid boolean value` でエラーになる | `--prop highPoint=true --prop highMarkerColor=FF0000`。lowPoint / firstPoint / lastPoint とその *MarkerColor も同じパターン。 |
| D-14 | スパークラインの横断的（cross-sectional）データは意味を持たない（地域や部門には順序がない） | 行が時系列（日付、月、四半期）でない限り、スパークラインは省略する。 |
| D-15 | 空のチャート `add` は CLI 層で拒否される（`Chart requires a 'data' property`）— サイレント受理に依存していた旧来のスキルはここで失敗する | チャート `add` 時には常に `series1.values=` / `dataRange=` / インラインの `data=` のいずれかを与えること。ゲート2の seriesCount チェックは、念のための追加検証として扱う。 |
| D-16 | `fullCalcOnLoad=true` は、エンドユーザーがファイルを開いた時に**実行時**の再計算を保証するが、XML 内のビルド時 `cachedValue` は更新しない。構築順序 `set B=100 → set E==B+D → fix B=150` では `E.cachedValue` が古いまま残る（ボードには "Net Change = 0" と見える）。 | すべての上流編集が確定した後、下流のすべての数式を再発行する（`officecli set "$FILE" /Sheet/E2 --prop formula==B2+D2`）か、`close` して再度開く。ゲート8で検証する。 |
| D-17 | 組み込みの計算エンジンは、配列述語形式の `SUMPRODUCT((A2:A97=X)*C2:C97*D2:D97)` を評価しない — cachedValue は `0`/`null` のままとなり、ゲート8で拒否される。実行時の Excel / WPS では正しく計算されるが、キャッシュが古いまま納品された XLSX は依然として `0` を出す。 | ヘルパー列＋`SUMIF` に書き換える: ソースシート上で `F2==C2*D2`、その後 `=SUMIF(B:B, "Region X", F:F)`。あるいは Summary シートで事前集計し、そこからチャートを作る。 |

### 継承（ポインターのみ）

シート横断の `!` の罠、作成後のチャート `anchor` ／シリーズの不変性 → officecli-xlsx の §Known Issues を参照。
