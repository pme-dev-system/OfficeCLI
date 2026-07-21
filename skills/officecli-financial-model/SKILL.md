---
name: officecli-financial-model
description: "ユーザーが Excel で財務モデル — 3ステートメントモデル、DCF バリュエーション、LBO、SaaS ユニットエコノミクス、感応度/シナリオ分析、デットスケジュール、資金調達プロジェクションなど — を構築したい場合にこのスキルを使用する。トリガーとなる語句: 'financial model', '3-statement model', 'P&L + BS + CF', 'DCF', 'WACC', 'NPV', 'terminal value', 'LBO', 'debt schedule', 'cash sweep', 'MOIC', 'IRR / XIRR', 'sensitivity table', 'scenario analysis', 'ARR model', 'unit economics', 'CAC / LTV', 'cap table forecast'。出力は単一の数式駆動の .xlsx。このスキルは officecli-xlsx の上に乗るシーンレイヤーであり、xlsx v2 のすべてのルール（4色コード、ビジュアルフロア、数値フォーマット、キャッシュドリフト、Known Issues、Delivery Gate の最小サイクル）を継承する。単純な予算トラッカー、CSV ダンプ、運用 KPI シートには invoke しないこと — それらは officecli-xlsx ベースにルーティングする。"
---

# OfficeCLI Financial-Model スキル

**このスキルは `officecli-xlsx` の上に乗るシーンレイヤーです。** xlsx のハードルール — シェルクォート、逐次実行、Help-First Rule、ビジュアル納品フロア、CFO 4色コード（青＝入力、黒＝数式、緑＝シート横断、黄色塗り＝前提条件）、数値フォーマット標準（年は文字列、ゼロは `-`、`%` は小数点1桁、負数は括弧表記）、前提セルの規律、CSV 一括インポート、チャートのデータフィード形式（a/b/c）、5ゲートの Delivery サイクル、キャッシュドリフトのガイダンス、Known Issues（シート横断の `!` トラップ、数式に対する batch + resident、レンダラーの注意点）— はすべて**継承済みであり、ここでは再教育しない**。このファイルが追加するのは、**財務モデル**が xlsx ベースの上に必要とするものだけ: 三層アーキテクチャ、3つのモデルタイプのレシピ（3ステートメント / DCF / LBO）、感応度＋シナリオのプロトコル、財務関数パターン、循環参照の規律、モデル固有の Delivery Gate 4〜6。

xlsx ベースルールで網羅されている箇所は、本文中で `→ see xlsx v2 §X` と記す。未読であれば先に `skills/officecli-xlsx/SKILL.md` を読むこと。

## セットアップ

`officecli` が未インストールの場合:

- **macOS / Linux**: `curl -fsSL https://d.officecli.ai/install.sh | bash`
- **Windows (PowerShell)**: `irm https://d.officecli.ai/install.ps1 | iex`

`officecli --version` で確認する（PATH が反映されていなければ新しいターミナルを開く）。インストールが失敗する場合は https://github.com/iOfficeAI/OfficeCLI/releases からバイナリをダウンロードする。

## Help-First Rule

このスキルが教えるのは財務モデルに必要な事項であり、全 CLI フラグではない。プロパティ名/エイリアス/enum が不確かな場合は、推測する**前に**ヘルプを確認する: `officecli help xlsx [element] [--json]`。ヘルプはインストール済みバージョンに対して権威があり、本スキルとヘルプが食い違う場合は**ヘルプが勝つ**。以下の各 `--prop X=` はすべて `officecli help xlsx <element>` で検証済み。

## メンタルモデルと継承

**xlsx v2 を継承する。** 先に `skills/officecli-xlsx/SKILL.md` を読むこと。このスキルは、`create` / `open` / `close`、値/数式の `set`、シート横断数式のための `batch` ヒアドキュメント、`/SheetName/A1` パス、名前付き範囲、5ゲートの Delivery サイクル、シート横断の `!` トラップ、そして**シート横断数式は non-resident（単一 batch または個別 `set`）で行い、resident のまま batch しては絶対にいけない**ことを既知の前提とする。

## シェルと実行の規律

シェルクォート、逐次実行、`$FILE` 慣習 → see xlsx v2 §Shell & Execution Discipline。ルールは同じ: すべての `[N]` パスをクォートする、`$` を含むプロパティはシングルクォートで囲む（本書に登場するすべての数値フォーマット `$#,##0;($#,##0);"-"` はシングルクォートが必要）、手書きの `\$`/`\t`/`\n` は禁止、コマンドは一度に一つ。以降の例では `$FILE`（`FILE="model.xlsx"`）を使用する。

## 核となる原則（アイデンティティ）

財務モデルとは、**意思決定グレードの数式駆動レイヤー**を備えた xlsx である: すべての出力は青字の前提条件まで途切れない連鎖でたどれ、すべてのステートメントは各期間でバランスし、すべてのバリュエーションは再監査可能である。一般的な xlsx に対する8つの差分:

1. **三層アーキテクチャは必須:** Inputs → Calc → Outputs。ゾーンを崩すと監査不能になる。
2. **前提条件はセルに置く。数式の中には決して埋め込まない。** `=B5*(1+Assumptions!GrowthRate)` であり、`=B5*1.05` ではない。
3. **ステートメントは各期間でバランスする。** `Assets − Liab − Equity = 0`、`CF.EndingCash = BS.Cash`。Gate 4 は `IMBALANCED` で失敗する。
4. **ハードコードは監査対象。** Calc シートのハードコード数値はゼロ、Gate 6 がカウントする。
5. **感応度/シナリオは一級市民。** 2軸グリッド、ドロップダウンの `INDEX/MATCH` スイッチ、または Base/Upside/Downside の列。Excel の Data Tables は確実にはサポートされない — 手動グリッドのみ。
6. **バリュエーションセルのキャッシュ値は納品を左右する。** キャッシュ済み結果がない（または `#OCLI_NOTEVAL!` センチネル付きの）バリュエーションセルは、再計算しない読者に空欄/誤った数値を送る。エバリュエーターは現在 NPV / XNPV / IRR / XIRR を計算する。読み戻しでキャッシュ値を検証すること。Gate 5 が抜き取り検査する。
7. **循環性は設計上の選択である。** 正当なリング（利息⇄現金、リボルバープラグ⇄期末現金）には `calc.iterate=true` を使う。意図しない循環性は壊れた代数であり、`iterate` で覆い隠してはならない。
8. **3回以上使用する前提条件には名前付き範囲を使う。** `WACC`、`TaxRate`、`TerminalGrowth`、`ExitMultiple`、`ChurnRate`。宣言済みで未使用の名前は無用の飾りであり、Gate 6 が検出する。

### 逆ハンドオフ — いつ xlsx ベースに戻るか

**xlsx ベース**にとどまるべきケース: 予算トラッカー、CSV からレポートへのダンプ、運用 KPI シート、単純なテンプレート、フォーキャストロジックのないキャップテーブル。**このスキル**を使うのは、依頼が 3ステートメント / DCF / WACC / NPV / TV / LBO / デットスケジュール / MOIC / IRR / ユニットエコノミクス / ARR ロールフォワード / 感応度グリッド / シナリオスイッチ / プロフォーマに言及している場合のみ。

## 三層アーキテクチャ（ハードルール）

本スキルのすべてのモデルは三層の上に構築する。**名前を付け、タブに色を付け、実行可能な監査で強制する。** ゾーンルールを破ることが、監査不能なモデルになる最も一般的な原因である。

| ゾーン | シート名（慣習） | タブ色 | 内容 | ハードコード | 数式 |
|---|---|---|---|---|---|
| **Inputs** | `Assumptions`、`Inputs`、`Drivers` | 黄色 `FFC000` | 生のドライバー: 成長率、利益率、税率、WACC、FTE、価格、運転資本日数 | すべてのセルに青 `0000FF` | 派生前提条件（例: `=MonthlyARPU*12`）にのみ許可 |
| **Calc** | `P&L`、`Balance Sheet`、`Cash Flow`、`DCF`、`Debt`、`ARR` | 青 `4472C4` | すべての導出とステートメント | **ゼロ**（Gate 6 が強制） | 同一シート内は黒 `000000`、シート横断は緑 `008000` |
| **Outputs** | `Summary`、`Dashboard`、`Sensitivity`、`Returns` | 緑 `70AD47` | KPI、感応度グリッド、チャート、リターンウォーターフォール | ラベルのみ（数値以外）；Gate 6 が数値ハードコードをカウント → 0 | 上記に準じ黒/緑 |

**ビルド順序はゾーン横断を意識する。** 前提条件を最初に、次に依存関係の連鎖（3ステートメントなら `IS → BS → CF`、DCF なら `FCF → WACC → NPV`）に沿って Calc をボトムアップに、最後に Outputs。Outputs を先に作ると至る所に `0` がキャッシュされ、下流もゼロを継承してしまう。

**実行可能なゾーン監査**（Gate 4 の前に実行）:

```bash
# Calc ゾーン: 数値ハードコードは0でなければならない。`cell:not(:has(formula))` はリテラル（非数式）セルを、`cell:has(formula)` は数式セルを選択する。
HARDCODE=$(officecli query "$FILE" 'cell[type=Number]' --json | jq '[.data.results[] | select(.format.formula == null) | select(.path | test("/(P&L|Balance Sheet|Cash Flow|DCF|Debt|ARR)/"))] | length')
[ "$HARDCODE" -eq 0 ] && echo "Zone audit OK" || { echo "REJECT: $HARDCODE hardcoded numeric cells on Calc sheets — move to Assumptions"; exit 1; }
# Assumptions ゾーン: 非ゼロであるべき。
INPUTS=$(officecli query "$FILE" '/Assumptions/cell[type=Number]' --json | jq '[.data.results[] | select(.format.formula == null)] | length')
[ "$INPUTS" -ge 5 ] && echo "Assumptions has $INPUTS hardcoded drivers" || echo "WARN: Assumptions has only $INPUTS inputs"
```

## 印刷用納品物（取締役会 / IC / LP）

依頼に「print」/「一页」/「董事会」/「投资人」/「IC memo」/「LP update」が含まれる場合、印刷パイプラインは Outputs ゾーン**のみ**を出力しなければならない。2つの成果物:

```bash
# 1. Print_Area を Outputs シート (Summary か Dashboard) にスコープする。
officecli add "$FILE" / --type namedrange --prop name=_xlnm.Print_Area --prop scope=Summary --prop 'refersTo=Summary!$A$1:$H$40'
# 2. Outputs 以外のシートをすべて非表示にする — Print_Area のスコープだけでは、印刷パイプラインが可視シートをすべて出力するのを止められない。
for S in Assumptions 'P&L' 'Balance Sheet' 'Cash Flow' DCF WACC Debt FCF 'S&U' Exit Returns; do
  officecli raw-set "$FILE" /workbook --xpath "//x:sheet[@name='$S']" --action setattr --xml "state=hidden" || true
done
# 3. Outputs シートに fit-to-page の横向きを設定する。
officecli raw-set "$FILE" /Summary --xpath "//x:worksheet" --action prepend --xml '<sheetPr xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><pageSetUpPr fitToPage="1"/></sheetPr>'
```

Calc シートに設定された `Print_Area` は削除すること — スコープが競合すると、Assumptions / ステートメントシートが漏れ出す複数ページ出力になる。

## ビルド順序とキャッシュドリフトのルール（3ステートメントで重大）

3つの事実がサイレントな誤数値を引き起こす: (1) 新しい数式がキャッシュ値なしで出荷される — Excel は開いたときに再計算するが、HTML プレビュー / 古いビューアはしない; (2) 上流と同じシーケンスで書かれた下流は、上流のキャッシュ前の状態から `0` をキャッシュする; (3) resident が開いた状態でのシート横断 `batch` は 3〜5 操作でデッドロックする。

**規律（すべてのレシピ共通）:**
- ビルド順序はデータ連鎖に従う: `P&L → BS → CF`（3ステートメント）；`FCF → WACC → NPV → Sensitivity`（DCF）；`S&U → Debt → P&L → CF → Returns`（LBO）。
- シート横断連鎖の後、**キャッシュ再更新パス:** すべてのサマリー / バリュエーション / バランスチェックセルに対して non-resident で `set` を再発行する。
- 抜き取り確認: `officecli get "$FILE" /Summary/B2 --json | jq '.data.results[0].format.cachedValue'` が妥当な非 null 値を返すこと。`null` は Excel が開いたときに計算することを意味する（納品としては OK）。セルが `#OCLI_NOTEVAL!` センチネルを示す場合: resident を閉じて再 set、それでも未評価なら cache-fallback（§財務関数パターン）。

## レシピ — 3つのモデルタイプ

以下の各レシピは**実行可能な骨格であり、ファイナンス理論ではない**。数値は差し替えて構わないが、構造は変更しないこと。すべてのレシピは `FILE="model.xlsx"` が設定済みで、`officecli create "$FILE"` + `officecli open "$FILE"` を実行済みであることを前提とする。最後に `officecli close "$FILE"` で締める。

### レシピ A — 3ステートメントモデル（P&L + BS + CF）

**このレシピが生成するもの。** 4シート: `Assumptions`、`P&L`、`Balance Sheet`、`Cash Flow`、加えて `Summary`。年次列は 2024A・2025E・2026E・2027E。BS にバランスチェック行、CF にキャッシュ照合行。すべてのステートメント行 = Assumptions への数式。

**ビルド順序（必須）。** `Assumptions → P&L → Balance Sheet → Cash Flow → Summary`。P&L より前に BS を作らないこと — `RetainedEarnings` は `NI` に依存する。BS より前に CF を作らないこと — `CF.OpeningCash = 前期の CF.EndingCash` の自己連鎖には、Y1 用に固定された BS の現金が必要。順序を誤ると、本スキルの Gate 4 バランスチェックはサイレントに失敗する。

**ステップ1 — シート＋タブ色＋ウィンドウ枠の固定。**

```bash
officecli add "$FILE" / --type sheet --prop name=Assumptions --prop tabColor=FFC000
officecli add "$FILE" / --type sheet --prop name=P&L --prop tabColor=4472C4
officecli add "$FILE" / --type sheet --prop name='Balance Sheet' --prop tabColor=4472C4
officecli add "$FILE" / --type sheet --prop name='Cash Flow' --prop tabColor=4472C4
officecli add "$FILE" / --type sheet --prop name=Summary --prop tabColor=70AD47
officecli set "$FILE" /Assumptions --prop freeze=B2
officecli set "$FILE" /P&L --prop freeze=B3
officecli set "$FILE" "/Balance Sheet" --prop freeze=B3
officecli set "$FILE" "/Cash Flow" --prop freeze=B3
```

**ステップ2 — 前提条件（青、主要ドライバーは黄色塗り）。** 年ヘッダーは2行目、ラベルは A 列を下に、青い数値入力は B:E。ドライバー: `RevenueGrowth`、`GrossMargin`、`OpExRatio`、`TaxRate`、`DaysReceivable/Inventory/Payable`、`CapExRatio`、`DepreciationYears`。B:E に `font.color=0000FF`。シナリオ切替対象の3〜5個のドライバーには黄色塗り（`fill=FFFF00`）。

**3回以上使用するドライバーには名前付き範囲を宣言し、それを参照する**（`StartingARR`、`TaxRate`、`OpeningCash`、`GrowthRate`、`GrossMargin`）。数式は `=Assumptions!B4` ではなく `=StartingARR`、`=EBT*Assumptions!B8` ではなく `=EBT*TaxRate`。宣言だけで未使用の名前は無用の飾りであり、Gate 6 で却下される。

**ステップ3 — P&L 行（すべて数式）。** 行: `Revenue` / `COGS` / `Gross Profit` / `OpEx` / `EBITDA` / `D&A` / `EBIT` / `Interest` / `EBT` / `Tax` / `Net Income`。すべての行が `Assumptions` または直前行のセルを参照する数式。以下は収益側ブロックの例 — **自分の行番号に差し替えること**。この例の行マップ: `B3=Revenue, B4=COGS, B5=Gross Profit, B7=OpEx, B9=EBITDA, B10=EBIT, B15=Net Income`。単一の non-resident batch として投入する:

```bash
cat <<'EOF' | officecli batch "$FILE"
[
  {"command":"set","path":"/P&L/B3","props":{"formula":"Assumptions!B5","font.color":"008000"}},
  {"command":"set","path":"/P&L/C3","props":{"formula":"B3*(1+Assumptions!C6)"}},
  {"command":"set","path":"/P&L/D3","props":{"formula":"C3*(1+Assumptions!D6)"}},
  {"command":"set","path":"/P&L/E3","props":{"formula":"D3*(1+Assumptions!E6)"}},
  {"command":"set","path":"/P&L/B4","props":{"formula":"-B3*(1-Assumptions!B7)"}},
  {"command":"set","path":"/P&L/B5","props":{"formula":"B3+B4"}}
]
EOF
```

Assumptions 参照（`B5`、`C6`、`B7`）もプレースホルダー行に過ぎない — より良いのは、**各ドライバーに名前付き範囲を定義する**こと（ステップ2）で、行レイアウトに関わらず数式が `=StartingRevenue*(1+RevenueGrowth_Y2)` のように読めるようにする。`OpEx` / `D&A` / `Interest` / `Tax` / `NI` についても同様に繰り返す。シート横断参照のセルにはすべて `font.color=008000`、同一シートのセルはデフォルトの `000000`。すべての $ 行に `numFmt='$#,##0;($#,##0);"-"'`。

**ステップ4 — Balance Sheet 行（すべて数式）。** Assets = `Cash + AR + Inventory + Net PP&E`。Liab = `AP + Debt`。Equity = `OpeningEquity + RetainedEarnings`。運転資本の行は Days の前提条件を使用: `AR = Revenue × DaysReceivable / 365`。`Net PP&E` はロールフォワード: `Beg + CapEx − Depreciation`。**`BS.Cash` は独立したプラグではない** — 必ず `'Cash Flow'!B<期末現金の行>`（ステップ5で入力）と等しくなければならない。

**Retained Earnings — 毎期生きた数式で。** `BS.RE(t) = BS.RE(t-1) + 'P&L'!NI(t) − Dividends(t)`。RE をハードコードすると整数ドルに丸められ、BS が毎期 ±1ドルずれる（CFO には「モデルがバランスしていない」と読まれる）。Y1 の Historical RE（前期 NI がない）は BS の恒等式から**生きた数式**として計算する: `BS!RE_Y1 = TotalAssets − TotalLiabilities − PaidInCapital`。Y1 セルには青字＋クラシックコメント; Y2..Y5 は NI 駆動のまま。

**ステップ5 — Cash Flow 行（すべて数式）。** Operating: `NI + D&A − ΔWorkingCapital`。Investing: `−CapEx`。Financing: `ΔDebt − Dividends`。Ending Cash = `Opening + Operating + Investing + Financing`。**Year 2 以降の Opening Cash = 前期の Ending Cash** — 同一シート上の自己連鎖: `C17 = B19`、`D17 = C19`、`E17 = D19`。Y1 の `OpeningCash` は Assumptions の入力。

**ステップ6 — バランスチェック＋キャッシュ照合行（納品の必須チェック）。** この例の行マップ: `Balance Sheet: B10=Total Assets, B15=Total Liab, B17=Total Equity, B18=Balance Check`；`Cash Flow: B5=BS.Cash（シート横断アンカー）, B19=CF.Ending Cash, B21=CF-BS Cash Recon`。自分のレイアウトの行番号に差し替えること — チェックの本質はロジックであり、セルアドレスではない。

```bash
cat <<'EOF' | officecli batch "$FILE"
[
  {"command":"set","path":"/Balance Sheet/B18","props":{"formula":"IF(ABS(B10-B15-B17)<0.01,\"OK\",\"IMBALANCED: \"&ROUND(B10-B15-B17,0))","bold":"true","font.color":"000000"}},
  {"command":"set","path":"/Cash Flow/B21","props":{"formula":"IF(ABS(B19-'Balance Sheet'!B5)<0.01,\"OK\",\"CF != BS CASH: \"&ROUND(B19-'Balance Sheet'!B5,0))","bold":"true"}}
]
EOF
```

C/D/E 列にも複製する。赤塗り（`fill=FFC7CE`）を `type=containsText --prop text=IMBALANCED` または `text="CF !="` で条件付き適用する。Gate 4 はこれらの行に対して query を実行し、`IMBALANCED` が1つでもあれば納品を拒否する。

**ステップ7 — キャッシュ再更新＋フォーマットパス。** `Summary` のすべてのサマリーセル、すべてのバランスチェック / 照合セル、BS / CF のすべてのシート横断参照を再 set する（non-resident、シートごとに1バッチ）。列幅を適用し（`col[A]=28`、`col[B:E]=15`）、すべてのドル行に `numberformat='$#,##0;($#,##0);"-"'`、セクションヘッダー行（REVENUE / COGS / ASSETS / LIABILITIES）にヘッダー塗り（`fill=1F3864`、`font.color=FFFFFF`、`bold=true`）を適用する。ヘッダー塗りはラベルセルだけでなく A:E 全体を覆うこと（→ xlsx v2 §visual floor）。

**ステップ8 — Summary / Dashboard の KPI＋チャート。** 最低4つの KPI: `Revenue 27E`、`EBITDA Margin 27E`、`Ending Cash 27E`、`Net Income CAGR` — それぞれステートメントセルを参照する数式で、緑フォント。

**取締役会/エグゼクティブ向けに納品する Dashboard にはチャートを最低3枚** — 1枚は下書き、3枚で納品物。マージン用チャートを追加する前に、`Summary!A10:E13` に Gross Margin / EBITDA Margin / NI Margin の比率行（`P&L` を参照する数式）を先に用意すること。

```bash
# (1) トップライン推移（Revenue + EBITDA）。
officecli add "$FILE" /Summary --type chart --prop chartType=column --prop dataRange='P&L!A2:E5' --prop title='Revenue & EBITDA' --prop width=14cm --prop height=8cm
# (2) マージン推移（Gross / EBITDA / NI margin）。
officecli add "$FILE" /Summary --type chart --prop chartType=line --prop dataRange='Summary!A10:E13' --prop title='Margin trend' --prop width=14cm --prop height=8cm
# (3) キャッシュ推移（Ending Cash ± Runway）。
officecli add "$FILE" /Summary --type chart --prop chartType=area --prop dataRange='Cash Flow!A19:E19' --prop title='Ending cash' --prop width=14cm --prop height=8cm
```

**検証（3つとも実行）:**

```bash
# バランスチェックは全期間で OK と表示されなければならない（範囲取得 → .data.results[0].children[] の下にセルがある）
officecli get "$FILE" "/Balance Sheet/B18:E18" --json | jq '.data.results[0].children[] | .format.cachedValue // .text'
# キャッシュ照合は全期間で OK でなければならない
officecli get "$FILE" "/Cash Flow/B21:E21" --json | jq '.data.results[0].children[] | .format.cachedValue // .text'
# Summary の KPI は妥当な数値であり、null であってはならない
officecli get "$FILE" "/Summary/B2:B5" --json | jq '.data.results[0].children[].format.cachedValue'
```

### レシピ B — DCF バリュエーション

**このレシピが生成するもの。** シート: `Assumptions`、`FCF`（10年予測）、`WACC`（パネル）、`DCF`（NPV + TV + エクイティブリッジ）、`Sensitivity`（2軸グリッド）。出力: `Implied Equity Value` + `Implied Per-Share`、`WACC × g` 感応度付き。

**ビルド順序。** `Assumptions → FCF → WACC → DCF → Sensitivity`。

**ステップ1 — 主要ドライバーの名前付き範囲。** DCF の可読性は名前次第。以下のすべての数式は `$B$6` ではなく `WACC`、`TaxRate`、`g` を使う:

```bash
cat <<'EOF' | officecli batch "$FILE"
[
  {"command":"add","parent":"/","type":"namedrange","props":{"name":"WACC","ref":"WACC!$B$12"}},
  {"command":"add","parent":"/","type":"namedrange","props":{"name":"TaxRate","ref":"Assumptions!$B$8"}},
  {"command":"add","parent":"/","type":"namedrange","props":{"name":"TerminalGrowth","ref":"Assumptions!$B$15"}},
  {"command":"add","parent":"/","type":"namedrange","props":{"name":"NetDebt","ref":"Assumptions!$B$20"}},
  {"command":"add","parent":"/","type":"namedrange","props":{"name":"SharesOut","ref":"Assumptions!$B$21"}}
]
EOF
```

**ステップ2 — FCF の構築（10年）。** 列 B:K = Y1..Y10。行: `Revenue`（成長率から）/ `EBIT`（収益 × 利益率）/ `EBIT × (1 − TaxRate)`（NOPAT）/ `+ D&A` / `− CapEx` / `− ΔNWC` / `= FCF`。Assumptions 駆動の比率を使う（`CapEx = Revenue × CapExRatio`）。すべてのセルは数式、黒フォント、`numFmt='$#,##0;($#,##0);"-"'`。

**ステップ3 — WACC パネル。** `WACC` シートに8行のパネル: `Risk-free rate` / `Equity risk premium` / `Beta` / `Cost of equity`（=Rf + β×ERP）/ `Pre-tax debt cost` / `After-tax debt cost`（=×(1−TaxRate)）/ `Equity weight` / `Debt weight` / `WACC`（=We×Re + Wd×Rd_after_tax）。入力は青、派生行は黒。

**ステップ4 — ターミナルバリュー＋NPV＋エクイティブリッジ。** 行マップ: `DCF: B/C 3=TV, 4=PV explicit FCF, 5=PV terminal, 6=EV, 7=Net Debt, 8=Equity Value, 9=Per-Share`；`FCF: row 2 = periods (1..10), row 11 = FCF, B:K = Y1..Y10`。自分の行番号に差し替えること。ノート欄のセルは `{"value":"text"}` を使い、`{"formula":"..."}` は決して使わない — 数式風の文章は開いたときに `#NAME?` を返す（レシピ C ステップ5後のコールアウト参照）。`DCF` シートで:

```bash
cat <<'EOF' | officecli batch "$FILE"
[
  {"command":"set","path":"/DCF/B3","props":{"value":"Terminal value (Gordon growth)"}},
  {"command":"set","path":"/DCF/C3","props":{"formula":"FCF!K11*(1+TerminalGrowth)/(WACC-TerminalGrowth)","font.color":"008000","numberformat":"$#,##0;($#,##0);\"-\""}},
  {"command":"set","path":"/DCF/B4","props":{"value":"PV of explicit-period FCF (10 yr)"}},
  {"command":"set","path":"/DCF/C4","props":{"formula":"SUMPRODUCT(FCF!B11:K11/(1+WACC)^FCF!B2:K2)","font.color":"008000","numberformat":"$#,##0;($#,##0);\"-\""}},
  {"command":"set","path":"/DCF/B5","props":{"value":"PV of terminal value"}},
  {"command":"set","path":"/DCF/C5","props":{"formula":"C3/(1+WACC)^10","numberformat":"$#,##0;($#,##0);\"-\""}},
  {"command":"set","path":"/DCF/B6","props":{"value":"Enterprise value"}},
  {"command":"set","path":"/DCF/C6","props":{"formula":"C4+C5","bold":"true","numberformat":"$#,##0;($#,##0);\"-\""}},
  {"command":"set","path":"/DCF/B7","props":{"value":"Less: Net debt"}},
  {"command":"set","path":"/DCF/C7","props":{"formula":"-NetDebt","font.color":"008000","numberformat":"$#,##0;($#,##0);\"-\""}},
  {"command":"set","path":"/DCF/B8","props":{"value":"Equity value"}},
  {"command":"set","path":"/DCF/C8","props":{"formula":"C6+C7","bold":"true","font.color":"000000","numberformat":"$#,##0;($#,##0);\"-\""}},
  {"command":"set","path":"/DCF/B9","props":{"value":"Implied per-share"}},
  {"command":"set","path":"/DCF/C9","props":{"formula":"C8/SharesOut","bold":"true","numberformat":"$0.00"}}
]
EOF
```

**`NPV` と `SUMPRODUCT` — どちらも正しくキャッシュされる。** エバリュエーターは `NPV(rate, cross_sheet_range)` を計算し正しい値をキャッシュする（読み戻しで検証すること）。`SUMPRODUCT(values/(1+rate)^periods)` は代数的に等価な形（期間行 `FCF!B2:K2 = 1..10` は一度きりのセットアップ）であり、ここでは監査の可読性のためのオプション選択として示しているに過ぎない — 明示的な割引系列の方がレビュアーにとって追いやすい場合があるからで、キャッシュの回避策ではない。どちらの形でも構わない。不規則な日付には `XNPV(rate, values, dates)` も同様にキャッシュされる；SUMPRODUCT の等価形は `SUMPRODUCT(values/(1+rate)^((dates-base_date)/365))`。

**ステップ5 — 2軸感応度グリッド（WACC × g）。** 5×5 グリッド。行 = WACC の値 `7.5% ... 11.5%`、列 = `g` の値 `1.5% ... 3.5%`。各セルは、グリッドの WACC と g を代入して DCF を再実行する独立した数式1個。テンプレート:

```bash
# セル D14（最初のデータセル、グリッドのアンカーは C14 = WACC ラベル、C15 = 最初の WACC 値）
# このセルの g である $D$13 と、このセルの WACC である $C15 を、複製した EV + equity 数式に代入する。
cat <<'EOF' | officecli batch "$FILE"
[
  {"command":"set","path":"/Sensitivity/D15","props":{"formula":"(NPV($C15,FCF!$B$11:$K$11)+(FCF!$K$11*(1+D$14)/($C15-D$14))/(1+$C15)^10+(-NetDebt))/SharesOut","numberformat":"$0.00"}}
]
EOF
```

D15:H19（5×5 グリッド）に数式をコピーする。14行目には g の値（青入力）、C列には WACC の値（青入力）。13行目と B列にはラベル。パッと読める3色グラデーション CF を適用する（緑＝アップサイド、赤＝ダウンサイド）:

```bash
officecli add "$FILE" /Sensitivity --type conditionalformatting \
  --prop type=colorScale --prop ref=D15:H19
```

**Excel の Data Tables は不可。** Excel ネイティブの `/Data/Table`（2変数テーブル）は CLI 経由では確実にサポートされない — 各グリッドセルは明示的な数式でなければならない。テンプレートをコピーし、`Data Table` の入力セルは試みないこと。

**検証。**

```bash
officecli get "$FILE" "/DCF/C8" --json | jq '.data.results[0].format.cachedValue'   # エクイティバリュー、妥当な$金額
officecli get "$FILE" "/DCF/C9" --json | jq '.data.results[0].format.cachedValue'   # 一株当たり、$XX.XX の範囲
officecli get "$FILE" "/Sensitivity/F17" --json | jq '.data.results[0].format.cachedValue'   # グリッド中心セル、妥当な値
```

`C8` や `C9` が `null` または `#OCLI_NOTEVAL!` センチネルを返す場合は、（non-resident で）再 set する — §ビルド順序とキャッシュドリフト参照。

### レシピ C — LBO モデル

**このレシピが生成するもの。** シート: `Assumptions`、`S&U`（Sources & Uses）、`Debt`（マルチトランシェスケジュール）、`P&L`（5年）、`CF`、`Exit` / `Returns`。出力: `MOIC`、`IRR`、4段階のリターンウォーターフォール。LBO はストレステストであり、循環参照（利息⇄現金）、最も深いシート横断連鎖、そして最も多用される名前付き範囲が予想される。

**ビルド順序。** `Assumptions → S&U → P&L → Debt → CF → Exit → Returns`。P&L を Debt より先に（デットの利息はカバレッジチェックのため P&L の EBIT に依存する）；Debt を CF より先に（CF は利息＋元本償却を使う）。ステップ5の前に `calc.iterate` を有効化する。

**ステップ1 — Sources & Uses（バランス必須、すべての手数料行を項目別に）。**

```
Uses    = Purchase_EV (EntryEBITDA × EntryMultiple) + Transaction_fees (Purchase_EV × TxnFeePct, typ 1.5–2.5%)
        + Financing_fees ((Senior + Mezz) × FinFeePct, typ 1–3%) + Refinanced_debt
Sources = Senior_TLB + Mezz + Revolver_drawn + Sponsor_equity
```

**スポンサーエクイティ — どちらか一方を選び、両方は使わない。** (a) **固定額:** `Sponsor_equity = Assumptions!SponsorEquity` とし、senior/mezz をスケーリングして Sources = Uses にする（手数料は debt が吸収し、サイレントなプラグにはしない）。(b) **逆算:** `Sponsor_equity = Uses − Senior − Mezz − Revolver − Refinanced`、ラベルは「Sponsor Equity (solved)」、独立した Assumptions 参照は持たない。ハードコードされた `SponsorEquity` に加えて `=Uses − Senior − Mezz` のプラグを持つと、手数料がサイレントに吸収される — 固定額 $140M に対しプラグ $194.67M = $54.67M の手数料が説明不能になり、CFO に一目で却下される。

```bash
# Sources = Uses のハードチェック。
officecli set "$FILE" /S&U/B12 --prop formula='IF(ABS(SUM(B4:B7)-SUM(B9:B11))<1,"BALANCED","S&U IMBALANCE: "&ROUND(SUM(B4:B7)-SUM(B9:B11),0))' --prop bold=true

# 固定額 vs プラグの整合性チェック（Gate 4 の追補；パターン(a)を選んだ場合のみ実行）。
STATED=$(officecli get "$FILE" /Assumptions/B12 --json | jq -r '.data.results[0].format.cachedValue // "null"')
PLUGGED=$(officecli get "$FILE" /S&U/B10 --json | jq -r '.data.results[0].format.cachedValue // "null"')   # B10 = S&U のスポンサーエクイティ行
DELTA=$(python3 -c "print(abs(float('$STATED') - float('$PLUGGED')))" 2>/dev/null || echo 99999)
python3 -c "import sys; sys.exit(0 if float('$DELTA') <= 1 else 1)" && echo "S&U sponsor OK (stated=$STATED plug=$PLUGGED)" || { echo "REJECT Gate 4 S&U: stated $STATED ≠ plug $PLUGGED (Δ=$DELTA) — fees silently absorbed"; exit 1; }
```

`S&U` のスポンサー以外のすべての行は、青い Assumptions 入力（目標 EBITDA、エントリーマルチプル、手数料率など）か派生数式のいずれか。Uses / Sources のハードコード数値は不可。

**ステップ2 — デットスケジュール（マルチトランシェ）。** トランシェ × 年で1行。列: `BeginningBalance` / `Mandatory amortization` / `Cash sweep` / `EndingBalance` / `AverageBalance` / `InterestExpense`。シニア TLB: 1% の強制償却＋余剰キャッシュはすべてスイープへ。Mezz: 償却0%、キャッシュペイの利息のみ。この例の行マップ（シニア TLB トランシェ、2年目の C 列）: `C4=Beginning Balance, C5=Mandatory Amort, C6=Ending Balance, C7=Cash Sweep, C8=Average Balance, C9=Interest Expense`。`CF!C20` = スイープ可能なフリーキャッシュ（CF シート上の2年目・スイープ前の期末現金）。レイアウトに応じてトランシェ行ブロックを差し替えること。

```bash
# 2年目 シニア TLB
cat <<'EOF' | officecli batch "$FILE"
[
  {"command":"set","path":"/Debt/C4","props":{"formula":"B6"}},
  {"command":"set","path":"/Debt/C5","props":{"formula":"-C4*Assumptions!$B$30","numberformat":"$#,##0;($#,##0);\"-\""}},
  {"command":"set","path":"/Debt/C6","props":{"formula":"C4+C5+C7","numberformat":"$#,##0;($#,##0);\"-\""}},
  {"command":"set","path":"/Debt/C7","props":{"formula":"-MIN(-CF!C20,C4+C5)"}},
  {"command":"set","path":"/Debt/C8","props":{"formula":"(C4+C6)/2","numberformat":"$#,##0;($#,##0);\"-\""}},
  {"command":"set","path":"/Debt/C9","props":{"formula":"-C8*Assumptions!$B$31","numberformat":"$#,##0;($#,##0);\"-\""}}
]
EOF
# スイープルールのコメントをクラシックコメントとして追加する（comment はセルのプロパティではなく、独立した --type comment）。
officecli add "$FILE" /Debt --type comment --prop ref=C7 --prop text='cash sweep capped at available cash and remaining tranche balance'
```

**リボルバー枠の上限。** 取引にリボルバートランシェがある場合、各期のリボルバー残高はコミットメント上限で制限される:
```
Revolver_Balance = MIN(Assumptions!RevolverCapacity, MAX(0, prior_revolver + draw − paydown))
```
`MIN(capacity, ...)` の外側がなければ、不足四半期に facility を静かに超過してドローしてしまう。

行インデックスは自分のレイアウトに合わせて調整する。各トランシェ（シニア / メザニン / リボルバー）と各年について繰り返す。

**ステップ3 — P&L（5年）＋ Debt からの利息。** P&L の利息行は Debt から取得する: `Interest = 'Debt'!TotalInterestRowY<N>`。これが**循環参照**を作る: Interest → NI → CF → Cash Sweep → Debt balance → Interest。

**書き込み順序に関する警告。** `calc.iterate=true` は_再計算_を制御するのであって、書き込みフェーズを制御するのではない。すでにリングを含むファイルにシート横断リングの締めくくりの脚を追加すると、`iterate` の値にかかわらずエンジンが CPU 100% でデッドロックする。複雑なリング（マルチトランシェ LBO、リボルバー＋TLB＋メザニン）では、以下の §Write-order surgery（デリング → 下流を書く → 再リング）を使う。リング数式を書く**前に** `calc.iterate=true` を有効化する:

```bash
officecli set "$FILE" / --prop calc.iterate=true --prop calc.iterateCount=100 --prop calc.iterateDelta=0.001
```

`iterate` は、自然に減衰するループ（利息が上がる→現金が減る→スイープが減る→残高が上がる、EBIT で有界）に対して逐次近似で収束する。`#REF!` や発散する値が出たら一旦止め、`iterateCount` を1000に上げるのではなく代数を修正すること。

**ステップ4 — CF＋キャッシュスイープ。** Ending cash = Opening + CFO − CapEx − Mandatory amort − Cash sweep。Cash sweep = `MIN(freeCashAfterCapEx, seniorDebtBalance + seniorMandatoryAmort)`。`MIN` の上限がゼロ未満へのスイープを防ぐ。

**ステップ5 — Exit＋Returns。** 行マップ: `Exit: B3=Exit EV, B4=Less: remaining debt, B5=Exit equity to sponsor`；`Returns: B3=MOIC, B4=IRR`。

```bash
# 値/数式 — 単一の non-resident batch。
cat <<'EOF' | officecli batch "$FILE"
[
  {"command":"set","path":"/Exit/B3","props":{"formula":"'P&L'!F8*Assumptions!$B$25","numberformat":"$#,##0;($#,##0);\"-\""}},
  {"command":"set","path":"/Exit/B4","props":{"formula":"-('Debt'!F6+'Debt'!F13)","font.color":"008000","numberformat":"$#,##0;($#,##0);\"-\""}},
  {"command":"set","path":"/Exit/B5","props":{"formula":"B3+B4","bold":"true","numberformat":"$#,##0;($#,##0);\"-\""}},
  {"command":"set","path":"/Returns/B3","props":{"formula":"'Exit'!B5/('S&U'!B9)","numberformat":"0.00\"x\""}},
  {"command":"set","path":"/Returns/B4","props":{"formula":"IRR({-'S&U'!B9,0,0,0,0,'Exit'!B5})","numberformat":"0.0%"}}
]
EOF
# クラシックコメント — アンカーセルごとに1つの --type comment。
officecli add "$FILE" /Exit --type comment --prop ref=B3 --prop text='Exit EV = Y5 EBITDA × exit multiple'
officecli add "$FILE" /Returns --type comment --prop ref=B3 --prop text='MOIC = exit equity / sponsor equity'
officecli add "$FILE" /Returns --type comment --prop ref=B4 --prop text='IRR — 5-yr, entry + exit only; use XIRR for mid-year dividends'
```

**コールアウト — ラベル: `comment` 要素 vs Notes 列 vs `formula`（3つの異なる仕組み）。**
- **ホバーツールチップ** → `officecli add ... --type comment --prop ref=<cell> --prop text='...'`。**`comment` キーは `set cell` の有効なプロパティではなく**（`officecli help xlsx cell` に存在しない）、`set cell` の props dict に埋め込んでもサイレントに無視される。専用の要素を使うこと。
- **隣接する Notes 列に表示するテキスト** → `{"command":"set","path":"/DCF/D3","props":{"value":"TV = FCF × (1+g) / (WACC−g)"}}` — **`formula` ではなく `value`**、単なる引用符付き文字列。
- **数式風の文章を実際の数式として書く** → 絶対にしない。`{"formula":"FCF10*(1+g)/(WACC-g)"}` は Excel で `#NAME?` になる（そのセルコンテキストでは `FCF10`、`g`、`WACC` は未バインドの識別子）。

年央配当や部分エグジットには、`IRR` の代わりに `XIRR({cashflows}, {dates})` を使う。

**ステップ6 — リターンウォーターフォール（任意、4段階 LP/GP）。** 段階: (1) LP 優先リターン 8%；(2) GP キャッチアップ 20% まで；(3) ハードル超過分の 80/20 分配；(4) 損失時は LP に100%。各段階は `MAX(0, MIN(...))` のクランプ。一般的なグリッドパターンは §Sensitivity & scenarios を参照。

**検証。**

```bash
officecli get "$FILE" /S&U/B12 --json | jq '.data.results[0].format.cachedValue // .data.results[0].text'   # BALANCED と表示されなければならない
officecli get "$FILE" /Returns/B3 --json | jq '.data.results[0].format.cachedValue'                # MOIC、典型的には2.0x〜4.0xを想定
officecli get "$FILE" /Returns/B4 --json | jq '.data.results[0].format.cachedValue'                # IRR、典型的には0.15〜0.30を想定
# Iterate は収束したか？
officecli query "$FILE" 'cell:contains("#REF!")' --json | jq '.data.results | length'   # 0でなければならない
```

## 感応度とシナリオ

**3パターンから1つ選ぶ:**
- **(a) Assumptions 上の Base / Upside / Downside 列** — 横並びのシナリオ、"Active" 列＋ `INDEX/MATCH` によるドロップダウンなしの切替。
- **(b) ドロップダウン＋ `INDEX/MATCH` スイッチ** — Summary 上の1つの入力規則ドロップダウンが `INDEX(Base:Downside, MATCH(Dropdown, ScenLabels, 0))` を通じてすべてのドライバーを駆動する。
- **(c) 2軸感応度グリッド** — 5×5 または 7×7、セルごとに独立した数式1つ、行/列見出しが2つのドライバー。WACC × g についてはレシピ B ステップ5を参照。

(a)+(b) を混在させると循環入力になる（シナリオがドロップダウンで選ばれつつ Active 列でも上書きされる）— どちらか一方を選ぶこと。

**グリッドのルール:** 各セルは、出力数式の独立したコピーに行ドライバーと列ドライバーを代入する。`WACC` 名前付き範囲（それはパネルのもの）は参照できず、グリッドの軸セルを参照する。

**ドロップダウンによるシナリオ切替。** Summary 上の1つの `validation` ドロップダウンが、すべての `Assumptions` 行を駆動する:

```bash
cat <<'EOF' | officecli batch "$FILE"
[
  {"command":"add","parent":"/Summary","type":"validation","props":{"sqref":"B1","type":"list","formula1":"Base,Upside,Downside"}},
  {"command":"set","path":"/Assumptions/B5","props":{"formula":"INDEX(C5:E5,MATCH(Summary!$B$1,$C$4:$E$4,0))"}}
]
EOF
# B5 にホバーツールチップを付けたい場合は、別途追加する:
officecli add "$FILE" /Assumptions --type comment --prop ref=B5 --prop text='Revenue growth — picked by Summary!B1 scenario dropdown'
```

すべての `Assumptions` ドライバー行に同じ `INDEX/MATCH` を設定する。C:E の Base / Upside / Downside 列は青のまま（ハードコードされたシナリオ入力）。

**フットボールフィールドチャートのパターン（DCF バリュエーションサマリー）。** 3〜5つのバリュエーション手法（DCF base、DCF bear、Trading comps、Precedent txns、LBO floor）について、横向きの Low→High バーを縦に積む。`Football` シート上: A列 = 手法ラベル、B列 = Low $、C列 = High $、D列 = `=C−B`（幅）。チャートは積み上げ横棒で、B列を不可視の最初の系列（白/塗りなし）、D列を可視の系列にする — `dataRange=Football!A3:D7`、`chartType=bar`。Excel はこれを手法ごとの浮遊バーとして読み込む。

## 財務関数パターン

簡潔なリファレンスであり、ファイナンスの教科書ではない。これらが何をするか分からない場合は、一旦止めてユーザーに確認すること。

| 関数 | 優先すべき対象 | 理由 |
|---|---|---|
| `XNPV(rate, values, dates)` | `NPV` | 不規則なキャッシュフロー日付（M&A クローズが年央、段階的トランシェ） |
| `XIRR(values, dates)` | `IRR` | 不規則な日付；複数回の符号変化をより適切に処理 |
| `INDEX(range, MATCH(lookup, key, 0))` | `VLOOKUP` | 挿入に強い（VLOOKUP はソース範囲に列が挿入されると壊れる） |
| `IFERROR(x/y, 0)` または `IF(y=0, 0, x/y)` | 素の除算 | 財務モデルではすべての `/` をガードする — `#DIV/0!` を出荷すると納品失敗 |
| `MIRR(values, financeRate, reinvestRate)` | 符号反転を伴う `IRR` | キャッシュフローパターンに符号変化が2回以上ある場合 |
| `SUMIFS(sumRange, criteriaRange1, criterion1, ...)` | 配列 `SUMPRODUCT((...))` | 条件付き合計として意図が明確；どちらも正しく評価される |

**エバリュエーターのカバレッジ — 検証すること、先回りしてハードコードしないこと。** CLI のエバリュエーターは、本スキルが使うファイナンス関数 — `NPV` / `XNPV` / `IRR` / `XIRR`、配列リテラル数式（`IRR({...})`）、`SUMPRODUCT(1/COUNTIF(range,range))` の重複除外カウント — を計算し、`evaluated:true` で正しい値をキャッシュする。`NPV→0` に書き換える義務も、重複除外カウントの `1/N` トラップもない。正直なルールは: 意図通りの数式を組み、読み戻しでキャッシュ値を検証すること（`get ... --json | jq '.data.results[0].format.cachedValue'`）。特定のセルが本当に `#OCLI_NOTEVAL!` センチネルを返す場合のみ（入力より前に書かれた数式、あるいは未サポートの関数）フォールバックする — まず `close` 後に non-resident で再 set し、それでも評価されなければ、計算済みの値を青字＋クラシックコメントでハードコードする（`officecli add "$FILE" /Sheet --type comment --prop ref=<cell> --prop text='cached valuation; refreshes on open in Excel — do not edit'`）、そして納品ノートで開示すること。

## 循環参照と反復計算

**`calc.iterate` を有効化するのは、循環性が代数的に正当化される場合のみ:** 利息⇄現金（LBO のリボルバー / キャッシュスイープ）、税シールド⇄NI（稀 — ほとんどの3ステートメントモデルは税引前に利息を計算し、これを回避する）、リボルバープラグ⇄期末現金（最低現金を伴う企業キャッシュウォーターフォール）。

```bash
officecli set "$FILE" / --prop calc.iterate=true --prop calc.iterateCount=100 --prop calc.iterateDelta=0.001
```

`iterateCount=100` / `iterateDelta=0.001` は Excel のデフォルトで、自然に減衰するループには問題ない。

### 書き込み順序サージェリー（デリング → 下流を書く → 再リング）

`calc.iterate` は再計算を制御するのであって、書き込みフェーズは制御しない。すでに配線済みのシート横断リング（Debt.Interest ⇄ CF.Cash ⇄ Debt.CashSweep）の締めくくりの脚を追加すると CPU 100% でデッドロックする；未収束のリングでは `view html` / `get` もハングする。

**3ステップの手順:**
1. **デリング** — Debt を書く際、リングの10〜20個のセルをリテラルの `0` にセットする（例: `C7=0`、`=-MIN(...)` ではなく）。リングを取り除く。
2. **下流を書く** — 非循環のすべての連鎖（P&L、CF、Exit、Returns、Summary、グリッド）を non-resident で、シートごとに1つのヒアドキュメントで構築する。すべてがゼロ化されたセルに対してキャッシュされる。
3. **再リング** — resident をすべて閉じ、各循環セルに実際の数式を再 set する、1セルにつき1回の `set`、non-resident で。

**受け入れ基準。** `get /Debt/C7 --json | jq '.data.results[0].format.cachedValue'` が非ゼロ非 null を返すこと。それでもデッドロックするセルがあれば、`=0` ＋クラシックコメント「circular; recalculates in Excel on F9」を残し、納品時にフラグを立てる。`iterateCount=1000` で覆い隠すことは絶対にしない。

**`#REF!` /発散値の応急処置として `iterate` を使わないこと。** `iterateCount` を1000に上げるとバグが隠れ、それらしく誤った値を出荷してしまう；`validate` はそれを検出しない。ループは代数的に断つこと（例: 平均残高ではなく期首残高に対する利息にする）。

**収束の検証。** ループのセルを読み、駆動する前提条件を上げてから戻し、再度読む — 値が一致しなければならない:

```bash
V1=$(officecli get "$FILE" /Debt/C9 --json | jq '.data.results[0].format.cachedValue')
officecli set "$FILE" /Assumptions/B31 --prop value=0.085
officecli set "$FILE" /Assumptions/B31 --prop value=0.0845
V2=$(officecli get "$FILE" /Debt/C9 --json | jq '.data.results[0].format.cachedValue')
[ "$V1" = "$V2" ] && echo "Iterate converged" || echo "WARN: drift V1=$V1 V2=$V2 — tighten iterateDelta or check algebra"
```

## 監査と Delivery Gate

**問題があると想定すること。** 最初のビルドが正しいことはほぼない。以下の全ゲートを実行し、すべてのチェックが成功メッセージを出力しなければならない。`validate` が通ることは納品を意味しない — モデルはスキーマを通過しても、10倍間違っていることがある。

### Gate 1〜3 — xlsx v2 から逐語的に継承

→ see xlsx v2 §QA minimum cycle（Gate 1〜3 は `view issues`、エラーセルの query、close 後の `validate` をカバーする）。xlsx v2 に書かれている通り、まずこれらを実行する。財務モデル固有の調整はない。

### Gate 4 — ステートメントの整合性（3ステートメント＆LBO）

レシピ A / C が生成するバランスチェックとキャッシュ照合の行は、すべての期間で `OK` / `BALANCED` を示さなければならない。チェック行に `query` を行い、`IMBALANCED` / `CF !=` が1つでもあれば拒否する:

```bash
BS_FAIL=$(officecli query "$FILE" 'cell:contains("IMBALANCED")' --json | jq '.data.results | length')
CF_FAIL=$(officecli query "$FILE" 'cell:contains("CF !=")' --json | jq '.data.results | length')
SU_FAIL=$(officecli query "$FILE" 'cell:contains("S&U IMBALANCE")' --json | jq '.data.results | length')
if [ "$BS_FAIL" -eq 0 ] && [ "$CF_FAIL" -eq 0 ] && [ "$SU_FAIL" -eq 0 ]; then
  echo "Gate 4 OK (balance + recon + S&U all pass)"
else
  echo "REJECT Gate 4: BS=$BS_FAIL CF=$CF_FAIL S&U=$SU_FAIL"; exit 1
fi
```

いずれかが失敗した場合、モデルはサイレントに間違っている — 納品前に上流の連鎖を修正すること。最も一般的な原因: シート横断数式が `\!`（シェルによって壊れた形）で保存されている — `officecli query "$FILE" 'cell:contains("\\\\!")'` を実行し、batch ヒアドキュメント経由で再入力する。

### Gate 5 — バリュエーションセルのキャッシュ値の健全性

NPV / IRR / XIRR / エクイティブリッジ / MOIC / サマリー KPI のセルが未評価のまま（典型的には入力より前に数式が書かれたことによる `#OCLI_NOTEVAL!` センチネル）出荷されると、開いても再計算しない読者に空欄/誤った数値を送ってしまう。すべてのバリュエーションセルを列挙し、キャッシュ値が存在することを確認する:

```bash
# パスのリストはレシピごとにカスタマイズする — これは DCF の例
for P in "/DCF/C4" "/DCF/C5" "/DCF/C6" "/DCF/C8" "/DCF/C9"; do
  V=$(officecli get "$FILE" "$P" --json | jq -r '.data.results[0].format.cachedValue // "null"')
  if [ "$V" = "null" ] || [ "${V#\#OCLI_NOTEVAL}" != "$V" ]; then
    echo "REJECT Gate 5: $P unevaluated (cached=$V) — re-set after close (see §Build-order & cache-drift)"; exit 1
  fi
  echo "Gate 5 $P: cached=$V OK"
done
```

LBO の場合は `/Exit/B5`、`/Returns/B3`、`/Returns/B4` をリストに追加する。3ステートメントの場合は `/Summary/B2:B5` を追加する。

### Gate 6 — ハードコード / ゾーンの規律

すべての Calc シートのハードコード数値はゼロ。実行可能な形:

```bash
# `cell:not(:has(formula))` はリテラルセルを選択し（`cell:has(formula)` は数式セルを選択する）。
HARDCODE=$(officecli query "$FILE" 'cell[type=Number]' --json \
  | jq '[.data.results[] | select(.format.formula == null) | select(.path | test("/(P&L|Balance Sheet|Cash Flow|DCF|Debt|FCF|WACC|Exit|Returns)/"))] | length')
[ "$HARDCODE" -eq 0 ] && echo "Gate 6 OK (no hardcodes on Calc sheets)" || { echo "REJECT Gate 6: $HARDCODE hardcoded numeric cells on Calc zone — move to Assumptions"; exit 1; }

# 名前付き範囲のカバレッジ＋無用の飾りの監査: ≥3個の範囲が宣言され、かつ各々が≥1個の数式から参照されている。
NR=$(officecli query "$FILE" namedrange --json | jq '.data.results | length')
[ "$NR" -ge 3 ] && echo "Gate 6 OK ($NR named ranges)" || echo "WARN Gate 6: only $NR named ranges"
DEAD=0
for NR_NAME in $(officecli query "$FILE" namedrange --json | jq -r '.data.results[].format.name'); do
  # 数式の「ソース」(`formula~=`) にマッチさせ、キャッシュ/表示された「結果」ではないこと — `:contains` だと
  # 計算値をスキャンしてしまい、すべての名前を dead と誤って報告する（偽の却下）。
  USES=$(officecli query "$FILE" "cell[formula~=$NR_NAME]" --json | jq '.data.results | length')
  [ "$USES" -ge 1 ] && echo "  $NR_NAME: $USES uses OK" || { echo "  WARN: $NR_NAME unused"; DEAD=$((DEAD+1)); }
done
[ "$DEAD" -eq 0 ] && echo "Gate 6 named-range audit OK" || { echo "REJECT Gate 6: $DEAD dead-decoration name(s)"; exit 1; }
```

### Gate 5b — HTML プレビューによるビジュアル監査（必須）

Gate 1〜4/6 は grep による防御に過ぎず、レンダリングされたシートを見ることはできない。`officecli view "$FILE" html` を実行し、返却された HTML を Read すること。すべてのシートを歩いて確認する（xlsx v2 のビジュアルフロアを継承）:

- 数値セルに `###` がないこと（列幅を広げる）。
- ラベル / セクションヘッダーが省略されていないこと（列幅を広げるか `alignment.wrapText=true`）。
- プレースホルダートークン（`TBD`、`{var}`、`xxxx`）がないこと — 下記の Gate 6.1 grep。
- バランスチェック / 照合行が、期間列すべてで `OK` / `BALANCED` を示していること。
- ダッシュボードのチャートがレンダリングされ、ARR/revenue 系列の y 軸が 0 起点で、ソースデータがステートメントシートと一致していること。
- 感応度グリッドの色が緑（アップサイド）→赤（ダウンサイド）で読めること — カラースケール CF が適用されていること。
- サマリー KPI に古いキャッシュ `0` が残っていないこと；残っていればキャッシュ再更新パスを実行する。

欠陥があれば REJECT。**人による目視プレビュー:** `officecli watch "$FILE"`、または Excel / WPS / Numbers で開く — 最終的な色とチャートの忠実度は対象ビューアでのみ完全にレンダリングされる。

### Gate 6.1 — トークン / プレースホルダーの一斉検査

```bash
LEAK=$(officecli view "$FILE" text | grep -niE 'TBD|\(fill in\)|xxxx|lorem|\{\{|placeholder|coming soon')
[ -z "$LEAK" ] && echo "Gate 6.1 OK (no placeholder tokens)" || { echo "REJECT Gate 6.1:"; echo "$LEAK"; exit 1; }
```

### 正直な限界

`validate` はスキーマエラーを検出するのであって、ファイナンスのエラーは検出しない。モデルは、`BS.Cash` がバランスを取るためにハードコードされていても、`NPV` が `0` でキャッシュされていても、感応度グリッドが FCF の前に構築されたためオール0になっていても、引用符なしの参照を持つ `P&L` という名前のシートで `#NAME?` の実行時エラーが起きていても、`validate` を通過してしまう。Gate 4 / 5 / 6 / 5b が存在するのは、スキーマレベルの `validate` がこれらを一切検出できないからである。

## Known Issues & Pitfalls

→ ベースの落とし穴（シート横断の `!` トラップ、batch JSON のドット付き名前ルール、レンダラーの注意点）: see xlsx v2 §Known Issues & Pitfalls — すべて適用される。

財務モデル固有:

- **COGS 上の AP の符号。** Accounts Payable: P&L 上で COGS が負数で保存されている場合、AP の数式は符号を反転させる必要がある — `=-COGS*DaysPayable/365`。符号を誤ると NWC が水増しされ、CF の向きが反転する。サイレントに `validate` を通過してしまう。
- **`#NAME?` は `query` / `validate` では検出されない。** シート名をクォートせずに `P&L!B3` を参照するシート横断数式（`&` が特殊文字であるため）は、実行時に `#NAME?` になる。シート横断参照は常に `'P&L'!B3` のように書くこと — シート名に `&`、スペース、`(`、`)` などが含まれる場合はシングルクォートで囲む。Gate 5b のビジュアルチェックのみが検出できる。
- **反復計算のサイレントな未収束。** `calc.iterate=true iterateCount=100` は、真の答えがその2倍であっても、上限に達した時点の値で収束する。常に収束検証を実行すること（§循環参照）。複雑な LBO リング（マルチトランシェのデット＋スイープ＋税シールド）は収束しない場合がある；リングのセルで `cachedValue=0` の場合は §Write-order surgery を使う。
- **循環的な書き込みでの batch-while-resident デッドロック。** resident を開いたまま `batch` でシート横断リングの締めくくりの脚を書くと、CPU 100% でデッドロックする。リングのセルに対する単一の `set` でさえハングすることがある。対処: resident を閉じ、§Write-order surgery の通り2パスでリングを書く。non-resident の単一ヒアドキュメントのみが安全な形である。
- **`view html` でのシート横断キャッシュ値の陳腐化。** 上流と同じシーケンスで書かれた下流は `0` をキャッシュする。Excel は開いたときに解決するが、HTML プレビューはしない。連鎖の後、下流すべてを non-resident で再 set すること（§ビルド順序とキャッシュドリフト）。
- **`NPV()` / `XNPV()` は正しく評価・キャッシュされる。** エバリュエーターは両方（同一シート・シート横断とも）を計算し、書き換えは不要。`SUMPRODUCT(values/(1+rate)^periods)` は依然として監査可読性のためのオプションの代替であり、キャッシュの回避策ではない。
- **感応度グリッドのビルド順序は依然として重要。** FCF/WACC の入力が存在する前に書かれたグリッドセルは空の入力に対して評価され、誤った/空欄の値をキャッシュしてしまう可能性がある。まず FCF＋WACC＋DCF を構築し、次にグリッドを別の non-resident batch で構築し、検証する（`jq '.data.results[0].format.cachedValue'`）。これは順序の規律の問題であり、エバリュエーターの限界ではない。
- **`BS.Cash` は常に CF の期末現金と等しい**（Y1 を含む: `BS.Cash = 'Cash Flow'!B19`）。独立したプラグや Assumptions 参照には決してしない — プラグされた `BS.Cash` はバランスエラーを隠してしまう。
- **Year 2 以降の `Opening Cash` = 前期の `Ending Cash`**（`C17=B19`、`D17=C19`）。独立した Y2以降の opening-cash 入力は BS からサイレントにずれていく。
- **ウォーターフォールチャートの「合計」バー。** `chartType=waterfall` はプログラム的に合計をマークできない — `colors=` の慣習を使う（濃色＝合計、中間色＝プラス、赤＝マイナス）。`help xlsx chart` を参照。
- **`SharesOut` が数式の場合の DCF 一株当たり値。** `=BasicShares + OptionPool × ExerciseAssumption` → 青字の前提条件セルを追加し、`SharesOut` 名前付き範囲を生の入力ではなく計算済みセルに向ける。

## ヘルプへのポインタ

迷ったときは: `officecli help xlsx [element] [--json]`。ヘルプが権威あるスキーマであり、本スキルは財務モデリングの差分に関する意思決定ガイドである。
