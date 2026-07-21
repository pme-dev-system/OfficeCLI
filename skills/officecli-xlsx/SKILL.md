---
name: officecli-xlsx
description: "このスキルは .xlsx ファイルが関わる場面 -- 入力・出力いずれか、または両方 -- で常に使用する。具体的には次のような場合が含まれる: スプレッドシート、財務モデル、ダッシュボード、トラッカーの作成; 任意の .xlsx ファイルからのデータの読み取り・パース・抽出; 既存ワークブックの編集・変更・更新; 数式、チャート、ピボットテーブル、テンプレートの操作; CSV/TSV データの Excel 形式へのインポート。ユーザーが「スプレッドシート」「ワークブック」「Excel」「財務モデル」「トラッカー」「ダッシュボード」に言及した場合、または .xlsx/.csv のファイル名を参照した場合にトリガーする。"
---

# OfficeCLI XLSX スキル

## セットアップ

`officecli` が見つからない場合:

- **macOS / Linux**: `curl -fsSL https://d.officecli.ai/install.sh | bash`
- **Windows (PowerShell)**: `irm https://d.officecli.ai/install.ps1 | iex`

`officecli --version` で確認する（PATH が反映されない場合は新しいターミナルを開く）。インストールに失敗した場合は https://github.com/iOfficeAI/OfficeCLI/releases からバイナリをダウンロードする。

## ⚠️ ヘルプ最優先ルール

**このスキルは「良い xlsx とは何か」を教えるものであり、すべてのコマンドフラグを網羅するものではない。プロパティ名、enum 値、エイリアスが不確かな場合は、推測する前に必ずヘルプを確認すること。**

```bash
officecli help xlsx                         # xlsx の全要素を一覧表示
officecli help xlsx <element>               # 要素の完全なスキーマ（例: pivottable、chart、cf）
officecli help xlsx <verb> <element>        # 動詞スコープ（例: add chart、set cell）
officecli help xlsx <element> --json        # 機械可読なスキーマ
```

ヘルプはインストール済みの CLI バージョンを反映する。このスキルとヘルプの内容が食い違う場合は、**ヘルプが正**である。

## シェルと実行の規律

**シェルクォーティング（zsh / bash）。** Excel のパスには `[]` が含まれ、数値書式には `$` が含まれる。どちらもシェルのメタ文字である。ルール:

- 要素パスは常にクォートする: `"/Sheet1/row[1]"` であって `/Sheet1/row[1]` ではない。
- `$` を含むプロパティ値には**シングルクォート**を使う: `numFmt='$#,##0'`。
- シート間参照の `!` を含む数式は、`batch` と `<<'EOF'` ヒアドキュメントを使う（後述の「既知の問題」参照）。
- プロパティ値中の `\n` と `\t` は CLI によって解釈される —— `\n` はセル内の実改行になり（`--prop wrapText=true` と組み合わせる）、`\t` はタブになる —— これは xlsx / docx / pptx で共通の挙動である。リテラルなバックスラッシュ + n を得たい場合（めったにないが）は二重にする（`\\n`）。（`$` は上位のシェル層の話であり、シングルクォートで対処する。）

**インクリメンタルな実行。** コマンドは1つずつ実行し、都度終了コードを確認する。`officecli` は呼び出しごとにファイルを変更するため、3番目のコマンドで失敗する50コマンドのスクリプトは、失敗が静かに連鎖する。1コマンド実行 → 出力確認 → 次へ進む、を徹底する。

## 成果物への要件

コマンドに手を出す前に、良い xlsx がどのようなものかを把握しておく。以下は、納品するすべてのワークブックが満たさなければならない品質基準である。

### すべての Excel ファイル

**数式エラーゼロ。** 納品するワークブックは `#REF!`、`#DIV/0!`、`#VALUE!`、`#NAME?`、`#N/A` を一つも含んではならない。例外なし —— 分母は `IFERROR` または `IF(x=0,...)` でガードする。

**数式であり、ハードコードされた値ではない。** 他のセルから計算できる数値は数式にする。`=SUM(B2:B9)` であるべき箇所に `5000` をハードコードすると、入力が変わってもワークブックが「生きたまま」であるという契約が壊れる。これはこのスキルにおける最も重要なルールである。

**プロフェッショナルなフォント。** ワークブック全体で一貫した、プロフェッショナルなフォント（Arial / Calibri / Times New Roman）を使う。あるシートが CSV 由来だからといって、4種類ものフォントを混在させない。

**明示的な幅。** 自動フィットは存在しない。ユーザーが読むすべての列には `width` を設定しなければならない —— デフォルトの 8.43 文字幅ではすべてが見切れる。妥当な初期値: ラベル 20〜25、数値 12〜15、日付 12、短いコード 8〜10。

**既存のテンプレートを尊重する。** 既にルック＆フィールが確立されているファイルを編集する場合は、それに合わせる。既存の慣習はこのガイドラインより優先される。

### 視覚的な納品基準（すべてのワークブックに適用）

完了を宣言する前に、`officecli view "$FILE" html` を実行し、返された HTML を Read して以下をすべて確認する:

- **どのセルにも `###` がないこと。** `###` は列がその中の最大値を表示するには狭すぎることを意味する。ユーザーが読むすべての列には明示的な `width` が必要である。納品ファイルに `###` が残っているのは未完成の仕事であり、「些細な見た目の問題」では決してない。
- **タイトルが切れていないこと。** シートタイトル、セクション見出し、長いラベルはすべて収まっていなければならない。列を広げるか、セルに `wrapText=true` を適用する。
- **プレースホルダートークンがデータとして表示されていないこと。** `$fy$24`、`{var}`、`<TODO>`、`xxxx` がセル、チャートタイトル、シリーズ名、凡例に出現してはならない。これらは置換されずに残ったビルド時トークンである。
- **円グラフ／ドーナツグラフのスライスが個別の塗り色を持っていること。** スライスが同色でレンダリングされる場合は、`bar` / `column` に切り替えるか、`colors=...` を明示的に設定する。
- **末尾に空のページや空のチャートアンカーがないこと。** 空のソースセル上に `anchor=D2:J18` を置くと、壊れたチャートのように見える。

上記のいずれかに該当する場合は、完了を宣言する前に立ち止まって修正すること。

**印刷レイアウト。** ユーザーが印刷したり役員向け資料として送付したりする可能性のあるシートには、ページ設定が必要である。デフォルトの縦向き＋改ページなしのままだと、横長の表やチャートが途中で分割される。シートの形状に応じて fit モードを選ぶ:

```bash
# サマリー／チャート／ダッシュボードシート（小さめ、行数 40 程度以下）: 1ページに収める。
officecli set "$FILE" "/Summary" --prop orientation=landscape --prop fitToPage=true
# 縦に長いデータ表（数十行以上）: 幅だけを1ページに収め、高さは自然にページ送りさせる。
# ここで fitToPage=true にすると全行が1ページに押し込まれ、読めなくなる（### 化した日付、5px の行高）。
officecli set "$FILE" "/Data" --prop orientation=landscape --prop fitToPage=1x0
```

`fitToPage=true` は `1x1`、つまり両軸を1ページに収める指定と同じである —— これはシートがすでに短い場合にのみ正しい。`1x0` は幅1ページ・高さ無制限を意味する。トリガー条件: シートがチャートを含む、8列を超える、またはユーザーの依頼に印刷／役員会／投資家という言葉が含まれる場合。

### 財務モデル限定 —— テンプレート、トラッカー、CSV インポート、業務用シートを作る場合はこの節をスキップしてよい

対象範囲: 予算、予測、3表連動モデル、バリュエーション、`$` が多用される分析系ワークブック全般。カスタマーサポート用トラッカーやオンボーディングテンプレートにはこの節は不要である。

**カラーコーディング —— 業界標準。** 5つの基本色を、装飾ではなく言語として使う。レビュアーは数式を読む前に、色だけを見てそのセルが何であるかがわかるべきである。

| 色 | 役割 | 例 |
|---|---|---|
| 青字 `0000FF` | ハードコードされた入力値、シナリオ変数 | `font.color=0000FF` |
| 黒字 `000000` | すべての数式・計算 | デフォルト |
| 緑字 `008000` | このワークブック内のシート間リンク | `font.color=008000` |
| 赤字 `FF0000` | 外部ファイル／外部ワークブックへのリンク | `font.color=FF0000` |
| 黄色塗り `FFFF00` | レビューが必要な重要な前提条件 | `fill=FFFF00` |

レビュアーは数式を読む前に、色だけを見てそのセルが何であるかがわかるべきである。これは見た目の好みではなく、コミュニケーション上の約束事である。

**数値書式 —— 好みではなく標準。**

- **年**はテキストであり、数値ではない。`2,026` ではなく `2026` と表示する —— `numFmt="@"` を使うか `type=string` を設定する。
- **通貨**は単位をヘッダーに持たせる（`Revenue ($mm)`）。セルごとに繰り返さない。
- **ゼロは `-` として表示**し、`0` としては表示しない。`$#,##0;($#,##0);"-"` を使う。
- **パーセンテージ**はデフォルトで小数点以下1桁とする: `0.0%`。
- **負の数は括弧を使う**: `-1,234` ではなく `(1,234)`。
- **バリュエーション倍率**は `0.0x` 形式を使う（EV/EBITDA、P/E など）。

**前提条件はセルに置き、数式の中に埋め込まない。** `=B5*(1+$B$6)` は正しく、`=B5*1.05` はバグである。青色でハードコードされた各入力値には、隣のセルまたはセルコメントに出典を明記する:

```
Source: Company 10-K, FY2024, Page 45, Revenue Note
Source: Bloomberg, 2026-05-02, AAPL US Equity
Source: Management guidance, Q2 2026 earnings call
```

出典のないハードコードされた数値は、未文書化の前提条件である —— レビュアーはそれを監査できない。

## 共通ワークフロー

6ステップ。ある程度の規模を持つビルドはすべてこの形をたどる。

1. **open/save ライフサイクル。** 最初に `officecli open <file>`、最後に `officecli save <file>` を実行してディスクにフラッシュする —— `save` は書き込みのみを行い、常駐（resident）は後続の編集のためにウォームな状態のまま残す；`officecli close <file>` は、ワンショットの引き渡しで常駐を解放したい場合にのみ使う。どちらも常に安全である（エラーになったり作業が失われたりすることはない）。多数のセルを扱う場合は `batch` を使う: **1ブロックあたり 50 操作以下を推奨。純粋な値設定ペイロードでは1ブロック80件以上でもテスト済みで失敗ゼロ。シート間数式のバッチはその例外であり、非常駐・単一ヒアドキュメントで実行する（後述の「既知の問題」参照）**。**フラッシュは officecli の外側の境界でのみ行う。** officecli 自身の読み取りは常にあなたの編集を反映するので、`save`/`close` は officecli 以外のプログラム（openpyxl/pandas、Excel、レンダラー、納品）がファイルを読む直前にのみ実行する。
2. **作成またはロード。** `officecli create "$FILE"`（新規）または `officecli view "$FILE" outline`（既存 —— まず全体像を把握する）。
3. **段階的に構築する。** 1コマンド実行 → 出力確認 → 次へ。構造的な操作（新しいシート、チャート、名前付き範囲、ピボット）の後は必ず `get` でその形を確認してから、その上に積み重ねる。
4. **書式設定。** 列幅、数値書式、ウィンドウ枠の固定、タブの色、ヘッダーの塗りつぶし。書式は任意の仕上げ作業ではなく、「成果物への要件」に従って成果物の一部である。
5. **保存し、それからキャッシュと向き合う。** `officecli save <file>` はディスクに書き込む。新規に追加された数式はキャッシュされた値を持たずに出荷される。人間がスプレッドシートアプリでファイルを開くと、アプリが再計算してキャッシュを埋める。**しかし、下流の `INDEX/MATCH`、`SUMPRODUCT`、あるいは上流の数式を参照するどの数式も、書き込み時に上流がキャッシュしていた値 —— しばしば `0` や古い値 —— をそのままキャッシュしてしまい、そのキャッシュされた嘘が再計算しないリーダーにまで残ってしまう。** 配列数式（`SUMPRODUCT`、動的条件付きの `SUMIFS`）やシート間の連鎖を含む複数数式のビルドの後は、**下流のすべてのセルを再タッチ**する（同じ数式で `set` を再実行する）ことで、新しくキャッシュされた上流の値からエンジンにキャッシュを再計算させる。⚠️ シート間の連鎖に対する常駐経由の再タッチは信頼性が低い（「バッチ／常駐の注意点」参照） —— 再タッチのパスには非常駐の `set` を優先する。その後 `officecli get` で下流のセルをいくつか確認し、`cachedValue=` が妥当であることを目視で確認する。`validate` は常駐が開いた状態でも安全であり、それ自体が保留中の編集をディスクにフラッシュする（docx / pptx と同様）。
6. **QA —— 問題があることを前提とする。** QA の節を参照。最後のコマンドが 0 で終了したから完了なのではなく、修正と検証のサイクルを1回まわして新たな問題がゼロになって初めて完了である。

## クイックスタート

最小限の xlsx: 3か月分の売上 + 合計の数式 + 列幅 + 通貨書式。そのままコピペせず、自分のファイル・自分のデータに合わせて調整すること。

```bash
officecli create "$FILE"
officecli open "$FILE"
officecli set "$FILE" /Sheet1/A1 --prop value=Month --prop bold=true
officecli set "$FILE" /Sheet1/B1 --prop value=Revenue --prop bold=true
officecli set "$FILE" /Sheet1/A2 --prop value=Jan
officecli set "$FILE" /Sheet1/A3 --prop value=Feb
officecli set "$FILE" /Sheet1/A4 --prop value=Mar
officecli set "$FILE" /Sheet1/B2 --prop value=42000 --prop numFmt='$#,##0'
officecli set "$FILE" /Sheet1/B3 --prop value=45000 --prop numFmt='$#,##0'
officecli set "$FILE" /Sheet1/B4 --prop value=48000 --prop numFmt='$#,##0'
officecli set "$FILE" /Sheet1/A5 --prop value=Total --prop bold=true
officecli set "$FILE" /Sheet1/B5 --prop formula="SUM(B2:B4)" --prop bold=true --prop numFmt='$#,##0'
officecli set "$FILE" "/Sheet1/col[A]" --prop width=12
officecli set "$FILE" "/Sheet1/col[B]" --prop width=15
officecli close "$FILE"
officecli validate "$FILE"
```

確認済み: `validate` は `no errors found` を返し、`B5` は `135000` に解決される。これがすべてのビルドの形である: open → セル／数式を set → 書式設定 → close → validate。

## CSV／一括インポート

**ネイティブな `import` コマンド（CSV/TSV には推奨）。** 最速の経路。CSV を一度の呼び出しでシートに読み込む。`--header` は AutoFilter とウィンドウ枠の固定を1行目に設定する。幅と `numFmt` は依然として後続のパスが必要（Dashboard スキルの D-12 参照）。

```bash
officecli import "$FILE" /Sheet1 --file data.csv --header
officecli import "$FILE" /Sheet1 --file data.tsv --format tsv --header
officecli import "$FILE" /Sheet1 --stdin --start-cell B2 < data.csv
```

**Python + batch のフォールバック** —— カスタムの型変換、数式の注入、または CSV が別のデータパイプラインの一部である場合に使う。600〜6000 以上のセルに対するレシピ:

```python
# gen_batch.py —— 1チャンクあたり80件の value-set 操作からなる batch チャンクを生成する
import csv, json
ops = []
with open("data.csv") as f:
    reader = csv.reader(f)
    for r, row in enumerate(reader, start=1):
        for c, val in enumerate(row):
            col = chr(ord('A') + c)
            ops.append({"command":"set","path":f"/Data/{col}{r}",
                        "props":{"value": val}})
for i in range(0, len(ops), 80):
    print(json.dumps(ops[i:i+80]))
```

```bash
python gen_batch.py | while IFS= read -r chunk; do
  printf '%s\n' "$chunk" | officecli batch "$FILE"
done
```

結果: 648行の小売 CSV（6490セル）が約30秒でロードされ、失敗ゼロ。チューニング: 1チャンク80操作から開始し、失敗するチャンクがあれば40に落とす。数値の型推論や数式は、この後にターゲットを絞った `set` で追加する —— このレシピにおける batch は純粋な値の注入である。

## 読み取りと分析

広く始めて、それから絞り込む。まず `outline` でどのシートが存在し、データがどこにあるかを把握し、場所がわかってから初めて `view` / `get` / `query` に踏み込む。

**レンダリングされたワークブックを開いて自分の作業を目視確認する。**
- `officecli view $FILE html` —— 返された HTML を Read してレンダリング結果を監査する。各シートはアドレス可能で、チャートはインラインでレンダリングされる。`###`、プレースホルダーの漏出、ピボットのレイアウト、行の高さの切れなどを検出できる。
- `officecli watch $FILE` は人間のユーザー向けにライブプレビューを立ち上げ続ける —— ユーザーが自分の判断で開く。ユーザーが一緒に見たい場合に使う。エージェント自身のセルフチェックには上記の `view html` を使う。
バッチ編集の後の**最初の視覚チェック**として `view html` を使い、根本原因を修正する。最終的な視覚確認は、ユーザー自身が Excel / WPS / Numbers ビューアで `.xlsx` を開いて行う。

**方向づけ。** シート、次元、数式の件数。

```bash
officecli view "$FILE" outline
```

**抽出。** コンテンツ QA や LLM のコンテキスト用のプレーンテキストダンプ。大きなファイルには `--start` / `--end` / `--cols` で範囲を絞る。

```bash
officecli view "$FILE" text --start 1 --end 50 --cols A,B,C
```

知っておくと役立つその他の `view` モード: `annotated`（セル値＋型／数式＋警告）、`stats`（数値のサマリー）、`issues`（壊れた数式、空のシート、欠落した参照）。

**ラウンドトリップダンプ。** `officecli dump "$FILE" [path]` はワークブック —— または1つのワークシート（`/Sheet1`、`/sheet[N]`）—— を再生可能な batch JSON にシリアライズする。`officecli batch new.xlsx --input dump.json` で再生する。既存ワークブックの構造から学んだり、生の OOXML を読む代わりにテンプレートを複製・調整したりする用途に使う。カバレッジは `dump --help` を参照。サブツリーのダンプにはワークブックレベルのリソース（設定、名前付き範囲）は含まれない —— 再生先はそれらをあらかじめ定義しておく必要がある。

```bash
officecli dump "$FILE" -o blueprint.json            # ワークブック全体
officecli dump "$FILE" /Sheet1 -o sheet.json        # 1つのワークシート
officecli batch new.xlsx --input blueprint.json
```

**単一の要素を確認する。** XPath 風のパスを使う。シェルは `[N]` をグロブとして展開するので、常にクォートする。

```bash
officecli get "$FILE" "/Sheet1/A1"            # 単一セル
officecli get "$FILE" "/Sheet1/A1:D10"        # 範囲
officecli get "$FILE" "/Sheet1/chart[1]"      # チャート
officecli get "$FILE" "/Sheet1/table[1]"      # ListObject
officecli get "$FILE" "/namedrange[1]"        # ワークブックレベルの名前付き範囲
```

子要素を展開するには `--depth N` を、機械可読な出力には `--json` を追加する。全要素の一覧は `officecli help xlsx`。

**ワークブック全体をクエリする。** CSS ライクなセレクタ。1件ずつ手でたどるのではなく、体系的なチェック（数式のカバレッジ、エラーセル、空のヘッダー）に使う。

```bash
officecli query "$FILE" 'cell:has(formula)'       # すべての数式セル
officecli query "$FILE" 'cell:contains("#REF!")'  # 壊れた参照
officecli query "$FILE" 'cell[type=Number]'       # 型によるフィルタ
officecli query "$FILE" 'Sheet1!B[value!=0]'      # シートスコープ
```

演算子: `=`、`!=`、`~=`（含む）、`>=`、`<=`、`[attr]`（存在確認）。

**マージセルのショートカット。** `officecli query $FILE merge` または `mergedrange` —— どちらも `mergeCell` のエイリアスである。`<mergeCell>` の各エントリを手でたどることなく、ワークブック内のすべての結合範囲を返す。

**行を一つずつたどるのが非効率なほどデータが大きい場合**は、Excel 自身の分析用要素に頼る:

- `officecli add`（`--type pivottable`）で**ピボットテーブル**を構築し、20個の SUMIF を書かずにグループ化・集計する。**スライサー**（`--type slicer`）を付ければ、読み手にフィルタ用の UI を提供できる。
- 行ごとのトレンドを示すには**スパークライン**（`--type sparkline`）を挿入する —— 行ごとに1本の折れ線グラフを作るよりも安価で、印刷時にもインラインで表示される。`type` は厳格な enum である: **`line | column | stacked`**（加えてエイリアス `winloss` / `win-loss` → `stacked`）。不正な `type=` の値は完全に失敗する —— もはや `line` へのサイレントなフォールバックはない。
- 正確なプロパティ名は `officecli help xlsx pivottable`、`officecli help xlsx slicer`、`officecli help xlsx sparkline` で確認する。

## 作成と編集

ビルドの9割はセル、数式、書式設定、そして1〜2個のチャートである。動詞は: `add`（新規要素）、`set`（プロパティ変更）、`remove`、`move`、`swap`、`batch`。

### セルと数式

値とその書式を1回の呼び出しで設定する。数式の先頭に `=` を書いてはならない —— CLI がそれを除去する。

```bash
officecli set "$FILE" /Sheet1/B5 --prop formula="SUM(B2:B4)" --prop numFmt='$#,##0'
officecli set "$FILE" /Sheet1/C5 --prop formula="B5/A5" --prop numFmt="0.0%"
```

構造的なプロパティ（幅、高さ、ウィンドウ枠の固定、タブの色）は行／列／シートのノードに存在する:

```bash
officecli set "$FILE" "/Sheet1/col[A]" --prop width=20
officecli set "$FILE" "/Sheet1/row[1]" --prop height=22
officecli set "$FILE" "/Sheet1" --prop freeze=A2 --prop tabColor=1F4E79
```

### 名前付き範囲

数式中の `$B$6` よりも名前付き範囲を優先する。それらは自己文書化されており（`GrowthRate` は `$B$6` よりわかりやすい）、数式を壊さずに前提セルを移動できる。`ref` の値には `!` と `$` の両方が含まれるため、batch ヒアドキュメント経由で追加する:

```bash
cat <<'EOF' | officecli batch "$FILE"
[
  {"command":"add","parent":"/","type":"namedrange","props":{"name":"GrowthRate","ref":"Sheet1!$B$6"}}
]
EOF
```

完全なスキーマは `officecli help xlsx namedrange` を参照。

**batch の JSON はシェルのエイリアスを受け付けない。** batch の `props` の内部では、常にフルのドット付き名を使う —— `"font.color": "FF0000"`、`"font.size": 14` であり、`"color": "FF0000"`（テキストか塗りか曖昧）は使わない。素のセル上でも、シェル形式は拒否される: `--prop color=1F4E79` は `ambiguous in cell context — use 'font.color' (text) or 'fill' (bg)` というエラーになる。ルール: batch の JSON でもセルのプロパティでも、`font.color` / `fill` を明示的に書く。`parent` はワークブックレベルの要素では `"/"`、シートスコープの要素では `"/SheetName"` とする。空文字列はこれと同等ではない。

### チャート

チャートの種類は `officecli help xlsx chart` にある —— enum は長い（20種類以上）。伝えたいメッセージに合ったものを選ぶ: カテゴリ比較には column、時系列には line、円グラフはスライスの割合が自明な場合のみ、相関には scatter。特定の問いに答える場合を除き、特殊なタイプは避ける。

**チャートデータを供給する3つの方法。1チャートにつき1つを選ぶこと —— 作成時に複数を混在させるのはよくある落とし穴である。**

| 方法 | 形 | 使いどころ |
|---|---|---|
| (a) インライン `data` | `--prop data="Sales:100,200,300" --prop categories="Jan,Feb,Mar"` | 編集しない小さなデモ用チャート。真実の源はセルではなくチャート XML の中にある。 |
| (b) 2次元 `dataRange` | `--prop dataRange="Sheet1!A1:B4"`（最初の列がカテゴリ、最初の行がヘッダー／シリーズ名） | 通常のケース。**2次元**でなければならない —— 単一列は "Chart requires data" で失敗する。 |
| (c) ドット区切りのシリーズごと指定 | `--prop series1.name=Sales --prop series1.values="Sheet1!B2:B4" --prop series1.categories="Sheet1!A2:A4"` | 各シリーズが連続しない範囲を指す複数シリーズのチャート、または明示的なシリーズ名が必要な場合。`series1.values` のみ（`categories` なし）だと、`1,2,3` を x 軸としたチャートになる。 |

**単一列の落とし穴。** `dataRange="Sheet1!B2:B13"` は「値の列」に見えるが、エンジンはこれを `Chart requires data` として拒否する。カテゴリ列を含むよう範囲を広げる（`A2:B13`）か、明示的な `series1.categories` を使う方法 (c) に切り替える。

**作成後にチャートを移動／リサイズする:** `set chart[N] --prop anchor="F5:N25"`（`--prop x= --prop y= --prop width= --prop height=` も使える）。**シリーズは依然として不変**である —— シリーズを追加・変更するには、`officecli remove` でチャートを削除し、`officecli add` で完全なシリーズリストを指定して再作成する。`remove chart[1]` は `chart[2] → chart[1]` のようにインデックスをシフトさせ、再追加は**末尾に追加される**点に注意 —— チャートの順序を維持するには、すべて削除してから順番通りに再構築する。

**アンカーのサイズ調整。** 自動フィットは存在しない。5〜6カテゴリ×2シリーズの column チャートには、すべてのラベルが切れずに表示されるようにおおよそ `A5:L22`（12列×18行）が必要である。狭すぎると X 軸ラベルが切れ、広すぎると印刷／エクスポート時にチャートがページをまたいで分割されることがある。迷ったら狭めに始め、`view html` でプレビューし（返された HTML パスを Read する）、段階的に広げる。ページレイアウト（後述）はこの修正のもう半分である。

**チャートの `dataRange` —— 常にシート名を接頭辞につける。** チャートが同じシート上にある場合でも、`A17:C22` ではなく `dataRange="Summary!A17:C22"` と書く。シート名なしの形式は動作が不安定である。接頭辞付きの形式は100%信頼できる。

officecli は、従来の Excel オブジェクトモデルにはない拡張チャートタイプを追加している: `boxWhisker`、`waterfall`、`funnel`、`histogram`、`treemap`、`sunburst`、`pareto`。データがそれを求めるときはこれらを使う。

**チャートのタイトル／シリーズ名／凡例／軸タイトルに未置換のテンプレートトークンを絶対に残さないこと。** `$fy$24`、`{var}`、`<TODO>`、`$VAR`、`{{placeholder}}` は凡例に**そのまま**レンダリングされる —— validate は通過するが、CFO には「FY2024」と表示されるべき場所に `$fy$24` が見える。常に確定したテキストかセル参照に紐づける（`title="FY2024 Revenue"` または `series1.name="Sheet1!A1"`）。

### 条件付き書式

一般的な3つのフレーバー。それぞれ独自のプロパティ形状を持つ（`officecli help xlsx cf` を参照）:

- **カラースケール**: 値に応じてグラデーションで色付けされたセル —— `minColor` / `midColor` / `maxColor` を伴う `type=colorscale`。
- **データバー**: 大きさを示すセル内バー —— `type=databar`。列全体で一貫したスケーリングをするために明示的な `min` / `max` を設定する。省略した場合のデフォルトも有効である。
- **数式ルール**（`formulacf` 要素）: 条件が真のとき行をハイライトする —— `formula="$C2>1000"` を伴う `type=formula` と塗り／フォント設定。

ルール: 条件付き書式は控えめに適用する。すべてのセルが色付けされたワークブックは、読み手に何も伝えない。

### データ入力規則

トラッカーやテンプレートの入力セルには、データ入力規則を必ず設定しなければならない。低コストで、下流のバグを丸ごと1クラス防げる。**リストソースの3パターン** —— 許可された値がどこにあるかで選ぶ。

**(a) インラインリスト** —— 許可された値が短く、ルール自体に固定されている場合。

```bash
officecli add "$FILE" /Sheet1 --type validation \
  --prop sqref="C2:C100" --prop type=list \
  --prop formula1="Yes,No,Maybe" \
  --prop showError=true --prop errorTitle="Invalid" --prop error="Select from list"
```

**(b) 名前付き範囲（シート間参照で推奨）** —— 許可された値が別のシートにあり、増える可能性がある場合。まず名前付き範囲を定義し、それを参照する。`ref` に `!` と `$` が含まれるため、batch ヒアドキュメントを使う:

```bash
cat <<'EOF' | officecli batch "$FILE"
[
  {"command":"add","parent":"/","type":"namedrange","props":{"name":"StatusList","ref":"Lookups!$A$2:$A$4"}},
  {"command":"add","parent":"/Sheet1","type":"validation","props":{"sqref":"B2:B100","type":"list","formula1":"=StatusList"}}
]
EOF
```

**(c) 直接のシート間範囲** —— 名前付き範囲を使わず、`formula1` の中に生の `Lookups!$A$2:$A$4` を書く。これも `!` と `$` を無傷で保つため batch ヒアドキュメントが必要。

```bash
cat <<'EOF' | officecli batch "$FILE"
[
  {"command":"add","parent":"/Sheet1","type":"validation","props":{"sqref":"C2:C100","type":"list","formula1":"Lookups!$A$2:$A$4"}}
]
EOF
```

シェル上で `--prop formula1=...` としてシート間バリアントを書くと、`!` がシェルによって `\!` に壊され、ドロップダウンは静かにリストなしにフォールバックする。`officecli get "$FILE" /Sheet1/validation[N]` で確認する —— `formula1=` はバックスラッシュのない素の `!` を表示していなければならない。

その他の一般的な `type` 値: `decimal`、`whole`、`date`、`textLength`、`custom`。演算子と全プロパティ一覧は `officecli help xlsx validation` を参照。

### その他の要素（ワンライナー）

- **テーブル**（ListObjects）—— 範囲を指定した `add --type table`。オートフィルターと構造化参照を提供する。`officecli help xlsx table`。
- **コメント** —— `add --type comment`。ハードコードされた前提条件を文書化するのに使う。`officecli help xlsx comment`。
- **シートの並べ替え** —— `swap` ではなく `officecli move` を使う。`swap` は行／セルのパスにしか効かない。

## チャート軸をロール別に操作する

チャートの軸をその場で編集する方が、チャートを作り直すより安価である。軸はインデックスではなく**ロール**（`value` = Y軸、`category` = X軸）でアドレス指定する —— XML の順序は安定していない。

```bash
officecli get "$FILE" "/Sheet1/chart[1]/axis[@role=value]"
officecli set "$FILE" "/Sheet1/chart[1]/axis[@role=value]" --prop min=0 --prop max=100000
officecli set "$FILE" "/Sheet1/chart[1]/axis[@role=category]" --prop title="Month"
```

安全なプロパティ: `title`、`min`、`max`、`majorGridlines`、`visible`、`labelRotation`。

## QA（必須）

**問題があることを前提にする。あなたの仕事はそれを見つけることである。**

最初のワークブックが正しいことはほとんどない。QA は確認作業ではなくバグハントとして扱う。最初の点検で何も問題が見つからなかったなら、それは十分に厳しく見ていない証拠である。数式は、ソースセルと突き合わせて確認するまでは問題なく見える。

### 「完了」前の最小サイクル

1. `officecli view "$FILE" issues` —— 空のシート、壊れた数式、欠落した参照。
2. `officecli view "$FILE" annotated`（サンプル範囲）—— 値＋型＋警告。
3. すべての Excel エラータイプについてクエリする:
   ```bash
   officecli query "$FILE" 'cell:contains("#REF!")'
   officecli query "$FILE" 'cell:contains("#DIV/0!")'
   officecli query "$FILE" 'cell:contains("#VALUE!")'
   officecli query "$FILE" 'cell:contains("#NAME?")'
   officecli query "$FILE" 'cell:contains("#N/A")'
   ```
4. `officecli validate "$FILE"` —— 常駐が開いた状態でも安全。`validate` はそれ自体で保留中の編集をディスクにフラッシュする。
5. **視覚パス —— HTML プレビューで全シートを歩く。** `officecli view "$FILE" html` を実行し、返された HTML パスを Read する。各シートはチャートをインラインでレンダリングして表示される。`###`、タイトルの切れ、プレースホルダートークン（`$fy$24`、`{var}`、`<TODO>`）、分割されたチャート、白いスライスの円グラフ、空のチャートアンカーがないか確認する —— **完了を宣言する前に立ち止まって修正する**。「validate を通過した」は納品ではない。「プレビューが本物のワークブックのように見える」ことが納品である。人間向けのプレビューには `officecli watch "$FILE"` を実行する（ユーザーが自分の判断でライブプレビューを開く）か、`.xlsx` を Excel / WPS / Numbers で直接開いてもらう。
6. **印刷レイアウトの修正（幅広の表／複数チャートのシート）。** シートがチャートまたは幅広の表を持ち、ユーザーが印刷する場合、シートごとのページレイアウトを設定する —— ただし、fit モードはシートの高さに合わせる:
   ```bash
   # 短いサマリー／チャートシート → 1ページに収める。
   officecli set "$FILE" "/Summary" --prop orientation=landscape --prop fitToPage=true
   # 縦に長いデータ表 → 幅だけを収める（fitToPage=true だと全行が1ページに押し込まれ読めなくなる）。
   officecli set "$FILE" "/Data" --prop orientation=landscape --prop fitToPage=1x0
   ```
   結果: チャートや幅広の表はチャートの途中で分割されずに印刷され、縦長の表は自然なページ区切りのまま読みやすさを保つ。チャートを持つシート、または8列を超える表を持つシートすべてに適用する。
7. 何か失敗した場合は修正し、**サイクル全体を再実行**する。1つの修正がもう1つの問題を生むことはよくある。

`officecli view issues` と `view html` は構造的な QA のペアである: `issues` は壊れた数式と空のシートを検出し、`view html`（返された HTML パスを Read する）は `###`、切れ、トークン漏出を検出する。チャートの塗り色／テーマの色合いはビューアによって異なることがある —— 色の忠実度が重要な場合は、ユーザーの実際のターゲットビューアでスポットチェックする。

### 数式検証チェックリスト

- [ ] 2〜3個の数式をランダムに選ぶ。それぞれに `officecli get` を実行する。数式の文字列が意図通りであること、**かつ** `cachedValue=` が期待通りであることを確認する —— 頭の中で計算する。
- [ ] **すべてのサマリーセルについてキャッシュ値の妥当性を確認する。** 集計を行うセル（COUNTA / COUNTIF / SUMPRODUCT / INDEX＆MATCH）はすべて、妥当な `cachedValue` を持たなければならない。進捗トラッカーが空のテンプレート上で `199 / 199 / 100%` を表示している場合、そのキャッシュは嘘をついている —— `set` 経由で数式を再タッチする（再計算を強制する）か、正しいキャッシュ値を手動で設定する。「validate は通過するが数値はでたらめ」を出荷してはならない。
- [ ] **数値列ごとに1セルをスポットチェックする。** `%` 列全体が整数の `0.0%` を表示している場合、分母が間違っているか分子が古いままキャッシュされている —— 1セルを調査し、パターンを修正する。
- [ ] 範囲がすべての行を含んでいるか: データが `B13` まであるのに `SUM(B2:B12)` になっているような off-by-one は最もよくあるバグである。
- [ ] シート間数式（`Sheet1!A1`）に `\!` が含まれていないか。`officecli get` が `Sheet1\!A1` を表示する場合、`!` はシェルによって壊されている —— 削除して batch／ヒアドキュメント経由で入れ直す。
- [ ] 名前付き範囲（`officecli get "$FILE" "/namedrange[1]"`）が名前の示す通りの場所を指しているか。
- [ ] すべての `/` の分母がガードされているか —— `IFERROR(x/y, 0)` または `IF(y=0, 0, x/y)`。
- [ ] チャートデータとソースセル: インラインデータを持つすべてのチャートについて、`officecli get` によるソースセルの値とデータポイントをスポットチェックする。
- [ ] チャートのタイトル／シリーズ名／凡例に未置換のトークン（`$...$`、`{var}`、`<TODO>`）が**ない**こと。`officecli get /Sheet1/chart[N]` でチャートを確認する。

### テンプレート QA

テンプレートを編集する際は、残存するプレースホルダーがないか確認する —— これらはコンテンツのように見え、`validate` をすり抜ける:

```bash
officecli query "$FILE" 'cell:contains("{{")'
officecli query "$FILE" 'cell:contains("xxxx")'
officecli query "$FILE" 'cell:contains("TBD")'
```

### 新鮮な目で

ワークブックを仕上げたら、新規に開き直す。`view text` / HTML プレビューを、新しいレビュアーになったつもりで頭から最後まで読む —— 数式、おかしく見える数値、書式の不整合、欠落したデータを探す。

### 正直な限界

`validate` はスキーマエラーを検出するのであって、設計上の誤りは検出しない。ワークブックはすべての数値が間違っていても `validate` を通過し得る。上記のチェックリスト —— 特に数式をソースセルと突き合わせるスポットチェック —— が、validate では捉えられないものを捉える方法である。

## 既知の問題と落とし穴

### シート間 `!` の罠（短縮版）

シェル（bash のヒストリ展開、zsh の分割）と CLI の引数パースは、`Sheet1!A1` の中の `!` を `\!` に壊す。`\!` を含む数式は静かに壊れており —— リテラルなテキストとしてレンダリングされ、何も参照しない。

**対処法。** シングルクォートのデリミタ（`<<'EOF'`）を使った batch ヒアドキュメントを使う。これによりすべてのシェル展開が無効になる:

```bash
cat <<'EOF' | officecli batch "$FILE"
[{"command":"set","path":"/Summary/B2","props":{"formula":"Revenue!B13"}}]
EOF
```

**検証。** 書き込み後、そのセルに `officecli get` を実行する。`formula=` はバックスラッシュのない素の `!` を表示していなければならない。

### CLI のバグ・バックログ（短縮版）

回避すべき CLI の制約とギャップ —— 出力ファイルの欠陥ではない。

- **チャートのシリーズは作成後は不変** —— シリーズを追加・変更するには: 完全なシリーズリストで `remove` してから `add` する（位置は可変: `set chart[N] --prop anchor=` / `x/y/width/height`）。`remove chart[N]` は後続のインデックスを下にシフトさせ、再追加は末尾に追加される。
- **シート間数式の batch は常駐経由でも問題なく動作する** —— 以前の「3〜5操作でもデッドロックする」という注意はもはや再現しない。純粋な値設定の batch も50〜80操作以上で安定している。もしハングに遭遇したら、非常駐の1つの大きな batch か、個別の `set` にフォールバックする。**同一ファイル／マシン上で複数の常駐プロセスが動いていると、依然として競合し得る** —— 別のエージェント／セッションが同じファイルの常駐を保持している場合、非決定的なハングが起こり得る。
- **条件付き書式の命名の非対称性** —— `--type` の要素名は `conditionalformatting` だが、パスのサフィックスは `/cf[N]` である。スキーマには `officecli help xlsx conditionalformatting`、パスには `/cf[N]` を使う。
- **add 時のシート `position` プロパティ** —— ヘルプには Add が `position` を処理すると書かれているが、このプロパティはしばしば無視される。シート作成後に `officecli move --index` / `--after` / `--before` で並べ替える。
- **`remove /sheet[N]` のカスケードガード** —— 別のシートのデータ入力規則／条件付き書式／スパークライン／ハイパーリンク／名前付き範囲から参照されているシートの削除／リネームを拒否する。まずそれらの依存要素を削除してから、シートを削除する。
- **batch の JSON はセルの `color` エイリアスを拒否する** —— batch の `props` の内部で `"color": "FF0000"` はエラー `ambiguous in cell context — use 'font.color' (text) or 'fill' (bg)` になる。シェルレベルの CLI はセル以外の要素に対しては `--prop color=...` / `--prop size=14` をエイリアスとして受け付けるが、batch の JSON でセルに対して書く場合は常にフルのドット付き名を書く: `"font.color"`、`"font.size"`、`"font.name"`。

### レンダラーの注意点（ビューア間の色の忠実度）

`officecli view html` は構造的な QA（あふれ、切れ、プレースホルダーの漏出、レイアウト）に適したツールである —— 返された HTML パスを Read する。チャートのレンダリングの一部の細部は、エンドユーザーがファイルを開くビューアによって異なる。観測された相違:

- **円グラフ／ドーナツグラフの塗り色が、一部のビューアでは単一のテーマの色合いに潰れることがある**（スライスが「すべて白」または「すべて同色」に見える）。ファイル自体はユーザーのターゲットビューアでは問題ない場合がある。
- **折れ線グラフ／棒グラフのシリーズの色が、一部のビューアではワークブックのテーマからずれることがある。**
- **フォームコントロールのチェックボックスが、一部のビューアでは二重枠として表示されることがある。**

色やチャートを「壊れている」と判断する前に、ユーザーの実際のターゲットビューアでファイルを開く。そこで正しく見えるなら、問題はデータではなくビューアのレンダリングであり、追いかける必要はない。CLI の構造的チェック（`###`、切れ、プレースホルダーテキスト、レイアウト）が引き続き権威あるものである。

### エスケープの階層（シェルクォーティングは上で扱った。これらはその追加分）

`$` はシェル層の話であり（前述の通りシングルクォートで対処する）、プロパティ値中の `\n` / `\t` は CLI によって実際の改行／タブに解釈される。さらに2つの層がある:

- **JSON レベル（batch）。** 標準的な JSON エスケープ —— `"\n"`、`"\t"`、`"\""`。最終文字列中の実際のバックスラッシュは `"\\\\"` になる。
- **Excel レベル。** セル内の `\n` は実際の改行である —— `--prop wrapText=true` と組み合わせて、Excel に折り返しを表示させる。シェルでクォートされたプロパティに直接書いても動く（`--prop value='a\nb'`）。batch JSON 内の `"\n"` も同じ結果になる。迷ったら、`officecli get` でセルを確認し、文字単位で比較する。

### その他のよくある落とし穴

| 落とし穴 | 対処法 |
|---|---|
| `--name "foo"` | すべての属性は `--prop` 経由: `--prop name="foo"` |
| プロパティ名を推測する | `officecli help xlsx <element>` —— 即興で済ませない |
| セルに `--prop color=...` | 曖昧 —— `font.color`（テキスト）または `fill`（背景）を使う。batch JSON の内部でも同様に、シェルのエイリアスではなく常にフルのドット付き名を使う |
| `#FF0000` の16進カラー | `#` を外す: `FF0000` |
| `--index` と `[N]` | `--index` は0始まり（配列）; `[N]` パスは1始まり（XPath） |
| zsh/bash で `[N]` をクォートし忘れる | すべてのパスをクォートする: `"/Sheet1/row[1]"` |
| スペースを含むシート名 | パス全体をクォートする: `"/My Sheet/A1"` |
| 年が `2,026` として表示される | `--prop type=string` または `numFmt="@"` |
| Excel で開いたままのファイルを変更する | 先に Excel 側でファイルを閉じる |
| `swap` がシートを並べ替えない | `swap` は行／セル用。シートには `move --after` / `--before` / `--index` を使う |
| 書き込み後にキャッシュ値が欠落する | 人間がファイルを開いたときに新規数式へキャッシュ値が付与される; `validate` はどちらの状態でも受け付ける |
