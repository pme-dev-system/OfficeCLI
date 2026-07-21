---
name: officecli
description: officecli CLI ツールを使って Office 文書（.docx, .xlsx, .pptx）の作成、分析、校正、変更を行います。ユーザーが Office 文書を作成、検査、書式チェック、問題の発見、チャートの追加、変更をしたい場合に使用してください。
---

# officecli

.docx、.xlsx、.pptx 向けの AI フレンドリーな CLI。単一バイナリ、依存関係なし、Office のインストール不要。

## インストール

`officecli` がインストールされていない場合：

```bash
# macOS / Linux
curl -fsSL https://d.officecli.ai/install.sh | bash

# Windows (PowerShell)
irm https://d.officecli.ai/install.ps1 | iex
```

`officecli --version` で確認してください。インストール後も見つからない場合は、新しいターミナルを開いてください。

---

## 戦略

**L1（読み取り）→ L2（DOM 編集）→ L3（生 XML）**。常により上位のレイヤーを優先してください。構造化出力には `--json` を追加します。

**文書作業の前に、専門スキル（このファイルの末尾）を確認してください。** 資金調達用デック、学術論文、財務モデル、ダッシュボード、Morph アニメーションはそれぞれ専用のスキルを先にロードする必要があります — `load_skill` を一度実行してから進めてください。

---

## ヘルプシステム（重要）

**プロパティ名、値の形式、コマンド構文について不明な場合は、推測せずに常にヘルプを実行してください。** 1 回のヘルプ照会は、推測→失敗→再試行のループより優れています。

`officecli help` ≡ `officecli --help`、そして `officecli <cmd> --help` ≡ `officecli help <cmd>` — 内容は同じです。

```bash
officecli help                                  # 全コマンド + グローバルオプション + スキーマのエントリーポイント
officecli help docx                             # docx の全要素を一覧表示
officecli help docx paragraph                   # 完全なスキーマ：プロパティ、エイリアス、例、読み戻し値
officecli help docx set paragraph               # 動詞フィルタ済み：`set` で使用可能なプロパティのみ
officecli help docx paragraph --json            # 構造化スキーマ（機械可読）
```

フォーマットエイリアス：`word`→`docx`、`excel`→`xlsx`、`ppt`/`powerpoint`→`pptx`。動詞：`add`、`set`、`get`、`query`、`remove`。MCP は同一のスキーマを単一の `command` 文字列パラメータ経由で公開します：`{"command":"help docx paragraph"}`（構造化された `{"format":...,"type":...}` オブジェクトではありません — MCP ツールはパラメータを `command` の 1 つだけ持ち、CLI へそのまま渡します）。

---

## パフォーマンス：レジデントモード

**すべてのコマンドは初回アクセス時に自動でレジデント（常駐プロセス）を起動します**（60 秒のアイドルタイムアウト）— ファイルロックの競合は自動的に回避されます。長いセッションでは明示的な `open`/`close`（12 分のアイドルタイムアウト）を引き続き推奨します：
```bash
officecli open report.docx       # 明示的にメモリに保持
officecli set report.docx ...    # ファイル I/O のオーバーヘッドなし
officecli close report.docx      # 保存して解放
```

自動起動をオプトアウトする場合：`OFFICECLI_NO_AUTO_RESIDENT=1`。

**フラッシュは officecli 以外との境界でのみ行ってください。** officecli 自身の読み取り（`get`/`query`/`view`/`dump`）は常に最新の編集内容を参照するため、ワークフロー途中で保存する必要はありません。`save`（レジデントは維持）または `close`（フラッシュして解放）を実行するのは、**officecli 以外のプログラムがファイルを読み込む直前**のみです — python-docx/openpyxl、Word、レンダラー、配信/アップロードなど。（アイドル状態のセッションは数秒以内に自動フラッシュされます。`OFFICECLI_RESIDENT_FLUSH=each` を設定すると、すべての変更が返却前にフラッシュされます。）

---

## クイックスタート

**PPT：**
```bash
officecli create slides.pptx
officecli add slides.pptx / --type slide --prop title="Q4 Report" --prop background=1A1A2E
officecli add slides.pptx '/slide[1]' --type shape --prop text="Revenue grew 25%" --prop x=2cm --prop y=5cm --prop font=Arial --prop size=24 --prop color=FFFFFF
```

**Word：**
```bash
officecli create report.docx
officecli add report.docx /body --type paragraph --prop text="Executive Summary" --prop style=Heading1
officecli add report.docx /body --type paragraph --prop text="Revenue increased by 25% year-over-year."
```

**Excel：**
```bash
officecli create data.xlsx
officecli set data.xlsx /Sheet1/A1 --prop value="Name" --prop bold=true
officecli set data.xlsx /Sheet1/A2 --prop value="Alice"
```

---

## L1：作成・読み取り・検査

```bash
officecli create <file>               # 空の .docx/.xlsx/.pptx を作成（拡張子からタイプを判定）
officecli view <file> <mode>          # outline | stats | issues | text | annotated | html
officecli get <file> <path> --depth N # ノードとその子要素を取得 [--json]
officecli query <file> <selector>     # CSS ライクなクエリ
officecli validate <file>             # OpenXML スキーマに対して検証
```

### view モード

| モード | 説明 | 便利なフラグ |
|------|-------------|-------------|
| `outline` | 文書構造 | |
| `stats` | 統計情報（ページ数、単語数、シェイプ数） | |
| `issues` | フォーマット/コンテンツ/構造上の問題 | `--type format\|content\|structure`、`--limit N` |
| `text` | プレーンテキスト抽出 | `--start N --end N`、`--max-lines N` |
| `annotated` | フォーマット注釈付きテキスト | |
| `html` | 静的 HTML スナップショット — `watch` と同じレンダラー、サーバー不要 | `--browser`、`--page N`（docx）、`--start N --end N`（pptx） |
| `screenshot` / `svg` / `pdf` / `forms` | ヘッドレスブラウザ経由の PNG / SVG（pptx スライド）/ エクスポータプラグイン経由の PDF / フォーマットハンドラプラグイン経由のフォームフィールド JSON | `-o`、`--screenshot-width/-height`、pptx の `--grid N` |

一度限りのスナップショット（CI アーティファクト、アーカイブ、差分比較）には `view html` を使用し、ライブリフレッシュやブラウザ側のクリック選択が必要な場合は `watch` を使用してください。

### get

要素の localName による任意の XML パス。子要素を展開するには `--depth N` を使用します。構造化出力には `--json` を追加します。デフォルトのテキスト出力は grep しやすい形式です：`path (type) "text" key=val key=val ...`

```bash
officecli get report.docx '/body/p[3]' --depth 2 --json
officecli get slides.pptx '/slide[1]' --depth 1          # スライド 1 の全シェイプを一覧表示
officecli get data.xlsx '/Sheet1/B2' --json
```

### 安定 ID アドレッシング

安定 ID を持つ要素は、位置インデックスではなく `@attr=value` パスを返します。複数ステップのワークフローではこちらを優先してください — 位置インデックスは挿入/削除で変動しますが、安定 ID は変動しません。

```
/slide[1]/shape[@id=550950021]                    # PPT シェイプ
/slide[1]/table[@id=1388430425]/tr[1]/tc[2]       # PPT 表
/body/p[@paraId=1A2B3C4D]                         # Word 段落
/comments/comment[@commentId=1]                    # Word コメント
```

PPT は `@name=`（例：`shape[@name=Title 1]`）も受け付け、morph の `!!` プレフィックスも認識します。安定 ID を持たない要素（slide、run、tr/tc、row）は位置インデックスにフォールバックします。

### query

CSS ライクなセレクタ：`[attr=value]`、`[attr!=value]`、`[attr~=text]`、`[attr>=value]`、`[attr<=value]`、`:contains("text")`、`:empty`、`:has(formula)`、`:no-alt`。ブール演算子 `and`/`or` は `query`/`set`/`remove` にわたってサポートされています：`cell[value>5000 or value<100]`、`cell[(type=Number or type=Date) and value>0]`。Excel の列名による行指定：`Sheet1!row[Salary>5000]`。`set` はセレクタと Excel ネイティブパスの両方を受け付けます（`get`/`query` とのパリティ）。裸の（スコープなしの）セレクタは `set`/`remove` では拒否されます。

```bash
officecli query report.docx 'paragraph[style=Normal] > run[font!=Arial]'
officecli query slides.pptx 'shape[fill=FF0000]'
```

---

## Watch とインタラクティブ選択

ファイルが変更されるたびに自動でリフレッシュされるライブ HTML プレビュー。ブラウザ上でクリック / シフトクリック / ボックスドラッグによってシェイプを選択でき、CLI は現在のブラウザ選択状態を読み取って処理できます。

```bash
officecli watch <file> [--port N]      # プレビューサーバーを起動（デフォルトポート 26315）
officecli unwatch <file>               # 停止
officecli goto <file> <path>           # watch 中のブラウザを要素へスクロール（docx：p / table / tr / tc）
```

出力された `http://localhost:N` の URL を開いてください。クリックで選択、shift/cmd/ctrl+クリックで複数選択、空白部分からドラッグでボックス選択できます。PPT/Word は青いアウトライン、Excel はネイティブ風の緑色の選択表示を使用します（セルをダブルクリックでインライン編集、チャートをドラッグで再配置）。

### `get <file> selected` — ユーザーのクリック内容を読み取る

```bash
officecli get <file> selected [--json]
```

現在選択されているものの DocumentNodes を返します。何も選択されていない場合は空の結果になります。watch が実行されていない場合、終了コードは 0 以外になります。

```bash
# ユーザーがブラウザでシェイプをクリックし、「これらを赤くして」と指示した場合
PATHS=$(officecli get deck.pptx selected --json | jq -r '.data.Results[].path')
for p in $PATHS; do officecli set deck.pptx "$p" --prop fill=FF0000; done
```

### 主要な特性

- **選択状態はファイル編集後も保持されます。** パスは安定した `@id=` 形式を使用します。
- **接続中の全ブラウザが 1 つの選択状態を共有します。** 最後の書き込みが優先されます。
- **同一ファイルにつき同時に 1 つの watch のみ。** 1 つのファイルにつき watch プロセスは同時に 1 つだけです。
- **グループシェイプは全体として選択されます。** グループ内の個々の子要素へのドリルダウンは v1 ではサポートされていません。
- **カバレッジ：** `.pptx` のシェイプ/画像/表/チャート/コネクタ/グループ、`.docx` のトップレベル段落と表。継承されたレイアウト/マスターの装飾や Word のネストされた要素（表セル、ラン レベル）はアドレス指定できません。**`.xlsx` は `data-path` を出力しません** — xlsx 上の `mark`/`selection` は常に `stale=true` に解決されます（v2 の候補機能）。

### マーク — レビュー待ちの編集提案

変更をファイルに反映する**前に**人間によるレビューが必要な場合は `mark` を使用してください。マークは watch プロセス内にのみ存在し、承認されたものは別の `set` パイプラインで適用されます。一度限りの変更には `set` を直接使用し、永続的なファイル注釈には `add --type comment`（Word ネイティブ）を使用してください。

```bash
officecli mark <file> <path> [--prop find=... color=... note=... tofix=... regex=true] [--json]
officecli unmark <file> [--path <p> | --all] [--json]
officecli get-marks <file> [--json]
```

プロパティ：`find`（リテラルまたは `regex=true` 時は正規表現。生の形式：`find='r"[abc]"'`）、`color`（16進数 / `rgb(...)` / 22 種類の名前付きホワイトリスト）、`note`、`tofix`（適用パイプラインを駆動）。**パス**は watch HTML 由来の `data-path` 形式である必要があります — 完全な適用パイプラインについてはサブスキルを参照してください。

---

## L2：DOM 操作

### set — プロパティを変更

```bash
officecli set <file> <path> --prop key=value [--prop ...]
```

**任意の XML 属性は要素パス経由で設定可能です**（`get --depth N` で調査）— 現在存在しない属性も設定できます。`find=` を指定しない場合、`set` は要素全体にフォーマットを適用します。

**値の形式：**

| タイプ | 形式 | 例 |
|------|--------|---------|
| 色 | 16進数（`#` あり/なし）、色名、RGB、テーマ | `FF0000`、`#FF0000`、`red`、`rgb(255,0,0)`、`accent1`..`accent6` |
| 間隔 | 単位指定 | `12pt`、`0.5cm`、`1.5x`、`150%` |
| 寸法 | EMU または接尾辞付き | `914400`、`2.54cm`、`1in`、`72pt`、`96px` |

**ドット区切り属性エイリアス** — `font.<attr>` 形式は shape/run/paragraph/table/row/cell/section/styles で受け付けられます。例：`--prop font.color=red --prop font.bold=true --prop font.size=14pt`。完全なリストは `officecli help <fmt> <element>` を実行してください。

### find — 一致したテキストのフォーマットまたは置換

`set`（および `query` の `--find`）ではトップレベルの `--find` / `--replace` を使用してください。旧来の `--prop find=X` も引き続き動作しますが、ヒントが表示されます。

```bash
# 一致したテキストをフォーマット（ランを自動分割）
officecli set doc.docx '/body/p[1]' --find weather --prop bold=true --prop color=red

# 正規表現マッチング（regex= は依然としてプロパティフラグ）
officecli set doc.docx '/body/p[1]' --find '\d+%' --prop regex=true --prop color=red

# テキストを置換（文書全体をスコープにする場合は `/` を使用）
officecli set doc.docx / --find draft --replace final

# docx：変更履歴付き検索/置換
officecli set doc.docx / --find draft --replace final --prop revision.author=Alice

# PPT — 同じ構文、パスのみ異なる
officecli set slides.pptx / --find draft --replace final
```

**パスが検索範囲を制御します：** `/` = 文書全体、`/body/p[1]` または `/slide[N]/shape[M]` = 特定の要素、`/header[1]` / `/footer[1]` = ヘッダー/フッター。

**注意事項：**
- デフォルトでは大文字小文字を区別します。区別しない場合：`--prop 'find=(?i)error' --prop regex=true`
- マッチはラン境界をまたいで機能します
- マッチしない場合は静かに成功扱いになります。`--json` には `"matched": N` が含まれます
- **Excel：** `find` + `replace` のみサポート（find + フォーマットプロパティの組み合わせは非対応）

### add — 要素の追加またはクローン

```bash
officecli add <file> <parent> --type <type> [--prop ...]
officecli add <file> <parent> --type <type> --after <path> [--prop ...]   # アンカーの後に挿入
officecli add <file> <parent> --type <type> --before <path> [--prop ...]  # アンカーの前に挿入
officecli add <file> <parent> --type <type> --index N [--prop ...]        # 0-based の位置（レガシー）
officecli add <file> <parent> --from <path>                               # 既存要素をクローン
```

`--after`、`--before`、`--index` は互いに排他的です。位置フラグを指定しない場合は末尾に追加されます。

**要素タイプ（エイリアス付き）：**

| フォーマット | タイプ |
|--------|-------|
| **pptx** | slide（非表示を含む）、shape（font.latin/ea/cs、direction=rtl、underline.color、highlight=COLOR（Add/Set/Get/HTML プレビュー）、effective.X+effective.X.src；rightArrow のエイリアス arrow；slideMaster/slideLayout の型付き add/set/remove）、picture（SVG、brightness/contrast/glow/shadow、rotation、link、tooltip）、chart（direction=rtl、pieOfPie、barOfPie、軸線/グリッド線の属性別セッター、animation+chartBuild=byCategory|bySeries、line の dropLines/hiLowLines/upDownBars、anchor=x,y,w,h 省略記法）、table（cell の direction=rtl、fill/background、PowerPoint 組み込みスタイルカタログ、/col[C] get + swap/copyFrom、row/col の Move/CopyFrom）、row（tr）、connector（from/to は完全パスの `@name=`/`@id=` 形式を受け付ける — 裸の `@name=Foo` は拒否され、`/slide[N]/shape[@name=Foo]` の形式が必須；startshape/endshape の SetByPath；デフォルトでエッジ間アンカリング、エッジを強制する場合は fromSide/toSide、生の cxn インデックスには fromIdx/toIdx）、group（link、tooltip、get/query/add/remove によるディープウォーク、ungroup=true でスライド絶対座標に解消）、align/distribute（targets= は位置指定だけでなく shape[@id=N] パスも受け付ける）、video/audio（loop、autoStart のエイリアス）、equation、notes（direction=rtl、lang）、comment（レガシー + p188 モダンスレッド化の往復対応）、animation（15 種の強調 + 16 種の終了プリセット、複数エフェクトチェーン、モーションパスプリセット、repeat/restart/autoReverse、チャートアニメーション）、transition（12 種の p15 プリセット + morph/p14）、paragraph（para）、run、zoom、ole（preview=、add-part+raw-set 経由の完全ダンプ往復）、placeholder（phType=...）、model3d（rotation=ax,ay,az；完全ダンプ往復）、smartart（add-part 経由のダンプ往復）、diagram（追加専用、mermaid → ネイティブシェイプまたはレンダリング画像、`--type diagram`/`flowchart`）。 |
| **docx** | paragraph（direction/font.latin/ea/cs、bold.cs/italic.cs/size.cs、lang.latin/ea/cs、wordWrap、framePr.\*、tabs 省略記法）、run（lang スロット、direction、underline.color、position ハーフポイント、**revision.type=ins\|del\|format\|moveFrom\|moveTo + revision.action=accept\|reject**（.author/.date 付き）— フィルタ済み accept/reject には `set /revision[...]` 上の裸の `@author=`/`@type=` セレクタを使用するが、`query 'revision[...]'` はドット区切りの `revision.author=`/`revision.type=` 形式が必要；move+revision はラン レベルのパスのみで段落レベルでは不可；**range=START:END** は段落/シェイプ パス上で、ラン をアドレス指定する代わりに明示的な 0-based の半開区間オフセットで文字範囲をフォーマットする — find= の兄弟にあたるオフセット指定）、table（direction=rtl、hMerge、row の cantSplit / cell の nowrap（Add・Set 両対応）、**仮想列操作**：/body/tbl[N]/col への add/remove/move/copyfrom）、row（tr）、cell（td）、image、header/footer（direction）、section（pageNumFmt の完全な列挙、direction=rtl、rtlGutter、pgBorders=box）、bookmark、comment、footnote、endnote、formfield、sdt、chart、equation、field（28 種類）、hyperlink、style（direction、indents、pbdr、Add/Set 上の lineSpacing）、toc、watermark、break、ole、**num/abstractNum/lvl**、**tab**、**textbox/shape**（主に追加専用 — Get は生の XML プレビューのみを返し、構造化された読み戻しは不可；Set は width/height/geometry/fill/line.\* に限定；位置は裸の x/y ではなく `anchor.x`/`anchor.y`；**textbox のみ** `textDirection`/rotation/gradient/shadow — docx の shape 自体には rotation も gradient もありません）、埋め込み **OLE の dump→batch 往復**、**diagram**（追加専用、mermaid → ネイティブシェイプまたはレンダリング画像、`--type diagram`/`flowchart`、追加時に x/y なし — `set /body/group[N]` で再配置）。docDefaults.rtl、autoHyphenation、`get /` はロケール + /comments /footnotes /endnotes を公開。生の OOXML スキャフォールディングには `create --minimal`。 |
| **xlsx** | sheet（visible/hidden/veryHidden、印刷余白、printTitleRows/Cols、RTL の sheetView、カスケード対応のリネーム）、row（c{N}= セル内容の省略記法；add は `--from /Sheet/col[L]` を受け付ける；挿入時の数式参照リライト）、col（数式参照リライト、移動時の名前付き範囲追従）、cell（type=richtext+runs、merge=range/sweep、direction=rtl、phonetic；**remove 時は --shift left\|up、add 時は shift=right\|down** — Excel UI ダイアログとのパリティ；数式自動検出；計算時の OFFSET/INDIRECT）、chart（軸別 RTL/タイトル、anchor=x,y,w,h、パレート）、image（SVG）、comment（direction=rtl）、table（listobject）、namedrange（definedname、volatile、`[@name=X]`；パース時に数式本体をインライン化）、pivottable（キャッシュの CoW + クロスピボット共有、labelFilter=field:type:value は追加時のみ、topN=integer は追加時のみ、fillDownLabels は repeatLabels のエイリアスであり別機能ではない、calculatedField）、sparkline、validation、autofilter、shape、textbox、CF（databar/colorscale/iconset/formulacf/cellIs/topN/aboveAverage）、ole、csv。Query は `merge`/`mergedrange` をサポート。ワークブック：password。Shape セレクタは grpSp 内の末端要素を列挙します。 |

### ピボットテーブル（xlsx）

```bash
officecli add data.xlsx /Sheet1 --type pivottable \
  --prop source="Sheet1!A1:E100" --prop rows=Region,Category \
  --prop cols=Year --prop values="Sales:sum,Qty:count" \
  --prop grandTotals=rows --prop subtotals=off --prop sort=asc
```

主なプロパティ：`rows`、`cols`、`values`（Field:func[:showDataAs]）、`filters`、`source`、`position`、`layout`（compact/outline/tabular）、`repeatLabels`、`blankRows`、`aggregate`、`showDataAs`（percent_of_total/row/col、running_total）、`grandTotals`、`subtotals`、`sort`。集計関数：sum、count、average、max、min、product、stdDev、stdDevp、var、varp、countNums。日付列は自動でグループ化されます。完全なスキーマは `officecli help xlsx pivottable` を実行してください。

### 文書レベルのプロパティ（全フォーマット共通）

```bash
officecli set doc.docx / --prop docDefaults.font=Arial --prop docDefaults.fontSize=11pt
officecli set doc.docx / --prop protection=forms --prop evenAndOddHeaders=true
officecli set data.xlsx / --prop calc.mode=manual --prop calc.refMode=r1c1
officecli set slides.pptx / --prop defaultFont=Arial --prop show.loop=true --prop print.what=handouts
```

全ての文書レベルプロパティ（docDefaults、docGrid、CJK 間隔、calc、print、show、theme、extended）は `officecli help <format> /` を実行してください。

### ソート（xlsx）

```bash
officecli set data.xlsx /Sheet1 --prop sort="C desc" --prop sortHeader=true
officecli set data.xlsx '/Sheet1/A1:D100' --prop sort="A asc" --prop sortHeader=true
```

形式：`COL DIR[, COL DIR ...]`。結合セルまたは数式を含む範囲は拒否されます。サイドカーメタデータ（ハイパーリンク、コメント、条件付き書式、図形描画）は行に自動で追従します。

### テキストアンカー挿入（`--after find:X` / `--before find:X`）

段落内のテキスト一致箇所を挿入位置として指定します。インライン型（run、picture、hyperlink）は段落内に挿入され、ブロック型（table、paragraph）は段落を自動分割します。PPT はインラインのみサポートします。

```bash
# Word：一致したテキストの後にインラインの run を挿入
officecli add doc.docx '/body/p[1]' --type run --after find:weather --prop text=" (sunny)"

# Word：一致したテキストの後にブロックの table を挿入（段落を自動分割）
officecli add doc.docx '/body/p[1]' --type table --after "find:First sentence." --prop rows=2 --prop cols=2
```

### クローン

`officecli add <file> / --from '/slide[1]'` — すべてのクロスパート関係を含めてコピーします。

### move、swap、remove

```bash
officecli move <file> <path> [--to <parent>] [--index N] [--after <path>] [--before <path>]
officecli swap <file> <path1> <path2>
officecli remove <file> '/body/p[4]'
```

`--after` または `--before` を使用する場合、`--to` は省略可能です — ターゲットのコンテナはアンカーから推測されます。

### batch — 1 回の保存サイクルで複数の操作を実行

**デフォルトでアトミック（v1.0.137 以降）：** 各アイテムは実行され、レポートされます（そのため `N succeeded, M failed` は意味を保ち、すべての失敗が表面化します）が、*いずれか*のアイテムが失敗するとバッチ全体がロールバックされます — ディスク上のファイルはバッチ実行前とバイト単位で同一に保たれます（スタンドアロン・レジデントモードの両方でライブ確認済み）。`--best-effort` を使用すると、旧来の「成功したものだけ適用する」動作に戻せます（不完全な結果より、1 つの非対応アイテムのために全体を失う方が悪い、可逆性の低い `dump→batch` の再生に有用）。`--stop-on-error` は実行の早期停止タイミングのみを変更し（残りのアイテムは `skipped` になります）、実行済みのものが保持されるかどうかには影響しません — 「最初の失敗で停止するが、既に成功したものは保持する」を望む場合は `--best-effort` と組み合わせてください。`--force` はこれとは無関係で、docx の保護回避専用です。失敗したアイテムには機械可読な `code` フィールドが付与されます（`error.code` と同じリスト）。ロールバックされたバッチの JSON サマリーには `"atomicRolledBack": true` が付与されます。

`officecli dump <file> [<path>]` は往復再生可能なバッチ JSON を出力します — `.docx`（全カバレッジ）、`.pptx`（生の raw-set パススルー経由でテキスト/表/画像/チャート/ノート/テーマ + OLE/3D/ビデオ/オーディオ/SmartArt/morph/p15 トランジション）、`.xlsx`（セル/数式/スタイル + 表、条件付き書式、入力規則、コメント、チャート、スパークライン、画像、シェイプ、ピボットテーブル；スライサー/chartEx/OLE は逐語キャリア経由）。パスのデフォルトは `/`（文書全体）です。サブツリーのパス（docx：`/body`、`/body/p[N]`、`/body/tbl[N]`、`/theme`、`/settings`、`/numbering`、`/styles`；xlsx：`/SheetName`、`/sheet[N]`）を渡すと、ダンプ範囲を限定できます。`officecli refresh <file.docx>` は再生後に TOC ページ番号 / PAGE / 相互参照を再計算します（Windows 上では Word バックエンド、それ以外ではヘッドレス HTML フォールバック）。`officecli plugins list` は `.doc`、`.hwpx`、`.pdf` エクスポートへの対応を拡張します。

```bash
echo '[
  {"command":"set","path":"/Sheet1/A1","props":{"value":"Name","bold":"true"}},
  {"command":"set","path":"/Sheet1/B1","props":{"value":"Score","bold":"true"}}
]' | officecli batch data.xlsx --json

officecli batch data.xlsx --commands '[{"op":"set","path":"/Sheet1/A1","props":{"value":"Done"}}]' --json
officecli batch data.xlsx --input updates.json --best-effort --json   # 一部が失敗しても成功したものは保持する
```

サポート対象：`add`、`set`、`get`、`query`、`remove`、`move`、`swap`、`view`、`raw`、`raw-set`、`validate`。フィールド：`command`（または `op`）、`path`、`parent`、`type`、`from`、`to`、`index`、`after`、`before`、`props`、`selector`、`mode`、`depth`、`part`、`xpath`、`action`、`xml`。

---

## L3：生 XML

L2 で表現できない場合に使用します。xmlns の宣言は不要です — プレフィックスは自動登録されます。

```bash
officecli raw <file> <part>                          # 生 XML を表示
officecli raw-set <file> <part> --xpath "..." --action replace --xml '<w:p>...</w:p>'
officecli add-part <file> <parent>                   # 新しい文書パートを作成（rId を返す）
```

`raw-set` のアクション：`append`、`prepend`、`insertbefore`、`insertafter`、`replace`、`remove`、`setattr`。利用可能なパートは `officecli help <format> raw` を実行してください。

---

## よくある落とし穴

| 落とし穴 | 正しいアプローチ |
|---------|-----------------|
| `--name "foo"` | `--prop name="foo"` を使用してください — すべての属性は `--prop` 経由で指定します |
| zsh/bash でクォートなしの `[N]` パス | 常にクォートしてください：`'/slide[1]'` または `"/slide[1]"`（シェルが角括弧をグロブ展開してしまいます） |
| コンテンツに PPT `shape[1]` を使う | `shape[1]` は通常タイトルのプレースホルダーです。コンテンツ用のシェイプには `shape[2]` 以降を使用してください |
| `/shape[myname]` | 名前によるインデックス指定はサポートされていません。数値インデックスまたは `@name=`（PPT のみ）を使用してください |
| プロパティ名の推測 | 正確な名前を確認するには `officecli help <format> <element>` を実行してください |
| 開いているファイルの変更 | まず PowerPoint/WPS でファイルを閉じてください |
| シェル文字列内の `\n` | `--prop text="..."` 内の改行には `\\n` を使用してください |
| シェルテキスト内の `$` | `--prop text="$15M"` は `$15` を除去してしまいます。シングルクォートを使用してください：`--prop text='$15M'`、またはヒアドキュメントのバッチ |

---

## 専門スキル

`officecli load_skill <name>` — 出力は SKILL.md であり、その規則に従ってください。

**ロードのルール**：
- 「When to use」の中で最も具体的に一致するものを選び、該当がなければフォーマットのデフォルト（`word` / `pptx` / `excel`）をロードしてください。
- シーンにはすでにフォーマットのデフォルトの規則が含まれています — 1 つの成果物につき 1 つのスキルのみロードし、重ねてロードしないでください。
- ロードされた規則はターン間で保持されます。毎回の返信で再ロードする必要はありません。
- 2 つの異なる成果物 → 2 回の個別ロード。

### Word (.docx)

| 名前 | 使用場面 |
|------|-------------|
| `word` | レポート、レター、メモ、提案書、一般的な文書 |
| `academic-paper` | 学術誌 / 会議 / 論文：APA / Chicago / IEEE / MLA 引用形式、数式、SEQ + PAGEREF 相互参照、複数段組みの学術誌レイアウト、参考文献。ビジネスレポートやレターには使用しないでください（そちらは `word` へ） |

### PowerPoint (.pptx)

| 名前 | 使用場面 |
|------|-------------|
| `pptx` | 一般的なデック：役員会向けレビュー、営業デック、全社会議、製品ローンチ |
| `pitch-deck` | **資金調達専用** — シード / シリーズ A〜C / SAFE / コンバーティブル / 戦略的資金調達。営業/製品/役員会向けデックには使用しないでください（そちらは `pptx` へ） |
| `morph-ppt` | シネマティックな Morph アニメーション付きプレゼンテーション。静的なデックには使用しないでください（そちらは `pptx` へ） |
| `morph-ppt-3d` | 3D Morph：GLB モデル、カメラワーク、奥行き。2D のみの Morph には使用しないでください（そちらは `morph-ppt` へ） |

### Excel (.xlsx)

| 名前 | 使用場面 |
|------|-------------|
| `excel` | 一般的なワークブック、数式、ピボット、トラッカー |
| `financial-model` | 財務モデル、シナリオ、予測。一般的なデータ分析には使用しないでください（そちらは `excel` へ） |
| `data-dashboard` | CSV/表形式データ → チャートとスパークライン付きの KPI / 分析 / 経営ダッシュボード。生データのトラッキングには使用しないでください（そちらは `excel` へ） |

例：資金調達デックのタスク → `officecli load_skill pitch-deck` → 出力された規則を使用。

---

## 注意事項

- パスは **1-based** です（XPath の慣例）：`'/body/p[3]'` = 3 番目の段落
- `--index` は **0-based** です（配列の慣例）：`--index 0` = 1 番目の位置
- **Excel の例外**：`add --type row` と `add --type col` では、`--index N` は **1-based** です（OOXML の RowIndex / 列文字インデックスに一致）。`--index 5` は行 5 / 列 5 に挿入されます。
- 変更後は `validate` および/または `view issues` で確認してください
- **不明な場合は**、推測せずに `officecli help <format> <element>` を実行してください
</content>
