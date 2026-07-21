---
name: officecli-docx
description: "このスキルは .docx ファイルが入力・出力のいずれか（あるいは両方）で関わる場面で常に使用してください。対象には、Word 文書・レポート・手紙・メモ・提案書の作成、任意の .docx ファイルからのテキスト読み取り／解析／抽出、既存文書の編集・修正・更新、テンプレート・変更履歴・コメント・ヘッダー/フッター・目次の操作が含まれます。ユーザーが「Word doc」「document」「report」「letter」「memo」に言及した場合、または .docx ファイル名を挙げた場合にトリガーしてください。"
---

# OfficeCLI DOCX スキル

## セットアップ

`officecli` が見つからない場合：

- **macOS / Linux**: `curl -fsSL https://d.officecli.ai/install.sh | bash`
- **Windows (PowerShell)**: `irm https://d.officecli.ai/install.ps1 | iex`

`officecli --version` で確認してください（PATH が反映されていない場合は新しいターミナルを開いてください）。インストールに失敗した場合は https://github.com/iOfficeAI/OfficeCLI/releases からバイナリをダウンロードしてください。

## ⚠️ ヘルプ優先ルール

**このスキルが教えるのは「良い docx とはどういうものか」であり、すべてのコマンドフラグではありません。プロパティ名、列挙値、エイリアスに確信が持てない場合は、推測する前にヘルプを確認してください。**

```bash
officecli help docx                         # すべての docx 要素を一覧表示
officecli help docx <element>               # 要素の完全なスキーマ（例: paragraph, field, numbering, watermark, toc）
officecli help docx <verb> <element>        # 動詞スコープのヘルプ（例: add field, set section）
officecli help docx <element> --json        # 機械可読なスキーマ
```

ヘルプはインストールされている CLI のバージョンに紐づいています。このスキルとヘルプの内容が食い違う場合、**ヘルプが正です**。

## メンタルモデル

`.docx` は XML パーツ（`document.xml`、`styles.xml`、`numbering.xml`、`header*.xml`、`footer*.xml`、`comments.xml` など）を格納した ZIP です。ユーザーが目にするもの — 見出し、表、ページ番号、目次、変更履歴 — はすべて、その ZIP 内の XML です。`officecli` はその上にセマンティックパス API（`/body/p[1]/r[2]`）を提供するため、生の XML に触れる必要はほとんどありません。どうしても必要な場合は `raw-set` を使ってください（XML 付録を参照）。

## シェルと実行の規律

docx のパスには `[]` が含まれ、一部のプロパティ値には `$` が含まれます。どちらもシェルのメタ文字です。エスケープは 3 つの層で行われます — それぞれを混同しないでください。

1. **シェル。** 要素パスは必ずクォートしてください: `"/body/p[1]"`（`/body/p[1]` は不可） — zsh/bash が `[N]` をグロブとして解釈します。`$` を含む値はシングルクォートで囲んでください: `--prop text='$50M'` — どんなに長い文字列でも、`$` を含む部分は一対のシングルクォートの中に収めます。クォートしない `$50M` は `M` だけに削られてしまいます。1 本の長い文字列の中で `'…$var…'` と `"…$50…"` を混在させると、`$50` が静かに消える原因になります。
2. **CLI（`text=`）。** `\n` と `\t` の 2 文字エスケープは `--prop text=` の中で実際に解釈されます — `\n` は `<w:br/>` の改行に、`\t` は `<w:tab/>` になります。これは docx / pptx / xlsx で一貫しています。リテラルなバックスラッシュ + n（ほとんど必要にならないケース）にしたい場合は二重にしてください（`\\n`）。この挙動は行レベルの表ショートカット `c1…cN` にも適用されます（セル内で `\n` → `<w:br/>`）。
3. **JSON（バッチ）。** `batch` ヒアドキュメントの JSON 文字列内でも、実際の改行を `"\n"` として渡せます。結果は同じです。

迷ったら、書き込み後に `view text` して文字単位で突き合わせてください。

**逐次実行。** `officecli` は呼び出しのたびにファイルを変更します。コマンドは 1 つずつ実行し、それぞれの終了コードを確認してください — 50 個のコマンドから成るスクリプトが 3 番目で失敗すると、その影響は静かに連鎖します。構造的な操作（新しいスタイル、表、目次、セクション区切り）の後は、次を積み重ねる前に `get` で確認してください。

**open/save のライフサイクル:** 最初に `officecli open <file>`、最後に `officecli save <file>` でディスクにフラッシュします — `save` は書き込みだけを行い、後続編集のためにレジデントプロセスを起動したままにします。ワンショットの引き渡しでレジデントプロセスを解放したい場合にのみ `officecli close <file>` を使ってください。どちらも常に安全です（エラーになったり作業内容を失ったりすることはありません）。同じスタイルの段落が多数ある場合は `batch` を使ってください（配列全体を 1 パスで処理）。**非 officecli の境界でのみフラッシュしてください:** officecli 自身の読み取りは常にあなたの編集を反映しますが、`save`/`close` は officecli 以外のプログラムがそのファイルを読む直前（python-docx、Word、レンダラー、納品など）にのみ実行してください。

**`$FILE` の慣習。** すべてのコマンドは `"$FILE"` を使用します — 一度だけ設定してください（`FILE="your-doc.docx"`）。リテラルな `doc.docx` / `review.docx` を出力に混ぜてはいけません — 常に実際の対象ファイルに置き換えてください。

## 成果物の必須要件

すべての文書が満たすべき納品基準 — コマンドに手を伸ばす前に把握しておいてください。

**明確な階層。** どんな軽微でない文書にも、タイトル → 見出し1 → 見出し2 → 本文という構造があり、無地の `Normal` 段落の壁ではありません。`view outline` がフラットな1本のリストを示す場合、階層が欠けています。

**明示的な見出しサイズ**（Word の既定スタイルサイズはテンプレート間でぶれます）: **H1 は 18pt 以上**（長文レポートでは 20pt）、H2 = 14pt 太字、H3 = 12pt 太字、本文 = 11〜12pt、行間 1.15〜1.5 倍。インラインでサイズを指定するより `style=Heading1` を優先してください。そうすればリテーマの際にスタイル定義を1回変えるだけで済みます — ただし、テンプレートのスタイルを信頼できない場合は明示的なサイズを設定してください。

**本文フォントは1種類、アクセントも1種類。** 読みやすい本文フォント（Calibri、Cambria、Georgia、Times New Roman）を1つ選び、見出しの強調や表のヘッダーにはアクセントカラーを使う — 虹色の装飾はしないでください。

**間隔はプロパティで制御。** 段落には `spaceBefore` / `spaceAfter` を使ってください。空の段落を連ねると改ページが崩れ、`view issues` に指摘されます。

**タイポグラフィの品質。** 新規コンテンツにはカーリークォート（`'` `'` `"` `"`）を使い、ASCII クォートは使いません — Unicode を直接使うか、`raw-set` の中では XML エンティティ（`&#x2018;`/`&#x2019;`/`&#x201C;`/`&#x201D;`）を使います。範囲にはエンダッシュ `–`（`2024–2026`）、挿入句にはエムダッシュ `—` を使います。

**1ページを超える文書にはヘッダー、フッター、ページ番号を。** ページ番号は必ずライブな `PAGE` フィールド（`--prop field=page`）で表現し、リテラルな "Page 1" というテキストにはしません — CLI が `<w:fldChar>` を代わりに挿入します（ヘッダーとフッターの節を参照）。

**既存テンプレートを尊重する。** すでに見た目が決まっているファイルを編集する場合は、それに合わせてください — 既存の慣習は、ここでのガイドラインより優先されます。

### 視覚的な納品ライン（すべての文書に適用）

完了と宣言する前に `officecli view "$FILE" html` を実行し、返された HTML パスを Read して、以下のすべてを確認してください:

- **プレースホルダートークンがデータとしてレンダリングされていない。** `$xxx$`、`{var}`、`{{name}}`、`<TODO>`、`lorem`、`xxxx` は、見出し・本文・表紙・目次・キャプション・ヘッダー・フッターのいずれにも決して現れてはいけません。人間に埋めてもらうためのリテラルな `{name}` は、目に見える指示段落の中に置くべきで（例: "送付前に `{name}` を置き換えてください"）、完成したコンテンツとして扱ってはいけません。
- **タイトルが途切れていたり、セルがはみ出していたりしない。** コンテンツを削るのではなく、列幅を広げるか `wrapText` を設定してください。
- **見出しが3つ以上ある文書には目次が存在する**（`--type toc`）。
- **表紙は60%以上、最終ページは40%以上埋まっている。** 薄い表紙にはサブタイトル／著者／日付／範囲／主要ハイライトを補い、「ありがとうございました」の最終ページには結論／次のステップ／連絡先／法的表記を補ってください。
- **文書テキストに `\$`、`\t`、`\n` のリテラルが残っていない。** `view text` にこれらが表示される場合、シェルエスケープの層が漏れています — 該当段落を削除し、入力し直してください。

いずれかが失敗した場合は、完了宣言の前に立ち止まって修正してください。

## 一般的なワークフロー

6ステップです。軽微でないビルドはすべてこの形をたどります。

1. **ファイルを開く。** `officecli open "$FILE"`（レジデントがデフォルト）。新規ファイルの場合は先に `officecli create "$FILE"` を実行します。
2. **状況を把握する。** 既存ファイルの場合: `officecli view "$FILE" outline` — 見出しツリー、セクション数、目次／透かし／変更履歴が既に存在するかどうか。盲目的に編集してはいけません。
3. **段階的に構築する。** 構造 → コンテンツ → フォーマットの順: スタイル・番号定義 → セクション／ページ設定 → 見出しと本文 → 表／画像／フィールド／目次 → ヘッダー／フッター → コメント。各構造操作の後、次を積み重ねる前に `get` で確認してください。
4. **仕様どおりにフォーマットする。** 明示的な見出しサイズ、間隔、幅、配置、タブ、リストのインデント — フォーマットは成果物の一部であり、任意の仕上げではありません。
5. **保存し、その後はキャッシュされたテキストより構造を信頼する。** `officecli save "$FILE"` が XML を書き込みます。TOC / PAGE / NUMPAGES / SEQ / PAGEREF の各フィールドは**キャッシュされた値**を保持しており、人間が再計算する（Word 上での F9）まで古い値または空のままかもしれません。表示されているテキストを信頼するのではなく、フィールドが *存在する* こと（`get --depth 3` で `<w:fldChar>` が見つかるか）を確認してください。
6. **QA — 問題があると想定する。** 1回の「修正して検証」サイクルで新たな問題が見つからなくなった時点で完了であり、最後のコマンドの終了コードが 0 だったからといって完了ではありません。QA の節を参照してください。

## クイックスタート

最小限の docx: 見出し、本文段落、サブ見出し、そしてライブなページ番号フィールドを持つフッター。そのままコピペせず、あなた自身のファイル・コンテンツに合わせて調整してください。

```bash
FILE="review.docx"
officecli create "$FILE"
officecli open "$FILE"
officecli add "$FILE" /body --type paragraph --prop text="Q4 2026 Review" --prop style=Heading1 --prop size=20pt --prop bold=true --prop spaceAfter=12pt
officecli add "$FILE" /body --type paragraph --prop text="Revenue grew 18% year-over-year, ahead of plan." --prop size=11pt --prop spaceAfter=8pt
officecli add "$FILE" /body --type paragraph --prop text="Key Drivers" --prop style=Heading2 --prop size=14pt --prop bold=true --prop spaceBefore=12pt --prop spaceAfter=6pt
officecli add "$FILE" /body --type paragraph --prop text="Enterprise renewals, upsell, and a new EMEA region." --prop size=11pt
officecli add "$FILE" / --type footer --prop type=default --prop size=9pt --prop text="Page " --prop field=page
officecli set "$FILE" "/footer[1]/p[1]" --prop align=center
officecli save "$FILE"
officecli validate "$FILE"
```

確認済み: `validate` は `no errors found` を返します。`get /footer[1] --depth 3` は 5-run 構成の PAGE フィールドチェーン（begin / instrText / separate / キャッシュ値 / end）を示します。

## 読み取りと解析

まず広く見渡し、そこから絞り込みます。`outline` は既に何があるかを教えてくれます。どこを見ればよいか分かったら `view text` / `get` / `query` に進んでください。

```bash
officecli view "$FILE" outline            # 見出しツリー、セクション数、表/画像の数、透かし、変更履歴の有無 — まずここで状況を把握
officecli view "$FILE" html               # 返された HTML パスを Read: 一連の編集の後の最初の視覚チェック（階層、空段落の間隔、目次の欠落）
officecli view "$FILE" text --start 1 --end 80   # コンテンツ QA 用のテキスト。パスは [/body/p[N]] の形で表示され、get で戻れる
officecli view "$FILE" annotated          # 値 + スタイル/フォント/サイズ + ラン単位の警告
officecli view "$FILE" stats              # 段落数、フォント使用状況、スタイル分布
officecli view "$FILE" issues             # 空段落、代替テキストの欠落、間隔の異常
```

`officecli watch "$FILE"` は、人間のユーザーが任意のタイミングで開けるようライブプレビューを起動し続けます — エージェントの自己確認には `view html` を使ってください。最終的な視覚検証は、ユーザーが Word / WPS / Pages で `.docx` を開くことです。

**要素を1つ確認する。** XPath 風のセマンティックパス（1-based）。シェルが `[N]` をグロブ展開するため、常にクォートしてください。最後の要素には（括弧付きの）`[last()]` を使ってください。`[last]` はエラーになります。機械可読出力には `--json` を付けてください。

```bash
officecli get "$FILE" /                          # 文書ルート: メタデータ、ページ設定
officecli get "$FILE" "/body/p[1]"                # 1つの段落
officecli get "$FILE" "/body/p[1]/r[1]"           # 1つのラン（文字レベルのフォーマット）
officecli get "$FILE" "/body/tbl[1]" --depth 3    # 行とセルを含む表
officecli get "$FILE" "/footer[1]" --depth 3      # フッター — fldChar の有無を確認
officecli get "$FILE" "/styles/Heading1"          # スタイル定義
officecli get "$FILE" /numbering --depth 2        # 番号設定の abstractNum と num のバインディング
```

**文書全体を横断してクエリする。** CSS ライクなセレクタで、手作業でたどるのではなく体系的にチェックできます。演算子: `=`、`!=`、`~=`（部分一致）、`>=`、`<=`、`[attr]`（存在確認）。完全なリファレンス: `officecli query --help`。

```bash
officecli query "$FILE" 'paragraph[style=Heading1]'       # すべての H1
officecli query "$FILE" 'p:contains("quarterly")'         # テキスト一致
officecli query "$FILE" 'p:empty'                         # 空段落（不要な残骸）
officecli query "$FILE" 'image:no-alt'                    # アクセシビリティの欠落箇所
officecli query "$FILE" 'paragraph[size>=24pt]'           # 数値比較
officecli query "$FILE" 'field[fieldType!=page]'          # PAGE 以外のフィールド
```

`query --json` は結果を `.data.results[]` にラップします — 件数を数えるには `jq '.data.results | length'`。

**大規模文書。** `view outline` で見出しをたどり、`query` でジャンプしてください。本文全体をコンテキストに流し込まないでください。

## 作成と編集

動詞: `add`（新規要素）、`set`（プロパティ変更）、`remove`、`move`、`swap`、`batch`、`raw-set`（最終手段の XML）。ビルドの9割は段落、ラン、表、いくつかの画像、目次、フッターです。

### 段落、ラン、スタイル

段落（`p`）はブロックであり、ラン（`r`）はその中で一貫した文字フォーマットを持つ範囲です。段落レベルのプロパティ（スタイル、配置、間隔、インデント）は `p` に、フォント／サイズ／色／太字は `r` に設定してください。

```bash
officecli add "$FILE" /body --type paragraph --prop text="Executive Summary" --prop style=Heading1 --prop size=18pt --prop bold=true --prop spaceAfter=12pt
officecli set "$FILE" "/body/p[1]/r[1]" --prop color=1F4E79
```

垂直方向の間隔には `spaceBefore` / `spaceAfter` を使ってください — 空段落を連ねるのは禁物です。左インデントには `--prop indent=720`（twips）、最初の行には `firstLineIndent=360`、ぶら下げインデントには `hangingIndent=720` を使ってください。先頭のスペースは `view issues` に検出されます。

### 表

表は `/body/tbl[N]` で、行は `tr[N]`、セルは `tc[N]` です。行数・列数を指定して追加し、その後埋めていきます。

```bash
officecli add "$FILE" /body --type table --prop rows=4 --prop cols=3 --prop width=100%
officecli set "$FILE" "/body/tbl[1]/tr[1]" --prop header=true --prop c1=Quarter --prop c2="Revenue" --prop c3="Growth"
officecli set "$FILE" "/body/tbl[1]/tr[1]/tc[1]/p[1]/r[1]" --prop bold=true
```

行レベルの `set` は `height`、`header`、そして `c1 / c2 / … / cN` のテキストショートカット（`cN` は任意の列数に一般化されます）に対応しています。セルのフォーマット（太字、塗り、色）はセルの段落／ランに対して設定します — 行レベルには**設定できません**。セル単位の罫線には、`tc` に対してセルレベルの `border.*`（`--prop border.bottom="single;6;000000;0"`）を、または内側の段落に対して段落レベルの `pbdr.*` を設定してください。

**水平線 = 段落の下罫線であり、1行の表ではありません。** 罫線代わりの表は、最小高さの空の枠としてレンダリングされます（ヘッダー／フッター内では特に悪化します）。代わりに段落に `pbdr.bottom`（`STYLE;SIZE;COLOR`）を使ってください:

```bash
officecli set "$FILE" "/body/p[3]" --prop pbdr.bottom="single;6;2E75B6"
```

### リスト（箇条書き、番号付き、マルチレベル）

単一レベルの箇条書き／番号付きリストには、段落に `listStyle` を設定してください（`listStyle` は段落プロパティであり、ランのプロパティでは**ありません** — よくある間違いです）:

```bash
officecli add "$FILE" /body --type paragraph --prop text="First item" --prop listStyle=bullet
```

マルチレベル（法律文書スタイルの 1 / 1.1 / 1.1.1）の場合は、`abstractNum` を追加し、次に `num` を追加し、段落ごとに `numId` を参照してください:

```bash
officecli add "$FILE" /numbering --type abstractnum --prop format=decimal     # → abstractNum id=0
officecli add "$FILE" /numbering --type num --prop abstractNumId=0             # → num id=1
officecli add "$FILE" /body --type paragraph --prop text="Section one" --prop numId=1 --prop ilvl=0
```

ID は 0 始まりです: 最初の `abstractNum` は id=0 であり、`num` は `abstractNumId=0` でそれを参照し、自身は id=1 として割り当てられます。存在しない `abstractNumId` を指定するとエラーになるため、作成後は必ず ID を確認してください。`officecli query "$FILE" 'paragraph[numId>0]'` で検証してください。レベルとフォーマットのオプションは `help docx abstractnum` / `help docx num` を参照してください。

### タブストップ（署名欄、リーダー付きの行）

タブストップは段落の第一級の子要素 `tab` です。`pos` は `6in`/`6cm`/twips を受け付け、`val` は `left`/`center`/`right`、`leader` は `none`/`dot`/`hyphen`/`underscore` のいずれかです。`help docx tab` を参照してください。

```bash
officecli add "$FILE" "/body/p[1]" --type tab --prop pos=6in --prop val=right --prop leader=dot
```

**リーダーの注意点。** `leader=dot` を指定するだけではドットは出力されません — リーダーは、テキストとタブストップの間にあるラン内に実際の `<w:tab/>` 文字が存在する場合にのみレンダリングされます。テキスト中の `\t` でそれを配置してください: タブストップを定義し（`add tab --prop pos=6in --prop val=right --prop leader=dot`）、その後 `--prop text="Chapter 1\t12"` とします — `\t` が `<w:tab/>` になり、右揃えのページ番号までドットが埋まります。（リテラルな `text="Chapter 1 ......... 12"` も動作はしますが、実際のタブストップの方がきれいに整列します。）

### フィールド（PAGE / NUMPAGES / DATE / MERGEFIELD / REF）

フィールドはレンダリング時に計算されるライブな値です。`fieldType` がフィールドを決め、`name` が対象（マージ名または `ref` ブックマーク）を与え、`format` / `instr` がスイッチを追加します。

| フィールド | 用途 | 例 |
|---|---|---|
| `page` | 現在のページ番号 | フッターで `--prop field=page`、インラインでは `--prop fieldType=page` |
| `numpages` | 総ページ数 | `--prop field=numpages` / `--prop fieldType=numpages` |
| `date` | 今日の日付 | `--prop fieldType=date --prop format='yyyy-MM-dd'` |
| `mergefield` | テンプレートのマージトークン | `--prop fieldType=mergefield --prop name=CustomerName` |
| `ref` | ブックマークへの相互参照 | `--prop fieldType=ref --prop name=bookmarkName` |

完全な `fieldType` の列挙（`pageref`、`seq`、`styleref`、`docproperty`、`createdate` などを含む30種類以上）は `help docx field` にあります。**`fieldInstr` という fieldType は存在しません** — 型付きショートカットで足りない場合は、生のフィールド命令文を指定するために `instruction` プロパティを使ってください。ピクチャースイッチ（`MERGEFIELD Amount \# "#,##0.00"`、`DATE \@ "yyyy年MM月"`）は `--prop instruction='…'` 経由で指定します（mergefield の `format` プロパティは警告付きで無視されます — `instruction` を使ってください）。

```bash
officecli add "$FILE" "/body/p[3]" --type field --prop fieldType=mergefield --prop name=customer_name
# «customer_name» としてレンダリングされる — 目に見えるプレースホルダーで、差し込み印刷時に Word 上で置き換えられる。
```

**MERGEFIELD テンプレート: プレースホルダーのリテラルを決して本文としてレンダリングしないでください。** 本文テキストとして表示された `{{customer_name}}` や `$NAME$` は、受け取った側に見えてしまう失敗したテンプレートです — 実際の MERGEFIELD を挿入するか（上記参照）、リテラルなトークンは明示的な指示段落の中に限定してください。`query 'field[fieldType=mergefield]'` で確認してください。

**SEQ / PAGEREF / TOC フィールドの値。** officecli は書き込み時にレンダリング済みのフィールド値を保持しません。それぞれのパスが必要とする方法で再計算してください:

- **SEQ 番号付け**（`Figure 1/2/3`）: `officecli set "$FILE" / --prop recalcFields=seq` は本文の文書順で SEQ フィールドを数え、キャッシュされた値を書き込みます（`evaluated` が true に切り替わります。スイッチ／フォーマットは `help docx document` を参照）。見出し相対の `\s` やヘッダー／フッター内の SEQ は Word に委ねられます。
- **PAGE / PAGEREF / NUMPAGES / TOC のページ番号**にはページネーションが必要で、officecli にはそのためのエンジンがありません — `officecli set "$FILE" /settings --prop updateFields=true` で、開いた際に Word へ委ねます。

複数の図を含む文書では両方を使ってください。学術論文については `officecli-academic-paper` スキルを参照してください。

### ヘッダーとフッター（ページ番号付け）

単一コマンドのパターン — CLI が `<w:fldChar>` を挿入するため、フィールドを手で組み立てる必要は一切ありません:

```bash
# 空の1ページ目フッター — differentFirstPage を自動的に有効化し、表紙にページ番号が出ないようにする
officecli add "$FILE" / --type footer --prop type=first --prop text=""
# ライブなページ番号付きのデフォルトフッター
officecli add "$FILE" / --type footer --prop type=default --prop align=center --prop size=9pt --prop text="Page " --prop field=page
```

両方が存在する場合、デフォルトフッターは `/footer[2]` になります。単独なら `/footer[1]` です。**確認方法**: `get --depth 3` は、リテラルな `"Page"` というテキストのランだけでなく `fldChar` の子要素を示すはずです（`view outline` はライブフィールドと静的テキストの両方に対して "Footer: Page" と表示するため、これに頼ってはいけません）。`set --prop differentFirstPage=true` は実行しないでください — このプロパティはサポートされておらず（静かにではなく終了コード2で拒否されます）、first タイプのフッターを追加すればそのビットは自動的に切り替わります。合成的な **Page X of Y** については、レシピ (b) を参照してください。

### 目次

見出しが3つ以上ある文書では、まず要求された範囲に実際のソースがあることを確認してください。目次のソースは、組み込みの Heading スタイルか、`outlineLvl` を持つカスタム段落スタイルでなければなりません。太字や大きな Normal テキストはソースにはなりません。有資格なソースが存在しない場合は、"Error! No table of contents entries found." と表示される目次を挿入するのではなく、いったん立ち止まって見出し構造を修正してください。

```bash
# 組み込みの Heading スタイルは目次のソースになる。
officecli add "$FILE" /body --type paragraph --prop text="Introduction" --prop style=Heading1
# カスタム段落スタイルにはアウトラインレベルが必要（0 = Heading 1）。
officecli add "$FILE" /styles --type style --prop id=ThesisH1 --prop type=paragraph --prop outlineLvl=0
# ソースが存在した後にのみ目次を追加する。
officecli add "$FILE" /body --type toc --prop levels="1-3" --prop title="Table of Contents" --prop hyperlinks=true --index 0
```

ページ番号にはページネーションが必要であり、OfficeCLI は目次のページ番号を自身で計算できません。`updateFields=true` を設定して、開いた際に Word が目次（および全フィールド）を再計算するようにしてください。Word 互換のフィールドエンジンがないため、目次は動的で未計算であると報告してください。ページ番号が確定していると主張したり、静的な目次を推測で埋めたりしないでください。

```bash
officecli set "$FILE" /settings --prop updateFields=true
```

`get`/`set`/`remove` では、目次は `/toc[1]` または `/tableofcontents` で直接アドレス指定できます。

### 画像

画像はラン内に配置されます。アクセシビリティのため代替テキストは必須です — 作成時に直接 `alt` を渡してください:

```bash
officecli add "$FILE" "/body/p[5]" --type picture --prop src=logo.png --prop width=1.5in --prop alt="Acme logo"
```

納品前に `officecli query "$FILE" 'image:no-alt'` が空であることを確認してください。

### チャート

データを扱う場合は、チャートのフラットな PNG スクリーンショットではなく、**ネイティブなチャート**（編集可能・テーマ対応・アクセシブル・Word 上で再レンダリング可能）を追加してください。系列ごとに `data="Label:v1,v2,…"`、系列1つにつき `data=` を1つ（または `series1=`/`series2=`）指定します。

```bash
officecli add "$FILE" /body --type chart --prop chartType=bar --prop title="Revenue by Region" --prop categories="EMEA,APAC,Americas" --prop data="2026:120,150,180"
```

`chartType` は bar / column / line / pie / area / scatter のいずれかです（軸・凡例・系列のスタイリングは `help docx chart`）。`--type picture` 経由の PNG は、officecli が構築できない特殊なチャートのフォールバックとしてのみ使ってください。

### ハイパーリンクとブックマーク

外部リンクは `hyperlink` 経由で行います:

```bash
officecli add "$FILE" "/body/p[2]" --type hyperlink --prop url="https://example.com" --prop text="our site"
```

**内部リンク**（ブックマークへのリンク）は、`url` に `#fragment` を使うのではなく `--prop anchor=bookmarkName` を使います:

```bash
officecli add "$FILE" "/body/p[2]" --type hyperlink --prop anchor=chapter1 --prop text="See Chapter 1"
```

`PAGEREF` フィールドと目に見えるテキストを組み合わせるのが代替手段です。`help docx hyperlink` / `help docx bookmark` を参照してください。

### セクションとページ設定

文書ルート `/` はページ設定（`pageWidth`、`pageHeight`、余白、単位は twips）を保持します。マルチセクション文書（横向きの挿入、段組み）には `section` 区切りを追加します — `help docx section` を参照してください。camelCase（`pageWidth`、正規形）と小文字エイリアス（`pagewidth`）の両方が受け付けられますが、camelCase を優先してください。

```bash
officecli set "$FILE" / --prop pageWidth=12240 --prop pageHeight=15840 --prop marginTop=1440 --prop marginLeft=1440
# 新聞スタイルのマルチカラムフロー（columnSpace は twips 単位。720 = 0.5in）:
officecli set "$FILE" / --prop columns=2 --prop columnSpace=720
```

### 強制的な改ページ

論理的な境界1つにつき、改ページの方法はちょうど1つだけ使ってください。デフォルトは対象の見出しに `pageBreakBefore=true` を設定することです。代わりに、見出しの前に明示的な `pagebreak` を1つ挿入することもできます。同じ境界に両方の方法を組み合わせては絶対にいけません — 結果として二重の改ページが発生し、白紙のページが生まれることがあります。`pagebreak` を追加した直後の `p[last()]` に `pageBreakBefore` を設定しないでください。

```bash
# デフォルト: 見出しに直接改ページを適用する。
officecli add "$FILE" /body --type paragraph --prop text="Introduction" --prop style=Heading1 --prop pageBreakBefore=true
```

`--prop break=newPage` は `pageBreakBefore=true` のエイリアスです。`view html` でプレビューし、白紙のページがないか確認してください。

### レポートレベルのレシピ

長文レポートで毎回出てくるパターンです。それぞれ実行済みで `validate` を通過しています。

**(a) リッチな表紙 — 60%以上の充填ラインを満たす。** 機密バナー、タイトル、サブタイトル、クライアント／プロジェクト／日付ブロック、キーテーマの帯を積み重ね、次のセクションを強制的に新しいページへ送ります:

```bash
officecli add "$FILE" /body --type paragraph --prop text="CONFIDENTIAL — CLIENT USE ONLY" --prop align=center --prop size=9pt --prop color=C00000 --prop spaceAfter=24pt
officecli add "$FILE" /body --type paragraph --prop text="Strategic Growth Review" --prop style=Title --prop size=32pt --prop bold=true --prop align=center --prop font=Cambria --prop spaceAfter=8pt
officecli add "$FILE" /body --type paragraph --prop text="FY26 Outlook and Scenario Planning" --prop italic=true --prop size=16pt --prop align=center --prop spaceAfter=36pt
officecli add "$FILE" /body --type paragraph --prop text='Prepared for: Acme Corp. Leadership Team' --prop align=center --prop size=11pt
officecli add "$FILE" /body --type paragraph --prop text='Engagement: 2026-04 — 2026-06' --prop align=center --prop size=11pt --prop spaceAfter=36pt
officecli add "$FILE" /body --type paragraph --prop text="Key themes: 1) margin resilience, 2) EMEA expansion, 3) capital allocation." --prop align=center --prop italic=true --prop size=10pt
officecli add "$FILE" /body --type paragraph --prop text="Executive Summary" --prop style=Heading1 --prop pageBreakBefore=true
```

**(b) Page X of Y のフッター — PAGE と NUMPAGES の合成。** フッター段落を追加し、その後3つの子操作でライブに `Page <X> of <Y>` を組み立てます。公式の `help docx footer` レシピです。

```bash
officecli add "$FILE" / --type footer --prop type=default --prop text="Page " --prop align=center --prop size=9pt
officecli add "$FILE" "/footer[1]/p[1]" --type field --prop fieldType=page
officecli add "$FILE" "/footer[1]/p[1]" --type run --prop text=" of "
officecli add "$FILE" "/footer[1]/p[1]" --type field --prop fieldType=numpages
officecli get "$FILE" "/footer[1]/p[1]" --depth 1 | grep -o fldChar | wc -l   # 4以上を期待。grep -c ではなく grep -o ... | wc -l を使うこと（1行の XML では grep -c は 1 を返してしまう）
```

**(c) 塗りつぶしと白い太字テキストを持つヘッダー行。** 順序が重要です — まずヘッダーセルのテキストを入力し（空のセルにはランが存在しないため、空セルに `set …/tc[N]/p[1]/r[1]` を実行すると "No r found" エラーになります）、次にセルの塗り、最後にランのフォーマットを設定します:

```bash
officecli add "$FILE" /body --type table --prop rows=5 --prop cols=4 --prop width=100%
officecli set "$FILE" "/body/tbl[1]/tr[1]" --prop header=true --prop c1=Quarter --prop c2=Revenue --prop c3=Growth --prop c4=Status
for col in 1 2 3 4; do
  officecli set "$FILE" "/body/tbl[1]/tr[1]/tc[$col]" --prop fill=1F4E79
  officecli set "$FILE" "/body/tbl[1]/tr[1]/tc[$col]/p[1]/r[1]" --prop bold=true --prop color=FFFFFF
done
for row in 3 5; do for col in 1 2 3 4; do
  officecli set "$FILE" "/body/tbl[1]/tr[$row]/tc[$col]" --prop fill=D9E2F3      # ゼブラストライプ
done; done
```

**(d) 財務表 — 数値を右揃え、合計を太字、合計行に下罫線。**

```bash
for row in 2 3 4 5; do for col in 2 3 4; do
  officecli set "$FILE" "/body/tbl[1]/tr[$row]/tc[$col]/p[1]" --prop align=right
done; done
for col in 1 2 3 4; do
  officecli set "$FILE" "/body/tbl[1]/tr[5]/tc[$col]/p[1]/r[1]" --prop bold=true
  officecli set "$FILE" "/body/tbl[1]/tr[4]/tc[$col]/p[1]" --prop pbdr.bottom="single;6;000000;0"
done
```

**(e) 複数の箇条書きを持つセル（SWOT／リスクマトリクス）。** `c1="a\nb"` は**1つの**段落内に `<w:br/>` の改行を作ります — 単純な複数行テキストには問題ありませんが、箇条書きには別々の段落が必要です。最初の項目を `set c1=` で仕込み、以降の箇条書きごとにセル配下へ（`listStyle=bullet` を伴う）`add paragraph` を実行してください:

```bash
officecli set "$FILE" "/body/tbl[1]/tr[1]" --prop c1="Installed base of 18k enterprise seats"
officecli add "$FILE" "/body/tbl[1]/tr[1]/tc[1]" --type paragraph --prop text="Margin structure above peer median" --prop listStyle=bullet
officecli set "$FILE" "/body/tbl[1]/tr[1]/tc[1]/p[1]" --prop listStyle=bullet
```

仕込んだ行が下に来てしまう場合は、並べ替えてください: `officecli move "$FILE" "/body/tbl[1]/tr[1]/tc[1]/p[N]" --index 0`。

**(f) フィールドエンジンなしでの目次。** CLI 単体のパイプラインでは、信頼できるページ番号は計算できません。ライブな目次のまま `updateFields=true` を設定し、Word または他の互換フィールドエンジンで文書を開くよう受け取り側に伝えてください。推測による静的なページ番号に置き換えないでください。

### テンプレート納品 — テンプレートノートとエンドユーザー向けコンテンツの分離

人事／法務／ベンダーのテンプレートには、納品してはいけない社内限定のガイダンス（"`{{CompanyName}}` を置き換える" など）が含まれています。実用的なパターンは2つあります:

- **末尾の「Template Notes」セクション**を明確な `Heading 1`（"Template Notes for HR Users"）の下に置き、その下にすべての指示を記載します。配布前に、見出しから下を `remove` してください（`query 'paragraph[style=Heading1]:contains("Template Notes")'` で位置を特定）。
- **ブックマークで区切られた社内セクション** — `__template_notes_start` / `_end` のブックマークの間。納品時に `raw-set` でアンカー間のすべてを削除します。

テンプレートの納品ゲート: 削除後、`query 'p:contains("Template Notes")'` と `query 'p:contains("{{")'` の両方が空を返すこと。ノート段落が1つでも残っていれば、下流の従業員が社内向けの文言を読んでしまいます。

### 高度／専門トピック（レポートを書いているだけなら読み飛ばして構いません）

レポート、メモ、手紙、提案書、人事テンプレートにはこの節は不要です。文書が学術的（数式、脚注、参考文献）、レビュー対象（コメント、変更履歴）、またはマーク付き（透かし）である場合にのみ読み進めてください。

**数式と脚注。** `--type equation` は LaTeX を受け付けます — `\frac`、`\sum`、ギリシャ文字、`\mathit`、`\mathcal` はすべてレンダリングされます。デフォルトでは独立した `/body/oMathPara[N]` の表示ブロックを作成します。段落の親パスを指定して `--prop mode=inline` を渡すと（`add "/body/p[N]" --type equation --prop formula=… --prop mode=inline`）、地の文にインラインの `<m:oMath>` を挿入します。脚注は段落インデックスで自動採番されます。参考文献のぶら下げインデント: 各エントリに `firstLineIndent=-720 indent=720`。

```bash
officecli add "$FILE" /body --type equation --prop formula="\\frac{a}{b} + \\sum_{i=1}^{n} x_i"
officecli add "$FILE" "/body/p[3]" --type footnote --prop text="See Appendix A for methodology."
```

**コメントと変更履歴。** 一括承認／却下: `set "$FILE" /revision --prop revision.action=accept`（または `--prop revision.action=reject`）。`/revision[@author=Alice]` や `/revision[@type=ins]` のようなセレクタで絞り込みます。個々の変更箇所は `query ins` と `query del` で特定してください（`trackedchange` はセレクタではありません）。ランに対して変更履歴を作成するには `--prop revision.type=ins|del --prop revision.author=…`（`revision.*` の完全な一覧は `help docx run` — `format`/`moveFrom`/`moveTo` も含む）。コメントの追加: `add "/body/p[4]" --type comment --prop author=… --prop text=…`。`--prop parentId=N` で返信スレッド化し、`set "/comments/comment[N]" --prop done=true` で解決済みとしてマークします（監査証跡を残すため、削除ではなく解決してください — `query 'comment[done=false]'` で未解決のものが一覧できます）。プロパティスキーマ: `help docx comment` / `help docx run`。

**透かし。** `add / --type watermark --prop text="DRAFT" --prop color=BFBFBF --prop opacity=0.8` を1コマンドで（デフォルトの不透明度は 0.5）。`set /watermark --prop opacity=…` で後から調整できます。

**スキルを切り替えるタイミング。** 章立ての草稿、脚注3個以下、数式2個以下、参考文献／相互参照なしの場合は docx にとどまってください。引用スタイル（APA / Chicago / IEEE / GB 7714）、本文と参考文献の自動リンク、`\ref` を伴う番号付き数式、「図表一覧」、自動更新される相互参照が必要な場合は **`academic-paper`** に切り替えてください。文書の目的が**データ収集**（入力可能なフォーム、ユーザー入力欄のある契約書、アンケート、差し込み印刷テンプレート — `<w:sdt>` コンテンツコントロール、`<w:ffData>`、`documentProtection=forms`）である場合は **`officecli-word-form`** に切り替えてください。

### Raw-set のエスケープハッチ（L1 / L2 / L3）

3段階の精度があります。目的を果たす最も低いレベルを使ってください。

- **L1 — 高レベルプロパティ**（`--prop text=…`、`--prop style=Heading1`）: デフォルトの選択肢。80%をカバーします。
- **L2 — ドット付き属性フォールバック**（`pbdr.top=`、`ind.left=`、`shd.fill=`、`padding.top=`、`font.size=`）: L1 に必要なノブがない場合。例: `--prop pbdr.bottom="single;6;1F4E79;0"`。スキーマとして有効な XML を出力します。
- **L3 — XML を伴う `raw-set`**: 最終手段であり、スキーマによる保護はありません。内部ハイパーリンク、合成フィールドなど、型付き動詞で表現できない形状にのみ使ってください（XML 付録を参照）。

罫線は `style;size;color;space` の形式を使います: `single;4;FF0000;1`。16進数カラーは `#` で始めません: `FF0000`。配色テーマ名（`accent1..6`、`dark1`/`dark2`、`light1`/`light2`、`hyperlink`）は 16進数カラーが使える場所ならどこでも使えます — テーマをまたいで色を安定させたい場合は 16進数を優先してください。

## QA（必須）

**問題があると想定してください — QA は確認作業ではなくバグハントです。** 最初の文書はほぼ確実に正しくありません。初回検査で問題ゼロなら、それは十分に注意深く見ていないということです。見出しは、`view outline` が H1 の直下に H3 があることを示すまでは問題なさそうに見えます。フッターは、`get --depth 3` が静的なランでありフィールドではないことを明らかにするまでは "Page 1" と表示されているように見えます。

### 「完了」と言う前の最小サイクル

1. `officecli view "$FILE" issues` — 空段落、代替テキストの欠落、フォーマットの異常。
2. `officecli view "$FILE" outline` — 見出し階層（H1 → H3 の飛び越えがないか）、目次の有無、セクション数。
3. `officecli view "$FILE" text --max-lines 400` — 誤字、迷い込んだ `\$`/`\t`/`\n` リテラル、プレースホルダートークン。
4. `officecli validate "$FILE"` — スキーマチェック（納品ゲートは、閉じられてディスク上にあるファイルに対してこれを再実行します）。
5. **視覚的パス — 文書全体をコンタクトシートとして**（ビジョン対応エージェントのみ — 画像を解釈できない場合はこのステップを飛ばしてください。ステップ1〜4があなたの上限であり、引き渡し時に文書を「視覚的に未検証」とフラグしてください）。`officecli view "$FILE" screenshot --grid auto -o /tmp/sheet.png` を実行し、それを Read してください。`--grid auto` は**すべてのページ**を1枚の画像にタイル状に並べます（列数は自動。強制するには `--grid 4`） — DOM だけでなく、ページネーション、白紙のページ、見出しのリズム、不揃いな余白、目次／表紙の配置を実際に「見る」ことができます。Windows+Word では各ページが実際の Word を通じてレンダリングされ、それ以外では HTML です。スクリーンショットが失敗した場合は `view html` にフォールバックし、ページをまたぐ改ページ／整列／リズムを「視覚的に未検証」としてフラグしてください。サムネイルは問題箇所の**特定**にのみ使ってください: 細かい判断（列の整列、行間、インデント、暗色 on 暗色、キャプションの配置）は、疑わしいページを `screenshot --page N` でフル解像度で確認してください（`--grid` は付けず、Windows 上の実際の Word で）。「validate 通過」は納品ではなく、「本物の文書に見える」ことも納品ではありません。
6. 何か失敗した場合は修正し、**サイクル全体を再実行**してください — 1つの修正が別の問題を生むことはよくあります。

### 納品ゲート（引き渡し前に実行 — いずれかの失敗 = REJECT、納品しないこと）

コピペして `FILE` を設定し、すべてのゲートが OK を出力するまで完了と宣言しないでください。

```bash
FILE="your-file.docx"

# ゲート1 — スキーマ。
officecli close "$FILE" 2>/dev/null
officecli validate "$FILE" | grep -q "no errors found" || { echo "REJECT Gate 1: validate failed"; exit 1; }
echo "Gate 1 OK"

# ゲート2 — トークン漏れ（シェルエスケープ／テンプレートトークン／リテラルな \$ \t \n）。grep -c は決して誤って PASS しない。
LEAK=$(officecli view "$FILE" text | grep -cE '(\$[A-Za-z_]+\$|\{\{[^}]+\}\}|<TODO>|xxxx|lorem|\\[\$tn])')
[ "$LEAK" -eq 0 ] && echo "Gate 2 OK" || { echo "REJECT Gate 2: $LEAK leak line(s)"; officecli view "$FILE" text | grep -nE '(\$[A-Za-z_]+\$|\{\{[^}]+\}\}|<TODO>|xxxx|lorem|\\[\$tn])'; exit 1; }
# Word 互換のフィールドエンジンが更新する前は、目次のプレースホルダーは正当な状態である。目次フィールドと updateFields 設定が構造的に存在することを確認すること。

# ゲート3 — フッターが期待される場合、ライブな PAGE フィールドが存在する。
FLD=$(officecli query "$FILE" 'field[fieldType=page]' --json | jq '.data.results | length')
[ "$FLD" -ge 1 ] && echo "Gate 3 OK" || { echo "REJECT Gate 3: no live PAGE field"; exit 1; }
echo "Delivery Gate PASS"
```

### フィールド／キャッシュ値の抜き取り確認

フィールドは、書き込み時点では古い、または空のままかもしれないキャッシュ値を保持しています — テキストではなく**構造**によって存在を確認してください。

- **フッターの PAGE:** `get /footer[N] --depth 3` は begin / instrText / separate / cached / end のランチェーンを列挙します — 単一の PAGE なら5ラン以上、合成の "Page X of Y" なら11ラン以上。テキスト `"Page"` を持つ単一のランの場合はフィールドが欠落しています — `--prop field=page` で再追加してください。
- **目次:** `get /toc[1] --depth 2` はフィールド構造を示します。ページ番号は再計算されるまで `1 1 1 1` や `Update field to see…` と読めることがあります（§目次 を参照 — `updateFields=true` を設定してください）。
- **MERGEFIELD:** `query 'field[fieldType=mergefield]'` — スロットごとに1つ、他の場所にリテラルな `{{name}}` がないこと。

### 正直な限界

`validate` はスキーマエラーを検出しますが、デザイン上のエラーは検出しません — 見出し階層が誤っていても、Heading 1 を偽装したサイズであっても、本文にプレースホルダートークンが残っていても、表紙のない文書に空の1ページ目フッターがあっても、文書は `validate` を通過できます。コンタクトシートの視覚的パス（`screenshot --grid`）とフィールド構造チェックこそが、検証では捕まえられないものを捕まえる方法です。

### QA の表示に関する注意点（追いかけなくてよいもの）

- `view text` は、実際にレンダリングされる番号に関わらず、番号付きリストのすべての項目に対して `"1."` を表示します — 実際の出力は正しく増分されます。
- `view issues` は、表紙の段落、中央揃えの見出し、リスト項目、参考文献のエントリに対して「本文段落に一行目インデントがない」と指摘しますが — 一行目インデントは APA／学術系の本文テキストにのみ必須であり、ブロックスタイルの一般的なビジネス文書ではこれらは想定どおりです。

## 既知の問題と落とし穴

何かが「壊れているように見える」場合は、追いかける前に原因を特定してください: **[AGENT-ERROR]** 文書が誤っている（修正する）・**[RENDERER-BUG]** 文書は正しく、ビューアが異なる表示をしているだけ（追いかけない）・**[SKILL gap]** スキルがそのルールを教えていなかった（issue を立てる）。

### レンダラーの癖（ビューア横断、[RENDERER-BUG] — 追いかけないこと）

色／フィールド／チャートが壊れていると判断する前に、ユーザーの対象ビューアでファイルを開いてください。そこで正しく見えるなら、それはビューア固有の癖です。

- **PAGE フィールドは、再計算されるまでリテラルな "Page"**（数字なし）としてレンダリングされることがある — `fldChar` の有無で判断し、数字では判断しないこと。
- **目次のキャッシュされたページ番号は、F9 まで "1 1 1 1"** と読めることがある。
- **円グラフ／ドーナツグラフの塗りが、一部のビューアで1色に潰れる**ことがある（列／棒グラフは正しくレンダリングされる）。
- **フォームコントロールのチェックボックスが二重枠でレンダリングされる**ことがある。**OMML 数式のベースライン**はビューア間でずれることがある（XML 自体は同一）。

### よくある落とし穴

| 落とし穴 | 正しいアプローチ |
|---|---|
| `--index` と `[N]` | `--index` は 0-based、`[N]` パスは 1-based |
| 同じ N で複数回 `add --index N` | 挿入のたびに後続コンテンツが下にシフトする。N を使い回すと後の項目が前の項目より前に来てしまう — 逆順で挿入するか、`paraId` を基準にした `move --after/--before` を使う |
| zsh/bash で `[N]` をクォートしない | すべてのパスをクォートする: `"/body/p[1]"` |
| 述語として `[last]` を使う | `[last()]` として括弧を付けなければならない |
| 間隔指定に生の twips を使う | 単位付きの値を使う: `12pt`、`0.5cm`、`1.5x` |
| 間隔調整のための空段落 | `spaceBefore` / `spaceAfter` を使う |
| セルフォーマットのための行レベル `set` | 行レベルの `set` は `height`、`header`、`c1..cN` テキストのみ対応。フォーマットはセルの段落／ランに設定する |
| ランに対する `listStyle` | これは段落プロパティである |
| 先頭スペースによるインデント | `indent=720` / `firstLineIndent=360` / `hangingIndent=720`（ドット付きの `ind.left` / `ind.firstLine` も使える） |
| `set differentFirstPage=true` による表紙のページ番号抑制 | 未サポート — first タイプのフッターを追加する: `--type footer --prop type=first --prop text=""` |
| 新しい章のページが必要 | 見出しに `pageBreakBefore=true` を使う。明示的な `pagebreak` は代替手段としてのみ使い、両方は絶対に併用しない |
| 1つのセル内に複数の箇条書き段落 | `c1="a\nb"` は `<w:br/>` の改行を作る（1段落）。別々の箇条書き段落にはレシピ (e) を使う |
| ドット付き属性で足りる場面での `raw-set` | L3 の raw-set より L2 のドット付き属性を優先する |
| 直後の段落が直前の Heading スタイルを継承する | 次の段落に明示的な `--prop style=Normal` を設定する |
| Word で開いているファイルの変更 | 先に Word 側で閉じる |
| `$`/`'` を含む文字列を batch へ echo すると壊れる | シングルクォートの区切り文字を使ったヒアドキュメント: `cat <<'EOF' \| officecli batch …` |

## Raw-set XML 付録（L3 パターン）

`raw-set` はリテラルな OOXML を注入します — スキーマによる保護はありません。`<w:pPr>` 内の要素順序: `pStyle`、`numPr`、`spacing`、`ind`、`jc`、`rPr`（最後）。スマートクォートはエンティティとして（`&#x2018;`/`&#x2019;`/`&#x201C;`/`&#x201D;`）。先頭または末尾にスペースを持つ `<w:t>` には `xml:space="preserve"` を追加してください。RSID は8桁の16進数です。ユーザーが別の名前を指定しない限り、変更履歴／コメントの作成者は "Claude" としてください。

**変更履歴の挿入／削除の追跡** — 優先すべきはランに対する高レベルの `--prop revision.type=ins|del` です。raw-set は、型付きパスで表現できないもの（他の作成者の変更の却下や復元、以下参照）にのみ使ってください。`<w:r>…</w:r>` 全体を置き換え、ラン内にタグを注入することは決してしないでください。元の `<w:rPr>` を両方にコピーしてフォーマットを保持してください。`<w:del>` の内部では `<w:delText>` を使います（命令文には `<w:delInstrText>`）:

```xml
<w:r><w:t>The term is </w:t></w:r>
<w:del w:id="1" w:author="Claude" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>30</w:delText></w:r></w:del>
<w:ins w:id="2" w:author="Claude" w:date="2026-01-01T00:00:00Z"><w:r><w:t>60</w:t></w:r></w:ins>
<w:r><w:t> days.</w:t></w:r>
```

段落／リスト項目のすべてのコンテンツを削除する場合は、段落マーク自体も削除済みとしてマークしてください（`<w:pPr><w:rPr>` 内の `<w:del/>`）— そうしないと、変更を承認した際に空の段落が残ってしまいます。**他の作成者による挿入を却下する**には、あなたの `<w:del>` を彼らの `<w:ins>` の内側にネストしてください。**彼らの削除を復元する**には、その後ろに `<w:ins>` を追加してください（彼らのものを直接変更しないでください）。

**ブックマークへの内部ハイパーリンク**（高レベルの `--prop anchor=` パスを優先。raw-set はコマンドで表現できないカスタムのラン スタイリングにのみ使う）:

```xml
<w:hyperlink w:anchor="chapter1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>See Chapter 1</w:t></w:r></w:hyperlink>
```

**1つのランに合成フィールド**（例: 単一コマンドのパスでは合成できない2つのフィールド） — `fldChar begin / instrText / separate / value / end` のチェーン:

```xml
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>1</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

**コメントマーカー**は `<w:r>` の兄弟要素であり、その内部には決して置きません（返信スレッド化と解決状態は高レベルの操作です — `--prop parentId=`/`done=`、上記のコメントの節を参照）:

```xml
<w:commentRangeStart w:id="0"/><w:r><w:t>annotated text</w:t></w:r><w:commentRangeEnd w:id="0"/>
<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>
```

開いた際のフィールド再計算は `officecli set "$FILE" /settings --prop updateFields=true` で強制してください（`<w:updateFields w:val="true"/>` を書き込みます。レイアウト依存のフィールド PAGE / PAGEREF / NUMPAGES / TOC ページ番号をカバーします — raw-set は不要です）。SEQ 番号付けには、Word を待たずに正しいキャッシュ値を今すぐ書き込む `set / --prop recalcFields=seq` を優先してください。

### ヘルプへのポインタ

迷ったときは: `officecli help docx`、`officecli help docx <element>`、`officecli help docx <verb> <element>`、エージェント向けには `--json`。ヘルプが正としてのスキーマであり、このスキルは判断のためのガイドです。
