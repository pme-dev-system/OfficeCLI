---
name: officecli-academic-paper
description: "このスキルは、学術スタイルの .docx 出力を作成する際に使用します: 正式な引用スタイル（APA、Chicago、IEEE、MLA）、番号付き数式、図表の相互参照、脚注/文末脚注、参考文献リスト、または複数段組のジャーナルレイアウトを伴うジャーナル論文・学会論文・学位論文の章など。トリガーとなるキーワード: 'research paper'、'journal paper'、'conference paper'、'manuscript'、'thesis'、'APA'、'MLA'、'Chicago'、'IEEE two-column'、'bibliography'、'hanging indent'、'citation style'、'abstract + keywords'、'equation numbering'、'cross-reference'、脚注/文末脚注を伴う論文。出力は単一の .docx ファイル。"
---

# OfficeCLI 学術論文スキル

**このスキルは `officecli-docx` の上に乗るシーンレイヤーです。** docx の基本ルール — スタイル構造、見出し階層、シェルのクォーティング、改ページルール、ライブ PAGE フィールド、Delivery Gate、レンダラーの癖 — はすべて継承されるものであり、再度教える対象ではありません。このファイルが追加するのは、学術論文に固有の内容だけです: 引用スタイル、数式、SEQ / PAGEREF による相互参照、複数段組のジャーナルレイアウト、参考文献のぶら下げインデント、要旨/キーワード/所属ブロック。

docx の基本ルールでカバーされている箇所は、本文中で `→ see docx v2 §X` と表記します。まだ読んでいなければ、先に docx v2 を読んでください。

## セットアップ

`officecli` が未インストールの場合:

- **macOS / Linux**: `curl -fsSL https://d.officecli.ai/install.sh | bash`
- **Windows (PowerShell)**: `irm https://d.officecli.ai/install.ps1 | iex`

`officecli --version` で確認してください（PATH が反映されていない場合は新しいターミナルを開いてください）。インストールに失敗した場合は https://github.com/iOfficeAI/OfficeCLI/releases からバイナリをダウンロードしてください。

## ⚠️ ヘルプ優先ルール

**このスキルが教えるのは学術論文が要求する内容であり、すべてのコマンドフラグではありません。** プロパティ名、列挙値、フィールドの指定が不確かな場合は、推測する前にヘルプを確認してください。

```bash
officecli help docx                          # 全 docx 要素
officecli help docx <element>                # 完全なスキーマ（例: section、equation、field、footnote）
officecli help docx <element> --json         # 機械可読形式
```

ヘルプはインストール済みの CLI バージョンに紐づいています。**このスキルとヘルプが食い違う場合は、ヘルプが優先します。** 本ファイル内のすべての `--prop X=` は `officecli help docx <element>` に対して grep 検証済みです — 後のバージョンでヘルプがプロパティを追加・改名した場合は、ヘルプを信頼してください。

## メンタルモデルと継承関係

**docx v2 を継承します。** 先に `skills/officecli-docx/SKILL.md` を読んでおくべきです。このスキルは、段落の追加、スタイル設定、表の構築、画像の挿入、TOC/フッター/ヘッダーの管理、強制改ページ、Delivery Gate の実行についてすでに理解していることを前提とします。これらのいずれかに馴染みがなければ、続ける前に docx v2 の別セッションを開いてください。

## シェルと実行の規律

**シェルのクォーティング、段階的な実行、`$FILE` 規約** → see docx v2 §Shell & Execution Discipline。同じルールがここでもそのまま適用されます — `[N]` を含むパスはクォートし、`$` を含む値（本文段落中の `$2.8B` や `@` を含む DOI も含む）はシングルクォートで囲み、実行例の中で `\$ \t \n` を手書きせず、コマンドは一度に一つずつ実行してください。以下の学術論文の例では、`$FILE` をシェル変数として使用します（`FILE="thesis.docx"`）。

## ここでの「学術」の意味（アイデンティティ）

学術論文とは、**学術的な層**が上乗せされた docx です: 検証可能な引用、精緻な数式、同期を保つ相互参照、整形された参考文献リスト。docx の基本ルールはそのまま適用され、学術性が加える差分は六つです:

1. **引用スタイルは契約です。** APA / Chicago / IEEE / MLA はそれぞれ、著者表記、日付の位置、参考文献リストの並び順、文中マーカーの形を規定します。最初に一つを選べば、以降のすべての判断（ぶら下げインデント、脚注か括弧書きか、`[1]` か `(Smith, 2024)` か）はそれに従います。
2. **数式は一級コンテンツです** — 本文中のインライン `oMath`、独立ブロックとしての表示用 `oMathPara`、必要に応じて番号付け。
3. **図表は自動採番されます。** `SEQ Figure` / `SEQ Table` フィールドがカウントし、`PAGEREF` が「図2参照」をその実際のページ番号にリンクします。
4. **参考文献はぶら下げインデントを使います**（一行目は左揃え、継続行はインデント）。字下げインデントではありません。左インデント単独でもありません。ぶら下げです。
5. **要旨/キーワード/所属ブロック**は、一枚目の三点セットであり、マーケティング的な意味でのカバーではありません。ブロックスタイルの要旨、字下げインデントなし、装飾なし。
6. **複数段組レイアウト**は IEEE / ACM / Nature など多くのジャーナルに見られます: 単段組の要旨 + 二段組の本文。

### 逆方向のハンドオフ — docx に戻るべきタイミング

会場（venue）や引用スタイルを持たない白書、政策提言書、技術レポート、人事テンプレートなどは **docx v2** に留まってください。**このスキル**を使うのは、文書が以下のうち少なくとも二つを備える場合のみです: 引用スタイルの参考文献、数式、SEQ/PAGEREF 相互参照、複数段組、要旨+キーワードブロック。

## ワークフロー — 5つの動詞

1. **会場（venue）の仕様を読む。** APA 7 / Chicago 17 / IEEE / MLA 9、またはジャーナル固有の仕様。行間、フォント、引用形式、参考文献の並び順 — 以降のすべてはこの一つの決定から従います。
2. **セクションを計画する。** 要旨 → キーワード → 序論 → 方法 → 結果 → 考察 → 結論 → 参考文献。見出しの数を見積もり TOC の要否を判断します（3つ以上の見出しなら TOC を追加、see docx v2 §Table of Contents）。
3. **前もってスタイルを設定する。** Heading1 / Heading2 / Heading3 / Caption / AbstractTitle / Bibliography。コンテンツを追加する前にすべてのスタイルを定義してください（→ see docx v2 §Paragraphs and styles — ここでも同じルール、省略した場合の失敗モードも同じ）。
4. **本文を順番通りに構築する。** 表紙/タイトルブロック → 要旨 → キーワード → TOC（必要なら）→ 読み順の本文セクション → SEQ キャプション付きの図表 → 参考文献 → 脚注は本文追加後に段落パスで最後に加えます。
5. **QA — Delivery Gate。** docx v2 の Gate 1〜3 を継承し、下記の学術専用 Gate 4〜5 を追加します。

## 要件（docx v2 の上に乗る学術の最低ライン）

docx v2 §Requirements for Outputs のすべてが適用されます。それに加え、学術論文は以下の追加ルールを満たさなければなりません:

### タイポグラフィと行間（会場依存）

- **フォント。** 本文は Times New Roman 11-12pt（デフォルト）、または会場指定（IEEE は Times 10pt 2段組、APA は Calibri 11pt を許容）。本文全体で同じフォントを使い、見出しに装飾フォントは使わない。
- **見出し階層。** H1 = 20pt bold、H2 = 14pt bold、H3 = 12pt bold italic、本文 = 11-12pt。（docx v2 と同じ数値ですが、学術論文では Word のデフォルトに頼ることが決してないため改めて記載します。）
- **行間。** APA 7 = 2倍（ダブルスペース）。Chicago / IEEE / ほとんどのジャーナル = 1.5倍。1.15倍を下回らないこと。本文段落と References の両方に設定します。
- **余白。** 特に会場の指定がない限り全辺 1 インチ（1440 twips）（一部ジャーナルは製本用に左 1.25in を要求 — 仕様を確認してください）。

### 要旨、参考文献、キャプションの配置

- **要旨はブロックスタイル。** `firstLineIndent` を使わない。段落間の区切りには `spaceAfter=12pt` を使う。`view issues` が要旨段落に対して「本文段落に字下げインデントがない」と報告した場合、それは誤検知なので無視してよい。
- **参考文献はぶら下げインデントを使う。** 各項目は `indent=720 hangingIndent=720`（左インデント 0.5"、一行目は同量だけ逆方向）を持つ一つの段落。一行目は左揃え、折り返し行は著者名の下でインデントされます。
- **図のキャプションは図の下に置く。** 表のキャプションは表の上に置く。これは、学術分野に不慣れな人が最も間違えやすい単一のルールです — APA、Chicago、IEEE、MLA すべてがこの点で一致しています。
- **引用のラウンドトリップ。** すべての文中引用キーは、参考文献リストの項目に解決できなければなりません。Delivery Gate 4 がこれを検証します。
- **SEQ の存在。** 番号付きの図表を持つ論文はすべて、生きた `SEQ Figure` / `SEQ Table` フィールドを持たなければなりません（文書中に図を挿入するとずれてしまうハードコードされた「Figure 1」というテキストではなく）。Delivery Gate 5 がこれを検証します。

### 表紙/一枚目ブロック

学術論文の表紙は、一般的なビジネス文書の表紙とは異なります。最低限の要素: タイトル（中央揃え、20-22pt bold）、著者、所属、投稿先またはジャーナル名、日付、要旨、キーワード。docx v2 §Visual delivery floor の「60% 埋める」ルールはここでも適用されます — 三行だけの表紙にページの半分が空白なら不合格です。一枚目のレシピについては下記の §要旨 / キーワード / 所属ブロック を参照してください。

### 節番号の慣習（スタイル依存 — 機械的に適用しないこと）

学術的な節番号は、リスト番号付けで計算するものではなく、**見出しテキストの一部**です。`officecli` の `numId`/`listStyle` の仕組みは Heading1 の再利用をまたぐと壊れやすいため、プレフィックスは手書きしてください。ただし、そのプレフィックスの形はスタイルによって異なります — 4つのスタイルすべてに同じ形式を使わないでください:

| スタイル | H1 形式 | H2 形式 | 例 |
|---|---|---|---|
| **APA 7** | **番号なし・中央揃え・太字** | 番号なし・左揃え・太字 | `Introduction` / `Methods`（中央揃え） |
| **Chicago** | `"N. Title"` 左揃え | `"N.M Title"` | `1. Introduction`、`2.1 Policy Formation` |
| **IEEE** | `"N. TITLE"` 全て大文字 + ローマ数字 | `A. Subtitle` タイトルケース | `I. INTRODUCTION`、`II. RELATED WORK`、`A. Datasets` |
| **MLA 9** | 番号なし・左揃え・太字 | 同上 | `Literature Review`（プレフィックスなし） |

APA 7 の L1 見出しは **中央揃え、太字、番号なし**、L2 は左揃え太字、L3 は左揃え太字斜体、L4/L5 はランイン。APA の見出しに `1. / 2.` のようなプレフィックスを付けないでください — それは Chicago/IEEE の慣習です。IEEE はローマ数字付きの全て大文字（`I. INTRODUCTION`）を求めます。各節の中では `A./B./C.` のサブ見出し（タイトルケース）を使います。アラビア数字で採番された本文節は Chicago スタイルのみです。

**四スタイル共通の例外**: References / Bibliography / Works Cited / Acknowledgments はスタイルに関わらず番号なし — `N.` プレフィックスは省略してください。

## クイックスタート — 最小構成の APA 論文

```bash
FILE="paper.docx"
officecli create "$FILE"
officecli open "$FILE"
officecli set "$FILE" / --prop defaultFont="Times New Roman"
officecli add "$FILE" /body --type paragraph --prop text="Remote Work and Team Cohesion" --prop align=center --prop size=20pt --prop bold=true --prop spaceAfter=24pt
officecli add "$FILE" /body --type paragraph --prop text="Alice Chen" --prop align=center --prop size=12pt
officecli add "$FILE" /body --type paragraph --prop text="Department of Psychology, Stanford University" --prop align=center --prop size=11pt --prop spaceAfter=24pt
officecli add "$FILE" /body --type paragraph --prop text="Abstract" --prop align=center --prop size=14pt --prop bold=true --prop spaceBefore=12pt --prop spaceAfter=6pt
officecli add "$FILE" /body --type paragraph --prop text="This study examines remote-work adoption on team cohesion across 18 months..." --prop size=12pt --prop lineSpacing=2x --prop spaceAfter=12pt
officecli add "$FILE" /body --type paragraph --prop text="Keywords: remote work, team cohesion, psychological safety" --prop italic=true --prop size=11pt --prop spaceAfter=18pt
officecli add "$FILE" /body --type paragraph --prop text="1. Introduction" --prop style=Heading1 --prop size=20pt --prop bold=true --prop spaceBefore=18pt --prop spaceAfter=12pt
officecli add "$FILE" /body --type paragraph --prop text="Remote-work research (Smith, 2024) has expanded since 2020..." --prop size=12pt --prop lineSpacing=2x --prop firstLineIndent=720
officecli add "$FILE" /body --type paragraph --prop text="References" --prop style=Heading1 --prop size=20pt --prop bold=true --prop spaceBefore=18pt --prop spaceAfter=12pt
officecli add "$FILE" /body --type paragraph --prop text="Smith, J. (2024). Remote work and cohesion. Journal of Applied Psychology, 109(3), 412-430." --prop size=12pt --prop lineSpacing=2x --prop indent=720 --prop hangingIndent=720
officecli add "$FILE" / --type footer --prop type=default --prop align=center --prop size=10pt --prop field=page
officecli close "$FILE"
officecli validate "$FILE"
```

十行の骨格です。実際の論文は、本文段落を増やし、参考文献項目を増やし（それぞれ同じ `indent=720 hangingIndent=720` のペアを付けて）、キャプション付きの図表を加え、Heading1 が3つ以上あれば TOC を加えることで育っていきます。クイックスタートは検証をクリアします。以降の節ではそれぞれの側面を詳しく説明します。

## 引用スタイルのレシピ

主流の4つのファミリーです。プロジェクト開始時に一つを選べば、以降のすべての判断はそれに従います。**スタイル別の決定表:**

| スタイル | 文中引用の形 | 参考文献リストの並び順 | 本文行間 | 脚注は使う？ |
|---|---|---|---|---|
| APA 7 | `(Smith, 2024)` または `Smith (2024)` | 著者のアルファベット順 | 2倍（ダブル） | まれ（内容注のみ） |
| Chicago 17（Notes-Bib） | 上付きの脚注番号 | 著者のアルファベット順 | 1.5倍〜2倍 | **主軸**（完全な引用情報を脚注に記載） |
| IEEE | `[1]`、`[2]`、...、`[N]` | 初出順 | 1.15倍〜1.5倍、2段組 | まれ |
| MLA 9 | `(Smith 412)` ページ番号 | 著者のアルファベット順、「Works Cited」 | 2倍 | まれ |

四スタイル共通のデフォルト: 参考文献段落は `indent=720 hangingIndent=720`（ぶら下げインデント 0.5"）を使用する。Heading1 が3つ以上あればライブ TOC を追加する（→ see docx v2 §Table of Contents）。`updateFields=true` を設定し、Word 互換のフィールドエンジンが更新するまで TOC のページ番号は未計算として報告する。

### APA 7（社会科学 — 心理学、教育学、経営学）

- 文中: `(Author, Year)` または、地の文なら `Author (Year)`。直接引用にはページ番号が必須: `(Smith, 2024, p. 15)`。著者3名以上: 初出後は `(Smith et al., 2024)`。
- 参考文献リストの並び順: **第一著者の姓のアルファベット順**。タイトルの表記: 論文タイトルはセンテンスケース、ジャーナル名はタイトルケース（斜体）。
- 参考文献の形: `Author, A. A., & Co-Author, B. B. (Year). Title of article. Journal Name, Volume(Issue), pages.` DOI は URL より優先し、`doi:` プレフィックスではなく https URL として記載する。
- すべてダブルスペース（`lineSpacing=2x`）とし、要旨と参考文献も含む。本文の一行目インデントは 0.5"（`firstLineIndent=720`）。

```bash
# 括弧書き引用を含む本文段落
officecli add "$FILE" /body --type paragraph --prop text="Remote work adoption accelerated during the pandemic (Kramer & Kramer, 2020)." --prop size=12pt --prop lineSpacing=2x --prop firstLineIndent=720
# ぶら下げインデント付きの参考文献項目
officecli add "$FILE" /body --type paragraph --prop text="Kramer, A., & Kramer, K. Z. (2020). The potential impact of the Covid-19 pandemic on occupational status. Journal of Vocational Behavior, 119, 103442." --prop size=12pt --prop lineSpacing=2x --prop indent=720 --prop hangingIndent=720
# 参考文献段落に付与する DOI ハイパーリンク
officecli add "$FILE" "/body/p[last()]" --type hyperlink --prop url="https://doi.org/10.1016/j.jvb.2020.103442" --prop text="https://doi.org/10.1016/j.jvb.2020.103442"
```

QA: `officecli query "$FILE" 'paragraph[hangingIndent]'` はすべての参考文献項目を返す。字下げインデントのままの参考文献項目がぶら下げインデントの代わりに残っているものがゼロであること。

### Chicago 17 — Notes-Bibliography（人文科学 — 歴史学、哲学、宗教学）

- 文中: 上付きの脚注番号。完全な引用情報を最初の脚注に記載する（`Timothy Brook, The Troubled Empire (Cambridge, MA: Harvard UP, 2010), 142.`）。それ以降は **短縮形**（`Brook, Troubled Empire, 150.`）。
- **再引用のルール（Chicago 17, op. cit. は非推奨）:**
  - 同じ出典・同じページを**直前に連続して**引用する場合 → `Ibid.`
  - 同じ出典の**直前連続・異なるページ** → `Ibid., 22.`
  - 連続していない再引用 → **短縮形**（`Brook, Troubled Empire, 150.`）を使い、`op. cit.` は使わない。Chicago 17 では `op. cit.` を廃止しており、直前の反復を除いて常に短縮形を使う。
- 巻末に参考文献リスト、**第一著者の姓のアルファベット順**（「Brook, Timothy.」）、ぶら下げインデント。脚注本文は閲覧者側の脚注デフォルト（通常 10pt）でレンダリングされ、参考文献項目は 12pt。（`footnote` 要素は `text` のみを公開しており、脚注ごとにサイズを設定することはできません — レンダラーのデフォルトを信頼してください。）
- 一次資料の多い論文では典型的に、単一の `Bibliography` Heading1 の下に `Primary Sources` と `Secondary Sources` の二つの Heading2 に分割する。書名は脚注・参考文献の両方で斜体。
- Chicago には理系分野で使われる Author-Date バリアントもある — 会場が Chicago Author-Date を指定している場合は、APA のレシピに従い句読点だけを変更する（著者と年の間にカンマを入れない: `(Smith 2024)`）。

```bash
# 脚注のアンカーとなる本文段落、続けて脚注そのもの
officecli add "$FILE" /body --type paragraph --prop text="The Ming dynasty's 海禁 policy shaped coastal trade for two centuries." --prop size=12pt --prop lineSpacing=1.5x --prop firstLineIndent=720
officecli add "$FILE" "/body/p[last()]" --type footnote --prop text="Timothy Brook, The Troubled Empire: China in the Yuan and Ming Dynasties (Cambridge, MA: Harvard University Press, 2010), 142."
# 次の脚注 — 短縮形
officecli add "$FILE" "/body/p[last()]" --type footnote --prop text="Brook, Troubled Empire, 150."
# 参考文献セクションの分割 — 一次資料を先に
officecli add "$FILE" /body --type paragraph --prop text="Bibliography" --prop style=Heading1 --prop size=20pt --prop bold=true --prop spaceBefore=18pt
officecli add "$FILE" /body --type paragraph --prop text="Primary Sources" --prop style=Heading2 --prop size=14pt --prop bold=true --prop spaceBefore=12pt
officecli add "$FILE" /body --type paragraph --prop text="Ming Shilu 明實錄. Taipei: Academia Sinica, 1966." --prop size=12pt --prop indent=720 --prop hangingIndent=720
officecli add "$FILE" /body --type paragraph --prop text="Secondary Sources" --prop style=Heading2 --prop size=14pt --prop bold=true --prop spaceBefore=12pt
officecli add "$FILE" /body --type paragraph --prop text="Brook, Timothy. The Troubled Empire: China in the Yuan and Ming Dynasties. Cambridge, MA: Harvard University Press, 2010." --prop size=12pt --prop indent=720 --prop hangingIndent=720
```

QA: `officecli query "$FILE" 'footnote'` の件数 ≥ 本文段落内の引用件数。

### IEEE（工学 — トランザクション誌、学会予稿集）

- 文中: `[1]`、`[2]`。アルファベット順ではなく**初出順**で採番する。同じ出典の再引用には同じ番号を再利用する。ページ参照は `[1, p. 15]`、範囲は `[1]-[3]`。
- 参考文献項目は角括弧の番号で始まる: `[1] A. Smith and B. Jones, "Title," IEEE Trans. X, vol. 5, no. 3, pp. 1-10, 2024, doi: ...`。著者はイニシャル先頭、ジャーナル名は IEEE のリストに従って省略形にする（`IEEE Trans. Neural Netw.` であり、フルネームではない）。
- 本文は**二段組**（下記 §複数段組 を参照）。要旨は折り返し前の単段組、10pt、行間 1.15倍、通常 200〜250 語。
- 本文段落の一行目インデントは 0.2"（`firstLineIndent=288` twips ≈ 14pt）。2段組の幅が狭いため APA の 0.5" より小さい。
- **節見出し: ローマ数字付きの全て大文字** — `I. INTRODUCTION`、`II. RELATED WORK`、`III. METHOD`。サブセクションはタイトルケースで `A. Datasets`、`B. Baselines`。IEEE で `1. Introduction`（アラビア数字）を使わないこと — それは Chicago スタイルです。
- **表はローマ数字で採番**: `Table I`、`Table II`、`Table III`。図はアラビア数字のまま（`Fig. 1`、`Fig. 2`）。`recalcFields=seq` は両方にアラビア数字のキャッシュ値を書き込みます — IEEE のローマ数字表については、recalc 後にキャッシュされた `<w:t>` を手動でローマ数字にパッチするか、アラビア数字のまま受け入れてカバーレターに注記してください。

```bash
# 参考文献1を引用する本文
officecli add "$FILE" /body --type paragraph --prop text="Attention-based anomaly detection has been applied to industrial sensor data [1], [2]." --prop size=10pt --prop lineSpacing=1.15x
# 参考文献リストの項目 — 本文中の番号
officecli add "$FILE" /body --type paragraph --prop text="[1] A. Smith and B. Jones, \"Attention for anomaly detection,\" IEEE Trans. Neural Netw., vol. 35, no. 2, pp. 412-430, 2024." --prop size=10pt --prop indent=720 --prop hangingIndent=720
officecli add "$FILE" /body --type paragraph --prop text="[2] C. Lee, \"Time-series anomaly survey,\" in Proc. ICML, 2023, pp. 1200-1215." --prop size=10pt --prop indent=720 --prop hangingIndent=720
```

QA: 本文中で最も大きい `[N]` は参考文献項目数と一致しなければならない。Grep: `officecli view "$FILE" text | grep -oE '\[[0-9]+\]' | sort -u | tail -5`。

### MLA 9（文学、言語、文化研究）

APA との違い: 文中引用は `(Author Page)` で**カンマなし**（例: `(Smith 412)`）。直接引用には常にページ番号を付ける。参考文献セクションのタイトルは **Works Cited**（References / Bibliography ではない）。項目は姓のアルファベット順、ぶら下げインデント、行間2倍、ピリオドで区切られた九つの「コア要素」: `Author. Title. Container, Other Contributors, Version, Number, Publisher, Date, Location.` — 該当しない要素は省略する。書名は斜体、記事タイトルは引用符。それ以外は APA の段落設定と同一。

## 数式（OMML — インライン vs 表示）

`--type equation` は LaTeX 風の数式を OMML にパースします。`--prop mode=` で選択する二つのモード:

| モード | XML | 見た目 | 用途 |
|---|---|---|---|
| `display`（デフォルト） | `/body` 上の `<m:oMathPara>` | 独立した中央揃えのブロック | 番号付き数式、定理の記述 |
| `inline` | 段落内のランに付加される `<m:oMath>` | 本文と一体化 | 地の文中の `if $x > 0$` のようなスタイル |

```bash
# 表示用数式（独立した段落、中央揃え）— わかりやすさのため明示的に mode=display を指定
officecli add "$FILE" /body --type equation --prop mode=display --prop formula="x^2 + y^2 = z^2"
# ギリシャ文字・下付き・積分を含む表示用数式 — 下記でレンダリングを確認
officecli add "$FILE" /body --type equation --prop mode=display --prop formula="\\lambda_1 + \\alpha"
officecli add "$FILE" /body --type equation --prop mode=display --prop formula="\\frac{1}{2\\pi} \\int_0^{\\infty} e^{-x^2} dx"
# 地の文中のインライン数式 — 本文段落に x_{t+1}, \lambda などの変数が現れる場合は必須:
officecli add "$FILE" /body --type paragraph --prop text="Given the weight " --prop size=11pt
officecli add "$FILE" "/body/p[last()]" --type equation --prop mode=inline --prop formula="W_t"
officecli add "$FILE" "/body/p[last()]" --type run --prop text=" we define the loss..."
```

**数式が（プレーンテキストの LaTeX トークンではなく）OMML 数式としてレンダリングされることを確認してください。** `close` の後、以下を実行します:
```bash
officecli view "$FILE" text | head -20       # λ₁ + α、∫₀∞、x² は Unicode の数式として表示されなければならない（レンダリング確認済み）
officecli raw "$FILE" /document | grep -c '<m:oMathPara'   # 表示用数式1つにつき ≥ 1
```
本文の地の文に生の `lambda_1`、`x_{t+1}`、`\alpha` のようなプレーンテキストトークンが含まれている場合（つまり `--type equation --prop mode=inline` で包む代わりに `paragraph --prop text=` にそのまま入力してしまった場合）、下流のビューアーはそれをリテラルな ASCII としてレンダリングします。**ルール: 地の文中の数学変数・ギリシャ文字・下付き文字はすべて `--type equation mode=inline` を経由させ、`paragraph --prop text=` で書かないこと。**

**LaTeX サブセットの落とし穴**（絶対厳守）:

1. `\left(...\right)` / `\left[...\right]` の**内側**に上付き/下付きがある場合 → パースエラー（`Error: cast object … Subscript`）。**外側**のスクリプト（`\left(x+y\right)^2`）は問題ない。プレーンな `(`、`)`、`[`、`]` は常に動作し、OMML が表示モードで自動的にサイズ調整する。
2. `/body/oMathPara[N]` に対する `move` は表示用数式の位置を並べ替える（囲んでいる段落を移動させる）。`--before <path>` は空の段落を残してしまうことがあるため、きれいに並べ替えるには `--index` か `--after` を使う方がよい。

**数式の番号付け** — ネイティブの `\eqno` はありません。ジャーナル標準のレイアウトは**一行**です: 数式を中央に、番号を列端に右揃えで配置（`Y = A(X) ⊗ X      (1)`）。これは段落の**タブストップ**二つ — 列の中間点にある `center` タブと列の右端にある `right` タブ — で構築し、一つの段落に `[tab] equation(inline) [tab] (1)` を並べます。中央揃えの表示用数式の後ろに別行の右揃え `(1)` を続ける方法は使わないでください — それでは番号が別の行に分かれてしまい、エージェントが期待通りの見た目を再現できない最も多い原因です。

```bash
# タブ位置は列幅（twips）に依存する。デフォルトの空白文書 = A4、余白 3.18cm
#   → 本文幅 = 8300 twips。
#   単段組: center タブ = 4150、right タブ = 8300
#   二段組（デフォルトの溝幅 720 twips）: 列幅 = (8300-720)/2 = 3790 → center 1895、right 3790
# （他のページサイズ/余白の場合: 本文幅 = pageWidth - marginLeft - marginRight で再計算すること。）
officecli add "$FILE" /body --type paragraph                                   # 数式段落（仮に p[N] に着地するとする）
officecli add "$FILE" "/body/p[N]" --type tab --prop pos=4150 --prop val=center
officecli add "$FILE" "/body/p[N]" --type tab --prop pos=8300 --prop val=right
officecli add "$FILE" "/body/p[N]" --type run --prop text=$'\t'                 # タブ → 中央にジャンプ
officecli add "$FILE" "/body/p[N]" --type equation --prop mode=inline --prop formula='Y = A(X) \otimes X'
officecli add "$FILE" "/body/p[N]" --type run --prop text=$'\t(1)'             # タブ → 右端にジャンプし、その後に番号
```

二段組の節では、二段組用のタブ位置（1895 / 3790）を使うだけでよく、同じ段落が列内で数式を中央揃えにし、`(1)` を列の右端に配置します。スキーマ: `officecli help docx tab`。

**`--type equation` を表セル `tc[N]` のパスに直接置かないこと** — ガードエラー（`table cells only accept paragraphs, tables, or SDTs`）で拒否され、不正な XML は書き込まれません。セル内に数式が必要な場合は `tc[N]/p[1]` を `mode=inline` で指定してください。

完全な数式スキーマ: `officecli help docx equation`。

## 図表と相互参照（SEQ + PAGEREF）

いずれも**ネイティブの fieldType**（`officecli help docx field` で列挙を確認）である二つのプリミティブ: 自動採番されるキャプションカウンターの `seq`、「7ページの図2参照」という背参照のための `pageref`。ネイティブフィールドは未評価のキャッシュ結果とともに挿入され、`set "$FILE" / --prop recalcFields=seq` を一度実行するだけで（下記 §SEQ 採番 参照）文書順に実際の番号が埋まります — フィールドごとの個別パッチは不要です。

### SEQ 自動採番 — 図と表

SEQ フィールドは名前（`identifier`）を持つカウンターです。すべての `SEQ Figure` は **recalc** のたびに Figure カウンターをインクリメントし、すべての `SEQ Table` は Table カウンターをインクリメントします。

**SEQ のキャッシュ値 — すべてのキャプションを追加し終えたら一つのコマンドで。** 追加されたばかりの SEQ フィールドには評価済みのキャッシュ番号がなく、`view text` は recalc するまで `#OCLI_NOTEVAL!{SEQ Figure}` というセンチネル（および `Format["evaluated"]=false`）を表示します。**フィールドを一つずつ手でパッチしないこと。** すべての図表キャプションを配置し終えたら、一度だけ実行します:
```bash
officecli set "$FILE" / --prop recalcFields=seq
```
これは本文の文書順に `SEQ Figure` / `SEQ Table` フィールドをカウントし、実際のキャッシュ値（`Figure 1 / Figure 2 / Figure 3`）を書き込んで `evaluated` を true に反転させます。文書の途中にキャプションを挿入した後は、番号の同期を保つために再実行してください。（見出し相対の `\s` や、ヘッダー/フッター内の SEQ は Word 自身の再計算に委ねられます — see `help docx document`。）確認: `officecli view "$FILE" text` に明確に昇順の異なる番号が表示されること。

```bash
# 画像の下にキャプションを配置する図。キャプション = "Figure <seq>: title" + 相互参照用の任意のブックマーク。
officecli add "$FILE" /body --type picture --prop src=arch.png --prop width=5in
officecli set "$FILE" "/body/p[last()]/r[last()]" --prop alt="Model architecture: attention over time-series sensors"
# キャプション段落（学術的慣習に従い図の下に配置）
officecli add "$FILE" /body --type paragraph --prop text="Figure " --prop style=Caption --prop size=10pt --prop italic=true --prop align=center
officecli add "$FILE" "/body/p[last()]" --type field --prop fieldType=seq --prop identifier=Figure
officecli add "$FILE" "/body/p[last()]" --type run --prop text=": Attention-based anomaly detection model."
# 他の段落が PAGEREF で参照できるようキャプションにブックマークを付与
officecli add "$FILE" /body --type bookmark --prop name=fig_arch
# （番号は、すべてのキャプションが揃った後に一度の `set / --prop recalcFields=seq` で埋められる。）
```

### PAGEREF — ブックマークによる相互参照

```bash
# 相互参照段落: 「図1参照（Xページ）」
officecli add "$FILE" /body --type paragraph --prop text="As shown in Figure 1 (see page " --prop size=11pt --prop lineSpacing=1.5x
officecli add "$FILE" "/body/p[last()]" --type field --prop fieldType=pageref --prop name=fig_arch
officecli add "$FILE" "/body/p[last()]" --type run --prop text=")."
```

### 表 — キャプションは上に

```bash
# キャプションを先に（表の上に）、その後に表
officecli add "$FILE" /body --type paragraph --prop text="Table " --prop style=Caption --prop size=10pt --prop italic=true --prop spaceAfter=6pt
officecli add "$FILE" "/body/p[last()]" --type field --prop fieldType=seq --prop identifier=Table
officecli add "$FILE" "/body/p[last()]" --type run --prop text=": Participant demographics (N=47)."
officecli add "$FILE" /body --type table --prop rows=5 --prop cols=4 --prop width=100%
# ... docx v2 §Tables に従いヘッダーと行を埋める
```

### SEQ + PAGEREF フィールドが着地したことの確認

```bash
# 本文の document パートに SEQ Figure または SEQ Table が少なくとも1つあること
officecli raw "$FILE" /document | grep -c 'w:instrText[^>]*>[^<]*SEQ'   # ≥ 1 を期待
officecli raw "$FILE" /document | grep -c 'w:instrText[^>]*>[^<]*PAGEREF' # 相互参照がなければ 0 でよい
```

ライブフィールドは、人間が Word で F9 を押すまで古びて見える**キャッシュ値**を持ちます。「Figure 1」は recalc 直後に `1`、`2`、... として表示されることを期待してください。recalc 前は、ビューアーによっては `0` や空白が表示されることがあります。フィールドの存在は、目に見える数字ではなく `fldChar` の存在で判断してください（→ see docx v2 §Field / cached-value spot-check）。

## 脚注 vs 文末脚注

**脚注（Footnote）** — アンカーとなる段落が存在するページの下部に配置されます。Chicago Notes-Bib の出典引用や、どのスタイルにおける内容補足にも使われます。

**文末脚注（Endnote）** — 文書の末尾（または参考文献の前）に配置されます。一部の会場で脚注の代わりに、あるいはページを煩雑にする長い文脈注に使われます。

```bash
# 段落 N にアンカーされた脚注
officecli add "$FILE" "/body/p[3]" --type footnote --prop text="Smith et al. reported similar findings in their 2023 review."
# 文末脚注 — 段落 N にアンカー（脚注と同様）。/endnote[@endnoteId=N] に着地する
officecli add "$FILE" "/body/p[3]" --type endnote --prop text="Extended derivation of equation (4) is available at the project repository."
```

どちらも `view annotated` の出力では空文字列のランとして表示されます（`r[N] ""`）— そのランは可視テキストではなく `<w:footnoteReference>` という XML 要素を持ちます。挿入の確認は `officecli query "$FILE" 'footnote'` または `officecli get "$FILE" "/footnotes/footnote[N]"` で行ってください。脚注は段落インデックスをずらしません — 本文の内容が揃った後、どの順序で追加しても構いません。完全なスキーマ: `officecli help docx footnote` / `officecli help docx endnote`。

## 参考文献セクション

すべての学術論文は参考文献リストで終わります。セクション名はスタイルによって異なります（APA / IEEE / Chicago Author-Date は **References**、Chicago Notes-Bib は **Bibliography**、MLA は **Works Cited**）。各項目は**ぶら下げインデント**を持つ独立した段落です。

```bash
# セクション見出し — 本文の Heading1 と同じ（慣習により本文の番号付けから除外される）
officecli add "$FILE" /body --type paragraph --prop text="References" --prop style=Heading1 --prop size=20pt --prop bold=true --prop spaceBefore=18pt --prop spaceAfter=12pt
# 各項目: ぶら下げインデント 720 twips（0.5"）。indent=720 を対にする（一行目は左揃え、折り返しはインデント）
officecli add "$FILE" /body --type paragraph --prop text="Smith, J. (2024). Remote work and cohesion. Journal of Applied Psychology, 109(3), 412-430." --prop size=12pt --prop lineSpacing=2x --prop indent=720 --prop hangingIndent=720
# 項目段落に付与する DOI ハイパーリンク（独立したランとして）
officecli add "$FILE" "/body/p[last()]" --type hyperlink --prop url="https://doi.org/10.1037/apl0001123" --prop text="https://doi.org/10.1037/apl0001123"
```

検証済み: `--prop indent=720 --prop hangingIndent=720` は `officecli help docx paragraph` に基づく正規のぶら下げインデントのペアです。旧来の `ind.firstLine=-720`（負の一行目インデント）形式は正規ではなく、出力時にスキーマ検証に失敗します — → see docx v2 §Schema-invalid-on-emit。

**ラウンドトリップ QA。** 文中引用マーカー（APA `(Author, Year)`、IEEE `[N]`、MLA `(Author N)`）と参考文献項目の件数を突き合わせてください。下記 Delivery Gate 4 を参照。引用されたすべてのキーは解決できなければならず、リストに載っているすべての項目は少なくとも一度は引用されているべきです。

## 複数段組（IEEE ジャーナルの二段組レシピ）

IEEE や多くの工学・物理学系ジャーナルは、本文テキストを二段組でレンダリングし、その上に単段組の要旨を配置します。仕組み: `type=continuous` と `columns=2` を持つセクション区切り、そして文書末尾でもう一つのセクション区切りを入れて単段組に**戻す**。

**この復元ステップは省略できません。** これを行わないと、参考文献を含む文書の残り全体が二段組でレンダリングされます。これが複数段組で最もよくある失敗です。

```bash
FILE="ieee.docx"
officecli create "$FILE"
officecli open "$FILE"

# 1. タイトル、著者、所属 — 単段組（デフォルトの最初のセクション）
officecli add "$FILE" /body --type paragraph --prop text="Attention-Based Anomaly Detection for Industrial Time Series" --prop align=center --prop size=18pt --prop bold=true --prop spaceAfter=12pt
officecli add "$FILE" /body --type paragraph --prop text="Alice Chen, Bob Martinez" --prop align=center --prop size=11pt
officecli add "$FILE" /body --type paragraph --prop text="Department of CS, Stanford University" --prop align=center --prop size=10pt --prop spaceAfter=18pt

# 2. 要旨 — まだ単段組、ブロックスタイル
officecli add "$FILE" /body --type paragraph --prop text="Abstract" --prop align=center --prop size=12pt --prop bold=true --prop spaceAfter=6pt
officecli add "$FILE" /body --type paragraph --prop text="We present an attention-based model for detecting anomalies in industrial sensor time series..." --prop size=10pt --prop lineSpacing=1.15x --prop spaceAfter=12pt

# 3. セクション区切り + ここから二段組
#    `/section[last()]` は最終セクションに解決される（p[last()] と同様）。明示的な /section[N] も動作する。
officecli add "$FILE" /body --type section --prop type=continuous
SECTION_COUNT=$(officecli query "$FILE" section --json | jq '.data.results | length')
# add の後、SECTION_COUNT は 2 になるはず — [1] が区切り前、[2] が区切り後（二段組の本文エリア）。
officecli set "$FILE" "/section[2]" --prop columns=2 --prop columnSpace=1cm

# 4. 本文 — IEEE はローマ数字 + 全て大文字の節タイトルを求める（P1.2）。
officecli add "$FILE" /body --type paragraph --prop text="I. INTRODUCTION" --prop style=Heading1 --prop size=10pt --prop bold=true
officecli add "$FILE" /body --type paragraph --prop text="Industrial anomaly detection has been studied since [1]..." --prop size=10pt --prop lineSpacing=1.15x --prop firstLineIndent=360

# 5. 二段組本文の末尾で、さらにセクション区切りを入れて参考文献/付録用に単段組へ戻す
# （参考文献も二段組にしたい場合はステップ5をスキップ — ただしほとんどの IEEE 論文は参考文献も二段組。）
# officecli add "$FILE" /body --type section --prop type=continuous
# 最終セクションには /section[last()] を使うか、明示的な /section[N] のため再カウントする:
# officecli set "$FILE" "/section[3]" --prop columns=1

# 6. フッター、close、検証
officecli add "$FILE" / --type footer --prop type=default --prop align=center --prop size=9pt --prop field=page
officecli close "$FILE"
officecli validate "$FILE"
```

**目視での検証。** `officecli view "$FILE" html` を実行し、返された HTML を Read でレンダリング結果を確認してください。要旨は全幅でレンダリングされ、序論以降は二段組でレンダリングされなければなりません。要旨が二つの狭い列に折り返されている場合は、最初のセクション区切りが要旨より前に着地しています — 移動してください。

**セクションのインデックス管理。** `add /body --type section` は一つの空段落を `/body` に挿入します（セクション区切りのマーカー）。以降のすべての `p[N]` インデックスはセクション区切り一つにつき +1 ずつシフトします。セクション区切りは事前に計画し、区切りを追加した後は `officecli get "$FILE" /body --depth 1` で再インデックスしてから続けてください。

完全なセクションスキーマ（`columns`、`columnSpace`、`orientation`、`pageNumFmt`、`titlePage`、`lineNumbers`）: `officecli help docx section`。

## 要旨 / キーワード / 所属ブロック

一枚目のメタデータの積み重ね: タイトル（中央揃え 20-22pt bold）→ 著者（中央揃え 12pt、複数所属の場合は上付きの `^1 ^2`）→ 所属（中央揃え 11pt、上付き文字に紐づく）→ 投稿先/日付 → **Abstract** 見出し（14pt bold）→ 要旨本文（ブロックスタイル、**`firstLineIndent` なし**、150〜300語）→ キーワード行（斜体 11pt）。docx v2 と同じ「表紙は60%以上埋める」ルールが適用されます。

```bash
# 上付きの所属マーカー（複数機関にまたがる論文）
officecli add "$FILE" /body --type paragraph --prop text="Alice Chen" --prop align=center --prop size=12pt
officecli add "$FILE" "/body/p[last()]" --type run --prop text="1" --prop superscript=true
officecli add "$FILE" "/body/p[last()]" --type run --prop text=", Bob Martinez"
officecli add "$FILE" "/body/p[last()]" --type run --prop text="2" --prop superscript=true
# ランニングヘッダー（表紙ではスキップ — type=first の空ヘッダーで実現。see docx v2 §headers）
officecli add "$FILE" / --type header --prop type=default --prop align=right --prop size=9pt --prop text="Short Running Title"
```

**Nature 系の二段組要旨**は稀ですが、必要な場合は Abstract 見出しの**前**に `section type=continuous columns=2` を開いてください。要旨が短い（100語未満）と列がガタガタになります。**奇数/偶数ページで異なるヘッダー**は高レベル API で公開されています: `officecli set "$FILE" /settings --prop evenAndOddHeaders=true` の後、`--type header --prop type=even` で偶数ページ用ヘッダーを追加します。完全なヘッダースキーマ: `officecli help docx header`。

## QA — Delivery Gate（実行可能）

**問題があると想定してください。あなたの仕事はそれを見つけることです。** 最初のレンダリングが正しいことはほとんどありません。完了と宣言する前に、このブロックを実行してください。

### Gate 1〜3 — docx v2 から継承

→ see docx v2 §Delivery Gate。スキーマ検証、トークン漏れの grep、ライブ PAGE フィールドの構造。まず docx v2 の Gate ブロックをそのままコピーして貼り付けてください。すべてのチェックが成功メッセージを表示しなければなりません。

### Gate 4 — 引用のラウンドトリップ

すべての文中引用キーは参考文献項目に解決できなければなりません。件数の不一致 = REJECT。

```bash
# IEEE の例（角括弧の数字）。APA (Author, Year) や MLA (Author Page) では正規表現を調整すること。
CITATIONS=$(officecli view "$FILE" text | grep -oE '\[[0-9]+\]' | sort -u | wc -l)
ENTRIES=$(officecli query "$FILE" 'paragraph[hangingIndent]' --json | jq '.data.results | length')
echo "In-text citation markers: $CITATIONS | Bibliography entries: $ENTRIES"
# 引用数が参考文献数を上回る場合（引用のみで参考文献がない）は REJECT。参考文献数 > 引用数 は一部の会場では許容される。
[ "$CITATIONS" -le "$ENTRIES" ] && echo "Gate 4 OK" || { echo "REJECT Gate 4: $CITATIONS in-text markers but only $ENTRIES bibliography entries"; exit 1; }
```

### Gate 5a — SEQ の存在 + キャッシュ番号が重複していないこと

図または表に番号が振られている論文の場合、本文はライブな `SEQ` フィールドを持ち、かつそのキャッシュ値は重複のない昇順の番号を示さなければなりません（さもないと、`view text` やキャッシュフィールドを再計算しない下流のビューアーはすべてに「Figure 1」を表示してしまいます）。

```bash
# query 経由で SEQ フィールドをカウントする（raw grep は一行の XML 上で複数マッチが潰れて過小カウントになる）。
SEQ_COUNT=$(officecli query "$FILE" 'field[fieldType=seq]' --json | jq '.data.results | length')
VISIBLE_FIG=$(officecli view "$FILE" text | grep -cE '(Figure|Table) [0-9]+')
if [ "$VISIBLE_FIG" -gt 0 ] && [ "$SEQ_COUNT" -eq 0 ]; then
  echo "REJECT Gate 5a: $VISIBLE_FIG visible Figure/Table labels but 0 SEQ fields."
  exit 1
fi
# キャッシュ値は重複してはならない。すべてのキャプションが揃った後に一度だけ `set / --prop recalcFields=seq` を実行すること。
# recalc 前はフィールドが #OCLI_NOTEVAL! センチネルをレンダリングし、recalc 後は Figure 1 / Figure 2 / Figure 3 になる:
DISTINCT=$(officecli view "$FILE" text | grep -oE '(Figure|Table) [0-9]+' | sort -u | wc -l)
[ "$SEQ_COUNT" -le "$DISTINCT" ] && echo "Gate 5a OK (SEQ=$SEQ_COUNT, distinct=$DISTINCT)" || { echo "REJECT Gate 5a: $SEQ_COUNT SEQ fields but only $DISTINCT distinct rendered labels — run 'set \"$FILE\" / --prop recalcFields=seq'"; exit 1; }
```

### Gate 5b — HTML プレビューによる目視監査（必須、任意ではない）

Gate 1〜5a はスキーマ、トークン漏れ、ライブフィールドの存在、引用数を捕捉します。**しかし物理的な組み立て上の欠陥は捕捉しません** — ページ順の入れ替わり、文書中盤での要旨の重複、SEQ フィールドが存在するのに3つの図がすべて「Fig. 1」とラベル付けされている、数式の変数が数式ではなくプレーンテキストの LaTeX（`lambda_1`、`x_{t+1}`）としてレンダリングされている、など。省略しないでください — Gate 1〜5a の合格は目視 OK を意味しません。

`officecli view "$FILE" html` を実行し、返された HTML パスを Read してください。論文のすべてのページについて、以下に答えてください:

> (a) ページは論理的な学術的順序になっているか？（タイトル → 要旨 → キーワード → 序論 → 本文 → 参考文献 — 前方へのジャンプや後方への漏れがないこと。）
> (b) 要旨は文書中に一度だけ現れるか、文書中盤で重複していないか？
> (c) Figure N / Table N のラベルは重複なく昇順になっているか？（Fig. 1, Fig. 2, Fig. 3 — すべて「Fig. 1」になっていないこと。表も同様。）
> (d) 数式は数式としてレンダリングされているか？（斜体化された変数、λ / α のようなギリシャ文字、適切な積分/分数 — プレーンテキストの `lambda_1`、`x_{t+1}`、`\int` ではないこと。）
> (e) IEEE 論文の場合: 節タイトルはローマ数字付きの全て大文字（`I. INTRODUCTION`）になっているか？ 表はローマ数字（`Table I`、`Table II`）になっているか？
> (f) APA 論文の場合: レベル1見出しは中央揃え太字で番号なし（`1. Introduction` になっていない）か？
> (g) 文中の「see Fig. N」/「see Table N」はすべて、実際にその番号を持つ図/表に解決されるか？
> (h) H1 / H2 / H3 を通じて見出し階層が視覚的に区別できる（サイズ + ウェイト）か？

見つけた欠陥はすべて報告してください。一つでも欠陥があれば → REJECT。修正するまで納品しないでください。

**人間によるプレビュー（任意）。** ユーザーに論文を目視でプレビューしてもらいたい場合は、`officecli watch "$FILE"` を実行してユーザーが任意のタイミングで開けるライブプレビューを提供するか、`.docx` を Word / WPS / Pages で直接開いてもらってください。最終的な目視確認では、対象のビューアーでファイルを開いてください。

### 正直な限界

`validate` はスキーマエラーを捕捉しますが、学術スタイル上のエラーは捕捉しません。IEEE 論文に APA 式の引用があっても、脚注を禁じているスタイルに脚注があっても、新しい図を挿入するとずれるハードコードされた番号の図があっても、文書は `validate` を通過します。上記の Gate — 特に Gate 4（ラウンドトリップ）と Gate 5（SEQ の存在）— が、`validate` では捕捉できないものを捕捉する手段です。

## 既知の問題と落とし穴（学術固有）

→ 基本的な落とし穴（シェルエスケープ、`\$ \t \n` のリテラル、表セルの書式適用順、改ページの境界、`shd.fill` / `ind.firstLine` のスキーマ無効形式、TOC のキャッシュ値、透かしの二段階処理）については docx v2 §Known Issues & Pitfalls を参照。

学術固有の問題:

- **`\left(...\right)` / `\left[...\right]` の内側に上付き/下付きがあるとパースエラーになる**（`cast object … Subscript`）。外側のスクリプト（`\left(x+y\right)^2`）は問題ない。プレーンな `(`、`)`、`[`、`]` を使うこと — OMML が表示モードで自動サイズ調整する。
- **`/body/oMathPara[N]` に対する `move` は数式を並べ替える**（囲んでいる段落が移動する）。`--before <path>` は空の段落を残すことがあるため、`--index` / `--after` を推奨。
- **セクション区切りによる +1 段落のオフセット。** `add /body --type section` はそれぞれ一つの空段落を `/body` に挿入する。区切り後のすべての `p[N]` インデックスは +1 ずつシフトする。区切りは計画的に行い、`add section` の後は `officecli get "$FILE" /body --depth 1` で再インデックスすること。
- **`/section[last()]` は最終セクションに解決される**（get・set の両方で `p[last()]` と同様）。明示的な `/section[N]` も動作する:
  ```bash
  SECTION_COUNT=$(officecli query "$FILE" section --json | jq '.data.results | length')
  # 最終セクションには /section[last()] を使うか、明示的に /section[2]、/section[3]、... を使う
  ```
  `add /body --type section` のたびにカウントが増える。区切りのたびに再クエリすること。
- **複数段組は自動的に元に戻らない。** `columns=2` のセクションの後は、別のセクション区切りを追加し、新しい `/section[N]`（N = 復元後のカウント）に明示的に `columns=1` を設定しなければならない — さもないと、参考文献を含む文書の残り全体が二段組でレンダリングされてしまう。各 N について `officecli get "$FILE" "/section[N]"` で確認すること。
- **`tc[N]` のセルパスに対する `--type equation` はガードエラーで拒否される**（`table cells only accept paragraphs, tables, or SDTs`）— 不正な XML は書き込まれない。表セル内では、代わりに `tc[N]/p[1]` を `--prop mode=inline` で指定すること。
- **SEQ キャプション番号は `set / --prop recalcFields=seq` で埋められる**（すべてのキャプションが揃った後に一度実行）ものであり、フィールドごとの raw-set パッチではない。recalc 前の SEQ フィールドは `view text` で `#OCLI_NOTEVAL!` センチネルを表示する。
- **ぶら下げインデントの正規形は `indent=720 hangingIndent=720`。** `ind.firstLine=-720` ではない。ドット付きの形式は `<w:jc>` の後に `<w:ind>` を出力してしまい、出力時にスキーマ検証に失敗する。
- **脚注参照のランは `view annotated` で空文字列として表示される。** 参照側の `<w:footnoteReference>` という XML 要素には可視テキストがなく、注釈本体は `/footnotes/footnote[N]` にある。`view text` を目で追うのではなく、`officecli query "$FILE" 'footnote'` で確認すること。
- **キャプションの配置:** 表のキャプションは表の**上**、図のキャプションは図の**下**。主要なスタイル（APA、Chicago、IEEE、MLA）すべてがこの点で一致している。表のキャプションを表の下に置くのは学術スタイル上の誤りであり、レンダリングの問題ではない — `validate` はこれを捕捉しない。
- **TOC のキャッシュレンダリング / シェルエスケープ:** → see docx v2 §Table of Contents, §Shell escape。

## レンダラーの癖（ビューアー間の差異）

→ see docx v2 §Renderer quirks。PAGE / TOC のキャッシュ値、OMML のベースラインのずれ、テーマカラー — これらの癖はすべて学術論文にも同様に当てはまります。数式や引用マーカーが壊れていると判断する前に、ユーザーの対象ビューアー（Word、WPS、Pages）でファイルを開いてください — そこで正しくレンダリングされるなら、それはビューアーの癖であり、スキルの欠陥ではありません。

## ヘルプへのポインタ

迷ったときは: `officecli help docx`、`officecli help docx <element>`、`officecli help docx <element> --json`。ヘルプが権威あるスキーマであり、このスキルは docx v2 の上に乗る学術的な差分についての判断ガイドです。
