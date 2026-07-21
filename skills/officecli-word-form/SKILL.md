---
name: officecli-word-form
description: "実在する Content Controls（SDT）+ レガシー FormField チェックボックス + MERGEFIELD 差し込み印刷プレースホルダー + 文書保護を備えた、記入可能な Word フォーム（.docx）を作成する際にこのスキルを使用する。トリガー: 'fillable form'、'form fields'、'content controls'、'SDT'、'word form'、'fill in'、'only editable fields'、'protect document'、'onboarding form'、'HR intake'、'survey template'、'contract / SOW template'、'mail-merge template'、'compliance checklist'、'medical intake questionnaire'。出力は、特定のフィールドだけが編集可能で残りはロックされた単一の .docx。このスキルは独立しており、docx の上のシーンレイヤーではない — ペイロードは `<w:sdt>` + `<w:ffData>` + `<w:fldChar>` + `documentProtection` であり、これらはいずれも docx ベーススキルではカバーされない。ユーザーが記入可能なフィールドを持たない通常のレポート、レター、メモ、学術論文、ピッチデッキなどの文書には発火させず、officecli-docx またはそのシーンレイヤーへ振り分けること。"
---

# OfficeCLI Word-Form スキル

**このスキルは独立しており、docx の上のシーンレイヤーではない。** フォームのペイロード — `<w:sdt>` コントロール、`<w:ffData>` レガシーフィールド、`<w:fldChar>` 差し込み印刷、`documentProtection` — は、docx の段落/見出し/スタイルのプリミティブとは別クラスの要素である。QA の観点も異なる: docx の Delivery Gate は視覚的レイアウトとライブな PAGE フィールドを気にするが、このスキルの QA はデータの配線（保護が強制されているか／alias+tag／項目が注入されているか／name ≤ 20 文字／アンダースコアのアンチパターンがないか）を気にする。**逆方向のハンドオフ:** ユーザーの文書に記入可能なフィールドが一切ない（レポート、レター、メモ、論文、提案書）場合は `officecli-docx` または docx のシーンスキルへ振り分け、このスキルは使わないこと。

## 開始前に（重要）

**`officecli` が未インストールの場合:**

`macOS / Linux`

```bash
if ! command -v officecli >/dev/null 2>&1; then
    curl -fsSL https://d.officecli.ai/install.sh | bash
fi
```

`Windows (PowerShell)`

```powershell
if (-not (Get-Command officecli -ErrorAction SilentlyContinue)) {
    irm https://d.officecli.ai/install.ps1 | iex
}
```

確認: `officecli --version`

初回インストール後もなお `officecli` が見つからない場合は、新しいターミナルを開いて確認コマンドを再実行すること。

上記のインストールコマンドが失敗する場合（セキュリティポリシーによるブロック、ネットワーク未接続、権限不足など）は、https://github.com/iOfficeAI/OfficeCLI/releases からお使いのプラットフォーム用バイナリを手動でダウンロードしてインストールし、確認コマンドを再実行すること。

## ヘルプ最優先ルール

このスキルが教えるのは「実際のフォームに何が必要か」であり、CLI の全フラグではない。プロパティ／エイリアス／列挙値に確信が持てない場合は、推測する前にヘルプを確認する: `officecli help docx [element] [--json]`（例: `sdt`、`formfield`、`field`）。ヘルプはインストール済み CLI バージョンに紐づいており、常に正とする — このスキルとヘルプの内容が食い違う場合は**ヘルプが勝つ**（特に `sdt` に設定できるプロパティ集合は時間とともに増えてきているため、ハードコードされたリストではなく `help docx sdt` を信頼すること）。

## メンタルモデルと継承関係

Word フォームとは、素の docx スキルが触れない 4 つの OpenXML ペイロード層を加えた `.docx` である: **`<w:sdt>`** コンテンツコントロール（種類: text / richtext / dropdown / combobox / date / picture / group）、**`<w:ffData>`** レガシー FormField（今なお実在のチェックボックスを得る唯一の手段 — SDT の `type=checkbox` は未実装）、**`<w:fldChar>`** 複合フィールド（MERGEFIELD、REF、PAGEREF、SEQ、IF — テンプレート作成時のもので、エンドユーザーが記入するものではない）、そして **`documentProtection`**（Word 上でフィールド以外のテキストを読み取り専用にするロック機構 — CLI 上では `protection=forms` によりフィールド以外の内容の編集がロックされる（`--force` または `raw-set` が必要）が、フォームフィールド（SDT）の編集は許可され続ける。これこそが forms 保護の意義である）。

**docx v2 からの継承はない。** docx の Delivery Gate（カバー充填率、ライブ PAGE チェック）はここでは適用されない — フォームの QA は `view forms` + `query sdt alias+tag` + `protectionEnforced` である。

**docx への逆ハンドオフ。** レポート／レター／メモ／論文／ピッチデッキなど、編集可能なフィールドを一切持たない文書は `officecli-docx` に戻すこと。文書の目的がデータ収集またはテンプレート差し込みである場合に**このスキル**を使う。

## シェル実行の規律

**一度に一コマンド。出力を読んでから次へ。** OfficeCLI はインクリメンタルであり、`add` / `set` / `remove` の一つひとつが即座にファイルを変更する。以下のレシピはすべて `FILE=form.docx` というシェル変数を用いる。

**3 つのシェルエスケープの落とし穴:**

1. **`[N]` を含むパスは必ずクォートする** — zsh/bash は角カッコをグロブ展開する。`officecli get "$FILE" /body/sdt[1]` は `no matches found` で失敗する。正: `officecli get "$FILE" '/body/sdt[1]'`。
2. **`$` を含むプロパティはシングルクォートで囲む** — `"Total: $50,000"` は `$50` の変数展開後に `"Total: ,000"` になってしまう。正: `'Total: $50,000'`。
3. **`--after find:<text>` は外側をシングルクォートで囲む。内側にダブルクォートを使ってはならない** — `--after find:"Client Signature:"` はクォート自体が検索文字列の一部になってしまい、マッチに失敗する。正: `--after 'find:Client Signature:'`。

**`WARNING: UNSUPPORTED`（終了コード 2）は、静かに間違った要素ができているサイン。** CLI は拒否されたプロパティ**なしで**要素を作成している。ビルドログに UNSUPPORTED が一つでもあれば、現行 CLI がその要素で受け付けないプロパティ名があるということ — 立ち止まって `help docx <element>` で正しいプロパティ名を確認し（`items`/`format`/`lock` など、ほとんどの SDT プロパティは現在受理される。`maxlength` は受理されない）、コマンドを修正して再実行すること。そのまま出荷してはならない。

**`protection=forms` は最後に実行する構造コマンド。** これを設定すると、`protection=forms` はフィールド以外の内容の編集をロックする（それらの編集には `--force` または `raw-set` が必要）が、フォームフィールド（SDT）の編集は許可され続ける — これが forms 保護の意義である。フィールド以外の内容の編集（例: `set /body/p[N] --prop text=`）は拒否される（`ERROR: Document is protected … use --force`）。フォームフィールドの編集（例: `set /body/sdt[N] --prop text=` / `--prop alias=`）は成功する（終了コード 0）。記入可能なフィールドを探すには `Query("editable")` を使うこと。したがって、フィールド以外（静的レイアウト）の編集をすべて終えてからロックすること。ロック後に静的な内容を編集する必要が生じた場合は `--force` を渡す（あるいは一時的に保護を解除して編集し、再ロックする）。

### `--after find:` 早見表

`--after find:<text>` は**最初の**出現箇所にマッチする。アンカーを誤ると挿入位置がずれ、デバッグに手間がかかる。3 つのルール:

1. **アンカーは文書全体で一意でなければならない。** 二言語契約書で「甲方签字」は両当事者にマッチしてしまう — 「甲方签字（Service Provider）」のような一意なフレーズ、または英語の完全なタイトルを使うこと。
2. **挿入後は `/body/p[last()]` が当てにならない** — find による挿入で `<w:body>` の子要素の順序が変わるため。新しく挿入した段落を続けて操作するには、その実際の paraId を読み取る: `officecli query "$FILE" paragraph --json | jq -r '.data.results[-1].format.paraId'`。
3. **中国語＋全角カッコ `（）` は `find` にリテラルでマッチする**が、不安な場合はまず `officecli view "$FILE" text | grep -n "锚点"` でファイル中の実バイト列を確認すること。

```bash
# Trap: first-match hits 甲方 only, 乙方 missed
officecli add "$FILE" /body --type sdt --after 'find:签字'

# Fix: two signatories, two unique anchors
officecli add "$FILE" /body --type sdt --prop alias=Party_A_Name --prop tag=party_a \
  --after 'find:甲方签字（Service Provider）'
PID_A=$(officecli query "$FILE" paragraph --json | jq -r '.data.results[-1].format.paraId')
officecli add "$FILE" "/body/p[@paraId='$PID_A']" --type sdt --prop alias=Party_A_Title --prop tag=party_a_title
```

`--after find:` によるインライン SDT は、新しい段落としてではなく、マッチした段落の子として追加される — ラベルと SDT が同じ行を共有すべき場合にこれを使うこと。

## 実在のフォームとは何か（本質の定義）

実在の記入可能フォームには、**構造化フィールド** + **文書保護**の両方が必要である。

| アプローチ | Word ユーザーに見えるもの | CLI で読み取り可能か | 実在のフォームか |
|---|---|---|---|
| SDT コントロール + `protection=forms` | グレーの枠付きフィールド。それ以外はロック | `query sdt` / `view forms` | **YES** |
| FormField チェックボックス + `protection=forms` | 実際にクリックできるチェックボックス。それ以外はロック | `query formfield` / `view forms` | **YES**（チェックボックスのみ） |
| MERGEFIELD プレースホルダー | `«CustomerName»` が下流のエンジンによって差し込まれる | `query field` | **YES**（テンプレート作成時） |
| アンダースコア `___` / 空行 | 見た目だけ。文書全体が編集可能 | No — 構造化フィールドが存在しない | **NO** |

**アンダースコアでフィールドを模倣してはならない。** `姓名：_______________` は構造化データを一切生成せず、あらゆる検証をすり抜けてしまう。常に `--type sdt` または `--type formfield` を使うこと。

**チェックボックスは formfield であり、SDT ではない。** `--type sdt --prop type=checkbox` は終了コード 1（`SDT type 'checkbox' is not implemented`）で終了する。すべてのレシピにおいて、チェックボックスは常に `--type formfield --prop type=checkbox` を使う。

**MERGEFIELD は別トラックである。** `view forms` は SDT と formfield のみを列挙し、`query field` は複合フィールドのみを列挙する。互いに素な 2 つのインベントリであり、両方が同一ファイル内に共存できる。

## 出力に対する要件（最低ライン）

すべてのフォームはこれらを満たさなければならない — Delivery Gate がそれぞれを実行可能なチェックとして強制する。

1. `protection=forms` が強制されている（`get $FILE /` → `protectionEnforced=True`）。
2. すべての SDT が `alias` と `tag` の両方を持つ。
3. すべての dropdown/combobox が `view forms` 上で空でない `items=...` を持つ。
4. すべての date SDT が意図した `format=...` を表示する。
5. ロックされたすべての SDT が、意図どおり `lock=sdtLocked` / `contentLocked` / `sdtContentLocked` を表示する。
6. ビルドログに `WARNING: UNSUPPORTED` が一つもない。
7. どの SDT にも `type=checkbox` が一つもない。
8. すべての formfield の `name` が 20 文字以下である。
9. アンダースコア行／空行のプレースホルダーが一つもない。
10. フィールドの種類がユーザーの意図に一致する（短いテキスト／段落テキスト／固定リスト／リスト＋自由入力／日付／真偽値）。

## 3 つの経路（中核となる意思決定）

SDT のプロパティは `add` において**第一級**である — 現在のセットは `officecli help docx sdt` で確認すること。現時点では `type, tag, alias, text, items, format, lock, placeholder/placeholderText, date.*` と、SDT の種類 `text / richtext / dropdown / combobox / date / group / picture` が含まれる。したがって、ほとんどのフォームは純粋な `add` だけで作れる — raw-set も Word テンプレートも不要。残る 2 つの経路は、本当に到達不可能なケース専用である。**コマンドを一行書く前に経路を選ぶこと。**

### 経路 A — 純粋な CLI（ほぼ全ケースのデフォルト）

**使用場面**: text / richtext / dropdown / combobox / date / picture / group いずれの SDT でも — そのオプション、日付フォーマット、ロックを含めて。プロパティをそのまま `add` に渡し、`get '/body/sdt[N]' --json` でそれぞれが永続化されたことを確認する。

```bash
# dropdown WITH its options and a non-default date format, in one add each — no raw-set:
officecli add "$FILE" /body --type sdt \
  --prop type=dropdown --prop alias="Department" --prop tag=dept \
  --prop items="Engineering,Finance,HR"
officecli add "$FILE" /body --type sdt \
  --prop type=date --prop alias="Start Date" --prop tag=start \
  --prop format="yyyy年MM月dd日"
# lock works on add AND set:
officecli add "$FILE" /body --type sdt --prop type=text --prop tag=full_name \
  --prop alias="Full Name" --prop text="Enter full name" --prop lock=sdtLocked
# protection comes last, once all fields exist:
# officecli set "$FILE" / --prop protection=forms
```

### 経路 B — CLI + `raw-set` ブリッジ（ヘルプが公開していない属性のみ）

**使用場面**: `help docx sdt` が列挙していない SDT 属性が必要な場合（例: 表示テキストと異なる `w:value` を持つ `<w:listItem>`、あるいは `--prop` の存在しない sdtPr の子要素）。`raw-set` は OfficeCLI の万能な OpenXML フォールバックであり、`officecli --help` にトップレベルコマンドとして列挙されている。通常のドロップダウンには `--prop items=`（経路 A）を使うこと。raw-set はここでは例外であり、原則ではない。

```bash
# Display text ≠ stored value — not expressible via items=, so inject the listItems directly:
officecli add "$FILE" /body --type sdt --prop type=dropdown --prop alias="Department" --prop tag=dept
officecli raw-set "$FILE" /document \
  --xpath "//w:sdt[w:sdtPr/w:tag/@w:val='dept']/w:sdtPr/w:dropDownList" \
  --action append \
  --xml '<w:listItem xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:displayText="Engineering" w:value="ENG"/><w:listItem xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:displayText="Finance" w:value="FIN"/>'
```

### 経路 C — Word テンプレート（API が届かないもののみ）

**使用場面**: **実在の SDT チェックボックス**（`type=checkbox` は依然として終了コード 1 — 代わりにレガシー FormField を使う。§Legacy FormField 参照）、`placeholderDocPart` によるプロンプトテキストのパーツ、あるいは `--prop` の到達範囲を超えるカスタム richtext の外観／パート横断のネスト。picture と grouped の SDT は**もうここには含まれない** — 経路 A で問題なく追加できる。

```bash
# One-time in Word: Developer tab → Insert Content Control → Save as template.docx
cp templates/onboarding_with_signature.docx "$FILE"
officecli open "$FILE"
officecli view "$FILE" forms                 # inspect embedded controls + paths
officecli set "$FILE" '/body/sdt[@sdtId=3]' --prop text="Jane Smith"
officecli set "$FILE" / --prop protection=forms
```

### 決定テーブル

| 必要なもの | 経路 | 補足 |
|---|---|---|
| デフォルト文字列を持つ text / richtext SDT | **A** | `--prop type/alias/tag/text` |
| ロックが必要な text SDT | **A** | `--prop lock=sdtLocked` は `add`（および `set`）で機能する |
| **オプション付きの** dropdown / combobox | **A** | `--prop items="A,B,C"` |
| デフォルトでないフォーマットの date SDT | **A** | `--prop format="yyyy年MM月dd日"` |
| 署名 picture SDT、grouped SDT | **A** | `--prop type=picture` / `type=group` を直接 add できる |
| 保存値が表示テキストと異なる dropdown | **B** | raw-set で `<w:listItem w:value=…>` を append |
| 実在のチェックボックス | **FormField** | `--type formfield --prop type=checkbox`（§Legacy FormField 参照） |
| 差し込み印刷プレースホルダー | **MERGEFIELD** | `--type field --prop fieldType=mergefield`（§MERGEFIELD 参照） |
| 実在の SDT チェックボックス／プレースホルダーパーツ／カスタム外観 | **C** | Word でスケルトンを作り、CLI で埋める |

## クイックスタート — 経路 A + FormField（最小限のインテークフォーム）

SDT テキストフィールド 2 つ、チェックボックス 1 つ、保護。貼り付けて調整すればよい — これが出荷に値する最小のフォームである。

```bash
FILE=intake.docx
officecli close "$FILE" 2>/dev/null; rm -f "$FILE"   # preflight: clear stale resident / prior file (cold-start after CLI upgrade commonly leaks a resident)
officecli create "$FILE"
officecli open "$FILE"

officecli set "$FILE" / --prop title="Employee Onboarding Intake" \
  --prop docDefaults.font="Calibri" --prop docDefaults.fontSize="12pt"

officecli add "$FILE" /body --type paragraph \
  --prop text="Employee Onboarding Intake" --prop style=Heading1 \
  --prop size=20 --prop bold=true --prop spaceAfter=18pt

officecli add "$FILE" /body --type paragraph \
  --prop text="Full Name:" --prop size=11 --prop bold=true --prop spaceAfter=4pt
officecli add "$FILE" /body --type sdt --prop type=text \
  --prop alias="Full Name" --prop tag=full_name --prop text="Enter full name"

officecli add "$FILE" /body --type paragraph \
  --prop text="Start Date:" --prop size=11 --prop bold=true --prop spaceAfter=4pt
officecli add "$FILE" /body --type sdt --prop type=date \
  --prop alias="Start Date" --prop tag=start_date

officecli add "$FILE" /body --type paragraph \
  --prop text="Read and agree to employee handbook" --prop size=11 --prop spaceAfter=4pt
officecli add "$FILE" /body --type formfield \
  --prop type=checkbox --prop name=agree_handbook --prop checked=false

officecli set "$FILE" '/body/sdt[1]' --prop lock=sdtlocked
officecli set "$FILE" '/body/sdt[2]' --prop lock=sdtlocked
officecli set "$FILE" / --prop protection=forms
officecli close "$FILE"
officecli view "$FILE" forms
```

## 経路 B — raw-set レシピ

3 つのレシピで、SDT フォームにおける複雑属性のニーズのほぼすべてをカバーする。

### B1 — ドロップダウン項目（append）

```bash
# Skeleton (Path A)
officecli add "$FILE" /body --type sdt --prop type=dropdown \
  --prop alias="Department" --prop tag=dept

# Inject items
officecli raw-set "$FILE" /document \
  --xpath "//w:sdt[w:sdtPr/w:tag/@w:val='dept']/w:sdtPr/w:dropDownList" \
  --action append \
  --xml '<w:listItem xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:displayText="Engineering" w:value="Engineering"/><w:listItem xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:displayText="Finance" w:value="Finance"/><w:listItem xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:displayText="HR" w:value="HR"/>'

# Verify
officecli get "$FILE" '/body/sdt[1]'   # expect: type=dropdown items=Engineering,Finance,HR
```

**テンプレート。** 差し替えるのは `<TAG>` / `<LABEL>` / `<VALUE>` だけ。ルートの `<w:listItem>` にはすべて `xmlns:w=...` が必須 — raw-set は名前空間プレフィックスを継承しない。複数の `<w:listItem>` を一回の呼び出しで連結でき、オプションの順序は保持される。

### B2 — コンボボックス項目（B1 と同様、xpath の末尾のみ異なる）

```bash
officecli add "$FILE" /body --type sdt --prop type=combobox \
  --prop alias="Current Medication" --prop tag=current_med

officecli raw-set "$FILE" /document \
  --xpath "//w:sdt[w:sdtPr/w:tag/@w:val='current_med']/w:sdtPr/w:comboBox" \
  --action append \
  --xml '<w:listItem xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:displayText="Antihypertensives" w:value="Antihypertensives"/><w:listItem xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:displayText="Insulin" w:value="Insulin"/><w:listItem xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:displayText="Other (specify)" w:value="Other"/>'
```

B1 との唯一の違い: xpath 末尾が `w:comboBox` か `w:dropDownList` か。コンボボックスはユーザーによる自由入力を許すが、ドロップダウンは許さない。

### B3 — 日付フォーマット（setattr）

```bash
officecli add "$FILE" /body --type sdt --prop type=date \
  --prop alias="Contract Start Date" --prop tag=contract_start

# Chinese: yyyy年MM月dd日
officecli raw-set "$FILE" /document \
  --xpath "//w:sdt[w:sdtPr/w:tag/@w:val='contract_start']/w:sdtPr/w:date/w:dateFormat" \
  --action setattr \
  --xml "w:val=yyyy年MM月dd日"

# US:    w:val=MM/dd/yyyy
# ISO:   w:val=yyyy-MM-dd  (already the default)
# Long:  w:val="MMMM d, yyyy"

officecli get "$FILE" '/body/sdt[N]'   # expect: type=date format=yyyy年MM月dd日
```

`setattr` は属性を一つだけ置き換える — `--xml` 内で値をクォートしないこと。変更されるのは `w:val` のみで、`<w:dateFormat>` のラッパーは保持される。

### raw-set のアクションとエラー

| `--action` | フォームでの用途 |
|---|---|
| `append` | ターゲットの末尾に新しい子要素を挿入（B1、B2 — listItem） |
| `setattr` | 属性を一つ変更。`--xml "key=value"`（B3 — dateFormat/@val） |
| `replace` | ターゲット全体を置換（まれ — `<w:date>` ラッパー全体をリセットする場合） |
| `remove` | ターゲットを削除（再投入の前にオプションをクリアする） |

| 症状 | 対処 |
|---|---|
| `raw-set: 0 element(s) affected` | XPath がマッチしなかった。`tag` の値と、その SDT がブロックかインラインかを確認する。`officecli raw $FILE /document` にフォールバックして実際の XML を読むこと。 |
| `Error: prefix 'w' is not defined` | フラグメントに `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"` がない — `--xml` 内のすべてのルート要素にこれが必要。 |
| append 後に項目の読み取りが空 | `<w:dropDownList/>` があらかじめ存在している必要がある（経路 A の `type=dropdown` で保証される）。それがなければ append の挿入先がない。 |
| 成功と同じ行に `VALIDATION: N new error(s) introduced` | append によりスキーマ的に不正な子要素が導入された。`raw-set` が終了コード 0 でも、これは停止して修正すべき事象として扱うこと。 |

## 経路 C — Word テンプレートワークフロー

CLI では表現できないフィールド（署名 `picture` SDT、実在の SDT チェックボックス、`placeholderDocPart` のプロンプトテキスト、grouped SDT、カスタム richtext スタイリング）については、一度だけ Word でスケルトンを作り、その後 CLI で埋める。

**Word での一度限りの作業:** ファイル → オプション → リボンのユーザー設定 → 開発。開発タブ → 画像／チェックボックス／グループ化コンテンツコントロールの挿入 → 右クリック → プロパティ → タイトル（`alias`）とタグを設定する。`template.docx` として保存する。

**CLI で埋める:**

```bash
cp templates/onboarding_with_signature.docx "$FILE"
officecli open "$FILE"
officecli view "$FILE" forms                                    # see /body/... paths + sdtId values
officecli set "$FILE" '/body/sdt[@sdtId=3]' --prop text="Jane Smith"
officecli set "$FILE" / --prop protection=forms
officecli close "$FILE"
```

## MERGEFIELD（データ駆動トラック）

`help docx field` は `mergefield`、`ref`、`pageref`、`seq`、`if` を含む約 30 種類の `fieldType` 列挙値を宣言しており、いずれも型付きプロパティとともに CLI で表現できる。MERGEFIELD は同一ファイル内で SDT と共存できるが、`query field` のみに報告される。`view forms` は MERGEFIELD を列挙**しない**（ユーザーが記入するものではないため）。

**基本形の MERGEFIELD:**

```bash
officecli add "$FILE" /body --type paragraph --prop text="Dear "
officecli add "$FILE" '/body/p[1]' --type field --prop fieldType=mergefield --prop name=CustomerName
officecli add "$FILE" '/body/p[1]' --type run --prop text=", "
officecli add "$FILE" '/body/p[1]' --type field --prop fieldType=mergefield --prop name=CompanyName
# Readback: "Dear «CustomerName», «CompanyName»"
```

**要素タイプによるショートカット**（等価）: `officecli add "$FILE" '/body/p[1]' --type mergefield --prop name=CustomerName`。

### よくあるフィールドパターン

| パターン | 呼び出しの形 |
|---|---|
| 差し込み印刷プレースホルダー | `--type field --prop fieldType=mergefield --prop name=<FieldName>` |
| 数値ピクチャ付き差し込み印刷（金額、パーセント） | `--type field --prop fieldType=mergefield --prop name=Amount --prop instr='MERGEFIELD Amount \# "#,##0.00"'`。型付きの `format` プロパティは mergefield では無視される（警告を出力する）— フィールドコード全体を埋め込むには `instr`（エイリアス `instruction`）を使うこと。確認: `query "$FILE" field --json \| jq '.data.results[].format.instruction'` に `\#` とピクチャが含まれていること。 |
| 日付ピクチャ付き差し込み印刷 | `--type field --prop fieldType=mergefield --prop name=StartDate --prop instr='MERGEFIELD StartDate \@ "yyyy-MM-dd"'` |
| ブックマークテキストへの相互参照 | `--type field --prop fieldType=ref --prop name=<BookmarkName>` |
| ブックマークのページ番号への相互参照 | `--type field --prop fieldType=pageref --prop name=<BookmarkName>` |
| 自動採番（Figure 1 / 2 / 3） | `--type field --prop fieldType=seq --prop identifier=Figure` |
| フッターのページ番号 | `--type field --prop fieldType=page` |
| 「Page X of Y」 | 2 つのフィールド: `fieldType=page` + `fieldType=numpages` |
| 条件付きテキスト | `--type field --prop fieldType=if --prop expression='{ MERGEFIELD Gender } = "Male"' --prop trueText="Mr." --prop falseText="Ms."` |

### IF 条件分岐（CLI で表現可能）

```bash
officecli add "$FILE" /body --type paragraph --prop text=""
officecli add "$FILE" '/body/p[last()]' --type field --prop fieldType=if \
  --prop expression='{ MERGEFIELD Gender } = "Male"' \
  --prop trueText="Mr." --prop falseText="Ms."
officecli add "$FILE" '/body/p[last()]' --type run --prop text=" "
officecli add "$FILE" '/body/p[last()]' --type field --prop fieldType=mergefield --prop name=LastName
# Merge-time result: "Mr. «LastName»" or "Ms. «LastName»"
```

`{ IF { MERGEFIELD X } = "Y" { REF bm } "fallback" }` のようなネストされたラッパーは `--prop` の連結では表現できない — 手作りの `<w:fldChar>` / `<w:instrText>` フラグメントを raw-set するか、Word テンプレート（経路 C）で一度だけ作ること。

**読み取り確認。** `query $FILE field` は `/field[N]` とインストラクション、`fieldType` を列挙する。`view $FILE forms` は MERGEFIELD を列挙**しない**（SDT と formfield のみ）— これらはテンプレート作成時のものであり、エンドユーザーが記入するものではない。`get $FILE '/body/p[1]'` はギュメ（« »）で囲まれたフィールド名をレンダリングする。

## レガシー FormField

**実在のチェックボックスが必要な場合**に FormField を使う。テキスト／ドロップダウンには SDT を優先する。

`help docx formfield`: `type`（text/checkbox/check/dropdown）、`name`（必須、**20 文字以下** — OpenXML スキーマの MaxLength。`add` はそれより長くても通してしまうが `validate` が拒否する）、`text`（text のみ、エイリアス `value`）、`checked`（checkbox のみ）。

```bash
# CHECKBOX — the only real checkbox (SDT type=checkbox is not implemented)
officecli add "$FILE" /body --type formfield --prop type=checkbox \
  --prop name=agree_terms --prop checked=false

# TEXT formfield
officecli add "$FILE" /body --type formfield --prop type=text \
  --prop name=emp_name --prop text="Enter name"

# DROPDOWN formfield — items NOT settable via CLI; use Word template or SDT Path B
officecli add "$FILE" /body --type formfield --prop type=dropdown --prop name=dept_select

# Read / modify by name (stable) or 1-based index
officecli get "$FILE" '/formfield[agree_terms]'
officecli set "$FILE" '/formfield[agree_terms]' --prop checked=true
officecli set "$FILE" '/formfield[emp_name]' --prop text="Jane Smith"
officecli set "$FILE" '/formfield[dept_select]' --prop text="Engineering"
```

FormField のパス（`/formfield[N]` または `/formfield[<name>]`）は SDT のパス（`/body/sdt[N]`）とは別系統である。両者は共存でき、`protection=forms` は両方をカバーする。

**スケール。** 単一の文書内で 50 個以上のチェックボックスでテスト済み — formfield 数に実用上の上限はなく、ビルドと `validate` はクリーンなまま。`name` ≤ 20 文字（K13）だけが唯一の厳格な制約である。

**レンダラーに関する注意 — formfield チェックボックスの `[RENDERER-BUG]`。** LibreOffice の PDF エクスポートは、formfield のチェックボックスを `☐☐`（二重の枠）として描画することがある。Word と WPS は単一のクリック可能なボックス（トグルで ☑）として描画する。これは LibreOffice のレンダラー特有の癖であり、**スキルや文書品質の問題ではない** — K19 参照。フォーム側で回避策を試みないこと。評価者が LibreOffice 生成の PDF をスクリーンショットして `☐☐` を目にした場合は、`[RENDERER-BUG]` に起因すると説明すること。

## 文書保護とロック

### フォーム保護の有効化

```bash
officecli set "$FILE" / --prop protection=forms
officecli get "$FILE" /                                  # look for: protectionEnforced=True
```

### 保護モード

| モード | Word ユーザーができること | CLI の挙動 |
|---|---|---|
| `forms` | SDT と formfield のみ記入可能 | フォームフィールド（SDT）の編集は成功（終了コード 0）。フィールド以外の内容の編集には `--force` または `raw-set` が必要 |
| `readOnly` | 読み取りのみ | フィールド以外の編集には `--force` が必要。raw-set は保護を回避する |
| `comments` | コメント追加のみ | フィールド以外の編集には `--force` が必要。raw-set は保護を回避する |
| `trackedChanges` | 変更履歴付きでのみ編集可能 | フィールド以外の編集には `--force` が必要。raw-set は保護を回避する |
| `none` | 完全な編集が可能 | すべての操作が機能する |

**重要:** 文書保護は Word ユーザーと CLI の両方を制限する。`protection=forms` の下では、フォームフィールド（SDT）の編集 — `set /body/sdt[N] --prop text=` / `--prop alias=` — は成功する（終了コード 0）。これが forms 保護の意義である。フィールド**以外**の内容の編集（例: `set /body/p[N] --prop text=`）のみが `ERROR: Document is protected (mode: forms). … use --force to override` で拒否される。それでも静的な内容を編集したい場合は `--force`（または保護を完全に回避できる唯一の動詞である `raw-set`）を渡すこと。Word ユーザーが記入できるフィールドを見つけるには `Query("editable")` を使う。

### ロック値（`add` と `set` の両方に設定可能）

```bash
officecli set "$FILE" '/body/sdt[1]' --prop lock=sdtlocked           # content editable; control cannot be deleted
officecli set "$FILE" '/body/sdt[1]' --prop lock=contentlocked       # content read-only; control can be deleted
officecli set "$FILE" '/body/sdt[1]' --prop lock=sdtcontentlocked    # both locked
# Omit lock entirely → unlocked (default)
```

`--prop lock=...` は `set` だけでなく `add` でも機能する — コントロールを作成する時点でインラインに設定できる。読み取り時は入力の大文字小文字によらず camelCase（`sdtLocked`）に正規化される — どちらの表記も受理される。

### lock × `protection=forms` の相互作用

| lock 値 | `protection=forms` 有効時 | Word ユーザーが編集できるか | Word ユーザーがコントロールを削除できるか |
|---|---|---|---|
| （なし） | yes | **できる** | **できる** |
| `sdtlocked` | yes | できる | できない |
| `contentlocked` | yes | できない | できる |
| `sdtcontentlocked` | yes | できない | できない |
| ブロックレベル SDT ラップの `contentlocked` | いずれでも | できない（保護の有無にかかわらずラップされた段落は読み取り専用） | できない |
| いずれでも | `readOnly` モード | できない | できない |

### ブロックレベルロック（段落をラップする SDT）

`protection=forms` は文書レベルのものである — いったん管理者が保護解除すれば、静的な段落（免責事項、法的宣誓文、契約条項）はすべて再び編集可能になってしまう。マスターテンプレートには多層防御が必要である: 重要な段落を `lock=contentLocked` を持つブロックレベルの `<w:sdt>` でラップしておけば、保護が解除された後もその内容は読み取り専用のままになる。

```bash
officecli add "$FILE" /body --type paragraph \
  --prop text="I authorize the above and acknowledge all clauses." --prop size=11 --prop spaceAfter=12pt
PID=$(officecli query "$FILE" paragraph --json | jq -r '.data.results[-1].format.paraId')

# raw-set actions: append | prepend | insertbefore | insertafter | replace | remove | setattr
# No `wrap` action — two-step instead: (1) insertbefore an empty <w:sdt><w:sdtContent/></w:sdt>,
# (2) move the original <w:p> inside by `replace` on the sdtContent with a copy of the paragraph XML.
# Simpler alternative: read the paragraph XML via `officecli raw`, then `replace` the whole <w:p> with <w:sdt>...<w:sdtContent>[original w:p]</w:sdtContent></w:sdt>:
# NOTE: this needs an XML-aware extraction, NOT awk. `raw /document` emits the whole document on ONE line,
# so an awk range like `/paraId=.../,/<\/w:p>/` grabs the entire <w:document> (schema-invalid). Parse the XML
# and pull the single <w:p> by paraId instead:
PARA_XML=$(officecli raw "$FILE" /document | python3 -c '
import sys, xml.etree.ElementTree as ET
xml = sys.stdin.read(); pid = sys.argv[1]
ns = {"w":"http://schemas.openxmlformats.org/wordprocessingml/2006/main","w14":"http://schemas.microsoft.com/office/word/2010/wordml"}
for p, u in ns.items(): ET.register_namespace(p, u)
root = ET.fromstring(xml); key = "{%s}paraId" % ns["w14"]
for p in root.iter("{%s}p" % ns["w"]):
    if p.attrib.get(key) == pid:
        sys.stdout.write(ET.tostring(p, encoding="unicode")); break
' "$PID")
officecli raw-set "$FILE" /document \
  --xpath "//w:p[@w14:paraId='$PID']" \
  --action replace \
  --xml "<w:sdt xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"><w:sdtPr><w:alias w:val=\"Authorization\"/><w:tag w:val=\"auth_para\"/><w:lock w:val=\"contentLocked\"/></w:sdtPr><w:sdtContent>${PARA_XML}</w:sdtContent></w:sdt>"
```

`query sdt --json | jq '.data.results[] | select(.format.lock == "contentLocked" and .format.type == "block")'` で確認すること。法的宣誓文、コンプライアンス免責事項、機密保持条項にのみ使う — 通常のインテークフィールドにはここまでの手当ては不要である。

### ロールごとにロックするフィールド（複数ロールのフォーム）

一つのフォームを 2 つのロール（患者 vs 医師、甲 vs 乙）が記入する場合、相手のロールが触れてはいけないフィールドに `lock=contentLocked` を使う。`protection=forms` の下では、`contentLocked` の SDT は Word 上で読み取り専用として表示される。想定するロールが保護を解除する（あるいは管理者がロールごとのコピーを差し替える）ことで、もう一方の半分を記入できるようになる。

```bash
# Patient section — editable (no lock, or sdtlocked to prevent accidental deletion only)
officecli set "$FILE" '/body/sdt[1]' --prop lock=sdtlocked      # patient_name
officecli set "$FILE" '/body/sdt[2]' --prop lock=sdtlocked      # patient_dob

# Physician section — locked against patient edits
officecli set "$FILE" '/body/sdt[14]' --prop lock=contentLocked # physician_diagnosis
officecli set "$FILE" '/body/sdt[15]' --prop lock=contentLocked # physician_signature
```

これは医療インテーク、二当事者契約、順次承認フォームにおける中核パターンである。

## レシピ — MERGEFIELD + 署名付き契約書 / SOW テンプレート

3 つのサブレシピにわたる行対応表: SDT[1]=project_name、SDT[2]=contract_start、SDT[3]=payment_schedule、SDT[4]=signatory_name（インライン）。同じ `$FILE` に対して (sow-a) → (sow-b) → (sow-c) の順に実行する。各サブレシピは 20 行未満に収めているため、シェルエスケープの一つのミスが一つのブロックを超えて連鎖することがない。

### レシピ (sow-a) 定型文 + 表紙 + 当事者

ファイルを作成し、docDefaults を設定し、タイトルと導入文を書き、下流の差し込み印刷が埋める 2 つの MERGEFIELD プレースホルダー（`CustomerName`、`ContractNo`）を配置する。

```bash
FILE=sow.docx
officecli create "$FILE"
officecli open "$FILE"
officecli set "$FILE" / --prop title="Statement of Work" \
  --prop docDefaults.font="Calibri" --prop docDefaults.fontSize="12pt"

officecli add "$FILE" /body --type paragraph --prop text="Statement of Work" \
  --prop style=Heading1 --prop size=20 --prop bold=true --prop spaceAfter=12pt
officecli add "$FILE" /body --type paragraph \
  --prop text="This Statement of Work ('SOW') is entered into between the parties identified below and governs the delivery of professional services." \
  --prop size=11 --prop spaceAfter=12pt

officecli add "$FILE" /body --type paragraph --prop text="Customer: "
officecli add "$FILE" '/body/p[last()]' --type field \
  --prop fieldType=mergefield --prop name=CustomerName
officecli add "$FILE" /body --type paragraph --prop text="Contract #: "
officecli add "$FILE" '/body/p[last()]' --type field \
  --prop fieldType=mergefield --prop name=ContractNo
```

### レシピ (sow-b) SDT フィールド — すべて経路 A 経由

3 つのブロックレベル SDT（プロジェクト／日付／ドロップダウン）と、`--after 'find:Client Signature:'` でアンカーされたインライン署名 SDT を追加する。日付フォーマットとドロップダウンの項目はそのまま `add` に渡す — raw-set は不要。

```bash
officecli add "$FILE" /body --type sdt --prop type=text \
  --prop alias="Project Name" --prop tag=project_name --prop text="Enter project name"
officecli add "$FILE" /body --type sdt --prop type=date \
  --prop alias="Contract Start Date" --prop tag=contract_start --prop format="MM/dd/yyyy"
officecli add "$FILE" /body --type sdt --prop type=dropdown \
  --prop alias="Payment Schedule" --prop tag=payment_schedule \
  --prop items="Full Prepayment,Net 30 Upon Delivery"
officecli add "$FILE" /body --type paragraph --prop text="Client Signature:" \
  --prop bold=true --prop spaceBefore=18pt --prop spaceAfter=4pt
officecli add "$FILE" /body --type sdt --prop type=text \
  --prop alias="Signatory Name" --prop tag=signatory_name --prop text="Authorized Signatory" \
  --after 'find:Client Signature:'
# (Only reach for raw-set if a listItem's stored value must differ from its display text — see Path B.)
```

### レシピ (sow-c) 透かし + ロック + 文書保護

CONFIDENTIAL の透かしを配置し（親は `/` であり `/body` ではない）、3 つのブロックレベル SDT をロックし、インラインの signatory_name SDT をロックする方法を示し（パスは `view forms` の後にのみ判明する）、最後のコマンドとして `protection=forms` で文書を封印する。

```bash
officecli add "$FILE" / --type watermark \
  --prop text="CONFIDENTIAL" --prop color=FF0000 --prop rotation=315

officecli set "$FILE" '/body/sdt[1]' --prop lock=sdtlocked
officecli set "$FILE" '/body/sdt[2]' --prop lock=sdtlocked
officecli set "$FILE" '/body/sdt[3]' --prop lock=sdtlocked
officecli view "$FILE" forms   # copy signatory_name path, then: set '/body/p[@paraId=...]/sdt[1]' --prop lock=sdtlocked

officecli set "$FILE" / --prop protection=forms
officecli close "$FILE"
officecli query "$FILE" field     # expect 2 MERGEFIELDs: CustomerName, ContractNo
```

## 設計原則（フォーム）

**コントロール種別の決定木:**

```
Date → type=date | Fixed list → type=dropdown | List + custom → type=combobox
Short text → type=text | Long text → type=richtext | Boolean → formfield checkbox
```

**タイポグラフィのスケール。** 単位の落とし穴: `spaceBefore` / `spaceAfter` / `spaceLine` はデフォルトで**トゥイップ**（1/20 pt）単位 — 常に `spaceBefore=18pt` のように書くこと。

| 要素 | サイズ | スタイル | 間隔 |
|---|---|---|---|
| フォームタイトル（H1） | 20pt | 太字 | `spaceBefore=0pt`、`spaceAfter=12pt` |
| セクション見出し（H2） | 14pt | 太字 | `spaceBefore=18pt`、`spaceAfter=8pt` |
| フィールドラベル | 11pt | 太字 | `spaceAfter=4pt` |
| 説明文／注記 | 11pt | 斜体 `color=666666` | `spaceAfter=18pt` |

**アクセシビリティの引き上げ。** 医療／高齢者向け／アクセシビリティ重視のフォームでは、フィールドラベルと説明文を **12pt** に引き上げる（デフォルトの 11pt は高齢のユーザーには小さすぎる）。セクション見出しは 14pt のまま。

**CJK フォーム:** `docDefaults.font="Microsoft YaHei"` を設定すること — Calibri には中国語グリフがない。

**フィールドの順序。** (1) 個人情報／ID、(2) ロール／分類、(3) 日付、(4) 補足的な自由記述、(5) 確認／署名。

**Yes/No + 条件付きフォローアップ**（コンプライアンス／医療インテークで頻出）: formfield チェックボックスの後に、`alias` に手がかりを持たせた richtext SDT を続ける — 例えば `--type formfield --prop type=checkbox --prop name=has_cond` の後に `--type sdt --prop type=richtext --prop alias="If yes, explain" --prop tag=cond_detail --prop text="If yes, explain here"`。

**署名ブロックの順序。** ラベルはそれ自体の段落に、SDT は次の段落に置く（ラベルに `spaceBefore=18pt`、SDT に `spaceAfter=4pt`）。`Label: SDT` のようにインラインで並べてはならない — Word 上でランが接触して視覚的にくっついて見えてしまう。

**ビルド順序。** create+open → メタデータ → 構造（見出し、ラベル段落）→ SDT/formfield のスケルトン（経路 A の 4 プロパティ）→ 経路 B の注入 → フィールドごとのロック → `protection=forms` を最後に → close。

**ヘッダー／フッターに関する注意。** ヘッダー／フッターはセクション作成時に**あらかじめ定義済み**である（default/first/even、それぞれ 3 個ずつ）。最初の変更は既存パートに対する `set` でなければならず、`add` であってはならない — `add $FILE /header ...` は `already exists` を返すか、無言で何もしない。まず `officecli query "$FILE" header --json` で `type` の値を確認し、それから `officecli set "$FILE" '/header[@type=default]' --prop text=...` を実行する。`add` を使うのは、独自のヘッダー／フッターを持つ追加セクションを新たに作る場合のみ。

## バッチモード（概要）

コントロールの多いフォームでは、バッチによりオーバーヘッドを削減できる。経路 A と経路 B は一つのバッチ内で共存できる。

```bash
cat <<'EOF' | officecli batch "$FILE"
[
  {"command":"add","parent":"/body","type":"sdt","props":{"type":"text","alias":"Full Name","tag":"full_name","text":"Enter name"}},
  {"command":"add","parent":"/body","type":"sdt","props":{"type":"dropdown","alias":"Department","tag":"dept"}},
  {"command":"raw-set","part":"/document","xpath":"//w:sdt[w:sdtPr/w:tag/@w:val='dept']/w:sdtPr/w:dropDownList","action":"append","xml":"<w:listItem xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" w:displayText=\"Engineering\" w:value=\"Engineering\"/><w:listItem xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" w:displayText=\"Finance\" w:value=\"Finance\"/>"},
  {"command":"set","path":"/body/sdt[1]","props":{"lock":"sdtlocked"}},
  {"command":"set","path":"/body/sdt[2]","props":{"lock":"sdtlocked"}}
]
EOF
officecli set "$FILE" / --prop protection=forms
```

- `xml` 内の内側の `"` は `\"` でエスケープする。`$var` を展開させないために、シングルクォートのヒアドキュメント `<<'EOF'` を使うこと。
- **P0 バッチの落とし穴:** バッチ内で本当にサポートされていないプロパティは、**警告なしに**静かに破棄される（対話的な `add` なら `WARNING: UNSUPPORTED` を出力し終了コード 2 になる）。防御策: `help docx sdt` が列挙するプロパティ（`type/tag/alias/text/items/format/lock/...` は `add` で問題なく使える。`maxlength` は使えない）だけを送り、バッチ後に読み取り確認で検証すること。
- `batch` は `add`、`set`、`get`、`query`、`remove`、`validate`、`raw-set` をサポートする。

## Delivery Gate（実行可能なゲート）

すべてのフォームの後に、以下のゲートをすべて実行すること。各ゲートは `OK` 行を出力しなければならない。一つでも `REJECT` があれば納品してはならない。

```bash
# Assumes FILE=<your-form.docx>, document has been closed with officecli close "$FILE"

# Gate 1 — Validate (must be clean; protection=forms no longer produces a schema error)
VAL_OUT=$(officecli validate "$FILE" 2>&1)
VAL_ERRS=$(echo "$VAL_OUT" | grep -c '\[Schema\]')
if [ "$VAL_ERRS" -eq 0 ]; then echo "Gate 1 OK (validate clean)"
else echo "REJECT Gate 1: $VAL_ERRS schema errors"; echo "$VAL_OUT"; exit 1
fi

# Gate 2 — Token / placeholder leak (labels used as visual underscore substitutes)
LEAK=$(officecli view "$FILE" text | grep -niE '_{3,}|TBD|\(fill in\)|\{\{|xxxx|lorem|placeholder')
[ -z "$LEAK" ] && echo "Gate 2 OK (no underscore / placeholder leak)" || { echo "REJECT Gate 2:"; echo "$LEAK"; exit 1; }

# Gate 3 — At least one structured field exists
SDT_N=$(officecli query "$FILE" sdt --json | jq '.data.results | length')
FF_N=$(officecli query "$FILE" formfield --json | jq '.data.results | length')
FLD_N=$(officecli query "$FILE" field --json | jq '.data.results | length')
TOTAL=$((SDT_N + FF_N + FLD_N))
[ "$TOTAL" -gt 0 ] && echo "Gate 3 OK ($SDT_N sdt + $FF_N formfield + $FLD_N field)" || { echo "REJECT Gate 3: 0 structured fields — this is not a form"; exit 1; }

# Gate 4 — Every SDT has alias + tag (skill-imposed H2)
# NOTE: `query`/`get --json` wrap prop fields under `.data.results[N].format.{prop}` — use `.data.results[0].format.alias` / `.format.tag`, never bare `.alias` or `.data.format`.
SDT_MISSING=$(officecli query "$FILE" sdt --json | jq '[.data.results[] | select(.format.alias == null or .format.alias == "" or .format.tag == null or .format.tag == "")] | length')
[ "$SDT_MISSING" -eq 0 ] && echo "Gate 4 OK (every SDT has alias+tag)" || { echo "REJECT Gate 4: $SDT_MISSING SDT(s) missing alias or tag"; exit 1; }

# Gate 5 — Protection enforced + per-field lock inventory
PROT=$(officecli get "$FILE" / --json | jq -r '.data.results[0].format.protection // "none"')
[ "$PROT" = "forms" ] && echo "Gate 5 OK (protection=forms enforced)" || { echo "REJECT Gate 5: protection is '$PROT', expected 'forms'"; exit 1; }
officecli view "$FILE" forms | head -40   # visual spot-check: every dropdown shows items=; every date shows format=; every locked SDT shows lock=

# Gate 6 — No type=checkbox leaked onto any SDT
BAD_CB=$(officecli query "$FILE" sdt --json | jq '[.data.results[] | select(.format.type == "checkbox")] | length')
[ "$BAD_CB" -eq 0 ] && echo "Gate 6 OK (no SDT checkbox — formfield only)" || { echo "REJECT Gate 6: $BAD_CB SDT with type=checkbox"; exit 1; }
```

**`view issues` がゲートでない理由。** これはプローズ寄りのチェック（先頭行のインデント、見出しサイズ）しか実行せず、フォームのすべてのラベルを `Body paragraph missing first-line indent` として誤検知してしまう — フォームに対しては偽陽性の雪崩となる。このスキルでは無視すること。`validate`（スキーマの整合性）と `view forms`（フィールドの一覧）を使うこと。

## 既知の問題

| # | 問題 | 挙動 | 回避策 |
|---|---|---|---|
| K1 | SDT `type=checkbox` は未実装 | `add ... --type sdt --prop type=checkbox` → `Error: SDT type 'checkbox' is not implemented`、終了コード 1（現在のエラーメッセージは group/picture を含むサポート済みの種類を列挙する） | `--type formfield --prop type=checkbox` を使う |
| K3 | SDT `maxlength` は `add` で UNSUPPORTED | `WARNING: UNSUPPORTED: maxlength`、終了コード 2。要素自体は作成される（`items` / `format` / `lock` / `placeholderText` は現在 `add` でサポートされている — 拒否されるのは `maxlength` のみ。`name` は受理されるが SDT では no-op — `alias`/`tag` を使うこと） | 長さの強制は下流で行う。初期コンテンツには `text` を使う |
| K4 | SDT の `items` / `format` / `type` は作成**後**には設定不可 | `set --prop items=...` → `UNSUPPORTED props (use raw-set instead)`（`add` 時にはこれらは設定可能 — 作成時にインラインで設定すること） | `add` 時に設定する。後で変更するには経路 B の `raw-set` か、`remove` して再 `add` する |
| K5 | FormField `maxlength` が UNSUPPORTED | `WARNING: UNSUPPORTED: maxlength`。formfield 自体は作成される | 長さの強制は下流の検証で行う |
| K6 | FormField dropdown の `items` が UNSUPPORTED | dropdown formfield は空のオプションリストで作成される | 代わりに `--prop items=` を持つ SDT dropdown を使う |
| K7 | Watermark の `width` / `height` は設定不可 | それらなしで透かしが作成される（`opacity` は `add` で設定可能で読み取りも可能。実サイズは自動計算される） | 正確なサイズが必要な場合は Word で開いて図形を調整する（Phase 2） |
| K9 | バッチモードは本当に UNSUPPORTED なプロパティを静かに破棄する | `WARNING` 行は出ない。プロパティ（例: `maxlength`）が破棄されていてもバッチは「N succeeded」と報告する | バッチ内の SDT エントリは `help docx sdt` が列挙するプロパティに留め、バッチ後に読み取り確認で検証する |
| K13 | FormField `name` が 20 文字超 | `add` は警告なしで終了コード 0 を返す。後で `validate` が `/w:ffData/w:name` に対して `[Schema] ... MaxLength=20` を報告する | `name` は 20 文字以下に保つ（OpenXML スキーマの制限）。SDT の `alias` / `tag` にはこの制限がない |
| K14 | 段落上の `shd.fill` はスキーマ的に不正な `<w:pPr>/<w:shd>` を出力する | `validate` はインスタンスごとに 2 件のスキーマエラーを報告する（`unexpected child element`、`required attribute 'val' missing`）。Word はそれでも描画してしまう | 代わりにラン上でハイライトを適用する（`shading=HEX`、フラットで正規）、あるいはラン内の `<w:rPr>` に `<w:shd w:val="clear" w:fill="HEX"/>` を raw-set する |
| K15 | `view forms` は MERGEFIELD を列挙**しない** | 出力に含まれるのは SDT と formfield のみ。MERGEFIELD はテンプレート作成時のものであり、エンドユーザーが記入するものではない | `query field` と `view forms` を互いに素な 2 つのインベントリとして扱う。すべてのレシピで両方を確認する |
| K16 | ヘッダー／フッターはセクション作成時にあらかじめ定義済み（default/first/even、それぞれ 3 個ずつ） | 最初の呼び出しで `add $FILE /header ...` は `already exists` を返すか、無言で何もしない | 最初の変更は既存パートに対する `set` を使う: `officecli query $FILE header --json` で `type` を確認し、`set '/header[@type=default]' --prop text=...` を実行する。`add` を使うのは、まったく新しいセクションのヘッダー／フッターの場合のみ |
| K17 | ヘッダーに注入された透かしはスキーマ的に不正な `<w:noProof>` の子要素を出力する | `validate` は `/header[N]/w:sdt/.../w:noProof` に追加の `[Schema]` エラーを報告する — 無害ではなく、修正すべき本物のスキーマエラーである | `add $FILE / --type watermark` の後、ヘッダーパートごとに一度: `officecli raw-set $FILE /word/header1.xml --xpath "//w:noProof" --action remove`（`header2.xml`、`header3.xml` が存在する場合は繰り返す） |
| K18 | `query`/`get --json` はプロパティフィールドを `.data.results[N].format.{prop}` の下にラップする | 生の `.alias` / `.tag` / `.protection`、あるいは `.data.format.*` に対して jq を書くと null や 0 マッチになり、Gate 4/5 が誤って「missing=N」を報告してしまう | `.data.results[].format.alias` / `.format.tag` を使う。`get /` の場合は（`.data.format.protection` ではなく）`.data.results[0].format.protection` を使う。`.format.type` / `.format.paraId` も同様 |
| K19 | LibreOffice は PDF エクスポートで formfield チェックボックスを `☐☐`（二重の枠）として描画する | 見た目だけの問題 — Word / WPS は単一のボックスとして描画し、クリックでトグルできる（☑）。LibreOffice のレンダラー特有の癖であり、[RENDERER-BUG] としてフラグされる | スキル側で「修正」しようとしないこと。評価者が LibreOffice 生成の PDF をスクリーンショットして `☐☐` を目にした場合は、フォーム品質の欠陥ではなく [RENDERER-BUG] に起因するとする |

## Phase 2 — Word での仕上げ

一部の仕上げは CLI の対象範囲外である。これらは人間にファイルを渡すこと。いずれも有効なフォームであるために必須ではない。

| 必要なもの | Word で開く理由 |
|---|---|
| 特定のロックを持つ実在の SDT チェックボックス | `type=checkbox` は終了コード 1 — 開発タブ → チェックボックスコンテンツコントロールを使う（または CLI 経由のレガシー FormField チェックボックス） |
| プロンプトテキスト（「Click here to enter a date」） | `/word/glossary/document.xml` 内の `placeholderDocPart` が必要 |
| カスタム richtext のデフォルト外観 | Word のスタイルペインで参照先のスタイルを調整する |
| 透かしのリサイズ | `width` / `height` は設定不可 — 図形のハンドルをドラッグする |

（`picture` と `group` の SDT は `--type sdt --prop type=picture` / `type=group` で直接 add できる — もはや Phase 2 の項目ではない。）

最初の 4 項目については、一度だけスケルトンを作り（経路 C）、それを再利用すること。

## ヘルプへの導線

迷ったときは: `officecli help docx`、`officecli help docx <element>`、`officecli help docx <element> --json`。ヘルプは正となるスキーマであり、このスキルはその上で実在の記入可能な Word フォームを構築するための意思決定ガイドである。
