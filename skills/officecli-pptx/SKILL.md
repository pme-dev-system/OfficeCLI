---
name: officecli-pptx
description: "このスキルは、.pptx ファイルが入力・出力のいずれか（または両方）で関わるあらゆる場面で使用します。対象には次が含まれます: スライドデック、ピッチデック、プレゼンテーションの作成；任意の .pptx ファイルからのテキストの読み取り・解析・抽出；既存プレゼンテーションの編集・修正・更新；スライドファイルの結合や分割；テンプレート、レイアウト、スピーカーノート、コメントの操作。ユーザーが 'deck'、'slides'、'presentation'、'pitch' に言及した場合、または .pptx ファイル名を参照した場合は必ず発動します。"
---

# OfficeCLI PPTX スキル

## セットアップ

`officecli` が見つからない場合：

- **macOS / Linux**: `curl -fsSL https://d.officecli.ai/install.sh | bash`
- **Windows (PowerShell)**: `irm https://d.officecli.ai/install.ps1 | iex`

`officecli --version` で確認してください（PATH が反映されていない場合は新しいターミナルを開いてください）。インストールに失敗した場合は https://github.com/iOfficeAI/OfficeCLI/releases からバイナリをダウンロードしてください。

## ⚠️ Help 優先ルール

**このスキルは「良いスライドとは何か」を教えるものであり、すべてのコマンドフラグを網羅するものではありません。プロパティ名、列挙値、エイリアスに確信が持てない場合は、推測する前にまず help を参照してください。**

```bash
officecli help pptx                         # pptx の全要素を一覧表示
officecli help pptx <element>               # 要素の完全なスキーマ（例: shape, chart, animation, connector, zoom, group）
officecli help pptx <verb> <element>        # 動詞スコープのヘルプ（例: add shape, set slide）
officecli help pptx <element> --json        # 機械可読なスキーマ
```

help はインストール済みの CLI バージョンを反映します。スキルの記載と help が食い違う場合、**help が正です**。即座に help を実行すべきトリガー: `UNSUPPORTED props:` 警告、未知のアニメーションプリセット、`connector.shape=` の列挙値のずれ、プロパティ名 vs エイリアスの混同（`lineWidth` vs `line.width`、`color` vs `font.color`）。

## シェルと実行の規律

**シェルクォーティング (zsh / bash)。** 要素パスは常にクォートしてください（`"/slide[1]/..."`）— zsh はクォートされていない `[1]` をグロブとして扱い `no matches found` になります。エスケープは 3 層で行われ、それぞれ分離して考える必要があります（2 層目は CLI が代わりに処理します）:

1. **シェル。** 値の中の `$` はシェルに属します — 値全体をシングルクォートしてください: `--prop text='$15M'`。ダブルクォートの `"$15M"` はシェルにより `M` に展開されてしまいます。CLI は `\$` を代わりにアンエスケープしてくれません。
2. **CLI (`text=`)。** 2 文字のエスケープ `\n` と `\t` は解釈されます — pptx / docx / xlsx を通じて一貫して、`\n` は改行（段落区切り）、`\t` はタブです。テキストに文字通りのバックスラッシュ+n を出したい場合は二重にします（`\\n`）。これが必要になることは稀です。
3. **JSON (batch heredoc)。** 以下のレシピは `cat <<EOF | officecli batch` を**クォートせずに**パイプし、本体内で `$SLIDE` / `$FILE` を展開させています。同じ理由で、クォートなしの heredoc は値中のリテラル `$` も展開してしまいます: `"$1.42"` は無音のうちに `.42` に、`"$2.4B"` は `.4B` になります（`$1` / `$2` は空のシェルパラメータ）。**通貨は `\$` としてエスケープ**してください — `"text":"\$1.42"` — これなら `$SLIDE` の展開は維持されます。JSON 内の `\n` / `\t` はどちらの書き方でも動作します。全体をクォートした `<<'EOF'` はすべての `$` を保護しますが、その場合 `$SLIDE` も展開されなくなるため、本体にシェル変数が一切ない場合にのみ使ってください。金額を書き込んだ後は `view text` で `$` が生き残っているか確認してください。

迷ったら、書き込み後に `view text` を実行して文字単位で比較してください。

**インクリメンタルな実行。** 1 コマンド実行 → 終了コード確認 → 続行。50 コマンドのスクリプトが 3 番目で失敗すると、それが静かに連鎖します。構造的な操作（新規スライド、チャート、アニメーション、コネクタ）の後は、さらに積み上げる前に `get` を実行してください。

## 出力物の要件

これらは、あらゆるデックが満たすべき納品物の基準です。いずれか一つでも違反すれば、内容の質にかかわらず「未完了」です。

### すべてのデック共通

**1 スライド 1 アイデア。** スライドが 2 つ目のタイトルで内容を説明する必要があるなら、分割してください。「X についてのすべて」を詰め込んだ密なスライドは、3 秒以内に聴衆を失います。関連する 1 アイデアのスライドをグループ化するにはセクション区切りを使い、メガスライドにはしないでください。

**明示的なタイプ階層 — テーマのデフォルトに頼らないこと。** テーマのデフォルトはマスター間でずれます。すべてのテキストシェイプに明示的にサイズを設定してください。

| 要素 | 最小 | 一般的 | シェイプの最小高さ |
|---|---|---|---|
| スライドタイトル | **≥ 36pt** 太字 | 36–44pt | ≥ 2cm |
| セクション / サブタイトル | ≥ 20pt | 20–24pt | ≥ 1.2cm |
| 本文テキスト | **≥ 18pt** | 18–22pt | ≥ 1cm |
| キャプション / 軸ラベル | ≥ 10pt ミュート色 | 10–12pt | ≥ 0.6cm |

経験則: **シェイプの最小高さ ≈ フォントサイズ(pt) × 0.05cm**。0.8cm 高のボックスに入れた 18pt のサブラベルはオーバーフローします — `view annotated` で検出できます。

タイトルは**本文サイズの 2 倍以上**にしてください（36pt / 20pt は成立しますが、28pt / 20pt では弱く見えます）。本文 ≥ 18pt の正当な例外は 4 つ: チャート軸ラベル、凡例、フッター / ページ番号、5 単語以下の KPI サブラベル（例: "Active users"）。説明文は ≥ 18pt にすること。本文は左揃え、中央揃えはタイトルとヒーロー数値のみ。「カードが収まらない」場合はフォントを縮めるのではなくカードを減らしてください。

**フォントは最大 2 種、パレットは 1 つ。** 見出し用フォント 1 つ + 本文用フォント 1 つ（例: Georgia + Calibri）— 3 つ目の*ディスプレイ*書体は、大きな数字表示やカバータイトルのみに限り可（その場合も見出し+本文のペアは維持すること）。ブランドカラーは、支配色 1 つ（重み 60–70%）+ 補助色 1 つ + アクセント 1 つ。本文コンテンツで 4 色以上を混在させないこと。**Design Principles にあるパレットとフォントの組み合わせは床（最低ライン）であり、メニューではありません:** ユーザーがブランドカラー/フォントや既存テンプレートを指定した場合はそちらを優先してください。そうでない場合、記載の組み合わせはキャリブレーション済みの「種」であり、混ぜても離れても構いません — 結果がそれらより*悪く*ならず、コントラストの床をクリアしている限り。

**すべてのスライドが情報を伝える非テキストのビジュアルを持つこと。** シェイプ、チャート、アイコン、意味を持つグラデーションバンド — 装飾ではなく。箇条書きだけのデックは Word 文書と入れ替え可能です。例外: 引用のみのスライド、コードブロック、単一のサマリーテーブルスライド。

**足すより引く — すべての要素が存在理由を持つこと。** 上記のビジュアルルールは箇条書きの壁を防ぐためのものであり、雑然とさせてよいという許可ではありません。情報を伝えない装飾的な統計、アイコン、埋め草のセクション（「データスロップ」）で埋めないこと。スライドが空に感じる場合は、内容を捏造するのではなくレイアウトと余白で対処してください — スコープを削る方を選び、大きな追加が必要なら勝手に足さずフラグを立てること。

**すべてのコンテンツスライドにスピーカーノート。** `--type notes --prop text="..."`。話者にはスクリプトが必要であり、聴衆はスライドをそのまま読み上げられるべきではありません。

**コピーは人間らしく、AI らしくないこと。** タイトルは内容を軸にし、オチを狙わないこと。「X ではない。Y だ。」のような作られた対比、演出された緊張感、偽の洞察（「魔法の瞬間」）、一語のドラマ（「勢い。」）は禁止。誇張形容詞（シームレス、堅牢、ゲームチェンジング）は削り、数字自身に語らせてください。

**既存テンプレートを尊重すること。** ファイルに既にテーマとマスターがある場合は、それに合わせてください。既存の慣習がこれらのガイドラインより優先されます。

### ビジュアル納品の床（すべてのデックに適用）

完了と宣言する前に、スライド単位のレンダリング（QA 参照）は以下を満たす必要があります:

- **プレースホルダートークンがコンテンツとしてレンダリングされていないこと。** `{{name}}`、`$fy$24`、`<TODO>`、`lorem`、`xxxx`、チャートタイトル内の空の `()`/`[]` が現れないこと。
- **端からのオーバーフローや、シェイプ内でのテキストのクリッピングがないこと。** `view issues` が両方をフラグします（`shape_off_slide` + テキストフィットのヒント）。クリップの修正: ボックスを大きくするか値を短くする — コンテンツを収まるように削らないこと。
- **カバーが方向づけとなる要素を備えていること。** タイトル + サブタイトル + 発表者/クライアント + 日付 + ブランドバンドまたはキーテイクアウェイの帯 — タイトルのみのカバーはスタブに見えます。周囲に余白をたっぷり取るのは正しく、リッチさは詰め込み過ぎとは違います。
- **コントラスト。** `view issues` は最も一般的なケース — シェイプ自身の濃い塗りの上の不透明な濃い色のテキスト（`low_contrast`）— を自動でフラグします。それ以外（アイコン / チャート系列の塗り、スキーム/継承色、または*別の*背景シェイプの上のテキスト）は検出できません。そのため、明度 30% 未満の塗り（`1E2761`、`36454F`、深い森/ベリー/チェリー系）では、すべての本文ラン、カード本文、チャート系列、アイコンが `FFFFFF` または明度 80% 超であることを確認してください — 中間グレー（`6B7B8D` ≈ 44%）はノートPCでは読めても投影では消えます。濃色塗りのパス後、`view html` でスポットチェックしてください。

いずれかが失敗した場合は、完了と宣言する前に STOP して修正してください。

## デザイン原則

デックは文書ではありません。聴衆は各スライドを把握するのに 3 秒しかありません。何かを追加する前に問うべきこと: 「聴衆が最大の要素だけを読んで一瞥するだけで、要点が伝わるか？」箇条書きを読まないと伝わらないなら、最大の要素が間違っています。

### グリッド、余白、ネガティブスペース

標準ワイドスクリーンは **33.87 × 19.05cm**。内部では 12 カラムグリッドとして扱ってください:

- **端の余白 ≥ 1.27cm**（0.5"）を全辺に。
- **ブロック間のギャップ ≥ 0.76cm**（0.3"）— カード / カラム / 行の間。値は 1 つ（0.76 または 1.27cm）に決めてどこでも使うこと — ギャップが混在すると未完成に見えます。
- **スライドあたり ≥ 20% のネガティブスペース。** すべてのピクセルを埋めるのは素人っぽく見えます。
- **構成せよ、Web センタリングするな。** 余白は構造の一部です — 下三分の一に開いたスペースを持つ上重心のスライドは、正しい構図であり空の欠陥ではありません。意図的な非対称（コンテンツを左、余白を右）はすべてを中央揃えするより意匠を凝らして見えます — スペースがあるからといって埋める必要はありません。
- カードグリッドの場合: `usable = 33.87 − 2·margin − (N−1)·gap`、その上で `col_width = usable / N`。x 座標を手で選ばないこと。

### フォントの組み合わせ

文書のトーンに合わせて組み合わせてください。新奇さでは選ばないこと。「Best For」列はプロンプトであり指令ではありません — このテーブルの外の組み合わせでも合っていれば問題ありません。この 8 種は種であり、全集合ではありません。

| Header | Body | Best For |
|---|---|---|
| Georgia | Calibri | フォーマルなビジネス、金融、経営層向けレポート |
| Arial Black | Arial | 大胆なマーケティング、製品ローンチ |
| Calibri | Calibri Light | クリーンなコーポレート、ミニマルデザイン |
| Cambria | Calibri | 伝統的なプロフェッショナル、法務、学術 |
| Trebuchet MS | Calibri | フレンドリーなテック、スタートアップ、SaaS |
| Impact | Arial | 大胆な見出し、イベントデック、基調講演 |
| Palatino | Garamond | エレガントなエディトリアル、ラグジュアリー、非営利団体 |
| Consolas | Calibri | 開発者ツール、技術/エンジニアリング |

両方のフォントを、テーマ継承ではなく、すべてのシェイプに明示的に設定してください（タイトルには `--prop font=Georgia`、本文には `--prop font=Calibri`）。

### 色とコントラスト

各列: **Primary**（支配色 — 重みの 60–70%、最初に目に入る色）、**Secondary**（補助トーン）、**Accent**（控えめに、一点強調用）、**Text**（明るい塗りの上の本文）、**Muted**（キャプション / 軸ラベル / フッター）。

| テーマ | Primary | Secondary | Accent | Text | Muted |
|---|---|---|---|---|---|
| Coral Energy | `F96167` | `F9E795` | `2F3C7E` | `333333` | `8B7E6A` |
| Midnight Executive | `1E2761` | `CADCFC` | `FFFFFF` | `333333` | `8899BB` |
| Forest & Moss | `2C5F2D` | `97BC62` | `F5F5F5` | `2D2D2D` | `6B8E6B` |
| Charcoal Minimal | `36454F` | `F2F2F2` | `212121` | `333333` | `7A8A94` |
| Warm Terracotta | `B85042` | `E7E8D1` | `A7BEAE` | `3D2B2B` | `8C7B75` |
| Berry & Cream | `6D2E46` | `A26769` | `ECE2D0` | `3D2233` | `8C6B7A` |
| Ocean Gradient | `065A82` | `1C7293` | `21295C` | `2B3A4E` | `6B8FAA` |
| Teal Trust | `028090` | `00A896` | `02C39A` | `2D3B3B` | `5E8C8C` |
| Sage Calm | `84B59F` | `69A297` | `50808E` | `2D3D35` | `7A9488` |
| Cherry Bold | `990011` | `FCF6F5` | `2F3C7E` | `333333` | `8B6B6B` |

デフォルトではなくトピックに応じて選んでください — 金融には Midnight Executive、製品ローンチには Coral Energy、安全/LOTO には Cherry Bold が合います。最も近い名前付きテーマがぴったりでない場合はブレンドしてください（例: Forest の primary + ゴールド `D4A843` の accent）。明るい塗りには **Text**、キャプション/軸/フッターには **Muted**、暗い塗りの上の本文には `FFFFFF` または Secondary を使ってください。

暗い背景では、テキストとチャート系列は上記の Hard rules のコントラストの床に従ってください。

### チャート選択の決定テーブル

誤ったチャートタイプは 3 秒テストを台無しにします:

| データの形状 | 使う | 避ける |
|---|---|---|
| カテゴリ比較（A vs B vs C） | `column`（縦） / `bar`（≥ 6 カテゴリ、横） | 円グラフ（スライスが混ざる）、折れ線（時間軸がない） |
| 時系列、1–3 系列 | `line` | 面グラフ（オクルージョン）、棒グラフ（離散を暗示） |
| 全体に対する部分、2–5 スライス | `pie` / `doughnut` | 8 スライス以上の円グラフ（読めない） |
| 相関 / 分布 | `scatter` | 折れ線（順序を暗示） |
| 複数カテゴリ × 指標、密 | 積み上げ `column` またはヒートマップ | 指標ごとに 1 チャート — 統合すべき |
| KPI スナップショット（単一の大きな数字） | **大きなテキストシェイプ**（60–72pt + 5 単語以下のサブラベル）、チャートではない | ゲージチャート、小さな棒グラフ |

経験則: 系列が 3 を超え、カテゴリが 8 を超える場合は、2 つのチャートに分割するか表に切り替えてください。

### アニメーション

ブランドとコンテンツが求める分だけ使ってください — フォーマルな金融デックはほぼゼロに寄り、製品ローンチはより表現力豊かで構いません。アニメーションはツールであり装飾ではありません。それがデックを損なわないようにする 3 つの床があります（どれも使用量の上限を定めるものではありません）:

- **目的があること** — それぞれが明かすか強調する（段階的な箇条書きの表示、積み上がっていくチャート）ためのものであり、装飾のためではない。理解を助けないなら削ること。
- **段階的に劣化すること** — pptx のアニメーションはビューア（Keynote / Slides / Web / モバイル）間でレンダリングが一貫せず、まったく再生されないこともあるため、すべてのスライドは*静止フレーム*としても正しく読めなければならない。必須のコンテンツを表示アニメーションの裏に隠さないこと。
- **実機で検証すること** — アニメーションはランタイムのみの機能であり、`view html` やスクリーンショットでは見えないため、出荷前に実際のプレゼンテーションビューアで確認すること。

好みの指針（禁止ではない）: `fade` / `appear` / 単発の `zoom-entrance`（キビキビした数百 ms のデュレーション）はほとんどのデックに合います。`bounce` / `swivel` / `spin` / `fly-from-edge` / 密な複数オブジェクトの振り付けは、たいてい素人っぽく見えます — ブランドが意図的にプレイフルな場合にのみ使ってください。

### レイアウトパターンとデータ表示

スライド間でレイアウトを変化させてください — 同じパターンの繰り返しはすべてのスライドを同じに感じさせます。以下はよくある構成要素であり全集合ではありません — スライドごとに 1 つ選ぶか、内容が求めるならテーブル外のレイアウトを組み立ててください:

| パターン | 使う場面 | 主要な採寸 |
|---|---|---|
| **2 カラム**（左テキスト、右ビジュアル） | コンセプト + エビデンス；機能 + スクリーンショット | 各カラム ≈ 14–15cm；ギャップ 1cm |
| **アイコン行**（塗りつぶし円のアイコン + 太字見出し + 説明） | 機能一覧、メリット、チームの役割 | アイコン円 1.5–2cm；行数は最大 3–4 |
| **2×2 または 2×3 グリッド**（カードタイル） | 象限分析、SWOT、選択肢比較 | ギャップ ≥ 0.76cm；一貫したカード高さ |
| **ハーフブリード画像**（左または右半分いっぱい、反対側にコンテンツをオーバーレイ） | ヒーロー的な瞬間、事例紹介の導入 | 画像幅 16–17cm；コンテンツカラム ≥ 14cm |
| **大きな統計コールアウト**（60–72pt の数字 + 5 単語以下のサブラベル） | 単一の KPI、マイルストーン、市場規模 | シェイプを使う、チャートではない；サブラベルは 14–16pt ミュート色 |

**データ表示のクイックルール:**
- 比較カラム（前/後、A vs B）は、2–3 の選択肢であれば表より優れています。
- タイムラインとプロセスフロー: 番号付きのステップシェイプ + コネクタを使い、箇条書きにはしないこと。

### 画像の扱い（スライドが写真/スクリーンショット/ロゴを使う場合のみ）

**まず画像を読む**（ファイルを開く）— 見た内容から扱いを選んでください。ファイル名だけで盲目的に配置しないこと。

- **フルブリード写真** → 領域を COVER するようにサイズ調整（端をクロップ）、ボーダーなし。
- **スクリーンショット / 図 / ロゴ** → FIT するようにサイズ調整（コンテンツを絶対にクロップしないこと）。透過画像や FIT 画像はコントラストのある塗りの上に置く — 背後に色付きの矩形を敷き、白地の上に浮かせないこと。
- **写真の上のテキスト** → 画像に直乗せしないこと。カードに乗せるか、画像とテキストの間に保護スクリム（不透明度 ~50–60% の暗い矩形、またはテキスト端からフェードするグラデーション）を挟むこと。
- 決して引き伸ばさない（アスペクト比を歪めない）；ごちゃついたスクリーンショットにテキストをオーバーレイしないこと。
- ユーザー提供の画像/ブランドアセットを優先し、指示がない限り絵文字や自作イラストは使わないこと。

### ビジュアルモチーフの一貫性

1 つの特徴的な要素（角丸の画像フレーム、塗りつぶし円の中のセクション番号、片側のボーダー帯、対角のアクセントストライプ）を選び、すべてのスライドに貫いてください — デック全体で一貫させること；1 枚だけスタイリングして残りを素のままにすると放棄されたように見えます。副次的なモチーフは、主モチーフと競合しない限り可。ビルドプランの中で最初に宣言してください: `## Motif: numbered circles in brand color`

### 避けるべきビジュアルの AI っぽさ

- **スライドタイトルの下に装飾的なアンダーラインを引かないこと。** 見出しの下のストライプ/罫線は、AI 生成スライドの最も一般的な特徴です — 代わりに余白か背景色の変化を使ってください。
- **角丸カードに色付きの左ボーダーアクセントストライプをつけないこと。** もう一つの典型的な AI スライドの特徴です — 代わりに単色塗り、トップのアクセントバンド、または余白による分離を使ってください。
- **アイコンとして絵文字を使わないこと** — ブランドが使っている場合を除く。シェイプか本物のアイコンアセットを使ってください。

コピーレベルの特徴は「コピーは人間らしく」を参照。

## 一般的なワークフロー

1. **開く/保存のライフサイクル。** 最初に `officecli open <file>`、最後に `officecli save <file>` で編集をディスクにフラッシュしてください。`save` は書き込むだけで、後続の編集のためにレジデントをウォームに保ちます。`officecli close <file>` に手を伸ばすのは、レジデントを即座に解放したい場合（ワンショットのハンドオフ）のみにしてください。両方とも常に安全です — エラーになったり作業が失われたりすることはありません。繰り返しのシェイプグリッドには `batch` を使ってください。**非 officecli の境界でのみフラッシュしてください:** officecli 自身の読み取りは常にあなたの編集を反映します；`save`/`close` は、非 officecli プログラム（python-pptx、PowerPoint、レンダラー、納品）がファイルを読む前にのみ実行してください。
2. **方向づけ。** 新規デック: `officecli create "$FILE"`。既存デック: まず `officecli view "$FILE" outline`。盲目的に編集しないこと。
3. **タイトルシーケンスを先に（計画するだけで、まだ構築しない）。** スライドやシェイプを作る前に、順序付けられたスライドタイトルの全リストを書き出してください。タイトルだけを読んだ人が議論の流れを追えないなら、今のうちに直してください — 14 枚組み上げた後より、リストの段階で直す方が安上がりです。1 つのタイトル文法を選び（すべて名詞句か、すべて行動文か、混在させないこと）、通して維持してください（「コピーは人間らしく」参照）。
4. **表示順に構築。** 聴衆視点の順序でスライドを追加してください: カバー → アジェンダ → セクション 1 の区切り → セクション 1 のコンテンツ → セクション 2 の区切り → … → クロージング。スライド追加時の `--index` も機能しますが、線形の追加の方がビルドスクリプトを読みやすく保ち、インデックス演算のバグを避けられます。**最終納品前に、スライド数 + 物語の流れがビルドプランと一致することを確認してください。** Gate 3 の順序サニティチェックは、カバーが 14 枚中 1 枚目のはずが 11 枚目になっているようなケースを捕捉します。
5. **スライドごとにインクリメンタルに。** スライド + 背景を作成し、次にタイトル、次に補助シェイプ / チャート / コネクタ。カスタムデザインには常に `layout=blank`。構造的な操作の後は毎回 `get /slide[N] --depth 1` でシェイプ ID を確認してください。
6. **仕様どおりにフォーマット。** Requirements テーブルに従うこと；フォーマットは磨き上げではなく納品物です。
7. **保存 + 検証。** `officecli save` がファイルをディスクにフラッシュします（または `officecli close` でフラッシュしつつセッションも終了）。出荷前には必ず対象のプレゼンテーションビューアで開いて確認してください — チャートの色、アニメーション、フォント、ズームはランタイム機能であり `view html` ではレンダリングできません。完全な検証は下記の QA を参照。
8. **QA — 問題があると仮定する。** 1 サイクルで新規の問題がゼロになるまで、修正と検証を繰り返してください。

## クイックスタート

最小限の実用的なデック: カバー + コンテンツスライド 1 枚 + ノート。`$FILE` はあなたのファイル名の代わりです。

```bash
FILE="deck.pptx"
officecli create "$FILE"
officecli open "$FILE"

# カバー — 暗い塗り、中央揃えのタイトル
officecli add "$FILE" / --type slide --prop layout=blank --prop background=1E2761
officecli add "$FILE" /slide[1] --type shape --prop text="FY26 Strategic Review" \
  --prop x=2cm --prop y=7cm --prop width=29.87cm --prop height=3cm \
  --prop font=Georgia --prop size=44 --prop bold=true --prop color=FFFFFF --prop align=center

# コンテンツ — 白い塗り、タイトル + 本文 + ノート
officecli add "$FILE" / --type slide --prop layout=blank --prop background=FFFFFF
officecli add "$FILE" /slide[2] --type shape --prop text="Revenue grew 18% YoY" \
  --prop x=1.5cm --prop y=1.2cm --prop width=30cm --prop height=2cm \
  --prop font=Georgia --prop size=36 --prop bold=true --prop color=1E2761
officecli add "$FILE" /slide[2] --type shape --prop text="Enterprise renewals + new EMEA region drove the beat; NRR held at 118%." \
  --prop x=1.5cm --prop y=4cm --prop width=30cm --prop height=3cm \
  --prop font=Calibri --prop size=20 --prop color=333333
officecli add "$FILE" /slide[2] --type notes --prop text="Lead with the 18% beat, preview EMEA."

officecli save "$FILE"
officecli validate "$FILE"
```

すべてのビルドの型: open → スライド+背景 → タイトル → 本文 → ノート → save → validate。

## 読み取りと分析

広く始めて、絞り込んでいく。まず `outline`、どこを見るべきか分かってから `view text` / `get` / `query`。

```bash
officecli view "$FILE" outline          # スライド数 + タイトル
officecli view "$FILE" annotated        # フォント、サイズ、表、チャートを含む完全なスライド別内訳
officecli view "$FILE" text --start 1 --end 5   # テキストダンプ（表セルのテキストも含む）
officecli view "$FILE" issues           # 空のスライド、オーバーフローのヒント
officecli view "$FILE" stats            # カウント + 合計（alt 未設定の画像を含む）
```

**単一要素の確認。** XPath 風のパス、1-based。常にクォートすること。位置指定の `[N]` より `@name=` / `@id=` セレクタを優先してください（並べ替えに対して安定します）。`[last()]` も使えます。機械可読な出力には `--json` を追加してください。

```bash
officecli get "$FILE" "/slide[1]" --depth 1              # ID と名前を含むシェイプ一覧
officecli get "$FILE" "/slide[1]/shape[@name=Title]"
officecli get "$FILE" "/slide[1]/table[1]" --depth 3     # 表の行 / セル
```

**デック全体を横断してクエリ。** CSS ライクなセレクタ；演算子は `=`、`!=`、`~=`、`>=`、`<=`、`[attr]`、`:contains()`、`:no-alt`。`help pptx query` にクエリ可能な要素タイプの一覧があります。

```bash
officecli query "$FILE" 'shape:contains("Revenue")'
officecli query "$FILE" 'picture:no-alt'                 # アクセシビリティのギャップ
officecli query "$FILE" 'shape[fill=1E2761]'             # 色のマッチ
officecli query "$FILE" 'shape[width>=10cm]'             # 数値
```

**`query --json` の出力スキーマ。** 結果は `.data.results[]` にラップされます — `jq -r '.data.results[0].format.id'` であり `.[0].id` ではありません。シェイプ名は `.name`；塗りは `.format.fill`；テキストカラーは `.format.textColor`。

**ビジュアルプレビュー（推奨）。**

```bash
officecli view "$FILE" html                # HTML プレビューのパスを表示；スライド単位のビジュアル監査には Read で読み込むこと（最良の構造的な正解データ）
officecli view "$FILE" svg --start 3 --end 3   # 単一スライドの SVG（チャートとグラデーションは SVG ではレンダリングされない）
```

**出力の読み方 — これは想定内であり欠陥ではない:**
- **`layout=blank` にはタイトルプレースホルダーがありません。** タイトルは単なる `shape` 要素なので、`view outline` が `(untitled)` と報告するのは**想定どおり**であり欠陥ではありません。スクリーンリーダーのアウトライン互換性が問題になる場合のみ `layout=title` + `placeholder[title]` を使ってください。

## 作成と編集

動詞: `add` / `set` / `remove` / `move` / `swap` / `batch` / `raw-set`。デックの 9 割はスライド、シェイプ、テキスト、少数のチャート、画像、コネクタで構成されます。

### スライドと背景

スライドは `/slide[N]`。カスタムデザインには常に `layout=blank` を渡してください。背景: 単色、グラデーション、または画像。

```bash
officecli add "$FILE" / --type slide --prop layout=blank --prop background=1E2761                 # 単色
officecli add "$FILE" / --type slide --prop layout=blank --prop "background=1E2761-CADCFC-180"   # グラデーション（開始色-終了色-角度）
officecli add "$FILE" / --type slide --prop layout=blank --prop background=image:/path/to/hero.jpg  # 画像背景（推奨）
```

### シェイプ

`shape` はテキスト、塗り、枠線、位置、任意のアニメーション/リンクを保持します。

```bash
officecli add "$FILE" /slide[2] --type shape --prop name=Title --prop text="Key Insight" \
  --prop x=2cm --prop y=2cm --prop width=20cm --prop height=3cm \
  --prop font=Georgia --prop size=36 --prop bold=true --prop color=1E2761 --prop fill=none
```

配置は明示的です — レイアウトエンジンはなく、グリッドの計算はあなた自身が持ちます。`--prop preset=` でジオメトリを選びます（`rect`、`roundRect`、`ellipse`、`triangle`、`arrow`、`star5` など）；カスタムの `M...Z` パスはサポートされないため、プリセットから選んでください。**作成時にシェイプに名前を付け**（`--prop name=HeroTitle`）、後から `"/slide[N]/shape[@name=HeroTitle]"` でアドレスしてください — 名前は z 順序の変更や削除→追加を跨いで生き残りますが、位置指定の `/shape[3]`（さらには `@id=` も）はずれます。構造的な変更の後、位置指定インデックスを使う前に必ず `get --depth 1` し直してください。

### シェイプ内のテキスト（段落、ラン、スタイリング）

シェイプは段落（`paragraph[K]`）とラン（`run[K]`）を持ちます。1 行のテキストならシェイプの `--prop text=` で十分です；テキスト中の `\n` は段落区切りに、`\t` はタブになります（詳細は「シェルと実行の規律」参照；リテラルには `\\n` を二重にしてください）。`add --type paragraph` はシェイプと同じスタイルプロパティ（text、align、bold、italic、size、color、font）を取ります。1 行*内*で混在したスタイリングをするには、スタイル付きのランを追加してください:

```bash
officecli add "$FILE" "/slide[2]/shape[@name=Card1]/paragraph[1]" --type run \
  --prop text=" (inline detail)" --prop size=14 --prop italic=true --prop color=8899BB
```

### チャート

Design Principles のチャート選択テーブルに従ってチャートタイプを選んでください。完全なプロパティ一覧（chartType 列挙、`seriesN.*`、`data=`/`categories=`、軸オプション）: `help pptx add chart`。ブランドカラーを使った典型的なマルチ系列:

```bash
officecli add "$FILE" /slide[3] --type chart --prop chartType=column \
  --prop series1.name=Revenue --prop series1.values="42,45,48" --prop series1.color=1E2761 \
  --prop series2.name=Growth  --prop series2.values="2,7,7"    --prop series2.color=CADCFC \
  --prop categories="Q1,Q2,Q3" \
  --prop x=2cm --prop y=4cm --prop width=20cm --prop height=10cm
```

注意点: (1) `()`、`[]`、`TBD` を含むチャートタイトルはリテラルなテキストとして出荷されます。(2) 一部のビューアはチャートの色をテーマのデフォルトに正規化します — 対象のビューアで確認してください。系列は作成後にも追加できます（`add --type series`）。

### 画像

```bash
officecli add "$FILE" /slide[4] --type picture --prop src=hero.jpg \
  --prop x=1cm --prop y=1cm --prop width=32cm --prop height=18cm \
  --prop alt="Product hero, gradient lit from right"
```

`officecli query "$FILE" 'picture:no-alt'` で確認してください — 納品前には空である必要があります。

### コネクタ（推奨 — フローチャート/意思決定ツリーが第一級サポート）

2 つのシェイプまたは自由座標の間に線を描きます。完全なプロパティ/列挙リファレンス（`shape`、`headEnd`/`tailEnd` の値、`from`/`to` の参照形式）: `help pptx add connector`。

```bash
officecli add "$FILE" /slide[5] --type connector \
  --prop "from=/slide[5]/shape[@name=BoxA]" --prop "to=/slide[5]/shape[@name=BoxB]" \
  --prop shape=elbow --prop color=333333 --prop tailEnd=triangle
```

**すべてのフローコネクタに矢印が必要です。** ないと `bentConnector3` は方向のない線としてレンダリングされます。`preset=rightArrow` のオーバーレイは水平フローにのみ有効です；分岐するエッジを持つ菱形/意思決定ツリーには `tailEnd=` が必要です。

### アニメーション（推奨）

上記のアニメーションの床（目的があること、段階的に劣化すること、実機で検証すること）に従って使ってください。プリセット名 + デュレーション構文: `help pptx animation`。

```bash
officecli set "$FILE" "/slide[2]/shape[@name=HeroCard]" --prop animation=fade-entrance-400
officecli set "$FILE" "/slide[2]/shape[@name=HeroCard]" --prop animation=none    # すべてクリア
```

### ハイパーリンク、ツールチップ、スライドジャンプ

デック内ジャンプには `--prop link=slide[N]`（1-based；対象スライドが存在すること）、名前付きナビゲーションには `link=nextslide` / `firstslide` / `lastslide` / `previousslide` / `endshow`、URL には `link=https://...`、ホバーテキストには `--prop tooltip="..."`。

### 表、プレースホルダー、グループ、ズーム — ワンライナー集

- **表** — `--type table --prop rows=N --prop cols=M`。行レベルの `set` は `height` と `c1/c2/c3`（セルテキストの初期値）をサポートします。ヘッダー行のスタイリングは表レベル（`firstRow=true` / `headerFill=`）であり、行のプロパティではありません。セルのフォーマットはセルの段落/ランに属します。表レベルのフォントを設定する前に行を投入してください（行操作でフォントのカスケードがリセットされます）。
- **プレースホルダー** — `"/slide[N]/placeholder[title]"` / `placeholder[body]`。プレースホルダーを持つレイアウトを使うスライドでのみ利用可能（`layout=blank` では不可）。
- **グループ**（推奨） — 子要素には `"/slide[N]/group[@name=G]/shape[1]"` でアドレスしてください。位置指定インデックスより並べ替えに対して安定します。
- **ズームスライド**（推奨） — `--type zoom --prop target=N`（対象ごとに 1 リンク；エイリアス `slide`）。複数対象のナビゲーションハブには N 個の別々のズームシェイプを発行してください。ズームはランタイム機能です — `view html` は静的なジオメトリを表示しますが、ズームのインタラクション自体はライブのプレゼンテーションビューアでのみ動作します。
- **スライドコメント** — `/slide[N]/comment[M]` に紐づくレビュー者の注釈。フルライフサイクル（`add / set / get / query / remove`）。プロパティ: `text`、`author`、`initials`（自動導出）、`date`（ISO 8601、デフォルトは UtcNow）、`x` / `y`（長さのアンカー）。
  ```bash
  officecli add "$FILE" "/slide[2]" --type comment --prop author="Alice" --prop text="Tighten this bullet" --prop x=20cm --prop y=3cm
  officecli query "$FILE" 'comment' --json | jq '.data.results | length'   # すべてのレビューコメントを数える
  officecli remove "$FILE" "/slide[2]/comment[1]"                           # 対応後にクローズ
  ```

### デックレベルのレシピ

プリミティブからは自明ではないパターン。それぞれ**ビジュアルな結果**を先に示し、次に実行可能なブロックを示します。`$FILE` = あなたのファイル名。直前に追加したスライドをアドレスするには `/slide[last()]` を使ってください。以下のレシピは**構造と座標計算**を示すものです — このトピック向けに選んだパレット/フォントに差し替えてください；紺色 `1E2761` + Georgia は例として使ったテーマにすぎず、必ず真似るべきハウススタイルではありません。

**Z 順序。** 後から追加されたシェイプが上に来ます。背景装飾を先に、タイトルを最後に追加してください。後から直す場合: `--prop zorder=back/front`（兄弟要素を再採番するので、さらに積み上げる前に `get --depth 1` し直してください）。

#### (a) カバー（およびセクション区切り）

**ビジュアルな結果。** 濃紺の塗り、中央揃えの 44pt タイトル、18pt のアイスブルーのメタ行。

```bash
officecli add "$FILE" / --type slide --prop layout=blank --prop background=1E2761
officecli add "$FILE" "/slide[last()]" --type shape --prop text="Strategic Growth Review" \
  --prop x=2cm --prop y=7cm --prop width=29.87cm --prop height=3cm \
  --prop font=Georgia --prop size=44 --prop bold=true --prop color=FFFFFF --prop align=center
officecli add "$FILE" "/slide[last()]" --type shape --prop text="Prepared for Acme Leadership — FY26 Outlook" \
  --prop x=2cm --prop y=11cm --prop width=29.87cm --prop height=1.2cm \
  --prop font=Calibri --prop size=18 --prop color=CADCFC --prop align=center
```

**セクション区切り** = カバーと同じ構成に加え、セクションタイトルの背後に来るよう最初に追加された巨大な半透明の番号（`size=120`、`opacity=0.15`）。

#### (b) データスライド（チャート + コメンタリーブロック）

**ビジュアルな結果。** 左三分の二: ブランド系列色の縦棒グラフ。右三分の一: 20pt 見出し + 18pt 本文の「Key Insight」カード — 聴衆は棒を読み解く前に要点を読みます。

```bash
officecli add "$FILE" / --type slide --prop layout=blank --prop background=FFFFFF
officecli add "$FILE" "/slide[last()]" --type shape --prop text="FY26 Revenue Beat Plan by 18%" \
  --prop x=1.5cm --prop y=1cm --prop width=30cm --prop height=1.8cm \
  --prop font=Georgia --prop size=36 --prop bold=true --prop color=1E2761

# チャート — 左 2/3（`$` があるためタイトルはシングルクォート）
officecli add "$FILE" "/slide[last()]" --type chart --prop chartType=column \
  --prop series1.name=Actual --prop series1.values="42,45,48,55" --prop series1.color=1E2761 \
  --prop series2.name=Plan --prop series2.values="40,42,45,48" --prop series2.color=CADCFC \
  --prop categories="Q1,Q2,Q3,Q4" --prop x=1.5cm --prop y=3.5cm --prop width=20cm --prop height=14cm --prop title='FY26 Revenue ($M)'

# コメンタリーカード — 右 1/3: 背景 + 見出し + 本文
officecli add "$FILE" "/slide[last()]" --type shape --prop preset=roundRect --prop fill=F5F7FA --prop line=none \
  --prop x=22.5cm --prop y=3.5cm --prop width=9.8cm --prop height=14cm
officecli add "$FILE" "/slide[last()]" --type shape --prop text="Key Insight" \
  --prop x=23cm --prop y=4cm --prop width=9cm --prop height=1.2cm \
  --prop font=Georgia --prop size=20 --prop bold=true --prop color=1E2761
officecli add "$FILE" "/slide[last()]" --type shape --prop text="EMEA launch + NRR at 118% drove 12pp of the 18pp beat." \
  --prop x=23cm --prop y=5.5cm --prop width=9cm --prop height=11cm \
  --prop font=Calibri --prop size=18 --prop color=333333
```

#### (c) フローチャート / プロセス図（ボックス + コネクタ）

**ビジュアルな結果。** y=8cm に横並びの角丸ボックス 4 つ、それぞれ 6×3cm、紺/アイスブルー交互、肘型コネクタと三角矢印で接続。

グリッド計算（4 ボックス、33.87cm スライド、余白 1.5cm）: `gap = (33.87 − 3 − 24) / 3 = 2.29cm`。x 座標: `1.5, 9.79, 18.08, 26.37`。

各ボックスは `valign=middle` により自身でラベルを持ちます（別のオーバーレイシェイプは不要）。座標計算をポータブルにするため `batch` heredoc を使ってください — `bc` も bash 配列も不要です。

```bash
cat <<EOF | officecli batch "$FILE"
[
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"name":"Step1","preset":"roundRect","fill":"1E2761","line":"none","x":"1.5cm","y":"8cm","width":"6cm","height":"3cm","text":"Step 1","font":"Georgia","size":"20","bold":"true","color":"FFFFFF","align":"center","valign":"middle"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"name":"Step2","preset":"roundRect","fill":"CADCFC","line":"none","x":"9.79cm","y":"8cm","width":"6cm","height":"3cm","text":"Step 2","font":"Georgia","size":"20","bold":"true","color":"1E2761","align":"center","valign":"middle"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"name":"Step3","preset":"roundRect","fill":"1E2761","line":"none","x":"18.08cm","y":"8cm","width":"6cm","height":"3cm","text":"Step 3","font":"Georgia","size":"20","bold":"true","color":"FFFFFF","align":"center","valign":"middle"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"name":"Step4","preset":"roundRect","fill":"CADCFC","line":"none","x":"26.37cm","y":"8cm","width":"6cm","height":"3cm","text":"Step 4","font":"Georgia","size":"20","bold":"true","color":"1E2761","align":"center","valign":"middle"}}
]
EOF

# コネクタパターン — ボックス間の任意のグラフに再利用可能
for pair in "Step1 Step2" "Step2 Step3" "Step3 Step4"; do
  A=${pair% *}; B=${pair#* }
  officecli add "$FILE" "/slide[$SLIDE]" --type connector \
    --prop "from=/slide[$SLIDE]/shape[@name=$A]" \
    --prop "to=/slide[$SLIDE]/shape[@name=$B]" \
    --prop shape=elbow --prop color=333333 --prop tailEnd=triangle
done
```

`shape=elbow` が正式です（`bentConnector2` / `bentConnector3` も受け付けられます）。

#### (d) 複数スライドのデックの骨格

コードブロックはありません — これはリズムの問題です。以下の並びは**動作するカデンスの一例（濃色の区切りと白いコンテンツの交互）を示すものであり、必須の順序ではありません** — まずコンテンツから実際のアーク（流れ）を導き（「タイトルシーケンスを先に」参照）、その上で合う区切り/コンテンツのリズムを借用してください:

- **10 枚のレビュー:** カバー・アジェンダ・KPI×3・区切り01・チャート・チャート・区切り02・フロー・タイムライン・クロージング
- **20 枚のピッチ:** 同じリズムを×2、Problem・Solution・Market・Product・Traction・Model・Team・Financials・Ask のセクションに分割
- すべての区切りは、そのセクションのコンテンツより**前に**現れなければなりません（Gate 3 の順序サニティ）
- カバー/区切り = (a)；チャートページ = (b)；プロセスページ = (c)；KPI ページ = (e)；意思決定ページ = (f)

#### (e) KPI コールアウト — 巨大数字のカードグリッド

**ビジュアルな結果。** 1 行に 3〜4 つの巨大な数字；各カードは単位のサブラベル + 小さなパーセント変化のチップ + 1 行のテイクアウェイで構成。経営層向けデックで最も一般的な要素。

**サイズのルール。** 60pt Georgia 太字は 9.78cm のカードにおおよそ 5 文字収まります（`$84.2`、`118%`、`24.5`）。より長い値（`$84.2M`）は分割してください: 大きな数字として `$84.2`、サブラベルとして `USD millions` — 単位のサフィックスを追うためにフォントを縮めないこと、単に折り返されるだけです。

グリッド計算（3 カード、余白 1.5cm、ギャップ 0.76cm）: `col_width = (33.87 − 3 − 1.52) / 3 = 9.78cm`。x 座標: `1.5, 12.04, 22.58`。リスクが 1 秒で読み取れるよう、単一の「要注意」カードにアクセントカラーを使ってください。

```bash
# 2 枚のカード: 紺色の標準カード + テラコッタ色の要注意カード。それぞれ = 背景 + 大きな数字 + サブラベル + チップ
cat <<EOF | officecli batch "$FILE"
[
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"preset":"roundRect","fill":"1E2761","line":"none","x":"1.5cm","y":"4cm","width":"9.78cm","height":"7cm"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"84.2","x":"1.5cm","y":"4.8cm","width":"9.78cm","height":"2.8cm","font":"Georgia","size":"60","bold":"true","color":"FFFFFF","align":"center"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"USD millions · ARR","x":"1.5cm","y":"8cm","width":"9.78cm","height":"0.8cm","font":"Calibri","size":"14","color":"CADCFC","align":"center"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"+24% YoY","x":"1.5cm","y":"9cm","width":"9.78cm","height":"0.8cm","font":"Calibri","size":"14","bold":"true","color":"CADCFC","align":"center"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"preset":"roundRect","fill":"B85042","line":"none","x":"22.58cm","y":"4cm","width":"9.78cm","height":"7cm"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"\$1.42","x":"22.58cm","y":"4.8cm","width":"9.78cm","height":"2.8cm","font":"Georgia","size":"60","bold":"true","color":"FFFFFF","align":"center"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"CAC payback (yrs)","x":"22.58cm","y":"8cm","width":"9.78cm","height":"0.8cm","font":"Calibri","size":"14","color":"FFFFFF","align":"center"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"+8% — watch","x":"22.58cm","y":"9cm","width":"9.78cm","height":"0.8cm","font":"Calibri","size":"14","bold":"true","color":"FFFFFF","align":"center"}}
]
EOF
```

#### (f) 意思決定ツリー — YES/NO の分岐

**ビジュアルな結果。** 上部中央に菱形；YES/NO の子ボックスが左右に分岐；両方が共通の終端ボックスに合流。レイアウト: 菱形は `x=13.94, y=2cm, 6×3cm`；YES は `3cm, 7.5cm`；NO は `22.87cm, 7.5cm`；終端は `13.94cm, 13cm`。慣習: 赤 = 停止/エスカレーション、青 = 標準、緑 = 安全な終端。**すべてのコネクタに矢印が必要です** — ないと読み手が方向を誤読します。

```bash
cat <<EOF | officecli batch "$FILE"
[
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"name":"Decide","preset":"diamond","fill":"1E2761","line":"none","x":"13.94cm","y":"2cm","width":"6cm","height":"3cm","text":"Hazardous energy present?","font":"Calibri","size":"14","bold":"true","color":"FFFFFF","align":"center","valign":"middle"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"name":"YesBox","preset":"roundRect","fill":"B85042","line":"none","x":"3cm","y":"7.5cm","width":"8cm","height":"3cm","text":"Lockout + Tagout + Verify","font":"Calibri","size":"16","bold":"true","color":"FFFFFF","align":"center","valign":"middle"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"name":"NoBox","preset":"roundRect","fill":"CADCFC","line":"none","x":"22.87cm","y":"7.5cm","width":"8cm","height":"3cm","text":"Proceed with standard PPE","font":"Calibri","size":"16","bold":"true","color":"1E2761","align":"center","valign":"middle"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"name":"Done","preset":"roundRect","fill":"2C5F2D","line":"none","x":"13.94cm","y":"13cm","width":"6cm","height":"2.5cm","text":"Begin service","font":"Calibri","size":"16","bold":"true","color":"FFFFFF","align":"center","valign":"middle"}}
]
EOF
```

続けて (c) のコネクタループパターンを使って 4 本のコネクタ（`Decide→YesBox`、`Decide→NoBox`、`YesBox→Done`、`NoBox→Done`）を追加。

## QA（必須）

**問題があると仮定すること。** 最初のレンダリングが正しいことはほぼありません。ゼロ件しか見つからなかったなら、見方が足りていません。

### 納品ゲート（いずれかが失敗 = REJECT、納品しないこと）

Gate 1〜2b はテキスト/スキーマレベル（レンダリングされたスライドは見えません）；Gate 3 のみが唯一のビジュアルチェックです。完了 = すべてのゲートが PASS **かつ** Gate 3 のループが収束していること。

各ゲートは「コマンドを実行し、その出力を判断する」形式です — officecli のコマンドはすべての OS（macOS / Linux / Windows）で同一なのでシェルスクリプトは不要です；判断はあなたが行います。

- **Gate 1 — スキーマ。** `officecli validate "<file>"`。スキーマエラーが 1 つでもあれば → REJECT して修正。
- **Gate 2 — オーバーフロー / フォーマット / 構造。** `officecli view "<file>" issues`。1 件でも問題（`[O1]`、`[C1]`、`[S1]` などのタグが付いた行）を列挙したら → REJECT、修正、クリーンになるまで再実行。
- **Gate 2b — 残存プレースホルダー。** `officecli view "<file>" text` を実行し、出力を `xxxx`、`lorem` / `ipsum`、`<TODO>`、`placeholder`、"this slide layout"、または空の `()` / `[]` についてスキャンしてください。1 件でもヒットしたら → REJECT。

### Gate 3 — ビジュアル監査（必須）

**1 つ**のパスを選んでください:

**スクリーンショット（デフォルト）** — ビジョン対応エージェント向け。各スライドを順にスクリーンショットしてください — `officecli view "<file>" screenshot --page 1 -o slide1.png`、続いて `--page 2`、… — ページ番号がデックを超えるまで（スクリーンショット 1 枚 = スライド 1 枚）。ページ 1 でエラーになる場合は下記のフォールバックを使ってください。

**すべての PNG をチェックリストに照らして敵対的に判定してください** — 「問題は存在すると仮定する；何も見つからないなら見方が足りていない」。問題ごとに `slide N: <issue>` の行を 1 つ、あるいは `PASS` を報告してください。この手順はどのように実行しても必須です。**もし**あなたのハーネスがサブエージェントを起動できるなら、判定は*独立した別の*サブエージェントに委ねてください — デックを構築したエージェントは「大丈夫そう」に偏りがちなので、別の目の方が批判的です — スクリーンショットとこのチェックリスト、同じ敵対的なフレーミングを渡してください。サブエージェントがない場合は、まったく同じことを自分で行ってください。

**フォールバック — HTML テキスト**（ビジョンがない、またはスクリーンショットが失敗した場合）: `view "$FILE" html` をテキストとして読んでください。DOM は**濃色 on 濃色 / 微細な重なり / 矢印 / ギャップ・マージンの数値 / カラムの整列**を証明できません — これらは PASS ではなく「ビジュアル未検証」としてフラグしてください。

**任意の `--grid N`** — レイアウトのリズムをユーザーが要求した場合、または `view outline` が異常なレイアウト分布を示す場合のみ: `officecli view "<file>" screenshot --grid 3 -o grid.png`。

**スライド単位のチェックリスト（問題があると仮定する）:**

- **重なり** — シェイプ / チャート / 巨大な装飾番号（01/02/03、100pt 以上）の衝突
- **テキストのオーバーフロー** — スライドまたはシェイプの境界でクリップされている（KPI カード、狭いボックス）
- **狭いテキストボックス** — コンテンツは技術的には収まるが、1〜2 語ずつの短い行に多く折り返されている；3cm の KPI カード内の長いサブラベル、きつすぎるカラム内の本文行
- **濃色 on 濃色** — 明度 30% 未満の塗りに、明度 80% 未満のテキスト/アイコン（コントラストのある円のない濃色アイコンを含む）
- **画像の扱い** — 引き伸ばされ/歪んだ写真、ごちゃついた画像の上に直乗せされたテキスト（カード/スクリムなし）、クロップされたスクリーンショットやロゴ、白地の上に浮いた透過画像
- **矢印の欠落** — 単なる線として描かれたフローチャートのコネクタ
- **装飾線 / タイトルの不一致** — 1 行タイトル用のアクセントバーなのにタイトルが 2 行に折り返された（またはその逆）
- **フッター / 出典の衝突** — 出典行、ページ番号、脚注が上のコンテンツに接触している
- **タイトなマージン / ギャップ** — スライド端から ~0.5" 以内の要素、または 2 枚のカードが ~0.3" 以内
- **不均一なギャップ** — 片側は大きな空きスペース、もう片側は詰まっている（リズムの崩れ）
- **カラム / 反復要素のずれ** — KPI カード/アイコンがベースラインからずれている、または幅が一貫していない
- **順序のサニティ** — 順序が物語と一致している（カバー → アジェンダ → セクション前の区切り → クロージング）

REJECT する場合は `slide N: <issue>` の行、そうでなければ "Gate 3 PASS"（HTML テキストのフォールバックの場合は「<unverified-items> は視覚的に未検証」を追加）。

**修正-検証（必須、最大 3 サイクル）。** 修正 → Gate 3 を再実行 → 新しい問題がゼロになるまで繰り返す；1 つの修正が別の問題を露呈することがよくあります。3 ラウンド経っても収束しない場合は**停止**してください — シーソー現象、テンプレートレベルの原因、またはエージェントの誤読の可能性が高いです。`slide N: <issue> — attempted: <fixes> — likely root: <template|design-conflict|ambiguous>` を報告し、ユーザーに判断を委ねてください。

**そしてフラッシュする（ゲートの一部）。** Gate 3 が収束したら、`officecli save "<file>"` で締めくくってください — これにより納品前に編集がディスクに書き込まれることが保証されます（ワンショットのハンドオフでレジデントも解放したい場合は代わりに `officecli close "<file>"` を使ってください）。これは省略可能な手順ではなく必須の最終手順です。常に安全です: エラーになったり作業が失われたりすることはありません。

## よくある落とし穴

サニティチェック用のチートシート — 初回で何がつまずくか。デザインとシェルの罠。

| 落とし穴 | 正しいアプローチ |
|---|---|
| zsh/bash でクォートされていない `[N]` | パスは常にクォート: `"/slide[1]"`。zsh はクォートされていない `[1]` をグロブして `no matches found` になる — 初回利用の #1 のつまずき |
| `--name "foo"` | すべての属性は `--prop` を通す: `--prop name="foo"` |
| `/shape[myname]`（角括弧に生の名前） | `@name=` セレクタを使う: `/shape[@name=myname]` または `/shape[@id=10007]` |
| パスは 1-based、`--index` は 0-based | `/slide[1]` = 最初のスライド；`--index 0` = 最初の位置 |
| `--prop text=` 内の `$` | シングルクォート: `--prop text='$15M'`。ダブルクォートの `"$15M"` はシェルにより `M` に展開される |
| `--prop text=` 内の `\n` / `\t` | CLI により解釈される: `\n` = 段落区切り、`\t` = タブ。リテラルには `\\n` を二重にする |
