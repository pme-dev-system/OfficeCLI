---
name: officecli-pitch-deck
description: "このスキルは、ユーザーが資金調達 / 投資家向けピッチデッキを作成する際に使用する — シード、Series A / B / C、コンバーティブルノート、SAFE ラウンド、戦略的資金調達など。トリガーとなるキーワード: 'pitch deck'、'investor deck'、'Series A deck'、'Series B deck'、'Series C deck'、'fundraising deck'、'seed pitch'、'VC deck'、'raising capital'、'term sheet presentation'。出力は単一の .pptx。このスキルは officecli-pptx の上に乗るシーンレイヤーであり — pptx v2 のあらゆるルール（visual floor、グリッド、パレット、コネクタ規範、Delivery Gate）を継承する。一般的な取締役会レビュー、セールスデック、全社会議、製品ローンチには使用しないこと — それらは officecli-pptx ベースへルーティングする。"
---

# OfficeCLI Pitch Deck スキル

**このスキルは `officecli-pptx` の上に乗るシーンレイヤーです。** pptx のハードルール — visual delivery floor（title ≥ 36pt / body ≥ 18pt / title ≥ 2× body）、33.87×19.05cm 上の 12 カラムグリッド、4 つの正規パレット、チャート選択の判断表、コネクタ規範（`shape` / `from` / `to` / `tailEnd=triangle`）、シェルエスケープ、resident + batch、Delivery Gate 1–5a — はすべて継承されるものであり、ここで再度教えるものではない。このファイルが追加するのは、**資金調達**特有に必要な部分だけ: ステージ診断（A / B / C）、5 つの業界別アーク・テンプレート、10 個のキースライド・レシピ（cover / problem / solution / market / product / model / traction / team / financials / ask）、ピッチ特有の数値表記規約、VC 向け出荷前チェック、そしてピッチ特有の fresh-eyes Gate 6 である。

pptx ベースのルールでカバーされている箇所は、本文中で `→ see pptx v2 §X` と表記する。まだ読んでいない場合は、先に `skills/officecli-pptx/SKILL.md` を読むこと。

## セットアップ

`officecli` が未インストールの場合:

- **macOS / Linux**: `curl -fsSL https://d.officecli.ai/install.sh | bash`
- **Windows (PowerShell)**: `irm https://d.officecli.ai/install.ps1 | iex`

`officecli --version` で確認する（PATH が反映されない場合は新しいターミナルを開くこと）。インストールに失敗する場合は https://github.com/iOfficeAI/OfficeCLI/releases からバイナリをダウンロードする。

## ⚠️ ヘルプ優先ルール

**このスキルは資金調達デックに何が必要かを教えるものであり、すべてのコマンドフラグを網羅するものではない。** プロパティ名、enum 値、プリセットが不確かな場合は、推測する前にヘルプを確認すること。

```bash
officecli help pptx                          # すべての pptx 要素
officecli help pptx <element>                # 完全なスキーマ（例: chart, shape, connector, picture）
officecli help pptx <element> --json         # 機械可読形式
```

ヘルプはインストール済みの CLI バージョンを反映する。このスキルとヘルプが食い違う場合は **ヘルプが優先する。** 本ファイル内のすべての `--prop X=` は `officecli help pptx <element>` に対して grep 検証済みだが、後続バージョンでヘルプがプロパティを追加・改名している場合はヘルプを信頼すること。

## メンタルモデルと継承

**pptx v2 を継承する。** 先に `skills/officecli-pptx/SKILL.md` を読んでいることが前提。このスキルは、スライド・シェイプ・チャート・コネクタの追加、`@name=` / `@id=` によるアドレッシング、パスのクォート、`batch` heredoc の使用、すべてのフローコネクタへの `--prop tailEnd=triangle` の記述、5 段階の Delivery Gate の実行方法を知っていることを前提とする。いずれかに不慣れな場合は、続行する前に pptx v2 セッションを開くこと。

## シェル & 実行の規律

**シェルクォート、段階的実行、`$FILE` 規約** → pptx v2 §Shell & Execution Discipline を参照。ルールはそのまま同じ — `[N]` を含むパスはクォートし、`$` を含む値（カバーや ask スライドの `$35M`、`$1.2B TAM` を含む）はシングルクォートし、実行可能な例では `\$ \t \n` を手書きしない、コマンドは一度に一つずつ実行する。以下の例は `$FILE`（`FILE="deck.pptx"`）を使用する。

**`$` を含むすべてのシェイプテキストはシングルクォートすること。** `--prop text="Series B · $35M"`（ダブルクォート）は誤り — zsh が `$35M` を展開して空文字になり、デックには `Series B · M` と何の警告もなく描画されてしまう。`--prop text='Series B · $35M'`（シングルクォート）が正しい。これはピッチデックのシェルエスケープ失敗モードの第 1 位である（`$35M`、`$18M ARR`、`$1.2B TAM` はカバー / ask / financials / milestones に登場する）。Gate 2 は取り除かれた `$35M` を検出できない — 痕跡が残らないためだ。Gate 2b は一般的な strip パターンを捕捉する。シングルクォートはそれらを未然に防ぐ。

## ここでの「ピッチデック」の定義（アイデンティティ）

ピッチデックとは、**資金調達レイヤー**を重ねた pptx である: VC 志向のナラティブアーク、検証可能な指標、ステージに応じたデータ密度、創業者としての信頼性の演出。スライドはライブの場で 1 枚あたり約 3 秒で消費される — これは pptx v2 のルールだ。ピッチデックはその上にもう一つの制約を加える: **すべてのスライドは 1 つの投資可能な命題を運ぶ**。あるスライドが「興味深い背景情報」に過ぎず、ask を前に進めないなら、切り捨てること。VC はそうする。ベースの pptx ルールは引き続き適用され、ピッチデックはさらに 6 つの差分を加える:

1. **ステージがすべてを決定する。** Series A / B / C はそれぞれ、スライド枚数、ナラティブの重み、必須指標、ユニットエコノミクスの洗練度に対する許容度を決める。CAC/LTV の数式が 6 ページも続く Series A デックは過剰包装に見え、ユニットエコノミクスを欠く Series B デックは未完成に見える。まずステージを決めること — その後のすべてはそこから導かれる。
2. **ナラティブアークはフィーチャーの羅列に勝る。** 10 の必須スライドを固定順序で: cover → problem → solution → market → product → model → traction → team → financials → ask。順序が狂えば VC は離脱する。
3. **数値は契約である。** TAM/SAM/SOM はクリーンな三層構造でなければならず、CAC/LTV には payback の行が必要で、ARR ≠ revenue、Use-of-Funds は四分割の円グラフでなければならない。数値が雑ならラウンドは死ぬ。
4. **チームスライドは過去の所属企業を運ぶ。** アバターグリッドだけでは学生プロジェクトに見える。過去の所属企業のロゴ / 名前 + 1 行の役割を加えること。これがないと、初めての創業者はまさに初めての創業者そのものに見えてしまう。
5. **トラクションチャートの y 軸は 0 から始める。** 現在値の 80% を `y_min` とする「ホッケースティック」は視覚的な嘘であり — 1 万件のデックを見てきた VC は 2 秒未満で見抜く。
6. **ask はスライドであり、脚注ではない。** `$XX M` の主役数値 + 四分割の Use-of-Funds + ランウェイ期間。「多少の資金を調達しています」は ask ではない。

### 逆ハンドオフ — pptx ベースに戻るべきタイミング

取締役会レビュー、全社会議、セールスデック、製品ローンチ、トレーニングデックなど — 資金調達に紐づかないものは **pptx v2 ベース**にとどまること。**このスキル**を使うのは、(a) ユーザーが特定のラウンド（シード / Series A / B / C）や VC ミーティングに言及しており、かつ (b) デックが {problem, traction, team with credentials, Use-of-Funds, stage-appropriate unit econ, financial projections} のうち少なくとも 4 つを必要とする場合のみ。

ユーザーが「資金調達デック」と言っていても、文脈が事業部門の四半期予算要求であれば、それは取締役会レビューである。pptx v2 Recipe (d) の 10 枚ブループリントへルーティングすること。逆にユーザーが「取締役会レビュー」と言っていても、文脈が小規模企業のブリッジラウンド調達であれば、こちらへルーティングすること。

## Series A / B / C ステージ診断（判断ツール）

**コマンドを 1 つでも書く前にこれを読むこと。** ユーザーの説明に合致する行を選ぶ — それ以降のすべて（スライド枚数、必要な指標、使用するレシピ、チームスライドが何を示すべきか）はこの一回の判断から導出される。

| ステージ | 売上帯 | チーム | スライド枚数 | 支配的ナラティブ（比重） | 必須データ | よくあるレッドフラグ |
|---|---|---|---|---|---|---|
| **Seed** | $0 – $1M ARR（多くはプレレベニュー） | 2 – 8 名 | 10 – 12 | Problem (30%) + Solution (25%) + Team (15%) + Market (15%) + Traction (15%) | Founder-market fit のストーリー; デザインパートナー / パイロットのロゴ 1 – 2 件; トップダウン TAM で可 | トラクションの過大主張（顧客 10 社 = 「市場実証済み」） |
| **Series A** | $1 – $5M ARR | 10 – 25 名 | 12 – 16 | Problem (20%) + Solution (20%) + **Market「なぜ今か」** (15%) + Product (15%) + Traction (20%) + Team (10%) | PMF の証拠（NRR > 110%、低チャーン）、ボトムアップ TAM/SAM、パイプライン / パイロットの成約実績 | ボトムアップ TAM が作り話に見える; CAC/LTV がまだ意味を持たないのに提示されている |
| **Series B** | $5 – $30M ARR | 30 – 100 名 | 18 – 22 | **Traction + Unit econ (30%)** + Market + Product + Team + Financials (ask) | 0 起点の ARR カーブ; NRR、CAC、LTV、payback（18 か月未満が理想）; コホートリテンション; ロゴウォール | ユニットエコノミクススライドがない; 説明なしに CAC payback が 24 か月超; Use-of-Funds に割合の記載がない |
| **Series C** | $30M+ ARR | 100 名以上 | 20 – 24 | **Financials + Scale + Moat (40%)** + Market expansion + Team depth | 複数年 GAAP、Rule of 40、GM の推移、国際展開計画、防御可能性 | moat スライドがない; マージンのストーリーを伴わない売上成長; チームスライドに CEO / CFO の前職経験がない |
| **Bridge / SAFE** | 任意 | 任意 | 8 – 10 | **具体的なブリッジの理由** + ランウェイ計算 + コミットメント | 前ラウンドの文脈; ブリッジ資金が達成するマイルストーン; コミット済み投資家額 | Series A のように扱う — スライドが多すぎて ask が薄まる |

**判断手順。** ユーザーの 1 ～ 2 文（「Series B、$18M ARR、顧客 120 社、$35M 調達」）から、ステージ行を 1 つだけ選ぶ。以降のこのスキル内の選択はすべて、あなたのステージ判断を参照する: どの業界別テンプレートを使うか、どのレシピが必須 / 任意か、Delivery Gate 6 のどのチェックが発火するか。

**コーナーケース。** A → B 間のブリッジラウンドやコンバーティブルは、ブリッジが「PMF を完成させる」（A 形状）ためのものか「ユニットエコノミクスの目標を達成する」（B 形状）ためのものかによって、A か B のどちらかに近い。同一ステージでの「エクステンション」ラウンドは、前ステージの骨格を再利用しつつ「前回ラウンド以降の進捗」1 枚を追加する。

**非 SaaS のステージ上書き。** Series B の ARR / ユニットエコノミクスの形は SaaS に合わせたものだ。他の業界では、売上帯 + ユニットエコノミクス相当指標 + Gate 6.3 の grep を置き換える:

| 業界 | Series B 時点の売上「帯」 | 「ユニットエコノミクス」相当 | Gate 6.3 の代替 |
|---|---|---|---|
| **バイオ / 臨床段階** | プレレベニュー、20–60 名 | バーンレート + 次のマイルストーン（IND / Ph1 readout / BLA）までのランウェイ | `shape:contains("ORR")` OR `contains("Pipeline")` OR `contains("BLA")` OR `contains("runway")` ≥ 1 |
| **ディープテック / フロンティア** | プレレベニューまたは初期パイロット売上 | 技術的マイルストーン + TRL レベル + SoTA 対比ベンチマーク | `shape:contains("TRL")` OR `contains("benchmark")` ≥ 1 |
| **マーケットプレイス / ネットワーク** | GMV $10–100M | テイクレート + コホートリテンション + 流動性 | `shape:contains("GMV")` + `contains("take rate")` ≥ 1 |
| **コンシューマーハードウェア** | $2–15M 売上（出荷台数ベース） | 貢献利益 + リピート率 + ブレンド CAC | `shape:contains("repeat")` OR `contains("contribution")` ≥ 1 |

これらの業界で Gate 6.3 を実行する際は類似 grep に置き換えること。SaaS の CAC/LTV に対する誤検知 WARN は想定内; 真の懸念は業界特有の代替指標が存在するかどうかである。特にバイオの Series B デックでは、バーン + マイルストーンまでのランウェイこそが「ユニットエコノミクス」のストーリーになる。

## 業界別アーク・テンプレート（5 ファミリー）

主要な 5 業界。VC が概念実証として求める証明が業界ごとに異なるため、スライドの比重も異なる。業界の行を選ぶこと。スライド骨格はそのままコピーして使える出発点である。スライド枚数は上記の対応するステージ行を前提とする。

### (1) B2B SaaS / エンタープライズソフトウェア

正規のアーク — VC の筋肉記憶の大半がこのテンプレートに基づいて形成されている。Series B の例（20 枚）: cover · TL;DR · problem · problem evidence · solution · product loop · market TAM/SAM/SOM · **ユニットエコノミクス（CAC / LTV / payback / GM）** · ARR trajectory · retention cohort · logo wall · team · competitors · financials 4-year · ask。必須: Series A 以降はユニットエコノミクススライド; Series B 以降はロゴウォール。

### (2) コンシューマー（B2C アプリ / コンシューマーハードウェア / D2C）

ナラティブ駆動。初期段階のデックは **プロダクト体験のスクリーンショット + 創業ストーリー + 「なぜ今か」**の市場タイミングに寄りかかる; ユニットエコノミクス（通常 SaaS より弱い）の比重は軽い。Series A の例（14 枚）: cover · hook（30 秒のプロダクトデモまたは 1 行のビジョン）· problem（実体験）· solution（プロダクト画像）· product-experience flow · 「なぜ今か」の市場ウィンドウ · 予約注文 / クラウドファンディング / 初期販売の実績 · retention / engagement（DAU、D30）· market（ボトムアップが信頼できない場合はトップダウンで可）· competitive positioning · founder story + team · press / endorsements · financials · ask。必須: プロダクトのビジュアルが 3 枚以上; 「なぜ今か」スライド（ウィンドウの正当化）; 売上だけでなくエンゲージメント指標。

### (3) ディープテック / フロンティアテック（AI 基盤モデル、量子、気候ハードウェア、ロボティクス）

技術的信頼性こそが売りである。プレレベニューのディープテックは「トラクション」の代わりに **技術的マイルストーン + 防御可能性**を用いる。Series B の例（22 枚）: cover · thesis（「これが実現したら何が変わるか」を 1 行で）· problem（現在の技術水準）· solution（技術的アプローチ）· **技術アーキテクチャ** · SoTA 対比ベンチマーク · pipeline / TRL レベル · market（ロングテール）· business model · early commercial traction（パイロット、LOI）· IP / 特許 · team（通常は博士号 / 元 FAANG リサーチ）· partners · financials · ask。必須: ベンチマークスライド; IP スライド; 博士号や前所属研究室名を密に記載したチームスライド。

### (4) マーケットプレイス / ネットワークビジネス（双方向プラットフォーム、ソーシャル、コマース）

流動性こそが指標である。「ユニットエコノミクス」の代わりに **GMV + テイクレート + コホートリテンション + 需給バランス**を用いる。Series A の例（15 枚）: cover · problem（現行の需給における摩擦）· solution · product demo（両サイド）· network effects diagram · early liquidity（初週の GMV、マッチングまでの時間）· cohort retention · geographic / category expansion plan · competitive positioning vs incumbents · take-rate model · team · financials · ask。必須: 流動性指標スライド; コホートリテンションチャート; ネットワーク効果図。

### (5) バイオ / ライフサイエンス / ヘルステック

規制上のパイプラインそのものが事業である。「プロダクトロードマップ」の代わりに **臨床パイプライン + 規制上の道筋 + 科学的エビデンス**を用いる。Series B の例（22 枚）: cover · unmet medical need · scientific rationale（作用機序）· preclinical / clinical data（ORR、安全性、エンドポイント）· **パイプラインチャート**（候補薬 × ステージ × 日付）· differentiation vs standard of care · IP / exclusivity · regulatory strategy（IND、BTD、fast-track）· market（有病率 × 価格設定）· commercial strategy（orphan / specialty / biosimilar）· partnerships / collaborations · team（過去に FDA 承認実績のある CSO / CMO）· financials（次のマイルストーンまでのバーン）· ask。必須: パイプラインチャート; 臨床データスライド; 過去の規制当局対応実績を持つチームスライド。

**業界横断ルール。** テンプレート間で要素を混在させてもよいが、主要業界の必須項目を落としてはならない。ユニットエコノミクスを欠く SaaS デック、パイプラインチャートを欠くバイオデック、流動性指標を欠くマーケットプレイスデック — いずれも即座に VC から失格とされる。

## スライドパターン（レイアウト規範）

パターンは**レイアウトの幾何構造**であり、以下のレシピは**ナラティブの意図**である。1 枚のスライドは、その視覚的な形として 1 つのパターン（下記 6 種の正規パターン）を選び、何を主張するかとして 1 つのレシピ（cover / problem / traction / ...）を選ぶ。複数のレシピが 1 つのパターンを共有できる — Problem / Why-Now / Traction-callout はいずれも 3-stat row（C.2）に依拠している。まずパターンを選び、次にレシピの内容で埋めること。

**スピーカーノートのルール。** すべてのコンテンツスライド（cover でも closing でもないもの）は、`officecli add "$FILE" /slide[N] --type notes --prop text='…'` によってスピーカーノートを必ず持たなければならない。ノートの欠落は出荷不可 — pptx v2 §Hard rules (H7) を継承する。プロパティ名を確認するには先に `officecli help pptx notes` を実行すること。

**パターン再利用の規律。** 連続する 2 枚のスライドで同じパターンを使わないこと — データが異なっていても、同じ幾何構造が続くとテンプレートのループに見える。C.2 と C.4 または C.5b を交互に使ってリズムを崩すこと。

**垂直方向の中央寄せ。** スライドに含まれる要素がパターンの最大数より少ない場合、y 座標を 2 – 3cm 下げて視覚的な重心を中央に寄せること。以下の表はフルコンテンツを前提とする。

### C.1 タイトル / カバー（ダークグラデーション）

グラデーション塗りの上に 3 – 4 個のテキストシェイプ。すべてのデックの 1 枚目。

```
+----------------------------------+
|                                  |
|          TITLE (centered)        |
|          tagline                 |
|                                  |
|   round · amount · date          |
|  ________________________        |  <- thin brand band
+----------------------------------+
```

| 要素 | X | Y | 幅 | 高さ | フォント / サイズ |
|---|---|---|---|---|---|
| Title | 2cm | 5cm | 29.87cm | 4cm | セリフ体太字、36pt 以上（44 が典型） |
| Tagline | 2cm | 10cm | 29.87cm | 2cm | サンセリフ 18–22pt |
| Meta (round · $ · date) | 2cm | 13cm | 29.87cm | 1.5cm | サンセリフ 12–16pt |

**使用場面。** そのスライドが 1 枚目（Cover レシピ 1）の場合 — 3 秒でのアイデンティティ把握。背景は 2 つのダークパレット色の間の 180° 線形グラデーション（例: Professional Navy `1E2761 → 0D1F35`）。タイトルが 2 行に折り返す場合は **高さを増やし（4cm → 5cm）、フォントを 36pt 未満に下げないこと** — ピッチのカバーで 36pt 未満は、内容に関わらず弱気に見える。トランジション: フェード。

### C.2 3-Stat コールアウト行

Title + 3 個の大きな数字 / ラベルのペアを横並びに。Problem / Why-Now / Traction-callout スライドの既定パターン。

```
+----------------------------------+
|  Title                           |
|                                  |
|   73%      12hr      $4.2B       |
|   label    label     label       |
|   source   source    source      |
+----------------------------------+
```

| 要素 | X | Y | 幅 | 高さ | フォント / サイズ |
|---|---|---|---|---|---|
| Title | 1.5cm | 1cm | 30.87cm | 3cm | セリフ体太字 36pt 以上 |
| Stat 1 number | 2cm | 5cm | 9cm | 4cm | セリフ体太字 60–64pt |
| Stat 1 label | 2cm | 9.5cm | 9cm | 2cm | サンセリフ 16pt 以上（H4 floor） |
| Stat 2 number / label | 12.5cm | (同上) | 9cm | (同上) | (同上) |
| Stat 3 number / label | 23cm | (同上) | 9cm | (同上) | (同上) |

**使用場面。** アンカーとなる数字が 2 – 3 個あり、「3 つの事実が論点を主張する」ストーリーの場合 — Problem、Why-Now、Market-callout、単一行の Traction。ラベル 16pt 以上は H4 floor（サブラベルの例外）; ラベルのない数字ははったりに見えるので、より多くのテキストを収めるためにラベルを 12–14pt に落としてはならない。

### C.3 4-Stat コールアウト行

C.2 と同じ幾何構造だが 4 列。数字は 60pt、幅は各 7cm。

```
+-------------------------------------+
|  Title                              |
|                                     |
|  73%   12hr   $9M   4.2x            |
|  lbl   lbl    lbl   lbl             |
+-------------------------------------+
```

| 要素 | X 位置 | Y | 幅 | 高さ | フォント / サイズ |
|---|---|---|---|---|---|
| Title | 1.5cm | 1cm | 30.87cm | 3cm | セリフ体太字 36pt |
| Stat numbers | 1.5 / 9.5 / 17.5 / 25.5cm | 5cm | 7cm | 4cm | セリフ体太字 60pt |
| Stat labels | (同 X) | 9.5cm | 7cm | 2cm | サンセリフ 16pt 以上 |

**使用場面。** ちょうど 4 つの並列指標がストーリーを語り、3 では説明不足に感じる場合。迷ったら C.2 を優先すること — 4 は常に 3 より窮屈に感じられ、折り返しリスクも現実にある。

> **折り返し警告。** 幅 7cm で 60pt の場合、`$` と `.` の両方を含むドル表記は失敗する: `$9.4M` は 5 文字だが、セリフ体太字での幅広い `$` と `.` により 2 行に折り返り、コールアウトが崩れる。60pt/7cm で安全なドル表記の形: `$9M`、`$96B`、`$4K`（3 – 4 文字）。ドル以外の形: `340%`、`4.2x`、`12.3` は 5 文字まで安全。6 文字以上の値（`197min`、`3 Days`）は折り返る — (a) フォントを 44–48pt に落とす、(b) 略記する（`197m`、`$9M`）、(c) C.2（1 stat あたり 9cm）に切り替える、のいずれかを行うこと。単一トークンのみ、内部にスペースを含めないこと。

### C.4 チャート + コンテキスト（左にチャート、右に stats）

チャートが左 55% を占め、右側に 2 – 3 個のスタック型コールアウト。Traction / Financials / コンテキスト付き Market-sizing の既定パターン。

```
+-------------------------------------+
|  Title                              |
|                                     |
|  +---------------+   +--------+     |
|  |               |   | Stat 1 |     |
|  |    chart      |   +--------+     |
|  |               |   | Stat 2 |     |
|  +---------------+   +--------+     |
+-------------------------------------+
```

| 要素 | X | Y | 幅 | 高さ |
|---|---|---|---|---|
| Title | 2cm | 1cm | 29.87cm | 3cm |
| Chart | 2cm | 4cm | 17cm | 13cm |
| Stats column | 21cm | 4cm+ | 11cm | 数字 2.5cm + ラベル 1.5cm（ペアあたり約 3.7cm） |

サブラベルは 16pt 以上（H4 floor）。stats が 5 個スタックされる場合は数字サイズを 44pt に落とす; 6 個以上なら別のパターンを選ぶこと。カラム / バーチャートのバッチ後処理: `officecli set "$FILE" "/slide[N]/chart[1]" --prop gap=80` でバー間隔を詰める。

**使用場面。** 1 つの主要チャートがストーリーを牽引し、2 – 3 個の数値アンカーがそれを補強する場合 — Traction（ARR カーブ + 現在の ARR + YoY + NRR）、Financials（4 年分のカラムチャート + 前提のコールアウト）、Market（バーチャート + SOM / CAGR / 手法）。

### C.5 円アイコングリッド（3 行縦並び）

3 行の縦並び、それぞれ = 左に円形アイコン + タイトル + 1 行の説明。

```
+---------------------------------------+
|  Title                                |
|                                       |
|  (o)  Label one                       |
|       description one                 |
|                                       |
|  (o)  Label two                       |
|       description two                 |
|                                       |
|  (o)  Label three                     |
|       description three               |
+---------------------------------------+
```

| 要素 | X | Y 位置 | 幅 | 高さ | フォント / サイズ |
|---|---|---|---|---|---|
| Icon circle | 2cm | 4.5 / 8.5 / 12.5cm | 2.5cm | 2.5cm | 楕円、アクセント塗り |
| Label | 5.5cm | (icon Y + 0) | 25cm | 1.2cm | サンセリフ太字 18pt |
| Description | 5.5cm | (icon Y + 1.3cm) | 25cm | 1.8cm | サンセリフ 16pt 以上（H4 floor）、ミュートトーン |

**使用場面。** 短い縦の要点が 3 つあり、行ごとに視覚的アンカーがあると効果的な場合 — Solution mechanism、Value pillars、Product loop。項目が並列でちょうど 4 個ある場合は C.5b（2×2 グリッド）を選ぶこと; アイコンを横並びに見せたい場合（例: 5 ステップのプロセス）は横型 5 列バリアントを選ぶこと。

### C.5b 2×2 フィーチャーグリッド（4 並列項目）

角丸カード 4 枚、2 列 × 2 行。ちょうど 4 個の並列項目（プロダクトの柱、サービス種別、フィーチャーの四象限）がある場合に使用する。

```
+-----------------------------+
|  Title                      |
|                             |
|  +---------+  +---------+   |
|  | (o) T1  |  | (o) T2  |   |
|  | body    |  | body    |   |
|  +---------+  +---------+   |
|  +---------+  +---------+   |
|  | (o) T3  |  | (o) T4  |   |
|  | body    |  | body    |   |
|  +---------+  +---------+   |
+-----------------------------+
```

| 要素 | X | Y | 幅 | 高さ | フォント / サイズ |
|---|---|---|---|---|---|
| Slide title | 2cm | 1cm | 29.87cm | 2.5cm | セリフ体太字 32pt |
| Card 1 bg (top-left) | 1.5cm | 4cm | 14.5cm | 7cm | roundRect |
| Card 2 bg (top-right) | 17.5cm | 4cm | 14.5cm | 7cm | roundRect |
| Card 3 bg (bottom-left) | 1.5cm | 12cm | 14.5cm | 7cm | roundRect |
| Card 4 bg (bottom-right) | 17.5cm | 12cm | 14.5cm | 7cm | roundRect |
| Icon ellipse (each card) | card_x + 0.5cm | card_y + 0.5cm | 2cm | 2cm | — |
| Card title (each) | card_x + 3.2cm | card_y + 0.6cm | 10.5cm | 1.8cm | サンセリフ太字 16pt |
| Card body (each) | card_x + 0.5cm | card_y + 3cm | 13cm | 3.5cm | サンセリフ 16pt 以上（H4 floor） |

**使用場面。** ちょうど 4 個の並列項目があり、視線が均等に各要素へ向くべき場合 — 4 つのプロダクトの柱、4 つのサービス層、4 種のステークホルダー。3 項目では 2×2 の中で寂しく見え、5 項目以上ではグリッドが崩れる — 3×2 に移行する（pptx v2 §(d) グリッド数学を参照）か、C.5 の行パターンを使うこと。

> **Z-order 規範（重要）。** 各カードの `roundRect` 背景は、そのカードのアイコン / タイトル / 本文シェイプの直前にバッチ JSON 内で追加しなければならない — pptx は挿入順に描画するため、テキストの後に追加された背景はテキストを覆い隠してしまう。`officecli batch` で構築する際は、カード単位の順序 `bg → ellipse → title → body` を厳守すること。パターンと z-order の詳細 → pptx v2 §Recipe (c) z-order canon を参照; 2×2 以外の個数の場合は pptx v2 §(d) からグリッド数学を再利用すること。

**ダーク背景バリアント。** カード塗りを `F0F4F8`（明色）から `1A2540` のような明るめのダーク色に変更し、本文テキストを `FFFFFF` / `E8E8E8` に上げる。パレット変数（例: `$MUTED`）はシングルクォート heredoc の中では展開されない — JSON 内には 16 進数のリテラル（`64748B`）を直接書くこと。

---

## キースライドレシピ（10 の必須要素）

すべてのピッチデックが持つ 10 枚のスライド。以下の各レシピは: **視覚的成果**（3m 離れて見たときのスライドの見え方）+ **実行可能なブロック**（18 行以内）+ **QA ワンライナー**を示す。すべてのレシピは、pptx v2 のパレット、グリッド数学、タイプ階層、すべてのコネクタへの `--prop tailEnd=triangle` を継承する。レシピは上記のスライドパターンを参照する: Cover は C.1 を再利用; Problem / Why-Now は C.2 を再利用; Traction / Financials は C.4 を再利用; Feature / pillar スライドは C.5b を再利用。`$FILE` はあなたのデックファイル。

**長いタイトルの折り返しルール。** 36pt 以上のタイトルが 2 行に折り返る場合: `height` を増やす（例: 2cm → 3.5cm）— フォントを 36pt 未満に下げないこと。ピッチデックで 36pt 未満のタイトルは、内容に関わらず弱気に見える。

> **チャートの `series1.color=` は `add` で機能する**（以下のすべてのチャートレシピに適用される）。`--prop series1.color=`（または `series2.color=`、…）をチャート `add` に渡すと、シリーズカラーが適用され exit 0 で終了する。気になる場合は読み戻しで確認できる: `officecli get "$FILE" "/slide[N]/chart[1]/series[1]" --json | jq '.data.results[0].format.color'`。

### (1) カバースライド — 会社名 · タグライン · ラウンド · 日付

**視覚的成果。** ダークネイビーの塗り、中央揃えの 44pt 会社名、その下に 20pt の 1 行タグライン、下部にラウンド + 金額 + 日付を記した小さな 16pt のメタ行。最下部にはアクセントカラーの薄いブランドバンド（高さ 0.5cm）。

```bash
officecli add "$FILE" / --type slide --prop layout=blank --prop background=1E2761
officecli add "$FILE" "/slide[1]" --type shape --prop name=BrandBand \
  --prop geometry=rect --prop fill=CADCFC \
  --prop x=0cm --prop y=18.5cm --prop width=33.87cm --prop height=0.55cm
officecli add "$FILE" "/slide[1]" --type shape --prop name=CoverTitle --prop text="Acme DevOps" \
  --prop x=2cm --prop y=7cm --prop width=29.87cm --prop height=3cm \
  --prop font=Georgia --prop size=44 --prop bold=true --prop color=FFFFFF --prop align=center --prop fill=none
officecli add "$FILE" "/slide[1]" --type shape --prop name=Tagline --prop text="Kubernetes observability, built for production at scale" \
  --prop x=2cm --prop y=10.5cm --prop width=29.87cm --prop height=1.5cm \
  --prop font=Calibri --prop size=20 --prop color=CADCFC --prop align=center --prop fill=none
officecli add "$FILE" "/slide[1]" --type shape --prop name=CoverMeta --prop text='Series B · $35M · April 2026' \
  --prop x=2cm --prop y=15cm --prop width=29.87cm --prop height=1.2cm \
  --prop font=Calibri --prop size=16 --prop color=FFFFFF --prop align=center --prop fill=none
```

**QA。** カバーには 4 つの独立した要素がある（ブランドバンド + タイトル + タグライン + メタ）。余白 80% のカバーは pptx の「カバー充填率 60% 以上」floor に違反する。

**コンシューマー・バリアント（3 秒での把握）。** コンシューマーデック（B2C アプリ / ハードウェア / D2C）は、1 つの支配的なモチーフ — ヒーロープロダクト写真、超大型の会社名（60–96pt）、または象徴的なマーク（三日月 / 抽象幾何形状）— を加えるべきである。44pt のタイトルを、80–96pt の名前 + 1 個のモチーフシェイプ（抽象マークには `--type shape --prop geometry=ellipse --prop fill=<accent>`、プロダクトヒーローにはスライドの約 40% を占める `picture`）に置き換えること。タグライン + ラウンド + 日付はそのまま。SaaS / B2B は省略してよい — タイポグラフィのみのカバーで十分。

### (2) Problem スライド — 業界の痛みを 1 文で + 3 個のデータカード

**視覚的成果。** 痛みを述べる 36pt のタイトル（「The Problem」ではなく）。その下に、スライド全体に等幅の 3 個のデータカード: それぞれ大きな数字（40pt）+ 1 行の修飾語（16pt）+ 出典脚注（12pt グレー）。

3 カード分のグリッド数学、余白 1.5cm、ギャップ 0.76cm: `usable = 33.87 − 3 − 2·0.76 = 29.35`、`col_width = 29.35 / 3 = 9.78cm`。x 位置: `1.5 / 12.04 / 22.58`。

```bash
SLIDE=2  # second slide, after cover. Adjust from your build order.
officecli add "$FILE" / --type slide --prop layout=blank --prop background=FFFFFF
officecli add "$FILE" "/slide[$SLIDE]" --type shape --prop text="Kubernetes debugging burns 12 engineering hours / incident" \
  --prop x=1.5cm --prop y=1.2cm --prop width=30.87cm --prop height=2.5cm \
  --prop font=Georgia --prop size=36 --prop bold=true --prop color=1E2761 --prop fill=none
cat <<EOF | officecli batch "$FILE"
[
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"name":"PC1","geometry":"roundRect","fill":"F5F7FA","x":"1.5cm","y":"5cm","width":"9.78cm","height":"10cm"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"73%","x":"1.5cm","y":"6cm","width":"9.78cm","height":"3cm","font":"Georgia","size":"60","bold":"true","color":"1E2761","align":"center","fill":"none"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"of incidents take > 1 hour to diagnose","x":"1.5cm","y":"9.5cm","width":"9.78cm","height":"3cm","font":"Calibri","size":"18","color":"333333","align":"center","fill":"none"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"Source: 2025 DORA Report","x":"1.5cm","y":"13cm","width":"9.78cm","height":"1cm","font":"Calibri","size":"12","italic":"true","color":"666666","align":"center","fill":"none"}}
]
EOF
# カード 2、3 を x=12.04cm、x=22.58cm で同じ 4 ブロックのパターンを繰り返す。
```

**QA。** `officecli query "$FILE" 'shape:contains("Source")'` が 3 以上を返す（すべての主張に出典が伴う）。出典がゼロなら、VC は単一の数字を信用しない。

### (2b) Why Now スライド — コンシューマー / シード / 早期 A の必須要素

**視覚的成果。** 横並びの 3 カード: それぞれ = **トリガーの見出し**（24pt 太字）+ **データポイント**（60pt の数字または日付）+ **1 行の含意**（16pt）+ **出典脚注**（12pt グレー）。Problem のグリッド数学を再利用（`col=9.78cm`、x = `1.5 / 12.04 / 22.58`）。§業界別 コンシューマー行 2 の必須要素; あらゆる業界の Seed / 早期 A において「市場ウィンドウ」自体がテーゼである場合に有効。

```bash
SLIDE=3
officecli add "$FILE" / --type slide --prop layout=blank --prop background=FFFFFF
officecli add "$FILE" "/slide[$SLIDE]" --type shape --prop text="Why now: three converging triggers" \
  --prop x=1.5cm --prop y=1.2cm --prop width=30.87cm --prop height=2.5cm \
  --prop font=Georgia --prop size=36 --prop bold=true --prop color=1E2761 --prop fill=none
# Card 1 (x=1.5cm) — trigger / data / implication / source. Repeat at x=12.04cm and x=22.58cm.
cat <<EOF | officecli batch "$FILE"
[
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"geometry":"roundRect","fill":"F5F7FA","x":"1.5cm","y":"5cm","width":"9.78cm","height":"10cm"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"BOM cost","x":"1.5cm","y":"5.5cm","width":"9.78cm","height":"1.2cm","font":"Calibri","size":"24","bold":"true","color":"1E2761","align":"center","fill":"none"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"−90%","x":"1.5cm","y":"7cm","width":"9.78cm","height":"3cm","font":"Georgia","size":"60","bold":"true","color":"B85042","align":"center","fill":"none"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"Wearable BOM fell 90% since 2021; sub-$40 retail now viable","x":"1.5cm","y":"11cm","width":"9.78cm","height":"2cm","font":"Calibri","size":"16","color":"333333","align":"center","fill":"none"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"Source: IDC Wearables Teardown 2025","x":"1.5cm","y":"13.5cm","width":"9.78cm","height":"1cm","font":"Calibri","size":"12","italic":"true","color":"666666","align":"center","fill":"none"}}
]
EOF
# Card 2 pattern: Oura IPO 2024 / +$2.4B valuation / category proven. Card 3: On-device LLM (Llama 3.2) / Q4-24 / privacy moat viable.
```

**QA。** 3 カードそれぞれの出典脚注に年 / 日付の引用があり、各カードは 30 語以内。`officecli query "$FILE" 'shape:contains("2024")'` + `'shape:contains("2025")'` の合計が 2 以上（タイミングのアンカーが可視化されている）。

### (3) Solution スライド — プロダクトを 1 文で + 3 ステップの「仕組み」

**視覚的成果。** プロダクトのパターンを名指しする 36pt タイトル（「Our Solution」ではなく）。その下: y=7cm に横並びの角丸ボックス 3 – 4 個、エルボーコネクタ + 三角矢印付き。各ボックス = 1 つの動詞（observe / correlate / resolve）。pptx Recipe (c) のフローチャートを再利用する — 新しいプリミティブではなく、オーケストレーションである。

```bash
# Title — "a product pattern, not a brand slogan".
# Good: "Auto-correlate K8s events across 3 data planes in 90 seconds"
# Bad:  "The future of observability"
SLIDE=4
officecli add "$FILE" / --type slide --prop layout=blank --prop background=FFFFFF
officecli add "$FILE" "/slide[$SLIDE]" --type shape --prop name=SolTitle \
  --prop text="Correlate K8s events across 3 data planes in 90 seconds" \
  --prop x=1.5cm --prop y=1.2cm --prop width=30.87cm --prop height=2.2cm \
  --prop font=Georgia --prop size=32 --prop bold=true --prop color=1E2761 --prop fill=none
# 3 boxes across: gap = (33.87 − 3 − 3·7) / 2 = 4.93cm; x = 1.5, 13.43, 25.36
# Connectors + arrowheads: --prop tailEnd=triangle ALWAYS (pptx Known Issues C-P-5..6).
# Full batch block → see pptx v2 §Creating and Editing (c) 4-step flowchart; swap N from 4 boxes to 3.
```

**プロダクトパターンのタイトルルール。** Solution のタイトルは、動詞 + 差別化されたメカニズム + 指標である。「Observe / Correlate / Resolve」は一般的すぎて、VC にはどこにでもある APM ベンダーに読める。「Correlate K8s events across 3 data planes in 90 seconds」は具体的で、VC には洞察として読まれる。

**QA。** コネクタ数をカウント: `officecli query "$FILE" 'connector' --json | jq '.data.results | length'` が (step_count − 1) 以上。すべてのコネクタは `tailEnd=triangle` を持たなければならない — `view annotated` で矢印の向きを確認できる。タイトルは 12 語以内（一息で読める長さ）。

### (4) Market スライド — TAM / SAM / SOM のネストされたカラム

**視覚的成果。** 36pt タイトル「Market: $X.YB growing Z% CAGR」。その下: TAM / SAM / SOM のラベルとドル値 + 成長率を付した 3 本の横棒（または 3 段のネスト長方形）。下部の脚注には**トップダウンかボトムアップかの出典**を明記する — 1 デックにつき 1 つの手法を選び、混在させないこと。

```bash
# Use a pptx column chart with 3 values. Categories = TAM,SAM,SOM. Source annotation is a separate shape.
SLIDE=5
officecli add "$FILE" / --type slide --prop layout=blank --prop background=FFFFFF
officecli add "$FILE" "/slide[$SLIDE]" --type shape --prop text="$42B observability market, 18% CAGR" \
  --prop x=1.5cm --prop y=1.2cm --prop width=30.87cm --prop height=2cm \
  --prop font=Georgia --prop size=36 --prop bold=true --prop color=1E2761 --prop fill=none
officecli add "$FILE" "/slide[$SLIDE]" --type chart --prop chartType=bar \
  --prop series1.name="USD (billions)" --prop series1.values="42,8.4,0.62" --prop series1.color=1E2761 \
  --prop categories="TAM,SAM,SOM (5-yr)" \
  --prop x=2cm --prop y=4cm --prop width=22cm --prop height=12cm \
  --prop title='Market sizing — bottom-up by enterprise count × ACV'
officecli add "$FILE" "/slide[$SLIDE]" --type shape --prop text='Source: Gartner 2025 APM Magic Quadrant; SAM = 20% of TAM (K8s-first shops); SOM = 7.4% of SAM over 5 years at 18-24% share.' \
  --prop x=2cm --prop y=16.5cm --prop width=29.87cm --prop height=2cm \
  --prop font=Calibri --prop size=12 --prop italic=true --prop color=666666 --prop fill=none
```

**QA。** トップダウンかボトムアップかは出典脚注に必ず明記すること。手法のない TAM は捏造に見える。

### (5) Product スライド — スクリーンショット + 3 つの箇条書き、または 3 カードのフィーチャーグリッド

**視覚的成果。** 2 つのレイアウト選択肢: (a) 左側にヒーロープロダクトのスクリーンショット（スライドの 60%）、右側に 1 行のフィーチャー箇条書き 3 個（それぞれ本文 18pt 以上、箇条書きの中の箇条書きは不可）。(b) アイコン / スクリーンショットのサムネイルをそれぞれ 1 つ持つフィーチャーカード 3 枚。コンシューマー / アプリ系のプロダクトには (a) を、B2B / インフラ系には (b) を選ぶこと。

```bash
# (a) screenshot + bullets — consumer pattern
officecli add "$FILE" "/slide[$SLIDE]" --type picture --prop src=product_hero.png \
  --prop x=1cm --prop y=4cm --prop width=18cm --prop height=13cm
officecli set "$FILE" "/slide[$SLIDE]/picture[1]" --prop alt="Product UI: dashboard with 12 K8s clusters, live correlation graph"
# Right column bullets (each as a separate shape so sizes stay explicit)
officecli add "$FILE" "/slide[$SLIDE]" --type shape --prop text="Auto-correlate across 3 data planes" \
  --prop x=20cm --prop y=5cm --prop width=12cm --prop height=1.5cm \
  --prop font=Calibri --prop size=20 --prop bold=true --prop color=1E2761 --prop fill=none
# Repeat for bullets 2 and 3 at y=7.5cm / y=10cm.
```

**QA。** 画像の alt テキストが存在する（`query 'picture:no-alt'` が空）。箇条書きはそれぞれ 18pt 以上。「Lorem」/「product name here」/`{{...}}` のようなトークンがないこと。

### (6) Business model スライド — ユニットエコノミクスまたは収益モデル

**視覚的成果。** 業界ごとの決定木:
- **SaaS / エンタープライズ（Series A 以降）** — CAC / LTV / Payback / GM の 4 つの KPI コールアウト（pptx Recipe (e) を再利用）。
- **コンシューマー / D2C** — AOV · リピート購入率 · 貢献利益 · ブレンド CAC。
- **マーケットプレイス** — GMV / テイクレート / 流動性指標 / コホートリテンション。
- **バイオ / ディープテック** — 想定レンジ付きの収益モデル（ライセンス / マイルストーン / ロイヤリティ分配）。

タイトルは支配的な指標を名指しする（例: 「LTV:CAC 4.7x · 14-month payback · 78% gross margin」）— 「Business Model」ではない。4 カード分のフルバッチブロック → pptx v2 §(e) KPI callouts を参照。

```bash
# SaaS pattern: KPI card values + sub-label + gray VC-floor context under each.
# Card 1 (LTV): big number "$420K", sub "Lifetime value", context "floor: ARPU × GM / churn"
# Card 2 (CAC): big number "$90K",  sub "Acquisition cost", context "fully-loaded S&M spend"
# Card 3 (Payback): big number "14 mo", sub "CAC payback", context "VC floor: < 18 mo"
# Card 4 (GM): big number "78%", sub "Gross margin", context "SaaS floor: 70%+"
# Grid math for 4 cards across: usable = 33.87 − 3 − 3·0.76 = 28.59, col = 7.15cm
# → Full batch template → pptx v2 §(e). Adapt card count 3→4 and card width 9.78cm→7.15cm.
```

**QA。** Series B 以降の場合、{CAC, LTV, payback, GM} のすべてが存在すること: `officecli query "$FILE" 'shape:contains("CAC")'` ≥ 1 AND `shape:contains("LTV")'` ≥ 1 AND `shape:contains("payback")'` ≥ 1 AND `shape:contains("gross margin")'` ≥ 1。

### (7) Traction スライド — 0 起点の ARR カーブ

**視覚的成果。** スライド幅の 60% を占める折れ線グラフ; ARR は y 軸で **0 起点**（現在値の 80% から始める — VC のホッケースティックの嘘ではない）。右側のコメンタリーカード: 1 つの大きな数字（現在の ARR）+ 成長率 + 2 – 3 個のマイルストーン。Series B 以降なら 2 段目にコホートリテンションの抜粋またはロゴウォール。

```bash
SLIDE=7
officecli add "$FILE" / --type slide --prop layout=blank --prop background=FFFFFF
officecli add "$FILE" "/slide[$SLIDE]" --type shape --prop text='ARR: $0 → $18M in 24 months' \
  --prop x=1.5cm --prop y=1.2cm --prop width=30.87cm --prop height=2cm \
  --prop font=Georgia --prop size=36 --prop bold=true --prop color=1E2761 --prop fill=none
officecli add "$FILE" "/slide[$SLIDE]" --type chart --prop chartType=line \
  --prop series1.name=ARR --prop series1.values="0.2,0.6,1.4,3.2,6.1,11.3,15.8,18.0" --prop series1.color=1E2761 \
  --prop categories="Q1-24,Q2-24,Q3-24,Q4-24,Q1-25,Q2-25,Q3-25,Q4-25" \
  --prop x=1.5cm --prop y=4cm --prop width=21cm --prop height=13cm \
  --prop title='Quarterly ARR ($M) — y-axis anchored at 0' \
  --prop axismin=0
# Right callout
officecli add "$FILE" "/slide[$SLIDE]" --type shape --prop geometry=roundRect --prop fill=1E2761 --prop line=none \
  --prop x=23.5cm --prop y=4cm --prop width=8.8cm --prop height=13cm
officecli add "$FILE" "/slide[$SLIDE]" --type shape --prop text='$18M' \
  --prop x=23.5cm --prop y=5cm --prop width=8.8cm --prop height=3cm \
  --prop font=Georgia --prop size=64 --prop bold=true --prop color=FFFFFF --prop align=center --prop fill=none
officecli add "$FILE" "/slide[$SLIDE]" --type shape --prop text="ARR · +312% YoY · NRR 128%" \
  --prop x=23.5cm --prop y=9cm --prop width=8.8cm --prop height=3cm \
  --prop font=Calibri --prop size=18 --prop color=CADCFC --prop align=center --prop fill=none
```

**`--prop axismin=0` は必須要件である** — これがないと、pptx は最低値付近から始まるよう y 軸を自動スケールしてしまう。それがホッケースティックの嘘である。Gate 6 が以下でこれを grep する。

**QA。** ARR カーブのチャートは `axismin=0` を持たなければならない。`officecli get "$FILE" "/slide[$SLIDE]/chart[1]" --json | jq .format.axisMin` が `0` を返す（入力プロパティは小文字の `axismin` だが、CLI は読み戻し時にキャメルケースの `axisMin` を返す）。

### (8) Team スライド — アバター + 名前 + 過去の所属企業（単なる壁ではなく）

**視覚的成果。** スライド中央に 3 または 4 枚のカードが横並び。各カード: 上部に写真（6×6cm）; 名前（20pt 太字）; 役職（16pt）; **過去の所属企業 + 肩書き**（16pt イタリック、1 行の要点）; 任意で LinkedIn URL のフッター（12pt）。顔写真と名前だけのチームスライドは素人臭く見える。

```bash
SLIDE=11
officecli add "$FILE" / --type slide --prop layout=blank --prop background=FFFFFF
officecli add "$FILE" "/slide[$SLIDE]" --type shape --prop text="Team: 3 prior exits, 42 years combined K8s" \
  --prop x=1.5cm --prop y=1.2cm --prop width=30.87cm --prop height=2cm \
  --prop font=Georgia --prop size=36 --prop bold=true --prop color=1E2761 --prop fill=none
# Card 1 — CEO
officecli add "$FILE" "/slide[$SLIDE]" --type picture --prop src=alice.jpg \
  --prop x=2cm --prop y=5cm --prop width=6cm --prop height=6cm
officecli set "$FILE" "/slide[$SLIDE]/picture[1]" --prop alt="Alice Chen, CEO — portrait"
officecli add "$FILE" "/slide[$SLIDE]" --type shape --prop text="Alice Chen" \
  --prop x=2cm --prop y=11.5cm --prop width=6cm --prop height=1cm \
  --prop font=Georgia --prop size=20 --prop bold=true --prop color=1E2761 --prop align=center --prop fill=none
officecli add "$FILE" "/slide[$SLIDE]" --type shape --prop text="CEO" \
  --prop x=2cm --prop y=12.8cm --prop width=6cm --prop height=0.8cm \
  --prop font=Calibri --prop size=16 --prop color=333333 --prop align=center --prop fill=none
officecli add "$FILE" "/slide[$SLIDE]" --type shape --prop text="ex-Datadog Director (Series C → IPO); led K8s observability GTM $40M → $200M ARR" \
  --prop x=2cm --prop y=13.8cm --prop width=6cm --prop height=2.5cm \
  --prop font=Calibri --prop size=14 --prop italic=true --prop color=333333 --prop align=center --prop fill=none
# Repeat for Card 2 (CTO, x=10cm) and Card 3 (VP Eng, x=18cm) — 3 cards × 5-6 shapes each.
```

過去の所属企業は**信頼性の密度**を運ぶ。VC は「ex-Datadog Director + led $40M → $200M」を 2 秒で読み取るが、「co-founder, passionate」は 0 秒で読み飛ばす。アドバイザーを載せる場合は、下に小さめの行としてロゴ 1 個ずつで表示すること。

**配置ヘルパー。** 3 カード: `col=9.78cm, x=1.5/12.04/22.58`。4 カード: `col=7.15cm, x=1.5/9.41/17.32/25.23`。5 カード: `col=5.85cm, x=1.5/7.75/14.0/20.25/26.5`（ギャップ 0.4cm、より詰めた配置）。6 枚以上や非対称な場合は 2 段グリッド（3×2 / 3×3）にする; pptx v2 §(d) グリッド数学を参照。

**QA。** `officecli query "$FILE" 'shape:contains("ex-")'` + `'shape:contains("prior")'` + `'shape:contains("former")'` がメンバー 1 人あたり 1 以上。ゼロなら、それはチームではなくポートフォリオである。

### (9) Financials スライド — 4 年計画 + 誠実な前提

**視覚的成果。** カラムチャート: 4 年 × (revenue, gross margin $, EBITDA)。右側のカード: 3 項目の前提パネル（ARPU の前提、勝率の前提、チャーンの前提）。タイトルは軌跡を名指しする（「$18M → $85M by FY29」）— 「Financial Projections」ではない。

pptx Recipe (b) のチャート + コメンタリーを再利用する。ピッチ特有の点: 右側の ASSUMPTIONS 列は**必須要件である** — 前提が見えない 4 年計画は願望に見える。VC はどのみちすべての数字の裏付けを尋ねてくる; あらかじめ提示しておくこと。

左 2/3 — スライド + タイトル + 3 系列のカラムチャート:

```bash
SLIDE=17
officecli add "$FILE" / --type slide --prop layout=blank --prop background=FFFFFF
officecli add "$FILE" "/slide[$SLIDE]" --type shape --prop text='$18M → $85M ARR by FY29' \
  --prop x=1.5cm --prop y=1.2cm --prop width=30.87cm --prop height=2cm \
  --prop font=Georgia --prop size=36 --prop bold=true --prop color=1E2761 --prop fill=none
officecli add "$FILE" "/slide[$SLIDE]" --type chart --prop chartType=column \
  --prop series1.name="Revenue ($M)"  --prop series1.values="18,34,58,85" --prop series1.color=1E2761 \
  --prop series2.name="Gross Margin ($M)" --prop series2.values="14,26,45,68" --prop series2.color=CADCFC \
  --prop series3.name="EBITDA ($M)"   --prop series3.values="-6,-2,8,22" --prop series3.color=B85042 \
  --prop categories="FY26,FY27,FY28,FY29" \
  --prop x=1.5cm --prop y=4cm --prop width=20cm --prop height=13cm \
  --prop title='4-year plan — revenue, GM, EBITDA ($M)'
```

右 1/3 — 前提コメンタリーカード:

```bash
officecli add "$FILE" "/slide[$SLIDE]" --type shape --prop geometry=roundRect --prop fill=F5F7FA --prop line=none \
  --prop x=22.5cm --prop y=4cm --prop width=9.8cm --prop height=13cm
officecli add "$FILE" "/slide[$SLIDE]" --type shape --prop text="Key Assumptions" \
  --prop x=23cm --prop y=4.5cm --prop width=8.8cm --prop height=1.2cm \
  --prop font=Georgia --prop size=20 --prop bold=true --prop color=1E2761 --prop fill=none
# 5 assumption bullets as 5 separate paragraph shapes at y=6, 7.5, 9, 10.5, 12cm — size=14, italic=true.
# Keep each bullet ≤ 14 words so 8.8cm width fits without wrap.
```

**前提パネルは必須要件である。** 前提が見えない 4 年計画は願望に見える。VC はどのみちすべての数字の裏付けを尋ねてくる — カーブを牽引する 3 – 4 個の前提を明示すること。

**QA。** `officecli query "$FILE" 'shape:contains("assumption")'` OR `'shape:contains("Assumes")'` ≥ 1。ゼロならパネルを追加すること。

### (10) The Ask — 主役数値 + 4 分割 Use-of-Funds + ランウェイ

**視覚的成果。** ダーク塗り（カバーと合わせる）。中央上部に主役数値: `$35M` を 96pt の白字で。その下に、4 分割の円グラフ、または **Engineering 40% / GTM 35% / G&A 15% / Reserve 10%** を列挙する 4 カード行。最下行: 「18-month runway to $40M ARR」（次のマイルストーン、「次のラウンドまで」ではない）。

```bash
SLIDE=20
officecli add "$FILE" / --type slide --prop layout=blank --prop background=1E2761
officecli add "$FILE" "/slide[$SLIDE]" --type shape --prop text='$35M Series B' \
  --prop x=2cm --prop y=2cm --prop width=29.87cm --prop height=4cm \
  --prop font=Georgia --prop size=88 --prop bold=true --prop color=FFFFFF --prop align=center --prop fill=none
officecli add "$FILE" "/slide[$SLIDE]" --type chart --prop chartType=pie \
  --prop series1.name="Use of Funds" --prop series1.values="40,35,15,10" \
  --prop categories="Engineering,Go-to-Market,G&A,Reserve" \
  --prop colors="CADCFC,B85042,97BC62,FFFFFF" \
  --prop x=6cm --prop y=7cm --prop width=12cm --prop height=10cm \
  --prop title="Use of Funds — 4 buckets"
officecli add "$FILE" "/slide[$SLIDE]" --type shape --prop text='18 months runway to $40M ARR and Series C' \
  --prop x=2cm --prop y=17cm --prop width=29.87cm --prop height=1.5cm \
  --prop font=Calibri --prop size=22 --prop color=CADCFC --prop align=center --prop fill=none
```

**4 分割の規約。** Engineering / GTM / G&A / Reserve が正規の内訳である。典型的な Series A の割合レンジ: Eng 40-50%、GTM 30-40%、G&A 10-15%、Reserve 5-10%。Series B では Eng から GTM へ 5-10 ポイントシフトする。

**QA。** `officecli query "$FILE" 'shape:contains("Use of Funds")'` ≥ 1。ask スライドに円グラフが存在すること。ask スライドにランウェイ + マイルストーンが存在すること。

### (11) Pipeline チャート — バイオ / ディープテックの必須要素

**視覚的成果。** 横型スイムレーン。左のカラム = 候補名; 右に 4 個のステージカラム（バイオなら Preclinical / Ph1 / Ph2 / Ph3、ディープテックなら TRL1-3 / TRL4-6 / TRL7-8 / TRL9）。各行のバーは現在のステージまで伸び、後段のステージほど濃い塗り。下部に NCT / 治験 ID のフッター。§業界別 行 5 バイオの必須要素; SaaS / コンシューマーは省略。

グリッド数学: usable `= 30.87cm`、候補名カラム `= 7cm`、ステージカラム `= (30.87 − 7) / 4 = 5.97cm` ずつ、行の高さ `= 2.3cm`。ステージカラムの x: `8.5 / 14.47 / 20.44 / 26.41`。

```bash
SLIDE=6
officecli add "$FILE" / --type slide --prop layout=blank --prop background=FFFFFF
officecli add "$FILE" "/slide[$SLIDE]" --type shape --prop text="Pipeline: 3 candidates across Ph1–Ph3" \
  --prop x=1.5cm --prop y=1.2cm --prop width=30.87cm --prop height=2cm \
  --prop font=Georgia --prop size=36 --prop bold=true --prop color=1E2761 --prop fill=none
# 4 stage headers + candidate row 1 (HLX-201 at Ph2, bar width = 3·5.97 = 17.91cm) in one batch.
cat <<EOF | officecli batch "$FILE"
[
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"Preclinical","x":"8.5cm","y":"4cm","width":"5.97cm","height":"1cm","font":"Calibri","size":"16","bold":"true","color":"333333","align":"center","fill":"F5F7FA"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"Phase 1","x":"14.47cm","y":"4cm","width":"5.97cm","height":"1cm","font":"Calibri","size":"16","bold":"true","color":"333333","align":"center","fill":"F5F7FA"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"Phase 2","x":"20.44cm","y":"4cm","width":"5.97cm","height":"1cm","font":"Calibri","size":"16","bold":"true","color":"333333","align":"center","fill":"F5F7FA"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"Phase 3","x":"26.41cm","y":"4cm","width":"5.97cm","height":"1cm","font":"Calibri","size":"16","bold":"true","color":"333333","align":"center","fill":"F5F7FA"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"text":"HLX-201 (lead)","x":"1.5cm","y":"5.5cm","width":"7cm","height":"1.5cm","font":"Calibri","size":"18","bold":"true","color":"1E2761","align":"left","fill":"none"}},
  {"command":"add","parent":"/slide[$SLIDE]","type":"shape","props":{"geometry":"roundRect","fill":"1E2761","x":"8.5cm","y":"5.7cm","width":"17.91cm","height":"1.1cm","line":"none"}}
]
EOF
# Repeat rows 2 & 3 at y=7.8cm / y=10.1cm with bar widths per stage (Ph1=5.97cm, Ph1-Ph2=11.94cm, Ph1-Ph3=17.91cm).
# NCT footer full-width at y=16.8cm.
officecli add "$FILE" "/slide[$SLIDE]" --type shape --prop text='NCT05021323 (HLX-201, Ph2, n=48) · NCT06142091 (HLX-304, Ph1, n=24) · IND-filed Q1-26 for HLX-412' \
  --prop x=1.5cm --prop y=16.8cm --prop width=30.87cm --prop height=1.2cm \
  --prop font=Calibri --prop size=12 --prop italic=true --prop color=666666 --prop fill=none
```

**QA。** `officecli query "$FILE" 'shape:contains("NCT")' --json | jq '.data.results | length'` ≥ 1。ステージが進むほどバーの色が濃くなる（`CADCFC` = preclinical のみ、`1E2761` = Ph2 到達）。

### (12) 競合比較表 — Series B 以降の必須要素

**視覚的成果。** 5 – 7 行 × 4 – 6 列。列 1 = 競合名（任意でロゴシェイプを添える）; 残りは差別化要素（速度 / 価格 / インテグレーション数 / マージン / カバレッジ）。**最終行 = 自社、アクセントカラー（CADCFC / 97BC62）でハイライト**; 競合行はグレー。Series B 以降のすべてのデックにこれが必要（SaaS: Datadog / New Relic / Splunk; バイオ: Kite / Novartis / BMS）。

```bash
SLIDE=13
officecli add "$FILE" / --type slide --prop layout=blank --prop background=FFFFFF
officecli add "$FILE" "/slide[$SLIDE]" --type shape --prop text="Competitive landscape" \
  --prop x=1.5cm --prop y=1.2cm --prop width=30.87cm --prop height=2cm \
  --prop font=Georgia --prop size=36 --prop bold=true --prop color=1E2761 --prop fill=none
# Inline table via --prop data= (per-cell r#c# is silently ignored / not honored — use data=). Single-quote the data value — '$15/host' would strip.
officecli add "$FILE" "/slide[$SLIDE]" --type table \
  --prop data='Competitor,Speed,Price,Integrations,Margin;Datadog,12 min,$15/host,680,75%;New Relic,18 min,$25/host,520,68%;Splunk,45 min,$45/GB,310,62%;You (Acme DevOps),90 sec,$8/host,1200,82%' \
  --prop style=medium1 --prop headerFill=1E2761 \
  --prop x=1.5cm --prop y=4cm --prop width=30.87cm --prop height=12cm
# Highlight your row: loop over /slide[$SLIDE]/table[1]/tr[5]/tc[1..5] and set cell fill to CADCFC.
```

**QA。** `officecli query "$FILE" 'table' --json | jq '.data.results | length'` ≥ 1。行数 ≥ 4（自社 + 名指しの競合 3 社以上）。自社の行がセル塗りにより視覚的に区別されている（Gate 5b の視覚チェック — テーブルスタイルだけでは 1 行をハイライトできない）。

## 数値表記規約（ピッチ特有）

簡潔な規約表 — **ファイナンスの教科書ではない**。これらの意味が分からない場合は、デックの作成を止めて数値をユーザーに確認すること; 推測しないこと。

| 指標 | 形式 | Floor / 規約 |
|---|---|---|
| **TAM** | `$X.YB`、手法は 1 つ | トップダウン（アナリストレポート）かボトムアップ（件数 × ACV）のどちらか。両方は不可; どちらもないのも不可。 |
| **SAM** | `$X.YB`、TAM に占める割合 | 垂直特化型 SaaS では通常 TAM の 15 – 30%; 水平型ではより高い |
| **SOM** | N 年目時点の `$X.YB` | 現実的な 5 年シェア: 早期段階では SAM の 5 – 15% |
| **ARR** | MRR × 12。revenue ではない。 | SaaS のみ; 帳簿上の契約、チャーン控除後 |
| **MRR** | 月次経常収益 | ARR / 12; 月次売上と混同しないこと |
| **NRR (Net Revenue Retention)** | %、直近 12 か月 | VC floor: 100% 超で合格、115% 超で好調、130% 超で卓越 |
| **CAC** | $、フルロード | セールス + マーケティング支出 / 新規獲得ロゴ数 |
| **LTV** | $ | ARPU × 粗利率 × (1 / チャーン率) |
| **LTV:CAC** | 比率 | VC floor: 3x で合格、4x 超で好調、5x 超で卓越 |
| **CAC payback** | 月数 | VC floor: 18 か月未満で合格、12 か月未満で好調 |
| **Gross margin** | % | SaaS floor 70%、好調は 80%以上; マーケットプレイスは 15-40%; ハードウェアは 30-50% |
| **Burn / runway** | $/月 + 月数 | 総バーンか純バーンか — どちらか明記; 特定のマイルストーンまでのランウェイ |
| **Use of Funds** | 4 分割円グラフ | Engineering / Go-to-Market / G&A / Reserve — Ask スライドレシピを参照 |

**ルール。** デック上のすべての数字は単位を伴う。`18%` や `18M` 単独では曖昧である — `$18M ARR` / `18% NRR growth` のように書くこと。数値枠内の `TBD`、`coming soon`、`(fill in)`、`lorem`、`xxxx` は即座に VC から失格とされる。Gate 6 が以下でこれらを grep する。

## VC 出荷前チェック（6 つのレッドフラグ / ポジティブシグナル）

VC が最初の 30 秒で読み取るもの。6 つの一行条件 — 以下の「FAIL」はいずれも即座にラウンドを殺す — 納品前に修正すること。

| # | レッドフラグ（該当すれば FAIL） | ポジティブシグナル（出荷可） |
|---|---|---|
| 1 | ラウンド + 金額 + 日付のないカバー | `Company · tagline · Series X · $YM · Date` を 4 行で |
| 2 | 出典 / 手法の引用なしに TAM > $100B | TAM が明確にボトムアップまたはトップダウンとラベル付けされ、2024 年以降の出典が可視化されている |
| 3 | トラクションチャートの y 軸が 0 起点でない（ホッケースティックの嘘） | 折れ線グラフが `axismin=0`; 成長曲線が誠実 |
| 4 | チームスライドが顔写真 + 名前のみで過去の所属企業なし | 全メンバーについて: 過去の所属企業 + 役職 + 実績指標 1 件 |
| 5 | Ask スライドに Use-of-Funds の内訳がない | `$XM` の主役数値 + 4 分割円グラフ（Eng / GTM / G&A / Reserve）+ ランウェイ + 次のマイルストーン |
| 6 | どこかに `TBD` / `lorem` / `xxxx` / `{{...}}` / `(fill in)` がある | `view text` がクリーン — プレースホルダートークンがゼロ |

**ラウンド固有によくある失敗。**
- **Series A 固有** — 架空の企業数 × ACV から算出したボトムアップ TAM（数を裏付ける参照顧客がない）; 12 か月未満のデータで `CAC / LTV` を提示（統計的に無意味）。
- **Series B 固有** — ユニットエコノミクススライドが皆無; 「まだスケール前だが、これが計画だ」というナラティブを伴わずに CAC payback が 24 か月超; ロゴウォールの顧客数 8 未満。
- **Series C 固有** — moat / 防御可能性のスライドがない; マージンの推移を伴わない売上成長; 国際展開を謳いながら具体的なローンチ計画 / 採用計画がない。

以下の Delivery Gate 6 ブロックが、上記チェック 1 – 6 を grep + query で実行する。Gate 5b の fresh-eyes は、grep では見えない視覚的判断（ホッケースティック、チームの信頼性）をカバーする。

## トラクションの三重パターン（ARR + マイルストーン + ロゴ）

Series B 以降では、トラクションはしばしば 2 枚のスライドにまたがる: 1 枚はチャート + コールアウト（上記レシピ 7）、もう 1 枚は**マイルストーンのタイムライン + ロゴウォール**。タイムライン = 4 – 6 個の日付を横一列に、それぞれ 1 行のイベント付きで。ロゴウォール = 4×N または 5×N グリッドに 12 – 20 個の顧客ロゴ、単一のブランドが目立たないよう落ち着いたモノクロで。

```bash
# Milestone timeline: 5 dates as circles on a horizontal line at y=8cm.
# Use pptx shapes (ellipse preset) + connectors (shape=straight) between them.
# Each milestone = ellipse at y=8cm + date label above + event description below.
# → See pptx v2 Recipe (d) row 9 (Roadmap timeline) for the canonical pattern.

# Logo wall: pictures in a 5×N grid. Typical spacing: logo width = 5cm, height = 2cm, gap = 0.4cm.
# grid math for 5 logos across, 1.5cm edge margin: usable = 33.87 − 3 − 4·0.4 = 29.27, col = 5.85cm
# (use 5cm logo width centered in each 5.85cm column)
```

**QA。** ロゴウォールは Series B 以降で 8 個以上、Series A で 4 個以上あるべき。それより少ないと「見かけより軽い」印象になり、20 を超えるとピクセルノイズになる。

## QA — Delivery Gate（実行可能）

**問題は必ずあるものと想定すること。** 初回のレンダリングがそのまま正しいことはほぼない。ピッチデックは 2 つの層で失敗する: **構造的な失敗**（スキーマ、トークン漏れ — pptx v2 Gates 1–3 が捕捉する）と、**ナラティブの失敗**（ステージの誤り、ユニットエコノミクスの欠落、TAM に出典がない — これらは pptx v2 Gate 5b + Gate 6 を不可欠にするチェックである）。すべてのチェックは成功メッセージを出力しなければならない。

### Gates 1–5a — pptx v2 からそのまま継承

→ pptx v2 §Delivery Gate L637-679 を参照。ブロック全体をコピー&ペーストすること:

- **Gate 1** — `validate` によるスキーマチェック（C-P-2 に基づき `ChartShapeProperties` の警告をホワイトリスト化）。
- **Gate 2** — `view text` の grep によるトークン漏れ検出（`$xxx$`、`{{...}}`、`<TODO>`、`lorem`、`xxxx`、空の `()`/`[]`、`\$`/`\t`/`\n` のリテラル）。
- **Gate 3** — ハイパーリンク `rPr` のスキーマトラップ（C-P-1）— `<a:rPr><a:hlinkClick>` がゼロであること。
- **Gate 4** — スライド順の健全性チェック — カバーが最初、区切りがセクションの前、closing が最後。
- **Gate 5a** — ダーク背景上のダークテキストのコントラスト — `{1E2761, 0A1628, 8B1A1A, 2C5F2D, 36454F}` のいずれかの塗りには、ほぼ白のテキストカラーを明記しなければならない。**その塗りの上に描画されるチャートも含む**: チャートの `title.textColor`、`legend.textColor`、軸テキストは既定で暗色になり、ダーク背景上では見えなくなる — 明示的に設定するか、ダークスライド内の明色カードの上にチャートを配置すること。

これら 5 つは省略も順序変更もしないこと。Gates 1–5a が捕捉するすべての pptx レイヤーの欠陥は、ピッチデックでも発生する。

**Gate 2b — ピッチ特有のシェル strip シグネチャ（必須）。** Gate 2 は zsh が黙って空文字に strip した `$35M` を見逃す（痕跡が残らないため）。Gate 2 の後にこれを実行すること:

```bash
# $XXM stripped by zsh leaves bare " M ARR" / " M raised" / "Series [A-C] · M" patterns.
STRIP=$(officecli view "$FILE" text | grep -niE '(^|[^A-Za-z0-9])M (ARR|raised|Series|runway|round|raise)|Series [A-C] · M( |$)|runway · M|raised · M|raising ·? M')
[ -z "$STRIP" ] && echo "Gate 2b OK (no \$-strip signatures)" || { echo "REJECT Gate 2b (likely zsh \$-strip — re-issue with single quotes):"; echo "$STRIP"; exit 1; }
```

修正方法: 問題の `add`/`set` を、値をシングルクォートして再実行すること（`--prop text='Series B · $35M'`、ダブルクォートではなく）。同じ strip は**チャートのシリーズ名 / 軸タイトル**にも起こる（`--prop name="营收 ($M)"` → 凡例に `营收 ()` と表示される）: `$` を含むすべてのチャートプロパティをシングルクォートすること。

### Gate 5b — HTML プレビューによる視覚監査（必須、任意ではない）

Gates 1–5a はトークン grep による防御である。**レンダリングされたスライドを見ることはできない。** このステップだけが唯一の視覚的な組み立てチェックである。省略しないこと。

`officecli view "$FILE" html` を実行し、返された HTML を読むこと。すべてのスライドを一通り確認し、各項目について答えること（pptx v2 Gate 5b のチェックリストを継承; ピッチ特有の追加項目には ⭐ を付す）:

- **重なり**: テキストシェイプ同士、またはチャートと重なっていないか？
- **暗色上の暗色**: 塗りの明度 < 30% かつテキストの明度 < 80% のテキストがないか？
- **ディバイダーの重なり**: 巨大な装飾数字（01/02/03 の 100pt 以上）がディバイダーのタイトルテキストと衝突していないか？
- **順序の健全性**: スライドの並びは、あなたのステージに適したナラティブのアウトラインと一致しているか？
- **矢印の欠落**: フローチャート / 決定木のコネクタは方向を示しているか、それとも単なる線か？
- ⭐ **トラクションの y 軸**: すべての ARR / revenue / growth の折れ線グラフの y 軸は 0 起点になっているか？（現在値の 80% ではない — それはホッケースティックの嘘である。）
- ⭐ **チームの信頼性**: すべてのチームスライドのカードが過去の所属企業または過去の肩書きを示しているか？（顔写真 + 名前だけのカードは reject。）
- ⭐ **TAM / 市場数値の信頼性**: TAM はニッチ市場に対して $100B 未満か、あるいは $100B 以上の場合、手法の出典が引用されているか？（出典のない `$500B TAM` の主張は自動的な reject のレッドフラグである。）
- ⭐ **Use-of-Funds の円グラフ**: ask スライドに 4 分割円グラフ（Engineering / GTM / G&A / Reserve）、または %付きの 4 カード行があるか？
- ⭐ **ナラティブの完全性**: 順序が cover → problem → solution → market → product → model → traction → team → financials → ask、またはあなたのステージに適した §Stage diagnosis からの順列になっているか？

**指示。** `officecli view "$FILE" html` を実行し、HTML を読むこと。すべてのスライドを以下の質問に沿って確認すること。チャートの色、アニメーション、ズームをレンダリングする場合 — それらは対象ビューア（PowerPoint / Keynote / WPS）でのみ表示されるため、これらのランタイム機能についてはユーザーに `.pptx` を直接開くよう依頼すること。

> すべてのスライドについて:
> (a) スライドは VC のナラティブ順序（cover → problem → solution → market → product → model → traction → team → financials → ask、あなたのステージに応じた調整を含む）になっているか？順序違反があれば指摘すること。
> (b) すべての ARR / revenue / growth の折れ線グラフの y 軸は 0 起点になっているか？ホッケースティックの視覚的な嘘を指摘すること。
> (c) チームスライドは各メンバーについて過去の所属企業の実績を示しているか？（顔写真 + 名前だけではなく。）
> (d) すべての TAM / SAM / SOM の主張に、可視化された出典や手法があるか？
> (e) ask スライドには 4 分割の Use of Funds（Engineering / GTM / G&A / Reserve）と、具体的な次のマイルストーン + ランウェイ期間があるか？
> (f) テキストの重なり、暗色上の暗色、スライド外へのはみ出し、矢印の欠落、プレースホルダートークン（`TBD` / `lorem` / `{{...}}` / `xxxx` / 空の `()`）はないか？

欠陥があれば、それぞれをスライド番号とともに報告すること。1 つでも欠陥があれば — REJECT; 修正するまで納品しないこと。

**人間によるプレビュー（任意）。** ユーザーにデックを視覚的にプレビューしてもらいたい場合は、`officecli watch "$FILE"` を実行してライブプレビューを提供し、ユーザー自身の判断で開いてもらうか、`.pptx` を PowerPoint / WPS / Keynote で直接開いてもらうこと。最終的な視覚確認には、対象のプレゼンテーションビューアでファイルを開くこと。

### Gate 6 — ピッチナラティブの健全性チェック（実行可能）

VC のレッドフラグをデックから grep するピッチ特有のチェック。いずれもトークンチェックであり、フルカバレッジのために Gate 5b の人間による確認と組み合わせること。

```bash
FILE="deck.pptx"

# 6.1 — no TBD / lorem / placeholder tokens (stronger than Gate 2 — pitch-specific scope)
LEAK=$(officecli view "$FILE" text | grep -niE 'TBD|lorem|\(fill in\)|xxxx|coming soon|placeholder')
[ -z "$LEAK" ] && echo "Gate 6.1 OK (no placeholder tokens)" || { echo "REJECT Gate 6.1:"; echo "$LEAK"; exit 1; }

# 6.2 — TAM / SAM / SOM presence (Series A+)
TAM_HIT=$(officecli query "$FILE" 'shape:contains("TAM")' --json | jq '.data.results | length')
[ "$TAM_HIT" -ge 1 ] && echo "Gate 6.2 OK (TAM slide present)" || echo "WARN Gate 6.2: no TAM mention — confirm stage is Seed / Bridge if intentional"

# 6.3 — Unit econ presence (Series B+): CAC OR LTV OR payback
CAC_HIT=$(officecli query "$FILE" 'shape:contains("CAC")' --json | jq '.data.results | length')
LTV_HIT=$(officecli query "$FILE" 'shape:contains("LTV")' --json | jq '.data.results | length')
if [ "$CAC_HIT" -ge 1 ] || [ "$LTV_HIT" -ge 1 ]; then
  echo "Gate 6.3 OK (unit econ surface)"
else
  echo "WARN Gate 6.3: no CAC / LTV — confirm stage Seed/A if intentional, REJECT if Series B+"
fi

# 6.4 — Use of Funds present on ask slide
UOF_HIT=$(officecli query "$FILE" 'shape:contains("Use of Funds")' --json | jq '.data.results | length')
[ "$UOF_HIT" -ge 1 ] && echo "Gate 6.4 OK (Use of Funds)" || { echo "REJECT Gate 6.4: ask slide missing Use of Funds"; exit 1; }

# 6.5 — Team prior-company signal (at least one of ex- / former / prior / previously)
PRIOR_HIT=$(officecli view "$FILE" text | grep -ciE '\b(ex-|former|prior|previously)\b')
[ "$PRIOR_HIT" -ge 1 ] && echo "Gate 6.5 OK (team prior-company)" || { echo "REJECT Gate 6.5: team slide has no prior-company credentials"; exit 1; }

# 6.6 — Traction chart y-axis anchored at 0 (at least one chart must set axismin=0, Series A+)
AXISMIN_HIT=$(officecli query "$FILE" 'chart' --json | jq '[.data.results[]? | select(.format.axisMin == "0" or .format.axisMin == 0 or .format.axismin == "0" or .format.axismin == 0)] | length')
[ "$AXISMIN_HIT" -ge 1 ] && echo "Gate 6.6 OK (traction chart axisMin=0)" || echo "WARN Gate 6.6: no chart sets axisMin=0 — confirm no ARR/revenue line chart, or add --prop axismin=0"

echo "Delivery Gate 6 PASS (token + narrative checks) — proceed to Gate 5b fresh-eyes (MANDATORY)"
```

**読み戻しに関する注意。** CLI は入力（`--prop axismin=0`）として小文字の `axismin` を受け付けるが、`query --json` の読み戻しではキャメルケースの `axisMin` を返す。上記の jq は将来互換性のため両方を受け付ける。

Gate 6 は grep の floor である。Gate 5b は視覚の ceiling である。両方が PASS を出力したときのみ出荷すること。

### 正直な限界

`validate` はスキーマエラーを捕捉するが、資金調達上の誤りは捕捉しない。デックは、$10M 市場に対する `$500B TAM`、過去の所属企業のない 4 人の共同創業者のチームスライド、80% 起点のホッケースティック y 軸、ユニットエコノミクスを欠く Series B ラウンドのピッチ、「多少の資金を調達しています」という ask スライドを抱えたまま `validate` を通過し得る。上記の Gates 5b + 6 は、`validate` がこれらのいずれも捕捉できないために存在する。

## 既知の問題と落とし穴

→ ベースの落とし穴（シェルエスケープ、resident における `[last()]`、コネクタ `@name=` の拒否 C-P-6、画像 alt の 2 段階手順 C-P-7、アニメーション削除 C-P-4、チャート色の正規化 C-P-7）: pptx v2 §Known Issues & Pitfalls C-P-1..7 を参照。

ピッチ特有:

- **ステージの誤認。** CAC/LTV の数式が 6 ページも続く Series A デックは過剰包装。ユニットエコノミクスを欠く Series B デックは未完成。不明な場合は、構築前に §Stage diagnosis を読み直すこと。
- **ホッケースティックの y 軸。** 折れ線グラフの y 軸が 0 起点でない場合、VC は 2 秒以内にそれを視覚的な嘘として読み取る。ARR / revenue / growth のチャートには常に `--prop axismin=0` を付けること。Gate 6.6 がこれをチェックする。
- **チームスライド = ポートフォリオ。** {顔写真 + 名前 + 役職} だけを示すカードは VC の信頼性審査に落ちる。すべてのカードに、過去の所属企業または過去の実績の 1 行が必要。Gate 6.5 がこれをチェックする。
- **手法のない TAM。** 「トップダウン」や「ボトムアップ」の出典脚注がない主張は捏造とみなされる。1 デックにつき 1 つの手法を選び、混在させないこと。
- **Use-of-Funds が 3 分割や 5 分割になっている。** 4 分割（Eng / GTM / G&A / Reserve）が規約であり、それから外れると雑に見える。Gate 6.4 が存在をチェックする。
- **ピッチデックを取締役会レビュー / セールスデックに流用する。** ナラティブアーク（problem → ask）は取締役会レビューにはそぐわない — 代わりに pptx v2 Recipe (d) の 10 枚構成へルーティングすること。上記の §Reverse handoff を参照。
- **pptx v2 Recipe (d′) の 20 枚構成は出発点であり、公式ではない。** これはステージ非依存の SaaS 向けである。あなたのステージ + 業界に応じて §Stage diagnosis と §業界別アーク・テンプレート で調整すること — 非 SaaS の Series A に対して (d′) を無調整のまま出荷しないこと。

## ヘルプの参照先

迷ったら: `officecli help pptx`、`officecli help pptx <element>`、`officecli help pptx <element> --json`。ヘルプが権威あるスキーマであり、このスキルは pptx v2 の上に乗る資金調達の差分に関する判断ガイドである。
</content>
