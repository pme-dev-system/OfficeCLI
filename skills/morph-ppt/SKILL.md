---
name: morph-ppt
description: "スムーズなスライド間アニメーションを持つ .pptx を作成したいとき — PowerPoint Morph トランジション、Keynote-style transition のような連続したモーション、スライドが進むにつれて成長 / 移動 / 回転する図形など — に使用するスキル。トリガー: 'morph', 'morph transition', 'smooth transition', 'continuous animation across slides', 'Keynote-style transition', 'animated slide sequence', 'shape continuity across slides'。出力は単一の .pptx。このスキルは officecli-pptx の上位に位置するシーンレイヤーであり、pptx v2 の全ルール（visual floor、グリッド、パレット、コネクタ規範、Delivery Gate 1–5a）を継承する。スライド間モーションを伴わない汎用デッキ、ピッチデック、役員レビュー用途には呼び出さないこと — それらは officecli-pptx ベースまたは officecli-pitch-deck にルーティングすること。"
---

# OfficeCLI Morph-PPT スキル

**このスキルは `officecli-pptx` の上位に位置するシーンレイヤーです。** pptx のハードルール — visual delivery floor（タイトル ≥ 36pt / 本文 ≥ 18pt / タイトル ≥ 本文の2倍）、33.87×19.05cm 上の 12 カラムグリッド、正規パレット、チャート選択の判断テーブル、コネクタ規範、シェルエスケープ、レジデント + バッチ、Delivery Gate 1–5a — はすべて継承されており、ここで再度教えることはしません。このファイルが追加するのは、**Morph** が上乗せで必要とするものだけです: スライド間の図形名バインディング、Scene Actor とコンテンツのプレフィックスの区別、ゴースト（幽霊化）の作法、`transition=morph` の CLI 特有の癖、52 種のビジュアルスタイルライブラリの参照方法、そして Morph 専用の fresh-eyes Gate 5b 拡張です。

pptx ベースルールで既にカバーされている箇所は、本文中で `→ see pptx v2 §X` と記します。まだ読んでいない場合は先に `skills/officecli-pptx/SKILL.md` を読んでください。

## セットアップ

`officecli` が未インストールの場合:

- **macOS / Linux**: `curl -fsSL https://d.officecli.ai/install.sh | bash`
- **Windows (PowerShell)**: `irm https://d.officecli.ai/install.ps1 | iex`

`officecli --version` で確認してください（PATH が反映されていない場合は新しいターミナルを開いてください）。インストールに失敗した場合は https://github.com/iOfficeAI/OfficeCLI/releases からバイナリをダウンロードしてください。

## ⚠️ ヘルプ優先ルール

**このスキルが教えるのは Morph のワークフロー — 図形名をいつ一致させるか、いつゴースト化するか、CLI がいつ自動プレフィックスを付けるか — であり、全コマンドフラグではありません。** プロパティ名、列挙値、プリセットが不明な場合は、推測する前にヘルプを確認してください。

```bash
officecli help pptx slide           # 権威ある情報源: transition, advanceTime, advanceClick, background
officecli help pptx transition      # transition / transitionDuration / transitionSpeed (Parent: slide)
officecli help pptx shape           # name, preset, x/y/width/height, fill, rotation, opacity, animation
officecli help pptx animation       # preset + trigger + duration の値
officecli help pptx <element> --json  # 機械可読なスキーマ
```

ヘルプはインストール済みの CLI バージョンを反映します。スキルとヘルプが食い違う場合は**ヘルプが優先**します。本ファイル中のすべての `--prop X=` は `officecli help pptx <element>` に対して grep で検証済みです。具体的な確認事項: `transition=morph` は `slide` 上の列挙値として存在します。`advanceTime` / `advanceClick` は有効です。`transition` は実在する要素（`officecli help pptx transition`、Parent: slide、set/get）で、`transition`、`transitionDuration`、`transitionSpeed` を公開します。トランジションの設定は高レベルパスの `set <slide> --prop transition=morph` で行い、速度・時間の調整は複合ショートハンド `transition=morph-slow`（または `-fast`、あるいは `transition=morph-<DUR_MS>`）で行います。速度・時間は `transition` プロパティ上のこのショートハンドを通じてのみ設定でき、独立したサブプロパティとしては設定できません。両方とも読み戻し時にラウンドトリップします: `transition=morph-slow`/`-fast` は `transitionSpeed=slow`/`fast` として読み戻され、`transition=morph-<DUR_MS>`（例: `morph-1500`）は `transitionDuration=1500` として読み戻されます。

## メンタルモデルと継承

**pptx v2 を継承します。** まず `skills/officecli-pptx/SKILL.md` を読んでいる前提です。このスキルは、以下の内容を既に理解しているものとして扱います: スライド + 図形 + チャート + コネクタの追加; `@name=` / `@id=` によるアドレッシング; パスのクォート; `batch` のヒアドキュメント利用; フローコネクタでの `tailEnd=triangle`; Delivery Gate 1–5a の実行; `[AGENT-ERROR]` / `[RENDERER-BUG]` / `[SKILL gap]` の切り分け。これらのいずれかに不慣れであれば、先に pptx v2 を読んでください。

**pptx v2 から継承（再教育しない）:**

- Visual delivery floor — タイトル ≥ 36pt / 本文 ≥ 18pt / タイトル ≥ 本文の2倍、カバーの豊かさ、コントラストフロア、`\$\t\n` リテラル禁止、1スライドあたりアニメーション ≤ 1件 / ≤ 600ms。
- グリッド計算 — 33.87 × 19.05cm、端の余白 ≥ 1.27cm、ブロック間ギャップ ≥ 0.76cm、余白（ネガティブスペース）≥ 20%。N カードグリッドの場合: `col = (33.87 − 2·margin − (N−1)·gap) / N`。
- 4 つの正規パレット（Executive navy / Forest & moss / Warm terracotta / Charcoal minimal）— morph デッキは `reference/styles/` から別のムードを選んでもよいが、コントラストルールは適用され続ける。
- チャート選択テーブル — 列 vs 棒 vs 折れ線 vs 円 vs 散布図 vs 大文字 KPI; `系列 > 3 かつカテゴリ > 8` の場合は分割。
- コネクタ規範 — `shape=straight|elbow|curve`、from/to は `@id=`（C-P-6）、すべてのフローに `tailEnd=triangle`。
- シェルエスケープの3層構造 — `$` はシングルクォート、バッチはヒアドキュメント、実際の改行は `<a:br/>`。
- レジデントモード + バッチ ≤ 12 操作、`<<'EOF'` のシングルクォートデリミタ。
- Delivery Gate 1-5a（スキーマ、トークン grep、ハイパーリンク rPr、スライド順序、dark-on-dark）— どのゲートも「完了」を宣言する前に OK を出力すること。
- 既知の問題 C-P-1..7（ハイパーリンク rPr、チャート spPr 警告、アニメーション時間の読み戻し、アニメーション削除、コネクタの列挙値、コネクタの `@name=`、チャート色のレンダラー正規化）。
- 帰属トリアージ — `[AGENT-ERROR]` vs `[RENDERER-BUG]` vs `[SKILL gap]`。

**Morph が担うアイデンティティ — このスキル固有部分（pptx v2 の上に乗る差分）:**

- **スライド間の図形名バインディング。** PowerPoint の Morph エンジンは、隣接するスライド間で**完全に一致する `name=`** によって図形をペアリングし、その位置・サイズ・回転・塗りつぶし・不透明度を補間します。名前が一致しなければアニメーションは発生せず、サイレントフェードになります。これはワークフロー上の規律であり、CLI 機能ではありません。
- **名前空間プレフィックス:** `!!scene-*`（永続的な装飾、決してゴースト化しない）/ `!!actor-*`（進化してから退場するコンテンツ）/ `#sN-*`（スライドごとのコンテンツ、N+1 スライド目でゴースト化）。名前は `add` する**前に**計画してください。
- **ゴースト位置 `x=36cm`**（33.87cm のキャンバスの右端の外側）。`!!` プレフィックス付きの図形を削除してはいけません — キャンバス外に移動させることで、Morph の退場アニメーションが再生され続けます。
- **`transition=morph` の自動プレフィックスの癖。** CLI は morph スライド上のすべての図形に自動で `!!` を前置します（`#s1-title` は `!!#s1-title` として保存されます）。`@name=` パスセレクタは**それでも解決可能**です — `get .../shape[@name=#s1-title]` はその図形を返します（マッチングはサフィックス/プレフィックス寛容です）。読み戻される名前はプレフィックス付きの形式になります。§既知の問題を参照。
- **隣接スライドの空間的多様性。** 変位 ≥ 5cm または回転 ≥ 15° がペア間になければ、Morph は目に見える補間を何も行いません。
- **レンダラーの現実。** Morph は PowerPoint 365 / Keynote / WPS でレンダリングされます。LibreOffice や多くの Web ビューアではプレーンフェードとしてレンダリングされます（ランタイム機能）。これはスキルの欠陥ではなく `[RENDERER-BUG]` です。

### 逆ハンドオフ — いつ pptx ベース（または姉妹スキル）に戻るか

スライド間モーションを伴わないデッキ（役員レビュー、営業資料、全社会議、トレーニング）は **pptx v2 ベース**にとどまってください。Morph を伴わない資金調達のナラティブアークは **officecli-pitch-deck** にとどまってください。ユーザーが明示的に "morph" / "smooth transitions" / "continuous animation" を求め、かつ連続する 2 枚以上のスライドが変形する視覚要素を共有している場合にのみ、このスキルを使用してください。「アニメーション付きデッキ」が単発の入場アニメーションを意味する場合は → pptx v2 §Animations であり、morph ではありません。

## シェル & 実行の規律

**シェルクォート、段階的実行、`$FILE` 慣習** → pptx v2 §Shell & Execution Discipline を参照。ルールは全く同じです。

**Morph 固有の追加事項:**

- **シェル値中の `!!` — シングルクォートで囲む。** Bash / zsh の history expansion がクォートなしの `!!foo` を食ってしまいます。常に `--prop 'name=!!scene-ring'`（シングルクォート）を使ってください。Python の `subprocess.run([...])` リストではクォート不要 — `"name=!!scene-ring"` をそのまま文字列として渡せます。
- **prop テキスト中の `$` — シングルクォートで囲む（価格トークン）。** `--prop text='$9/mo'` や `--prop text='$199/yr'` — `--prop text="$9/mo"` は**絶対に禁止**（zsh/bash が `$9` を空変数として食い、テキストが `.` や末尾のピリオドとしてレンダリングされてしまいます）。ダブルクォート prop 中の `${VAR}`、`$USER`、`\n`、`\r`、`\t` も同様です。下記の Gate 2 morph 補則がこの漏れのシグネチャを grep します。
- **シェル値中の `#` — 安全だが念のためクォート。** `#` はシェル単語の先頭でのみコメントリーダーになります。`--prop name=#s1-title` は動作しますが、`--prop 'name=#s1-title'` を習慣にしておけば推測せずに済みます。
- **複数図形スライドではバッチのヒアドキュメントが最もクリーンな方法。** `<<'EOF' | officecli batch $FILE` はシェル展開をすべて無効化します — JSON 本文中の `$`、`!!`、`#`、`'` に対して安全です。
- **`--json` レスポンスはペイロードを `.data.results[]` でラップします。** `query` と `get` はいずれも `.data.results[]` 配列を返します。単一ノードの `format` は `.data.results[0].format.X` に、そのノードの子は `.data.results[0].children[]`（各子の format は `.data.results[0].children[].format.X`）にあります。常に `.data.results[0]` を経由してください — 素の `.data.children[]` や `.data.format` は null をサイレントに返します。
- **変数:** すべてのビルドスクリプトの先頭で `FILE="deck.pptx"` を定義し、以下の全例で `$FILE` を使用します。
- **ゲートのシェルパターン — カウントしてから if/else。** `grep … && echo LEAK || echo OK` は書かないでください — grep がゼロマッチで exit 1 になると、`||` の分岐が空 stdout のまま発火し「OK」を紛らわしく表示してしまいます（あるいは前段のパイプから「LEAK」を表示してしまいます）。正規形はこちら: `COUNT=$(cmd | wc -l); if [ "$COUNT" -gt 0 ]; then echo "LEAK: …"; else echo "OK"; fi`。

## このスキルが担う2つの基本要素

- **Scene Actor** = `!!` 名を持つ永続的な図形（装飾またはコンテンツ）で、隣接するスライド間で**同一名でペアリング**され、Morph が補間できるようにしたもの。すべての `!!scene-*` / `!!actor-*` 図形は Scene Actor です。
- **Choreography（振り付け）** = アクターがどう進化していくかの計画 — 誰がどこへ移動し、誰が入場し、誰がどのスライドペアで退場するか。コードを書く**前**に §Morph Pair Planning のテーブルに記述します。

ユーザーが morph によるモーションを求め、かつ連続する 2 枚以上のスライドが変形する視覚要素を共有している場合にこのスキルを使用してください。対象ビューアに関する注意点: morph には PowerPoint 365 / Keynote / WPS が必要です — ユーザーが LibreOffice のみを使う場合は先に警告してください（§レンダラーの正直さ を参照）。

**スピーカーノートのルール。** すべてのコンテンツスライド（表紙・締めくくり以外）は `officecli add "$FILE" /slide[N] --type notes --prop text='…'` によるスピーカーノートを**必ず**持たなければなりません。ノートの欠落は納品不可 — pptx v2 §Hard rules (H7) を継承します。Morph デッキは視覚的にミニマルになりがちなので、ノートがナレーションを担います。

## Morph とは何か？（コアメカニクス）

PowerPoint の Morph トランジションは、**同一の図形名**によってマッチングされた隣接スライド間で図形プロパティを補間することで、スムーズなモーションを生み出します。

```
Slide 1: shape name="!!scene-ring" x=5cm  width=8cm   fill=E94560 opacity=0.3
Slide 2: shape name="!!scene-ring" x=20cm width=12cm fill=E94560 opacity=0.6
         ↓  transition=morph on slide 2
Result:  Ring smoothly moves, grows, and fades darker over ~1 second
```

Morph はスライド N+1 が `transition=morph` を持っている場合にのみ動作します。作成時に `officecli add / --type slide --prop transition=morph` を通じて適用するか、後から `officecli set "/slide[N]" --prop transition=morph` で適用してください。このプロパティを省略したスライド 2 以降は、マスターが定義する内容（通常はトランジションなし）にフォールバックし、モーションはサイレントに失われます。

**3プレフィックス命名システム（交渉の余地なし）:**

| プレフィックス | 役割 | ライフサイクル | 例 |
|---|---|---|---|
| `!!scene-*` | 背景 / 装飾 — デッキ全体を通じて持続 | 一度設定し、位置/サイズを調整してモーションを作る; **ゴースト化はまれ** | `!!scene-ring`, `!!scene-bg-band`, `!!scene-grid` |
| `!!actor-*` | コンテンツ / 前景 — セクションを通じて進化 | スライド N で導入され、N+1, N+2… で変更され、退場スライドで **`x=36cm` にゴースト化** | `!!actor-feature-box`, `!!actor-metric`, `!!actor-headline` |
| `#sN-*` | スライドごとのコンテンツ（タイトル、箇条書き、キャプション） | スライド N で新規追加され、スライド N+1 で **`x=36cm` にゴースト化** | `#s1-title`, `#s2-kpi`, `#s3-caption` |

**ハードルール:** `!!scene-*` と `!!actor-*` の名前は**絶対に**衝突してはいけません（例: 同じデッキ内の `!!scene-card` + `!!actor-card` — morph エンジンが混同します）。曖昧さを解消してください: `!!scene-card-bg` vs `!!actor-card-content`。

**チャートも morph でペアリング可能です。** `officecli add … --type chart` は `--prop name=!!…`（名前は読み戻し可能）を受け付けるため、隣接スライドで同一の `!!` 名を持つチャートは図形名 morph ペアリングに参加し、チャートフレームの位置/サイズが補間されます。ただし morph はチャートフレーム内の*プロットされたデータ*自体は補間できません。棒グラフや折れ線グラフが伸びていくようなナラティブで、棒そのものをアニメーションさせたい場合: (a) チャートをそのままの状態でプレーンなフェードインとして受け入れるか、(b) 値に合わせて手動でサイズ調整した N 個の `!!actor-bar-K` 矩形を作り、それらを morph させてください — 各矩形は隣接スライドで同じ `!!actor-bar-K` 名を保ちつつ、幅/高さ/塗りつぶしを進化させます。

**ゴーストの蓄積はサイレントです。** `!!` プレフィックス付きの図形がいずれかのスライドに一度現れると、明示的に `x=36cm` へ移動させない限り、以降のすべての morph スライドで表示され続けます。`final-check` ヘルパーは可視領域に残留する `!!` 図形を検出**しません** — **Gate 5b のスクリーンショット監査のみ**が検出します。コーディングの**前**に、ペアテーブルで各アクターの退場スライドを計画してください。

**空間的多様性のルール。** 隣接するスライドは**明らかに異なる**構図を持たなければなりません — 少なくとも 3 つの morph ペアリング済み図形について、変位 ≥ 5cm または回転 ≥ 15° またはサイズ変化 ≥ 30% が必要です。これがなければ morph は目に見える補間を何も行わず、トランジションはフェード（サイレント失敗）に落ち込みます。

**同時タイミングの制約。** 1つの morph ペア内のすべての `!!` 図形は同時にアニメーションします。図形 A を図形 B より先に動かして時間差をつけたい場合は、中間のキーフレームスライドを挿入してください — 図形ごとの遅延ノブは存在しません。

**ペア vs 入場 vs 退場 — 3つの挙動、1つのルール。** 同一のメカニズム（図形名の一致）が3種類の結果を生みます:

| 挙動 | 元スライド A | 対象スライド B | `!!` を持つのは誰か |
|---|---|---|---|
| **ペアリング morph**（補間） | `!!foo` を持つ | `!!foo` を持つ | 両スライド、同一名 |
| **入場**（フェード / morph-in） | — （対応物なし） | `!!foo` を持つ | 対象のみ — 新規図形 |
| **ゴースト経由の退場**（スライドアウト） | 可視 `x` に `!!foo` を持つ | `x=36cm` に `!!foo` を持つ | 両方 — 同一名、B はキャンバス外 |

**`!!` プレフィックスとゴースト化の対象は退場する（入場ではない）コンテンツです。** `!!actor-*` 図形は、忘れると静かに「消えて」しまいます — スライド B でその名前が失われると、対応のない退場（プレーンフェード）として解釈されます。常に明示的に `x=36cm` へゴースト化することで、退場アニメーションが右端へ視覚的にスライドオフするようにしてください。実行可能な一例:

```bash
# スライド2: アクターは x=5cm に表示 — スライド3: 同じ名前をキャンバス外にゴースト化 → 視覚的なスライドオフモーション
officecli add "$FILE" "/slide[3]" --type shape --prop 'name=!!actor-metric' \
  --prop text="42%" --prop x=36cm --prop y=8cm --prop width=6cm --prop height=3cm
```

**コンテンツ（`#sN-*`）はスライドごとに新規追加されます。** テキストが毎スライド変化するため、Morph はタイトル/本文に対して意味のあるペアリングを行いません — 代わりにクロスフェードします。これが `#sN-*` がスライドごとに異なる名前を持つ理由です（意図的に非ペアリングです）。スライド N+1 でゴースト化する必要があります。Scene Actor（`!!`）が連続性を担い、コンテンツ（`#`）がメッセージを担います。

## Morph ペア計画（コーディング前、必須）

morph ペアを計画する前に、デッキの想定読者/目的/ナラティブが未定義であれば、`reference/decision-rules.md` の計画プロンプトを実行して `brief.md` を先に出力してください — ナラティブの背骨を持たない morph アークは「モーション付きスライド」に収縮してしまい、「モーション付きストーリー」にはなりません。

`officecli add` を書く**前**に、`brief.md` 内のテーブルですべてのトランジションを計画してください。ビルド途中での図形リネームは、ゴースト蓄積バグの原因ナンバーワンです。

| ペア | スライド A（開始） | スライド B（終了） | 関与するアクター | スライド B でのゴースト |
|---|---|---|---|---|
| 1→2 | `!!scene-ring` が中央 5cm、`#s1-title` が表示 | Ring が x=20cm に移動、8→12cm に成長; `#s2-subtitle` が現れる | `!!scene-ring` が進化 | `#s1-title` → x=36cm |
| 2→3 | `!!actor-feature-box` が大きい（幅14cm） | Feature box が小さくなり（6cm）、`!!actor-metric` が入場 | `!!scene-ring`, `!!actor-feature-box`, `!!actor-metric` | `#s2-subtitle` → x=36cm |
| 3→4 | コンテンツセクション A | セクション B のディバイダー | — | `!!actor-feature-box` + `!!actor-metric` → x=36cm（セクション退場）; `#s3-*` → x=36cm |

**計画ルール:**

1. すべての `!!` 名を事前に決定してください — morph ペアリングされる各図形は両スライドで**完全に同一の名前**を使う必要があります。
2. すべての `!!` 図形を `!!scene-*` または `!!actor-*` に分類してください。Scene 図形は持続し、Actor には計画された退場スライドが必要です。
3. **セクション境界の遷移:** 新しいトピックセクションに移る際は、そのセクションの最初のスライドで前セクションの `!!actor-*` をすべてゴースト化してください。残るのは `!!scene-*`（デッキ全体の装飾）のみです。
4. テーブルが完成するまでビルドを開始しないでください。ビルド途中で計画が変わった場合は、テーブルを描き直し、影響を受けるスライドを再検証してください。

## Morph レシピ（4パターン）

4つのパターンで morph デッキの約95%をカバーします。`$FILE="deck.pptx"` を通して使用します。各ブロックは自己完結型で20行以下です。

### (a) 単一要素 morph — サイズ / 位置

**視覚的な結果。** スライド1の中央にヒーロータイトル（48pt、y=8cm）を配置し、スライド2でそれを32ptに縮小して左上隅（x=1.5cm, y=1cm）に移動 — スライド2の新しいコンテンツが中央の主役を務められるようにします。1つの図形、クリーンなモーション、アクターなし。

```bash
FILE="deck.pptx"
officecli create "$FILE"; officecli open "$FILE"

# スライド1 — ヒーロー
officecli add "$FILE" / --type slide --prop layout=blank --prop background=1E2761
officecli add "$FILE" /slide[1] --type shape --prop 'name=!!actor-headline' \
  --prop text="The one idea" --prop x=4cm --prop y=8cm --prop width=26cm --prop height=3cm \
  --prop font=Georgia --prop size=48 --prop bold=true --prop color=FFFFFF --prop align=center --prop fill=none

# スライド2 — headline が縮小 + 移動; 新しい本文が主役に
officecli add "$FILE" / --type slide --prop layout=blank --prop background=1E2761 --prop transition=morph
officecli add "$FILE" /slide[2] --type shape --prop 'name=!!actor-headline' \
  --prop text="The one idea" --prop x=1.5cm --prop y=1cm --prop width=12cm --prop height=1.5cm \
  --prop font=Georgia --prop size=24 --prop bold=true --prop color=FFFFFF --prop align=left --prop fill=none
officecli add "$FILE" /slide[2] --type shape --prop 'name=#s2-body' \
  --prop text="Here is the supporting evidence." --prop x=1.5cm --prop y=5cm --prop width=30cm --prop height=2cm \
  --prop font=Calibri --prop size=20 --prop color=CADCFC --prop fill=none

officecli close "$FILE"; officecli validate "$FILE"
```

### (b) 複数要素の協調 morph — Actor / Choreography

**視覚的な結果。** 3つの Scene Actor（`!!scene-ring`, `!!scene-dot`, `!!scene-band`）を3スライドにわたって再配置し、カメラパンのような感覚を作ります。スライドごとの新しいタイトルは `#sN-*` ゴーストパターンでフェードイン/アウトします。継続的な視覚的背景を持つナラティブでこれを使ってください。

```bash
# スライド1 — アンカー構図（既にレシピaで構築済み; ここでアクターを追加）
officecli add "$FILE" /slide[1] --type shape --prop 'name=!!scene-ring' --prop preset=ellipse \
  --prop fill=E94560 --prop opacity=0.3 --prop x=5cm --prop y=3cm --prop width=8cm --prop height=8cm
officecli add "$FILE" /slide[1] --type shape --prop 'name=!!scene-dot' --prop preset=ellipse \
  --prop fill=0F3460 --prop x=28cm --prop y=15cm --prop width=1cm --prop height=1cm

# スライド2 — morph: ring が移動 + 成長、dot が左にスライド（両方とも空間的多様性 ≥ 5cm）
officecli set "$FILE" "/slide[2]" --prop transition=morph
officecli add "$FILE" /slide[2] --type shape --prop 'name=!!scene-ring' --prop preset=ellipse \
  --prop fill=E94560 --prop opacity=0.6 --prop x=20cm --prop y=2cm --prop width=12cm --prop height=12cm
officecli add "$FILE" /slide[2] --type shape --prop 'name=!!scene-dot' --prop preset=ellipse \
  --prop fill=0F3460 --prop x=3cm --prop y=16cm --prop width=1.5cm --prop height=1.5cm
# スライド1のコンテンツをゴースト化（morph 後もパスは解決可能 — 既知の問題を参照）
officecli set "$FILE" "/slide[2]/shape[@name=#s1-title]" --prop x=36cm 2>/dev/null || true

# morph ペアを検証: スライド1と2で同一の名前
officecli get "$FILE" /slide[1] --depth 1 --json | jq -r '.data.results[0].children[]?.format.name // empty'
officecli get "$FILE" /slide[2] --depth 1 --json | jq -r '.data.results[0].children[]?.format.name // empty'
# 比較 — `!!scene-ring` と `!!scene-dot` は両方に、バイト単位で同一で現れなければならない。
# 注: morph は名前に `!!` プレフィックスを付けて保存する; プレフィックス付きの形式同士を比較すること。
```

### (c) 連続する複数スライド morph（ストーリーアーク）— ヘルパーを使用

**視覚的な結果。** 1つの連続したストーリーを語る5スライドのアーク: 同じ2つの Scene Actor がナラティブの進行とともにキャンバス上を漂い、コンテンツ（`#sN-*`）はスライドごとに更新され次のスライドでゴースト化されます。これを手作業で構築すると約60コマンドになります — `reference/morph-helpers.py` を使ってビルドスクリプトを短く自動検証可能にしてください。

```python
#!/usr/bin/env python3
# clone + ghost + verify を行う提供済みヘルパーライブラリを呼び出す
import subprocess, sys, os
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
HELPERS = os.path.join(SCRIPT_DIR, "reference", "morph-helpers.py")
FILE = "deck.pptx"

def helper(*args):
    subprocess.run([sys.executable, HELPERS, *[str(a) for a in args]], check=True)

# ... スライド1は2つの Scene Actor（!!scene-ring, !!scene-dot）+ #s1-title で構築済みとする
# ヘルパーがスライド2〜5を構築: 前スライドからの clone + transition=morph の適用 + 前スライドの #sN- コンテンツのゴースト化
# `clone` はクローンされたスライドの図形リストを出力する — それを読んで、
# 前スライドの #s(n-1)- コンテンツを担う図形インデックスを選び、それらの明示的なインデックスを `ghost` に渡す。
for n in range(2, 6):
    helper("clone", FILE, n - 1, n)          # clone + transition=morph の設定 + 図形リスト表示（#s(n-1)- のインデックスに注目）
    helper("ghost", FILE, n, 1, 2)           # #s(n-1)- のコンテンツ図形をインデックスでゴースト化（ここでは図形1と2）
    # …続けてこのスライドの #sN- コンテンツを通常通り officecli add で追加…
helper("final-check", FILE)                   # 構造的なパス; 可視領域に残留する !! は検出しない
```

ヘルパーのシグネチャとソース: `reference/morph-helpers.py`（`clone`, `ghost`, `verify`, `final-check`）。シェル版は `reference/morph-helpers.sh` — プラットフォームごとにどちらか一方を選び、混在させないでください。

**ヘルパー vs 生の `officecli` の使い分け。** 2〜3スライドのデッキでは、生のコマンド（レシピ a, b）の方が明快です。clone/ghost/verify のサイクルが繰り返される5スライド以上では、ヘルパーがコマンド数を約40%削減し、組み込みの検証を提供します。どちらの場合も、納品前に必ず `officecli validate` で全スライドをクローズしてください。

### (d) Morph + フェードのハイブリッド — morph スライド上での入場

**視覚的な結果。** `!!scene-ring` が連続的に移動する morph ペアと同時に、新しいスライドごとのカードがフェードインする例。morph ペアの背景が視線を運び、新しい前景コンテンツが生の出現よりも柔らかい入場を必要とする場合に使用します。

```bash
# スライド2は既に transition=morph と !!scene-ring を持っている。フェード入場する新しいカードを追加。
officecli add "$FILE" /slide[2] --type shape --prop 'name=#s2-card' --prop preset=roundRect \
  --prop fill=F5F7FA --prop line=none --prop x=2cm --prop y=12cm --prop width=10cm --prop height=5cm

# 新しいカードに morph と同時のフェード入場を適用。
# 'fade-entrance-300-with' = フェードイン、300ms、trigger=withPrevious（morph トランジションと同時再生）。
officecli set "$FILE" "/slide[2]/shape[@name=#s2-card]" --prop animation=fade-entrance-300-with
officecli get "$FILE" "/slide[2]/shape[@name=#s2-card]" --json | jq '.data.results[0].format.animation'  # 読み戻しの健全性チェック — トリガーのサフィックスは落ち、"fade-entrance-300" として読み戻される
```

**なぜこれで動くのか。** Morph は `!!scene-*` 図形のみをアニメーションさせます（スライド1にペアがあるため）。新しい `#s2-card` にはスライド1の対応物がないため、morph はデフォルトでフェードさせようとします — `fade-entrance-300-with` はそのフェードを明示的かつタイミング付きにします。アニメーションは pptx v2 の floor に従ってください: ≤ 600ms、バウンス/スウィベル/端からのフライインなし（正規のプリセット一覧は `officecli help pptx animation` を参照）。

## Choreography — アニメーションの種類とスタッガードタイミング

Morph が複数の図形をどうアニメーションさせるかが、観客に見える内容を決定します。各ペアに適したメカニズムを選んでください:

| アニメーションの種類 | 達成方法（スライドAとスライドBの間） |
|---|---|
| 単純な移動 | 両スライドで同じ `!!` 名、同じサイズ、異なる `x`/`y` — morph が位置を補間 |
| スケール変換 | 同じ名前、異なる `width`/`height` — morph がサイズを補間（中心も再配置） |
| 移動 + スケール | `x`, `y`, `width`, `height` を同時に変更 — morph がすべての次元を一度に処理 |
| 色 / 不透明度の変化 | 同じ名前、異なる `fill` または `opacity` — morph が塗りつぶしをクロスフェード |
| 回転 | 同じ名前、異なる `rotation`（度） — morph が最短弧に沿って回転 |
| フォントサイズの変化 | テキスト図形で同じ名前、異なる `size`（pt） — PowerPoint 365 では補間されるが、Keynote / WPS / LibreOffice では信頼性が低い（クロスフェードに劣化する場合がある）。可搬性のあるモーションのためには、`size` の変化に対応する `width`/`height` の差分か `x`/`y` の変位を組み合わせてください — サイズ補間が効かない場合でも空間的な変化がモーションを可視に保ちます |
| 入場（フェードイン） | 図形がスライドBにのみ存在（Aに対応物なし） — morph がフェードイン |
| 退場（フェードアウト） | 図形がスライドAにのみ存在（Bに対応物なし） — morph がフェードアウト |

**複数図形のタイミング制約。** 1つの morph ペア内のすべての `!!` 図形は**同時に**アニメーションします — CLI には図形ごとの遅延/時間ノブは存在しません（ヘルプで確認済み: スライドに `morph.duration` / `morph.delay` はありません）。図形Aを図形Bより先に動かして時間差をつけたい場合は、中間スライドを挟んで**トランジションを2つのペアに分割**してください:

```
Slide 2 → Slide 3:  !!actor-A moves (!!actor-B stays put)
Slide 3 → Slide 4:  !!actor-B moves (!!actor-A stays put or ghosts)
```

スライド3は明示的な中間キーフレームです。図形の `animation=` プロパティのタイミング設定でスタッガーを偽装しようとしないでください — Morph は図形ごとのアニメーションより先に実行されます。

**十分な多様性のヒューリスティック（ベストプラクティス — 創造的な柔軟性）。** morph が「モーション」として読めるようにするには、支配的なペアリング図形について {x, y, width, height, rotation, fill, opacity} のうち少なくとも3つを変更し、変位 ≥ 5cm または回転 ≥ 15° またはサイズ変化 ≥ 30% を満たしてください。1図形 × 3プロパティは有効な創造的パターンです（1つのヒーロー要素に焦点を当てる場合）。

**Delivery Gate 5b-morph-2 はより厳格です。** このゲートは、ペアにわたって {x, y, width, height, rotation, フォントサイズ} のうち少なくとも1つが変化する、3つ以上の**異なる** `!!` プレフィックス付き図形をハードにアサートします — 「これは本当に morph なのか、それとも見せかけの morph なのか」の整合性チェックです。ヒューリスティックは創造的な意図に情報を与えますが、Gate が納品を決定します。**ブランドの定常的な装飾（固定されたヘッダーストリップ、フッターバー、ロゴバッジ）は3図形のノルマにカウントされません** — これらは動かないことになっているためで、モーションは他の3つの命名済み図形から生まれる必要があります。迷ったら、より厳格な Gate に従ってください。

**デッキ長のリズム。** すべてのトランジションを morph で埋めると、映画的ではなく落ち着きのない印象になります。デッキの長さに応じて morph の瞬間のペースを調整してください:
- **8-10スライド（密度が高い）:** 3-5 の morph の瞬間; モーションは集中してもよい。
- **12-18スライド（儀式的）:** 合計 3-5 の morph を、4-6 スライドごとに配置; セクションディバイダーで `transition=morph` を使い、アニメーションが継続的な動揺ではなく章の句読点として読めるようにする。
- **18+スライド（幕構成）:** 幕間に1つの長いセクションディバイダー morph（5-10秒の意図的なモーションと短い静止）を挟んだ3幕構成にし、各幕の中に2-3のより静かな morph を配置する。1スライドごとの `!!actor-*` の入れ替わりよりも `!!scene-*` の連続性に重きを置く。

## Scene Actor の空間ルール

Scene Actor とキャンバス上を移動するアクターは、morph 中は予測可能なゾーンにとどまらなければなりません — さもないとコンテンツと交差し、雑然として見えます。

**セーフゾーン（Scene Actor の静止位置と morph のパスに推奨）:**

```
Top-right corner:   x ≥ 24cm, y ≤ 6cm
Bottom-right:       x ≥ 24cm, y ≥ 12cm
Bottom-left:        x ≤ 2cm,  y ≥ 12cm
Off-canvas (ghost): x ≥ 33.87cm  (canvas right edge; use x=36cm for explicit ghost)
```

**アクターをコンテンツの核心部で静止させないでください:** `x = 2~28cm, y = 3~16cm`。アクターは morph 中にこの核心部を**通過**してもかまいません（それがモーションです）が、それ自身がコンテンツ（メッセージを担う `!!actor-*`）でない限り、高い不透明度でそこに留まったままスライドを終えるべきではありません。

**Scene Actor を配置する前に、既存の図形の境界を確認してください:**

```bash
officecli get "$FILE" "/slide[$N]" --depth 1 --json | \
  jq -r '.data.results[0].children[]? | "\(.format.name // .path)  x=\(.format.x) y=\(.format.y) w=\(.format.width) h=\(.format.height)"'
```

アクターの目標位置が `#sN-*` コンテンツ図形の境界ボックス（`x` から `x + width`、`y` から `y + height`）と重ならないことを確認してください。重なる場合は、アクターの `opacity` を ≤ 0.15 に下げるか、セーフゾーンへ移動してください。

## スタイルライブラリ参照ワークフロー

`reference/styles/` には52種のビジュアルスタイルディレクトリ（dark / light / warm / vivid / bw / mixed のムード）があります — デザインインスピレーションであって、テンプレートではありません。ライブラリはコンテンツダンプとしてではなく**オンデマンド参照**として使ってください。

**なぜコピーではなく参照なのか。** 52個の `build.sh` はそれぞれ完全なスタイルデモですが、座標はそのデモ固有のコンテンツ長に合わせて手動調整されています。異なる長さのコンテンツを持つデッキにそのままコピーすると、重なりやずれが発生します（`INDEX.md` L5-11 で指摘済み）。ライブラリの価値は**デザインロジック**にあります: ムードに対するパレット選択、シグネチャシェイプ、振り付けパターン。そのロジックを自分自身のグリッド計算に適用してください。

**4ステップの参照方法:**

1. **INDEX を閲覧。** `reference/styles/INDEX.md` は52スタイルをパレットカテゴリとムードでグループ化しています（例: `dark--premium-navy` = 権威的/洗練; `warm--earth-organic` = 有機的/落ち着き）。Quick Lookup テーブルには各スタイルの**主要な16進数トリオ**（bg / fg / accent）も表示されており、ユーザーがブランドカラーを指定した場合、すべての `style.md` を開かずに16進数の列から最も近い一致を見つけられます。トピックのムードに合致する、またはユーザー指定の16進数と整合する1つのスタイルを選んでください。
2. **哲学を読む。** `reference/styles/<style-id>/style.md` を開き、デザイン意図 — タイプの組み合わせ、色のロジック、シグネチャ要素 — を確認します。
3. **テクニックをざっと見る。** `reference/styles/<style-id>/build.sh` はテクニック参照（シグネチャシェイプ、パレットの16進数コード、振り付けのアイデア）のためだけに開いてください — **座標は `INDEX.md` L5-11 の記載通り既知のバグあり**です; コピーしないでください。
4. **自分自身のキャンバスに適用。** pptx v2 のグリッド計算と visual floor を使ってデッキを構築してください; パレットとシグネチャジェスチャーだけを借りてください。

**ポインタ:** `→ see reference/styles/<style-id>/` — スタイルの build.sh から座標をインラインコピーしないでください。

## Delivery Gate（pptx v2 + morph 追加を継承）

**Gate 1–5a: pptx v2 からの全面移植。** → pptx v2 §Delivery Gate を参照。スキーマ（C-P-2 チャート spPr のホワイトリスト化）、トークン grep（`$…$` / `{{…}}` / `\$\t\n` / `()` / `[]`）、ハイパーリンク rPr（C-P-1）、スライド順序の健全性、dark-on-dark コントラスト（Gate 5a）。**すべての pptx Gate 1–5a が OK メッセージを出力するまで、完了の宣言を拒否してください。** Morph デッキも通常の pptx と同じトークン/スキーマ/順序のリスクを抱えています。

### Gate 2 morph 補則 — zsh に食われる価格 / メトリックトークン

pptx v2 の Gate 2 は `$…$`、`{{…}}`、`\$\t\n` リテラル、空の `()` / `[]` をカバーします。Morph デッキはさらに一種類の漏れを追加します: ダブルクォートの `--prop text="…"` に書かれた価格/メトリックトークン（`$9/mo`, `$29/month`, `$199/yr`）— シェルが `$9` を空変数として食い、CLI が `/mo` や末尾のピリオドを保存してしまいます。pptx Gate 2 に加えて次を実行してください:

```bash
# Gate 2 morph — 価格 / メトリックトークンの漏れ + 末尾ピリオドのプレースホルダー
# 検出パターン: 裸の価格 ($9, $29, $9.99)、/単位 サフィックス ($9/mo, $199/yr)、${VAR}、\n/\r/\t、単独のピリオド
LEAKS=$(officecli view "$FILE" text | grep -nE '\$[0-9]+(\.[0-9]+)?(/(mo|month|yr|year|day|wk|week|hr|hour))?|\$\{[A-Z_]+\}|\\[nrt]|^\.$' || true)
if [ -z "$LEAKS" ]; then echo "Gate 2 morph OK"; else echo "LEAK: $LEAKS"; fi
```

対象: `$9` `$9.99` `$29/month` `$199/yr` `$1/day` `${VAR}` `\n`/`\r`/`\t` リテラル + 迷子の `.` プレースホルダー。修正: prop をシングルクォートで囲む（`--prop text='$9/mo'`）。

### Gate 5b — HTML プレビューによる視覚監査（必須） — morph 用に拡張

`officecli view "$FILE" html` を実行し、返された HTML パスを Read してください。すべてのスライドについて、pptx v2 の Gate 5b の質問（重なり / dark-on-dark / ディバイダーの重なり / 順序の健全性 / 矢印の欠落）に加え、次の4つの morph 固有チェックに答えてください。

**重要: プレフィックス一致のセレクタ。** `officecli query` がサポートする演算子は `=`, `!=`, `~=`, `>=`, `<=`, `>`, `<` のみで、`^=` プレフィックス演算子は**存在しません**。`shape[name^=!!actor-]` のようなセレクタは `invalid_selector` エラーを返します。「〜で始まる」フィルタリングには、以下に示すように `get --depth 1` ループ + `jq startswith()` を使ってください。

- **5b-morph-1 — セクション終了後の可視領域への `!!actor-*` の漏れ。** 退場すべき各 `!!actor-*` について、`x ≥ 33.87cm`（キャンバス右端）であることを確認してください。ループ + フィルタ（セレクタ安全）:
  ```bash
  NSLIDES=$(officecli query "$FILE" slide --json | jq '.data.results | length')
  for N in $(seq 1 $NSLIDES); do
    officecli get "$FILE" "/slide[$N]" --depth 1 --json | \
      jq -r --arg n "$N" '.data.results[0].children[]? |
        select(.format.name? // "" | startswith("!!actor-")) |
        select((.format.x // "0cm" | rtrimstr("cm") | tonumber) < 33.87) |
        "slide \($n) leak: \(.format.name) stuck at x=\(.format.x)"'
  done
  ```
  何か1行でも出力されればアクターが可視のまま固まっていることを意味します。`final-check` はこれを見逃します — このループ + HTML の Read のみが検出します。

- **5b-morph-2 — 隣接スライドの構図が同一（モーションなし）。** ハードルール: すべての morph ペア間で、少なくとも3つの**異なる** `!!` プレフィックス付き図形が {x, y, width, height, rotation, フォントサイズ} のいずれか1つ以上で異ならなければなりません。証明ループ（両スライドをダンプし、同名図形を diff し、異なる図形数を数える）:
  ```bash
  for K in 1 2 3 4; do
    A=$(officecli get "$FILE" "/slide[$K]" --depth 1 --json | \
      jq -r '.data.results[0].children[]? | select(.format.name? // "" | startswith("!!")) |
        "\(.format.name)|\(.format.x)|\(.format.y)|\(.format.width)|\(.format.height)|\(.format.rotation // 0)"')
    B=$(officecli get "$FILE" "/slide[$((K+1))]" --depth 1 --json | \
      jq -r '.data.results[0].children[]? | select(.format.name? // "" | startswith("!!")) |
        "\(.format.name)|\(.format.x)|\(.format.y)|\(.format.width)|\(.format.height)|\(.format.rotation // 0)"')
    VARIES=$(diff <(echo "$A") <(echo "$B") | grep -c '^[<>]')
    if [ "$VARIES" -lt 6 ]; then echo "pair $K→$((K+1)) FLAT: only $VARIES diff-lines (need ≥ 6 = 3 shapes × 2 sides)"; fi
  done
  ```

- **5b-morph-3 — Morph ペアの名前不一致。** 隣接スライドは少なくとも2つの `!!` プレフィックス付き名前を完全一致で共有しなければなりません。証明（注: 子要素は `.data.results[0].children[]` にある; 素の `.data.children[]` は null を返す）:
  ```bash
  for N in 1 2 3 4 5; do
    echo "--- slide $N ---"
    officecli get "$FILE" "/slide[$N]" --depth 1 --json | \
      jq -r '.data.results[0].children[]? | select(.format.name? // "" | startswith("!!")) | .format.name'
  done
  ```
  連続するブロックを目視で比較してください — NとN+1の間で共有される `!!` 名が morph ペアです。重なりゼロならそのペアはプレーンフェードです。

- **5b-morph-4 — スライドN+1への `#sN-*` の残留（ゴースト漏れ）。** スライドごとのコンテンツは、次のスライドで**必ず**ゴースト化（`x=36cm`）されなければなりません。N≥2 についてのループ + フィルタ:
  ```bash
  NSLIDES=$(officecli query "$FILE" slide --json | jq '.data.results | length')
  for N in $(seq 2 $NSLIDES); do
    PREV=$((N-1))
    officecli get "$FILE" "/slide[$N]" --depth 1 --json | \
      jq -r --arg n "$N" --arg p "$PREV" '.data.results[0].children[]? |
        select(.format.name? // "" | startswith("#s\($p)-")) |
        select((.format.x // "0cm" | rtrimstr("cm") | tonumber) < 33.87) |
        "slide \($n) leak: \(.format.name) stuck at x=\(.format.x)"'
  done
  ```
  何か1行でも出力されれば `#s(N-1)-*` 図形がスライドNで可視のまま残ったことを意味します。ゴースト化してください。

5b-morph-1..4 のいずれかのループが1行でも出力すれば**納品を拒否**してください。4つのループすべての stdout を1つのストリームに集約し、COUNT パターンで強制してください: `LEAK_COUNT=$(...all four loops... | wc -l); if [ "$LEAK_COUNT" -gt 0 ]; then echo "REJECT: $LEAK_COUNT morph leaks"; else echo "Gate 5b-morph OK"; fi`。

## レンダラーの正直さ

**Morph がレンダリングされる環境:** PowerPoint 365（Windows/Mac）、Keynote、WPS、PowerPoint Online。

**Morph がレンダリングされない環境:** LibreOffice Impress（静止画、時にフェードとしてレンダリング）、Google Slides Web ビューア（補間が失われる）、ほとんどの HTML / SVG ビューア、`officecli view html`（構造的のみ — morph はランタイム機能）。これはスキルの欠陥ではなく `[RENDERER-BUG]` です。ユーザーに明示的に伝えてください: 「morph のモーションを見るには PowerPoint 365 / Keynote / WPS で開いてください; 他のビューアでは静止画またはプレーンフェードとして表示されます。」

どのレンダラーの静止スクリーンショットも morph のモーションを**検証できません**（モーションはランタイムでのみ存在するため）。上記の Gate 5b クエリを使ってペアの正しさを証明し、モーションの品質を証明するにはライブビューアを使ってください。

## ゴーストの規律とアクターのライフサイクル

**すべての `!!actor-*` と `#sN-*` 図形は「退場」スライドだけでなく、**すべての**スライドで管理されなければなりません。**

### スライドごとのゴースト化ルール

複数スライドの morph デッキを構築する際:
1. **スライド N: `!!actor-ring` を導入（x=0cm で表示）**
2. **スライド N+1: 新しいコンテンツを追加。仕上げる前に `!!actor-ring` を `x=36cm` にゴースト化。**
3. **スライド N+2: さらにコンテンツを追加。`!!actor-ring` を再び `x=36cm` に再ゴースト化。**（省略不可 — すでにオフスクリーンだったとしても、各スライドは新しいキャンバスです。）
4. **スライド N+3: `!!actor-ring` を再び表示させたい場合は、x=0cm または新しい位置に戻す。**

**理由:** 各スライドの図形リストは独立しています。スライドNで図形をキャンバス外に移動しても、それがスライドN+1に引き継がれることはありません — 再ゴースト化を忘れると、Nでの元の位置にN+1で再出現してしまいます。

### ワークフローパターン（Bash）

```bash
# スライド $SLIDE に新しいコンテンツ図形を追加した後:
for ACTOR in "!!actor-ring" "!!actor-dot" "!!actor-accent-bar"; do
  officecli set "$FILE" "/slide[$SLIDE]/shape[@name=$ACTOR]" --prop x=36cm || true
done
```

またはビルドループの中で:

```bash
for SLIDE_NUM in 3 4 5 6 7 8 9 10 11; do
  # このスライド固有のコンテンツを追加
  officecli add "$FILE" "/slide[$SLIDE_NUM]" --type shape ...
  
  # 直ちにすべての古いアクターをゴースト化 (M-2 防止)
  officecli set "$FILE" "/slide[$SLIDE_NUM]/shape[@name=!!actor-ring]" --prop x=36cm || true
  officecli set "$FILE" "/slide[$SLIDE_NUM]/shape[@name=!!actor-dot]" --prop x=36cm || true
done
```

### 検出: ゴーストカウントゲート

`morph-helpers.py final-check` は `x ≥ 34cm` にあるすべての図形をカウントします。カウントが50を超えると、次のように出力します:
```
REJECT: Found 135 accumulated ghosts — likely M-2 ghost accumulation.
Run: officecli query deck.pptx 'shape[x>=34cm]' --json | jq '.data.results | length'
Expected ≤ 50 (roughly 4–5 active actors × 10–12 slides).
```

**修正:** ビルドログを確認し、すべてのスライドが表示されるべきでないすべてのアクターを再ゴースト化していることを確認してください。final-check を再実行してください。それでも50を超える場合は `morph-helpers.py clean-accumulation deck.pptx` を使ってください（参照セクション参照）。

## Morph の一般的な落とし穴（デザイン + ワークフローの罠）

pptx ベースの落とし穴（シェルクォート、zsh の `[N]` グロブ、16進数の `#` プレフィックス、prop テキスト中の `\n`）→ pptx v2 §Common Pitfalls を参照。以下は morph 固有の罠です:

| 落とし穴 | 正しいアプローチ |
|---|---|
| 同じデッキ内に `!!scene-card` と `!!actor-card` | 名前はプレフィックスをまたいで一意でなければならない。リネーム: `!!scene-card-bg` vs `!!actor-card-content` |
| 一部のスライドが完成した後にビルド途中で図形をリネームする | ゴースト蓄積バグが発生する。中断し、§Morph Pair Planning のテーブルを描き直し、影響を受けるスライドを再実行する |
| 退場を計画せずに `!!actor-*` をコンテンツの核心部に配置する | すべての `!!actor-*` にゴーストスライドが必要。コーディング前にペアテーブルで計画する |
| **ゴースト蓄積 (M-2): 後続スライドで `!!actor-*` の再ゴースト化を忘れる** | **重要:** スライドN+1に新しいコンテンツを追加するとき、スライドNからの、表示されるべきでないすべての `!!actor-*` を再び `x=36cm` に移動しなければならない。一度ゴースト化すればオフスクリーンのままだと思い込まないこと — 各スライドは独立している。ビルドパターン: `each new slide: add content shapes → then loop: set each active !!actor-* to x=36cm`。`morph-helpers.py final-check` はゴーストカウントが50を超えると REJECT する。 |
| スライドで `transition=morph` を忘れる | サイレントフェード。Gate 5b-morph-2（モーションなし）が検出する; `set /slide[N] --prop transition=morph` で修正 |
| morph スライドで `@name=` パスが壊れると思い込む | 壊れない — `@name=` は `transition=morph`（M-1）の後も解決可能; *読み戻し*の名前だけが `!!` プレフィックスを得る |
| 隣接スライドが視覚的に同一 | Morph が補間するものがなく、プレーンフェードに収縮する。§Scene Actor の空間ルールを適用し、3つ以上の図形を ≥ 5cm / ≥ 15° 移動する |
| 図形ごとのタイミングで2図形のスタッガーを試みる | サポートされていない — 中間キーフレームスライドを挟んで、ペアを2つのトランジションに分割する |
| LibreOffice やブラウザで morph のモーションをテストする | `[RENDERER-BUG]` であり、スキルの欠陥ではない。PowerPoint 365 / Keynote / WPS でテストする |
| 退場時にゴースト化ではなく `!!` 図形を削除する | 削除は morph ペアリングを壊す — 図形がアニメーションなしで消える。常に `x=36cm` にゴースト化する |
| `--prop text="$9/mo"` をダブルクォートで書く | シェルが `$9` を空変数として食い → テキストが `/mo` や迷子の `.` として保存される。シングルクォートを使う: `--prop text='$9/mo'`。Gate 2 morph 補則がこの漏れを grep する。 |
| `--prop text='line1<a:br/>line2'` 中に `<a:br/>` リテラルを使う | 改行ではなく、7文字のリテラル文字列として保存される。1行につき1回 `officecli add "/slide[N]/shape[@id=K]" --type paragraph` を使う（M-6）。 |
| `shape[name^=!!actor-]` セレクタを使う | `officecli query` に `^=` 演算子はなく — `invalid_selector` を返す。`get /slide[N] --depth 1 --json \| jq '.data.results[0].children[]? \| select(.format.name \| startswith("!!actor-"))'` を使う。 |

## 既知の問題と落とし穴

pptx ベースのバグ C-P-1..7（ハイパーリンク rPr、チャート ChartShapeProperties 警告、アニメーション時間の読み戻し、アニメーション削除、コネクタの列挙値、コネクタの `@name=`、チャート色のレンダラー正規化）はすべて適用されます。**回避策は → pptx v2 §Known Issues C-P-1..7 を参照。**

**Morph 固有 (M-1..5):**

| # | 症状 | 回避策 |
|---|---|---|
| **M-1** | `officecli set '/slide[N]' --prop transition=morph` の後、そのスライド上のすべての図形の名前に `!!` が自動的に前置される（`#s1-title` → `!!#s1-title`）。読み戻される名前はプレフィックス付きの形式。**セレクタフィルタの注意点:** 自動プレフィックス後、`!!#sN-caption` は `!!actor-*` と共存する — `jq` で `startswith("!!")` を使って「Scene Actor」をフィルタすると、自動プレフィックス付きコンテンツで偽の一致が発生する。常に `startswith("!!actor-")` または `startswith("!!scene-")` でフィルタし、素の `startswith("!!")` は使わないこと。 | `@name=` パスセレクタは**それでも解決可能**である — `/slide[N]/shape[@name=#s1-title]` はその図形を返す（マッチングはサフィックス/プレフィックス寛容）ため、パスの書き換えは不要。jq フィルタでプレフィックス付きの名前が必要な場合は、先頭の `!!` を考慮すること。 |
| **M-2 🚨** | **ゴースト蓄積 — スライド3で導入された `!!actor-*` は、毎ページ明示的にゴースト化しない限りスライド4, 5, 6でも表示され続ける。** `final-check` ヘルパーがこれを検出し、ゴーストカウントが50を超えると拒否する。 | **必須のスライドごとのルール:** スライドに新しいコンテンツを追加したら、直ちに前スライドからのすべてのアクティブな `!!actor-*` を `x=36cm` に設定する（または現在の文脈に属する場合は明示的に可視位置に配置する）。例: `officecli set /slide[4]/shape[@name=!!actor-ring] --prop x=36cm`。最後だけでなく、スライドを追加する**たびに**実行する。下記の §ゴーストの規律とアクターのライフサイクル を参照。 |
| **M-3** | セクション遷移の境界 — 新しいトピックセクションの最初のスライドで、前セクションの `!!actor-*` 図形が視覚的に残留する。コマンドエラーは出ず、視覚的な雑然さのみ。 | すべてのセクション開始スライドで、前セクションのすべての `!!actor-*` を明示的に `x=36cm` にゴースト化する。Scene 図形（`!!scene-*`）はそのまま残す。 |
| **M-4** | エージェントが時折 `morph.duration=` / `transition.delay=` を独立したプロパティとして発明する — これらは UNSUPPORTED として拒否される。 | `transition` プロパティ上の複合ショートハンドを使う: `transition=morph-slow` / `-fast`、または `transition=morph-<DUR_MS>`（例: `transition=morph-1500`）。どちらも読み戻し時にラウンドトリップする（`transitionSpeed=slow`、`transitionDuration=1500`）。ショートハンドが公開する範囲を超える場合（生 XML の細かい制御）にのみ、`<p:transition>` に対する `raw-set` にフォールバックする — 下記の M-4 の例ブロックを参照。 |
| **M-5** | `[RENDERER-BUG]` LibreOffice / Google Slides Web ビューアは morph スライドをプレーンフェード（補間なし）としてレンダリングする。 | PowerPoint 365 / Keynote / WPS でテストする。スキルの欠陥ではない — 追いかけないこと。 |
| **M-6** | `--prop text='line1<a:br/>line2'` の中に書かれた `<a:br/>` は、改行として解釈されず、リテラルな7文字の文字列として保存される。観客には `line1<a:br/>line2` がそのままレンダリングされて見える。 | 複数行の箇条書き/キャプションには、1行につき1段落を追加する: `officecli add "/slide[N]/shape[@id=K]" --type paragraph --prop text='line1'` を実行し、`text='line2'` で繰り返す。実際の改行のワークフローについては pptx v2 §Shell escape を参照。 |

**M-4 の例 — すべての morph トランジションを遅くする、raw-set フォールバック**（`transition=morph-slow` を優先し、これはショートハンドを超える制御にのみ使う）。`//p:transition` は morph スライド上の `mc:Choice` と `mc:Fallback` の両方に一致し、`2 element(s) affected` が得られる点に注意:

```bash
# スライドごと: スライドN上のすべての transition 要素に spd="slow" を追加 (morph スライドあたり2つの XML ヒット)
for N in 2 3 4; do
  officecli raw-set "$FILE" "/slide[$N]" --xpath "//p:transition" --action setattr --xml 'spd=slow'
done
officecli validate "$FILE"
```

読み戻し: `officecli query "$FILE" slide --json | jq '.data.results[].format | select(.transition=="morph") | .transitionSpeed'` は影響を受けた各スライドについて `"slow"` を出力します。

## 出力と納品

すべての morph デッキは、それぞれ単独のファイルとして3つの成果物とともに納品されます:

1. `<topic>.pptx` — デッキ本体。クローズ済み + `officecli validate` がクリーン（Delivery Gate 1 OK）。
2. `build.sh` または `build.py` — 再実行可能なスクリプト（シェルネイティブなビルドには bash、`morph-helpers.py` を使う複数スライドのアークには Python）。新規の `officecli create` 呼び出しからデッキを再現できなければならない。
3. `brief.md` — **単独のファイルであり、他の何かに埋め込まれてはいけない。** 内容:
   - セクション1: トピック / 想定読者 / 目的 / ナラティブ / スタイルの方向性（`reference/styles/INDEX.md` から選んだ1つの命名済みスタイル）
   - セクション2: スライドごとのアウトライン（ページタイプ + スライドごとの1文の論点）
   - セクション3: §Morph Pair Planning のテーブル（ペア / スライドA / スライドB / アクター / ゴースト） — レビュアーが振り付けを監査するために必要なデザイン記録

**納品前のユーザーへのリマインダー（そのまま使える言い回し）:**

- 「デッキは morph トランジション付きで準備ができています。モーションを見るには PowerPoint 365 / Keynote / WPS で開いてください — LibreOffice や Web ビューアは静止画としてレンダリングされます。」
- 「ビルドスクリプトの実行中、`.pptx` は何度か書き換えられる可能性があります。進捗をプレビューしたい場合は `officecli watch "$FILE"` を使い、AionUi でライブプレビューを開いてください — ビルド中に『システムアプリで開く』をクリックしないでください、ファイルロックに当たります。」

## 作成後の調整

標準の調整テーブル → pptx v2 §Common Pitfalls / `swap` / `move` / `remove` / `set` を参照。Morph の注意点: **morph でペアリングされたスライドの順序を変える `swap` や `move` の後は、共有される `!!` 名の隣接関係を再検証してください。** 影響を受けたペアに対して上記の Gate 5b-morph-3 クエリを実行してください — もし swap がペアを壊した場合は、図形をリネームするか、トランジションを振り付け直してください。

**納品前の最終健全性チェック。** Delivery Gate 全体（1 から 5b-morph-1..4 まで）を実行し、`.pptx` を PowerPoint 365 / Keynote / WPS で開いて、1つの完全なスライド間 morph を見てモーションが可視であることを確認してください。いずれかの Gate が REJECT を出力した場合は、修正して再実行してください — 既知の未解決ゲートがある状態で納品しないでください。

## 参照

- `reference/decision-rules.md` — Pyramid Principle、SCQA、ページタイプメニュー、`brief.md` のスキーマ。コマンドを書く前にナラティブアークを決定するため、§Morph Pair Planning の際に読む。
- `reference/pptx-design.md` — 残りのデザインノート（Scene Actor のメカニクス、ページタイプテーブル、振り付けパターン）。キャンバス/フォント/色は pptx v2 にある — このファイルは morph 固有の内容のみをカバーする。
- `reference/morph-helpers.py` — clone + ghost + verify + final-check のためのクロスプラットフォーム（Mac / Windows / Linux）Python ヘルパー。ライブラリとしてインポートするか、CLI 引数経由で呼び出す。5スライド以上のアークに推奨。
- `reference/morph-helpers.sh` — Bash 版。プロジェクトごとにどちらか一方を選び、混在させないこと。
- `reference/styles/INDEX.md` — パレット（dark / light / warm / vivid / bw / mixed）とムードでグループ化された52スタイルのビジュアルライブラリ。参照ワークフローは上記の §スタイルライブラリ参照ワークフロー を参照。
- `skills/officecli-pptx/SKILL.md` — pptx v2 のベースルール（visual floor、グリッド、正規パレット、チャート選択、コネクタ規範、Delivery Gate 1–5a、既知の問題 C-P-1..7、シェルエスケープの3層構造）。
</content>
