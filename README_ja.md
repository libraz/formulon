# Formulon

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon/actions/workflows/ci.yml)
[![npm](https://img.shields.io/npm/v/@libraz/formulon)](https://www.npmjs.com/package/@libraz/formulon)
[![PyPI](https://img.shields.io/pypi/v/formulon)](https://pypi.org/project/formulon/)
[![codecov](https://codecov.io/gh/libraz/formulon/branch/main/graph/badge.svg)](https://codecov.io/gh/libraz/formulon)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon/blob/main/LICENSE)
[![C++17](https://img.shields.io/badge/C%2B%2B-17-blue?logo=c%2B%2B)](https://en.cppreference.com/w/cpp/17)
[![Platform](https://img.shields.io/badge/platform-Linux%20%7C%20macOS%20%7C%20WebAssembly-lightgrey)](https://github.com/libraz/formulon)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-2563eb)](https://formulon.libraz.net)

**Formulon は Excel 互換の計算エンジンです。** C++17 製のコアエンジンが、既定では **Windows Excel 365 (ja-JP)** の挙動に合わせて数式を評価します。実 Excel から取得した oracle データで互換性を確認し、既知の差分はすべて理由つきで追跡しています。同じエンジンをブラウザ (WebAssembly)、Python、ネイティブ CLI から使えるため、どの実行環境でも同じワークブックを同じ結果に再計算できます。

Excel 本体、Microsoft ランタイム、COM オートメーションは実行時には不要です。WASM 版はブラウザと Node で動作し、Python 版は `wasmtime` 経由で同じ WASM コアを呼び出します。ネイティブ CLI は `darwin-arm64` / `linux-x64` / `linux-arm64` 向けに配布しています。

## インストール

```bash
npm install @libraz/formulon   # JavaScript / TypeScript (WASM)
pip install formulon            # Python
```

CLI バイナリは [GitHub Releases](https://github.com/libraz/formulon/releases) から取得できます。

## 特徴

- **互換性は実際の Excel と照合して確かめています。** 既定の profile は `win-365-ja_JP` です。数式の結果は Mac Excel 365 (ja-JP)、ピボットテーブルと印刷レイアウトは Windows Excel 365 (ja-JP) を基準に固定しています。ピボットテーブルの作成を自動化するには Windows COM が必要なためです。いずれも検証済みの Microsoft 365 環境から採取しています。出力は実 Excel から再生成した golden と照合します。許容している差分、たとえば超越関数の ulp 差、揮発関数、Excel 側の不整合を Formulon が意図的に採らないケースは、[`tests/divergence.yaml`](tests/divergence.yaml) に理由と確認済み Excel ビルドを記録します。
- **どの環境でも同じ C++ コアが計算します。** ブラウザ、Python、CLI で別々の計算ロジックを持たず、同じエンジンを配布しています。実装が分かれないので、環境ごとに結果がずれることもありません。
- **WASM のサイズに上限を設けています。** CI は非圧縮 **3.00 MiB**、Brotli **768 KiB** を超えると失敗し、**2.75 MiB** / **736 KiB** を超えると警告を出します。実際に効いてくるのは配信時の Brotli サイズなので、非圧縮と対等に検査します。現在値は `make size-check` で確認できます。
- **依存は小さく保っています。** ランタイム依存は `miniz` (zip/deflate)、`pugixml` (XML + XPath 1.0)、`PCRE2` (`REGEX*`)、`double-conversion` (Grisu3 `dtoa`) の 4 つです。線形代数、UTF-8 処理、数値変換の多くはリポジトリ内で実装しています。
- **C++ は監査しやすさを優先して書いています。** `Expected<T, Error>` ベースのエラー処理、RAII、`-fno-exceptions -fno-rtti`、Google C++ Style を採用しています。

## 使いどころ

Excel を起動せずにスプレッドシートを計算したい場面で使えます。

- バッチジョブやデータパイプラインで `.xlsx` をヘッドレス再計算する
- Web アプリの中で Excel 風の数式を評価する
- 社内ツール、ボット、ノートブックに計算機能を組み込む
- 数式の検証や、レガシースプレッドシートの移行に使う

組み込みの実装例として [formulon-cell](https://github.com/libraz/formulon-cell) があります。`@libraz/formulon` の上に作ったブラウザ向けのスプレッドシート UI で、npm パッケージを実ブラウザで一通り動かす結合テストも兼ねています。

## やらないこと

Formulon は以下を **意図的にサポートしません**。

| 項目 | 理由 |
|------|------|
| VBA の実行 | セキュリティ上の理由からです。`vbaProject.bin` はバイト列としてだけ保存し、実行はしません。 |
| 旧 `.xls` (BIFF8 / Excel 97-2003) | Excel 365 互換の対象外です。 |
| グラフ / 図形のレンダリング | 描画レイヤの責務。計算エンジンの仕事ではありません。 |
| PowerQuery (M) / DAX | 別のエンジンが扱う別の問題領域です。 |
| Pivot キャッシュの再生成 | 保存済み `pivotCacheRecords` はそのまま保持し、ソース範囲から作り直すことはしません。PivotTable の計算結果自体は API から都度評価できます。 |
| スプレッドシート UI 本体 | 描画は呼び出し側の責任です。UI から使うための API(ビューポート単位の再計算 `partialRecalc`、範囲単位の条件付き書式評価、spill 情報など)はエンジン側で提供しています。 |

これらは「まだやっていない」機能ではなく、スコープ外として固定しています。

## パッケージ

| 配布元 | パッケージ名 | 内容 |
|--------|-------------|------|
| npm | [`@libraz/formulon`](https://www.npmjs.com/package/@libraz/formulon) | WASM ESM モジュール。型定義同梱。Node 22+ / ブラウザ / Worker 対応。既定のビルドはシングルスレッドで cross-origin isolation 不要、`@libraz/formulon/threads` は `recalcParallel` 用のワーカースレッド付きです。 |
| PyPI | [`formulon`](https://pypi.org/project/formulon/) | Python 3.9+ の `py3-none-any` wheel。`formulon_capi.wasm` と pure-Python wrapper を同梱し、`wasmtime` は `pip` が解決します。 |
| GitHub Releases | `formulon-<version>-<platform-arch>.tar.gz` | 単体 CLI バイナリ (`eval` / `recalc` / `dump` / `paginate`)。`darwin-arm64` / `linux-x64` / `linux-arm64` 向け。 |

同じ入力からはどの配布形態でも同じ結果が出ます。3 つとも、ワークシートのパートをパース前にまるごとメモリへ展開します（zip リーダーの上限は 1 エントリ 100 MiB、1 回の読み込みで 256 MiB）。WASM ビルド（npm と PyPI のパッケージ）は、そのバッファから必ず DOM ツリーを構築します。ネイティブ CLI は 256 KiB を超えるワークシートではストリーミングパーサに切り替え、DOM ツリーを作りません。このパーサはバイナリサイズを食うため、WASM には入れていません。ただし CLI も展開済みの XML バッファは先に持つので、ピークメモリが一定になるわけではありません。どの配布形態でもシートは 1 枚ずつ読むため、ピークはワークブック単位ではなくワークシート単位です。WASM での実際の上限はホストの 32-bit アドレス空間です。どちらのパーサでも結果は同じです。

## コマンドライン

リリースのバイナリを `PATH` に置くと、次の 4 つのコマンドが使えます。`eval` は数式 1 つの評価、`recalc` は再計算したブックの書き出し、`dump` はテキストでの内容出力、`paginate` は印刷範囲・改ページ・ページ数の算出です。

```bash
formulon eval '=SUM(1,2,3)'
formulon recalc input.xlsx -o output.xlsx
formulon recalc --threads 4 input.xlsx -o output.xlsx
formulon dump output.xlsx --formulas
formulon paginate output.xlsx --sheet 0
```

4 つのコマンドはいずれも `--` でオプションの解釈を終了できます。`--` より前にオプションを指定し、その後には `eval` では数式を、`recalc`・`dump`・`paginate` では入力パスを 1 つだけ渡します。これにより `-` で始まる相対パスも扱えます（例: `formulon dump --sheets -- -input.xlsx`）。

`recalc` は入出力とも `.xlsx` / `.xlsb` を受け付けます。成功すると stderr に `formulon: recalc: ok, wrote M bytes to 'OUT'` を出力します。このステータス行は `--quiet` で消せますが、読み込み・保存時に情報が失われる旨の警告は XLSB・OOXML とも `--quiet` を付けても表示されます。既定では直列に再計算します。`--threads N` を渡すと並列 SCC スケジューラに切り替わり（`0` は自動検出、`1` は呼び出し元スレッドのまま、`2..8` はワーカー数の上限）、ステータス行にパスごとの計測値も出ます。

## ステータス

カタログに登録した Excel 関数は **522 個すべてを認識**します。ただし、「関数名を知っている」ことと「Excel 互換の実装がある」ことは分けて扱います。現在の内訳は `make function-status` で確認できます。

上位 2 区分（実装済み・unavailable stub）は排他的で、合計は 522 に一致します。環境依存の行は**実装済み 507 件の内数**です。固定の golden では記述しきれないため別に示していますが、独立した区分ではありません（507 + 15 = 522 で、524 にはなりません）。

| 区分 | 件数 | 意味 | 例 |
|------|------|------|----|
| 実装済み | 507 | 通常の計算エンジン内で評価できる関数。unit / oracle で検証しています。 | 数学、統計、検索、テキスト、動的配列など |
| &nbsp;&nbsp;↳ うち環境依存 | 2 | 実装済みだが、ホスト環境やワークブック状態によって値が変わるため固定 golden だけでは完全に記述できない関数。上記 507 に含まれます。 | `INFO`, `CELL` |
| unavailable stub | 15 | Formulon が内蔵しない外部サービス、ネットワーク、COM、OLAP 接続などが必要な関数。決まったエラーを返します。 | `PY`, `WEBSERVICE`, `STOCKHISTORY`, `IMAGE`, `RTD`, `TRANSLATE`, `DETECTLANGUAGE`, `COPILOT`, `CUBE*` |

oracle は **104 カテゴリ** あります。数式 track と条件付き書式 track は Mac Excel 365 ja-JP から、workbook track は Windows Excel 365 ja-JP から再生成します。workbook track の golden には採取 ID が付いていて、すべての suite が同じ検証済み Microsoft 365 セッションで採取されたことを確認できます。

現在のローカル検証結果:

| 検査 | 結果 |
|------|------|
| `ctest -LE "SLOW\|BENCH\|TSAN"` — `make test`、PR ゲート | すべて passed |
| `ctest -LE "BENCH\|TSAN"` — `make test-slow`、`SLOW` 層を追加 | すべて passed |
| primary formula oracle | `4546/4546` passed / `125` documented skips |
| 条件付き書式 oracle | `23/23` |
| workbook oracle (pivot + print) | `73/73` passed / `9` documented skips |
| 取り込み済み外部エンジンコーパス (クロスチェック) | `12510/12510` passed / `168` documented divergences |

CTest スイートを分けているラベルは 3 つです。`SLOW`（数分かかる結合・並行性テスト）、`TSAN`（ThreadSanitizer での実行）、`BENCH`（しきい値を調整できるマイクロベンチの回帰チェックで、必要なときだけ実行）です。ラベルのないテストはすべて高速層で、CI はこれを合否判定に使います。負荷試験専用の層はありません。libFuzzer ハーネスも `SLOW` ラベルを持ちますが、`-DFM_BUILD_FUZZ=ON` を指定したビルド (`make fuzz`) にしか存在せず、既定ビルドにも CI にも含まれません。libFuzzer ランタイムを同梱する Clang が必要で、Apple の toolchain はこれを持たないため、macOS では別途 LLVM が要ります。macOS では AddressSanitizer も既定で無効です。最近の macOS の動的リンカと組み合わせると、shadow memory の初期化中にデッドロックするためです。そのため macOS での fuzz 実行で検出できるのはクラッシュ・タイムアウト・未定義動作までで、ヒープ破壊は検出できません。

残っている skip は、明示済みの divergence、ホストサービス依存、揮発・環境依存ケース、またはドライバ制約です。黙って未実装 stub に落としているものではありません。522 関数のうち `518` は closure 6 条件 (`behaviors_declared` / `cases_cover_behaviors` / `golden_present` / `divergence_documented` / `not_in_pilot` / `behavior_drift`) を全て満たします。残る 4 件 (`ARRAYTOTEXT`, `FILTERXML`, `GETPIVOTDATA`, `PHONETIC`) が満たさないのは `behaviors_declared` だけで、挙動の分類がまだ書き切れていないためです。`JIS` は `DBCS` の別名として宣言し、closure を満たしています。Excel は ja-JP の数式バーで入力された `JIS` を保存・評価の前に `DBCS` へ書き換えるため、`JIS` を直接呼ぶ oracle case は作れません。closure harness は宣言をそのまま信用せず、別名の参照先の関数を実際に評価して判定します。

数式の結果に加えて、**ピボットテーブルと印刷範囲・改ページ**には専用の **workbook oracle track** があり、WSL2 から Windows COM へ渡すブリッジ経由で採取します。残る 9 件の skip はいずれも同じ Excel の癖です。印刷倍率またはズームが 50% 以下のとき、Excel の改ページプレビューは幾何的なページ分割に従わない列の自動改ページを出すため、観測される改ページ位置は倍率を下げても縮まらず、25% では逆に増えます。skip した各ケースには、照合した Microsoft 365 の観測値を記録しています。

新規ワークブックはデフォルトで `win-365-ja_JP` profile を使います。必要に応じて profile-id API (`mac-365-ja_JP` / `win-365-ja_JP`) で切り替えられます。英語ロケール profile は、対応する EN oracle データとロケール固有挙動の検証が揃うまで公開しません。

OOXML reader / writer はシート、スタイル、条件付き書式、コメント、ハイパーリンク、結合セル、入力規則、定義済み名前、テーブル、ピボットテーブルを round-trip します。MS-XLSB reader / writer はセル値、スタイル、シート間 3-D 参照、および一般的なトークン化数式をカバーします。配列定数リテラルと 2007 年以降の future function ID は、OOXML 経路に比べてまだ限定的です。ワークブック操作は C ABI と各言語バインディングから利用できます。CLI は意図的に `eval` / `recalc` / `dump` / `paginate` のみを公開します。

不具合報告・oracle 差分レポート・ご意見はいつでも歓迎しています。

## コントリビューション

いちばん助かるのは、**手元の Excel から oracle データを提供していただくこと**です。Mac ja-JP 以外の Excel 365 をお持ちなら、`make oracle-contribute` で Excel を駆動して golden を取得し、PR の手順まで進められます。詳しくは [CONTRIBUTING.md](CONTRIBUTING.md) を参照してください。

## ライセンス

Apache License 2.0。[LICENSE](LICENSE) および [NOTICE](NOTICE) を参照してください。
