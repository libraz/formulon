# Formulon

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon/actions/workflows/ci.yml)
[![npm](https://img.shields.io/npm/v/@libraz/formulon)](https://www.npmjs.com/package/@libraz/formulon)
[![PyPI](https://img.shields.io/pypi/v/formulon)](https://pypi.org/project/formulon/)
[![codecov](https://codecov.io/gh/libraz/formulon/branch/main/graph/badge.svg)](https://codecov.io/gh/libraz/formulon)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon/blob/main/LICENSE)
[![C++17](https://img.shields.io/badge/C%2B%2B-17-blue?logo=c%2B%2B)](https://en.cppreference.com/w/cpp/17)
[![Platform](https://img.shields.io/badge/platform-Linux%20%7C%20macOS%20%7C%20WebAssembly-lightgrey)](https://github.com/libraz/formulon)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-blue)](https://formulon.libraz.net)

**Formulon は Excel 互換の計算エンジンです。** C++17 製のコアエンジンが、既定では **Windows Excel 365 (ja-JP)** の挙動に合わせて数式を評価します。実 Excel から取得した oracle データで互換性を確認し、既知の差分はすべて理由つきで追跡しています。同じエンジンをブラウザ (WebAssembly)、Python、ネイティブ CLI から使えるため、どの実行環境でも同じワークブックを同じ結果に再計算できます。

Excel 本体、Microsoft ランタイム、COM オートメーションは実行時には不要です。WASM 版はブラウザと Node で動作し、Python 版は `wasmtime` 経由で同じ WASM コアを呼び出します。ネイティブ CLI は `darwin-arm64` / `linux-x64` / `linux-arm64` 向けに配布しています。

## インストール

```bash
npm install @libraz/formulon   # JavaScript / TypeScript (WASM)
pip install formulon            # Python
```

CLI バイナリは [GitHub Releases](https://github.com/libraz/formulon/releases) から取得できます。

## 特徴

- **互換性を oracle で確認します。** 既定の profile は `win-365-ja_JP` です。primary oracle は Mac Excel 365 (ja-JP) で、Windows Excel 365 (ja-JP) は variant golden として管理しています。出力は実 Excel から再生成した golden と照合します。許容している差分、たとえば超越関数の ulp 差、揮発関数、Excel 側の不整合を Formulon が意図的に採らないケースは、[`tests/divergence.yaml`](tests/divergence.yaml) に理由と確認済み Excel ビルドを記録します。
- **C++ コア 1 本で動きます。** ブラウザ、Python、CLI のために別々の計算ロジックを持たず、同じエンジンを配布します。実装が複数に分かれて結果がずれる、という問題を避けています。
- **WASM サイズに上限を置いています。** 目標は 1.65 MB、Brotli 圧縮で 530 KB。ハード上限は 1.9 MB / 600 KB です。CI で上限を確認し、超える機能は入れません。
- **依存は小さく保っています。** ランタイム依存は `miniz` (zip/deflate)、`pugixml` (XML + XPath 1.0)、`PCRE2` (`REGEX*`)、`double-conversion` (Grisu3 `dtoa`) の 4 つです。線形代数、UTF-8 処理、数値変換の多くはリポジトリ内で実装しています。
- **監査しやすい C++ を優先します。** `Expected<T, Error>` ベースのエラー処理、RAII、`-fno-exceptions -fno-rtti`、Google C++ Style を採用しています。

## 使いどころ

Excel を起動せずにスプレッドシートを計算したい場面で使えます。

- バッチジョブやデータパイプラインで `.xlsx` をヘッドレス再計算する
- Web アプリの中で Excel 風の数式を評価する
- 社内ツール、ボット、ノートブックに計算機能を組み込む
- 数式の検証や、レガシースプレッドシートの移行に使う

## やらないこと

Formulon は以下を **意図的にサポートしません**。

| 項目 | 理由 |
|------|------|
| VBA の実行 | セキュリティ上の理由。`vbaProject.bin` はバイト列としてだけ保存し、実行はしません。 |
| 旧 `.xls` (BIFF8 / Excel 97-2003) | Excel 365 互換というスコープの外。 |
| グラフ / 図形のレンダリング | 描画レイヤの責務。計算エンジンの仕事ではありません。 |
| PowerQuery (M) / DAX | 別エンジン・別問題。 |
| Pivot キャッシュの再計算 | 構造は保持しますが、再計算は対象外。 |
| スプレッドシート UI 本体 | 薄い UI 統合 API は計画していますが、描画自体は呼び出し側の責任です。 |

これらは「まだやっていない」機能ではなく、スコープ外として固定しています。

## パッケージ

| 配布元 | パッケージ名 | 内容 |
|--------|-------------|------|
| npm | [`@libraz/formulon`](https://www.npmjs.com/package/@libraz/formulon) | WASM ESM モジュール。型定義同梱。Node 22+ / ブラウザ / Worker 対応。 |
| PyPI | [`formulon`](https://pypi.org/project/formulon/) | Python 3.9+ の `py3-none-any` wheel。`formulon_capi.wasm` と pure-Python wrapper を同梱し、`wasmtime` は `pip` が解決します。 |
| GitHub Releases | `formulon-cli-<platform-arch>` | 単体 CLI バイナリ (`eval` / `recalc` / `dump`)。`darwin-arm64` / `linux-x64` / `linux-arm64` 向け。 |

## ステータス

カタログ済み Excel 関数は **522 / 522 を認識**します。ただし、「関数名を知っている」ことと「Excel 互換の実装がある」ことは分けて扱います。現在の内訳は `make function-status` で確認できます。

上位 2 区分（実装済み・unavailable stub）は排他的で、合計は 522 に一致します。環境依存の行は**実装済み 507 の内訳（部分集合）**であり、固定 golden で完全に記述できないため個別に示しているもので、追加のカテゴリではありません（507 + 15 = 522。524 ではありません）。

| 区分 | 件数 | 意味 | 例 |
|------|------|------|----|
| 実装済み | 507 | 通常の計算エンジン内で評価できる関数。unit / oracle で検証しています。 | 数学、統計、検索、テキスト、動的配列など |
| &nbsp;&nbsp;↳ うち環境依存 | 2 | 実装済みだが、ホスト環境やワークブック状態によって値が変わるため固定 golden だけでは完全に記述できない関数。上記 507 に含まれます。 | `INFO`, `CELL` |
| unavailable stub | 15 | Formulon が内蔵しない外部サービス、ネットワーク、COM、OLAP 接続などが必要な関数。固定のエラー面だけを返します。 | `PY`, `WEBSERVICE`, `STOCKHISTORY`, `IMAGE`, `RTD`, `TRANSLATE`, `DETECTLANGUAGE`, `COPILOT`, `CUBE*` |

oracle は **92 カテゴリ** あります。primary oracle は Mac Excel 365 ja-JP から再生成し、Windows Excel 365 ja-JP は `win-365-ja_JP` variant golden でカバーしています。

現在のローカル検証結果:

- fast test: `14342/14342` passed
- primary formula oracle: `4026/4026` passed / `166` documented skips
- workbook oracle (pivot + print): `35/41` passed / `6` documented skips (win-365-ja_JP scoped)

残っている skip は、明示済みの divergence、ホストサービス依存、揮発・環境依存ケース、またはドライバ制約です。黙って未実装 stub に落としているものではありません。522 関数のうち `515` は closure 6 条件 (`behaviors_declared` / `cases_cover_behaviors` / `golden_present` / `divergence_documented` / `not_in_pilot` / `behavior_drift`) を全て満たします。残る `7` 件 (`FILTERXML`, `ARRAYTOTEXT`, `CONCAT`, `CHAR`, `TRUE`, `GETPIVOTDATA`, `PHONETIC`) は oracle metadata 側の課題 — Excel での golden 採取がまだ、または behavior taxonomy が未整備 — であり、実装と Excel の出力が不一致というわけではありません。

数式の結果に加えて、**ピボットテーブルと印刷範囲・改ページ**には専用の **workbook oracle track** があります。信頼できる PivotTable 自動化には Windows Excel の COM が必要なため、この track の primary は `win-365-ja_JP` です。golden 採取は完了しており、pivot suite は `28/28` closure、`print_basic` / `print_pagination` / `print_fit` / `print_matrix` の 4 suite が `formulon_workbook_oracle_tests` ハーネスで `35/41` ケース pass。残り `6` ケースは Excel の PageBreakPreview に `PageSetup.Zoom <= 50` で発生する既知の COM quirk (低スケールで列の auto-break が直感に逆らって増える挙動。VBA コミュニティで documented) のため `win-365-ja_JP` scope で divergence-skip 登録しています。

新規ワークブックはデフォルトで `win-365-ja_JP` profile を使います。必要に応じて profile-id API (`mac-365-ja_JP` / `win-365-ja_JP`) で切り替えられます。英語ロケール profile は、対応する EN oracle データとロケール固有挙動の検証が揃うまで公開しません。

実装面では、バイトコードコンパイラとスタックマシン VM が tree-walker と並列に動作し、parity を検証しています。OOXML reader / writer はシート、スタイル、条件付き書式、コメント、ハイパーリンク、結合セル、入力規則、定義済み名前、テーブル、ピボットテーブルを round-trip します。MS-XLSB reader / writer も実装済みです。シート追加、リネーム、移動、数式書き換えを伴う行・列の挿入 / 削除、partial recalc、反復計算ソルバの進捗コールバックは、C ABI 経由で WASM / Python / CLI から使えます。

不具合報告・oracle 差分レポート・ご意見はいつでも歓迎しています。

## コントリビューション

いちばん助かるのは、**手元の Excel から oracle データを提供していただくこと**です。Mac ja-JP 以外の Excel 365 をお持ちなら、`make oracle-contribute` で Excel を駆動して golden を取得し、PR の手順まで進められます。詳しくは [CONTRIBUTING.md](CONTRIBUTING.md) を参照してください。

## ライセンス

Apache License 2.0。[LICENSE](LICENSE) および [NOTICE](NOTICE) を参照してください。
