# Formulon

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon/actions/workflows/ci.yml)
[![npm](https://img.shields.io/npm/v/@libraz/formulon)](https://www.npmjs.com/package/@libraz/formulon)
[![PyPI](https://img.shields.io/pypi/v/formulon)](https://pypi.org/project/formulon/)
[![codecov](https://codecov.io/gh/libraz/formulon/branch/main/graph/badge.svg)](https://codecov.io/gh/libraz/formulon)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon/blob/main/LICENSE)
[![C++17](https://img.shields.io/badge/C%2B%2B-17-blue?logo=c%2B%2B)](https://en.cppreference.com/w/cpp/17)
[![Platform](https://img.shields.io/badge/platform-Linux%20%7C%20macOS%20%7C%20WebAssembly-lightgrey)](https://github.com/libraz/formulon)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-blue)](https://formulon.libraz.net)

**Formulon は Excel 互換の計算エンジンです。** C++17 で書かれたコアエンジンが、Excel 365 (ja-JP) の挙動を bit 単位で再現することを目指しています。同じエンジンをブラウザ (WebAssembly)、Python、ネイティブ CLI から呼び出せるので、どの環境でもワークブックは同じ値に再計算されます。

Excel 本体・Microsoft ランタイム・COM オートメーションは不要です。WASM はブラウザ、Node、Python では `wasmtime` 経由で動作し、ネイティブ CLI パッケージは `darwin-arm64` / `linux-x64` / `linux-arm64` を対象に配布しています。

## インストール

```bash
npm install @libraz/formulon   # JavaScript / TypeScript (WASM)
pip install formulon            # Python
```

CLI バイナリは [GitHub Releases](https://github.com/libraz/formulon/releases) から取得できます。

## 特徴

- **「目標」ではなく、検証済みの互換性。** デフォルトのプロファイルは `win-365-ja_JP`、一次 oracle は Mac Excel 365 (ja-JP)、Windows Excel 365 (ja-JP) もバリアント golden でカバーしています。出力は実 Excel から再生成した golden と bit 単位で照合され、許容している差分 (超越関数の ulp 差、揮発関数のスナップショット、Formulon が意図的に Excel 側の不整合を訂正しているケース) は [`tests/divergence.yaml`](tests/divergence.yaml) に理由と最終確認 Excel ビルドつきで全件記録しています。
- **C++ コア 1 本、どこでも同じ結果。** JS 系の競合はブラウザ用とサーバ用にロジックを二重に持ちがちですが、Formulon は WASM / Python / CLI のすべてに同じエンジンを配ります。二つめの実装が drift する余地がありません。
- **厳格な WASM サイズ予算。** 目標 1.65 MB (Brotli 530 KB)、ハード上限 1.9 MB (Brotli 600 KB)。CI でハード上限を強制し、超える機能は載せません。
- **小さな依存。** ランタイム依存は `miniz` (zip/deflate)、`pugixml` (XML + XPath 1.0)、`PCRE2` (`REGEX*` 用の Excel 互換方言)、`double-conversion` (Grisu3 最短往復 `dtoa`) の 4 つのみ。線形代数、UTF-8 処理、数値変換の大半は内製です。
- **読める / 監査できるコード。** `Expected<T, Error>` ベースのエラー処理、RAII、`-fno-exceptions -fno-rtti`、Google C++ Style。

## 使いどころ

Excel を起動せずにスプレッドシートを計算したい、すべての場面で:

- バッチジョブやデータパイプラインで `.xlsx` をヘッドレスに再計算する
- Web アプリ内 (ブラウザ) で Excel 風の数式を評価する
- 社内ツール・ボット・ノートブックに計算機能を組み込む
- 数式の検証、レガシースプレッドシートの移行

## やらないこと (恒久的 non-goal)

Formulon は以下を **意図的にサポートしません**:

| 項目 | 理由 |
|------|------|
| VBA の実行 | セキュリティ上の理由。`vbaProject.bin` はバイト列としてだけ保存し、実行はしません。 |
| 旧 `.xls` (BIFF8 / Excel 97–2003) | Excel 365 互換というスコープの外。 |
| Chart / Drawing のレンダリング | 描画レイヤの責務。計算エンジンの仕事ではありません。 |
| PowerQuery (M) / DAX | 別エンジン・別問題。 |
| Pivot キャッシュの再計算 | 構造は保持しますが、再計算は対象外。 |
| スプレッドシート UI 本体 | 薄い UI 統合 API は計画していますが、描画自体は呼び出し側の責任です。 |

これらは「まだやっていない」ではなく **恒久的** non-goal です。スコープは意図的に有限です。

## パッケージ

| 配布元 | パッケージ名 | 内容 |
|--------|-------------|------|
| npm | [`@libraz/formulon`](https://www.npmjs.com/package/@libraz/formulon) | WASM ESM モジュール (型定義同梱)。Node 18+ / ブラウザ / Worker 対応。 |
| PyPI | [`formulon`](https://pypi.org/project/formulon/) | Python 3.9+ の `py3-none-any` wheel。`formulon_capi.wasm` と pure-Python wrapper を同梱し、platform-specific runtime は `pip` が `wasmtime` として解決します。 |
| GitHub Releases | `formulon-cli-<platform-arch>` | 単体 CLI バイナリ (`eval` / `recalc` / `dump`)。`darwin-arm64` / `linux-x64` / `linux-arm64` 向け。 |

## ステータス

カタログ済み Excel 関数 **522 / 522 を認識**します。ただし、関数名を認識することと Excel 互換の実処理を持つことは分けています。関数カタログは availability を明示し、実装済み、実装済みだが検証未完了、環境依存、意図的な unavailable stub を区別します。unavailable stub は、Formulon が内蔵しないホストサービスや接続を必要とする関数に限定しています。例: PY、WEBSERVICE、STOCKHISTORY、IMAGE、RTD、TRANSLATE、DETECTLANGUAGE、COPILOT、CUBE* OLAP ファミリ。現在の内訳は `make function-status` で確認できます。

oracle は **92 カテゴリ** を定義し、Mac Excel 365 ja-JP から再生成。Windows Excel 365 ja-JP は `win-365-ja_JP` バリアント golden でカバーしています。新規ワークブックはデフォルトで `win-365-ja_JP` profile を使い、profile-id API (`mac-365-ja_JP` / `win-365-ja_JP`) で切替可能です。英語ロケール profile は、対応する EN oracle データとロケール固有挙動の検証が揃うまで意図的に未公開です。

実装面では、バイトコードコンパイラとスタックマシン VM が tree-walker と並列に動作し、parity 検証を行っています。OOXML reader / writer は sheets / styles / 条件付き書式 / コメント / ハイパーリンク / 結合セル / 入力規則 / 定義済み名前 / テーブル / ピボットテーブルを round-trip します。MS-XLSB reader / writer も実装済み。ワークブック操作 (シート追加 / リネーム / 移動、数式書き換えを伴う行・列の挿入 / 削除、partial recalc、反復計算ソルバの進捗コールバック) は C ABI を経由して WASM / Python / CLI のすべての surface に露出しています。

不具合報告・oracle 差分レポート・ご意見はいつでも歓迎しています。

## コントリビューション

いちばん助かるのは、**手元の Excel から oracle データを寄贈していただくこと**です。Mac ja-JP 以外の Excel 365 をお持ちなら `make oracle-contribute` 一発で Excel を駆動して golden を取得し、PR の手順までガイドします。詳細とコミュニティ駆動である理由は [CONTRIBUTING.md](CONTRIBUTING.md) を参照してください。

## ライセンス

Apache License 2.0。[LICENSE](LICENSE) および [NOTICE](NOTICE) を参照してください。
