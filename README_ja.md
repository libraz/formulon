# Formulon

> **ステータス: 開発中。本番利用にはまだ適しません。**
> 最初のタグ付きリリースまで、API・ファイル構成・パッケージングは予告なく変更される可能性があります。

Formulon はヘッドレスで動く Excel 互換の計算エンジンです。C++17 製のコア
が **Mac Excel 365 (ja-JP) に対して bit 単位で一致する**ことを目標とし、
既知の差分はすべて明示的に追跡されています。同じエンジンをブラウザ
(WebAssembly)・Python・ネイティブ CLI 向けにパッケージ配布し、どの環境
でもワークブックが同じ値に再計算される、という一点を保証することを目指
します。

Excel のインストール、Microsoft ランタイム、COM オートメーションは一切
不要です。macOS / Linux / Windows / ブラウザ / Node で動作します。

## なぜ Formulon か

- **目標ではなく、検証済みの互換性。** Mac Excel 365 (ja-JP) を
  behavioral oracle として扱います。出力は実製品から再生成した golden
  データと bit 単位で照合され、許容している差分 (超越関数の ulp 差、
  揮発関数のスナップショット、Formulon が意図的に Excel 側の不整合を
  訂正しているケースなど) は
  [`tests/divergence.yaml`](tests/divergence.yaml) にケース単位で理由と
  最終確認 Excel バージョンつきで記録されています。
- **C++ コア 1 本、どこでも同じ結果。** JavaScript 系の競合はブラウザ用
  のロジックとサーバ用のロジックを二重に持ちがちですが、Formulon は全
  surface (WASM / Python / CLI) に同じエンジンを配ります。二つめの実装
  が drift する余地がありません。
- **厳格な WASM サイズ予算。** 目標 **1.65 MB (Brotli 530 KB)**、
  上限 **1.8 MB (Brotli 600 KB)**。CI で強制され、守れない機能は載せ
  ません。
- **小さな依存。** エンジンのランタイム依存は `miniz` (zip) と `pugixml`
  (XML + XPath 1.0) の 2 個のみ。線形代数・数値フォーマット・UTF-8 処理
  は内製です。
- **読めて監査できるコード。** `Expected<T, Error>` ベースのエラー処理、
  RAII、`-fno-exceptions -fno-rtti`、Google C++ Style。

## 想定ユースケース

Excel を起動せずにスプレッドシートを計算したい場面:

- バッチジョブやデータパイプラインで `.xlsx` をヘッドレスに再計算する
- Web アプリ内 (ブラウザ) で Excel 風の数式を評価する
- 社内ツール・ボット・ノートブックに計算機能を埋め込む
- 数式の検証やレガシースプレッドシートの移行

## スコープ外 (恒久的 non-goals)

Formulon は以下を **意図的にサポートしません**:

| 項目 | 理由 |
|------|------|
| VBA の実行 | セキュリティ。`vbaProject.bin` はバイト列として保持するだけで、実行はしません。 |
| 旧 `.xls` (BIFF8、Excel 97–2003) | Excel 365 互換というスコープ外。 |
| Chart・Drawing の画像化 | 描画レイヤの責務。計算エンジンの仕事ではありません。 |
| PowerQuery (M) / DAX | 別エンジン・別問題。 |
| Pivot キャッシュの再計算 | 構造は保持しますが、再計算は対象外。 |
| スプレッドシート UI 本体 | 薄い UI 統合レイヤは計画していますが、描画は利用側の責任です。 |

これらは「まだやっていない」ではなく **恒久的** non-goals です。スコープ
は意図的に有限にしています。

## パッケージ

| 媒体 | 名前 | 備考 |
|------|------|------|
| npm | `@libraz/formulon` | WASM ESM モジュール、型定義同梱。Node 18+ / ブラウザ / Worker 対応。 |
| PyPI | `formulon` | macOS / Linux / Windows 向け CPython 3.9–3.13 wheel。実行時は ctypes と標準ライブラリのみに依存。 |
| GitHub Releases | `formulon-cli-<os>-<arch>` | 単体 CLI バイナリ (`eval` / `recalc` / `dump`)。 |

## ステータス

2026-05 時点: **カタログ済みの Excel 関数 522 / 522 を実装済み (100%)**。
Math & Trig、統計、論理、文字列、日付時刻、ルックアップ、財務、エン
ジニアリング、情報、データベース、キューブ、および 2024/2025 追加分
(GROUPBY、PIVOTBY、TRANSLATE、COPILOT など) を網羅しています。**92
カテゴリの oracle** を定義し、Mac Excel 365 ja-JP から再生成しています。
バイトコードコンパイラとスタックマシン VM が tree-walker と並列に動作し、
parity 検証を行っています。OOXML reader / writer はシート、スタイル、
条件付き書式、コメント、ハイパーリンク、結合セル、入力規則、定義済み
名前、テーブル、ピボットテーブルを round-trip し、MS-XLSB の reader /
writer も実装済みです。ワークブックレベルの操作 (シートの追加 / リネー
ム / 移動、数式書き換えを伴う行・列の挿入 / 削除、partial recalc、
反復計算ソルバの進捗コールバック) は C ABI を経由して WASM / Python /
CLI 各 surface に露出しています。

フィードバック・不具合報告・オラクル差分レポートは歓迎しますが、現時点
で本番ワークロードでの利用はお控えください。

## コントリビューション

現時点でいちばん助かるのは **手元のロケールの Excel oracle データを
寄贈していただくこと**です。Mac ja-JP 以外の Excel 365 が手元にあれば、
`make oracle-contribute` ひとつで Excel を駆動して golden を取得し、
PR まで案内します。詳細とコミュニティ駆動である理由は
[CONTRIBUTING.md](CONTRIBUTING.md) を参照してください。

## ライセンス

Apache License 2.0。[LICENSE](LICENSE) および [NOTICE](NOTICE) を参照し
てください。
