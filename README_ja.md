# Formulon

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon/actions/workflows/ci.yml)
[![npm](https://img.shields.io/npm/v/@libraz/formulon)](https://www.npmjs.com/package/@libraz/formulon)
[![PyPI](https://img.shields.io/pypi/v/formulon)](https://pypi.org/project/formulon/)
[![codecov](https://codecov.io/gh/libraz/formulon/branch/main/graph/badge.svg)](https://codecov.io/gh/libraz/formulon)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon/blob/main/LICENSE)
[![C++17](https://img.shields.io/badge/C%2B%2B-17-blue?logo=c%2B%2B)](https://en.cppreference.com/w/cpp/17)
[![Platform](https://img.shields.io/badge/platform-Linux%20%7C%20macOS%20%7C%20WebAssembly-lightgrey)](https://github.com/libraz/formulon)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-2563eb)](https://formulon.libraz.net)

**Formulon は、Excel なしで Excel のワークブックと数式を再計算するエンジンです。結果は実際の Excel 365 と照合しています。**
Excel のインストールも Windows も COM 操作も要りません。1 つの C++17 コアを、ブラウザと Node 向けの WebAssembly、Python パッケージ、ネイティブ CLI として配布しているので、どこで動かしても同じワークブックは同じ値になります。

**こんなときに使えます:**

- **サーバーでワークブックを再計算する** — バッチ処理、CI、データパイプラインで `.xlsx` / `.xlsb` を読み込み、入力を変えて、再計算した値で保存できます。
- **ブラウザで表計算ロジックを動かす** — 数式もワークブック全体もクライアント側で評価するので、アップロードされたファイルが利用者のマシンの外に出ません。
- **計算モデルは Excel のまま使う** — チームが保守している既存のワークブックを Node や Python から呼び出せます。数式を別の言語で書き直す必要はありません。
- **ロケールごとに Excel と同じ答えを得る** — 日本語のバイト数計算、ドイツ語の小数点カンマ、ローカライズされた `TRUE` / `FALSE` は、選んだプロファイルに従います。
- **AI エージェントに表計算の道具を渡す** — [formulon-mcp](https://github.com/libraz/formulon-mcp) がエンジンを MCP 経由で提供します。

Excel との既知の差分は [`tests/divergence.yaml`](https://github.com/libraz/formulon/blob/main/tests/divergence.yaml) にケースごとに記録し、それぞれに理由と最後に確認した Excel ビルドを添えています。

📖 **[ドキュメント](https://formulon.libraz.net/ja/)** &nbsp;·&nbsp; **[はじめに](https://formulon.libraz.net/ja/start/)** &nbsp;·&nbsp; **[ロケールプロファイル](https://formulon.libraz.net/ja/compatibility/locale-profiles)** &nbsp;·&nbsp; **[数式カバレッジ](https://formulon.libraz.net/ja/compatibility/formula-coverage)** &nbsp;·&nbsp; **[API](https://formulon.libraz.net/ja/api/)**

## 主な機能

- **数式エンジン** — Excel 関数名 526 件を認識し、511 件をローカルで実装しています。動的配列、`LET` / `LAMBDA`、`REGEX*` にも対応します。残りはクラウドや COM のサービスを呼ぶ関数で、決まったエラーを返します。[カバレッジ](https://formulon.libraz.net/ja/compatibility/formula-coverage)
- **ワークブック入出力** — `.xlsx` と `.xlsb` を読み込み、再計算し、書き出します。スタイル、条件付き書式、テーブル、ピボットテーブル、印刷レイアウトも扱います。[ファイル形式](https://formulon.libraz.net/ja/compatibility/file-format-support)
- **ロケールプロファイル** — 7 ロケール × Mac / Windows ホストの、14 種類の Excel 挙動プロファイルを選べます。[ロケールプロファイル](https://formulon.libraz.net/ja/compatibility/locale-profiles)
- **Excel oracle** — 数式の結果は、7 ロケールの Mac Excel 365 から取得したゴールデンデータとビット単位で照合しています。ピボットと印刷レイアウトは Windows Excel 365 と照合しています。[Oracle テスト](https://formulon.libraz.net/ja/compatibility/oracle-testing)
- **サイズ上限つきの WASM** — 非圧縮 3.75 MiB、Brotli 1024 KiB を超えると CI が失敗します。[サイズ予算](https://formulon.libraz.net/ja/development/size-budgets)

## インストール

```bash
npm install @libraz/formulon   # ブラウザと Node 22+
pip install formulon            # Python 3.9+
```

`darwin-arm64`、`linux-x64`、`linux-arm64` 向けの CLI バイナリは [GitHub Releases](https://github.com/libraz/formulon/releases) にあります。

## クイックスタート

```js
import createFormulon from '@libraz/formulon';

const Module = await createFormulon();
console.log(Module.evalFormula('=SUM(1,2,3)').value.number); // 6
```

```python
import formulon

print(formulon.eval_formula("=SUM(1,2,3)").to_python())  # 6.0
```

```bash
formulon eval '=SUM(1,2,3)'
formulon recalc input.xlsx -o output.xlsx
```

ワークブックの読み込み、編集、保存は [ワークブックを再計算する](https://formulon.libraz.net/ja/start/recalculate) で説明しています。

## ロケールプロファイルの切り替え

新しいワークブックの既定は `win-365-en_US` です。数式は常に英語の関数名と保存形式の区切り文字で書きます。プロファイルが変えるのは、文字列の解釈と結果の表示です。ワークブックごとに切り替えて再計算します。

```js
const wb = Module.Workbook.createDefault();
try {
  wb.setFormula(0, 0, 0, '=LENB("日本")');  // A1
  wb.setFormula(0, 1, 0, '=VALUE("1,5")');  // A2
  wb.setFormula(0, 2, 0, '=ISEVEN(2)');     // A3

  for (const id of ['mac-365-en_US', 'mac-365-ja_JP', 'mac-365-de_DE']) {
    wb.setExcelProfileId(id);
    wb.recalc(); // a profile change marks formulas dirty
    console.log(id, [0, 1, 2].map((row) => wb.getDisplayText(0, row, 0).text));
  }
} finally {
  wb.delete();
}
// mac-365-en_US [ '2', '#VALUE!', 'TRUE' ]
// mac-365-ja_JP [ '4', '#VALUE!', 'TRUE' ]
// mac-365-de_DE [ '2', '1,5', 'WAHR' ]
```

```python
with formulon.Workbook.create_default() as wb:
    wb.set_formula(0, 0, 0, '=LENB("日本")')
    wb.set_formula(0, 1, 0, '=VALUE("1,5")')
    wb.set_formula(0, 2, 0, "=ISEVEN(2)")

    for profile in ("mac-365-en_US", "mac-365-ja_JP", "mac-365-de_DE"):
        wb.set_excel_profile_id(profile)
        wb.recalc()
        print(profile, [wb.get_display_text(0, row, 0)[0] for row in range(3)])
```

プロファイル ID は `{mac,win}-365-{ja_JP,en_US,de_DE,fr_FR,zh_CN,ko_KR,th_TH,es_ES,es_MX,pt_BR,ru_RU,zh_TW,it_IT,nl_NL}` です。`mac-*` のすべてと `win-365-ja_JP` は Excel で測定したもので、それ以外の `win-*` は Mac での測定値からの推定です。プロファイルはファイルに保存されないため、アプリケーション側で対象の ID を保持し、読み込み後に設定し直してください。

## 対象外

VBA の実行、旧形式の `.xls`、グラフの描画、Power Query / DAX、元データからのピボットキャッシュ更新、外部データへのライブ接続、表計算 UI は、恒久的に対象外です。詳しくは [対象外の機能](https://formulon.libraz.net/ja/compatibility/non-goals) を参照してください。

## コントリビュート

いちばん役立つのは、お使いのロケールの Excel oracle データです。特に Windows Excel 365 のデータを歓迎します。`make oracle-contribute` が Excel を操作してゴールデンデータを取得します。[CONTRIBUTING.md](CONTRIBUTING.md) と [Oracle データの提供](https://formulon.libraz.net/ja/development/oracle-contribution) を参照してください。

## ライセンス

[Apache License 2.0](LICENSE)。[NOTICE](NOTICE) も参照してください。

## 関連プロジェクト

- [formulon-cell](https://github.com/libraz/formulon-cell) — `@libraz/formulon` を使ったブラウザ向け表計算 UI
- [formulon-mcp](https://github.com/libraz/formulon-mcp) — AI エージェントにワークブック操作ツールを提供する MCP サーバー
