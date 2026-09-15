# 開発

[English](development.md)

## セットアップ

実行時の依存は `openpyxl` だけ。テストには `pytest`、lint には `ruff`
（設定は `pyproject.toml`、行長 100）を使う。

uv なら事前インストールは不要。以下の各コマンドが初回実行時に必要なものを
`.venv/`（gitignore 済み）へ解決する。pip の場合:

```sh
pip install openpyxl pytest ruff
```

## 実行

```sh
uv run md2xlsx.py examples/sample.md -o out.xlsx
python md2xlsx.py examples/sample.md -o out.xlsx -d   # -d でシートをテキスト表示
```

## テスト

```sh
uv run pytest
uv run pytest tests/test_md2xlsx.py::test_decorators   # 1 件だけ

python -m pytest tests/
```

テストはブラックボックス方式。Markdown 文字列を一時ファイルに書き、変換し、
`load_workbook` で読み戻してセルの値やスタイルを検証する。新しく書くときは
`tests/test_md2xlsx.py` の `build(tmp_path, text)` ヘルパーに倣うこと。
テストはリポジトリのルートを `sys.path` に足して `md2xlsx` を import する
ので、インストール手順は不要。

## Lint

```sh
uv run ruff check .
ruff check .
```

`DTZ001` は意図的に無視している。Excel の日付にはタイムゾーンがないので、
naive な `datetime` が正しい。

## コードの構成

すべて `md2xlsx.py` に入っていて、上から順にざっくり 3 層:

1. **セルのヘルパー。** `put()` がセル書き込みの唯一の入口。`{キー:値}` の
   装飾を剥がし、型付きの形（数式・通貨・日付・パーセント）を試し、次に
   ハイパーリンク、最後にインラインの太字/斜体/code 付きテキストとして
   落とす。この順番のせいで、数式セルは太字にできず、リンクセルは数値に
   ならない。
2. **`Sheet`。** `#[名前]` ごとに 1 つ。現在行を追跡し、パイプ行や CSV 行を
   セルにし、`finish()` で列幅・ドキュメントモードの罫線・ウィンドウ枠の
   固定・オートフィルタを適用する。
3. **`Conv`。** 1 行ずつ処理するステートマシン。```` ```table ```` の
   フェンス、`#[シート]` とそのディレクティブを処理し、本文の各行を
   見出し / テーブル行 / 空行 / テキストに振り分ける。

## 文法を変えるとき

文法の説明は複数の場所にあり、同期を保つ必要がある:

- `docs/syntax.md` と `docs/syntax.ja.md` — リファレンス
- `SKILL.md` — Claude が読むチートシート
- `examples/sample.md` — 全機能を使う例。コミット済みの `examples/sample.xlsx`
  は `python md2xlsx.py examples/sample.md` で再生成する
- `tests/test_md2xlsx.py` — ケースを追加する
