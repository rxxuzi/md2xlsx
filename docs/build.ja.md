# 実行ファイル（exe）にする

[English](build.md)

`md2xlsx.py` は依存 1 つの単一スクリプトなので、単体で動くバイナリに
簡単にできる。バイナリを使う側には Python も `openpyxl` も要らない。

## PyInstaller（おすすめ）

このリポジトリを clone した中で:

```sh
uv run --with pyinstaller pyinstaller --onefile --name md2xlsx md2xlsx.py
```

`uv run` が `pyproject.toml` から `openpyxl` を `.venv/` に入れ、その上に
PyInstaller を足してくれるので、追加の設定は不要。uv を使わない場合:

```sh
pip install openpyxl pyinstaller
pyinstaller --onefile --name md2xlsx md2xlsx.py
```

`dist/md2xlsx.exe`（Windows）または `dist/md2xlsx`（macOS / Linux）が
できる。Windows 11、Python 3.12、PyInstaller 6.22 での実測:

| | |
|---|---|
| ビルド時間 | 約 25 秒 |
| サイズ | 約 8 MB |
| 起動 | 約 1.4 秒（`--onefile` は実行のたびに一時ディレクトリへ展開するため） |

使い方はスクリプトと同じ:

```sh
dist/md2xlsx.exe input.md
dist/md2xlsx.exe input.md -o out.xlsx
```

注意点:

- バイナリはビルドした OS でしか動かない。Windows 用 `.exe` は Windows で、
  macOS 用は macOS で作る。3 つまとめて作るなら下の CI の例を参照。
- Windows Defender / SmartScreen が PyInstaller の `--onefile` バイナリを
  未知のソフトとして警告することがある。配布するならコード署名が本当の
  解決策。`--onedir`（単一ファイルではなくフォルダ）にすると警告されにくく、
  起動も速い。
- PyInstaller は `build/`、`dist/`、`md2xlsx.spec` を作る。`dist/` は
  `.gitignore` 済みなので、リポジトリ内でビルドするなら `build/` と
  `*.spec` も足すこと。

## Nuitka

スクリプトを C にコンパイルする。起動が速く、少し小さくなるが、C コンパイラ
（Windows なら MSVC か MinGW。Nuitka が MinGW のダウンロードを提案してくれる）
が必要で、ビルドに数分かかる。

```sh
uv run --with nuitka python -m nuitka --onefile --output-filename=md2xlsx md2xlsx.py
```

## 自分のマシンで `md2xlsx` コマンドが欲しいだけなら

バイナリは要らない:

```sh
uv tool install .          # clone した中で
uv tool install git+https://github.com/rxxuzi/md2xlsx
```

`PATH`（`~/.local/bin`）に `md2xlsx` のランチャーが入る。裏では uv 管理の
Python が動いているので、自分用であって配布用ではない。

## GitHub Actions で全 OS 分をビルドする

未検証のたたき台。`v0.2.0` のようなタグを push すると各 OS でバイナリを
ビルドし、GitHub Release に添付する。

```yaml
# .github/workflows/release.yml
name: release
on:
  push:
    tags: ["v*"]
permissions:
  contents: write
jobs:
  build:
    strategy:
      matrix:
        os: [windows-latest, macos-latest, ubuntu-latest]
    runs-on: ${{ matrix.os }}
    steps:
      - uses: actions/checkout@v4
      - uses: astral-sh/setup-uv@v6
      - run: uv run --with pyinstaller pyinstaller --onefile --name md2xlsx-${{ runner.os }} md2xlsx.py
      - uses: softprops/action-gh-release@v2
        with:
          files: dist/*
```

`runner.os` は `Windows` / `macOS` / `Linux` に展開されるので、3 つの
バイナリが別名になり Release 上で上書きし合わない。使う前に各 action の
最新バージョンを確認すること。
