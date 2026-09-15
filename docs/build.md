# Building a standalone executable

[日本語版](build.ja.md)

`md2xlsx.py` is a single script with one dependency, so it packages into a
self-contained binary easily. Users of that binary need neither Python nor
`openpyxl`.

## PyInstaller (recommended)

From a clone of this repo:

```sh
uv run --with pyinstaller pyinstaller --onefile --name md2xlsx md2xlsx.py
```

`uv run` installs `openpyxl` from `pyproject.toml` into `.venv/` and adds
PyInstaller on top, so no extra configuration is needed. Without uv:

```sh
pip install openpyxl pyinstaller
pyinstaller --onefile --name md2xlsx md2xlsx.py
```

The result is `dist/md2xlsx.exe` (Windows) or `dist/md2xlsx` (macOS /
Linux). Measured on Windows 11 with Python 3.12 and PyInstaller 6.22:

| | |
|---|---|
| Build time | ~25 s |
| Binary size | ~8 MB |
| Startup | ~1.4 s (`--onefile` unpacks to a temp dir on every run) |

Run it exactly like the script:

```sh
dist/md2xlsx.exe input.md
dist/md2xlsx.exe input.md -o out.xlsx
```

Things to know:

- A binary only runs on the OS it was built on. Build on Windows for a
  Windows `.exe`, on macOS for a macOS binary, and so on. See the CI recipe
  below for building all three at once.
- Windows Defender / SmartScreen sometimes flags `--onefile` binaries from
  PyInstaller as unknown software. If you distribute the file, code-signing
  it is the real fix; `--onedir` (a folder instead of a single file) is
  flagged less often and also starts faster.
- PyInstaller writes `build/`, `dist/` and `md2xlsx.spec`. `dist/` is
  already in `.gitignore`; add `build/` and `*.spec` if you build in the
  repo.

## Nuitka

Compiles the script to C. The binary starts faster and is a bit smaller,
but you need a C compiler (MSVC or MinGW on Windows; Nuitka offers to
download MinGW) and the build takes several minutes.

```sh
uv run --with nuitka python -m nuitka --onefile --output-filename=md2xlsx md2xlsx.py
```

## Just want a `md2xlsx` command on your own machine?

You don't need a binary for that:

```sh
uv tool install .          # from a clone
uv tool install git+https://github.com/rxxuzi/md2xlsx
```

This puts a `md2xlsx` launcher on your `PATH` (`~/.local/bin`). It still
uses a uv-managed Python under the hood, so it is for your machine, not for
handing to someone else.

## Building for all platforms with GitHub Actions

An untested starting point. Pushing a tag like `v0.2.0` builds a binary on
each OS and attaches them to a GitHub Release.

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

`runner.os` expands to `Windows`, `macOS` or `Linux`, so the three binaries
get distinct names and don't overwrite each other on the release. Check the
action versions against their current releases before using this.
