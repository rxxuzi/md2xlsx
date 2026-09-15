---
name: md2xlsx
description: Build .xlsx workbooks by writing flavored Markdown and running md2xlsx.py, instead of hand-coding openpyxl. Use this whenever the user wants an Excel file, spreadsheet, workbook, 表, シート, or Excel 仕様書 from text or structured content — multi-sheet reports, task lists with dropdowns, KPI tables with colors, or Japanese-style column-indented specification documents (Excel方眼紙 / 設計書). Also use when the user mentions md2xlsx, "markdown to excel", or asks to convert notes, a plan, or a spec into a spreadsheet. Prefer this over raw openpyxl for any workbook whose content can be described in Markdown.
---

# md2xlsx

Write the workbook as ordinary Markdown with a few extensions, then
convert. This page is the whole dialect plus the layout rules that make
the `.xlsx` come out clean.

```sh
python <skill_dir>/md2xlsx.py input.md -o output.xlsx -d
```

`<skill_dir>` is this folder. `-o` sets the output (default: `input.xlsx`),
`-d` also prints every sheet as a text grid so you can check the layout
without opening the file. Requires `openpyxl`. If it is missing, either
install it (`pip install openpyxl --break-system-packages`) or let `uv`
supply it:

```sh
uv run --with openpyxl <skill_dir>/md2xlsx.py input.md -o output.xlsx -d
```

## Workflow

1. Decide the sheets: one line per sheet with its type (table or
   document, see below) and, for a table, its columns.
2. Write the `.md`, one `#[Sheet]` per sheet.
3. Run `md2xlsx.py … -d`. Mistakes come back rustc-style on stderr with
   the line, a caret under the token and a `help:` line. An `error:` means
   nothing was written — fix it and rerun. A `warning:` still writes the
   file but points at a layout problem; fix it too unless the user asked
   for exactly that.
4. Read the `-d` grid: headings in the right columns, body text one column
   right of its heading, tables starting on row 1, formulas where the
   totals should be.
5. If the file contains `$$...$$` formulas, run the xlsx skill's
   `recalc.py` on the output so cached values exist and errors surface.
6. Present the `.xlsx` to the user.

## Layout

The converter already supplies the look: a blue header row, auto-sized
columns, thin row rules, a full grid in document mode. A clean workbook
adds almost nothing on top of that.

### One shape per sheet

**Table sheet** — a single table, header on row 1, nothing above or
below it. For anything where every row has the same fields: lists,
references, tasks, KPIs, schedules.

```markdown
#[Tasks]
@freeze(1)
@filter

| ID | Task | Owner | Status |
|----|------|-------|--------|
| 1 | API design | Tanaka | {dropdown:todo,doing,done}done |
| 2 | DB design | Suzuki | {dropdown:todo,doing,done}doing |
```

**Document sheet** — `@document`, then write it like a normal Markdown
document: headings for the structure, plain sentences and bullet lists
for the content. Headings become the indent columns (`#` → A, `##` → B,
`###` → C); body text sits in the column right of its heading. For
procedures, explanations, specifications.

```markdown
#[Spec]
@document

# 1. Functional
## 1.1 Login
The user signs in with an ID and a password.
- Enter user ID
- Enter password
  - masked
## 1.2 Dashboard
Shown right after login: KPI summary and notifications.
# 2. Non-functional
## 2.1 Performance
1. Response under 3s
2. 1000 concurrent users
```

Headings are for sections only. A sentence, a step, a note, a command
is body text — a plain line or a `-` item — never a `###`. Two or three
heading levels is normal; a document sheet with no plain lines at all is
wrong.

A document that needs a real table gets a second, table sheet. Mixing
`##` headings and small tables in one default-mode sheet is possible but
rarely reads well in Excel; prefer one combined table with a category
column.

### Keep it plain

- Colors are for data that carries state — status, pass/fail,
  over-budget — and one accent per workbook is enough. Headers are
  already styled; decorators on header cells are overridden.
- Notes are text: another column, a plain line under the heading, or a
  `{comment:…}` on the cell they explain. Not a colored label.
- Every table row is one item. Never leave a cell empty to continue the
  row above; put it in one cell or use a document sheet.
- Widths only when auto-width is clearly wrong. Never on a document
  sheet — column A holds the title.
- `---` only to separate two tables in a default-mode sheet.
- Short sheet names, few sheets.

## Syntax

### File shape

A file is a sequence of sheets: a `#[Name]` line, any directives, then
content. Blank lines separate blocks.

- Content before the first `#[...]` (or a file with none) goes to
  `Sheet1`.
- Sheet names are cut to 31 chars and cannot contain `[ ] : * ? / \`.
- A blank line ends a table.

### Directives

Only recognised directly under `#[Name]` (or at the very top of the
file), one per line. Anywhere else `@...` is plain text.

| Directive | Effect |
|---|---|
| `@freeze(N)` | Freeze the top N rows |
| `@filter` | Auto-filter from `A1` to the last used cell — the table must start on row 1 |
| `@style(...)` | Column styles, below |
| `@document` | Document mode, below |

`@style` takes comma-separated `COLUMN=values`; values are
space-separated and recognised by shape, in any order: a number is the
width, a color name or `#RRGGBB` the background (always background —
font color is only `{fg:…}` per cell), `bold` / `italic` the font,
`left` / `center` / `right` the alignment. Unlisted columns auto-size
(8–60, CJK counts as 2).

```markdown
@style(A=6 center, B=30, D=14 lightgreen)
```

Header rows are never touched by `@style`; a cell's own `{bg:…}` /
`{align:…}` wins over the column. Fill and font apply to used rows only.

### Text, lists and headings

| Write | Default mode | Document mode |
|---|---|---|
| plain line | Text in column A | Text one column right of the last heading |
| `- item` / `* item` | Same, marker dropped | Same, marker dropped; each 2 spaces of indent = one more column |
| `1. item` | Same, number kept | Same, number kept |
| `# Title` | Literal text `# Title` (not a heading) | Bold, column A |
| `## Title` | Bold 13pt, column A | Bold, column B |
| `### Sub` … `###### x` | Bold 11pt, column A | Regular, column C … F |
| `---` | Blank row | Blank row |

Document mode also draws a thin black grid over every used cell. Nothing
is auto-numbered — write `1.`, `1.1` yourself. Heading lines and list
items may contain `|` literally; a plain line containing `|` is a table
row.

### Pipe tables

```markdown
| Item | Qty | Price |
|------|-----|-------|
| Apple | 3 | ¥120 |
```

- Any line containing `|` (not starting with `#`) is a row.
- The `|---|` line marks the row **above** it as the header (white bold
  on blue, centered). No separator → no header row.
- Rows may have different lengths. `:---:` colons are ignored.
- A cell cannot contain `|`, and `\|` does not escape it. Rephrase, or
  move the text to a document-sheet line.

### CSV / TSV tables

For pasted data. Row 1 is always the header. Delimiter auto-detected
(tab vs comma); `table:csv` / `table:tsv` forces it. Python `csv`
quoting, so `"a,b"` is one cell.

````markdown
```table
Item,Qty,Price
Apple,3,¥120
```
````

### Cell values

Every cell — table or plain line — is typed by its text:

| Write | Get |
|---|---|
| `1200` `3.14` `-5` | Number |
| `$$=SUM(B2:B5)$$` | Live formula (must be the *whole* cell) |
| `$1,200` `$1200` | Number, format `$#,##0` |
| `¥1,200` `￥1200` | Number, format `¥#,##0` |
| `30%` `29.5%` | `0.3`, format `0.0%` |
| `2024-01-15` | Real date (exactly `YYYY-MM-DD`; `2024/1/15` stays text) |
| `[Label](https://…)` | Hyperlink, also inside longer text |
| `**bold**` `*italic*` `` `code` `` | Font style on that cell |
| anything else | Text |

- `1,200` with no currency sign is **text**. Write `1200`.
- Currency and percent forms take no sign: `-$50` is text; use `-50`.
- A formula cannot share a cell with text — label and formula go in two
  cells.
- Formula row numbers count from the sheet's row 1, headers included.
  Anything above the table shifts every row by one.
- Typed cells are exclusive: a formula / currency / date / percent cell
  ignores inline `**bold**`, and a link cell is never a number.
- `` `code` `` gives a monospace font only when the cell is not also
  bold or italic.

### Cell decorators

Prefix the cell text with one or more `{key:value}` tags, any order:

```markdown
| {bg:lightyellow}{comment:started 2024-02-01}Design | {dropdown:todo,doing,done}doing |
```

| Tag | Effect |
|---|---|
| `{bg:yellow}` | Background fill |
| `{fg:red}` | Font color |
| `{align:center}` | `left` / `center` / `right` |
| `{comment:text}` | Cell comment |
| `{dropdown:A,B,C}` | Data-validation list; the text after it is the current value |

Colors: `red green blue yellow orange purple pink gray grey white black
lightblue lightgreen lightred lightyellow`, or `#RRGGBB`.

- Dropdown options are comma-separated, so an option cannot contain `,`.
- A decorator needs text after it; an empty cell cannot be styled.
- On a header cell, `{bg}` / `{fg}` / `{align}` are overridden by the
  header look; `{comment}` and `{dropdown}` survive.

### Precedence

Strongest first: table header look → cell decorators / inline markup →
`@style` column → defaults. Column `bold` / `italic` are *added* to a
cell's font, not replacing it.

## Templates

### Report with totals

````markdown
#[売上レポート]
@freeze(1)
@filter

```table
四半期,売上,利益,利益率
Q1,$1200,¥360,30%
Q2,$1450,¥420,29%
Q3,$1380,¥400,29%
Q4,$1600,¥512,32%
```

| **合計** | $$=SUM(B2:B5)$$ | $$=SUM(C2:C5)$$ | $$=AVERAGE(D2:D5)$$ |
````

The totals row is a separate one-row pipe table right after the fence,
so it has no header and the ranges (`B2:B5`) line up with the data.

### Task list with dropdowns

```markdown
#[タスク管理]
@freeze(1)
@style(A=6 center, B=30)

| ID | タスク | 担当 | ステータス |
|----|--------|------|-----------|
| 1 | API設計 | 田中 | {dropdown:未着手,進行中,完了}完了 |
| 2 | DB設計 | 鈴木 | {dropdown:未着手,進行中,完了}進行中 |
| 3 | {comment:2024-02-01開始}テスト | 佐藤 | {dropdown:未着手,進行中,完了}未着手 |
```

### Schedule with real dates

```markdown
#[スケジュール]
@freeze(1)

| イベント | 日付 | 予算 |
|----------|------|------|
| キックオフ | 2024-01-15 | ¥500,000 |
| レビュー | 2024-03-20 | ¥150,000 |
| リリース | 2024-06-01 | ¥1,200,000 |
```

### Specification

```markdown
#[基本設計書]
@document

# 1. 機能要件
## 1.1 ログイン機能
ユーザー ID とパスワードで認証する。
- ユーザーIDを入力する
  - 半角英数 8 文字以上
- パスワードを入力する
  - マスク表示
## 1.2 ダッシュボード
ログイン直後に表示する画面。KPI サマリと通知一覧を出す。
# 2. 非機能要件
## 2.1 パフォーマンス
1. レスポンスタイム: 3 秒以内
2. 同時接続数: 1000 ユーザー
```

### Procedure

```markdown
#[手順書]
@document

# 環境構築
## 1. 依存関係
Python 3.8 以上が必要。
- pip install openpyxl
## 2. 動作確認
- python md2xlsx.py examples/sample.md
- 出力された sample.xlsx を開く
```

## Rules of thumb

- Formulas: stick to Excel-2007-era functions (`SUM`, `SUMIFS`, `INDEX`,
  `MATCH`, `IFERROR`). Avoid `XLOOKUP`, `FILTER`, `UNIQUE`; LibreOffice
  recalc cannot evaluate them.
- Before writing a formula, count rows from the top of that sheet
  including headings, blank rows and the header row.
- If a table would need an empty cell somewhere, it is a document sheet.
- If a document sheet has no plain lines, its headings are doing the
  body's job — demote them.

Full reference with output samples: `docs/syntax.md`; looks and
precedence in depth: `docs/style.md`; every feature in one file:
`examples/sample.md`.
