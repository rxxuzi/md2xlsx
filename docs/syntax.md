# Syntax reference

[日本語版](syntax.ja.md)

A complete description of the Markdown dialect `md2xlsx.py` understands.
For a file that uses every feature at once, see
[`examples/sample.md`](../examples/sample.md).

## Sheets

```markdown
#[Sales]
...content...

#[Tasks]
...content...
```

Each `#[name]` starts a new sheet. Content before the first `#[...]` (or a
file with no `#[...]` at all) goes to `Sheet1`. Names longer than 31
characters are truncated, since that is Excel's limit.

## Directives

Place directly under `#[name]`, one per line. Blank lines between them are
fine. Directives anywhere else are treated as plain text.

| Directive | Effect |
|---|---|
| `@freeze(N)` | Freeze the top N rows |
| `@filter` | Auto-filter on the used range |
| `@style(A=10, B=30 yellow)` | Column styles, see below |
| `@document` | Document mode, see below |

`@filter` always covers `A1` to the last used cell, so put the table at the
top of the sheet when you use it; a heading above the table would become
part of the filter range.

### Column styles

`@style` takes comma-separated `COLUMN=values` pairs. The values are
space-separated and recognised by shape, like a CSS shorthand: a number is
the width, a color name or `#RRGGBB` the background, `bold` / `italic` the
font, `left` / `center` / `right` the alignment.

```markdown
@style(A=6 center, B=30, D=14 lightyellow bold)
```

Unlisted columns auto-size. How this interacts with header rows and cell
decorators is in [Styling](style.md).

Auto-sized columns are clamped to 8–60 characters wide. CJK characters
count as two.

## Tables

### Pipe tables

```markdown
| Item | Qty |
|------|-----|
| Apple | 3 |
```

Any line containing `|` (and not starting with `#`) is a table row. The
`|---|` separator line marks the row *above* it as the header (bold, white
on blue, centered). A table without a separator line has no header row. A
blank line ends the table.

### CSV / TSV tables

For people who don't want to draw pipes, or are pasting from a spreadsheet.
The first row is always the header. The delimiter is auto-detected (tab if
the block has more tabs than commas); `table:csv` / `table:tsv` forces one.

````markdown
```table
Item,Qty
Apple,3
```
````

Quoting follows Python's `csv` module, so `"a,b"` is one cell.

## Cell values

Applied to every cell, whether in a table or a plain line.

| You write | You get |
|---|---|
| `1200`, `3.14` | Number |
| `$$=SUM(B2:B5)$$` | Live Excel formula |
| `$1,200` | Number with `$#,##0` format |
| `¥1,200` | Number with `¥#,##0` format (`￥` also works) |
| `30%` | `0.3` with `0.0%` format |
| `2024-01-15` | Real date cell (`YYYY-MM-DD` only) |
| `[Label](https://...)` | Hyperlink (works inside longer text too) |
| `**bold**` `*italic*` `` `code` `` | Font style on the cell |

Notes:

- A formula must be the whole cell. Excel cannot mix literal text and a
  formula in one cell, so `Total: $$=SUM(...)$$` stays text; use two cells.
- Formulas are written as-is and not evaluated. Excel computes them on
  open. Previewers that don't recalculate (GitHub, some mobile viewers) may
  show those cells as empty until the file is opened in Excel or
  LibreOffice.
- Row numbers in formulas count from the sheet's first row. A `##` heading
  above a table shifts the table down by one.
- `1,200` without a currency sign stays text. `$1,200.50` is stored as
  `1200.5` but displayed without decimals.
- Cell types are exclusive: a formula or currency cell does not get inline
  bold/italic, and a link cell does not get numeric typing.
- `` `code` `` uses a monospace font only when the cell has no bold or
  italic.

## Cell decorators

Prefix a cell's text with one or more `{key:value}` tags.

```markdown
| {bg:lightyellow}{comment:started 2024-02-01}Design | {dropdown:todo,doing,done}doing |
```

| Tag | Effect |
|---|---|
| `{bg:yellow}` | Background fill |
| `{fg:red}` | Font color |
| `{comment:text}` | Cell comment |
| `{dropdown:A,B,C}` | Data-validation list |
| `{align:right}` | `left` / `center` / `right` |

Colors: `red green blue yellow orange purple pink gray grey white black
lightblue lightgreen lightred lightyellow`, or any `#RRGGBB`.

Dropdown options are comma-separated, so an option cannot itself contain a
comma.

## Headings and text

In the default mode:

| Markdown | Result |
|---|---|
| `## Title` | Bold 13pt row in column A |
| `### Sub` | Bold 11pt row in column A |
| plain line | Text in column A |
| `- item` / `* item` | Text in column A, marker dropped |
| `1. item` | Text in column A, number kept |
| `---` | Blank row |

A single `# Title` is **not** a heading in the default mode; it is written
to the sheet literally. `#` is reserved for `#[sheet]` and for document
mode.

## Document mode

`@document` turns the sheet into a column-indented outline with a
full grid border, the way specification documents are usually laid out in
Excel in Japan. Headings give the structure: the number of `#` selects
the column (`#` → A, `##` → B, `###` → C, …). Plain lines and list items
are the body: they go one column to the right of the last heading, and
each two spaces of indent on a list item move it one more column.
Nothing is auto-numbered, write your own.

```markdown
#[Spec]
@document

# 1. Functional
## 1.1 Login
The user signs in with an ID and a password.
- Enter user ID
- Enter password
  - masked
# 2. Non-functional
## 2.1 Performance
Response under 3s.
```

| A | B | C | D |
|---|---|---|---|
| 1. Functional | | | |
| | 1.1 Login | | |
| | | The user signs in with an ID and a password. | |
| | | Enter user ID | |
| | | Enter password | |
| | | | masked |
| 2. Non-functional | | | |
| | 2.1 Performance | | |
| | | Response under 3s. | |

`#` and `##` are bold, deeper headings and all body text are regular
weight. Tables still work inside a document-mode sheet; they start at
column A and get the grid border instead of the default row underline.
