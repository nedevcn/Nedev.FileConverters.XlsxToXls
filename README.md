# Nedev.XlsxToXls

A high-performance **XLSX → XLS** converter library for .NET 10 with **zero third-party dependencies**. It reads Office Open XML (`.xlsx`) workbooks and writes Excel 97–2003 binary (`.xls`, BIFF8) using only built-in BCL types.

---

## Features

- **Zero third-party dependencies** — uses only `System.IO.Compression`, `System.Xml`, `System.Buffers`, and core .NET types.
- **Performance-oriented** — `ArrayPool<byte>` for buffers, streaming `XmlReader` for XLSX, `Span<byte>` for BIFF output to minimize allocations.
- **.NET 10** — targets `net10.0`; build from the `src` folder.

---

## API

### Conversion

| Method | Description |
|--------|-------------|
| `XlsxToXlsConverter.Convert(Stream xlsxStream, Stream xlsStream)` | Converts from a readable XLSX stream to a writable XLS stream. |
| `XlsxToXlsConverter.ConvertFile(string xlsxPath, string xlsPath)` | Converts a file to another file by path. |

### Example

```csharp
using Nedev.XlsxToXls;

// Stream-based
using var xlsx = File.OpenRead("input.xlsx");
using var xls = File.Create("output.xls");
XlsxToXlsConverter.Convert(xlsx, xls);

// File-based
XlsxToXlsConverter.ConvertFile("input.xlsx", "output.xls");
```

---

## Supported (Conversion Completeness)

### Workbook & sheets

| Feature | XLSX source | BIFF output |
|---------|-------------|-------------|
| Multiple worksheets | `xl/workbook.xml` + rels | BOUNDSHEET, separate sheet streams |
| Sheet names | `name` on `<sheet>` | Truncated to 31 chars in BIFF |
| Codepage | — | 1252 (Latin) |

### Cell data

| Feature | XLSX source | BIFF output |
|---------|-------------|-------------|
| Numbers | `<v>` with number format | NUMBER |
| Text (shared strings) | `t="s"` + SST | LABELSST / SST + CONTINUE |
| Inline / direct text | `t="str"`, `t="inlineStr"` | LABEL |
| Empty cells | `<c>` without value | BLANK |
| Booleans | `t="b"` | BOOLERR (boolean) |
| Errors | `t="e"` (#DIV/0!, #N/A, etc.) | BOOLERR (error), mapped to BIFF codes |
| Unicode | UTF-8 in XLSX | 16-bit in LABEL / SST |
| Formulas (basic) | `<f>` (formula) + cached `<v>` | FORMULA (+ STRING record for string results); limited parser (refs/areas, basic operators, a few functions) |

### Cell formatting (from `xl/styles.xml`)

| Feature | XLSX source | BIFF output |
|---------|-------------|-------------|
| Fonts | `fonts/font` | FONT |
| Number formats | `numFmts/numFmt` | FORMAT |
| Cell XFs | `cellXfs/xf` | XF (style + cell XFs), cell `s` → XF index |
| Minimum fonts | — | At least 4 fonts ensured |

### Rows, columns & layout

| Feature | XLSX source | BIFF output |
|---------|-------------|-------------|
| Used range | Computed from rows/cells | DIMENSION |
| Default column width | — | DEFCOLWIDTH (8) |
| Column width / visibility | `<col>` (width, hidden) | COLINFO |
| Row height / visibility | `<row>` (ht, hidden) | ROW |
| Merged cells | `<mergeCells>` | MERGEDCELLS |

### Sheet-level settings

| Feature | XLSX source | BIFF output |
|---------|-------------|-------------|
| Freeze panes | `sheetViews/sheetView/pane` | WINDOW2 + PANE |
| Horizontal page breaks | `rowBreaks/brk` | HORIZONTALPAGEBREAKS |
| Vertical page breaks | `colBreaks/brk` | VERTICALPAGEBREAKS |
| Page setup | `pageSetup` (orientation, scale, fitToWidth/Height) | PAGESETUP |
| Margins | `pageMargins` | LEFTMARGIN, RIGHTMARGIN, TOPMARGIN, BOTTOMMARGIN |
| Print area | `definedName` Print_Area / _xlnm.Print_Area in workbook.xml | NAME (Lbl) + ptgArea3D |
| Print titles (rows/cols) | `definedName` Print_Titles / _xlnm.Print_Titles in workbook.xml | NAME (Lbl) + ptgArea3D |

### Hyperlinks, comments & data validation

| Feature | XLSX source | BIFF output |
|---------|-------------|-------------|
| Cell/range hyperlinks (URLs) | `<hyperlink ref="..." r:id="...">` + sheet rels | HYPERLINK (URL moniker) |
| Cell comments (notes) | `commentsN.xml` (authors + commentList) | NOTE + OBJ + TXO/CONTINUE (cell, author, text) |
| Data validation | `dataValidations` / `dataValidation` (sqref, type, formula1/2) | DATAVALIDATIONS + DATAVALIDATION; **list** type with explicit comma-separated list supported (formula as tStr); other types written with flags/ranges/prompt/error strings, simple formulas compiled to RPN when possible |

### Shared string table (SST)

- Large SSTs are split across **SST + CONTINUE** records (BIFF record data &lt; 8224 bytes).

---

## Not supported (current limitations)

- **Formulas (advanced)** — only a subset of Excel formulas is compiled (no full Excel function set; shared formula edge cases may be imperfect).
- **Data validation (advanced)** — explicit lists and simple formulas/ranges are supported; complex formulas/functions and edge cases may not compile even though basic RPN emission is attempted.
- **Conditional formatting** — not implemented.
- **Charts, images, drawings** — not implemented.
- **Threaded comments** — only legacy comments (commentsN.xml) are read.

---

## XLS limits applied

| Limit | Value | Behavior |
|-------|--------|----------|
| Max rows | 65,536 | No truncation; out-of-range may produce invalid BIFF. |
| Max columns | 256 (A–IV) | No truncation. |
| Sheet name length | 31 | Truncated. |

---

## Build

From the repository root:

```bash
cd src
dotnet build
```

Output: `src/bin/Debug/net10.0/Nedev.XlsxToXls.dll`.

---

## Project layout

```
XlsxToXls/
├── src/
│   ├── XlsxToXls.csproj
│   ├── XlsxToXlsConverter.cs   # Public API + BIFF orchestration
│   └── Internal/
│       ├── BiffWriter.cs      # BIFF8 record writing
│       ├── OleCompoundWriter.cs
│       ├── StylesData.cs
│       ├── StylesReader.cs    # xl/styles.xml
│       └── XlsxReader.cs      # XLSX read (sheets, cells, comments, hyperlinks, etc.)
└── README.md
```

---

## License

See repository or package metadata for license terms.
