# Spreadsheet

<!-- MDOC !-->

A fast Elixir library for reading spreadsheet files, powered by Rust and [Calamine](https://docs.rs/calamine/latest/calamine/).

## Features

- 🚀 **Fast Performance** - Native Rust implementation via NIFs
- 📁 **Multiple Formats** - Supports .xls, .xla, .xlsx, .xlsm, .xlam, .xlsb, and .ods files
- 💾 **Paths or Binaries** - Parse from a file path or from content already in memory
- 🗂️ **Sheet Management** - List and filter sheet names, including hidden sheets
- 📅 **Consistent Types** - Numbers are floats, dates are `NaiveDateTime`, times are `Time`, durations are `Duration`, in every format
- 🛡️ **Bounded Memory** - A per-sheet `:max_cells` limit rejects sparse or crafted files before they are expanded
- 🔧 **Simple API** - `{:ok, result} | {:error, reason}` for anything wrong with the file

## Usage

### Getting Sheet Names

List all sheet names from a file:

```elixir
iex> Spreadsheet.sheet_names("financial_data.xlsx")
{:ok, ["Q1 Revenue", "Q2 Revenue", "Summary"]}
```

Or from binary content using the `:format` option:

```elixir
iex> content = File.read!("financial_data.xlsx")
iex> Spreadsheet.sheet_names(content, format: :binary)
{:ok, ["Q1 Revenue", "Q2 Revenue", "Summary"]}
```

### Filtering Hidden Sheets

Exclude hidden (and very hidden) sheets from the results:

```elixir
iex> Spreadsheet.sheet_names("workbook.xlsx", hidden: false)
{:ok, ["Visible Sheet"]}
```

### Parsing Sheet Data

Parse a specific sheet by name from a file:

```elixir
iex> Spreadsheet.parse("sales.xlsx", sheet: "Q1 Data")
{:ok, [
  ["Product", "Sales", "Date"],
  ["Widget A", 1500.0, ~N[2024-01-15 00:00:00]],
  ["Widget B", 2300.0, ~N[2024-01-20 00:00:00]]
]}
```

Parse all sheets from a file (returns a list of `{sheet_name, rows}` tuples in workbook order):

```elixir
iex> Spreadsheet.parse("sales.xlsx")
{:ok, [
  {"Q1 Data", [
    ["Product", "Sales", "Date"],
    ["Widget A", 1500.0, ~N[2024-01-15 00:00:00]]
  ]},
  {"Q2 Data", [
    ["Product", "Sales", "Date"],
    ["Widget B", 2300.0, ~N[2024-04-15 00:00:00]]
  ]}
]}
```

Or from binary content using the `:format` option:

```elixir
iex> content = File.read!("sales.xlsx")
iex> Spreadsheet.parse(content, sheet: "Q1 Data", format: :binary)
{:ok, [
  ["Product", "Sales", "Date"],
  ["Widget A", 1500.0, ~N[2024-01-15 00:00:00]],
  ["Widget B", 2300.0, ~N[2024-01-20 00:00:00]]
]}
```

### Options

| Option       | Functions                 | Default      | Description |
|--------------|---------------------------|--------------|-------------|
| `:format`    | `parse/2`, `sheet_names/2` | `:filename`  | `:filename` (the first argument is a path) or `:binary` (it is the file's content). |
| `:hidden`    | `parse/2`, `sheet_names/2` | `true`       | When `false`, hidden and very hidden sheets are skipped. Ignored by `parse/2` when `:sheet` is given. |
| `:sheet`     | `parse/2`                  | `nil`        | Name of the sheet to parse. `nil` parses every sheet. |
| `:max_cells` | `parse/2`, `sheet_names/2` | `10_000_000` | Largest sheet accepted, as rows × columns counted from A1. `:infinity` disables the limit. |

### Rows Are Anchored at A1

Rows always start at the top-left corner of the sheet: `rows |> Enum.at(r) |> Enum.at(c)`
is the cell in spreadsheet row `r + 1`, column `c + 1`, wherever the data starts.
Leading empty rows and columns are padded with `nil`, and every row has the same
length. A sheet whose data occupies `C3:D4` returns:

```elixir
{:ok, [
  [nil, nil, nil, nil],
  [nil, nil, nil, nil],
  [nil, nil, "name", "value"],
  [nil, nil, "a", 1.0]
]}
```

Empty sheets, chart sheets and other sheets without cells return `[]`. They are
still listed by `sheet_names/2`. The exception is `.xls` chart sheets, which
can carry the chart's data and return it as rows.

### Data Types

| Cell                                              | .xlsx, .xlsm, .xlam, .xlsb, .xls                            | .ods |
|---------------------------------------------------|-------------------------------------------------------------|------|
| Empty                                             | `nil`                                                       | `nil` |
| Number (including whole numbers)                  | float, e.g. `1234.0`                                        | float |
| Text                                              | `String.t()`                                                | `String.t()` |
| Boolean                                           | `true` / `false`                                            | `true` / `false` |
| Date or date-time                                 | `%NaiveDateTime{}` (a date alone is midnight)               | `%NaiveDateTime{}` (a date alone is midnight) |
| Time of day                                       | `%Time{}` (serial in `[0, 1)` without a duration format)    | `%Duration{}` (see below) |
| Duration (e.g. `[h]:mm:ss`)                       | `%Duration{}` with `hour`, `minute`, `second`, `microsecond` | `%Duration{}` |
| Date outside years 0–9999, or a negative serial   | the raw serial number as a float                            | the raw ISO 8601 string |
| Formula error (`#DIV/0!`, `#N/A`, …)              | `{:error, "#DIV/0!"}`                                       | `""` (see Known Limitations) |
| Formula                                           | its cached result, typed as above                           | its cached result |

Notes:

- **Numbers are always floats**, whatever the format stores (.xls and .xlsb store
  some numbers as integers internally). Compare with `==`, or convert with
  `trunc/1` when you need integers.
- **Durations keep their sign**: `-1:30:00` is `%Duration{hour: -1, minute: -30}`.
  Hours are not folded into days, so `36:00:00` is `%Duration{hour: 36}`.
- **ODS cannot tell a time of day from a duration** (both are `office:time-value`),
  so ODS time cells are always `%Duration{}`. An ODS duration with day, week, month
  or year parts keeps them in the matching `%Duration{}` fields.
- **Sub-second precision**: Excel formats store milliseconds; values with a
  fractional second get `microsecond: {value, 3}`, others `{0, 0}`, so they compare
  equal to sigils such as `~N[2024-01-15 10:00:00]`. ODS values keep the precision
  written in the file (up to microseconds).
- **Values that cannot be represented** are returned as found rather than as `nil`:
  Excel date serials that are negative or fall outside years 0–9999 as the serial
  (a float), and ODS date or time values that do not parse as the original string.
- Excel errors are `{:error, code}` tuples so they never look like text.

### Errors and Exceptions

`parse/2` and `sheet_names/2` return `{:error, reason}` (a string) for anything
wrong with the input file:

- the file does not exist or cannot be read (`"could not open \"x.xlsx\": No such file or directory"`);
- the first argument looks like file content but `:format` is `:filename`
  (the message suggests `format: :binary`), or vice versa;
- the content is not a supported spreadsheet, or calamine cannot parse it;
- the requested sheet does not exist (`"sheet \"Q3\" not found"`);
- a sheet exceeds `:max_cells`, or an .xls file is malformed in a way that would
  make the parser allocate without bound;
- `:format` is neither `:filename` nor `:binary` (kept as an error tuple for
  backwards compatibility).

They raise `ArgumentError` for programming errors: unknown options, a `:sheet`
that is not a string or `nil`, a `:hidden` that is not a boolean, or a
`:max_cells` that is not a positive integer or `:infinity`.

### Limiting Sheet Size

A sheet's rows are fully built in memory, and padding to A1 means a single cell
in `XFD1048576` produces a 1,048,576 × 16,384 rectangle. `:max_cells` (default
10 million) rejects such sheets with an error before anything is allocated:

```elixir
iex> Spreadsheet.parse("sparse.xlsx")
{:error, "sheet \"Data\" has 200000 rows × 16384 columns (3276800000 cells), above the :max_cells limit of 10000000"}

iex> Spreadsheet.parse("big.xlsx", max_cells: 50_000_000)
{:ok, [...]}
```

The limit applies per sheet, to rows × columns counted from A1. For .xlsx and
.xlsb it is checked against the actual cells (the file's declared dimensions are
ignored). .xls and .ods workbooks are decoded in full when they are opened, so
they are scanned before decoding and the limit also applies to `sheet_names/2`
and to every sheet, not just the one requested; for those formats it also bounds
cells the file declares or repeats (`table:number-columns-repeated`, an .xls
DIMENSIONS record) and rejects malformed .xls containers. Only raise or disable
the limit for files you trust.

### Known Limitations

- **.xls float precision**: some .xls numbers stored in Excel's compact RK form
  decode with a binary rounding difference, e.g. `0.12119999999999999` instead
  of `0.1212`. This comes from calamine's decoding.
- **ODS error cells**: LibreOffice stores formula errors in .ods files as empty
  text with a private `calcext:value-type="error"` attribute, which calamine
  reads as `""`.
- **Time-only cells at serial 0**: Excel formats only tell calamine whether a cell
  is a date-time or a duration, not whether its format shows a date. A
  date-formatted cell whose serial is in `[0, 1)` (e.g. `1904-01-01` in a 1904-epoch
  workbook) is returned as a `%Time{}`.
- **.xls and .ods memory**: calamine decodes every sheet of these formats when the
  file is opened, and copies the requested sheet, so parsing one sheet still
  costs memory for the whole workbook.
- **Not limited**: `:max_cells` bounds cells, not the size of individual strings
  or shared-string tables.

For details on the underlying parser, see the [Calamine documentation](https://docs.rs/calamine/latest/calamine/).

<!-- MDOC !-->

## Installation

Add `spreadsheet` to your list of dependencies in `mix.exs`:

```elixir
def deps do
  [
    {:spreadsheet, "~> 0.7.0"}
  ]
end
```

Upgrading from 0.6? See the breaking changes in the [CHANGELOG](CHANGELOG.md).

### No Rust Installation Required

By default, the library uses precompiled NIFs, so **you don't need Rust installed**. If you want to force compilation from source, set the application environment:

```elixir
config :rustler_precompiled, :force_build, spreadsheet: true
```

## Performance

- **Native Speed**: Rust-powered parsing, with cells converted to their final
  Elixir values inside the NIF (no second pass in Elixir)
- **Cross-Platform**: Precompiled NIFs for major platforms (Linux, macOS, Windows)

Parsing is not streamed: each sheet's rows are returned as one list, so memory
grows with the sheet's size. Use `:max_cells` to bound it.

### Benchmarks

Benchmarks comparing Spreadsheet against other popular Elixir XLSX libraries:

**Small Files (10 rows × 5 columns):**
- **~11-12x faster** than pure Elixir implementations
- **~160-600x less memory** usage

**Large Files (10,000 rows × 20 columns):**
- **~10-20x faster** parsing
- **~200-1200x less memory** usage

Memory figures come from Benchee and only count memory allocated by the calling
Elixir process; memory used inside the NIF is not included. For detailed results
and methodology, see [benchmarks/RESULTS.md](benchmarks/RESULTS.md).

## Why Spreadsheet?

Compared to other Elixir spreadsheet libraries:

- **Faster**: Native Rust implementation outperforms pure Elixir parsers
- **More Formats**: Supports more file formats including .ods, .xlsb and legacy .xls files
- **Zero Setup**: No need to install external dependencies

### Alternatives

- [XlsxReader](https://hex.pm/packages/xlsx_reader) - Pure Elixir, XLSX only
- [Xlsxir](https://hex.pm/packages/xlsxir) - XML-based parsing, XLSX only

## Development

Test fixtures in `test/fixtures` are generated by
`elixir test/fixtures/build_fixtures.exs` (the .xls ones need LibreOffice).

### Publishing a new version

Follow the [rustler_precompiled guide](https://hexdocs.pm/rustler_precompiled/precompilation_guide.html):

1. Bump `@version` in `mix.exs` and `version` in `native/spreadsheet/Cargo.toml`
   (the release workflow fails if they differ from each other or from the tag)
2. Release a new tag (`v<version>`)
3. Push the code with the new tag: `git push origin main --tags`
4. Wait for all NIFs to be built
5. Download precompiled NIFs: `mix rustler_precompiled.download Spreadsheet.Calamine --all`
6. Release the package to Hex.pm


## Copyright and License

Copyright (c) 2022-2026 Wilhelm H Kirschbaum

This work is free. You can redistribute it and/or modify it under the
terms of the MIT License. See the [LICENSE](./LICENSE) file for more details.
