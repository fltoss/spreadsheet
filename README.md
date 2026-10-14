# Spreadsheet

<!-- MDOC !-->

A fast Elixir library for reading spreadsheet files, powered by Rust and [Calamine](https://docs.rs/calamine/latest/calamine/).

## Features

- 🚀 **Fast Performance** - Native Rust implementation via NIFs
- 📁 **Multiple Formats** - Supports .xls, .xla, .xlsx, .xlsm, .xlam, .xlsb, and .ods files
- 💾 **Paths or Binaries** - Parse from a file path or from content already in memory
- 🗂️ **Sheet Management** - List and filter sheet names, including hidden sheets
- 📅 **Consistent Types** - Numbers are floats, dates are `NaiveDateTime`, times are `Time`, durations are `Duration`, in every format
- 🛡️ **Bounded Memory** - A per-call `:max_cells` budget rejects sparse or crafted files before they are expanded
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
| `:hidden`    | `parse/2`, `sheet_names/2` | `true`       | When `false`, hidden and very hidden sheets are skipped. `nil` means the default. Ignored by `parse/2` when `:sheet` is given. |
| `:sheet`     | `parse/2`                  | `nil`        | Name of the sheet to parse. `nil` parses every sheet. `sheet_names/2` accepts and ignores it, so one options list works for both. |
| `:max_cells` | `parse/2`, `sheet_names/2` | `50_000_000` | Cell budget of the call: the rows × columns (from A1) of every sheet built, summed. `:infinity` disables it. See "Limiting Memory" below. |

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
| Date or date-time                                 | `%NaiveDateTime{}` (a date alone is midnight; milliseconds) | `%NaiveDateTime{}` (a date alone is midnight; sub-second precision as written in the file, up to 6 digits) |
| Time of day                                       | `%Time{}` (serial in `[0, 1)` without a duration format)    | `%Duration{}` (see below) |
| Duration (e.g. `[h]:mm:ss`)                       | `%Duration{}` with `hour`, `minute`, `second`, `microsecond` | `%Duration{}` (sub-second precision as written in the file, up to 6 digits) |
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
- the sheets exceed `:max_cells`, or an .xls or .ods file is malformed in a way
  that would make the parser allocate without bound or loop forever;
- `:format` is neither `:filename` nor `:binary` (kept as an error tuple for
  backwards compatibility).

They raise `ArgumentError` for programming errors: unknown options, a `:sheet`
that is not a string or `nil`, a `:hidden` that is not a boolean or `nil` (only
checked when it is used), or a `:max_cells` that is not a positive integer or
`:infinity`. Duplicate options do not raise: the first one wins.

### Limiting Memory with `:max_cells`

Rows are fully built in memory, and padding to A1 means a single cell in
`XFD1048576` produces a 1,048,576 × 16,384 rectangle. `:max_cells` is a cell
budget for one call: every sheet the call builds costs its rows × columns,
counted from A1, and a sheet that does not fit in what is left returns an error
before anything is allocated for it. The default of 50 million fits, for
example, one sheet of 1,048,576 rows × 47 columns, or 500,000 rows × 100
columns:

```elixir
iex> Spreadsheet.parse("sparse.xlsx")
{:error, "sheet \"Data\" has 200000 rows × 16384 columns (3276800000 cells), above the :max_cells limit of 50000000"}

iex> Spreadsheet.parse("big.xlsx", max_cells: 100_000_000)
{:ok, [...]}
```

As a guide, each cell of the result takes about 17 bytes when it is `nil`,
33 bytes for a float and 57 bytes for typical mixed data (short strings, dates,
floats), so the default allows a result of roughly 1–3 GB. Lower it for
untrusted uploads.

- **.xlsx and .xlsb** sheets are checked against their actual cells while they
  are read (the file's declared dimensions are ignored). A sheet that repeats
  cell positions far beyond its rectangle is rejected as malformed, whatever
  the limit.
- **.xls and .ods** workbooks are decoded in full when they are opened, so they
  are scanned before decoding and every sheet of the workbook counts against
  the budget, even when you parse one sheet or only call `sheet_names/2`. For
  .ods the budget also bounds the cells repeat attributes expand to
  (`table:number-columns-repeated`).
- **Malformed files** are rejected whatever the limit, including with
  `max_cells: :infinity`: .xls compound file headers or sector chains that would
  make the parser allocate without bound or loop forever, .xls DIMENSIONS
  records beyond 65,536 rows × 256 columns, and .ods `<text:s>` space runs
  above 100 million spaces per workbook.

Only raise or disable the limit for files you trust.

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

Building from source needs a Rust toolchain and `rustler` in your own
dependencies (it is an optional dependency of this package). The crate uses
rustler 0.38, so pin the same version:

```elixir
{:rustler, "~> 0.38.0", runtime: false}
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
