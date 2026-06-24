# Spreadsheet

<!-- MDOC !-->

A fast, memory-efficient Elixir library for parsing spreadsheet files powered by Rust and [Calamine](https://docs.rs/calamine/latest/calamine/).

## Features

- 🚀 **Fast Performance** - Native Rust implementation via NIFs for high-speed parsing
- 📁 **Multiple Formats** - Supports .xls, .xla, .xlsx, .xlsm, .xlam, .xlsb, and .ods files
- 💾 **Memory Efficient** - Parse from file paths or binary content
- 🗂️ **Sheet Management** - List and filter sheet names, including hidden sheets
- 📅 **Smart Type Handling** - Automatic conversion of dates to NaiveDateTime and numbers to Float
- 🔧 **Simple API** - Clean, functional interface with {:ok, result} | {:error, reason} patterns

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

Exclude hidden sheets from the results:

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

Parse all sheets from a file (returns a list of `{sheet_name, sheet_data}` tuples in order):

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

# Parse all sheets from binary
iex> Spreadsheet.parse(content, format: :binary)
{:ok, [
  {"Q1 Data", [...]},
  {"Q2 Data", [...]}
]}
```

### Data Type Handling

The library automatically converts data types:
- **Dates**: Converted to `NaiveDateTime` structs. This is consistent across all
  formats — `.ods` files, which store dates as ISO 8601 strings, are normalised
  to `NaiveDateTime` just like `.xlsx`/`.xls`. Date-only values become a
  `NaiveDateTime` at midnight.
- **Numbers**: Converted to `Integer` or `Float` values
- **Durations**: Cells formatted as elapsed time (e.g. `[h]:mm:ss`) are returned
  as ISO 8601 duration strings (e.g. `"PT1H30M0S"`) rather than being misread as
  a date
- **Empty cells**: Returned as `nil`
- **Text**: Returned as strings
- **Formula errors**: `#REF!`, `#DIV/0!`, `#N/A`, etc. are returned as
  `{:error, reason}` tuples so they stay distinguishable from text cells

For detailed information on the underlying parsing engine, see the [Calamine documentation](https://docs.rs/calamine/latest/calamine/).

<!-- MDOC !-->

## Installation

Add `spreadsheet` to your list of dependencies in `mix.exs`:

```elixir
def deps do
  [
    {:spreadsheet, "~> 0.6.0"}
  ]
end
```

### No Rust Installation Required

By default, the library uses precompiled NIFs, so **you don't need Rust installed**. If you want to force compilation from source, set the application environment:

```elixir
config :rustler_precompiled, :force_build, spreadsheet: true
```

## Performance

Spreadsheet is built for performance and handles large files efficiently:

- **Native Speed**: Rust-powered parsing for maximum throughput
- **Memory Efficient**: Streaming approach minimizes memory usage
- **Cross-Platform**: Precompiled NIFs for major platforms (Linux, macOS, Windows)
- **Production Ready**: Battle-tested in high-volume data processing environments

### Benchmarks

Comprehensive benchmarks comparing Spreadsheet against other popular Elixir XLSX libraries show significant performance advantages:

**Small Files (10 rows × 5 columns):**
- **~11-12x faster** than pure Elixir implementations
- **~160-600x less memory** usage

**Large Files (10,000 rows × 20 columns):**
- **~10-20x faster** parsing
- **~200-1200x less memory** usage

For detailed benchmark results and methodology, see [benchmarks/RESULTS.md](benchmarks/RESULTS.md).

## Why Spreadsheet?

Compared to other Elixir spreadsheet libraries:

- **Faster**: Native Rust implementation outperforms pure Elixir parsers
- **More Formats**: Supports more file formats including .ods and legacy .xls files
- **Better Memory Usage**: Efficient memory handling for large files
- **Zero Setup**: No need to install external dependencies

### Alternatives

- [XlsxReader](https://hex.pm/packages/xlsx_reader) - Pure Elixir, XLSX only
- [Xlsxir](https://hex.pm/packages/xlsxir) - XML-based parsing, XLSX only

## Development

### Publishing a new version

Follow the [rustler_precompiled guide](https://hexdocs.pm/rustler_precompiled/precompilation_guide.html):

1. Release a new tag
2. Push the code with the new tag: `git push origin main --tags`
3. Wait for all NIFs to be built
4. Download precompiled NIFs: `mix rustler_precompiled.download Spreadsheet.Calamine --all`
5. Release the package to Hex.pm


## Copyright and License

Copyright (c) 2022-2026 Wilhelm H Kirschbaum

This work is free. You can redistribute it and/or modify it under the
terms of the MIT License. See the [LICENSE.md](./LICENSE.md) file for more details.
