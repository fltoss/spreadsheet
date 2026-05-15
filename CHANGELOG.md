# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Updated
- Updated Elixir dependencies: `erlex` 0.2.8 → 0.2.9, `ex_ast` 0.11.2 → 0.12.0
- Updated transitive Rust dependencies (`cc`, `hashbrown`, `js-sys`, `libc`, `quick-xml`, `wasm-bindgen`)

## [0.4.10] - 2026-05-12
- Relax `rustler_precompiled` package version

## [0.4.9] - 2026-05-12

### Added
- Allow `decimal` package version 2.0 or 3.0 (`~> 2.0 or ~> 3.0`)

### Updated
- Updated `rustler_precompiled` 0.8.1 → 0.9.0
- Updated Rust dependencies:
  - [calamine](https://github.com/tafia/calamine/blob/master/CHANGELOG.md) 0.34.0 → 0.35.0
    - Added `has_1904_epoch()` method on XLS/XLSX/XLSB structs to detect Excel date epoch
    - Made `RowDeserializer` public for custom `DeserializeSeed` implementations
    - Performance: optimised `<v>` cell value parsing and cell reference parsing
    - Bug fixes: redundant `<v>` in XLSX inlineStr cells, XML namespace prefixes in relation ids, rich text `<r>` after plain `<t>`, XLS date formats for formula values, XLSX shared formula expansion corrupting function names (e.g. `LOG10`), inlineStr CDATA handling, BIFF5 formula parse panic, panic on unsupported XLS formula functions
  - [rustler](https://github.com/rusterlium/rustler/blob/master/CHANGELOG.md) 0.37.3 → 0.37.4
    - Fixed OTP 22/23 heap overflow by gating `ErlNifResourceTypeInit` members/dyncall on `nif_version_2_16`

## [0.4.7] - 2026-03-18

### Updated
- Updated Rust dependencies:
  - [calamine](https://github.com/tafia/calamine/blob/master/CHANGELOG.md) 0.33.0 → 0.34.0
    - Performance improvements: cell reader buffer reuse, optimised XLS sector chain reads, cached path/attr lookups, minimised allocations
    - Bug fixes: capitalized workbook/book stream handling in XLS, BIFF5 record parsing, empty `<v/>` values in XLSX, millisecond rounding
  - [rustler](https://github.com/rusterlium/rustler/blob/master/CHANGELOG.md) 0.37.2 → 0.37.3

## [0.4.6] - 2026-03-02

### Updated
- Updated Rust dependencies:
  - [calamine](https://github.com/tafia/calamine/blob/master/CHANGELOG.md) 0.32.0 → 0.33.0
  - [chrono](https://github.com/chronotope/chrono/blob/main/CHANGELOG.md) 0.4.43 → 0.4.44

## [0.4.5] - 2026-01-20

### Updated
- Updated Rust dependencies

## [0.4.3] - 2025-12-05

### Updated
- Updated Rust dependencies:
  - [rustler](https://github.com/rusterlium/rustler/blob/master/CHANGELOG.md) 0.37.0 → 0.37.2

## [0.4.2] - 2025-11-24

### Updated
- Updated Rust dependencies in native spreadsheet module:
  - [calamine](https://github.com/tafia/calamine/blob/master/CHANGELOG.md) 0.31.0 → 0.32.0

## [0.4.1] - 2025-11-17

### Changed
- Dependency updates

## [0.4.0] - 2025-10-29

### Added
- **NEW**: `parse/2` now supports parsing all sheets at once when called without the `:sheet` option
  - Returns a list of `{sheet_name, sheet_data}` tuples in sheet order
  - Supports `:hidden` option to exclude hidden sheets
  - Works with both file paths and binary content via `:format` option
  - Example: `Spreadsheet.parse("workbook.xlsx")` returns `{:ok, [{"Sheet1", [...]}, {"Sheet2", [...]}]}`

### Changed
- **BREAKING**: Unified API interface for `sheet_names/2` and `parse/2` functions
  - `sheet_names/2` now accepts a `:format` option (`:filename` or `:binary`) instead of having separate `_from_binary` variant
  - `parse/2` now uses a `:sheet` option to specify which sheet to parse, instead of a separate parameter
  - Default format is `:filename` for backwards compatibility with existing code that doesn't pass options
  - Examples:
    - `Spreadsheet.sheet_names(content, format: :binary)` replaces `Spreadsheet.sheet_names_from_binary(content)`
    - `Spreadsheet.parse(content, sheet: "Sheet1", format: :binary)` replaces `Spreadsheet.parse_from_binary(content, "Sheet1")`
    - `Spreadsheet.parse("file.xlsx", sheet: "Sheet1")` replaces `Spreadsheet.parse("file.xlsx", "Sheet1")`

### Deprecated
- `sheet_names_from_binary/2` - Use `sheet_names/2` with `format: :binary` instead
- `parse_from_binary/2` - Use `parse/2` with `sheet:` and `format: :binary` options instead

### Fixed
- Improved error messages for invalid format options

## [0.3.0] - 2025-10-13

### Updated
- Updated Rust dependencies in native spreadsheet module:
  - [calamine](https://github.com/tafia/calamine/blob/master/CHANGELOG.md) 0.30.0 → 0.31.0
  - [rustler](https://github.com/rusterlium/rustler/blob/master/CHANGELOG.md) 0.36.2 → 0.37.0
  - [chrono](https://github.com/chronotope/chrono/blob/main/CHANGELOG.md) 0.4.41 → 0.4.42
  - [serde](https://github.com/serde-rs/serde/releases) 1.0.219 → 1.0.228
- Changed calamine feature flag from `dates` to `chrono`

## [0.2.4] - 2025-09-03

### Updated
- Updated calamine dependency to version 0.30.0

## [0.2.3] - 2025-01-08

### Updated
- Updated calamine dependency to version 0.29.0

## [0.2.2] - 2025-07-23

### Added
- Support for filtering hidden sheets in `sheet_names/2` and `sheet_names_from_binary/2` functions

