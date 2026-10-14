# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.7.0] - 2026-10-14

This release fixes a VM crash on crafted or very sparse files and makes cell
types consistent across formats, which changes some return values. The public
functions, their arities and the supported Elixir requirement (`~> 1.18`) are
unchanged. Read **Breaking changes** before upgrading.

### Breaking changes

- **Rows are anchored at A1.** Row index 0 is spreadsheet row 1 and column
  index 0 is column A. Previously the rows started at the first non-empty
  cell, so data beginning at `C3` came back at `rows[0][0]` and its position
  was lost. Leading empty rows and columns are now padded with `nil`.
  *Migrate:* sheets whose data starts at A1 are unaffected. Otherwise, index
  cells by their spreadsheet position, or drop the padding yourself, e.g.
  `Enum.drop_while(rows, &Enum.all?(&1, fn cell -> is_nil(cell) end))` for
  leading rows.
- **Numbers are always floats.** `.xls` and `.xlsb` files store some numbers
  as integers (RK records), which used to come back as integers while the same
  cell in `.xlsx`/`.ods` was a float. Every number is now a float, e.g.
  `1234.0`.
  *Migrate:* compare with `==` instead of `===` or pattern matching on
  integers, or convert with `trunc/1` / `round/1` where you need integers.
- **Durations are `%Duration{}` structs, and keep their sign.** Duration-
  formatted cells (e.g. `[h]:mm:ss`) in `.xlsx`/`.xls`/`.xlsb` and
  `office:time-value` cells in `.ods` used to be ISO 8601 strings such as
  `"PT1H30M0S"` (`.ods` returned the file's own string, e.g.
  `"PT01H30M00S"`), and negative durations were clamped to `"PT0H0M0S"`. They
  are now `%Duration{hour: 1, minute: 30}`, and `-1:30` is
  `%Duration{hour: -1, minute: -30}`. Hours are not folded into days.
  *Migrate:* use the struct fields, or `Duration.to_iso8601/1` if you need a
  string (its format differs from the old strings, e.g. `"PT1H30M"`).
- **Time-of-day cells are `%Time{}`.** An `.xlsx`/`.xls`/`.xlsb` date-time
  cell whose serial is in `[0, 1)` (a time without a date, e.g. `h:mm:ss`)
  used to be a `NaiveDateTime` on 1899-12-31. It is now e.g. `~T[12:00:00]`.
  `.ods` cannot distinguish a time of day from a duration, so `.ods` time
  cells are `%Duration{}` (see above).
  *Migrate:* match on `%Time{}`, or combine it with a date via
  `NaiveDateTime.new/2`.
- **Dates that cannot be represented return the serial number instead of
  `nil`.** A date-formatted Excel cell whose serial is negative or falls
  outside years 0–9999 used to be `nil` (or a bogus date for negative
  serials); it is now the raw serial as a float, e.g. `-1.0`. `.ods` date
  values that do not parse are still returned as the raw string.
  *Migrate:* handle a float where you expected `nil` for such cells.
- **Sheets above 10 million cells return an error.** The new `:max_cells`
  option (below) defaults to `10_000_000`, counted as rows × columns from A1.
  Larger sheets now return `{:error, "sheet \"Name\" has R rows × C columns
  (N cells), above the :max_cells limit of 10000000"}` instead of being
  parsed. For `.xls` and `.ods`, which are decoded in full when opened,
  `sheet_names/2` can return this error too.
  *Migrate:* pass `max_cells: n` with a higher limit, or `max_cells:
  :infinity`, for large files you trust.
- **Invalid options raise `ArgumentError`.** Unknown options (e.g. a typo such
  as `sheets:`) used to be silently ignored, and an option of the wrong type
  (`sheet: :Data`, `hidden: "no"`) raised an opaque `ArgumentError` from the
  NIF ("argument error"). Unknown options now raise via `Keyword.validate!/2`,
  and invalid values raise an `ArgumentError` naming the option. `sheet: nil`
  still means "all sheets", and an invalid `:format` still returns
  `{:error, "Invalid format option: ..."}`.
  *Migrate:* fix the options the exception names.
- **Error messages changed.** A missing file now returns
  `{:error, "could not open \"path\": No such file or directory"}` instead of
  `"Cannot detect file format"`, and a missing sheet returns
  `{:error, "sheet \"Name\" not found"}` for every format instead of
  `"Xlsx error: Worksheet 'Name' not found"` and similar.
  *Migrate:* if you match on error strings, update them.

### Fixed

- **A crafted or very sparse file could abort the whole VM.** A 2 KB `.xlsx`
  with values in `A1` and `XFD200000` made calamine allocate ~100 GB for the
  sheet's rectangle; the failed allocation aborted the BEAM rather than
  raising. `.xls` (cell positions, a forged DIMENSIONS record, compound file
  header fields and looping sector chains) and `.ods` (row and column repeat
  attributes, `<text:s>` space runs) had equivalent problems when the file was
  opened. All of these are now rejected with `{:error, reason}` before memory
  is allocated, as long as `:max_cells` is not `:infinity`.
- Data that does not start at A1 keeps its position (see Breaking changes).
- Negative durations keep their sign instead of becoming zero.
- A missing or unreadable file reports the real I/O error instead of "Cannot
  detect file format".
- Passing file contents as a path (with the default `format: :filename`) now
  returns `{:error, reason}` suggesting `format: :binary` instead of raising,
  and passing a path as content with `format: :binary` suggests
  `format: :filename`.
- Paths no longer have to be valid UTF-8 on Unix.
- File extensions are matched case-insensitively (`REPORT.XLSX` is opened as
  `.xlsx` directly, so errors name the real problem).
- VBA modules and `.xlsb` chart sheets are skipped without reading them (they
  return `[]`), so their content can no longer make a whole-workbook parse
  fail. `.xls` chart sheets are still read, since they can carry the chart's
  data.
- Files that make calamine panic (such as `.xls` formulas with truncated
  operands) return `{:error, "could not read workbook: ..."}` instead of
  raising an opaque `:nif_panicked` `ErlangError`.
- NaN or infinite numbers (only possible in malformed files) return
  `{:error, "#NUM!"}` instead of an invalid term.

### Added

- `:max_cells` option for `parse/2` and `sheet_names/2` (default
  `10_000_000`, `:infinity` disables it): the largest sheet accepted, as rows ×
  columns counted from A1. `.xlsx`/`.xlsb` sheets are checked against their
  actual cells (the declared `<dimension>` is ignored); `.xls`/`.ods` files are
  pre-scanned before calamine opens them.
- `Spreadsheet.cell/0` and `Spreadsheet.rows/0` types.
- Tests and generated fixtures (`test/fixtures/build_fixtures.exs`) covering
  `.xlsb`, data not starting at A1, time-of-day cells, positive and negative
  durations, error cells from LibreOffice exports, the 1904 date epoch,
  out-of-range dates, hidden sheets in `.xls`/`.ods`, non-ASCII sheet names,
  chart sheets, missing files and sheets, option validation and crafted
  sparse files.
- CI runs `mix format --check-formatted`, `mix credo --strict`,
  `cargo fmt --check`, `cargo clippy -D warnings` and the Rust unit tests, and
  caches Rust builds. The release workflow fails when the tag does not match
  the version in `mix.exs` and `Cargo.toml`.

### Changed

- Cells are converted to their final Elixir values inside the NIF, which
  removes a second pass over every cell in Elixir. Parsing a 100,000 × 10
  `.xlsx` sheet is about 20% faster.
- The six internal NIFs are collapsed into two (`sheet_names/3`, `parse/4`)
  that take `{:path, path}` or `{:binary, content}`; both still run on dirty
  CPU schedulers.
- Local NIF builds target NIF 2.16, like the precompiled binaries.
- The module documentation is now the README, including a type mapping table
  for every format and a known limitations section.

### Removed

- **Dependency change:** the runtime dependency on `decimal` is gone (it was
  never used). If your project used `Decimal` through this package, add
  `{:decimal, "~> 2.0"}` to your own dependencies.
- The unused `igniter` dev/test dependency, the rustler `serde` feature and
  the direct `serde` crate dependency.
- Internal modules and functions that were never public API (`@moduledoc
  false`): `Spreadsheet.Parser`, and the `Spreadsheet.Calamine` NIFs
  `sheet_names_from_path/2`, `sheet_names_from_binary/2`, `parse_from_path/2`,
  `parse_from_binary/2`, `parse_all_from_path/2` and `parse_all_from_binary/2`.
  Use `Spreadsheet.sheet_names/2` and `Spreadsheet.parse/2`.
- The unused `arm-unknown-linux-gnueabihf` target from
  `native/spreadsheet/.cargo/config.toml`.

### Known limitations

- `.ods` error cells written by LibreOffice are returned as `""`: they are
  stored as empty text with a private `calcext:value-type="error"` attribute,
  which calamine does not read. Needs a calamine change.
- Some `.xls` numbers decode with a binary rounding difference (e.g.
  `0.12119999999999999` for `0.1212`) in calamine's RK decoding. Unchanged.
- `.xls` and `.ods` workbooks are still decoded in full by calamine when
  opened, and the requested sheet is copied, so parsing one sheet costs memory
  for the whole workbook. Pre-scanning those formats closes the crash cases
  above but adds a second pass over the file; fixing this properly needs
  limits (or lazy sheet loading) in calamine.

## [0.6.4] - 2026-10-14

No breaking changes: the public API, supported Elixir requirement (`~> 1.18`)
and all runtime `mix.exs` constraints are unchanged. Dependency updates only.

### Added
- Legacy `.xls` files declaring code page 10007 (Mac Cyrillic) or 21010
  (UTF-16LE) now open; previously calamine failed with "Codepage … not found"
  (`codepage` 0.1.2 → 0.1.3).

### Updated
- Updated Rust dependencies:
  - `encoding_rs` 0.8.35 → 0.8.42: rewritten ASCII fast path, UTF-8 validation
    via `simdutf8`, and a fix for a boundary bug in streaming legacy decoders
    (calamine uses the non-streaming API, so results are unchanged). Pulls in
    four new crates: `simdutf8`, `scopeguard`, `core_detect` and
    `multiversion_no_op`.
  - `flate2` 1.1.9 → 1.1.10: a stream truncated at end-of-file is now an error
    instead of a silent short read, plus an infinite-loop fix in gzip header
    writing. Drops `miniz_oxide` and `adler2` from the lockfile.
  - flate2 1.1.10 moved zlib-rs runtime CPU feature detection (SIMD) behind a
    new default feature, which `zip` disables. `flate2` is now listed directly
    in `Cargo.toml` with `runtime_detection` enabled, so decompression keeps
    using SIMD. The backend is still zlib-rs.
  - Transitive: `cc`, `cfg-if`, `crc32fast`, `equivalent`, `find-msvc-tools`,
    `indexmap`, `inventory`, `js-sys`, `libc`, `syn`, `unicode-ident`,
    `wasm-bindgen` (+ macro crates), `zlib-rs`, `zopfli`.
  - `rustler` stays on 0.38.0: crate 0.39.0 is out, but the hex package is
    still 0.38.0 and the two must match.
  - Verified parsing is unaffected: every fixture in `test/files` (`.ods`,
    `.xls`, `.xlsx` and both hidden-sheet variants) was parsed before and
    after the update via `sheet_names/2`, `sheet_names_from_binary/2`,
    `parse/2` and `parse_from_binary/2`, producing identical results.
- Tested against `rustler_precompiled` 0.10.0, which the existing `~> 0.8`
  constraint already allows (adds NIF 2.18 / OTP 29 support; the options this
  library passes and the checksum file format are unchanged).
- Tested against `decimal` 3.2.0.
- Updated Elixir dev/test dependencies: `ex_doc` 0.40.4, `igniter` 0.8.4, and
  their transitives (`ex_ast`, `finch`, `hpax`, `makeup`, `mint`, `req`,
  `sourceror`, `spitfire`). Dropped the stale `inflex` and `toml` lock
  entries, which no environment resolved.

### Removed
- Dropped the unused `dialyxir` dev/test dependency.

### Internal
- CI actions updated to their current major versions.

## [0.6.3] - 2026-08-25

No breaking changes: the public API, supported Elixir requirement (`~> 1.18`)
and all runtime Elixir dependencies are unchanged. Dependency updates only.

### Updated
- Updated transitive Rust dependencies (`android_system_properties`, `cc`,
  `crc32fast`, `fast-float2`, `find-msvc-tools`, `futures-core`,
  `futures-task`, `futures-util`, `js-sys`, `log`, `syn`, `wasm-bindgen`,
  `wasm-bindgen-macro`, `wasm-bindgen-macro-support`, `wasm-bindgen-shared`,
  `zlib-rs`). The direct Rust dependencies (`calamine` 0.36.1, `rustler`
  0.38.0, `chrono` 0.4.45, `serde` 1.0.229) were already at their latest
  releases and are unchanged.
  - Verified parsing is unaffected: every fixture in `test/files` (`.ods`,
    `.xls`, `.xlsx` and both hidden-sheet variants) was parsed before and
    after the update via `sheet_names/2`, `sheet_names_from_binary/2`,
    `parse/2` and `parse_from_binary/2`, producing identical results.
- Updated Elixir dev/test dependencies (no runtime deps changed):
  `earmark_parser` 1.4.45 → 1.4.46, `req` 0.7.1 → 0.7.3, `spitfire` 0.3.13
  → 0.4.0. Dropped the stale `castore` lock entry, which no environment
  resolved.

## [0.6.2] - 2026-07-30

No breaking changes: the public API, supported Elixir requirement (`~> 1.18`)
and all runtime Elixir dependencies are unchanged. Dependency updates only.

### Updated
- Updated Rust dependencies:
  - [calamine](https://github.com/tafia/calamine/blob/master/Changelog.md) 0.36.0 → 0.36.1
    - Bug fixes only: fixed a 0.36.0 regression that broke Strict OOXML
      (ISO/IEC 29500) `_rels/.rels` links, fixed XLSB parsing of unhandled
      workbook-global records (previously could silently drop a worksheet or
      panic), and XLS string cells now preserve empty strings instead of
      discarding them.
  - [serde](https://github.com/serde-rs/serde/releases) 1.0.228 → 1.0.229
  - Updated transitive Rust dependencies (`cc`, `futures-core`, `futures-task`,
    `futures-util`, `libc`, `proc-macro2`, `quote`, `syn`)
- Updated Elixir dev/test dependencies (no runtime deps changed):
  `igniter` 0.8.2 → 0.8.3, `ex_ast` 0.12.10 → 0.13.1, `glob_ex` 0.1.11 →
  0.1.12, `req` 0.6.3 → 0.7.1

## [0.6.1] - 2026-07-17

No breaking changes: the public API, supported Elixir requirement (`~> 1.18`)
and all runtime Elixir dependencies are unchanged.

### Changed
- CI now tests every supported Elixir version (1.18, 1.19 and 1.20) against
  every Erlang/OTP version each of them supports (26 through 29), using the
  latest patch releases (1.18.4 / 1.19.5 / 1.20.2, OTP 26.2 / 27.3 / 28.5 /
  29.0).

### Updated
- Updated Rust dependencies:
  - [calamine](https://github.com/tafia/calamine/blob/master/Changelog.md) 0.35.0 → 0.36.0
    - No API changes affecting this library (verified: only
      `open_workbook_auto*`, `sheets_metadata` and `worksheet_range` are used).
    - Bug fixes: shared-formula cell-name replacement for defined names, custom
      number-format detection for the `*` (fill/repeat) operator, better XLS
      error messages for truncated files.
    - Behaviour note: XLSX text cells now trim leading/trailing ASCII
      whitespace unless `xml:space="preserve"` is set, matching Excel and the
      ECMA-376 spec. Files written by Excel, LibreOffice, openpyxl etc. set
      this attribute when whitespace is intended, so their output is
      unaffected; only spec-violating writers are normalised.
  - Updated transitive Rust dependencies (`atoi_simd`, `cc`, `memchr`,
    `quick-xml`, `syn`, `wasm-bindgen`, `zip`, `zlib-rs`, …)
- Updated Elixir dev/test dependencies (transitive only, no runtime deps
  changed): `ex_ast` 0.12.0 → 0.12.10, `hpax` 1.0.3 → 1.0.4, `makeup` 1.2.1 →
  1.2.2, `mint` 1.9.0 → 1.9.3, `req` 0.6.2 → 0.6.3

## [0.6.0] - 2026-06-24

### Fixed
- **Breaking:** `.ods` dates are now normalised to `NaiveDateTime` (date-only
  values become midnight), matching the `.xlsx`/`.xls` output. Previously they
  leaked back as raw ISO 8601 strings, so the type of a date cell depended on the
  source format.
- Duration-formatted cells (e.g. `[h]:mm:ss`) are no longer misread as bogus
  dates. Calamine surfaces them as `DateTime`/`TimeDelta` values; they are now
  returned as ISO 8601 duration strings (e.g. `"PT1H30M0S"`).

### Changed
- Internal cleanup: deduplicated the path/binary NIF bodies behind generic Rust
  helpers, and the all-sheets/format-error handling in the Elixir wrapper. No
  user-facing behaviour change.
- Removed the inert `doctest Spreadsheet` declaration (the `@doc` examples are
  illustrative and were never executed as doctests).

### Updated
- Updated Elixir dependencies:
  - [rustler](https://github.com/rusterlium/rustler/blob/master/CHANGELOG.md) 0.37.3 → 0.38.0
    - Added `Term::size` support on Erlang 29 (`nif_version_2_18`), direct UTF-8
      atom support (`nif_version_2_17`), `NewBinary::from_iter`, and
      `deserialize_seq` for binary deserialization
    - Deprecated the previous atom-creation signatures (alongside the new UTF-8
      variants) and removed some deprecated codegen features. No code changes
      were required — this library uses only `rustler::nif`, `NifTaggedEnum`, and
      `rustler::init!`, none of which are affected. Not breaking for callers.
  - `credo` 1.7.18 → 1.7.19 (dev/test)
  - `decimal` 3.1.0 → 3.1.1
  - `ex_doc` 0.40.2 → 0.40.3 (dev)
  - `igniter` 0.8.0 → 0.8.2 (dev/test)
- Updated Rust dependencies:
  - [rustler](https://github.com/rusterlium/rustler/blob/master/CHANGELOG.md) 0.37.4 → 0.38.0 (kept in sync with the Elixir package)
  - [chrono](https://github.com/chronotope/chrono/blob/main/CHANGELOG.md) 0.4.44 → 0.4.45
  - Updated transitive Rust dependencies (`cc`, `js-sys`, `log`, `memchr`,
    `quote`, `shlex`, `syn`, `wasm-bindgen`, `zlib-rs`, …)

## [0.5.0] - 2026-05-15

### Added
- `Spreadsheet.Calamine.parse_all_from_path/2` and `parse_all_from_binary/2` —
  open the workbook once and return every sheet, replacing the previous
  open-per-sheet loop in `Spreadsheet.parse/2`.

### Changed
- **Breaking:** Excel formula errors (`#REF!`, `#DIV/0!`, `#N/A`, …) are now
  returned as `{:error, reason}` tuples instead of bare strings, so they are
  distinguishable from text cells.
- All NIFs are now scheduled on `DirtyCpu` to avoid stalling the BEAM scheduler
  when parsing large workbooks.

### Fixed
- Datetimes that calamine cannot convert no longer leak the literal string
  `"Invalid DateTime"` into the result — they become `nil`.
- Sub-second precision in `Data::DateTime` cells is preserved (was previously
  truncated to whole seconds by the format string).
- Hidden-sheet filtering now uses sheet metadata in both branches, removing a
  latent inconsistency between `hidden: true` and `hidden: false`.

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

