# spreadsheet NIF

The Rust side of `Spreadsheet`, loaded as `Spreadsheet.Calamine` (see
`lib/spreadsheet/calamine.ex`). It wraps [calamine](https://docs.rs/calamine)
and returns finished Elixir terms.

- `src/lib.rs` - the two NIFs: `sheet_names/3` and `parse/4`.
- `src/workbook.rs` - opening a workbook from `{:path, path}` or
  `{:binary, content}`, trying formats in calamine's detection order.
- `src/format.rs` - everything per format: extensions, which pre-scan runs,
  how calamine opens it, which sheets hold cells and where they come from
  (streamed cells for .xlsx/.xlsb, decoded ranges for .xls/.ods).
- `src/sheet.rs` - building rows anchored at A1.
- `src/budget.rs` - the `:max_cells` budget.
- `src/error.rs` - every error message the crate returns.
- `src/cell.rs` / `src/encode.rs` - converting cells to Elixir values
  (`NaiveDateTime`, `Time`, `Duration`, ...).
- `src/prescan/` - pre-scans of .xls (`cfb.rs`: the compound file
  container; `biff.rs`: BIFF records) and .ods (`ods.rs`) files, which
  calamine decodes in full when opening them.

Precompiled binaries are built by `.github/workflows/release.yml` for NIF 2.16;
set `config :rustler_precompiled, :force_build, spreadsheet: true` to build
from source (this repository's `config/config.exs` does so for development).

Run `cargo fmt`, `cargo clippy --all-targets -- -D warnings` and the crate's
unit tests from this directory.
