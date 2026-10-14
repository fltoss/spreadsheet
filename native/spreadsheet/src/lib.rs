mod budget;
mod cell;
mod encode;
mod error;
mod format;
mod prescan;
mod sheet;
mod workbook;

use budget::Budget;
use error::Error;
use rustler::{Encoder, Env, Term};
use std::any::Any;
use std::panic::{self, AssertUnwindSafe};
use workbook::Source;

/// Lists sheet names, in workbook order.
///
/// `max_cells` (`nil` for no limit) applies here too: .xls and .ods workbooks
/// are fully decoded when opened.
#[rustler::nif(schedule = "DirtyCpu")]
fn sheet_names(
    source: Source,
    include_hidden: bool,
    max_cells: Option<u64>,
) -> Result<Vec<String>, String> {
    guard(|| {
        let workbook = workbook::open(&source, &Budget::new(max_cells))?;
        Ok(workbook::sheet_names(&workbook, include_hidden))
    })
}

/// Parses one sheet (`sheet` is a name) into rows, or every sheet (`sheet` is
/// `nil`) into `[{name, rows}]`. Cells are encoded straight into their final
/// Elixir values. `max_cells` bounds the rows × columns (from A1) of every
/// sheet the call builds, together; a sheet that does not fit is rejected
/// before its rows are built.
///
/// Runs on a dirty CPU scheduler: path reads are a small part of the work next
/// to decompression and XML/BIFF decoding, so one scheduler type fits both
/// sources.
#[rustler::nif(schedule = "DirtyCpu")]
fn parse<'a>(
    env: Env<'a>,
    source: Source,
    sheet: Option<String>,
    include_hidden: bool,
    max_cells: Option<u64>,
) -> Result<Term<'a>, String> {
    guard(|| {
        let mut budget = Budget::new(max_cells);
        let mut workbook = workbook::open(&source, &budget)?;
        match sheet {
            Some(name) => sheet::read_sheet(env, &mut workbook, &name, &mut budget),
            None => {
                let sheets = workbook::sheet_names(&workbook, include_hidden)
                    .into_iter()
                    .map(|name| {
                        let rows = sheet::read_sheet(env, &mut workbook, &name, &mut budget)?;
                        Ok((name, rows))
                    })
                    .collect::<Result<Vec<_>, Error>>()?;
                Ok(sheets.encode(env))
            }
        }
    })
}

/// Runs a NIF body, turning errors into their messages. calamine panics on
/// some malformed files (e.g. truncated formula operands in .xls); those are
/// reported like any other unreadable file instead of raising an opaque
/// `:nif_panicked` error.
fn guard<T>(body: impl FnOnce() -> Result<T, Error>) -> Result<T, String> {
    panic::catch_unwind(AssertUnwindSafe(body))
        .unwrap_or_else(|payload| Err(Error::Panic(panic_message(&*payload).to_string())))
        .map_err(|error| error.to_string())
}

fn panic_message(payload: &(dyn Any + Send)) -> &str {
    payload
        .downcast_ref::<&str>()
        .copied()
        .or_else(|| payload.downcast_ref::<String>().map(String::as_str))
        .unwrap_or("the reader panicked")
}

rustler::init!("Elixir.Spreadsheet.Calamine");
