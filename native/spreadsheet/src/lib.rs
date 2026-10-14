mod cell;
mod encode;
mod limits;
mod sheet;
mod workbook;

use calamine::{Reader, SheetVisible, Sheets};
use limits::CellLimit;
use rustler::{Encoder, Env, Term};
use std::io::{Read, Seek};
use workbook::{with_workbook, Source, WorkbookOperation};

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
    with_workbook(source, CellLimit(max_cells), SheetNames { include_hidden })
}

/// Parses one sheet (`sheet` is a name) into rows, or every sheet (`sheet` is
/// `nil`) into `[{name, rows}]`. Cells are encoded straight into their final
/// Elixir values. Sheets whose rows × columns (from A1) exceed `max_cells`
/// are rejected before their rows are built.
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
    let limit = CellLimit(max_cells);
    match sheet {
        Some(name) => with_workbook(source, limit, ParseSheet { env, name, limit }),
        None => with_workbook(
            source,
            limit,
            ParseAllSheets {
                env,
                include_hidden,
                limit,
            },
        ),
    }
}

struct SheetNames {
    include_hidden: bool,
}

impl WorkbookOperation for SheetNames {
    type Output = Vec<String>;

    fn run<RS: Read + Seek>(self, workbook: Sheets<RS>) -> Result<Self::Output, String> {
        Ok(visible_sheet_names(&workbook, self.include_hidden))
    }
}

struct ParseSheet<'a> {
    env: Env<'a>,
    name: String,
    limit: CellLimit,
}

impl<'a> WorkbookOperation for ParseSheet<'a> {
    type Output = Term<'a>;

    fn run<RS: Read + Seek>(self, mut workbook: Sheets<RS>) -> Result<Self::Output, String> {
        sheet::read_sheet(self.env, &mut workbook, &self.name, self.limit)
    }
}

struct ParseAllSheets<'a> {
    env: Env<'a>,
    include_hidden: bool,
    limit: CellLimit,
}

impl<'a> WorkbookOperation for ParseAllSheets<'a> {
    type Output = Term<'a>;

    fn run<RS: Read + Seek>(self, mut workbook: Sheets<RS>) -> Result<Self::Output, String> {
        let sheets = visible_sheet_names(&workbook, self.include_hidden)
            .into_iter()
            .map(|name| {
                let rows = sheet::read_sheet(self.env, &mut workbook, &name, self.limit)?;
                Ok((name, rows))
            })
            .collect::<Result<Vec<_>, String>>()?;
        Ok(sheets.encode(self.env))
    }
}

fn visible_sheet_names<RS: Read + Seek>(
    workbook: &Sheets<RS>,
    include_hidden: bool,
) -> Vec<String> {
    workbook
        .sheets_metadata()
        .iter()
        .filter(|sheet| include_hidden || sheet.visible == SheetVisible::Visible)
        .map(|sheet| sheet.name.clone())
        .collect()
}

rustler::init!("Elixir.Spreadsheet.Calamine");
