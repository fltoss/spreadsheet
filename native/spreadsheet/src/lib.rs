mod cell;
mod encode;
mod workbook;

use calamine::{Data, Range, Reader, SheetVisible, Sheets};
use cell::ToCellValue;
use rustler::{Encoder, Env, Term};
use std::io::{Read, Seek};
use workbook::{with_workbook, Source, WorkbookOperation};

/// Lists sheet names, in workbook order.
#[rustler::nif(schedule = "DirtyCpu")]
fn sheet_names(source: Source, include_hidden: bool) -> Result<Vec<String>, String> {
    with_workbook(source, SheetNames { include_hidden })
}

/// Parses one sheet (`sheet` is a name) into rows, or every sheet (`sheet` is
/// `nil`) into `[{name, rows}]`. Cells are encoded straight into their final
/// Elixir values.
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
) -> Result<Term<'a>, String> {
    match sheet {
        Some(name) => with_workbook(source, ParseSheet { env, name }),
        None => with_workbook(
            source,
            ParseAllSheets {
                env,
                include_hidden,
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
}

impl<'a> WorkbookOperation for ParseSheet<'a> {
    type Output = Term<'a>;

    fn run<RS: Read + Seek>(self, mut workbook: Sheets<RS>) -> Result<Self::Output, String> {
        parse_sheet(self.env, &mut workbook, &self.name)
    }
}

struct ParseAllSheets<'a> {
    env: Env<'a>,
    include_hidden: bool,
}

impl<'a> WorkbookOperation for ParseAllSheets<'a> {
    type Output = Term<'a>;

    fn run<RS: Read + Seek>(self, mut workbook: Sheets<RS>) -> Result<Self::Output, String> {
        let sheets = visible_sheet_names(&workbook, self.include_hidden)
            .into_iter()
            .map(|name| {
                let rows = parse_sheet(self.env, &mut workbook, &name)?;
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

fn parse_sheet<'a, RS: Read + Seek>(
    env: Env<'a>,
    workbook: &mut Sheets<RS>,
    name: &str,
) -> Result<Term<'a>, String> {
    let range = workbook
        .worksheet_range(name)
        .map_err(|error| error.to_string())?;
    Ok(encode_rows(env, &range))
}

fn encode_rows<'a>(env: Env<'a>, range: &Range<Data>) -> Term<'a> {
    let nil = rustler::types::atom::nil().encode(env);
    range
        .rows()
        .map(|row| {
            row.iter()
                .map(|cell| encode::encode_cell(env, cell.to_cell_value(), nil))
                .collect::<Vec<Term<'a>>>()
                .encode(env)
        })
        .collect::<Vec<Term<'a>>>()
        .encode(env)
}

rustler::init!("Elixir.Spreadsheet.Calamine");
