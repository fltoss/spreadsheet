use calamine::{open_workbook_auto, open_workbook_auto_from_rs, Data, Reader, SheetVisible};
use rustler::{Binary, NifTaggedEnum};
use std::io::Cursor;

fn filter_sheet_names_by_visibility(
    sheets_metadata: &[calamine::Sheet],
    show_hidden: bool,
) -> Vec<String> {
    sheets_metadata
        .iter()
        .filter(|sheet| show_hidden || sheet.visible == SheetVisible::Visible)
        .map(|sheet| sheet.name.clone())
        .collect()
}

#[rustler::nif(schedule = "DirtyCpu")]
fn sheet_names_from_binary(content: Binary, show_hidden: bool) -> Result<Vec<String>, String> {
    let cursor = Cursor::new(content.as_slice());

    match open_workbook_auto_from_rs(cursor) {
        Ok(workbook) => Ok(filter_sheet_names_by_visibility(
            workbook.sheets_metadata(),
            show_hidden,
        )),
        Err(e) => Err(e.to_string()),
    }
}

#[rustler::nif(schedule = "DirtyCpu")]
fn sheet_names_from_path(path: &str, show_hidden: bool) -> Result<Vec<String>, String> {
    match open_workbook_auto(path) {
        Ok(workbook) => Ok(filter_sheet_names_by_visibility(
            workbook.sheets_metadata(),
            show_hidden,
        )),
        Err(e) => Err(e.to_string()),
    }
}

#[rustler::nif(schedule = "DirtyCpu")]
fn parse_from_path(path: &str, sheet_name: &str) -> Result<Vec<Vec<ColumnData>>, String> {
    open_workbook_auto(path)
        .map_err(|e| e.to_string())
        .and_then(|mut workbook| {
            workbook
                .worksheet_range(sheet_name)
                .map_err(|e| e.to_string())
                .map(extract_rows)
        })
}

#[rustler::nif(schedule = "DirtyCpu")]
fn parse_from_binary(content: Binary, sheet_name: &str) -> Result<Vec<Vec<ColumnData>>, String> {
    open_workbook_auto_from_rs(Cursor::new(content.as_slice()))
        .map_err(|e| e.to_string())
        .and_then(|mut workbook| {
            workbook
                .worksheet_range(sheet_name)
                .map_err(|e| e.to_string())
                .map(extract_rows)
        })
}

#[rustler::nif(schedule = "DirtyCpu")]
fn parse_all_from_path(
    path: &str,
    show_hidden: bool,
) -> Result<Vec<(String, Vec<Vec<ColumnData>>)>, String> {
    let mut workbook = open_workbook_auto(path).map_err(|e| e.to_string())?;
    let names = filter_sheet_names_by_visibility(workbook.sheets_metadata(), show_hidden);
    names
        .into_iter()
        .map(|name| {
            workbook
                .worksheet_range(&name)
                .map_err(|e| e.to_string())
                .map(|range| (name, extract_rows(range)))
        })
        .collect()
}

#[rustler::nif(schedule = "DirtyCpu")]
fn parse_all_from_binary(
    content: Binary,
    show_hidden: bool,
) -> Result<Vec<(String, Vec<Vec<ColumnData>>)>, String> {
    let mut workbook =
        open_workbook_auto_from_rs(Cursor::new(content.as_slice())).map_err(|e| e.to_string())?;
    let names = filter_sheet_names_by_visibility(workbook.sheets_metadata(), show_hidden);
    names
        .into_iter()
        .map(|name| {
            workbook
                .worksheet_range(&name)
                .map_err(|e| e.to_string())
                .map(|range| (name, extract_rows(range)))
        })
        .collect()
}

#[derive(NifTaggedEnum)]
enum ColumnData {
    Int(i64),
    Float(f64),
    String(String),
    Bool(bool),
    DateTime(String),
    DateTimeIso(String),
    DurationIso(String),
    Error(String),
    Empty,
}

fn extract_rows(range: calamine::Range<Data>) -> Vec<Vec<ColumnData>> {
    range
        .rows()
        .map(|row| row.iter().map(extract_column).collect())
        .collect()
}

fn extract_column(cell: &Data) -> ColumnData {
    match cell {
        Data::Int(val) => ColumnData::Int(*val),
        Data::Float(val) => ColumnData::Float(*val),
        Data::String(val) => ColumnData::String(val.to_string()),
        Data::Bool(val) => ColumnData::Bool(*val),
        Data::DateTime(val) => match val.as_datetime() {
            Some(ndt) => ColumnData::DateTime(ndt.format("%Y-%m-%dT%H:%M:%S%.f").to_string()),
            None => ColumnData::Empty,
        },
        Data::DateTimeIso(val) => ColumnData::DateTimeIso(val.to_string()),
        Data::DurationIso(val) => ColumnData::DurationIso(val.to_string()),
        Data::Error(val) => ColumnData::Error(val.to_string()),
        Data::Empty => ColumnData::Empty,
    }
}

rustler::init!("Elixir.Spreadsheet.Calamine");
