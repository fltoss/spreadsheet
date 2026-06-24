use calamine::{
    open_workbook_auto, open_workbook_auto_from_rs, Data, Reader, SheetVisible, Sheets,
};
use rustler::{Binary, NifTaggedEnum};
use std::io::{Cursor, Read, Seek};

/// A parsed sheet: rows of cells.
type Sheet = Vec<Vec<ColumnData>>;
/// A sheet paired with its name, as returned by the bulk parsers.
type NamedSheets = Vec<(String, Sheet)>;

// The path- and binary-based NIFs differ only in how the workbook is opened;
// both yield a `Sheets<RS>`, so the actual work lives in these generic helpers.

fn collect_sheet_names<RS: Read + Seek>(workbook: &Sheets<RS>, show_hidden: bool) -> Vec<String> {
    workbook
        .sheets_metadata()
        .iter()
        .filter(|sheet| show_hidden || sheet.visible == SheetVisible::Visible)
        .map(|sheet| sheet.name.clone())
        .collect()
}

fn parse_sheet<RS: Read + Seek>(
    workbook: &mut Sheets<RS>,
    sheet_name: &str,
) -> Result<Sheet, String> {
    workbook
        .worksheet_range(sheet_name)
        .map_err(|e| e.to_string())
        .map(extract_rows)
}

fn parse_all_sheets<RS: Read + Seek>(
    mut workbook: Sheets<RS>,
    show_hidden: bool,
) -> Result<NamedSheets, String> {
    collect_sheet_names(&workbook, show_hidden)
        .into_iter()
        .map(|name| parse_sheet(&mut workbook, &name).map(|rows| (name, rows)))
        .collect()
}

#[rustler::nif(schedule = "DirtyCpu")]
fn sheet_names_from_binary(content: Binary, show_hidden: bool) -> Result<Vec<String>, String> {
    let workbook =
        open_workbook_auto_from_rs(Cursor::new(content.as_slice())).map_err(|e| e.to_string())?;
    Ok(collect_sheet_names(&workbook, show_hidden))
}

#[rustler::nif(schedule = "DirtyCpu")]
fn sheet_names_from_path(path: &str, show_hidden: bool) -> Result<Vec<String>, String> {
    let workbook = open_workbook_auto(path).map_err(|e| e.to_string())?;
    Ok(collect_sheet_names(&workbook, show_hidden))
}

#[rustler::nif(schedule = "DirtyCpu")]
fn parse_from_path(path: &str, sheet_name: &str) -> Result<Sheet, String> {
    let mut workbook = open_workbook_auto(path).map_err(|e| e.to_string())?;
    parse_sheet(&mut workbook, sheet_name)
}

#[rustler::nif(schedule = "DirtyCpu")]
fn parse_from_binary(content: Binary, sheet_name: &str) -> Result<Sheet, String> {
    let mut workbook =
        open_workbook_auto_from_rs(Cursor::new(content.as_slice())).map_err(|e| e.to_string())?;
    parse_sheet(&mut workbook, sheet_name)
}

#[rustler::nif(schedule = "DirtyCpu")]
fn parse_all_from_path(path: &str, show_hidden: bool) -> Result<NamedSheets, String> {
    let workbook = open_workbook_auto(path).map_err(|e| e.to_string())?;
    parse_all_sheets(workbook, show_hidden)
}

#[rustler::nif(schedule = "DirtyCpu")]
fn parse_all_from_binary(content: Binary, show_hidden: bool) -> Result<NamedSheets, String> {
    let workbook =
        open_workbook_auto_from_rs(Cursor::new(content.as_slice())).map_err(|e| e.to_string())?;
    parse_all_sheets(workbook, show_hidden)
}

#[derive(NifTaggedEnum, Debug)]
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

fn extract_rows(range: calamine::Range<Data>) -> Sheet {
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
        // Calamine represents duration-formatted cells (e.g. `[h]:mm:ss`) as a
        // `DateTime` with a `TimeDelta` type. Feeding those to `as_datetime()`
        // yields a nonsensical date, so route them to a duration instead.
        Data::DateTime(val) if val.is_duration() => match val.as_duration() {
            Some(dur) => ColumnData::DurationIso(duration_to_iso(dur)),
            None => ColumnData::Empty,
        },
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

/// Format a chrono `Duration` as an ISO 8601 duration (e.g. `PT1H30M0S`),
/// matching the shape calamine already uses for ODS `DurationIso` cells.
fn duration_to_iso(dur: chrono::Duration) -> String {
    let total_ms = dur.num_milliseconds().max(0);
    let ms = total_ms % 1000;
    let total_secs = total_ms / 1000;
    let secs = total_secs % 60;
    let mins = (total_secs / 60) % 60;
    let hours = total_secs / 3600;

    if ms > 0 {
        format!("PT{hours}H{mins}M{secs}.{ms:03}S")
    } else {
        format!("PT{hours}H{mins}M{secs}S")
    }
}

rustler::init!("Elixir.Spreadsheet.Calamine");

#[cfg(test)]
mod tests {
    use super::*;
    use calamine::{ExcelDateTime, ExcelDateTimeType};

    #[test]
    fn datetime_cells_become_datetimes() {
        // 2025/10/13 12:00:00 is serial 45943.5.
        let dt = ExcelDateTime::new(45943.5, ExcelDateTimeType::DateTime, false);
        match extract_column(&Data::DateTime(dt)) {
            ColumnData::DateTime(s) => assert_eq!(s, "2025-10-13T12:00:00"),
            other => panic!("expected DateTime, got {other:?}"),
        }
    }

    #[test]
    fn duration_cells_are_not_misread_as_datetimes() {
        // 1.5 hours = 0.0625 of a day, formatted as a duration ([h]:mm).
        let dur = ExcelDateTime::new(0.0625, ExcelDateTimeType::TimeDelta, false);
        match extract_column(&Data::DateTime(dur)) {
            ColumnData::DurationIso(s) => assert_eq!(s, "PT1H30M0S"),
            other => panic!("expected DurationIso, got {other:?}"),
        }
    }
}
