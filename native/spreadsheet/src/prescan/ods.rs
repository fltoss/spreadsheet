//! Pre-scan of .ods files.
//!
//! calamine decodes every table of an .ods workbook while opening it. Repeat
//! attributes are expanded into real cells: `table:number-columns-repeated`
//! while reading each row, and `table:number-rows-repeated` when the table's
//! rectangle is built (capped at 100 million cells, still several GB). A few
//! KB of compressed XML can therefore exhaust memory before a single sheet is
//! returned, and `<text:s text:c="…"/>` expands to that many spaces.
//!
//! This module streams `content.xml` the way calamine 0.36 reads it and
//! rejects the workbook when:
//!
//! - a table's output rectangle (rows × columns counted from A1 up to the last
//!   non-empty cell, with repeats expanded) or the cells calamine stores while
//!   reading its rows exceed the `:max_cells` budget;
//! - `<text:s>` elements expand to more than [`MAX_SPACES`] spaces across the
//!   workbook, repeats included. This is independent of `:max_cells`, which
//!   counts cells, not characters.
//!
//! Anything calamine would fail on by itself is left to calamine.

use crate::budget::Budget;
use crate::error::Error;
use quick_xml::events::{BytesStart, Event};
use quick_xml::Reader;
use std::io::{BufRead, BufReader, Read, Seek, SeekFrom};
use zip::ZipArchive;

/// calamine caps rows at Excel's maximum.
const MAX_ROWS: u64 = 1_048_576;
/// calamine caps columns at Excel's maximum.
const MAX_COLUMNS: u64 = 16_384;

/// The most spaces `<text:s>` elements may expand to in one workbook (100 MB
/// of text). Real files use a few spaces per cell.
pub const MAX_SPACES: u128 = 100_000_000;

const VALUE_ATTRIBUTES: [&[u8]; 5] = [
    b"office:value",
    b"office:string-value",
    b"office:date-value",
    b"office:time-value",
    b"office:boolean-value",
];

pub fn check<RS: Read + Seek + ?Sized>(reader: &mut RS, budget: &Budget) -> Result<(), Error> {
    let result = scan(reader, budget);
    // calamine expects to start reading from the beginning.
    let _ = reader.seek(SeekFrom::Start(0));
    result
}

fn scan<RS: Read + Seek + ?Sized>(reader: &mut RS, budget: &Budget) -> Result<(), Error> {
    let Ok(mut archive) = ZipArchive::new(reader) else {
        return Ok(());
    };
    let Ok(content) = archive.by_name("content.xml") else {
        return Ok(());
    };
    scan_content(BufReader::new(content), budget)
}

/// Scans `content.xml`.
fn scan_content<R: BufRead>(content: R, budget: &Budget) -> Result<(), Error> {
    let mut xml = Reader::from_reader(content);
    let config = xml.config_mut();
    config.check_end_names = false;
    config.trim_text(false);
    config.check_comments = false;
    config.expand_empty_elements = true;

    let mut spaces: u128 = 0;
    let mut buffer = Vec::with_capacity(1024);
    loop {
        buffer.clear();
        match xml.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) if element.name().as_ref() == b"table:table" => {
                let Some(name) = attribute(&element, b"table:name") else {
                    continue;
                };
                let name = String::from_utf8_lossy(&name).into_owned();
                let Some(table) = scan_table(&mut xml)? else {
                    return Ok(());
                };
                table.check(&name, budget)?;
                spaces += table.spaces;
                if spaces > MAX_SPACES {
                    return Err(Error::InvalidOds(format!(
                        "<text:s> elements expand to more than {MAX_SPACES} spaces \
                         (reached in sheet {name:?})"
                    )));
                }
            }
            Ok(Event::Eof) | Err(_) => return Ok(()),
            Ok(_) => (),
        }
    }
}

fn attribute(element: &BytesStart, key: &[u8]) -> Option<Vec<u8>> {
    element
        .attributes()
        .with_checks(false)
        .flatten()
        .find(|attribute| attribute.key.as_ref() == key)
        .map(|attribute| attribute.value.into_owned())
}

/// A repeat count, or `None` when it does not parse (calamine fails then).
fn repeat_count(element: &BytesStart, key: &[u8]) -> Option<u64> {
    match attribute(element, key) {
        Some(value) => std::str::from_utf8(&value).ok()?.parse().ok(),
        None => Some(1),
    }
}

#[derive(Default, Debug, PartialEq)]
struct Table {
    /// Rows up to and including the last row with content.
    rows: u64,
    /// Columns up to and including the last column with content.
    columns: u64,
    /// Cells calamine pushes while reading rows.
    stored_cells: u128,
    /// Spaces produced by `<text:s>` elements, repeats included.
    spaces: u128,
}

impl Table {
    fn check(&self, name: &str, budget: &Budget) -> Result<(), Error> {
        budget.check_rectangle(name, self.rows, self.columns)?;
        budget.check_stored_cells(name, self.stored_cells)
    }
}

/// Reads one `table:table`. `Ok(None)` means calamine would fail on it.
fn scan_table<R: BufRead>(xml: &mut Reader<R>) -> Result<Option<Table>, Error> {
    let mut table = Table::default();
    let mut row_index: u64 = 0;
    let mut buffer = Vec::with_capacity(1024);
    loop {
        buffer.clear();
        match xml.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) if element.name().as_ref() == b"table:table-row" => {
                let Some(repeats) = repeat_count(&element, b"table:number-rows-repeated") else {
                    return Ok(None);
                };
                let repeats = repeats.min(MAX_ROWS.saturating_sub(row_index));
                let Some(row) = scan_row(xml)? else {
                    return Ok(None);
                };
                table.stored_cells += u128::from(row.stored_cells);
                table.spaces += row.spaces * u128::from(repeats.max(1));
                if row.content_columns > 0 && repeats > 0 {
                    table.rows = row_index + repeats;
                    table.columns = table.columns.max(row.content_columns);
                }
                row_index += repeats;
            }
            Ok(Event::End(element)) if element.name().as_ref() == b"table:table" => {
                return Ok(Some(table));
            }
            Ok(Event::Eof) | Err(_) => return Ok(None),
            Ok(_) => (),
        }
    }
}

#[derive(Default)]
struct Row {
    stored_cells: u64,
    content_columns: u64,
    spaces: u128,
}

/// Mirrors calamine's `read_row`: empty cells are only stored when a later
/// cell follows them, and repeats are capped at 16,384 columns.
fn scan_row<R: BufRead>(xml: &mut Reader<R>) -> Result<Option<Row>, Error> {
    let mut row = Row::default();
    let mut pending_empty: u64 = 0;
    let mut buffer = Vec::with_capacity(1024);
    loop {
        buffer.clear();
        match xml.read_event_into(&mut buffer) {
            Ok(Event::Start(element))
                if matches!(
                    element.name().as_ref(),
                    b"table:table-cell" | b"table:covered-table-cell"
                ) =>
            {
                let Some(repeats) = repeat_count(&element, b"table:number-columns-repeated") else {
                    return Ok(None);
                };
                let has_content = cell_has_content(&element);
                let end = element.name().as_ref().to_vec();
                let Some(spaces) = scan_cell(xml, &end)? else {
                    return Ok(None);
                };

                row.stored_cells += pending_empty.min(MAX_COLUMNS - row.stored_cells);
                pending_empty = 0;
                let repeats = repeats.min(MAX_COLUMNS - row.stored_cells);
                if has_content {
                    row.stored_cells += repeats;
                    row.content_columns = row.stored_cells;
                    row.spaces += spaces * u128::from(repeats);
                } else {
                    pending_empty = repeats;
                }
            }
            Ok(Event::End(element)) if element.name().as_ref() == b"table:table-row" => {
                return Ok(Some(row));
            }
            Ok(Event::Eof) | Err(_) => return Ok(None),
            Ok(_) => (),
        }
    }
}

/// A cell calamine stores a value or formula for.
fn cell_has_content(element: &BytesStart) -> bool {
    element
        .attributes()
        .with_checks(false)
        .flatten()
        .any(|attribute| {
            let key = attribute.key.as_ref();
            VALUE_ATTRIBUTES.contains(&key)
                || (key == b"office:value-type" && attribute.value.as_ref() == b"string")
                || (key == b"table:formula" && !attribute.value.is_empty())
        })
}

/// Skips to the end of the cell, adding up `<text:s text:c="n"/>` spaces.
fn scan_cell<R: BufRead>(xml: &mut Reader<R>, end: &[u8]) -> Result<Option<u128>, Error> {
    let mut spaces: u128 = 0;
    let mut buffer = Vec::with_capacity(256);
    loop {
        buffer.clear();
        match xml.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) if element.name().as_ref() == b"text:s" => {
                let Some(count) = repeat_count(&element, b"text:c") else {
                    return Ok(None);
                };
                spaces += u128::from(count);
            }
            Ok(Event::End(element)) if element.name().as_ref() == end => return Ok(Some(spaces)),
            Ok(Event::Eof) | Err(_) => return Ok(None),
            Ok(_) => (),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn table(rows: &str) -> Table {
        let xml = format!("<table:table table:name=\"T\">{rows}</table:table>");
        let mut reader = Reader::from_reader(xml.as_bytes());
        reader.config_mut().expand_empty_elements = true;
        let mut buffer = Vec::new();
        reader.read_event_into(&mut buffer).unwrap();
        scan_table(&mut reader).unwrap().unwrap()
    }

    const VALUE: &str = r#"office:value-type="float" office:value="1""#;

    #[test]
    fn rectangles_count_from_a1_with_repeats() {
        let rows = format!(
            "<table:table-row table:number-rows-repeated=\"2\"><table:table-cell/></table:table-row>\
             <table:table-row><table:table-cell table:number-columns-repeated=\"3\"/>\
             <table:table-cell {VALUE} table:number-columns-repeated=\"2\"/>\
             <table:table-cell table:number-columns-repeated=\"9\"/></table:table-row>"
        );
        assert_eq!(
            table(&rows),
            Table {
                rows: 3,
                columns: 5,
                stored_cells: 5,
                spaces: 0
            }
        );
    }

    #[test]
    fn repeated_rows_with_content_expand() {
        let rows = format!(
            "<table:table-row table:number-rows-repeated=\"1000\"><table:table-cell {VALUE}/></table:table-row>"
        );
        let table = table(&rows);
        assert_eq!(
            (table.rows, table.columns, table.stored_cells),
            (1000, 1, 1)
        );
    }

    #[test]
    fn repeats_are_capped_like_calamine() {
        let rows = format!(
            "<table:table-row table:number-rows-repeated=\"99999999\"><table:table-cell {VALUE} table:number-columns-repeated=\"99999999\"/></table:table-row>"
        );
        let table = table(&rows);
        assert_eq!((table.rows, table.columns), (MAX_ROWS, MAX_COLUMNS));
    }

    #[test]
    fn repeated_spaces_are_counted() {
        let rows = "<table:table-row><table:table-cell office:value-type=\"string\" \
                    table:number-columns-repeated=\"4\"><text:p>a<text:s text:c=\"1000\"/>b</text:p>\
                    </table:table-cell></table:table-row>";
        assert_eq!(table(rows).spaces, 4000);
    }

    #[test]
    fn space_runs_are_capped_without_a_limit() {
        let cell = |count: u128| {
            format!(
                "<table:table table:name=\"T\"><table:table-row>\
                 <table:table-cell office:value-type=\"string\"><text:p>\
                 <text:s text:c=\"{count}\"/></text:p></table:table-cell>\
                 </table:table-row></table:table>"
            )
        };
        let unlimited = Budget::new(None);
        assert!(scan_content(cell(MAX_SPACES).as_bytes(), &unlimited).is_ok());
        let error = scan_content(cell(MAX_SPACES + 1).as_bytes(), &unlimited).unwrap_err();
        assert_eq!(
            error.to_string(),
            "invalid .ods file: <text:s> elements expand to more than 100000000 spaces \
             (reached in sheet \"T\")"
        );
    }
}
