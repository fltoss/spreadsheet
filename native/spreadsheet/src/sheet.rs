//! Turning a sheet's cells into a list of rows anchored at A1: row 0 is
//! spreadsheet row 1 and column 0 is column A, whatever cell the data starts
//! in. Leading empty rows and columns are padded with `nil`.
//!
//! The output rectangle is checked against `:max_cells` before it is built.
//! For .xlsx and .xlsb the cells are streamed from calamine's cell readers and
//! the rectangle is computed from the cells themselves (the `<dimension>`
//! element can be forged), instead of letting `Range::from_sparse` allocate
//! the rectangle first.

use crate::cell::ToCellValue;
use crate::encode::encode_cell;
use crate::limits::CellLimit;
use calamine::{Cell, Data, DataRef, Range, Reader, SheetType, Sheets, XlsxError};
use rustler::{Encoder, Env, Term};
use std::fmt::Display;
use std::io::{Read, Seek};

/// Reads sheet `name` and encodes its rows.
pub fn read_sheet<'a, RS: Read + Seek>(
    env: Env<'a>,
    workbook: &mut Sheets<RS>,
    name: &str,
    limit: CellLimit,
) -> Result<Term<'a>, String> {
    let sheet_type = workbook
        .sheets_metadata()
        .iter()
        .find(|sheet| sheet.name == name)
        .map(|sheet| sheet.typ)
        .ok_or_else(|| format!("sheet {name:?} not found"))?;

    // Chart sheets and VBA modules hold no cells.
    if matches!(sheet_type, SheetType::ChartSheet | SheetType::Vba) {
        return Ok(empty_list(env));
    }

    match workbook {
        Sheets::Xlsx(xlsx) => match xlsx.worksheet_cells_reader(name) {
            Ok(mut reader) => {
                let cells = collect_cells(|| reader.next_cell(), "Xlsx")?;
                encode_cells(env, name, cells, limit)
            }
            // calamine treats other sheet kinds as empty too.
            Err(XlsxError::NotAWorksheet(_)) => Ok(empty_list(env)),
            Err(error) => Err(format!("Xlsx error: {error}")),
        },
        Sheets::Xlsb(xlsb) => {
            let mut reader = xlsb
                .worksheet_cells_reader(name)
                .map_err(|error| format!("Xlsb error: {error}"))?;
            let cells = collect_cells(|| reader.next_cell(), "Xlsb")?;
            encode_cells(env, name, cells, limit)
        }
        Sheets::Xls(xls) => {
            let range = xls
                .worksheet_range(name)
                .map_err(|error| format!("Xls error: {error}"))?;
            encode_range(env, name, &range, limit)
        }
        Sheets::Ods(ods) => {
            let range = ods
                .worksheet_range(name)
                .map_err(|error| format!("Ods error: {error}"))?;
            encode_range(env, name, &range, limit)
        }
    }
}

fn empty_list(env: Env) -> Term {
    Vec::<Term>::new().encode(env)
}

/// Collects a sheet's non-empty cells from a calamine cell reader.
fn collect_cells<'cell, E: Display>(
    mut next_cell: impl FnMut() -> Result<Option<Cell<DataRef<'cell>>>, E>,
    format: &str,
) -> Result<Vec<Cell<DataRef<'cell>>>, String> {
    let mut cells = Vec::new();
    while let Some(cell) = next_cell().map_err(|error| format!("{format} error: {error}"))? {
        if *cell.get_value() != DataRef::Empty {
            cells.push(cell);
        }
    }
    Ok(cells)
}

/// Encodes sparse cells (.xlsx, .xlsb) once their rectangle passes the limit.
fn encode_cells<'a>(
    env: Env<'a>,
    name: &str,
    mut cells: Vec<Cell<DataRef<'_>>>,
    limit: CellLimit,
) -> Result<Term<'a>, String> {
    let Some((last_row, last_column)) = cells.iter().fold(None, |extent, cell| {
        let (row, column) = cell.get_position();
        let (last_row, last_column) = extent.unwrap_or((0, 0));
        Some((last_row.max(row), last_column.max(column)))
    }) else {
        return Ok(empty_list(env));
    };
    let rows = u64::from(last_row) + 1;
    let columns = u64::from(last_column) + 1;
    limit.check_rectangle(name, rows, columns)?;

    // Readers yield cells in file order, which is normally row-major but not
    // guaranteed. A stable sort keeps the last of duplicate cells last, and
    // the last one wins, as in calamine's `Range::from_sparse`.
    cells.sort_by_key(|cell| cell.get_position());

    let mut encoder = RowEncoder::new(env, columns as usize);
    let mut encoded = Vec::with_capacity(rows as usize);
    let mut remaining = cells.as_slice();
    for row in 0..=last_row {
        let count = remaining
            .iter()
            .take_while(|cell| cell.get_position().0 == row)
            .count();
        let (row_cells, rest) = remaining.split_at(count);
        remaining = rest;
        if row_cells.is_empty() {
            encoded.push(encoder.empty_row());
        } else {
            encoded.push(encoder.sparse_row(row_cells));
        }
    }
    Ok(encoded.encode(env))
}

/// Encodes a calamine range (.xls and .ods, which calamine loads up front).
fn encode_range<'a>(
    env: Env<'a>,
    name: &str,
    range: &Range<Data>,
    limit: CellLimit,
) -> Result<Term<'a>, String> {
    let (Some(start), Some(end)) = (range.start(), range.end()) else {
        return Ok(empty_list(env));
    };
    let rows = u64::from(end.0) + 1;
    let columns = u64::from(end.1) + 1;
    limit.check_rectangle(name, rows, columns)?;

    let mut encoder = RowEncoder::new(env, columns as usize);
    let mut encoded = Vec::with_capacity(rows as usize);
    for _ in 0..start.0 {
        encoded.push(encoder.empty_row());
    }
    for row in range.rows() {
        if row.iter().all(|cell| *cell == Data::Empty) {
            encoded.push(encoder.empty_row());
        } else {
            encoded.push(encoder.dense_row(start.1 as usize, row));
        }
    }
    Ok(encoded.encode(env))
}

/// Encodes rows of terms, sharing one all-`nil` row term for every row that
/// has no cells.
struct RowEncoder<'a> {
    env: Env<'a>,
    nil: Term<'a>,
    columns: usize,
    row: Vec<Term<'a>>,
    empty_row: Option<Term<'a>>,
}

impl<'a> RowEncoder<'a> {
    fn new(env: Env<'a>, columns: usize) -> Self {
        let nil = rustler::types::atom::nil().encode(env);
        RowEncoder {
            env,
            nil,
            columns,
            row: Vec::with_capacity(columns),
            empty_row: None,
        }
    }

    fn empty_row(&mut self) -> Term<'a> {
        let (env, nil, columns) = (self.env, self.nil, self.columns);
        *self
            .empty_row
            .get_or_insert_with(|| vec![nil; columns].encode(env))
    }

    fn encode(&self, cell: &impl ToCellValue) -> Term<'a> {
        encode_cell(self.env, cell.to_cell_value(), self.nil)
    }

    /// A row of a dense range whose first cell is at `first_column`.
    fn dense_row(&mut self, first_column: usize, cells: &[Data]) -> Term<'a> {
        self.row.clear();
        self.row.resize(first_column, self.nil);
        for cell in cells {
            let term = self.encode(cell);
            self.row.push(term);
        }
        self.row.resize(self.columns, self.nil);
        self.row.encode(self.env)
    }

    /// A row from cells sorted by column.
    fn sparse_row(&mut self, cells: &[Cell<DataRef<'_>>]) -> Term<'a> {
        self.row.clear();
        self.row.resize(self.columns, self.nil);
        for cell in cells {
            let term = self.encode(cell.get_value());
            self.row[cell.get_position().1 as usize] = term;
        }
        self.row.encode(self.env)
    }
}
