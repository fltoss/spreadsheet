//! Turning a sheet's cells into a list of rows anchored at A1: row 0 is
//! spreadsheet row 1 and column 0 is column A, whatever cell the data starts
//! in. Leading empty rows and columns are padded with `nil`.
//!
//! Each sheet's rectangle is spent from the `:max_cells` budget before it is
//! built.

use crate::budget::Budget;
use crate::cell::ToCellValue;
use crate::encode::encode_cell;
use crate::error::Error;
use crate::format::{self, SheetCells, SparseCells};
use crate::workbook::Workbook;
use calamine::{Cell, Data, DataRef, Range};
use rustler::{Encoder, Env, Term};

/// Reads sheet `name` and encodes its rows.
pub fn read_sheet<'a>(
    env: Env<'a>,
    workbook: &mut Workbook,
    name: &str,
    budget: &mut Budget,
) -> Result<Term<'a>, Error> {
    match format::read_cells(workbook, name, budget)? {
        SheetCells::Empty => Ok(empty_list(env)),
        SheetCells::Sparse(sparse) => encode_cells(env, name, sparse, budget),
        SheetCells::Range(range) => encode_range(env, name, &range, budget),
    }
}

fn empty_list(env: Env) -> Term {
    Vec::<Term>::new().encode(env)
}

/// Encodes sparse cells (.xlsx, .xlsb).
fn encode_cells<'a>(
    env: Env<'a>,
    name: &str,
    sparse: SparseCells<'_>,
    budget: &mut Budget,
) -> Result<Term<'a>, Error> {
    let SparseCells {
        mut cells,
        rows,
        columns,
        sorted,
    } = sparse;
    budget.spend(name, rows, columns)?;

    // Readers yield cells in file order, which is normally row-major but not
    // guaranteed. A stable sort keeps the last of duplicate cells last, and
    // the last one wins, as in calamine's `Range::from_sparse`.
    if !sorted {
        cells.sort_by_key(|cell| cell.get_position());
    }

    let mut encoder = RowEncoder::new(env, columns as usize);
    let mut encoded = Vec::with_capacity(rows as usize);
    let mut remaining = cells.as_slice();
    for row in 0..rows as u32 {
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
    budget: &mut Budget,
) -> Result<Term<'a>, Error> {
    let (Some(start), Some(end)) = (range.start(), range.end()) else {
        return Ok(empty_list(env));
    };
    let rows = u64::from(end.0) + 1;
    let columns = u64::from(end.1) + 1;
    budget.spend(name, rows, columns)?;

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
