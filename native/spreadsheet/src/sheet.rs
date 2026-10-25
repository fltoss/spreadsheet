//! Turning a sheet's cells into a list of rows. With `Anchor::A1`, row 0 is
//! spreadsheet row 1 and column 0 is column A, whatever cell the data starts
//! in, and leading empty rows and columns are padded with `nil`. With
//! `Anchor::Data`, row 0 and column 0 are the first that hold a value.
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
use rustler::{Encoder, Env, NifUnitEnum, Term};
use std::collections::HashMap;

/// Where a sheet's rows start: the `:anchor` option, `:a1` or `:data`.
#[derive(Clone, Copy, Debug, PartialEq, NifUnitEnum)]
pub enum Anchor {
    A1,
    Data,
}

/// Reads sheet `name` and encodes its rows.
pub fn read_sheet<'a>(
    env: Env<'a>,
    workbook: &mut Workbook,
    name: &str,
    budget: &mut Budget,
    anchor: Anchor,
) -> Result<Term<'a>, Error> {
    match format::read_cells(workbook, name, budget, anchor)? {
        SheetCells::Empty => Ok(empty_list(env)),
        SheetCells::Sparse(sparse) => encode_cells(env, name, sparse, budget),
        SheetCells::Range(range) => encode_range(env, name, &range, budget, anchor),
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
        first_row,
        first_column,
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
    for row in first_row..first_row + rows as u32 {
        let count = remaining
            .iter()
            .take_while(|cell| cell.get_position().0 == row)
            .count();
        let (row_cells, rest) = remaining.split_at(count);
        remaining = rest;
        if row_cells.is_empty() {
            encoded.push(encoder.empty_row());
        } else {
            encoded.push(encoder.sparse_row(first_column, row_cells));
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
    anchor: Anchor,
) -> Result<Term<'a>, Error> {
    let (Some(start), Some(end)) = (range.start(), range.end()) else {
        return Ok(empty_list(env));
    };
    let (first_row, first_column) = match anchor {
        Anchor::A1 => (0, 0),
        Anchor::Data => match range_origin(range) {
            Some(origin) => origin,
            None => return Ok(empty_list(env)),
        },
    };
    let rows = u64::from(end.0 - first_row) + 1;
    let columns = u64::from(end.1 - first_column) + 1;
    budget.spend(name, rows, columns)?;

    // With `Anchor::A1` the range is padded up to A1; with `Anchor::Data` its
    // leading empty rows and columns are dropped.
    let skipped_rows = first_row.saturating_sub(start.0) as usize;
    let skipped_columns = first_column.saturating_sub(start.1) as usize;
    let padding_columns = start.1.saturating_sub(first_column) as usize;

    let mut encoder = RowEncoder::new(env, columns as usize);
    let mut encoded = Vec::with_capacity(rows as usize);
    for _ in first_row..start.0 {
        encoded.push(encoder.empty_row());
    }
    for row in range.rows().skip(skipped_rows) {
        if row.iter().all(|cell| *cell == Data::Empty) {
            encoded.push(encoder.empty_row());
        } else {
            encoded.push(encoder.dense_row(padding_columns, &row[skipped_columns..]));
        }
    }
    Ok(encoded.encode(env))
}

/// The absolute row and column of a range's first non-empty row and column,
/// or `None` when every cell is empty.
fn range_origin(range: &Range<Data>) -> Option<(u32, u32)> {
    let start = range.start()?;
    range
        .used_cells()
        .map(|(row, column, _)| (row as u32, column as u32))
        .reduce(|(first_row, first_column), (row, column)| {
            (first_row.min(row), first_column.min(column))
        })
        .map(|(row, column)| (start.0 + row, start.1 + column))
}

/// Encodes rows of terms, sharing one all-`nil` row term for every row that
/// has no cells, and one term for every use of the same shared string.
struct RowEncoder<'a> {
    env: Env<'a>,
    nil: Term<'a>,
    columns: usize,
    row: Vec<Term<'a>>,
    empty_row: Option<Term<'a>>,
    /// Terms of shared strings, keyed by their address in the workbook's
    /// shared string table, which outlives the encoder.
    shared_strings: HashMap<(usize, usize), Term<'a>>,
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
            shared_strings: HashMap::new(),
        }
    }

    fn empty_row(&mut self) -> Term<'a> {
        let (env, nil, columns) = (self.env, self.nil, self.columns);
        *self
            .empty_row
            .get_or_insert_with(|| vec![nil; columns].encode(env))
    }

    fn shared_string(&mut self, text: &str) -> Term<'a> {
        let env = self.env;
        *self
            .shared_strings
            .entry((text.as_ptr() as usize, text.len()))
            .or_insert_with(|| text.encode(env))
    }

    fn encode(&self, cell: &impl ToCellValue) -> Term<'a> {
        encode_cell(self.env, cell.to_cell_value(), self.nil)
    }

    /// A row of a dense range, after `padding_columns` empty cells.
    fn dense_row(&mut self, padding_columns: usize, cells: &[Data]) -> Term<'a> {
        self.row.clear();
        self.row.resize(padding_columns, self.nil);
        for cell in cells {
            let term = self.encode(cell);
            self.row.push(term);
        }
        self.row.resize(self.columns, self.nil);
        self.row.encode(self.env)
    }

    /// A row from cells sorted by column, whose first column is `first_column`.
    fn sparse_row(&mut self, first_column: u32, cells: &[Cell<DataRef<'_>>]) -> Term<'a> {
        self.row.clear();
        self.row.resize(self.columns, self.nil);
        for cell in cells {
            let term = match cell.get_value() {
                DataRef::SharedString(text) => self.shared_string(text),
                value => self.encode(value),
            };
            self.row[(cell.get_position().1 - first_column) as usize] = term;
        }
        self.row.encode(self.env)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_range_starts_at_its_first_non_empty_row_and_column() {
        let mut range = Range::new((1, 1), (5, 4));
        range.set_value((3, 3), Data::Float(1.0));
        range.set_value((2, 4), Data::String("a".to_string()));
        assert_eq!(range_origin(&range), Some((2, 3)));

        let empty: Range<Data> = Range::new((1, 1), (2, 2));
        assert_eq!(range_origin(&empty), None);
    }
}
