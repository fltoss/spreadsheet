//! Turning a sheet's cells into a list of rows anchored at A1: row 0 is
//! spreadsheet row 1 and column 0 is column A, whatever cell the data starts
//! in. Leading empty rows and columns are padded with `nil`.

use crate::cell::ToCellValue;
use crate::encode::encode_cell;
use calamine::{Data, Range};
use rustler::{Encoder, Env, Term};

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

    /// Encodes a row whose cells start at column `first_column`.
    fn row<'cell, T: ToCellValue + 'cell>(
        &mut self,
        first_column: usize,
        cells: impl Iterator<Item = &'cell T>,
    ) -> Term<'a> {
        self.row.clear();
        self.row.resize(first_column, self.nil);
        for cell in cells {
            self.row
                .push(encode_cell(self.env, cell.to_cell_value(), self.nil));
        }
        self.row.resize(self.columns, self.nil);
        self.row.encode(self.env)
    }
}

/// Encodes a calamine range (.xls and .ods, which calamine loads up front).
pub fn encode_range<'a>(env: Env<'a>, range: &Range<Data>) -> Term<'a> {
    let (Some(start), Some(end)) = (range.start(), range.end()) else {
        return Vec::<Term<'a>>::new().encode(env);
    };
    let (first_row, first_column) = (start.0 as usize, start.1 as usize);
    let rows = end.0 as usize + 1;
    let columns = end.1 as usize + 1;

    let mut encoder = RowEncoder::new(env, columns);
    let mut encoded = Vec::with_capacity(rows);
    for _ in 0..first_row {
        encoded.push(encoder.empty_row());
    }
    for row in range.rows() {
        if row.iter().all(|cell| *cell == Data::Empty) {
            encoded.push(encoder.empty_row());
        } else {
            encoded.push(encoder.row(first_column, row.iter()));
        }
    }
    encoded.encode(env)
}
