//! The `:max_cells` limit, and the checks that enforce it before calamine
//! allocates memory for a sheet.
//!
//! - .xlsx and .xlsb sheets are read through calamine's streaming cell
//!   readers, so the limit is checked on the real cells (see `sheet.rs`).
//! - .xls and .ods workbooks are fully decoded by calamine when they are
//!   opened, allocating each sheet's whole rectangle up front. `xls` and `ods`
//!   pre-scan the raw file and reject it before calamine sees it.

pub mod ods;
pub mod xls;

/// Upper bound on the cells in a sheet's output rectangle (rows × columns,
/// counted from A1). `None` means no limit.
#[derive(Clone, Copy, Debug)]
pub struct CellLimit(pub Option<u64>);

impl CellLimit {
    pub fn is_unlimited(&self) -> bool {
        self.0.is_none()
    }

    /// Checks a `rows × columns` rectangle.
    pub fn check_rectangle(&self, sheet: &str, rows: u64, columns: u64) -> Result<(), String> {
        let cells = u128::from(rows) * u128::from(columns);
        match self.0 {
            Some(maximum) if cells > u128::from(maximum) => Err(format!(
                "sheet {sheet:?} has {rows} rows × {columns} columns ({cells} cells), \
                 above the :max_cells limit of {maximum}"
            )),
            _ => Ok(()),
        }
    }

    /// Checks a count of cells (or characters) a file makes the parser
    /// allocate, described by `what`.
    pub fn check_count(&self, sheet: &str, count: u128, what: &str) -> Result<(), String> {
        match self.0 {
            Some(maximum) if count > u128::from(maximum) => Err(format!(
                "sheet {sheet:?} {what} {count} cells, above the :max_cells limit of {maximum}"
            )),
            _ => Ok(()),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rectangles_above_the_limit_are_rejected() {
        let limit = CellLimit(Some(100));
        assert!(limit.check_rectangle("Sheet1", 10, 10).is_ok());
        assert_eq!(
            limit.check_rectangle("Sheet1", 200_000, 16_384),
            Err(
                "sheet \"Sheet1\" has 200000 rows × 16384 columns (3276800000 cells), \
                 above the :max_cells limit of 100"
                    .to_string()
            )
        );
        assert!(CellLimit(None)
            .check_rectangle("Sheet1", u64::MAX, u64::MAX)
            .is_ok());
    }
}
