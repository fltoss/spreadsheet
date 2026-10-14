//! The `:max_cells` budget.

use crate::error::Error;

/// Upper bound on the cells in a sheet's output rectangle (rows × columns,
/// counted from A1). `None` means no limit.
#[derive(Clone, Copy, Debug)]
pub struct Budget {
    maximum: Option<u64>,
}

impl Budget {
    pub fn new(maximum: Option<u64>) -> Self {
        Budget { maximum }
    }

    /// Checks a `rows × columns` rectangle.
    pub fn check_rectangle(&self, sheet: &str, rows: u64, columns: u64) -> Result<(), Error> {
        let cells = u128::from(rows) * u128::from(columns);
        match self.maximum {
            Some(maximum) if cells > u128::from(maximum) => Err(Error::TooManyCells {
                sheet: sheet.to_string(),
                rows,
                columns,
                maximum,
            }),
            _ => Ok(()),
        }
    }

    /// Checks the cells an .ods table makes calamine store while reading it.
    pub fn check_stored_cells(&self, sheet: &str, count: u128) -> Result<(), Error> {
        match self.maximum {
            Some(maximum) if count > u128::from(maximum) => Err(Error::TooManyStoredCells {
                sheet: sheet.to_string(),
                count,
                maximum,
            }),
            _ => Ok(()),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rectangles_above_the_limit_are_rejected() {
        let budget = Budget::new(Some(100));
        assert!(budget.check_rectangle("Sheet1", 10, 10).is_ok());
        assert_eq!(
            budget
                .check_rectangle("Sheet1", 200_000, 16_384)
                .unwrap_err()
                .to_string(),
            "sheet \"Sheet1\" has 200000 rows × 16384 columns (3276800000 cells), \
             above the :max_cells limit of 100"
        );
        assert!(Budget::new(None)
            .check_rectangle("Sheet1", u64::MAX, u64::MAX)
            .is_ok());
    }
}
