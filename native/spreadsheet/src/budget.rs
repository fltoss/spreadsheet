//! The `:max_cells` budget: the cells one call may build, counted as each
//! sheet's rows × columns from A1 and summed over every sheet it parses.

use crate::error::Error;

#[derive(Clone, Copy, Debug)]
pub struct Budget {
    /// `None` means no limit.
    maximum: Option<u64>,
    /// Cells spent on sheets already built.
    used: u64,
}

impl Budget {
    pub fn new(maximum: Option<u64>) -> Self {
        Budget { maximum, used: 0 }
    }

    /// Cells left to spend, `u64::MAX` without a limit.
    pub fn remaining(&self) -> u64 {
        self.maximum
            .map_or(u64::MAX, |maximum| maximum.saturating_sub(self.used))
    }

    /// Checks that a `rows × columns` rectangle fits in what is left.
    pub fn check_rectangle(&self, sheet: &str, rows: u64, columns: u64) -> Result<(), Error> {
        match self.maximum {
            Some(maximum)
                if u128::from(rows) * u128::from(columns) > u128::from(self.remaining()) =>
            {
                Err(Error::TooManyCells {
                    sheet: sheet.to_string(),
                    rows,
                    columns,
                    used: self.used,
                    maximum,
                })
            }
            _ => Ok(()),
        }
    }

    /// Checks a `rows × columns` rectangle and spends it.
    pub fn spend(&mut self, sheet: &str, rows: u64, columns: u64) -> Result<(), Error> {
        self.check_rectangle(sheet, rows, columns)?;
        self.used = self.used.saturating_add(rows.saturating_mul(columns));
        Ok(())
    }

    /// Checks a count of cell records collected for a sheet, which can exceed
    /// its rectangle when positions repeat.
    pub fn check_records(&self, sheet: &str, count: u64) -> Result<(), Error> {
        match self.maximum {
            Some(maximum) if count > self.remaining() => Err(Error::TooManyCellRecords {
                sheet: sheet.to_string(),
                count,
                used: self.used,
                maximum,
            }),
            _ => Ok(()),
        }
    }

    /// Checks the cells an .ods table makes calamine store while reading it.
    pub fn check_stored_cells(&self, sheet: &str, count: u128) -> Result<(), Error> {
        match self.maximum {
            Some(maximum) if count > u128::from(self.remaining()) => {
                Err(Error::TooManyStoredCells {
                    sheet: sheet.to_string(),
                    count,
                    used: self.used,
                    maximum,
                })
            }
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

    #[test]
    fn the_budget_covers_every_sheet_of_a_call() {
        let mut budget = Budget::new(Some(100));
        budget.spend("A", 6, 10).unwrap();
        assert_eq!(budget.remaining(), 40);
        assert_eq!(
            budget.spend("B", 5, 10).unwrap_err().to_string(),
            "sheet \"B\" has 5 rows × 10 columns (50 cells), above the 40 cells left \
             of the :max_cells limit of 100 after 60 cells in earlier sheets"
        );
        budget.spend("C", 4, 10).unwrap();
        assert_eq!(budget.remaining(), 0);
        assert!(budget.check_records("D", 1).is_err());
    }

    #[test]
    fn unlimited_budgets_never_run_out() {
        let mut budget = Budget::new(None);
        budget.spend("A", u64::MAX, 2).unwrap();
        assert_eq!(budget.remaining(), u64::MAX);
        assert!(budget.check_records("B", u64::MAX).is_ok());
    }
}
