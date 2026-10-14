//! The errors this crate reports. Each one reaches Elixir as
//! `{:error, message}`, worded by the `Display` impl below.

use std::fmt;

#[derive(Debug)]
pub enum Error {
    /// The file could not be opened; `reason` is the I/O error.
    Open {
        path: String,
        reason: String,
    },
    /// The path names a directory.
    Directory {
        path: String,
    },
    /// `format: :filename` was given what looks like file contents.
    ContentGivenAsPath,
    /// A path that is not UTF-8, where paths must be UTF-8 (not on Unix).
    #[cfg_attr(unix, allow(dead_code))]
    PathNotUtf8,
    /// No format could open the content (calamine's own wording).
    UnknownFormat,
    /// `format: :binary` was given what looks like a path; wraps the error
    /// opening it produced.
    PathGivenAsContent(Box<Error>),
    /// calamine could not read the file, worded as calamine words it.
    Calamine(calamine::Error),
    SheetNotFound(String),
    /// A sheet's rectangle (rows × columns from A1) does not fit in what is
    /// left of the `:max_cells` budget after `used` cells of earlier sheets.
    TooManyCells {
        sheet: String,
        rows: u64,
        columns: u64,
        used: u64,
        maximum: u64,
    },
    /// A sheet has more cell records than the `:max_cells` budget has left.
    TooManyCellRecords {
        sheet: String,
        count: u64,
        used: u64,
        maximum: u64,
    },
    /// An .ods table makes calamine store more cells than the `:max_cells`
    /// budget has left.
    TooManyStoredCells {
        sheet: String,
        count: u128,
        used: u64,
        maximum: u64,
    },
    /// An .xlsx or .xlsb sheet repeats cell positions far beyond its
    /// rectangle, which only a malformed file does.
    RepeatedCells {
        sheet: String,
        count: u64,
        rows: u64,
        columns: u64,
    },
    /// A malformed .xls file that would make calamine allocate without bound
    /// or loop forever.
    InvalidXls(String),
    /// An .ods file that would make calamine allocate without bound.
    InvalidOds(String),
    /// calamine panicked.
    Panic(String),
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Error::Open { path, reason } => write!(formatter, "could not open {path}: {reason}"),
            Error::Directory { path } => write!(formatter, "could not open {path}: is a directory"),
            Error::ContentGivenAsPath => formatter.write_str(
                "could not open the given path: the value looks like file contents \
                 rather than a path; did you mean format: :binary?",
            ),
            Error::PathNotUtf8 => formatter.write_str(
                "could not open the given path: it is not valid UTF-8; \
                 did you mean format: :binary?",
            ),
            Error::UnknownFormat => formatter.write_str("Cannot detect file format"),
            Error::PathGivenAsContent(error) => write!(
                formatter,
                "{error} (the content looks like a file path; did you mean format: :filename?)"
            ),
            Error::Calamine(error) => write!(formatter, "{error}"),
            Error::SheetNotFound(sheet) => write!(formatter, "sheet {sheet:?} not found"),
            Error::TooManyCells {
                sheet,
                rows,
                columns,
                used,
                maximum,
            } => write!(
                formatter,
                "sheet {sheet:?} has {rows} rows × {columns} columns ({cells} cells), {limit}",
                cells = u128::from(*rows) * u128::from(*columns),
                limit = Limit(*used, *maximum)
            ),
            Error::TooManyCellRecords {
                sheet,
                count,
                used,
                maximum,
            } => write!(
                formatter,
                "sheet {sheet:?} has {count} cell records, {limit}",
                limit = Limit(*used, *maximum)
            ),
            Error::TooManyStoredCells {
                sheet,
                count,
                used,
                maximum,
            } => write!(
                formatter,
                "sheet {sheet:?} stores {count} cells, {limit}",
                limit = Limit(*used, *maximum)
            ),
            Error::RepeatedCells {
                sheet,
                count,
                rows,
                columns,
            } => write!(
                formatter,
                "invalid sheet {sheet:?}: {count} cell records repeat positions within \
                 {rows} rows × {columns} columns"
            ),
            Error::InvalidXls(reason) => write!(formatter, "invalid .xls file: {reason}"),
            Error::InvalidOds(reason) => write!(formatter, "invalid .ods file: {reason}"),
            Error::Panic(message) => write!(formatter, "could not read workbook: {message}"),
        }
    }
}

impl From<calamine::Error> for Error {
    fn from(error: calamine::Error) -> Self {
        Error::Calamine(error)
    }
}

/// "above the :max_cells limit of M", naming what earlier sheets used.
struct Limit(u64, u64);

impl fmt::Display for Limit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let Limit(used, maximum) = *self;
        if used == 0 {
            write!(formatter, "above the :max_cells limit of {maximum}")
        } else {
            write!(
                formatter,
                "above the {left} cells left of the :max_cells limit of {maximum} \
                 after {used} cells in earlier sheets",
                left = maximum.saturating_sub(used)
            )
        }
    }
}
