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
    /// A sheet's rectangle (rows × columns from A1) exceeds `:max_cells`.
    TooManyCells {
        sheet: String,
        rows: u64,
        columns: u64,
        maximum: u64,
    },
    /// A file makes the parser allocate more than `:max_cells` cells (or
    /// characters), described by `what`.
    TooManyAllocated {
        sheet: String,
        what: String,
        count: u128,
        maximum: u64,
    },
    /// A malformed .xls file that would make calamine allocate without bound
    /// or loop forever.
    InvalidXls(String),
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
            Error::UnknownFormat => formatter.write_str(
                "Cannot detect file format: not a valid .xls, .xlsx, .xlsb or .ods file",
            ),
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
                maximum,
            } => write!(
                formatter,
                "sheet {sheet:?} has {rows} rows × {columns} columns ({cells} cells), \
                 above the :max_cells limit of {maximum}",
                cells = u128::from(*rows) * u128::from(*columns)
            ),
            Error::TooManyAllocated {
                sheet,
                what,
                count,
                maximum,
            } => write!(
                formatter,
                "sheet {sheet:?} {what} {count} cells, above the :max_cells limit of {maximum}"
            ),
            Error::InvalidXls(reason) => write!(formatter, "invalid .xls file: {reason}"),
            Error::Panic(message) => write!(formatter, "could not read workbook: {message}"),
        }
    }
}

impl From<calamine::Error> for Error {
    fn from(error: calamine::Error) -> Self {
        Error::Calamine(error)
    }
}
