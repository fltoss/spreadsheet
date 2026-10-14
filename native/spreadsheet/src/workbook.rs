//! Opening workbooks from a file path or from in-memory content.
//!
//! Both sources go through the same format detection and produce the same
//! [`Workbook`] type, so the NIFs never care where the bytes came from.

use crate::budget::Budget;
use crate::error::Error;
use crate::format::{Format, OpenError};
use calamine::{Reader, SheetVisible, Sheets};
use rustler::{Atom, Binary, Decoder, NifResult, Term};
use std::fs::File;
use std::io::{BufReader, Cursor, Read, Seek};
use std::path::{Path, PathBuf};

mod atoms {
    rustler::atoms! {
        path,
        binary,
    }
}

/// A reader calamine can open: a buffered file or an in-memory slice.
pub trait ReadSeek: Read + Seek {}

impl<T: Read + Seek> ReadSeek for T {}

/// An opened workbook, whatever its source.
pub type Workbook<'a> = Sheets<Box<dyn ReadSeek + 'a>>;

/// Where the workbook comes from, as passed by `Spreadsheet`:
/// `{:path, path}` or `{:binary, content}`.
pub enum Source<'a> {
    Path(Binary<'a>),
    Binary(Binary<'a>),
}

impl<'a> Decoder<'a> for Source<'a> {
    fn decode(term: Term<'a>) -> NifResult<Self> {
        let (tag, data): (Atom, Binary<'a>) = term.decode()?;
        if tag == atoms::path() {
            Ok(Source::Path(data))
        } else if tag == atoms::binary() {
            Ok(Source::Binary(data))
        } else {
            Err(rustler::Error::BadArg)
        }
    }
}

/// Opens the workbook described by `source`. .xls and .ods files are fully
/// decoded when opened, so they are pre-scanned against `budget` first.
pub fn open<'a>(source: &Source<'a>, budget: &Budget) -> Result<Workbook<'a>, Error> {
    match source {
        Source::Path(path) => open_path(&decode_path(path.as_slice())?, budget),
        Source::Binary(content) => open_binary(content.as_slice(), budget),
    }
}

/// Sheet names in workbook order, optionally without hidden sheets.
pub fn sheet_names(workbook: &Workbook, include_hidden: bool) -> Vec<String> {
    workbook
        .sheets_metadata()
        .iter()
        .filter(|sheet| include_hidden || sheet.visible == SheetVisible::Visible)
        .map(|sheet| sheet.name.clone())
        .collect()
}

const ZIP_MAGIC: &[u8] = b"PK\x03\x04";
const OLE_MAGIC: &[u8] = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1";

fn open_path<'a>(path: &Path, budget: &Budget) -> Result<Workbook<'a>, Error> {
    // Open the file up front so a missing or unreadable file reports the real
    // I/O error instead of calamine's "Cannot detect file format".
    let metadata = std::fs::metadata(path).map_err(|error| open_error(path, &error))?;
    if metadata.is_dir() {
        return Err(Error::Directory {
            path: quote_path(path),
        });
    }
    File::open(path).map_err(|error| open_error(path, &error))?;

    let formats = match Format::from_extension(path) {
        Some(format) => vec![format],
        None => Format::DETECTION_ORDER.to_vec(),
    };
    open_with(&formats, budget, || {
        let file = File::open(path).map_err(|error| open_error(path, &error))?;
        Ok(Box::new(BufReader::new(file)))
    })
}

fn open_binary<'a>(content: &'a [u8], budget: &Budget) -> Result<Workbook<'a>, Error> {
    open_with(&Format::DETECTION_ORDER, budget, || {
        Ok(Box::new(Cursor::new(content)))
    })
    .map_err(|error| {
        if looks_like_path(content) {
            Error::PathGivenAsContent(Box::new(error))
        } else {
            error
        }
    })
}

/// Tries each candidate format in turn with a fresh reader. With a single
/// candidate (known extension) its own error is returned, which is more
/// useful than a generic detection failure.
fn open_with<'a>(
    formats: &[Format],
    budget: &Budget,
    mut make_reader: impl FnMut() -> Result<Box<dyn ReadSeek + 'a>, Error>,
) -> Result<Workbook<'a>, Error> {
    let mut last_error = None;
    for format in formats {
        match format.open(make_reader()?, budget) {
            Ok(workbook) => return Ok(workbook),
            Err(OpenError::Rejected(error)) => return Err(error),
            Err(OpenError::NotThisFormat(error)) => last_error = Some(error),
        }
    }
    match (formats.len(), last_error) {
        (1, Some(error)) => Err(Error::Calamine(error)),
        _ => Err(Error::UnknownFormat),
    }
}

/// Paths arrive as raw bytes. On Unix any byte sequence is a valid path, so no
/// UTF-8 requirement is imposed; elsewhere the path must be UTF-8.
fn decode_path(bytes: &[u8]) -> Result<PathBuf, Error> {
    if looks_like_content(bytes) {
        return Err(Error::ContentGivenAsPath);
    }

    #[cfg(unix)]
    {
        use std::os::unix::ffi::OsStrExt;
        Ok(PathBuf::from(std::ffi::OsStr::from_bytes(bytes)))
    }

    #[cfg(not(unix))]
    {
        std::str::from_utf8(bytes)
            .map(PathBuf::from)
            .map_err(|_| Error::PathNotUtf8)
    }
}

/// Spreadsheet contents (zip or OLE containers, or anything with NUL bytes)
/// can never be a usable path.
fn looks_like_content(bytes: &[u8]) -> bool {
    bytes.starts_with(ZIP_MAGIC) || bytes.starts_with(OLE_MAGIC) || bytes.contains(&0)
}

/// Short, printable UTF-8 ending in a spreadsheet extension.
fn looks_like_path(content: &[u8]) -> bool {
    match std::str::from_utf8(content) {
        Ok(text) if text.len() <= 4096 && !text.chars().any(char::is_control) => {
            Format::from_extension(Path::new(text)).is_some()
        }
        _ => false,
    }
}

fn open_error(path: &Path, error: &std::io::Error) -> Error {
    let message = error.to_string();
    // "No such file or directory (os error 2)" -> "No such file or directory"
    let reason = message.split(" (os error").next().unwrap_or(&message);
    Error::Open {
        path: quote_path(path),
        reason: reason.to_string(),
    }
}

fn quote_path(path: &Path) -> String {
    format!("{:?}", path.to_string_lossy())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn spreadsheet_contents_are_not_paths() {
        assert!(looks_like_content(b"PK\x03\x04rest of a zip"));
        assert!(looks_like_content(OLE_MAGIC));
        assert!(looks_like_content(b"a\0b"));
        assert!(!looks_like_content(b"/tmp/report.xlsx"));
    }

    #[test]
    fn short_strings_with_extensions_look_like_paths() {
        assert!(looks_like_path(b"/tmp/Report.XLSX"));
        assert!(looks_like_path(b"data.ods"));
        assert!(!looks_like_path(b"PK\x03\x04"));
        assert!(!looks_like_path(b"report.csv"));
    }

    #[test]
    #[cfg(unix)]
    fn missing_files_report_the_io_error() {
        let error = open_path(Path::new("/definitely/not/here.xlsx"), &Budget::new(None))
            .err()
            .unwrap();
        assert_eq!(
            error.to_string(),
            "could not open \"/definitely/not/here.xlsx\": No such file or directory"
        );
    }
}
