//! Opening workbooks from a file path or from in-memory content.
//!
//! Both sources go through the same format detection, so every NIF works on a
//! `calamine::Sheets<RS>` regardless of where the bytes came from.

use crate::limits::{self, CellLimit};
use calamine::{open_workbook_from_rs, Error, Ods, Sheets, Xls, Xlsb, Xlsx};
use rustler::{Atom, Binary, Decoder, NifResult, Term};
use std::any::Any;
use std::fs::File;
use std::io::{BufReader, Cursor, Read, Seek};
use std::panic::{self, AssertUnwindSafe};
use std::path::{Path, PathBuf};

mod atoms {
    rustler::atoms! {
        path,
        binary,
    }
}

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

/// Work to run against an opened workbook. Generic over the reader so the
/// same code serves path- and binary-backed workbooks.
pub trait WorkbookOperation {
    type Output;

    fn run<RS: Read + Seek>(self, workbook: Sheets<RS>) -> Result<Self::Output, String>;
}

/// Opens the workbook described by `source` and runs `operation` on it.
///
/// .xls and .ods files are fully decoded when opened, so `limit` is enforced
/// on them here, before calamine allocates anything.
pub fn with_workbook<O: WorkbookOperation>(
    source: Source,
    limit: CellLimit,
    operation: O,
) -> Result<O::Output, String> {
    // calamine panics on some malformed files (e.g. truncated formula
    // operands in .xls). Report those like any other unreadable file instead
    // of raising an opaque `:nif_panicked` error.
    panic::catch_unwind(AssertUnwindSafe(|| match source {
        Source::Path(path) => {
            let path = decode_path(path.as_slice())?;
            operation.run(open_path(&path, limit)?)
        }
        Source::Binary(content) => operation.run(open_binary(content.as_slice(), limit)?),
    }))
    .unwrap_or_else(|payload| {
        Err(format!(
            "could not read workbook: {}",
            panic_message(&*payload)
        ))
    })
}

fn panic_message(payload: &(dyn Any + Send)) -> &str {
    payload
        .downcast_ref::<&str>()
        .copied()
        .or_else(|| payload.downcast_ref::<String>().map(String::as_str))
        .unwrap_or("the reader panicked")
}

#[derive(Clone, Copy, Debug, PartialEq)]
enum Format {
    Xls,
    Xlsx,
    Xlsb,
    Ods,
}

/// The order calamine's `open_workbook_auto` tries formats in when it cannot
/// rely on the file extension.
const ALL_FORMATS: [Format; 4] = [Format::Xls, Format::Xlsx, Format::Xlsb, Format::Ods];

const ZIP_MAGIC: &[u8] = b"PK\x03\x04";
const OLE_MAGIC: &[u8] = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1";

fn open_path(path: &Path, limit: CellLimit) -> Result<Sheets<BufReader<File>>, String> {
    // Open the file up front so a missing or unreadable file reports the real
    // I/O error instead of calamine's "Cannot detect file format".
    let metadata = std::fs::metadata(path).map_err(|error| open_error(path, &error))?;
    if metadata.is_dir() {
        return Err(format!(
            "could not open {}: is a directory",
            quote_path(path)
        ));
    }
    File::open(path).map_err(|error| open_error(path, &error))?;

    let formats = match path
        .extension()
        .and_then(|extension| extension.to_str())
        .map(str::to_ascii_lowercase)
        .as_deref()
    {
        Some("xls" | "xla") => &[Format::Xls][..],
        Some("xlsx" | "xlsm" | "xlam") => &[Format::Xlsx][..],
        Some("xlsb") => &[Format::Xlsb][..],
        Some("ods") => &[Format::Ods][..],
        _ => &ALL_FORMATS[..],
    };

    open_with(formats, limit, || {
        File::open(path)
            .map(BufReader::new)
            .map_err(|error| open_error(path, &error))
    })
}

fn open_binary(content: &[u8], limit: CellLimit) -> Result<Sheets<Cursor<&[u8]>>, String> {
    open_with(&ALL_FORMATS, limit, || Ok(Cursor::new(content))).map_err(|error| {
        if looks_like_path(content) {
            format!("{error} (the content looks like a file path; did you mean format: :filename?)")
        } else {
            error
        }
    })
}

/// Tries each candidate format in turn with a fresh reader. With a single
/// candidate (known extension) its own error is returned, which is more
/// useful than a generic detection failure.
fn open_with<RS, F>(
    formats: &[Format],
    limit: CellLimit,
    mut make_reader: F,
) -> Result<Sheets<RS>, String>
where
    RS: Read + Seek,
    F: FnMut() -> Result<RS, String>,
{
    let mut last_error = None;
    for format in formats {
        match open_as(*format, make_reader()?, limit) {
            Ok(workbook) => return Ok(workbook),
            Err(OpenError::Rejected(message)) => return Err(message),
            Err(OpenError::Calamine(error)) => last_error = Some(error),
        }
    }
    match (formats.len(), last_error) {
        (1, Some(error)) => Err(error.to_string()),
        _ => Err("Cannot detect file format: not a valid .xls, .xlsx, .xlsb or .ods file".into()),
    }
}

enum OpenError {
    /// The file exceeds a limit; do not try other formats.
    Rejected(String),
    /// calamine could not open the file as this format.
    Calamine(Error),
}

fn open_as<RS: Read + Seek>(
    format: Format,
    mut reader: RS,
    limit: CellLimit,
) -> Result<Sheets<RS>, OpenError> {
    match format {
        Format::Xls => limits::xls::check(&mut reader, limit).map_err(OpenError::Rejected)?,
        Format::Ods => limits::ods::check(&mut reader, limit).map_err(OpenError::Rejected)?,
        Format::Xlsx | Format::Xlsb => (),
    }
    let workbook = match format {
        Format::Xls => open_workbook_from_rs::<Xls<RS>, RS>(reader)
            .map(Sheets::Xls)
            .map_err(Error::Xls),
        Format::Xlsx => open_workbook_from_rs::<Xlsx<RS>, RS>(reader)
            .map(Sheets::Xlsx)
            .map_err(Error::Xlsx),
        Format::Xlsb => open_workbook_from_rs::<Xlsb<RS>, RS>(reader)
            .map(Sheets::Xlsb)
            .map_err(Error::Xlsb),
        Format::Ods => open_workbook_from_rs::<Ods<RS>, RS>(reader)
            .map(Sheets::Ods)
            .map_err(Error::Ods),
    };
    workbook.map_err(OpenError::Calamine)
}

/// Paths arrive as raw bytes. On Unix any byte sequence is a valid path, so no
/// UTF-8 requirement is imposed; elsewhere the path must be UTF-8.
fn decode_path(bytes: &[u8]) -> Result<PathBuf, String> {
    if looks_like_content(bytes) {
        return Err(
            "could not open the given path: the value looks like file contents \
                    rather than a path; did you mean format: :binary?"
                .into(),
        );
    }

    #[cfg(unix)]
    {
        use std::os::unix::ffi::OsStrExt;
        Ok(PathBuf::from(std::ffi::OsStr::from_bytes(bytes)))
    }

    #[cfg(not(unix))]
    {
        std::str::from_utf8(bytes).map(PathBuf::from).map_err(|_| {
            "could not open the given path: it is not valid UTF-8; \
             did you mean format: :binary?"
                .to_string()
        })
    }
}

/// Spreadsheet contents (zip or OLE containers, or anything with NUL bytes)
/// can never be a usable path.
fn looks_like_content(bytes: &[u8]) -> bool {
    bytes.starts_with(ZIP_MAGIC) || bytes.starts_with(OLE_MAGIC) || bytes.contains(&0)
}

/// Short, printable UTF-8 ending in a spreadsheet extension.
fn looks_like_path(content: &[u8]) -> bool {
    const EXTENSIONS: [&str; 7] = [".xls", ".xla", ".xlsx", ".xlsm", ".xlam", ".xlsb", ".ods"];
    match std::str::from_utf8(content) {
        Ok(text) if text.len() <= 4096 && !text.chars().any(char::is_control) => {
            let lowercase = text.to_ascii_lowercase();
            EXTENSIONS
                .iter()
                .any(|extension| lowercase.ends_with(extension))
        }
        _ => false,
    }
}

fn open_error(path: &Path, error: &std::io::Error) -> String {
    let message = error.to_string();
    // "No such file or directory (os error 2)" -> "No such file or directory"
    let message = message.split(" (os error").next().unwrap_or(&message);
    format!("could not open {}: {message}", quote_path(path))
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
        let error = open_path(Path::new("/definitely/not/here.xlsx"), CellLimit(None))
            .err()
            .unwrap();
        assert_eq!(
            error,
            "could not open \"/definitely/not/here.xlsx\": No such file or directory"
        );
    }
}
