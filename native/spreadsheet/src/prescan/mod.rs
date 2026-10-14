//! Pre-scans of the formats calamine decodes in full when it opens them.
//!
//! .xls and .ods workbooks are decoded sheet by sheet while calamine opens
//! them, allocating each sheet's whole rectangle up front, and their
//! containers are read with sizes and loops taken from the file. These scans
//! read the raw file first, the way calamine 0.36 does, and reject it before
//! calamine sees it. .xlsx and .xlsb sheets are streamed instead, so they are
//! checked while their cells are collected (see `format.rs`).
//!
//! When a file is not something calamine would open (wrong container,
//! truncated records, ...) the scans step aside and let calamine report its
//! own error.
//!
//! The structural checks (container loops and sizes, forged DIMENSIONS
//! records, `<text:s>` space runs) always run; only the cell counts depend
//! on `:max_cells`, so `max_cells: :infinity` still rejects malformed files.

mod biff;
mod cfb;
mod ods;

use crate::budget::Budget;
use crate::error::Error;
use std::io::{Read, Seek, SeekFrom};

pub use ods::check as ods;

/// `Ok(None)` means "not applicable": calamine will reject the file itself.
type Scan<T> = Result<Option<T>, Error>;

/// Rejects .xls files whose compound file container or DIMENSIONS records
/// would make calamine allocate without bound or loop forever (always), or
/// whose sheets exceed the `:max_cells` budget.
pub fn xls<RS: Read + Seek + ?Sized>(reader: &mut RS, budget: &Budget) -> Result<(), Error> {
    let result = cfb::workbook_stream(reader).and_then(|stream| match stream {
        Some(stream) => biff::scan_workbook(&stream, budget),
        None => Ok(None),
    });
    // calamine expects to start reading from the beginning.
    let _ = reader.seek(SeekFrom::Start(0));
    result.map(|_| ())
}

fn invalid_xls(reason: &str) -> Error {
    Error::InvalidXls(reason.to_string())
}

fn read_u16(bytes: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([bytes[offset], bytes[offset + 1]])
}

fn read_u32(bytes: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        bytes[offset],
        bytes[offset + 1],
        bytes[offset + 2],
        bytes[offset + 3],
    ])
}

#[cfg(test)]
mod tests {
    use super::*;

    /// The pre-scans mirror calamine's internals and read .ods files with the
    /// same quick-xml and zip. A version change needs a new audit.
    #[test]
    fn mirrored_crate_versions() {
        let lock = include_str!("../../Cargo.lock");
        for (name, version) in [
            ("calamine", "0.36.1"),
            ("quick-xml", "0.41.0"),
            ("zip", "8.6.0"),
        ] {
            let package = format!("name = \"{name}\"\n");
            let entry = format!("{package}version = \"{version}\"\n");
            assert!(
                lock.contains(&entry) && lock.matches(&package).count() == 1,
                "{name} no longer resolves to exactly {version}: re-audit src/prescan against \
                 its source (and calamine's use of it), then update this test"
            );
        }
    }

    #[test]
    fn non_compound_files_are_left_to_calamine() {
        let mut reader = std::io::Cursor::new(b"PK\x03\x04 not an xls file".to_vec());
        assert!(xls(&mut reader, &Budget::new(Some(1))).is_ok());
    }
}
