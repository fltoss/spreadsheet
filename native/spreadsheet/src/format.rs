//! Everything that differs between the spreadsheet formats: how each one is
//! recognised, which pre-scan guards it, how calamine opens it, which sheets
//! hold cells and where those cells come from.

use crate::budget::Budget;
use crate::error::Error;
use crate::prescan;
use crate::workbook::{ReadSeek, Workbook};
use calamine::{
    open_workbook_from_rs, Cell, Data, DataRef, Range, Reader, SheetType, Sheets, XlsxError,
};
use std::path::Path;

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum Format {
    Xls,
    Xlsx,
    Xlsb,
    Ods,
}

/// Local file header signature that zip-based formats start with.
pub const ZIP_MAGIC: &[u8] = b"PK\x03\x04";

const ODS_MIMETYPE: &[u8] = b"application/vnd.oasis.opendocument.spreadsheet";

/// Why a format could not open a file.
pub enum OpenError {
    /// The pre-scan rejected the file; do not try other formats.
    Rejected(Error),
    /// calamine could not open the file as this format.
    NotThisFormat(calamine::Error),
}

/// A sheet's cells, as its format's reader provides them.
pub enum SheetCells<'a> {
    /// The sheet holds no cells (chart sheets, VBA modules, empty sheets).
    Empty,
    /// The non-empty cells of an .xlsx or .xlsb sheet, in file order, which
    /// is normally but not necessarily row-major.
    Sparse(Vec<Cell<DataRef<'a>>>),
    /// The range calamine decoded when it opened an .xls or .ods workbook.
    Range(Range<Data>),
}

impl Format {
    /// The order calamine's `open_workbook_auto` tries formats in when it
    /// cannot rely on the file extension.
    pub const DETECTION_ORDER: [Format; 4] = [Format::Xls, Format::Xlsx, Format::Xlsb, Format::Ods];

    /// The format a file extension names, ignoring case.
    pub fn from_extension(path: &Path) -> Option<Format> {
        let extension = path.extension()?.to_str()?.to_ascii_lowercase();
        match extension.as_str() {
            "xls" | "xla" => Some(Format::Xls),
            "xlsx" | "xlsm" | "xlam" => Some(Format::Xlsx),
            "xlsb" => Some(Format::Xlsb),
            "ods" => Some(Format::Ods),
            _ => None,
        }
    }

    /// For content that starts like a zip archive: .ods when its first entry
    /// is an OpenDocument spreadsheet `mimetype`, .xlsx otherwise.
    pub fn zip_based(content: &[u8]) -> Option<Format> {
        if !content.starts_with(ZIP_MAGIC) {
            return None;
        }
        // The `mimetype` entry is stored first and uncompressed: its name
        // starts at offset 30 and its content follows a few bytes later.
        let start = &content[..content.len().min(128)];
        let ods = start.get(30..38) == Some(b"mimetype".as_slice())
            && start
                .windows(ODS_MIMETYPE.len())
                .any(|window| window == ODS_MIMETYPE);
        Some(if ods { Format::Ods } else { Format::Xlsx })
    }

    fn of(workbook: &Workbook) -> Format {
        match workbook {
            Sheets::Xls(_) => Format::Xls,
            Sheets::Xlsx(_) => Format::Xlsx,
            Sheets::Xlsb(_) => Format::Xlsb,
            Sheets::Ods(_) => Format::Ods,
        }
    }

    /// Opens `reader` as this format, after the pre-scan for the formats
    /// calamine decodes in full when opening them.
    pub fn open<'a>(
        self,
        mut reader: Box<dyn ReadSeek + 'a>,
        budget: &Budget,
    ) -> Result<Workbook<'a>, OpenError> {
        match self {
            Format::Xls => prescan::xls(&mut *reader, budget).map_err(OpenError::Rejected)?,
            Format::Ods => prescan::ods(&mut *reader, budget).map_err(OpenError::Rejected)?,
            Format::Xlsx | Format::Xlsb => (),
        }
        let workbook = match self {
            Format::Xls => open_workbook_from_rs(reader)
                .map(Sheets::Xls)
                .map_err(calamine::Error::Xls),
            Format::Xlsx => open_workbook_from_rs(reader)
                .map(Sheets::Xlsx)
                .map_err(calamine::Error::Xlsx),
            Format::Xlsb => open_workbook_from_rs(reader)
                .map(Sheets::Xlsb)
                .map_err(calamine::Error::Xlsb),
            Format::Ods => open_workbook_from_rs(reader)
                .map(Sheets::Ods)
                .map_err(calamine::Error::Ods),
        };
        workbook.map_err(OpenError::NotThisFormat)
    }

    /// Whether sheets of `sheet_type` are read. VBA modules hold no cells.
    /// .xlsb chart sheets hold none either, and calamine's cell reader fails
    /// on them. .xls chart sheets are read, because they can carry the
    /// chart's data; .xlsx ones are reported as empty by calamine.
    fn reads(self, sheet_type: SheetType) -> bool {
        match sheet_type {
            SheetType::Vba => false,
            SheetType::ChartSheet => self != Format::Xlsb,
            _ => true,
        }
    }
}

/// Reads the cells of sheet `name`.
pub fn read_cells<'a>(workbook: &'a mut Workbook, name: &str) -> Result<SheetCells<'a>, Error> {
    let sheet_type = workbook
        .sheets_metadata()
        .iter()
        .find(|sheet| sheet.name == name)
        .map(|sheet| sheet.typ)
        .ok_or_else(|| Error::SheetNotFound(name.to_string()))?;
    if !Format::of(workbook).reads(sheet_type) {
        return Ok(SheetCells::Empty);
    }

    match workbook {
        Sheets::Xlsx(xlsx) => match xlsx.worksheet_cells_reader(name) {
            Ok(mut reader) => collect_cells(|| reader.next_cell(), calamine::Error::Xlsx),
            // calamine treats other sheet kinds as empty too.
            Err(XlsxError::NotAWorksheet(_)) => Ok(SheetCells::Empty),
            Err(error) => Err(calamine::Error::Xlsx(error).into()),
        },
        Sheets::Xlsb(xlsb) => {
            let mut reader = xlsb
                .worksheet_cells_reader(name)
                .map_err(calamine::Error::Xlsb)?;
            collect_cells(|| reader.next_cell(), calamine::Error::Xlsb)
        }
        Sheets::Xls(xls) => Ok(SheetCells::Range(
            xls.worksheet_range(name).map_err(calamine::Error::Xls)?,
        )),
        Sheets::Ods(ods) => Ok(SheetCells::Range(
            ods.worksheet_range(name).map_err(calamine::Error::Ods)?,
        )),
    }
}

/// Collects a sheet's non-empty cells from a calamine cell reader.
fn collect_cells<'a, E>(
    mut next_cell: impl FnMut() -> Result<Option<Cell<DataRef<'a>>>, E>,
    to_error: impl Fn(E) -> calamine::Error,
) -> Result<SheetCells<'a>, Error> {
    let mut cells = Vec::new();
    while let Some(cell) = next_cell().map_err(&to_error)? {
        if *cell.get_value() != DataRef::Empty {
            cells.push(cell);
        }
    }
    Ok(SheetCells::Sparse(cells))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn damaged_zips_report_the_zip_based_format() {
        let mut ods = b"PK\x03\x04".to_vec();
        ods.resize(30, 0);
        ods.extend_from_slice(b"mimetypeapplication/vnd.oasis.opendocument.spreadsheetPK");
        assert_eq!(Format::zip_based(&ods), Some(Format::Ods));
        assert_eq!(
            Format::zip_based(b"PK\x03\x04truncated"),
            Some(Format::Xlsx)
        );
        assert_eq!(Format::zip_based(b"plain text"), None);
    }

    #[test]
    fn extensions_ignore_case() {
        assert_eq!(
            Format::from_extension(Path::new("a/REPORT.XLSM")),
            Some(Format::Xlsx)
        );
        assert_eq!(Format::from_extension(Path::new("a.csv")), None);
    }
}
