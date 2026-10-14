//! BIFF records of an .xls workbook stream.
//!
//! calamine decodes every sheet of an .xls workbook while opening it, and
//! allocates each sheet's full rectangle (`Range::from_sparse`) as well as a
//! `Vec` sized by the sheet's DIMENSIONS record. Both are attacker controlled:
//! a 6 KB file can ask for hundreds of gigabytes, and a failed allocation
//! aborts the whole VM. This scan reads the records the way calamine 0.36
//! does and reports each sheet's cell extent and declared dimensions.

use super::{invalid_xls as invalid, read_u16, read_u32, Scan};
use crate::budget::Budget;
use crate::error::Error;

/// BIFF records, merged with their CONTINUE records as calamine does.
struct Records<'a> {
    stream: &'a [u8],
}

impl<'a> Iterator for Records<'a> {
    /// `Err(())` is a malformed record, which makes calamine fail.
    type Item = Result<(u16, &'a [u8]), ()>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.stream.len() < 4 {
            return if self.stream.is_empty() {
                None
            } else {
                Some(Err(()))
            };
        }
        let record_type = read_u16(self.stream, 0);
        let length = read_u16(self.stream, 2) as usize;
        if self.stream.len() < length + 4 {
            return Some(Err(()));
        }
        let (record, rest) = self.stream.split_at(length + 4);
        self.stream = rest;
        while self.stream.len() > 4 && read_u16(self.stream, 0) == 0x003C {
            let length = read_u16(self.stream, 2) as usize;
            if self.stream.len() < length + 4 {
                return Some(Err(()));
            }
            self.stream = &self.stream[length + 4..];
        }
        Some(Ok((record_type, &record[4..])))
    }
}

pub fn scan_workbook(stream: &[u8], budget: &Budget) -> Scan<()> {
    // calamine decodes every sheet when it opens the file, so all of them
    // count against the budget, even when only one sheet is parsed.
    let mut budget = *budget;
    let mut sheets = Vec::new();
    let mut biff8 = true;

    for record in (Records { stream }) {
        let Ok((record_type, data)) = record else {
            return Ok(None);
        };
        match record_type {
            // FilePass: encrypted, calamine refuses it.
            0x002F => {
                if data.len() < 2 {
                    return Err(invalid("truncated FilePass record"));
                }
                if read_u16(data, 0) != 0 {
                    return Ok(None);
                }
            }
            // BOF
            0x0809 => {
                if data.len() < 2 {
                    return Err(invalid("truncated BOF record"));
                }
                let version = read_u16(data, 0);
                let document_type = if data.len() >= 4 {
                    read_u16(data, 2)
                } else {
                    0
                };
                biff8 = match version {
                    0x0200 | 0x0002 | 0x0007 | 0x0300 | 0x0400 | 0x0500 => false,
                    0 => document_type != 0x1000,
                    _ => true,
                };
            }
            // BoundSheet8
            0x0085 => {
                if data.len() < 6 {
                    return Err(invalid("truncated BoundSheet record"));
                }
                if data[4] & 0b0011_1111 > 2 || !matches!(data[5], 0x00 | 0x01 | 0x02 | 0x06) {
                    return Ok(None);
                }
                if data.len() < if biff8 { 8 } else { 7 } {
                    return Ok(None);
                }
                sheets.push((read_u32(data, 0) as usize, sheet_name(&data[6..], biff8)));
            }
            // EOF
            0x000A => break,
            _ => (),
        }
    }

    for (position, name) in sheets {
        let Some(sheet_stream) = stream.get(position..) else {
            return Err(invalid(&format!(
                "sheet {name:?} starts past the end of the workbook"
            )));
        };
        match scan_sheet(sheet_stream, &name, &mut budget)? {
            Some(()) => (),
            None => return Ok(None),
        }
    }
    Ok(Some(()))
}

/// Tracks the largest cell position of a sheet.
#[derive(Default)]
struct Extent {
    last: Option<(u32, u32)>,
}

impl Extent {
    fn add(&mut self, row: u32, column: u32) {
        let (last_row, last_column) = self.last.unwrap_or((0, 0));
        self.last = Some((last_row.max(row), last_column.max(column)));
    }
}

fn scan_sheet(stream: &[u8], name: &str, budget: &mut Budget) -> Scan<()> {
    let mut dimension_error = None;
    let extent = match sheet_extent(stream, name, &mut dimension_error) {
        Some(extent) => extent,
        // calamine fails on this sheet, but only after reserving what the
        // DIMENSIONS record asked for.
        None => return dimension_error.map_or(Ok(None), Err),
    };
    if let Some((last_row, last_column)) = extent.last {
        budget.spend(name, u64::from(last_row) + 1, u64::from(last_column) + 1)?;
    }
    // Report the cells first: they are what the caller can act on.
    dimension_error.map_or(Ok(Some(())), Err)
}

/// The largest cell position in a sheet substream, or `None` when calamine
/// would fail on a record.
fn sheet_extent(stream: &[u8], name: &str, dimension_error: &mut Option<Error>) -> Option<Extent> {
    let mut extent = Extent::default();
    let position = |data: &[u8]| (u32::from(read_u16(data, 0)), u32::from(read_u16(data, 2)));

    for record in (Records { stream }) {
        let (record_type, data) = record.ok()?;
        let minimum_length = match record_type {
            0x0200 => {
                if let Err(message) = check_dimensions(data, name)? {
                    dimension_error.get_or_insert(message);
                }
                continue;
            }
            // Number
            0x0203 => 14,
            // Label, RString
            0x0204 | 0x00D6 => 6,
            // BoolErr
            0x0205 => 8,
            // Rk, LabelSst
            0x027E | 0x00FD => 10,
            // Formula
            0x0006 => 20,
            // MulRk
            0x00BD => {
                if data.len() < 6 {
                    return None;
                }
                let (row, first_column) = position(data);
                let last_column = read_u16(data, data.len() - 2);
                let count = last_column
                    .wrapping_sub(first_column as u16)
                    .wrapping_add(1) as usize;
                if data.len() != 6 + 6 * count {
                    return None;
                }
                if count > 0 {
                    extent.add(row, first_column + count as u32 - 1);
                }
                continue;
            }
            // EOF
            0x000A => break,
            _ => continue,
        };
        if data.len() < minimum_length {
            return None;
        }
        let (row, column) = position(data);
        extent.add(row, column);
    }
    Some(extent)
}

/// The largest sheet an .xls (BIFF8) file can describe.
const MAX_ROWS: u32 = 65_536;
const MAX_COLUMNS: u32 = 256;

/// calamine reserves `rows × columns` cells from the DIMENSIONS record
/// before reading any cell. Reproduce its arithmetic (including wrapping) and
/// reject sizes no .xls file can have. This is a structural check, independent
/// of `:max_cells`: a formatting-only template may legitimately declare
/// `A1:IV65536` and hold a handful of values. `None` when calamine rejects the
/// record.
fn check_dimensions(data: &[u8], name: &str) -> Option<Result<(), Error>> {
    let (first_row, last_row, mut first_column, last_column) = match data.len() {
        10 => (
            u32::from(read_u16(data, 0)),
            u32::from(read_u16(data, 2)),
            u32::from(read_u16(data, 4)),
            u32::from(read_u16(data, 6)),
        ),
        14 => (
            read_u32(data, 0),
            read_u32(data, 4),
            u32::from(read_u16(data, 8)),
            u32::from(read_u16(data, 10)),
        ),
        _ => return None,
    };
    if 0xFF < first_column || last_column < first_column {
        first_column = 0;
    }
    let (start, end) = if 1 <= last_row && 1 <= last_column {
        ((first_row, first_column), (last_row - 1, last_column - 1))
    } else {
        ((first_row, first_column), (first_row, first_column))
    };
    let rows = end.0.wrapping_sub(start.0).wrapping_add(1);
    let columns = end.1.wrapping_sub(start.1).wrapping_add(1);
    if rows > MAX_ROWS || columns > MAX_COLUMNS {
        return Some(Err(invalid(&format!(
            "sheet {name:?} declares {rows} rows × {columns} columns in its DIMENSIONS \
             record, above the .xls maximum of {MAX_ROWS} rows × {MAX_COLUMNS} columns"
        ))));
    }
    Some(Ok(()))
}

/// Best-effort sheet name for error messages (`ShortXLUnicodeString`).
fn sheet_name(data: &[u8], biff8: bool) -> String {
    let count = data[0] as usize;
    let name: String = if biff8 {
        let high_byte = data[1] & 1 != 0;
        let characters = &data[2..];
        if high_byte {
            char::decode_utf16(
                characters
                    .as_chunks::<2>()
                    .0
                    .iter()
                    .take(count)
                    .map(|pair| u16::from_le_bytes(*pair)),
            )
            .map(|character| character.unwrap_or(char::REPLACEMENT_CHARACTER))
            .collect()
        } else {
            characters
                .iter()
                .take(count)
                .map(|byte| *byte as char)
                .collect()
        }
    } else {
        data[1..]
            .iter()
            .take(count)
            .map(|byte| *byte as char)
            .collect()
    };
    name.replace('\0', "")
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(record_type: u16, data: &[u8]) -> Vec<u8> {
        let mut bytes = record_type.to_le_bytes().to_vec();
        bytes.extend_from_slice(&(data.len() as u16).to_le_bytes());
        bytes.extend_from_slice(data);
        bytes
    }

    fn cell(record_type: u16, row: u16, column: u16, length: usize) -> Vec<u8> {
        let mut data = vec![0; length];
        data[..2].copy_from_slice(&row.to_le_bytes());
        data[2..4].copy_from_slice(&column.to_le_bytes());
        record(record_type, &data)
    }

    /// A BIFF8 workbook stream with one sheet named "S" made of `sheet_records`.
    fn workbook(sheet_records: &[Vec<u8>]) -> Vec<u8> {
        let bof = record(0x0809, &[0x00, 0x06, 0x05, 0x00]);
        // BoundSheet record (4 + 9 bytes) and EOF record (4 bytes) follow the BOF.
        let sheet_position = bof.len() + 13 + 4;
        let mut bound_sheet = (sheet_position as u32).to_le_bytes().to_vec();
        bound_sheet.extend_from_slice(&[0, 0, 1, 0, b'S']);
        let mut stream = bof;
        stream.extend(record(0x0085, &bound_sheet));
        stream.extend(record(0x000A, &[]));
        assert_eq!(stream.len(), sheet_position);
        stream.extend(record(0x0809, &[0x00, 0x06, 0x10, 0x00]));
        for sheet_record in sheet_records {
            stream.extend_from_slice(sheet_record);
        }
        stream.extend(record(0x000A, &[]));
        stream
    }

    #[test]
    fn small_sheets_pass() {
        let stream = workbook(&[cell(0x0203, 1, 1, 14), cell(0x027E, 9, 3, 10)]);
        assert!(matches!(
            scan_workbook(&stream, &Budget::new(Some(40))),
            Ok(Some(()))
        ));
    }

    #[test]
    fn sparse_cells_are_rejected() {
        let stream = workbook(&[cell(0x0203, 0, 0, 14), cell(0x027E, 0xFFFF, 0xFFFF, 10)]);
        let error = scan_workbook(&stream, &Budget::new(Some(10_000_000))).unwrap_err();
        assert_eq!(
            error.to_string(),
            "sheet \"S\" has 65536 rows × 65536 columns (4294967296 cells), \
             above the :max_cells limit of 10000000"
        );
    }

    #[test]
    fn mul_rk_records_count_every_column() {
        // Row 0, columns 2..=4.
        let mut data = vec![0, 0, 2, 0];
        data.extend(vec![0; 18]);
        data.extend_from_slice(&4u16.to_le_bytes());
        let stream = workbook(&[record(0x00BD, &data)]);
        assert!(scan_workbook(&stream, &Budget::new(Some(5))).is_ok());
        assert!(scan_workbook(&stream, &Budget::new(Some(4))).is_err());
    }

    /// A BIFF8 DIMENSIONS record: rows `first..last` (exclusive) and columns
    /// `first..last` (exclusive).
    fn dimensions(first_row: u32, last_row: u32, first_column: u16, last_column: u16) -> Vec<u8> {
        let mut data = first_row.to_le_bytes().to_vec();
        data.extend_from_slice(&last_row.to_le_bytes());
        data.extend_from_slice(&first_column.to_le_bytes());
        data.extend_from_slice(&last_column.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        record(0x0200, &data)
    }

    #[test]
    fn forged_dimensions_are_rejected_without_a_limit() {
        let stream = workbook(&[dimensions(0, u32::MAX, 0, 255)]);
        let error = scan_workbook(&stream, &Budget::new(None)).unwrap_err();
        assert_eq!(
            error.to_string(),
            "invalid .xls file: sheet \"S\" declares 4294967295 rows × 255 columns in its \
             DIMENSIONS record, above the .xls maximum of 65536 rows × 256 columns"
        );
        // last row before first row wraps around in calamine.
        let stream = workbook(&[dimensions(10, 5, 0, 1)]);
        assert!(scan_workbook(&stream, &Budget::new(None)).is_err());
    }

    #[test]
    fn declared_dimensions_do_not_count_against_the_limit() {
        // A formatting-only template: A1:IV65536 declared, two values.
        let stream = workbook(&[
            dimensions(0, 65_536, 0, 256),
            cell(0x0203, 0, 0, 14),
            cell(0x027E, 1, 1, 10),
        ]);
        assert!(matches!(
            scan_workbook(&stream, &Budget::new(Some(4))),
            Ok(Some(()))
        ));
    }

    #[test]
    fn oversized_dimensions_are_reported_even_when_a_later_record_is_malformed() {
        let mut data = 0u32.to_le_bytes().to_vec();
        data.extend_from_slice(&u32::MAX.to_le_bytes());
        data.extend_from_slice(&[0, 0, 255, 0, 0, 0]);
        // A Number record too short for calamine.
        let stream = workbook(&[record(0x0200, &data), record(0x0203, &[0; 4])]);
        assert!(scan_workbook(&stream, &Budget::new(Some(100))).is_err());
    }

    #[test]
    fn truncated_records_are_left_to_calamine() {
        let mut stream = workbook(&[cell(0x0203, 0, 0, 14)]);
        stream.truncate(stream.len() - 3);
        assert!(matches!(
            scan_workbook(&stream, &Budget::new(Some(1))),
            Ok(None)
        ));
    }
}
