//! Pre-scan of .xls files.
//!
//! calamine decodes every sheet of an .xls workbook while opening it, and
//! allocates each sheet's full rectangle (`Range::from_sparse`) as well as a
//! `Vec` sized by the sheet's DIMENSIONS record. Both are attacker controlled:
//! a 6 KB file can ask for hundreds of gigabytes, and a failed allocation
//! aborts the whole VM. Its compound file (CFB) reader also sizes vectors from
//! header fields and follows sector chains without loop detection.
//!
//! This module walks the same structures the way calamine 0.36 does (the CFB
//! traversal mirrors calamine's `cfb.rs`, MIT licensed) and rejects the file
//! before calamine opens it when:
//!
//! - a sheet's cells span more than `:max_cells` cells counted from A1;
//! - a sheet's DIMENSIONS record declares more than `:max_cells` cells;
//! - the compound file header or sector chains would make calamine allocate
//!   without bound or loop forever.
//!
//! When the file is not something calamine would open (not a compound file,
//! no workbook stream, truncated records, ...) the scan steps aside and lets
//! calamine report its own error.

use super::CellLimit;
use std::io::{Read, Seek, SeekFrom};

const OLE_SIGNATURE: &[u8] = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1";
const RESERVED_SECTORS: u32 = 0xFFFF_FFFA;
const DIFAT_SECTOR: u32 = 0xFFFF_FFFC;
const END_OF_CHAIN: u32 = 0xFFFF_FFFE;
const MINI_SECTOR_SIZE: usize = 64;
const MINI_STREAM_CUTOFF: usize = 4096;

/// calamine reserves a `Vec<u32>` with a capacity read from header bytes
/// 62..66 (a misread of the DIFAT length that mostly reflects the mini FAT
/// length). Legitimate files stay far below this; it bounds the reservation to
/// 1 GiB.
const MAX_DIFAT_CAPACITY: u64 = 1 << 28;

/// `Ok(None)` means "not applicable": calamine will reject the file itself.
type Scan<T> = Result<Option<T>, String>;

pub fn check<RS: Read + Seek>(reader: &mut RS, limit: CellLimit) -> Result<(), String> {
    if limit.is_unlimited() {
        return Ok(());
    }
    let result = scan(reader, limit);
    // calamine expects to start reading from the beginning.
    let _ = reader.seek(SeekFrom::Start(0));
    result.map(|_| ())
}

fn scan<RS: Read + Seek>(reader: &mut RS, limit: CellLimit) -> Scan<()> {
    let Some(file) = CompoundFile::open(reader)? else {
        return Ok(None);
    };
    let Some(stream) = file.workbook_stream(reader)? else {
        return Ok(None);
    };
    scan_workbook(&stream, limit)
}

fn invalid(reason: &str) -> String {
    format!("invalid .xls file: {reason}")
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

fn to_u32s(bytes: &[u8]) -> Result<impl Iterator<Item = u32> + '_, String> {
    if !bytes.len().is_multiple_of(4) {
        return Err(invalid("truncated sector"));
    }
    Ok(bytes
        .as_chunks::<4>()
        .0
        .iter()
        .map(|chunk| u32::from_le_bytes(*chunk)))
}

struct Directory {
    name: String,
    start: u32,
    length: usize,
}

struct CompoundFile {
    sector_size: usize,
    header_size: u64,
    /// calamine's `maxid`: the highest sector id it will read.
    maximum_sector: u64,
    fats: Vec<u32>,
    mini_fats: Vec<u32>,
    mini_stream: Vec<u8>,
    directories: Vec<Directory>,
}

impl CompoundFile {
    fn open<RS: Read + Seek>(reader: &mut RS) -> Scan<Self> {
        let Ok(file_length) = reader.seek(SeekFrom::End(0)) else {
            return Ok(None);
        };
        if reader.seek(SeekFrom::Start(0)).is_err() {
            return Ok(None);
        }
        let mut header = [0u8; 512];
        if reader.read_exact(&mut header).is_err() || &header[..8] != OLE_SIGNATURE {
            return Ok(None);
        }

        let version = read_u16(&header, 26);
        let (sector_size, header_size) = match read_u16(&header, 30) {
            0x0009 => (512, 512),
            0x000C => (4096, 4096),
            _ => return Ok(None),
        };
        if read_u16(&header, 32) != 0x0006 {
            return Ok(None);
        }

        let maximum_sector = file_length / sector_size as u64;
        let fat_length = u64::from(read_u32(&header, 44));
        let directory_start = read_u32(&header, 48);
        let mini_fat_start = read_u32(&header, 60);
        let mini_fat_length = read_u32(&header, 64) as usize;
        let difat_start = read_u32(&header, 68);
        let difat_capacity = u64::from(read_u32(&header, 62));

        if fat_length > maximum_sector + 1 {
            return Err(invalid(&format!(
                "the header declares {fat_length} FAT sectors but the file has {maximum_sector}"
            )));
        }
        if difat_capacity > MAX_DIFAT_CAPACITY {
            return Err(invalid("the header's mini FAT fields are out of range"));
        }

        let mut file = CompoundFile {
            sector_size,
            header_size,
            maximum_sector,
            fats: Vec::new(),
            mini_fats: Vec::new(),
            mini_stream: Vec::new(),
            directories: Vec::new(),
        };

        let mut difat: Vec<u32> = to_u32s(&header[76..512])?.collect();
        let mut sector_id = difat_start;
        let mut steps = 0;
        while sector_id < RESERVED_SECTORS {
            steps += 1;
            if steps > maximum_sector + 1 {
                return Err(invalid("the DIFAT sector chain loops"));
            }
            let Some(sector) = file.read_sector(reader, sector_id) else {
                return Ok(None);
            };
            difat.extend(to_u32s(&sector)?);
            sector_id = difat.pop().ok_or_else(|| invalid("empty DIFAT sector"))?;
        }

        let fat_sectors: Vec<u32> = difat.into_iter().filter(|id| *id < DIFAT_SECTOR).collect();
        if fat_sectors.len() as u64 > maximum_sector + 1 {
            return Err(invalid(
                "the DIFAT lists more FAT sectors than the file has",
            ));
        }
        for id in fat_sectors {
            let Some(sector) = file.read_sector(reader, id) else {
                return Ok(None);
            };
            file.fats.extend(to_u32s(&sector)?);
        }

        let Some(directory_bytes) = file.chain(reader, directory_start, usize::MAX)? else {
            return Ok(None);
        };
        for entry in directory_bytes.chunks(128) {
            if entry.len() < 128 {
                return Err(invalid("truncated directory entry"));
            }
            let name: String = char::decode_utf16(
                entry[..64]
                    .as_chunks::<2>()
                    .0
                    .iter()
                    .map(|pair| u16::from_le_bytes(*pair)),
            )
            .map(|character| character.unwrap_or(char::REPLACEMENT_CHARACTER))
            .take_while(|character| *character != '\0')
            .collect();
            let length = if sector_size == 512 {
                read_u32(entry, 120) as usize
            } else {
                u64::from_le_bytes(entry[120..128].try_into().expect("8 bytes")) as usize
            };
            file.directories.push(Directory {
                name,
                start: read_u32(entry, 116),
                length,
            });
        }
        let Some(root) = file.directories.first() else {
            return Ok(None);
        };
        if version != 3 && root.start == END_OF_CHAIN {
            return Ok(None);
        }

        if mini_fat_length > 0 {
            let (root_start, root_length) = (root.start, root.length);
            let Some(mini_stream) = file.chain(reader, root_start, root_length)? else {
                return Ok(None);
            };
            let Some(mini_fat) =
                file.chain(reader, mini_fat_start, mini_fat_length * sector_size)?
            else {
                return Ok(None);
            };
            file.mini_fats = to_u32s(&mini_fat)?.collect();
            file.mini_stream = mini_stream;
        }

        Ok(Some(file))
    }

    fn read_sector<RS: Read + Seek>(&self, reader: &mut RS, id: u32) -> Option<Vec<u8>> {
        if u64::from(id) > self.maximum_sector {
            return None;
        }
        let offset = self.header_size + u64::from(id) * self.sector_size as u64;
        reader.seek(SeekFrom::Start(offset)).ok()?;
        let mut sector = Vec::with_capacity(self.sector_size);
        reader
            .by_ref()
            .take(self.sector_size as u64)
            .read_to_end(&mut sector)
            .ok()?;
        Some(sector)
    }

    /// Follows a regular sector chain, as calamine's `Sectors::get_chain`.
    fn chain<RS: Read + Seek>(
        &self,
        reader: &mut RS,
        mut sector_id: u32,
        length: usize,
    ) -> Scan<Vec<u8>> {
        let mut chain = Vec::new();
        let mut steps = 0;
        while sector_id != END_OF_CHAIN {
            steps += 1;
            if steps > self.maximum_sector + 1 {
                return Err(invalid("a sector chain loops"));
            }
            let Some(sector) = self.read_sector(reader, sector_id) else {
                return Ok(None);
            };
            let take = sector.len().min(length - chain.len());
            chain.extend_from_slice(&sector[..take]);
            if chain.len() >= length {
                break;
            }
            sector_id = *self
                .fats
                .get(sector_id as usize)
                .ok_or_else(|| invalid("a sector chain points outside the FAT"))?;
        }
        Ok(Some(chain))
    }

    /// Follows a mini stream chain.
    fn mini_chain(&self, mut sector_id: u32, length: usize) -> Scan<Vec<u8>> {
        let mini_sector_count = self.mini_stream.len() / MINI_SECTOR_SIZE;
        let mut chain = Vec::new();
        let mut steps = 0;
        while sector_id != END_OF_CHAIN {
            steps += 1;
            if steps > mini_sector_count + 1 {
                return Err(invalid("a mini sector chain loops"));
            }
            if u64::from(sector_id) > self.maximum_sector * 8 {
                return Ok(None);
            }
            let start = sector_id as usize * MINI_SECTOR_SIZE;
            let sector = self
                .mini_stream
                .get(start..start + MINI_SECTOR_SIZE)
                .ok_or_else(|| invalid("a mini sector lies outside the mini stream"))?;
            let take = MINI_SECTOR_SIZE.min(length - chain.len());
            chain.extend_from_slice(&sector[..take]);
            if chain.len() >= length {
                break;
            }
            sector_id = *self
                .mini_fats
                .get(sector_id as usize)
                .ok_or_else(|| invalid("a mini sector chain points outside the mini FAT"))?;
        }
        Ok(Some(chain))
    }

    /// The workbook stream, under any of the names calamine accepts.
    fn workbook_stream<RS: Read + Seek>(&self, reader: &mut RS) -> Scan<Vec<u8>> {
        for name in ["Workbook", "Book", "WORKBOOK", "BOOK"] {
            let Some(directory) = self.directories.iter().find(|entry| entry.name == name) else {
                continue;
            };
            let stream = if directory.length < MINI_STREAM_CUTOFF {
                self.mini_chain(directory.start, directory.length)?
            } else {
                self.chain(reader, directory.start, directory.length)?
            };
            if stream.is_some() {
                return Ok(stream);
            }
        }
        Ok(None)
    }
}

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

fn scan_workbook(stream: &[u8], limit: CellLimit) -> Scan<()> {
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
        match scan_sheet(sheet_stream, &name, limit)? {
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

fn scan_sheet(stream: &[u8], name: &str, limit: CellLimit) -> Scan<()> {
    let mut dimension_error = None;
    let extent = match sheet_extent(stream, name, limit, &mut dimension_error) {
        Some(extent) => extent,
        // calamine fails on this sheet, but only after reserving what the
        // DIMENSIONS record asked for.
        None => return dimension_error.map_or(Ok(None), Err),
    };
    if let Some((last_row, last_column)) = extent.last {
        limit.check_rectangle(name, u64::from(last_row) + 1, u64::from(last_column) + 1)?;
    }
    // Report the cells first: they are what the caller can act on.
    dimension_error.map_or(Ok(Some(())), Err)
}

/// The largest cell position in a sheet substream, or `None` when calamine
/// would fail on a record.
fn sheet_extent(
    stream: &[u8],
    name: &str,
    limit: CellLimit,
    dimension_error: &mut Option<String>,
) -> Option<Extent> {
    let mut extent = Extent::default();
    let position = |data: &[u8]| (u32::from(read_u16(data, 0)), u32::from(read_u16(data, 2)));

    for record in (Records { stream }) {
        let (record_type, data) = record.ok()?;
        let minimum_length = match record_type {
            0x0200 => {
                if let Err(message) = check_dimensions(data, name, limit)? {
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

/// calamine reserves `rows × columns` cells from the DIMENSIONS record
/// before reading any cell; reproduce its arithmetic (including wrapping).
/// `None` when calamine rejects the record.
fn check_dimensions(data: &[u8], name: &str, limit: CellLimit) -> Option<Result<(), String>> {
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
    let cells = (rows as usize).saturating_mul(columns as usize) as u128;
    Some(limit.check_count(
        name,
        cells,
        &format!("declares {rows} rows × {columns} columns in its DIMENSIONS record:"),
    ))
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
        assert_eq!(scan_workbook(&stream, CellLimit(Some(40))), Ok(Some(())));
    }

    #[test]
    fn sparse_cells_are_rejected() {
        let stream = workbook(&[cell(0x0203, 0, 0, 14), cell(0x027E, 0xFFFF, 0xFFFF, 10)]);
        let error = scan_workbook(&stream, CellLimit(Some(10_000_000))).unwrap_err();
        assert_eq!(
            error,
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
        assert!(scan_workbook(&stream, CellLimit(Some(5))).is_ok());
        assert!(scan_workbook(&stream, CellLimit(Some(4))).is_err());
    }

    #[test]
    fn declared_dimensions_are_checked() {
        let mut data = 0u32.to_le_bytes().to_vec();
        data.extend_from_slice(&u32::MAX.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&255u16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        let stream = workbook(&[record(0x0200, &data)]);
        let error = scan_workbook(&stream, CellLimit(Some(10_000_000))).unwrap_err();
        assert!(error.starts_with(
            "sheet \"S\" declares 4294967295 rows × 255 columns in its DIMENSIONS record:"
        ));
    }

    #[test]
    fn oversized_dimensions_are_reported_even_when_a_later_record_is_malformed() {
        let mut data = 0u32.to_le_bytes().to_vec();
        data.extend_from_slice(&u32::MAX.to_le_bytes());
        data.extend_from_slice(&[0, 0, 255, 0, 0, 0]);
        // A Number record too short for calamine.
        let stream = workbook(&[record(0x0200, &data), record(0x0203, &[0; 4])]);
        assert!(scan_workbook(&stream, CellLimit(Some(100))).is_err());
    }

    #[test]
    fn truncated_records_are_left_to_calamine() {
        let mut stream = workbook(&[cell(0x0203, 0, 0, 14)]);
        stream.truncate(stream.len() - 3);
        assert_eq!(scan_workbook(&stream, CellLimit(Some(1))), Ok(None));
    }

    #[test]
    fn non_compound_files_are_left_to_calamine() {
        let mut reader = std::io::Cursor::new(b"PK\x03\x04 not an xls file".to_vec());
        assert_eq!(check(&mut reader, CellLimit(Some(1))), Ok(()));
    }
}
