//! The compound file (CFB) container of .xls files.
//!
//! calamine's CFB reader sizes vectors from header fields and follows sector
//! chains without loop detection. This walk mirrors calamine 0.36's `cfb.rs`
//! (MIT licensed) to find the workbook stream, and rejects headers and chains
//! that would make calamine allocate without bound or loop forever.

use super::{invalid_xls as invalid, read_u16, read_u32, Scan};
use crate::error::Error;
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

/// The workbook stream of a compound file, or `None` when calamine would
/// fail to open the container by itself.
pub fn workbook_stream<RS: Read + Seek + ?Sized>(reader: &mut RS) -> Scan<Vec<u8>> {
    let Some(file) = CompoundFile::open(reader)? else {
        return Ok(None);
    };
    file.workbook_stream(reader)
}

fn to_u32s(bytes: &[u8]) -> Result<impl Iterator<Item = u32> + '_, Error> {
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
    fn open<RS: Read + Seek + ?Sized>(reader: &mut RS) -> Scan<Self> {
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

    fn read_sector<RS: Read + Seek + ?Sized>(&self, reader: &mut RS, id: u32) -> Option<Vec<u8>> {
        if u64::from(id) > self.maximum_sector {
            return None;
        }
        let offset = self.header_size + u64::from(id) * self.sector_size as u64;
        reader.seek(SeekFrom::Start(offset)).ok()?;
        let mut sector = Vec::with_capacity(self.sector_size);
        (&mut *reader)
            .take(self.sector_size as u64)
            .read_to_end(&mut sector)
            .ok()?;
        Some(sector)
    }

    /// Follows a regular sector chain, as calamine's `Sectors::get_chain`.
    fn chain<RS: Read + Seek + ?Sized>(
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
    fn workbook_stream<RS: Read + Seek + ?Sized>(&self, reader: &mut RS) -> Scan<Vec<u8>> {
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
