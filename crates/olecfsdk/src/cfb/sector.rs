use std::cell::Cell;

use crate::{Error, Result};

use super::header::Header;

pub const MAX_REGULAR_SECTOR: u32 = 0xffff_fffa;

#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct SectorId(u32);

impl SectorId {
    pub fn new(value: u32) -> Result<Self> {
        if value > MAX_REGULAR_SECTOR {
            return Err(Error::invalid(
                0,
                format!("invalid regular sector ID 0x{value:08x}"),
            ));
        }
        Ok(Self(value))
    }

    pub fn get(self) -> u32 {
        self.0
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct MiniSectorId(u32);

impl MiniSectorId {
    pub fn new(value: u32) -> Result<Self> {
        if value > MAX_REGULAR_SECTOR {
            return Err(Error::invalid(
                0,
                format!("invalid mini-sector ID 0x{value:08x}"),
            ));
        }
        Ok(Self(value))
    }

    pub fn get(self) -> u32 {
        self.0
    }
}

pub(crate) struct SectorSource<'a> {
    bytes: &'a [u8],
    sector_len: usize,
    sector_count: usize,
    partial_sector: Option<SectorId>,
    partial_len: usize,
    partial_accessed: Cell<bool>,
}

impl<'a> SectorSource<'a> {
    pub(crate) fn new(bytes: &'a [u8], original_len: usize, header: &Header) -> Result<Self> {
        let sector_len = header.sector_len();
        if bytes.len() < sector_len {
            return Err(Error::invalid(
                0,
                "CFB file is shorter than its header sector",
            ));
        }
        if !bytes.len().is_multiple_of(sector_len) || original_len > bytes.len() {
            return Err(Error::invalid(0, "internal CFB sector image is not padded"));
        }
        let complete_sector_count = bytes.len() / sector_len;
        let partial_len = original_len % sector_len;
        let partial_sector = if partial_len == 0 {
            None
        } else {
            let id = complete_sector_count
                .checked_sub(2)
                .ok_or_else(|| Error::invalid(0, "partial CFB header sector"))?;
            Some(SectorId::new(id as u32)?)
        };
        Ok(Self {
            bytes,
            sector_len,
            sector_count: complete_sector_count - 1,
            partial_sector,
            partial_len,
            partial_accessed: Cell::new(false),
        })
    }

    pub(crate) fn sector_count(&self) -> usize {
        self.sector_count
    }

    pub(crate) fn sector_len(&self) -> usize {
        self.sector_len
    }

    pub(crate) fn sector(&self, id: SectorId) -> Result<&'a [u8]> {
        let index = usize::try_from(id.get())
            .map_err(|_| Error::invalid(0, "sector ID does not fit usize"))?;
        if index >= self.sector_count {
            return Err(Error::invalid(
                0,
                format!("sector {index} is beyond EOF ({})", self.sector_count),
            ));
        }
        let start = index
            .checked_add(1)
            .and_then(|value| value.checked_mul(self.sector_len))
            .ok_or_else(|| Error::invalid(0, "sector offset overflow"))?;
        let end = start
            .checked_add(self.sector_len)
            .ok_or_else(|| Error::invalid(0, "sector end overflow"))?;
        if self.partial_sector == Some(id) {
            self.partial_accessed.set(true);
        }
        Ok(&self.bytes[start..end])
    }

    pub(crate) fn full_sector(&self, id: SectorId) -> Result<&'a [u8]> {
        if self.valid_len(id) != self.sector_len {
            return Err(Error::invalid(
                0,
                "truncated CFB allocation or directory sector",
            ));
        }
        self.sector(id)
    }

    pub(crate) fn valid_len(&self, id: SectorId) -> usize {
        if self.partial_sector == Some(id) {
            self.partial_len
        } else {
            self.sector_len
        }
    }

    pub(crate) fn is_partial(&self, id: SectorId) -> bool {
        self.partial_sector == Some(id)
    }

    pub(crate) fn has_partial_sector(&self) -> bool {
        self.partial_sector.is_some()
    }

    pub(crate) fn unaccessed_partial_data(&self) -> &'a [u8] {
        if self.partial_accessed.get() || self.partial_len == 0 {
            return &[];
        }
        let start = self.bytes.len() - self.sector_len;
        &self.bytes[start..start + self.partial_len]
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cfb::header::{
        BYTE_ORDER_LE, FREE_SECTOR, MAGIC, MINI_SECTOR_SHIFT, MINI_STREAM_CUTOFF,
    };

    fn header() -> Header {
        Header {
            signature: MAGIC,
            clsid: [0; 16],
            minor_version: 0x003e,
            major_version: 3,
            byte_order: BYTE_ORDER_LE,
            sector_shift: 9,
            mini_sector_shift: MINI_SECTOR_SHIFT,
            reserved: [0; 6],
            number_of_directory_sectors: 0,
            number_of_fat_sectors: 0,
            first_directory_sector: FREE_SECTOR,
            transaction_signature: 0,
            mini_stream_cutoff: MINI_STREAM_CUTOFF,
            first_mini_fat_sector: FREE_SECTOR,
            number_of_mini_fat_sectors: 0,
            first_difat_sector: FREE_SECTOR,
            number_of_difat_sectors: 0,
            difat: [FREE_SECTOR; 109],
        }
    }

    #[test]
    fn checked_sector_addressing_excludes_header_sector() {
        let mut bytes = vec![0; 3 * 512];
        bytes[512] = 7;
        bytes[1024] = 11;
        let source = SectorSource::new(&bytes, bytes.len(), &header()).unwrap();
        assert_eq!(source.sector_count(), 2);
        assert_eq!(source.sector(SectorId::new(0).unwrap()).unwrap()[0], 7);
        assert_eq!(source.sector(SectorId::new(1).unwrap()).unwrap()[0], 11);
        assert!(source.sector(SectorId::new(2).unwrap()).is_err());
    }

    #[test]
    fn trailing_bytes_are_outside_the_sector_address_space() {
        let mut bytes = vec![0; 3 * 512];
        let original_len = 2 * 512 + 3;
        bytes[2 * 512..original_len].copy_from_slice(&[3, 5, 7]);
        let source = SectorSource::new(&bytes, original_len, &header()).unwrap();
        assert_eq!(source.sector_count(), 2);
        assert_eq!(source.unaccessed_partial_data(), [3, 5, 7]);
    }
}
