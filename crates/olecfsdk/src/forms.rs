//! MS-OFORMS control persistence structures.

use bitflags::bitflags;

use crate::{Error, Result};

bitflags! {
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub struct MorphDataPropertyMask: u64 {
        const VARIOUS_PROPERTY_BITS = 1 << 0;
        const BACK_COLOR = 1 << 1;
        const FORE_COLOR = 1 << 2;
        const MAX_LENGTH = 1 << 3;
        const BORDER_STYLE = 1 << 4;
        const SCROLL_BARS = 1 << 5;
        const DISPLAY_STYLE = 1 << 6;
        const MOUSE_POINTER = 1 << 7;
        const SIZE = 1 << 8;
        const PASSWORD_CHAR = 1 << 9;
        const LIST_WIDTH = 1 << 10;
        const BOUND_COLUMN = 1 << 11;
        const TEXT_COLUMN = 1 << 12;
        const COLUMN_COUNT = 1 << 13;
        const LIST_ROWS = 1 << 14;
        const COLUMN_INFO_COUNT = 1 << 15;
        const MATCH_ENTRY = 1 << 16;
        const LIST_STYLE = 1 << 17;
        const SHOW_DROP_BUTTON_WHEN = 1 << 18;
        const UNUSED1 = 1 << 19;
        const DROP_BUTTON_STYLE = 1 << 20;
        const MULTI_SELECT = 1 << 21;
        const VALUE = 1 << 22;
        const CAPTION = 1 << 23;
        const PICTURE_POSITION = 1 << 24;
        const BORDER_COLOR = 1 << 25;
        const SPECIAL_EFFECT = 1 << 26;
        const MOUSE_ICON = 1 << 27;
        const PICTURE = 1 << 28;
        const ACCELERATOR = 1 << 29;
        const UNUSED2 = 1 << 30;
        const RESERVED = 1 << 31;
        const GROUP_NAME = 1 << 32;
        const UNUSED3 = 0xffff_fffe_0000_0000;
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct MorphDataControlEnvelope {
    pub minor_version: u8,
    pub major_version: u8,
    pub property_mask: MorphDataPropertyMask,
    /// The mask-driven DataBlock and ExtraDataBlock. Their combined boundary
    /// is authoritative even before every optional property is promoted.
    pub data_and_extra: Vec<u8>,
    /// StreamData, TextProps, and optional column information after cbMorphData.
    pub following_data: Vec<u8>,
}

impl MorphDataControlEnvelope {
    const MASK_SIZE: usize = 8;

    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        if bytes.len() < 4 + Self::MASK_SIZE {
            return Err(Error::invalid(0, "truncated MS-OFORMS MorphDataControl"));
        }
        let minor_version = bytes[0];
        let major_version = bytes[1];
        if minor_version != 0 || major_version != 2 {
            return Err(Error::invalid(0, "MorphDataControl version must be 0.2"));
        }
        let property_size = usize::from(u16::from_le_bytes([bytes[2], bytes[3]]));
        if property_size < Self::MASK_SIZE {
            return Err(Error::invalid(
                2,
                "cbMorphData is smaller than MorphDataPropMask",
            ));
        }
        let property_end = 4usize
            .checked_add(property_size)
            .ok_or_else(|| Error::Limit("MorphDataControl property boundary overflow".into()))?;
        if property_end > bytes.len() {
            return Err(Error::invalid(2, "cbMorphData exceeds the control stream"));
        }
        let property_mask = MorphDataPropertyMask::from_bits_retain(u64::from_le_bytes(
            bytes[4..12].try_into().expect("eight-byte mask"),
        ));
        Ok(Self {
            minor_version,
            major_version,
            property_mask,
            data_and_extra: bytes[12..property_end].to_vec(),
            following_data: bytes[property_end..].to_vec(),
        })
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        if self.minor_version != 0 || self.major_version != 2 {
            return Err(Error::invalid(0, "MorphDataControl version must be 0.2"));
        }
        let property_size = Self::MASK_SIZE
            .checked_add(self.data_and_extra.len())
            .ok_or_else(|| Error::Limit("MorphDataControl property size overflow".into()))?;
        let property_size = u16::try_from(property_size)
            .map_err(|_| Error::Limit("cbMorphData exceeds u16".into()))?;
        let mut bytes =
            Vec::with_capacity(4 + usize::from(property_size) + self.following_data.len());
        bytes.push(self.minor_version);
        bytes.push(self.major_version);
        bytes.extend_from_slice(&property_size.to_le_bytes());
        bytes.extend_from_slice(&self.property_mask.bits().to_le_bytes());
        bytes.extend_from_slice(&self.data_and_extra);
        bytes.extend_from_slice(&self.following_data);
        Ok(bytes)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn morph_data_envelope_preserves_mask_driven_boundary() {
        let mut bytes = vec![0, 2, 12, 0];
        bytes.extend_from_slice(
            &(MorphDataPropertyMask::SIZE | MorphDataPropertyMask::DISPLAY_STYLE)
                .bits()
                .to_le_bytes(),
        );
        bytes.extend_from_slice(&[1, 2, 3, 4]);
        bytes.extend_from_slice(&[5, 6]);
        let value = MorphDataControlEnvelope::from_bytes(&bytes).unwrap();
        assert_eq!(value.data_and_extra, [1, 2, 3, 4]);
        assert_eq!(value.following_data, [5, 6]);
        assert_eq!(value.to_bytes().unwrap(), bytes);
    }

    #[test]
    fn morph_data_envelope_rejects_invalid_version_and_boundary() {
        assert!(
            MorphDataControlEnvelope::from_bytes(&[0, 1, 8, 0, 0, 0, 0, 0, 0, 0, 0, 0]).is_err()
        );
        assert!(
            MorphDataControlEnvelope::from_bytes(&[0, 2, 7, 0, 0, 0, 0, 0, 0, 0, 0, 0]).is_err()
        );
        assert!(
            MorphDataControlEnvelope::from_bytes(&[0, 2, 20, 0, 0, 0, 0, 0, 0, 0, 0, 0]).is_err()
        );
    }
}
