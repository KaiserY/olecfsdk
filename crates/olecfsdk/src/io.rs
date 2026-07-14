use std::io::{Read, Seek, SeekFrom, Write};

use crate::{Error, Result, limits::Limits};

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum BinaryFormat {
    #[default]
    Unknown,
    Cfb,
    PropertySet,
    Vba,
    Xls,
    Ppt,
    Doc,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum ParseMode {
    #[default]
    Strict,
    Compatible,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct IoContext {
    pub format: BinaryFormat,
    pub version: u32,
    pub code_page: Option<u16>,
    pub mode: ParseMode,
    pub limits: Limits,
}

impl Default for IoContext {
    fn default() -> Self {
        Self {
            format: BinaryFormat::Unknown,
            version: 0,
            code_page: None,
            mode: ParseMode::Strict,
            limits: Limits::default(),
        }
    }
}

pub trait SdkRead: Sized {
    fn read_from<R: Read + Seek>(reader: &mut Reader<R>) -> Result<Self>;
}

pub trait SdkWrite {
    fn write_to<W: Write + Seek>(&self, writer: &mut Writer<W>) -> Result<()>;
}

pub trait SdkSize {
    fn sdk_size(&self) -> u64;
}

pub trait SdkEnumValue: Copy {
    type Repr: Copy + std::fmt::Display;
    fn from_raw(value: Self::Repr) -> Option<Self>;
    fn raw(self) -> Self::Repr;
}

pub struct Reader<R> {
    inner: R,
    start: u64,
    end: u64,
    context: IoContext,
}

impl<R: Read + Seek> Reader<R> {
    pub fn new(mut inner: R) -> Result<Self> {
        let start = inner.stream_position()?;
        let end = inner.seek(SeekFrom::End(0))?;
        inner.seek(SeekFrom::Start(start))?;
        Ok(Self {
            inner,
            start,
            end,
            context: IoContext::default(),
        })
    }

    pub fn with_context(mut inner: R, context: IoContext) -> Result<Self> {
        let start = inner.stream_position()?;
        let end = inner.seek(SeekFrom::End(0))?;
        inner.seek(SeekFrom::Start(start))?;
        Ok(Self {
            inner,
            start,
            end,
            context,
        })
    }

    pub fn with_bounds(mut inner: R, start: u64, len: u64) -> Result<Self> {
        let end = start
            .checked_add(len)
            .ok_or_else(|| Error::invalid(start, "reader bounds overflow"))?;
        let actual_end = inner.seek(SeekFrom::End(0))?;
        if end > actual_end {
            return Err(Error::invalid(start, "reader bounds exceed input"));
        }
        inner.seek(SeekFrom::Start(start))?;
        Ok(Self {
            inner,
            start,
            end,
            context: IoContext::default(),
        })
    }

    pub fn position(&mut self) -> Result<u64> {
        Ok(self.inner.stream_position()?)
    }

    pub fn remaining(&mut self) -> Result<u64> {
        let position = self.position()?;
        self.end
            .checked_sub(position)
            .ok_or_else(|| Error::invalid(position, "reader moved beyond its bounds"))
    }

    pub fn seek_to(&mut self, position: u64) -> Result<()> {
        if position < self.start || position > self.end {
            return Err(Error::invalid(
                position,
                "seek position is outside bounded input",
            ));
        }
        self.inner.seek(SeekFrom::Start(position))?;
        Ok(())
    }

    pub fn start(&self) -> u64 {
        self.start
    }

    pub fn end(&self) -> u64 {
        self.end
    }

    pub fn context(&self) -> &IoContext {
        &self.context
    }

    pub fn context_mut(&mut self) -> &mut IoContext {
        &mut self.context
    }

    pub fn sub_reader(&mut self, len: u64) -> Result<Reader<&mut R>> {
        let start = self.position()?;
        if len > self.remaining()? {
            return Err(Error::invalid(start, "sub-reader exceeds bounded input"));
        }
        let end = start
            .checked_add(len)
            .ok_or_else(|| Error::invalid(start, "sub-reader end overflow"))?;
        Ok(Reader {
            inner: &mut self.inner,
            start,
            end,
            context: self.context,
        })
    }

    pub fn read_vec(&mut self, len: usize) -> Result<Vec<u8>> {
        self.ensure_allocation(len, 1)?;
        let mut value = vec![0; len];
        self.read_exact(&mut value)?;
        Ok(value)
    }

    pub fn ensure_allocation(&self, count: usize, element_size: usize) -> Result<()> {
        let bytes = count
            .checked_mul(element_size)
            .ok_or_else(|| Error::Limit("binary allocation size overflow".into()))?;
        if bytes > self.context.limits.max_allocation {
            return Err(Error::Limit(format!(
                "binary allocation {bytes} exceeds {}",
                self.context.limits.max_allocation
            )));
        }
        Ok(())
    }

    pub fn read_alignment(&mut self, alignment: usize) -> Result<Vec<u8>> {
        if alignment == 0 || !alignment.is_power_of_two() {
            return Err(Error::invalid(
                self.position()?,
                "alignment must be a power of two",
            ));
        }
        let position = usize::try_from(self.position()?)
            .map_err(|_| Error::Limit("reader position does not fit usize".into()))?;
        let padding = position.next_multiple_of(alignment) - position;
        self.read_vec(padding)
    }

    pub fn read_array<const N: usize>(&mut self) -> Result<[u8; N]> {
        let mut value = [0; N];
        self.read_exact(&mut value)?;
        Ok(value)
    }

    fn ensure(&mut self, len: usize) -> Result<()> {
        if self.remaining()? < len as u64 {
            return Err(Error::invalid(self.position()?, "truncated bounded input"));
        }
        Ok(())
    }

    pub fn read_u8(&mut self) -> Result<u8> {
        Ok(self.read_array::<1>()?[0])
    }
    pub fn read_i8(&mut self) -> Result<i8> {
        Ok(self.read_u8()? as i8)
    }
    pub fn read_u16(&mut self) -> Result<u16> {
        Ok(u16::from_le_bytes(self.read_array()?))
    }
    pub fn read_i16(&mut self) -> Result<i16> {
        Ok(i16::from_le_bytes(self.read_array()?))
    }
    pub fn read_u32(&mut self) -> Result<u32> {
        Ok(u32::from_le_bytes(self.read_array()?))
    }
    pub fn read_i32(&mut self) -> Result<i32> {
        Ok(i32::from_le_bytes(self.read_array()?))
    }
    pub fn read_u64(&mut self) -> Result<u64> {
        Ok(u64::from_le_bytes(self.read_array()?))
    }
    pub fn read_i64(&mut self) -> Result<i64> {
        Ok(i64::from_le_bytes(self.read_array()?))
    }
    pub fn read_f32(&mut self) -> Result<f32> {
        Ok(f32::from_le_bytes(self.read_array()?))
    }
    pub fn read_f64(&mut self) -> Result<f64> {
        Ok(f64::from_le_bytes(self.read_array()?))
    }
}

impl<R: Read + Seek> Read for Reader<R> {
    fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
        let position = self.inner.stream_position()?;
        let remaining = self.end.saturating_sub(position);
        let allowed = usize::try_from(remaining.min(buf.len() as u64)).unwrap_or(buf.len());
        self.inner.read(&mut buf[..allowed])
    }

    fn read_exact(&mut self, buf: &mut [u8]) -> std::io::Result<()> {
        self.ensure(buf.len()).map_err(std::io::Error::other)?;
        self.inner.read_exact(buf)
    }
}

pub struct Writer<W> {
    inner: W,
    context: IoContext,
}

impl<W: Write + Seek> Writer<W> {
    pub fn new(inner: W) -> Self {
        Self {
            inner,
            context: IoContext::default(),
        }
    }
    pub fn with_context(inner: W, context: IoContext) -> Self {
        Self { inner, context }
    }
    pub fn position(&mut self) -> Result<u64> {
        Ok(self.inner.stream_position()?)
    }
    pub fn into_inner(self) -> W {
        self.inner
    }
    pub fn context(&self) -> &IoContext {
        &self.context
    }
    pub fn alignment_padding(&mut self, alignment: usize) -> Result<usize> {
        if alignment == 0 || !alignment.is_power_of_two() {
            return Err(Error::invalid(
                self.position()?,
                "alignment must be a power of two",
            ));
        }
        let position = usize::try_from(self.position()?)
            .map_err(|_| Error::Limit("writer position does not fit usize".into()))?;
        Ok(position.next_multiple_of(alignment) - position)
    }
    pub fn write_alignment(&mut self, alignment: usize, value: u8) -> Result<usize> {
        let padding = self.alignment_padding(alignment)?;
        self.write_all(&vec![value; padding])?;
        Ok(padding)
    }
    pub fn write_u8(&mut self, value: u8) -> Result<()> {
        Ok(self.write_all(&[value])?)
    }
    pub fn write_i8(&mut self, value: i8) -> Result<()> {
        self.write_u8(value as u8)
    }
    pub fn write_u16(&mut self, value: u16) -> Result<()> {
        Ok(self.write_all(&value.to_le_bytes())?)
    }
    pub fn write_i16(&mut self, value: i16) -> Result<()> {
        Ok(self.write_all(&value.to_le_bytes())?)
    }
    pub fn write_u32(&mut self, value: u32) -> Result<()> {
        Ok(self.write_all(&value.to_le_bytes())?)
    }
    pub fn write_i32(&mut self, value: i32) -> Result<()> {
        Ok(self.write_all(&value.to_le_bytes())?)
    }
    pub fn write_u64(&mut self, value: u64) -> Result<()> {
        Ok(self.write_all(&value.to_le_bytes())?)
    }
    pub fn write_i64(&mut self, value: i64) -> Result<()> {
        Ok(self.write_all(&value.to_le_bytes())?)
    }
    pub fn write_f32(&mut self, value: f32) -> Result<()> {
        Ok(self.write_all(&value.to_le_bytes())?)
    }
    pub fn write_f64(&mut self, value: f64) -> Result<()> {
        Ok(self.write_all(&value.to_le_bytes())?)
    }
}

impl<W: Write + Seek> Write for Writer<W> {
    fn write(&mut self, buf: &[u8]) -> std::io::Result<usize> {
        self.inner.write(buf)
    }
    fn flush(&mut self) -> std::io::Result<()> {
        self.inner.flush()
    }
}

#[cfg(test)]
mod tests {
    use std::io::Cursor;

    use super::*;
    use crate::{SdkEnum, SdkObject};

    #[derive(Debug, PartialEq, Eq, SdkObject)]
    struct Header {
        a: u16,
        b: u32,
        raw: [u8; 3],
        sectors: [u32; 3],
    }

    #[derive(Clone, Copy, Debug, PartialEq, Eq, SdkEnum)]
    #[sdk(repr = "u16")]
    enum Kind {
        One = 1,
        Two = 2,
    }

    #[derive(Debug, PartialEq, Eq, SdkObject)]
    struct CountedValues {
        count: u16,
        #[sdk(count = "count")]
        values: Vec<u32>,
    }

    #[derive(Debug, PartialEq, Eq, SdkObject)]
    struct ConditionalAndPadding {
        flags: u16,
        #[sdk(condition = "flags", mask = 0x0001)]
        extra: Option<u32>,
        payload_len: u16,
        #[sdk(count = "payload_len")]
        payload: Vec<u8>,
        #[sdk(align = 4)]
        padding: Vec<u8>,
    }

    #[derive(Debug, PartialEq, Eq, SdkObject)]
    struct RemainingBytes {
        tag: u16,
        #[sdk(remaining)]
        tail: Vec<u8>,
    }

    bitflags::bitflags! {
        #[derive(Clone, Copy, Debug, PartialEq, Eq)]
        struct TestFlags: u16 {
            const KNOWN = 0x0001;
        }
    }

    #[derive(Debug, PartialEq, Eq, SdkObject)]
    struct Flagged {
        #[sdk(bitflags = "u16")]
        flags: TestFlags,
    }

    #[test]
    fn derived_binary_round_trip() {
        let value = Header {
            a: 7,
            b: 11,
            raw: [1, 2, 3],
            sectors: [13, 17, u32::MAX],
        };
        let mut writer = Writer::new(Cursor::new(Vec::new()));
        value.write_to(&mut writer).unwrap();
        Kind::Two.write_to(&mut writer).unwrap();
        assert_eq!(value.sdk_size(), 21);
        let bytes = writer.into_inner().into_inner();
        let mut reader = Reader::new(Cursor::new(bytes)).unwrap();
        assert_eq!(Header::read_from(&mut reader).unwrap(), value);
        assert_eq!(Kind::read_from(&mut reader).unwrap(), Kind::Two);
    }

    #[test]
    fn derive_bitflags_retains_unknown_bits() {
        let bytes = [0x01, 0x80];
        let mut reader = Reader::new(Cursor::new(bytes)).unwrap();
        let value = Flagged::read_from(&mut reader).unwrap();
        assert_eq!(value.flags.bits(), 0x8001);
        let mut writer = Writer::new(Cursor::new(Vec::new()));
        value.write_to(&mut writer).unwrap();
        assert_eq!(writer.into_inner().into_inner(), bytes);
    }

    #[test]
    fn derive_reads_and_validates_counted_vectors() {
        let value = CountedValues {
            count: 3,
            values: vec![7, 11, 13],
        };
        let mut writer = Writer::new(Cursor::new(Vec::new()));
        value.write_to(&mut writer).unwrap();
        assert_eq!(value.sdk_size(), 14);
        let mut reader = Reader::new(Cursor::new(writer.into_inner().into_inner())).unwrap();
        assert_eq!(CountedValues::read_from(&mut reader).unwrap(), value);

        let invalid = CountedValues {
            count: 2,
            values: vec![1],
        };
        assert!(
            invalid
                .write_to(&mut Writer::new(Cursor::new(Vec::new())))
                .is_err()
        );
    }

    #[test]
    fn context_and_sub_reader_keep_hard_bounds() {
        let context = IoContext {
            format: BinaryFormat::Xls,
            version: 8,
            limits: Limits {
                max_allocation: 4,
                ..Limits::default()
            },
            ..IoContext::default()
        };
        let mut reader = Reader::with_context(Cursor::new(vec![1, 2, 3, 4, 5]), context).unwrap();
        {
            let mut child = reader.sub_reader(3).unwrap();
            assert_eq!(child.context().format, BinaryFormat::Xls);
            assert_eq!(child.read_vec(3).unwrap(), [1, 2, 3]);
            assert!(child.read_u8().is_err());
        }
        assert_eq!(reader.read_u8().unwrap(), 4);
        assert!(reader.read_vec(5).is_err());
    }

    #[test]
    fn bounded_seek_cannot_escape_reader_limits() {
        let mut reader = Reader::with_bounds(Cursor::new(vec![1, 2, 3, 4, 5]), 1, 3).unwrap();
        assert_eq!(reader.read_u8().unwrap(), 2);
        reader.seek_to(3).unwrap();
        assert_eq!(reader.read_u8().unwrap(), 4);
        assert!(reader.seek_to(0).is_err());
        assert!(reader.seek_to(5).is_err());
        reader.seek_to(4).unwrap();
        assert_eq!(reader.remaining().unwrap(), 0);
    }

    #[test]
    fn derive_supports_conditions_and_preserved_alignment() {
        let value = ConditionalAndPadding {
            flags: 1,
            extra: Some(0x1122_3344),
            payload_len: 3,
            payload: vec![5, 6, 7],
            padding: vec![0],
        };
        assert_eq!(value.sdk_size(), 12);
        let mut writer = Writer::new(Cursor::new(Vec::new()));
        value.write_to(&mut writer).unwrap();
        let mut reader = Reader::new(Cursor::new(writer.into_inner().into_inner())).unwrap();
        assert_eq!(
            ConditionalAndPadding::read_from(&mut reader).unwrap(),
            value
        );

        let invalid = ConditionalAndPadding {
            flags: 0,
            extra: Some(1),
            payload_len: 0,
            payload: Vec::new(),
            padding: Vec::new(),
        };
        assert!(
            invalid
                .write_to(&mut Writer::new(Cursor::new(Vec::new())))
                .is_err()
        );
    }

    #[test]
    fn derive_supports_bounded_remaining_bytes() {
        let value = RemainingBytes {
            tag: 0x1234,
            tail: vec![1, 2, 3, 4],
        };
        let mut writer = Writer::new(Cursor::new(Vec::new()));
        value.write_to(&mut writer).unwrap();
        assert_eq!(value.sdk_size(), 6);
        let mut reader = Reader::new(Cursor::new(writer.into_inner().into_inner())).unwrap();
        assert_eq!(RemainingBytes::read_from(&mut reader).unwrap(), value);
    }
}
