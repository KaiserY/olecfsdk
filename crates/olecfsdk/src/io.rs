use std::io::{Read, Seek, SeekFrom, Write};

use crate::{Error, Result};

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
}

impl<R: Read + Seek> Reader<R> {
    pub fn new(mut inner: R) -> Result<Self> {
        let start = inner.stream_position()?;
        let end = inner.seek(SeekFrom::End(0))?;
        inner.seek(SeekFrom::Start(start))?;
        Ok(Self { inner, start, end })
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
        Ok(Self { inner, start, end })
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

    pub fn start(&self) -> u64 {
        self.start
    }

    pub fn end(&self) -> u64 {
        self.end
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
}

impl<W: Write + Seek> Writer<W> {
    pub fn new(inner: W) -> Self {
        Self { inner }
    }
    pub fn position(&mut self) -> Result<u64> {
        Ok(self.inner.stream_position()?)
    }
    pub fn into_inner(self) -> W {
        self.inner
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
    }

    #[derive(Clone, Copy, Debug, PartialEq, Eq, SdkEnum)]
    #[sdk(repr = "u16")]
    enum Kind {
        One = 1,
        Two = 2,
    }

    #[test]
    fn derived_binary_round_trip() {
        let value = Header {
            a: 7,
            b: 11,
            raw: [1, 2, 3],
        };
        let mut writer = Writer::new(Cursor::new(Vec::new()));
        value.write_to(&mut writer).unwrap();
        Kind::Two.write_to(&mut writer).unwrap();
        assert_eq!(value.sdk_size(), 9);
        let bytes = writer.into_inner().into_inner();
        let mut reader = Reader::new(Cursor::new(bytes)).unwrap();
        assert_eq!(Header::read_from(&mut reader).unwrap(), value);
        assert_eq!(Kind::read_from(&mut reader).unwrap(), Kind::Two);
    }
}
