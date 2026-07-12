//! Shared binary strings and code-page handling used across legacy Office formats.

use crate::{Error, Result};

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct CodePage(pub u16);

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct EncodedString {
    pub code_page: CodePage,
    /// Exact persisted bytes, including a terminator when the owning format
    /// defines it as part of the string packet.
    pub bytes: Vec<u8>,
}

impl CodePage {
    pub const WINDOWS_1252: Self = Self(1252);
    pub const UTF_16LE: Self = Self(1200);

    pub fn is_supported(self) -> bool {
        codepage::to_encoding(self.0).is_some()
    }

    pub fn decode(self, bytes: &[u8]) -> Result<String> {
        let encoding = codepage::to_encoding(self.0)
            .ok_or_else(|| Error::invalid(0, format!("unsupported Office code page {}", self.0)))?;
        let (text, had_errors) = encoding.decode_without_bom_handling(bytes);
        if had_errors {
            return Err(Error::invalid(
                0,
                format!("invalid byte sequence for Office code page {}", self.0),
            ));
        }
        Ok(text.into_owned())
    }

    pub fn encode(self, text: &str) -> Result<Vec<u8>> {
        let encoding = codepage::to_encoding(self.0)
            .ok_or_else(|| Error::invalid(0, format!("unsupported Office code page {}", self.0)))?;
        let (bytes, _, had_errors) = encoding.encode(text);
        if had_errors {
            return Err(Error::invalid(
                0,
                format!("text is not representable in Office code page {}", self.0),
            ));
        }
        Ok(bytes.into_owned())
    }
}

impl EncodedString {
    pub fn new(code_page: CodePage, bytes: Vec<u8>) -> Self {
        Self { code_page, bytes }
    }

    pub fn text(&self) -> Result<String> {
        self.code_page.decode(&self.bytes)
    }

    pub fn from_text(code_page: CodePage, text: &str) -> Result<Self> {
        Ok(Self {
            code_page,
            bytes: code_page.encode(text)?,
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn encoded_string_preserves_bytes_and_decodes_by_code_page() {
        let value = EncodedString::new(CodePage::WINDOWS_1252, b"caf\xe9".to_vec());
        assert_eq!(value.text().unwrap(), "caf\u{e9}");
        assert_eq!(
            EncodedString::from_text(CodePage::WINDOWS_1252, "caf\u{e9}").unwrap(),
            value
        );
    }

    #[test]
    fn unsupported_code_page_is_explicit() {
        assert!(!CodePage(0xffff).is_supported());
        assert!(CodePage(0xffff).decode(b"text").is_err());
    }
}
