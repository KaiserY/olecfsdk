//! Typed file root for the PowerPoint binary format.

use std::path::Path;

use crate::{Error, Result, cfb::CompoundFile, limits::Limits};

use super::{CurrentUserStream, PicturesStream, PowerPointDocument};

const DOCUMENT_STREAM: &str = "/PowerPoint Document";
const CURRENT_USER_STREAM: &str = "/Current User";
const PICTURES_STREAM: &str = "/Pictures";

/// Complete typed root for a PowerPoint binary file.
///
/// The document stream remains a recursive [`super::PptRecordSequence`]; no
/// content is flattened into text, slide summaries, or image shortcuts.
#[derive(Clone, Debug, PartialEq)]
pub struct PptFile {
    pub compound_file: CompoundFile,
    pub document: PowerPointDocument,
    pub current_user: Option<CurrentUserStream>,
    pub pictures: Option<PicturesStream>,
}

impl PptFile {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_compound_file(CompoundFile::open(path)?)
    }

    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, Limits::default())
    }

    pub fn from_bytes_with_limits(bytes: &[u8], limits: Limits) -> Result<Self> {
        Self::from_compound_file_with_limits(
            CompoundFile::from_bytes_with_limits(bytes, limits)?,
            limits,
        )
    }

    pub fn from_compound_file(compound_file: CompoundFile) -> Result<Self> {
        Self::from_compound_file_with_limits(compound_file, Limits::default())
    }

    pub fn from_compound_file_with_limits(
        compound_file: CompoundFile,
        limits: Limits,
    ) -> Result<Self> {
        let document = compound_file
            .stream(DOCUMENT_STREAM)
            .ok_or_else(|| Error::invalid(0, "PowerPoint Document stream is missing"))
            .and_then(|bytes| PowerPointDocument::from_bytes_with_limits(bytes, limits))?;
        let current_user = compound_file
            .stream(CURRENT_USER_STREAM)
            .map(CurrentUserStream::from_bytes)
            .transpose()?;
        let pictures = compound_file
            .stream(PICTURES_STREAM)
            .map(|bytes| PicturesStream::from_bytes_with_limits(bytes, limits))
            .transpose()?;
        Ok(Self {
            compound_file,
            document,
            current_user,
            pictures,
        })
    }

    /// Rebuilds all managed streams from their typed trees and returns CFB.
    pub fn to_compound_file(&self) -> Result<CompoundFile> {
        let mut compound = self.compound_file.clone();
        compound.replace_stream(DOCUMENT_STREAM, self.document.to_bytes()?)?;
        sync_optional_stream(
            &mut compound,
            CURRENT_USER_STREAM,
            self.current_user.as_ref().map(CurrentUserStream::to_bytes),
        )?;
        sync_optional_stream(
            &mut compound,
            PICTURES_STREAM,
            self.pictures.as_ref().map(PicturesStream::to_bytes),
        )?;
        Ok(compound)
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.to_compound_file()?.to_bytes()
    }

    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.to_compound_file()?.save(path)
    }
}

fn sync_optional_stream(
    compound: &mut CompoundFile,
    path: &str,
    bytes: Option<Result<Vec<u8>>>,
) -> Result<()> {
    match bytes {
        Some(bytes) => {
            compound.create_or_replace_stream(path, bytes?)?;
        }
        None if compound.is_stream(path) => {
            compound.remove_stream(path)?;
        }
        None => {}
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cfb::Version;

    #[test]
    fn file_root_round_trips_the_typed_document_stream() {
        let mut compound = CompoundFile::new(Version::V3).unwrap();
        compound
            .create_or_replace_stream(DOCUMENT_STREAM, Vec::new())
            .unwrap();
        let file = PptFile::from_compound_file(compound).unwrap();
        assert!(file.document.records.records.is_empty());
        let reopened = PptFile::from_bytes(&file.to_bytes().unwrap()).unwrap();
        assert_eq!(reopened.document, file.document);
    }
}
