//! MS-OVBA structures shared by VBA projects embedded in Office files.

use std::path::PathBuf;

use crate::{
    Error, Result,
    cfb::{CompoundFile, Entry},
    common::CodePage,
    limits::Limits,
};

pub mod cache;
pub mod compression;
pub mod directory;
pub mod module;
pub mod project;

use cache::{SrpStream, SrpStreamName, VbaProjectStream};
use compression::CompressedContainer;
use directory::{DirStream, ModuleDescriptor};
use module::ModuleStream;
use project::{ProjectLkStream, ProjectStream, ProjectWmStream};

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct VbaProject {
    pub vba_storage_path: PathBuf,
    pub project: Option<ProjectStream>,
    pub project_wm: Option<ProjectWmStream>,
    pub project_lk: Option<ProjectLkStream>,
    pub directory_container: CompressedContainer,
    pub directory: DirStream,
    pub cache: VbaProjectStream,
    pub srp_streams: Vec<SrpStream>,
    pub modules: Vec<VbaModule>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct VbaModule {
    pub descriptor: ModuleDescriptor,
    pub stream_path: PathBuf,
    pub stream: ModuleStream,
}

impl VbaProject {
    pub fn is_present(compound_file: &CompoundFile) -> bool {
        compound_file
            .entries()
            .iter()
            .any(|entry| entry.is_storage() && entry.name.eq_ignore_ascii_case("VBA"))
    }

    pub fn from_compound_file(compound_file: &CompoundFile) -> Result<Self> {
        Self::from_compound_file_with_limits(compound_file, Limits::default())
    }

    pub fn from_compound_file_with_limits(
        compound_file: &CompoundFile,
        limits: Limits,
    ) -> Result<Self> {
        let vba_storage = compound_file
            .entries()
            .iter()
            .find(|entry| {
                entry.is_storage()
                    && entry.name.eq_ignore_ascii_case("VBA")
                    && child_stream(compound_file, &entry.path, "dir").is_some()
                    && child_stream(compound_file, &entry.path, "_VBA_PROJECT").is_some()
            })
            .ok_or_else(|| Error::invalid(0, "compound file has no VBA storage"))?;
        Self::from_compound_file_at_with_limits(compound_file, &vba_storage.path, limits)
    }

    pub fn from_compound_file_at(
        compound_file: &CompoundFile,
        vba_storage_path: impl AsRef<std::path::Path>,
    ) -> Result<Self> {
        Self::from_compound_file_at_with_limits(compound_file, vba_storage_path, Limits::default())
    }

    pub fn from_compound_file_at_with_limits(
        compound_file: &CompoundFile,
        vba_storage_path: impl AsRef<std::path::Path>,
        limits: Limits,
    ) -> Result<Self> {
        let vba_storage_path = vba_storage_path.as_ref();
        let vba_storage = compound_file
            .entries()
            .iter()
            .find(|entry| entry.is_storage() && entry.path == vba_storage_path)
            .ok_or_else(|| Error::invalid(0, "compound file has no VBA storage at path"))?;
        let vba_storage_path = vba_storage.path.clone();
        let directory_entry = child_stream(compound_file, &vba_storage_path, "dir")
            .ok_or_else(|| Error::invalid(0, "VBA storage has no dir stream"))?;
        let directory_container =
            CompressedContainer::from_bytes_with_limits(&directory_entry.data, limits)?;
        let directory_bytes = directory_container.decompress()?;
        let directory = DirStream::from_bytes_with_limits(&directory_bytes, limits)?;
        let cache_entry = child_stream(compound_file, &vba_storage_path, "_VBA_PROJECT")
            .ok_or_else(|| Error::invalid(0, "VBA storage has no _VBA_PROJECT stream"))?;
        let cache = VbaProjectStream::from_bytes_with_limits(&cache_entry.data, limits)?;
        let mut srp_streams = Vec::new();
        for entry in compound_file.entries().iter().filter(|entry| {
            entry.is_stream() && entry.path.parent() == Some(vba_storage_path.as_path())
        }) {
            if let Some(name) = SrpStreamName::parse(&entry.name)? {
                if entry.data.len() as u64 > limits.max_stream_size {
                    return Err(Error::Limit(format!(
                        "VBA SRP stream length {} exceeds {}",
                        entry.data.len(),
                        limits.max_stream_size
                    )));
                }
                srp_streams.push(SrpStream {
                    path: entry.path.clone(),
                    name,
                    implementation_specific_cache: entry.data.clone(),
                });
            }
        }

        let project = stream_anywhere(compound_file, "PROJECT")
            .map(|entry| ProjectStream::from_bytes_with_limits(&entry.data, limits))
            .transpose()?;
        let project_wm = stream_anywhere(compound_file, "PROJECTwm")
            .map(|entry| ProjectWmStream::from_bytes(&entry.data))
            .transpose()?;
        let project_lk = stream_anywhere(compound_file, "PROJECTlk")
            .map(|entry| ProjectLkStream::from_bytes_with_limits(&entry.data, limits))
            .transpose()?;
        let code_page = CodePage(directory.code_page().unwrap_or(1252));
        let mut modules = Vec::new();
        for descriptor in directory.modules() {
            let stream_name = descriptor.stream_name_with_code_page(code_page)?;
            let entry =
                child_stream(compound_file, &vba_storage_path, &stream_name).ok_or_else(|| {
                    Error::invalid(0, format!("missing VBA module stream {stream_name}"))
                })?;
            let text_offset = descriptor.text_offset.ok_or_else(|| {
                Error::invalid(0, format!("VBA module {stream_name} has no text offset"))
            })?;
            modules.push(VbaModule {
                descriptor,
                stream_path: entry.path.clone(),
                stream: ModuleStream::from_bytes_with_limits(&entry.data, text_offset, limits)?,
            });
        }
        Ok(Self {
            vba_storage_path,
            project,
            project_wm,
            project_lk,
            directory_container,
            directory,
            cache,
            srp_streams,
            modules,
        })
    }

    /// Writes the MS-OVBA interoperable representation atomically.
    ///
    /// Version-dependent caches are discarded, every module source starts at
    /// offset zero, the corresponding dir records are updated and recompressed,
    /// and all SRP streams are removed.
    pub fn write_interoperable_to_compound_file(
        &self,
        compound_file: &mut CompoundFile,
    ) -> Result<()> {
        let mut directory = self.directory.clone();
        let offset_count = directory.set_module_offsets(0);
        if offset_count != self.modules.len() {
            return Err(Error::invalid(
                0,
                format!(
                    "VBA dir has {offset_count} module offsets but {} parsed modules",
                    self.modules.len()
                ),
            ));
        }
        let directory_bytes = directory.to_bytes()?;
        let directory_container = CompressedContainer::from_uncompressed(&directory_bytes);
        let encoded_directory = directory_container.to_bytes()?;
        let encoded_cache = self.cache.to_interoperable_bytes()?;
        let encoded_modules = self
            .modules
            .iter()
            .map(|module| {
                Ok((
                    module.stream_path.clone(),
                    module.stream.to_interoperable_bytes()?,
                ))
            })
            .collect::<Result<Vec<_>>>()?;

        let mut updated = compound_file.clone();
        updated.replace_stream(self.vba_storage_path.join("dir"), encoded_directory)?;
        updated.replace_stream(self.vba_storage_path.join("_VBA_PROJECT"), encoded_cache)?;
        for (path, bytes) in encoded_modules {
            updated.replace_stream(path, bytes)?;
        }
        for srp in &self.srp_streams {
            updated.remove_entry(&srp.path)?;
        }
        *compound_file = updated;
        Ok(())
    }

    pub fn replace_module_source(&mut self, stream_name: &str, source: &[u8]) -> Result<Vec<u8>> {
        let module = self
            .modules
            .iter_mut()
            .find(|module| {
                module
                    .stream_path
                    .file_name()
                    .and_then(|value| value.to_str())
                    .is_some_and(|value| value.eq_ignore_ascii_case(stream_name))
            })
            .ok_or_else(|| {
                Error::invalid(0, format!("VBA project has no module stream {stream_name}"))
            })?;
        module.stream.replace_source_bytes(source)
    }
}

fn child_stream<'a>(
    compound_file: &'a CompoundFile,
    parent: &std::path::Path,
    name: &str,
) -> Option<&'a Entry> {
    compound_file.entries().iter().find(|entry| {
        entry.is_stream()
            && entry.path.parent() == Some(parent)
            && entry.name.eq_ignore_ascii_case(name)
    })
}

fn stream_anywhere<'a>(compound_file: &'a CompoundFile, name: &str) -> Option<&'a Entry> {
    compound_file
        .entries()
        .iter()
        .find(|entry| entry.is_stream() && entry.name.eq_ignore_ascii_case(name))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{
        cfb::Version,
        vba::directory::{
            DirRecord, MarkerRecordKind, MbcsStringRecordKind, U16RecordKind, U32RecordKind,
        },
        vba::project::LicenseInfo,
    };

    #[test]
    fn interoperable_write_rebuilds_offsets_and_discards_all_caches() {
        let directory = DirStream {
            records: vec![
                DirRecord::U16 {
                    kind: U16RecordKind::ProjectCodePage,
                    value: 1252,
                },
                DirRecord::MbcsString {
                    kind: MbcsStringRecordKind::ModuleName,
                    bytes: b"Module1".to_vec(),
                },
                DirRecord::MbcsString {
                    kind: MbcsStringRecordKind::ModuleStreamName,
                    bytes: b"Module1".to_vec(),
                },
                DirRecord::U32 {
                    kind: U32RecordKind::ModuleOffset,
                    value: 2,
                },
                DirRecord::Marker {
                    kind: MarkerRecordKind::ModuleTerminator,
                    reserved: 0,
                },
                DirRecord::Terminator,
            ],
            reserved: 0,
        };
        let encoded_directory =
            CompressedContainer::from_uncompressed(&directory.to_bytes().unwrap())
                .to_bytes()
                .unwrap();
        let cache = VbaProjectStream {
            reserved1: VbaProjectStream::RESERVED1,
            version: 0x1234,
            reserved2: 0,
            reserved3: 7,
            performance_cache: vec![9, 8, 7],
        };
        let mut module = vec![0xaa, 0xbb];
        module.extend_from_slice(
            &CompressedContainer::from_uncompressed(b"Sub Main()\r\nEnd Sub")
                .to_bytes()
                .unwrap(),
        );

        let mut compound = CompoundFile::new(Version::V3).unwrap();
        compound.create_storage("/VBA").unwrap();
        compound
            .create_stream("/VBA/dir", encoded_directory)
            .unwrap();
        compound
            .create_stream("/VBA/_VBA_PROJECT", cache.to_bytes().unwrap())
            .unwrap();
        compound.create_stream("/VBA/Module1", module).unwrap();
        compound
            .create_stream("/VBA/__SRP_A1", vec![1, 2, 3])
            .unwrap();
        let project_lk = ProjectLkStream {
            version: ProjectLkStream::VERSION,
            licenses: vec![LicenseInfo {
                class_id: [0x44; 16],
                license_key: b"opaque-license".to_vec(),
                license_required: 1,
            }],
        };
        compound
            .create_stream("/PROJECTlk", project_lk.to_bytes().unwrap())
            .unwrap();

        let mut project = VbaProject::from_compound_file(&compound).unwrap();
        assert_eq!(project.srp_streams.len(), 1);
        assert_eq!(project.project_lk, Some(project_lk.clone()));
        assert_eq!(
            project
                .replace_module_source("module1", b"Sub Changed()\r\nEnd Sub")
                .unwrap(),
            b"Sub Main()\r\nEnd Sub"
        );
        assert!(project.replace_module_source("missing", b"").is_err());
        project
            .write_interoperable_to_compound_file(&mut compound)
            .unwrap();
        assert!(compound.stream("/VBA/__SRP_A1").is_none());

        let reopened = VbaProject::from_compound_file(&compound).unwrap();
        assert_eq!(
            reopened.cache.version,
            VbaProjectStream::INTEROPERABLE_VERSION
        );
        assert!(reopened.cache.performance_cache.is_empty());
        assert!(reopened.srp_streams.is_empty());
        assert_eq!(reopened.project_lk, Some(project_lk));
        assert_eq!(reopened.directory.module_offsets().collect::<Vec<_>>(), [0]);
        assert!(reopened.modules[0].stream.performance_cache.is_empty());
        assert_eq!(
            reopened.modules[0].stream.source_bytes().unwrap(),
            b"Sub Changed()\r\nEnd Sub"
        );
    }
}
