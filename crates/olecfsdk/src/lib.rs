//! Typed Rust SDK for Microsoft Office compound binary file formats.

extern crate self as olecfsdk;

pub mod cfb;
pub mod common;
pub mod doc;
pub mod error;
pub mod forms;
pub mod io;
pub mod limits;
pub mod office_art;
pub mod parse;
pub mod ppt;
pub mod property_set;
pub mod save;
pub mod shared;
pub mod shared_content;
pub mod vba;
pub mod xls;

pub use error::{Error, Result};
pub use olecfsdk_derive::{SdkBitfield, SdkEnum, SdkObject};
pub use parse::{
    ParseDiagnostic, ParseDiagnosticCode, ParseDiagnosticLocation, ParseDiagnosticSeverity,
    ParseOptions, ParseOutcome, SpecificationReference,
};
pub use save::{CompatibilityWritePolicy, SaveOptions};
