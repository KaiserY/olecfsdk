//! Typed Rust SDK for Microsoft Office compound binary file formats.

extern crate self as olecfsdk;

pub mod cfb;
pub mod common;
pub mod error;
pub mod io;
pub mod limits;
pub mod office_art;
pub mod property_set;
pub mod vba;
pub mod xls;

pub use error::{Error, Result};
pub use olecfsdk_derive::{SdkEnum, SdkObject};
