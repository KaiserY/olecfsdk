//! Typed Rust SDK for Microsoft Office compound binary file formats.

extern crate self as olecfsdk;

pub mod cfb;
pub mod error;
pub mod io;
pub mod limits;

pub use error::{Error, Result};
pub use olecfsdk_derive::{SdkEnum, SdkObject};
