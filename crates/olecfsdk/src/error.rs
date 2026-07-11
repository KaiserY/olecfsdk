use std::io;

#[derive(Debug, thiserror::Error)]
pub enum Error {
    #[error("I/O error: {0}")]
    Io(#[from] io::Error),
    #[error("invalid data at byte {offset}: {message}")]
    InvalidData { offset: u64, message: String },
    #[error("resource limit exceeded: {0}")]
    Limit(String),
}

impl Error {
    pub fn invalid(offset: u64, message: impl Into<String>) -> Self {
        Self::InvalidData {
            offset,
            message: message.into(),
        }
    }
}

pub type Result<T> = std::result::Result<T, Error>;
