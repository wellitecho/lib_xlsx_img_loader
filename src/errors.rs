use thiserror::Error;


#[derive(Error, Debug)]
pub enum XlsxPathParseError {
    #[error("invalid format. (expected {expected:?}, found {found:?})")]
    InvalidFormat {
        expected: String,
        found: String,
    },
    #[error("path provided is not a file: {0}")]
    FileNotFound(String),
}


#[derive(Error, Debug)]
pub enum IoError {
    #[error("Failed to create temp dir.")]
    CreateTempDirError{
        msg: String,
        #[source]
        source: std::io::Error,
    },
    #[error(transparent)]
    UnzipXlsxError(#[from] std::io::Error),
}

