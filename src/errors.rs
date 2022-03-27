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
pub enum ImgLoaderError {
    #[error("cannot create temp dir: {0}")]
    CreateTempDirError(String),
    #[error("failed to unzip xlsx file")]
    UnzipXlsxError,
}