//! retrieve {sheetname: {(col, row): imgpath}} for an xlsx file
//!
//! given an xlsx filepath, this lib copies it,
//!
//! rename it to .zip file and unzip it, then parses .xml files contained,
//!
//! finally, retrieve a map of {sheetname: {(col, row): imgpath}}

mod errors;
mod parse_xml;
mod structs;
mod unzip_utils;

pub use structs::{ImgLoader, XlsxPath};

// #[cfg(test)]
// mod tests {
//     #[test]
//     fn it_works() {
//         let result = 2 + 2;
//         assert_eq!(result, 4);
//     }
// }
