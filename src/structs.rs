use super::errors::{IoError, XlsxPathParseError};
use super::unzip_utils::UnzippedPaths;
use super::*;

use derive_more::Display;
use std::collections::HashMap;
use std::ffi::OsStr;
use std::path::{Path, PathBuf};
use std::primitive::str;
use std::str::FromStr;

/// main struct to contain the retrieved info
///
///
/// **xlsx_path**: the given xlsx file path, parsed from user input
///
/// **worksheet_name_id_map**: a map of {sheetname: sheet_id}
///
/// **worksheet_name_img_map**: a full map of {sheetname: {(col, row): imgpath}}
///
pub struct ImgLoader {
    pub xlsx_path: XlsxPath,
    pub worksheet_name_id_map: HashMap<i64, String>,
    pub worksheet_name_img_map:
        HashMap<String, HashMap<(i64, i64), Vec<PathBuf>>>,
}

impl ImgLoader {
    /// construct a new ImgLoader
    ///
    /// note: a temp/ dir will be created in the current dir
    pub fn new(xlsx_path: XlsxPath) -> Result<Self, IoError> {
        let temp_dir = Path::new("./temp");
        if !temp_dir.exists() {
            if let Err(e) = std::fs::create_dir_all(temp_dir) {
                return Err(IoError::CreateTempDirError {
                    msg: format!(
                        "cannot create temp dir: {}",
                        temp_dir.display()
                    ),
                    source: e,
                });
            }
        }

        match unzip_utils::unzip_xlsx(&xlsx_path, temp_dir) {
            Err(e) => return Err(e),
            Ok(UnzippedPaths {
                unzip_dir: _,
                media_dir,
                workbook_xml,
                drawing_dir,
                drawing_rels_dir,
                worksheet_rels_dir,
            }) => {
                let mut worksheet_name_img_map = HashMap::new();

                // parse workbook_xml, get worksheet names and ids
                let worksheet_name_id_map =
                    parse_xml::get_worksheet_name_id_map(
                        &workbook_xml.as_path(),
                    );

                // parse .rels files in worksheet_rels_dir, and get drawing xml and worksheet map
                let sheet_and_drawing_xml_map =
                    parse_xml::get_sheet_and_drawing_xml_map(
                        &worksheet_rels_dir.as_path(),
                        &drawing_dir.as_path(),
                    );

                for (sheet_id, sheet_name) in worksheet_name_id_map.clone() {
                    let sheet_rels_filename =
                        format!("sheet{sheet_id}.xml.rels");
                    if let Some(drawing_xml) =
                        sheet_and_drawing_xml_map.get(&sheet_rels_filename)
                    {
                        let drawing_xml_basename =
                            drawing_xml.file_name().unwrap().to_str().unwrap();
                        let drawing_rels_filename =
                            format!("{drawing_xml_basename}.rels");
                        let drawing_rels_filepath =
                            drawing_rels_dir.join(drawing_rels_filename);

                        let col_row_rid = parse_xml::get_col_row_r_id(
                            &drawing_xml.as_path(),
                        );

                        if let Ok(rid_img_dict) = parse_xml::get_rid_img_dict(
                            &drawing_rels_filepath.as_path(),
                        ) {
                            let col_row_abs_img_dict =
                                parse_xml::generate_col_row_abs_img_dict(
                                    col_row_rid,
                                    rid_img_dict,
                                    media_dir.as_path(),
                                );

                            worksheet_name_img_map
                                .insert(sheet_name, col_row_abs_img_dict);
                        }
                    }
                }
                Ok(ImgLoader {
                    xlsx_path,
                    worksheet_name_id_map,
                    worksheet_name_img_map,
                })
            }
        }
    }
}

/// get file extension as string lowercase
fn get_file_ext_lower<S>(filepath: S) -> String
where
    S: AsRef<Path>,
{
    filepath
        .as_ref()
        .extension()
        .and_then(OsStr::to_str)
        .map_or(String::new(), str::to_lowercase)
}

#[derive(Debug, Display, PartialEq, Eq)]
/// a NewType containing a string ended with .xlsx
pub struct XlsxPath(String);

impl XlsxPath {
    /// convert XlsxPath's inner String to PathBuf
    pub fn as_pathbuf(&self) -> PathBuf {
        PathBuf::from(&self.0)
    }
}

impl FromStr for XlsxPath {
    type Err = XlsxPathParseError;

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        let p = Path::new(s);
        if !p.exists() {
            Err(XlsxPathParseError::FileNotFound(s.to_owned()))
        } else {
            let ext = get_file_ext_lower(p);

            if ext != "xlsx" {
                Err(XlsxPathParseError::InvalidFormat {
                    expected: "xlsx".to_owned(),
                    found: ext,
                })
            } else {
                Ok(XlsxPath(s.to_owned()))
            }
        }
    }
}
