use super::errors::{ImgLoaderError, XlsxPathParseError};
use super::*;

use derive_more::Display;
use std::collections::HashMap;
use std::ffi::OsStr;
use std::path::{Path, PathBuf};
use std::primitive::str;
use std::str::FromStr;

pub struct ImgLoader {
    pub xlsx_path: XlsxPath,
    pub worksheet_name_id_map: HashMap<i64, String>,
    pub workbook_name_img_map: HashMap<String, HashMap<(i64, i64), Vec<PathBuf>>>,
}

impl ImgLoader {
    pub fn new(xlsx_path: XlsxPath) -> Result<Self, ImgLoaderError> {
        let temp_dir = Path::new("./temp");
        if !temp_dir.exists() {
            if let Err(e) = std::fs::create_dir_all(temp_dir) {
                return Err(ImgLoaderError::CreateTempDirError(
                    temp_dir.to_str().unwrap().to_owned(),
                ));
            }
        }

        let xlsx_path_buf = xlsx_path.as_pathbuf();

        match unzip_utils::unzip_xlsx(&xlsx_path_buf, temp_dir) {
            Err(e) => return Err(ImgLoaderError::UnzipXlsxError),
            Ok(unzip_dir) => {
                let mut workbook_name_img_map = HashMap::new();
                let media_dir = unzip_dir.join("xl").join("media");

                // TODO: parse workbook_xml, get worksheet names and ids
                let workbook_xml = unzip_dir.join("xl").join("workbook.xml");
                let worksheet_name_id_map =
                    parse_xml::get_worksheet_name_id_map(&workbook_xml.as_path());

                // parse .rels files in worksheet_rels_dir, and get drawing xml and worksheet map
                let drawing_dir = unzip_dir.join("xl").join("drawings");
                let worksheet_rels_dir = unzip_dir.join("xl").join("worksheets").join("_rels");
                let sheet_and_drawing_xml_map = parse_xml::get_sheet_and_drawing_xml_map(
                    &worksheet_rels_dir.as_path(),
                    &drawing_dir.as_path(),
                );

                for (sheet_id, sheet_name) in worksheet_name_id_map.clone() {
                    let sheet_rels_filename = format!("sheet{sheet_id}.xml.rels");
                    if let Some(drawing_xml) = sheet_and_drawing_xml_map.get(&sheet_rels_filename) {
                        let drawing_xml_basename =
                            drawing_xml.file_name().unwrap().to_str().unwrap();
                        let drawing_rels_filename = format!("{drawing_xml_basename}.rels");
                        let drawing_rels_filepath = unzip_dir
                            .join("xl")
                            .join("drawings")
                            .join("_rels")
                            .join(drawing_rels_filename);

                        let col_row_rid = parse_xml::get_col_row_r_id(&drawing_xml.as_path());

                        if let Ok(rid_img_dict) =
                            parse_xml::get_rid_img_dict(&drawing_rels_filepath.as_path())
                        {
                            let col_row_abs_img_dict = parse_xml::generate_col_row_abs_img_dict(
                                col_row_rid,
                                rid_img_dict,
                                media_dir.as_path(),
                            );

                            workbook_name_img_map.insert(sheet_name, col_row_abs_img_dict);
                        }
                    }
                }
                Ok(ImgLoader {
                    xlsx_path,
                    worksheet_name_id_map,
                    workbook_name_img_map,
                })
            }
        }
    }
}

#[derive(Debug, Display, PartialEq, Eq)]
pub struct XlsxPath(String);

impl XlsxPath {
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
            let ext = p
                .extension()
                .and_then(OsStr::to_str)
                .map_or(String::new(), str::to_lowercase);

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