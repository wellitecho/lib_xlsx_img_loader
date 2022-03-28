// use anyhow::{anyhow, Error, Result};
use std::fs;
use std::io;
use std::path::{Path, PathBuf};

use super::errors::IoError;
use super::structs::XlsxPath;

pub struct UnzippedPaths {
    pub unzip_dir: PathBuf,
    pub media_dir: PathBuf,
    pub workbook_xml: PathBuf,
    pub drawing_dir: PathBuf,
    pub drawing_rels_dir: PathBuf,
    pub worksheet_rels_dir: PathBuf,
}


pub fn unzip_xlsx<N>(
    xlsx_file: &XlsxPath,
    temp_dir: N,
) -> Result<UnzippedPaths, IoError>
where
    N: AsRef<Path>,
{
    let xlsx_file = xlsx_file.as_pathbuf();

    let file_stem = xlsx_file.file_stem().unwrap().to_str().unwrap();

    let unzip_dir = temp_dir.as_ref().join(file_stem);
    let rename_to_zip = temp_dir.as_ref().join(format!("{file_stem}.zip"));
    // force to overwrite
    fs::copy(xlsx_file, &rename_to_zip)?;

    let file = fs::File::open(&rename_to_zip)?;
    let mut archive = zip::ZipArchive::new(file).unwrap();

    for i in 0..archive.len() {
        let mut file = archive.by_index(i).unwrap();
        let outpath = match file.enclosed_name() {
            Some(path) => unzip_dir.join(path),
            None => continue,
        };

        if (*file.name()).ends_with("/") || (*file.name()).ends_with("\\") {
            fs::create_dir_all(&outpath)?;
        } else {
            if let Some(p) = outpath.parent() {
                if !p.exists() {
                    fs::create_dir_all(&p)?;
                }
            }
            if outpath.exists() {
                fs::remove_file(&outpath)?;
            }
            let mut outfile = fs::File::create(&outpath)?;

            io::copy(&mut file, &mut outfile)?;
        }
    }
    let xl_dir = unzip_dir.join("xl");
    fs::remove_file(&rename_to_zip)?;
    Ok(UnzippedPaths {
        unzip_dir,
        media_dir: xl_dir.join("media"),
        workbook_xml: xl_dir.join("workbook.xml"),
        drawing_dir: xl_dir.join("drawings"),
        drawing_rels_dir: xl_dir.join("drawings").join("_rels"),
        worksheet_rels_dir: xl_dir.join("worksheets").join("_rels"),
    })
}
