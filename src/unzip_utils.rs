use anyhow::{anyhow, Error, Result};
use std::ffi::OsStr;
use std::fs;
use std::io;
use std::path::{Path, PathBuf};
use std::primitive::str;

pub fn unzip_xlsx(xlsx_file: &PathBuf, temp_dir: &Path) -> Result<PathBuf, Error> {
    if !xlsx_file.is_file() {
        return Err(anyhow!("xlsx file does not exist"));
    }
    let ext = xlsx_file
        .extension()
        .and_then(OsStr::to_str)
        .map_or(String::new(), str::to_lowercase);
    if ext != "xlsx" {
        return Err(anyhow!("Invalid xlsx extension, expect '.xlsx' file"));
    }

    let parent = xlsx_file.parent().unwrap();
    let file_stem = xlsx_file.file_stem().unwrap().to_str().unwrap();

    let temp_dir_xlsx = temp_dir.join(file_stem);
    let rename_to_zip = parent.join(format!("{file_stem}.zip"));
    // force to overwrite
    fs::copy(xlsx_file, &rename_to_zip).unwrap();

    let file = fs::File::open(rename_to_zip).unwrap();
    let mut archive = zip::ZipArchive::new(file).unwrap();

    for i in 0..archive.len() {
        let mut file = archive.by_index(i).unwrap();
        let outpath = match file.enclosed_name() {
            Some(path) => temp_dir_xlsx.join(path),
            None => continue,
        };

        if (*file.name()).ends_with("/") || (*file.name()).ends_with("\\") {
            fs::create_dir_all(&outpath).unwrap();
        } else {
            if let Some(p) = outpath.parent() {
                if !p.exists() {
                    fs::create_dir_all(&p).unwrap();
                }
            }
            if outpath.exists() {
                fs::remove_file(&outpath).unwrap();
            }
            let mut outfile = fs::File::create(&outpath).unwrap();

            io::copy(&mut file, &mut outfile).unwrap();
        }
    }

    Ok(temp_dir_xlsx)
}
