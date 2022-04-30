Given an xlsx file, 

1. construct an XlsxPath with _XlsxPath::from_str_,
2.  then use _ImgLoader::new(XlsxPath)_ to copy this xlsx file and unzip it, then parse the xml files to get a map of SheetName -> {(col, row) : imagePath}


# Example
```rust
use lib_xlsx_img_loader::{ImgLoader, XlsxPath};
use read_input::prelude::*;
use std::str::FromStr;

fn main() {
    println!("input xlsx to unzip: ");
    let input_xlsx = input::<String>().get();
    match XlsxPath::from_str(&input_xlsx) {
        Ok(xlsx_path) => {
            if let Ok(ImgLoader {
                ref xlsx_path,
                ref unzip_dir,
                ref worksheet_name_id_map,
                ref worksheet_name_img_map,
                ref worksheet_id_img_map
            }) = ImgLoader::new(&xlsx_path) {
                dbg!(&worksheet_name_id_map);
                dbg!(worksheet_name_img_map);
                println!("Hey, it works!");
            }
        }
        Err(e) => println!("{e}"),
    }
    
}
```