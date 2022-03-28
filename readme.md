Given an xlsx file, 

1. construct an XlsxPath with _XlsxPath::from_str_,
2.  then use _ImgLoader::new(XlsxPath)_ to copy this xlsx file and unzip it, then parse the xml files to get a map of SheetName -> {(col, row) : imagePath}
