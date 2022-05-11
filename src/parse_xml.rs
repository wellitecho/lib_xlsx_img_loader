// use anyhow::Error;
use core::iter::IntoIterator;
use roxmltree::{Document, Namespace, Node};
use std::collections::HashMap;
use std::fs;
use std::path::{Path, PathBuf};

use super::errors::IoError;

#[derive(Debug)]
pub struct CellImgId {
    col: Option<i64>,
    row: Option<i64>,
    r_id: Option<String>,
}

pub fn compute_abs_img_path(
    relative_img_path: &Path,
    media_dir: &Path,
) -> PathBuf {
    let basename = relative_img_path.file_name().unwrap();
    let abs_img = media_dir.join(basename).canonicalize().unwrap();
    // dbg!(&abs_img);
    abs_img
}

fn get_namespace_str<'a>(
    namespaces: &'a [Namespace],
    ident: &'a str,
) -> Option<&'a str> {
    namespaces
        .into_iter()
        .find_map(|n| n.name().and_then(|k| (k == ident).then(|| n.uri())))
}

fn get_node_with_tag_namespace<'a>(
    parent_node: &'a Node,
    namespace_str: &'a str,
    tag_name: &'a str,
) -> Option<Node<'a, 'a>> {
    parent_node
        .descendants()
        .into_iter()
        .find_map(|n| n.has_tag_name((namespace_str, tag_name)).then(|| n))
}

fn get_node_with_attr_namespace<'a>(
    node: &'a Node,
    namespace_str: &'a str,
    attr_name: &'a str,
) -> Option<Node<'a, 'a>> {
    node.descendants()
        .into_iter()
        .find_map(|n| n.has_attribute((namespace_str, attr_name)).then(|| n))
}

fn convert_node_text_to_i64(node: &Node) -> Option<i64> {
    node.text()
        .and_then(|txt| txt.parse::<i64>().ok())
        .and_then(|num| Some(num))
}

pub fn get_col_row_r_id(col_row_r_id_file: &Path) -> Vec<CellImgId> {
    let mut entries: Vec<CellImgId> = Vec::new();
    let file_str = std::fs::read_to_string(col_row_r_id_file).unwrap();
    let doc = Document::parse(&file_str).unwrap();
    let namespaces = doc.root_element().namespaces();

    // let ns_default = Some("http://schemas.openxmlformats.org/package/2006/relationships");
    // let xdr_str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    // let ns_r = get_namespace_str(namespaces, "r");
    let ns_xdr = get_namespace_str(namespaces, "xdr");
    let ns_a = get_namespace_str(namespaces, "a");

    let ns_r = Some(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    );

    if ns_xdr.is_some() && ns_a.is_some() && ns_r.is_some() {
        for two_cell_anchor in doc.descendants().into_iter().filter(|c| {
            c.has_tag_name((ns_xdr.unwrap(), "twoCellAnchor"))
            // c.lookup_namespace_uri(Some("xdr")).is_some()
            //     && c.has_tag_name((ns_xdr.unwrap(), "twoCellAnchor"))
        }) {
            let mut col: Option<i64> = None;
            let mut row: Option<i64> = None;
            let mut r_id: Option<String> = None;

            if let Some(from_node) = get_node_with_tag_namespace(
                &two_cell_anchor,
                ns_xdr.unwrap(),
                "from",
            ) {
                if let Some(col_node) = get_node_with_tag_namespace(
                    &from_node,
                    ns_xdr.unwrap(),
                    "col",
                ) {
                    col = convert_node_text_to_i64(&col_node)
                        .and_then(|num| Some(num + 1));
                }
                if let Some(row_node) = get_node_with_tag_namespace(
                    &from_node,
                    ns_xdr.unwrap(),
                    "row",
                ) {
                    row = convert_node_text_to_i64(&row_node)
                        .and_then(|num| Some(num + 1));
                }
            }

            if let Some(pic_node) = get_node_with_tag_namespace(
                &two_cell_anchor,
                ns_xdr.unwrap(),
                "pic",
            ) {
                if let Some(blip_fill_node) = get_node_with_tag_namespace(
                    &pic_node,
                    ns_xdr.unwrap(),
                    "blipFill",
                ) {
                    if let Some(blip_node) = get_node_with_tag_namespace(
                        &blip_fill_node,
                        ns_a.unwrap(),
                        "blip",
                    ) {
                        if let Some(embed_ele) = get_node_with_attr_namespace(
                            &blip_node,
                            ns_r.unwrap(),
                            "embed",
                        ) {
                            let attrs = embed_ele.attributes();
                            r_id = Some(attrs[0].value().to_owned());
                            //dbg!(r_id);
                        }
                    }
                }
            }

            entries.push(CellImgId { col, row, r_id })
        }
    }

    entries
}

pub fn get_rid_img_dict(
    rels_file: &Path,
) -> Result<HashMap<String, String>, IoError> {
    if let Ok(doc) = Document::parse(&(fs::read_to_string(rels_file)?)) {
        Ok(doc.descendants()
                .into_iter()
                .filter_map(|n| {
                    n.has_tag_name((
                        "http://schemas.openxmlformats.org/package/2006/relationships",
                        "Relationship",
                    ))
                    .then(|| {
                        (
                            n.attribute("Id").unwrap().to_owned(),
                            n.attribute("Target").unwrap().to_owned(),
                        )
                    })
                })
                .collect::<HashMap<String, String>>())
    } else {
        Ok(HashMap::new())
    }
}

pub fn generate_col_row_abs_img_dict(
    col_row_rid: Vec<CellImgId>,
    rid_img_dict: HashMap<String, String>,
    media_dir: &Path,
) -> HashMap<(i64, i64), Vec<PathBuf>> {
    // dbg!(&col_row_rid);
    // dbg!(&rid_img_dict);
    // dbg!(&media_dir);
    let mut col_row_abs_img_dict: HashMap<(i64, i64), Vec<PathBuf>> =
        HashMap::new();
    for entry in col_row_rid {
        let img = rid_img_dict.get(&(entry.r_id.unwrap()));
        if let Some(relative_img_path) = img {
            let relative_img_path = Path::new(relative_img_path);
            let abs_img_path =
                compute_abs_img_path(relative_img_path, media_dir);
            if let Some(imgs) = col_row_abs_img_dict
                .get_mut(&(entry.col.unwrap(), entry.row.unwrap()))
            {
                imgs.push(abs_img_path);
            } else {
                col_row_abs_img_dict.insert(
                    (entry.col.unwrap(), entry.row.unwrap()),
                    vec![abs_img_path],
                );
            }
        }
    }

    col_row_abs_img_dict
}

/// xl/workbook.xml contains the info: worksheet id and worksheet name
pub fn get_worksheet_name_id_map(workbook_xml: &Path) -> HashMap<i64, String> {
    // let mut worksheet_name_id_map: HashMap<String, String> = HashMap::new();
    let file_str = std::fs::read_to_string(workbook_xml).unwrap();
    let doc = Document::parse(&file_str).unwrap();
    let worksheet_name_id_map = doc
        .descendants()
        .into_iter()
        .filter_map(|n| {
            if n.has_tag_name("sheet") {
                let ws_name = n.attribute("name").unwrap();
                let ws_id = n.attribute("sheetId").unwrap();
                Some((ws_id.parse::<i64>().unwrap(), ws_name.to_owned()))
            } else {
                None
            }
        })
        .collect::<HashMap<i64, String>>();

    worksheet_name_id_map
}

pub fn get_sheet_and_drawing_xml_map(
    worksheet_rels_dir: &Path,
    drawing_dir: &Path,
) -> HashMap<String, PathBuf> {
    let mut sheet_drawing_xml_map = HashMap::new();
    let mut rels_files = Vec::new();
    for rel_file in worksheet_rels_dir.read_dir().unwrap() {
        if let Ok(entry) = rel_file {
            let entry_path = entry.path();
            if entry_path.is_file() {
                let ext = entry_path.extension().unwrap().to_str().unwrap();
                if ext.to_lowercase() == "rels" {
                    rels_files.push(entry_path);
                }
            }
        }
    }

    for rels_file in rels_files {
        let rels_file_clone = rels_file.clone();
        let rels_file_last_component =
            rels_file_clone.file_name().unwrap().to_str().unwrap();
        let file_str = std::fs::read_to_string(&rels_file).unwrap();
        let doc = Document::parse(&file_str).unwrap();
        if let Some(relationship_node) = doc
            .descendants()
            .into_iter()
            .find_map(|n| n.has_tag_name("Relationship").then(|| n))
        {
            if let Some(type_str) = relationship_node.attribute("Type") {
                if type_str
                    == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
                {
                    let drawing_xml_relative = relationship_node.attribute("Target").unwrap();
                    let drawing_xml_last_filename = Path::new(drawing_xml_relative)
                        .file_name()
                        .unwrap()
                        .to_str()
                        .unwrap();
                    let drawing_xml = drawing_dir.join(drawing_xml_last_filename);
                    sheet_drawing_xml_map.insert(rels_file_last_component.to_owned(), drawing_xml);
                }
            }
        }
    }

    sheet_drawing_xml_map
}



fn get_node_with_tag<'a>(
    parent_node: &'a Node,
    tag_name: &'a str,
) -> Option<Node<'a, 'a>> {
    parent_node
        .descendants()
        .into_iter()
        .find_map(|n| n.has_tag_name((tag_name)).then(|| n))
}


pub fn get_col_row_r_id_sans_xdr(col_row_r_id_file: &Path) -> Vec<CellImgId> {
    let mut entries: Vec<CellImgId> = Vec::new();
    let file_str = std::fs::read_to_string(col_row_r_id_file).unwrap();
    let doc = Document::parse(&file_str).unwrap();
    let namespaces = doc.root_element().namespaces();

    // let ns_default = Some("http://schemas.openxmlformats.org/package/2006/relationships");
    // let xdr_str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    // let ns_r = get_namespace_str(namespaces, "r");
    let ns_xdr = get_namespace_str(namespaces, "xdr");
    let mut ns_a = get_namespace_str(namespaces, "a");

    if !ns_a.is_some() {
        ns_a = Some("http://schemas.openxmlformats.org/drawingml/2006/main");
    }
    let ns_r = Some(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    );

    if ns_a.is_some() && ns_r.is_some() {
        for two_cell_anchor in doc.descendants().into_iter().filter(|c| {
            c.has_tag_name("twoCellAnchor")
            // c.lookup_namespace_uri(Some("xdr")).is_some()
            //     && c.has_tag_name((ns_xdr.unwrap(), "twoCellAnchor"))
        }) {
            let mut col: Option<i64> = None;
            let mut row: Option<i64> = None;
            let mut r_id: Option<String> = None;

            if let Some(from_node) = get_node_with_tag(
                &two_cell_anchor,
                "from",
            ) {
                if let Some(col_node) = get_node_with_tag(
                    &from_node,
                    "col",
                ) {
                    col = convert_node_text_to_i64(&col_node)
                        .and_then(|num| Some(num + 1));
                }
                if let Some(row_node) = get_node_with_tag(
                    &from_node,
                    "row",
                ) {
                    row = convert_node_text_to_i64(&row_node)
                        .and_then(|num| Some(num + 1));
                }
            }

            if let Some(pic_node) = get_node_with_tag(
                &two_cell_anchor,
                "pic",
            ) {
                if let Some(blip_fill_node) = get_node_with_tag(
                    &pic_node,
                    "blipFill",
                ) {
                    if let Some(blip_node) = get_node_with_tag_namespace(
                        &blip_fill_node,
                        ns_a.unwrap(),
                        "blip",
                    ) {
                        if let Some(embed_ele) = get_node_with_attr_namespace(
                            &blip_node,
                            ns_r.unwrap(),
                            "embed",
                        ) {
                            let attrs = embed_ele.attributes();
                            r_id = Some(attrs[0].value().to_owned());
                          
                        }
                    }
                }
            }

            entries.push(CellImgId { col, row, r_id })
        }
    }

    entries
}