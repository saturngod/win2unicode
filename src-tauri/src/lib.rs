mod win_to_myanmar3;

use std::fs::File;
use std::io::{Read, Write};
use std::path::Path;
use std::collections::HashSet;
use win_to_myanmar3::win_to_myanmar3;

const TARGET_FONT: &str = "Myanmar Text";

#[tauri::command]
fn convert_file(source_path: String, target_path: String, source_font: String) -> Result<(), String> {
    let source = Path::new(&source_path);
    let target = Path::new(&target_path);

    if !source.exists() {
        return Err("Source file not found.".to_string());
    }

    let extension = source
        .extension()
        .and_then(|ext| ext.to_str())
        .unwrap_or("")
        .to_ascii_lowercase();

    match extension.as_str() {
        "txt" => convert_text_file(source, target).map_err(|e| e.to_string()),
        "docx" => convert_office_file(source, target, &source_font).map_err(|e| e.to_string()),
        "xlsx" => convert_xlsx_file(source, target, &source_font).map_err(|e| e.to_string()),
        "pptx" => convert_office_file(source, target, &source_font).map_err(|e| e.to_string()),
        _ => Err("Unsupported file type. Please select txt, docx, xlsx, or pptx.".to_string()),
    }
}

fn convert_text_file(source: &Path, target: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let content = std::fs::read_to_string(source)?;
    let converted = win_to_myanmar3(&content);
    std::fs::write(target, converted)?;
    Ok(())
}

fn convert_office_file(source: &Path, target: &Path, source_font: &str) -> Result<(), Box<dyn std::error::Error>> {
    use zip::write::FileOptions;
    use zip::{CompressionMethod, ZipArchive, ZipWriter};

    let source_file = File::open(source)?;
    let mut archive = ZipArchive::new(source_file)?;

    let target_file = File::create(target)?;
    let mut writer = ZipWriter::new(target_file);
    let options = FileOptions::default().compression_method(CompressionMethod::Deflated);

    for i in 0..archive.len() {
        let mut file = archive.by_index(i)?;
        let name = file.name().to_string();

        if file.is_dir() {
            writer.add_directory(name, options)?;
            continue;
        }

        let mut contents = Vec::new();
        file.read_to_end(&mut contents)?;

        let updated = if name == "word/document.xml" {
            Some(process_docx_xml(&contents, source_font))
        } else if name.starts_with("ppt/slides/") && name.ends_with(".xml") {
            Some(process_pptx_slide(&contents, source_font))
        } else {
            None
        };

        let output_bytes = updated.unwrap_or(contents);
        writer.start_file(name, options)?;
        writer.write_all(&output_bytes)?;
    }

    writer.finish()?;
    Ok(())
}

fn convert_xlsx_file(source: &Path, target: &Path, source_font: &str) -> Result<(), Box<dyn std::error::Error>> {
    use zip::write::FileOptions;
    use zip::{CompressionMethod, ZipArchive, ZipWriter};

    let source_file = File::open(source)?;
    let mut archive = ZipArchive::new(source_file)?;

    let mut entries: Vec<(String, Vec<u8>, bool)> = Vec::with_capacity(archive.len());

    for i in 0..archive.len() {
        let mut file = archive.by_index(i)?;
        let name = file.name().to_string();
        let is_dir = file.is_dir();
        if is_dir {
            entries.push((name, Vec::new(), true));
            continue;
        }
        let mut contents = Vec::new();
        file.read_to_end(&mut contents)?;
        entries.push((name, contents, false));
    }

    let styles_xml = entries
        .iter()
        .find(|(name, _, is_dir)| !is_dir && name == "xl/styles.xml")
        .map(|(_, data, _)| data.clone())
        .unwrap_or_default();

    let (source_font_ids, xf_font_ids) = parse_xlsx_styles(&styles_xml, source_font);

    let mut shared_indices: HashSet<usize> = HashSet::new();

    for (name, data, is_dir) in &entries {
        if *is_dir {
            continue;
        }
        if name.starts_with("xl/worksheets/") && name.ends_with(".xml") {
            collect_shared_string_indices(data, &source_font_ids, &xf_font_ids, &mut shared_indices);
        }
    }

    let target_file = File::create(target)?;
    let mut writer = ZipWriter::new(target_file);
    let options = FileOptions::default().compression_method(CompressionMethod::Deflated);

    for (name, data, is_dir) in entries {
        if is_dir {
            writer.add_directory(name, options)?;
            continue;
        }

        let updated = if name == "xl/sharedStrings.xml" {
            Some(process_shared_strings(&data, source_font, &shared_indices))
        } else if name == "xl/styles.xml" {
            Some(process_xlsx_styles(&data, source_font))
        } else {
            None
        };

        let output_bytes = updated.unwrap_or(data);
        writer.start_file(name, options)?;
        writer.write_all(&output_bytes)?;
    }

    writer.finish()?;
    Ok(())
}

fn process_docx_xml(contents: &[u8], source_font: &str) -> Vec<u8> {
    use quick_xml::events::{BytesStart, BytesText, Event};
    use quick_xml::{Reader, Writer};

    let mut reader = Reader::from_reader(contents);
    reader.trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(contents.len()));

    let mut buf = Vec::new();
    let mut in_run = false;
    let mut run_has_font = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let elem = e.into_owned();
                let name = elem.name().as_ref().to_vec();

                if name.as_slice() == b"w:r" {
                    in_run = true;
                    run_has_font = false;
                }

                if name.as_slice() == b"w:rFonts" && in_run {
                    let mut new_elem = BytesStart::new("w:rFonts");
                    for attr in elem.attributes().flatten() {
                        let key = attr.key.as_ref();
                        let value = attr.unescape_value().unwrap_or_default().to_string();
                        let is_font_attr = key == b"w:hAnsi" || key == b"w:ascii";
                        if is_font_attr && value == source_font {
                            run_has_font = true;
                            new_elem.push_attribute((key, TARGET_FONT.as_bytes()));
                        } else {
                            new_elem.push_attribute((key, value.as_bytes()));
                        }
                    }
                    writer.write_event(Event::Start(new_elem)).ok();
                } else {
                    writer.write_event(Event::Start(elem)).ok();
                }
            }
            Ok(Event::Empty(e)) => {
                let elem = e.into_owned();
                let name = elem.name().as_ref().to_vec();

                if name.as_slice() == b"w:rFonts" && in_run {
                    let mut new_elem = BytesStart::new("w:rFonts");
                    for attr in elem.attributes().flatten() {
                        let key = attr.key.as_ref();
                        let value = attr.unescape_value().unwrap_or_default().to_string();
                        let is_font_attr = key == b"w:hAnsi" || key == b"w:ascii";
                        if is_font_attr && value == source_font {
                            run_has_font = true;
                            new_elem.push_attribute((key, TARGET_FONT.as_bytes()));
                        } else {
                            new_elem.push_attribute((key, value.as_bytes()));
                        }
                    }
                    writer.write_event(Event::Empty(new_elem)).ok();
                } else {
                    writer.write_event(Event::Empty(elem)).ok();
                }
            }
            Ok(Event::Text(e)) => {
                if in_run && run_has_font {
                    let text = e.unescape().unwrap_or_default().to_string();
                    let converted = win_to_myanmar3(&text);
                    let new_text = BytesText::new(&converted);
                    writer.write_event(Event::Text(new_text)).ok();
                } else {
                    writer.write_event(Event::Text(e.into_owned())).ok();
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"w:r" {
                    in_run = false;
                    run_has_font = false;
                }
                writer.write_event(Event::End(e.into_owned())).ok();
            }
            Ok(Event::CData(e)) => {
                writer.write_event(Event::CData(e.into_owned())).ok();
            }
            Ok(Event::Decl(e)) => {
                writer.write_event(Event::Decl(e.into_owned())).ok();
            }
            Ok(Event::PI(e)) => {
                writer.write_event(Event::PI(e.into_owned())).ok();
            }
            Ok(Event::Comment(e)) => {
                writer.write_event(Event::Comment(e.into_owned())).ok();
            }
            Ok(Event::DocType(e)) => {
                writer.write_event(Event::DocType(e.into_owned())).ok();
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
        }
        buf.clear();
    }

    writer.into_inner()
}

fn process_shared_strings(contents: &[u8], source_font: &str, indices_to_convert: &HashSet<usize>) -> Vec<u8> {
    use quick_xml::events::{BytesStart, BytesText, Event};
    use quick_xml::{Reader, Writer};

    let mut reader = Reader::from_reader(contents);
    reader.trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(contents.len()));

    let mut buf = Vec::new();
    let mut in_run = false;
    let mut run_has_font = false;
    let mut in_si = false;
    let mut si_index: usize = 0;
    let mut convert_si = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let elem = e.into_owned();
                let name = elem.name().as_ref().to_vec();

                if name.as_slice() == b"si" {
                    in_si = true;
                    convert_si = indices_to_convert.contains(&si_index);
                    si_index += 1;
                }

                if name.as_slice() == b"r" {
                    in_run = true;
                    run_has_font = false;
                }

                if name.as_slice() == b"rFont" && in_run {
                    let mut new_elem = BytesStart::new("rFont");
                    for attr in elem.attributes().flatten() {
                        let key = attr.key.as_ref();
                        let value = attr.unescape_value().unwrap_or_default().to_string();
                        let is_val_attr = key == b"val" || key.ends_with(b":val");
                        if is_val_attr && value == source_font {
                            run_has_font = true;
                            new_elem.push_attribute((key, TARGET_FONT.as_bytes()));
                        } else {
                            new_elem.push_attribute((key, value.as_bytes()));
                        }
                    }
                    writer.write_event(Event::Start(new_elem)).ok();
                } else {
                    writer.write_event(Event::Start(elem)).ok();
                }
            }
            Ok(Event::Empty(e)) => {
                let elem = e.into_owned();
                let name = elem.name().as_ref().to_vec();

                if name.as_slice() == b"rFont" && in_run {
                    let mut new_elem = BytesStart::new("rFont");
                    for attr in elem.attributes().flatten() {
                        let key = attr.key.as_ref();
                        let value = attr.unescape_value().unwrap_or_default().to_string();
                        let is_val_attr = key == b"val" || key.ends_with(b":val");
                        if is_val_attr && value == source_font {
                            run_has_font = true;
                            new_elem.push_attribute((key, TARGET_FONT.as_bytes()));
                        } else {
                            new_elem.push_attribute((key, value.as_bytes()));
                        }
                    }
                    writer.write_event(Event::Empty(new_elem)).ok();
                } else {
                    writer.write_event(Event::Empty(elem)).ok();
                }
            }
            Ok(Event::Text(e)) => {
                if (in_run && run_has_font) || (in_si && convert_si) {
                    let text = e.unescape().unwrap_or_default().to_string();
                    let converted = win_to_myanmar3(&text);
                    let new_text = BytesText::new(&converted);
                    writer.write_event(Event::Text(new_text)).ok();
                } else {
                    writer.write_event(Event::Text(e.into_owned())).ok();
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"r" {
                    in_run = false;
                    run_has_font = false;
                }
                if e.name().as_ref() == b"si" {
                    in_si = false;
                    convert_si = false;
                }
                writer.write_event(Event::End(e.into_owned())).ok();
            }
            Ok(Event::CData(e)) => {
                writer.write_event(Event::CData(e.into_owned())).ok();
            }
            Ok(Event::Decl(e)) => {
                writer.write_event(Event::Decl(e.into_owned())).ok();
            }
            Ok(Event::PI(e)) => {
                writer.write_event(Event::PI(e.into_owned())).ok();
            }
            Ok(Event::Comment(e)) => {
                writer.write_event(Event::Comment(e.into_owned())).ok();
            }
            Ok(Event::DocType(e)) => {
                writer.write_event(Event::DocType(e.into_owned())).ok();
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
        }
        buf.clear();
    }

    writer.into_inner()
}

fn process_xlsx_styles(contents: &[u8], source_font: &str) -> Vec<u8> {
    use quick_xml::events::{BytesStart, Event};
    use quick_xml::{Reader, Writer};

    let mut reader = Reader::from_reader(contents);
    reader.trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(contents.len()));

    let mut buf = Vec::new();

    fn tag_matches(name: &[u8], local: &[u8]) -> bool {
        if name == local {
            return true;
        }
        if name.ends_with(local) {
            let idx = name.len() - local.len();
            return idx > 0 && name[idx - 1] == b':';
        }
        false
    }

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let elem = e.into_owned();
                let name = elem.name().as_ref().to_vec();

                if tag_matches(&name, b"name") || tag_matches(&name, b"rFont") {
                    let tag = String::from_utf8_lossy(&name).to_string();
                    let mut new_elem = BytesStart::new(tag.as_str());
                    for attr in elem.attributes().flatten() {
                        let key = attr.key.as_ref();
                        let value = attr.unescape_value().unwrap_or_default().to_string();
                        let is_val_attr = key == b"val" || key.ends_with(b":val");
                        if is_val_attr && value == source_font {
                            new_elem.push_attribute((key, TARGET_FONT.as_bytes()));
                        } else {
                            new_elem.push_attribute((key, value.as_bytes()));
                        }
                    }
                    writer.write_event(Event::Start(new_elem)).ok();
                } else {
                    writer.write_event(Event::Start(elem)).ok();
                }
            }
            Ok(Event::Empty(e)) => {
                let elem = e.into_owned();
                let name = elem.name().as_ref().to_vec();

                if tag_matches(&name, b"name") || tag_matches(&name, b"rFont") {
                    let tag = String::from_utf8_lossy(&name).to_string();
                    let mut new_elem = BytesStart::new(tag.as_str());
                    for attr in elem.attributes().flatten() {
                        let key = attr.key.as_ref();
                        let value = attr.unescape_value().unwrap_or_default().to_string();
                        let is_val_attr = key == b"val" || key.ends_with(b":val");
                        if is_val_attr && value == source_font {
                            new_elem.push_attribute((key, TARGET_FONT.as_bytes()));
                        } else {
                            new_elem.push_attribute((key, value.as_bytes()));
                        }
                    }
                    writer.write_event(Event::Empty(new_elem)).ok();
                } else {
                    writer.write_event(Event::Empty(elem)).ok();
                }
            }
            Ok(Event::Text(e)) => {
                writer.write_event(Event::Text(e.into_owned())).ok();
            }
            Ok(Event::End(e)) => {
                writer.write_event(Event::End(e.into_owned())).ok();
            }
            Ok(Event::CData(e)) => {
                writer.write_event(Event::CData(e.into_owned())).ok();
            }
            Ok(Event::Decl(e)) => {
                writer.write_event(Event::Decl(e.into_owned())).ok();
            }
            Ok(Event::PI(e)) => {
                writer.write_event(Event::PI(e.into_owned())).ok();
            }
            Ok(Event::Comment(e)) => {
                writer.write_event(Event::Comment(e.into_owned())).ok();
            }
            Ok(Event::DocType(e)) => {
                writer.write_event(Event::DocType(e.into_owned())).ok();
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
        }
        buf.clear();
    }

    writer.into_inner()
}

fn parse_xlsx_styles(contents: &[u8], source_font: &str) -> (HashSet<usize>, Vec<usize>) {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut reader = Reader::from_reader(contents);
    reader.trim_text(true);
    let mut buf = Vec::new();

    let mut in_fonts = false;
    let mut in_cell_xfs = false;
    let mut current_font_id: usize = 0;
    let mut source_font_ids: HashSet<usize> = HashSet::new();
    let mut xf_font_ids: Vec<usize> = Vec::new();

    fn tag_matches(name: &[u8], local: &[u8]) -> bool {
        if name == local {
            return true;
        }
        if name.ends_with(local) {
            let idx = name.len() - local.len();
            return idx > 0 && name[idx - 1] == b':';
        }
        false
    }

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name().as_ref().to_vec();
                if name.as_slice() == b"fonts" {
                    in_fonts = true;
                } else if name.as_slice() == b"cellXfs" {
                    in_cell_xfs = true;
                }

                if in_fonts && (tag_matches(&name, b"name") || tag_matches(&name, b"rFont")) {
                    for attr in e.attributes().flatten() {
                        let key = attr.key.as_ref();
                        let value = attr.unescape_value().unwrap_or_default().to_string();
                        let is_val_attr = key == b"val" || key.ends_with(b":val");
                        if is_val_attr && value == source_font {
                            source_font_ids.insert(current_font_id);
                        }
                    }
                } else if in_cell_xfs && name.as_slice() == b"xf" {
                    let mut font_id = 0usize;
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"fontId" {
                            if let Ok(val) = attr.unescape_value() {
                                font_id = val.parse::<usize>().unwrap_or(0);
                            }
                        }
                    }
                    xf_font_ids.push(font_id);
                }
            }
            Ok(Event::Empty(e)) => {
                let name = e.name().as_ref().to_vec();
                if in_fonts && (tag_matches(&name, b"name") || tag_matches(&name, b"rFont")) {
                    for attr in e.attributes().flatten() {
                        let key = attr.key.as_ref();
                        let value = attr.unescape_value().unwrap_or_default().to_string();
                        let is_val_attr = key == b"val" || key.ends_with(b":val");
                        if is_val_attr && value == source_font {
                            source_font_ids.insert(current_font_id);
                        }
                    }
                } else if in_cell_xfs && name.as_slice() == b"xf" {
                    let mut font_id = 0usize;
                    for attr in e.attributes().flatten() {
                        if attr.key.as_ref() == b"fontId" {
                            if let Ok(val) = attr.unescape_value() {
                                font_id = val.parse::<usize>().unwrap_or(0);
                            }
                        }
                    }
                    xf_font_ids.push(font_id);
                }
            }
            Ok(Event::End(e)) => {
                let name = e.name().as_ref().to_vec();
                if name.as_slice() == b"font" && in_fonts {
                    current_font_id += 1;
                } else if name.as_slice() == b"fonts" {
                    in_fonts = false;
                } else if name.as_slice() == b"cellXfs" {
                    in_cell_xfs = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }

    (source_font_ids, xf_font_ids)
}

fn collect_shared_string_indices(
    contents: &[u8],
    source_font_ids: &HashSet<usize>,
    xf_font_ids: &[usize],
    out: &mut HashSet<usize>,
) {
    use quick_xml::events::Event;
    use quick_xml::Reader;

    let mut reader = Reader::from_reader(contents);
    reader.trim_text(true);
    let mut buf = Vec::new();

    let mut current_cell_style: Option<usize> = None;
    let mut current_cell_type: Option<String> = None;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let name = e.name().as_ref().to_vec();
                if name.as_slice() == b"c" {
                    current_cell_style = None;
                    current_cell_type = None;
                    for attr in e.attributes().flatten() {
                        let key = attr.key.as_ref();
                        let value = attr.unescape_value().unwrap_or_default().to_string();
                        if key == b"s" {
                            current_cell_style = value.parse::<usize>().ok();
                        } else if key == b"t" {
                            current_cell_type = Some(value);
                        }
                    }
                }
            }
            Ok(Event::Text(e)) => {
                if let Some(ref cell_type) = current_cell_type {
                    if cell_type == "s" {
                        if let Some(style_idx) = current_cell_style {
                            if let Some(font_id) = xf_font_ids.get(style_idx) {
                                if source_font_ids.contains(font_id) {
                                    let text = e.unescape().unwrap_or_default().to_string();
                                    if let Ok(idx) = text.parse::<usize>() {
                                        out.insert(idx);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            Ok(Event::End(e)) => {
                if e.name().as_ref() == b"c" {
                    current_cell_style = None;
                    current_cell_type = None;
                }
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
            _ => {}
        }
        buf.clear();
    }
}

fn process_pptx_slide(contents: &[u8], source_font: &str) -> Vec<u8> {
    use quick_xml::events::{BytesStart, BytesText, Event};
    use quick_xml::{Reader, Writer};

    let mut reader = Reader::from_reader(contents);
    reader.trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(contents.len()));

    let mut buf = Vec::new();
    let mut in_run = false;
    let mut run_has_font = false;

    fn tag_matches(name: &[u8], local: &[u8]) -> bool {
        if name == local {
            return true;
        }
        if name.ends_with(local) {
            let idx = name.len() - local.len();
            return idx > 0 && name[idx - 1] == b':';
        }
        false
    }

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let elem = e.into_owned();
                let name = elem.name().as_ref().to_vec();

                if tag_matches(&name, b"r") {
                    in_run = true;
                    run_has_font = false;
                }

                if tag_matches(&name, b"rPr") && in_run {
                    let tag = String::from_utf8_lossy(&name).to_string();
                    let mut new_elem = BytesStart::new(tag.as_str());
                    for attr in elem.attributes().flatten() {
                        let key = attr.key.as_ref();
                        let value = attr.unescape_value().unwrap_or_default().to_string();
                        let is_typeface = key == b"typeface" || key.ends_with(b":typeface");
                        if is_typeface && value == source_font {
                            run_has_font = true;
                            new_elem.push_attribute((key, TARGET_FONT.as_bytes()));
                        } else {
                            new_elem.push_attribute((key, value.as_bytes()));
                        }
                    }
                    writer.write_event(Event::Start(new_elem)).ok();
                } else if tag_matches(&name, b"latin") && in_run {
                    let tag = String::from_utf8_lossy(&name).to_string();
                    let mut new_elem = BytesStart::new(tag.as_str());
                    for attr in elem.attributes().flatten() {
                        let key = attr.key.as_ref();
                        let value = attr.unescape_value().unwrap_or_default().to_string();
                        let is_typeface = key == b"typeface" || key.ends_with(b":typeface");
                        if is_typeface && value == source_font {
                            run_has_font = true;
                            new_elem.push_attribute((key, TARGET_FONT.as_bytes()));
                        } else {
                            new_elem.push_attribute((key, value.as_bytes()));
                        }
                    }
                    writer.write_event(Event::Start(new_elem)).ok();
                } else {
                    writer.write_event(Event::Start(elem)).ok();
                }
            }
            Ok(Event::Empty(e)) => {
                let elem = e.into_owned();
                let name = elem.name().as_ref().to_vec();

                if (tag_matches(&name, b"rPr") || tag_matches(&name, b"latin")) && in_run {
                    let tag = String::from_utf8_lossy(&name).to_string();
                    let mut new_elem = BytesStart::new(tag.as_str());
                    for attr in elem.attributes().flatten() {
                        let key = attr.key.as_ref();
                        let value = attr.unescape_value().unwrap_or_default().to_string();
                        let is_typeface = key == b"typeface" || key.ends_with(b":typeface");
                        if is_typeface && value == source_font {
                            run_has_font = true;
                            new_elem.push_attribute((key, TARGET_FONT.as_bytes()));
                        } else {
                            new_elem.push_attribute((key, value.as_bytes()));
                        }
                    }
                    writer.write_event(Event::Empty(new_elem)).ok();
                } else {
                    writer.write_event(Event::Empty(elem)).ok();
                }
            }
            Ok(Event::Text(e)) => {
                if in_run && run_has_font {
                    let text = e.unescape().unwrap_or_default().to_string();
                    let converted = win_to_myanmar3(&text);
                    let new_text = BytesText::new(&converted);
                    writer.write_event(Event::Text(new_text)).ok();
                } else {
                    writer.write_event(Event::Text(e.into_owned())).ok();
                }
            }
            Ok(Event::End(e)) => {
                if tag_matches(e.name().as_ref(), b"r") {
                    in_run = false;
                    run_has_font = false;
                }
                writer.write_event(Event::End(e.into_owned())).ok();
            }
            Ok(Event::CData(e)) => {
                writer.write_event(Event::CData(e.into_owned())).ok();
            }
            Ok(Event::Decl(e)) => {
                writer.write_event(Event::Decl(e.into_owned())).ok();
            }
            Ok(Event::PI(e)) => {
                writer.write_event(Event::PI(e.into_owned())).ok();
            }
            Ok(Event::Comment(e)) => {
                writer.write_event(Event::Comment(e.into_owned())).ok();
            }
            Ok(Event::DocType(e)) => {
                writer.write_event(Event::DocType(e.into_owned())).ok();
            }
            Ok(Event::Eof) => break,
            Err(_) => break,
        }
        buf.clear();
    }

    writer.into_inner()
}


#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_opener::init())
        .invoke_handler(tauri::generate_handler![convert_file])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
