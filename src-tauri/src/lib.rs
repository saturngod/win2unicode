mod win_to_myanmar3;

use std::fs::File;
use std::io::{Read, Write};
use std::path::Path;
use std::collections::HashSet;
use win_to_myanmar3::win_to_myanmar3;
use tauri::{AppHandle, Emitter};
use serde::Serialize;

const TARGET_FONT: &str = "Myanmar Text";

#[derive(Clone, Serialize)]
struct ProgressEvent {
    current: usize,
    total: usize,
    percentage: f64,
    message: String,
}

fn emit_progress(handle: &AppHandle, current: usize, total: usize, message: &str) {
    let percentage = if total > 0 {
        (current as f64 / total as f64) * 100.0
    } else {
        0.0
    };
    let event = ProgressEvent {
        current,
        total,
        percentage,
        message: message.to_string(),
    };
    log::debug!("Progress: {}/{} ({:.1}%) - {}", current, total, percentage, message);

    if let Err(e) = handle.emit("conversion-progress", event) {
        log::error!("Failed to emit progress: {}", e);
    }
}

#[tauri::command]
fn convert_file(
    handle: AppHandle,
    source_path: String,
    target_path: String,
    source_font: String,
) -> Result<(), String> {
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

    log::info!("Starting conversion: {} -> {}", source_path, target_path);
    log::info!("File extension: {}", extension);
    log::info!("Source font: {}", source_font);

    let result = match extension.as_str() {
        "txt" => {
            emit_progress(&handle, 1, 50, "Reading text file...");
            convert_text_file(&handle, source, target).map_err(|e| e.to_string())
        }
        "docx" => {
            emit_progress(&handle, 1, 50, "Reading DOCX file...");
            convert_office_file(&handle, source, target, &source_font, 50).map_err(|e| e.to_string())
        }
        "xlsx" => {
            emit_progress(&handle, 1, 50, "Reading XLSX file...");
            convert_xlsx_file(&handle, source, target, &source_font, 50).map_err(|e| e.to_string())
        }
        "pptx" => {
            emit_progress(&handle, 1, 50, "Reading PPTX file...");
            convert_office_file(&handle, source, target, &source_font, 50).map_err(|e| e.to_string())
        }
        _ => Err("Unsupported file type. Please select txt, docx, xlsx, or pptx.".to_string()),
    };

    if result.is_ok() {
        emit_progress(&handle, 50, 50, "Conversion completed successfully!");
        log::info!("Conversion completed successfully");
    } else {
        log::error!("Conversion failed");
    }

    result
}

#[tauri::command]
fn convert_text(input: String) -> Result<String, String> {
    Ok(win_to_myanmar3(&input))
}

fn convert_text_file(handle: &AppHandle, source: &Path, target: &Path) -> Result<(), Box<dyn std::error::Error>> {
    log::debug!("Reading text file: {:?}", source);
    let content = std::fs::read_to_string(source)?;
    emit_progress(handle, 25, 50, "Converting text content...");

    log::debug!("Converting content with win_to_myanmar3");
    let converted = win_to_myanmar3(&content);

    emit_progress(handle, 45, 50, "Writing converted file...");
    log::debug!("Writing converted file to: {:?}", target);
    std::fs::write(target, converted)?;
    Ok(())
}

fn convert_office_file(handle: &AppHandle, source: &Path, target: &Path, source_font: &str, total_steps: usize) -> Result<(), Box<dyn std::error::Error>> {
    use zip::write::FileOptions;
    use zip::{CompressionMethod, ZipArchive, ZipWriter};

    log::debug!("Opening office file: {:?}", source);
    let source_file = File::open(source)?;
    let mut archive = ZipArchive::new(source_file)?;

    let total_files = archive.len();
    log::debug!("Total files in archive: {}", total_files);

    let target_file = File::create(target)?;
    let mut writer = ZipWriter::new(target_file);
    let options = FileOptions::default().compression_method(CompressionMethod::Deflated);

    for i in 0..archive.len() {
        let progress = 3 + (i * 40 / total_files.max(1));
        emit_progress(handle, progress, total_steps, &format!("Processing file {}/{}", i + 1, total_files));
        let mut file = archive.by_index(i)?;
        let name = file.name().to_string();
        log::trace!("Processing archive entry: {}", name);

        if file.is_dir() {
            writer.add_directory(name, options)?;
            continue;
        }

        let mut contents = Vec::new();
        file.read_to_end(&mut contents)?;

        let updated = if name == "word/document.xml" {
            log::debug!("Processing DOCX document.xml");
            Some(process_docx_xml(handle, &contents, source_font, total_steps))
        } else if name.starts_with("ppt/slides/") && name.ends_with(".xml") {
            log::debug!("Processing PPTX slide: {}", name);
            Some(process_pptx_slide(handle, &contents, source_font, total_steps))
        } else {
            None
        };

        let output_bytes = updated.unwrap_or(contents);
        writer.start_file(name, options)?;
        writer.write_all(&output_bytes)?;
    }

    emit_progress(handle, 48, 50, "Finalizing office file...");
    log::debug!("Writing final office file");
    writer.finish()?;
    emit_progress(handle, 49, 50, "File written successfully");
    Ok(())
}

fn convert_xlsx_file(handle: &AppHandle, source: &Path, target: &Path, source_font: &str, _total_steps: usize) -> Result<(), Box<dyn std::error::Error>> {
    use zip::write::FileOptions;
    use zip::{CompressionMethod, ZipArchive, ZipWriter};

    log::debug!("Opening XLSX file: {:?}", source);
    let source_file = File::open(source)?;
    let mut archive = ZipArchive::new(source_file)?;

    let mut entries: Vec<(String, Vec<u8>, bool)> = Vec::with_capacity(archive.len());

    emit_progress(handle, 3, 50, "Reading XLSX structure...");
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

    emit_progress(handle, 10, 50, "Parsing XLSX styles...");
    let styles_xml = entries
        .iter()
        .find(|(name, _, is_dir)| !is_dir && name == "xl/styles.xml")
        .map(|(_, data, _)| data.clone())
        .unwrap_or_default();

    let (source_font_ids, xf_font_ids) = parse_xlsx_styles(&styles_xml, source_font);
    log::debug!("Found {} source font IDs, {} XF font IDs", source_font_ids.len(), xf_font_ids.len());

    emit_progress(handle, 15, 50, "Analyzing worksheet data...");
    let mut shared_indices: HashSet<usize> = HashSet::new();

    for (name, data, is_dir) in &entries {
        if *is_dir {
            continue;
        }
        if name.starts_with("xl/worksheets/") && name.ends_with(".xml") {
            log::trace!("Analyzing worksheet: {}", name);
            collect_shared_string_indices(data, &source_font_ids, &xf_font_ids, &mut shared_indices);
        }
    }
    log::debug!("Found {} shared string indices to convert", shared_indices.len());

    emit_progress(handle, 20, 50, "Processing shared strings and styles...");
    let target_file = File::create(target)?;
    let mut writer = ZipWriter::new(target_file);
    let options = FileOptions::default().compression_method(CompressionMethod::Deflated);

    let entry_count = entries.len();
    for (idx, (name, data, is_dir)) in entries.into_iter().enumerate() {
        if is_dir {
            writer.add_directory(&name, options)?;
            continue;
        }

        if idx % 3 == 0 {
            let progress = 20 + ((idx * 25) / entry_count.max(1));
            emit_progress(handle, progress, 50, &format!("Writing entry {}/{}", idx + 1, entry_count));
        }

        let updated = if name == "xl/sharedStrings.xml" {
            log::debug!("Processing shared strings XML");
            Some(process_shared_strings(&data, source_font, &shared_indices))
        } else if name == "xl/styles.xml" {
            log::debug!("Processing styles XML");
            Some(process_xlsx_styles(&data, source_font))
        } else {
            None
        };

        let output_bytes = updated.unwrap_or(data);
        writer.start_file(&name, options)?;
        writer.write_all(&output_bytes)?;
    }

    emit_progress(handle, 48, 50, "Finalizing XLSX file...");
    log::debug!("Writing final XLSX file");
    writer.finish()?;
    emit_progress(handle, 49, 50, "XLSX file written successfully");
    Ok(())
}

fn process_docx_xml(handle: &AppHandle, contents: &[u8], source_font: &str, _total_steps: usize) -> Vec<u8> {
    use quick_xml::events::{BytesStart, BytesText, Event};
    use quick_xml::{Reader, Writer};

    log::debug!("Starting DOCX XML processing, size: {} bytes", contents.len());
    let mut reader = Reader::from_reader(contents);
    reader.trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(contents.len()));

    let mut buf = Vec::new();
    let mut in_run = false;
    let mut run_has_font = false;
    let mut event_count = 0usize;
    let mut last_progress = 0usize;

    loop {
        event_count += 1;

        // Emit progress every 200 events processed (for 2% increments)
        if event_count % 200 == 0 {
            let progress = 5 + ((event_count / 200) * 40 / 100);
            if progress != last_progress && progress <= 45 {
                emit_progress(handle, progress, 50, &format!("Converting content... ({} nodes processed)", event_count));
                last_progress = progress;
            }

            // Allow other events to process
            std::thread::yield_now();
        }
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

    log::debug!("DOCX XML processing completed, {} events processed", event_count);
    emit_progress(handle, 46, 50, "Document content processed");

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

fn process_pptx_slide(handle: &AppHandle, contents: &[u8], source_font: &str, _total_steps: usize) -> Vec<u8> {
    use quick_xml::events::{BytesStart, BytesText, Event};
    use quick_xml::{Reader, Writer};

    log::debug!("Starting PPTX slide XML processing, size: {} bytes", contents.len());
    let mut reader = Reader::from_reader(contents);
    reader.trim_text(false);
    let mut writer = Writer::new(Vec::with_capacity(contents.len()));

    let mut buf = Vec::new();
    let mut in_run = false;
    let mut run_has_font = false;
    let mut event_count = 0usize;
    let mut last_progress = 0usize;

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
        event_count += 1;

        // Emit progress every 200 events processed (for 2% increments)
        if event_count % 200 == 0 {
            let progress = 5 + ((event_count / 200) * 40 / 100);
            if progress != last_progress && progress <= 45 {
                emit_progress(handle, progress, 50, &format!("Converting slide content... ({} nodes processed)", event_count));
                last_progress = progress;
            }

            // Allow other events to process
            std::thread::yield_now();
        }
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

    log::debug!("PPTX slide XML processing completed, {} events processed", event_count);
    emit_progress(handle, 46, 50, "Slide content processed");

    writer.into_inner()
}


#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_opener::init())
        .plugin(
            tauri_plugin_log::Builder::new()
                .level(log::LevelFilter::Debug)
                .build(),
        )
        .invoke_handler(tauri::generate_handler![convert_file, convert_text])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
