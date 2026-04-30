use std::{
    collections::HashMap,
    fmt,
    io::{Cursor, Read},
    path::{Path, PathBuf},
};

use egui::{Color32, Pos2, Vec2, pos2, vec2};
use flate2::read::DeflateDecoder;
use quick_xml::{
    Reader,
    encoding::Decoder,
    events::{BytesStart, Event},
};
use rich_canvas::{
    CanvasMode, ImageBlock, LayoutRole, RenderBox, RenderBoxKind, RichCanvas, TextAlignment,
    TextRun, TextStyle, TextVerticalAlignment,
};

const ODP_MIME_TYPE: &str = "application/vnd.oasis.opendocument.presentation";
const MAX_ENTRY_SIZE: usize = 64 * 1024 * 1024;
const DEFAULT_SLIDE_SIZE: Vec2 = Vec2::new(1280.0, 720.0);
const PX_PER_CM: f32 = 1280.0 / 28.0;
const PX_PER_IN: f32 = PX_PER_CM * 2.54;

pub const DEFAULT_ODP_PATH: &str =
    concat!(env!("CARGO_MANIFEST_DIR"), "/sample_docs/test_slides.odp");

#[derive(Debug)]
pub struct LoadedOdp {
    pub document_name: String,
    pub slides: Vec<RichCanvas>,
}

#[derive(Debug)]
pub enum OdpLoadError {
    Io(std::io::Error),
    InvalidPackage(String),
    MissingEntry(&'static str),
    Xml(String),
}

impl fmt::Display for OdpLoadError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Io(error) => write!(formatter, "{error}"),
            Self::InvalidPackage(message) => write!(formatter, "{message}"),
            Self::MissingEntry(entry) => write!(formatter, "missing ODP package entry {entry}"),
            Self::Xml(message) => write!(formatter, "{message}"),
        }
    }
}

impl From<std::io::Error> for OdpLoadError {
    fn from(error: std::io::Error) -> Self {
        Self::Io(error)
    }
}

pub fn load_default_odp() -> Result<LoadedOdp, OdpLoadError> {
    load_odp(Path::new(DEFAULT_ODP_PATH))
}

pub fn load_odp(path: &Path) -> Result<LoadedOdp, OdpLoadError> {
    OdpImporter::from_path(path)?.import()
}

struct OdpImporter {
    package: ZipPackage,
    parts: OdpDocumentParts,
}

impl OdpImporter {
    fn from_path(path: &Path) -> Result<Self, OdpLoadError> {
        let package = ZipPackage::open(path)?;
        let parts = OdpDocumentParts::load(&package)?;
        Ok(Self { package, parts })
    }

    fn import(self) -> Result<LoadedOdp, OdpLoadError> {
        let styles = StyleContext::from_parts(&self.parts)?;
        let mut slides =
            SlideImporter::new(&self.package, &styles).parse(&self.parts.content_xml)?;

        if slides.is_empty() {
            return Err(OdpLoadError::InvalidPackage(
                "ODP presentation contains no draw:page slides".to_owned(),
            ));
        }

        for slide in &mut slides {
            slide.relayout(CanvasMode::SlideDeck);
        }

        Ok(LoadedOdp {
            document_name: self.package.document_name(),
            slides,
        })
    }
}

struct OdpDocumentParts {
    content_xml: String,
    styles_xml: String,
}

impl OdpDocumentParts {
    fn load(package: &ZipPackage) -> Result<Self, OdpLoadError> {
        let mimetype = package.entry_text("mimetype")?;
        if mimetype.trim() != ODP_MIME_TYPE {
            return Err(OdpLoadError::InvalidPackage(format!(
                "expected ODP MIME type '{ODP_MIME_TYPE}', found '{}'",
                mimetype.trim()
            )));
        }

        package.entry_bytes("META-INF/manifest.xml")?;
        let content_xml = package.entry_text("content.xml")?;
        if !content_xml.contains("office:presentation") {
            return Err(OdpLoadError::InvalidPackage(
                "content.xml does not contain office:presentation".to_owned(),
            ));
        }

        Ok(Self {
            content_xml,
            styles_xml: package.entry_text("styles.xml")?,
        })
    }
}

struct StyleContext {
    page_size: Vec2,
    text_styles: HashMap<String, TextStyleDef>,
    paragraph_alignments: HashMap<String, TextAlignment>,
    text_vertical_alignments: HashMap<String, TextVerticalAlignment>,
}

impl StyleContext {
    fn from_parts(parts: &OdpDocumentParts) -> Result<Self, OdpLoadError> {
        let page_layouts = parse_page_layouts(&parts.styles_xml)?;
        let master_pages = parse_master_pages(&parts.styles_xml)?;
        Ok(Self {
            page_size: default_page_size(&master_pages, &page_layouts),
            text_styles: parse_text_styles_from_documents(&[
                &parts.styles_xml,
                &parts.content_xml,
            ])?,
            paragraph_alignments: parse_paragraph_alignments_from_documents(&[
                &parts.styles_xml,
                &parts.content_xml,
            ])?,
            text_vertical_alignments: parse_text_vertical_alignments_from_documents(&[
                &parts.styles_xml,
                &parts.content_xml,
            ])?,
        })
    }
}

#[derive(Debug)]
struct ZipPackage {
    path: PathBuf,
    bytes: Vec<u8>,
    entries: HashMap<String, ZipEntry>,
}

#[derive(Clone, Copy, Debug)]
struct ZipEntry {
    method: u16,
    compressed_size: usize,
    uncompressed_size: usize,
    local_header_offset: usize,
}

impl ZipPackage {
    fn open(path: &Path) -> Result<Self, OdpLoadError> {
        let bytes = std::fs::read(path)?;
        let entries = read_central_directory(&bytes)?;
        Ok(Self {
            path: path.to_path_buf(),
            bytes,
            entries,
        })
    }

    fn entry_text(&self, name: &'static str) -> Result<String, OdpLoadError> {
        let bytes = self.entry_bytes(name)?;
        String::from_utf8(bytes)
            .map_err(|error| OdpLoadError::InvalidPackage(format!("{name} is not UTF-8: {error}")))
    }

    fn entry_bytes(&self, name: &str) -> Result<Vec<u8>, OdpLoadError> {
        let entry = self.entries.get(name).ok_or(match name {
            "content.xml" => OdpLoadError::MissingEntry("content.xml"),
            "styles.xml" => OdpLoadError::MissingEntry("styles.xml"),
            "mimetype" => OdpLoadError::MissingEntry("mimetype"),
            _ => OdpLoadError::InvalidPackage(format!("missing package entry {name}")),
        })?;

        if entry.uncompressed_size > MAX_ENTRY_SIZE {
            return Err(OdpLoadError::InvalidPackage(format!(
                "{name} is too large to import safely"
            )));
        }

        let data_start = local_file_data_offset(&self.bytes, entry.local_header_offset)?;
        let data_end = data_start
            .checked_add(entry.compressed_size)
            .ok_or_else(|| OdpLoadError::InvalidPackage("ZIP entry size overflow".to_owned()))?;
        let compressed = self.bytes.get(data_start..data_end).ok_or_else(|| {
            OdpLoadError::InvalidPackage(format!("{name} points outside {}", self.path.display()))
        })?;

        match entry.method {
            0 => Ok(compressed.to_vec()),
            8 => {
                let mut output = Vec::with_capacity(entry.uncompressed_size);
                DeflateDecoder::new(Cursor::new(compressed)).read_to_end(&mut output)?;
                if output.len() > MAX_ENTRY_SIZE {
                    return Err(OdpLoadError::InvalidPackage(format!(
                        "{name} exceeds the decompression limit"
                    )));
                }
                Ok(output)
            }
            method => Err(OdpLoadError::InvalidPackage(format!(
                "{name} uses unsupported ZIP compression method {method}"
            ))),
        }
    }

    fn document_name(&self) -> String {
        self.path
            .file_name()
            .and_then(|name| name.to_str())
            .unwrap_or("test_slides.odp")
            .to_owned()
    }
}

fn read_central_directory(bytes: &[u8]) -> Result<HashMap<String, ZipEntry>, OdpLoadError> {
    let eocd_offset = bytes
        .windows(4)
        .rposition(|window| window == [0x50, 0x4b, 0x05, 0x06])
        .ok_or_else(|| {
            OdpLoadError::InvalidPackage("missing ZIP end of central directory".to_owned())
        })?;
    let eocd = bytes
        .get(eocd_offset..eocd_offset + 22)
        .ok_or_else(|| OdpLoadError::InvalidPackage("truncated ZIP end record".to_owned()))?;
    let entry_count = read_u16(eocd, 10)? as usize;
    let central_dir_offset = read_u32(eocd, 16)? as usize;

    let mut cursor = central_dir_offset;
    let mut entries = HashMap::new();
    for _ in 0..entry_count {
        let header = bytes.get(cursor..cursor + 46).ok_or_else(|| {
            OdpLoadError::InvalidPackage("truncated central directory".to_owned())
        })?;
        if read_u32(header, 0)? != 0x0201_4b50 {
            return Err(OdpLoadError::InvalidPackage(
                "invalid central directory signature".to_owned(),
            ));
        }

        let method = read_u16(header, 10)?;
        let compressed_size = read_u32(header, 20)? as usize;
        let uncompressed_size = read_u32(header, 24)? as usize;
        let name_len = read_u16(header, 28)? as usize;
        let extra_len = read_u16(header, 30)? as usize;
        let comment_len = read_u16(header, 32)? as usize;
        let local_header_offset = read_u32(header, 42)? as usize;
        let name_start = cursor + 46;
        let name_end = name_start + name_len;
        let name = std::str::from_utf8(bytes.get(name_start..name_end).ok_or_else(|| {
            OdpLoadError::InvalidPackage("truncated central directory name".to_owned())
        })?)
        .map_err(|error| {
            OdpLoadError::InvalidPackage(format!("invalid ZIP entry name: {error}"))
        })?;

        if is_unsafe_package_path(name) {
            return Err(OdpLoadError::InvalidPackage(format!(
                "unsafe package path rejected: {name}"
            )));
        }

        entries.insert(
            name.to_owned(),
            ZipEntry {
                method,
                compressed_size,
                uncompressed_size,
                local_header_offset,
            },
        );
        cursor = name_end + extra_len + comment_len;
    }

    Ok(entries)
}

fn local_file_data_offset(bytes: &[u8], offset: usize) -> Result<usize, OdpLoadError> {
    let header = bytes
        .get(offset..offset + 30)
        .ok_or_else(|| OdpLoadError::InvalidPackage("truncated local file header".to_owned()))?;
    if read_u32(header, 0)? != 0x0403_4b50 {
        return Err(OdpLoadError::InvalidPackage(
            "invalid local file header signature".to_owned(),
        ));
    }
    Ok(offset + 30 + read_u16(header, 26)? as usize + read_u16(header, 28)? as usize)
}

fn is_unsafe_package_path(name: &str) -> bool {
    name.starts_with('/')
        || name.starts_with('\\')
        || name.split('/').any(|part| part == "..")
        || name.split('\\').any(|part| part == "..")
}

fn read_u16(bytes: &[u8], offset: usize) -> Result<u16, OdpLoadError> {
    let value = bytes
        .get(offset..offset + 2)
        .ok_or_else(|| OdpLoadError::InvalidPackage("truncated u16".to_owned()))?;
    Ok(u16::from_le_bytes([value[0], value[1]]))
}

fn read_u32(bytes: &[u8], offset: usize) -> Result<u32, OdpLoadError> {
    let value = bytes
        .get(offset..offset + 4)
        .ok_or_else(|| OdpLoadError::InvalidPackage("truncated u32".to_owned()))?;
    Ok(u32::from_le_bytes([value[0], value[1], value[2], value[3]]))
}

#[derive(Clone, Copy, Debug)]
struct PageLayout {
    size: Vec2,
}

fn parse_page_layouts(xml: &str) -> Result<HashMap<String, PageLayout>, OdpLoadError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut layouts = HashMap::new();
    let mut active_name = None;

    loop {
        match reader.read_event() {
            Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"page-layout" => {
                active_name = attr(&event, reader.decoder(), b"name");
            }
            Ok(Event::Empty(event))
                if local_name(event.name().as_ref()) == b"page-layout-properties" =>
            {
                if let Some(name) = active_name.as_ref() {
                    if let (Some(width), Some(height)) = (
                        attr(&event, reader.decoder(), b"page-width")
                            .and_then(|v| parse_length(&v)),
                        attr(&event, reader.decoder(), b"page-height")
                            .and_then(|v| parse_length(&v)),
                    ) {
                        layouts.insert(
                            name.clone(),
                            PageLayout {
                                size: vec2(width, height),
                            },
                        );
                    }
                }
            }
            Ok(Event::End(event)) if local_name(event.name().as_ref()) == b"page-layout" => {
                active_name = None;
            }
            Ok(Event::Eof) => break,
            Err(error) => return Err(OdpLoadError::Xml(error.to_string())),
            _ => {}
        }
    }

    Ok(layouts)
}

fn parse_master_pages(xml: &str) -> Result<HashMap<String, String>, OdpLoadError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut masters = HashMap::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(event)) | Ok(Event::Empty(event))
                if local_name(event.name().as_ref()) == b"master-page" =>
            {
                if let (Some(name), Some(layout)) = (
                    attr(&event, reader.decoder(), b"name"),
                    attr(&event, reader.decoder(), b"page-layout-name"),
                ) {
                    masters.insert(name, layout);
                }
            }
            Ok(Event::Eof) => break,
            Err(error) => return Err(OdpLoadError::Xml(error.to_string())),
            _ => {}
        }
    }

    Ok(masters)
}

fn default_page_size(
    master_pages: &HashMap<String, String>,
    page_layouts: &HashMap<String, PageLayout>,
) -> Vec2 {
    master_pages
        .values()
        .find_map(|layout_name| page_layouts.get(layout_name).map(|layout| layout.size))
        .unwrap_or(DEFAULT_SLIDE_SIZE)
}

#[cfg(test)]
fn parse_text_styles(xml: &str) -> Result<HashMap<String, TextStyle>, OdpLoadError> {
    parse_text_styles_from_documents(&[xml]).map(text_style_defs_to_styles)
}

fn parse_text_styles_from_documents(
    xml_documents: &[&str],
) -> Result<HashMap<String, TextStyleDef>, OdpLoadError> {
    let mut definitions = HashMap::new();
    for xml in xml_documents {
        definitions.extend(parse_text_style_defs(xml)?);
    }

    let mut resolved = HashMap::new();
    for name in definitions.keys() {
        resolved.insert(
            name.clone(),
            resolve_text_style_def(name, &definitions, &mut Vec::new()),
        );
    }

    Ok(resolved)
}

#[derive(Clone, Debug, Default)]
struct TextStyleDef {
    parent: Option<String>,
    font_size: Option<f32>,
    color: Option<Color32>,
    background_color: Option<Option<Color32>>,
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<bool>,
}

impl TextStyleDef {
    fn apply_def(&mut self, other: &Self) {
        if other.font_size.is_some() {
            self.font_size = other.font_size;
        }
        if other.color.is_some() {
            self.color = other.color;
        }
        if other.background_color.is_some() {
            self.background_color = other.background_color;
        }
        if other.bold.is_some() {
            self.bold = other.bold;
        }
        if other.italic.is_some() {
            self.italic = other.italic;
        }
        if other.underline.is_some() {
            self.underline = other.underline;
        }
    }

    fn apply_to_style(&self, style: &mut TextStyle) {
        if let Some(font_size) = self.font_size {
            style.font_size = font_size;
        }
        if let Some(color) = self.color {
            style.color = color;
        }
        if let Some(background_color) = self.background_color {
            style.background_color = background_color;
        }
        if let Some(bold) = self.bold {
            style.bold = bold;
        }
        if let Some(italic) = self.italic {
            style.italic = italic;
        }
        if let Some(underline) = self.underline {
            style.underline = underline;
        }
    }
}

fn parse_text_style_defs(xml: &str) -> Result<HashMap<String, TextStyleDef>, OdpLoadError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut definitions = HashMap::new();
    let mut active_style = None;

    loop {
        match reader.read_event() {
            Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"style" => {
                active_style = attr(&event, reader.decoder(), b"name").map(|name| {
                    let parent = attr(&event, reader.decoder(), b"parent-style-name");
                    (
                        name,
                        TextStyleDef {
                            parent,
                            ..Default::default()
                        },
                    )
                });
            }
            Ok(Event::Empty(event)) if local_name(event.name().as_ref()) == b"style" => {
                if let Some(name) = attr(&event, reader.decoder(), b"name") {
                    definitions.insert(
                        name,
                        TextStyleDef {
                            parent: attr(&event, reader.decoder(), b"parent-style-name"),
                            ..Default::default()
                        },
                    );
                }
            }
            Ok(Event::Empty(event)) | Ok(Event::Start(event))
                if local_name(event.name().as_ref()) == b"text-properties" =>
            {
                if let Some((_name, definition)) = active_style.as_mut() {
                    if let Some(size) =
                        attr(&event, reader.decoder(), b"font-size").and_then(|v| parse_length(&v))
                    {
                        definition.font_size = Some(size);
                    }
                    if let Some(weight) = attr(&event, reader.decoder(), b"font-weight") {
                        definition.bold = Some(weight == "bold");
                    }
                    if let Some(font_style) = attr(&event, reader.decoder(), b"font-style") {
                        definition.italic = Some(font_style == "italic");
                    }
                    if let Some(underline) = attr(&event, reader.decoder(), b"text-underline-style")
                    {
                        definition.underline = Some(underline != "none");
                    }
                    if let Some(color) =
                        attr(&event, reader.decoder(), b"color").and_then(|v| parse_color(&v))
                    {
                        definition.color = Some(color);
                    }
                    if let Some(background_color) =
                        attr(&event, reader.decoder(), b"background-color")
                            .and_then(|v| parse_background_color(&v))
                    {
                        definition.background_color = Some(background_color);
                    }
                }
            }
            Ok(Event::End(event)) if local_name(event.name().as_ref()) == b"style" => {
                if let Some((name, definition)) = active_style.take() {
                    definitions.insert(name, definition);
                }
            }
            Ok(Event::Eof) => break,
            Err(error) => return Err(OdpLoadError::Xml(error.to_string())),
            _ => {}
        }
    }

    Ok(definitions)
}

fn resolve_text_style_def(
    name: &str,
    definitions: &HashMap<String, TextStyleDef>,
    seen: &mut Vec<String>,
) -> TextStyleDef {
    if seen.iter().any(|seen_name| seen_name == name) {
        return TextStyleDef::default();
    }
    seen.push(name.to_owned());

    let Some(definition) = definitions.get(name) else {
        return TextStyleDef::default();
    };
    let mut resolved = definition
        .parent
        .as_deref()
        .map(|parent| resolve_text_style_def(parent, definitions, seen))
        .unwrap_or_default();
    resolved.apply_def(definition);
    resolved.parent = None;
    resolved
}

#[cfg(test)]
fn text_style_defs_to_styles(defs: HashMap<String, TextStyleDef>) -> HashMap<String, TextStyle> {
    defs.into_iter()
        .map(|(name, definition)| {
            let mut style = TextStyle::body();
            definition.apply_to_style(&mut style);
            (name, style)
        })
        .collect()
}

#[cfg(test)]
fn parse_paragraph_alignments(xml: &str) -> Result<HashMap<String, TextAlignment>, OdpLoadError> {
    parse_paragraph_alignments_from_documents(&[xml])
}

fn parse_paragraph_alignments_from_documents(
    xml_documents: &[&str],
) -> Result<HashMap<String, TextAlignment>, OdpLoadError> {
    let mut definitions = HashMap::new();
    for xml in xml_documents {
        definitions.extend(parse_paragraph_alignment_defs(xml)?);
    }

    let mut resolved = HashMap::new();
    for name in definitions.keys() {
        if let Some(alignment) = resolve_paragraph_alignment(name, &definitions, &mut Vec::new()) {
            resolved.insert(name.clone(), alignment);
        }
    }

    Ok(resolved)
}

#[derive(Clone, Debug, Default)]
struct ParagraphStyleDef {
    parent: Option<String>,
    alignment: Option<TextAlignment>,
}

fn parse_paragraph_alignment_defs(
    xml: &str,
) -> Result<HashMap<String, ParagraphStyleDef>, OdpLoadError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut definitions = HashMap::new();
    let mut active_style = None;

    loop {
        match reader.read_event() {
            Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"style" => {
                active_style = attr(&event, reader.decoder(), b"name").map(|name| {
                    let parent = attr(&event, reader.decoder(), b"parent-style-name");
                    (
                        name,
                        ParagraphStyleDef {
                            parent,
                            alignment: None,
                        },
                    )
                });
            }
            Ok(Event::Empty(event)) if local_name(event.name().as_ref()) == b"style" => {
                if let Some(name) = attr(&event, reader.decoder(), b"name") {
                    definitions.insert(
                        name,
                        ParagraphStyleDef {
                            parent: attr(&event, reader.decoder(), b"parent-style-name"),
                            alignment: None,
                        },
                    );
                }
            }
            Ok(Event::Empty(event)) | Ok(Event::Start(event))
                if local_name(event.name().as_ref()) == b"paragraph-properties" =>
            {
                if let (Some((_name, definition)), Some(alignment)) = (
                    active_style.as_mut(),
                    attr(&event, reader.decoder(), b"text-align")
                        .and_then(|value| parse_text_alignment(&value)),
                ) {
                    definition.alignment = Some(alignment);
                }
            }
            Ok(Event::End(event)) if local_name(event.name().as_ref()) == b"style" => {
                if let Some((name, definition)) = active_style.take() {
                    definitions.insert(name, definition);
                }
            }
            Ok(Event::Eof) => break,
            Err(error) => return Err(OdpLoadError::Xml(error.to_string())),
            _ => {}
        }
    }

    Ok(definitions)
}

fn resolve_paragraph_alignment(
    name: &str,
    definitions: &HashMap<String, ParagraphStyleDef>,
    seen: &mut Vec<String>,
) -> Option<TextAlignment> {
    if seen.iter().any(|seen_name| seen_name == name) {
        return None;
    }
    seen.push(name.to_owned());

    let definition = definitions.get(name)?;
    definition.alignment.or_else(|| {
        definition
            .parent
            .as_deref()
            .and_then(|parent| resolve_paragraph_alignment(parent, definitions, seen))
    })
}

fn parse_text_vertical_alignments_from_documents(
    xml_documents: &[&str],
) -> Result<HashMap<String, TextVerticalAlignment>, OdpLoadError> {
    let mut definitions = HashMap::new();
    for xml in xml_documents {
        definitions.extend(parse_text_vertical_alignment_defs(xml)?);
    }

    let mut resolved = HashMap::new();
    for name in definitions.keys() {
        if let Some(alignment) =
            resolve_text_vertical_alignment(name, &definitions, &mut Vec::new())
        {
            resolved.insert(name.clone(), alignment);
        }
    }

    Ok(resolved)
}

#[derive(Clone, Debug, Default)]
struct TextVerticalAlignmentDef {
    parent: Option<String>,
    alignment: Option<TextVerticalAlignment>,
}

fn parse_text_vertical_alignment_defs(
    xml: &str,
) -> Result<HashMap<String, TextVerticalAlignmentDef>, OdpLoadError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut definitions = HashMap::new();
    let mut active_style = None;

    loop {
        match reader.read_event() {
            Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"style" => {
                active_style = attr(&event, reader.decoder(), b"name").map(|name| {
                    (
                        name,
                        TextVerticalAlignmentDef {
                            parent: attr(&event, reader.decoder(), b"parent-style-name"),
                            alignment: None,
                        },
                    )
                });
            }
            Ok(Event::Empty(event)) if local_name(event.name().as_ref()) == b"style" => {
                if let Some(name) = attr(&event, reader.decoder(), b"name") {
                    definitions.insert(
                        name,
                        TextVerticalAlignmentDef {
                            parent: attr(&event, reader.decoder(), b"parent-style-name"),
                            alignment: None,
                        },
                    );
                }
            }
            Ok(Event::Empty(event)) | Ok(Event::Start(event))
                if local_name(event.name().as_ref()) == b"graphic-properties" =>
            {
                if let (Some((_name, definition)), Some(alignment)) = (
                    active_style.as_mut(),
                    attr(&event, reader.decoder(), b"textarea-vertical-align")
                        .and_then(|value| parse_text_vertical_alignment(&value)),
                ) {
                    definition.alignment = Some(alignment);
                }
            }
            Ok(Event::End(event)) if local_name(event.name().as_ref()) == b"style" => {
                if let Some((name, definition)) = active_style.take() {
                    definitions.insert(name, definition);
                }
            }
            Ok(Event::Eof) => break,
            Err(error) => return Err(OdpLoadError::Xml(error.to_string())),
            _ => {}
        }
    }

    Ok(definitions)
}

fn resolve_text_vertical_alignment(
    name: &str,
    definitions: &HashMap<String, TextVerticalAlignmentDef>,
    seen: &mut Vec<String>,
) -> Option<TextVerticalAlignment> {
    if seen.iter().any(|seen_name| seen_name == name) {
        return None;
    }
    seen.push(name.to_owned());

    let definition = definitions.get(name)?;
    definition.alignment.or_else(|| {
        definition
            .parent
            .as_deref()
            .and_then(|parent| resolve_text_vertical_alignment(parent, definitions, seen))
    })
}

struct SlideImporter<'a> {
    package: &'a ZipPackage,
    styles: &'a StyleContext,
    next_box_id: u64,
}

impl<'a> SlideImporter<'a> {
    fn new(package: &'a ZipPackage, styles: &'a StyleContext) -> Self {
        Self {
            package,
            styles,
            next_box_id: 1,
        }
    }

    fn parse(&mut self, xml: &str) -> Result<Vec<RichCanvas>, OdpLoadError> {
        let mut reader = Reader::from_str(xml);
        reader.config_mut().trim_text(false);
        let mut slides = Vec::new();
        let mut canvas = None;
        let mut frame = None;
        let mut in_notes = 0usize;
        let mut text_depth = 0usize;
        let mut style_stack = Vec::new();
        let mut z_index = 0i32;

        loop {
            match reader.read_event() {
                Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"page" => {
                    if in_notes == 0 {
                        canvas = Some(new_slide_canvas(self.styles.page_size));
                        z_index = 0;
                    }
                }
                Ok(Event::End(event)) if local_name(event.name().as_ref()) == b"page" => {
                    if in_notes == 0 {
                        if let Some(mut slide) = canvas.take() {
                            slide.relayout(CanvasMode::SlideDeck);
                            slides.push(slide);
                        }
                    }
                }
                Ok(Event::Start(event)) if local_name(event.name().as_ref()) == b"notes" => {
                    in_notes += 1;
                }
                Ok(Event::End(event)) if local_name(event.name().as_ref()) == b"notes" => {
                    in_notes = in_notes.saturating_sub(1);
                }
                Ok(Event::Start(event))
                    if canvas.is_some()
                        && in_notes == 0
                        && local_name(event.name().as_ref()) == b"frame" =>
                {
                    frame = Some(FrameImport::from_event(&event, reader.decoder()));
                    if let Some(import) = frame.as_mut() {
                        for style_name in frame_alignment_style_names(&event, reader.decoder()) {
                            import.apply_alignment_style(
                                Some(style_name.as_str()),
                                &self.styles.paragraph_alignments,
                            );
                            import.apply_text_style(
                                Some(style_name.as_str()),
                                &self.styles.text_styles,
                            );
                            import.apply_vertical_alignment_style(
                                Some(style_name.as_str()),
                                &self.styles.text_vertical_alignments,
                            );
                        }
                    }
                }
                Ok(Event::End(event)) if local_name(event.name().as_ref()) == b"frame" => {
                    if let (Some(slide), Some(import)) = (canvas.as_mut(), frame.take()) {
                        if let Some(mut render_box) =
                            import.into_render_box(self.next_box_id, self.package)?
                        {
                            render_box.z_index = z_index;
                            slide.push(render_box);
                            self.next_box_id += 1;
                            z_index += 1;
                        }
                    }
                }
                Ok(Event::Start(event))
                    if frame.is_some() && local_name(event.name().as_ref()) == b"text-box" =>
                {
                    text_depth += 1;
                }
                Ok(Event::End(event)) if local_name(event.name().as_ref()) == b"text-box" => {
                    text_depth = text_depth.saturating_sub(1);
                }
                Ok(Event::Start(event))
                    if frame.is_some()
                        && text_depth > 0
                        && local_name(event.name().as_ref()) == b"p" =>
                {
                    if let Some(import) = frame.as_mut() {
                        if !import.runs.is_empty() {
                            import.push_text("\n", current_style(&style_stack));
                        }
                        import.apply_alignment_style(
                            attr(&event, reader.decoder(), b"style-name").as_deref(),
                            &self.styles.paragraph_alignments,
                        );
                        let paragraph_style = style_for_event(
                            &event,
                            reader.decoder(),
                            &self.styles.text_styles,
                            import.base_text_style(),
                        );
                        style_stack.push(paragraph_style);
                    }
                }
                Ok(Event::End(event)) if local_name(event.name().as_ref()) == b"p" => {
                    style_stack.pop();
                }
                Ok(Event::Start(event))
                    if frame.is_some()
                        && text_depth > 0
                        && local_name(event.name().as_ref()) == b"span" =>
                {
                    style_stack.push(style_for_event(
                        &event,
                        reader.decoder(),
                        &self.styles.text_styles,
                        current_style(&style_stack),
                    ));
                }
                Ok(Event::End(event)) if local_name(event.name().as_ref()) == b"span" => {
                    style_stack.pop();
                }
                Ok(Event::Start(event)) | Ok(Event::Empty(event))
                    if frame.is_some() && local_name(event.name().as_ref()) == b"image" =>
                {
                    if let Some(import) = frame.as_mut() {
                        import.image_href = attr(&event, reader.decoder(), b"href");
                    }
                }
                Ok(Event::Text(text)) if frame.is_some() && text_depth > 0 => {
                    let decoded = text
                        .decode()
                        .map_err(|error| OdpLoadError::Xml(error.to_string()))?;
                    if decoded.trim().is_empty() {
                        continue;
                    }
                    if let Some(import) = frame.as_mut() {
                        import.push_text(decoded.as_ref(), current_style(&style_stack));
                    }
                }
                Ok(Event::Eof) => break,
                Err(error) => return Err(OdpLoadError::Xml(error.to_string())),
                _ => {}
            }
        }

        Ok(slides)
    }
}

#[derive(Debug)]
struct FrameImport {
    position: Pos2,
    size: Vec2,
    class_name: Option<String>,
    alignment: Option<TextAlignment>,
    vertical_alignment: Option<TextVerticalAlignment>,
    text_style: Option<TextStyle>,
    runs: Vec<TextRun>,
    image_href: Option<String>,
}

impl FrameImport {
    fn from_event(event: &BytesStart<'_>, decoder: Decoder) -> Self {
        Self {
            position: pos2(
                attr(event, decoder, b"x")
                    .and_then(|v| parse_length(&v))
                    .unwrap_or_default(),
                attr(event, decoder, b"y")
                    .and_then(|v| parse_length(&v))
                    .unwrap_or_default(),
            ),
            size: vec2(
                attr(event, decoder, b"width")
                    .and_then(|v| parse_length(&v))
                    .unwrap_or(320.0),
                attr(event, decoder, b"height")
                    .and_then(|v| parse_length(&v))
                    .unwrap_or(90.0),
            ),
            class_name: attr(event, decoder, b"class"),
            alignment: None,
            vertical_alignment: None,
            text_style: None,
            runs: Vec::new(),
            image_href: None,
        }
    }

    fn apply_alignment_style(
        &mut self,
        style_name: Option<&str>,
        paragraph_alignments: &HashMap<String, TextAlignment>,
    ) {
        if let Some(alignment) = style_name.and_then(|name| paragraph_alignments.get(name)) {
            self.alignment = Some(*alignment);
        }
    }

    fn apply_text_style(
        &mut self,
        style_name: Option<&str>,
        text_styles: &HashMap<String, TextStyleDef>,
    ) {
        if let Some(definition) = style_name.and_then(|name| text_styles.get(name)) {
            let style = self.text_style.get_or_insert_with(TextStyle::body);
            definition.apply_to_style(style);
        }
    }

    fn apply_vertical_alignment_style(
        &mut self,
        style_name: Option<&str>,
        text_vertical_alignments: &HashMap<String, TextVerticalAlignment>,
    ) {
        if let Some(alignment) = style_name.and_then(|name| text_vertical_alignments.get(name)) {
            self.vertical_alignment = Some(*alignment);
        }
    }

    fn base_text_style(&self) -> TextStyle {
        self.text_style.clone().unwrap_or_else(TextStyle::body)
    }

    fn push_text(&mut self, text: &str, style: TextStyle) {
        if text.is_empty() {
            return;
        }
        self.runs.push(TextRun::new(text, style));
    }

    fn into_render_box(
        self,
        id: u64,
        package: &ZipPackage,
    ) -> Result<Option<RenderBox>, OdpLoadError> {
        if let Some(href) = self.image_href {
            let entry_name = href.trim_start_matches("./");
            let mut render_box = RenderBox::image(
                id,
                LayoutRole::Absolute,
                entry_name,
                self.size.max(Vec2::splat(24.0)),
            );
            let bytes = package.entry_bytes(entry_name)?;
            let image = ImageBlock::from_encoded_bytes(
                PathBuf::from(entry_name),
                &bytes,
                Some(self.size),
                0.0,
                0.0,
                0.0,
                0.0,
                false,
            )
            .map_err(|error| {
                OdpLoadError::InvalidPackage(format!(
                    "image {entry_name} could not be decoded: {error}"
                ))
            })?;
            render_box.kind = RenderBoxKind::Image(image);
            render_box.position = self.position;
            render_box.size = self.size.max(Vec2::splat(24.0));
            render_box.authored_size = Some(render_box.size);
            render_box.style.fill = Color32::from_rgb(246, 248, 250);
            return Ok(Some(render_box));
        }

        if self.runs.is_empty() {
            return Ok(None);
        }

        let mut render_box = RenderBox::text(id, LayoutRole::Absolute, self.runs);
        render_box.position = self.position;
        render_box.size = self.size.max(vec2(80.0, 40.0));
        render_box.authored_size = Some(render_box.size);
        if let Some(alignment) = self.alignment {
            render_box.set_text_alignment(alignment);
        }
        if let Some(alignment) = self.vertical_alignment {
            render_box.set_text_vertical_alignment(alignment);
        }
        if self.class_name.as_deref() == Some("title") {
            for run in match &mut render_box.kind {
                RenderBoxKind::Text(block) => &mut block.runs,
                _ => unreachable!(),
            } {
                run.style.font_size = run.style.font_size.max(44.0);
                run.style.bold = true;
            }
        }
        Ok(Some(render_box))
    }
}

fn new_slide_canvas(size: Vec2) -> RichCanvas {
    let mut canvas = RichCanvas::new(size);
    canvas.background = Color32::from_rgb(0x23, 0x23, 0x23);
    canvas.page.fill = Color32::WHITE;
    canvas.page.margin = Vec2::ZERO;
    canvas
}

fn current_style(stack: &[TextStyle]) -> TextStyle {
    stack.last().cloned().unwrap_or_else(TextStyle::body)
}

fn style_for_event(
    event: &BytesStart<'_>,
    decoder: Decoder,
    text_styles: &HashMap<String, TextStyleDef>,
    mut fallback: TextStyle,
) -> TextStyle {
    if let Some(definition) = attr(event, decoder, b"style-name")
        .as_deref()
        .and_then(|name| text_styles.get(name))
    {
        definition.apply_to_style(&mut fallback);
    }
    fallback
}

fn attr(event: &BytesStart<'_>, decoder: Decoder, local: &[u8]) -> Option<String> {
    event
        .attributes()
        .with_checks(false)
        .flatten()
        .find(|attribute| local_name(attribute.key.as_ref()) == local)
        .and_then(|attribute| attribute.decode_and_unescape_value(decoder).ok())
        .map(|value| value.into_owned())
}

fn attr_qualified(event: &BytesStart<'_>, decoder: Decoder, qualified: &[u8]) -> Option<String> {
    event
        .attributes()
        .with_checks(false)
        .flatten()
        .find(|attribute| attribute.key.as_ref() == qualified)
        .and_then(|attribute| attribute.decode_and_unescape_value(decoder).ok())
        .map(|value| value.into_owned())
}

fn frame_alignment_style_names(event: &BytesStart<'_>, decoder: Decoder) -> Vec<String> {
    [
        b"draw:style-name".as_slice(),
        b"presentation:style-name".as_slice(),
        b"draw:text-style-name".as_slice(),
    ]
    .into_iter()
    .filter_map(|name| attr_qualified(event, decoder, name))
    .collect()
}

fn local_name(name: &[u8]) -> &[u8] {
    name.iter()
        .position(|byte| *byte == b':')
        .map_or(name, |index| &name[index + 1..])
}

fn parse_length(value: &str) -> Option<f32> {
    let trimmed = value.trim();
    let (number, factor) = if let Some(number) = trimmed.strip_suffix("cm") {
        (number, PX_PER_CM)
    } else if let Some(number) = trimmed.strip_suffix("mm") {
        (number, PX_PER_CM / 10.0)
    } else if let Some(number) = trimmed.strip_suffix("in") {
        (number, PX_PER_IN)
    } else if let Some(number) = trimmed.strip_suffix("pt") {
        (number, PX_PER_IN / 72.0)
    } else if let Some(number) = trimmed.strip_suffix("px") {
        (number, 1.0)
    } else {
        (trimmed, 1.0)
    };
    number.parse::<f32>().ok().map(|value| value * factor)
}

fn parse_color(value: &str) -> Option<Color32> {
    let hex = value.strip_prefix('#')?;
    if hex.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&hex[0..2], 16).ok()?;
    let g = u8::from_str_radix(&hex[2..4], 16).ok()?;
    let b = u8::from_str_radix(&hex[4..6], 16).ok()?;
    Some(Color32::from_rgb(r, g, b))
}

fn parse_background_color(value: &str) -> Option<Option<Color32>> {
    match value {
        "transparent" => Some(None),
        value => parse_color(value).map(Some),
    }
}

fn parse_text_alignment(value: &str) -> Option<TextAlignment> {
    match value {
        "start" | "left" => Some(TextAlignment::Left),
        "center" => Some(TextAlignment::Center),
        "end" | "right" => Some(TextAlignment::Right),
        "justify" => Some(TextAlignment::Justify),
        _ => None,
    }
}

fn parse_text_vertical_alignment(value: &str) -> Option<TextVerticalAlignment> {
    match value {
        "top" => Some(TextVerticalAlignment::Top),
        "middle" | "center" => Some(TextVerticalAlignment::Center),
        "bottom" => Some(TextVerticalAlignment::Bottom),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::{
        fs,
        path::PathBuf,
        time::{SystemTime, UNIX_EPOCH},
    };

    #[test]
    fn loads_default_odp_into_canvas_slides() {
        let loaded = load_default_odp().expect("default ODP should load");
        assert_eq!(loaded.document_name, "test_slides.odp");
        assert_eq!(loaded.slides.len(), 2);
        assert!(loaded.slides.iter().all(|slide| !slide.boxes.is_empty()));
        assert!(
            loaded.slides[0]
                .boxes
                .iter()
                .any(|render_box| render_box.plain_text().as_deref() == Some("Test Slides"))
        );
        assert!(loaded.slides[0].boxes.iter().any(|render_box| matches!(
            &render_box.kind,
            RenderBoxKind::Text(block)
                if block.plain_text() == "Test Slides" && block.alignment == TextAlignment::Center
        )));
        let subtitle_style = loaded.slides[0].boxes.iter().find_map(|render_box| {
            let RenderBoxKind::Text(block) = &render_box.kind else {
                return None;
            };
            (block.plain_text() == "Just a sample of test slides").then_some(&block.runs[0].style)
        });
        let subtitle_style = subtitle_style.expect("first slide subtitle should load");
        assert_close(subtitle_style.font_size, 32.0 * PX_PER_IN / 72.0);
        assert!(loaded.slides[1].boxes.iter().any(|render_box| {
            let RenderBoxKind::Text(block) = &render_box.kind else {
                return false;
            };
            block
                .runs
                .iter()
                .any(|run| run.text.contains("Highlight") && run.style.background_color.is_some())
        }));
        assert!(loaded.slides[1].boxes.iter().any(RenderBox::is_image));
    }

    #[test]
    fn parses_supported_length_units_and_colors() {
        assert_close(parse_length("28cm").unwrap(), 1280.0);
        assert_close(parse_length("10mm").unwrap(), PX_PER_CM);
        assert_close(parse_length("1in").unwrap(), PX_PER_IN);
        assert_close(parse_length("72pt").unwrap(), PX_PER_IN);
        assert_close(parse_length("42px").unwrap(), 42.0);
        assert!(parse_length("bad").is_none());

        assert_eq!(
            parse_color("#c9211e"),
            Some(Color32::from_rgb(0xc9, 0x21, 0x1e))
        );
        assert_eq!(parse_color("c9211e"), None);
        assert_eq!(parse_color("#nope00"), None);
        assert_eq!(
            parse_background_color("#fff200"),
            Some(Some(Color32::from_rgb(0xff, 0xf2, 0x00)))
        );
        assert_eq!(parse_background_color("transparent"), Some(None));
    }

    #[test]
    fn parses_page_size_from_master_page_layout() {
        let xml = r#"
            <office:document-styles>
                <style:page-layout style:name="PM1">
                    <style:page-layout-properties fo:page-width="28cm" fo:page-height="15.75cm"/>
                </style:page-layout>
                <style:master-page style:name="Default" style:page-layout-name="PM1"/>
            </office:document-styles>
        "#;

        let layouts = parse_page_layouts(xml).expect("page layouts should parse");
        let masters = parse_master_pages(xml).expect("master pages should parse");
        let page_size = default_page_size(&masters, &layouts);

        assert_close(page_size.x, 1280.0);
        assert_close(page_size.y, 720.0);
    }

    #[test]
    fn parses_text_styles_by_expanded_local_names() {
        let xml = r##"
            <office:automatic-styles>
                <style:style style:name="T6" style:family="text">
                    <style:text-properties fo:color="#c9211e" fo:font-size="28pt"
                        fo:font-style="italic" fo:font-weight="bold"
                        fo:background-color="#fff200"
                        style:text-underline-style="solid"/>
                </style:style>
            </office:automatic-styles>
        "##;

        let styles = parse_text_styles(xml).expect("text styles should parse");
        let style = styles.get("T6").expect("style T6 should exist");

        assert_eq!(style.color, Color32::from_rgb(0xc9, 0x21, 0x1e));
        assert_eq!(
            style.background_color,
            Some(Color32::from_rgb(0xff, 0xf2, 0x00))
        );
        assert!(style.bold);
        assert!(style.italic);
        assert!(style.underline);
        assert_close(style.font_size, 28.0 * PX_PER_IN / 72.0);
    }

    #[test]
    fn resolves_text_style_inheritance_across_styles_and_content() {
        let styles_xml = r##"
            <office:document-styles>
                <office:styles>
                    <style:style style:name="BaseText" style:family="presentation">
                        <style:text-properties fo:font-size="32pt" fo:color="#c9211e"
                            fo:font-weight="bold"/>
                    </style:style>
                </office:styles>
            </office:document-styles>
        "##;
        let content_xml = r#"
            <office:document-content>
                <office:automatic-styles>
                    <style:style style:name="ChildText" style:family="text"
                        style:parent-style-name="BaseText">
                        <style:text-properties fo:font-style="italic"/>
                    </style:style>
                </office:automatic-styles>
            </office:document-content>
        "#;

        let styles = parse_text_styles_from_documents(&[styles_xml, content_xml])
            .map(text_style_defs_to_styles)
            .expect("text styles should parse");
        let child = styles.get("ChildText").expect("child style should resolve");

        assert_close(child.font_size, 32.0 * PX_PER_IN / 72.0);
        assert_eq!(child.color, Color32::from_rgb(0xc9, 0x21, 0x1e));
        assert!(child.bold);
        assert!(child.italic);
        assert!(!child.underline);
    }

    #[test]
    fn parses_paragraph_alignment_styles() {
        let xml = r#"
            <office:automatic-styles>
                <style:style style:name="ParentCentered" style:family="paragraph">
                    <style:paragraph-properties fo:text-align="center"/>
                </style:style>
                <style:style style:name="P1" style:family="paragraph">
                    <style:paragraph-properties fo:text-align="start"/>
                </style:style>
                <style:style style:name="P2" style:family="paragraph">
                    <style:paragraph-properties fo:text-align="center"/>
                </style:style>
                <style:style style:name="P3" style:family="paragraph">
                    <style:paragraph-properties fo:text-align="end"/>
                </style:style>
                <style:style style:name="P4" style:family="paragraph">
                    <style:paragraph-properties fo:text-align="justify"/>
                </style:style>
                <style:style style:name="P5" style:family="paragraph"
                    style:parent-style-name="ParentCentered"/>
            </office:automatic-styles>
        "#;

        let alignments = parse_paragraph_alignments(xml).expect("paragraph styles should parse");

        assert_eq!(alignments.get("P1"), Some(&TextAlignment::Left));
        assert_eq!(alignments.get("P2"), Some(&TextAlignment::Center));
        assert_eq!(alignments.get("P3"), Some(&TextAlignment::Right));
        assert_eq!(alignments.get("P4"), Some(&TextAlignment::Justify));
        assert_eq!(alignments.get("P5"), Some(&TextAlignment::Center));
        assert_eq!(parse_text_alignment("unknown"), None);
    }

    #[test]
    fn parses_slide_text_and_ignores_notes_pages() {
        let content_xml = r#"
            <office:document-content>
                <office:body>
                    <office:presentation>
                        <draw:page draw:name="page1">
                            <draw:frame presentation:class="title" svg:x="1cm" svg:y="2cm"
                                svg:width="10cm" svg:height="3cm">
                                <draw:text-box>
                                    <text:p>Visible title</text:p>
                                </draw:text-box>
                            </draw:frame>
                            <presentation:notes>
                                <draw:frame svg:x="0cm" svg:y="0cm" svg:width="1cm" svg:height="1cm">
                                    <draw:text-box><text:p>Hidden note</text:p></draw:text-box>
                                </draw:frame>
                            </presentation:notes>
                        </draw:page>
                    </office:presentation>
                </office:body>
            </office:document-content>
        "#;
        let package = empty_package();
        let styles = StyleContext {
            page_size: DEFAULT_SLIDE_SIZE,
            text_styles: HashMap::new(),
            paragraph_alignments: HashMap::new(),
            text_vertical_alignments: HashMap::new(),
        };
        let slides = SlideImporter::new(&package, &styles)
            .parse(content_xml)
            .expect("slides should parse");

        assert_eq!(slides.len(), 1);
        assert_eq!(slides[0].boxes.len(), 1);
        let title = &slides[0].boxes[0];
        assert_eq!(title.plain_text().as_deref(), Some("Visible title"));
        assert_close(title.position.x, PX_PER_CM);
        assert_close(title.position.y, PX_PER_CM * 2.0);
        assert!(matches!(&title.kind, RenderBoxKind::Text(block) if block.runs[0].style.bold));
    }

    #[test]
    fn parses_text_vertical_alignment_styles() {
        let xml = r#"
            <office:automatic-styles>
                <style:style style:name="ParentMiddle" style:family="presentation">
                    <style:graphic-properties draw:textarea-vertical-align="middle"/>
                </style:style>
                <style:style style:name="Top" style:family="presentation">
                    <style:graphic-properties draw:textarea-vertical-align="top"/>
                </style:style>
                <style:style style:name="Center" style:family="presentation"
                    style:parent-style-name="ParentMiddle"/>
                <style:style style:name="Bottom" style:family="presentation">
                    <style:graphic-properties draw:textarea-vertical-align="bottom"/>
                </style:style>
            </office:automatic-styles>
        "#;

        let alignments = parse_text_vertical_alignments_from_documents(&[xml])
            .expect("vertical alignments should parse");

        assert_eq!(alignments.get("Top"), Some(&TextVerticalAlignment::Top));
        assert_eq!(
            alignments.get("Center"),
            Some(&TextVerticalAlignment::Center)
        );
        assert_eq!(
            alignments.get("Bottom"),
            Some(&TextVerticalAlignment::Bottom)
        );
        assert_eq!(parse_text_vertical_alignment("unknown"), None);
    }

    #[test]
    fn loads_text_alignment_from_frame_and_paragraph_styles() {
        let content_xml = r#"
            <office:document-content>
                <office:body>
                    <office:presentation>
                        <draw:page draw:name="page1">
                            <draw:frame draw:text-style-name="FrameCentered"
                                svg:x="0cm" svg:y="0cm" svg:width="10cm" svg:height="2cm">
                                <draw:text-box><text:p>Frame centered</text:p></draw:text-box>
                            </draw:frame>
                            <draw:frame draw:text-style-name="FrameCentered"
                                svg:x="0cm" svg:y="3cm" svg:width="10cm" svg:height="2cm">
                                <draw:text-box>
                                    <text:p text:style-name="ParagraphRight">Paragraph right</text:p>
                                </draw:text-box>
                            </draw:frame>
                        </draw:page>
                    </office:presentation>
                </office:body>
            </office:document-content>
        "#;
        let package = empty_package();
        let styles = StyleContext {
            page_size: DEFAULT_SLIDE_SIZE,
            text_styles: HashMap::new(),
            paragraph_alignments: HashMap::from([
                ("FrameCentered".to_owned(), TextAlignment::Center),
                ("ParagraphRight".to_owned(), TextAlignment::Right),
            ]),
            text_vertical_alignments: HashMap::new(),
        };

        let slides = SlideImporter::new(&package, &styles)
            .parse(content_xml)
            .expect("slides should parse");

        assert_eq!(slides.len(), 1);
        assert_eq!(slides[0].boxes.len(), 2);
        assert!(matches!(
            &slides[0].boxes[0].kind,
            RenderBoxKind::Text(block) if block.alignment == TextAlignment::Center
        ));
        assert!(matches!(
            &slides[0].boxes[1].kind,
            RenderBoxKind::Text(block) if block.alignment == TextAlignment::Right
        ));
    }

    #[test]
    fn loads_text_vertical_alignment_from_frame_styles() {
        let content_xml = r#"
            <office:document-content>
                <office:body>
                    <office:presentation>
                        <draw:page draw:name="page1">
                            <draw:frame presentation:style-name="TopAligned"
                                svg:x="0cm" svg:y="0cm" svg:width="10cm" svg:height="2cm">
                                <draw:text-box><text:p>Top</text:p></draw:text-box>
                            </draw:frame>
                            <draw:frame presentation:style-name="MiddleAligned"
                                svg:x="0cm" svg:y="3cm" svg:width="10cm" svg:height="2cm">
                                <draw:text-box><text:p>Middle</text:p></draw:text-box>
                            </draw:frame>
                        </draw:page>
                    </office:presentation>
                </office:body>
            </office:document-content>
        "#;
        let package = empty_package();
        let styles = StyleContext {
            page_size: DEFAULT_SLIDE_SIZE,
            text_styles: HashMap::new(),
            paragraph_alignments: HashMap::new(),
            text_vertical_alignments: HashMap::from([
                ("TopAligned".to_owned(), TextVerticalAlignment::Top),
                ("MiddleAligned".to_owned(), TextVerticalAlignment::Center),
            ]),
        };

        let slides = SlideImporter::new(&package, &styles)
            .parse(content_xml)
            .expect("slides should parse");

        assert_eq!(slides.len(), 1);
        assert_eq!(slides[0].boxes.len(), 2);
        assert!(matches!(
            &slides[0].boxes[0].kind,
            RenderBoxKind::Text(block)
                if block.vertical_alignment == TextVerticalAlignment::Top
        ));
        assert!(matches!(
            &slides[0].boxes[1].kind,
            RenderBoxKind::Text(block)
                if block.vertical_alignment == TextVerticalAlignment::Center
        ));
    }

    #[test]
    fn resolves_presentation_style_alignment_from_styles_xml_parent() {
        let styles_xml = r#"
            <office:document-styles>
                <office:styles>
                    <style:style style:name="Default-title" style:family="presentation">
                        <style:paragraph-properties fo:text-align="center"/>
                    </style:style>
                </office:styles>
            </office:document-styles>
        "#;
        let content_xml = r#"
            <office:document-content>
                <office:automatic-styles>
                    <style:style style:name="pr1" style:family="presentation"
                        style:parent-style-name="Default-title"/>
                </office:automatic-styles>
                <office:body>
                    <office:presentation>
                        <draw:page draw:name="page1">
                            <draw:frame presentation:style-name="pr1"
                                svg:x="0cm" svg:y="0cm" svg:width="10cm" svg:height="2cm">
                                <draw:text-box><text:p>Inherited center</text:p></draw:text-box>
                            </draw:frame>
                        </draw:page>
                    </office:presentation>
                </office:body>
            </office:document-content>
        "#;
        let parts = OdpDocumentParts {
            content_xml: content_xml.to_owned(),
            styles_xml: styles_xml.to_owned(),
        };
        let styles = StyleContext::from_parts(&parts).expect("style context should parse");
        let package = empty_package();

        let slides = SlideImporter::new(&package, &styles)
            .parse(content_xml)
            .expect("slides should parse");

        assert!(matches!(
            &slides[0].boxes[0].kind,
            RenderBoxKind::Text(block) if block.alignment == TextAlignment::Center
        ));
    }

    #[test]
    fn resolves_frame_text_style_from_styles_xml_parent() {
        let styles_xml = r##"
            <office:document-styles>
                <office:styles>
                    <style:style style:name="Default-title" style:family="presentation">
                        <style:text-properties fo:font-size="44pt" fo:color="#c9211e"
                            fo:font-weight="bold"/>
                    </style:style>
                </office:styles>
            </office:document-styles>
        "##;
        let content_xml = r#"
            <office:document-content>
                <office:automatic-styles>
                    <style:style style:name="pr1" style:family="presentation"
                        style:parent-style-name="Default-title"/>
                </office:automatic-styles>
                <office:body>
                    <office:presentation>
                        <draw:page draw:name="page1">
                            <draw:frame presentation:style-name="pr1"
                                svg:x="0cm" svg:y="0cm" svg:width="10cm" svg:height="2cm">
                                <draw:text-box><text:p>Inherited title style</text:p></draw:text-box>
                            </draw:frame>
                        </draw:page>
                    </office:presentation>
                </office:body>
            </office:document-content>
        "#;
        let parts = OdpDocumentParts {
            content_xml: content_xml.to_owned(),
            styles_xml: styles_xml.to_owned(),
        };
        let styles = StyleContext::from_parts(&parts).expect("style context should parse");
        let package = empty_package();

        let slides = SlideImporter::new(&package, &styles)
            .parse(content_xml)
            .expect("slides should parse");

        let RenderBoxKind::Text(block) = &slides[0].boxes[0].kind else {
            panic!("expected text box");
        };
        let style = &block.runs[0].style;
        assert_close(style.font_size, 44.0 * PX_PER_IN / 72.0);
        assert_eq!(style.color, Color32::from_rgb(0xc9, 0x21, 0x1e));
        assert!(style.bold);
    }

    #[test]
    fn draw_text_style_does_not_reset_presentation_text_style() {
        let styles_xml = r##"
            <office:document-styles>
                <office:styles>
                    <style:style style:name="Default-subtitle" style:family="presentation">
                        <style:text-properties fo:font-size="32pt" fo:color="#c9211e"/>
                    </style:style>
                </office:styles>
            </office:document-styles>
        "##;
        let content_xml = r##"
            <office:document-content>
                <office:automatic-styles>
                    <style:style style:name="pr2" style:family="presentation"
                        style:parent-style-name="Default-subtitle"/>
                    <style:style style:name="P1" style:family="paragraph">
                        <loext:graphic-properties draw:fill-color="#ffffff"/>
                    </style:style>
                </office:automatic-styles>
                <office:body>
                    <office:presentation>
                        <draw:page draw:name="page1">
                            <draw:frame presentation:style-name="pr2" draw:text-style-name="P1"
                                svg:x="0cm" svg:y="0cm" svg:width="10cm" svg:height="2cm">
                                <draw:text-box><text:p>Subtitle</text:p></draw:text-box>
                            </draw:frame>
                        </draw:page>
                    </office:presentation>
                </office:body>
            </office:document-content>
        "##;
        let parts = OdpDocumentParts {
            content_xml: content_xml.to_owned(),
            styles_xml: styles_xml.to_owned(),
        };
        let styles = StyleContext::from_parts(&parts).expect("style context should parse");
        let package = empty_package();

        let slides = SlideImporter::new(&package, &styles)
            .parse(content_xml)
            .expect("slides should parse");

        let RenderBoxKind::Text(block) = &slides[0].boxes[0].kind else {
            panic!("expected text box");
        };
        let style = &block.runs[0].style;
        assert_close(style.font_size, 32.0 * PX_PER_IN / 72.0);
        assert_eq!(style.color, Color32::from_rgb(0xc9, 0x21, 0x1e));
    }

    #[test]
    fn paragraph_and_span_styles_overlay_frame_text_style() {
        let styles_xml = r##"
            <office:document-styles>
                <office:styles>
                    <style:style style:name="FrameStyle" style:family="presentation">
                        <style:text-properties fo:font-size="32pt" fo:color="#c9211e"/>
                    </style:style>
                </office:styles>
            </office:document-styles>
        "##;
        let content_xml = r#"
            <office:document-content>
                <office:automatic-styles>
                    <style:style style:name="ParagraphBold" style:family="paragraph">
                        <style:text-properties fo:font-weight="bold"/>
                    </style:style>
                    <style:style style:name="SpanItalic" style:family="text">
                        <style:text-properties fo:font-style="italic"/>
                    </style:style>
                </office:automatic-styles>
                <office:body>
                    <office:presentation>
                        <draw:page draw:name="page1">
                            <draw:frame presentation:style-name="FrameStyle"
                                svg:x="0cm" svg:y="0cm" svg:width="10cm" svg:height="2cm">
                                <draw:text-box>
                                    <text:p text:style-name="ParagraphBold">
                                        <text:span text:style-name="SpanItalic">Styled text</text:span>
                                    </text:p>
                                </draw:text-box>
                            </draw:frame>
                        </draw:page>
                    </office:presentation>
                </office:body>
            </office:document-content>
        "#;
        let parts = OdpDocumentParts {
            content_xml: content_xml.to_owned(),
            styles_xml: styles_xml.to_owned(),
        };
        let styles = StyleContext::from_parts(&parts).expect("style context should parse");
        let package = empty_package();

        let slides = SlideImporter::new(&package, &styles)
            .parse(content_xml)
            .expect("slides should parse");

        let RenderBoxKind::Text(block) = &slides[0].boxes[0].kind else {
            panic!("expected text box");
        };
        let style = &block.runs[0].style;
        assert_close(style.font_size, 32.0 * PX_PER_IN / 72.0);
        assert_eq!(style.color, Color32::from_rgb(0xc9, 0x21, 0x1e));
        assert!(style.bold);
        assert!(style.italic);
    }

    #[test]
    fn missing_image_is_a_load_error() {
        let content_xml = r#"
            <office:document-content>
                <office:body>
                    <office:presentation>
                        <draw:page draw:name="page1">
                            <draw:frame svg:x="0cm" svg:y="0cm" svg:width="4cm" svg:height="3cm">
                                <draw:image xlink:href="Pictures/missing.jpg"/>
                            </draw:frame>
                        </draw:page>
                    </office:presentation>
                </office:body>
            </office:document-content>
        "#;
        let package = empty_package();
        let styles = StyleContext {
            page_size: DEFAULT_SLIDE_SIZE,
            text_styles: HashMap::new(),
            paragraph_alignments: HashMap::new(),
            text_vertical_alignments: HashMap::new(),
        };
        let error = SlideImporter::new(&package, &styles)
            .parse(content_xml)
            .expect_err("missing images should stop import");

        assert!(
            matches!(error, OdpLoadError::InvalidPackage(message) if message.contains("Pictures/missing.jpg"))
        );
    }

    #[test]
    fn load_odp_reports_missing_content_entry() {
        let path = write_test_zip(&[
            ("mimetype", ODP_MIME_TYPE.as_bytes()),
            ("META-INF/manifest.xml", b"<manifest/>"),
        ]);
        let error = load_odp(&path).expect_err("content.xml must be required");
        fs::remove_file(path).ok();

        assert!(matches!(error, OdpLoadError::MissingEntry("content.xml")));
    }

    #[test]
    fn load_odp_reports_missing_manifest_entry() {
        let path = write_test_zip(&[("mimetype", ODP_MIME_TYPE.as_bytes())]);
        let error = load_odp(&path).expect_err("manifest.xml must be required");
        fs::remove_file(path).ok();

        assert!(matches!(
            error,
            OdpLoadError::InvalidPackage(message) if message.contains("META-INF/manifest.xml")
        ));
    }

    #[test]
    fn load_odp_rejects_content_without_presentation_body() {
        let path = write_test_zip(&[
            ("mimetype", ODP_MIME_TYPE.as_bytes()),
            ("META-INF/manifest.xml", b"<manifest/>"),
            ("content.xml", b"<office:document-content/>"),
        ]);
        let error = load_odp(&path).expect_err("presentation body must be required");
        fs::remove_file(path).ok();

        assert!(
            matches!(error, OdpLoadError::InvalidPackage(message) if message.contains("office:presentation"))
        );
    }

    #[test]
    fn zip_reader_rejects_unsafe_package_paths() {
        let zip = stored_zip(&[("../content.xml", b"bad")]);
        let error = read_central_directory(&zip).expect_err("path traversal must be rejected");

        assert!(
            matches!(error, OdpLoadError::InvalidPackage(message) if message.contains("unsafe package path"))
        );
    }

    #[test]
    fn zip_reader_extracts_stored_entries() {
        let zip = stored_zip(&[
            ("mimetype", ODP_MIME_TYPE.as_bytes()),
            ("content.xml", b"<office:presentation/>"),
        ]);
        let entries = read_central_directory(&zip).expect("central directory should parse");
        let package = ZipPackage {
            path: PathBuf::from("memory.odp"),
            bytes: zip,
            entries,
        };

        assert_eq!(
            package
                .entry_text("mimetype")
                .expect("mimetype should load"),
            ODP_MIME_TYPE
        );
        assert_eq!(
            package
                .entry_text("content.xml")
                .expect("content should load"),
            "<office:presentation/>"
        );
    }

    fn empty_package() -> ZipPackage {
        ZipPackage {
            path: PathBuf::from("empty.odp"),
            bytes: Vec::new(),
            entries: HashMap::new(),
        }
    }

    fn write_test_zip(entries: &[(&str, &[u8])]) -> PathBuf {
        let mut path = std::env::temp_dir();
        let suffix = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("system time should be after epoch")
            .as_nanos();
        path.push(format!("lro_slides_loader_test_{suffix}.odp"));
        fs::write(&path, stored_zip(entries)).expect("test ZIP should be written");
        path
    }

    fn stored_zip(entries: &[(&str, &[u8])]) -> Vec<u8> {
        let mut zip = Vec::new();
        let mut central_records = Vec::new();

        for (name, data) in entries {
            let local_header_offset = zip.len() as u32;
            push_u32(&mut zip, 0x0403_4b50);
            push_u16(&mut zip, 20);
            push_u16(&mut zip, 0);
            push_u16(&mut zip, 0);
            push_u16(&mut zip, 0);
            push_u16(&mut zip, 0);
            push_u32(&mut zip, 0);
            push_u32(&mut zip, data.len() as u32);
            push_u32(&mut zip, data.len() as u32);
            push_u16(&mut zip, name.len() as u16);
            push_u16(&mut zip, 0);
            zip.extend_from_slice(name.as_bytes());
            zip.extend_from_slice(data);

            let mut central = Vec::new();
            push_u32(&mut central, 0x0201_4b50);
            push_u16(&mut central, 20);
            push_u16(&mut central, 20);
            push_u16(&mut central, 0);
            push_u16(&mut central, 0);
            push_u16(&mut central, 0);
            push_u16(&mut central, 0);
            push_u32(&mut central, 0);
            push_u32(&mut central, data.len() as u32);
            push_u32(&mut central, data.len() as u32);
            push_u16(&mut central, name.len() as u16);
            push_u16(&mut central, 0);
            push_u16(&mut central, 0);
            push_u16(&mut central, 0);
            push_u16(&mut central, 0);
            push_u32(&mut central, 0);
            push_u32(&mut central, local_header_offset);
            central.extend_from_slice(name.as_bytes());
            central_records.push(central);
        }

        let central_dir_offset = zip.len() as u32;
        for record in &central_records {
            zip.extend_from_slice(record);
        }
        let central_dir_size = zip.len() as u32 - central_dir_offset;

        push_u32(&mut zip, 0x0605_4b50);
        push_u16(&mut zip, 0);
        push_u16(&mut zip, 0);
        push_u16(&mut zip, entries.len() as u16);
        push_u16(&mut zip, entries.len() as u16);
        push_u32(&mut zip, central_dir_size);
        push_u32(&mut zip, central_dir_offset);
        push_u16(&mut zip, 0);

        zip
    }

    fn push_u16(output: &mut Vec<u8>, value: u16) {
        output.extend_from_slice(&value.to_le_bytes());
    }

    fn push_u32(output: &mut Vec<u8>, value: u32) {
        output.extend_from_slice(&value.to_le_bytes());
    }

    fn assert_close(actual: f32, expected: f32) {
        assert!(
            (actual - expected).abs() <= 0.01,
            "expected {actual} to be close to {expected}"
        );
    }
}
