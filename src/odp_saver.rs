use std::{
    fmt, fs,
    io::{self, Write},
    path::Path,
};

use egui::{Color32, Vec2};
use image::{ColorType, ImageEncoder, codecs::png::PngEncoder};
use rich_canvas::{
    BoxStrokeKind, BoxStyle, RenderBox, RenderBoxKind, RichCanvas, TextAlignment, TextRun,
    TextStyle, TextVerticalAlignment,
};

const ODP_MIME_TYPE: &str = "application/vnd.oasis.opendocument.presentation";
const PX_PER_CM: f32 = 1280.0 / 28.0;
const PX_PER_IN: f32 = PX_PER_CM * 2.54;

#[derive(Debug)]
pub enum OdpSaveError {
    Io(io::Error),
    Image(image::ImageError),
    InvalidDocument(String),
}

impl fmt::Display for OdpSaveError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Io(error) => write!(formatter, "{error}"),
            Self::Image(error) => write!(formatter, "{error}"),
            Self::InvalidDocument(message) => write!(formatter, "{message}"),
        }
    }
}

impl From<io::Error> for OdpSaveError {
    fn from(error: io::Error) -> Self {
        Self::Io(error)
    }
}

impl From<image::ImageError> for OdpSaveError {
    fn from(error: image::ImageError) -> Self {
        Self::Image(error)
    }
}

pub fn save_odp(path: &Path, slides: &[RichCanvas]) -> Result<(), OdpSaveError> {
    if slides.is_empty() {
        return Err(OdpSaveError::InvalidDocument(
            "cannot save an ODP with no slides".to_owned(),
        ));
    }

    let package = OdpPackage::from_slides(slides)?;
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent)?;
    }
    fs::write(path, package.into_zip_bytes())
        .map_err(|error| OdpSaveError::Io(io::Error::new(error.kind(), error.to_string())))
}

struct OdpPackage {
    slide_size: Vec2,
    image_entries: Vec<SavedImageEntry>,
    content_xml: String,
}

impl OdpPackage {
    fn from_slides(slides: &[RichCanvas]) -> Result<Self, OdpSaveError> {
        let slide_size = slides
            .first()
            .map(|slide| slide.page.size)
            .unwrap_or(Vec2::new(1280.0, 720.0));
        let mut builder = ContentBuilder::new(slide_size);
        let content_xml = builder.content_xml(slides)?;

        Ok(Self {
            slide_size,
            image_entries: builder.image_entries,
            content_xml,
        })
    }

    fn into_zip_bytes(self) -> Vec<u8> {
        let mut zip = StoredZipWriter::new();
        zip.add("mimetype", ODP_MIME_TYPE.as_bytes());
        zip.add("META-INF/manifest.xml", self.manifest_xml().as_bytes());
        zip.add("content.xml", self.content_xml.as_bytes());
        zip.add("styles.xml", self.styles_xml().as_bytes());
        zip.add("meta.xml", self.meta_xml().as_bytes());
        for image in &self.image_entries {
            zip.add(&image.path, &image.bytes);
        }
        zip.finish()
    }

    fn manifest_xml(&self) -> String {
        let mut xml = String::from(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.presentation" manifest:full-path="/"/>
  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="content.xml"/>
  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="styles.xml"/>
  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="meta.xml"/>
"#,
        );
        for image in &self.image_entries {
            xml.push_str(&format!(
                r#"  <manifest:file-entry manifest:media-type="image/png" manifest:full-path="{}"/>
"#,
                xml_escape(&image.path)
            ));
        }
        xml.push_str("</manifest:manifest>\n");
        xml
    }

    fn styles_xml(&self) -> String {
        let width = length_cm(self.slide_size.x);
        let height = length_cm(self.slide_size.y);
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  office:version="1.0">
  <office:styles/>
  <office:automatic-styles>
    <style:page-layout style:name="pm1">
      <style:page-layout-properties fo:page-width="{width}" fo:page-height="{height}"/>
    </style:page-layout>
  </office:automatic-styles>
  <office:master-styles>
    <style:master-page style:name="Default" style:page-layout-name="pm1"/>
  </office:master-styles>
</office:document-styles>
"#
        )
    }

    fn meta_xml(&self) -> &'static str {
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
  office:version="1.0">
  <office:meta>
    <meta:generator>LibeRustOffice Slides</meta:generator>
  </office:meta>
</office:document-meta>
"#
    }
}

struct ContentBuilder {
    text_styles: Vec<SavedTextStyle>,
    graphic_styles: Vec<SavedGraphicStyle>,
    image_entries: Vec<SavedImageEntry>,
}

impl ContentBuilder {
    fn new(_slide_size: Vec2) -> Self {
        Self {
            text_styles: Vec::new(),
            graphic_styles: Vec::new(),
            image_entries: Vec::new(),
        }
    }

    fn content_xml(&mut self, slides: &[RichCanvas]) -> Result<String, OdpSaveError> {
        let mut pages = String::new();
        for (slide_index, slide) in slides.iter().enumerate() {
            pages.push_str(&format!(
                r#"      <draw:page draw:name="Slide {}" draw:master-page-name="Default">
"#,
                slide_index + 1
            ));

            let mut boxes = slide.boxes.iter().collect::<Vec<_>>();
            boxes.sort_by_key(|render_box| render_box.z_index);
            for render_box in boxes {
                self.push_render_box(&mut pages, render_box)?;
            }

            pages.push_str("      </draw:page>\n");
        }

        Ok(format!(
            r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
  xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
  xmlns:xlink="http://www.w3.org/1999/xlink"
  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
  office:version="1.0">
  <office:automatic-styles>
{}  </office:automatic-styles>
  <office:body>
    <office:presentation>
{pages}    </office:presentation>
  </office:body>
</office:document-content>
"#,
            self.automatic_styles_xml()
        ))
    }

    fn push_render_box(
        &mut self,
        xml: &mut String,
        render_box: &RenderBox,
    ) -> Result<(), OdpSaveError> {
        if !render_box.visible {
            return Ok(());
        }

        match &render_box.kind {
            RenderBoxKind::Text(block) => {
                let graphic_style_name =
                    self.graphic_style_name(&render_box.style, block.vertical_alignment);
                let frame_size = render_box.authored_size.unwrap_or(render_box.size);
                xml.push_str(&format!(
                    r#"        <draw:frame draw:style-name="{graphic_style_name}" svg:x="{}" svg:y="{}" svg:width="{}" svg:height="{}">
          <draw:text-box>
            <text:p text:style-name="{}">"#,
                    length_cm(render_box.position.x),
                    length_cm(render_box.position.y),
                    length_cm(frame_size.x * render_box.scale.x),
                    length_cm(frame_size.y * render_box.scale.y),
                    self.paragraph_style_name(block.alignment)
                ));
                for run in &block.runs {
                    self.push_text_run(xml, run);
                }
                xml.push_str(
                    r#"</text:p>
          </draw:text-box>
        </draw:frame>
"#,
                );
            }
            RenderBoxKind::Image(block) => {
                let entry_path = format!("Pictures/image-{}.png", self.image_entries.len() + 1);
                let bytes = encode_color_image_png(block)?;
                self.image_entries.push(SavedImageEntry {
                    path: entry_path.clone(),
                    bytes,
                });
                xml.push_str(&format!(
                    r#"        <draw:frame svg:x="{}" svg:y="{}" svg:width="{}" svg:height="{}">
          <draw:image xlink:href="{}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>
        </draw:frame>
"#,
                    length_cm(render_box.position.x),
                    length_cm(render_box.position.y),
                    length_cm(render_box.size.x * render_box.scale.x),
                    length_cm(render_box.size.y * render_box.scale.y),
                    xml_escape(&entry_path)
                ));
            }
            RenderBoxKind::Table(_) | RenderBoxKind::Container => {}
        }

        Ok(())
    }

    fn push_text_run(&mut self, xml: &mut String, run: &TextRun) {
        for (index, part) in run.text.split('\n').enumerate() {
            if index > 0 {
                xml.push_str("</text:p>\n            <text:p>");
            }
            if part.is_empty() {
                continue;
            }
            let style_name = self.text_style_name(&run.style);
            xml.push_str(&format!(
                r#"<text:span text:style-name="{style_name}">{}</text:span>"#,
                xml_escape(part)
            ));
        }
    }

    fn text_style_name(&mut self, style: &TextStyle) -> String {
        if let Some(existing) = self
            .text_styles
            .iter()
            .find(|saved| saved.style == *style)
            .map(|saved| saved.name.clone())
        {
            return existing;
        }

        let name = format!("T{}", self.text_styles.len() + 1);
        self.text_styles.push(SavedTextStyle {
            name: name.clone(),
            style: style.clone(),
        });
        name
    }

    fn graphic_style_name(
        &mut self,
        style: &BoxStyle,
        vertical_alignment: TextVerticalAlignment,
    ) -> String {
        if let Some(existing) = self
            .graphic_styles
            .iter()
            .find(|saved| saved.matches(style, vertical_alignment))
            .map(|saved| saved.name.clone())
        {
            return existing;
        }

        let name = format!("G{}", self.graphic_styles.len() + 1);
        self.graphic_styles.push(SavedGraphicStyle {
            name: name.clone(),
            fill: style.fill,
            stroke: style.stroke,
            stroke_width_millipx: millipx(style.stroke_width),
            stroke_kind: style.stroke_kind,
            vertical_alignment,
        });
        name
    }

    fn paragraph_style_name(&self, alignment: TextAlignment) -> &'static str {
        match alignment {
            TextAlignment::Left => "PLeft",
            TextAlignment::Center => "PCenter",
            TextAlignment::Right => "PRight",
            TextAlignment::Justify => "PJustify",
        }
    }

    fn automatic_styles_xml(&self) -> String {
        let mut xml = String::new();
        for (name, alignment) in [
            ("PLeft", "left"),
            ("PCenter", "center"),
            ("PRight", "right"),
            ("PJustify", "justify"),
        ] {
            xml.push_str(&format!(
                r#"    <style:style style:name="{name}" style:family="paragraph">
      <style:paragraph-properties fo:text-align="{alignment}"/>
    </style:style>
"#
            ));
        }
        for saved in &self.text_styles {
            xml.push_str(&saved.to_xml());
        }
        for saved in &self.graphic_styles {
            xml.push_str(&saved.to_xml());
        }
        xml
    }
}

#[derive(Clone)]
struct SavedTextStyle {
    name: String,
    style: TextStyle,
}

impl SavedTextStyle {
    fn to_xml(&self) -> String {
        let mut attrs = format!(
            r#"fo:font-size="{}" fo:color="{}""#,
            length_pt(self.style.font_size),
            color_hex(self.style.color)
        );
        if self.style.bold {
            attrs.push_str(r#" fo:font-weight="bold""#);
        }
        if self.style.italic {
            attrs.push_str(r#" fo:font-style="italic""#);
        }
        if self.style.underline {
            attrs.push_str(r#" style:text-underline-style="solid""#);
        }
        if let Some(background) = self.style.background_color {
            attrs.push_str(&format!(
                r#" fo:background-color="{}""#,
                color_hex(background)
            ));
        }

        format!(
            r#"    <style:style style:name="{}" style:family="text">
      <style:text-properties {attrs}/>
    </style:style>
"#,
            xml_escape(&self.name)
        )
    }
}

#[derive(Clone)]
struct SavedGraphicStyle {
    name: String,
    fill: Color32,
    stroke: Color32,
    stroke_width_millipx: i32,
    stroke_kind: BoxStrokeKind,
    vertical_alignment: TextVerticalAlignment,
}

impl SavedGraphicStyle {
    fn matches(&self, style: &BoxStyle, vertical_alignment: TextVerticalAlignment) -> bool {
        self.fill == style.fill
            && self.stroke == style.stroke
            && self.stroke_width_millipx == millipx(style.stroke_width)
            && self.stroke_kind == style.stroke_kind
            && self.vertical_alignment == vertical_alignment
    }

    fn to_xml(&self) -> String {
        let fill_attrs = if self.fill == Color32::TRANSPARENT {
            r#"draw:fill="none""#.to_owned()
        } else {
            format!(
                r#"draw:fill="solid" draw:fill-color="{}""#,
                color_hex(self.fill)
            )
        };
        let stroke_attrs = if self.stroke == Color32::TRANSPARENT || self.stroke_width_millipx <= 0
        {
            r#"draw:stroke="none""#.to_owned()
        } else {
            let stroke_type = match self.stroke_kind {
                BoxStrokeKind::Solid => "solid",
                BoxStrokeKind::Dash => "dash",
            };
            format!(
                r#"draw:stroke="{stroke_type}" svg:stroke-color="{}" svg:stroke-width="{}""#,
                color_hex(self.stroke),
                length_cm(self.stroke_width_millipx as f32 / 1000.0)
            )
        };

        format!(
            r#"    <style:style style:name="{}" style:family="graphic">
      <style:graphic-properties {fill_attrs} {stroke_attrs} draw:textarea-vertical-align="{}"/>
    </style:style>
"#,
            xml_escape(&self.name),
            vertical_alignment_value(self.vertical_alignment)
        )
    }
}

struct SavedImageEntry {
    path: String,
    bytes: Vec<u8>,
}

fn encode_color_image_png(block: &rich_canvas::ImageBlock) -> Result<Vec<u8>, OdpSaveError> {
    let mut rgba = Vec::with_capacity(block.color_image.pixels.len() * 4);
    for pixel in &block.color_image.pixels {
        rgba.extend_from_slice(&[pixel.r(), pixel.g(), pixel.b(), pixel.a()]);
    }

    let mut bytes = Vec::new();
    PngEncoder::new(&mut bytes).write_image(
        &rgba,
        block.color_image.size[0] as u32,
        block.color_image.size[1] as u32,
        ColorType::Rgba8.into(),
    )?;
    Ok(bytes)
}

fn length_cm(px: f32) -> String {
    format!("{:.4}cm", px.max(0.0) / PX_PER_CM)
}

fn length_pt(px: f32) -> String {
    format!("{:.2}pt", px / (PX_PER_IN / 72.0))
}

fn millipx(px: f32) -> i32 {
    (px * 1000.0).round() as i32
}

fn vertical_alignment_value(alignment: TextVerticalAlignment) -> &'static str {
    match alignment {
        TextVerticalAlignment::Top => "top",
        TextVerticalAlignment::Center => "middle",
        TextVerticalAlignment::Bottom => "bottom",
    }
}

fn color_hex(color: Color32) -> String {
    format!("#{:02x}{:02x}{:02x}", color.r(), color.g(), color.b())
}

fn xml_escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

struct StoredZipWriter {
    bytes: Vec<u8>,
    entries: Vec<CentralDirectoryEntry>,
}

impl StoredZipWriter {
    fn new() -> Self {
        Self {
            bytes: Vec::new(),
            entries: Vec::new(),
        }
    }

    fn add(&mut self, name: &str, data: &[u8]) {
        let offset = self.bytes.len() as u32;
        let crc = crc32(data);
        write_u32(&mut self.bytes, 0x0403_4b50);
        write_u16(&mut self.bytes, 20);
        write_u16(&mut self.bytes, 0);
        write_u16(&mut self.bytes, 0);
        write_u16(&mut self.bytes, 0);
        write_u16(&mut self.bytes, 0);
        write_u32(&mut self.bytes, crc);
        write_u32(&mut self.bytes, data.len() as u32);
        write_u32(&mut self.bytes, data.len() as u32);
        write_u16(&mut self.bytes, name.len() as u16);
        write_u16(&mut self.bytes, 0);
        self.bytes.extend_from_slice(name.as_bytes());
        self.bytes.extend_from_slice(data);
        self.entries.push(CentralDirectoryEntry {
            name: name.to_owned(),
            crc,
            size: data.len() as u32,
            local_header_offset: offset,
        });
    }

    fn finish(mut self) -> Vec<u8> {
        let central_directory_offset = self.bytes.len() as u32;
        let entries = std::mem::take(&mut self.entries);
        for entry in &entries {
            write_u32(&mut self.bytes, 0x0201_4b50);
            write_u16(&mut self.bytes, 20);
            write_u16(&mut self.bytes, 20);
            write_u16(&mut self.bytes, 0);
            write_u16(&mut self.bytes, 0);
            write_u16(&mut self.bytes, 0);
            write_u16(&mut self.bytes, 0);
            write_u32(&mut self.bytes, entry.crc);
            write_u32(&mut self.bytes, entry.size);
            write_u32(&mut self.bytes, entry.size);
            write_u16(&mut self.bytes, entry.name.len() as u16);
            write_u16(&mut self.bytes, 0);
            write_u16(&mut self.bytes, 0);
            write_u16(&mut self.bytes, 0);
            write_u16(&mut self.bytes, 0);
            write_u32(&mut self.bytes, 0);
            write_u32(&mut self.bytes, entry.local_header_offset);
            self.bytes.extend_from_slice(entry.name.as_bytes());
        }
        let central_directory_size = self.bytes.len() as u32 - central_directory_offset;
        write_u32(&mut self.bytes, 0x0605_4b50);
        write_u16(&mut self.bytes, 0);
        write_u16(&mut self.bytes, 0);
        write_u16(&mut self.bytes, entries.len() as u16);
        write_u16(&mut self.bytes, entries.len() as u16);
        write_u32(&mut self.bytes, central_directory_size);
        write_u32(&mut self.bytes, central_directory_offset);
        write_u16(&mut self.bytes, 0);
        self.bytes
    }
}

struct CentralDirectoryEntry {
    name: String,
    crc: u32,
    size: u32,
    local_header_offset: u32,
}

fn crc32(bytes: &[u8]) -> u32 {
    let mut crc = 0xffff_ffffu32;
    for byte in bytes {
        crc ^= u32::from(*byte);
        for _ in 0..8 {
            let mask = 0u32.wrapping_sub(crc & 1);
            crc = (crc >> 1) ^ (0xedb8_8320 & mask);
        }
    }
    !crc
}

fn write_u16(output: &mut Vec<u8>, value: u16) {
    output.write_all(&value.to_le_bytes()).expect("Vec write");
}

fn write_u32(output: &mut Vec<u8>, value: u32) {
    output.write_all(&value.to_le_bytes()).expect("Vec write");
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::odp_loader;
    use egui::{pos2, vec2};
    use rich_canvas::{LayoutRole, RenderBox, TextStyle};

    #[test]
    fn saved_odp_can_be_loaded_by_current_importer() {
        let mut slide = RichCanvas::new(vec2(1280.0, 720.0));
        let mut text = RenderBox::text(
            1,
            LayoutRole::Absolute,
            vec![TextRun::new("Saved slide", TextStyle::heading())],
        );
        text.position = pos2(80.0, 90.0);
        text.size = vec2(420.0, 120.0);
        slide.push(text);

        let path = std::env::temp_dir().join("liberustoffice_saved_roundtrip.odp");
        save_odp(&path, &[slide]).expect("save succeeds");
        let loaded = odp_loader::load_odp(&path).expect("saved ODP loads");

        assert_eq!(loaded.slides.len(), 1);
        assert_eq!(loaded.slides[0].boxes.len(), 1);
        assert_eq!(
            loaded.slides[0].boxes[0].plain_text().as_deref(),
            Some("Saved slide")
        );
    }

    #[test]
    fn loading_and_saving_default_odp_preserves_text_scheme() {
        let original_path = Path::new(odp_loader::DEFAULT_ODP_PATH);
        let loaded = odp_loader::load_odp(original_path).expect("default ODP loads");
        let saved_path = std::env::temp_dir().join("test_slides_saved.odp");
        let expected_package =
            OdpPackage::from_slides(&loaded.slides).expect("expected package can be built");

        save_odp(&saved_path, &loaded.slides).expect("save succeeds");

        let saved_content_xml =
            stored_zip_entry_text(&saved_path, "content.xml").expect("saved content.xml exists");
        let saved_loaded = odp_loader::load_odp(&saved_path).expect("saved ODP reloads");

        assert_eq!(saved_content_xml, expected_package.content_xml);
        assert_eq!(
            text_scheme(&saved_loaded.slides),
            text_scheme(&loaded.slides)
        );
    }

    #[test]
    fn saving_loaded_default_odp_preserves_slide_page_size() {
        let loaded = odp_loader::load_default_odp().expect("default ODP loads");
        let saved_path = std::env::temp_dir().join("test_slides_page_size_regression.odp");

        save_odp(&saved_path, &loaded.slides).expect("save succeeds");

        let styles_xml =
            stored_zip_entry_text(&saved_path, "styles.xml").expect("saved styles.xml exists");

        assert!(styles_xml.contains(r#"fo:page-width="28.0000cm""#));
        assert!(styles_xml.contains(r#"fo:page-height="15.7500cm""#));
        assert!(!styles_xml.contains(r#"fo:page-width="30.6250cm""#));
        assert!(!styles_xml.contains(r#"fo:page-height="18.3750cm""#));
    }

    #[test]
    fn text_box_fill_and_border_are_saved_as_graphic_style() {
        let mut slide = RichCanvas::new(vec2(1280.0, 720.0));
        let mut text = RenderBox::text(
            1,
            LayoutRole::Absolute,
            vec![TextRun::new("Styled box", TextStyle::body())],
        );
        if let RenderBoxKind::Text(block) = &mut text.kind {
            block.vertical_alignment = TextVerticalAlignment::Bottom;
        }
        text.style.fill = Color32::from_rgb(12, 34, 56);
        text.style.stroke = Color32::from_rgb(200, 120, 30);
        text.style.stroke_width = 2.5;
        slide.push(text);

        let package = OdpPackage::from_slides(&[slide]).expect("package can be built");

        assert!(package.content_xml.contains(r#"draw:style-name="G1""#));
        assert!(
            package
                .content_xml
                .contains(r##"draw:fill-color="#0c2238""##)
        );
        assert!(
            package
                .content_xml
                .contains(r##"svg:stroke-color="#c8781e""##)
        );
        assert!(
            package
                .content_xml
                .contains(r#"svg:stroke-width="0.0547cm""#)
        );
        assert!(
            package
                .content_xml
                .contains(r#"draw:textarea-vertical-align="bottom""#)
        );
    }

    #[test]
    fn crc32_matches_known_value() {
        assert_eq!(crc32(b"123456789"), 0xcbf4_3926);
    }

    #[derive(Debug, PartialEq, Eq)]
    struct TextBoxScheme {
        alignment: TextAlignment,
        vertical_alignment: TextVerticalAlignment,
        runs: Vec<TextRunScheme>,
    }

    #[derive(Debug, PartialEq, Eq)]
    struct TextRunScheme {
        text: String,
        font_size_centipx: i32,
        color: (u8, u8, u8, u8),
        background_color: Option<(u8, u8, u8, u8)>,
        bold: bool,
        italic: bool,
        underline: bool,
    }

    fn text_scheme(slides: &[RichCanvas]) -> Vec<Vec<TextBoxScheme>> {
        slides
            .iter()
            .map(|slide| {
                let mut boxes = slide.boxes.iter().collect::<Vec<_>>();
                boxes.sort_by_key(|render_box| render_box.z_index);
                boxes
                    .into_iter()
                    .filter_map(|render_box| {
                        let RenderBoxKind::Text(block) = &render_box.kind else {
                            return None;
                        };
                        Some(TextBoxScheme {
                            alignment: block.alignment,
                            vertical_alignment: block.vertical_alignment,
                            runs: block.runs.iter().map(text_run_scheme).collect(),
                        })
                    })
                    .collect()
            })
            .collect()
    }

    fn text_run_scheme(run: &TextRun) -> TextRunScheme {
        TextRunScheme {
            text: run.text.clone(),
            font_size_centipx: (run.style.font_size * 100.0).round() as i32,
            color: color_tuple(run.style.color),
            background_color: run.style.background_color.map(color_tuple),
            bold: run.style.bold,
            italic: run.style.italic,
            underline: run.style.underline,
        }
    }

    fn color_tuple(color: Color32) -> (u8, u8, u8, u8) {
        (color.r(), color.g(), color.b(), color.a())
    }

    fn stored_zip_entry_text(path: &Path, entry_name: &str) -> Option<String> {
        let bytes = fs::read(path).ok()?;
        let entry_name_bytes = entry_name.as_bytes();
        let mut offset = 0usize;

        while offset + 30 <= bytes.len() {
            if read_u32_le(&bytes, offset)? != 0x0403_4b50 {
                break;
            }

            let method = read_u16_le(&bytes, offset + 8)?;
            let compressed_size = read_u32_le(&bytes, offset + 18)? as usize;
            let uncompressed_size = read_u32_le(&bytes, offset + 22)? as usize;
            let name_len = read_u16_le(&bytes, offset + 26)? as usize;
            let extra_len = read_u16_le(&bytes, offset + 28)? as usize;
            let name_start = offset + 30;
            let name_end = name_start.checked_add(name_len)?;
            let data_start = name_end.checked_add(extra_len)?;
            let data_end = data_start.checked_add(compressed_size)?;
            if data_end > bytes.len() {
                return None;
            }

            if &bytes[name_start..name_end] == entry_name_bytes {
                if method != 0 || compressed_size != uncompressed_size {
                    return None;
                }
                return String::from_utf8(bytes[data_start..data_end].to_vec()).ok();
            }

            offset = data_end;
        }

        None
    }

    fn read_u16_le(bytes: &[u8], offset: usize) -> Option<u16> {
        Some(u16::from_le_bytes(
            bytes.get(offset..offset + 2)?.try_into().ok()?,
        ))
    }

    fn read_u32_le(bytes: &[u8], offset: usize) -> Option<u32> {
        Some(u32::from_le_bytes(
            bytes.get(offset..offset + 4)?.try_into().ok()?,
        ))
    }
}
