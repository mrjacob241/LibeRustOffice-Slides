use std::{fmt, fs, io, path::Path};

use egui::Color32;
use rich_canvas::{RenderBox, RenderBoxKind, RichCanvas, TextAlignment, TextRun};

const PX_PER_CM: f32 = 1280.0 / 28.0;
const PT_PER_IN: f32 = 72.0;
const CM_PER_IN: f32 = 2.54;
const PT_PER_PX: f32 = PT_PER_IN / (PX_PER_CM * CM_PER_IN);

#[derive(Debug)]
pub enum PdfExportError {
    Io(io::Error),
    InvalidDocument(String),
}

impl fmt::Display for PdfExportError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Io(error) => write!(formatter, "{error}"),
            Self::InvalidDocument(message) => write!(formatter, "{message}"),
        }
    }
}

impl From<io::Error> for PdfExportError {
    fn from(error: io::Error) -> Self {
        Self::Io(error)
    }
}

pub fn export_pdf(path: &Path, slides: &[RichCanvas]) -> Result<(), PdfExportError> {
    let bytes = export_pdf_bytes(slides)?;
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent)?;
    }
    fs::write(path, bytes)?;
    Ok(())
}

fn export_pdf_bytes(slides: &[RichCanvas]) -> Result<Vec<u8>, PdfExportError> {
    if slides.is_empty() {
        return Err(PdfExportError::InvalidDocument(
            "cannot export a PDF with no slides".to_owned(),
        ));
    }

    let mut allocator = ObjectAllocator::new(7);
    let mut pages = Vec::new();
    for slide in slides {
        pages.push(PdfPageBuild::from_slide(&mut allocator, slide));
    }

    let mut pdf = RawPdf::new(allocator.next_id);
    pdf.set_object(1, b"<< /Type /Catalog /Pages 2 0 R >>".to_vec());
    pdf.set_object(3, type1_font("Helvetica"));
    pdf.set_object(4, type1_font("Helvetica-Bold"));
    pdf.set_object(5, type1_font("Helvetica-Oblique"));
    pdf.set_object(6, type1_font("Helvetica-BoldOblique"));

    let page_refs = pages
        .iter()
        .map(|page| format!("{} 0 R", page.page_id))
        .collect::<Vec<_>>()
        .join(" ");
    pdf.set_object(
        2,
        format!(
            "<< /Type /Pages /Kids [ {page_refs} ] /Count {} >>",
            pages.len()
        )
        .into_bytes(),
    );

    for page in pages {
        for image in &page.images {
            pdf.set_object(image.object_id, image_xobject(image));
        }
        pdf.set_object(page.content_id, stream_object(&page.content));
        pdf.set_object(page.page_id, page_object(&page));
    }

    Ok(pdf.finish())
}

fn type1_font(base_font: &str) -> Vec<u8> {
    format!("<< /Type /Font /Subtype /Type1 /BaseFont /{base_font} >>").into_bytes()
}

fn page_object(page: &PdfPageBuild) -> Vec<u8> {
    let xobjects = if page.images.is_empty() {
        String::new()
    } else {
        let entries = page
            .images
            .iter()
            .map(|image| format!("/{} {} 0 R", image.name, image.object_id))
            .collect::<Vec<_>>()
            .join(" ");
        format!("/XObject << {entries} >>")
    };

    format!(
        concat!(
            "<< /Type /Page /Parent 2 0 R ",
            "/MediaBox [ 0 0 {:.4} {:.4} ] ",
            "/Resources << /Font << /F1 3 0 R /F2 4 0 R /F3 5 0 R /F4 6 0 R >> {} >> ",
            "/Contents {} 0 R >>"
        ),
        page.width_pt, page.height_pt, xobjects, page.content_id
    )
    .into_bytes()
}

fn stream_object(bytes: &[u8]) -> Vec<u8> {
    let mut object = format!("<< /Length {} >>\nstream\n", bytes.len()).into_bytes();
    object.extend_from_slice(bytes);
    object.extend_from_slice(b"\nendstream");
    object
}

fn image_xobject(image: &PdfImage) -> Vec<u8> {
    let mut object = format!(
        concat!(
            "<< /Type /XObject /Subtype /Image ",
            "/Width {} /Height {} /ColorSpace /DeviceRGB /BitsPerComponent 8 ",
            "/Length {} >>\nstream\n"
        ),
        image.width_px,
        image.height_px,
        image.rgb.len()
    )
    .into_bytes();
    object.extend_from_slice(&image.rgb);
    object.extend_from_slice(b"\nendstream");
    object
}

struct PdfPageBuild {
    page_id: usize,
    content_id: usize,
    width_pt: f32,
    height_pt: f32,
    images: Vec<PdfImage>,
    content: Vec<u8>,
}

impl PdfPageBuild {
    fn from_slide(allocator: &mut ObjectAllocator, slide: &RichCanvas) -> Self {
        let page_id = allocator.alloc();
        let content_id = allocator.alloc();
        let page_size = slide.page.size;
        let width_pt = px_to_pt(page_size.x);
        let height_pt = px_to_pt(page_size.y);
        let mut content = Vec::new();
        let mut images = Vec::new();

        push_rect_fill(&mut content, 0.0, 0.0, width_pt, height_pt, Color32::WHITE);

        let mut boxes = slide.boxes.iter().collect::<Vec<_>>();
        boxes.sort_by_key(|render_box| render_box.z_index);
        for render_box in boxes {
            if !render_box.visible {
                continue;
            }
            match &render_box.kind {
                RenderBoxKind::Text(block) => {
                    push_text_box(
                        &mut content,
                        render_box,
                        &block.runs,
                        block.alignment,
                        height_pt,
                    );
                }
                RenderBoxKind::Image(block) => {
                    let object_id = allocator.alloc();
                    let name = format!("Im{}", images.len() + 1);
                    push_image_box(&mut content, render_box, &name, height_pt);
                    images.push(PdfImage::from_color_image(
                        object_id,
                        name,
                        block.color_image.size,
                        &block.color_image.pixels,
                    ));
                }
                RenderBoxKind::Table(_) | RenderBoxKind::Container => {}
            }
        }

        Self {
            page_id,
            content_id,
            width_pt,
            height_pt,
            images,
            content,
        }
    }
}

struct PdfImage {
    object_id: usize,
    name: String,
    width_px: usize,
    height_px: usize,
    rgb: Vec<u8>,
}

impl PdfImage {
    fn from_color_image(
        object_id: usize,
        name: String,
        size: [usize; 2],
        pixels: &[Color32],
    ) -> Self {
        let mut rgb = Vec::with_capacity(pixels.len() * 3);
        for pixel in pixels {
            let alpha = pixel.a() as u16;
            let inverse_alpha = 255 - alpha;
            rgb.push(((pixel.r() as u16 * alpha + 255 * inverse_alpha) / 255) as u8);
            rgb.push(((pixel.g() as u16 * alpha + 255 * inverse_alpha) / 255) as u8);
            rgb.push(((pixel.b() as u16 * alpha + 255 * inverse_alpha) / 255) as u8);
        }
        Self {
            object_id,
            name,
            width_px: size[0],
            height_px: size[1],
            rgb,
        }
    }
}

fn push_image_box(content: &mut Vec<u8>, render_box: &RenderBox, image_name: &str, page_h: f32) {
    let x = px_to_pt(render_box.position.x);
    let y = page_h - px_to_pt(render_box.position.y + render_box.size.y * render_box.scale.y);
    let w = px_to_pt(render_box.size.x * render_box.scale.x);
    let h = px_to_pt(render_box.size.y * render_box.scale.y);
    content.extend_from_slice(
        format!(
            "q {:.4} 0 0 {:.4} {:.4} {:.4} cm /{} Do Q\n",
            w, h, x, y, image_name
        )
        .as_bytes(),
    );
}

fn push_text_box(
    content: &mut Vec<u8>,
    render_box: &RenderBox,
    runs: &[TextRun],
    alignment: TextAlignment,
    page_h: f32,
) {
    let box_x = px_to_pt(render_box.position.x);
    let box_y = px_to_pt(render_box.position.y);
    let box_w = px_to_pt(render_box.size.x * render_box.scale.x);
    let padding = render_box.style.padding * PT_PER_PX;
    let max_width = (box_w - padding.x * 2.0).max(12.0);
    let lines = wrap_runs(runs, max_width / PT_PER_PX);
    let mut y = page_h - box_y - padding.y;

    for line in lines {
        let line = merge_adjacent_segments(line);
        let line_width = line
            .iter()
            .map(|segment| px_to_pt(segment.width_px()))
            .sum::<f32>();
        let x_offset = match alignment {
            TextAlignment::Center => ((max_width - line_width) / 2.0).max(0.0),
            TextAlignment::Right => (max_width - line_width).max(0.0),
            TextAlignment::Left | TextAlignment::Justify => 0.0,
        };
        let mut x = box_x + padding.x + x_offset;
        let line_height = line
            .iter()
            .map(|segment| px_to_pt(segment.style.font_size) * 1.2)
            .fold(12.0_f32, f32::max);
        y -= line_height;

        for segment in line {
            if segment.text.is_empty() {
                continue;
            }
            let font_size = px_to_pt(segment.style.font_size);
            let segment_width = px_to_pt(segment.width_px());
            if let Some(background) = segment.style.background_color {
                let highlight_y = y - font_size * 0.08;
                let highlight_h = font_size * 1.04;
                push_rect_fill(
                    content,
                    x,
                    highlight_y,
                    segment_width,
                    highlight_h,
                    background,
                );
            }
            let escaped = pdf_escape_text(&segment.text);
            let font = font_name(segment.style.bold, segment.style.italic);
            content.extend_from_slice(
                format!(
                    "BT /{} {:.4} Tf {:.4} {:.4} {:.4} rg {:.4} {:.4} Td ({}) Tj ET\n",
                    font,
                    font_size,
                    color_unit(segment.style.color.r()),
                    color_unit(segment.style.color.g()),
                    color_unit(segment.style.color.b()),
                    x,
                    y,
                    escaped
                )
                .as_bytes(),
            );
            if segment.style.underline {
                let underline_y = y - font_size * 0.15;
                push_stroke_line(
                    content,
                    x,
                    underline_y,
                    x + segment_width,
                    underline_y,
                    segment.style.color,
                );
            }
            x += px_to_pt(segment.width_px());
        }
    }
}

fn wrap_runs(runs: &[TextRun], max_width_px: f32) -> Vec<Vec<TextSegment>> {
    let mut lines = Vec::new();
    let mut line = Vec::new();
    let mut line_width = 0.0;

    for run in runs {
        for token in text_tokens(&run.text) {
            if token == "\n" {
                lines.push(std::mem::take(&mut line));
                line_width = 0.0;
                continue;
            }
            let segment = TextSegment {
                text: token,
                style: run.style.clone(),
            };
            let width = segment.width_px();
            if !line.is_empty() && line_width + width > max_width_px {
                lines.push(std::mem::take(&mut line));
                line_width = 0.0;
            }
            line_width += width;
            line.push(segment);
        }
    }

    if !line.is_empty() || lines.is_empty() {
        lines.push(line);
    }
    lines
}

fn merge_adjacent_segments(segments: Vec<TextSegment>) -> Vec<TextSegment> {
    let mut merged: Vec<TextSegment> = Vec::new();
    for segment in segments {
        if let Some(last) = merged.last_mut() {
            if last.style == segment.style {
                last.text.push_str(&segment.text);
                continue;
            }
        }
        merged.push(segment);
    }
    merged
}

fn text_tokens(text: &str) -> Vec<String> {
    let mut tokens = Vec::new();
    let mut current = String::new();
    for ch in text.chars() {
        if ch == '\n' {
            if !current.is_empty() {
                tokens.push(std::mem::take(&mut current));
            }
            tokens.push("\n".to_owned());
        } else {
            current.push(ch);
            if ch.is_whitespace() {
                tokens.push(std::mem::take(&mut current));
            }
        }
    }
    if !current.is_empty() {
        tokens.push(current);
    }
    tokens
}

struct TextSegment {
    text: String,
    style: rich_canvas::TextStyle,
}

impl TextSegment {
    fn width_px(&self) -> f32 {
        text_width_px(&self.text, self.style.font_size)
    }
}

fn text_width_px(text: &str, font_size: f32) -> f32 {
    text.chars()
        .map(|ch| {
            let factor = match ch {
                ' ' | '\t' => 0.28,
                'i' | 'l' | 'I' | '!' | '|' | '.' | ',' | ':' | ';' => 0.25,
                'f' | 'j' | 'r' | 't' => 0.36,
                'm' | 'w' | 'M' | 'W' => 0.82,
                'A'..='Z' => 0.66,
                '0'..='9' => 0.55,
                _ => 0.52,
            };
            font_size * factor
        })
        .sum()
}

fn push_rect_fill(content: &mut Vec<u8>, x: f32, y: f32, w: f32, h: f32, color: Color32) {
    content.extend_from_slice(
        format!(
            "{:.4} {:.4} {:.4} rg {:.4} {:.4} {:.4} {:.4} re f\n",
            color_unit(color.r()),
            color_unit(color.g()),
            color_unit(color.b()),
            x,
            y,
            w,
            h
        )
        .as_bytes(),
    );
}

fn push_stroke_line(content: &mut Vec<u8>, x1: f32, y1: f32, x2: f32, y2: f32, color: Color32) {
    content.extend_from_slice(
        format!(
            "{:.4} {:.4} {:.4} RG 0.75 w {:.4} {:.4} m {:.4} {:.4} l S\n",
            color_unit(color.r()),
            color_unit(color.g()),
            color_unit(color.b()),
            x1,
            y1,
            x2,
            y2
        )
        .as_bytes(),
    );
}

fn font_name(bold: bool, italic: bool) -> &'static str {
    match (bold, italic) {
        (true, true) => "F4",
        (true, false) => "F2",
        (false, true) => "F3",
        (false, false) => "F1",
    }
}

fn pdf_escape_text(text: &str) -> String {
    let mut escaped = String::new();
    for ch in text.chars() {
        match ch {
            '(' => escaped.push_str("\\("),
            ')' => escaped.push_str("\\)"),
            '\\' => escaped.push_str("\\\\"),
            '\r' | '\n' => {}
            ch if ch.is_ascii() && !ch.is_control() => escaped.push(ch),
            _ => escaped.push('?'),
        }
    }
    escaped
}

fn color_unit(value: u8) -> f32 {
    value as f32 / 255.0
}

fn px_to_pt(px: f32) -> f32 {
    px * PT_PER_PX
}

struct ObjectAllocator {
    next_id: usize,
}

impl ObjectAllocator {
    fn new(next_id: usize) -> Self {
        Self { next_id }
    }

    fn alloc(&mut self) -> usize {
        let id = self.next_id;
        self.next_id += 1;
        id
    }
}

struct RawPdf {
    objects: Vec<Option<Vec<u8>>>,
}

impl RawPdf {
    fn new(next_id: usize) -> Self {
        Self {
            objects: vec![None; next_id],
        }
    }

    fn set_object(&mut self, id: usize, bytes: Vec<u8>) {
        if id >= self.objects.len() {
            self.objects.resize(id + 1, None);
        }
        self.objects[id] = Some(bytes);
    }

    fn finish(self) -> Vec<u8> {
        let mut bytes = b"%PDF-1.4\n%\x80\x80\x80\x80\n".to_vec();
        let mut offsets = vec![0usize; self.objects.len()];

        for id in 1..self.objects.len() {
            if let Some(object) = &self.objects[id] {
                offsets[id] = bytes.len();
                bytes.extend_from_slice(format!("{id} 0 obj\n").as_bytes());
                bytes.extend_from_slice(object);
                bytes.extend_from_slice(b"\nendobj\n");
            }
        }

        let xref_start = bytes.len();
        bytes.extend_from_slice(format!("xref\n0 {}\n", self.objects.len()).as_bytes());
        bytes.extend_from_slice(b"0000000000 65535 f \n");
        for offset in offsets.iter().skip(1) {
            bytes.extend_from_slice(format!("{offset:010} 00000 n \n").as_bytes());
        }
        bytes.extend_from_slice(
            format!(
                "trailer\n<< /Size {} /Root 1 0 R >>\nstartxref\n{xref_start}\n%%EOF\n",
                self.objects.len()
            )
            .as_bytes(),
        );
        bytes
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::odp_loader;
    use egui::{pos2, vec2};
    use rich_canvas::{LayoutRole, RenderBox, TextStyle};

    #[test]
    fn exports_pdf_with_one_page_per_slide() {
        let slides = vec![sample_slide(), sample_slide()];
        let bytes = export_pdf_bytes(&slides).expect("PDF export should succeed");
        let text = String::from_utf8_lossy(&bytes);

        assert!(bytes.starts_with(b"%PDF-1.4"));
        assert_eq!(text.matches("/Type /Page ").count(), 2);
        assert!(text.contains("xref"));
        assert!(text.contains("trailer"));
    }

    #[test]
    fn exports_default_landscape_slide_size() {
        let loaded = odp_loader::load_default_odp().expect("default ODP loads");
        let bytes = export_pdf_bytes(&loaded.slides).expect("PDF export should succeed");
        let text = String::from_utf8_lossy(&bytes);

        assert!(text.contains("/MediaBox [ 0 0 793.7008 446.4567 ]"));
    }

    #[test]
    fn text_width_estimate_does_not_expand_spaces_like_letters() {
        let font_size = 30.0;

        assert!(text_width_px("This is a ", font_size) < text_width_px("Thisisaxx", font_size));
        assert!(text_width_px("Test Slides", 44.0) < 270.0);
    }

    #[test]
    fn adjacent_same_style_segments_are_merged_for_pdf_spacing() {
        let style = TextStyle::body();
        let merged = merge_adjacent_segments(vec![
            TextSegment {
                text: "Sample ".to_owned(),
                style: style.clone(),
            },
            TextSegment {
                text: "Slide ".to_owned(),
                style: style.clone(),
            },
            TextSegment {
                text: "1".to_owned(),
                style,
            },
        ]);

        assert_eq!(merged.len(), 1);
        assert_eq!(merged[0].text, "Sample Slide 1");
    }

    #[test]
    fn highlighted_text_writes_background_rectangle() {
        let mut slide = RichCanvas::new(vec2(1280.0, 720.0));
        let mut style = TextStyle::body();
        style.background_color = Some(Color32::YELLOW);
        let mut text = RenderBox::text(
            1,
            LayoutRole::Absolute,
            vec![TextRun::new("Highlight.", style)],
        );
        text.position = pos2(80.0, 80.0);
        slide.push(text);

        let bytes = export_pdf_bytes(&[slide]).expect("PDF export should succeed");
        let pdf = String::from_utf8_lossy(&bytes);

        assert!(pdf.contains("1.0000 1.0000 0.0000 rg"));
        assert!(pdf.contains("(Highlight.) Tj"));
    }

    #[test]
    fn default_odp_can_be_written_as_pdf_file() {
        let loaded = odp_loader::load_default_odp().expect("default ODP loads");
        let path = std::env::temp_dir().join("liberustoffice_default_export.pdf");

        export_pdf(&path, &loaded.slides).expect("PDF export should succeed");

        let bytes = std::fs::read(path).expect("PDF file should exist");
        assert!(bytes.starts_with(b"%PDF-1.4"));
    }

    fn sample_slide() -> RichCanvas {
        let mut slide = RichCanvas::new(vec2(1280.0, 720.0));
        let mut title = RenderBox::text(
            1,
            LayoutRole::Absolute,
            vec![TextRun::new("PDF Export", TextStyle::heading())],
        );
        title.position = pos2(80.0, 70.0);
        title.size = vec2(600.0, 100.0);
        slide.push(title);
        slide
    }
}
