use std::{fmt, fs, io, path::Path};

use egui::{Color32, Pos2};
use rich_canvas::{
    BoxStrokeKind, RenderBox, RenderBoxKind, RichCanvas, TextAlignment, TextRun,
    TextVerticalAlignment,
};

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
        for annotation in &page.link_annotations {
            pdf.set_object(annotation.object_id, link_annotation_object(annotation));
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

    let annots = if page.link_annotations.is_empty() {
        String::new()
    } else {
        let refs = page
            .link_annotations
            .iter()
            .map(|annotation| format!("{} 0 R", annotation.object_id))
            .collect::<Vec<_>>()
            .join(" ");
        format!("/Annots [ {refs} ] ")
    };

    format!(
        concat!(
            "<< /Type /Page /Parent 2 0 R ",
            "/MediaBox [ 0 0 {:.4} {:.4} ] ",
            "/Resources << /Font << /F1 3 0 R /F2 4 0 R /F3 5 0 R /F4 6 0 R >> {} >> ",
            "{}/Contents {} 0 R >>"
        ),
        page.width_pt, page.height_pt, xobjects, annots, page.content_id
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

fn link_annotation_object(annotation: &PdfLinkAnnotation) -> Vec<u8> {
    format!(
        concat!(
            "<< /Type /Annot /Subtype /Link ",
            "/Rect [ {:.4} {:.4} {:.4} {:.4} ] ",
            "/Border [ 0 0 0 ] ",
            "/A << /S /URI /URI ({}) >> >>"
        ),
        annotation.rect.x,
        annotation.rect.y,
        annotation.rect.x + annotation.rect.w,
        annotation.rect.y + annotation.rect.h,
        pdf_escape_string(&annotation.url)
    )
    .into_bytes()
}

struct PdfPageBuild {
    page_id: usize,
    content_id: usize,
    width_pt: f32,
    height_pt: f32,
    images: Vec<PdfImage>,
    link_annotations: Vec<PdfLinkAnnotation>,
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
        let mut link_annotations = Vec::new();

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
                        allocator,
                        &mut content,
                        &mut link_annotations,
                        render_box,
                        &block.runs,
                        block.alignment,
                        block.vertical_alignment,
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
            link_annotations,
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

struct PdfLinkAnnotation {
    object_id: usize,
    rect: PdfRect,
    url: String,
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
    let rect = pdf_rect_from_rendered_vertices(rendered_box_vertices(render_box), page_h);
    content.extend_from_slice(
        format!(
            "q {:.4} 0 0 {:.4} {:.4} {:.4} cm /{} Do Q\n",
            rect.w, rect.h, rect.x, rect.y, image_name
        )
        .as_bytes(),
    );
}

fn push_text_box(
    allocator: &mut ObjectAllocator,
    content: &mut Vec<u8>,
    link_annotations: &mut Vec<PdfLinkAnnotation>,
    render_box: &RenderBox,
    runs: &[TextRun],
    alignment: TextAlignment,
    vertical_alignment: TextVerticalAlignment,
    page_h: f32,
) {
    let vertices = rendered_text_box_vertices(render_box);
    let frame_rect = pdf_rect_from_rendered_vertices(vertices, page_h);
    push_text_box_frame(content, render_box, frame_rect);

    let padding = render_box.style.padding * PT_PER_PX;
    let max_width = (frame_rect.w - padding.x * 2.0).max(12.0);
    let lines = wrap_runs(runs, max_width / PT_PER_PX)
        .into_iter()
        .map(merge_adjacent_segments)
        .collect::<Vec<_>>();
    let line_metrics = lines
        .iter()
        .map(|line| pdf_line_metrics(line))
        .collect::<Vec<_>>();
    let text_height = line_metrics
        .iter()
        .map(|metrics| metrics.height)
        .sum::<f32>();
    let content_height = (frame_rect.h - padding.y * 2.0).max(0.0);
    let vertical_offset = match vertical_alignment {
        TextVerticalAlignment::Top => 0.0,
        TextVerticalAlignment::Center => (content_height - text_height).max(0.0) * 0.5,
        TextVerticalAlignment::Bottom => (content_height - text_height).max(0.0),
    };
    let mut line_top = frame_rect.y + frame_rect.h - padding.y - vertical_offset;

    for (line, metrics) in lines.into_iter().zip(line_metrics) {
        let line_width = line
            .iter()
            .map(|segment| px_to_pt(segment.width_px()))
            .sum::<f32>();
        let x_offset = match alignment {
            TextAlignment::Center => ((max_width - line_width) / 2.0).max(0.0),
            TextAlignment::Right => (max_width - line_width).max(0.0),
            TextAlignment::Left | TextAlignment::Justify => 0.0,
        };
        let mut x = frame_rect.x + padding.x + x_offset;
        let y = line_top - metrics.baseline_offset;

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
            if let Some(url) = &segment.style.hyperlink {
                link_annotations.push(PdfLinkAnnotation {
                    object_id: allocator.alloc(),
                    rect: PdfRect {
                        x,
                        y: y - font_size * 0.25,
                        w: segment_width,
                        h: font_size * 1.2,
                    },
                    url: url.clone(),
                });
            }
            x += px_to_pt(segment.width_px());
        }
        line_top -= metrics.height;
    }
}

#[derive(Clone, Copy)]
struct PdfLineMetrics {
    height: f32,
    baseline_offset: f32,
}

fn pdf_line_metrics(line: &[TextSegment]) -> PdfLineMetrics {
    let max_font_size = line
        .iter()
        .map(|segment| px_to_pt(segment.style.font_size))
        .fold(10.0_f32, f32::max);
    PdfLineMetrics {
        height: (max_font_size * 1.2).max(12.0),
        baseline_offset: max_font_size,
    }
}

#[derive(Clone, Copy, Debug)]
struct PdfRect {
    x: f32,
    y: f32,
    w: f32,
    h: f32,
}

fn rendered_box_vertices(render_box: &RenderBox) -> [Pos2; 4] {
    let rect = render_box.rect(Pos2::ZERO, 1.0);
    [
        rect.left_top(),
        rect.right_top(),
        rect.right_bottom(),
        rect.left_bottom(),
    ]
}

fn rendered_text_box_vertices(render_box: &RenderBox) -> [Pos2; 4] {
    let size = if render_box.lock_size {
        render_box.authored_size.unwrap_or(render_box.size)
    } else {
        render_box.size
    };
    let rect = egui::Rect::from_min_size(
        render_box.position,
        egui::vec2(size.x * render_box.scale.x, size.y * render_box.scale.y),
    );
    [
        rect.left_top(),
        rect.right_top(),
        rect.right_bottom(),
        rect.left_bottom(),
    ]
}

fn pdf_rect_from_rendered_vertices(vertices: [Pos2; 4], page_h: f32) -> PdfRect {
    let min_x = vertices
        .iter()
        .map(|vertex| vertex.x)
        .fold(f32::INFINITY, f32::min);
    let max_x = vertices
        .iter()
        .map(|vertex| vertex.x)
        .fold(f32::NEG_INFINITY, f32::max);
    let min_y = vertices
        .iter()
        .map(|vertex| vertex.y)
        .fold(f32::INFINITY, f32::min);
    let max_y = vertices
        .iter()
        .map(|vertex| vertex.y)
        .fold(f32::NEG_INFINITY, f32::max);

    let x = px_to_pt(min_x);
    let y = page_h - px_to_pt(max_y);
    let w = px_to_pt(max_x - min_x);
    let h = px_to_pt(max_y - min_y);
    PdfRect { x, y, w, h }
}

fn push_text_box_frame(content: &mut Vec<u8>, render_box: &RenderBox, rect: PdfRect) {
    if render_box.style.fill != Color32::TRANSPARENT {
        push_rect_fill(
            content,
            rect.x,
            rect.y,
            rect.w,
            rect.h,
            render_box.style.fill,
        );
    }

    if render_box.style.stroke == Color32::TRANSPARENT || render_box.style.stroke_width <= 0.0 {
        return;
    }

    push_rect_stroke(
        content,
        rect,
        render_box.style.stroke,
        px_to_pt(render_box.style.stroke_width),
        render_box.style.stroke_kind,
    );
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

fn push_rect_stroke(
    content: &mut Vec<u8>,
    rect: PdfRect,
    color: Color32,
    stroke_width: f32,
    stroke_kind: BoxStrokeKind,
) {
    let inset = stroke_width * 0.5;
    let stroke_rect = PdfRect {
        x: rect.x + inset,
        y: rect.y + inset,
        w: (rect.w - stroke_width).max(0.0),
        h: (rect.h - stroke_width).max(0.0),
    };
    let dash = match stroke_kind {
        BoxStrokeKind::Solid => String::new(),
        BoxStrokeKind::Dash => format!(
            "[{:.4} {:.4}] 0 d ",
            (6.0 * PT_PER_PX).max(stroke_width * 2.0),
            (4.0 * PT_PER_PX).max(stroke_width)
        ),
    };
    let reset_dash = match stroke_kind {
        BoxStrokeKind::Solid => "",
        BoxStrokeKind::Dash => " [] 0 d",
    };
    content.extend_from_slice(
        format!(
            "{:.4} {:.4} {:.4} RG {:.4} w {dash}{:.4} {:.4} {:.4} {:.4} re S{reset_dash}\n",
            color_unit(color.r()),
            color_unit(color.g()),
            color_unit(color.b()),
            stroke_width.max(0.1),
            stroke_rect.x,
            stroke_rect.y,
            stroke_rect.w,
            stroke_rect.h
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
    pdf_escape_string_with_filter(text, false)
}

fn pdf_escape_string(text: &str) -> String {
    pdf_escape_string_with_filter(text, true)
}

fn pdf_escape_string_with_filter(text: &str, keep_control_as_space: bool) -> String {
    let mut escaped = String::new();
    for ch in text.chars() {
        match ch {
            '(' => escaped.push_str("\\("),
            ')' => escaped.push_str("\\)"),
            '\\' => escaped.push_str("\\\\"),
            '\r' | '\n' if keep_control_as_space => escaped.push(' '),
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
    use rich_canvas::{BoxStrokeKind, LayoutRole, RenderBox, TextStyle, TextVerticalAlignment};

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
    fn text_box_fill_and_border_are_written_to_pdf() {
        let mut slide = RichCanvas::new(vec2(1280.0, 720.0));
        let mut text = RenderBox::text(
            1,
            LayoutRole::Absolute,
            vec![TextRun::new("Styled box", TextStyle::body())],
        );
        text.position = pos2(80.0, 80.0);
        text.size = vec2(320.0, 90.0);
        text.style.fill = Color32::from_rgb(12, 34, 56);
        text.style.stroke = Color32::from_rgb(200, 120, 30);
        text.style.stroke_width = 2.5;
        slide.push(text);

        let bytes = export_pdf_bytes(&[slide]).expect("PDF export should succeed");
        let pdf = String::from_utf8_lossy(&bytes);

        assert!(pdf.contains("0.0471 0.1333 0.2196 rg"));
        assert!(pdf.contains("0.7843 0.4706 0.1176 RG"));
        assert!(pdf.contains("1.5502 w"));
        assert!(pdf.contains("(Styled box) Tj"));
    }

    #[test]
    fn dashed_text_box_border_is_written_to_pdf() {
        let mut slide = RichCanvas::new(vec2(1280.0, 720.0));
        let mut text = RenderBox::text(
            1,
            LayoutRole::Absolute,
            vec![TextRun::new("Dashed box", TextStyle::body())],
        );
        text.style.stroke = Color32::BLACK;
        text.style.stroke_kind = BoxStrokeKind::Dash;
        slide.push(text);

        let bytes = export_pdf_bytes(&[slide]).expect("PDF export should succeed");
        let pdf = String::from_utf8_lossy(&bytes);

        assert!(pdf.contains("[3.7205 2.4803] 0 d"));
        assert!(pdf.contains("[] 0 d"));
        assert!(pdf.contains("(Dashed box) Tj"));
    }

    #[test]
    fn pdf_text_box_frame_uses_rendered_canvas_vertices() {
        let mut text = RenderBox::text(
            1,
            LayoutRole::Absolute,
            vec![TextRun::new("Geometry", TextStyle::body())],
        );
        text.position = pos2(80.0, 90.0);
        text.size = vec2(320.0, 100.0);
        text.scale = vec2(1.25, 1.5);

        let vertices = rendered_box_vertices(&text);
        assert_close(vertices[0].x, 80.0);
        assert_close(vertices[0].y, 90.0);
        assert_close(vertices[2].x, 480.0);
        assert_close(vertices[2].y, 240.0);

        let rect = pdf_rect_from_rendered_vertices(vertices, px_to_pt(720.0));
        assert_close(rect.x, px_to_pt(80.0));
        assert_close(rect.y, px_to_pt(720.0 - 240.0));
        assert_close(rect.w, px_to_pt(400.0));
        assert_close(rect.h, px_to_pt(150.0));
    }

    #[test]
    fn pdf_text_box_frame_uses_authored_size() {
        let mut text = RenderBox::text(
            1,
            LayoutRole::Absolute,
            vec![TextRun::new("Authored geometry", TextStyle::body())],
        );
        text.position = pos2(80.0, 90.0);
        text.authored_size = Some(vec2(320.0, 100.0));
        text.size = vec2(360.0, 140.0);
        text.lock_size = true;
        text.scale = vec2(1.25, 1.5);

        let vertices = rendered_text_box_vertices(&text);
        assert_close(vertices[0].x, 80.0);
        assert_close(vertices[0].y, 90.0);
        assert_close(vertices[2].x, 480.0);
        assert_close(vertices[2].y, 240.0);

        let rect = pdf_rect_from_rendered_vertices(vertices, px_to_pt(720.0));
        assert_close(rect.x, px_to_pt(80.0));
        assert_close(rect.y, px_to_pt(720.0 - 240.0));
        assert_close(rect.w, px_to_pt(400.0));
        assert_close(rect.h, px_to_pt(150.0));
    }

    #[test]
    fn pdf_text_uses_text_box_vertical_alignment() {
        let mut slide = RichCanvas::new(vec2(1280.0, 720.0));
        let mut text = RenderBox::text(
            1,
            LayoutRole::Absolute,
            vec![TextRun::new("Centered", TextStyle::body())],
        );
        text.position = pos2(80.0, 80.0);
        text.size = vec2(320.0, 240.0);
        text.style.fill = Color32::TRANSPARENT;
        text.style.stroke = Color32::TRANSPARENT;
        if let RenderBoxKind::Text(block) = &mut text.kind {
            block.vertical_alignment = TextVerticalAlignment::Center;
        }
        slide.push(text);

        let bytes = export_pdf_bytes(&[slide]).expect("PDF export should succeed");
        let pdf = String::from_utf8_lossy(&bytes);

        assert!(pdf.contains("60.7677 317.4803 Td (Centered) Tj"));
    }

    #[test]
    fn hyperlink_text_writes_pdf_link_annotation() {
        let mut slide = RichCanvas::new(vec2(1280.0, 720.0));
        let mut style = TextStyle::body();
        style.hyperlink = Some("https://example.com/path?q=1".to_owned());
        let mut text = RenderBox::text(
            1,
            LayoutRole::Absolute,
            vec![TextRun::new("Example link", style)],
        );
        text.position = pos2(80.0, 80.0);
        text.size = vec2(320.0, 90.0);
        slide.push(text);

        let bytes = export_pdf_bytes(&[slide]).expect("PDF export should succeed");
        let pdf = String::from_utf8_lossy(&bytes);

        assert!(pdf.contains("/Annots ["));
        assert!(pdf.contains("/Subtype /Link"));
        assert!(pdf.contains("/S /URI /URI (https://example.com/path?q=1)"));
        assert!(pdf.contains("(Example link) Tj"));
    }

    #[test]
    fn pdf_string_escape_handles_link_urls() {
        assert_eq!(
            pdf_escape_string(r"https://example.com/a(b)\c"),
            r"https://example.com/a\(b\)\\c"
        );
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

    fn assert_close(actual: f32, expected: f32) {
        assert!(
            (actual - expected).abs() < 0.001,
            "expected {actual} to be close to {expected}"
        );
    }
}
