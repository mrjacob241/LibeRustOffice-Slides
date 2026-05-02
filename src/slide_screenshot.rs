use std::{
    fmt,
    path::{Path, PathBuf},
    sync::{Arc, Mutex},
};

use eframe::{App, NativeOptions, egui};
use egui::{Color32, Rect};
use image::{ImageBuffer, Rgba, RgbaImage};
use rich_canvas::{RenderBox, RenderBoxKind, RichCanvas, TextAlignment, TextStyle};

use crate::odp_loader::{self, OdpLoadError};

#[derive(Debug)]
pub enum SlideScreenshotError {
    Io(std::io::Error),
    Image(image::ImageError),
    Odp(OdpLoadError),
    MissingReference(PathBuf),
    InvalidScale,
    Window(String),
    Capture(String),
}

impl fmt::Display for SlideScreenshotError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Io(error) => write!(formatter, "{error}"),
            Self::Image(error) => write!(formatter, "{error}"),
            Self::Odp(error) => write!(formatter, "{error}"),
            Self::MissingReference(path) => {
                write!(formatter, "missing reference screenshot {}", path.display())
            }
            Self::InvalidScale => write!(formatter, "screenshot scale must be finite and positive"),
            Self::Window(message) => write!(formatter, "{message}"),
            Self::Capture(message) => write!(formatter, "{message}"),
        }
    }
}

impl From<std::io::Error> for SlideScreenshotError {
    fn from(error: std::io::Error) -> Self {
        Self::Io(error)
    }
}

impl From<image::ImageError> for SlideScreenshotError {
    fn from(error: image::ImageError) -> Self {
        Self::Image(error)
    }
}

impl From<OdpLoadError> for SlideScreenshotError {
    fn from(error: OdpLoadError) -> Self {
        Self::Odp(error)
    }
}

#[derive(Clone, Debug)]
pub struct SlideScreenshot {
    pub slide_index: usize,
    pub path: PathBuf,
    pub width: u32,
    pub height: u32,
}

#[derive(Clone, Debug)]
pub struct SlideScreenshotComparison {
    pub slide_index: usize,
    pub rendered_path: PathBuf,
    pub reference_path: PathBuf,
    pub rendered_width: u32,
    pub rendered_height: u32,
    pub reference_width: u32,
    pub reference_height: u32,
    pub dimensions_match: bool,
}

pub fn render_default_odp_screenshots_and_compare_references()
-> Result<Vec<SlideScreenshotComparison>, SlideScreenshotError> {
    let loaded = odp_loader::load_default_odp()?;
    let output_dir = Path::new(env!("CARGO_MANIFEST_DIR")).join("target/window_slide_screenshots");
    let reference_dir = Path::new(env!("CARGO_MANIFEST_DIR")).join("sample_docs");
    render_slides_with_window_and_compare_references(
        loaded.slides,
        &output_dir,
        &reference_dir,
        egui::vec2(1041.0, 590.0),
    )
}

#[allow(dead_code)]
pub fn render_default_odp_window_screenshots() -> Result<Vec<SlideScreenshot>, SlideScreenshotError>
{
    let loaded = odp_loader::load_default_odp()?;
    let output_dir = Path::new(env!("CARGO_MANIFEST_DIR")).join("target/window_slide_screenshots");
    render_slide_window_screenshots(loaded.slides, &output_dir, egui::vec2(1041.0, 590.0))
}

pub fn render_slides_with_window_and_compare_references(
    slides: Vec<RichCanvas>,
    output_dir: &Path,
    reference_dir: &Path,
    window_size: egui::Vec2,
) -> Result<Vec<SlideScreenshotComparison>, SlideScreenshotError> {
    let rendered = render_slide_window_screenshots(slides, output_dir, window_size)?;
    compare_screenshots_with_references(&rendered, reference_dir)
}

pub fn render_slide_window_screenshots(
    slides: Vec<RichCanvas>,
    output_dir: &Path,
    window_size: egui::Vec2,
) -> Result<Vec<SlideScreenshot>, SlideScreenshotError> {
    std::fs::create_dir_all(output_dir)?;

    let output_dir = output_dir.to_path_buf();
    let result = Arc::new(Mutex::new(WindowCaptureResult::default()));
    let app_result = Arc::clone(&result);
    let options = NativeOptions {
        viewport: egui::ViewportBuilder::default()
            .with_inner_size(window_size)
            .with_resizable(false)
            .with_decorations(false)
            .with_title("LibeRustOffice slide screenshot capture"),
        vsync: false,
        ..Default::default()
    };

    eframe::run_native(
        "LibeRustOffice slide screenshot capture",
        options,
        Box::new(move |cc| {
            rich_canvas::configure_editor_fonts(&cc.egui_ctx);
            Ok(Box::new(WindowCaptureApp {
                slides,
                output_dir,
                result: app_result,
                current_slide: 0,
                requested_slide: None,
            }))
        }),
    )
    .map_err(|error| SlideScreenshotError::Window(error.to_string()))?;

    let result = Arc::try_unwrap(result)
        .map_err(|_| SlideScreenshotError::Capture("capture result still has owners".to_owned()))?
        .into_inner()
        .map_err(|_| {
            SlideScreenshotError::Capture("capture result lock was poisoned".to_owned())
        })?;
    if let Some(error) = result.error {
        return Err(SlideScreenshotError::Capture(error));
    }
    Ok(result.screenshots)
}

#[allow(dead_code)]
pub fn render_slides_and_compare_references(
    slides: &[RichCanvas],
    output_dir: &Path,
    reference_dir: &Path,
    scale: f32,
) -> Result<Vec<SlideScreenshotComparison>, SlideScreenshotError> {
    let rendered = render_slide_screenshots(slides, output_dir, scale)?;
    compare_screenshots_with_references(&rendered, reference_dir)
}

#[allow(dead_code)]
pub fn render_slide_screenshots(
    slides: &[RichCanvas],
    output_dir: &Path,
    scale: f32,
) -> Result<Vec<SlideScreenshot>, SlideScreenshotError> {
    if !scale.is_finite() || scale <= 0.0 {
        return Err(SlideScreenshotError::InvalidScale);
    }

    std::fs::create_dir_all(output_dir)?;
    slides
        .iter()
        .enumerate()
        .map(|(index, slide)| {
            let image = render_slide_image(slide, scale)?;
            let path = output_dir.join(format!("slide_{}.png", index + 1));
            image.save(&path)?;
            Ok(SlideScreenshot {
                slide_index: index + 1,
                path,
                width: image.width(),
                height: image.height(),
            })
        })
        .collect()
}

pub fn compare_screenshots_with_references(
    screenshots: &[SlideScreenshot],
    reference_dir: &Path,
) -> Result<Vec<SlideScreenshotComparison>, SlideScreenshotError> {
    screenshots
        .iter()
        .map(|screenshot| {
            let reference_path =
                reference_dir.join(format!("ref_slide_{}.png", screenshot.slide_index));
            if !reference_path.exists() {
                return Err(SlideScreenshotError::MissingReference(reference_path));
            }

            let rendered = image::open(&screenshot.path)?.to_rgba8();
            debug_assert_eq!(rendered.width(), screenshot.width);
            debug_assert_eq!(rendered.height(), screenshot.height);
            let reference = image::open(&reference_path)?.to_rgba8();
            Ok(SlideScreenshotComparison {
                slide_index: screenshot.slide_index,
                rendered_path: screenshot.path.clone(),
                reference_path,
                rendered_width: rendered.width(),
                rendered_height: rendered.height(),
                reference_width: reference.width(),
                reference_height: reference.height(),
                dimensions_match: rendered.dimensions() == reference.dimensions(),
            })
        })
        .collect()
}

#[derive(Default)]
struct WindowCaptureResult {
    screenshots: Vec<SlideScreenshot>,
    error: Option<String>,
}

struct WindowCaptureApp {
    slides: Vec<RichCanvas>,
    output_dir: PathBuf,
    result: Arc<Mutex<WindowCaptureResult>>,
    current_slide: usize,
    requested_slide: Option<usize>,
}

impl App for WindowCaptureApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        self.handle_screenshot_events(ctx);
        if self.current_slide >= self.slides.len() {
            ctx.send_viewport_cmd(egui::ViewportCommand::Close);
            return;
        }

        egui::CentralPanel::default()
            .frame(egui::Frame::NONE)
            .show(ctx, |ui| {
                let rect = ui.max_rect();
                let painter = ui.painter_at(rect);
                self.slides[self.current_slide].paint_page_preview(&painter, rect);
            });

        if self.requested_slide != Some(self.current_slide) {
            ctx.send_viewport_cmd(egui::ViewportCommand::Screenshot(egui::UserData::new(
                self.current_slide,
            )));
            self.requested_slide = Some(self.current_slide);
        }
        ctx.request_repaint();
    }
}

impl WindowCaptureApp {
    fn handle_screenshot_events(&mut self, ctx: &egui::Context) {
        let events = ctx.input(|input| input.events.clone());
        for event in events {
            let egui::Event::Screenshot {
                user_data, image, ..
            } = event
            else {
                continue;
            };

            let Some(slide_index) = user_data
                .data
                .as_ref()
                .and_then(|data| data.downcast_ref::<usize>())
                .copied()
            else {
                continue;
            };

            if let Err(error) = self.save_screenshot(slide_index, &image) {
                if let Ok(mut result) = self.result.lock() {
                    result.error = Some(error.to_string());
                }
                ctx.send_viewport_cmd(egui::ViewportCommand::Close);
                return;
            }
            self.current_slide = slide_index + 1;
            self.requested_slide = None;
        }
    }

    fn save_screenshot(
        &mut self,
        slide_index: usize,
        image: &egui::ColorImage,
    ) -> Result<(), SlideScreenshotError> {
        let path = self
            .output_dir
            .join(format!("slide_{}.png", slide_index + 1));
        color_image_to_rgba(image).save(&path)?;

        let screenshot = SlideScreenshot {
            slide_index: slide_index + 1,
            path,
            width: image.size[0] as u32,
            height: image.size[1] as u32,
        };
        let mut result = self.result.lock().map_err(|_| {
            SlideScreenshotError::Capture("capture result lock was poisoned".to_owned())
        })?;
        result.screenshots.push(screenshot);
        Ok(())
    }
}

fn color_image_to_rgba(image: &egui::ColorImage) -> RgbaImage {
    let mut rgba = RgbaImage::new(image.size[0] as u32, image.size[1] as u32);
    for (index, color) in image.pixels.iter().enumerate() {
        let x = (index % image.size[0]) as u32;
        let y = (index / image.size[0]) as u32;
        rgba.put_pixel(x, y, color_to_rgba(*color));
    }
    rgba
}

pub fn render_slide_image(
    slide: &RichCanvas,
    scale: f32,
) -> Result<RgbaImage, SlideScreenshotError> {
    if !scale.is_finite() || scale <= 0.0 {
        return Err(SlideScreenshotError::InvalidScale);
    }

    let width = (slide.page.size.x * scale).round().max(1.0) as u32;
    let height = (slide.page.size.y * scale).round().max(1.0) as u32;
    let mut image = ImageBuffer::from_pixel(width, height, color_to_rgba(slide.page.fill));

    let mut boxes = slide.boxes.iter().collect::<Vec<_>>();
    boxes.sort_by_key(|render_box| render_box.z_index);
    for render_box in boxes {
        if !render_box.visible {
            continue;
        }
        draw_render_box(&mut image, render_box, scale);
    }

    Ok(image)
}

fn draw_render_box(target: &mut RgbaImage, render_box: &RenderBox, scale: f32) {
    let rect = render_box.rect(egui::Pos2::ZERO, scale);
    if render_box.style.fill != Color32::TRANSPARENT {
        fill_rect(target, rect, render_box.style.fill);
    }

    match &render_box.kind {
        RenderBoxKind::Image(block) => draw_color_image(target, rect, block),
        RenderBoxKind::Text(block) => draw_text_block(target, rect, block, scale),
        RenderBoxKind::Table(_) | RenderBoxKind::Container => {}
    }

    if render_box.style.stroke != Color32::TRANSPARENT && render_box.style.stroke_width > 0.0 {
        stroke_rect(
            target,
            rect,
            render_box.style.stroke,
            (render_box.style.stroke_width * scale).round().max(1.0) as u32,
        );
    }
}

#[derive(Clone, Debug)]
struct ScreenshotGlyph {
    ch: char,
    style: TextStyle,
}

#[derive(Clone, Debug)]
struct ScreenshotTextLine {
    glyphs: Vec<ScreenshotGlyph>,
    width: f32,
    height: f32,
}

fn draw_text_block(
    target: &mut RgbaImage,
    rect: Rect,
    block: &rich_canvas::RichTextBlock,
    scale: f32,
) {
    let lines = layout_text_lines(block, rect.width().max(1.0), scale);
    if lines.is_empty() {
        return;
    }

    let total_height = lines.iter().map(|line| line.height).sum::<f32>();
    let mut cursor_y = match block.vertical_alignment {
        rich_canvas::TextVerticalAlignment::Top => rect.min.y,
        rich_canvas::TextVerticalAlignment::Center => {
            rect.min.y + (rect.height() - total_height).max(0.0) / 2.0
        }
        rich_canvas::TextVerticalAlignment::Bottom => rect.max.y - total_height,
    };

    for line in lines {
        let mut cursor_x = match block.alignment {
            TextAlignment::Left | TextAlignment::Justify => rect.min.x,
            TextAlignment::Center => rect.min.x + (rect.width() - line.width).max(0.0) / 2.0,
            TextAlignment::Right => rect.max.x - line.width,
        };

        for glyph in line.glyphs {
            let metrics = glyph_metrics(&glyph.style, scale);
            if glyph.ch == ' ' {
                cursor_x += metrics.advance;
                continue;
            }
            if let Some(background) = glyph.style.background_color {
                fill_rect(
                    target,
                    Rect::from_min_size(
                        egui::pos2(cursor_x, cursor_y),
                        egui::vec2(metrics.advance, line.height),
                    ),
                    background,
                );
            }
            draw_bitmap_glyph(
                target,
                glyph.ch,
                cursor_x,
                cursor_y,
                metrics.pixel,
                glyph.style.color,
                glyph.style.bold,
                glyph.style.italic,
            );
            if glyph.style.underline {
                fill_rect(
                    target,
                    Rect::from_min_size(
                        egui::pos2(cursor_x, cursor_y + metrics.height + metrics.pixel),
                        egui::vec2(metrics.advance - metrics.pixel, metrics.pixel.max(1.0)),
                    ),
                    glyph.style.color,
                );
            }
            cursor_x += metrics.advance;
        }

        cursor_y += line.height;
        if cursor_y > rect.max.y {
            break;
        }
    }
}

fn layout_text_lines(
    block: &rich_canvas::RichTextBlock,
    max_width: f32,
    scale: f32,
) -> Vec<ScreenshotTextLine> {
    let mut lines = Vec::new();
    let mut current = ScreenshotTextLine {
        glyphs: Vec::new(),
        width: 0.0,
        height: 0.0,
    };

    for run in &block.runs {
        for ch in run.text.chars() {
            if ch == '\n' {
                push_text_line(&mut lines, &mut current);
                continue;
            }

            let metrics = glyph_metrics(&run.style, scale);
            if !current.glyphs.is_empty()
                && current.width + metrics.advance > max_width
                && ch != ' '
            {
                push_text_line(&mut lines, &mut current);
            }
            current.width += metrics.advance;
            current.height = current.height.max(metrics.line_height);
            current.glyphs.push(ScreenshotGlyph {
                ch,
                style: run.style.clone(),
            });
        }
    }
    push_text_line(&mut lines, &mut current);
    lines
}

fn push_text_line(lines: &mut Vec<ScreenshotTextLine>, current: &mut ScreenshotTextLine) {
    if current.glyphs.is_empty() {
        current.height = current.height.max(12.0);
    }
    lines.push(std::mem::replace(
        current,
        ScreenshotTextLine {
            glyphs: Vec::new(),
            width: 0.0,
            height: 0.0,
        },
    ));
}

#[derive(Clone, Copy, Debug)]
struct GlyphMetrics {
    pixel: f32,
    advance: f32,
    height: f32,
    line_height: f32,
}

fn glyph_metrics(style: &TextStyle, scale: f32) -> GlyphMetrics {
    let pixel = ((style.font_size * scale) / 9.0).round().max(1.0);
    GlyphMetrics {
        pixel,
        advance: pixel * 6.0,
        height: pixel * 7.0,
        line_height: pixel * 9.0,
    }
}

fn draw_bitmap_glyph(
    target: &mut RgbaImage,
    ch: char,
    x: f32,
    y: f32,
    pixel: f32,
    color: Color32,
    bold: bool,
    italic: bool,
) {
    let pattern = glyph_pattern(ch);
    let pixel_size = pixel.round().max(1.0) as u32;
    for (row, bits) in pattern.iter().enumerate() {
        let italic_offset = if italic {
            (6usize.saturating_sub(row)) as f32 * pixel * 0.2
        } else {
            0.0
        };
        for col in 0..5 {
            if bits & (1 << (4 - col)) == 0 {
                continue;
            }
            let glyph_rect = Rect::from_min_size(
                egui::pos2(
                    x + col as f32 * pixel + italic_offset,
                    y + row as f32 * pixel,
                ),
                egui::vec2(
                    pixel_size as f32 + u32::from(bold) as f32,
                    pixel_size as f32,
                ),
            );
            fill_rect(target, glyph_rect, color);
        }
    }
}

fn glyph_pattern(ch: char) -> [u8; 7] {
    match ch.to_ascii_uppercase() {
        'A' => [
            0b01110, 0b10001, 0b10001, 0b11111, 0b10001, 0b10001, 0b10001,
        ],
        'B' => [
            0b11110, 0b10001, 0b10001, 0b11110, 0b10001, 0b10001, 0b11110,
        ],
        'C' => [
            0b01110, 0b10001, 0b10000, 0b10000, 0b10000, 0b10001, 0b01110,
        ],
        'D' => [
            0b11110, 0b10001, 0b10001, 0b10001, 0b10001, 0b10001, 0b11110,
        ],
        'E' => [
            0b11111, 0b10000, 0b10000, 0b11110, 0b10000, 0b10000, 0b11111,
        ],
        'F' => [
            0b11111, 0b10000, 0b10000, 0b11110, 0b10000, 0b10000, 0b10000,
        ],
        'G' => [
            0b01110, 0b10001, 0b10000, 0b10111, 0b10001, 0b10001, 0b01110,
        ],
        'H' => [
            0b10001, 0b10001, 0b10001, 0b11111, 0b10001, 0b10001, 0b10001,
        ],
        'I' => [
            0b11111, 0b00100, 0b00100, 0b00100, 0b00100, 0b00100, 0b11111,
        ],
        'J' => [
            0b00111, 0b00010, 0b00010, 0b00010, 0b10010, 0b10010, 0b01100,
        ],
        'K' => [
            0b10001, 0b10010, 0b10100, 0b11000, 0b10100, 0b10010, 0b10001,
        ],
        'L' => [
            0b10000, 0b10000, 0b10000, 0b10000, 0b10000, 0b10000, 0b11111,
        ],
        'M' => [
            0b10001, 0b11011, 0b10101, 0b10101, 0b10001, 0b10001, 0b10001,
        ],
        'N' => [
            0b10001, 0b11001, 0b10101, 0b10011, 0b10001, 0b10001, 0b10001,
        ],
        'O' => [
            0b01110, 0b10001, 0b10001, 0b10001, 0b10001, 0b10001, 0b01110,
        ],
        'P' => [
            0b11110, 0b10001, 0b10001, 0b11110, 0b10000, 0b10000, 0b10000,
        ],
        'Q' => [
            0b01110, 0b10001, 0b10001, 0b10001, 0b10101, 0b10010, 0b01101,
        ],
        'R' => [
            0b11110, 0b10001, 0b10001, 0b11110, 0b10100, 0b10010, 0b10001,
        ],
        'S' => [
            0b01111, 0b10000, 0b10000, 0b01110, 0b00001, 0b00001, 0b11110,
        ],
        'T' => [
            0b11111, 0b00100, 0b00100, 0b00100, 0b00100, 0b00100, 0b00100,
        ],
        'U' => [
            0b10001, 0b10001, 0b10001, 0b10001, 0b10001, 0b10001, 0b01110,
        ],
        'V' => [
            0b10001, 0b10001, 0b10001, 0b10001, 0b01010, 0b01010, 0b00100,
        ],
        'W' => [
            0b10001, 0b10001, 0b10001, 0b10101, 0b10101, 0b10101, 0b01010,
        ],
        'X' => [
            0b10001, 0b01010, 0b00100, 0b00100, 0b00100, 0b01010, 0b10001,
        ],
        'Y' => [
            0b10001, 0b01010, 0b00100, 0b00100, 0b00100, 0b00100, 0b00100,
        ],
        'Z' => [
            0b11111, 0b00001, 0b00010, 0b00100, 0b01000, 0b10000, 0b11111,
        ],
        '0' => [
            0b01110, 0b10001, 0b10011, 0b10101, 0b11001, 0b10001, 0b01110,
        ],
        '1' => [
            0b00100, 0b01100, 0b00100, 0b00100, 0b00100, 0b00100, 0b01110,
        ],
        '2' => [
            0b01110, 0b10001, 0b00001, 0b00010, 0b00100, 0b01000, 0b11111,
        ],
        '3' => [
            0b11110, 0b00001, 0b00001, 0b01110, 0b00001, 0b00001, 0b11110,
        ],
        '4' => [
            0b00010, 0b00110, 0b01010, 0b10010, 0b11111, 0b00010, 0b00010,
        ],
        '5' => [
            0b11111, 0b10000, 0b10000, 0b11110, 0b00001, 0b00001, 0b11110,
        ],
        '6' => [
            0b01110, 0b10000, 0b10000, 0b11110, 0b10001, 0b10001, 0b01110,
        ],
        '7' => [
            0b11111, 0b00001, 0b00010, 0b00100, 0b01000, 0b01000, 0b01000,
        ],
        '8' => [
            0b01110, 0b10001, 0b10001, 0b01110, 0b10001, 0b10001, 0b01110,
        ],
        '9' => [
            0b01110, 0b10001, 0b10001, 0b01111, 0b00001, 0b00001, 0b01110,
        ],
        '.' => [0, 0, 0, 0, 0, 0b01100, 0b01100],
        ',' => [0, 0, 0, 0, 0, 0b01100, 0b01000],
        ':' => [0, 0b01100, 0b01100, 0, 0b01100, 0b01100, 0],
        ';' => [0, 0b01100, 0b01100, 0, 0b01100, 0b01000, 0],
        '-' | '–' => [0, 0, 0, 0b11111, 0, 0, 0],
        '_' => [0, 0, 0, 0, 0, 0, 0b11111],
        '(' => [
            0b00010, 0b00100, 0b01000, 0b01000, 0b01000, 0b00100, 0b00010,
        ],
        ')' => [
            0b01000, 0b00100, 0b00010, 0b00010, 0b00010, 0b00100, 0b01000,
        ],
        '/' => [
            0b00001, 0b00010, 0b00010, 0b00100, 0b01000, 0b01000, 0b10000,
        ],
        '\'' => [0b01100, 0b01100, 0b01000, 0, 0, 0, 0],
        '"' => [0b01010, 0b01010, 0b01010, 0, 0, 0, 0],
        '&' => [
            0b01100, 0b10010, 0b10100, 0b01000, 0b10101, 0b10010, 0b01101,
        ],
        '+' => [0, 0b00100, 0b00100, 0b11111, 0b00100, 0b00100, 0],
        '!' => [0b00100, 0b00100, 0b00100, 0b00100, 0b00100, 0, 0b00100],
        '?' => [0b01110, 0b10001, 0b00001, 0b00010, 0b00100, 0, 0b00100],
        _ => [0b11111, 0b10001, 0b00010, 0b00100, 0b00100, 0, 0b00100],
    }
}

fn draw_color_image(target: &mut RgbaImage, rect: Rect, block: &rich_canvas::ImageBlock) {
    let src_width = block.color_image.size[0];
    let src_height = block.color_image.size[1];
    if src_width == 0 || src_height == 0 {
        return;
    }

    let min_x = rect.min.x.round().max(0.0) as i32;
    let min_y = rect.min.y.round().max(0.0) as i32;
    let max_x = rect.max.x.round().min(target.width() as f32) as i32;
    let max_y = rect.max.y.round().min(target.height() as f32) as i32;
    let width = (max_x - min_x).max(0);
    let height = (max_y - min_y).max(0);
    if width == 0 || height == 0 {
        return;
    }

    for y in min_y..max_y {
        let v = ((y - min_y) as f32 / height as f32).clamp(0.0, 1.0);
        let src_y = (v * (src_height.saturating_sub(1)) as f32).round() as usize;
        for x in min_x..max_x {
            let u = ((x - min_x) as f32 / width as f32).clamp(0.0, 1.0);
            let src_x = (u * (src_width.saturating_sub(1)) as f32).round() as usize;
            let source = block.color_image.pixels[src_y * src_width + src_x];
            blend_pixel(target, x as u32, y as u32, source);
        }
    }
}

fn fill_rect(target: &mut RgbaImage, rect: Rect, color: Color32) {
    let min_x = rect.min.x.floor().max(0.0) as u32;
    let min_y = rect.min.y.floor().max(0.0) as u32;
    let max_x = rect.max.x.ceil().min(target.width() as f32) as u32;
    let max_y = rect.max.y.ceil().min(target.height() as f32) as u32;

    for y in min_y..max_y {
        for x in min_x..max_x {
            blend_pixel(target, x, y, color);
        }
    }
}

fn stroke_rect(target: &mut RgbaImage, rect: Rect, color: Color32, width: u32) {
    for offset in 0..width {
        let inset = offset as f32;
        fill_rect(
            target,
            Rect::from_min_max(
                egui::pos2(rect.min.x + inset, rect.min.y + inset),
                egui::pos2(rect.max.x - inset, rect.min.y + inset + 1.0),
            ),
            color,
        );
        fill_rect(
            target,
            Rect::from_min_max(
                egui::pos2(rect.min.x + inset, rect.max.y - inset - 1.0),
                egui::pos2(rect.max.x - inset, rect.max.y - inset),
            ),
            color,
        );
        fill_rect(
            target,
            Rect::from_min_max(
                egui::pos2(rect.min.x + inset, rect.min.y + inset),
                egui::pos2(rect.min.x + inset + 1.0, rect.max.y - inset),
            ),
            color,
        );
        fill_rect(
            target,
            Rect::from_min_max(
                egui::pos2(rect.max.x - inset - 1.0, rect.min.y + inset),
                egui::pos2(rect.max.x - inset, rect.max.y - inset),
            ),
            color,
        );
    }
}

fn blend_pixel(target: &mut RgbaImage, x: u32, y: u32, source: Color32) {
    let source_alpha = source.a() as f32 / 255.0;
    if source_alpha <= 0.0 {
        return;
    }
    let target_pixel = target.get_pixel_mut(x, y);
    let target_channels = target_pixel.0;
    let inverse_alpha = 1.0 - source_alpha;
    *target_pixel = Rgba([
        (source.r() as f32 * source_alpha + target_channels[0] as f32 * inverse_alpha).round()
            as u8,
        (source.g() as f32 * source_alpha + target_channels[1] as f32 * inverse_alpha).round()
            as u8,
        (source.b() as f32 * source_alpha + target_channels[2] as f32 * inverse_alpha).round()
            as u8,
        255,
    ]);
}

fn color_to_rgba(color: Color32) -> Rgba<u8> {
    Rgba([color.r(), color.g(), color.b(), color.a()])
}

#[cfg(test)]
mod tests {
    use super::*;
    use egui::vec2;
    use rich_canvas::{LayoutRole, TextRun};

    #[test]
    #[ignore = "opens an eframe window and captures real rendered frames"]
    fn default_odp_screenshots_can_be_compared_to_references() {
        let comparisons = render_default_odp_screenshots_and_compare_references()
            .expect("default screenshots should render and compare");

        assert!(!comparisons.is_empty());
        for comparison in &comparisons {
            eprintln!(
                "slide {}: rendered={} reference={} rendered_size={}x{} reference_size={}x{} dimensions_match={}",
                comparison.slide_index,
                comparison.rendered_path.display(),
                comparison.reference_path.display(),
                comparison.rendered_width,
                comparison.rendered_height,
                comparison.reference_width,
                comparison.reference_height,
                comparison.dimensions_match
            );
            assert!(comparison.rendered_path.exists());
            assert!(comparison.reference_path.exists());
            assert!(comparison.rendered_width > 0);
            assert!(comparison.rendered_height > 0);
            assert!(comparison.reference_width > 0);
            assert!(comparison.reference_height > 0);
            assert!(comparison.dimensions_match);
        }
    }

    #[test]
    fn rendered_screenshot_contains_text_pixels() {
        let mut slide = RichCanvas::new(vec2(320.0, 180.0));
        let mut text = RenderBox::text(
            1,
            LayoutRole::Absolute,
            vec![TextRun::new("Visible text", rich_canvas::TextStyle::body())],
        );
        text.position = egui::pos2(20.0, 20.0);
        text.size = vec2(240.0, 80.0);
        slide.push(text);

        let image = render_slide_image(&slide, 1.0).expect("text slide should render");
        let non_white_pixels = image
            .pixels()
            .filter(|pixel| pixel.0 != [255, 255, 255, 255])
            .count();

        assert!(non_white_pixels > 0);
    }
}
