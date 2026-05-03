use std::{
    collections::VecDeque,
    fs::OpenOptions,
    io::Write,
    path::{Path, PathBuf},
    sync::Arc,
    time::{SystemTime, UNIX_EPOCH},
};

use egui::{
    Align, Align2, Color32, ColorImage, FontFamily, FontId, FontSelection, Galley, Painter, Pos2,
    Rect, Response, RichText, Sense, Shape, Stroke, StrokeKind, Style, TextureHandle,
    TextureOptions, Ui, Vec2,
    epaint::{Mesh, Vertex, text::PlacedRow},
    text::{CCursor, LayoutJob, TextWrapping},
    vec2,
};

const EDITOR_REGULAR_FONT_NAME: &str = "editor_regular";
const EDITOR_BOLD_FONT_NAME: &str = "editor_bold";
const ZOOM_TELEMETRY_ENV: &str = "LIBERUSTOFFICE_ZOOM_TELEMETRY";
const ZOOM_SMOOTHING_WINDOW_SECONDS: f64 = 0.08;
const ZOOM_SMOOTHING_RESET_SECONDS: f64 = 0.16;
const ZOOM_SCROLL_SCALE: f32 = 0.0005;
const IMAGE_SELECTION_STROKE_WIDTH: f32 = 1.8;
const IMAGE_SELECTION_HANDLE_RADIUS: f32 = 4.5;
const IMAGE_ASPECT_HANDLE_RADIUS: f32 = IMAGE_SELECTION_HANDLE_RADIUS * 1.35;
const IMAGE_SELECTION_HANDLE_HIT_RADIUS: f32 = 28.0;
const IMAGE_SELECTION_CENTER_HANDLE_HIT_RADIUS: f32 = IMAGE_SELECTION_HANDLE_HIT_RADIUS * 2.0;
const TEXT_MOVE_HANDLE_SIZE: f32 = 12.0;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CanvasMode {
    TextDocument,
    SlideDeck,
}

#[derive(Clone, Debug)]
pub struct RichCanvas {
    pub size: Vec2,
    pub zoom: f32,
    pub background: Color32,
    pub page: PageFrame,
    pub boxes: Vec<RenderBox>,
    last_canvas_view: Option<CanvasViewState>,
    last_scroll_offset: Vec2,
    zoom_scroll_samples: VecDeque<(f64, f32)>,
}

const IMAGE_MIN_SIZE: f32 = 24.0;

#[derive(Clone, Copy, Debug)]
struct CanvasViewState {
    viewport_rect: Rect,
    canvas_origin: Pos2,
    canvas_size: Vec2,
    canvas_content_origin: Vec2,
}

fn zoom_scroll_offset_for_pointer(
    canvas_size: Vec2,
    current_zoom: f32,
    next_zoom: f32,
    view: CanvasViewState,
    pointer: Pos2,
    current_scroll_offset: Vec2,
) -> Option<Vec2> {
    let viewport_pos = pointer - view.viewport_rect.min;
    let current_canvas_origin =
        view.viewport_rect.min + view.canvas_content_origin - current_scroll_offset;
    let canvas_rect = Rect::from_min_size(current_canvas_origin, view.canvas_size);
    if !canvas_rect.contains(pointer) {
        return None;
    }

    let canvas_point =
        (current_scroll_offset + viewport_pos - view.canvas_content_origin) / current_zoom;
    let viewport_size = view.viewport_rect.size();
    let next_canvas_size = canvas_size * next_zoom;
    let next_allocated_size = vec2(
        next_canvas_size.x.max(viewport_size.x),
        next_canvas_size.y.max(viewport_size.y),
    );
    let next_canvas_origin = (next_allocated_size - next_canvas_size).max(Vec2::ZERO) * 0.5;

    let max_scroll_offset = (next_allocated_size - viewport_size).max(Vec2::ZERO);
    Some(
        (next_canvas_origin + canvas_point * next_zoom - viewport_pos)
            .clamp(Vec2::ZERO, max_scroll_offset),
    )
}

fn write_zoom_telemetry(line: String) {
    if std::env::var_os(ZOOM_TELEMETRY_ENV).is_none() {
        return;
    }

    eprintln!("{line}");

    if std::fs::create_dir_all("logs").is_err() {
        return;
    }

    if let Ok(mut file) = OpenOptions::new()
        .create(true)
        .append(true)
        .open("logs/zoom_telemetry.log")
    {
        let _ = writeln!(file, "{line}");
    }
}

fn zoom_telemetry_timestamp_ms() -> u128 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_millis())
        .unwrap_or_default()
}

pub struct CanvasUiResponse {
    pub response: Response,
    pub clicked_box: Option<u64>,
    pub clicked_caret: Option<usize>,
    pub dragged_box: Option<u64>,
    pub dragged_caret: Option<usize>,
    pub hovered_hyperlink: Option<CanvasHyperlink>,
    pub image_resize_handle: Option<ImageResizeHandle>,
    pub text_resize_handle: Option<ImageResizeHandle>,
    pub pointer_pos: Option<Pos2>,
    pub drag_started: bool,
    pub drag_stopped: bool,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CanvasHyperlink {
    pub box_id: u64,
    pub char_index: usize,
    pub url: String,
}

#[derive(Clone, Copy, Debug)]
pub struct CanvasSelection {
    pub box_id: u64,
    pub caret_index: Option<usize>,
    pub range: Option<TextRange>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ImageResizeHandle {
    Center,
    Left,
    Right,
    Top,
    Bottom,
    TopLeft,
    TopRight,
    BottomRight,
    BottomLeft,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct TextRange {
    pub anchor: usize,
    pub caret: usize,
}

impl TextRange {
    pub fn sorted(self) -> (usize, usize) {
        if self.anchor <= self.caret {
            (self.anchor, self.caret)
        } else {
            (self.caret, self.anchor)
        }
    }

    pub fn is_empty(self) -> bool {
        self.anchor == self.caret
    }
}

impl RichCanvas {
    pub fn new(size: Vec2) -> Self {
        Self {
            size,
            zoom: 1.0,
            background: Color32::from_rgb(232, 228, 222),
            page: PageFrame::new(size),
            boxes: Vec::new(),
            last_canvas_view: None,
            last_scroll_offset: Vec2::ZERO,
            zoom_scroll_samples: VecDeque::new(),
        }
    }

    pub fn push(&mut self, render_box: RenderBox) {
        self.boxes.push(render_box);
    }

    pub fn relayout(&mut self, mode: CanvasMode) {
        match mode {
            CanvasMode::TextDocument => self.layout_text_document(),
            CanvasMode::SlideDeck => self.layout_slide_deck(),
        }
    }

    fn layout_text_document(&mut self) {
        let page_margin = self.page.margin.x;
        let mut cursor_y = page_margin;

        for render_box in &mut self.boxes {
            if matches!(render_box.layout_role, LayoutRole::FlowBlock) {
                // Flow blocks are constrained by the page width and stacked vertically.
                render_box.position = Pos2::new(page_margin, cursor_y);
                render_box.size.x = (self.page.size.x - page_margin * 2.0).max(120.0);
                render_box.measure();
                cursor_y += render_box.size.y + render_box.style.outer_margin.y;
            } else {
                render_box.measure();
            }
        }

        self.page.size.y = self.page.size.y.max(cursor_y + page_margin);
        self.size = self.page.outer_size();
    }

    fn layout_slide_deck(&mut self) {
        for render_box in &mut self.boxes {
            render_box.measure();
            if let Some(animation) = &render_box.animation {
                // Slide mode keeps explicit positions and applies preview transforms on top.
                if animation.is_preview_oscillation() {
                    render_box.position += animation.preview_offset();
                }
            }
        }

        self.size = self.page.outer_size();
    }

    fn focused_zoom_scroll_offset(
        &self,
        pointer: Pos2,
        next_zoom: f32,
        current_scroll_offset: Vec2,
    ) -> Option<Vec2> {
        zoom_scroll_offset_for_pointer(
            self.size,
            self.zoom,
            next_zoom,
            self.last_canvas_view?,
            pointer,
            current_scroll_offset,
        )
    }

    fn smoothed_zoom_scroll(&mut self, now: f64, scroll_delta: f32) -> f32 {
        if self
            .zoom_scroll_samples
            .back()
            .is_some_and(|(last_time, last_delta)| {
                now - *last_time > ZOOM_SMOOTHING_RESET_SECONDS
                    || last_delta.signum() != scroll_delta.signum()
            })
        {
            self.zoom_scroll_samples.clear();
        }

        self.zoom_scroll_samples.push_back((now, scroll_delta));
        while self
            .zoom_scroll_samples
            .front()
            .is_some_and(|(time, _)| now - *time > ZOOM_SMOOTHING_WINDOW_SECONDS)
        {
            self.zoom_scroll_samples.pop_front();
        }

        let total: f32 = self
            .zoom_scroll_samples
            .iter()
            .map(|(_, delta)| *delta)
            .sum();
        total / self.zoom_scroll_samples.len().max(1) as f32
    }

    #[cfg(test)]
    fn set_last_canvas_view_for_test(&mut self, view: CanvasViewState) {
        self.last_canvas_view = Some(view);
    }

    pub fn ui(&mut self, ui: &mut Ui, selection: Option<CanvasSelection>) -> CanvasUiResponse {
        let scroll_area_id_salt = "rich_canvas_scroll_area";
        let (pointer_hover_pos, pointer_latest_pos, pointer_interact_pos) = ui.input(|input| {
            (
                input.pointer.hover_pos(),
                input.pointer.latest_pos(),
                input.pointer.interact_pos(),
            )
        });
        let ctrl_scroll = ui.input(|input| {
            if input.modifiers.ctrl {
                input.smooth_scroll_delta.y + input.raw_scroll_delta.y
            } else {
                0.0
            }
        });
        let input_time = ui.input(|input| input.time);

        if ctrl_scroll.abs() > f32::EPSILON {
            ui.input_mut(|input| {
                input.raw_scroll_delta = Vec2::ZERO;
                input.smooth_scroll_delta = Vec2::ZERO;
            });

            let smoothed_scroll = self.smoothed_zoom_scroll(input_time, ctrl_scroll);
            let previous_zoom = self.zoom;
            let next_zoom =
                (self.zoom * (smoothed_scroll * ZOOM_SCROLL_SCALE).exp()).clamp(0.4, 3.0);
            if (next_zoom - previous_zoom).abs() > f32::EPSILON {
                let previous_scroll_offset = self.last_scroll_offset;
                let pointer = pointer_hover_pos;
                let canvas_rect = self.last_canvas_view.map(|view| {
                    let canvas_origin = view.viewport_rect.min + view.canvas_content_origin
                        - previous_scroll_offset;
                    Rect::from_min_size(canvas_origin, view.canvas_size)
                });
                let rendered_canvas_origin = self.last_canvas_view.map(|view| view.canvas_origin);
                let pointer_over_canvas = pointer
                    .zip(canvas_rect)
                    .is_some_and(|(pointer, canvas_rect)| canvas_rect.contains(pointer));
                let computed_offset = pointer.and_then(|pointer| {
                    self.focused_zoom_scroll_offset(pointer, next_zoom, previous_scroll_offset)
                });
                let zoom_scroll_offset = computed_offset;
                write_zoom_telemetry(format!(
                    concat!(
                        "zoom_event ts_ms={} ctrl_scroll={:.3} zoom={:.4}->{:.4} ",
                        "smoothed_scroll={:.3} sample_count={} ",
                        "hover={:?} latest={:?} interact={:?} pointer_source=hover ",
                        "last_view={:?} rendered_canvas_origin={:?} canvas_rect={:?} pointer_over_canvas={} ",
                        "scroll_before={:?} computed_offset={:?} scroll_applied={} size={:?}"
                    ),
                    zoom_telemetry_timestamp_ms(),
                    ctrl_scroll,
                    previous_zoom,
                    next_zoom,
                    smoothed_scroll,
                    self.zoom_scroll_samples.len(),
                    pointer_hover_pos,
                    pointer_latest_pos,
                    pointer_interact_pos,
                    self.last_canvas_view,
                    rendered_canvas_origin,
                    canvas_rect,
                    pointer_over_canvas,
                    previous_scroll_offset,
                    computed_offset,
                    computed_offset.is_some(),
                    self.size,
                ));
                self.zoom = next_zoom;
                return self.show_canvas_ui(ui, selection, scroll_area_id_salt, zoom_scroll_offset);
            } else {
                write_zoom_telemetry(format!(
                    concat!(
                        "zoom_event ts_ms={} ctrl_scroll={:.3} zoom={:.4}->{:.4} ",
                        "smoothed_scroll={:.3} sample_count={} clamped=true hover={:?} latest={:?} interact={:?}"
                    ),
                    zoom_telemetry_timestamp_ms(),
                    ctrl_scroll,
                    previous_zoom,
                    next_zoom,
                    smoothed_scroll,
                    self.zoom_scroll_samples.len(),
                    pointer_hover_pos,
                    pointer_latest_pos,
                    pointer_interact_pos,
                ));
            }
        }

        self.show_canvas_ui(ui, selection, scroll_area_id_salt, None)
    }

    pub fn paint_page_preview(&mut self, painter: &Painter, target_rect: Rect) {
        painter.rect_filled(target_rect, 2.0, self.background);
        let scale = (target_rect.width() / self.page.size.x)
            .min(target_rect.height() / self.page.size.y)
            .max(0.001);
        let page_size = self.page.size * scale;
        let page_rect = Rect::from_center_size(target_rect.center(), page_size);

        let clipped = painter.with_clip_rect(target_rect);
        clipped.rect_filled(page_rect, self.page.corner_radius * scale, self.page.fill);
        clipped.rect_stroke(
            page_rect,
            self.page.corner_radius * scale,
            Stroke::new((1.0 * scale).clamp(0.5, 1.0), self.page.stroke),
            StrokeKind::Inside,
        );

        let mut visible_indices: Vec<usize> = self
            .boxes
            .iter()
            .enumerate()
            .filter_map(|(index, item)| item.visible.then_some(index))
            .collect();
        visible_indices.sort_by_key(|index| self.boxes[*index].z_index);

        for index in visible_indices {
            self.boxes[index].paint(&clipped, page_rect.min, scale);
        }
    }

    fn show_canvas_ui(
        &mut self,
        ui: &mut Ui,
        selection: Option<CanvasSelection>,
        scroll_area_id_salt: &'static str,
        zoom_scroll_offset: Option<Vec2>,
    ) -> CanvasUiResponse {
        let mut scroll_area = egui::ScrollArea::both()
            .id_salt(scroll_area_id_salt)
            .auto_shrink([false; 2]);
        if let Some(offset) = zoom_scroll_offset {
            scroll_area = scroll_area.scroll_offset(offset);
        }

        let output = scroll_area.show_viewport(ui, |ui, viewport| {
            // The allocated painter covers at least the viewport. If the canvas is smaller
            // than the viewport, keep it centered without changing canvas coordinates.
            let canvas_size = self.size * self.zoom;
            let available_size = ui.available_size();
            let allocated_size = vec2(
                canvas_size.x.max(available_size.x),
                canvas_size.y.max(available_size.y),
            );
            let (response, painter) = ui.allocate_painter(allocated_size, Sense::click_and_drag());
            painter.rect_filled(response.rect, 0.0, self.background);
            let extra_space = (allocated_size - canvas_size).max(Vec2::ZERO);
            let canvas_content_origin = extra_space * 0.5;
            let canvas_origin = response.rect.min + canvas_content_origin - viewport.min.to_vec2();

            self.paint(&painter, canvas_origin);
            if let Some(selection) = selection {
                self.paint_selection(&painter, canvas_origin, selection);
            }

            let (clicked_box, clicked_caret) = if response.clicked() {
                let click = response.interact_pointer_pos();
                let clicked_box = click.and_then(|pointer| self.hit_test(canvas_origin, pointer));
                let clicked_caret = clicked_box.and_then(|id| {
                    click.and_then(|pointer| {
                        self.text_caret_index_at(id, canvas_origin, pointer, &painter)
                    })
                });
                (clicked_box, clicked_caret)
            } else {
                (None, None)
            };

            let hovered_hyperlink = response.hover_pos().and_then(|pointer| {
                let box_id = self.hit_test(canvas_origin, pointer)?;
                let char_index =
                    self.text_caret_index_at(box_id, canvas_origin, pointer, &painter)?;
                self.text_hyperlink_at(box_id, char_index)
                    .map(|url| CanvasHyperlink {
                        box_id,
                        char_index,
                        url,
                    })
            });
            let pointer = response.interact_pointer_pos();
            let image_resize_handle = selection.and_then(|selection| {
                pointer.and_then(|pos| {
                    self.hit_test_image_resize_handle(canvas_origin, selection, pos)
                })
            });
            let text_resize_handle = selection.and_then(|selection| {
                pointer
                    .and_then(|pos| self.hit_test_text_resize_handle(canvas_origin, selection, pos))
            });
            let dragged_box =
                if response.drag_started() || response.dragged() || response.drag_stopped() {
                    pointer.and_then(|pos| self.hit_test(canvas_origin, pos))
                } else {
                    None
                };
            let dragged_caret = dragged_box.and_then(|id| {
                pointer.and_then(|pos| self.text_caret_index_at(id, canvas_origin, pos, &painter))
            });

            let drag_started = response.drag_started();
            let drag_stopped = response.drag_stopped();
            (
                CanvasUiResponse {
                    response,
                    clicked_box,
                    clicked_caret,
                    dragged_box,
                    dragged_caret,
                    hovered_hyperlink,
                    image_resize_handle,
                    text_resize_handle,
                    pointer_pos: pointer,
                    drag_started,
                    drag_stopped,
                },
                canvas_origin,
                canvas_size,
                canvas_content_origin,
            )
        });
        let (response, canvas_origin, canvas_size, canvas_content_origin) = output.inner;
        self.last_canvas_view = Some(CanvasViewState {
            viewport_rect: output.inner_rect,
            canvas_origin,
            canvas_size,
            canvas_content_origin,
        });
        self.last_scroll_offset = output.state.offset;
        response
    }

    pub fn paint(&mut self, painter: &Painter, origin: Pos2) {
        let canvas_rect = Rect::from_min_size(origin, self.size * self.zoom);
        painter.rect_filled(canvas_rect, 12.0, self.background);
        let page_rect = self.page.rect(origin, self.zoom);
        let shadow_rect = page_rect.translate(vec2(8.0, 10.0) * self.zoom);
        painter.rect_filled(
            shadow_rect,
            self.page.corner_radius * self.zoom.clamp(0.5, 1.5),
            Color32::from_black_alpha(18),
        );
        painter.rect_filled(
            page_rect,
            self.page.corner_radius * self.zoom.clamp(0.5, 1.5),
            self.page.fill,
        );
        painter.rect_stroke(
            page_rect,
            self.page.corner_radius * self.zoom.clamp(0.5, 1.5),
            Stroke::new(self.zoom.clamp(1.0, 2.0), self.page.stroke),
            StrokeKind::Inside,
        );

        let clip = painter.clip_rect();
        let mut box_indices: Vec<usize> = self
            .boxes
            .iter()
            .enumerate()
            .filter_map(|(index, item)| item.visible.then_some(index))
            .collect();
        box_indices.sort_by_key(|index| self.boxes[*index].z_index);
        let content_origin = self.page.content_origin(origin, self.zoom);

        for index in box_indices {
            let render_box = &mut self.boxes[index];
            let rect = render_box.rect(content_origin, self.zoom);
            // Cull against the current clip rect so large documents only paint what is visible.
            if rect.intersects(clip.expand(48.0)) {
                render_box.paint(painter, content_origin, self.zoom);
            }
        }
    }

    fn paint_selection(&self, painter: &Painter, origin: Pos2, selection: CanvasSelection) {
        let Some(render_box) = self.boxes.iter().find(|item| item.id == selection.box_id) else {
            return;
        };

        let rect = render_box.rect(self.page.content_origin(origin, self.zoom), self.zoom);
        if render_box.is_image() {
            paint_image_selection_overlay(painter, rect, render_box.rotation);
            return;
        }

        if render_box.is_text() {
            paint_text_selection_overlay(painter, rect);
        } else {
            painter.rect_stroke(
                rect.expand(4.0 * self.zoom.clamp(0.5, 1.5)),
                (render_box.style.corner_radius + 4.0) * self.zoom.clamp(0.5, 1.5),
                Stroke::new(
                    (2.0 * self.zoom).clamp(1.5, 3.0),
                    Color32::from_rgb(181, 82, 38),
                ),
                StrokeKind::Outside,
            );
        }

        if let Some(range) = selection.range.filter(|range| !range.is_empty()) {
            if let Some(rects) = render_box.selection_rects(
                self.page.content_origin(origin, self.zoom),
                self.zoom,
                range,
                painter,
            ) {
                for rect in rects {
                    painter.rect_filled(
                        rect,
                        4.0 * self.zoom.clamp(0.5, 1.0),
                        Color32::from_rgba_unmultiplied(181, 82, 38, 56),
                    );
                }
            }
        }

        if let Some(caret_index) = selection.caret_index {
            if selection.range.is_none_or(TextRange::is_empty) {
                if let Some(caret_rect) = render_box.caret_rect(
                    self.page.content_origin(origin, self.zoom),
                    self.zoom,
                    caret_index,
                    painter,
                ) {
                    painter.rect_filled(caret_rect, 1.0, Color32::from_rgb(36, 38, 41));
                }
            }
        }
    }

    fn hit_test_image_resize_handle(
        &self,
        origin: Pos2,
        selection: CanvasSelection,
        pointer_pos: Pos2,
    ) -> Option<ImageResizeHandle> {
        let render_box = self
            .boxes
            .iter()
            .find(|item| item.id == selection.box_id && item.visible && item.is_image())?;

        let rect = render_box.rect(self.page.content_origin(origin, self.zoom), self.zoom);
        image_handle_points(rect, render_box.rotation)
            .into_iter()
            .find_map(|(handle, point)| {
                let hit_radius = if handle == ImageResizeHandle::Center {
                    IMAGE_SELECTION_CENTER_HANDLE_HIT_RADIUS
                } else {
                    IMAGE_SELECTION_HANDLE_HIT_RADIUS
                };
                (point.distance(pointer_pos) <= hit_radius).then_some(handle)
            })
    }

    fn hit_test_text_resize_handle(
        &self,
        origin: Pos2,
        selection: CanvasSelection,
        pointer_pos: Pos2,
    ) -> Option<ImageResizeHandle> {
        let render_box = self
            .boxes
            .iter()
            .find(|item| item.id == selection.box_id && item.visible && item.is_text())?;
        if render_box.rotation.abs() > f32::EPSILON {
            return None;
        }

        let rect = render_box.rect(self.page.content_origin(origin, self.zoom), self.zoom);
        text_box_handle_points(rect)
            .into_iter()
            .find_map(|(handle, point, shape)| match shape {
                TextBoxHandleShape::Square => {
                    let hit_rect = Rect::from_center_size(
                        point,
                        Vec2::splat(IMAGE_SELECTION_HANDLE_HIT_RADIUS * 2.0),
                    );
                    hit_rect.contains(pointer_pos).then_some(handle)
                }
                TextBoxHandleShape::Circle => (point.distance(pointer_pos)
                    <= IMAGE_SELECTION_HANDLE_HIT_RADIUS)
                    .then_some(handle),
            })
    }

    fn hit_test(&self, origin: Pos2, pointer_pos: Pos2) -> Option<u64> {
        let content_origin = self.page.content_origin(origin, self.zoom);
        let mut visible_boxes: Vec<&RenderBox> =
            self.boxes.iter().filter(|item| item.visible).collect();
        visible_boxes.sort_by_key(|item| item.z_index);

        visible_boxes
            .into_iter()
            .rev()
            .find(|item| self.hit_test_render_box(item, content_origin, pointer_pos))
            .map(|item| item.id)
    }

    fn hit_test_render_box(&self, render_box: &RenderBox, origin: Pos2, pointer_pos: Pos2) -> bool {
        let rect = render_box.rect(origin, self.zoom);
        if render_box.is_image() && render_box.rotation.abs() > f32::EPSILON {
            rotated_rect_contains(rect, render_box.rotation, pointer_pos)
        } else {
            rect.contains(pointer_pos)
        }
    }

    pub fn box_mut(&mut self, id: u64) -> Option<&mut RenderBox> {
        self.boxes.iter_mut().find(|item| item.id == id)
    }

    pub fn box_ref(&self, id: u64) -> Option<&RenderBox> {
        self.boxes.iter().find(|item| item.id == id)
    }

    pub fn text_caret_index_at(
        &self,
        id: u64,
        origin: Pos2,
        pointer_pos: Pos2,
        painter: &Painter,
    ) -> Option<usize> {
        self.box_ref(id).and_then(|render_box| {
            render_box.caret_index_at(
                self.page.content_origin(origin, self.zoom),
                self.zoom,
                pointer_pos,
                painter,
            )
        })
    }

    pub fn text_hyperlink_at(&self, id: u64, char_index: usize) -> Option<String> {
        self.box_ref(id)
            .and_then(|render_box| render_box.hyperlink_at(char_index))
    }
}

fn image_resize_handle_points(image_rect: Rect) -> [(ImageResizeHandle, Pos2); 9] {
    image_handle_points(image_rect, 0.0)
}

fn image_handle_points(image_rect: Rect, rotation_degrees: f32) -> [(ImageResizeHandle, Pos2); 9] {
    let center = image_rect.center();
    let rotate = |point: Pos2| rotate_point(point, center, rotation_degrees);
    [
        (ImageResizeHandle::Center, image_rect.center()),
        (ImageResizeHandle::TopLeft, rotate(image_rect.left_top())),
        (ImageResizeHandle::Top, rotate(image_rect.center_top())),
        (ImageResizeHandle::TopRight, rotate(image_rect.right_top())),
        (ImageResizeHandle::Right, rotate(image_rect.right_center())),
        (
            ImageResizeHandle::BottomRight,
            rotate(image_rect.right_bottom()),
        ),
        (
            ImageResizeHandle::Bottom,
            rotate(image_rect.center_bottom()),
        ),
        (
            ImageResizeHandle::BottomLeft,
            rotate(image_rect.left_bottom()),
        ),
        (ImageResizeHandle::Left, rotate(image_rect.left_center())),
    ]
}

fn rotated_rect_corners(rect: Rect, rotation_degrees: f32) -> [Pos2; 4] {
    let center = rect.center();
    [
        rotate_point(rect.left_top(), center, rotation_degrees),
        rotate_point(rect.right_top(), center, rotation_degrees),
        rotate_point(rect.right_bottom(), center, rotation_degrees),
        rotate_point(rect.left_bottom(), center, rotation_degrees),
    ]
}

fn rotate_point(point: Pos2, center: Pos2, rotation_degrees: f32) -> Pos2 {
    let angle = rotation_degrees.to_radians();
    let (sin, cos) = angle.sin_cos();
    let delta = point - center;
    Pos2::new(
        center.x + delta.x * cos - delta.y * sin,
        center.y + delta.x * sin + delta.y * cos,
    )
}

fn rotated_rect_contains(rect: Rect, rotation_degrees: f32, pointer_pos: Pos2) -> bool {
    let local_pointer = rotate_point(pointer_pos, rect.center(), -rotation_degrees);
    rect.contains(local_pointer)
}

fn paint_image_selection_overlay(painter: &Painter, image_rect: Rect, rotation_degrees: f32) {
    let stroke_color = Color32::from_rgb(20, 120, 220);
    if rotation_degrees.abs() > f32::EPSILON {
        paint_rotated_box_handle_overlay(painter, image_rect, rotation_degrees, stroke_color);
    } else {
        paint_box_handle_overlay(painter, image_rect, stroke_color);
    }
}

fn paint_text_selection_overlay(painter: &Painter, text_rect: Rect) {
    let stroke_color = Color32::from_rgb(205, 54, 54);
    painter.rect_stroke(
        text_rect,
        0.0,
        Stroke::new(IMAGE_SELECTION_STROKE_WIDTH, stroke_color),
        StrokeKind::Outside,
    );

    for (_handle, handle_point, shape) in text_box_handle_points(text_rect) {
        match shape {
            TextBoxHandleShape::Square => {
                let handle_rect =
                    Rect::from_center_size(handle_point, Vec2::splat(TEXT_MOVE_HANDLE_SIZE));
                painter.rect_filled(handle_rect, 0.0, Color32::WHITE);
                painter.rect_stroke(
                    handle_rect,
                    0.0,
                    Stroke::new(1.4, stroke_color),
                    StrokeKind::Outside,
                );
            }
            TextBoxHandleShape::Circle => {
                painter.circle_filled(handle_point, IMAGE_SELECTION_HANDLE_RADIUS, Color32::WHITE);
                painter.circle_stroke(
                    handle_point,
                    IMAGE_SELECTION_HANDLE_RADIUS,
                    Stroke::new(1.4, stroke_color),
                );
            }
        }
    }
}

fn paint_box_handle_overlay(painter: &Painter, rect: Rect, stroke_color: Color32) {
    painter.rect_stroke(
        rect,
        0.0,
        Stroke::new(IMAGE_SELECTION_STROKE_WIDTH, stroke_color),
        StrokeKind::Outside,
    );

    for (handle, handle_point) in image_resize_handle_points(rect) {
        paint_image_resize_handle(painter, handle, handle_point, stroke_color);
    }
}

fn paint_rotated_box_handle_overlay(
    painter: &Painter,
    rect: Rect,
    rotation_degrees: f32,
    stroke_color: Color32,
) {
    let corners = rotated_rect_corners(rect, rotation_degrees);
    for index in 0..corners.len() {
        painter.line_segment(
            [corners[index], corners[(index + 1) % corners.len()]],
            Stroke::new(IMAGE_SELECTION_STROKE_WIDTH, stroke_color),
        );
    }

    for (handle, handle_point) in image_handle_points(rect, rotation_degrees) {
        paint_image_resize_handle(painter, handle, handle_point, stroke_color);
    }
}

fn paint_image_resize_handle(
    painter: &Painter,
    handle: ImageResizeHandle,
    center: Pos2,
    stroke_color: Color32,
) {
    if matches!(
        handle,
        ImageResizeHandle::TopLeft | ImageResizeHandle::BottomRight
    ) {
        let radius = IMAGE_ASPECT_HANDLE_RADIUS;
        let points = vec![
            Pos2::new(center.x, center.y - radius),
            Pos2::new(center.x + radius, center.y),
            Pos2::new(center.x, center.y + radius),
            Pos2::new(center.x - radius, center.y),
        ];
        painter.add(Shape::convex_polygon(
            points,
            Color32::WHITE,
            Stroke::new(1.4, stroke_color),
        ));
    } else {
        painter.circle_filled(center, IMAGE_SELECTION_HANDLE_RADIUS, Color32::WHITE);
        painter.circle_stroke(
            center,
            IMAGE_SELECTION_HANDLE_RADIUS,
            Stroke::new(1.4, stroke_color),
        );
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum TextBoxHandleShape {
    Square,
    Circle,
}

fn text_box_handle_points(text_rect: Rect) -> [(ImageResizeHandle, Pos2, TextBoxHandleShape); 8] {
    [
        (
            ImageResizeHandle::Center,
            text_rect.left_top(),
            TextBoxHandleShape::Square,
        ),
        (
            ImageResizeHandle::Top,
            text_rect.center_top(),
            TextBoxHandleShape::Circle,
        ),
        (
            ImageResizeHandle::TopRight,
            text_rect.right_top(),
            TextBoxHandleShape::Circle,
        ),
        (
            ImageResizeHandle::Right,
            text_rect.right_center(),
            TextBoxHandleShape::Circle,
        ),
        (
            ImageResizeHandle::BottomRight,
            text_rect.right_bottom(),
            TextBoxHandleShape::Circle,
        ),
        (
            ImageResizeHandle::Bottom,
            text_rect.center_bottom(),
            TextBoxHandleShape::Circle,
        ),
        (
            ImageResizeHandle::BottomLeft,
            text_rect.left_bottom(),
            TextBoxHandleShape::Circle,
        ),
        (
            ImageResizeHandle::Left,
            text_rect.left_center(),
            TextBoxHandleShape::Circle,
        ),
    ]
}

pub fn configure_editor_fonts(ctx: &egui::Context) {
    let Some(regular_font) = load_font_bytes(&[
        concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../fonts/LiberationSans-Regular.ttf"
        ),
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]) else {
        return;
    };
    let Some(bold_font) = load_font_bytes(&[
        concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../fonts/LiberationSans-Bold.ttf"
        ),
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
        "/usr/share/fonts/truetype/noto/NotoSans-Bold.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    ]) else {
        return;
    };

    let mut fonts = egui::FontDefinitions::default();
    fonts.font_data.insert(
        EDITOR_REGULAR_FONT_NAME.into(),
        Arc::new(egui::FontData::from_owned(regular_font)),
    );
    fonts.font_data.insert(
        EDITOR_BOLD_FONT_NAME.into(),
        Arc::new(egui::FontData::from_owned(bold_font)),
    );
    fonts.families.insert(
        FontFamily::Name(EDITOR_REGULAR_FONT_NAME.into()),
        vec![EDITOR_REGULAR_FONT_NAME.into()],
    );
    fonts.families.insert(
        FontFamily::Name(EDITOR_BOLD_FONT_NAME.into()),
        vec![EDITOR_BOLD_FONT_NAME.into()],
    );
    ctx.set_fonts(fonts);
}

fn load_font_bytes(candidates: &[&str]) -> Option<Vec<u8>> {
    for path in candidates {
        if let Ok(bytes) = std::fs::read(path) {
            return Some(bytes);
        }
    }
    None
}

#[derive(Clone, Debug)]
pub struct PageFrame {
    pub size: Vec2,
    pub margin: Vec2,
    pub fill: Color32,
    pub stroke: Color32,
    pub corner_radius: f32,
    pub page_offset: Vec2,
}

impl PageFrame {
    pub fn new(size: Vec2) -> Self {
        Self {
            size,
            margin: vec2(36.0, 36.0),
            fill: Color32::WHITE,
            stroke: Color32::from_rgb(210, 214, 220),
            corner_radius: 8.0,
            page_offset: vec2(48.0, 48.0),
        }
    }

    pub fn rect(&self, origin: Pos2, zoom: f32) -> Rect {
        Rect::from_min_size(origin + self.page_offset * zoom, self.size * zoom)
    }

    pub fn content_origin(&self, origin: Pos2, zoom: f32) -> Pos2 {
        // Child box coordinates are page-relative, not canvas-relative.
        origin + self.page_offset * zoom
    }

    pub fn outer_size(&self) -> Vec2 {
        self.size + self.page_offset * 2.0 + vec2(24.0, 24.0)
    }
}

#[derive(Clone, Debug)]
pub struct RenderBox {
    pub id: u64,
    pub kind: RenderBoxKind,
    pub layout_role: LayoutRole,
    pub position: Pos2,
    pub size: Vec2,
    pub authored_size: Option<Vec2>,
    pub lock_size: bool,
    pub scale: Vec2,
    pub rotation: f32,
    pub visible: bool,
    pub z_index: i32,
    pub style: BoxStyle,
    pub animation: Option<AnimationSpec>,
}

impl RenderBox {
    pub fn text(id: u64, role: LayoutRole, runs: Vec<TextRun>) -> Self {
        let mut style = BoxStyle::default();
        style.fill = Color32::TRANSPARENT;
        style.stroke = Color32::TRANSPARENT;

        Self {
            id,
            kind: RenderBoxKind::Text(RichTextBlock {
                runs,
                alignment: TextAlignment::Left,
                vertical_alignment: TextVerticalAlignment::Top,
            }),
            layout_role: role,
            position: Pos2::ZERO,
            size: vec2(320.0, 80.0),
            authored_size: Some(vec2(320.0, 80.0)),
            lock_size: false,
            scale: vec2(1.0, 1.0),
            rotation: 0.0,
            visible: true,
            z_index: 0,
            style,
            animation: None,
        }
    }

    pub fn image(id: u64, role: LayoutRole, label: impl Into<String>, size: Vec2) -> Self {
        let mut style = BoxStyle::default();
        style.padding = Vec2::ZERO;
        Self {
            id,
            kind: RenderBoxKind::Image(ImageBlock::placeholder(label, size)),
            layout_role: role,
            position: Pos2::ZERO,
            size,
            authored_size: Some(size),
            lock_size: false,
            scale: vec2(1.0, 1.0),
            rotation: 0.0,
            visible: true,
            z_index: 0,
            style,
            animation: None,
        }
    }

    pub fn image_from_path(
        id: u64,
        role: LayoutRole,
        path: impl AsRef<Path>,
    ) -> Result<Self, image::ImageError> {
        let image = ImageBlock::load_from_path(path)?;
        let size = image.size;
        let mut style = BoxStyle::default();
        style.padding = Vec2::ZERO;
        Ok(Self {
            id,
            kind: RenderBoxKind::Image(image),
            layout_role: role,
            position: Pos2::ZERO,
            size,
            authored_size: Some(size),
            lock_size: false,
            scale: vec2(1.0, 1.0),
            rotation: 0.0,
            visible: true,
            z_index: 0,
            style,
            animation: None,
        })
    }

    pub fn table(id: u64, role: LayoutRole, table: TableBlock) -> Self {
        Self {
            id,
            kind: RenderBoxKind::Table(table),
            layout_role: role,
            position: Pos2::ZERO,
            size: vec2(360.0, 200.0),
            authored_size: Some(vec2(360.0, 200.0)),
            lock_size: false,
            scale: vec2(1.0, 1.0),
            rotation: 0.0,
            visible: true,
            z_index: 0,
            style: BoxStyle::default(),
            animation: None,
        }
    }

    pub fn is_text(&self) -> bool {
        matches!(self.kind, RenderBoxKind::Text(_))
    }

    pub fn is_image(&self) -> bool {
        matches!(self.kind, RenderBoxKind::Image(_))
    }

    pub fn image_ref(&self) -> Option<&ImageBlock> {
        match &self.kind {
            RenderBoxKind::Image(block) => Some(block),
            _ => None,
        }
    }

    pub fn image_mut(&mut self) -> Option<&mut ImageBlock> {
        match &mut self.kind {
            RenderBoxKind::Image(block) => Some(block),
            _ => None,
        }
    }

    pub fn plain_text(&self) -> Option<String> {
        match &self.kind {
            RenderBoxKind::Text(block) => Some(block.plain_text()),
            _ => None,
        }
    }

    pub fn set_plain_text(&mut self, text: String) -> bool {
        match &mut self.kind {
            RenderBoxKind::Text(block) => {
                block.set_plain_text(text);
                true
            }
            _ => false,
        }
    }

    pub fn text_style_state(&self) -> Option<TextStyleState> {
        match &self.kind {
            RenderBoxKind::Text(block) => Some(block.style_state()),
            _ => None,
        }
    }

    pub fn text_style_state_in_range(&self, range: TextRange) -> Option<TextStyleState> {
        match &self.kind {
            RenderBoxKind::Text(block) => Some(block.style_state_in_range(range)),
            _ => None,
        }
    }

    pub fn toggle_bold(&mut self, range: Option<TextRange>) -> bool {
        match &mut self.kind {
            RenderBoxKind::Text(block) => {
                block.toggle_bold(range);
                true
            }
            _ => false,
        }
    }

    pub fn toggle_italic(&mut self, range: Option<TextRange>) -> bool {
        match &mut self.kind {
            RenderBoxKind::Text(block) => {
                block.toggle_italic(range);
                true
            }
            _ => false,
        }
    }

    pub fn toggle_underline(&mut self, range: Option<TextRange>) -> bool {
        match &mut self.kind {
            RenderBoxKind::Text(block) => {
                block.toggle_underline(range);
                true
            }
            _ => false,
        }
    }

    pub fn set_text_color(&mut self, range: Option<TextRange>, color: Color32) -> bool {
        match &mut self.kind {
            RenderBoxKind::Text(block) => {
                block.set_text_color(range, color);
                true
            }
            _ => false,
        }
    }

    pub fn set_highlight_color(
        &mut self,
        range: Option<TextRange>,
        color: Option<Color32>,
    ) -> bool {
        match &mut self.kind {
            RenderBoxKind::Text(block) => {
                block.set_highlight_color(range, color);
                true
            }
            _ => false,
        }
    }

    pub fn set_text_font_size(&mut self, range: Option<TextRange>, font_size: f32) -> bool {
        match &mut self.kind {
            RenderBoxKind::Text(block) => {
                block.set_font_size(range, font_size);
                true
            }
            _ => false,
        }
    }

    pub fn set_text_alignment(&mut self, alignment: TextAlignment) -> bool {
        match &mut self.kind {
            RenderBoxKind::Text(block) => {
                block.alignment = alignment;
                true
            }
            _ => false,
        }
    }

    pub fn set_text_vertical_alignment(&mut self, alignment: TextVerticalAlignment) -> bool {
        match &mut self.kind {
            RenderBoxKind::Text(block) => {
                block.vertical_alignment = alignment;
                true
            }
            _ => false,
        }
    }

    pub fn text_len(&self) -> Option<usize> {
        match &self.kind {
            RenderBoxKind::Text(block) => Some(block.char_len()),
            _ => None,
        }
    }

    pub fn insert_text_at(&mut self, char_index: usize, text: &str) -> Option<usize> {
        match &mut self.kind {
            RenderBoxKind::Text(block) => Some(block.insert_text_at(char_index, text)),
            _ => None,
        }
    }

    pub fn delete_char_before(&mut self, char_index: usize) -> Option<usize> {
        match &mut self.kind {
            RenderBoxKind::Text(block) => block.delete_char_before(char_index),
            _ => None,
        }
    }

    pub fn delete_char_at(&mut self, char_index: usize) -> Option<usize> {
        match &mut self.kind {
            RenderBoxKind::Text(block) => block.delete_char_at(char_index),
            _ => None,
        }
    }

    pub fn delete_range(&mut self, range: TextRange) -> Option<usize> {
        match &mut self.kind {
            RenderBoxKind::Text(block) => block.delete_range(range),
            _ => None,
        }
    }

    pub fn replace_range(&mut self, range: TextRange, text: &str) -> Option<usize> {
        match &mut self.kind {
            RenderBoxKind::Text(block) => Some(block.replace_range(range, text)),
            _ => None,
        }
    }

    pub fn move_caret_horizontal(
        &self,
        char_index: usize,
        delta: isize,
        ctx: &egui::Context,
        _zoom: f32,
    ) -> Option<usize> {
        let galley = self.text_galley(ctx)?;
        let cursor = CCursor::new(char_index);
        let next = if delta < 0 {
            galley.cursor_left_one_character(&cursor)
        } else {
            galley.cursor_right_one_character(&cursor)
        };
        Some(next.index)
    }

    pub fn move_caret_vertical(
        &self,
        char_index: usize,
        direction: isize,
        ctx: &egui::Context,
        _zoom: f32,
        desired_x: Option<f32>,
    ) -> Option<(usize, Option<f32>)> {
        let galley = self.text_galley(ctx)?;
        let cursor = CCursor::new(char_index);
        let (next, next_x) = if direction < 0 {
            galley.cursor_up_one_row(&cursor, desired_x)
        } else {
            galley.cursor_down_one_row(&cursor, desired_x)
        };
        Some((next.index, next_x))
    }

    pub fn line_boundary(
        &self,
        char_index: usize,
        to_end: bool,
        ctx: &egui::Context,
        _zoom: f32,
    ) -> Option<usize> {
        let galley = self.text_galley(ctx)?;
        let cursor = CCursor::new(char_index);
        let next = if to_end {
            galley.cursor_end_of_row(&cursor)
        } else {
            galley.cursor_begin_of_row(&cursor)
        };
        Some(next.index)
    }

    pub fn measure(&mut self) {
        match &self.kind {
            RenderBoxKind::Text(block) => {
                if self.lock_size {
                    if let Some(size) = self.authored_size {
                        self.size = size;
                    }
                    return;
                }
                // Text height is estimated from wrapped line count rather than real font metrics.
                let content_width = (self.size.x - self.style.padding.x * 2.0).max(80.0);
                let layout = block.layout_lines(content_width);
                let measured_height = (layout.total_height + self.style.padding.y * 2.0).max(48.0);
                self.size.y = self.size.y.max(measured_height);
            }
            RenderBoxKind::Image(block) => {
                self.size = block.size + self.style.padding * 2.0;
            }
            RenderBoxKind::Table(block) => {
                self.size = vec2(
                    block.column_count() as f32 * 120.0 + self.style.padding.x * 2.0,
                    block.row_count() as f32 * 36.0 + self.style.padding.y * 2.0,
                );
            }
            RenderBoxKind::Container => {}
        }
    }

    pub fn rect(&self, origin: Pos2, zoom: f32) -> Rect {
        let min = origin + self.position.to_vec2() * zoom;
        Rect::from_min_size(
            min,
            vec2(self.size.x * self.scale.x, self.size.y * self.scale.y) * zoom,
        )
    }

    pub fn paint(&mut self, painter: &Painter, origin: Pos2, zoom: f32) {
        let rect = self.rect(origin, zoom);
        let radius = self.style.corner_radius * zoom.clamp(0.5, 1.5);

        match &mut self.kind {
            RenderBoxKind::Image(block) if self.rotation.abs() > f32::EPSILON => {
                // Rotation is currently implemented only for image placeholders.
                block.paint_rotated(painter, rect, &self.style, zoom, self.rotation, self.id);
            }
            RenderBoxKind::Text(block) => {
                paint_box_background_and_stroke(painter, rect, radius, &self.style, zoom);
                block.paint(painter, rect, &self.style, zoom)
            }
            RenderBoxKind::Image(block) => {
                paint_box_background_and_stroke(painter, rect, radius, &self.style, zoom);
                block.paint(painter, rect, self.id)
            }
            RenderBoxKind::Table(block) => {
                paint_box_background_and_stroke(painter, rect, radius, &self.style, zoom);
                block.paint(painter, rect, &self.style, zoom)
            }
            RenderBoxKind::Container => {}
        }
    }

    pub fn caret_rect(
        &self,
        origin: Pos2,
        zoom: f32,
        char_index: usize,
        painter: &Painter,
    ) -> Option<Rect> {
        let galley = self.text_galley_from_painter(painter)?;
        let rect = self.rect(origin, zoom);
        if let RenderBoxKind::Text(block) = &self.kind {
            let text_origin = block.text_origin(rect, &self.style, zoom, galley.size().y * zoom);
            let row =
                block.painted_row_for_index(&galley, char_index, zoom, painter, text_origin)?;
            let cursor_rect = row.galley.pos_from_cursor(CCursor::new(row.local_index));
            let min = row.origin + cursor_rect.min.to_vec2();
            let max = row.origin + cursor_rect.max.to_vec2();
            let height = (max.y - min.y).max(8.0 * zoom);
            return Some(Rect::from_min_size(
                Pos2::new(min.x - 0.75 * zoom, min.y),
                vec2(1.5 * zoom, height),
            ));
        };
        let text_origin = rect.min + self.style.padding * zoom;
        let cursor_rect = galley.pos_from_cursor(CCursor::new(char_index));
        let min = text_origin + cursor_rect.min.to_vec2() * zoom;
        let max = text_origin + cursor_rect.max.to_vec2() * zoom;
        let height = (max.y - min.y).max(8.0 * zoom);
        Some(Rect::from_min_size(
            Pos2::new(min.x - 0.75 * zoom, min.y),
            vec2(1.5 * zoom, height),
        ))
    }

    pub fn selection_rects(
        &self,
        origin: Pos2,
        zoom: f32,
        range: TextRange,
        painter: &Painter,
    ) -> Option<Vec<Rect>> {
        let galley = self.text_galley_from_painter(painter)?;
        let rect = self.rect(origin, zoom);
        let text_origin = match &self.kind {
            RenderBoxKind::Text(block) => {
                block.text_origin(rect, &self.style, zoom, galley.size().y * zoom)
            }
            _ => rect.min + self.style.padding * zoom,
        };
        let (start, end) = range.sorted();
        if start == end {
            return Some(Vec::new());
        }

        let mut rects = Vec::new();
        let mut row_start = 0usize;
        for row in &galley.rows {
            let row_end = row_start + row.char_count_including_newline();
            let row_content_end = row_start + row.char_count_excluding_newline();
            let sel_start = start.max(row_start).min(row_content_end);
            let sel_end = end.max(row_start).min(row_content_end);
            if sel_start < sel_end {
                let local_start = sel_start - row_start;
                let local_end = sel_end - row_start;
                if let RenderBoxKind::Text(block) = &self.kind {
                    let painted_row =
                        block.painted_row(&galley, row, row_start, zoom, painter, text_origin);
                    if let Some(painted_row) = painted_row {
                        let painted = &painted_row.galley.rows[0];
                        let x0 = painted.x_offset(local_start);
                        let x1 = painted.x_offset(local_end);
                        let min = painted_row.origin + vec2(x0, painted.min_y());
                        let max = painted_row.origin + vec2(x1, painted.max_y());
                        rects.push(Rect::from_min_max(min, max));
                    }
                } else {
                    let x0 = row.pos.x + row.x_offset(local_start);
                    let x1 = row.pos.x + row.x_offset(local_end);
                    let min = text_origin + vec2(x0, row.min_y()) * zoom;
                    let max = text_origin + vec2(x1, row.max_y()) * zoom;
                    rects.push(Rect::from_min_max(min, max));
                }
            }
            row_start = row_end;
        }

        Some(rects)
    }

    pub fn caret_index_at(
        &self,
        origin: Pos2,
        zoom: f32,
        pointer_pos: Pos2,
        painter: &Painter,
    ) -> Option<usize> {
        let galley = self.text_galley_from_painter(painter)?;
        let rect = self.rect(origin, zoom);
        if !rect.contains(pointer_pos) {
            return None;
        }

        let text_origin = match &self.kind {
            RenderBoxKind::Text(block) => {
                block.text_origin(rect, &self.style, zoom, galley.size().y * zoom)
            }
            _ => rect.min + self.style.padding * zoom,
        };
        if let RenderBoxKind::Text(block) = &self.kind {
            return block.painted_caret_index_at(&galley, pointer_pos, zoom, painter, text_origin);
        }
        let local_pos = (pointer_pos - text_origin) / zoom.max(0.001);
        Some(galley.cursor_from_pos(local_pos).index)
    }

    pub fn hyperlink_at(&self, char_index: usize) -> Option<String> {
        let RenderBoxKind::Text(block) = &self.kind else {
            return None;
        };
        block.hyperlink_at(char_index)
    }

    fn text_galley(&self, ctx: &egui::Context) -> Option<Arc<Galley>> {
        let block = match &self.kind {
            RenderBoxKind::Text(block) => block,
            _ => return None,
        };
        let max_width = (self.size.x * self.scale.x - self.style.padding.x * 2.0).max(80.0);
        let job = block.layout_job(max_width, 1.0, ctx.style().as_ref());
        Some(ctx.fonts_mut(|fonts| fonts.layout_job(job)))
    }

    fn text_galley_from_painter(&self, painter: &Painter) -> Option<Arc<Galley>> {
        let block = match &self.kind {
            RenderBoxKind::Text(block) => block,
            _ => return None,
        };
        let max_width = (self.size.x * self.scale.x - self.style.padding.x * 2.0).max(80.0);
        let job = block.layout_job(max_width, 1.0, painter.ctx().style().as_ref());
        Some(painter.layout_job(job))
    }
}

fn paint_box_background_and_stroke(
    painter: &Painter,
    rect: Rect,
    radius: f32,
    style: &BoxStyle,
    zoom: f32,
) {
    if style.fill != Color32::TRANSPARENT {
        painter.rect_filled(rect, radius, style.fill);
    }

    if style.stroke == Color32::TRANSPARENT || style.stroke_width <= 0.0 {
        return;
    }

    let stroke = Stroke::new((style.stroke_width * zoom).max(0.5), style.stroke);
    match style.stroke_kind {
        BoxStrokeKind::Solid => {
            painter.rect_stroke(rect, radius, stroke, StrokeKind::Inside);
        }
        BoxStrokeKind::Dash => paint_dashed_rect(painter, rect, stroke, zoom),
    }
}

fn paint_dashed_rect(painter: &Painter, rect: Rect, stroke: Stroke, zoom: f32) {
    let inset = stroke.width * 0.5;
    let rect = rect.shrink(inset);
    let dash = (6.0 * zoom).max(stroke.width * 2.0);
    let gap = (4.0 * zoom).max(stroke.width);
    let corners = [
        rect.left_top(),
        rect.right_top(),
        rect.right_bottom(),
        rect.left_bottom(),
    ];

    for index in 0..corners.len() {
        paint_dashed_line(
            painter,
            corners[index],
            corners[(index + 1) % corners.len()],
            stroke,
            dash,
            gap,
        );
    }
}

fn paint_dashed_line(
    painter: &Painter,
    start: Pos2,
    end: Pos2,
    stroke: Stroke,
    dash: f32,
    gap: f32,
) {
    let vector = end - start;
    let length = vector.length();
    if length <= f32::EPSILON {
        return;
    }
    let direction = vector / length;
    let mut offset = 0.0;
    while offset < length {
        let dash_end = (offset + dash).min(length);
        painter.line_segment(
            [start + direction * offset, start + direction * dash_end],
            stroke,
        );
        offset += dash + gap;
    }
}

#[derive(Clone, Debug)]
pub enum RenderBoxKind {
    Text(RichTextBlock),
    Image(ImageBlock),
    Table(TableBlock),
    Container,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum LayoutRole {
    FlowBlock,
    Absolute,
}

#[derive(Clone, Debug)]
pub struct RichTextBlock {
    pub runs: Vec<TextRun>,
    pub alignment: TextAlignment,
    pub vertical_alignment: TextVerticalAlignment,
}

struct PaintedTextRow {
    origin: Pos2,
    galley: Arc<Galley>,
    local_index: usize,
}

impl RichTextBlock {
    pub fn plain_text(&self) -> String {
        self.runs
            .iter()
            .map(|run| run.text.as_str())
            .collect::<Vec<_>>()
            .join("")
    }

    pub fn set_plain_text(&mut self, text: String) {
        let style = self
            .runs
            .first()
            .map(|run| run.style.clone())
            .unwrap_or_else(TextStyle::body);
        self.runs = vec![TextRun::new(text, style)];
    }

    pub fn char_len(&self) -> usize {
        self.plain_text().chars().count()
    }

    pub fn style_state(&self) -> TextStyleState {
        TextStyleState {
            bold: self.runs.iter().all(|run| run.style.bold),
            italic: self.runs.iter().all(|run| run.style.italic),
            underline: self.runs.iter().all(|run| run.style.underline),
            font_size: self
                .runs
                .first()
                .map(|run| run.style.font_size)
                .unwrap_or_else(|| TextStyle::body().font_size),
            color: self
                .runs
                .first()
                .map(|run| run.style.color)
                .unwrap_or_else(|| TextStyle::body().color),
            background_color: self.runs.first().and_then(|run| run.style.background_color),
            alignment: self.alignment,
        }
    }

    pub fn style_state_in_range(&self, range: TextRange) -> TextStyleState {
        let chars = self.styled_chars();
        let (start, end) = range.sorted();
        let slice = &chars[start.min(chars.len())..end.min(chars.len())];
        TextStyleState {
            bold: !slice.is_empty() && slice.iter().all(|(_, style)| style.bold),
            italic: !slice.is_empty() && slice.iter().all(|(_, style)| style.italic),
            underline: !slice.is_empty() && slice.iter().all(|(_, style)| style.underline),
            font_size: slice
                .first()
                .map(|(_, style)| style.font_size)
                .unwrap_or_else(|| self.style_state().font_size),
            color: slice
                .first()
                .map(|(_, style)| style.color)
                .unwrap_or_else(|| self.style_state().color),
            background_color: slice
                .first()
                .and_then(|(_, style)| style.background_color)
                .or_else(|| self.style_state().background_color),
            alignment: self.alignment,
        }
    }

    pub fn toggle_bold(&mut self, range: Option<TextRange>) {
        self.map_styles(range, |style, next| style.bold = next, |state| state.bold);
    }

    pub fn toggle_italic(&mut self, range: Option<TextRange>) {
        self.map_styles(
            range,
            |style, next| style.italic = next,
            |state| state.italic,
        );
    }

    pub fn toggle_underline(&mut self, range: Option<TextRange>) {
        self.map_styles(
            range,
            |style, next| style.underline = next,
            |state| state.underline,
        );
    }

    pub fn set_text_color(&mut self, range: Option<TextRange>, color: Color32) {
        let mut chars = self.styled_chars();
        let (start, end) = match range.filter(|range| !range.is_empty()) {
            Some(range) => range.sorted(),
            None => (0, chars.len()),
        };

        let len = chars.len();
        let start = start.min(len);
        let end = end.min(len);
        for (_, style) in chars.iter_mut().take(end).skip(start) {
            style.color = color;
        }

        self.rebuild_from_styled_chars(chars);
    }

    pub fn set_highlight_color(&mut self, range: Option<TextRange>, color: Option<Color32>) {
        let mut chars = self.styled_chars();
        let (start, end) = match range.filter(|range| !range.is_empty()) {
            Some(range) => range.sorted(),
            None => (0, chars.len()),
        };

        let len = chars.len();
        let start = start.min(len);
        let end = end.min(len);
        for (_, style) in chars.iter_mut().take(end).skip(start) {
            style.background_color = color;
        }

        self.rebuild_from_styled_chars(chars);
    }

    pub fn set_font_size(&mut self, range: Option<TextRange>, font_size: f32) {
        let mut chars = self.styled_chars();
        let (start, end) = match range.filter(|range| !range.is_empty()) {
            Some(range) => range.sorted(),
            None => (0, chars.len()),
        };

        let len = chars.len();
        let start = start.min(len);
        let end = end.min(len);
        for (_, style) in chars.iter_mut().take(end).skip(start) {
            style.font_size = font_size;
        }

        self.rebuild_from_styled_chars(chars);
    }

    pub fn insert_text_at(&mut self, char_index: usize, text: &str) -> usize {
        let mut chars = self.styled_chars();
        let insert_at = char_index.min(chars.len());
        let inherited_style = chars
            .get(insert_at.saturating_sub(1))
            .or_else(|| chars.get(insert_at))
            .map(|(_, style)| style.clone())
            .unwrap_or_else(|| {
                self.runs
                    .first()
                    .map(|run| run.style.clone())
                    .unwrap_or_else(TextStyle::body)
            });
        for (offset, ch) in text.chars().enumerate() {
            chars.insert(insert_at + offset, (ch, inherited_style.clone()));
        }
        self.rebuild_from_styled_chars(chars);
        insert_at + text.chars().count()
    }

    pub fn delete_char_before(&mut self, char_index: usize) -> Option<usize> {
        if char_index == 0 {
            return None;
        }

        let mut chars = self.styled_chars();
        let delete_at = char_index.min(chars.len()).saturating_sub(1);
        chars.remove(delete_at);
        self.rebuild_from_styled_chars(chars);
        Some(delete_at)
    }

    pub fn delete_char_at(&mut self, char_index: usize) -> Option<usize> {
        let mut chars = self.styled_chars();
        if char_index >= chars.len() {
            return None;
        }

        chars.remove(char_index);
        self.rebuild_from_styled_chars(chars);
        Some(char_index)
    }

    pub fn delete_range(&mut self, range: TextRange) -> Option<usize> {
        let (start, end) = range.sorted();
        if start == end {
            return Some(start);
        }
        let mut chars = self.styled_chars();
        let start = start.min(chars.len());
        let end = end.min(chars.len());
        chars.drain(start..end);
        self.rebuild_from_styled_chars(chars);
        Some(start)
    }

    pub fn replace_range(&mut self, range: TextRange, text: &str) -> usize {
        let insertion_point = self.delete_range(range).unwrap_or(range.sorted().0);
        self.insert_text_at(insertion_point, text)
    }

    pub fn max_font_size(&self) -> f32 {
        self.runs
            .iter()
            .map(|run| run.style.font_size)
            .fold(18.0, f32::max)
    }

    pub fn layout_lines(&self, max_width: f32) -> TextLayout {
        let mut lines = Vec::new();
        let mut current_line = TextLine::default();
        let mut line_start_index = 0usize;
        let mut ended_with_break = false;
        let mut run_start = 0usize;

        for run in &self.runs {
            for segment in run.segments(run_start) {
                if segment.forced_break {
                    if current_line.segments.is_empty() {
                        current_line.start_index = line_start_index;
                        current_line.end_index = line_start_index;
                    }
                    current_line.finalize();
                    lines.push(current_line);
                    current_line = TextLine::default();
                    line_start_index = segment.end;
                    ended_with_break = true;
                    continue;
                }

                if current_line.width > 0.0
                    && current_line.width + segment.width > max_width
                    && !segment.text.trim().is_empty()
                {
                    // Wrap only before non-space content so leading whitespace is dropped.
                    current_line.finalize();
                    lines.push(current_line);
                    current_line = TextLine::default();
                    line_start_index = segment.start;
                    ended_with_break = false;

                    if segment.is_space {
                        continue;
                    }
                }

                ended_with_break = false;
                current_line.push(segment);
            }

            run_start += run.text.chars().count();
        }

        if !current_line.segments.is_empty() || ended_with_break {
            if current_line.segments.is_empty() {
                current_line.start_index = line_start_index;
                current_line.end_index = line_start_index;
            }
            current_line.finalize();
            lines.push(current_line);
        }

        if lines.is_empty() {
            lines.push(TextLine {
                start_index: 0,
                end_index: 0,
                width: 0.0,
                height: self.max_font_size() * 1.25,
                baseline: self.max_font_size(),
                segments: Vec::new(),
            });
        }

        let total_height = lines.iter().map(|line| line.height).sum();
        TextLayout {
            lines,
            total_height,
        }
    }

    pub fn paint(&self, painter: &Painter, rect: Rect, style: &BoxStyle, zoom: f32) {
        let content_rect = rect.shrink2(style.padding * zoom);
        let content_width = (content_rect.width() / zoom.max(0.001)).max(80.0);
        let layout_style = painter.ctx().style();
        let layout_style = layout_style.as_ref();
        let galley = painter.layout_job(self.layout_job(content_width, 1.0, layout_style));
        let text_origin = self.text_origin(rect, style, zoom, galley.size().y * zoom);
        let clipped = painter.with_clip_rect(content_rect);
        let styled_chars = self.styled_chars();

        let mut row_start = 0usize;
        for row in &galley.rows {
            let row_end = row_start + row.char_count_including_newline();
            let row_content_end = row_start + row.char_count_excluding_newline();
            let slice_start = row_start.min(styled_chars.len());
            let slice_end = row_content_end.min(styled_chars.len());
            if slice_start < slice_end {
                let row_job = self.layout_job_for_chars(
                    &styled_chars[slice_start..slice_end],
                    zoom,
                    layout_style,
                );
                let row_galley = clipped.layout_job(row_job);
                let row_pos = text_origin + vec2(row.pos.x, row.min_y()) * zoom;
                clipped.galley(row_pos, row_galley, Color32::TRANSPARENT);
            }
            row_start = row_end;
        }
    }

    fn text_origin(&self, rect: Rect, style: &BoxStyle, zoom: f32, text_height: f32) -> Pos2 {
        let content_rect = rect.shrink2(style.padding * zoom);
        let y = match self.vertical_alignment {
            TextVerticalAlignment::Top => content_rect.top(),
            TextVerticalAlignment::Center => {
                content_rect.top() + ((content_rect.height() - text_height).max(0.0) * 0.5)
            }
            TextVerticalAlignment::Bottom => {
                content_rect.top() + (content_rect.height() - text_height).max(0.0)
            }
        };
        match self.alignment {
            TextAlignment::Left | TextAlignment::Justify => Pos2::new(content_rect.left(), y),
            TextAlignment::Center => Pos2::new(content_rect.center().x, y),
            TextAlignment::Right => Pos2::new(content_rect.right(), y),
        }
    }

    fn layout_job(&self, max_width: f32, zoom: f32, style: &Style) -> LayoutJob {
        let mut job = LayoutJob {
            wrap: TextWrapping {
                max_width,
                ..Default::default()
            },
            halign: self.alignment.into(),
            justify: self.alignment == TextAlignment::Justify,
            break_on_newline: true,
            ..Default::default()
        };

        for run in &self.runs {
            append_text_run_to_job(&mut job, &run.text, &run.style, zoom, style);
        }

        job
    }

    fn layout_job_for_chars(
        &self,
        chars: &[(char, TextStyle)],
        zoom: f32,
        style: &Style,
    ) -> LayoutJob {
        let mut job = LayoutJob {
            wrap: TextWrapping {
                max_width: f32::INFINITY,
                ..Default::default()
            },
            break_on_newline: false,
            ..Default::default()
        };

        let mut current_text = String::new();
        let mut current_style = chars.first().map(|(_, style)| style.clone());
        for (ch, char_style) in chars {
            if current_style.as_ref() != Some(char_style) {
                if let Some(style_to_flush) = current_style.take() {
                    append_text_run_to_job(&mut job, &current_text, &style_to_flush, zoom, style);
                    current_text.clear();
                }
                current_style = Some(char_style.clone());
            }
            current_text.push(*ch);
        }

        if let Some(style_to_flush) = current_style {
            append_text_run_to_job(&mut job, &current_text, &style_to_flush, zoom, style);
        }

        job
    }

    fn painted_row_for_index(
        &self,
        galley: &Galley,
        char_index: usize,
        zoom: f32,
        painter: &Painter,
        text_origin: Pos2,
    ) -> Option<PaintedTextRow> {
        let mut row_start = 0usize;
        let mut fallback = None;

        for row in &galley.rows {
            let row_end = row_start + row.char_count_including_newline();
            let row_content_end = row_start + row.char_count_excluding_newline();
            let painted_row = self.painted_row(galley, row, row_start, zoom, painter, text_origin);
            if char_index <= row_end {
                return painted_row.map(|mut painted_row| {
                    painted_row.local_index = char_index
                        .saturating_sub(row_start)
                        .min(row_content_end.saturating_sub(row_start));
                    painted_row
                });
            }
            fallback = painted_row.map(|mut painted_row| {
                painted_row.local_index = row_content_end.saturating_sub(row_start);
                painted_row
            });
            row_start = row_end;
        }

        fallback
    }

    fn painted_row(
        &self,
        _galley: &Galley,
        row: &PlacedRow,
        row_start: usize,
        zoom: f32,
        painter: &Painter,
        text_origin: Pos2,
    ) -> Option<PaintedTextRow> {
        let styled_chars = self.styled_chars();
        let row_content_end = row_start + row.char_count_excluding_newline();
        let slice_start = row_start.min(styled_chars.len());
        let slice_end = row_content_end.min(styled_chars.len());
        let row_job = self.layout_job_for_chars(
            &styled_chars[slice_start..slice_end],
            zoom,
            painter.ctx().style().as_ref(),
        );
        let row_galley = painter.layout_job(row_job);
        if row_galley.rows.is_empty() {
            return None;
        }

        Some(PaintedTextRow {
            origin: text_origin + vec2(row.pos.x, row.min_y()) * zoom,
            galley: row_galley,
            local_index: 0,
        })
    }

    fn painted_caret_index_at(
        &self,
        galley: &Galley,
        pointer_pos: Pos2,
        zoom: f32,
        painter: &Painter,
        text_origin: Pos2,
    ) -> Option<usize> {
        let mut row_start = 0usize;
        let mut best: Option<(f32, usize)> = None;

        for row in &galley.rows {
            let row_end = row_start + row.char_count_including_newline();
            let row_content_end = row_start + row.char_count_excluding_newline();
            if let Some(painted_row) =
                self.painted_row(galley, row, row_start, zoom, painter, text_origin)
            {
                let painted = &painted_row.galley.rows[0];
                let row_min_y = painted_row.origin.y + painted.min_y();
                let row_max_y = painted_row.origin.y + painted.max_y();
                let distance = if pointer_pos.y < row_min_y {
                    row_min_y - pointer_pos.y
                } else if pointer_pos.y > row_max_y {
                    pointer_pos.y - row_max_y
                } else {
                    0.0
                };
                let local_pos = pointer_pos - painted_row.origin;
                let local_index = painted_row.galley.cursor_from_pos(local_pos).index;
                let index = row_start + local_index.min(row_content_end.saturating_sub(row_start));
                if best.is_none_or(|(best_distance, _)| distance < best_distance) {
                    best = Some((distance, index.min(row_end)));
                }
            }
            row_start = row_end;
        }

        best.map(|(_, index)| index)
    }

    fn styled_chars(&self) -> Vec<(char, TextStyle)> {
        let mut chars = Vec::new();
        for run in &self.runs {
            for ch in run.text.chars() {
                chars.push((ch, run.style.clone()));
            }
        }
        chars
    }

    fn hyperlink_at(&self, char_index: usize) -> Option<String> {
        let chars = self.styled_chars();
        if chars.is_empty() {
            return None;
        }

        let index = char_index.min(chars.len().saturating_sub(1));
        chars[index].1.hyperlink.clone().or_else(|| {
            char_index
                .checked_sub(1)
                .and_then(|index| chars.get(index))
                .and_then(|(_, style)| style.hyperlink.clone())
        })
    }

    fn rebuild_from_styled_chars(&mut self, chars: Vec<(char, TextStyle)>) {
        if chars.is_empty() {
            self.runs = vec![TextRun::new(
                "",
                self.runs
                    .first()
                    .map(|run| run.style.clone())
                    .unwrap_or_else(TextStyle::body),
            )];
            return;
        }

        let mut runs = Vec::new();
        let mut current_text = String::new();
        let mut current_style = chars[0].1.clone();

        for (ch, style) in chars {
            if style == current_style {
                current_text.push(ch);
            } else {
                runs.push(TextRun::new(
                    std::mem::take(&mut current_text),
                    current_style,
                ));
                current_style = style;
                current_text.push(ch);
            }
        }

        runs.push(TextRun::new(current_text, current_style));
        self.runs = runs;
    }

    fn map_styles(
        &mut self,
        range: Option<TextRange>,
        mut apply: impl FnMut(&mut TextStyle, bool),
        state: impl Fn(TextStyleState) -> bool,
    ) {
        let whole_state = self.style_state();
        let target_next = !state(match range {
            Some(range) if !range.is_empty() => self.style_state_in_range(range),
            _ => whole_state,
        });

        let mut chars = self.styled_chars();
        let (start, end) = match range.filter(|range| !range.is_empty()) {
            Some(range) => range.sorted(),
            None => (0, chars.len()),
        };

        let len = chars.len();
        let start = start.min(len);
        let end = end.min(len);
        for (_, style) in chars.iter_mut().take(end).skip(start) {
            apply(style, target_next);
        }

        self.rebuild_from_styled_chars(chars);
    }
}

fn append_text_run_to_job(
    job: &mut LayoutJob,
    text: &str,
    run_style: &TextStyle,
    zoom: f32,
    style: &Style,
) {
    if text.is_empty() {
        return;
    }

    let family = if run_style.bold {
        FontFamily::Name(EDITOR_BOLD_FONT_NAME.into())
    } else {
        FontFamily::Name(EDITOR_REGULAR_FONT_NAME.into())
    };
    let mut rich_text = RichText::new(text).font(FontId::new(run_style.font_size * zoom, family));

    if run_style.bold {
        rich_text = rich_text.strong();
    }
    if run_style.italic {
        rich_text = rich_text.italics();
    }
    if run_style.underline {
        rich_text = rich_text.underline();
    }
    rich_text = rich_text.color(run_style.color);
    if let Some(background_color) = run_style.background_color {
        rich_text = rich_text.background_color(background_color);
    }

    rich_text.append_to(job, style, FontSelection::Default, Align::Center);
}

#[derive(Clone, Debug)]
pub struct TextRun {
    pub text: String,
    pub style: TextStyle,
}

impl TextRun {
    pub fn new(text: impl Into<String>, style: TextStyle) -> Self {
        Self {
            text: text.into(),
            style,
        }
    }

    pub fn estimated_width(&self) -> f32 {
        let spacing_factor = self.style.font_size * 0.55;
        self.text.chars().count() as f32 * spacing_factor
    }

    fn segments(&self, start_offset: usize) -> Vec<LineSegment> {
        let mut parts = Vec::new();
        let mut current = String::new();
        let mut current_start = start_offset;
        let mut char_index = start_offset;

        for ch in self.text.chars() {
            if ch == '\n' {
                if !current.is_empty() {
                    parts.push(LineSegment::new(
                        std::mem::take(&mut current),
                        current_start,
                        char_index,
                        self.style.clone(),
                    ));
                }
                parts.push(LineSegment::forced_break(char_index, self.style.clone()));
                char_index += 1;
                current_start = char_index;
                continue;
            }

            if current.is_empty() {
                current_start = char_index;
            }
            current.push(ch);
            char_index += 1;
            if ch.is_whitespace() {
                // Keep whitespace attached to the preceding token so wrapping is deterministic.
                parts.push(LineSegment::new(
                    std::mem::take(&mut current),
                    current_start,
                    char_index,
                    self.style.clone(),
                ));
            }
        }

        if !current.is_empty() {
            parts.push(LineSegment::new(
                current,
                current_start,
                char_index,
                self.style.clone(),
            ));
        }

        parts
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct TextStyle {
    pub font_size: f32,
    pub color: Color32,
    pub background_color: Option<Color32>,
    pub hyperlink: Option<String>,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum TextAlignment {
    #[default]
    Left,
    Center,
    Right,
    Justify,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum TextVerticalAlignment {
    #[default]
    Top,
    Center,
    Bottom,
}

impl From<TextAlignment> for Align {
    fn from(alignment: TextAlignment) -> Self {
        match alignment {
            TextAlignment::Left => Align::LEFT,
            TextAlignment::Center => Align::Center,
            TextAlignment::Right => Align::RIGHT,
            TextAlignment::Justify => Align::LEFT,
        }
    }
}

impl TextStyle {
    pub fn heading() -> Self {
        Self {
            font_size: 32.0,
            color: Color32::from_rgb(36, 38, 41),
            background_color: None,
            hyperlink: None,
            bold: false,
            italic: false,
            underline: false,
        }
    }

    pub fn body() -> Self {
        Self {
            font_size: 20.0,
            color: Color32::from_rgb(62, 69, 77),
            background_color: None,
            hyperlink: None,
            bold: false,
            italic: false,
            underline: false,
        }
    }

    pub fn accent() -> Self {
        Self {
            font_size: 22.0,
            color: Color32::from_rgb(181, 82, 38),
            background_color: None,
            hyperlink: None,
            bold: false,
            italic: false,
            underline: false,
        }
    }
}

#[derive(Clone)]
pub struct ImageBlock {
    pub path: PathBuf,
    pub size: Vec2,
    pub margin_left: f32,
    pub margin_right: f32,
    pub margin_top: f32,
    pub margin_bottom: f32,
    pub center_horizontally: bool,
    pub color_image: ColorImage,
    texture: Option<TextureHandle>,
}

impl std::fmt::Debug for ImageBlock {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("ImageBlock")
            .field("path", &self.path)
            .field("size", &self.size)
            .field("margin_left", &self.margin_left)
            .field("margin_right", &self.margin_right)
            .field("margin_top", &self.margin_top)
            .field("margin_bottom", &self.margin_bottom)
            .field("center_horizontally", &self.center_horizontally)
            .field("color_image_size", &self.color_image.size)
            .finish()
    }
}

impl ImageBlock {
    pub fn from_encoded_bytes(
        path: PathBuf,
        bytes: &[u8],
        requested_size: Option<Vec2>,
        margin_left: f32,
        margin_right: f32,
        margin_top: f32,
        margin_bottom: f32,
        center_horizontally: bool,
    ) -> Result<Self, image::ImageError> {
        let rgba_image = image::load_from_memory(bytes)?.to_rgba8();
        let pixel_size = vec2(rgba_image.width() as f32, rgba_image.height() as f32);
        let color_image = ColorImage::from_rgba_unmultiplied(
            [rgba_image.width() as usize, rgba_image.height() as usize],
            rgba_image.as_raw(),
        );

        Ok(Self {
            path,
            size: requested_size
                .unwrap_or(pixel_size)
                .max(Vec2::splat(IMAGE_MIN_SIZE)),
            margin_left,
            margin_right,
            margin_top,
            margin_bottom,
            center_horizontally,
            color_image,
            texture: None,
        })
    }

    pub fn load_from_path(path: impl AsRef<Path>) -> Result<Self, image::ImageError> {
        let path = path.as_ref().to_path_buf();
        let bytes = std::fs::read(&path)?;
        Self::from_encoded_bytes(path, &bytes, None, 0.0, 0.0, 0.0, 0.0, false)
    }

    pub fn reload_from_path(&mut self, path: impl AsRef<Path>) -> Result<(), image::ImageError> {
        let replacement = Self::from_encoded_bytes(
            path.as_ref().to_path_buf(),
            &std::fs::read(path.as_ref())?,
            Some(self.size),
            self.margin_left,
            self.margin_right,
            self.margin_top,
            self.margin_bottom,
            self.center_horizontally,
        )?;

        self.path = replacement.path;
        self.color_image = replacement.color_image;
        self.texture = None;
        Ok(())
    }

    pub fn placeholder(label: impl Into<String>, size: Vec2) -> Self {
        let label = label.into();
        let color_image = placeholder_color_image(label.as_bytes());
        Self {
            path: PathBuf::from(label),
            size,
            margin_left: 0.0,
            margin_right: 0.0,
            margin_top: 0.0,
            margin_bottom: 0.0,
            center_horizontally: false,
            color_image,
            texture: None,
        }
    }

    pub fn texture_handle(&mut self, painter: &Painter, image_id: u64) -> TextureHandle {
        if self.texture.is_none() {
            let texture_name = format!("embedded-document-image-{image_id}");
            self.texture = Some(painter.ctx().load_texture(
                texture_name,
                self.color_image.clone(),
                TextureOptions::LINEAR,
            ));
        }

        self.texture
            .as_ref()
            .expect("texture is initialized above")
            .clone()
    }

    pub fn invalidate_texture(&mut self) {
        self.texture = None;
    }

    fn paint(&mut self, painter: &Painter, rect: Rect, image_id: u64) {
        let texture = self.texture_handle(painter, image_id);
        painter.image(
            texture.id(),
            rect,
            Rect::from_min_max(Pos2::ZERO, Pos2::new(1.0, 1.0)),
            Color32::WHITE,
        );
    }

    fn paint_rotated(
        &mut self,
        painter: &Painter,
        rect: Rect,
        style: &BoxStyle,
        zoom: f32,
        rotation_degrees: f32,
        image_id: u64,
    ) {
        let texture = self.texture_handle(painter, image_id);
        let corners = rotated_rect_corners(rect, rotation_degrees);
        painter.add(Shape::convex_polygon(
            corners.to_vec(),
            style.fill,
            Stroke::new((style.stroke_width * zoom).clamp(1.0, 2.0), style.stroke),
        ));

        let mut mesh = Mesh::with_texture(texture.id());
        mesh.vertices = vec![
            Vertex {
                pos: corners[0],
                uv: Pos2::new(0.0, 0.0),
                color: Color32::WHITE,
            },
            Vertex {
                pos: corners[1],
                uv: Pos2::new(1.0, 0.0),
                color: Color32::WHITE,
            },
            Vertex {
                pos: corners[2],
                uv: Pos2::new(1.0, 1.0),
                color: Color32::WHITE,
            },
            Vertex {
                pos: corners[3],
                uv: Pos2::new(0.0, 1.0),
                color: Color32::WHITE,
            },
        ];
        mesh.indices = vec![0, 1, 2, 0, 2, 3];
        painter.add(Shape::Mesh(mesh.into()));
    }
}

fn placeholder_color_image(seed: &[u8]) -> ColorImage {
    let width = 320;
    let height = 180;
    let tint = seed.iter().fold(0u8, |acc, byte| acc.wrapping_add(*byte));
    let mut pixels = Vec::with_capacity(width * height);
    for y in 0..height {
        for x in 0..width {
            let stripe = ((x / 18 + y / 18) % 2) as u8;
            let base = if stripe == 0 { 226 } else { 210 };
            pixels.push(Color32::from_rgb(base, base.saturating_add(tint / 12), 236));
        }
    }
    ColorImage::new([width, height], pixels)
}

#[derive(Clone, Debug)]
pub struct TableBlock {
    pub rows: Vec<Vec<String>>,
}

impl TableBlock {
    fn row_count(&self) -> usize {
        self.rows.len().max(1)
    }

    fn column_count(&self) -> usize {
        self.rows.iter().map(Vec::len).max().unwrap_or(1)
    }

    fn paint(&self, painter: &Painter, rect: Rect, style: &BoxStyle, zoom: f32) {
        let rows = self.row_count();
        let cols = self.column_count();
        let scaled_padding = style.padding * zoom;
        // Cells are distributed evenly; this is a placeholder table renderer, not content-sized.
        let cell_w = (rect.width() - scaled_padding.x * 2.0) / cols as f32;
        let cell_h = (rect.height() - scaled_padding.y * 2.0) / rows as f32;
        let top_left = rect.min + scaled_padding;

        for row in 0..rows {
            for col in 0..cols {
                let min = top_left + vec2(col as f32 * cell_w, row as f32 * cell_h);
                let cell = Rect::from_min_size(min, vec2(cell_w, cell_h));
                painter.rect_stroke(
                    cell,
                    0.0,
                    Stroke::new(1.0, style.stroke),
                    StrokeKind::Inside,
                );

                if let Some(text) = self.rows.get(row).and_then(|items| items.get(col)) {
                    painter.text(
                        cell.center(),
                        Align2::CENTER_CENTER,
                        text,
                        FontId::new(16.0 * zoom, FontFamily::Proportional),
                        Color32::from_rgb(48, 52, 56),
                    );
                }
            }
        }
    }
}

#[derive(Clone, Debug)]
pub struct BoxStyle {
    pub fill: Color32,
    pub stroke: Color32,
    pub stroke_width: f32,
    pub stroke_kind: BoxStrokeKind,
    pub padding: Vec2,
    pub outer_margin: Vec2,
    pub corner_radius: f32,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BoxStrokeKind {
    Solid,
    Dash,
}

impl Default for BoxStyle {
    fn default() -> Self {
        Self {
            fill: Color32::from_rgb(255, 252, 248),
            stroke: Color32::from_rgb(170, 176, 184),
            stroke_width: 1.0,
            stroke_kind: BoxStrokeKind::Solid,
            padding: vec2(18.0, 16.0),
            outer_margin: vec2(0.0, 16.0),
            corner_radius: 10.0,
        }
    }
}

#[derive(Clone, Debug)]
pub struct AnimationSpec {
    pub phase: f32,
    pub amplitude: Vec2,
    pub kind: AnimationKind,
}

#[derive(Clone, Debug)]
pub enum AnimationKind {
    PreviewOscillation,
    Entrance {
        effect: EntranceEffect,
        direction: Option<FlyInDirection>,
        duration_seconds: f32,
    },
    Emphasis {
        effect: EmphasisEffect,
        duration_seconds: f32,
    },
    Exit {
        effect: ExitEffect,
        direction: Option<FlyInDirection>,
        duration_seconds: f32,
    },
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum EntranceEffect {
    Appear,
    VenetianBlinds,
    Box,
    Checkerboard,
    Circle,
    Oval,
    FlyIn,
    FlyInSlow,
    DissolveIn,
    FadeIn,
    FadeInAndZoom,
    Zoom,
    Expand,
    SpinIn,
    Bounce,
    SpiralIn,
    Boomerang,
    Sling,
    Glide,
    Float,
    Magnify,
    Wipe,
    Wheel,
    RandomBars,
    Split,
    Plus,
    Diamond,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum EmphasisEffect {
    Spin,
    GrowShrink,
    Pulse,
    Teeter,
    Transparency,
    FillColor,
    LineColor,
    FontColor,
    BoldFlash,
    Blink,
    ColorPulse,
    GrowWithColor,
    Lighten,
    Desaturate,
    Wave,
    Flicker,
    VerticalHighlight,
    HorizontalHighlight,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ExitEffect {
    Disappear,
    FadeOut,
    FlyOut,
    FlyOutSlow,
    WipeOut,
    Split,
    Box,
    Circle,
    Diamond,
    DissolveOut,
    RandomBars,
    Checkerboard,
    VenetianBlinds,
    Wheel,
    ClockWipe,
    PeekOut,
    Zoom,
    FadeOutAndZoom,
    Collapse,
    Compress,
    Stretchy,
    SpinOut,
    Swivel,
    Sling,
    SpiralOut,
    Boomerang,
    Bounce,
    Float,
    Glide,
    Fold,
    Thread,
    Random,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FlyInDirection {
    FromLeft,
    FromRight,
    FromTop,
    FromBottom,
}

impl AnimationSpec {
    pub fn preview_oscillation(phase: f32, amplitude: Vec2) -> Self {
        Self {
            phase,
            amplitude,
            kind: AnimationKind::PreviewOscillation,
        }
    }

    pub fn entrance_appear(duration_seconds: f32) -> Self {
        Self::entrance(EntranceEffect::Appear, None, duration_seconds)
    }

    pub fn entrance_fly_in(direction: FlyInDirection, duration_seconds: f32) -> Self {
        Self::entrance(EntranceEffect::FlyIn, Some(direction), duration_seconds)
    }

    pub fn entrance(
        effect: EntranceEffect,
        direction: Option<FlyInDirection>,
        duration_seconds: f32,
    ) -> Self {
        Self {
            phase: 0.0,
            amplitude: Vec2::ZERO,
            kind: AnimationKind::Entrance {
                effect,
                direction,
                duration_seconds,
            },
        }
    }

    pub fn emphasis(effect: EmphasisEffect, duration_seconds: f32) -> Self {
        Self {
            phase: 0.0,
            amplitude: Vec2::ZERO,
            kind: AnimationKind::Emphasis {
                effect,
                duration_seconds,
            },
        }
    }

    pub fn exit(
        effect: ExitEffect,
        direction: Option<FlyInDirection>,
        duration_seconds: f32,
    ) -> Self {
        Self {
            phase: 0.0,
            amplitude: Vec2::ZERO,
            kind: AnimationKind::Exit {
                effect,
                direction,
                duration_seconds,
            },
        }
    }

    pub fn is_preview_oscillation(&self) -> bool {
        matches!(self.kind, AnimationKind::PreviewOscillation)
    }

    pub fn preview_offset(&self) -> Vec2 {
        // Preview animation is just a lightweight oscillating transform.
        vec2(
            self.phase.sin() * self.amplitude.x,
            self.phase.cos() * self.amplitude.y,
        )
    }
}

#[derive(Clone, Copy, Debug, Default)]
pub struct TextStyleState {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub font_size: f32,
    pub color: Color32,
    pub background_color: Option<Color32>,
    pub alignment: TextAlignment,
}

#[derive(Clone, Debug)]
pub struct TextLayout {
    pub lines: Vec<TextLine>,
    pub total_height: f32,
}

#[derive(Clone, Debug, Default)]
pub struct TextLine {
    pub start_index: usize,
    pub end_index: usize,
    pub width: f32,
    pub height: f32,
    pub baseline: f32,
    pub segments: Vec<LineSegment>,
}

impl TextLine {
    fn push(&mut self, segment: LineSegment) {
        if self.segments.is_empty() {
            self.start_index = segment.start;
        }
        self.end_index = segment.end;
        self.width += segment.width;
        self.height = self.height.max(segment.style.font_size * 1.25);
        self.baseline = self.baseline.max(segment.style.font_size);
        self.segments.push(segment);
    }

    fn finalize(&mut self) {
        // Trim trailing spaces from the measured width so wrapped lines end cleanly.
        while self
            .segments
            .last()
            .is_some_and(|segment| segment.is_space && !segment.forced_break)
        {
            if let Some(last) = self.segments.pop() {
                self.width -= last.width;
            }
        }

        self.end_index = self
            .segments
            .last()
            .map(|segment| segment.end)
            .unwrap_or(self.start_index);

        if self.height == 0.0 {
            self.height = 24.0;
            self.baseline = 18.0;
        }
    }
}

#[derive(Clone, Debug)]
pub struct LineSegment {
    pub start: usize,
    pub end: usize,
    pub text: String,
    pub style: TextStyle,
    pub width: f32,
    pub forced_break: bool,
    pub is_space: bool,
}

impl LineSegment {
    fn new(text: String, start: usize, end: usize, style: TextStyle) -> Self {
        let width = text.chars().count() as f32 * style.font_size * 0.55;
        let is_space = text.chars().all(char::is_whitespace);

        Self {
            start,
            end,
            text,
            style,
            width,
            forced_break: false,
            is_space,
        }
    }

    fn forced_break(index: usize, style: TextStyle) -> Self {
        Self {
            start: index,
            end: index + 1,
            text: String::new(),
            style,
            width: 0.0,
            forced_break: true,
            is_space: false,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use egui::pos2;

    fn assert_close(actual: f32, expected: f32) {
        assert!(
            (actual - expected).abs() < 0.001,
            "expected {expected}, got {actual}"
        );
    }

    #[test]
    fn zoom_scroll_offset_keeps_canvas_point_under_mouse() {
        let canvas_size = vec2(1000.0, 800.0);
        let current_zoom = 1.0;
        let view = CanvasViewState {
            viewport_rect: Rect::from_min_size(pos2(100.0, 50.0), vec2(500.0, 400.0)),
            canvas_origin: pos2(-120.0, -80.0),
            canvas_size: vec2(1000.0, 800.0),
            canvas_content_origin: Vec2::ZERO,
        };

        let pointer = pos2(350.0, 250.0);
        let next_zoom = 2.0;
        let current_scroll_offset = vec2(220.0, 130.0);
        let next_offset = zoom_scroll_offset_for_pointer(
            canvas_size,
            current_zoom,
            next_zoom,
            view,
            pointer,
            current_scroll_offset,
        )
        .expect("pointer is over the rendered canvas");

        assert_close(next_offset.x, 690.0);
        assert_close(next_offset.y, 460.0);

        let top_left_zoom_offset = current_scroll_offset * next_zoom;
        assert!(
            (next_offset - top_left_zoom_offset).length() > 1.0,
            "zoom offset unexpectedly matches top-left anchored zoom"
        );
    }

    #[test]
    fn zoom_scroll_offset_is_none_when_mouse_is_outside_canvas() {
        let mut canvas = RichCanvas::new(vec2(1000.0, 800.0));
        canvas.zoom = 1.0;
        canvas.set_last_canvas_view_for_test(CanvasViewState {
            viewport_rect: Rect::from_min_size(pos2(100.0, 50.0), vec2(500.0, 400.0)),
            canvas_origin: pos2(120.0, 80.0),
            canvas_size: vec2(1000.0, 800.0),
            canvas_content_origin: Vec2::ZERO,
        });

        assert!(
            canvas
                .focused_zoom_scroll_offset(pos2(10.0, 10.0), 2.0, Vec2::ZERO)
                .is_none()
        );
    }

    #[test]
    fn rotated_rect_contains_uses_local_image_axes() {
        let rect = Rect::from_min_size(pos2(100.0, 100.0), vec2(120.0, 60.0));
        let inside = rotate_point(pos2(210.0, 130.0), rect.center(), 45.0);
        let outside_axis_aligned_corner = pos2(100.0, 100.0);

        assert!(rotated_rect_contains(rect, 45.0, inside));
        assert!(!rotated_rect_contains(
            rect,
            45.0,
            outside_axis_aligned_corner
        ));
    }

    #[test]
    fn rotated_image_handles_follow_rotated_corners() {
        let rect = Rect::from_min_size(pos2(100.0, 100.0), vec2(120.0, 60.0));
        let handles = image_handle_points(rect, 90.0);
        let top_left = handles
            .iter()
            .find(|(handle, _)| *handle == ImageResizeHandle::TopLeft)
            .map(|(_, point)| *point)
            .expect("top-left handle exists");

        assert_close(top_left.x, 190.0);
        assert_close(top_left.y, 70.0);
    }

    #[test]
    fn text_caret_row_navigation_is_zoom_independent() {
        let ctx = egui::Context::default();
        configure_editor_fonts(&ctx);
        let _ = ctx.run(Default::default(), |ctx| {
            let mut text_box = RenderBox::text(
                1,
                LayoutRole::Absolute,
                vec![TextRun::new(
                    "A long line that wraps into multiple rows in a narrow text box",
                    TextStyle::body(),
                )],
            );
            text_box.size.x = 180.0;

            let low_zoom = text_box
                .move_caret_vertical(0, 1, ctx, 0.5, None)
                .expect("low zoom row movement should resolve");
            let high_zoom = text_box
                .move_caret_vertical(0, 1, ctx, 2.0, None)
                .expect("high zoom row movement should resolve");

            assert_eq!(low_zoom, high_zoom);
        });
    }

    #[test]
    fn set_text_font_size_applies_to_selected_range_only() {
        let mut text_box = RenderBox::text(
            1,
            LayoutRole::Absolute,
            vec![TextRun::new("hello", TextStyle::body())],
        );

        assert!(text_box.set_text_font_size(
            Some(TextRange {
                anchor: 1,
                caret: 4
            }),
            36.0
        ));

        let block = match &text_box.kind {
            RenderBoxKind::Text(block) => block,
            _ => panic!("expected text block"),
        };
        let chars = block.styled_chars();
        assert_eq!(chars[0].1.font_size, 20.0);
        assert_eq!(chars[1].1.font_size, 36.0);
        assert_eq!(chars[2].1.font_size, 36.0);
        assert_eq!(chars[3].1.font_size, 36.0);
        assert_eq!(chars[4].1.font_size, 20.0);
    }

    #[test]
    fn hyperlink_lookup_uses_current_or_previous_character() {
        let mut link_style = TextStyle::body();
        link_style.hyperlink = Some("https://example.com".to_owned());
        let text_box = RenderBox::text(
            1,
            LayoutRole::Absolute,
            vec![
                TextRun::new("Go ", TextStyle::body()),
                TextRun::new("there", link_style),
            ],
        );

        assert_eq!(
            text_box.hyperlink_at(3).as_deref(),
            Some("https://example.com")
        );
        assert_eq!(
            text_box.hyperlink_at(8).as_deref(),
            Some("https://example.com")
        );
        assert_eq!(text_box.hyperlink_at(1), None);
    }

    #[test]
    fn set_highlight_color_applies_to_selected_range_only() {
        let mut text_box = RenderBox::text(
            1,
            LayoutRole::Absolute,
            vec![TextRun::new("hello", TextStyle::body())],
        );

        assert!(text_box.set_highlight_color(
            Some(TextRange {
                anchor: 1,
                caret: 4
            }),
            Some(Color32::YELLOW)
        ));

        let block = match &text_box.kind {
            RenderBoxKind::Text(block) => block,
            _ => panic!("expected text block"),
        };
        let chars = block.styled_chars();
        assert_eq!(chars[0].1.background_color, None);
        assert_eq!(chars[1].1.background_color, Some(Color32::YELLOW));
        assert_eq!(chars[2].1.background_color, Some(Color32::YELLOW));
        assert_eq!(chars[3].1.background_color, Some(Color32::YELLOW));
        assert_eq!(chars[4].1.background_color, None);
    }

    #[test]
    fn text_style_state_reports_range_font_size() {
        let mut text_box = RenderBox::text(
            1,
            LayoutRole::Absolute,
            vec![TextRun::new("hello", TextStyle::body())],
        );
        text_box.set_text_font_size(
            Some(TextRange {
                anchor: 2,
                caret: 5,
            }),
            42.0,
        );

        let style_state = text_box
            .text_style_state_in_range(TextRange {
                anchor: 2,
                caret: 5,
            })
            .expect("text style state should exist");

        assert_eq!(style_state.font_size, 42.0);
    }
}
