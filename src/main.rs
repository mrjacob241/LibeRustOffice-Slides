use eframe::{App, Frame, NativeOptions, egui};
mod odp_loader;
use rich_canvas::{
    AnimationSpec, CanvasMode, CanvasSelection, ImageResizeHandle, LayoutRole, RenderBox,
    RichCanvas, TableBlock, TextAlignment, TextRange, TextRun, TextStyle, TextStyleState,
};
use std::sync::Arc;

const APP_TITLE: &str = "LibeRustOffice Slides v0.01";
const SIDE_PANEL_MIN_WIDTH: f32 = 180.0;
const SIDE_PANEL_DEFAULT_WIDTH: f32 = 220.0;
const SIDE_PANEL_MAX_WIDTH: f32 = 420.0;
const COLLAPSED_SIDE_RAIL_WIDTH: f32 = 28.0;
const IMAGE_PANEL_SECTION_GAP: f32 = 12.0;
const EDITOR_REGULAR_FONT_NAME: &str = "editor_regular";
const PT_TO_PX: f32 = 4.0 / 3.0;
const MIN_FONT_SIZE_PT: f32 = 6.0;
const MAX_FONT_SIZE_PT: f32 = 144.0;
const RECENT_COLOR_SLOT_COUNT: usize = 10;
const TEXT_COLOR_PALETTE: [egui::Color32; 10] = [
    egui::Color32::BLACK,
    egui::Color32::RED,
    egui::Color32::BLUE,
    egui::Color32::GREEN,
    egui::Color32::from_rgb(255, 165, 0),
    egui::Color32::WHITE,
    egui::Color32::from_rgb(128, 0, 128),
    egui::Color32::from_rgb(0, 150, 160),
    egui::Color32::from_rgb(120, 80, 40),
    egui::Color32::from_rgb(90, 90, 90),
];
const HIGHLIGHT_COLOR_PALETTE: [Option<egui::Color32>; 10] = [
    None,
    Some(egui::Color32::YELLOW),
    Some(egui::Color32::from_rgb(255, 210, 120)),
    Some(egui::Color32::from_rgb(180, 230, 180)),
    Some(egui::Color32::from_rgb(190, 220, 255)),
    Some(egui::Color32::from_rgb(255, 180, 180)),
    Some(egui::Color32::from_rgb(220, 190, 255)),
    Some(egui::Color32::from_rgb(185, 245, 245)),
    Some(egui::Color32::from_rgb(235, 235, 160)),
    Some(egui::Color32::from_rgb(210, 210, 210)),
];

#[derive(Clone, Copy, PartialEq, Eq)]
enum LeftPanelTab {
    Slides,
    Scheduler,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum RightPanelTab {
    Properties,
    Image,
    Animation,
    Scheduler,
}

struct LibeRustOfficeSlidesApp {
    canvas: RichCanvas,
    slides: Vec<RichCanvas>,
    active_slide: usize,
    phase: f32,
    selected_box: Option<u64>,
    caret_index: usize,
    selection_anchor: Option<usize>,
    caret_x_target: Option<f32>,
    document_name: String,
    left_panel_open: bool,
    left_panel_tab: LeftPanelTab,
    right_panel_open: bool,
    right_panel_tab: RightPanelTab,
    link_image_dimensions: bool,
    image_resize_drag: Option<ImageResizeDrag>,
    text_box_drag: Option<TextBoxDrag>,
    text_rgb_picker: [u8; 3],
    highlight_rgb_picker: [u8; 3],
    recent_text_custom_colors: [Option<egui::Color32>; RECENT_COLOR_SLOT_COUNT],
    recent_highlight_custom_colors: [Option<egui::Color32>; RECENT_COLOR_SLOT_COUNT],
    fit_zoom_pending: bool,
    status: String,
    next_box_id: u64,
}

#[derive(Clone, Copy, Debug)]
struct ImageResizeDrag {
    box_id: u64,
    handle: ImageResizeHandle,
    start_pointer: egui::Pos2,
    start_position: egui::Pos2,
    start_size: egui::Vec2,
    start_scale: egui::Vec2,
    start_image_size: egui::Vec2,
    start_rotation: f32,
    zoom: f32,
}

#[derive(Clone, Copy, Debug)]
struct TextBoxDrag {
    box_id: u64,
    handle: ImageResizeHandle,
    start_pointer: egui::Pos2,
    start_position: egui::Pos2,
    start_size: egui::Vec2,
    zoom: f32,
}

impl LibeRustOfficeSlidesApp {
    fn new(cc: &eframe::CreationContext<'_>) -> Self {
        configure_local_editor_fonts(&cc.egui_ctx);
        let loaded = odp_loader::load_default_odp().unwrap_or_else(|error| {
            panic!(
                "Failed to load default ODP '{}': {error}",
                odp_loader::DEFAULT_ODP_PATH
            )
        });
        let slides = loaded.slides;
        let document_name = loaded.document_name;
        let status = format!("Loaded default ODP {}", odp_loader::DEFAULT_ODP_PATH);
        let canvas = slides
            .first()
            .cloned()
            .expect("ODP loader returned no slides");
        let next_box_id = slides
            .iter()
            .flat_map(|slide| slide.boxes.iter().map(|render_box| render_box.id))
            .max()
            .unwrap_or(6)
            + 1;
        Self {
            canvas,
            slides,
            active_slide: 0,
            phase: 0.0,
            selected_box: None,
            caret_index: 0,
            selection_anchor: None,
            caret_x_target: None,
            document_name,
            left_panel_open: true,
            left_panel_tab: LeftPanelTab::Slides,
            right_panel_open: true,
            right_panel_tab: RightPanelTab::Properties,
            link_image_dimensions: true,
            image_resize_drag: None,
            text_box_drag: None,
            text_rgb_picker: [36, 38, 41],
            highlight_rgb_picker: [255, 242, 0],
            recent_text_custom_colors: [None; RECENT_COLOR_SLOT_COUNT],
            recent_highlight_custom_colors: [None; RECENT_COLOR_SLOT_COUNT],
            fit_zoom_pending: true,
            status,
            next_box_id,
        }
    }

    fn window_title(&self) -> String {
        format!("{APP_TITLE}  ::  {}", self.document_name)
    }

    fn reload_default_odp(&mut self) {
        let loaded = odp_loader::load_default_odp().unwrap_or_else(|error| {
            panic!(
                "Failed to load default ODP '{}': {error}",
                odp_loader::DEFAULT_ODP_PATH
            )
        });
        self.slides = loaded.slides;
        self.active_slide = 0;
        self.canvas = self
            .slides
            .first()
            .cloned()
            .expect("ODP loader returned no slides");
        self.phase = 0.0;
        self.selected_box = None;
        self.caret_index = 0;
        self.selection_anchor = None;
        self.caret_x_target = None;
        self.image_resize_drag = None;
        self.text_box_drag = None;
        self.fit_zoom_pending = true;
        self.document_name = loaded.document_name;
        self.status = format!("Reloaded default ODP {}", odp_loader::DEFAULT_ODP_PATH);
        self.next_box_id = self
            .slides
            .iter()
            .flat_map(|slide| slide.boxes.iter().map(|render_box| render_box.id))
            .max()
            .unwrap_or(0)
            + 1;
    }

    fn sync_active_slide(&mut self) {
        if let Some(slide) = self.slides.get_mut(self.active_slide) {
            *slide = self.canvas.clone();
        }
    }

    fn select_slide(&mut self, index: usize) {
        if index >= self.slides.len() || index == self.active_slide {
            return;
        }

        self.sync_active_slide();
        self.active_slide = index;
        self.canvas = self.slides[index].clone();
        self.selected_box = None;
        self.caret_index = 0;
        self.selection_anchor = None;
        self.caret_x_target = None;
        self.image_resize_drag = None;
        self.text_box_drag = None;
        self.fit_zoom_pending = true;
        self.status = format!("Selected slide {}", index + 1);
    }

    fn apply_pending_fit_zoom(&mut self, viewport_size: egui::Vec2) {
        if !self.fit_zoom_pending || viewport_size.x <= 0.0 || viewport_size.y <= 0.0 {
            return;
        }

        let canvas_size = self.canvas.size;
        if canvas_size.x <= 0.0 || canvas_size.y <= 0.0 {
            return;
        }

        let usable_size = viewport_size * 0.98;
        let fit_zoom = (usable_size.x / canvas_size.x)
            .min(usable_size.y / canvas_size.y)
            .clamp(0.4, 3.0);
        self.canvas.zoom = fit_zoom;
        if let Some(slide) = self.slides.get_mut(self.active_slide) {
            slide.zoom = fit_zoom;
        }
        self.fit_zoom_pending = false;
    }

    fn advance_animation_phase(&mut self) {
        self.phase += 0.03;

        for render_box in &mut self.canvas.boxes {
            if let Some(animation) = &mut render_box.animation {
                animation.phase = self.phase;
            }
        }

        self.canvas.relayout(CanvasMode::SlideDeck);
    }

    fn draw_file_menu(&mut self, ui: &mut egui::Ui) {
        ui.menu_button("File", |ui| {
            if ui.button("Reload Default ODP").clicked() {
                self.reload_default_odp();
                ui.close();
            }

            ui.add_enabled(false, egui::Button::new("Open"));
            ui.add_enabled(false, egui::Button::new("Save"));
            ui.add_enabled(false, egui::Button::new("Save as..."));
        });
    }

    fn draw_insert_menu(&mut self, ui: &mut egui::Ui) {
        ui.menu_button("Insert", |ui| {
            if ui.button("Text Box").clicked() {
                self.insert_text_box();
                ui.close();
            }
            if ui.button("Image Placeholder").clicked() {
                self.insert_image_box();
                ui.close();
            }
            if ui.button("Table").clicked() {
                self.insert_table_box();
                ui.close();
            }
        });
    }

    fn draw_toolbar(&mut self, ui: &mut egui::Ui) {
        ui.horizontal(|ui| {
            let has_text_selection = self
                .selection_range()
                .is_some_and(|range| !range.is_empty());

            if ui
                .add_enabled(has_text_selection, format_toolbar_text_button("B", false))
                .on_hover_text("Bold")
                .clicked()
            {
                self.toggle_selected_formatting(FormattingMode::Bold);
            }
            if ui
                .add_enabled(has_text_selection, format_toolbar_text_button("I", false))
                .on_hover_text("Italic")
                .clicked()
            {
                self.toggle_selected_formatting(FormattingMode::Italic);
            }
            if ui
                .add_enabled(has_text_selection, format_toolbar_text_button("U", false))
                .on_hover_text("Underline")
                .clicked()
            {
                self.toggle_selected_formatting(FormattingMode::Underline);
            }

            ui.separator();
            self.draw_font_size_controls(ui);
            ui.separator();
            self.draw_color_menu(ui);
            self.draw_highlight_menu(ui);
            ui.separator();
            self.draw_alignment_buttons(ui);
            ui.separator();
            ui.label(format!("Zoom: {:.0}%", self.canvas.zoom * 100.0));
            ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                ui.label(&self.status);
            });
        });
    }

    fn draw_left_panel(&mut self, ctx: &egui::Context) {
        let panel = egui::SidePanel::left("left_menu_panel");

        if self.left_panel_open {
            panel
                .resizable(true)
                .default_width(SIDE_PANEL_DEFAULT_WIDTH)
                .width_range(SIDE_PANEL_MIN_WIDTH..=SIDE_PANEL_MAX_WIDTH)
                .show(ctx, |ui| {
                    ui.horizontal(|ui| {
                        ui.heading("Slides");
                        ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                            if ui
                                .button("<")
                                .on_hover_text("Collapse left panel")
                                .clicked()
                            {
                                self.left_panel_open = false;
                            }
                        });
                    });

                    ui.separator();
                    ui.horizontal(|ui| {
                        ui.selectable_value(
                            &mut self.left_panel_tab,
                            LeftPanelTab::Slides,
                            "Slides",
                        );
                        ui.selectable_value(
                            &mut self.left_panel_tab,
                            LeftPanelTab::Scheduler,
                            "Scheduler",
                        );
                    });
                    ui.separator();

                    match self.left_panel_tab {
                        LeftPanelTab::Slides => self.draw_slide_navigator(ui),
                        LeftPanelTab::Scheduler => {
                            ui.label("Scheduler presets");
                            if ui.button("Draft Deck").clicked() {
                                self.status = "Scheduler preset: Draft Deck".to_owned();
                            }
                            if ui.button("Design Pass").clicked() {
                                self.status = "Scheduler preset: Design Pass".to_owned();
                            }
                            if ui.button("Presentation Polish").clicked() {
                                self.status = "Scheduler preset: Presentation Polish".to_owned();
                            }
                        }
                    }
                });
            return;
        }

        panel
            .resizable(false)
            .exact_width(COLLAPSED_SIDE_RAIL_WIDTH)
            .show(ctx, |ui| {
                ui.vertical_centered(|ui| {
                    if ui.button(">").on_hover_text("Expand left panel").clicked() {
                        self.left_panel_open = true;
                    }
                });
            });
    }

    fn draw_alignment_buttons(&mut self, ui: &mut egui::Ui) {
        let selected_alignment = self
            .selected_text_style_state()
            .map(|state| state.alignment);
        let enabled = selected_alignment.is_some();

        let left = ui
            .add_enabled(
                enabled,
                format_alignment_button("|← ", selected_alignment == Some(TextAlignment::Left)),
            )
            .on_hover_text("Align left");
        if left.clicked() {
            self.set_selected_alignment(TextAlignment::Left);
        }

        let justify = ui
            .add_enabled(
                enabled,
                format_alignment_button("|↔|", selected_alignment == Some(TextAlignment::Justify)),
            )
            .on_hover_text("Fill line");
        if justify.clicked() {
            self.set_selected_alignment(TextAlignment::Justify);
        }

        let center = ui
            .add_enabled(
                enabled,
                format_alignment_button(" ↔ ", selected_alignment == Some(TextAlignment::Center)),
            )
            .on_hover_text("Align center");
        if center.clicked() {
            self.set_selected_alignment(TextAlignment::Center);
        }

        let right = ui
            .add_enabled(
                enabled,
                format_alignment_button(" →|", selected_alignment == Some(TextAlignment::Right)),
            )
            .on_hover_text("Align right");
        if right.clicked() {
            self.set_selected_alignment(TextAlignment::Right);
        }
    }

    fn draw_font_size_controls(&mut self, ui: &mut egui::Ui) {
        let style_state = self.selected_text_style_state();
        let enabled = style_state.is_some();
        let active_font_size = style_state.map(|state| state.font_size).unwrap_or_default();

        ui.label("Size");
        if ui
            .add_enabled(
                enabled && active_font_size > MIN_FONT_SIZE_PT * PT_TO_PX,
                format_font_size_step_button("-"),
            )
            .on_hover_text("Decrease font size")
            .clicked()
        {
            self.step_selected_font_size(-1.0);
        }

        let label = if enabled {
            format!("{:.0} pt", active_font_size / PT_TO_PX)
        } else {
            "-- pt".to_owned()
        };
        ui.add_enabled(false, egui::Label::new(label));

        if ui
            .add_enabled(
                enabled && active_font_size < MAX_FONT_SIZE_PT * PT_TO_PX,
                format_font_size_step_button("+"),
            )
            .on_hover_text("Increase font size")
            .clicked()
        {
            self.step_selected_font_size(1.0);
        }
    }

    fn draw_color_menu(&mut self, ui: &mut egui::Ui) {
        let Some(style_state) = self.selected_text_style_state() else {
            ui.add_enabled(false, egui::Button::new("   "));
            return;
        };

        let active_color = style_state.color;
        ui.scope(|ui| {
            ui.spacing_mut().button_padding = egui::vec2(12.0, 2.0);
            ui.visuals_mut().widgets.inactive.weak_bg_fill = active_color;
            ui.visuals_mut().widgets.hovered.weak_bg_fill = active_color;
            ui.visuals_mut().widgets.active.weak_bg_fill = active_color;
            ui.visuals_mut().widgets.inactive.bg_stroke =
                egui::Stroke::new(1.5, egui::Color32::WHITE);
            ui.visuals_mut().widgets.hovered.bg_stroke =
                egui::Stroke::new(1.5, egui::Color32::WHITE);
            ui.visuals_mut().widgets.active.bg_stroke =
                egui::Stroke::new(1.5, egui::Color32::WHITE);

            ui.menu_button("   ", |ui| {
                ui.horizontal(|ui| {
                    for color in TEXT_COLOR_PALETTE {
                        let response = color_swatch(ui, Some(color), color == active_color);

                        if response.clicked() {
                            self.set_selected_text_color(color);
                            ui.close();
                        }
                    }
                });
                ui.horizontal(|ui| {
                    for color in self.recent_text_custom_colors {
                        let custom_selected = color == Some(active_color);
                        let response = recent_color_swatch(ui, color, custom_selected);
                        if response.clicked() {
                            let Some(color) = color else {
                                continue;
                            };
                            self.set_selected_text_color(color);
                            ui.close();
                        }
                    }
                });
                ui.separator();
                ui.horizontal(|ui| {
                    if let Some(color) = rgb_picker_menu(ui, &mut self.text_rgb_picker) {
                        push_recent_color(&mut self.recent_text_custom_colors, color);
                        self.set_selected_text_color(color);
                        ui.close();
                    }
                });
            });
        });
    }

    fn draw_highlight_menu(&mut self, ui: &mut egui::Ui) {
        let Some(style_state) = self.selected_text_style_state() else {
            ui.add_enabled(false, egui::Button::new("HL"));
            return;
        };

        let active_color = style_state.background_color;
        let swatch_color = active_color.unwrap_or(egui::Color32::from_gray(35));
        ui.scope(|ui| {
            ui.spacing_mut().button_padding = egui::vec2(12.0, 2.0);
            ui.visuals_mut().widgets.inactive.weak_bg_fill = swatch_color;
            ui.visuals_mut().widgets.hovered.weak_bg_fill = swatch_color;
            ui.visuals_mut().widgets.active.weak_bg_fill = swatch_color;
            ui.visuals_mut().widgets.inactive.bg_stroke =
                egui::Stroke::new(1.5, egui::Color32::WHITE);
            ui.visuals_mut().widgets.hovered.bg_stroke =
                egui::Stroke::new(1.5, egui::Color32::WHITE);
            ui.visuals_mut().widgets.active.bg_stroke =
                egui::Stroke::new(1.5, egui::Color32::WHITE);

            ui.menu_button("HL", |ui| {
                ui.horizontal(|ui| {
                    for color in HIGHLIGHT_COLOR_PALETTE {
                        let response = color_swatch(ui, color, color == active_color);

                        if response.clicked() {
                            self.set_selected_highlight_color(color);
                            ui.close();
                        }
                    }
                });
                ui.horizontal(|ui| {
                    for color in self.recent_highlight_custom_colors {
                        let custom_selected = color == active_color;
                        let response = recent_color_swatch(ui, color, custom_selected);
                        if response.clicked() {
                            let Some(color) = color else {
                                continue;
                            };
                            self.set_selected_highlight_color(Some(color));
                            ui.close();
                        }
                    }
                });
                ui.separator();
                ui.horizontal(|ui| {
                    if let Some(color) = rgb_picker_menu(ui, &mut self.highlight_rgb_picker) {
                        push_recent_color(&mut self.recent_highlight_custom_colors, color);
                        self.set_selected_highlight_color(Some(color));
                        ui.close();
                    }
                });
            });
        });
    }

    fn draw_slide_navigator(&mut self, ui: &mut egui::Ui) {
        self.sync_active_slide();
        ui.label(format!("File: {}", self.document_name));
        ui.label(format!("Slides: {}", self.slides.len()));
        ui.label(format!("Objects: {}", self.canvas.boxes.len()));
        ui.label(format!("Status: {}", self.status));
        ui.separator();

        let mut clicked_slide = None;
        for index in 0..self.slides.len() {
            let thumb_size = egui::vec2(ui.available_width().max(120.0), 92.0);
            let (rect, response) = ui.allocate_exact_size(thumb_size, egui::Sense::click());
            let painter = ui.painter_at(rect);
            let selected = index == self.active_slide;
            painter.rect_filled(
                rect,
                4.0,
                if selected {
                    egui::Color32::from_rgb(54, 62, 72)
                } else {
                    egui::Color32::from_rgb(0x23, 0x23, 0x23)
                },
            );
            let slide_rect = egui::Rect::from_center_size(
                rect.center() + egui::vec2(10.0, 0.0),
                egui::vec2(128.0, 72.0),
            );
            painter.text(
                egui::pos2(rect.left() + 10.0, slide_rect.center().y),
                egui::Align2::CENTER_CENTER,
                (index + 1).to_string(),
                egui::FontId::proportional(16.0),
                egui::Color32::WHITE,
            );
            self.slides[index].paint_page_preview(&painter, slide_rect);
            painter.rect_stroke(
                slide_rect,
                2.0,
                egui::Stroke::new(
                    if selected { 2.0 } else { 1.0 },
                    if selected {
                        egui::Color32::from_rgb(20, 120, 220)
                    } else {
                        egui::Color32::from_gray(170)
                    },
                ),
                egui::StrokeKind::Inside,
            );

            if response.clicked() {
                clicked_slide = Some(index);
            }

            ui.add_space(8.0);
        }

        if let Some(index) = clicked_slide {
            self.select_slide(index);
        }
    }
    fn draw_right_panel(&mut self, ctx: &egui::Context) {
        let panel = egui::SidePanel::right("right_menu_panel");

        if self.right_panel_open {
            panel
                .resizable(true)
                .default_width(SIDE_PANEL_DEFAULT_WIDTH)
                .width_range(SIDE_PANEL_MIN_WIDTH..=SIDE_PANEL_MAX_WIDTH)
                .show(ctx, |ui| {
                    ui.horizontal(|ui| {
                        ui.heading("Properties");
                        ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                            if ui
                                .button(">")
                                .on_hover_text("Collapse right panel")
                                .clicked()
                            {
                                self.right_panel_open = false;
                            }
                        });
                    });

                    ui.separator();
                    ui.horizontal(|ui| {
                        ui.selectable_value(
                            &mut self.right_panel_tab,
                            RightPanelTab::Properties,
                            "Style",
                        );
                        ui.selectable_value(
                            &mut self.right_panel_tab,
                            RightPanelTab::Animation,
                            "Animation",
                        );
                        ui.selectable_value(
                            &mut self.right_panel_tab,
                            RightPanelTab::Image,
                            "Images",
                        );
                        ui.selectable_value(
                            &mut self.right_panel_tab,
                            RightPanelTab::Scheduler,
                            "Scheduler",
                        );
                    });
                    ui.separator();

                    match self.right_panel_tab {
                        RightPanelTab::Properties => self.draw_selected_properties(ui),
                        RightPanelTab::Image => self.draw_image_panel(ui),
                        RightPanelTab::Animation => self.draw_animation_panel(ui),
                        RightPanelTab::Scheduler => {
                            ui.label("Scheduler presets");
                            if ui.button("Morning Edit").clicked() {
                                self.status = "Scheduler preset: Morning Edit".to_owned();
                            }
                            if ui.button("Deadline Cleanup").clicked() {
                                self.status = "Scheduler preset: Deadline Cleanup".to_owned();
                            }
                            if ui.button("Final Rehearsal").clicked() {
                                self.status = "Scheduler preset: Final Rehearsal".to_owned();
                            }
                        }
                    }
                });
            return;
        }

        panel
            .resizable(false)
            .exact_width(COLLAPSED_SIDE_RAIL_WIDTH)
            .show(ctx, |ui| {
                ui.vertical_centered(|ui| {
                    if ui.button("<").on_hover_text("Expand right panel").clicked() {
                        self.right_panel_open = true;
                    }
                });
            });
    }

    fn draw_selected_properties(&mut self, ui: &mut egui::Ui) {
        let Some(selected_id) = self.selected_box else {
            ui.label("No object selected");
            return;
        };
        let selected_text_range = self.selection_range().filter(|range| !range.is_empty());

        let Some(render_box) = self.canvas.box_mut(selected_id) else {
            ui.label("No object selected");
            return;
        };

        ui.label(format!("Object #{selected_id}"));
        ui.separator();

        let mut changed = false;
        changed |= draw_f32_drag_field(ui, "X", &mut render_box.position.x, -4096.0, 4096.0);
        changed |= draw_f32_drag_field(ui, "Y", &mut render_box.position.y, -4096.0, 4096.0);
        changed |= draw_f32_drag_field(ui, "Width", &mut render_box.size.x, 1.0, 4096.0);
        changed |= draw_f32_drag_field(ui, "Height", &mut render_box.size.y, 1.0, 4096.0);
        changed |= draw_f32_drag_field(ui, "Rotation", &mut render_box.rotation, -360.0, 360.0);

        if let Some(style_state) = render_box.text_style_state() {
            ui.separator();
            let mut font_size_pt = style_state.font_size / PT_TO_PX;
            if draw_f32_drag_field(
                ui,
                "Font size",
                &mut font_size_pt,
                MIN_FONT_SIZE_PT,
                MAX_FONT_SIZE_PT,
            ) {
                changed |= render_box.set_text_font_size(
                    selected_text_range,
                    (font_size_pt * PT_TO_PX)
                        .clamp(MIN_FONT_SIZE_PT * PT_TO_PX, MAX_FONT_SIZE_PT * PT_TO_PX),
                );
            }
            ui.label(format!(
                "Bold: {}",
                if style_state.bold { "on" } else { "off" }
            ));
            ui.label(format!(
                "Italic: {}",
                if style_state.italic { "on" } else { "off" }
            ));
            ui.label(format!(
                "Underline: {}",
                if style_state.underline { "on" } else { "off" }
            ));
        }

        if changed {
            self.status = format!("Updated object #{selected_id}");
            self.canvas.relayout(CanvasMode::SlideDeck);
        }
    }

    fn draw_image_panel(&mut self, ui: &mut egui::Ui) {
        let Some(selected_id) = self.selected_box else {
            ui.label("No image selected");
            ui.label("Select an image object to inspect its parameters.");
            return;
        };

        let Some(render_box) = self.canvas.box_mut(selected_id) else {
            ui.label("No image selected");
            return;
        };

        if !render_box.is_image() {
            ui.label("No image selected");
            ui.label("Select an image object to inspect its parameters.");
            return;
        }

        let mut position = render_box.position;
        let mut scale = render_box.scale;
        let mut rotation = render_box.rotation;
        let mut path = render_box
            .image_ref()
            .map(|image| image.path.display().to_string())
            .unwrap_or_default();
        let mut image_width = render_box
            .image_ref()
            .map(|image| image.size.x)
            .unwrap_or(1.0);
        let mut image_height = render_box
            .image_ref()
            .map(|image| image.size.y)
            .unwrap_or(1.0);
        let mut margin_left = render_box
            .image_ref()
            .map(|image| image.margin_left)
            .unwrap_or_default();
        let mut margin_right = render_box
            .image_ref()
            .map(|image| image.margin_right)
            .unwrap_or_default();
        let mut margin_top = render_box
            .image_ref()
            .map(|image| image.margin_top)
            .unwrap_or_default();
        let mut margin_bottom = render_box
            .image_ref()
            .map(|image| image.margin_bottom)
            .unwrap_or_default();
        let mut center_horizontally = render_box
            .image_ref()
            .map(|image| image.center_horizontally)
            .unwrap_or_default();
        let aspect_ratio = if image_height > 0.0 {
            image_width / image_height
        } else {
            1.0
        };

        ui.label(format!("Selected image #{selected_id}"));

        let mut image_changed = false;
        let mut reload_error = None;

        ui.add_space(IMAGE_PANEL_SECTION_GAP);
        ui.label("Source");
        ui.horizontal(|ui| {
            ui.label("Path");
            let response = ui.text_edit_singleline(&mut path);
            if response.lost_focus() && response.changed() {
                match render_box
                    .image_mut()
                    .and_then(|image| image.reload_from_path(path.trim()).err())
                {
                    Some(error) => reload_error = Some(error.to_string()),
                    None => image_changed = true,
                }
            }
        });

        ui.add_space(IMAGE_PANEL_SECTION_GAP);
        ui.label("Position");
        image_changed |= draw_f32_drag_field(ui, "X", &mut position.x, -4096.0, 4096.0);
        image_changed |= draw_f32_drag_field(ui, "Y", &mut position.y, -4096.0, 4096.0);

        ui.add_space(IMAGE_PANEL_SECTION_GAP);
        ui.label("Image dimensions");
        ui.horizontal(|ui| {
            ui.label("Keep ratio");
            image_changed |= ui.checkbox(&mut self.link_image_dimensions, "").changed();
        });

        let width_changed = draw_f32_drag_field(ui, "Width", &mut image_width, 1.0, 4096.0);
        if width_changed && self.link_image_dimensions && aspect_ratio > 0.0 {
            image_height = (image_width / aspect_ratio).max(1.0);
        }
        image_changed |= width_changed;

        let height_changed = draw_f32_drag_field(ui, "Height", &mut image_height, 1.0, 4096.0);
        if height_changed && self.link_image_dimensions && aspect_ratio > 0.0 {
            image_width = (image_height * aspect_ratio).max(1.0);
        }
        image_changed |= height_changed;

        ui.add_space(IMAGE_PANEL_SECTION_GAP);
        ui.label("Scale");
        image_changed |= draw_f32_drag_field(ui, "Scale X", &mut scale.x, 0.05, 20.0);
        image_changed |= draw_f32_drag_field(ui, "Scale Y", &mut scale.y, 0.05, 20.0);
        image_changed |= draw_f32_drag_field(ui, "Rotation", &mut rotation, -360.0, 360.0);

        ui.add_space(IMAGE_PANEL_SECTION_GAP);
        ui.label("Margins");
        image_changed |= draw_f32_drag_field(ui, "Margin L", &mut margin_left, 0.0, 2048.0);
        image_changed |= draw_f32_drag_field(ui, "Margin R", &mut margin_right, 0.0, 2048.0);
        image_changed |= draw_f32_drag_field(ui, "Margin T", &mut margin_top, 0.0, 2048.0);
        image_changed |= draw_f32_drag_field(ui, "Margin B", &mut margin_bottom, 0.0, 2048.0);

        ui.add_space(IMAGE_PANEL_SECTION_GAP);
        ui.label("Anchor");
        ui.horizontal(|ui| {
            ui.label("Centered");
            image_changed |= ui.checkbox(&mut center_horizontally, "").changed();
        });

        if image_changed {
            render_box.position = position;
            render_box.scale = egui::vec2(scale.x.max(0.05), scale.y.max(0.05));
            render_box.rotation = rotation;
            if let Some(image) = render_box.image_mut() {
                image.size = egui::vec2(image_width.max(1.0), image_height.max(1.0));
                image.margin_left = margin_left.max(0.0);
                image.margin_right = margin_right.max(0.0);
                image.margin_top = margin_top.max(0.0);
                image.margin_bottom = margin_bottom.max(0.0);
                image.center_horizontally = center_horizontally;
            }
            self.canvas.relayout(CanvasMode::SlideDeck);
            self.status = format!("Updated image #{selected_id} properties");
        }

        if let Some(error) = reload_error {
            self.status = format!("Image reload failed: {error}");
        }
    }

    fn draw_animation_panel(&mut self, ui: &mut egui::Ui) {
        let Some(selected_id) = self.selected_box else {
            ui.label("No object selected");
            return;
        };

        let Some(render_box) = self.canvas.box_mut(selected_id) else {
            ui.label("No object selected");
            return;
        };

        let mut enabled = render_box.animation.is_some();
        if ui.checkbox(&mut enabled, "Preview animation").changed() {
            render_box.animation = enabled.then_some(AnimationSpec {
                phase: self.phase,
                amplitude: egui::vec2(6.0, 0.0),
            });
            self.status = format!("Updated animation for object #{selected_id}");
        }

        if let Some(animation) = &mut render_box.animation {
            ui.separator();
            let mut changed = false;
            changed |=
                draw_f32_drag_field(ui, "Amplitude X", &mut animation.amplitude.x, -80.0, 80.0);
            changed |=
                draw_f32_drag_field(ui, "Amplitude Y", &mut animation.amplitude.y, -80.0, 80.0);
            if changed {
                self.status = format!("Updated animation for object #{selected_id}");
            }
        }
    }

    fn insert_text_box(&mut self) {
        let id = self.allocate_box_id();
        let mut text_box = RenderBox::text(
            id,
            LayoutRole::Absolute,
            vec![TextRun::new("Text box", TextStyle::body())],
        );
        text_box.position = egui::pos2(140.0, 220.0);
        text_box.size.x = 320.0;
        self.canvas.push(text_box);
        self.canvas.relayout(CanvasMode::SlideDeck);
        self.select_box(id, Some(0));
        self.status = format!("Inserted text box #{id}");
    }

    fn insert_image_box(&mut self) {
        let id = self.allocate_box_id();
        let mut image = RenderBox::image(
            id,
            LayoutRole::Absolute,
            "Image Placeholder",
            egui::vec2(320.0, 180.0),
        );
        image.position = egui::pos2(660.0, 250.0);
        image.style.fill = egui::Color32::from_rgb(246, 248, 250);
        self.canvas.push(image);
        self.canvas.relayout(CanvasMode::SlideDeck);
        self.select_box(id, None);
        self.status = format!("Inserted image placeholder #{id}");
    }

    fn insert_table_box(&mut self) {
        let id = self.allocate_box_id();
        let mut table = RenderBox::table(
            id,
            LayoutRole::Absolute,
            TableBlock {
                rows: vec![
                    vec!["Item".into(), "Status".into()],
                    vec!["Topic".into(), "Draft".into()],
                    vec!["Visual".into(), "Pending".into()],
                ],
            },
        );
        table.position = egui::pos2(150.0, 430.0);
        table.style.fill = egui::Color32::from_rgb(252, 250, 246);
        self.canvas.push(table);
        self.canvas.relayout(CanvasMode::SlideDeck);
        self.select_box(id, None);
        self.status = format!("Inserted table #{id}");
    }

    fn allocate_box_id(&mut self) -> u64 {
        let id = self.next_box_id;
        self.next_box_id += 1;
        id
    }

    fn select_box(&mut self, box_id: u64, clicked_caret: Option<usize>) {
        self.selected_box = Some(box_id);
        let text_len = self
            .canvas
            .box_ref(box_id)
            .and_then(RenderBox::text_len)
            .unwrap_or(0);
        self.caret_index = clicked_caret.unwrap_or(text_len).min(text_len);
        self.selection_anchor = None;
        self.caret_x_target = None;
        self.right_panel_tab = if self.canvas.box_ref(box_id).is_some_and(RenderBox::is_image) {
            RightPanelTab::Image
        } else {
            RightPanelTab::Properties
        };
    }

    fn start_image_resize_drag(
        &mut self,
        box_id: u64,
        handle: ImageResizeHandle,
        pointer: egui::Pos2,
    ) {
        let Some(render_box) = self.canvas.box_ref(box_id).filter(|item| item.is_image()) else {
            return;
        };
        let Some(image) = render_box.image_ref() else {
            return;
        };

        self.image_resize_drag = Some(ImageResizeDrag {
            box_id,
            handle,
            start_pointer: pointer,
            start_position: render_box.position,
            start_size: render_box.size,
            start_scale: render_box.scale,
            start_image_size: image.size,
            start_rotation: render_box.rotation,
            zoom: self.canvas.zoom.max(0.01),
        });
        self.text_box_drag = None;
        self.selected_box = Some(box_id);
        self.selection_anchor = None;
        self.caret_x_target = None;
        self.right_panel_tab = RightPanelTab::Image;
    }

    fn start_text_box_drag(&mut self, box_id: u64, handle: ImageResizeHandle, pointer: egui::Pos2) {
        let Some(render_box) = self
            .canvas
            .box_ref(box_id)
            .filter(|item| item.is_text() && item.rotation.abs() <= f32::EPSILON)
        else {
            return;
        };

        self.text_box_drag = Some(TextBoxDrag {
            box_id,
            handle,
            start_pointer: pointer,
            start_position: render_box.position,
            start_size: render_box.size,
            zoom: self.canvas.zoom.max(0.01),
        });
        self.image_resize_drag = None;
        self.selected_box = Some(box_id);
        self.selection_anchor = None;
        self.caret_x_target = None;
        self.right_panel_tab = RightPanelTab::Properties;
    }

    fn drag_text_box_to(&mut self, pointer: egui::Pos2) {
        let Some(drag) = self.text_box_drag else {
            return;
        };
        let delta = (pointer - drag.start_pointer) / drag.zoom;
        let Some(render_box) = self.canvas.box_mut(drag.box_id) else {
            self.text_box_drag = None;
            return;
        };

        if drag.handle == ImageResizeHandle::Center {
            render_box.position = drag.start_position + delta;
            render_box.size = drag.start_size;
            let position = render_box.position;
            self.canvas.relayout(CanvasMode::SlideDeck);
            self.status = format!("Text box moved to {:.0}, {:.0}", position.x, position.y);
            return;
        }

        let left = matches!(
            drag.handle,
            ImageResizeHandle::Left | ImageResizeHandle::TopLeft | ImageResizeHandle::BottomLeft
        );
        let right = matches!(
            drag.handle,
            ImageResizeHandle::Right | ImageResizeHandle::TopRight | ImageResizeHandle::BottomRight
        );
        let top = matches!(
            drag.handle,
            ImageResizeHandle::Top | ImageResizeHandle::TopLeft | ImageResizeHandle::TopRight
        );
        let bottom = matches!(
            drag.handle,
            ImageResizeHandle::Bottom
                | ImageResizeHandle::BottomLeft
                | ImageResizeHandle::BottomRight
        );

        let mut position = drag.start_position;
        let mut size = drag.start_size;
        if left {
            let next_width = (drag.start_size.x - delta.x).max(80.0);
            position.x += drag.start_size.x - next_width;
            size.x = next_width;
        } else if right {
            size.x = (drag.start_size.x + delta.x).max(80.0);
        }

        if top {
            let next_height = (drag.start_size.y - delta.y).max(48.0);
            position.y += drag.start_size.y - next_height;
            size.y = next_height;
        } else if bottom {
            size.y = (drag.start_size.y + delta.y).max(48.0);
        }

        render_box.position = position;
        render_box.size = size;
        self.canvas.relayout(CanvasMode::SlideDeck);
        self.status = format!("Text box resized to {:.0} x {:.0}", size.x, size.y);
    }

    fn drag_image_resize_to(&mut self, pointer: egui::Pos2) {
        let Some(drag) = self.image_resize_drag else {
            return;
        };
        let delta = (pointer - drag.start_pointer) / drag.zoom;
        if drag.handle == ImageResizeHandle::Center {
            let Some(render_box) = self.canvas.box_mut(drag.box_id) else {
                self.image_resize_drag = None;
                return;
            };

            render_box.position = drag.start_position + delta;
            let position = render_box.position;
            self.canvas.relayout(CanvasMode::SlideDeck);
            self.status = format!("Image moved to {:.0}, {:.0}", position.x, position.y);
            return;
        }

        let local_delta = rotate_vec(delta, -drag.start_rotation);
        let start_visual_size = egui::vec2(
            drag.start_size.x * drag.start_scale.x,
            drag.start_size.y * drag.start_scale.y,
        );
        let mut left_edge = -start_visual_size.x * 0.5;
        let mut right_edge = start_visual_size.x * 0.5;
        let mut top_edge = -start_visual_size.y * 0.5;
        let mut bottom_edge = start_visual_size.y * 0.5;
        let min_size = 24.0;

        if matches!(
            drag.handle,
            ImageResizeHandle::TopLeft | ImageResizeHandle::BottomRight
        ) {
            let (width, height) =
                aspect_preserving_image_resize_size(start_visual_size, local_delta, drag.handle);
            match drag.handle {
                ImageResizeHandle::TopLeft => {
                    left_edge = right_edge - width;
                    top_edge = bottom_edge - height;
                }
                ImageResizeHandle::BottomRight => {
                    right_edge = left_edge + width;
                    bottom_edge = top_edge + height;
                }
                _ => {}
            }
        } else {
            match drag.handle {
                ImageResizeHandle::Left | ImageResizeHandle::BottomLeft => {
                    left_edge = (left_edge + local_delta.x).min(right_edge - min_size);
                }
                ImageResizeHandle::Right | ImageResizeHandle::TopRight => {
                    right_edge = (right_edge + local_delta.x).max(left_edge + min_size);
                }
                _ => {}
            }

            match drag.handle {
                ImageResizeHandle::Top | ImageResizeHandle::TopRight => {
                    top_edge = (top_edge + local_delta.y).min(bottom_edge - min_size);
                }
                ImageResizeHandle::Bottom | ImageResizeHandle::BottomLeft => {
                    bottom_edge = (bottom_edge + local_delta.y).max(top_edge + min_size);
                }
                _ => {}
            }
        }

        let visual_size = egui::vec2(right_edge - left_edge, bottom_edge - top_edge);
        let start_center = drag.start_position + start_visual_size * 0.5;
        let center_shift = rotate_vec(
            egui::vec2(
                (left_edge + right_edge) * 0.5,
                (top_edge + bottom_edge) * 0.5,
            ),
            drag.start_rotation,
        );
        let position = start_center + center_shift - visual_size * 0.5;

        let Some(render_box) = self.canvas.box_mut(drag.box_id) else {
            self.image_resize_drag = None;
            return;
        };
        let next_size = egui::vec2(
            (visual_size.x / drag.start_scale.x.max(0.01)).max(24.0),
            (visual_size.y / drag.start_scale.y.max(0.01)).max(24.0),
        );
        render_box.position = position;
        render_box.size = next_size;
        if let Some(image) = render_box.image_mut() {
            let image_scale = egui::vec2(
                (next_size.x / drag.start_size.x.max(1.0)).max(0.01),
                (next_size.y / drag.start_size.y.max(1.0)).max(0.01),
            );
            image.size = egui::vec2(
                (drag.start_image_size.x * image_scale.x).max(24.0),
                (drag.start_image_size.y * image_scale.y).max(24.0),
            );
        }
        let status_size = egui::vec2(
            next_size.x * render_box.scale.x,
            next_size.y * render_box.scale.y,
        );
        self.canvas.relayout(CanvasMode::SlideDeck);
        self.status = format!(
            "Image resized to {:.0} x {:.0}",
            status_size.x, status_size.y
        );
    }

    fn selection_range(&self) -> Option<TextRange> {
        match (self.selected_box, self.selection_anchor) {
            (Some(_), Some(anchor)) => Some(TextRange {
                anchor,
                caret: self.caret_index,
            }),
            _ => None,
        }
    }

    fn handle_keyboard_editing(&mut self, ctx: &egui::Context) {
        let Some(selected_id) = self.selected_box else {
            return;
        };

        let events = ctx.input(|input| input.events.clone());
        let mut relayout = false;
        let zoom = self.canvas.zoom;

        for event in events {
            let active_range = self.selection_range().filter(|r| !r.is_empty());
            match event {
                egui::Event::Text(text) if !text.is_empty() => {
                    let Some(render_box) = self.canvas.box_mut(selected_id) else {
                        continue;
                    };
                    let next_caret = if let Some(range) = active_range {
                        render_box.replace_range(range, &text)
                    } else {
                        render_box.insert_text_at(self.caret_index, &text)
                    };
                    if let Some(next_caret) = next_caret {
                        self.caret_index = next_caret;
                        self.selection_anchor = None;
                        self.caret_x_target = None;
                        relayout = true;
                    }
                }
                egui::Event::Key {
                    key,
                    pressed: true,
                    modifiers,
                    ..
                } => {
                    let Some(render_box) = self.canvas.box_mut(selected_id) else {
                        continue;
                    };

                    match key {
                        egui::Key::Backspace => {
                            let next_caret = if let Some(range) = active_range {
                                render_box.delete_range(range)
                            } else {
                                render_box.delete_char_before(self.caret_index)
                            };
                            if let Some(next_caret) = next_caret {
                                self.caret_index = next_caret;
                                self.selection_anchor = None;
                                self.caret_x_target = None;
                                relayout = true;
                            }
                        }
                        egui::Key::Delete => {
                            let next_caret = if let Some(range) = active_range {
                                render_box.delete_range(range)
                            } else {
                                render_box.delete_char_at(self.caret_index)
                            };
                            if let Some(next_caret) = next_caret {
                                self.caret_index = next_caret;
                                self.selection_anchor = None;
                                self.caret_x_target = None;
                                relayout = true;
                            }
                        }
                        egui::Key::Enter => {
                            let next_caret = if let Some(range) = active_range {
                                render_box.replace_range(range, "\n")
                            } else {
                                render_box.insert_text_at(self.caret_index, "\n")
                            };
                            if let Some(next_caret) = next_caret {
                                self.caret_index = next_caret;
                                self.selection_anchor = None;
                                self.caret_x_target = None;
                                relayout = true;
                            }
                        }
                        egui::Key::ArrowLeft => {
                            if let Some(range) = active_range {
                                self.caret_index = range.sorted().0;
                                self.selection_anchor = None;
                                self.caret_x_target = None;
                                continue;
                            }
                            if let Some(next_caret) =
                                render_box.move_caret_horizontal(self.caret_index, -1, ctx, zoom)
                            {
                                self.caret_index = next_caret;
                                self.caret_x_target = None;
                            }
                        }
                        egui::Key::ArrowRight => {
                            if let Some(range) = active_range {
                                self.caret_index = range.sorted().1;
                                self.selection_anchor = None;
                                self.caret_x_target = None;
                                continue;
                            }
                            if let Some(next_caret) =
                                render_box.move_caret_horizontal(self.caret_index, 1, ctx, zoom)
                            {
                                self.caret_index = next_caret;
                                self.caret_x_target = None;
                            }
                        }
                        egui::Key::ArrowUp => {
                            if let Some((next_caret, next_x)) = render_box.move_caret_vertical(
                                self.caret_index,
                                -1,
                                ctx,
                                zoom,
                                self.caret_x_target,
                            ) {
                                self.caret_index = next_caret;
                                self.caret_x_target = next_x;
                            }
                        }
                        egui::Key::ArrowDown => {
                            if let Some((next_caret, next_x)) = render_box.move_caret_vertical(
                                self.caret_index,
                                1,
                                ctx,
                                zoom,
                                self.caret_x_target,
                            ) {
                                self.caret_index = next_caret;
                                self.caret_x_target = next_x;
                            }
                        }
                        egui::Key::Home => {
                            if let Some(next_caret) =
                                render_box.line_boundary(self.caret_index, false, ctx, zoom)
                            {
                                self.caret_index = next_caret;
                                self.caret_x_target = None;
                            }
                        }
                        egui::Key::End => {
                            if let Some(next_caret) =
                                render_box.line_boundary(self.caret_index, true, ctx, zoom)
                            {
                                self.caret_index = next_caret;
                                self.caret_x_target = None;
                            }
                        }
                        egui::Key::A if modifiers.command => {
                            if let Some(len) = render_box.text_len() {
                                self.caret_index = len;
                                self.selection_anchor = Some(0);
                                self.caret_x_target = None;
                            }
                        }
                        _ => {}
                    }
                }
                _ => {}
            }
        }

        let Some(render_box) = self.canvas.box_ref(selected_id) else {
            return;
        };

        if let Some(len) = render_box.text_len() {
            self.caret_index = self.caret_index.min(len);
        }

        if relayout {
            self.canvas.relayout(CanvasMode::SlideDeck);
            self.status = "Edited text".to_owned();
        }
    }

    fn toggle_selected_formatting(&mut self, mode: FormattingMode) {
        let Some(selected_id) = self.selected_box else {
            return;
        };
        let range = self.selection_range().filter(|range| !range.is_empty());

        let changed = match self.canvas.box_mut(selected_id) {
            Some(render_box) => match mode {
                FormattingMode::Bold => render_box.toggle_bold(range),
                FormattingMode::Italic => render_box.toggle_italic(range),
                FormattingMode::Underline => render_box.toggle_underline(range),
            },
            None => false,
        };

        if changed {
            self.canvas.relayout(CanvasMode::SlideDeck);
            self.status = format!("Formatted object #{selected_id}");
        }
    }

    fn set_selected_alignment(&mut self, alignment: TextAlignment) {
        let Some(selected_id) = self.selected_box else {
            return;
        };

        let changed = self
            .canvas
            .box_mut(selected_id)
            .is_some_and(|render_box| render_box.set_text_alignment(alignment));

        if changed {
            self.canvas.relayout(CanvasMode::SlideDeck);
            self.status = format!("Aligned object #{selected_id}");
        }
    }

    fn set_selected_text_color(&mut self, color: egui::Color32) {
        let Some(selected_id) = self.selected_box else {
            return;
        };
        let range = self.selection_range().filter(|range| !range.is_empty());

        let changed = self
            .canvas
            .box_mut(selected_id)
            .is_some_and(|render_box| render_box.set_text_color(range, color));

        if changed {
            self.canvas.relayout(CanvasMode::SlideDeck);
            self.status = format!("Changed text color for object #{selected_id}");
        }
    }

    fn set_selected_highlight_color(&mut self, color: Option<egui::Color32>) {
        let Some(selected_id) = self.selected_box else {
            return;
        };
        let range = self.selection_range().filter(|range| !range.is_empty());

        let changed = self
            .canvas
            .box_mut(selected_id)
            .is_some_and(|render_box| render_box.set_highlight_color(range, color));

        if changed {
            self.canvas.relayout(CanvasMode::SlideDeck);
            self.status = format!("Changed highlight for object #{selected_id}");
        }
    }

    fn step_selected_font_size(&mut self, step_pt: f32) {
        let Some(style_state) = self.selected_text_style_state() else {
            return;
        };
        let current_pt = style_state.font_size / PT_TO_PX;
        self.set_selected_font_size_pt(current_pt + step_pt);
    }

    fn set_selected_font_size_pt(&mut self, font_size_pt: f32) {
        let Some(selected_id) = self.selected_box else {
            return;
        };
        let range = self.selection_range().filter(|range| !range.is_empty());
        let font_size = (font_size_pt * PT_TO_PX)
            .clamp(MIN_FONT_SIZE_PT * PT_TO_PX, MAX_FONT_SIZE_PT * PT_TO_PX);

        let changed = self
            .canvas
            .box_mut(selected_id)
            .is_some_and(|render_box| render_box.set_text_font_size(range, font_size));

        if changed {
            self.canvas.relayout(CanvasMode::SlideDeck);
            self.status = format!("Changed font size for object #{selected_id}");
        }
    }

    fn selected_text_style_state(&self) -> Option<TextStyleState> {
        let selected_id = self.selected_box?;
        let render_box = self.canvas.box_ref(selected_id)?;
        if let Some(range) = self.selection_range().filter(|range| !range.is_empty()) {
            render_box.text_style_state_in_range(range)
        } else {
            render_box.text_style_state()
        }
    }

    fn show_formatting_popup(&mut self, ctx: &egui::Context, canvas_min: egui::Pos2) {
        let Some(selected_id) = self.selected_box else {
            return;
        };
        let range = self.selection_range().filter(|range| !range.is_empty());
        let Some(range) = range else {
            return;
        };
        let Some(render_box) = self
            .canvas
            .box_ref(selected_id)
            .filter(|item| item.is_text())
        else {
            return;
        };

        let content_origin = self
            .canvas
            .page
            .content_origin(canvas_min, self.canvas.zoom);
        let rect = render_box.rect(content_origin, self.canvas.zoom);
        let popup_pos = egui::pos2(rect.center().x - 264.0, rect.top() - 104.0);
        let style_state = render_box
            .text_style_state_in_range(range)
            .unwrap_or_default();
        let visuals = PopupVisuals::from_state(style_state);

        egui::Area::new(egui::Id::new("text_format_popup_slides"))
            .order(egui::Order::Foreground)
            .fixed_pos(popup_pos)
            .show(ctx, |ui| {
                egui::Frame::popup(ui.style())
                    .corner_radius(8.0)
                    .show(ui, |ui| {
                        ui.spacing_mut().item_spacing = egui::vec2(10.0, 10.0);
                        ui.horizontal(|ui| {
                            if format_text_button(ui, "B", visuals.bold, "Bold") {
                                self.toggle_selected_formatting(FormattingMode::Bold);
                            }
                            if format_text_button(ui, "I", visuals.italic, "Italic") {
                                self.toggle_selected_formatting(FormattingMode::Italic);
                            }
                            if format_text_button(ui, "U", visuals.underline, "Underline") {
                                self.toggle_selected_formatting(FormattingMode::Underline);
                            }

                            ui.separator();
                            self.draw_font_size_controls(ui);
                            ui.separator();
                            self.draw_color_menu(ui);
                            self.draw_highlight_menu(ui);
                            ui.separator();
                            self.draw_alignment_buttons(ui);
                        });
                    });
            });
    }
}

impl App for LibeRustOfficeSlidesApp {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut Frame) {
        ctx.send_viewport_cmd(egui::ViewportCommand::Title(self.window_title()));
        self.advance_animation_phase();
        self.handle_keyboard_editing(ctx);
        ctx.request_repaint();

        egui::TopBottomPanel::top("app_menu_bar").show(ctx, |ui| {
            ui.horizontal(|ui| {
                self.draw_file_menu(ui);
                empty_menu(ui, "Edit");
                self.draw_insert_menu(ui);
                empty_menu(ui, "Settings");
                ui.menu_button("Help", |ui| {
                    ui.label("Author: mrjacob241");
                    ui.label("Made with Codex");
                });
            });
        });

        egui::TopBottomPanel::top("editor_toolbar_bar").show(ctx, |ui| {
            self.draw_toolbar(ui);
        });

        self.draw_left_panel(ctx);
        self.draw_right_panel(ctx);

        egui::CentralPanel::default().show(ctx, |ui| {
            self.apply_pending_fit_zoom(ui.available_size());
            let selection = self.selected_box.map(|box_id| CanvasSelection {
                box_id,
                caret_index: self
                    .canvas
                    .box_ref(box_id)
                    .filter(|item| item.is_text())
                    .map(|_| self.caret_index),
                range: self.selection_range(),
            });
            let canvas_response = self.canvas.ui(ui, selection);
            if canvas_response.drag_started {
                if let (Some(box_id), Some(handle), Some(pointer)) = (
                    self.selected_box,
                    canvas_response.image_resize_handle,
                    canvas_response.pointer_pos,
                ) {
                    self.start_image_resize_drag(box_id, handle, pointer);
                } else if let (Some(box_id), Some(handle), Some(pointer)) = (
                    self.selected_box,
                    canvas_response.text_resize_handle,
                    canvas_response.pointer_pos,
                ) {
                    self.start_text_box_drag(box_id, handle, pointer);
                } else if let (Some(box_id), Some(caret)) =
                    (canvas_response.dragged_box, canvas_response.dragged_caret)
                {
                    self.selected_box = Some(box_id);
                    self.caret_index = caret;
                    self.selection_anchor = Some(caret);
                    self.caret_x_target = None;
                }
            } else if canvas_response.response.dragged() {
                if let Some(pointer) = canvas_response.pointer_pos {
                    if self.image_resize_drag.is_some() {
                        self.drag_image_resize_to(pointer);
                    } else if self.text_box_drag.is_some() {
                        self.drag_text_box_to(pointer);
                    }
                }
                if self.image_resize_drag.is_none() && self.text_box_drag.is_none() {
                    if let (Some(box_id), Some(caret)) =
                        (canvas_response.dragged_box, canvas_response.dragged_caret)
                    {
                        if self.selected_box == Some(box_id) {
                            self.caret_index = caret;
                        }
                    }
                }
            } else if canvas_response.drag_stopped {
                self.image_resize_drag = None;
                self.text_box_drag = None;
            } else if let Some(clicked_id) = canvas_response.clicked_box {
                self.select_box(clicked_id, canvas_response.clicked_caret);
                self.status = format!("Selected object #{clicked_id}");
            } else if canvas_response.response.clicked() {
                self.selected_box = None;
                self.caret_index = 0;
                self.selection_anchor = None;
                self.caret_x_target = None;
                self.image_resize_drag = None;
                self.text_box_drag = None;
                self.status = "Selected slide".to_owned();
            }

            self.show_formatting_popup(ctx, canvas_response.response.rect.min);
        });
    }
}

#[derive(Clone, Copy)]
enum FormattingMode {
    Bold,
    Italic,
    Underline,
}

#[derive(Clone, Copy)]
struct PopupVisuals {
    bold: bool,
    italic: bool,
    underline: bool,
}

impl PopupVisuals {
    fn from_state(state: TextStyleState) -> Self {
        Self {
            bold: state.bold,
            italic: state.italic,
            underline: state.underline,
        }
    }
}

fn empty_menu(ui: &mut egui::Ui, label: &str) {
    ui.menu_button(label, |_ui| {});
}

fn draw_f32_drag_field(
    ui: &mut egui::Ui,
    label: &str,
    value: &mut f32,
    min_value: f32,
    max_value: f32,
) -> bool {
    let mut changed = false;
    ui.horizontal(|ui| {
        ui.label(label);
        let response = ui.add(
            egui::DragValue::new(value)
                .speed(1.0)
                .range(min_value..=max_value)
                .max_decimals(1),
        );
        changed = response.changed();
    });
    changed
}

fn rotate_vec(vector: egui::Vec2, rotation_degrees: f32) -> egui::Vec2 {
    let angle = rotation_degrees.to_radians();
    let (sin, cos) = angle.sin_cos();
    egui::vec2(
        vector.x * cos - vector.y * sin,
        vector.x * sin + vector.y * cos,
    )
}

fn aspect_preserving_image_resize_size(
    start_size: egui::Vec2,
    local_delta: egui::Vec2,
    handle: ImageResizeHandle,
) -> (f32, f32) {
    let min_size = 24.0;
    let aspect_ratio = (start_size.x / start_size.y.max(1.0)).max(0.01);
    let (raw_width, raw_height) = match handle {
        ImageResizeHandle::TopLeft => (start_size.x - local_delta.x, start_size.y - local_delta.y),
        ImageResizeHandle::BottomRight => {
            (start_size.x + local_delta.x, start_size.y + local_delta.y)
        }
        _ => (start_size.x, start_size.y),
    };
    let raw_width = raw_width.max(min_size);
    let raw_height = raw_height.max(min_size);
    let width_change = ((raw_width - start_size.x) / start_size.x.max(1.0)).abs();
    let height_change = ((raw_height - start_size.y) / start_size.y.max(1.0)).abs();

    let (mut width, mut height) = if width_change >= height_change {
        (raw_width, raw_width / aspect_ratio)
    } else {
        (raw_height * aspect_ratio, raw_height)
    };

    if width < min_size {
        width = min_size;
        height = width / aspect_ratio;
    }
    if height < min_size {
        height = min_size;
        width = height * aspect_ratio;
    }

    (width, height)
}

fn format_alignment_button(label: &str, active: bool) -> egui::Button<'_> {
    let font_id = egui::FontId::new(
        15.0,
        egui::FontFamily::Name(EDITOR_REGULAR_FONT_NAME.into()),
    );
    let mut text = egui::RichText::new(label).font(font_id);
    if active {
        text = text.color(egui::Color32::BLACK);
    }
    let mut button = egui::Button::new(text)
        .frame(true)
        .min_size(egui::vec2(44.0, 32.0));
    if active {
        button = button.fill(egui::Color32::from_rgb(232, 220, 205));
    }
    button
}

fn format_toolbar_text_button(label: &str, active: bool) -> egui::Button<'_> {
    let mut text = egui::RichText::new(label).size(20.0);
    if active {
        text = text.color(egui::Color32::BLACK);
    }
    let mut button = egui::Button::new(text)
        .frame(true)
        .min_size(egui::vec2(44.0, 36.0));
    if active {
        button = button.fill(egui::Color32::from_rgb(232, 220, 205));
    }
    button
}

fn format_font_size_step_button(label: &str) -> egui::Button<'_> {
    egui::Button::new(egui::RichText::new(label).size(16.0).strong())
        .frame(true)
        .min_size(egui::vec2(28.0, 30.0))
}

fn color_swatch(ui: &mut egui::Ui, color: Option<egui::Color32>, selected: bool) -> egui::Response {
    let (rect, response) = ui.allocate_exact_size(egui::vec2(24.0, 22.0), egui::Sense::click());
    let swatch_rect = rect.shrink(2.0);

    if let Some(color) = color {
        ui.painter().rect_filled(swatch_rect, 4.0, color);
    } else {
        ui.painter()
            .rect_filled(swatch_rect, 4.0, egui::Color32::from_gray(245));
        ui.painter().line_segment(
            [swatch_rect.left_bottom(), swatch_rect.right_top()],
            egui::Stroke::new(1.5, egui::Color32::RED),
        );
    }

    ui.painter().rect_stroke(
        swatch_rect,
        4.0,
        egui::Stroke::new(
            if selected { 2.0 } else { 1.0 },
            egui::Color32::from_gray(90),
        ),
        egui::StrokeKind::Inside,
    );

    response
}

fn recent_color_swatch(
    ui: &mut egui::Ui,
    color: Option<egui::Color32>,
    selected: bool,
) -> egui::Response {
    let (rect, response) = ui.allocate_exact_size(egui::vec2(24.0, 22.0), egui::Sense::click());
    let swatch_rect = rect.shrink(2.0);
    ui.painter().rect_filled(
        swatch_rect,
        4.0,
        color.unwrap_or(egui::Color32::from_gray(45)),
    );
    ui.painter().rect_stroke(
        swatch_rect,
        4.0,
        egui::Stroke::new(
            if selected { 2.0 } else { 1.0 },
            egui::Color32::from_gray(90),
        ),
        egui::StrokeKind::Inside,
    );

    response
}

fn push_recent_color(
    colors: &mut [Option<egui::Color32>; RECENT_COLOR_SLOT_COUNT],
    color: egui::Color32,
) {
    if let Some(index) = colors.iter().position(|recent| *recent == Some(color)) {
        colors[..=index].rotate_right(1);
        colors[0] = Some(color);
        return;
    }

    colors.rotate_right(1);
    colors[0] = Some(color);
}

fn rgb_picker_menu(ui: &mut egui::Ui, picker: &mut [u8; 3]) -> Option<egui::Color32> {
    let mut picked = None;

    ui.menu_button("RGB", |ui| {
        ui.set_min_width(280.0);
        let color = egui::Color32::from_rgb(picker[0], picker[1], picker[2]);

        ui.horizontal(|ui| {
            ui.vertical(|ui| {
                rgb_slider(ui, "R", &mut picker[0]);
                rgb_slider(ui, "G", &mut picker[1]);
                rgb_slider(ui, "B", &mut picker[2]);
            });

            let (rect, response) =
                ui.allocate_exact_size(egui::vec2(72.0, 78.0), egui::Sense::click());
            let preview_stroke_color = if response.hovered() {
                egui::Color32::from_rgb(255, 165, 0)
            } else {
                egui::Color32::from_gray(90)
            };
            ui.painter().rect_filled(rect.shrink(3.0), 4.0, color);
            ui.painter().rect_stroke(
                rect.shrink(3.0),
                4.0,
                egui::Stroke::new(
                    if response.hovered() { 2.0 } else { 1.0 },
                    preview_stroke_color,
                ),
                egui::StrokeKind::Inside,
            );
            if response.clicked() {
                picked = Some(color);
                ui.close();
            }
        });

        ui.horizontal(|ui| {
            let _ = color_swatch(ui, Some(color), false);
            ui.label(format!("R {}  G {}  B {}", picker[0], picker[1], picker[2]));
            if ui.button("Apply").clicked() {
                picked = Some(color);
                ui.close();
            }
        });
    });

    picked
}

fn rgb_slider(ui: &mut egui::Ui, label: &str, value: &mut u8) {
    ui.add(
        egui::Slider::new(value, 0..=255)
            .text(label)
            .show_value(true),
    );
}

fn format_text_button(ui: &mut egui::Ui, label: &str, active: bool, tooltip: &str) -> bool {
    let mut text = egui::RichText::new(label).size(18.0);
    if active {
        text = text.color(egui::Color32::BLACK);
    }
    let mut button = egui::Button::new(text)
        .frame(true)
        .min_size(egui::vec2(42.0, 42.0));
    if active {
        button = button.fill(egui::Color32::from_rgb(232, 220, 205));
    }

    ui.add(button).on_hover_text(tooltip).clicked()
}

fn configure_local_editor_fonts(ctx: &egui::Context) {
    let regular_font = include_bytes!("../fonts/LiberationSans-Regular.ttf").to_vec();
    let bold_font = include_bytes!("../fonts/LiberationSans-Bold.ttf").to_vec();

    let mut fonts = egui::FontDefinitions::default();
    fonts.font_data.insert(
        EDITOR_REGULAR_FONT_NAME.into(),
        Arc::new(egui::FontData::from_owned(regular_font)),
    );
    fonts.font_data.insert(
        "editor_bold".into(),
        Arc::new(egui::FontData::from_owned(bold_font)),
    );
    fonts.families.insert(
        egui::FontFamily::Name(EDITOR_REGULAR_FONT_NAME.into()),
        vec![EDITOR_REGULAR_FONT_NAME.into()],
    );
    fonts.families.insert(
        egui::FontFamily::Name("editor_bold".into()),
        vec!["editor_bold".into()],
    );
    ctx.set_fonts(fonts);
}

fn main() -> eframe::Result<()> {
    let native_options = NativeOptions {
        viewport: egui::ViewportBuilder::default().with_inner_size([1280.0, 780.0]),
        ..Default::default()
    };

    eframe::run_native(
        APP_TITLE,
        native_options,
        Box::new(|cc| Ok(Box::new(LibeRustOfficeSlidesApp::new(cc)))),
    )
}
