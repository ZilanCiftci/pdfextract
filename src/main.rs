#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use std::error::Error;
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::{
    atomic::{AtomicUsize, Ordering},
    Arc,
};
use std::time::Duration;

use iced::alignment::{Horizontal, Vertical};
use iced::widget::{
    button, column, container, progress_bar, row, scrollable, text, text_input, Space,
};
use iced::{time, Application, Command, Element, Length, Settings, Size, Subscription, Theme};
use lopdf::{Dictionary, Document, Object};
use rfd::FileDialog;
use rust_xlsxwriter::Workbook;

type BoxError = Box<dyn Error>;

// ── data ─────────────────────────────────────────────────────────────────────

#[derive(Debug, Clone)]
struct CommentRow {
    author: String,
    comment_text: String,
    page: u32,
    filename: String,
}

// ── app state ─────────────────────────────────────────────────────────────────

#[derive(Default)]
struct App {
    pdf_files: Vec<PathBuf>,
    output_path: String,
    status: String,
    running: bool,
    progress: f32,
    progress_done: usize,
    progress_total: usize,
    progress_counter: Option<Arc<AtomicUsize>>,
}

#[derive(Debug, Clone)]
enum Message {
    AddFiles,
    AddFolder,
    ClearAll,
    RemoveFile(usize),
    ChooseOutput,
    OutputChanged(String),
    Run,
    ExtractionFinished(Result<String, String>),
    Tick,
}

fn main() -> iced::Result {
    let icon = load_window_icon();
    App::run(Settings {
        window: iced::window::Settings {
            size: Size::new(680.0, 720.0),
            min_size: Some(Size::new(540.0, 520.0)),
            icon,
            ..Default::default()
        },
        ..Default::default()
    })
}

impl Application for App {
    type Executor = iced::executor::Default;
    type Message = Message;
    type Theme = Theme;
    type Flags = ();

    fn new(_flags: Self::Flags) -> (Self, Command<Self::Message>) {
        (Self::default(), Command::none())
    }

    fn title(&self) -> String {
        "PDF Kommentarsextraktor".to_string()
    }

    fn theme(&self) -> Self::Theme {
        Theme::Light
    }

    fn update(&mut self, message: Self::Message) -> Command<Self::Message> {
        match message {
            Message::AddFiles => {
                if let Some(files) = FileDialog::new()
                    .add_filter("PDF-filer", &["pdf", "PDF"])
                    .pick_files()
                {
                    for f in files {
                        if !self.pdf_files.contains(&f) {
                            self.pdf_files.push(f);
                        }
                    }
                }
                Command::none()
            }
            Message::AddFolder => {
                if let Some(folder) = FileDialog::new().pick_folder() {
                    let mut found = Vec::new();
                    collect_pdf_files_recursive(&folder, &mut found).ok();
                    found.sort();
                    for f in found {
                        if !self.pdf_files.contains(&f) {
                            self.pdf_files.push(f);
                        }
                    }
                }
                Command::none()
            }
            Message::ClearAll => {
                self.pdf_files.clear();
                Command::none()
            }
            Message::RemoveFile(idx) => {
                if idx < self.pdf_files.len() {
                    self.pdf_files.remove(idx);
                }
                Command::none()
            }
            Message::ChooseOutput => {
                if let Some(path) = FileDialog::new()
                    .add_filter("Excel-arbetsbok", &["xlsx"])
                    .set_file_name("kommentarer.xlsx")
                    .save_file()
                {
                    let mut p = path.to_string_lossy().into_owned();
                    if !p.to_ascii_lowercase().ends_with(".xlsx") {
                        p.push_str(".xlsx");
                    }
                    self.output_path = p;
                }
                Command::none()
            }
            Message::OutputChanged(value) => {
                self.output_path = value;
                Command::none()
            }
            Message::Run => {
                let can_run = !self.running
                    && !self.pdf_files.is_empty()
                    && !self.output_path.is_empty();
                if !can_run {
                    return Command::none();
                }

                let files = self.pdf_files.clone();
                let output = PathBuf::from(&self.output_path);
                let counter = Arc::new(AtomicUsize::new(0));
                self.running = true;
                self.status = "Bearbetar…".to_string();
                self.progress = 0.0;
                self.progress_done = 0;
                self.progress_total = files.len();
                self.progress_counter = Some(counter.clone());

                Command::perform(
                    async move { run_extraction_with_progress(&files, &output, counter) },
                    Message::ExtractionFinished,
                )
            }
            Message::ExtractionFinished(result) => {
                self.running = false;
                self.progress_counter = None;
                self.progress = if self.progress_total == 0 { 0.0 } else { 1.0 };
                match result {
                    Ok(msg) => self.status = msg,
                    Err(e) => self.status = format!("Fel: {e}"),
                }
                Command::none()
            }
            Message::Tick => {
                if let Some(counter) = &self.progress_counter {
                    let done = counter.load(Ordering::Relaxed);
                    self.progress_done = done.min(self.progress_total);
                    if self.progress_total > 0 {
                        self.progress = self.progress_done as f32 / self.progress_total as f32;
                    }
                }
                Command::none()
            }
        }
    }

    fn subscription(&self) -> Subscription<Self::Message> {
        if self.running {
            time::every(Duration::from_millis(200)).map(|_| Message::Tick)
        } else {
            Subscription::none()
        }
    }

    fn view(&self) -> Element<'_, Self::Message> {
        let header = container(
            row![
                Space::with_width(Length::Fixed(6.0)),
                text("PDF Kommentarsextraktor").size(22),
            ]
            .align_items(iced::Alignment::Center),
        )
        .padding(14)
        .width(Length::Fill);

        let mut list_col = column![].spacing(6);
        if self.pdf_files.is_empty() {
            list_col = list_col.push(text(
                "Inga filer ännu – lägg till PDF-filer eller en mapp ovan.",
            ));
        } else {
            for (i, path) in self.pdf_files.iter().enumerate() {
                let name = path
                    .file_name()
                    .and_then(|s| s.to_str())
                    .unwrap_or("?")
                    .to_string();

                let mut remove = button(text("✕")).padding([2, 8]);
                if !self.running {
                    remove = remove.on_press(Message::RemoveFile(i));
                }

                list_col = list_col.push(
                    row![remove, text(name)]
                        .spacing(8)
                        .align_items(iced::Alignment::Center),
                );
            }
        }

        let list = scrollable(list_col)
            .height(Length::Fixed(150.0))
            .width(Length::Fill);

        let input_card = container(
            column![
                text("INDATA – PDF-FILER").size(12),
                Space::with_height(Length::Fixed(6.0)),
                row![
                    button(text("Lägg till PDF-filer"))
                        .style(black_button_style())
                        .on_press(Message::AddFiles),
                    button(text("Lägg till mapp"))
                        .style(black_button_style())
                        .on_press(Message::AddFolder),
                    Space::with_width(Length::Fill),
                    button(text("Rensa allt"))
                        .style(black_button_style())
                        .on_press(Message::ClearAll),
                ]
                .spacing(8),
                Space::with_height(Length::Fixed(6.0)),
                text(format!("{} fil(er) valda", self.pdf_files.len())),
                Space::with_height(Length::Fixed(6.0)),
                container(list)
                    .padding(8),
            ]
            .spacing(4),
        )
        .padding(12)
        .width(Length::Fill);

        let output_exists = !self.output_path.is_empty() && Path::new(&self.output_path).exists();

        let output_card = container(
            column![
                text("UTDATA – EXCEL-FIL").size(12),
                Space::with_height(Length::Fixed(6.0)),
                row![
                    button(text("Välj fil…"))
                        .style(black_button_style())
                        .on_press(Message::ChooseOutput),
                    text_input("Välj eller skriv en .xlsx-sökväg…", &self.output_path)
                        .on_input(Message::OutputChanged)
                        .padding(8)
                        .size(14)
                        .width(Length::Fill),
                ]
                .spacing(8)
                .align_items(iced::Alignment::Center),
                if output_exists {
                    text("Filen finns – nya kommentarer läggs till i slutet.").size(12)
                } else {
                    text("").size(12)
                },
            ]
            .spacing(4),
        )
        .padding(12)
        .width(Length::Fill);

        let can_run = !self.running && !self.pdf_files.is_empty() && !self.output_path.is_empty();
        let mut run_btn = button(
            text("Extrahera kommentarer")
                .size(16)
                .horizontal_alignment(Horizontal::Center)
                .vertical_alignment(Vertical::Center),
        )
        .padding([10, 18])
        .width(Length::Fixed(260.0));
        if can_run {
            run_btn = run_btn
                .style(black_button_style())
                .on_press(Message::Run);
        } else {
            run_btn = run_btn.style(black_button_style());
        }

        let content = column![
            input_card,
            Space::with_height(Length::Fixed(12.0)),
            output_card,
            Space::with_height(Length::Fixed(14.0)),
            container(run_btn)
                .width(Length::Fill)
                .align_x(Horizontal::Center),
            if self.running {
                container(
                    column![
                        progress_bar(0.0..=1.0, self.progress)
                            .width(Length::Fill),
                        text(format!(
                            "{} / {} ({}%)",
                            self.progress_done,
                            self.progress_total,
                            (self.progress * 100.0).round() as u32
                        ))
                        .size(12),
                    ]
                    .spacing(6),
                )
                .padding([0, 12])
            } else {
                container(text(""))
            },
            if self.running {
                container(text("Bearbetar…"))
                    .padding([4, 0])
                    .width(Length::Fill)
                    .align_x(Horizontal::Center)
            } else {
                container(text(""))
            },
            Space::with_height(Length::Fixed(6.0)),
            status_banner(&self.status),
        ]
        .spacing(8)
        .padding(16)
        .width(Length::Fill);

        container(column![header, content])
            .width(Length::Fill)
            .height(Length::Fill)
            .into()
    }
}

fn load_window_icon() -> Option<iced::window::Icon> {
    let bytes = include_bytes!("../assets/icon.png");
    let img = image::load_from_memory(bytes).ok()?;
    let rgba = img.to_rgba8();
    let (w, h) = rgba.dimensions();
    iced::window::icon::from_rgba(rgba.into_raw(), w, h).ok()
}

fn black_button_style() -> iced::theme::Button {
    iced::theme::Button::Custom(Box::new(BlackButton))
}

struct BlackButton;

impl button::StyleSheet for BlackButton {
    type Style = Theme;

    fn active(&self, _style: &Self::Style) -> button::Appearance {
        button::Appearance {
            background: Some(iced::Background::Color(iced::Color::from_rgb8(0x21, 0x21, 0x21))),
            text_color: iced::Color::WHITE,
            border: iced::Border {
                color: iced::Color::from_rgb8(0x21, 0x21, 0x21),
                width: 1.0,
                radius: 6.0.into(),
            },
            shadow_offset: iced::Vector::new(0.0, 0.0),
            shadow: iced::Shadow::default(),
        }
    }

    fn hovered(&self, style: &Self::Style) -> button::Appearance {
        let mut s = self.active(style);
        s.background = Some(iced::Background::Color(iced::Color::from_rgb8(0x11, 0x11, 0x11)));
        s
    }

    fn disabled(&self, style: &Self::Style) -> button::Appearance {
        let mut s = self.active(style);
        s.background = Some(iced::Background::Color(iced::Color::from_rgb8(0xC8, 0xC8, 0xC8)));
        s.text_color = iced::Color::from_rgb8(0x66, 0x66, 0x66);
        s.border = iced::Border {
            color: iced::Color::from_rgb8(0xC8, 0xC8, 0xC8),
            width: 1.0,
            radius: 6.0.into(),
        };
        s
    }
}

fn status_banner<'a>(status: &str) -> Element<'a, Message> {
    if status.is_empty() {
        return container(text("")).into();
    }

    let is_err = status.starts_with("Fel");
    let (bg, border, fg) = if is_err {
        (
            iced::Color::from_rgb8(0xFD, 0xE7, 0xE9),
            iced::Color::from_rgb8(0xF1, 0xC2, 0xC7),
            iced::Color::from_rgb8(0xB1, 0x0E, 0x1C),
        )
    } else {
        (
            iced::Color::from_rgb8(0xDF, 0xF6, 0xDD),
            iced::Color::from_rgb8(0xBF, 0xE5, 0xBA),
            iced::Color::from_rgb8(0x0E, 0x70, 0x0E),
        )
    };

    let banner = container(text(status).style(iced::theme::Text::Color(fg)))
        .padding(8)
        .width(Length::Fill)
        .style(iced::theme::Container::Custom(Box::new(StatusStyle {
            background: bg,
            border,
        })));

    banner.into()
}

struct StatusStyle {
    background: iced::Color,
    border: iced::Color,
}

impl container::StyleSheet for StatusStyle {
    type Style = Theme;

    fn appearance(&self, _style: &Self::Style) -> container::Appearance {
        container::Appearance {
            background: Some(iced::Background::Color(self.background)),
            border: iced::Border {
                color: self.border,
                width: 1.0,
                radius: 6.0.into(),
            },
            text_color: None,
            shadow: iced::Shadow::default(),
        }
    }
}

// ── extraction (runs on background thread) ────────────────────────────────────

fn run_extraction_with_progress(
    files: &[PathBuf],
    output: &Path,
    counter: Arc<AtomicUsize>,
) -> Result<String, String> {
    run_extraction_inner(files, output, Some(counter)).map_err(|e| e.to_string())
}

fn run_extraction_inner(
    files: &[PathBuf],
    output: &Path,
    counter: Option<Arc<AtomicUsize>>,
) -> Result<String, BoxError> {
    if output
        .extension()
        .and_then(|s| s.to_str())
        .unwrap_or("")
        .to_ascii_lowercase()
        != "xlsx"
    {
        return Err("Utdatafilen måste ha filändelsen .xlsx".into());
    }

    // If the target workbook already exists, load its rows and append to them.
    let mut rows = Vec::new();
    let appended = output.exists();
    let existing_count = if appended {
        let prev = read_existing_xlsx(output)?;
        let n = prev.len();
        rows.extend(prev);
        n
    } else {
        0
    };

    let mut new_count = 0;
    for pdf in files {
        let filename = pdf
            .file_name()
            .and_then(|s| s.to_str())
            .unwrap_or("unknown.pdf")
            .to_string();
        let file_rows = extract_comments_from_pdf(pdf, &filename)?;
        new_count += file_rows.len();
        rows.extend(file_rows);

        if let Some(c) = &counter {
            c.fetch_add(1, Ordering::Relaxed);
        }
    }

    write_xlsx(output, &rows)?;

    if appended {
        Ok(format!(
            "Klart. La till {new_count} ny(a) kommentar(er) från {} fil(er) till de befintliga {existing_count} raden/raderna. Totalt {} rad(er) i {}",
            files.len(),
            rows.len(),
            output.display()
        ))
    } else {
        Ok(format!(
            "Klart. Bearbetade {} fil(er), skrev {new_count} kommentarsrad(er) till {}",
            files.len(),
            output.display()
        ))
    }
}

// ── read existing workbook (for appending) ────────────────────────────────────

fn read_existing_xlsx(path: &Path) -> Result<Vec<CommentRow>, BoxError> {
    use calamine::{open_workbook, Reader, Xlsx};

    let mut workbook: Xlsx<_> = open_workbook(path)?;
    let mut out = Vec::new();

    let sheet = match workbook.sheet_names().first().cloned() {
        Some(s) => s,
        None => return Ok(out),
    };

    let range = workbook.worksheet_range(&sheet)?;
    for (i, row) in range.rows().enumerate() {
        if i == 0 {
            continue; // header
        }
        let author = cell_to_string(row.first());
        let filename = cell_to_string(row.get(1));
        let page = cell_to_u32(row.get(2));
        let comment = cell_to_string(row.get(3));

        if author.is_empty() && comment.is_empty() && filename.is_empty() {
            continue;
        }
        out.push(CommentRow {
            author,
            comment_text: comment,
            page,
            filename,
        });
    }
    Ok(out)
}

fn cell_to_string(cell: Option<&calamine::Data>) -> String {
    use calamine::Data;
    match cell {
        Some(Data::String(s)) => s.clone(),
        Some(Data::Int(i)) => i.to_string(),
        Some(Data::Float(f)) => {
            if f.fract() == 0.0 {
                (*f as i64).to_string()
            } else {
                f.to_string()
            }
        }
        Some(Data::Bool(b)) => b.to_string(),
        _ => String::new(),
    }
}

fn cell_to_u32(cell: Option<&calamine::Data>) -> u32 {
    use calamine::Data;
    match cell {
        Some(Data::Int(i)) => (*i).max(0) as u32,
        Some(Data::Float(f)) => f.max(0.0) as u32,
        Some(Data::String(s)) => s.trim().parse().unwrap_or(0),
        _ => 0,
    }
}

// ── PDF helpers ───────────────────────────────────────────────────────────────

fn collect_pdf_files_recursive(dir: &Path, out: &mut Vec<PathBuf>) -> Result<(), BoxError> {
    collect_pdf_files_depth(dir, out, 0);
    Ok(())
}

fn collect_pdf_files_depth(dir: &Path, out: &mut Vec<PathBuf>, depth: usize) {
    // Guard against pathologically deep trees / junction loops.
    if depth > 64 {
        return;
    }

    let entries = match fs::read_dir(dir) {
        Ok(e) => e,
        Err(_) => return, // skip folders we can't read (permissions, etc.)
    };

    for entry in entries.flatten() {
        // Use the entry's own file type so we don't follow symlinks/junctions,
        // which on Windows can point back to an ancestor and cause infinite
        // recursion (stack overflow).
        let file_type = match entry.file_type() {
            Ok(t) => t,
            Err(_) => continue,
        };
        if file_type.is_symlink() {
            continue;
        }

        let path = entry.path();
        if file_type.is_dir() {
            collect_pdf_files_depth(&path, out, depth + 1);
        } else if file_type.is_file() && is_pdf_file(&path) {
            out.push(path);
        }
    }
}

fn is_pdf_file(path: &Path) -> bool {
    path.extension()
        .and_then(|s| s.to_str())
        .map(|s| s.eq_ignore_ascii_case("pdf"))
        .unwrap_or(false)
}

fn extract_comments_from_pdf(path: &Path, filename: &str) -> Result<Vec<CommentRow>, BoxError> {
    let doc = Document::load(path)?;
    let mut rows = Vec::new();

    for (page_num, page_id) in doc.get_pages() {
        let page_obj = doc.get_object(page_id)?;
        let page_dict = match page_obj.as_dict() {
            Ok(d) => d,
            Err(_) => continue,
        };

        let annots_obj = match page_dict.get(b"Annots") {
            Ok(obj) => obj,
            Err(_) => continue,
        };

        let annots_arr = match deref_object(&doc, annots_obj) {
            Some(Object::Array(arr)) => arr,
            _ => continue,
        };

        for annot_ref in annots_arr {
            let annot_obj = match deref_object(&doc, &annot_ref) {
                Some(obj) => obj,
                None => continue,
            };

            let annot_dict = match annot_obj.as_dict() {
                Ok(d) => d,
                Err(_) => continue,
            };

            let contents =
                get_string_from_dict(&doc, annot_dict, b"Contents").unwrap_or_default();
            if contents.trim().is_empty() {
                continue;
            }

            let author = get_string_from_dict(&doc, annot_dict, b"T").unwrap_or_default();
            if author.to_ascii_lowercase().contains("autocad") {
                continue;
            }
            rows.push(CommentRow {
                author,
                comment_text: contents,
                page: page_num,
                filename: filename.to_string(),
            });
        }
    }

    Ok(rows)
}

fn deref_object(doc: &Document, obj: &Object) -> Option<Object> {
    match obj {
        Object::Reference(id) => doc.get_object(*id).ok().cloned(),
        _ => Some(obj.clone()),
    }
}

fn get_string_from_dict(doc: &Document, dict: &Dictionary, key: &[u8]) -> Option<String> {
    let obj = dict.get(key).ok()?;
    object_to_string(doc, obj)
}

fn object_to_string(doc: &Document, obj: &Object) -> Option<String> {
    match obj {
        Object::String(bytes, _) => Some(decode_pdf_text(bytes)),
        Object::Name(name) => Some(String::from_utf8_lossy(name).to_string()),
        Object::Reference(id) => {
            let deref = doc.get_object(*id).ok()?;
            object_to_string(doc, deref)
        }
        _ => None,
    }
}

fn decode_pdf_text(bytes: &[u8]) -> String {
    if bytes.len() >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF {
        let mut data = Vec::new();
        for chunk in bytes[2..].chunks(2) {
            if chunk.len() == 2 {
                data.push(u16::from_be_bytes([chunk[0], chunk[1]]));
            }
        }
        return String::from_utf16_lossy(&data);
    }
    if bytes.len() >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE {
        let mut data = Vec::new();
        for chunk in bytes[2..].chunks(2) {
            if chunk.len() == 2 {
                data.push(u16::from_le_bytes([chunk[0], chunk[1]]));
            }
        }
        return String::from_utf16_lossy(&data);
    }

    // Try UTF-8; fall back to Latin-1 / PDFDocEncoding (covers å ä ö etc.)
    if let Ok(s) = std::str::from_utf8(bytes) {
        return s.to_string();
    }
    bytes.iter().map(|&b| b as char).collect()
}

// ── XLSX writer ───────────────────────────────────────────────────────────────

fn write_xlsx(output: &Path, rows: &[CommentRow]) -> Result<(), BoxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    worksheet.write_string(0, 0, "author")?;
    worksheet.write_string(0, 1, "filename")?;
    worksheet.write_string(0, 2, "page")?;
    worksheet.write_string(0, 3, "commenttext")?;

    for (i, row) in rows.iter().enumerate() {
        let r = (i + 1) as u32;
        worksheet.write_string(r, 0, &row.author)?;
        worksheet.write_string(r, 1, &row.filename)?;
        worksheet.write_number(r, 2, row.page as f64)?;
        worksheet.write_string(r, 3, &row.comment_text)?;
    }

    workbook.save(output)?;
    Ok(())
}
