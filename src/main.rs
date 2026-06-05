use std::env;
use std::error::Error;
use std::fs;
use std::path::{Path, PathBuf};

use lopdf::{Dictionary, Document, Object};
use rust_xlsxwriter::Workbook;

type Result<T> = std::result::Result<T, Box<dyn Error>>;

#[derive(Debug, Clone)]
struct CommentRow {
    author: String,
    comment_text: String,
    filename: String,
}

fn main() {
    if let Err(e) = run() {
        eprintln!("Error: {e}");
        std::process::exit(1);
    }
}

fn run() -> Result<()> {
    let args: Vec<String> = env::args().collect();

    if args.len() != 3 {
        print_usage(&args[0]);
        return Err("invalid arguments".into());
    }

    let input = PathBuf::from(&args[1]);
    let output = PathBuf::from(&args[2]);

    if output.extension().and_then(|s| s.to_str()).unwrap_or("").to_ascii_lowercase() != "xlsx" {
        return Err("output file must have .xlsx extension".into());
    }

    let pdf_files = collect_pdf_files(&input)?;
    if pdf_files.is_empty() {
        return Err("no PDF files found".into());
    }

    let mut rows = Vec::new();
    for pdf in &pdf_files {
        let filename = pdf
            .file_name()
            .and_then(|s| s.to_str())
            .unwrap_or("unknown.pdf")
            .to_string();

        let file_rows = extract_comments_from_pdf(pdf, &filename)?;
        rows.extend(file_rows);
    }

    write_xlsx(&output, &rows)?;
    println!(
        "Done. Processed {} PDF file(s), wrote {} comment row(s) to {}",
        pdf_files.len(),
        rows.len(),
        output.display()
    );

    Ok(())
}

fn print_usage(bin_name: &str) {
    eprintln!("Usage:");
    eprintln!("  {bin_name} <input-pdf-or-folder> <output.xlsx>");
    eprintln!();
    eprintln!("Examples:");
    eprintln!("  {bin_name} C:/docs/sample.pdf C:/out/comments.xlsx");
    eprintln!("  {bin_name} C:/docs C:/out/comments.xlsx");
}

fn collect_pdf_files(input: &Path) -> Result<Vec<PathBuf>> {
    if input.is_file() {
        if is_pdf_file(input) {
            return Ok(vec![input.to_path_buf()]);
        }
        return Err("input file is not a PDF".into());
    }

    if !input.is_dir() {
        return Err("input path does not exist".into());
    }

    let mut files = Vec::new();
    collect_pdf_files_recursive(input, &mut files)?;
    Ok(files)
}

fn collect_pdf_files_recursive(dir: &Path, out: &mut Vec<PathBuf>) -> Result<()> {
    for entry in fs::read_dir(dir)? {
        let entry = entry?;
        let path = entry.path();

        if path.is_dir() {
            collect_pdf_files_recursive(&path, out)?;
        } else if is_pdf_file(&path) {
            out.push(path);
        }
    }
    Ok(())
}

fn is_pdf_file(path: &Path) -> bool {
    path.extension()
        .and_then(|s| s.to_str())
        .map(|s| s.eq_ignore_ascii_case("pdf"))
        .unwrap_or(false)
}

fn extract_comments_from_pdf(path: &Path, filename: &str) -> Result<Vec<CommentRow>> {
    let doc = Document::load(path)?;
    let mut rows = Vec::new();

    for (_, page_id) in doc.get_pages() {
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
            let annot_obj = match deref_object(&doc, annot_ref) {
                Some(obj) => obj,
                None => continue,
            };

            let annot_dict = match annot_obj.as_dict() {
                Ok(d) => d,
                Err(_) => continue,
            };

            let contents = get_string_from_dict(&doc, annot_dict, b"Contents").unwrap_or_default();
            if contents.trim().is_empty() {
                continue;
            }

            let author = get_string_from_dict(&doc, annot_dict, b"T").unwrap_or_default();

            rows.push(CommentRow {
                author,
                comment_text: contents,
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

    String::from_utf8_lossy(bytes).to_string()
}

fn write_xlsx(output: &Path, rows: &[CommentRow]) -> Result<()> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    worksheet.write_string(0, 0, "author")?;
    worksheet.write_string(0, 1, "commenttext")?;
    worksheet.write_string(0, 2, "filename")?;

    for (i, row) in rows.iter().enumerate() {
        let r = (i + 1) as u32;
        worksheet.write_string(r, 0, &row.author)?;
        worksheet.write_string(r, 1, &row.comment_text)?;
        worksheet.write_string(r, 2, &row.filename)?;
    }

    workbook.save(output)?;
    Ok(())
}
