# PDF Comment Extractor (Rust)

This tool extracts PDF annotation comments and writes them to an `.xlsx` file with these columns:
- author
- commenttext
- filename

## What it does
- Input can be a single PDF file or a folder (scans subfolders recursively)
- Reads PDF annotations from `/Annots`
- Extracts `/T` as `author` and `/Contents` as `commenttext`
- Writes all rows into one Excel file

## Build prerequisites (Windows)
1. Install Rust toolchain from https://rustup.rs/
2. Open a new terminal after installation

## Build (small release exe)
```powershell
cargo build --release
```

Output binary:
- `target\release\pdfextract.exe`

## Run examples
```powershell
# Single PDF
.\target\release\pdfextract.exe C:\docs\sample.pdf C:\out\comments.xlsx

# Folder (recursive)
.\target\release\pdfextract.exe C:\docs C:\out\comments.xlsx
```

## Notes about exe size
The release profile is tuned for size (`opt-level="z"`, LTO, single codegen unit, strip symbols, panic abort). Actual size depends on toolchain and target, but Rust is typically much smaller than Python-based single-file executables for this use case.
