# Extrahera PDF-kommentarer (Rust)

Det här verktyget läser PDF-annoteringar och exporterar dem till en `.xlsx`-fil.

## Vad programmet gör

- Välj en eller flera PDF-filer, eller en mapp (skannar undermappar rekursivt)
- Läser annoteringar via `/Annots`
- Hämtar författare från `/T` och kommentar från `/Contents`
- Exkluderar annoteringar där författaren innehåller `AutoCAD` (t.ex. "AutoCAD SHX Text")
- Skriver allt till en Excel-fil och kan lägga till i en befintlig fil

## Exportformat (kolumner)

1. `author`
2. `filename`
3. `page`
4. `commenttext`

## Bygga (Windows)

1. Installera Rust: https://rustup.rs/
2. Öppna en ny terminal

```powershell
cargo build --release
```

Körbar fil:

- `target\release\pdfextract.exe`

## Användning

Starta `.exe` och använd GUI:t för att välja PDF-filer/mapp och var Excel-filen ska sparas.
