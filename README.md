# VBA-PdfTXT

Pure VBA PDF text extractor. Drops into any Excel, Word, or Access VBA project as a single `.bas` file.

---

## What it handles

### Compression filters (complete PDF filter set for text streams)
| Filter | Abbreviation | Notes |
|---|---|---|
| FlateDecode | `/Fl` | zlib/Deflate: used by Word, LibreOffice, Chrome, LaTeX, iText |
| ASCII85Decode | `/A85` | Used by ReportLab (`pageCompression=1`), Ghostscript `ps2pdf`, Acrobat Distiller PostScript path |
| LZWDecode | `/LZW` | Legacy compression: Acrobat Distiller ≤ 3.x, old WordPerfect, early laser printer drivers; supports both EarlyChange=1 (PDF default) and EarlyChange=0 |
| ASCIIHexDecode | `/AHx` | Hex-encoded streams: old Distiller settings, often chained as `[/ASCIIHexDecode /LZWDecode]` |
| RunLengthDecode | `/RL` | PackBits RLE: completes the non-image filter set |

Filter chains (arrays like `[/ASCII85Decode /FlateDecode]` or `[/ASCIIHexDecode /LZWDecode]`) are decoded in the correct order at all three internal pass sites (CMap pass, content stream pass, ObjStm pass).

Image-only filters (DCTDecode/JPEG, JBIG2, JPXDecode, CCITTFax) are detected and the stream is skipped: they never contain extractable text.

### Text encodings
- **Latin-1 literal strings** `(text)`: standard single-byte encoding
- **Hex strings with ToUnicode CMap** `<hexdata>`: full CMap lookup including `beginbfchar`, `beginbfrange`, and multi-byte ranges
- **2-byte CID mode** (Identity-H/V fonts used by InDesign, modern Word): automatic detection via even-length hex strings
- **UTF-16 BOM in hex strings**: `FEFF...` prefix detected and decoded as UTF-16BE
- **Octal escapes** in literal strings: `\101` → `A`
- **Standard escape sequences**: `\n \r \t \\ \( \)`

### Content stream operators
| Operator | Meaning |
|---|---|
| `Tj` | Show string |
| `TJ` | Show array of strings (with kerning values: numbers ignored, text concatenated) |
| `'` | Move to next line and show string |
| `"` | Set word/char spacing, move to next line, show string |
| `Tm` | Set text matrix (position) |
| `Td` / `TD` | Move text position |
| `T*` | Move to next line |
| `TL` | Set leading |
| `BT` / `ET` | Begin/end text block |

### Structural features
- **Multi-stream pages**: all content streams per page collected and spatially sorted
- **Reading-order reconstruction**: Y-then-X sort approximates top-to-bottom, left-to-right reading order
- **Object streams** (PDF 1.5+ ObjStm): compressed font/CMap objects extracted from cross-reference streams
- **Indirect /Filter references**: `/Filter 6 0 R` resolved one level deep (SAP NetWeaver, some enterprise generators)
- **PNG predictor** (values 10–15) and **TIFF predictor** (value 2): reversed after FlateDecode
- **Linearised PDFs**: handled by linear stream scan (no XREF dependency)
- **Form XObjects** (`/Subtype /Form`): picked up automatically as content streams (headers, footers, stamps, repeated elements)
- **Tagged PDF / Marked content** (`BDC`/`EMC`): tags silently ignored, text operators extracted normally
- **Encrypted PDFs**: detected via `/Encrypt` dict; returns empty string immediately

### Output cleaning
`PDF_CleanText` is applied to the final joined output:
- Strips null bytes (codes 0–31, 127, 128–159) that appear when hex strings have no CMap mapping
- Collapses consecutive double-spaces from explicit character spacing
- Normalises line breaks

---

## What it does not handle (and why)

| Scenario | Reason |
|---|---|
| **Encrypted PDFs** (password-protected) | Stream data is ciphertext; decryption requires AES/RC4: out of scope, use OCR fallback |
| **Type 3 fonts without `/ToUnicode`** | Glyphs named `/a0`, `/a1`… have no standardised Unicode mapping; the character identity is only in the drawing procedures (Bézier paths / bitmaps): fundamentally undecipherable without the font author's intent. These runs produce empty output (not garbage). |
| **Image-only PDFs** (scanned documents) | No text operators exist; OCR required |
| **`bfrange` array destinations** | Rare CMap variant; silently produces missing characters, no crash |

---

## Architecture: 25 functions

### Public API (2)
| Function | Signature | Description |
|---|---|---|
| `PDF_ExtractText` | `(sFilePath As String) As String` | Main entry point. Returns extracted text or `""` on encrypted/invalid files. |
| `PDF_DiagnoseStreams` | `(sFilePath As String) As String` | Diagnostic: lists every stream found, its filter chain, byte counts, and whether it was parsed as a content stream. |

### Internal pipeline (23)
| Function | Role |
|---|---|
| `PDF_ProcessAllStreams` | Orchestrates the three-pass pipeline; returns `PDF_CleanText(PDF_SortAndJoin(...))` |
| `PDF_ParseCMap` | Parses CMap object into a lookup string |
| `PDF_ApplyCMap` | Maps raw byte string through CMap lookup |
| `PDF_ExtractObjStmCMaps` | Extracts CMap data from PDF 1.5+ ObjStm streams |
| `PDF_ExtractTextOps` | State-machine content stream parser: Tj/TJ/Tm/Td/TD/T*/TL/'/"/BT/ET |
| `PDF_HexDecode` | Decodes `<hex>` strings including UTF-16 BOM handling |
| `PDF_SortAndJoin` | Spatially sorts text runs (Y desc, X asc) and joins to string |
| `PDF_CleanText` | Strips control characters, collapses double-spaces |
| `PDF_ResolveFilterRef` | Resolves indirect `/Filter N M R` references one level |
| `PDF_ParseDecodeParms` | Extracts `/Predictor`, `/Columns`, `/EarlyChange` from stream dict |
| `PDF_ApplyPredictor` | Reverses PNG (10–15) and TIFF (2) predictor encoding |
| `PDF_DecodeASCII85` | ASCII85 (Base-85) decoder: groups of 5 chars → 4 bytes, `z` shorthand, partial final group |
| `PDF_DecodeASCIIHex` | ASCIIHexDecode: hex digit pairs → bytes, whitespace-tolerant, `>` terminator |
| `PDF_DecompressLZW` | LZW decoder: variable-width 9–12 bit codes, MSB-first, KwKwK edge case, EarlyChange=0/1 |
| `PDF_DecodeRunLength` | PackBits RLE decoder: literal runs, repeat runs, 128=EOD |
| `PDF_DecompressDeflate` | Calls `VBA_Inflate` with zlib header skip |
| `VBA_Inflate` | Pure VBA DEFLATE implementation (fixed + dynamic Huffman, LZ77 back-references) |
| `INF_ReadBits` | LSB-first bit reader used by `VBA_Inflate` |
| `INF_BuildTree` | Builds canonical Huffman decode tree |
| `INF_DecodeSymbol` | Decodes one symbol from the bit stream |
| `HuffNode` / `HuffMakeTree` | Huffman tree node type and builder |

---

## Usage

```vba
Dim sText As String
sText = PDF_ExtractText("C:\reports\invoice.pdf")
If Len(sText) = 0 Then
    ' Encrypted, image-only, corrupt, or could be an extreme unhandled edge-case as well
Else
    Debug.Print sText
End If
```

Diagnostics:
```vba
Debug.Print PDF_DiagnoseStreams("C:\reports\invoice.pdf")
```

---

## Compatibility notes

### PDF generators confirmed to work
| Generator | Filter used | Notes |
|---|---|---|
| Microsoft Word (all versions) | FlateDecode | |
| LibreOffice Writer | FlateDecode | |
| Google Chrome / Chromium print-to-PDF | FlateDecode | |
| LaTeX (pdflatex, xelatex, lualatex) | FlateDecode | CMap extracted for ligatures |
| Adobe InDesign | FlateDecode + CID | 2-byte CID mode |
| iText / iTextSharp | FlateDecode | |
| Apache PDFBox | FlateDecode | |
| ReportLab (Python) | ASCII85Decode + FlateDecode | `pageCompression=1` default |
| Ghostscript `ps2pdf` | ASCII85Decode + FlateDecode | |
| Acrobat Distiller (modern) | FlateDecode or ASCII85+Flate | |
| Acrobat Distiller ≤ 3.x (legacy) | LZWDecode or ASCIIHex+LZW | PDF 1.1–1.2 era |
| Old WordPerfect PDF export | LZWDecode | |
| Uncompressed hand-crafted PDFs | None (raw) | |


## Performance (approximate, Core i5, 32-bit VBA host)

| Content type | Pages | Typical time |
|---|---|---|
| FlateDecode (Word, LibreOffice) | 1–10 | 50–200 ms |
| FlateDecode (Word, LibreOffice) | 50–100 | 1–3 s |
| ASCII85 + FlateDecode (ReportLab) | 1–5 | 300–700 ms |
| LZWDecode (legacy Distiller) | 1–10 | 100–400 ms |
| ASCIIHex + LZW chain | 1–5 | 150–500 ms |
| RunLengthDecode | 1–5 | 50–150 ms |
| Uncompressed | any | < 50 ms |

---

## Installation

1. Download `VBA-PdfTXT.bas`
2. In your VBA project: **File → Import File** → select `VBA-PdfTXT.bas`
3. Call `PDF_ExtractText(filePath)` from any module

No references to set. No DLLs to register. Works in Excel, Word, Access, and any other Office VBA host.

---

## License

MIT License — see [LICENSE](LICENSE) for details.

---

## Credits

Copyright © 2026, [rafael-yml](https://rafael-yml.lovable.app/)
