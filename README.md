# VBA-PdfTXT

A pure VBA PDF text extraction library. No external dependencies, no DLLs, no COM automation.

---

## Why this exists

Every VBA approach to reading PDFs either requires a third-party library, shells out to an external process, or relies on Adobe Acrobat being installed. This module reads the file directly, decompresses the content streams with a from-scratch DEFLATE implementation, parses the text operators, and returns the text within the VBA runtime.

---

## What it handles

* **FlateDecode (DEFLATE) compressed streams**: full RFC 1951 implementation in pure VBA, including fixed and dynamic Huffman trees
* **LZWDecode streams**: full MSB-first 9–12 bit LZW decoder with both EarlyChange modes (PDF default `EarlyChange=1` and TIFF-compatible `EarlyChange=0`)
* **ASCII85Decode, ASCIIHexDecode, RunLengthDecode**: all legacy PDF filter types supported
* **Filter chains**: multiple filters applied in sequence (e.g. ASCIIHex + LZW)
* **2-byte CID-encoded fonts**: the encoding Word, LibreOffice, and most modern PDF generators use; maps glyph IDs back to Unicode via the embedded ToUnicode CMap
* **Literal 1-byte encoded fonts**: standard Latin PDFs with `(text) Tj` operators
* **Object streams (`/ObjStm`)**: PDF 1.5+ compressed object bundles are unpacked to extract embedded ToUnicode CMaps that would otherwise be missed
* **Multi-column layout reconstruction**: tracks the text matrix (`Tm`, `Td`, `TD`, `T*`) to recover X/Y positions, sorts runs spatially, and joins columns on the same line with a tab character
* **Ligatures**: `fl`, `fi`, `ffi`, and other multi-codepoint CMap destinations are emitted as their full character sequences
* **Both coordinate systems**: auto-detects whether Y increases upward (standard PDF) or downward (Word/LibreOffice exports with a flipped CTM) and sorts accordingly
* **Multiple content streams per page**: all streams are collected before sorting, so the final output reflects visual reading order across the whole document
* **Escaped characters in strings**: `\n`, `\r`, `\t`, `\\`, `\(`, `\)`, and octal escapes (`\141`) are all handled

---

## What it does not handle (and why)

| Scenario | Reason |
|---|---|
| **Encrypted PDFs** (password-protected) | Stream data is ciphertext; decryption requires AES/RC4: out of scope, use OCR fallback |
| **Type 3 fonts without `/ToUnicode`** | Glyphs named `/a0`, `/a1`… have no standardised Unicode mapping; the character identity is only in the drawing procedures (Bézier paths / bitmaps): fundamentally undecipherable without the font author's intent. These runs produce empty output (not garbage). |
| **Image-only PDFs** (scanned documents) | No text operators exist; OCR required |
| **`bfrange` array destinations** | Rare CMap variant; silently produces missing characters, no crash |

---

## Installation

1. In the VBA editor, go to **File → Import File** and select `PdfTXT.bas`
2. That's it. No references to set, no extra modules needed.

---

## Usage

```vb
Dim sText As String
sText = PDF_ExtractText("C:\path\to\file.pdf")

If Len(sText) = 0 Then
    ' Encrypted, image-only, or some issue reading it
Else
    Debug.Print sText
End If
```

The returned string uses:

* **Line feed** (`Chr(10)`) between lines
* **Tab** (`Chr(9)`) between items on the same line (e.g. a label and its value in a two-column layout)

---

## Diagnostics

### PDF_DiagnoseStreams

Lists every stream in the file with its position, length, content-stream classification, and the first 100 characters of its dictionary header. Useful for spotting encrypted, image-only, or unusual PDFs at a glance.

```vb
Debug.Print PDF_DiagnoseStreams("C:\path\to\file.pdf")
```

### PDF_DiagnoseVerbose

Runs the full extraction pipeline step by step and reports what happens at each stage. Use this when `PDF_ExtractText` returns empty unexpectedly.

```vb
Debug.Print PDF_DiagnoseVerbose("C:\path\to\file.pdf")
```

Output stages:

| Stage | What it reports |
|-------|----------------|
| `[S1]` | File size and `%PDF` magic byte check |
| `[S2]` | Latin-1 string conversion length |
| `[S3]` | Encryption check result |
| `[S4]` | CMap data extracted from Object Streams |
| `[S5]` | Per-stream: content flag, compressed/decompressed sizes, predictor, first 80 decompressed chars, text run byte count |
| `[S6]` | Totals: streams scanned, content streams found, streams with text |
| `[S7]` | Sorted output length |
| `[S8]` | Final cleaned output length and first 200 characters |

If an exception occurs at any stage it is reported as `[EXCEPTION] Err=N: description` and the output up to that point is still returned, so you can see exactly where the pipeline stopped.

---

## Architecture

The module is a single `.bas` file with 26 functions. The public surface is three; everything else is internal.

```
PDF_ExtractText          <- main entry point
PDF_DiagnoseStreams      <- stream-listing diagnostic
PDF_DiagnoseVerbose      <- step-by-step pipeline diagnostic

PDF_ReadFileBytes         reads raw file bytes into a Byte array
PDF_BytesToLatin1         converts Byte array to a 1:1 character string
PDF_ProcessAllStreams     two-pass: (0) ObjStm CMaps, (1) regular CMaps, (2) extract text
PDF_ExtractObjStmCMaps    unpacks /ObjStm bundles to find embedded ToUnicode CMaps
PDF_IsContentStream       rejects font files, images, ICC profiles, XRef streams etc.
PDF_ParseDecodeParms      reads /Predictor, /Columns, /EarlyChange from stream dicts
PDF_ApplyPredictor        reverses PNG (10-15) and TIFF (2) predictor encoding
PDF_ResolveFilterRef      resolves indirect /Filter and /DecodeParms references
PDF_DecompressDeflate     entry point for DEFLATE decompression
VBA_Inflate               RFC 1951 DEFLATE decompressor
INF_ReadBits              bit-level reader for Inflate
INF_BuildTable            builds Huffman decode table
INF_DecodeHuff            decodes one symbol from a Huffman stream
PDF_DecodeASCII85         decodes ASCII85-encoded streams
PDF_DecompressLZW         MSB-first 9-12 bit LZW decoder, both EarlyChange modes
PDF_DecodeASCIIHex        decodes ASCIIHex-encoded streams
PDF_DecodeRunLength       decodes PackBits / RunLength encoded streams
PDF_ExtractTextOps        parses PDF content stream operators, tracks text position
PDF_HexDecode             decodes <hex> strings, preserving null bytes for CID pairing
PDF_CleanText             strips control characters from final output
PDF_ParseCMap             parses beginbfchar / beginbfrange CMap sections
PDF_ApplyCMap             maps 2-byte CID pairs to Unicode using a parsed CMap
PDF_SortAndJoin           sorts positioned text runs into visual reading order
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

## Known edge cases

**Two-column layouts**: labels and values on the same visual line are joined with a tab. Split on `Chr(9)` to get them as separate fields.

**Line spacing tolerance**: two text runs are treated as the same line if their Y coordinates are within 8 PDF points. This handles most layouts comfortably. If you have a document with very tight superscripts or very wide line spacing that merges or splits incorrectly, adjust the `Y_TOL` constant in `PDF_SortAndJoin` (stored x100, so the default `800` = 8.0 points).

**Windows code page**: `PDF_ApplyCMap` uses `Asc()` to recover byte values from VBA's internal string representation. This correctly round-trips all byte values on systems using a single-byte ANSI code page (CP1252, CP1250, etc.). Behaviour on DBCS code page systems (Japanese, Chinese, Korean) has not been tested.

**Object streams**: ToUnicode CMaps embedded inside `/ObjStm` bundles (common in PDFs generated by Chrome, modern Word, and LibreOffice) are extracted in a dedicated pre-pass and merged with any CMaps found in regular streams before text extraction begins.

---

## License

MIT License — see [LICENSE](LICENSE) for details.

---

## Credits

Copyright © 2026, [rafael-yml](https://rafael-yml.lovable.app/)