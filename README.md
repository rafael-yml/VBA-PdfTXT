# VBA-PdfTXT

A pure VBA PDF text extraction library. No external dependencies, no DLLs, no COM automation.

---

## Why this exists

Every VBA approach to reading PDFs either requires a third-party library, shells out to an external process, or relies on Adobe Acrobat being installed. This module reads the file directly, decompresses the content streams with a from-scratch DEFLATE implementation, parses the text operators, and returns the text within the VBA runtime.

---

## What it handles

- **FlateDecode (DEFLATE) compressed streams** full RFC 1951 implementation in pure VBA, including fixed and dynamic Huffman trees
- **2-byte CID-encoded fonts** the encoding Word, LibreOffice, and most modern PDF generators use; maps glyph IDs back to Unicode via the embedded ToUnicode CMap
- **Literal 1-byte encoded fonts** standard Latin PDFs with `(text) Tj` operators
- **Multi-column layout reconstruction** tracks the text matrix (`Tm`, `Td`, `TD`, `T*`) to recover X/Y positions, sorts runs spatially, and joins columns on the same line with a tab character
- **Ligatures** `fl`, `fi`, `ffi`, and other multi-codepoint CMap destinations are emitted as their full character sequences
- **Both coordinate systems** auto-detects whether Y increases upward (standard PDF) or downward (Word/LibreOffice exports with a flipped CTM) and sorts accordingly
- **Multiple content streams per page** all streams are collected before sorting, so the final output reflects visual reading order across the whole document

---

## What it does not handle

- **Encrypted PDFs** files with a `/Encrypt` dictionary are not supported; the function returns an empty string
- **Scanned / image-only PDFs** there is no OCR; if a PDF contains no text operators, nothing is extracted
- **Type 1 / TrueType fonts without a ToUnicode CMap** some very old or highly customised PDFs omit the CMap; characters may come out as raw glyph indices
- **Object streams (`/ObjStm`)** cross-reference streams used in some PDF 1.5+ files are not parsed; content inside them will be missed
- **Right-to-left text** Arabic, Hebrew etc. will extract but word order may be reversed

---

## Installation

1. In the VBA editor, go to **File → Import File** and select `VBA-PdfTXT.bas`
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
- **Line feed** (`Chr(10)`) between lines
- **Tab** (`Chr(9)`) between items on the same line (e.g. a label and its value in a two-column layout)

---

## Diagnostics

If a PDF returns empty text unexpectedly, call the diagnostic function to inspect its streams:

```vb
Debug.Print PDF_DiagnoseStreams("C:\path\to\file.pdf")
```

This prints each stream's position, length, whether it was identified as a content stream, and the first 100 characters of its dictionary header for spotting encrypted, image-only, or ObjStm-packed files.

---

## Architecture

The module is a single `.bas` file with 17 functions. The public surface is just two; everything else is internal.

```
PDF_ExtractText          ← public entry point
PDF_DiagnoseStreams      ← public diagnostic helper

PDF_ReadFileBytes         reads raw file bytes into a Byte array
PDF_BytesToLatin1         converts Byte array to a 1:1 character string
PDF_ProcessAllStreams     two-pass: (1) find CMap, (2) extract positioned runs
PDF_IsContentStream       rejects font files, images, ICC profiles etc.
PDF_DecompressDeflate     entry point for DEFLATE decompression
VBA_Inflate               RFC 1951 DEFLATE decompressor
INF_ReadBits              bit-level reader for Inflate
INF_BuildTable            builds Huffman decode table
INF_DecodeHuff            decodes one symbol from a Huffman stream
PDF_ExtractTextOps        parses PDF content stream operators, tracks text position
PDF_HexDecode             decodes <hex> strings, preserving null bytes for CID pairing
PDF_CleanText             strips control characters from final output
PDF_ParseCMap             parses beginbfchar / beginbfrange CMap sections
PDF_ApplyCMap             maps 2-byte CID pairs to Unicode using a parsed CMap
PDF_SortAndJoin           sorts positioned text runs into visual reading order
```

---

## Performance

The bottleneck is the pure-VBA DEFLATE decompressor. On a typical modern machine:

| PDF type | Pages | Approx. time |
|---|---|---|
| Simple text, no compression | any | < 100 ms |
| Compressed, standard encoding | 1–5 | ~200–500 ms |
| Compressed, CID encoding (Word export) | 1–5 | ~300–700 ms |

For bulk processing, consider calling `PDF_ExtractText` in a loop with `DoEvents` between files to keep the host application responsive.

---

## Known edge cases

**Two-column layouts** labels and values on the same visual line are joined with a tab. Split on `Chr(9)` to get them as separate fields.

**Line spacing tolerance** two text runs are treated as the same line if their Y coordinates are within 8 PDF points. This handles most layouts comfortably. If you have a document with very tight superscripts or very wide line spacing that merges or splits incorrectly, adjust the `Y_TOL` constant in `PDF_SortAndJoin` (stored ×100, so the default 800 = 8.0 points).

**Windows code page** `PDF_ApplyCMap` uses `Asc()` to recover byte values from VBA's internal string representation. This correctly round-trips all byte values on systems using a single-byte ANSI code page (CP1252, CP1250, etc.). Behaviour on DBCS code page systems (Japanese, Chinese, Korean) has not been tested.

---

Copyright © 2026, [rafael-yml](https://rafael-yml.lovable.app/)

---

## License

MIT License — see [LICENSE](LICENSE) for details.
