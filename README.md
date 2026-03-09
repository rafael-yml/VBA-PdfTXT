# VBA-PdfTXT

A pure VBA PDF text extraction library. No external dependencies, no DLLs, no COM automation — the entire implementation lives in a single `.bas` file.

---

## Why this exists

Every VBA approach to reading PDFs either requires a third-party library, shells out to an external process, or relies on Adobe Acrobat being installed. This module reads the file directly, decompresses content streams with a from-scratch DEFLATE implementation, resolves font CMaps (including those buried inside PDF 1.5+ object streams), and returns plain Unicode text entirely within the VBA runtime.

---

## What it handles

* **FlateDecode (DEFLATE) compressed streams** — full RFC 1951 implementation in pure VBA, including stored blocks, fixed Huffman trees, and dynamic Huffman trees
* **ASCII85Decode streams** — pure VBA Base-85 decoder per PDF spec §7.4.3, including the `z` zero shorthand, partial final groups, and the `~>` end marker; also handles the chained `[/ASCII85Decode /FlateDecode]` filter array used by ReportLab and Ghostscript's `ps2pdf`
* **PNG and TIFF predictors** — all five PNG row filters (None, Sub, Up, Average, Paeth) and the TIFF horizontal-differencing predictor, applied automatically when `/DecodeParms` specifies them
* **Indirect `/Filter` references** — some generators (e.g. SAP NetWeaver) write `/Filter 6 0 R` instead of `/Filter /FlateDecode`; the reference is resolved before decompression
* **2-byte CID-encoded fonts** — the encoding used by Word, LibreOffice, and most modern PDF generators; glyph IDs are mapped back to Unicode via the embedded ToUnicode CMap
* **Literal 1-byte encoded fonts** — standard Latin PDFs using `(text) Tj` operators with octal escape sequences
* **ToUnicode CMaps in Object Streams (PDF 1.5+)** — Chrome, Word, and LibreOffice often pack font objects into compressed `/ObjStm` bundles; CMaps inside them are now extracted correctly, so CID-encoded PDFs from these generators decode properly
* **Multi-column layout reconstruction** — tracks the text matrix (`Tm`, `Td`, `TD`, `T*`, `'`) to recover X/Y positions, sorts runs spatially, and joins columns on the same line with a tab character
* **Ligatures and multi-codepoint CMap destinations** — `fl`, `fi`, `ffi`, and other multi-codepoint sequences are emitted as their full character sequences
* **Both PDF coordinate systems** — auto-detects whether Y increases upward (standard PDF origin) or downward (Word/LibreOffice flipped CTM) and sorts accordingly
* **Multiple content streams per page** — all streams are collected before sorting, so the final output reflects visual reading order across the whole page

---

## What it does not handle

* **Encrypted PDFs** — files with a `/Encrypt` dictionary are detected early and the function returns an empty string
* **Scanned / image-only PDFs** — there is no OCR; if a PDF contains no text operators, nothing is extracted
* **Type 1 / TrueType fonts without a ToUnicode CMap** — some very old or highly customised PDFs omit the CMap; glyphs may come out as raw byte values or be silently skipped
* **Right-to-left text** — Arabic, Hebrew, etc. will be extracted but word order within a line may be reversed
* **`"` (double-quote) TJ operator** — the operator that sets word and character spacing then shows a string in one step is not currently handled; it is uncommon outside specialised generators
* **bfrange array-destination form** — a rare CMap variant where a single range maps to an explicit array of destination codes (instead of a sequential run) is not parsed; sequential runs, which cover the vast majority of real-world CMaps, are fully supported

---

## Installation

1. In the VBA editor, go to **File → Import File** and select `VBA-PdfTXT.bas`
2. That is it. No references to set, no extra modules needed.

---

## Usage

```vba
Dim sText As String
sText = PDF_ExtractText("C:\path\to\file.pdf")

If Len(sText) = 0 Then
    ' Encrypted, image-only, or some other issue reading the file
Else
    Debug.Print sText
End If
```

The returned string uses:

* **Line feed** (`Chr(10)`) between lines
* **Tab** (`Chr(9)`) between items on the same visual line (e.g. a label and its value in a two-column layout)

---

## Diagnostics

If a PDF returns empty text unexpectedly, call the diagnostic function to inspect its streams:

```vba
Debug.Print PDF_DiagnoseStreams("C:\path\to\file.pdf")
```

This prints each stream's position, length, filter type, whether it was identified as a content stream, and whether it is an object stream — useful for spotting encrypted, image-only, or non-standard PDFs.

---

## Architecture

The module is a single `.bas` file with 22 functions. The public surface is just two; everything else is internal.

```
PDF_ExtractText           ← public entry point
PDF_DiagnoseStreams        ← public diagnostic helper

PDF_ReadFileBytes          reads raw file bytes into a Byte array
PDF_BytesToLatin1          converts a Byte array to a 1:1 Latin-1 string
PDF_ProcessAllStreams       two-pass: (1) collect CMaps, (2) extract positioned runs
PDF_ExtractObjStmCMaps     extracts ToUnicode CMaps from PDF 1.5+ /ObjStm bundles
PDF_IsContentStream        rejects font files, images, ICC profiles, XMP, etc.
PDF_ResolveFilterRef       resolves indirect /Filter object references (e.g. SAP PDFs)
PDF_ParseDecodeParms       parses /Predictor and /Columns from a stream dictionary
PDF_ApplyPredictor         reverses PNG and TIFF predictor encoding (PDF spec Table 8)
PDF_DecompressDeflate      entry point for DEFLATE decompression (strips zlib header)
PDF_DecodeASCII85          decodes ASCII85 (Base-85) encoded streams (PDF spec §7.4.3)
VBA_Inflate                RFC 1951 DEFLATE decompressor (all three block types)
INF_ReadBits               bit-level reader for VBA_Inflate
INF_BuildTable             builds a Huffman decode table from code-length arrays
INF_DecodeHuff             decodes one symbol using a Huffman table
PDF_ExtractTextOps         parses PDF content stream operators, tracks text position
PDF_HexDecode              decodes <hex> strings, preserving null bytes for CID pairing
PDF_CleanText              strips control characters and collapses double spaces
PDF_ParseCMap              parses beginbfchar / beginbfrange CMap sections
PDF_ApplyCMap              maps 1-byte or 2-byte CID values to Unicode via a parsed CMap
PDF_SortAndJoin            sorts positioned text runs into visual reading order
```

---

## Performance

The bottleneck is the pure-VBA DEFLATE decompressor. On a typical modern machine:

| PDF type | Pages | Approx. time |
|---|---|---|
| Simple text, uncompressed | any | < 100 ms |
| Compressed, standard 1-byte encoding | 1–5 | ~200–500 ms |
| Compressed, CID encoding (Word/LibreOffice export) | 1–5 | ~300–800 ms |
| Compressed, CID encoding with ObjStm fonts | 1–5 | ~400–900 ms |
| ASCII85 + FlateDecode (ReportLab, Ghostscript ps2pdf) | 1–5 | ~300–700 ms |

For bulk processing, call `PDF_ExtractText` in a loop with `DoEvents` between files to keep the host application responsive.

---

## Known edge cases

**Two-column layouts** — labels and values on the same visual line are joined with a tab. Split on `Chr(9)` to get them as separate fields.

**Line spacing tolerance** — two text runs are treated as the same line if their Y coordinates are within 8 PDF points. This handles most layouts comfortably. If a document with very tight superscripts or very wide line spacing merges or splits lines incorrectly, adjust the `Y_TOL` constant in `PDF_SortAndJoin` (stored ×100, so the default 800 = 8.0 points).

**Windows code page** — `PDF_ApplyCMap` uses `Asc()` to recover byte values from VBA's internal string representation. This correctly round-trips all byte values on systems using a single-byte ANSI code page (CP1252, CP1250, etc.). Behaviour on DBCS code page systems (Japanese, Chinese, Korean) has not been tested.

**Incremental PDF updates** — the parser scans the file linearly and ignores the cross-reference table. In PDFs with incremental updates, superseded object versions may also be parsed alongside their replacements. In practice this is harmless for text extraction.

---

Copyright © 2026, [rafael-yml](https://rafael-yml.lovable.app/)

---

## License

MIT License — see [LICENSE](LICENSE) for details.
