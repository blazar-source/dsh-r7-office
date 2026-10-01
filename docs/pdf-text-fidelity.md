# PDF text fidelity (r7_convert)

## The defect

On a stock R7-Office install, `r7_convert` to PDF used to produce documents whose
body text was **absent** — not vectorised, absent. Independent renderers disagreed
about the result because there was nothing to render: pages came out blank apart
from list numbers and the page footer, and copy/paste returned nothing.

Measured on `R7_MCP_DOCX_Acceptance.docx` with the pre-fix converter:

```
bytes 68877  pages 5  textGlyphs 28  pathOperators 1690  ToUnicode 13  cyrillic 0
per page: 1: 13 glyphs / 1690 paint ops   2..5: 0 glyphs
PDFium: 57 extracted characters, 0 Cyrillic; page 1 rendered 651 non-white pixels
        (pages 4-5, which hold a table, rendered 217k+)
```

The "path operators" were fills with no path to paint:

```
q
0 0 0 rg
/E1 gs
f          <- nothing between gs and f
Q
```

repeated once per glyph (1669 times on page 1). Latin text was lost identically, so
this was never a Cyrillic-specific or font-specific problem.

## Root cause

`x2t` has two invocation forms:

```
x2t "path_to_params_xml"
x2t "path_to_file_1" "path_to_file_2" ["path_to_font_selection"]
```

The two-argument form makes the DOCX/PPTX/HTML renderer fall back to the
exe-relative `converter/DoctRenderer.config`, whose `<allfonts>` entry points at
`<install>\editors\sdkjs\common\AllFonts.js`. On a normal installation that file is a
**0-byte stub** (the real font list is per-machine and lives in the user profile,
written by R7-Office Desktop's `CApplicationFontsWorker` on first launch:

```
%LOCALAPPDATA%\R7-Office\Editors\data\fonts\AllFonts.js     (~181 KB, version 2)
```

With an empty font list the renderer resolves no glyph outlines and emits an empty
fill per glyph instead of a text object. XLSX does not use that renderer, which is
why spreadsheets always converted correctly — the only genuinely healthy format.

The single-argument params-XML form accepts `<m_sAllFontsPath>` and fixes it.

## The fix

`src/r7/adapter.js`:

1. discovers a usable `AllFonts.js` (user profile, `R7_ALL_FONTS_JS`, install
   tree; 0-byte stubs are rejected), and
2. runs `x2t` through the params-XML form with `<m_sAllFontsPath>`, falling back to
   the CLI form only if that fails, then
3. repairs malformed `ToUnicode` CMap entry counts, and
4. reports `pdfTextQuality` and **throws** rather than returning a text-less PDF.

Unusable input degrades loudly. On a fresh install where R7 Desktop has never run,
`AllFonts.js` does not exist and the caller sees:

```
R7 x2t produced a PDF with no extractable text: 28 text glyphs versus 4753 glyph
fills that painted nothing (verdict "mixed"). Cause: no R7 font list (AllFonts.js)
was found, so the DOCX/PPTX/HTML renderer had no fonts to draw with. Run R7-Office
Desktop once so it generates %LOCALAPPDATA%\R7-Office\Editors\data\fonts\AllFonts.js,
or point the R7_ALL_FONTS_JS environment variable at a valid AllFonts.js. Pass
allowOutlinedPdf:true to keep the unusable PDF and inspect pdfTextQuality for
details.
```

The error carries `error.pdfTextQuality` with the measurements. Passing
`allowOutlinedPdf: true` (adapter option) keeps the broken file instead.

## Second, independent defect: malformed ToUnicode counts

`x2t` sometimes writes a `beginbfchar` header whose declared entry count does not
match the entries that follow. Observed for the plain-Arial font that carries body
text: `944 beginbfchar` wrapping 94 entries. Conforming readers reject the block, so
the text renders correctly but extracts as raw CIDs.

This is **not** a universal x2t defect, so the repair is detection-gated rather than
unconditional — correct CMaps are left untouched and the file is not rewritten at
all when there is nothing to fix:

| document | malformed blocks |
|---|---|
| DOCX acceptance, no `w:rFonts` in the package | **1** (`944` declared / `94` actual) |
| DOCX acceptance re-themed with `w:rFonts` + `w:lang` on every run | 0 |
| minimal Cyrillic DOCX | 0 |
| PPTX acceptance | 0 |
| XLSX acceptance | 0 |

## Verifying any PDF

`r7_pdf_inspect` (MCP tool) and `inspectPdf()` (`src/r7/pdf-inspect.js`) report
`textGlyphs`, `pathOperators`, `emptyFillOperators`, per-font `toUnicodeMappings` /
`cyrillicMappings` / `embedded`, `hasExtractableText`, `hasCyrillicText`,
`textLossSuspected` and a `verdict` of `text` | `outlined` | `mixed`.

The signature of this defect is `emptyFillOperators` high while `textGlyphs` is near
zero. A genuinely vectorised PDF has `textGlyphs === 0` but *no* empty fills, because
each glyph outline builds a path before it is filled.

Pass `repair: true` to have the tool apply the ToUnicode count repair in place.

## Results after the fix

Through the public `r7_convert` tool:

| format | route | bytes | pages | textGlyphs | emptyFill | cmapMaps | cyrillic |
|---|---|---|---|---|---|---|---|
| DOCX | `R7Adapter.convert` (params XML) | 141821 | 5 | 4781 | 0 | 163 | 109 |
| PPTX | `R7Adapter.convert` (params XML) | 120687 | 7 | 814 | 0 | 175 | 114 |
| XLSX | `xlsxEngine.exportPdf` | 297258 | 6 | 874 | 0 | 219 | 145 |

Independent extraction with PDFium (`pypdfium2`, a different engine from the project):

| PDF | extracted chars | Cyrillic | control chars |
|---|---|---|---|
| DOCX | 5907 | 4415 (74.7 %) | 0 |
| PPTX | 992 | 707 (71.3 %) | 0 |
| XLSX | 984 | 508 (51.6 %) | 0 |

Pre-fix the same DOCX PDF yielded 57 characters and 0 Cyrillic.

## Known limits

- The fix depends on an `AllFonts.js` produced by R7-Office Desktop. The plugin does
  not generate one (the format's `__fonts_infos` / `g_fonts_selection_bin` blobs come
  from R7's own font worker). Without it, conversion **fails loudly** instead of
  emitting a blank PDF.
- `x2t` writes the length object *after* the stream it belongs to, so a PDF parser
  must honour `/Length` and fall back to an `endstream`+`endobj` scan for forward
  references. A naive `/Length\s+(\d+)/` regex backtracks on `/Length 94 0 R` and
  reports a length of 1, silently truncating every compressed stream.
- R7's 0-byte `AllFonts.js` stub under `Program Files` cannot be replaced by the
  plugin (the directory is not writable without elevation) — the params-XML form
  sidesteps it instead.
