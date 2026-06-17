# JSDoc

[![CI](https://img.shields.io/github/actions/workflow/status/Alpaq92/JSDoc/ci.yml?branch=main&label=CI)](https://github.com/Alpaq92/JSDoc/actions/workflows/ci.yml)
[![Tests](https://img.shields.io/github/actions/workflow/status/Alpaq92/JSDoc/ci.yml?branch=main&label=tests)](https://github.com/Alpaq92/JSDoc/actions/workflows/ci.yml)
[![Live demo](https://img.shields.io/badge/demo-live-success)](https://alpaq92.github.io/JSDoc/)
[![License: 0BSD](https://img.shields.io/badge/License-0BSD-blue.svg)](LICENSE)

**Reads — and writes — legacy Microsoft Word `.doc` files** (Word 97–2003, the old binary OLE2 format), right in the browser, with no dependencies. Written clean-room from Microsoft's published specs, so it ships under `0BSD`.

```js
docToText(input)          // → body text (string), or null if it can't be read
docToText.sections(input) // → { body, footnotes, headers, … } — text of each story
docToText.html(input)     // → styled HTML per story (bold/italic/size/colour/font + tables)
docToText.model(input)    // → { body, … } styled paragraph/run model (feeds the writer)
docToText.images(input)   // → [{ mime, bytes }] — embedded PNG/JPEG
textToDoc(input)          // → Uint8Array — write a .doc back (string or a styled model)
```

`docToText` takes an `ArrayBuffer`, `Uint8Array`, or Node `Buffer`; it returns the body text, or `null` when the file is something it won't touch (Word 6/95, encrypted, not a `.doc`, corrupt) — your cue to fall back to a download link. `textToDoc` is the inverse. No GPL code anywhere near it, so the whole thing drops cleanly into a permissive codebase.

## Try it

**▶ Live demo: <https://alpaq92.github.io/JSDoc/>**

[`index.html`](index.html) is a no-build demo: drop a `.doc` onto the page (or hit **Try a sample**) and it's parsed locally — nothing is uploaded. Three views — **Formatted** (styling, tables, lists, footnotes/headers, and inline images, rendered from the same paragraph model the writer uses), **Plain text**, and **Edit** (editable in place) — plus **Download** as `.txt`, `.html`, or a real `.doc`.

- **Locally:** serve the folder (`npx serve`, or `python -m http.server`) and open `index.html`.
- **On GitHub Pages:** Settings → Pages → deploy from `main` / root, then it's live at <https://alpaq92.github.io/JSDoc/>. The sample in [`samples/`](samples/) is bundled, so the page needs no network at all.

## Using it

In the browser, `docToText` is a UMD global:

```html
<script src="src/docToText.js"></script>
<script>
  const buf = await (await fetch('/file.doc')).arrayBuffer();
  const text = docToText(buf);
  text === null ? showDownloadLink() : showText(text);
</script>
```

In Node it's a plain `require`:

```js
const docToText = require('./src/docToText.js');
const text = docToText(require('fs').readFileSync('file.doc'));
```

## What you get

The main body text and its paragraph breaks. Smart quotes and non-Latin scripts come through correctly, and field codes are stripped — you keep the result (say, the page number) and lose the `PAGE` instruction behind it.

`docToText.sections(input)` returns the document's separate stories — `{ body, footnotes, headers, annotations, endnotes, textboxes, headerTextboxes }` (each a string) — so **headers, footers, footnotes, endnotes, comments, and text boxes** come through too; they sit right after the body in the same piece table. `body` is exactly what `docToText()` returns.

`docToText.html(input)` returns those same stories as **styled HTML** — each run wrapped in a `<span>` carrying its **bold / italic / underline / strikethrough, size, colour, and font** — with `\t` between table cells and `\n` at row/paragraph breaks. Formatting is fully resolved through the stylesheet (paragraph style → character style → direct run properties), so formatting that lives in a *style* — a heading's bold, a hyperlink's blue/underline — isn't lost, not just directly-applied sprms. The demo renders this as the Formatted view.

`docToText.images(input)` returns `[{ mime, bytes }]` for **embedded raster images** (PNG/JPEG), carved by signature from the reassembled CFB streams and validated to their real end marker. (The demo renders images inline via `docToText.model()`; this API returns them as a flat set.)

**Tracked changes are resolved as "accept all":** deleted text is dropped (the `sprmCFRMarkDel` revision mark in the CHPX bin table) and inserted text kept.

Not handled yet (and where each would slot in):

- **Tables** come out as one row per line with tab-separated columns — close to the original grid, though merged or empty cells can nudge the columns.
- **WMF/EMF images** — only PNG/JPEG are extracted (see `docToText.images`); the common WMF/EMF *metafiles* can't be rendered in-browser without a heavy, non-permissive converter, so they're skipped. Exact inline image *placement* isn't reconstructed either — images come out as a set.
- **Exact page layout** (line/page-break positions, columns, precise spacing) — that needs a real layout engine, not just property resolution.

## Writing a `.doc`

The reverse direction. **`textToDoc(input)`** ([src/textToDoc.js](src/textToDoc.js)) writes a Word 97–2003 **binary `.doc`** and returns a `Uint8Array`. `input` is either a plain string or a **styled model** from `docToText.model()`, so you can round-trip formatting: `textToDoc(docToText.model(buf).body)`. The demo's **Download .doc** button uses the model, so a loaded document comes back out with its formatting intact.

```js
const textToDoc = require('./src/textToDoc.js');
const docToText = require('./src/docToText.js');
fs.writeFileSync('plain.doc', Buffer.from(textToDoc('Hello\nWorld')));        // from text
fs.writeFileSync('rich.doc',  Buffer.from(textToDoc(docToText.model(buf).body))); // round-trip formatting
```

**How, and why it's not from scratch.** A `.doc` synthesised entirely from spec round-trips through lenient parsers but **real word processors reject it** — they require a valid stylesheet, section table, and character/paragraph property tables, and getting every one right blind is the wall (Apache POI doesn't build a `.doc` from scratch either). So the writer **injects** content into a tiny **bundled blank-document skeleton** — a genuine app-saved empty `.doc`, stripped to its `WordDocument` + `1Table` streams (FIB, stylesheet, sections, fonts), embedded in the module. It reuses those structures and swaps in the text, piece table, and freshly built CHPX/PAPX property pages. Pass your own blank `.doc` as a second argument to use a different skeleton: `textToDoc(input, myTemplateBytes)`.

**Writes back:** paragraphs with **alignment** (`sprmPJc`), **spacing & indentation** (space before/after, left/right/first-line indent, and line spacing — `sprmPDyaBefore`/`sprmPDyaAfter`, `sprmPDxaLeft`/`sprmPDxaRight`/`sprmPDxaLeft1`, `sprmPDyaLine`), **lists** (bullet / numbered, with nesting — `sprmPIlfo`/`sprmPIlvl` against the skeleton's built-in list definitions), **character formatting** (bold / italic / underline / strike / size / colour / **font**, as CHPX sprms — fonts the skeleton lacks are appended to its `SttbfFfn`), **tables** (cell marks + `sprmPFInTable` / `sprmPFTtp` / `sprmTDefTable` with borders), and **inline images** (PNG/JPEG re-embedded as an OfficeArt picture in a `Data` stream, sized and placed in the right paragraph). Literal tabs are preserved. **Not yet:** hyperlink URLs.

The output is checked three ways: read back by our `docToText` (`.model()` re-reads the table cells to prove they're real cells, not flattened tabs) *and* the unrelated `word-extractor` ([test/styled.test.js](test/styled.test.js)), and — since the lenient parsers were exactly the problem — confirmed to **open as real bold text and a real table in a word processor** (SoftMaker TextMaker), driven through its COM automation ([scripts/read-with-textmaker.ps1](scripts/read-with-textmaker.ps1)).

## How it works

Two layers, both taken straight from the spec:

1. **The container** ([MS-CFB](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/53989ce4-7b05-4f8d-829b-d08d6148375b)) — a `.doc` is an OLE2 compound file, essentially a little FAT-style filesystem. Parse the header, follow the sector chains, and pull out the `WordDocument` and table streams.
2. **The text** ([MS-DOC](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/ccd7b486-7881-484c-a137-51170af7cc22)) — read the FIB header to locate the piece table, then walk the pieces, decoding each as either Windows-1252 or UTF-16 and joining them up to the body length. That's the [§2.4.1 "Retrieving Text"](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/01d5d8c4-cf9c-4ef9-80fd-439e763cfe01) algorithm.

A small detail worth calling out: the compressed (8-bit) encoding [isn't quite Windows-1252](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/aa2e55a2-f4f2-4795-bab5-6d9d7a0ed249) — the spec remaps 24 specific bytes and leaves the rest as raw code points. The lookup table is copied from the spec, so bytes like `0x80`, `0x8E`, and `0x9E` decode exactly the way Word means them, not the way a generic cp1252 decoder would.

Both specs are free under Microsoft's [Open Specification Promise](https://go.microsoft.com/fwlink/?LinkId=214445), which expressly allows copying them to build an implementation — that's what makes a clean-room `0BSD` build legitimate, rather than porting GPL tools like catdoc or antiword (which were never read).

## Tests

```bash
npm test            # offline, no dependencies
npm run test:oracle # compares against word-extractor on real .doc files
```

`npm test` builds a spec-valid `.doc` in memory and checks the result against a known answer — covering both storage paths, both encodings, field codes, and the graceful-`null` cases. `npm run test:oracle` downloads real Word documents and diffs our output against [word-extractor](https://github.com/morungos/node-word-extractor) (MIT) as an independent reference, used only for comparison and never copied. Verified across 15 real files.

## License

[`0BSD`](LICENSE) — public-domain-equivalent, no attribution required, so it can live anywhere.

The demo's bundled sample, [`samples/detailed-sample.doc`](samples/detailed-sample.doc), is a small document exercising the reader's range — styled text, a bullet list, a table, and an embedded image — with public-domain Lorem Ipsum filler text. `word-extractor` (MIT) is a dev-only test dependency (the oracle in `test/oracle.test.js`) and isn't part of the shipped code.

The writer's **skeleton** (a structural empty `.doc`) and **inline-picture machinery** are reverse-engineered from blank/one-image documents saved by a real word processor (SoftMaker FreeOffice), reduced to their structural bytes and embedded into `src/textToDoc.js` as base64. They carry no authored content, so no copyrightable expression — and the source documents aren't bundled. [`scripts/embed-template.js`](scripts/embed-template.js) regenerates the skeleton embed from such a blank `.doc`.
