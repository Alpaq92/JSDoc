# JSDoc

[![CI](https://img.shields.io/github/actions/workflow/status/Alpaq92/JSDoc/ci.yml?branch=main&label=CI)](https://github.com/Alpaq92/JSDoc/actions/workflows/ci.yml)
[![Tests](https://img.shields.io/github/actions/workflow/status/Alpaq92/JSDoc/ci.yml?branch=main&label=tests)](https://github.com/Alpaq92/JSDoc/actions/workflows/ci.yml)
[![Live demo](https://img.shields.io/badge/demo-live-success)](https://alpaq92.github.io/JSDoc/)
[![License: 0BSD](https://img.shields.io/badge/License-0BSD-blue.svg)](LICENSE)

Reads the text out of legacy Microsoft Word `.doc` files (Word 97–2003, the old binary OLE2 format) — right in the browser, with no dependencies. The whole API is one function:

```js
docToText(input) // → string, or null if the file can't be read
```

Give it an `ArrayBuffer`, `Uint8Array`, or Node `Buffer`; you get back the document's body text, or `null` when the file is something it won't touch (Word 6/95, encrypted, not a `.doc`, or corrupt) — your cue to fall back to a download link.

It's written from scratch against Microsoft's published format specs, with no GPL code anywhere near it, so it ships under `0BSD` and drops cleanly into a permissive codebase.

## Try it

**▶ Live demo: <https://alpaq92.github.io/JSDoc/>**

[`index.html`](index.html) is a no-build demo: drop a `.doc` onto the page (or hit **Try a sample**) and the text appears, extracted locally — nothing is uploaded. Three views: **Formatted** (rebuilds tables), **Plain text**, and **Edit** (editable in place).

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

`docToText.images(input)` returns `[{ mime, bytes }]` for **embedded raster images** (PNG/JPEG), carved by signature from the reassembled CFB streams and validated to their real end marker. The demo shows them as a gallery.

**Tracked changes are resolved as "accept all":** deleted text is dropped (the `sprmCFRMarkDel` revision mark in the CHPX bin table) and inserted text kept.

Not handled yet (and where each would slot in):

- **Tables** come out as one row per line with tab-separated columns — close to the original grid, though merged or empty cells can nudge the columns.
- **WMF/EMF images** — only PNG/JPEG are extracted (see `docToText.images`); the common WMF/EMF *metafiles* can't be rendered in-browser without a heavy, non-permissive converter, so they're skipped. Exact inline image *placement* isn't reconstructed either — images come out as a set.
- **Exact page layout** (line/page-break positions, columns, precise spacing) — that needs a real layout engine, not just property resolution.

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

`word-extractor` (MIT) is a dev-only test dependency and isn't part of the shipped code. The demo's sample, [`samples/license-comparison.doc`](samples/license-comparison.doc), is a renamed copy of `test03.doc` from word-extractor (MIT, © 2016–2021 Stuart Watt) — permissive, so it doesn't affect the `0BSD` license of the extractor.
