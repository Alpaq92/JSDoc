// SPDX-License-Identifier: 0BSD
/*
 * styled.test.js — the styled writer: docToText.model() -> textToDoc() must
 * round-trip character formatting (bold/italic/underline/strike/size/colour)
 * and reconstruct tables (cell marks + table PAPX). We read the written .doc
 * back with docToText (.html for styling, .model for the table structure, plain
 * for text) and cross-check with word-extractor. Real-editor rendering (it opens
 * as an actual table/bold text in a word processor) is verified separately
 * against SoftMaker TextMaker — see scripts/read-with-textmaker.ps1.
 */
'use strict';
var fs = require('fs');
var path = require('path');
var textToDoc = require('../src/textToDoc.js');
var docToText = require('../src/docToText.js');

var failures = 0;
function check(name, cond) { console.log('  ' + (cond ? 'ok  ' : 'FAIL') + ' ' + name); if (!cond) failures++; }

console.log('styled writer — model round-trip\n');

// 1) Hand-built styled model -> every run property survives a re-parse.
var out = textToDoc([
  { runs: [{ text: 'Title', b: true, size: 16, color: 0x123456, font: 'Courier New' }], kind: 'p' },
  { runs: [{ text: 'plain ' }, { text: 'italic', i: true }, { text: ' ' }, { text: 'under', u: true }, { text: ' ' }, { text: 'struck', strike: true }], kind: 'p' }
]);
check('hand-built: produces a CFB', out[0] === 0xD0 && out[1] === 0xCF);
var html = docToText.html(out).body;
check('hand-built: bold', /font-weight:bold/.test(html));
check('hand-built: italic', /font-style:italic/.test(html));
check('hand-built: underline', /text-decoration:[^"]*underline/.test(html));
check('hand-built: strike', /text-decoration:[^"]*line-through/.test(html));
check('hand-built: size 16pt', /font-size:16pt/.test(html));
check('hand-built: colour', /color:rgb\(86,\s*52,\s*18\)/.test(html));   // 0x123456 COLORREF -> rgb(0x56,0x34,0x12)
check('hand-built: font (appended to table)', /font-family:'Courier New'/.test(html));
check('hand-built: text intact', docToText(out).replace(/\r?\n/g, ' ').indexOf('Title plain italic under struck') !== -1);

// 2) Round-trip the bundled sample's styled model; tables must come back as tables.
var sample = docToText.model(fs.readFileSync(path.join(__dirname, '..', 'samples', 'detailed-sample.doc'))).body;
var doc2 = textToDoc(sample);
var html2 = docToText.html(doc2).body;
check('sample: bold preserved', /font-weight:bold/.test(html2));
check('sample: colour preserved', /color:rgb/.test(html2));
var re = docToText.model(doc2).body, kinds = {};
re.forEach(function (p) { kinds[p.kind] = (kinds[p.kind] || 0) + 1; });
check('sample: table cells round-trip (real 0x07 cells, not tabs)', kinds.cell > 0 && kinds.rowEnd > 0);

// 3) Inline image + paragraph alignment round-trip. A 1x1 PNG in a centred
// paragraph: the image must embed (and read back) and the alignment survive.
var PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M8AAAMBAQDJ/pLvAAAAAElFTkSuQmCC', 'base64');
var imgDoc = textToDoc([
  { runs: [{ text: 'Above the image.' }], kind: 'p' },
  { runs: [{ image: { mime: 'image/png', bytes: new Uint8Array(PNG) } }], kind: 'p', align: 1 }
]);
var imgData = textToDoc.readCfb(imgDoc).byName['Data'];
check('image embeds into a Data stream (exact bytes)', !!imgData && Buffer.from(imgData).indexOf(PNG) !== -1);
check('centred paragraph round-trips (sprmPJc)', docToText.model(imgDoc).body.some(function (p) { return p.align === 1; }));

// 4) Lists: bullet + numbered paragraphs map to the skeleton's built-in lists
// (sprmPIlfo/sprmPIlvl) and read back as list paragraphs with their level.
var listDoc = textToDoc([
  { runs: [{ text: 'one' }], kind: 'p', list: { kind: 'bullet', ilvl: 0 } },
  { runs: [{ text: 'sub' }], kind: 'p', list: { kind: 'bullet', ilvl: 1 } },
  { runs: [{ text: 'step' }], kind: 'p', list: { kind: 'number', ilvl: 0 } }
]);
var lm = docToText.model(listDoc).body.filter(function (p) { return p.list; });
check('lists round-trip (3 list paragraphs, with level)', lm.length === 3 && lm[1].list.ilvl === 1);
check('bullet list detected', lm[0].list.kind === 'bullet');

// 5) Paragraph spacing & indentation: left/right/first-line indent, space
// before/after, and line spacing (twips) survive as PAPX sprms; an unspaced
// paragraph stays bare (no .pp).
var spaced = textToDoc([
  { runs: [{ text: 'Indented, spaced, 1.5 line.' }], kind: 'p', pp: { indL: 720, ind1: -360, spB: 120, spA: 240, line: 360, lineMult: 1 } },
  { runs: [{ text: 'Plain.' }], kind: 'p' }
]);
var sp = docToText.model(spaced).body;
var spPP = (sp.filter(function (p) { return p.pp; })[0] || {}).pp;
check('indent/spacing round-trip (left+first-line+before/after+line)',
  !!spPP && spPP.indL === 720 && spPP.ind1 === -360 && spPP.spB === 120 && spPP.spA === 240 && spPP.line === 360);
check('unspaced paragraph carries no .pp', sp.some(function (p) { return p.runs[0] && p.runs[0].text === 'Plain.' && !p.pp; }));

// 6) Hyperlinks: a run with a URL becomes a HYPERLINK field (begin/instruction/
// separator/result/end, the marks flagged sprmCFSpec, plus a PlcfFldMom). It
// reads back with the URL on the result run, and the instruction stays hidden.
var linkDoc = textToDoc([
  { runs: [{ text: 'See ' }, { text: 'the site', url: 'https://example.com/docs', u: true }, { text: '.' }], kind: 'p' }
]);
var lr = null;
docToText.model(linkDoc).body.forEach(function (p) { (p.runs || []).forEach(function (r) { if (r.url) lr = r; }); });
check('hyperlink URL round-trips on the result run', !!lr && lr.url === 'https://example.com/docs' && lr.text === 'the site');
var ltext = docToText(linkDoc);
check('hyperlink instruction stays hidden from the text', ltext.indexOf('HYPERLINK') === -1 && ltext.replace(/\s+/g, ' ').indexOf('See the site.') !== -1);

// 7) Footnotes: a body anchor (ftnRef) + a footnote-document paragraph round-trip.
// The reader re-emits the anchor, the footnote text lands in the footnote story
// (ccpFtn + PlcffndRef/PlcffndTxt), and the body text stays clean.
var ftnDoc = textToDoc({
  body: [{ runs: [{ text: 'Body' }, { ftnRef: 0 }, { text: ' end.' }], kind: 'p' }],
  footnotes: [{ runs: [{ text: ' A footnote.' }], kind: 'p' }]
});
var fm = docToText.model(ftnDoc);
check('footnote anchor round-trips in the body', fm.body.some(function (p) { return (p.runs || []).some(function (r) { return r.ftnRef === 0; }); }));
check('footnote text lands in the footnote story', (docToText.sections(ftnDoc).footnotes || '').indexOf('A footnote.') !== -1);
check('footnote stays out of the body text', docToText(ftnDoc).indexOf('A footnote.') === -1 && docToText(ftnDoc).replace(/\s+/g, ' ').indexOf('Body end.') !== -1);

// 8) Headers/footers: a header + footer round-trip into the header document
// (PlcfHdd story 7 / 9), read back as model.header/.footer, kept out of the body.
var hfDoc = textToDoc({
  body: [{ runs: [{ text: 'Body.' }], kind: 'p' }],
  header: [{ runs: [{ text: 'The Header' }], kind: 'p' }],
  footer: [{ runs: [{ text: 'The Footer' }], kind: 'p' }]
});
var hfm = docToText.model(hfDoc);
function flatText(ps) { return ps ? ps.map(function (p) { return (p.runs || []).map(function (r) { return r.text || ''; }).join(''); }).join(' ').trim() : ''; }
check('header round-trips into the header document', flatText(hfm.header).indexOf('The Header') !== -1);
check('footer round-trips into the header document', flatText(hfm.footer).indexOf('The Footer') !== -1);
check('header/footer stay out of the body', docToText(hfDoc).indexOf('The Header') === -1 && docToText(hfDoc).indexOf('The Footer') === -1);

// 9) Endnotes (a near-clone of footnotes) + coexistence: a body endnote anchor +
// endnote text round-trip (ccpEdn + PlcfendRef/PlcfendTxt), kept distinct from
// footnotes, and a single doc carries footnote + endnote + header + footer at once.
var combo = textToDoc({
  body: [{ runs: [{ text: 'A' }, { ftnRef: 0 }, { text: ' B' }, { endRef: 0 }, { text: ' C' }, { comRef: 0 }, { text: ' D.' }], kind: 'p' }],
  footnotes: [{ runs: [{ text: ' fn' }], kind: 'p' }],
  endnotes: [{ runs: [{ text: ' en' }], kind: 'p' }],
  annotations: [{ runs: [{ text: ' cm' }], kind: 'p' }],
  header: [{ runs: [{ text: 'hd' }], kind: 'p' }],
  footer: [{ runs: [{ text: 'ft' }], kind: 'p' }]
});
var cm = docToText.model(combo);
function hasRef(body, key, v) { return body.some(function (p) { return (p.runs || []).some(function (r) { return r[key] === v; }); }); }
check('footnote, endnote and comment anchors round-trip distinctly', hasRef(cm.body, 'ftnRef', 0) && hasRef(cm.body, 'endRef', 0) && hasRef(cm.body, 'comRef', 0));
check('endnote + comment text land in their stories', (docToText.sections(combo).endnotes || '').indexOf('en') !== -1 && (docToText.sections(combo).annotations || '').indexOf('cm') !== -1);
check('all five stories coexist (footnote/endnote/comment/header/footer), body clean',
  flatText(cm.footnotes).indexOf('fn') !== -1 && flatText(cm.endnotes).indexOf('en') !== -1 && flatText(cm.annotations).indexOf('cm') !== -1 &&
  flatText(cm.header).indexOf('hd') !== -1 && flatText(cm.footer).indexOf('ft') !== -1 &&
  docToText(combo).replace(/\s+/g, ' ').indexOf('A B C D.') !== -1);

// 10) Text box: a body anchor (tbxRef) + box text round-trip via the OfficeArt
// drawing (TBX_DGG) + FSPA (PlcfspaMom) + PlcftxbxTxt. The box text stays out of
// the body, and ccpTxbx is stable across repeated round-trips (no growth).
var tbxDoc = textToDoc({ body: [{ runs: [{ text: 'See ' }, { tbxRef: 0 }, { text: ' box.' }], kind: 'p' }], textboxes: [{ runs: [{ text: 'Boxed' }], kind: 'p' }] });
var tbm = docToText.model(tbxDoc);
check('text-box anchor round-trips in the body', tbm.body.some(function (p) { return (p.runs || []).some(function (r) { return r.tbxRef === 0; }); }));
check('text-box text lands in the textbox story', (docToText.sections(tbxDoc).textboxes || '').indexOf('Boxed') !== -1);
check('text-box text stays out of the body', docToText(tbxDoc).indexOf('Boxed') === -1);
var tbRt1 = (docToText.sections(textToDoc(tbm)).textboxes || '').length;
var tbRt2 = (docToText.sections(textToDoc(docToText.model(textToDoc(tbm)))).textboxes || '').length;
check('text-box round-trip is stable (ccpTxbx constant)', tbRt1 > 0 && tbRt1 === tbRt2);

// 11) Page setup: margins + page size round-trip via the first section's SEPX
// (sprmSDyaTop/Bottom, sprmSDxaLeft/Right, sprmSXaPage/SYaPage), plus orientation
// (sprmSBOrientation) and column count (sprmSCcolumns) appended to a relocated SEPX.
var pgDoc = textToDoc({ body: [{ runs: [{ text: 'P.' }], kind: 'p' }], page: { top: 1440, bottom: 1440, left: 1800, right: 1800, width: 12240, height: 15840 } });
var pgm = docToText.model(pgDoc).page;
// All six values differ from the skeleton's A4/1417 defaults, so this fails unless
// the writer actually applied input.page.
check('page setup round-trips (margins + page size)', !!pgm && pgm.top === 1440 && pgm.bottom === 1440 && pgm.left === 1800 && pgm.right === 1800 && pgm.width === 12240 && pgm.height === 15840);
// Orientation (landscape — TextMaker-verified: reports landscape, swaps dims) and the
// column count round-trip via the appended/relocated SEPX; a plain portrait page stays bare.
var lsm = docToText.model(textToDoc({ body: [{ runs: [{ text: 'L' }], kind: 'p' }], page: { width: 15840, height: 12240, landscape: true, cols: 2 } })).page;
check('orientation (landscape) round-trips (sprmSBOrientation)', !!lsm && lsm.landscape === true && lsm.width === 15840 && lsm.height === 12240);
check('column count round-trips (sprmSCcolumns)', !!lsm && lsm.cols === 2);
check('plain portrait page carries no orientation/column flags', !pgm.landscape && pgm.cols == null);
// The reader exposes every section's setup as model.sections; a single-section
// document gives a one-element array equal to model.page.
check('reader exposes model.sections (single section -> one entry === page)', (function () { var ss = docToText.model(pgDoc).sections; return Array.isArray(ss) && ss.length === 1 && JSON.stringify(ss[0]) === JSON.stringify(pgm); })());

// 13) Table column widths: a row's rgdxaCenter (sprmTDefTable) round-trips, so
// unequal columns aren't flattened to equal widths.
var twDoc = textToDoc([
  { runs: [{ text: 'A' }], kind: 'cell' },
  { runs: [{ text: 'B' }], kind: 'cell' },
  { runs: [{ text: 'C' }], kind: 'rowEnd', tblw: [0, 1500, 4000, 9000] }
]);
var twRows = docToText.model(twDoc).body.filter(function (p) { return p.kind === 'rowEnd'; });
check('table column widths round-trip (sprmTDefTable rgdxaCenter)', twRows.length === 1 && JSON.stringify(twRows[0].tblw) === '[0,1500,4000,9000]');
// The real sample's three unequal columns survive intact (not flattened to equal).
var twSample = docToText.model(fs.readFileSync(path.join(__dirname, '..', 'samples', 'detailed-sample.doc')));
var twOrig = twSample.body.filter(function (p) { return p.kind === 'rowEnd'; });
var twRe = docToText.model(textToDoc(twSample)).body.filter(function (p) { return p.kind === 'rowEnd'; });
check('detailed-sample column widths survive round-trip (unequal preserved)',
  twOrig.length > 0 && twRe.length === twOrig.length &&
  JSON.stringify(twRe[0].tblw) === JSON.stringify(twOrig[0].tblw) &&
  JSON.stringify(twOrig[0].tblw) !== JSON.stringify([0, 3000, 6000, 9000]));

// 13b) Table cell shading + merging. Shading is a per-cell background COLORREF
// (sprmTDefTableShd, 0xD612); merging is Word's native form — a row with fewer,
// wider cells (a one-cell row spans the whole table).
var shRow = docToText.model(textToDoc([
  { runs: [{ text: 'r' }], kind: 'cell', shd: 0x0000FF },   // red   (COLORREF 0x00BBGGRR)
  { runs: [{ text: 'y' }], kind: 'cell', shd: 0x00FFFF },   // yellow
  { runs: [{ text: 'p' }], kind: 'rowEnd', tblw: [0, 3000, 6000, 9000] }
])).body.filter(function (p) { return p.kind === 'rowEnd'; })[0];
check('cell shading round-trips (sprmTDefTableShd COLORREF)',
  !!shRow && JSON.stringify(shRow.tblShd) === JSON.stringify([0x0000FF, 0x00FFFF, null]));
// Merge flags (fFirstMerged / fMerged in each TC80's tcgrf) round-trip as tblMerge.
var mgRow = docToText.model(textToDoc([
  { runs: [{ text: 'A' }], kind: 'cell', hmerge: 'start' },
  { runs: [{ text: 'B' }], kind: 'cell', hmerge: 'cont' },
  { runs: [{ text: 'C' }], kind: 'rowEnd', tblw: [0, 3000, 6000, 9000] }
])).body.filter(function (p) { return p.kind === 'rowEnd'; })[0];
check('cell merge flags round-trip (TC80 tcgrf -> tblMerge)',
  !!mgRow && Array.isArray(mgRow.tblMerge) && !!mgRow.tblMerge[0] && mgRow.tblMerge[0].h === 'start' && mgRow.tblMerge[1].h === 'cont');
// The generated showcase sample exercises everything in one document: a one-cell merged
// header, shaded status cells, an empty cell, character extras, tab stops, a shaded +
// bordered callout, and bookmarks — all surviving a read-back.
var mshModel = docToText.model(fs.readFileSync(path.join(__dirname, '..', 'samples', 'feature-showcase.doc')));
var mshRows = mshModel.body.filter(function (p) { return p.kind === 'rowEnd'; });
check('sample: merged header is a one-cell, full-width row',
  mshRows.filter(function (r) { return JSON.stringify(r.tblw) === '[0,9000]'; }).length === 1);
check('sample: status cells carry shading (green/red/yellow)',
  mshRows.filter(function (r) { return r.tblShd && r.tblShd.some(function (c) { return c != null; }); }).length >= 3);
var mshRuns = []; mshModel.body.forEach(function (p) { (p.runs || []).forEach(function (rr) { mshRuns.push(rr); }); });
check('sample: shows superscript and subscript',
  mshRuns.some(function (r) { return r.va === 'super'; }) && mshRuns.some(function (r) { return r.va === 'sub'; }));
check('sample: shows highlighted runs',
  mshRuns.filter(function (r) { return r.highlight != null; }).length >= 2);
check('sample: has tab-stop paragraphs (dot leader)',
  mshModel.body.some(function (p) { return p.pp && p.pp.tabs && p.pp.tabs.some(function (t) { return t.leader === 'dot'; }); }));
check('sample: has a shaded + bordered callout paragraph',
  mshModel.body.some(function (p) { return p.pp && p.pp.shd != null && p.pp.borders && p.pp.borders.top; }));
check('sample: has an empty table cell',
  mshModel.body.filter(function (p) { return (p.kind === 'cell' || p.kind === 'rowEnd') && (!p.runs || !p.runs.length); }).length >= 1);
check('sample: has bookmarks landing on the named text',
  Array.isArray(mshModel.bookmarks) && mshModel.bookmarks.length >= 2 &&
  mshModel.bookmarks.some(function (b) { return b.name === 'highlightedTerm'; }));
check('sample: shows underline styles, small caps, all caps and hidden text',
  mshRuns.some(function (r) { return r.uStyle === 'double'; }) && mshRuns.some(function (r) { return r.uStyle === 'wavy'; }) &&
  mshRuns.some(function (r) { return r.smallCaps; }) && mshRuns.some(function (r) { return r.caps; }) && mshRuns.some(function (r) { return r.hidden; }));
check('sample: has a vertically merged cell (restart + cont)',
  mshRows.some(function (r) { return r.tblMerge && r.tblMerge.some(function (c) { return c && c.v === 'restart'; }); }) &&
  mshRows.some(function (r) { return r.tblMerge && r.tblMerge.some(function (c) { return c && c.v === 'cont'; }); }));
// A REAL Word-saved table (not our writer): its "Merge Cells" header reads back as one
// full-width cell, and its shaded cell's fill is read from sprmTDefTableShd. Proves the
// reader handles genuine Word output — in particular the 2-byte sprmTDefTable cb, whose
// length must be right to reach the shading sprm that follows the table definition.
var wtRows = docToText.model(fs.readFileSync(path.join(__dirname, '..', 'samples', 'word-shaded-table.doc')))
  .body.filter(function (p) { return p.kind === 'rowEnd'; });
check('real Word table: merged header reads as one full-width cell',
  wtRows.some(function (r) { return JSON.stringify(r.tblw) === '[0,9000]'; }));
check('real Word table: shaded cell fill reads from sprmTDefTableShd (red)',
  wtRows.some(function (r) { return r.tblShd && r.tblShd.indexOf(0x0000FF) !== -1; }));

// 13c) Empty table cells survive. A row "A | (empty) | C" must read back as ONE row of
// three cells (empty middle preserved), not split into two — the reader keys the row
// terminator off sprmPFTtp instead of counting consecutive cell marks, so a blank cell's
// back-to-back 0x07s no longer collapse or break the row.
var ecBody = docToText.model(textToDoc([
  { runs: [{ text: 'A' }], kind: 'cell' }, { runs: [], kind: 'cell' }, { runs: [{ text: 'C' }], kind: 'rowEnd', tblw: [0, 3000, 6000, 9000] }
])).body;
var ecCells = ecBody.filter(function (p) { return p.kind === 'cell' || p.kind === 'rowEnd'; });
check('empty middle cell preserved, row not split',
  ecBody.filter(function (p) { return p.kind === 'rowEnd'; }).length === 1 && ecCells.length === 3);
check('empty cell is empty, neighbours intact',
  ecCells[0].runs[0] && ecCells[0].runs[0].text === 'A' && ecCells[1].runs.length === 0 && ecCells[2].runs[0] && ecCells[2].runs[0].text === 'C');

// 14) Palette colours: a run coloured via the 16-colour palette (sprmCIco) rather
// than an explicit RGB (sprmCCv) still reads back as that colour, in the model and
// the styled HTML. The writer emits sprmCCv, so build a coloured run then
// length-preservingly swap its sprm for sprmCIco(6 = red) + a harmless filler.
var icoSrc = textToDoc([{ runs: [{ text: 'ICO', color: 0x123456 }], kind: 'p' }]);
var icoNeedle = [0x70, 0x68, 0x56, 0x34, 0x12, 0x00];   // sprmCCv + COLORREF 0x123456
var icoAt = -1;
for (var ii = 0; ii + icoNeedle.length <= icoSrc.length && icoAt < 0; ii++) {
  var icoOk = true; for (var jj = 0; jj < icoNeedle.length; jj++) if (icoSrc[ii + jj] !== icoNeedle[jj]) { icoOk = false; break; }
  if (icoOk) icoAt = ii;
}
var icoDoc = Uint8Array.from(icoSrc);
[0x42, 0x2A, 0x06, 0x35, 0x08, 0x00].forEach(function (b, k) { icoDoc[icoAt + k] = b; });   // sprmCIco(red) + sprmCFBold(0)
var icoRun = null;
docToText.model(icoDoc).body.forEach(function (p) { (p.runs || []).forEach(function (r) { if (r.text === 'ICO') icoRun = r; }); });
check('sprmCIco palette colour maps to RGB in the model (red)', icoAt >= 0 && !!icoRun && icoRun.color === 0x0000FF);
check('sprmCIco palette colour shows in the styled HTML', /rgb\(255,\s*0,\s*0\)/.test(docToText.html(icoDoc).body));

// 15) Document properties: title/subject/author written into a \x05SummaryInformation
// property set and read back from the model. (TextMaker's COM confirms it reads them
// too: BuiltInDocumentProperties Title/Author/Subject.)
var propDoc = textToDoc({ body: [{ runs: [{ text: 'X' }], kind: 'p' }], props: { title: 'My Title', author: 'Jane Doe', subject: 'Subj' } });
var pr = docToText.model(propDoc).props;
check('document properties round-trip (SummaryInformation)', !!pr && pr.title === 'My Title' && pr.author === 'Jane Doe' && pr.subject === 'Subj');
check('a doc with no properties carries no .props', docToText.model(textToDoc('plain')).props === undefined);
// The reader also parses a real word processor's property set (the bundled sample
// carries an author), proving it's not just reading back our own writer's bytes.
var realProps = docToText.model(fs.readFileSync(path.join(__dirname, '..', 'samples', 'detailed-sample.doc'))).props;
check('detailed-sample author parses from its real SummaryInformation', !!realProps && typeof realProps.author === 'string' && realProps.author.length > 0);

// 16) Paragraph keep/break flags: keep-with-next, keep-lines-together and page-break-
// before (sprmPFKeepFollow / sprmPFKeep / sprmPFPageBreakBefore) round-trip via pp;
// plain paragraphs stay bare. (Round-trip + opens-clean; invisible to TextMaker's COM.)
var kbDoc = textToDoc([
  { runs: [{ text: 'H' }], kind: 'p', pp: { keepNext: 1 } },
  { runs: [{ text: 'B' }], kind: 'p', pp: { keepLines: 1, pageBreak: 1 } },
  { runs: [{ text: 'P' }], kind: 'p' }
]);
var kb = docToText.model(kbDoc).body;
check('keep-with-next round-trips (sprmPFKeepFollow)', !!kb[0].pp && kb[0].pp.keepNext === 1 && !kb[0].pp.keepLines);
check('keep-together + page-break-before round-trip', !!kb[1].pp && kb[1].pp.keepLines === 1 && kb[1].pp.pageBreak === 1);
check('a plain paragraph carries no keep/break flags', kb[2].pp == null);

// 16b) Tab stops: custom paragraph tab stops (sprmPChgTabsPapx -> PChgTabsPapxOperand)
// round-trip — position (twips), alignment (jc) and leader (tlc, from each TBD byte).
// Structure verified against the [MS-DOC] spec; no local fixture uses explicit tab stops
// (default tab stops aren't stored per paragraph), so this is round-trip + spec coverage.
var tabPP = docToText.model(textToDoc([{ runs: [{ text: 'A\tB\tC' }], kind: 'p', pp: { tabs: [
  { pos: 1440, align: 'left', leader: 'none' },
  { pos: 5040, align: 'decimal', leader: 'dot' },
  { pos: 8640, align: 'right', leader: 'underscore' }
] } }])).body[0].pp;
check('tab stops round-trip (pos / alignment / leader)',
  !!tabPP && Array.isArray(tabPP.tabs) && tabPP.tabs.length === 3 &&
  tabPP.tabs[0].pos === 1440 && tabPP.tabs[0].align === 'left' && tabPP.tabs[0].leader === 'none' &&
  tabPP.tabs[1].pos === 5040 && tabPP.tabs[1].align === 'decimal' && tabPP.tabs[1].leader === 'dot' &&
  tabPP.tabs[2].pos === 8640 && tabPP.tabs[2].align === 'right' && tabPP.tabs[2].leader === 'underscore');
check('a paragraph with no tab stops carries no pp.tabs',
  (docToText.model(textToDoc('plain')).body[0].pp || {}).tabs === undefined);

// 16c) Bookmarks: named CP ranges (SttbfBkmk + Plcfbkf + Plcfbkl) round-trip as
// model.bookmarks = [{ name, start, end }], including nested/overlapping ranges paired
// via FBKF.ibkl. Layout verified against the [MS-DOC] worked example; no local fixture
// has bookmarks, so this is round-trip + spec coverage.
var bks = docToText.model(textToDoc({
  body: [{ runs: [{ text: 'The quick brown fox.' }], kind: 'p' }],
  bookmarks: [{ name: 'fox', start: 16, end: 19 }, { name: 'all', start: 0, end: 20 }, { name: 'quick', start: 4, end: 9 }]
})).bookmarks;
function bget(nm) { for (var bj = 0; bj < (bks || []).length; bj++) if (bks[bj].name === nm) return bks[bj]; return null; }
check('bookmarks round-trip (name + CP range)',
  Array.isArray(bks) && bks.length === 3 &&
  !!bget('fox') && bget('fox').start === 16 && bget('fox').end === 19 &&
  !!bget('quick') && bget('quick').start === 4 && bget('quick').end === 9 &&
  !!bget('all') && bget('all').start === 0 && bget('all').end === 20);
check('bookmarks come back sorted by start (overlap/nesting resolved via ibkl)',
  bks[0].start <= bks[1].start && bks[1].start <= bks[2].start);
check('a doc with no bookmarks carries no model.bookmarks',
  docToText.model(textToDoc('plain')).bookmarks === undefined);

// 16d) Paragraph shading (sprmPShd) + box borders (sprmPBrcTop/Left/Bottom/Right): the
// fill COLORREF and each side's colour / width (pt) / type round-trip through pp. Brc /
// Shd byte layout taken from the [MS-DOC] spec (the cell-shading work already validated
// the Shd path against a Word-saved reference).
var pbPP = docToText.model(textToDoc([{ runs: [{ text: 'box' }], kind: 'p', pp: {
  shd: 0x00FFFF,
  borders: { top: { color: 0x0000FF, width: 1, type: 1 }, bottom: { color: 0x00FF00, width: 1.5, type: 1 } }
} }])).body[0].pp;
check('paragraph shading round-trips (sprmPShd fill COLORREF)', !!pbPP && pbPP.shd === 0x00FFFF);
check('paragraph border sides round-trip (colour + width, only set sides)',
  !!pbPP && !!pbPP.borders &&
  !!pbPP.borders.top && pbPP.borders.top.color === 0x0000FF && pbPP.borders.top.width === 1 &&
  !!pbPP.borders.bottom && pbPP.borders.bottom.color === 0x00FF00 && pbPP.borders.bottom.width === 1.5 &&
  !pbPP.borders.left && !pbPP.borders.right);
check('a plain paragraph carries no shd / borders',
  (function () { var pp = docToText.model(textToDoc('plain')).body[0].pp || {}; return pp.shd === undefined && pp.borders === undefined; })());

// 14b) Character formatting completeness: underline styles (sprmCKul kind -> uStyle),
// small caps (sprmCFSmallCaps) and all caps (sprmCFCaps). The toggle path matches
// bold/italic; the caps opcodes are confirmed present 300x across the fixture corpus.
var chrRuns = docToText.model(textToDoc([{ runs: [
  { text: 'dbl', u: true, uStyle: 'double' }, { text: ' dot', u: true, uStyle: 'dotted' },
  { text: ' wav', u: true, uStyle: 'wavy' }, { text: ' plain', u: true },
  { text: ' sc', smallCaps: true }, { text: ' ac', caps: true }
], kind: 'p' }])).body[0].runs;
function cget(txt) { for (var ci = 0; ci < chrRuns.length; ci++) if (chrRuns[ci].text.trim() === txt) return chrRuns[ci]; return null; }
check('underline styles round-trip (double / dotted / wavy)',
  !!cget('dbl') && cget('dbl').uStyle === 'double' && cget('dot').uStyle === 'dotted' && cget('wav').uStyle === 'wavy');
check('a plain single underline carries no uStyle', !!cget('plain') && cget('plain').u === true && cget('plain').uStyle == null);
check('small caps and all caps round-trip', !!cget('sc') && cget('sc').smallCaps === true && !!cget('ac') && cget('ac').caps === true);
check('styled HTML renders underline style + caps',
  (function () { var h = docToText.html(textToDoc([{ runs: [{ text: 'x', u: true, uStyle: 'double' }, { text: 'y', caps: true }], kind: 'p' }])).body; return /text-decoration-style:double/.test(h) && /text-transform:uppercase/.test(h); })());

// 14c) Hidden text (sprmCFVanish) round-trips as run.hidden (dimmed in the demo, but
// kept in the extracted text — it's content, not chrome).
var hidRun = docToText.model(textToDoc([{ runs: [{ text: 'a' }, { text: 'b', hidden: true }], kind: 'p' }])).body[0].runs;
check('hidden text round-trips (sprmCFVanish -> run.hidden)', hidRun.length === 2 && !hidRun[0].hidden && hidRun[1].hidden === true);
check('hidden text stays in the plain-text output', docToText(textToDoc([{ runs: [{ text: 'keep', hidden: true }], kind: 'p' }])).indexOf('keep') === 0);

// 14d) Plain-text tables keep empty cells as empty columns — the row terminator is found
// by its sprmPFTtp mark, not by counting cell marks, so "A | (empty) | C" is "A\t\tC"
// (an empty middle column) instead of being split across two lines.
check('plain-text empty cell stays an empty column (not a split row)',
  docToText(textToDoc([{ runs: [{ text: 'A' }], kind: 'cell' }, { runs: [], kind: 'cell' }, { runs: [{ text: 'C' }], kind: 'rowEnd', tblw: [0, 3000, 6000, 9000] }])).indexOf('A\t\tC') === 0);

// 14e) Vertical cell merge (TC80 fVertRestart / fVertMerge) round-trips as tblMerge[ci].v
// ('restart' / 'cont'), which the demo renders as a rowspan.
var vmRows = docToText.model(textToDoc([
  { runs: [{ text: 'M' }], kind: 'cell', vmerge: 'restart' }, { runs: [{ text: 'b1' }], kind: 'rowEnd', tblw: [0, 4500, 9000] },
  { runs: [], kind: 'cell', vmerge: 'cont' }, { runs: [{ text: 'b2' }], kind: 'rowEnd', tblw: [0, 4500, 9000] }
])).body.filter(function (p) { return p.kind === 'rowEnd'; });
check('vertical merge round-trips (restart / cont in tblMerge)',
  vmRows.length === 2 && !!vmRows[0].tblMerge && !!vmRows[0].tblMerge[0] && vmRows[0].tblMerge[0].v === 'restart' &&
  !!vmRows[1].tblMerge && !!vmRows[1].tblMerge[0] && vmRows[1].tblMerge[0].v === 'cont');

// 17) Floating-shape positions: the reader exposes each text box's FSPA bounding box
// (page coordinates, twips) as model.shapes, so the demo can place it where the
// document actually puts it rather than at its text anchor.
var shp = docToText.model(fs.readFileSync(path.join(__dirname, '..', 'samples', 'detailed-sample.doc'))).shapes;
check('text-box FSPA position is exposed (model.shapes)',
  Array.isArray(shp) && shp.length === 1 && shp[0].xL > 0 && shp[0].yT > 0 && shp[0].xR > shp[0].xL && shp[0].yB > shp[0].yT);

// 18) Character extras: superscript/subscript (sprmCSs) and highlight (sprmCHighlight)
// round-trip through the model. Highlight is a 16-colour palette index, surfaced in the
// model as a COLORREF the same way text colour is. (The reader is also confirmed to read
// real superscript out of a word-processor-saved doc, not just our own bytes.)
var chx = docToText.model(textToDoc([{ runs: [
  { text: 'E=mc' }, { text: 'sup', va: 'super' }, { text: ' H' }, { text: 'sub', va: 'sub' },
  { text: 'O ' }, { text: 'lit', highlight: 0x00FFFF }   // yellow (COLORREF 0x00BBGGRR)
], kind: 'p' }])).body[0].runs;
function rget(txt) { for (var ci = 0; ci < chx.length; ci++) if (chx[ci].text === txt) return chx[ci]; return null; }
check('superscript round-trips (sprmCSs=1 -> va:super)', !!rget('sup') && rget('sup').va === 'super');
check('subscript round-trips (sprmCSs=2 -> va:sub)', !!rget('sub') && rget('sub').va === 'sub');
check('a plain run carries no super/subscript', !!rget('E=mc') && rget('E=mc').va == null);
check('highlight round-trips as a palette COLORREF (sprmCHighlight)', !!rget('lit') && rget('lit').highlight === 0x00FFFF);
// The styled-HTML output renders both.
var chHtml = docToText.html(textToDoc([{ runs: [
  { text: 'a' }, { text: 'b', va: 'super' }, { text: ' ' }, { text: 'c', highlight: 0x00FFFF }
], kind: 'p' }])).body;
check('styled HTML renders superscript (vertical-align)', /vertical-align:super/.test(chHtml));
check('styled HTML renders highlight (background-color)', /background-color:/.test(chHtml));

// 11b) List-marker synthesis. Word keeps a list item's marker ("1." / "a)" / "•")
// in the list definition, not the text stream, so the reader regenerates it.
// The number-format + counting logic is unit-tested directly (the bullet-only
// skeleton can't write a numbered list); bullet markers are checked end-to-end.
var L = docToText._lists;
check('fmtNum decimal', L.fmtNum(4, 0) === '4');
check('fmtNum upper roman', L.fmtNum(4, 1) === 'IV');
check('fmtNum lower roman', L.fmtNum(9, 2) === 'ix');
check('fmtNum upper letter', L.fmtNum(1, 3) === 'A' && L.fmtNum(27, 3) === 'AA');
check('fmtNum lower letter', L.fmtNum(2, 4) === 'b');
// A two-level decimal list: level 0 = "N.", level 1 = "N.M".
var num = L.makeNumberer({ 1: [{ nfc: 0, startAt: 1, tmpl: [0, 46] }, { nfc: 0, startAt: 1, tmpl: [0, 46, 1] }] });
var seq = [[1, 0], [1, 0], [1, 1], [1, 1], [1, 0], [1, 1]].map(function (s) { return num(s[0], s[1]); });
check('numberer counts + restarts deeper levels (1. 2. 2.1 2.2 3. 3.1)', seq.join(' ') === '1. 2. 2.1 2.2 3. 3.1');
var alpha = L.makeNumberer({ 5: [{ nfc: 4, startAt: 1, tmpl: [0, 41] }] });   // "a)" lower-letter
check('numberer formats a lettered list (a) b) c))', [alpha(5, 0), alpha(5, 0), alpha(5, 0)].join(' ') === 'a) b) c)');
check('numberer ignores an unknown ilfo', L.makeNumberer({})(7, 0) === '');
// End-to-end: the writer's bullet list comes back with a • marker in both text and model.
var listDoc = textToDoc([
  { runs: [{ text: 'Shopping' }], kind: 'p' },
  { runs: [{ text: 'Milk' }], kind: 'p', list: { kind: 'bullet', ilvl: 0 } },
  { runs: [{ text: 'Bread' }], kind: 'p', list: { kind: 'bullet', ilvl: 0 } }
]);
check('bullet marker appears in plain text', /(^|\n)• Milk(\n|$)/.test(docToText(listDoc)));
var listModel = docToText.model(listDoc).body.filter(function (p) { return p.list; });
check('bullet marker on the model paragraphs', listModel.length === 2 && listModel.every(function (p) { return p.list.marker === '•'; }));
// The writer synthesizes the list tables (PlfLst/PlfLfo), so a numbered list now
// round-trips end-to-end with counted decimal markers.
var numDoc = textToDoc([
  { runs: [{ text: 'A' }], kind: 'p', list: { kind: 'number', ilvl: 0 } },
  { runs: [{ text: 'B' }], kind: 'p', list: { kind: 'number', ilvl: 0 } },
  { runs: [{ text: 'C' }], kind: 'p', list: { kind: 'number', ilvl: 0 } }
]);
check('numbered list round-trips with 1. 2. 3.', /(^|\n)1\. A\n2\. B\n3\. C(\n|$)/.test(docToText(numDoc)));
check('numbered markers on the model', docToText.model(numDoc).body.filter(function (p) { return p.list; }).map(function (p) { return p.list.marker; }).join(' ') === '1. 2. 3.');

// 11c) Advanced character formatting: double strikethrough (sprmCFDStrike), character
// spacing (sprmCDxaSpace, expanded/condensed) and position (sprmCHpsPos, raised/lowered),
// including negative values.
var advDoc = textToDoc([{ runs: [
  { text: 'ds', dstrike: true }, { text: 'ex', spacing: 2 }, { text: 'co', spacing: -0.7 },
  { text: 'up', position: 3 }, { text: 'dn', position: -3 }
], kind: 'p' }]);
var advRuns = docToText.model(advDoc).body[0].runs;
function findRun(t) { return advRuns.find(function (r) { return r.text === t; }) || {}; }
check('double strikethrough round-trips', findRun('ds').dstrike === true);
check('expanded spacing round-trips', findRun('ex').spacing === 2);
check('condensed (negative) spacing round-trips', findRun('co').spacing === -0.7);
check('raised position round-trips', findRun('up').position === 3);
check('lowered (negative) position round-trips', findRun('dn').position === -3);
var advHtml = docToText.html(advDoc).body;
check('styled HTML: double-strike style', /text-decoration:line-through;text-decoration-style:double/.test(advHtml));
check('styled HTML: letter-spacing', /letter-spacing:2\.0pt/.test(advHtml));
check('styled HTML: raised via vertical-align pt', /vertical-align:3\.0pt/.test(advHtml));

// 12) Independent oracle: word-extractor must still parse the styled .doc AND read
// the footnote + header + endnote we wrote (proves those PLCs are structurally
// valid, not orphaned text the body parser happens to skip).
(function () {
  var WordExtractor;
  try { WordExtractor = require('word-extractor'); }
  catch (e) { console.log('\n  skip word-extractor cross-check (not installed)'); return done(); }
  var we = new WordExtractor();
  we.extract(Buffer.from(doc2)).then(function (d) {
    check('word-extractor parses the styled .doc', d.getBody().indexOf('Lorem') !== -1);
    return we.extract(Buffer.from(ftnDoc));
  }).then(function (d2) {
    check('word-extractor reads the written footnote (getFootnotes)', (d2.getFootnotes() || '').indexOf('A footnote.') !== -1);
    return we.extract(Buffer.from(hfDoc));
  }).then(function (d3) {
    check('word-extractor reads the written header (getHeaders)', (d3.getHeaders() || '').indexOf('The Header') !== -1);
    return we.extract(Buffer.from(combo));
  }).then(function (d4) {
    check('word-extractor reads the written endnote (getEndnotes)', (d4.getEndnotes() || '').indexOf('en') !== -1);
    check('word-extractor reads the written comment (getAnnotations)', (d4.getAnnotations() || '').indexOf('cm') !== -1);
    done();
  }).catch(function (e) { check('word-extractor cross-check (' + e.message + ')', false); done(); });
})();

function done() { console.log(failures === 0 ? '\nALL PASSED' : '\n' + failures + ' FAILURE(S)'); process.exit(failures === 0 ? 0 : 1); }
