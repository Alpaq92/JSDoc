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
// (sprmSDyaTop/Bottom, sprmSDxaLeft/Right, sprmSXaPage/SYaPage). Twips.
var pgDoc = textToDoc({ body: [{ runs: [{ text: 'P.' }], kind: 'p' }], page: { top: 1440, bottom: 1440, left: 1800, right: 1800, width: 12240, height: 15840 } });
var pgm = docToText.model(pgDoc).page;
// All six values differ from the skeleton's A4/1417 defaults, so this fails unless
// the writer actually applied input.page.
check('page setup round-trips (margins + page size)', !!pgm && pgm.top === 1440 && pgm.bottom === 1440 && pgm.left === 1800 && pgm.right === 1800 && pgm.width === 12240 && pgm.height === 15840);

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
