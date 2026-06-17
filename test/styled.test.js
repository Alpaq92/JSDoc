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
  body: [{ runs: [{ text: 'A' }, { ftnRef: 0 }, { text: ' B' }, { endRef: 0 }, { text: ' C.' }], kind: 'p' }],
  footnotes: [{ runs: [{ text: ' fn' }], kind: 'p' }],
  endnotes: [{ runs: [{ text: ' en' }], kind: 'p' }],
  header: [{ runs: [{ text: 'hd' }], kind: 'p' }],
  footer: [{ runs: [{ text: 'ft' }], kind: 'p' }]
});
var cm = docToText.model(combo);
function hasRef(body, key, v) { return body.some(function (p) { return (p.runs || []).some(function (r) { return r[key] === v; }); }); }
check('endnote anchor round-trips, distinct from footnote', hasRef(cm.body, 'endRef', 0) && hasRef(cm.body, 'ftnRef', 0));
check('endnote text lands in the endnote story', (docToText.sections(combo).endnotes || '').indexOf('en') !== -1);
check('footnote + endnote + header + footer coexist, body clean',
  flatText(cm.footnotes).indexOf('fn') !== -1 && flatText(cm.endnotes).indexOf('en') !== -1 &&
  flatText(cm.header).indexOf('hd') !== -1 && flatText(cm.footer).indexOf('ft') !== -1 &&
  docToText(combo).replace(/\s+/g, ' ').indexOf('A B C.') !== -1);

// 10) Independent oracle: word-extractor must still parse the styled .doc AND read
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
    done();
  }).catch(function (e) { check('word-extractor cross-check (' + e.message + ')', false); done(); });
})();

function done() { console.log(failures === 0 ? '\nALL PASSED' : '\n' + failures + ' FAILURE(S)'); process.exit(failures === 0 ? 0 : 1); }
