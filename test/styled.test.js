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

// 2) Round-trip the license sample's styled model; tables must come back as tables.
var lic = docToText.model(fs.readFileSync(path.join(__dirname, '..', 'samples', 'license-comparison.doc'))).body;
var doc2 = textToDoc(lic);
var html2 = docToText.html(doc2).body;
check('license: bold preserved', /font-weight:bold/.test(html2));
check('license: colour preserved', /color:rgb/.test(html2));
var re = docToText.model(doc2).body, kinds = {};
re.forEach(function (p) { kinds[p.kind] = (kinds[p.kind] || 0) + 1; });
check('license: table cells round-trip (real 0x07 cells, not tabs)', kinds.cell > 0 && kinds.rowEnd > 0);

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

// 5) Independent oracle: word-extractor must still parse the styled .doc.
(function () {
  var WordExtractor;
  try { WordExtractor = require('word-extractor'); }
  catch (e) { console.log('\n  skip word-extractor cross-check (not installed)'); return done(); }
  new WordExtractor().extract(Buffer.from(doc2)).then(function (d) {
    check('word-extractor parses the styled .doc', d.getBody().indexOf('GPL') !== -1);
    done();
  }).catch(function (e) { check('word-extractor parses the styled .doc (' + e.message + ')', false); done(); });
})();

function done() { console.log(failures === 0 ? '\nALL PASSED' : '\n' + failures + ' FAILURE(S)'); process.exit(failures === 0 ? 0 : 1); }
