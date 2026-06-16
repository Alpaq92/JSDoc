// SPDX-License-Identifier: 0BSD
/*
 * writer.test.js — validate textToDoc() by round-tripping the .doc it writes.
 *
 * There's no Word in this environment, so "is the output a valid .doc?" is
 * answered by reading it back with two independent parsers: our own docToText
 * (must reproduce the input exactly) and the unrelated word-extractor (must
 * parse it and find the text). Both succeeding is strong evidence the binary
 * is well-formed.
 *
 * The docToText round-trip is offline and always runs; the word-extractor check
 * is best-effort (skips if that dev dependency isn't installed).
 */
'use strict';
var textToDoc = require('../src/textToDoc.js');
var docToText = require('../src/docToText.js');

var failures = 0;
function check(name, cond) {
  console.log('  ' + (cond ? 'ok  ' : 'FAIL') + ' ' + name);
  if (!cond) failures++;
}

console.log('textToDoc — write a .doc and read it back\n');

var cases = [
  'Hello, World!',
  'Line one\nLine two\nLine three',
  'Smart quotes “like this”, an em dash —, and accents éàü.',
  'Unicode: π ∑ ✓ 😀',
  ''
];

cases.forEach(function (input, i) {
  var bytes;
  try { bytes = textToDoc(input); } catch (e) { check('case ' + i + ' writes', false); return; }
  check('case ' + i + ' produces bytes', bytes && bytes.length >= 512 && bytes[0] === 0xD0 && bytes[1] === 0xCF);
  var back = docToText(bytes);
  // textToDoc appends a trailing paragraph mark; compare on normalized lines.
  var want = input.replace(/\r\n?|\n/g, '\n');
  var got = (back || '').replace(/\r\n?|\n/g, '\n').replace(/\n$/, '');
  check('case ' + i + ' round-trips through docToText', got === want);
});

// Independent oracle: word-extractor must also parse our .doc.
(function () {
  var WordExtractor;
  try { WordExtractor = require('word-extractor'); }
  catch (e) { console.log('\n  skip word-extractor cross-check (not installed)'); return done(); }
  var input = 'Independent reader check.\nSecond paragraph.';
  var buf = Buffer.from(textToDoc(input));
  new WordExtractor().extract(buf).then(function (doc) {
    var body = doc.getBody();
    check('word-extractor parses our .doc and finds the text',
      body.indexOf('Independent reader check.') !== -1 && body.indexOf('Second paragraph.') !== -1);
    done();
  }).catch(function (e) {
    check('word-extractor parses our .doc (' + e.message + ')', false);
    done();
  });
})();

function done() {
  console.log(failures === 0 ? '\nALL PASSED' : '\n' + failures + ' FAILURE(S)');
  process.exit(failures === 0 ? 0 : 1);
}
