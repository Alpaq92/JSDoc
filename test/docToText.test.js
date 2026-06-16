// SPDX-License-Identifier: 0BSD
/*
 * Offline, dependency-free tests for docToText().
 *   - Round-trips a synthetic spec-valid .doc with a known expected output,
 *     covering the FAT + mini-FAT paths, a Prc-prefixed CLX, compressed
 *     cp1252 + uncompressed UTF-16LE pieces, fields, and control marks.
 *   - Confirms graceful null on unsupported / malformed input.
 * Run: node test/docToText.test.js   (exit code 0 = all passed)
 */
'use strict';

var docToText = require('../src/docToText.js');
var fixture = require('./make-fixture.js');

var failures = 0;
function check(name, cond) {
  if (cond) { console.log('  ok   ' + name); }
  else { console.log('  FAIL ' + name); failures++; }
}
function show(s) { return JSON.stringify(s); }

console.log('docToText — offline tests');

// 1. Happy path: synthetic .doc round-trip.
var doc = fixture.buildDoc();
var got = docToText(doc.buffer);
check('synthetic .doc extracts expected text', got === fixture.EXPECTED);
if (got !== fixture.EXPECTED) {
  console.log('   expected ' + show(fixture.EXPECTED));
  console.log('   got      ' + show(got));
}

// 1b. Accepts a Uint8Array too (not just ArrayBuffer).
check('accepts Uint8Array input', docToText(doc) === fixture.EXPECTED);

// 1c. Accepts a Node Buffer.
check('accepts Node Buffer input', docToText(Buffer.from(doc)) === fixture.EXPECTED);

// 2. Field codes dropped, results kept.
check('field instruction dropped, result kept',
  got.indexOf('PAGE') === -1 && got.indexOf('1Tab') !== -1);

// 3. Smart quotes decoded via cp1252 table.
check('cp1252 smart quotes decoded', got.indexOf('“World”') !== -1);

// 4. UTF-16-only characters decoded.
check('UTF-16LE non-Latin decoded', got.indexOf('π') !== -1); // pi

// --- graceful degradation (must return null, never throw) ---
function nullCase(name, input) {
  var r;
  try { r = docToText(input); } catch (e) { r = '<threw: ' + e.message + '>'; }
  check(name + ' -> null', r === null);
}

nullCase('null input', null);
nullCase('empty buffer', new Uint8Array(0).buffer);
nullCase('short non-CFB', new Uint8Array(600).buffer); // zeros, bad signature

// Corrupt the CFB signature.
var badSig = fixture.buildDoc(); badSig[0] = 0x00;
nullCase('bad CFB signature', badSig.buffer);

// Encrypted document (fEncrypted bit 8 in the FIB flags).
var enc = fixture.buildDoc();
var encDv = new DataView(enc.buffer);
var fibFlagsOff = 512 + 4 * 512 + 10; // WordDocument sector + FIB flags offset
encDv.setUint16(fibFlagsOff, encDv.getUint16(fibFlagsOff, true) | 0x0100, true);
nullCase('encrypted (fEncrypted)', enc.buffer);

// Word 6/95 (older nFib) -> out of scope.
var old = fixture.buildDoc();
new DataView(old.buffer).setUint16(512 + 4 * 512 + 2, 0x0065, true);
nullCase('Word 6/95 nFib', old.buffer);

console.log(failures === 0
  ? '\nALL PASSED'
  : '\n' + failures + ' FAILURE(S)');
process.exit(failures === 0 ? 0 : 1);
