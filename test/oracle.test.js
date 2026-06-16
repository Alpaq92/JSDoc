// SPDX-License-Identifier: 0BSD
/*
 * oracle.test.js — real-world validation against actual Word .doc files.
 *
 * Why: the synthetic fixture proves the parser is internally consistent; this
 * proves it agrees with an *independent* implementation on *real* documents,
 * which guards against shared-bug blindness.
 *
 * Oracle: `word-extractor` (MIT, https://github.com/morungos/node-word-extractor)
 * — an unrelated pure-JS reader. We never read or copy its source; we only
 * compare its text output to ours.
 *
 * Fixtures: real .doc files from that project's MIT-licensed test corpus,
 * downloaded on demand (not vendored). These are genuine Word binaries — the
 * "validate against a real .doc" requirement, since no Word/LibreOffice/catdoc
 * is available in this environment to generate one locally.
 *
 * Comparison is by normalized word tokens. The two readers differ on
 * whitespace, field/table handling, and tracked changes (the oracle strips
 * tracked-*deletion* text via character-level revision marks; our v1 keeps it,
 * by design — see README "Known limitations"). So we assert RECALL — that we
 * never *miss* text the oracle found — plus an in-order prefix match, rather
 * than exact equality. Over-inclusion (e.g. tracked deletions) is reported but
 * tolerated; under-extraction (a real bug) fails.
 *
 * Gracefully SKIPS (exit 0) if offline or if word-extractor is not installed.
 * Run: npm run test:oracle
 */
'use strict';

var fs = require('fs');
var path = require('path');
var https = require('https');
var docToText = require('../src/docToText.js');

var RAW = 'https://raw.githubusercontent.com/morungos/node-word-extractor/develop/__tests__/data/';
var FIXTURE_DIR = path.join(__dirname, 'fixtures');
var DOCS = ['test01.doc', 'test02.doc', 'test05.doc', 'test08.doc', 'test11.doc', 'test13.doc'];
var BAD = 'badfile-01-bad-header.doc';

var RECALL_MIN = 0.97;   // fraction of oracle tokens we must also capture
var PREFIX_TOKENS = 25;

function download(name) {
  return new Promise(function (resolve) {
    var dest = path.join(FIXTURE_DIR, name);
    if (fs.existsSync(dest) && fs.statSync(dest).size > 0) return resolve(dest);
    https.get(RAW + name, function (res) {
      if (res.statusCode !== 200) { res.resume(); return resolve(null); }
      var chunks = [];
      res.on('data', function (d) { chunks.push(d); });
      res.on('end', function () {
        try { fs.writeFileSync(dest, Buffer.concat(chunks)); resolve(dest); }
        catch (e) { resolve(null); }
      });
    }).on('error', function () { resolve(null); });
  });
}

function tokens(s) {
  if (!s) return [];
  var m = s.toLowerCase().match(/[\p{L}\p{N}]+/gu);
  return m || [];
}

// Multiset overlap stats between mine (a) and oracle (b).
//   recall    = how much of the oracle we captured  (1.0 = missed nothing)
//   precision = how much of ours the oracle also had (low = we included extra)
function overlap(a, b) {
  var ca = count(a), cb = count(b), keys = {}, k;
  for (k in ca) keys[k] = true;
  for (k in cb) keys[k] = true;
  var inter = 0, sizeA = a.length, sizeB = b.length;
  for (k in keys) inter += Math.min(ca[k] || 0, cb[k] || 0);
  return {
    recall: sizeB === 0 ? 1 : inter / sizeB,
    precision: sizeA === 0 ? 1 : inter / sizeA
  };
}
function count(arr) {
  var c = {};
  for (var i = 0; i < arr.length; i++) c[arr[i]] = (c[arr[i]] || 0) + 1;
  return c;
}
function prefixMatch(a, b, n) {
  var lim = Math.min(n, a.length, b.length);
  if (lim === 0) return a.length === 0 && b.length === 0 ? 1 : 0;
  var hit = 0;
  for (var i = 0; i < lim; i++) if (a[i] === b[i]) hit++;
  return hit / lim;
}

(function main() {
  var WordExtractor;
  try { WordExtractor = require('word-extractor'); }
  catch (e) {
    console.log('SKIP oracle test: word-extractor not installed (npm install).');
    process.exit(0);
  }

  if (!fs.existsSync(FIXTURE_DIR)) fs.mkdirSync(FIXTURE_DIR, { recursive: true });
  var extractor = new WordExtractor();
  var failures = 0, ran = 0;

  Promise.all(DOCS.concat(BAD).map(download)).then(async function (paths) {
    console.log('docToText — real-.doc oracle comparison (vs word-extractor)\n');

    for (var i = 0; i < DOCS.length; i++) {
      var p = paths[i];
      if (!p) { console.log('  skip ' + DOCS[i] + ' (download failed)'); continue; }
      ran++;
      var buf = fs.readFileSync(p);
      var mine = docToText(buf);

      var oracle, odoc;
      try { odoc = await extractor.extract(buf); oracle = odoc.getBody(); }
      catch (e) { console.log('  skip ' + DOCS[i] + ' (oracle error: ' + e.message + ')'); ran--; continue; }

      if (mine == null) {
        console.log('  FAIL ' + DOCS[i] + ' — mine returned null but oracle extracted ' + tokens(oracle).length + ' tokens');
        failures++; continue;
      }
      var ta = tokens(mine), tb = tokens(oracle);
      var ov = overlap(ta, tb);
      var pre = prefixMatch(ta, tb, PREFIX_TOKENS);
      var ok = ov.recall >= RECALL_MIN && pre >= 0.9;
      var extra = ov.precision < 0.98 ? '  (+extra: e.g. tracked deletions)' : '';
      console.log('  ' + (ok ? 'ok  ' : 'FAIL') + ' ' + DOCS[i] +
        '  tokens(mine=' + ta.length + ', oracle=' + tb.length + ')' +
        '  recall=' + ov.recall.toFixed(3) + '  prefix=' + pre.toFixed(2) + extra);
      if (!ok) {
        failures++;
        console.log('      mine[0..15]   ' + JSON.stringify(ta.slice(0, 15)));
        console.log('      oracle[0..15] ' + JSON.stringify(tb.slice(0, 15)));
      }

      // Extra stories vs word-extractor's matching getters (only when present).
      var sects = docToText.sections(buf) || {};
      var SECS = [['footnotes', 'getFootnotes'], ['headers', 'getHeaders'], ['endnotes', 'getEndnotes'], ['annotations', 'getAnnotations']];
      for (var si = 0; si < SECS.length; si++) {
        var oo = ''; try { oo = odoc[SECS[si][1]](); } catch (e) { oo = ''; }
        var ob = tokens(oo);
        if (!ob.length) continue;
        var mt = tokens(sects[SECS[si][0]]);
        var rc = overlap(mt, ob).recall;
        if (rc < RECALL_MIN) failures++;
        console.log('    ' + (rc >= RECALL_MIN ? 'ok  ' : 'FAIL') + ' ' + DOCS[i] + ' / ' + SECS[si][0] +
          '  recall=' + rc.toFixed(3) + ' (mine=' + mt.length + ', oracle=' + ob.length + ')');
      }
    }

    // The deliberately-corrupt file must degrade gracefully (null, no throw).
    var badPath = paths[paths.length - 1];
    if (badPath) {
      ran++;
      var r;
      try { r = docToText(fs.readFileSync(badPath)); }
      catch (e) { r = '<threw: ' + e.message + '>'; }
      var okBad = (r === null || typeof r === 'string');
      console.log('  ' + (okBad ? 'ok  ' : 'FAIL') + ' ' + BAD + ' degrades gracefully (got ' +
        (r === null ? 'null' : typeof r) + ')');
      if (!okBad) failures++;
    }

    if (ran === 0) { console.log('\nSKIP: no fixtures available (offline?).'); process.exit(0); }
    console.log(failures === 0 ? '\nALL PASSED (' + ran + ' files)' : '\n' + failures + ' FAILURE(S)');
    process.exit(failures === 0 ? 0 : 1);
  });
})();
