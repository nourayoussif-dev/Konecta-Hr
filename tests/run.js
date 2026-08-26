'use strict';
/**
 * Zero-dependency test runner. `node tests/run.js`
 *
 * Two kinds of check:
 *   test()   — an assertion. Failing means a regression; exit code 1.
 *   review() — documents behaviour we believe is WRONG but have not changed,
 *              because changing payroll maths needs a human decision.
 *              Reported loudly, never fails the build.
 */
const results = { pass: 0, fail: 0, review: [] };
let CURRENT = '';

function describe(name, fn) { CURRENT = name; console.log(`\n\x1b[1m${name}\x1b[0m`); fn(); }

function test(name, fn) {
  try {
    fn();
    results.pass++;
    console.log(`  \x1b[32m✓\x1b[0m ${name}`);
  } catch (e) {
    results.fail++;
    console.log(`  \x1b[31m✗ ${name}\x1b[0m`);
    console.log(`    \x1b[31m${e.message}\x1b[0m`);
  }
}

function review(name, detail) {
  results.review.push({ group: CURRENT, name, detail });
  console.log(`  \x1b[33m⚠\x1b[0m ${name}`);
}

function eq(actual, expected, msg) {
  if (actual !== expected) {
    throw new Error(`${msg || 'expected'}: got ${JSON.stringify(actual)}, want ${JSON.stringify(expected)}`);
  }
}

function close(actual, expected, tolerance, msg) {
  if (Math.abs(actual - expected) > tolerance) {
    throw new Error(`${msg || 'expected'}: got ${actual}, want ${expected} (±${tolerance})`);
  }
}

function throws(fn, msg) {
  let threw = false;
  try { fn(); } catch (e) { threw = true; }
  if (!threw) throw new Error(msg || 'expected a throw, got none');
}

global.describe = describe;
global.test = test;
global.review = review;
global.eq = eq;
global.close = close;
global.throws = throws;

require('./payrollengine.test.js');
require('./coreaccess.test.js');

console.log('\n' + '─'.repeat(64));
console.log(`\x1b[32m${results.pass} passed\x1b[0m` +
            (results.fail ? `, \x1b[31m${results.fail} failed\x1b[0m` : '') +
            (results.review.length ? `, \x1b[33m${results.review.length} flagged for review\x1b[0m` : ''));

if (results.review.length) {
  console.log('\n\x1b[1m\x1b[33mFLAGGED FOR REVIEW — behaviour that looks wrong, left unchanged\x1b[0m');
  console.log('These need a decision from the payroll owner before any code changes.\n');
  results.review.forEach((r, i) => {
    console.log(`\x1b[33m${i + 1}. ${r.name}\x1b[0m  \x1b[2m(${r.group})\x1b[0m`);
    console.log(r.detail.split('\n').map(l => '   ' + l).join('\n') + '\n');
  });
}

process.exit(results.fail ? 1 : 0);
