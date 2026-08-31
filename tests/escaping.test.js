'use strict';
// Server-side escaping helpers. The client esc() is the primary XSS defence;
// these back it up so nothing dangerous is ever stored or rendered in a PDF.
const { load } = require('./harness');
const g = load(['Code.gs']);

describe('Server-side HTML escaping', () => {
  test('escapeHtml_ now covers quotes as well as angle brackets', () => {
    eq(g.escapeHtml_('<img onerror="x" \'y\'>'),
       '&lt;img onerror=&quot;x&quot; &#39;y&#39;&gt;');
    eq(g.escapeHtml_('a & b'), 'a &amp; b');
  });

  test('scrubText_ strips angle brackets so no tag reaches the sheet', () => {
    eq(g.scrubText_('<img src=x onerror=alert(1)>'), 'img src=x onerror=alert(1)');
    eq(g.scrubText_('Ahmed'), 'Ahmed');
    eq(g.scrubText_(null), '');
    eq(g.scrubText_(undefined), '');
  });

  test('a name that was a payload becomes inert once scrubbed', () => {
    const planted = '<script>google.script.run.hrSaveRecord(1,{})</script>';
    const stored = g.scrubText_(planted);
    if (/[<>]/.test(stored)) throw new Error('angle bracket survived scrub');
  });
});
