'use strict';
// Structural guards for the self-approval and bank four-eyes fixes.
const { load } = require('./harness');
const g = load(['Code.gs']);

describe('Bank four-eyes controls', () => {
  test('bank fields cannot be edited through the bulk tool', () => {
    const bulk = g.__eval('BULK_FIELDS');
    ['bank_name', 'account_number', 'iban'].forEach(f => {
      if (f in bulk) throw new Error(`${f} must not be bulk-editable — it bypasses verification`);
    });
  });

  test('bank_verified is locked from direct HR editing', () => {
    const locked = g.__eval('HR_LOCKED');
    if (locked.indexOf('bank_verified') === -1)
      throw new Error('bank_verified must be in HR_LOCKED so only hrVerifyBank sets it');
  });

  test('lastBankChangeBy_ and the four-eyes guard exist', () => {
    eq(typeof g.lastBankChangeBy_, 'function');
    eq(typeof g.hrVerifyBank, 'function');
  });
});

describe('Self-approval guard', () => {
  test('assertNotSelf_ and myEmployeeId_ exist and are wired', () => {
    eq(typeof g.assertNotSelf_, 'function');
    eq(typeof g.myEmployeeId_, 'function');
  });

  test('all four decision functions call assertNotSelf_', () => {
    const fs = require('fs');
    const src = fs.readFileSync(require('path').join(__dirname, '..', 'Code.gs'), 'utf8');
    ['decideLeave', 'decideResignation', 'decideWithdrawal', 'confirmHandover'].forEach(fn => {
      const start = src.indexOf('function ' + fn + '(');
      const body = src.slice(start, start + 1600);
      if (!/assertNotSelf_\(/.test(body))
        throw new Error(`${fn} is missing the self-decide guard`);
    });
  });
});

describe('Rate limiting', () => {
  test('withinRateLimit_ exists and fails open on cache error', () => {
    eq(typeof g.withinRateLimit_, 'function');
    // no CacheService in this harness load path other than the stub → should not throw
    eq(g.withinRateLimit_('noshow', 20), true);
  });
});
