'use strict';
/**
 * Loads Apps Script .gs files into a sandbox so their pure logic can be
 * tested on a laptop, with no Google account and no network.
 *
 * Apps Script has no modules: every .gs file shares one global scope. We
 * reproduce that here by evaluating the requested files into a single vm
 * context, then handing back that context as the namespace.
 *
 * Google services are stubbed. Anything that genuinely needs a spreadsheet is
 * out of scope for these tests by design — the point is to pin down the
 * arithmetic, which is where the money is.
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');

function notImplemented(service) {
  return new Proxy({}, {
    get(_, prop) {
      if (prop === Symbol.toPrimitive || typeof prop === 'symbol') return undefined;
      throw new Error(
        `${service}.${String(prop)}() was called. This code path touches Google ` +
        `services and cannot run in the offline harness — extract the pure ` +
        `logic, or stub ${service} explicitly in your test.`
      );
    }
  });
}

/** Minimal stand-ins for the Apps Script globals the engine may touch. */
function makeSandbox(overrides) {
  const TZ = 'Africa/Cairo';
  const sandbox = {
    console,
    // Share the host Date constructor. The vm context otherwise gets its own,
    // and a Date created by a test fails `v instanceof Date` inside the .gs
    // code (fmt_ silently falls through to String(v)).
    Date,
    Logger: { log() {} },

    Session: {
      getScriptTimeZone: () => TZ,
      getActiveUser: () => ({ getEmail: () => 'test@konecta.com' }),
      getEffectiveUser: () => ({ getEmail: () => 'test@konecta.com' })
    },

    Utilities: {
      // Enough of the format pattern for the codebase's actual usage.
      formatDate(date, tz, pattern) {
        const p = n => String(n).padStart(2, '0');
        const map = {
          'yyyy-MM-dd': `${date.getFullYear()}-${p(date.getMonth() + 1)}-${p(date.getDate())}`,
          'yyyy-MM': `${date.getFullYear()}-${p(date.getMonth() + 1)}`
        };
        if (map[pattern]) return map[pattern];
        throw new Error(`harness: unsupported date pattern "${pattern}"`);
      },
      formatString(fmt, ...args) {
        let i = 0;
        return fmt.replace(/%[,\d.]*[sdf]/g, m => {
          const v = args[i++];
          if (/f$/.test(m)) {
            const dec = (m.match(/\.(\d+)f$/) || [, '2'])[1];
            const n = Number(v).toFixed(Number(dec));
            return /,/.test(m) ? Number(n).toLocaleString('en-US', {
              minimumFractionDigits: Number(dec), maximumFractionDigits: Number(dec)
            }) : n;
          }
          return String(v);
        });
      }
    },

    SpreadsheetApp: notImplemented('SpreadsheetApp'),
    DriveApp:       notImplemented('DriveApp'),
    DocumentApp:    notImplemented('DocumentApp'),
    CalendarApp:    notImplemented('CalendarApp'),
    MailApp:        notImplemented('MailApp'),
    HtmlService:    notImplemented('HtmlService'),
    ScriptApp:      notImplemented('ScriptApp'),

    CacheService: {
      getScriptCache: () => ({ get: () => null, put() {}, removeAll() {} })
    },
    LockService: {
      getScriptLock: () => ({ tryLock: () => true, waitLock() {}, releaseLock() {} })
    },
    PropertiesService: {
      getScriptProperties: () => ({ getProperty: () => null, setProperty() {} })
    }
  };

  Object.assign(sandbox, overrides || {});
  return sandbox;
}

/**
 * @param {string[]} files  .gs filenames relative to the repo root
 * @param {object}  [overrides]  extra/replacement sandbox globals
 * @returns {object} the shared global scope after evaluation
 */
function load(files, overrides) {
  const sandbox = makeSandbox(overrides);
  const context = vm.createContext(sandbox);
  for (const f of files) {
    const src = fs.readFileSync(path.join(ROOT, f), 'utf8');
    try {
      vm.runInContext(src, context, { filename: f });
    } catch (e) {
      throw new Error(`harness: failed loading ${f}: ${e.message}`);
    }
  }
  return sandbox;
}

module.exports = { load };
