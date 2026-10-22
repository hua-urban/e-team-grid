const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '../backend/Code.gs'), 'utf8');

function makeHarness() {
  const sheet = { reads: 0, writes: [], values: [['id', 'name', 'updated_at', 'last_editor', 'data_json']] };
  let lockCalls = 0;
  const context = {
    console,
    Date,
    JSON,
    Object,
    Array,
    String,
    Number,
    isFinite,
    Utilities: {
      newBlob(value) { return { getBytes: () => Buffer.from(value, 'utf8') }; },
      formatDate() { return '2026-10-22 12:00'; }
    },
    ContentService: {
      MimeType: { JSON: 'application/json' },
      createTextOutput(text) { return { text, setMimeType() { return this; }, getContent() { return this.text; } }; }
    },
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        sheet.reads++;
        return { getSheetByName() { sheet.reads++; return sheet; } };
      }
    },
    LockService: {
      getScriptLock() {
        lockCalls++;
        return { tryLock: () => true, releaseLock() {} };
      }
    },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    Logger: { log() {} }
  };
  sheet.getDataRange = () => {
    sheet.reads++;
    return { getValues: () => sheet.values };
  };
  sheet.getRange = (row, col) => {
    sheet.reads++;
    const range = {
      setNumberFormat() { return this; },
      setValue(value) { sheet.writes.push({ row, col, value }); return this; }
    };
    return range;
  };
  vm.createContext(context);
  vm.runInContext(source, context, { filename: 'backend/Code.gs' });
  return { call: raw => context.doPost({ postData: { contents: raw } }).getContent(), sheet, getLockCalls: () => lockCalls };
}

function body(overrides = {}) {
  return JSON.stringify({ id: '003', changes: [], data: { self: { name: '測試夥伴', en: '', role: '', area: '' }, themes: [] }, ...overrides });
}

test('rejects oversized JSON before reading Sheets or acquiring a lock', () => {
  const h = makeHarness();
  const result = JSON.parse(h.call(' '.repeat(64 * 1024 + 1)));
  assert.equal(result.ok, false);
  assert.equal(h.sheet.reads, 0);
  assert.equal(h.getLockCalls(), 0);
});

test('rejects null, non-object, invalid ID, invalid data, and oversized nested values before Sheets', () => {
  const invalid = [
    'null',
    '[]',
    body({ id: '../003' }),
    body({ data: null }),
    body({ data: [] }),
    body({ data: { value: null } }),
    body({ changes: Array.from({ length: 129 }, () => ({})) }),
    body({ data: { text: 'x'.repeat(4097) } })
  ];
  for (const raw of invalid) {
    const h = makeHarness();
    assert.equal(JSON.parse(h.call(raw)).ok, false, raw.slice(0, 80));
    assert.equal(h.sheet.reads, 0);
    assert.equal(h.getLockCalls(), 0);
  }
});

test('keeps public writes for safe dynamic IDs and creates a row', () => {
  const h = makeHarness();
  const raw = body({
    id: 'new_member-2026',
    name: '新增夥伴',
    data: { self: { name: '新增夥伴', en: '', role: '顧問', area: '' }, themes: [] }
  });
  const result = JSON.parse(h.call(raw));
  assert.deepEqual(result, { ok: true });
  assert.ok(h.sheet.reads > 0);
  assert.equal(h.getLockCalls(), 1);
  assert.deepEqual(h.sheet.writes.map(w => w.value), ['new_member-2026', '新增夥伴', JSON.stringify({ self: { name: '新增夥伴', en: '', role: '顧問', area: '' }, themes: [] }), '2026-10-22 12:00']);
});
