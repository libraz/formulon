// Empty-value rules of the conditional-format rule input (api-conventions §5 rule 30): `""` is refused for the id, the
// two formulas and a threshold value, and is a value for the text of a text rule, as Excel stores it.

import assert from 'node:assert/strict';
import test from 'node:test';
import { inflateRawSync } from 'node:zlib';

import { getModule } from './smoke_support.mjs';

async function withBook(fn) {
  const f = await getModule();
  const wb = f.Workbook.createDefault();
  try {
    return await fn(wb, wb.sheet(0), f);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

const invalid = (f) => (e) => e.name === 'FormulonError' && e.code === f.StatusCode.InvalidArgument;

// Reads one stored or deflated entry of a ZIP archive by name; null when absent.
const readZipEntry = (bytes, name) => {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = bytes.length - 22;
  while (view.getUint32(eocd, true) !== 0x06054b50) eocd -= 1;
  let at = view.getUint32(eocd + 16, true);
  for (let n = view.getUint16(eocd + 10, true); n > 0; n -= 1) {
    const method = view.getUint16(at + 10, true);
    const size = view.getUint32(at + 20, true);
    const nameLen = view.getUint16(at + 28, true);
    const extraLen = view.getUint16(at + 30, true);
    const commentLen = view.getUint16(at + 32, true);
    const local = view.getUint32(at + 42, true);
    const entry = Buffer.from(bytes.subarray(at + 46, at + 46 + nameLen)).toString('utf8');
    if (entry === name) {
      const start = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
      const raw = bytes.subarray(start, start + size);
      return Buffer.from(method === 8 ? inflateRawSync(raw) : raw).toString('utf8');
    }
    at += 46 + nameLen + extraLen + commentLen;
  }
  return null;
};

// ---- id, formula1, formula2 ----

test('addCfRule: id "" is refused and adds nothing, omitted generates a GUID', async () => {
  await withBook((_wb, ws, f) => {
    const rule = { type: f.CfRuleType.Expression, ranges: 'A1:A4', formula1: 'TRUE' };
    assert.throws(() => ws.addCfRule({ ...rule, id: '' }), invalid(f));
    assert.equal(ws.listCfRules().length, 0);
    const id = ws.addCfRule(rule);
    assert.match(id, /^\{[0-9A-F-]{36}\}$/);
    assert.equal(ws.listCfRules().length, 1);
  });
});

test('addCfRule: formula1 "" is refused, omitted fills a text rule formula', async () => {
  await withBook((_wb, ws, f) => {
    assert.throws(() => ws.addCfRule({ type: f.CfRuleType.Expression, ranges: 'A1:A4', formula1: '' }), invalid(f));
    assert.throws(
      () =>
        ws.addCfRule({
          type: f.CfRuleType.CellIs,
          operator: f.ComparisonOperator.GreaterThan,
          ranges: 'A1:A4',
          formula1: '',
        }),
      invalid(f),
    );
    const text = { type: f.CfRuleType.ContainsText, ranges: 'A1:A4', text: 'ab' };
    assert.throws(() => ws.addCfRule({ ...text, formula1: '' }), invalid(f));
    assert.equal(ws.listCfRules().length, 0);
    const id = ws.addCfRule(text);
    assert.equal(ws.getCfRule(id).formula1, '=NOT(ISERROR(SEARCH("ab",A1)))');
  });
});

test('addCfRule: formula2 "" is refused, omitted and null are absent', async () => {
  await withBook((_wb, ws, f) => {
    const rule = {
      type: f.CfRuleType.CellIs,
      operator: f.ComparisonOperator.Between,
      ranges: 'A1:A4',
      formula1: '1',
    };
    assert.throws(() => ws.addCfRule({ ...rule, formula2: '' }), invalid(f));
    assert.equal(ws.listCfRules().length, 0);
    const absent = ws.addCfRule({ ...rule, operator: f.ComparisonOperator.GreaterThan });
    const nulled = ws.addCfRule({ ...rule, operator: f.ComparisonOperator.GreaterThan, formula2: null });
    assert.equal(ws.getCfRule(absent).formula2, null);
    assert.equal(ws.getCfRule(nulled).formula2, null);
  });
});

test('updateCfRule: "" formula is refused and changes nothing', async () => {
  await withBook((_wb, ws, f) => {
    const id = ws.addCfRule({
      type: f.CfRuleType.CellIs,
      operator: f.ComparisonOperator.Between,
      ranges: 'A1:A4',
      formula1: '1',
      formula2: '9',
    });
    assert.throws(() => ws.updateCfRule(id, { formula1: '' }), invalid(f));
    assert.throws(() => ws.updateCfRule(id, { formula2: '' }), invalid(f));
    assert.equal(ws.getCfRule(id).formula1, '=1');
    assert.equal(ws.getCfRule(id).formula2, '=9');
  });
});

// ---- threshold value ----

test('addCfRule: a threshold value "" is refused, null reads none', async () => {
  await withBook((_wb, ws, f) => {
    const colorScale = (value) => ({
      type: f.CfRuleType.ColorScale,
      ranges: 'A1:A4',
      colorScale: {
        thresholds: [{ type: f.CfValueObjectType.Min }, { type: f.CfValueObjectType.Number, value }],
        colors: ['#FF0000', '#00FF00'],
      },
    });
    assert.throws(() => ws.addCfRule(colorScale('')), invalid(f));
    assert.throws(
      () =>
        ws.addCfRule({
          type: f.CfRuleType.DataBar,
          ranges: 'A1:A4',
          dataBar: {
            min: { type: f.CfValueObjectType.Min },
            max: { type: f.CfValueObjectType.Number, value: '' },
            fill: '#638EC6',
          },
        }),
      invalid(f),
    );
    assert.equal(ws.listCfRules().length, 0);
    const id = ws.addCfRule(colorScale('10'));
    assert.equal(ws.getCfRule(id).colorScale.thresholds[0].value, null);
  });
});

test('addCfRule and updateCfRule: a colour-scale AutoMin or AutoMax threshold is refused and changes nothing', async () => {
  await withBook((_wb, ws, f) => {
    const scale = (first, second) => ({
      type: f.CfRuleType.ColorScale,
      ranges: 'A1:A4',
      colorScale: { thresholds: [{ type: first }, { type: second }], colors: ['#FF0000', '#00FF00'] },
    });
    const T = f.CfValueObjectType;
    assert.throws(() => ws.addCfRule(scale(T.AutoMin, T.Max)), invalid(f));
    assert.equal(ws.listCfRules().length, 0);
    const id = ws.addCfRule(scale(T.Min, T.Max));
    const patch = {
      colorScale: { thresholds: [{ type: T.Min }, { type: T.AutoMax }], colors: ['#FF0000', '#00FF00'] },
    };
    assert.throws(() => ws.updateCfRule(id, patch), invalid(f));
    assert.deepEqual(
      ws.getCfRule(id).colorScale.thresholds.map((t) => t.type),
      [T.Min, T.Max],
    );
  });
});

// ---- text ----

test('addCfRule: text "" is a value, omitted or null is absent', async () => {
  await withBook((_wb, ws, f) => {
    const rule = { type: f.CfRuleType.ContainsText, ranges: 'A1:A4' };
    const empty = ws.addCfRule({ ...rule, text: '' });
    assert.equal(ws.getCfRule(empty).text, '');
    assert.equal(ws.getCfRule(empty).formula1, '=NOT(ISERROR(SEARCH("",A1)))');
    const absent = ws.addCfRule(rule);
    const nulled = ws.addCfRule({ ...rule, text: null });
    assert.equal(ws.getCfRule(absent).text, null);
    assert.equal(ws.getCfRule(nulled).text, null);
  });
});

test('updateCfRule: text "" is kept', async () => {
  await withBook((_wb, ws, f) => {
    const id = ws.addCfRule({ type: f.CfRuleType.BeginsWith, ranges: 'A1:A4', text: 'ab' });
    ws.updateCfRule(id, { text: '' });
    assert.equal(ws.getCfRule(id).text, '');
  });
});

test('a text "" rule is saved as text="" and round-trips in xlsx and xlsb', async () => {
  await withBook(async (wb, ws, f) => {
    ws.addCfRule({ type: f.CfRuleType.ContainsText, ranges: 'A1:A4', text: '' });
    const xlsx = wb.save();
    const xml = readZipEntry(xlsx, 'xl/worksheets/sheet1.xml');
    assert.match(xml, /type="containsText"/);
    assert.match(xml, /operator="containsText"/);
    assert.match(xml, /text=""/);
    for (const bytes of [xlsx, wb.save(f.WorkbookFormat.Xlsb)]) {
      const reloaded = f.Workbook.load(bytes);
      try {
        const rules = reloaded.sheet(0).listCfRules();
        assert.equal(rules.length, 1);
        assert.equal(rules[0].text, '');
      } finally {
        reloaded.dispose();
      }
    }
  });
});

test('a containsText "" rule matches every cell that is not an error', async () => {
  await withBook((wb, ws, f) => {
    ws.setText('A1', 'abc');
    ws.setNumber('A2', 5);
    ws.setFormula('A4', '=1/0');
    wb.recalc();
    ws.addCfRule({ type: f.CfRuleType.ContainsText, ranges: 'A1:A4', text: '' });
    const rows = () => ws.evaluateCfRange('A1:A4').map(({ row }) => row);
    assert.deepEqual(rows(), [0, 1, 2]);
  });
});
