// Conditional-format rules addressed by position, and priority placement, on a file whose two rules share one id.

import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

import { getModule, isFormulonError } from './smoke_support.mjs';

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..', '..');
const fixturePath = path.join(repoRoot, 'tests', 'fixtures', 'excel', 'cf_x14_shared_id.xlsx');
const SHARED_ID = '{00000000-000E-0000-0000-000001000000}';
const INT32_MAX = 2147483647;

// Two data-bar rules sharing one id: A1:A10 at priority 1 and B1:B10 at priority 2.
async function withShared(fn) {
  const f = await getModule();
  const wb = f.Workbook.load(new Uint8Array(readFileSync(fixturePath)));
  try {
    return await fn(wb.sheet(0), f, wb);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

const view = (ws) =>
  ws.listCfRules().map((r) => ({
    id: r.id,
    priority: r.priority,
    col: r.ranges[0].firstCol,
    stopIfTrue: r.stopIfTrue,
  }));
const priorities = (ws) => ws.listCfRules().map((r) => r.priority);
const byCol = (ws, col) => view(ws).find((r) => r.col === col);
const expression = (f, extra = {}) => ({ type: f.CfRuleType.Expression, ranges: 'D1:D10', formula1: 'TRUE', ...extra });
const hasCode = (code) => (e) => isFormulonError(e, code);

test('the fixture lists two rules that share an id', async () => {
  await withShared((ws) => {
    assert.deepEqual(
      view(ws).map((r) => [r.id, r.priority, r.col]),
      [
        [SHARED_ID, 1, 0],
        [SHARED_ID, 2, 1],
      ],
    );
  });
});

test('updateCfRule with a position changes that rule and not the first rule sharing its id', async () => {
  await withShared((ws) => {
    const before = view(ws);
    ws.updateCfRule(1, { ranges: 'F1:F10' });
    const after = view(ws);
    assert.deepEqual(after[0], before[0]);
    assert.equal(after[1].col, 5);
    assert.equal(after[1].id, SHARED_ID);
    assert.equal(after[1].priority, 2);
  });
});

test('updateCfRule with position 0 changes the first rule only', async () => {
  await withShared((ws) => {
    const before = view(ws);
    ws.updateCfRule(0, { ranges: 'G1:G10' });
    const after = view(ws);
    assert.equal(after[0].col, 6);
    assert.deepEqual(after[1], before[1]);
  });
});

test('removeCfRule with position 0 leaves the second rule with its priority and the shared id', async () => {
  await withShared((ws) => {
    ws.removeCfRule(0);
    assert.deepEqual(
      view(ws).map((r) => [r.id, r.priority, r.col]),
      [[SHARED_ID, 2, 1]],
    );
  });
});

test('removeCfRule with position 1 leaves the first rule', async () => {
  await withShared((ws) => {
    ws.removeCfRule(1);
    assert.deepEqual(
      view(ws).map((r) => [r.id, r.priority, r.col]),
      [[SHARED_ID, 1, 0]],
    );
  });
});

test('an id names the first of the rules sharing it', async () => {
  await withShared((ws) => {
    assert.equal(ws.getCfRule(SHARED_ID).ranges[0].firstCol, 0);
    ws.updateCfRule(SHARED_ID, { stopIfTrue: true });
    assert.equal(byCol(ws, 0).stopIfTrue, true);
    assert.equal(byCol(ws, 1).stopIfTrue, false);
    ws.removeCfRule(SHARED_ID);
    assert.deepEqual(
      view(ws).map((r) => r.col),
      [1],
    );
    // Once the first is gone the id reaches the remaining one.
    ws.removeCfRule(SHARED_ID);
    assert.equal(ws.listCfRules().length, 0);
  });
});

test('a position outside the list fails with IndexOutOfRange and changes nothing', async () => {
  await withShared((ws, f) => {
    const before = view(ws);
    const calls = {
      update: (position) => ws.updateCfRule(position, { stopIfTrue: true }),
      remove: (position) => ws.removeCfRule(position),
      priority: (position) => ws.setCfRulePriority(position, 1),
    };
    for (const position of [2, 5, -1]) {
      for (const [name, call] of Object.entries(calls)) {
        assert.throws(() => call(position), hasCode(f.StatusCode.IndexOutOfRange), `${name} ${position}`);
      }
    }
    assert.deepEqual(view(ws), before);
  });
});

test('a fractional position is a RangeError', async () => {
  await withShared((ws) => {
    const before = view(ws);
    assert.throws(() => ws.updateCfRule(1.5, { stopIfTrue: true }), RangeError);
    assert.throws(() => ws.removeCfRule(1.5), RangeError);
    assert.throws(() => ws.setCfRulePriority(1.5, 1), RangeError);
    assert.deepEqual(view(ws), before);
  });
});

test('an explicit duplicate priority is accepted on add and on update', async () => {
  await withShared((ws, f) => {
    const id = ws.addCfRule(expression(f, { priority: 2 }));
    assert.equal(ws.getCfRule(id).priority, 2);
    assert.deepEqual(priorities(ws), [1, 2, 2]);
    ws.updateCfRule(1, { priority: 1 });
    assert.deepEqual(priorities(ws), [1, 1, 2]);
  });
});

test('a listed rule added back with its id cleared keeps its fields', async () => {
  await withShared((ws) => {
    const copy = { ...ws.listCfRules()[0], id: undefined };
    const id = ws.addCfRule(copy);
    assert.notEqual(id, SHARED_ID);
    const added = ws.getCfRule(id);
    assert.equal(added.priority, 1);
    assert.equal(added.ranges[0].firstCol, 0);
    assert.equal(ws.listCfRules().length, 3);
  });
});

test('setCfRulePriority to 1 shifts every other rule up by one', async () => {
  await withShared((ws, f) => {
    const id = ws.addCfRule(expression(f));
    assert.equal(ws.getCfRule(id).priority, 3);
    ws.setCfRulePriority(id, 1);
    assert.equal(ws.getCfRule(id).priority, 1);
    assert.equal(byCol(ws, 0).priority, 2);
    assert.equal(byCol(ws, 1).priority, 3);
  });
});

test('setCfRulePriority keeps tied rules together and in order', async () => {
  await withShared((ws, f) => {
    ws.updateCfRule(1, { priority: 1 });
    const id = ws.addCfRule(expression(f));
    assert.equal(ws.getCfRule(id).priority, 2);
    ws.setCfRulePriority(ws.listCfRules().length - 1, 1);
    assert.equal(ws.getCfRule(id).priority, 1);
    assert.deepEqual(
      view(ws)
        .filter((r) => r.col < 2)
        .map((r) => [r.col, r.priority]),
      [
        [0, 2],
        [1, 2],
      ],
    );
  });
});

test('setCfRulePriority leaves lower numbers and the old slot alone', async () => {
  await withShared((ws, f) => {
    const id = ws.addCfRule(expression(f));
    ws.setCfRulePriority(id, 2);
    assert.equal(byCol(ws, 0).priority, 1);
    assert.equal(ws.getCfRule(id).priority, 2);
    assert.equal(byCol(ws, 1).priority, 3);
    ws.setCfRulePriority(id, 9);
    assert.equal(ws.getCfRule(id).priority, 9);
    assert.equal(byCol(ws, 1).priority, 3);
  });
});

test('setCfRulePriority above the maximum shifts nothing', async () => {
  await withShared((ws) => {
    ws.setCfRulePriority(0, 10);
    assert.equal(byCol(ws, 0).priority, 10);
    assert.equal(byCol(ws, 1).priority, 2);
  });
});

test('setCfRulePriority to the current priority is a no-op', async () => {
  await withShared((ws) => {
    const before = view(ws);
    for (const [i, rule] of before.entries()) {
      ws.setCfRulePriority(i, rule.priority);
      assert.deepEqual(view(ws), before);
    }
  });
});

test('setCfRulePriority with an id acts on the first rule sharing it', async () => {
  await withShared((ws) => {
    ws.setCfRulePriority(SHARED_ID, 5);
    assert.equal(byCol(ws, 0).priority, 5);
    assert.equal(byCol(ws, 1).priority, 2);
  });
});

test('setCfRulePriority below 1 fails with InvalidArgument and changes nothing', async () => {
  await withShared((ws, f) => {
    const before = view(ws);
    for (const bad of [0, -1]) {
      assert.throws(() => ws.setCfRulePriority(1, bad), hasCode(f.StatusCode.InvalidArgument), `priority ${bad}`);
    }
    assert.deepEqual(view(ws), before);
  });
});

test('setCfRulePriority that would overflow fails with InvalidArgument and changes nothing', async () => {
  await withShared((ws, f) => {
    ws.updateCfRule(1, { priority: INT32_MAX });
    const before = view(ws);
    assert.throws(() => ws.setCfRulePriority(0, INT32_MAX), hasCode(f.StatusCode.InvalidArgument));
    assert.throws(() => ws.setCfRulePriority(0, 2), hasCode(f.StatusCode.InvalidArgument));
    assert.deepEqual(view(ws), before);
  });
});

test('priorities set by setCfRulePriority survive a save and reload', async () => {
  await withShared(async (ws, f, wb) => {
    ws.setCfRulePriority(1, 1);
    const saved = wb.save();
    const reloaded = f.Workbook.load(saved);
    try {
      const rules = reloaded.sheet(0).listCfRules();
      assert.deepEqual(
        rules.map((r) => [r.ranges[0].firstCol, r.priority]).sort((a, b) => a[0] - b[0]),
        [
          [0, 2],
          [1, 1],
        ],
      );
    } finally {
      reloaded.dispose();
    }
  });
});

test('an automatic priority is one past the highest on the sheet', async () => {
  await withShared((ws, f) => {
    const id = ws.addCfRule(expression(f));
    assert.equal(ws.getCfRule(id).priority, 3);
  });
});
