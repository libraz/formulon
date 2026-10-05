import test from 'node:test';
import { assert, getModule } from './smoke_support.mjs';

// ---- Sheet-annotation surfaces -----------------------------------------
//
// Conditional formats, data validations, merges, comments and hyperlinks
// are declared in dist/formulon.d.ts and reach the engine through the same
// staged bundle as the calc core, so a staging or binding fault in any of
// them is invisible to the cases above.

// `CfValueObjectInput.type`: 0 num, 1 percent, 2 percentile, 3 min,
// 4 max, 5 formula, 6 autoMin, 7 autoMax.
const COLOR_SCALE_RULE = Object.freeze({
  sqref: [{ firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 0 }],
  type: 2,
  colorScale: {
    thresholds: [
      { type: 0, value: '0' },
      { type: 1, value: '50' },
      { type: 0, value: '100' },
    ],
    colors: [
      { r: 248, g: 105, b: 107, a: 255 },
      { r: 255, g: 235, b: 132, a: 255 },
      { r: 99, g: 190, b: 123, a: 255 },
    ],
  },
});

const DATA_BAR_RULE = Object.freeze({
  sqref: [{ firstRow: 0, firstCol: 1, lastRow: 4, lastCol: 1 }],
  type: 3,
  dataBar: {
    min: { type: 0, value: '5' },
    max: { type: 0, value: '95' },
    fill: { r: 0, g: 112, b: 192, a: 255 },
  },
});

const ICON_SET_RULE = Object.freeze({
  sqref: [{ firstRow: 0, firstCol: 2, lastRow: 4, lastCol: 2 }],
  type: 4,
  iconSet: {
    name: 0,
    thresholds: [
      { type: 1, value: '0' },
      { type: 1, value: '33' },
      { type: 1, value: '67' },
    ],
  },
});

// The threshold value strings, in rule order, as one flat list. Each one
// is handed to the C ABI as a borrowed `const char*`, so a store that
// relocates while the later thresholds are pulled publishes a pointer to
// freed bytes and the strings come back wrong (or empty).
function cfvoValues(rules) {
  const byType = new Map(rules.map((rule) => [rule.type, rule]));
  const colorScale = byType.get(2);
  const dataBar = byType.get(3);
  const iconSet = byType.get(4);
  assert.ok(colorScale?.colorScale, 'colorScale payload missing');
  assert.ok(dataBar?.dataBar, 'dataBar payload missing');
  assert.ok(iconSet?.iconSet, 'iconSet payload missing');
  return [
    ...colorScale.colorScale.thresholds.map((t) => t.value),
    dataBar.dataBar.min.value,
    dataBar.dataBar.max.value,
    ...iconSet.iconSet.thresholds.map((t) => t.value),
  ];
}

const EXPECTED_CFVO_VALUES = Object.freeze(['0', '50', '100', '5', '95', '0', '33', '67']);

test('CFVO value strings survive getConditionalFormats and a save/load cycle', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  let bytes;
  try {
    assert.equal(wb.getConditionalFormats(0).length, 0);
    for (const rule of [COLOR_SCALE_RULE, DATA_BAR_RULE, ICON_SET_RULE]) {
      const added = wb.addConditionalFormat(0, rule);
      assert.ok(added.status.ok, `addConditionalFormat type=${rule.type}: ${JSON.stringify(added.status)}`);
    }

    const inSession = wb.getConditionalFormats(0);
    assert.ok(inSession.status.ok, JSON.stringify(inSession.status));
    assert.equal(inSession.length, 3);
    assert.deepEqual(cfvoValues(inSession), EXPECTED_CFVO_VALUES);
    // The threshold types travel with the values; a type read back as 0
    // would make a percent threshold mean an absolute number.
    const byType = new Map(inSession.map((rule) => [rule.type, rule]));
    assert.deepEqual(
      byType.get(2).colorScale.thresholds.map((t) => t.type),
      [0, 1, 0],
    );
    assert.deepEqual(
      byType.get(4).iconSet.thresholds.map((t) => t.type),
      [1, 1, 1],
    );
    assert.equal(byType.get(2).colorScale.colors.length, 3);
    assert.equal(byType.get(2).colorScale.colors[2].g, 190);
    assert.equal(byType.get(3).dataBar.fill.b, 192);

    const saved = wb.save();
    assert.ok(saved.status.ok, JSON.stringify(saved.status));
    bytes = saved.bytes;
  } finally {
    wb.delete();
  }

  const reloaded = Module.Workbook.loadBytes(bytes);
  try {
    const rules = reloaded.getConditionalFormats(0);
    assert.equal(rules.length, 3);
    assert.deepEqual(cfvoValues(rules), EXPECTED_CFVO_VALUES);
  } finally {
    reloaded.delete();
  }
});

test('addValidation defaults every omitted boolean to false', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    // A list rule stating only what it needs: the booleans are left out
    // so the binding's own defaults are what is read back.
    assert.ok(
      wb.addValidation(0, {
        ranges: [{ firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 0 }],
        type: 3,
        formula1: '"Yes,No"',
      }).ok,
    );
    const rules = wb.getValidations(0);
    assert.ok(rules.status.ok, JSON.stringify(rules.status));
    assert.equal(rules.length, 1);
    const rule = rules[0];
    assert.equal(typeof rule.allowBlank, 'boolean');
    assert.equal(rule.allowBlank, false);
    assert.equal(rule.showInputMessage, false);
    assert.equal(rule.showErrorMessage, false);
    // `showDropDown` is the documented exception: absent means the
    // in-cell arrow is shown, which is what Excel does.
    assert.equal(rule.showDropDown, true);
    assert.equal(rule.type, 3);
    assert.equal(rule.op, 0);
    assert.equal(rule.errorStyle, 0);
    assert.equal(rule.formula1, '"Yes,No"');
    assert.equal(rule.formula2, '');
    assert.equal(rule.ranges.length, 1);
    assert.equal(rule.ranges[0].lastRow, 2);

    // An explicit `true` is not overwritten by the default.
    assert.ok(wb.addValidation(0, { type: 3, allowBlank: true, formula1: '"A,B"' }).ok);
    assert.equal(wb.getValidations(0)[1].allowBlank, true);
  } finally {
    wb.delete();
  }
});

test('merges, comments and hyperlinks round-trip through save and load', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  let bytes;
  try {
    assert.ok(wb.setText(0, 0, 0, '見出し').ok);
    assert.ok(wb.addMerge(0, { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 3 }).ok);
    assert.ok(wb.setComment(0, 2, 1, '校閲者', '要確認').ok);
    assert.ok(wb.addHyperlink(0, 3, 0, 'https://example.com/', '例', 'ツールチップ', '').ok);

    const merges = wb.getMerges(0);
    assert.ok(merges.status.ok, JSON.stringify(merges.status));
    assert.equal(merges.length, 1);
    assert.equal(merges[0].lastCol, 3);

    const comment = wb.getCommentResult(0, 2, 1);
    assert.ok(comment.status.ok, JSON.stringify(comment.status));
    assert.equal(comment.comment.author, '校閲者');
    assert.equal(comment.comment.text, '要確認');
    // An absent comment and an unreadable sheet both answer with a null
    // comment, and are told apart by the status they carry.
    const none = wb.getCommentResult(0, 9, 9);
    assert.equal(none.status.ok, false);
    assert.equal(none.comment, null);
    const badSheet = wb.getCommentResult(99, 0, 0);
    assert.equal(badSheet.status.ok, false);
    assert.equal(badSheet.comment, null);
    assert.notEqual(none.status.status, badSheet.status.status);

    const saved = wb.save();
    assert.ok(saved.status.ok, JSON.stringify(saved.status));
    bytes = saved.bytes;
  } finally {
    wb.delete();
  }

  const reloaded = Module.Workbook.loadBytes(bytes);
  try {
    assert.equal(reloaded.getMerges(0).length, 1);
    assert.equal(reloaded.getMerges(0)[0].lastCol, 3);

    const comments = reloaded.getComments(0);
    assert.ok(comments.status.ok, JSON.stringify(comments.status));
    assert.equal(comments.length, 1);
    assert.equal(comments[0].row, 2);
    assert.equal(comments[0].col, 1);
    assert.equal(comments[0].text, '要確認');

    const links = reloaded.getHyperlinks(0);
    assert.ok(links.status.ok, JSON.stringify(links.status));
    assert.equal(links.length, 1);
    assert.equal(links[0].row, 3);
    assert.equal(links[0].target, 'https://example.com/');
    assert.equal(links[0].tooltip, 'ツールチップ');
  } finally {
    reloaded.delete();
  }
});
