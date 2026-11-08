import test from 'node:test';
import { assert, expectFailure, getModule } from './smoke_support.mjs';

// ---- Sheet-annotation surfaces -----------------------------------------
//
// Conditional formats, data validations, merges, comments and hyperlinks
// are declared in dist/formulon.d.ts and reach the engine through the same
// staged bundle as the calc core, so a staging or binding fault in any of
// them is invisible to the cases above.

// `CfValueObjectInput.type`: 0 num, 1 percent, 2 percentile, 3 min,
// 4 max, 5 formula, 6 autoMin, 7 autoMax.
const COLOR_SCALE_RULE = Object.freeze({
  ranges: [{ firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 0 }],
  type: 2,
  colorScale: {
    thresholds: [
      { type: 0, value: '0' },
      { type: 1, value: '50' },
      { type: 0, value: '100' },
    ],
    colors: ['#F8696B', '#FFEB84', '#63BE7B'],
  },
});

const DATA_BAR_RULE = Object.freeze({
  ranges: [{ firstRow: 0, firstCol: 1, lastRow: 4, lastCol: 1 }],
  type: 3,
  dataBar: {
    min: { type: 0, value: '5' },
    max: { type: 0, value: '95' },
    fill: '#0070C0',
  },
});

const ICON_SET_RULE = Object.freeze({
  ranges: [{ firstRow: 0, firstCol: 2, lastRow: 4, lastCol: 2 }],
  type: 4,
  iconSet: {
    // 4Arrows: a 4-icon set carries three thresholds.
    name: 8,
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

test('CFVO value strings survive getCfRule and a save/load cycle', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  let bytes;
  try {
    assert.equal(wb.sheet(0).listCfRules().length, 0);
    for (const rule of [COLOR_SCALE_RULE, DATA_BAR_RULE, ICON_SET_RULE]) {
      wb.sheet(0).addCfRule(rule);
    }

    const inSession = wb.sheet(0).listCfRules();
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
    assert.equal(byType.get(2).colorScale.colors[2].argb, 0xff63be7b);
    assert.equal(byType.get(3).dataBar.fill.argb, 0xff0070c0);

    const saved = wb.save();
    bytes = saved;
  } finally {
    wb.dispose();
  }

  const reloaded = Module.Workbook.load(bytes);
  try {
    const rules = reloaded.sheet(0).listCfRules();
    assert.equal(rules.length, 3);
    assert.deepEqual(cfvoValues(rules), EXPECTED_CFVO_VALUES);
  } finally {
    reloaded.dispose();
  }
});

test('addValidation defaults every omitted boolean to true and keeps an explicit false', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    // A list rule stating only what it needs: the booleans are left out so Excel's dialog
    // defaults are what is read back.
    wb.sheet(0).addValidation({
      ranges: [{ firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 0 }],
      type: 3,
      formula1: '"Yes,No"',
    });
    const rules = wb.sheet(0).listValidations();
    assert.equal(rules.length, 1);
    const rule = rules[0];
    assert.equal(typeof rule.allowBlank, 'boolean');
    assert.equal(rule.allowBlank, true);
    assert.equal(rule.showInputMessage, true);
    assert.equal(rule.showErrorMessage, true);
    assert.equal(rule.showInCellDropdown, true);
    assert.equal(rule.type, 3);
    assert.equal(rule.operator, 0);
    assert.equal(rule.errorStyle, 0);
    assert.equal(rule.formula1, '="Yes,No"');
    assert.equal(rule.formula2, '');
    assert.equal(rule.ranges.length, 1);
    assert.equal(rule.ranges[0].lastRow, 2);

    // An explicit `false` is not overwritten by the default.
    wb.sheet(0).addValidation({
      ranges: [{ firstRow: 3, firstCol: 0, lastRow: 3, lastCol: 0 }],
      type: 3,
      allowBlank: false,
      showInputMessage: false,
      showErrorMessage: false,
      showInCellDropdown: false,
      formula1: '"A,B"',
    });
    const explicit = wb.sheet(0).listValidations()[1];
    assert.equal(explicit.allowBlank, false);
    assert.equal(explicit.showInputMessage, false);
    assert.equal(explicit.showErrorMessage, false);
    assert.equal(explicit.showInCellDropdown, false);
  } finally {
    wb.dispose();
  }
});

test('setHyperlink upserts one link per cell and a null link removes it', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).setHyperlink(1, 1, { target: 'https://example.com/a', display: 'A' });
  wb.sheet(0).setHyperlink(1, 1, { target: 'https://example.com/b' });
  const links = wb.sheet(0).listHyperlinks();
  assert.equal(links.length, 1);
  assert.equal(links[0].target, 'https://example.com/b');
  assert.equal(links[0].display, null);

  wb.sheet(0).setHyperlinkRange(
    { firstRow: 3, firstCol: 0, lastRow: 4, lastCol: 1 },
    { target: 'https://example.com/range' },
  );
  assert.equal(wb.sheet(0).listHyperlinks().length, 2);
  wb.sheet(0).setHyperlinkRange({ firstRow: 3, firstCol: 0, lastRow: 4, lastCol: 1 }, null);
  assert.equal(wb.sheet(0).listHyperlinks().length, 1);

  wb.sheet(0).setHyperlink(1, 1, null);
  assert.equal(wb.sheet(0).getHyperlink(1, 1), null);
  // Removing a link that is not there is a no-op.
  wb.sheet(0).setHyperlink(1, 1, null);
  assert.equal(wb.sheet(0).listHyperlinks().length, 0);
});

test('merges, comments and hyperlinks round-trip through save and load', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  let bytes;
  try {
    wb.sheet(0).setText(0, 0, '見出し');
    wb.sheet(0).addMerge({ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 3 });
    wb.sheet(0).setComment(2, 1, { author: '校閲者', text: '要確認' });
    wb.sheet(0).setHyperlink(3, 0, { target: 'https://example.com/', display: '例', tooltip: 'ツールチップ' });

    const merges = wb.sheet(0).listMerges();
    assert.equal(merges.length, 1);
    assert.equal(merges[0].lastCol, 3);

    const comment = wb.sheet(0).getComment(2, 1);
    assert.equal(comment.author, '校閲者');
    assert.equal(comment.text, '要確認');
    // An absent comment answers null; an unreadable sheet throws.
    const none = wb.sheet(0).getComment(9, 9);
    assert.equal(none, null);
    expectFailure(() => wb.sheet(99).getComment(0, 0), Module.StatusCode.SheetIndexOutOfRange);

    const saved = wb.save();
    bytes = saved;
  } finally {
    wb.dispose();
  }

  const reloaded = Module.Workbook.load(bytes);
  try {
    assert.equal(reloaded.sheet(0).listMerges().length, 1);
    assert.equal(reloaded.sheet(0).listMerges()[0].lastCol, 3);

    const comments = reloaded.sheet(0).listComments();
    assert.equal(comments.length, 1);
    assert.equal(comments[0].row, 2);
    assert.equal(comments[0].col, 1);
    assert.equal(comments[0].text, '要確認');

    const links = reloaded.sheet(0).listHyperlinks();
    assert.equal(links.length, 1);
    assert.equal(links[0].rect.firstRow, 3);
    assert.equal(links[0].target, 'https://example.com/');
    assert.equal(links[0].tooltip, 'ツールチップ');
  } finally {
    reloaded.dispose();
  }
});
