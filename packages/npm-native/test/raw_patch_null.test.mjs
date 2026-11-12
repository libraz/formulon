import assert from 'node:assert/strict';
import { existsSync } from 'node:fs';
import { createRequire } from 'node:module';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';
import { FormulonError } from '../dist/common.mjs';

// The null rule of a raw patch: an omitted key keeps, `null` clears only a key the patch type declares clearable,
// and `null` on any other key is a TypeError at the boundary.

const testDir = path.dirname(fileURLToPath(import.meta.url));
const addonPath =
  process.env.FORMULON_NODE_ADDON ??
  [
    path.resolve(testDir, '..', 'dist', 'prebuilds', `${process.platform}-${process.arch}`, 'formulon.node'),
    path.resolve(testDir, '..', 'dist', 'formulon.node'),
    path.resolve(testDir, '..', 'formulon.node'),
  ].find((candidate) => existsSync(candidate));
assert.ok(addonPath, 'native addon not found; set FORMULON_NODE_ADDON or stage dist/formulon.node');
const addon = createRequire(import.meta.url)(addonPath);
addon.__setErrorClass(FormulonError);

const RANGE = { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 };

const HEADER_FOOTER_TEXTS = ['oddHeader', 'oddFooter', 'evenHeader', 'evenFooter', 'firstHeader', 'firstFooter'];
const SPARKLINE_COLORS = [
  'colorSeries',
  'colorNegative',
  'colorAxis',
  'colorMarkers',
  'colorFirst',
  'colorLast',
  'colorHigh',
  'colorLow',
];

const FAMILIES = {
  'Sheet.updateView': {
    keys: [
      'zoomPct',
      'freezeRows',
      'freezeCols',
      'showGridLines',
      'showRowColumnHeaders',
      'showZeros',
      'rightToLeft',
      'tabSelected',
      'viewMode',
      'topLeftRow',
      'topLeftCol',
      'paneTopLeftRow',
      'paneTopLeftCol',
      'splitXTwips',
      'splitYTwips',
      'showFormulas',
      'showOutlineSymbols',
    ],
    clearable: [],
    apply: (wb, patch) => wb.updateView(0, patch),
  },
  'Sheet.updatePageSetup': {
    keys: ['orientation', 'paperSize', 'scalePct', 'fitToWidth', 'fitToHeight', 'fitToPage'],
    clearable: ['orientation'],
    apply: (wb, patch) => wb.updatePageSetup(0, patch),
  },
  'Sheet.updatePageMargins': {
    keys: ['leftIn', 'rightIn', 'topIn', 'bottomIn', 'headerIn', 'footerIn'],
    clearable: [],
    apply: (wb, patch) => wb.updatePageMargins(0, patch),
  },
  'Sheet.updatePrintOptions': {
    keys: ['gridLines', 'headings', 'horizontalCentered', 'verticalCentered'],
    clearable: [],
    apply: (wb, patch) => wb.updatePrintOptions(0, patch),
  },
  'Sheet.updateHeaderFooter': {
    keys: [...HEADER_FOOTER_TEXTS, 'differentOddEven', 'differentFirst', 'scaleWithDoc', 'alignWithMargins'],
    clearable: HEADER_FOOTER_TEXTS,
    apply: (wb, patch) => wb.updateHeaderFooter(0, patch),
  },
  'Sheet.updateFormatDefaults': {
    keys: ['defaultColumnWidthChars', 'defaultRowHeightPt', 'baseColumnWidthChars', 'customHeight', 'zeroHeight'],
    clearable: ['defaultColumnWidthChars', 'defaultRowHeightPt'],
    apply: (wb, patch) => wb.updateFormatDefaults(0, patch),
  },
  'Shape.update': {
    keys: [
      'name',
      'description',
      'hidden',
      'anchor',
      'presetGeometry',
      'text',
      'textBox',
      'fill',
      'line',
      'rotationDeg',
    ],
    // `description` takes `null` on every drawing object.
    clearable: ['description', 'fill', 'line'],
    setup: (wb) => wb.addShape(0, { presetGeometry: 'rect', anchor: { fromRow: 1, fromCol: 1, toRow: 4, toCol: 4 } }),
    apply: (wb, patch, id) => wb.updateShape(0, id, patch),
  },
  'Sheet.updateProtectedRange': {
    keys: [
      'name',
      'spinCount',
      'ranges',
      'legacyPasswordHash',
      'algorithmName',
      'hashValue',
      'saltValue',
      'securityDescriptor',
    ],
    clearable: ['legacyPasswordHash', 'algorithmName', 'hashValue', 'saltValue', 'securityDescriptor'],
    // A clearable key is given a value first, and reads back absent after `null`.
    seed: {
      legacyPasswordHash: 'CC1A',
      algorithmName: 'SHA-512',
      hashValue: 'aA==',
      saltValue: 'cw==',
      securityDescriptor: 'O:WDG:WD',
    },
    setup: (wb) => {
      wb.addProtectedRange(0, { name: 'Totals', ranges: [{ firstRow: 1, firstCol: 1, lastRow: 2, lastCol: 2 }] });
      return 'Totals';
    },
    apply: (wb, patch, name) => wb.updateProtectedRange(0, name, patch),
    read: (wb, name) => wb.getProtectedRange(0, name),
  },
  'Sheet.updateSparklineGroup': {
    keys: [
      ...SPARKLINE_COLORS,
      'manualMin',
      'manualMax',
      'lineWidthPt',
      'type',
      'markers',
      'high',
      'low',
      'first',
      'last',
      'negative',
      'displayXAxis',
      'displayHidden',
      'rightToLeft',
      'minAxisType',
      'maxAxisType',
      'dateAxis',
      'emptyCellsAs',
      'dateAxisRef',
      'sparklines',
    ],
    clearable: [...SPARKLINE_COLORS, 'manualMin', 'manualMax', 'dateAxisRef'],
    setup: (wb) => wb.addSparklineGroup(0, { sparklines: [{ formula: 'Sheet1!A1:D1', location: 'F1' }] }),
    apply: (wb, patch, id) => wb.updateSparklineGroup(0, id, patch),
  },
};

function withWorkbook(fn) {
  const wb = addon.Workbook.createDefault();
  try {
    fn(wb);
  } finally {
    wb.dispose();
  }
}

for (const [where, family] of Object.entries(FAMILIES)) {
  for (const key of family.keys) {
    const clearable = family.clearable.includes(key);
    test(`${where}: null on ${key} ${clearable ? 'is accepted' : 'is a TypeError'}`, () => {
      withWorkbook((wb) => {
        const handle = family.setup?.(wb);
        const run = () => family.apply(wb, { [key]: null }, handle);
        if (clearable && family.seed?.[key] !== undefined) {
          family.apply(wb, { [key]: family.seed[key] }, handle);
          assert.notEqual(family.read(wb, handle)[key] ?? null, null, `${key} was not stored`);
          run();
          assert.equal(family.read(wb, handle)[key] ?? null, null, `null did not clear ${key}`);
        } else if (clearable) {
          try {
            run();
          } catch (e) {
            assert.ok(!(e instanceof TypeError), `null on the clearable ${key} threw ${e}`);
          }
        } else {
          assert.throws(run, (e) => e instanceof TypeError, `null on ${key} must be a TypeError`);
        }
      });
    });
  }
}

// An element of a range array is labelled with its array key and index.
function badElement(call, label) {
  withWorkbook((wb) => {
    assert.throws(
      () => call(wb),
      (e) => e instanceof TypeError && e.message.includes(`${label}[1]`) && !e.message.includes('ranges[]'),
      `a bad element at index 1 must be labelled ${label}[1]`,
    );
  });
}

test('a bad range element is labelled with the validation ranges key and its index', () => {
  badElement((wb) => wb.addValidation(0, { type: 1, ranges: [RANGE, 5] }), 'validation.ranges');
});

test('a bad range element is labelled with the print area ranges key and its index', () => {
  badElement((wb) => wb.setPrintArea(0, [RANGE, 5]), 'ranges');
});

test('a bad range element is labelled with the selection ranges key and its index', () => {
  badElement((wb) => wb.setSelection(0, { active: { row: 0, col: 0 }, ranges: [RANGE, 5] }), 'selection.ranges');
});
