// The null rule of a patch: an omitted key keeps, `null` clears only a key the patch type declares clearable, and
// `null` on any other writable key is a TypeError at the JS boundary. Read-only keys are ignored whatever their value.
// FORMULON_NPM_ENTRY selects the staged subpath of the WASM package: "." (default) or "./threads".

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { describe, it } from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const pkgRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const entry = process.env.FORMULON_NPM_ENTRY ?? '.';

async function loadPackage() {
  const pkg = JSON.parse(await readFile(path.join(pkgRoot, 'package.json'), 'utf8'));
  const target = pkg.exports?.[entry]?.import ?? pkg.main;
  if (!target) throw new Error('package.json has neither an "exports" import target nor "main"');
  return import(pathToFileURL(path.resolve(pkgRoot, target)).href);
}

const f = await (await loadPackage()).createFormulon();

const headerFooterTexts = ['evenFooter', 'evenHeader', 'firstFooter', 'firstHeader', 'oddFooter', 'oddHeader'];
const sparklineColors = [
  'colorAxis',
  'colorFirst',
  'colorHigh',
  'colorLast',
  'colorLow',
  'colorMarkers',
  'colorNegative',
  'colorSeries',
];

const FAMILIES = {
  'Sheet.updateView': {
    keys: [
      'freezeCols',
      'freezeRows',
      'paneTopLeft',
      'rightToLeft',
      'showFormulas',
      'showGridLines',
      'showOutlineSymbols',
      'showRowColumnHeaders',
      'showZeros',
      'splitXTwips',
      'splitYTwips',
      'tabSelected',
      'topLeft',
      'viewMode',
      'zoomPct',
    ],
    readOnly: [],
    clearable: [],
    apply: (ws, patch) => ws.updateView(patch),
  },
  'Sheet.updatePageSetup': {
    keys: [
      'fitToHeight',
      'fitToHeightStated',
      'fitToPage',
      'fitToPageStated',
      'fitToWidth',
      'fitToWidthStated',
      'orientation',
      'paperSize',
      'paperSizeStated',
      'scalePct',
      'scalePctStated',
    ],
    readOnly: ['fitToHeightStated', 'fitToPageStated', 'fitToWidthStated', 'paperSizeStated', 'scalePctStated'],
    clearable: ['orientation'],
    apply: (ws, patch) => ws.updatePageSetup(patch),
  },
  'Sheet.updatePageMargins': {
    keys: [
      'bottomIn',
      'bottomInStated',
      'footerIn',
      'footerInStated',
      'headerIn',
      'headerInStated',
      'leftIn',
      'leftInStated',
      'rightIn',
      'rightInStated',
      'topIn',
      'topInStated',
    ],
    readOnly: ['bottomInStated', 'footerInStated', 'headerInStated', 'leftInStated', 'rightInStated', 'topInStated'],
    clearable: [],
    apply: (ws, patch) => ws.updatePageMargins(patch),
  },
  'Sheet.updatePrintOptions': {
    keys: [
      'gridLines',
      'gridLinesStated',
      'headings',
      'headingsStated',
      'horizontalCentered',
      'horizontalCenteredStated',
      'verticalCentered',
      'verticalCenteredStated',
    ],
    readOnly: ['gridLinesStated', 'headingsStated', 'horizontalCenteredStated', 'verticalCenteredStated'],
    clearable: [],
    apply: (ws, patch) => ws.updatePrintOptions(patch),
  },
  'Sheet.updateHeaderFooter': {
    keys: [
      'alignWithMargins',
      'alignWithMarginsStated',
      'differentFirst',
      'differentFirstStated',
      'differentOddEven',
      'differentOddEvenStated',
      'evenFooter',
      'evenHeader',
      'firstFooter',
      'firstHeader',
      'oddFooter',
      'oddHeader',
      'scaleWithDoc',
      'scaleWithDocStated',
    ],
    readOnly: ['alignWithMarginsStated', 'differentFirstStated', 'differentOddEvenStated', 'scaleWithDocStated'],
    clearable: headerFooterTexts,
    apply: (ws, patch) => ws.updateHeaderFooter(patch),
  },
  'Sheet.updateFormatDefaults': {
    keys: [
      'baseColumnWidthChars',
      'customHeight',
      'defaultColumnWidthChars',
      'defaultRowHeightPt',
      'effectiveDefaultColumnWidthChars',
      'effectiveDefaultRowHeightPt',
      'zeroHeight',
    ],
    readOnly: ['effectiveDefaultColumnWidthChars', 'effectiveDefaultRowHeightPt'],
    clearable: ['defaultColumnWidthChars', 'defaultRowHeightPt'],
    apply: (ws, patch) => ws.updateFormatDefaults(patch),
  },
  'Shape.update': {
    keys: [
      'anchor',
      'description',
      'fill',
      'hidden',
      'kind',
      'line',
      'name',
      'objectId',
      'presetGeometry',
      'rotationDeg',
      'text',
      'textBox',
    ],
    readOnly: ['kind', 'objectId'],
    // `description` takes `null` on every drawing object.
    clearable: ['description', 'fill', 'line'],
    setup: (ws) =>
      ws.addShape({
        presetGeometry: 'rect',
        anchor: { kind: f.AnchorKind.OneCell, fromRow: 2, widthEmu: 914400, heightEmu: 457200 },
      }),
    apply: (_ws, patch, shape) => shape.update(patch),
  },
  'Sheet.updateProtectedRange': {
    keys: [
      'algorithmName',
      'hashValue',
      'legacyPasswordHash',
      'name',
      'ranges',
      'saltValue',
      'securityDescriptor',
      'spinCount',
    ],
    readOnly: [],
    clearable: ['algorithmName', 'hashValue', 'legacyPasswordHash', 'saltValue', 'securityDescriptor'],
    // A clearable key is given a value first, and reads back absent after `null`.
    seed: {
      algorithmName: 'SHA-512',
      hashValue: 'aA==',
      legacyPasswordHash: 'CC1A',
      saltValue: 'cw==',
      securityDescriptor: 'O:WDG:WD',
    },
    setup: (ws) =>
      ws.addProtectedRange({
        name: 'Totals',
        ranges: [{ firstRow: 1, firstCol: 1, lastRow: 2, lastCol: 2 }],
      }),
    apply: (ws, patch) => ws.updateProtectedRange(0, patch),
    read: (ws) => ws.listProtectedRanges()[0],
  },
  'Sheet.updateSparklineGroup': {
    keys: [
      ...sparklineColors,
      'dateAxis',
      'dateAxisRange',
      'displayHidden',
      'displayXAxis',
      'emptyCellsAs',
      'first',
      'high',
      'id',
      'last',
      'lineWidthPt',
      'low',
      'manualMax',
      'manualMin',
      'markers',
      'maxAxisType',
      'minAxisType',
      'negative',
      'rightToLeft',
      'sparklines',
      'type',
    ],
    readOnly: ['id'],
    clearable: [...sparklineColors, 'dateAxisRange', 'manualMax', 'manualMin'],
    setup: (ws) => ws.addSparklineGroup({ sparklines: [{ source: 'Sheet1!A1:D1', location: 'F1' }] }),
    apply: (ws, patch, id) => ws.updateSparklineGroup(id, patch),
  },
};

/** Runs `body` on a fresh default workbook's first sheet and disposes the workbook. */
function withSheet(body) {
  const wb = f.Workbook.createDefault();
  try {
    return body(wb.sheet(0));
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

for (const [where, family] of Object.entries(FAMILIES)) {
  const writable = family.keys.filter((key) => !family.readOnly.includes(key));
  describe(`${where}: null on a patch key`, () => {
    it('names only clearable keys that are writable keys of the patch', () => {
      for (const key of family.clearable) assert.ok(writable.includes(key), `${key} is not a writable key`);
    });

    for (const key of writable) {
      const clearable = family.clearable.includes(key);
      it(clearable ? `${key} accepts null` : `${key} refuses null with a TypeError`, () => {
        withSheet((ws) => {
          const handle = family.setup?.(ws);
          const run = () => family.apply(ws, { [key]: null }, handle);
          if (clearable && family.seed?.[key] !== undefined) {
            family.apply(ws, { [key]: family.seed[key] }, handle);
            assert.notEqual(family.read(ws, handle)[key] ?? null, null, `${key} was not stored`);
            run();
            assert.equal(family.read(ws, handle)[key] ?? null, null, `null did not clear ${key}`);
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
  });
}
