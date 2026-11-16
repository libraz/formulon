// A stored sparkline rectangle loads as its top-left cell; a location that is not one range drops its group at
// .xlsx load, counted in the load diagnostics, and a save carries only the groups that read.

import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

import { getModule } from './smoke_support.mjs';
import { rewritePackage } from './zip_rewrite.mjs';

const SHEET = 'xl/worksheets/sheet1.xml';
const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..', '..');
const EXCEL_XLSB = path.join(
  repoRoot,
  'tests',
  'fixtures',
  'excel',
  'xlsb_parity',
  'sparklines',
  'multicell_location.xlsb',
);

// The text of the saved first sheet, read through an identity rewrite.
function sheetText(bytes) {
  let text = '';
  rewritePackage(bytes, {
    [SHEET]: (xml) => {
      text = xml;
      return xml;
    },
  });
  return text;
}

// A saved book with one group per location (each over its own row of A:D), with the given sqref texts rewritten.
async function craftedPackage(f, locations, rewrites) {
  const wb = f.Workbook.createDefault();
  try {
    const ws = wb.sheet(0);
    locations.forEach((location, i) => {
      ws.addSparklineGroup({ sparklines: [{ source: `Sheet1!A${i + 2}:D${i + 2}`, location }] });
    });
    return rewritePackage(wb.save(), {
      [SHEET]: (xml) => rewrites.reduce((text, [from, to]) => text.replace(from, to), xml),
    });
  } finally {
    wb.dispose();
  }
}

const sqref = (text) => `<xm:sqref>${text}</xm:sqref>`;

function locations(wb) {
  return wb
    .sheet(0)
    .listSparklineGroups()
    .map((g) => g.sparklines[0].location);
}

test('a rectangle loads as its top-left cell without a diagnostic', async () => {
  const f = await getModule();
  const wb = f.Workbook.load(await craftedPackage(f, ['F2', 'B2'], [[sqref('B2'), sqref('B3:C4')]]));
  try {
    assert.deepEqual(locations(wb), [
      { row: 1, col: 5 },
      { row: 2, col: 1 },
    ]);
    assert.equal(wb.getDiagnostics().skippedFeatureCount, 0);
  } finally {
    wb.dispose();
  }
});

test('groups whose location is not one range are dropped at load and counted', async () => {
  const f = await getModule();
  const bytes = await craftedPackage(
    f,
    ['F2', 'B2', 'C2', 'D2'],
    [
      [sqref('B2'), sqref('A1:B2')],
      [sqref('C2'), sqref('zz')],
      [sqref('D2'), sqref('A1 C3')],
    ],
  );
  const wb = f.Workbook.load(bytes);
  try {
    assert.deepEqual(locations(wb), [
      { row: 1, col: 5 },
      { row: 0, col: 0 },
    ]);
    assert.equal(wb.getDiagnostics().skippedFeatureCount, 2);
    const reloaded = f.Workbook.load(wb.save());
    try {
      assert.deepEqual(locations(reloaded), [
        { row: 1, col: 5 },
        { row: 0, col: 0 },
      ]);
      assert.equal(reloaded.getDiagnostics().skippedFeatureCount, 0);
    } finally {
      reloaded.dispose();
    }
  } finally {
    wb.dispose();
  }
});

test('a book whose only group is unreadable saves no sparklineGroups block', async () => {
  const f = await getModule();
  const wb = f.Workbook.load(await craftedPackage(f, ['B2'], [[sqref('B2'), sqref('zz')]]));
  try {
    assert.deepEqual(wb.sheet(0).listSparklineGroups(), []);
    assert.equal(wb.getDiagnostics().skippedFeatureCount, 1);
    assert.doesNotMatch(sheetText(wb.save()), /sparklineGroups/);
  } finally {
    wb.dispose();
  }
});

test('an Excel .xlsb rectangle loads as its top-left cell', async () => {
  const f = await getModule();
  const wb = f.Workbook.load(new Uint8Array(readFileSync(EXCEL_XLSB)));
  try {
    assert.deepEqual(locations(wb), [{ row: 0, col: 6 }]);
    assert.equal(wb.getDiagnostics().skippedFeatureCount, 0);
  } finally {
    wb.dispose();
  }
});

test('an .xlsx rectangle is written to .xlsb as its top-left cell', async () => {
  const f = await getModule();
  const wb = f.Workbook.load(await craftedPackage(f, ['B2'], [[sqref('B2'), sqref('A1:B2')]]));
  try {
    const back = f.Workbook.load(wb.save(f.WorkbookFormat.Xlsb));
    try {
      assert.deepEqual(locations(back), [{ row: 0, col: 0 }]);
      assert.equal(back.getDiagnostics().skippedFeatureCount, 0);
    } finally {
      back.dispose();
    }
  } finally {
    wb.dispose();
  }
});
