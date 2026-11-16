// Shape flips, picture rotation and flips, and rotation stored as given, on the native entry.

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { describe, it } from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const pkgRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

// The staged entry named by the package's own "main" field.
async function loadPackage() {
  const pkg = JSON.parse(await readFile(path.join(pkgRoot, 'package.json'), 'utf8'));
  if (!pkg.main) throw new Error('package.json is missing "main"');
  return import(pathToFileURL(path.resolve(pkgRoot, pkg.main)).href);
}

const f = await (await loadPackage()).createFormulon();

const PNG_1X1 = Uint8Array.from(
  atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='),
  (c) => c.charCodeAt(0),
);

function withSheet(body) {
  const wb = f.Workbook.createDefault();
  try {
    return body(wb.sheet(0), wb);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

const expectTypeError = (fn) => assert.throws(fn, (e) => e instanceof TypeError, 'expected a TypeError');
const expectCode = (fn, code) =>
  assert.throws(fn, (e) => {
    assert.equal(e.name, 'FormulonError', `expected a FormulonError, got ${e}`);
    assert.equal(e.code, code, `code ${e.code}`);
    return true;
  });
const box = { fromRow: 1, fromCol: 1, toRow: 5, toCol: 4 };
const formats = () => [f.WorkbookFormat.Xlsx, f.WorkbookFormat.Xlsb];

describe('shape flips and raw rotation', () => {
  it('flipH and flipV read false until set, and a patch to false clears them', () => {
    withSheet((ws) => {
      const plain = ws.addShape({ anchor: box });
      assert.equal(plain.getProperties().flipH, false);
      assert.equal(plain.getProperties().flipV, false);
      const shape = ws.addShape({ anchor: box, flipH: true });
      assert.equal(shape.getProperties().flipH, true);
      assert.equal(shape.getProperties().flipV, false);
      shape.update({ flipV: true, flipH: false });
      assert.equal(shape.getProperties().flipH, false);
      assert.equal(shape.getProperties().flipV, true);
      shape.update({ flipV: true });
      assert.equal(shape.getProperties().flipV, true);
    });
  });

  it('a connector takes flips too', () => {
    withSheet((ws) => {
      const line = ws.addShape({
        presetGeometry: 'straightConnector1',
        anchor: { kind: f.AnchorKind.OneCell, fromRow: 1, widthEmu: 914400, heightEmu: 457200 },
        flipV: true,
      });
      assert.equal(line.getProperties().kind, f.DrawingObjectKind.Connector);
      assert.equal(line.getProperties().flipV, true);
    });
  });

  it('null or a non-boolean flip is a TypeError on add and update', () => {
    withSheet((ws) => {
      const shape = ws.addShape({ anchor: box });
      for (const key of ['flipH', 'flipV']) {
        expectTypeError(() => ws.addShape({ anchor: box, [key]: null }));
        expectTypeError(() => shape.update({ [key]: null }));
        expectTypeError(() => shape.update({ [key]: 'yes' }));
        expectTypeError(() => shape.update({ [key]: 1 }));
      }
      assert.equal(shape.getProperties().flipH, false);
    });
  });

  it('rotationDeg is stored as given and round-trips through both save formats', () => {
    withSheet((ws, wb) => {
      const shape = ws.addShape({ anchor: box, rotationDeg: 360, flipH: true });
      assert.equal(shape.getProperties().rotationDeg, 360);
      shape.update({ rotationDeg: -90 });
      assert.equal(shape.getProperties().rotationDeg, -90);
      shape.update({ rotationDeg: 400 });
      assert.equal(shape.getProperties().rotationDeg, 400);
      shape.update({ rotationDeg: -30, flipV: true });
      for (const format of formats()) {
        const loaded = f.Workbook.load(wb.save(format));
        try {
          const props = loaded.sheet(0).shape(shape.objectId).getProperties();
          assert.equal(props.rotationDeg, -30, `format ${format}`);
          assert.equal(props.flipH, true, `format ${format}`);
          assert.equal(props.flipV, true, `format ${format}`);
        } finally {
          loaded.dispose();
        }
      }
      shape.update({ rotationDeg: 0 });
      assert.equal(shape.getProperties().rotationDeg, 0);
    });
  });

  it('a rotation that is not finite or does not fit 1/60000 degree in an int32 is refused', () => {
    withSheet((ws) => {
      const shape = ws.addShape({ anchor: box, rotationDeg: 10 });
      for (const bad of [35792, -35792, 1e12]) {
        expectCode(() => shape.update({ rotationDeg: bad }), f.StatusCode.InvalidArgument);
        expectCode(() => ws.addShape({ anchor: box, rotationDeg: bad }), f.StatusCode.InvalidArgument);
      }
      for (const bad of [Number.NaN, Number.POSITIVE_INFINITY]) {
        assert.throws(() => shape.update({ rotationDeg: bad }));
      }
      assert.equal(shape.getProperties().rotationDeg, 10);
      shape.update({ rotationDeg: 35791 });
      assert.equal(shape.getProperties().rotationDeg, 35791);
    });
  });
});

describe('picture rotation and flips', () => {
  it('a new picture is unrotated and unflipped, and its properties keep the image keys', () => {
    withSheet((ws) => {
      const props = ws.addPicture(PNG_1X1).getProperties();
      assert.equal(props.rotationDeg, 0);
      assert.equal(props.flipH, false);
      assert.equal(props.flipV, false);
      assert.equal(props.widthPx, 1);
      assert.deepEqual(props.bytes, PNG_1X1);
    });
  });

  it('update sets the transform, keeps the anchor and applies beside the common keys', () => {
    withSheet((ws) => {
      const picture = ws.addPicture(PNG_1X1, { anchor: { kind: f.AnchorKind.OneCell, fromRow: 2, fromCol: 3 } });
      const before = picture.getDrawingProperties().anchor;
      picture.update({ rotationDeg: -90, flipH: true });
      assert.deepEqual(picture.getDrawingProperties().anchor, before);
      let props = picture.getProperties();
      assert.equal(props.rotationDeg, -90);
      assert.equal(props.flipH, true);
      assert.equal(props.flipV, false);
      picture.update({ rotationDeg: 360, flipH: false, flipV: true, name: 'Turned' });
      props = picture.getProperties();
      assert.equal(props.rotationDeg, 360);
      assert.equal(props.flipH, false);
      assert.equal(props.flipV, true);
      assert.equal(picture.name, 'Turned');
      // A read record passes back; the image keys are ignored.
      picture.update(props);
      assert.equal(picture.getProperties().rotationDeg, 360);
    });
  });

  it('the transform survives a save and reload in both formats', () => {
    withSheet((ws, wb) => {
      const picture = ws.addPicture(PNG_1X1);
      picture.update({ rotationDeg: 45, flipH: true, flipV: true });
      for (const format of formats()) {
        const loaded = f.Workbook.load(wb.save(format));
        try {
          const props = loaded.sheet(0).picture(picture.objectId).getProperties();
          assert.equal(props.rotationDeg, 45, `format ${format}`);
          assert.equal(props.flipH, true, `format ${format}`);
          assert.equal(props.flipV, true, `format ${format}`);
        } finally {
          loaded.dispose();
        }
      }
    });
  });

  it('null, a non-number rotation, a non-boolean flip and an unknown key are a TypeError', () => {
    withSheet((ws) => {
      const picture = ws.addPicture(PNG_1X1);
      expectTypeError(() => picture.update({ rotationDeg: null }));
      expectTypeError(() => picture.update({ rotationDeg: '90' }));
      expectTypeError(() => picture.update({ flipH: null }));
      expectTypeError(() => picture.update({ flipV: 'yes' }));
      expectTypeError(() => picture.update({ presetGeometry: 'rect' }));
      assert.equal(picture.getProperties().rotationDeg, 0);
    });
  });

  it('a rotation that does not fit is InvalidArgument and changes nothing', () => {
    withSheet((ws) => {
      const picture = ws.addPicture(PNG_1X1);
      picture.update({ rotationDeg: 20 });
      for (const bad of [35792, -35792]) {
        expectCode(() => picture.update({ rotationDeg: bad, name: 'Never' }), f.StatusCode.InvalidArgument);
      }
      assert.throws(() => picture.update({ rotationDeg: Number.NaN }));
      assert.equal(picture.getProperties().rotationDeg, 20);
      assert.notEqual(picture.name, 'Never');
    });
  });
});
