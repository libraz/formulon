// Shape flips, picture rotation and flips, and rotation stored as given, on the public entry.
// FORMULON_NPM_ENTRY selects the staged subpath: "." (default) or "./threads".

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { describe, it } from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const pkgRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const entry = process.env.FORMULON_NPM_ENTRY ?? '.';

async function loadPackage() {
  const pkg = JSON.parse(await readFile(path.join(pkgRoot, 'package.json'), 'utf8'));
  const target = pkg.exports?.[entry]?.import;
  if (!target) throw new Error(`package.json "exports" has no import target for ${entry}`);
  return import(pathToFileURL(path.resolve(pkgRoot, target)).href);
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

describe('shape adjust values, text direction and connections', () => {
  const callout = (ws, extra = {}) => ws.addShape({ anchor: box, presetGeometry: 'wedgeRectCallout', ...extra });

  it('adjustValues reads by guide name and is empty without any', () => {
    withSheet((ws) => {
      assert.deepEqual(ws.addShape({ anchor: box }).getProperties().adjustValues, {});
      const shape = callout(ws, { adjustValues: { adj1: -20833, adj2: 62500 } });
      assert.deepEqual(shape.getProperties().adjustValues, { adj1: -20833, adj2: 62500 });
    });
  });

  it('a patch merges by name: a number sets, null removes, a guide not named is kept', () => {
    withSheet((ws) => {
      const shape = callout(ws, { adjustValues: { adj1: 1000, adj2: 2000 } });
      shape.update({ adjustValues: { adj1: 5000, adj2: null } });
      assert.deepEqual(shape.getProperties().adjustValues, { adj1: 5000 });
      shape.update({ adjustValues: { adj3: 16667 } });
      assert.deepEqual(shape.getProperties().adjustValues, { adj1: 5000, adj3: 16667 });
      shape.update({ adjustValues: {} });
      assert.deepEqual(shape.getProperties().adjustValues, { adj1: 5000, adj3: 16667 });
    });
  });

  it('adjustValues survive a save and reload in both formats', () => {
    withSheet((ws, wb) => {
      const shape = callout(ws, { adjustValues: { adj1: 1234 } });
      for (const format of formats()) {
        const loaded = f.Workbook.load(wb.save(format));
        try {
          const props = loaded.sheet(0).shape(shape.objectId).getProperties();
          assert.deepEqual(props.adjustValues, { adj1: 1234 }, `format ${format}`);
        } finally {
          loaded.dispose();
        }
      }
    });
  });

  it('null, a non-number value, a malformed name and more than 64 guides are refused', () => {
    withSheet((ws) => {
      const shape = callout(ws);
      expectTypeError(() => callout(ws, { adjustValues: null }));
      expectTypeError(() => shape.update({ adjustValues: null }));
      expectTypeError(() => callout(ws, { adjustValues: { adj1: null } }));
      expectTypeError(() => callout(ws, { adjustValues: { adj1: 'wide' } }));
      expectCode(() => callout(ws, { adjustValues: { '': 1 } }), f.StatusCode.InvalidArgument);
      expectCode(() => shape.update({ adjustValues: { '1adj': 1 } }), f.StatusCode.InvalidArgument);
      const many = Object.fromEntries(Array.from({ length: 65 }, (_, i) => [`g${i}`, i]));
      expectCode(() => callout(ws, { adjustValues: many }), f.StatusCode.SecurityResourceLimit);
      expectCode(() => shape.update({ adjustValues: many }), f.StatusCode.SecurityResourceLimit);
      assert.deepEqual(shape.getProperties().adjustValues, {});
    });
  });

  it('textVert is null until set, applies on add and update, and null clears it', () => {
    withSheet((ws) => {
      assert.equal(ws.addShape({ anchor: box }).getProperties().textVert, null);
      const shape = ws.addShape({ anchor: box, textVert: f.ShapeTextVert.Vert270 });
      assert.equal(shape.getProperties().textVert, f.ShapeTextVert.Vert270);
      shape.update({ textVert: f.ShapeTextVert.EaVert });
      assert.equal(shape.getProperties().textVert, f.ShapeTextVert.EaVert);
      shape.update({ textVert: f.ShapeTextVert.Horz });
      assert.equal(shape.getProperties().textVert, f.ShapeTextVert.Horz);
      shape.update({ textVert: null });
      assert.equal(shape.getProperties().textVert, null);
    });
  });

  it('textVert on add is a TypeError for null and InvalidArgument out of range or on a connector', () => {
    withSheet((ws) => {
      expectTypeError(() => ws.addShape({ anchor: box, textVert: null }));
      expectCode(() => ws.addShape({ anchor: box, textVert: 7 }), f.StatusCode.InvalidArgument);
      const line = { anchor: box, presetGeometry: 'straightConnector1' };
      expectCode(() => ws.addShape({ ...line, textVert: f.ShapeTextVert.Vert }), f.StatusCode.InvalidArgument);
      const connector = ws.addShape(line);
      expectCode(() => connector.update({ textVert: f.ShapeTextVert.Vert }), f.StatusCode.InvalidArgument);
      assert.equal(connector.getProperties().textVert, null);
    });
  });

  it('textVert survives a save and reload in both formats', () => {
    withSheet((ws, wb) => {
      const shape = ws.addShape({ anchor: box, textVert: f.ShapeTextVert.WordArtVertRtl });
      for (const format of formats()) {
        const loaded = f.Workbook.load(wb.save(format));
        try {
          const props = loaded.sheet(0).shape(shape.objectId).getProperties();
          assert.equal(props.textVert, f.ShapeTextVert.WordArtVertRtl, `format ${format}`);
        } finally {
          loaded.dispose();
        }
      }
    });
  });

  it('a shape and a free connector read no connection, and the read keys are ignored on update', () => {
    withSheet((ws) => {
      const shape = ws.addShape({ anchor: box });
      const connector = ws.addShape({ anchor: box, presetGeometry: 'straightConnector1' });
      for (const object of [shape, connector]) {
        const props = object.getProperties();
        assert.equal(props.startConnection, null);
        assert.equal(props.endConnection, null);
        object.update({ startConnection: null, endConnection: { objectId: 1, siteIndex: 0 } });
      }
    });
  });
});
