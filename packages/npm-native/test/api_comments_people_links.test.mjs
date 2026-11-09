// Notes, threaded comments, people and hyperlinks through the public entry.

import assert from 'node:assert/strict';
import test from 'node:test';

import { getModule } from './smoke_support.mjs';

async function withWorkbook(fn) {
  const f = await getModule();
  const wb = f.Workbook.createDefault();
  try {
    return await fn(wb, f);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

const GUID = /^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/;

test('addPerson without id returns a generated GUID that the list carries', async () => {
  await withWorkbook((wb) => {
    const id = wb.addPerson({ displayName: 'Alice' });
    assert.match(id, GUID);
    const other = wb.addPerson({ displayName: 'Bob' });
    assert.match(other, GUID);
    assert.notEqual(id, other);
    assert.deepEqual(
      wb.listPeople().map((p) => p.id),
      [id, other],
    );
  });
});

test('absent userId and providerId read null', async () => {
  await withWorkbook((wb) => {
    wb.addPerson({ displayName: 'Alice' });
    const [person] = wb.listPeople();
    assert.equal(person.userId, null);
    assert.equal(person.providerId, null);
  });
});

test('setComment null removes the note and is a no-op when absent', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    ws.setComment('B2', { text: 'hello', author: 'me' });
    assert.equal(ws.getComment('B2').text, 'hello');
    ws.setComment('B2', null);
    assert.equal(ws.getComment('B2'), null);
    ws.setComment('B2', null);
    ws.setComment(5, 5, null);
    assert.equal(ws.getComment(5, 5), null);
    assert.equal(ws.removeComment, undefined);
    assert.equal(ws.removeHyperlink, undefined);
  });
});

test('setCellPhoneticProperties null resets to the unstated triple', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    ws.setText('A1', '漢字');
    ws.setCellPhoneticRuns('A1', [{ start: 0, end: 2, text: 'かんじ' }]);
    ws.setCellPhoneticProperties('A1', { fontIndex: 0, type: 0, alignment: 0 });
    assert.equal(ws.getCellPhoneticProperties('A1').stated, true);
    ws.setCellPhoneticProperties('A1', null);
    assert.equal(ws.getCellPhoneticProperties('A1').stated, false);
  });
});

test('listComments elements carry cell', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    ws.setComment('C4', { text: 'note', author: 'me' });
    const [comment] = ws.listComments();
    assert.deepEqual({ ...comment.cell }, { row: 3, col: 2 });
    assert.equal(comment.text, 'note');
    assert.equal('row' in comment, false);
    assert.equal('col' in comment, false);
  });
});

test('setHyperlink keeps a read link extent', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    ws.setHyperlink('A1', { target: 'https://example.com/', range: 'A1:B2' });
    const read = ws.getHyperlink('A1');
    assert.deepEqual({ ...read.range }, { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 });
    ws.setHyperlink('A1', read);
    assert.deepEqual({ ...ws.getHyperlink('A1').range }, { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 });
    ws.setHyperlink('D4', { target: 'https://example.com/' });
    assert.deepEqual({ ...ws.getHyperlink('D4').range }, { firstRow: 3, firstCol: 3, lastRow: 3, lastCol: 3 });
    assert.equal(ws.setHyperlinkRange, undefined);
    ws.setHyperlink('A1', null);
    assert.equal(ws.getHyperlink('A1'), null);
  });
});

test('setHyperlink range not at the key cell is InvalidArgument', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    assert.throws(
      () => ws.setHyperlink('A1', { target: 'https://example.com/', range: 'B2:C3' }),
      (e) => e.name === 'FormulonError' && e.code === f.StatusCode.InvalidArgument,
    );
    assert.equal(ws.getHyperlink('A1'), null);
  });
});

function threadWithMention(wb, ws, mention) {
  const person = wb.addPerson({ displayName: 'Bob' });
  const id = ws.addThreadedComment({
    cell: 'A1',
    personId: person,
    text: 'hello @Bob now',
    created: { year: 2026, month: 1, day: 2, hour: 3, minute: 4, second: 5 },
    mentions: [{ personId: person, ...mention }],
  });
  return { person, id };
}

test('text.slice(m.start, m.end) is the mention', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    const { id } = threadWithMention(wb, ws, {
      mentionId: '{6F1E0D2C-1111-4111-8111-0000000000AA}',
      start: 6,
      end: 10,
    });
    const comment = ws.getThreadedComment(id);
    const [m] = comment.mentions;
    assert.equal(m.start, 6);
    assert.equal(m.end, 10);
    assert.equal(comment.text.slice(m.start, m.end), '@Bob');
    assert.equal('length' in m, false);
  });
});

test('a mention without mentionId is generated', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    const { id } = threadWithMention(wb, ws, { start: 6, end: 10 });
    const [m] = ws.getThreadedComment(id).mentions;
    assert.match(m.mentionId, GUID);
  });
});

test('a read mention feeds back into a threaded comment patch', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    const { id } = threadWithMention(wb, ws, { start: 6, end: 10 });
    const read = ws.getThreadedComment(id).mentions;
    ws.updateThreadedComment(id, { mentions: read });
    assert.deepEqual(ws.getThreadedComment(id).mentions, read);
  });
});
