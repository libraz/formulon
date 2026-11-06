// Reads the member names of an exported interface from a declaration file,
// for the tests that compare a runtime class against its declaration.

import { readFileSync } from 'node:fs';

function interfaceBody(text, name) {
  const start = text.search(new RegExp(`export interface ${name}\\s*\\{`));
  if (start < 0) throw new Error(`no interface ${name}`);
  const open = text.indexOf('{', start);
  let depth = 0;
  for (let i = open; i < text.length; i++) {
    if (text[i] === '{') depth++;
    else if (text[i] === '}' && --depth === 0) return text.slice(open + 1, i);
  }
  throw new Error(`unbalanced interface ${name}`);
}

/** Member names of one interface: methods, properties and `[Symbol.dispose]`. */
export function memberNames(dtsPath, name) {
  const text = readFileSync(dtsPath, 'utf8')
    .replace(/\/\*[\s\S]*?\*\//g, '')
    .replace(/\/\/[^\n]*/g, '');
  const body = interfaceBody(text, name);
  const statements = [];
  let depth = 0;
  let current = '';
  for (const ch of body) {
    if (ch === ';' && depth === 0) {
      statements.push(current);
      current = '';
      continue;
    }
    if ('({'.includes(ch)) depth++;
    else if (')}'.includes(ch)) depth--;
    current += ch;
  }
  const names = new Set();
  for (const statement of statements) {
    const m = statement.trim().match(/^(?:readonly\s+)?(\[Symbol\.dispose\]|[A-Za-z_$][\w$]*)\??\s*[(<:]/);
    if (m) names.add(m[1]);
  }
  return names;
}

/** The own keys of a class prototype as the declarations spell them. */
export function prototypeKeys(proto) {
  return [
    ...Object.getOwnPropertyNames(proto).filter((k) => k !== 'constructor'),
    ...Object.getOwnPropertySymbols(proto).map((s) => (s === Symbol.dispose ? '[Symbol.dispose]' : String(s))),
  ].sort();
}
