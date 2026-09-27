// The marshalled bytes of every fidelity part, hashed, for one build of this package.
//
// Used by RELEASING.md's post-publish check: run it against the working tree and against the
// tarball npm actually served, and compare. A pre-release test run says the code was good; only
// this says the artifact is the code that was tested. Point it at a build with OBJECTS_TS_DIST.
import { readFileSync, readdirSync, statSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { fileURLToPath, pathToFileURL } from 'node:url';
import path from 'node:path';

const { unmarshalString, marshalString } = await import(
  process.env.OBJECTS_TS_DIST ? pathToFileURL(process.env.OBJECTS_TS_DIST).href : '../../dist/index.mjs');

const root = fileURLToPath(new URL('../fixtures/fidelity/', import.meta.url));
const parts = [];
const walk = (dir) => {
  for (const entry of readdirSync(dir)) {
    const full = path.join(dir, entry);
    if (statSync(full).isDirectory()) walk(full);
    else if (entry !== 'SOURCE.md') parts.push(full);
  }
};
walk(root);
parts.sort();

const read = (file) => {
  const bytes = readFileSync(file);
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes);
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes);
  return bytes.toString('utf8').replace(/^﻿/, '');
};

let marshalled = 0;
for (const file of parts) {
  const name = path.relative(root, file).split(path.sep).join('/');
  let out;
  try { out = await marshalString(await unmarshalString(read(file))); }
  catch (error) { console.log(`${createHash('sha256').update(`threw:${error.message.split('\n')[0]}`).digest('hex').slice(0, 16)}  ${name}`); continue; }
  marshalled++;
  console.log(`${createHash('sha256').update(out).digest('hex').slice(0, 16)}  ${name}`);
}
console.error(`${parts.length} parts, ${marshalled} marshalled`);
