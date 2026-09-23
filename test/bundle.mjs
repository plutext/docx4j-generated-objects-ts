// CR-005: the package bundles with no bundler configuration, and a bundle carries the 103 mapping
// modules rather than the 309 .mjs files in modules/.
//
// The defect this pins: `import(`../modules/${name}.mjs`)` is a template, so a bundler cannot see
// what it names. Vite leaves it a runtime URL and the browser 404s (the editor's first open, which
// asked for this CR); esbuild expands it to a glob over ../modules/*.mjs, which works but drags in
// the `.el.mjs` and `.factory.mjs` sibling of every module. A registry of literal imports takes
// exactly what loads.
import { mkdtempSync, rmSync, readFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';

const here = fileURLToPath(new URL('.', import.meta.url));
const esbuild = join(here, '../node_modules/.bin/esbuild');
const out = mkdtempSync(join(tmpdir(), 'objects-ts-bundle-'));
try {
  // No flags beyond target and format: configuration is what the CR exists to remove.
  execFileSync(esbuild, [join(here, 'bundle/consumer.mjs'), '--bundle', '--platform=node',
    '--format=esm', `--outfile=${join(out, 'app.mjs')}`], { stdio: 'pipe' });
  const bundle = readFileSync(join(out, 'app.mjs'), 'utf8');

  assert.ok(!bundle.includes('__glob'), 'no glob expansion: the bundler saw literal specifiers');
  assert.ok(!/modules\/org_\w+\.(el|factory)\.mjs/.test(bundle), 'the el and factory siblings stay out of the bundle');
  assert.ok(bundle.includes('org_docx4j_wml'), 'the mappings are in the bundle, not fetched at run time');

  // The bundle runs where nothing of this package is on disk beside it - the add-in's situation.
  // DOMParser comes from the harness because the Jsonix runtime's UMD injects @xmldom/xmldom through
  // amdefine, which a bundler cannot satisfy, so a bundled runtime falls back to the browser's
  // globals: a real obstacle to bundling for Node, but the runtime's (plutext/jsonix), not this
  // package's, and a browser supplies these anyway. Without it the bundle reaches for XMLHttpRequest.
  const xmldom = fileURLToPath(new URL('../node_modules/@xmldom/xmldom/lib/index.js', import.meta.url));
  const ran = execFileSync(process.execPath, ['--input-type=module', '-e',
    `import { DOMParser, XMLSerializer } from ${JSON.stringify(xmldom)};`
    + 'globalThis.DOMParser ??= DOMParser; globalThis.XMLSerializer ??= XMLSerializer;'
    + `await import(${JSON.stringify(join(out, 'app.mjs'))});`], { stdio: 'pipe', cwd: out }).toString().trim();
  assert.equal(ran, 'org_docx4j_wml.Document 1', 'the bundle unmarshals where nothing of the package is on disk');
  console.log(`bundle: esbuild --bundle with no configuration, ${(bundle.length / 1048576).toFixed(1)}MB, unmarshals with nothing beside it`);
} finally {
  rmSync(out, { recursive: true, force: true });
}
