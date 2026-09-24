// CR-005: the package bundles with no bundler configuration, and a bundle carries the 103 mapping
// modules rather than the 309 .mjs files in modules/. Two witnesses, because a bundler picks the
// runtime's entry from the consumer's syntax rather than the output format (jsonix-CR-005 section
// 8): an `import` consumer, and a `require()` consumer of the UMD mappings, which is the only one
// that reaches jsonix.js's footer.
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

  // The bundle runs where nothing of this package is on disk beside it - the add-in's situation -
  // and with no globals supplied. Until @docx4j/jsonix 3.3.0 this needed globalThis.DOMParser and
  // XMLSerializer, because the runtime's UMD injected @xmldom/xmldom through amdefine, which no
  // bundler can satisfy: it fell back to browser globals and reached for XMLHttpRequest
  // (jsonix-CR-005). The absence of that shim here is the acceptance check for that fix.
  const ran = execFileSync(process.execPath, [join(out, 'app.mjs')], { stdio: 'pipe', cwd: out }).toString().trim();
  assert.equal(ran, 'org_docx4j_wml.Document 1', 'the bundle unmarshals where nothing of the package is on disk');

  // The require() witness: a CommonJS consumer of the UMD mappings, bundled to cjs, which is the
  // only shape that takes the runtime's UMD footer.
  execFileSync(esbuild, [join(here, 'bundle/consumer.cjs'), '--bundle', '--platform=node',
    '--format=cjs', `--outfile=${join(out, 'app.cjs')}`], { stdio: 'pipe' });
  const ranCjs = execFileSync(process.execPath, [join(out, 'app.cjs')], { stdio: 'pipe', cwd: out }).toString().trim();
  assert.equal(ranCjs, 'org_docx4j_relationships.Relationships 1 true',
    'a require() consumer bundled to cjs unmarshals and marshals, with no globals supplied');

  console.log(`bundle: esbuild --bundle with no configuration, ${(bundle.length / 1048576).toFixed(1)}MB esm`
    + ' and a cjs require() consumer, both unmarshalling with nothing beside them');
} finally {
  rmSync(out, { recursive: true, force: true });
}
