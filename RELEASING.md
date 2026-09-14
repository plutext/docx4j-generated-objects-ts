# Releasing

`@docx4j/generated-objects-ts` is published to npm from GitHub Actions
(`.github/workflows/push-to-npm.yml`), like `@docx4j/jsonix` and `@docx4j/jsonix-schema-compiler`.

## Publishing target

Publishing uses npm **trusted publishing**: no npm token is stored anywhere. The settings of the npm
package `@docx4j/generated-objects-ts` (npmjs.com, scope `@docx4j`) name the publisher GitHub Actions and
the GitHub repository allowed to publish it: organization or user `plutext`, repository
`docx4j-generated-objects-ts`, workflow `push-to-npm.yml`, no environment. Renaming the workflow file
breaks publishing until the npmjs.com setting is changed to match.

A trusted publisher is configured on an existing package, so the **first** version is published by hand
(`npm publish --access public` from a clean checkout, after `npm install`; `prepublishOnly` runs typecheck
and test), and the trusted publisher is added on npmjs.com afterwards.

npm versions cannot be reused once published (even after an unpublish): a problem found after
publishing ships as the next patch version.

The package ships `dist/` (built from `src/`), the generated `modules/`, `README.md`, `LICENSE` and
`NOTICE`; its only runtime dependency is `@docx4j/jsonix`. The compiler is not a dependency: `modules/`
is generated from the sibling checkout (see `generate.md`) and committed.

## Steps

```bash
# 1. If modules/ is to be regenerated for the release, do it first and commit it citing the compiler
#    and docx4j commits (see generate.md); a regeneration with unchanged inputs is an empty diff.

# 2. Set the version (no tag or commit yet)
npm version 0.1.1 --no-git-tag-version

# 3. From a clean node_modules, check against the registry runtime and inspect the package
rm -rf node_modules dist && npm install
npm run typecheck && npm test
npm pack --dry-run

# 4. Commit, tag and push (the tag is the bare version, as the workflow checks it; a leading v is tolerated)
git commit -am "Version 0.1.1"
git tag -a 0.1.1 -m "Version 0.1.1"
git push origin main 0.1.1

# 5. Publish: create a GitHub release for the tag
gh release create 0.1.1 --title "0.1.1" --notes "..."
```

Publishing the release runs `push-to-npm.yml`, which installs, fails unless the release tag equals
`package.json`'s version, runs typecheck and test, checks the tree is unchanged, and runs `npm pack` and
`npm publish` (with provenance, via OIDC).

If the workflow fails before the publish step, fix the problem, move the tag, and re-run it from the
Actions tab (or delete and recreate the release).

`@docx4j/core-ts` depends on this package: after a release that changes the facade or the declarations,
update its dependency range.
