#!/usr/bin/env sh
# Regenerates the Jsonix mappings and TypeScript declarations in this directory from docx4j's
# schemas. See README.md. Run from the repository root:
#
#   OfficeOpenXML/generate.sh [path/to/docx4j/xsd/ROOT.xsd]
#
# Requires JDK 11+ and a built full jar (./mvnw -pl .,compiler,full install -DskipTests).
set -eu
ROOT_XSD="${1:-../docx4j/xsd/ROOT.xsd}"
JAR="$(ls full/target/jsonix-schema-compiler-full-*.jar | grep -v -- '-sources\|original-' | head -1)"
OUT="$(mktemp -d)"
trap 'rm -rf "$OUT"' EXIT
# -Xinheritance/-Xannotate/-Xinject-code: docx4j's schemas carry customizations for those XJC
# plugins (they shape the generated Java, not the XML model); XJC refuses them unless enabled.
java -jar "$JAR" -d "$OUT" -generateTypeScript -Xinheritance -Xannotate -Xinject-code \
  "$ROOT_XSD" -b OfficeOpenXML/bindings.xjb
# XJC also writes Java sources (org/...) into the target directory; keep only the mappings and declarations.
rm -f OfficeOpenXML/*.js OfficeOpenXML/*.mjs OfficeOpenXML/*.d.ts OfficeOpenXML/*.d.mts
cp "$OUT"/*.js "$OUT"/*.mjs "$OUT"/*.d.ts "$OUT"/*.d.mts OfficeOpenXML/
echo "generated $(ls OfficeOpenXML/*.js | wc -l) modules into OfficeOpenXML/"
