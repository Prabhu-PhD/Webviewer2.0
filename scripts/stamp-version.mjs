// Replaces every ?v=<token> query in web/content.html with ?v=<package.json version>
// Run as a prebuild step so the deployed file always matches the declared version.
import { readFileSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { join, dirname } from "node:path";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const pkg  = JSON.parse(readFileSync(join(root, "package.json"), "utf8"));
const ver  = pkg.version.replace(/\./g, "");   // "0.1.0" → "010"

const htmlPath = join(root, "web", "content.html");
const before   = readFileSync(htmlPath, "utf8");
const after    = before.replace(/\?v=\d+/g, `?v=${ver}`);

if (before !== after) {
  writeFileSync(htmlPath, after, "utf8");
  console.log(`stamp-version: updated ?v= tokens to ${ver} in content.html`);
} else {
  console.log(`stamp-version: ?v= tokens already at ${ver}`);
}
