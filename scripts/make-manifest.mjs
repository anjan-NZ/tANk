// Rewrites the dev manifest to point at wherever the built files are hosted.
// Usage: node scripts/make-manifest.mjs https://owner.github.io/repo
import { readFileSync, writeFileSync, mkdirSync } from "node:fs";

const base = (process.argv[2] ?? process.env.TANK_BASE_URL ?? "").replace(/\/+$/, "");
if (!base) {
  console.error("Give me the public base URL, e.g. node scripts/make-manifest.mjs https://owner.github.io/repo");
  process.exit(1);
}

const src = readFileSync("manifest.xml", "utf8");
const out = src.replace(/https:\/\/localhost:3000/g, base);

if (out === src) console.warn("Nothing was replaced. Is the dev manifest still pointing at localhost:3000?");

mkdirSync("dist", { recursive: true });
writeFileSync("dist/manifest.xml", out);

// The installers are published next to the app so people can grab them directly.
// The .cmd pair is the one to point people at: no exe, so nothing for Smart App
// Control to block. Written with CRLF because this runs on a Linux runner and a
// Windows script host is fussy about line endings.
const crlf = (s) => s.replace(/\r?\n/g, "\r\n");

for (const name of ["install.ps1", "uninstall.ps1", "install.cmd", "uninstall.cmd"]) {
  const body = readFileSync("scripts/" + name, "utf8").replace(/__BASE_URL__/g, base);
  writeFileSync("dist/" + name, name.endsWith(".cmd") ? crlf(body) : body);
}

console.log("dist/manifest.xml   -> " + base + "/index.html");
console.log("dist/install.cmd    -> downloads from " + base + "/manifest.xml");
console.log("dist/install.ps1    -> downloads from " + base + "/manifest.xml");
