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

// The installer is published next to the app so people can grab it with one command.
const installer = readFileSync("scripts/install.ps1", "utf8").replace(/__BASE_URL__/g, base);
writeFileSync("dist/install.ps1", installer);
writeFileSync("dist/uninstall.ps1", readFileSync("scripts/uninstall.ps1", "utf8"));

console.log("dist/manifest.xml   -> " + base + "/index.html");
console.log("dist/install.ps1    -> downloads from " + base + "/manifest.xml");
