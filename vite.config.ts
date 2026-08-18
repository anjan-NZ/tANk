import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import devCerts from "office-addin-dev-certs";
import { readFileSync } from "node:fs";

// Office refuses to load a task pane over plain http, so dev runs on the
// office-addin-dev-certs local CA. Falls back to http when certs are missing
// (useful for eyeballing the UI in a normal browser tab).
//
// Only ever asked for when serving. On a build it would reach into the machine's
// certificate store, which hangs forever on a CI runner with nobody to accept it.
export default defineConfig(async ({ command }) => {
  let https: { key: Buffer; cert: Buffer } | undefined;
  if (command === "serve") {
    try {
      const opts = await devCerts.getHttpsServerOptions();
      https = { key: opts.key as unknown as Buffer, cert: opts.cert as unknown as Buffer };
    } catch {
      https = undefined;
    }
  }

  const pkg = JSON.parse(readFileSync("package.json", "utf8"));
  // GITHUB_REPOSITORY is set by the build; locally there is nothing to check against.
  const repo = process.env.GITHUB_REPOSITORY ?? "";

  return {
    define: {
      __APP_VERSION__: JSON.stringify(pkg.version),
      __REPO_SLUG__: JSON.stringify(repo),
    },
    plugins: [react()],
    base: "./",
    server: { port: 3000, strictPort: true, https },
    // No source map in the published build, it would serve the readable source.
    build: { outDir: "dist", sourcemap: false },
  };
});
