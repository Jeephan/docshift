import { build } from "esbuild";
import { rm, mkdir, cp, readFile, writeFile } from "node:fs/promises";

const DIST = "dist";

await rm(DIST, { recursive: true, force: true });
await mkdir(DIST, { recursive: true });

await build({
  entryPoints: ["src/app.js"],
  bundle: true,
  minify: true,
  format: "esm",
  target: ["es2020", "safari14"],
  outfile: `${DIST}/app.js`,
  legalComments: "none",
  alias: { buffer: "./src/shims/buffer.js" },
  logLevel: "info",
});

await build({
  entryPoints: ["src/styles.css"],
  bundle: true,
  minify: true,
  outfile: `${DIST}/styles.css`,
  logLevel: "info",
});

await cp("index.html", `${DIST}/index.html`);
await cp("public/favicon.svg", `${DIST}/favicon.svg`);

const size = (await readFile(`${DIST}/app.js`)).length;
console.log(`\n构建完成 → ${DIST}/   app.js ${(size / 1024).toFixed(0)} KB`);
