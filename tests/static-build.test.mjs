import assert from "node:assert/strict";
import test from "node:test";
import { readFile, readdir } from "node:fs/promises";

test("o pacote independente contém entrada HTML, scripts, worker e cabeçalhos", async () => {
  const html = await readFile(new URL("../dist/index.html", import.meta.url), "utf8");
  assert.match(html, /Arquivo Ágil/);
  assert.match(html, /id="root"/);
  assert.match(html, /src="\/assets\/[^"]+\.js"/);
  assert.doesNotMatch(html, /chatgpt-auth|codex-preview|main\.tsx/);
  const assets = await readdir(new URL("../dist/assets/", import.meta.url));
  assert.ok(assets.some((name) => /^soc-import\.worker-.*\.js$/.test(name)), "o importador precisa ser empacotado");
  const headers = await readFile(new URL("../dist/_headers", import.meta.url), "utf8");
  assert.match(headers, /X-Content-Type-Options: nosniff/);
});
