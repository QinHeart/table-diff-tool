import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

async function readBuiltPage(pathname = "index.html") {
  return readFile(new URL(`../dist/${pathname}`, import.meta.url), "utf8");
}

test("builds the classic table comparison page", async () => {
  const html = await readBuiltPage();
  assert.match(html, /<title>表格差异对比<\/title>/i);
  assert.match(html, /data-result-layout="classic"/);
  assert.match(html, /src="\/assets\/[^"]+\.js"/);
  assert.match(html, /href="\/assets\/[^"]+\.css"/);
});

test("builds the candidate page separately", async () => {
  const html = await readBuiltPage("candidate/index.html");
  assert.match(html, /<title>表格差异对比<\/title>/i);
  assert.match(html, /data-result-layout="candidate"/);
});
