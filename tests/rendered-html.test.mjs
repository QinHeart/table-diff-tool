import assert from "node:assert/strict";
import test from "node:test";

async function render(pathname = "/") {
  const workerUrl = new URL("../dist/server/index.js", import.meta.url);
  workerUrl.searchParams.set("test", `${process.pid}-${Date.now()}`);
  const { default: worker } = await import(workerUrl.href);
  return worker.fetch(
    new Request(`http://localhost${pathname}`, { headers: { accept: "text/html" } }),
    { ASSETS: { fetch: async () => new Response("Not found", { status: 404 }) } },
    { waitUntil() {}, passThroughOnException() {} },
  );
}

test("server-renders the table comparison tool", async () => {
  const response = await render();
  assert.equal(response.status, 200);
  assert.match(response.headers.get("content-type") ?? "", /^text\/html\b/i);
  const html = await response.text();
  assert.match(html, /<title>表格差异对比<\/title>/i);
  assert.match(html, /两份文件，快速看清差异/);
  assert.match(html, /载入虚构示例/);
  assert.match(html, /文件只在当前浏览器中处理/);
  assert.doesNotMatch(html, /codex-preview|SkeletonPreview|Your site is taking shape/i);
});

test("keeps the original page and exposes the result-display candidate separately", async () => {
  const response = await render("/candidate");
  assert.equal(response.status, 200);
  const html = await response.text();
  assert.match(html, /<title>表格差异对比 · 结果展示候选版<\/title>/i);
  assert.match(html, /两份文件，快速看清差异/);
});
