import assert from "node:assert/strict";
import test from "node:test";
import * as XLSX from "xlsx";
import { buildExportRows, compareSheets, getCommonHeaders, suggestKeyField } from "../app/diff-core.ts";

test("finds common headers and suggests identifier columns", () => {
  assert.deepEqual(getCommonHeaders(["SKU", "标题", "库存"], ["SKU", "库存", "价格"]), ["SKU", "库存"]);
  assert.equal(suggestKeyField(["商品名称", "ASIN", "库存"]), "ASIN");
});

test("classifies changed, added, removed and unchanged rows", () => {
  const result = compareSheets({
    beforeRows: [
      { SKU: "0001", 价格: "10", 库存: "5" },
      { SKU: "0002", 价格: "20", 库存: "8" },
      { SKU: "0003", 价格: "30", 库存: "9" },
    ],
    afterRows: [
      { SKU: "0001", 价格: "10", 库存: "7" },
      { SKU: "0002", 价格: "20", 库存: "8" },
      { SKU: "0004", 价格: "40", 库存: "6" },
    ],
    keyField: "SKU",
    compareFields: ["价格", "库存"],
  });

  assert.deepEqual(result.summary, {
    total: 4,
    changed: 1,
    added: 1,
    removed: 1,
    unchanged: 1,
    duplicate: 0,
    issues: 0,
  });
  assert.equal(result.entries.find((entry) => entry.key === "0001")?.changes[0].field, "库存");
});

test("keeps identifier leading zeros and reports unsafe duplicate keys", () => {
  const result = compareSheets({
    beforeRows: [{ SKU: "0007", 名称: "A" }, { SKU: "0007", 名称: "B" }, { SKU: "", 名称: "空" }],
    afterRows: [{ SKU: "0007", 名称: "C" }],
    keyField: "SKU",
    compareFields: ["名称"],
  });

  assert.equal(result.entries[0].key, "0007");
  assert.equal(result.entries[0].status, "duplicate");
  assert.equal(result.summary.issues, 2);
  assert.deepEqual(result.emptyKeyRows, { before: 1, after: 0 });
});

test("builds a valid Excel export and protects formula-like identifiers", () => {
  const result = compareSheets({
    beforeRows: [{ SKU: "0001", 价格: "10" }],
    afterRows: [{ SKU: "0001", 价格: "12" }, { SKU: "=CMD", 价格: "20" }],
    keyField: "SKU",
    compareFields: ["价格"],
  });
  const { summaryRows, detailRows } = buildExportRows(result, "SKU");
  assert.ok(summaryRows.some((row) => row.SKU === "0001"));
  assert.ok(summaryRows.some((row) => row.SKU === "'=CMD"));
  assert.equal(detailRows.length, 2);

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(summaryRows), "差异摘要");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(detailRows), "字段变化");
  const buffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx" });
  assert.ok(buffer.length > 1000);
});
