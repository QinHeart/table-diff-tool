export type CellRow = Record<string, string>;

export type DifferenceStatus =
  | "changed"
  | "added"
  | "removed"
  | "unchanged"
  | "duplicate";

export interface FieldChange {
  field: string;
  before: string;
  after: string;
}

export interface DifferenceEntry {
  key: string;
  status: DifferenceStatus;
  before: CellRow | null;
  after: CellRow | null;
  changes: FieldChange[];
  duplicateBefore: number;
  duplicateAfter: number;
}

export interface ComparisonResult {
  entries: DifferenceEntry[];
  summary: Record<DifferenceStatus, number> & { total: number; issues: number };
  emptyKeyRows: { before: number; after: number };
}

const EXPORT_STATUS_LABEL: Record<DifferenceStatus, string> = {
  changed: "有变化",
  added: "新增",
  removed: "删除",
  unchanged: "相同",
  duplicate: "需处理",
};

function text(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value);
}

function comparisonValue(value: unknown, ignoreOuterWhitespace: boolean): string {
  const normalized = text(value);
  return ignoreOuterWhitespace ? normalized.trim() : normalized;
}

function buildIndex(rows: CellRow[], keyField: string) {
  const groups = new Map<string, CellRow[]>();
  let emptyKeyRows = 0;

  for (const row of rows) {
    const key = text(row[keyField]).trim();
    if (!key) {
      emptyKeyRows += 1;
      continue;
    }
    const existing = groups.get(key) ?? [];
    existing.push(row);
    groups.set(key, existing);
  }

  return { groups, emptyKeyRows };
}

export function getCommonHeaders(before: string[], after: string[]): string[] {
  const afterSet = new Set(after);
  return before.filter((header) => afterSet.has(header));
}

export function getUnionHeaders(before: string[], after: string[]): string[] {
  return [...new Set([...before, ...after])];
}

export function suggestKeyField(headers: string[]): string {
  const preferred = [
    /^sku$/i,
    /^asin$/i,
    /^id$/i,
    /订单号/,
    /商品编码/,
    /产品编号/,
    /唯一标识/,
  ];

  for (const pattern of preferred) {
    const found = headers.find((header) => pattern.test(header.trim()));
    if (found) return found;
  }

  return headers[0] ?? "";
}

export function describeDifference(entry: DifferenceEntry): string {
  if (entry.status === "duplicate") {
    return `旧文件 ${entry.duplicateBefore} 行，新文件 ${entry.duplicateAfter} 行使用了同一标识`;
  }
  if (entry.status === "added") return "仅存在于新文件";
  if (entry.status === "removed") return "仅存在于旧文件";
  if (entry.status === "unchanged") return "所选字段完全相同";
  return `${entry.changes.length} 个字段发生变化`;
}

function safeExportValue(value: string): string {
  return /^[=+\-@]/.test(value) ? `'${value}` : value;
}

export function buildExportRows(comparison: ComparisonResult, keyField: string) {
  const relevant = comparison.entries.filter((entry) => entry.status !== "unchanged");
  const summaryRows: Array<Record<string, string | number>> = relevant.map((entry) => ({
    状态: EXPORT_STATUS_LABEL[entry.status],
    [keyField]: safeExportValue(entry.key),
    变化字段数: entry.changes.length,
    变化字段: safeExportValue(entry.changes.map((change) => change.field).join("、")),
    说明: describeDifference(entry),
  }));
  const detailRows: Array<Record<string, string>> = relevant.flatMap((entry) => {
    if (entry.status === "duplicate") {
      return [{
        状态: EXPORT_STATUS_LABEL[entry.status],
        [keyField]: safeExportValue(entry.key),
        字段: "匹配列重复",
        旧值: `${entry.duplicateBefore} 行`,
        新值: `${entry.duplicateAfter} 行`,
      }];
    }
    return entry.changes.map((change) => ({
      状态: EXPORT_STATUS_LABEL[entry.status],
      [keyField]: safeExportValue(entry.key),
      字段: change.field,
      旧值: safeExportValue(change.before),
      新值: safeExportValue(change.after),
    }));
  });

  return { summaryRows, detailRows };
}

export function compareSheets({
  beforeRows,
  afterRows,
  keyField,
  compareFields,
  ignoreOuterWhitespace = true,
}: {
  beforeRows: CellRow[];
  afterRows: CellRow[];
  keyField: string;
  compareFields: string[];
  ignoreOuterWhitespace?: boolean;
}): ComparisonResult {
  const beforeIndex = buildIndex(beforeRows, keyField);
  const afterIndex = buildIndex(afterRows, keyField);
  const keys = [...new Set([...beforeIndex.groups.keys(), ...afterIndex.groups.keys()])].sort(
    (a, b) => a.localeCompare(b, "zh-CN", { numeric: true, sensitivity: "base" }),
  );

  const entries = keys.map<DifferenceEntry>((key) => {
    const beforeGroup = beforeIndex.groups.get(key) ?? [];
    const afterGroup = afterIndex.groups.get(key) ?? [];
    const before = beforeGroup[0] ?? null;
    const after = afterGroup[0] ?? null;

    if (beforeGroup.length > 1 || afterGroup.length > 1) {
      return {
        key,
        status: "duplicate",
        before,
        after,
        changes: [],
        duplicateBefore: beforeGroup.length,
        duplicateAfter: afterGroup.length,
      };
    }

    if (!before) {
      return {
        key,
        status: "added",
        before: null,
        after,
        changes: compareFields.map((field) => ({ field, before: "", after: text(after?.[field]) })),
        duplicateBefore: 0,
        duplicateAfter: 1,
      };
    }

    if (!after) {
      return {
        key,
        status: "removed",
        before,
        after: null,
        changes: compareFields.map((field) => ({ field, before: text(before[field]), after: "" })),
        duplicateBefore: 1,
        duplicateAfter: 0,
      };
    }

    const changes = compareFields.flatMap<FieldChange>((field) => {
      const beforeValue = text(before[field]);
      const afterValue = text(after[field]);
      return comparisonValue(beforeValue, ignoreOuterWhitespace) ===
        comparisonValue(afterValue, ignoreOuterWhitespace)
        ? []
        : [{ field, before: beforeValue, after: afterValue }];
    });

    return {
      key,
      status: changes.length ? "changed" : "unchanged",
      before,
      after,
      changes,
      duplicateBefore: 1,
      duplicateAfter: 1,
    };
  });

  const summary = {
    total: entries.length,
    changed: 0,
    added: 0,
    removed: 0,
    unchanged: 0,
    duplicate: 0,
    issues: 0,
  };

  for (const entry of entries) summary[entry.status] += 1;
  summary.issues = summary.duplicate + beforeIndex.emptyKeyRows + afterIndex.emptyKeyRows;

  return {
    entries,
    summary,
    emptyKeyRows: { before: beforeIndex.emptyKeyRows, after: afterIndex.emptyKeyRows },
  };
}
