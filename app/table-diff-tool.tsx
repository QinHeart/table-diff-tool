"use client";

import { DragEvent, useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import {
  CellRow,
  ComparisonResult,
  DifferenceEntry,
  DifferenceStatus,
  buildExportRows,
  compareSheets,
  describeDifference,
  getCommonHeaders,
  getUnionHeaders,
  suggestKeyField,
} from "./diff-core";

interface SheetData {
  name: string;
  headers: string[];
  rows: CellRow[];
}

interface WorkbookData {
  fileName: string;
  sheets: SheetData[];
}

type Side = "before" | "after";
type ResultLayout = "classic" | "candidate";
type ResultFilter = "all" | "attention" | DifferenceStatus;
type CandidateSort = "priority" | "key" | "changes";

const STATUS_LABEL: Record<DifferenceStatus, string> = {
  changed: "有变化",
  added: "新增",
  removed: "删除",
  unchanged: "相同",
  duplicate: "需处理",
};

const DEMO_BEFORE: CellRow[] = [
  { SKU: "DEMO-0001", 商品名称: "便携标签打印机", 站点: "UK", 售价: "29.99", 库存: "86", 状态: "在售" },
  { SKU: "DEMO-0002", 商品名称: "透明标签纸", 站点: "UK", 售价: "9.99", 库存: "240", 状态: "在售" },
  { SKU: "DEMO-0003", 商品名称: "收纳标签套装", 站点: "DE", 售价: "12.90", 库存: "48", 状态: "在售" },
  { SKU: "DEMO-0004", 商品名称: "迷你热敏纸", 站点: "FR", 售价: "8.50", 库存: "0", 状态: "待下架" },
];

const DEMO_AFTER: CellRow[] = [
  { SKU: "DEMO-0001", 商品名称: "便携蓝牙标签打印机", 站点: "UK", 售价: "27.99", 库存: "112", 状态: "在售" },
  { SKU: "DEMO-0002", 商品名称: "透明标签纸", 站点: "UK", 售价: "9.99", 库存: "240", 状态: "在售" },
  { SKU: "DEMO-0003", 商品名称: "收纳标签套装", 站点: "DE", 售价: "12.90", 库存: "31", 状态: "在售" },
  { SKU: "DEMO-0005", 商品名称: "彩色标签纸", 站点: "FR", 售价: "10.50", 库存: "75", 状态: "新品" },
];

function makeUniqueHeaders(values: unknown[], columnCount: number): string[] {
  const seen = new Map<string, number>();
  return Array.from({ length: columnCount }, (_, index) => {
    const base = String(values[index] ?? "").trim() || `未命名列 ${index + 1}`;
    const count = (seen.get(base) ?? 0) + 1;
    seen.set(base, count);
    return count === 1 ? base : `${base}（${count}）`;
  });
}

function worksheetToData(name: string, worksheet: XLSX.WorkSheet): SheetData {
  const matrix = XLSX.utils.sheet_to_json<unknown[]>(worksheet, {
    header: 1,
    raw: false,
    defval: "",
    blankrows: false,
  });

  if (!matrix.length) return { name, headers: [], rows: [] };

  const columnCount = matrix.reduce((max, row) => Math.max(max, row.length), 0);
  const headers = makeUniqueHeaders(matrix[0] ?? [], columnCount);
  const rows = matrix.slice(1).flatMap<CellRow>((values) => {
    const row = Object.fromEntries(headers.map((header, index) => [header, String(values[index] ?? "")])) as CellRow;
    return Object.values(row).some((value) => value.trim() !== "") ? [row] : [];
  });

  return { name, headers, rows };
}

async function readWorkbook(file: File): Promise<WorkbookData> {
  const extension = file.name.split(".").pop()?.toLowerCase();
  if (!extension || !["xlsx", "xls", "csv", "tsv"].includes(extension)) {
    throw new Error("请选择 Excel、CSV 或 TSV 文件");
  }

  const workbook = XLSX.read(await file.arrayBuffer(), {
    type: "array",
    raw: false,
  });
  const sheets = workbook.SheetNames.map((name) => worksheetToData(name, workbook.Sheets[name]));
  if (!sheets.some((sheet) => sheet.headers.length)) throw new Error("文件中没有可读取的表格数据");
  return { fileName: file.name, sheets };
}

function makeDemoWorkbook(fileName: string, rows: CellRow[]): WorkbookData {
  return {
    fileName,
    sheets: [{ name: "商品资料", headers: Object.keys(rows[0]), rows }],
  };
}

function FileCard({
  side,
  workbook,
  sheetName,
  onFile,
  onSheet,
}: {
  side: Side;
  workbook: WorkbookData | null;
  sheetName: string;
  onFile: (side: Side, file: File) => void;
  onSheet: (name: string) => void;
}) {
  const sheet = workbook?.sheets.find((item) => item.name === sheetName) ?? null;
  const [dragging, setDragging] = useState(false);

  function handleDrop(event: DragEvent<HTMLLabelElement>) {
    event.preventDefault();
    setDragging(false);
    const file = event.dataTransfer.files[0];
    if (file) onFile(side, file);
  }

  return (
    <section className={`file-card ${dragging ? "is-dragging" : ""}`}>
      <div className="file-card-heading">
        <span className="step-number">{side === "before" ? "A" : "B"}</span>
        <div>
          <h2>{side === "before" ? "旧文件" : "新文件"}</h2>
          <p>{side === "before" ? "作为比较基准" : "查看相对变化"}</p>
        </div>
      </div>

      <label
        className="drop-zone"
        onDragOver={(event) => {
          event.preventDefault();
          setDragging(true);
        }}
        onDragLeave={() => setDragging(false)}
        onDrop={handleDrop}
      >
        <input
          type="file"
          accept=".xlsx,.xls,.csv,.tsv"
          onChange={(event) => {
            const file = event.target.files?.[0];
            if (file) onFile(side, file);
            event.target.value = "";
          }}
        />
        <strong>{workbook ? "更换文件" : "选择或拖入文件"}</strong>
        <span>支持 XLSX、XLS、CSV、TSV</span>
      </label>

      {workbook && sheet ? (
        <div className="file-meta" aria-live="polite">
          <div className="file-name" title={workbook.fileName}>{workbook.fileName}</div>
          <div className="file-stats">
            <span>{sheet.rows.length.toLocaleString("zh-CN")} 行</span>
            <span>{sheet.headers.length} 列</span>
          </div>
          {workbook.sheets.length > 1 ? (
            <label className="sheet-picker">
              工作表
              <select value={sheetName} onChange={(event) => onSheet(event.target.value)}>
                {workbook.sheets.map((item) => <option key={item.name}>{item.name}</option>)}
              </select>
            </label>
          ) : (
            <span className="single-sheet">工作表：{sheet.name}</span>
          )}
        </div>
      ) : null}
    </section>
  );
}

const CANDIDATE_STATUS_LABEL: Record<DifferenceStatus, string> = {
  changed: "修改",
  added: "新增",
  removed: "删除",
  unchanged: "相同",
  duplicate: "数据问题",
};

function shownValue(value: string) {
  return value || "（空）";
}

function rowNameFor(entry: DifferenceEntry) {
  const row = entry.after ?? entry.before;
  return row?.["商品名称"] || row?.["名称"] || "";
}

function changedFieldSummary(entry: DifferenceEntry) {
  if (entry.status === "duplicate") return "匹配列重复";
  if (entry.status === "unchanged") return "—";
  if (entry.status === "added") return "整行新增";
  if (entry.status === "removed") return "整行删除";
  const names = entry.changes.slice(0, 2).map((change) => change.field).join("、");
  return entry.changes.length > 2 ? `${names} +${entry.changes.length - 2}` : names;
}

function DifferenceDrawer({
  entry,
  keyField,
  onClose,
}: {
  entry: DifferenceEntry | null;
  keyField: string;
  onClose: () => void;
}) {
  useEffect(() => {
    if (!entry) return;
    const closeOnEscape = (event: KeyboardEvent) => {
      if (event.key === "Escape") onClose();
    };
    window.addEventListener("keydown", closeOnEscape);
    return () => window.removeEventListener("keydown", closeOnEscape);
  }, [entry, onClose]);

  if (!entry) return null;
  const rowName = rowNameFor(entry);

  return (
    <div className="drawer-layer">
      <button className="drawer-backdrop" type="button" aria-label="关闭差异详情" onClick={onClose} />
      <aside className="detail-drawer" role="dialog" aria-modal="true" aria-labelledby="detail-drawer-title">
        <header className="drawer-header">
          <div>
            <span className={`status-pill status-pill-${entry.status}`}>{CANDIDATE_STATUS_LABEL[entry.status]}</span>
            <h2 id="detail-drawer-title">{entry.key}</h2>
            {rowName ? <p>{rowName}</p> : null}
          </div>
          <button type="button" className="drawer-close" onClick={onClose} aria-label="关闭详情">×</button>
        </header>

        <div className="drawer-summary">
          <span>{keyField}</span>
          <strong>{entry.key}</strong>
          <p>{describeDifference(entry)}</p>
        </div>

        <div className="drawer-content">
          {entry.status === "duplicate" ? (
            <div className="duplicate-message">请先在源文件中处理重复的 {keyField}，再重新比较。</div>
          ) : entry.changes.length ? entry.changes.map((change) => (
            <section className="drawer-change" key={change.field}>
              <h3>{change.field}</h3>
              <div className="drawer-values">
                <div><span>旧值</span><p className="detail-old">{shownValue(change.before)}</p></div>
                <span className="drawer-arrow" aria-hidden="true">→</span>
                <div><span>新值</span><p className="detail-new">{shownValue(change.after)}</p></div>
              </div>
            </section>
          )) : (
            <div className="duplicate-message">所选字段没有差异。</div>
          )}
        </div>
      </aside>
    </div>
  );
}

export function TableDiffTool({ resultLayout = "classic" }: { resultLayout?: ResultLayout }) {
  const defaultResultFilter: ResultFilter = resultLayout === "candidate" ? "attention" : "all";
  const [beforeBook, setBeforeBook] = useState<WorkbookData | null>(null);
  const [afterBook, setAfterBook] = useState<WorkbookData | null>(null);
  const [beforeSheetName, setBeforeSheetName] = useState("");
  const [afterSheetName, setAfterSheetName] = useState("");
  const [keyField, setKeyField] = useState("");
  const [selectedFields, setSelectedFields] = useState<string[]>([]);
  const [ignoreOuterWhitespace, setIgnoreOuterWhitespace] = useState(true);
  const [comparison, setComparison] = useState<ComparisonResult | null>(null);
  const [filter, setFilter] = useState<ResultFilter>(defaultResultFilter);
  const [query, setQuery] = useState("");
  const [error, setError] = useState("");
  const [loadingSide, setLoadingSide] = useState<Side | null>(null);
  const [exportMessage, setExportMessage] = useState("");
  const [candidateSort, setCandidateSort] = useState<CandidateSort>("priority");
  const [pageSize, setPageSize] = useState(50);
  const [currentPage, setCurrentPage] = useState(1);
  const [selectedEntry, setSelectedEntry] = useState<DifferenceEntry | null>(null);

  const beforeSheet = beforeBook?.sheets.find((sheet) => sheet.name === beforeSheetName) ?? null;
  const afterSheet = afterBook?.sheets.find((sheet) => sheet.name === afterSheetName) ?? null;
  const commonHeaders = useMemo(
    () => getCommonHeaders(beforeSheet?.headers ?? [], afterSheet?.headers ?? []),
    [beforeSheet, afterSheet],
  );
  const unionHeaders = useMemo(
    () => getUnionHeaders(beforeSheet?.headers ?? [], afterSheet?.headers ?? []),
    [beforeSheet, afterSheet],
  );

  function configureForSheets(nextBefore: SheetData | null, nextAfter: SheetData | null) {
    if (!nextBefore || !nextAfter) {
      setKeyField("");
      setSelectedFields([]);
    } else {
      const common = getCommonHeaders(nextBefore.headers, nextAfter.headers);
      const union = getUnionHeaders(nextBefore.headers, nextAfter.headers);
      const suggested = suggestKeyField(common);
      setKeyField(suggested);
      setSelectedFields(union.filter((header) => header !== suggested));
    }
    setComparison(null);
    setFilter(defaultResultFilter);
    setCurrentPage(1);
    setSelectedEntry(null);
  }

  function changeCandidateFilter(nextFilter: ResultFilter) {
    setFilter(nextFilter);
    setCurrentPage(1);
    setSelectedEntry(null);
  }

  async function handleFile(side: Side, file: File) {
    setError("");
    setLoadingSide(side);
    try {
      const workbook = await readWorkbook(file);
      const firstSheet = workbook.sheets.find((sheet) => sheet.headers.length) ?? workbook.sheets[0];
      if (side === "before") {
        setBeforeBook(workbook);
        setBeforeSheetName(firstSheet.name);
        configureForSheets(firstSheet, afterSheet);
      } else {
        setAfterBook(workbook);
        setAfterSheetName(firstSheet.name);
        configureForSheets(beforeSheet, firstSheet);
      }
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : "文件读取失败");
    } finally {
      setLoadingSide(null);
    }
  }

  function loadDemo() {
    const before = makeDemoWorkbook("旧版商品资料（虚构示例）.xlsx", DEMO_BEFORE);
    const after = makeDemoWorkbook("新版商品资料（虚构示例）.xlsx", DEMO_AFTER);
    const fields = before.sheets[0].headers.filter((header) => header !== "SKU");
    setBeforeBook(before);
    setAfterBook(after);
    setBeforeSheetName("商品资料");
    setAfterSheetName("商品资料");
    setKeyField("SKU");
    setSelectedFields(fields);
    setComparison(compareSheets({
      beforeRows: DEMO_BEFORE,
      afterRows: DEMO_AFTER,
      keyField: "SKU",
      compareFields: fields,
    }));
    setFilter(defaultResultFilter);
    setQuery("");
    setError("");
    setExportMessage("");
    setCurrentPage(1);
    setSelectedEntry(null);
  }

  function runComparison() {
    if (!beforeSheet || !afterSheet || !keyField) {
      setError("请先选择两份文件和匹配列");
      return;
    }
    if (!selectedFields.length) {
      setError("请至少选择一个需要比较的字段");
      return;
    }
    setComparison(compareSheets({
      beforeRows: beforeSheet.rows,
      afterRows: afterSheet.rows,
      keyField,
      compareFields: selectedFields,
      ignoreOuterWhitespace,
    }));
    setFilter(defaultResultFilter);
    setQuery("");
    setError("");
    setExportMessage("");
    setCurrentPage(1);
    setSelectedEntry(null);
    window.setTimeout(() => document.getElementById("comparison-results")?.scrollIntoView({ behavior: "smooth" }), 50);
  }

  function exportDifferences() {
    if (!comparison) return;
    const { summaryRows, detailRows } = buildExportRows(comparison, keyField);

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(summaryRows), "差异摘要");
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(detailRows), "字段变化");
    const content = XLSX.write(workbook, { type: "array", bookType: "xlsx" });
    const url = URL.createObjectURL(new Blob([content], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }));
    const link = document.createElement("a");
    link.href = url;
    link.download = `表格差异_${new Date().toISOString().slice(0, 10)}.xlsx`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
    setExportMessage(`已生成 ${summaryRows.length} 条差异记录`);
  }

  const visibleEntries = useMemo(() => {
    if (!comparison) return [];
    const normalizedQuery = query.trim().toLocaleLowerCase("zh-CN");
    const entries = comparison.entries.filter((entry) => {
      const statusMatches = filter === "all"
        || (filter === "attention" && entry.status !== "unchanged")
        || entry.status === filter;
      if (!statusMatches) return false;
      if (!normalizedQuery) return true;
      const haystack = [
        entry.key,
        ...Object.values(entry.before ?? {}),
        ...Object.values(entry.after ?? {}),
        ...entry.changes.flatMap((change) => [change.field, change.before, change.after]),
      ].join(" ").toLocaleLowerCase("zh-CN");
      return haystack.includes(normalizedQuery);
    });
    if (resultLayout === "candidate") {
      const priority: Record<DifferenceStatus, number> = {
        duplicate: 0,
        changed: 1,
        added: 2,
        removed: 3,
        unchanged: 4,
      };
      entries.sort((a, b) => {
        if (candidateSort === "key") {
          return a.key.localeCompare(b.key, "zh-CN", { numeric: true, sensitivity: "base" });
        }
        if (candidateSort === "changes") {
          return b.changes.length - a.changes.length
            || priority[a.status] - priority[b.status]
            || a.key.localeCompare(b.key, "zh-CN", { numeric: true, sensitivity: "base" });
        }
        return priority[a.status] - priority[b.status]
          || a.key.localeCompare(b.key, "zh-CN", { numeric: true, sensitivity: "base" });
      });
    }
    return entries;
  }, [candidateSort, comparison, filter, query, resultLayout]);

  const ready = Boolean(beforeSheet && afterSheet && keyField && selectedFields.length);
  const currentStep = comparison ? 3 : beforeSheet && afterSheet ? 2 : 1;
  const attentionCount = comparison
    ? comparison.summary.changed + comparison.summary.added + comparison.summary.removed + comparison.summary.issues
    : 0;
  const attentionEntryCount = comparison
    ? comparison.summary.changed + comparison.summary.added + comparison.summary.removed + comparison.summary.duplicate
    : 0;
  const totalPages = Math.max(1, Math.ceil(visibleEntries.length / pageSize));
  const safeCurrentPage = Math.min(currentPage, totalPages);
  const pageStart = (safeCurrentPage - 1) * pageSize;
  const paginatedEntries = visibleEntries.slice(pageStart, pageStart + pageSize);

  return (
    <main className="app-shell">
      <header className="tool-header">
        <h1>表格差异对比</h1>
        <button className="demo-button" type="button" onClick={loadDemo}>载入示例</button>
      </header>

      <nav className="progress" aria-label="操作进度">
        {["上传文件", "设置比较", "查看结果"].map((label, index) => (
          <div className={currentStep >= index + 1 ? "is-active" : ""} key={label}>
            <span>{index + 1}</span>{label}
          </div>
        ))}
      </nav>

      <div className="upload-grid">
        <FileCard
          side="before"
          workbook={beforeBook}
          sheetName={beforeSheetName}
          onFile={handleFile}
          onSheet={(name) => {
            const nextSheet = beforeBook?.sheets.find((sheet) => sheet.name === name) ?? null;
            setBeforeSheetName(name);
            configureForSheets(nextSheet, afterSheet);
          }}
        />
        <FileCard
          side="after"
          workbook={afterBook}
          sheetName={afterSheetName}
          onFile={handleFile}
          onSheet={(name) => {
            const nextSheet = afterBook?.sheets.find((sheet) => sheet.name === name) ?? null;
            setAfterSheetName(name);
            configureForSheets(beforeSheet, nextSheet);
          }}
        />
      </div>

      {loadingSide ? <p className="notice" role="status">正在读取{loadingSide === "before" ? "旧" : "新"}文件…</p> : null}
      {error ? <p className="error-notice" role="alert">{error}</p> : null}

      {beforeSheet && afterSheet ? (
        <section className="settings-panel">
          <div className="section-heading">
            <div>
              <p className="section-kicker">第 2 步</p>
              <h2>设置比较方式</h2>
            </div>
            <p>{commonHeaders.length} 个同名字段可用</p>
          </div>

          {commonHeaders.length ? (
            <>
              <div className="settings-row">
                <label className="select-field">
                  <span>用哪一列匹配每一行</span>
                  <select
                    value={keyField}
                    onChange={(event) => {
                      const nextKey = event.target.value;
                      setKeyField(nextKey);
                      setSelectedFields(unionHeaders.filter((header) => header !== nextKey));
                      setComparison(null);
                    }}
                  >
                    {commonHeaders.map((header) => <option key={header}>{header}</option>)}
                  </select>
                  <small>这一列应当是唯一的，例如 SKU、ASIN 或订单号</small>
                </label>
                <label className="toggle-field">
                  <input
                    type="checkbox"
                    checked={ignoreOuterWhitespace}
                    onChange={(event) => { setIgnoreOuterWhitespace(event.target.checked); setComparison(null); }}
                  />
                  <span>忽略单元格首尾空格</span>
                </label>
              </div>

              <fieldset className="field-selector">
                <legend>需要比较的字段</legend>
                <div className="field-actions">
                  <button type="button" onClick={() => { setSelectedFields(unionHeaders.filter((header) => header !== keyField)); setComparison(null); }}>全选</button>
                  <button type="button" onClick={() => { setSelectedFields([]); setComparison(null); }}>清空</button>
                  <span>已选 {selectedFields.length} 项</span>
                </div>
                <div className="field-chips">
                  {unionHeaders.filter((header) => header !== keyField).map((header) => (
                    <label key={header}>
                      <input
                        type="checkbox"
                        checked={selectedFields.includes(header)}
                        onChange={(event) => {
                          setSelectedFields((current) => event.target.checked
                            ? [...current, header]
                            : current.filter((field) => field !== header));
                          setComparison(null);
                        }}
                      />
                      <span>{header}</span>
                    </label>
                  ))}
                </div>
              </fieldset>

              <button className="compare-button" type="button" disabled={!ready} onClick={runComparison}>
                开始对比
              </button>
            </>
          ) : (
            <p className="empty-state">两份文件没有同名字段，暂时无法选择匹配列。请检查表头。</p>
          )}
        </section>
      ) : null}

      {comparison ? (
        <section className="results" id="comparison-results">
          <div className="section-heading results-heading">
            <div>
              <p className="section-kicker">第 3 步</p>
              <h2>比较结果</h2>
            </div>
            <button type="button" className="export-button" onClick={exportDifferences}>导出差异 Excel</button>
          </div>

          {exportMessage ? <p className="success-notice" role="status">{exportMessage}</p> : null}

          {resultLayout === "candidate" ? (
            <>
              <p className="compact-result-summary">
                <strong>{attentionCount} 项需要关注</strong>
                <span>共比较 {comparison.summary.total} 个 {keyField}，{comparison.summary.unchanged} 个完全相同</span>
              </p>

              {comparison.summary.issues ? (
                <p className="issue-note">
                  有 {comparison.summary.duplicate} 个重复标识；旧文件 {comparison.emptyKeyRows.before} 行、新文件 {comparison.emptyKeyRows.after} 行缺少匹配值。请先修正这些数据再重新比较。
                </p>
              ) : null}

              <div className="result-toolbar candidate-toolbar">
                <div className="filter-tabs candidate-filter-tabs" role="group" aria-label="结果筛选">
                  {([
                    ["attention", `需要关注 ${attentionEntryCount}`],
                    ["changed", `修改 ${comparison.summary.changed}`],
                    ["added", `新增 ${comparison.summary.added}`],
                    ["removed", `删除 ${comparison.summary.removed}`],
                    ["duplicate", `数据问题 ${comparison.summary.duplicate}`],
                    ["unchanged", `相同 ${comparison.summary.unchanged}`],
                    ["all", `全部 ${comparison.summary.total}`],
                  ] as const).map(([status, label]) => (
                    <button
                      type="button"
                      key={status}
                      className={filter === status ? "is-active" : ""}
                      onClick={() => changeCandidateFilter(status)}
                    >
                      {label}
                    </button>
                  ))}
                </div>
                <label className="search-field">
                  <span className="sr-only">搜索结果</span>
                  <input
                    value={query}
                    onChange={(event) => {
                      setQuery(event.target.value);
                      setCurrentPage(1);
                      setSelectedEntry(null);
                    }}
                    placeholder={`搜索 ${keyField}、字段或内容`}
                  />
                </label>
              </div>

              <div className="compact-table-controls">
                <p aria-live="polite">
                  {visibleEntries.length
                    ? `显示第 ${pageStart + 1}–${Math.min(pageStart + pageSize, visibleEntries.length)} 条，共 ${visibleEntries.length} 条`
                    : "当前没有记录"}
                </p>
                <div>
                  <label>
                    排序
                    <select
                      value={candidateSort}
                      onChange={(event) => {
                        setCandidateSort(event.target.value as CandidateSort);
                        setCurrentPage(1);
                      }}
                    >
                      <option value="priority">按处理优先级</option>
                      <option value="key">按 {keyField}</option>
                      <option value="changes">按变化数量</option>
                    </select>
                  </label>
                  <label>
                    每页
                    <select
                      value={pageSize}
                      onChange={(event) => {
                        setPageSize(Number(event.target.value));
                        setCurrentPage(1);
                      }}
                    >
                      <option value={25}>25 条</option>
                      <option value={50}>50 条</option>
                      <option value={100}>100 条</option>
                    </select>
                  </label>
                </div>
              </div>

              {paginatedEntries.length ? (
                <>
                  <div className="compact-table-shell" aria-live="polite">
                  <div className="compact-table-scroll">
                    <div className="compact-result-table" role="table" aria-label="表格差异结果">
                      <div className="compact-table-header" role="row">
                        <span role="columnheader">状态</span>
                        <span role="columnheader">{keyField}</span>
                        <span role="columnheader">名称</span>
                        <span role="columnheader">变化字段</span>
                        <span role="columnheader">首项变化</span>
                        <span role="columnheader">数量</span>
                        <span role="columnheader">详情</span>
                      </div>
                      {paginatedEntries.map((entry) => {
                        const firstChange = entry.changes[0];
                        const rowName = rowNameFor(entry);
                        return (
                          <div className={`compact-result-row status-${entry.status}`} role="row" key={`${entry.status}-${entry.key}`}>
                            <span role="cell"><span className="status-pill">{CANDIDATE_STATUS_LABEL[entry.status]}</span></span>
                            <code role="cell" title={entry.key}>{entry.key}</code>
                            <span className="compact-name" role="cell" title={rowName}>{rowName || "—"}</span>
                            <span className="compact-fields" role="cell" title={entry.changes.map((change) => change.field).join("、")}>
                              {changedFieldSummary(entry)}
                            </span>
                            <span className="compact-first-change" role="cell">
                              {entry.status === "changed" && firstChange ? (
                                <>
                                  <span className="compact-old" title={firstChange.before}>{shownValue(firstChange.before)}</span>
                                  <span aria-hidden="true">→</span>
                                  <span className="compact-new" title={firstChange.after}>{shownValue(firstChange.after)}</span>
                                </>
                              ) : (
                                <span className="compact-description">{describeDifference(entry)}</span>
                              )}
                            </span>
                            <strong role="cell">{entry.status === "changed" ? entry.changes.length : "—"}</strong>
                            <span role="cell">
                              <button type="button" className="view-detail-button" onClick={() => setSelectedEntry(entry)}>查看</button>
                            </span>
                          </div>
                        );
                      })}
                    </div>
                  </div>

                  <div className="pagination-bar">
                    <span>第 {safeCurrentPage} / {totalPages} 页</span>
                    <div>
                      <button type="button" disabled={safeCurrentPage <= 1} onClick={() => setCurrentPage((page) => Math.max(1, page - 1))}>上一页</button>
                      <button type="button" disabled={safeCurrentPage >= totalPages} onClick={() => setCurrentPage((page) => Math.min(totalPages, page + 1))}>下一页</button>
                    </div>
                  </div>
                </div>
                </>
              ) : (
                <p className="empty-state compact-empty">没有符合当前筛选条件的结果。</p>
              )}

              <DifferenceDrawer entry={selectedEntry} keyField={keyField} onClose={() => setSelectedEntry(null)} />
            </>
          ) : (
            <>
              <div className="summary-grid">
                {([
                  ["changed", "有变化", comparison.summary.changed],
                  ["added", "新增", comparison.summary.added],
                  ["removed", "删除", comparison.summary.removed],
                  ["unchanged", "相同", comparison.summary.unchanged],
                  ["duplicate", "需处理", comparison.summary.issues],
                ] as const).map(([status, label, count]) => (
                  <button
                    type="button"
                    key={status}
                    className={`summary-card status-${status} ${filter === status ? "is-selected" : ""}`}
                    onClick={() => setFilter(filter === status ? "all" : status)}
                  >
                    <span>{label}</span>
                    <strong>{count}</strong>
                  </button>
                ))}
              </div>

              {comparison.summary.issues ? (
                <p className="issue-note">
                  有 {comparison.summary.duplicate} 个重复标识；旧文件 {comparison.emptyKeyRows.before} 行、新文件 {comparison.emptyKeyRows.after} 行缺少匹配值。重复或空标识不会被当作正常变化。
                </p>
              ) : null}

              <div className="result-toolbar">
                <div className="filter-tabs" role="group" aria-label="结果筛选">
                  {(["all", "changed", "added", "removed", "unchanged", "duplicate"] as const).map((status) => (
                    <button
                      type="button"
                      key={status}
                      className={filter === status ? "is-active" : ""}
                      onClick={() => setFilter(status)}
                    >
                      {status === "all" ? `全部 ${comparison.summary.total}` : STATUS_LABEL[status]}
                    </button>
                  ))}
                </div>
                <label className="search-field">
                  <span className="sr-only">搜索结果</span>
                  <input value={query} onChange={(event) => setQuery(event.target.value)} placeholder={`搜索 ${keyField}、字段或内容`} />
                </label>
              </div>

              <div className="result-list" aria-live="polite">
                {visibleEntries.length ? visibleEntries.map((entry) => (
                  <details className={`result-item status-${entry.status}`} key={`${entry.status}-${entry.key}`}>
                    <summary>
                      <span className="status-pill">{STATUS_LABEL[entry.status]}</span>
                      <code>{entry.key}</code>
                      <span className="change-labels">
                        {entry.changes.slice(0, 3).map((change) => <span key={change.field}>{change.field}</span>)}
                        {entry.changes.length > 3 ? <span>+{entry.changes.length - 3}</span> : null}
                      </span>
                      <span className="result-description">{describeDifference(entry)}</span>
                      <span className="disclosure">展开</span>
                    </summary>
                    <div className="change-table">
                      {entry.status === "duplicate" ? (
                        <div className="duplicate-message">请先在源文件中处理重复的 {keyField}，再重新比较。</div>
                      ) : entry.changes.length ? entry.changes.map((change) => (
                        <div className="change-row" key={change.field}>
                          <strong>{change.field}</strong>
                          <div><span>旧值</span><p>{shownValue(change.before)}</p></div>
                          <div><span>新值</span><p>{shownValue(change.after)}</p></div>
                        </div>
                      )) : (
                        <div className="duplicate-message">所选字段没有差异。</div>
                      )}
                    </div>
                  </details>
                )) : (
                  <p className="empty-state">没有符合当前筛选条件的结果。</p>
                )}
              </div>
            </>
          )}
        </section>
      ) : null}

      <footer>本地处理 · 不上传文件 · 关闭页面后不保留表格内容</footer>
    </main>
  );
}
