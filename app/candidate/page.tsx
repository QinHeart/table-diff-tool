import type { Metadata } from "next";
import { TableDiffTool } from "../table-diff-tool";

export const metadata: Metadata = {
  title: "表格差异对比 · 结果展示候选版",
  description: "更适合初次使用者的表格差异结果展示候选版。",
};

export default function CandidatePage() {
  return <TableDiffTool resultLayout="candidate" />;
}
