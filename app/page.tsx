import type { Metadata } from "next";
import { TableDiffTool } from "./table-diff-tool";

export const metadata: Metadata = {
  title: "表格差异对比",
  description: "在浏览器本地对比两份 Excel 或 CSV，快速找出新增、删除和字段变化。",
};

export default function Home() {
  return <TableDiffTool />;
}
