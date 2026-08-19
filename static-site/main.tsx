import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { TableDiffTool } from "../app/table-diff-tool";
import "../app/globals.css";

const root = document.getElementById("root");

if (!root) {
  throw new Error("页面缺少应用挂载节点");
}

const resultLayout =
  document.documentElement.dataset.resultLayout === "candidate"
    ? "candidate"
    : "classic";

createRoot(root).render(
  <StrictMode>
    <TableDiffTool resultLayout={resultLayout} />
  </StrictMode>,
);
