# 表格差异对比

在浏览器本地比较两份 Excel、CSV 或 TSV 文件，快速找出新增、删除、字段变化和完全相同的记录。

## 在线使用

- [打开表格差异对比](https://qinheart.github.io/table-diff-tool/)
- [打开适合大量数据的结果视图](https://qinheart.github.io/table-diff-tool/candidate/)

上传的文件只在当前浏览器中读取，不会发送到 GitHub 或其他服务器。

## 当前功能

- 上传或拖入新旧两份表格
- 多工作表选择
- 自动建议 SKU、ASIN、订单号等匹配列
- 自由选择需要比较的字段
- 分类显示新增、删除、变化、相同和重复标识
- 搜索、筛选并展开查看字段级旧值与新值
- 导出差异摘要和字段变化 Excel
- 保留作为文本存储或格式化显示的标识前导零
- 对空匹配值和重复匹配值给出单独警告
- 文件只在当前浏览器中处理，不上传、不保存

页面内置完全虚构的商品资料示例，便于直接体验完整流程。

## 本地运行

需要 Node.js 22.13 或更高版本。

```bash
npm install
npm run dev
```

打开终端显示的本地地址，通常为 <http://localhost:3000/>。

## 检查

```bash
npm run lint
npm test
```

## 更新 GitHub Pages

```bash
npm run build:pages
git add docs
git commit -m "Update GitHub Pages"
git push
```

GitHub Pages 直接发布 `main` 分支的 `docs` 目录，不依赖其他托管服务。

表格读取和导出使用 [SheetJS Community Edition](https://docs.sheetjs.com/)；打包完成后，实际处理不依赖第三方上传服务。
