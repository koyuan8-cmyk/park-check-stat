2026-08-04: v3.5 — 页面设置 + 页眉 + 行距 + 布局全面优化。
  - 页面设置：每个 Sheet 强制 A4、适应1页宽×1页高 (fitToPage="1")、水平居中、垂直不居中、统一页边距 0.4"/0.5"。
  - 页眉：浮动图片 (DrawingML, openpyxl 精确格式) + VML legacyDrawingHF 双通道。`_add_logo_image()` 将桌面 logo.jpg 嵌入每个 Sheet 的 A1 位置。
  - 行距：所有 Sheet 数据行 28pt、表头 32pt、合计 30pt。商家计费表同步加宽。
  - 布局：每个 Sheet 第 1 行放 Logo (48pt)，第 2-3 行空白间距 (10pt)，数据从第 4 行开始。所有函数生成的行号统一 +2。
  - `_add_header_logo_to_xlsx` 全部改为字符串操作，不再用 ElementTree，避免 XML 结构破坏。
  - HTML 版 `addHeaderToXlsxBuf` 完整重写，4 步流程对应 Python：页面设置强制替换 → VML 页眉 → DrawingML 浮动图 → Content_Types。
  - `addPageFitToXlsxBuf` 废弃，pageSetUpPr 匹配 `<sheetPr/>` 自闭合标签，消除 SheetJS 兼容性问题。
  文件：generate_branch.py, SKILL.md, 停车票核销报表生成器-v3.5.html
2026-08-05: v3.5 HTML 修复 + 封装为 skill。
  - 修复1：车场报表(3个sheet) 第4/5行标题/期间未居中、未加大字体 — scell 行号 1,2 → 3,4
  - 修复2：页眉多余 "&G" 文字 — headerFooter 去掉 &G，只保留 &L（VML 图片不需要 &G 指令）
  - 修复3：车场报表第5行日期取最早月份 → 取最晚月份
  - Skill: parking-report-v3.4 → parking-report-v3.5，目录重命名
  - 文件：SKILL.md, 停车票核销报表生成器-v3.5.html, generate_branch.py
2026-08-03: HTML 版页眉修正——移除 headerFooter &L&G 注入与 VML ClientData 锚点（对齐 Python 版），文件：~/Desktop/停车票核销报表生成器-v3.4.html
