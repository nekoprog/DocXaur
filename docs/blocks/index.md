# Blocks Overview

Blocks are the **OOXML renderables** you compose inside sections:

- **Section** — Parent container for blocks; emits `<w:sectPr>`
- **Paragraph** — `<w:p>` with inline styling
- **Image** — Inline DrawingML picture referencing a relationship
- **Table** — `<w:tbl>` with grid, rows, and cells (cell-level styling)

## Typical Usage

```ts
const doc = new DocXaur();
const section = doc.addSection();

section.heading("Report", 1);
section.paragraph().text("Introduction…");

await section.image("/images/logo.png", { width: "5cm", align: "left" });

const table = section.table({
  columns: [{ width: "4cm" }, { width: "3cm" }, { width: "3cm" }],
  borders: true,
});

table.row("Item", "Qty", "Total");
table.row("Widget A", "3", "$30");

await doc.download("report.docx");
```

> Future blocks: **headers, footers, bookmarks**, and more.
