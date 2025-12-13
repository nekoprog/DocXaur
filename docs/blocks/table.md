# Table

Build `<w:tbl>` with column grid and cell-level styling.

## Create

```ts
const table = section.table({
  columns: [
    { width: "4cm", hAlign: "left", bold: true },
    { width: "3cm", hAlign: "right", fontSize: 12 },
    { width: "3cm", hAlign: "center" },
  ],
  borders: true,
});
```

## Rows

```ts
table.row("Name", "Qty", "Total");

table.row(
  { text: "Widget A", fontColor: "0066CC" },
  { text: "3", hAlign: "center", bold: true },
  { text: "$30", bold: true },
);
```

## Cell Styling (per cell)

- `text`: string
- `fontName`, `fontSize`, `fontColor`, `bold`, `italic`, `underline`
- `hAlign`: "left" | "center" | "right" | "justify"
- `vAlign`: "top" | "center" | "bottom"
- `cellColor`: hex without `#`
- `height`: string (e.g., `"1cm"`)
- `colspan`, `rowspan` (`rowspan: 0` = continuation)

## Examples

### Header Row with Color

```ts
table.row(
  { text: "Item", bold: true, cellColor: "4472C4", fontColor: "FFFFFF" },
  { text: "Qty", bold: true, cellColor: "4472C4", fontColor: "FFFFFF" },
  { text: "Total", bold: true, cellColor: "4472C4", fontColor: "FFFFFF" },
);
```

### Rowspan / Colspan

```ts
table.row(
  { text: "Group A", rowspan: 2, vAlign: "center", bold: true },
  { text: "Mon" },
  { text: "Tue" },
);
table.row(
  { text: "", rowspan: 0 }, // continuation
  { text: "Wed" },
  { text: "Thu" },
);
```
