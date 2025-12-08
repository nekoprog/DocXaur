# DocXaur Documentation

A semantic, consistent DOCX generation library for Deno Fresh Islands.

## ⚠️ Important Notice

**This library ONLY works in Fresh Islands (browser environment)**. It will NOT
work in server-side routes, middleware, or Deno runtime.

## Installation

```typescript
import { DocXaur } from "jsr:@yourscope/docxaur";
```

---

## Table of Contents

1. [Quick Start](#quick-start)
2. [DocXaur Class](#docxaur-class)
3. [Section Class](#section-class)
4. [Paragraph Class](#paragraph-class)
5. [Table Class](#table-class)
6. [Advanced Examples](#advanced-examples)

---

## Quick Start

```typescript
// islands/DocumentGenerator.tsx
import { DocXaur } from "jsr:@yourscope/docxaur";

export default function DocumentGenerator() {
  const generateDoc = async () => {
    const doc = new DocXaur({ title: "My Document" });
    const section = doc.addSection();

    section.heading("Hello World");
    section.paragraph().text("This is a paragraph.");

    await doc.download("my-document.docx");
  };

  return <button onClick={generateDoc}>Generate Document</button>;
}
```

---

## DocXaur Class

Main class for creating DOCX documents.

### Constructor

```typescript
new DocXaur(options?: DocumentOptions)
```

#### DocumentOptions

| Option        | Type     | Default      | Description            |
| ------------- | -------- | ------------ | ---------------------- |
| `title`       | `string` | `"Document"` | Document title         |
| `creator`     | `string` | `"DocXaur"`  | Document creator       |
| `description` | `string` | `""`         | Document description   |
| `subject`     | `string` | `""`         | Document subject       |
| `keywords`    | `string` | `""`         | Document keywords      |
| `fontName`    | `string` | `"Calibri"`  | Default font family    |
| `fontSize`    | `number` | `11`         | Default font size (pt) |

#### Example

```typescript
const doc = new DocXaur({
  title: "Sales Report Q4 2024",
  creator: "John Doe",
  description: "Quarterly sales analysis",
  subject: "Sales",
  keywords: "sales, Q4, 2024, report",
  fontName: "Arial",
  fontSize: 12,
});
```

### Methods

#### `addSection(options?: SectionOptions): Section`

Adds a new section to the document.

**SectionOptions:**

| Option     | Type       | Default          | Description     |
| ---------- | ---------- | ---------------- | --------------- |
| `pageSize` | `PageSize` | A4 Portrait      | Page dimensions |
| `margins`  | `Margins`  | 2.54cm all sides | Page margins    |

**PageSize:**

| Property      | Type                        | Default      | Description      |
| ------------- | --------------------------- | ------------ | ---------------- |
| `width`       | `number`                    | `21`         | Page width (cm)  |
| `height`      | `number`                    | `29.7`       | Page height (cm) |
| `orientation` | `"portrait" \| "landscape"` | `"portrait"` | Page orientation |

**Margins:**

| Property | Type     | Default | Description        |
| -------- | -------- | ------- | ------------------ |
| `top`    | `number` | `2.54`  | Top margin (cm)    |
| `right`  | `number` | `2.54`  | Right margin (cm)  |
| `bottom` | `number` | `2.54`  | Bottom margin (cm) |
| `left`   | `number` | `2.54`  | Left margin (cm)   |

**Example:**

```typescript
// Default A4 Portrait
const section1 = doc.addSection();

// Custom Landscape with narrow margins
const section2 = doc.addSection({
  pageSize: {
    width: 29.7,
    height: 21,
    orientation: "landscape",
  },
  margins: {
    top: 1.5,
    right: 1.5,
    bottom: 1.5,
    left: 1.5,
  },
});
```

#### `async download(filename?: string): Promise<void>`

Downloads the document to the user's computer.

**Parameters:**

- `filename` (optional): Name of the file. Default: `"document.docx"`

**Example:**

```typescript
await doc.download("my-report.docx");
```

#### `async toBlob(): Promise<Blob>`

Returns the document as a Blob for custom handling.

**Example:**

```typescript
const blob = await doc.toBlob();
// Upload to server, etc.
```

#### `getDefaultFont(): string`

Returns the default font name.

#### `getDefaultSize(): number`

Returns the default font size.

---

## Section Class

Represents a section in the document. Contains paragraphs, headings, images,
tables, and breaks.

### Methods

#### `heading(content: string, level?: 1|2|3|4|5|6, options?: ParagraphOptions): this`

Adds a heading to the section.

**Parameters:**

- `content`: Heading text
- `level`: Heading level (1-6). Default: `1`
- `options`: Additional paragraph options

**Default sizes:**

- Level 1: 24pt
- Level 2: 20pt
- Level 3: 18pt
- Level 4: 16pt
- Level 5: 14pt
- Level 6: 12pt

**Example:**

```typescript
section.heading("Chapter 1: Introduction", 1);
section.heading("Section 1.1", 2);
section.heading("Subsection 1.1.1", 3, { color: "FF0000" });
```

#### `paragraph(options?: ParagraphOptions): Paragraph`

Creates a new paragraph builder.

**ParagraphOptions:**

| Option           | Type                                         | Default          | Description                 |
| ---------------- | -------------------------------------------- | ---------------- | --------------------------- |
| `align`          | `"left" \| "center" \| "right" \| "justify"` | `"left"`         | Text alignment              |
| `bold`           | `boolean`                                    | `false`          | Bold text                   |
| `italic`         | `boolean`                                    | `false`          | Italic text                 |
| `underline`      | `boolean`                                    | `false`          | Underlined text             |
| `fontSize`       | `number`                                     | Document default | Font size (pt)              |
| `fontColor`      | `string`                                     | `"000000"`       | Font color (hex without #)  |
| `fontName`       | `string`                                     | Document default | Font family                 |
| `spacing.before` | `number`                                     | `0`              | Space before paragraph (pt) |
| `spacing.after`  | `number`                                     | `0`              | Space after paragraph (pt)  |
| `spacing.line`   | `number`                                     | `1.0`            | Line spacing multiplier     |
| `breakBefore`    | `number`                                     | `0`              | Line breaks before content  |
| `breakAfter`     | `number`                                     | `0`              | Line breaks after content   |

**Example:**

```typescript
section.paragraph({ align: "center", fontSize: 14 })
  .text("Hello ")
  .text("World", { bold: true, color: "FF0000" });
```

#### `async image(url: string, options?: ImageOptions): Promise<this>`

Adds an image from a URL.

**ImageOptions:**

| Option   | Type                            | Default       | Description                       |
| -------- | ------------------------------- | ------------- | --------------------------------- |
| `width`  | `string`                        | `"10cm"`      | Image width (cm, pt, mm, in, px)  |
| `height` | `string`                        | Auto (square) | Image height (cm, pt, mm, in, px) |
| `align`  | `"left" \| "center" \| "right"` | `"center"`    | Image alignment                   |

**Supported URLs:**

- HTTP/HTTPS URLs: `https://example.com/image.jpg`
- Absolute paths (from Fresh static folder): `/images/logo.png`

**Example:**

```typescript
// From URL
await section.image("https://example.com/photo.jpg", {
  width: "15cm",
  align: "center",
});

// From static folder (place in static/images/)
await section.image("/images/logo.png", {
  width: "5cm",
  height: "5cm",
});
```

#### `table(options: TableOptions): Table`

Creates a table builder.

**TableOptions:**

| Option    | Type                            | Required | Description                           |
| --------- | ------------------------------- | -------- | ------------------------------------- |
| `columns` | `TableColumn[]`                 | Yes      | Column definitions                    |
| `width`   | `string`                        | No       | Total table width                     |
| `align`   | `"left" \| "center" \| "right"` | No       | Table alignment (default: `"center"`) |
| `borders` | `boolean`                       | No       | Show borders (default: `true`)        |

**TableColumn:**

| Property    | Type                            | Required | Description                                         |
| ----------- | ------------------------------- | -------- | --------------------------------------------------- |
| `width`     | `string`                        | Yes      | Column width (cm, pt, mm, in, %)                    |
| `hAlign`    | `"left" \| "center" \| "right"` | No       | Default horizontal alignment for column             |
| `vAlign`    | `"top" \| "center" \| "bottom"` | No       | Default vertical alignment for column               |
| `fontName`  | `string`                        | No       | Default font family for column                      |
| `fontSize`  | `number`                        | No       | Default font size for column (pt)                   |
| `fontColor` | `string`                        | No       | Default text color for column (hex without #)       |
| `cellColor` | `string`                        | No       | Default background color for column (hex without #) |
| `bold`      | `boolean`                       | No       | Default bold styling for column                     |
| `italic`    | `boolean`                       | No       | Default italic styling for column                   |
| `underline` | `boolean`                       | No       | Default underline styling for column                |

**Example:**

```typescript
const table = section.table({
  columns: [
    {
      width: "3cm",
      hAlign: "left",
      vAlign: "center",
      bold: true,
    },
    {
      width: "8cm",
      hAlign: "left",
      fontName: "Arial",
      fontSize: 11,
    },
    {
      width: "3cm",
      hAlign: "right",
      fontColor: "0066CC",
      bold: true,
    },
  ],
  borders: true,
  align: "center",
});

// Simple rows inherit column defaults
table.row("Name", "Description", "$100");

// Individual cells can override column defaults
table.row(
  "Product A",
  { text: "Special item", fontColor: "FF0000" }, // Override column fontColor
  { text: "$200", bold: false }, // Override column bold
);
```

#### `lineBreak(count?: number): this`

Adds line breaks.

**Parameters:**

- `count`: Number of line breaks. Default: `1`

**Example:**

```typescript
section.lineBreak(2); // Add 2 line breaks
```

#### `pageBreak(count?: number): this`

Adds page breaks.

**Parameters:**

- `count`: Number of page breaks. Default: `1`

**Example:**

```typescript
section.pageBreak(); // Start new page
```

---

## Paragraph Class

Builder for creating paragraphs with mixed inline styles.

### Methods

#### `text(text: string, style?: TextStyle): this`

Adds text with optional inline styling.

**TextStyle:**

| Property    | Type      | Description                |
| ----------- | --------- | -------------------------- |
| `bold`      | `boolean` | Bold text                  |
| `italic`    | `boolean` | Italic text                |
| `underline` | `boolean` | Underlined text            |
| `size`      | `number`  | Font size (pt)             |
| `color`     | `string`  | Font color (hex without #) |
| `font`      | `string`  | Font family                |

**Example:**

```typescript
section.paragraph()
  .text("Normal text ")
  .text("bold text ", { bold: true })
  .text("red italic text", { italic: true, color: "FF0000" });
```

#### `tab(): this`

Adds a tab character.

**Example:**

```typescript
section.paragraph()
  .text("Name:")
  .tab()
  .text("John Doe");
```

#### `lineBreak(count?: number): this`

Adds line breaks within the paragraph.

**Example:**

```typescript
section.paragraph()
  .text("First line")
  .lineBreak()
  .text("Second line")
  .lineBreak(2)
  .text("Fourth line");
```

#### `pageBreak(count?: number): this`

Adds a page break within the paragraph.

**Example:**

```typescript
section.paragraph()
  .text("End of page 1")
  .pageBreak()
  .text("Start of page 3")
  .pageBreak(2)
  .text("Start of page 6");
```

#### `apply(...operations: ((builder: this) => this)[]): this`

Conditionally applies operations.

**Example:**

```typescript
const showDate = true;
const showTime = false;

section.paragraph()
  .text("Report generated")
  .apply(
    ...showDate ? [(p) => p.text(" on 2024-01-15")] : [],
    ...showTime ? [(p) => p.text(" at 10:30 AM")] : [],
  );
```

---

## Table Class

Builder for creating tables with rows and cells. **All styling options are
applied at the cell level.**

### Methods

#### `row(...cells: (string | TableCellData)[]): this`

Adds a row to the table. Each cell can have its own individual styling.

**Simple usage (strings):**

```typescript
// Simple text cells (uses column defaults for alignment)
table.row("Cell 1", "Cell 2", "Cell 3");
```

**Advanced usage (TableCellData objects):**

```typescript
// Each cell with its own styling
table.row(
  { text: "Name", bold: true, cellColor: "4472C4", fontColor: "FFFFFF" },
  { text: "Age", bold: true, fontSize: 12, fontName: "Arial" },
  { text: "City", italic: true, hAlign: "left" },
);
```

**TableCellData Properties:**

| Property    | Type                            | Required | Description                                              |
| ----------- | ------------------------------- | -------- | -------------------------------------------------------- |
| `text`      | `string`                        | ✅ Yes   | Cell content                                             |
| `fontName`  | `string`                        | No       | Font family (e.g., "Arial", "Times New Roman")           |
| `fontSize`  | `number`                        | No       | Font size in points (e.g., 12)                           |
| `fontColor` | `string`                        | No       | Text color (hex without #, e.g., "FF0000")               |
| `cellColor` | `string`                        | No       | Background color (hex without #, e.g., "E0E0E0")         |
| `bold`      | `boolean`                       | No       | Bold text                                                |
| `italic`    | `boolean`                       | No       | Italic text                                              |
| `underline` | `boolean`                       | No       | Underlined text                                          |
| `hAlign`    | `"left" \| "center" \| "right"` | No       | Horizontal alignment (default: column align or "center") |
| `vAlign`    | `"top" \| "center" \| "bottom"` | No       | Vertical alignment (default: "center")                   |
| `colspan`   | `number`                        | No       | Number of columns to span                                |
| `rowspan`   | `number`                        | No       | Number of rows to span (use 0 for continuation)          |
| `height`    | `string`                        | No       | Cell height (e.g., "1cm", "20pt")                        |

**Examples:**

```typescript
// Header row with blue background and white text
table.row(
  { text: "Name", bold: true, cellColor: "4472C4", fontColor: "FFFFFF" },
  { text: "Age", bold: true, cellColor: "4472C4", fontColor: "FFFFFF" },
  { text: "City", bold: true, cellColor: "4472C4", fontColor: "FFFFFF" },
);

// Data row with mixed styles
table.row(
  "John Doe",
  { text: "25", hAlign: "center", bold: true },
  { text: "New York", fontColor: "666666" },
);

// Row with custom fonts and sizes
table.row(
  { text: "Total", fontName: "Arial", fontSize: 14, bold: true },
  { text: "$1,250", fontName: "Courier New", fontSize: 12 },
  { text: "USD", fontSize: 10, italic: true },
);

// Row with cell colors
table.row(
  { text: "Warning", cellColor: "FFFF00", bold: true },
  { text: "Error", cellColor: "FF0000", fontColor: "FFFFFF", bold: true },
  { text: "Success", cellColor: "00FF00", bold: true },
);

// Row with different heights
table.row(
  { text: "Tall cell", height: "2cm", vAlign: "top" },
  { text: "Normal", height: "1cm" },
  { text: "Short", height: "0.5cm" },
);
```

#### `apply(...operations: ((builder: this) => this)[]): this`

Conditionally applies operations to the table.

**Example:**

```typescript
const includeTotal = true;

table
  .row("Item", "Price")
  .row("Apple", "$1.50")
  .row("Banana", "$0.75")
  .apply(
    ...includeTotal
      ? [
        (t) =>
          t.row(
            { text: "Total", bold: true, cellColor: "FFFF00" },
            { text: "$2.25", bold: true, cellColor: "FFFF00" },
          ),
      ]
      : [],
  );
```

---

## Advanced Examples

### Example 1: Sales Report with Styled Table

```typescript
const doc = new DocXaur({
  title: "Sales Report Q4 2024",
  creator: "Sales Team",
  fontName: "Arial",
  fontSize: 11,
});

const section = doc.addSection();

// Title
section.heading("Q4 2024 Sales Report", 1, { align: "center" });

// Subtitle
section.paragraph({ align: "center", size: 12, color: "666666" })
  .text("Generated on ")
  .text("January 15, 2025", { bold: true });

section.lineBreak(2);

// Introduction
section.heading("Executive Summary", 2);
section.paragraph()
  .text("Total revenue for Q4 2024 reached ")
  .text("$1.2M", { bold: true, color: "00AA00" })
  .text(", representing a ")
  .text("15% increase", { bold: true })
  .text(" over Q3 2024.");

section.lineBreak();

// Sales Table with column defaults
section.heading("Sales by Region", 2);

const table = section.table({
  columns: [
    { width: "4cm", hAlign: "left" },
    { width: "3cm", hAlign: "right" },
    { width: "3cm", hAlign: "right" },
    { width: "3cm", hAlign: "right" },
  ],
  borders: true,
  align: "center",
});

// Header row - blue background with white text
table.row(
  { text: "Region", bold: true, cellColor: "4472C4", fontColor: "FFFFFF" },
  { text: "Q3 2024", bold: true, cellColor: "4472C4", fontColor: "FFFFFF" },
  { text: "Q4 2024", bold: true, cellColor: "4472C4", fontColor: "FFFFFF" },
  { text: "Growth", bold: true, cellColor: "4472C4", fontColor: "FFFFFF" },
);

// Data rows
table.row("North America", "$450K", "$520K", "+15.6%");
table.row("Europe", "$380K", "$425K", "+11.8%");
table.row("Asia Pacific", "$280K", "$355K", "+26.8%");

// Total row - light blue background with bold text
table.row(
  { text: "Total", bold: true, cellColor: "D9E1F2" },
  { text: "$1.11M", bold: true, cellColor: "D9E1F2" },
  { text: "$1.30M", bold: true, cellColor: "D9E1F2" },
  { text: "+17.1%", bold: true, fontColor: "00AA00", cellColor: "D9E1F2" },
);

await doc.download("sales-report-q4-2024.docx");
```

### Example 2: Invoice with Custom Cell Styling

```typescript
const doc = new DocXaur({ title: "Invoice" });
const section = doc.addSection();

// Company logo
await section.image("/images/company-logo.png", {
  width: "5cm",
  align: "left",
});

section.lineBreak();

// Invoice header
section.paragraph({ align: "right", size: 16, bold: true })
  .text("INVOICE");

section.paragraph({ align: "right", size: 10 })
  .text("Invoice #: ")
  .text("INV-2024-001", { bold: true })
  .lineBreak()
  .text("Date: ")
  .text("January 15, 2025", { bold: true });

section.lineBreak(2);

// Bill to
section.heading("Bill To:", 3);
section.paragraph()
  .text("Acme Corporation")
  .lineBreak()
  .text("123 Business St")
  .lineBreak()
  .text("New York, NY 10001");

section.lineBreak(2);

// Items table with column defaults
const table = section.table({
  columns: [
    { width: "1.5cm", hAlign: "center" },
    { width: "7cm", hAlign: "left" },
    { width: "2cm", hAlign: "right" },
    { width: "2cm", hAlign: "right" },
    { width: "2.5cm", hAlign: "right", bold: true },
  ],
  borders: true,
});

// Header with gray background
table.row(
  { text: "#", bold: true, cellColor: "E0E0E0" },
  { text: "Description", bold: true, cellColor: "E0E0E0" },
  { text: "Qty", bold: true, cellColor: "E0E0E0" },
  { text: "Price", bold: true, cellColor: "E0E0E0" },
  { text: "Total", bold: true, cellColor: "E0E0E0" },
);

// Data rows (last column automatically bold due to column default)
table.row("1", "Web Development Services", "40", "$100", "$4,000");
table.row("2", "UI/UX Design", "20", "$120", "$2,400");
table.row("3", "Hosting (Annual)", "1", "$500", "$500");

// Subtotal row
table.row(
  { text: "", colspan: 4, cellColor: "F0F0F0" },
  { text: "Subtotal:", cellColor: "F0F0F0" },
);
table.row(
  { text: "", colspan: 4 },
  "$6,900",
);

// Total row with blue background and white text
table.row(
  { text: "", colspan: 4, cellColor: "4472C4" },
  { text: "TOTAL:", fontColor: "FFFFFF", cellColor: "4472C4" },
);
table.row(
  { text: "", colspan: 4, cellColor: "4472C4" },
  { text: "$6,900", fontSize: 14, fontColor: "FFFFFF", cellColor: "4472C4" },
);

section.lineBreak(2);
section.paragraph({ size: 9, color: "666666" })
  .text("Payment due within 30 days. Thank you for your business!");

await doc.download("invoice-2024-001.docx");
```

### Example 3: Product Comparison with Column Defaults

```typescript
const doc = new DocXaur({ title: "Product Comparison" });
const section = doc.addSection();

section.heading("Product Feature Comparison");

// Use column defaults for consistent styling
const table = section.table({
  columns: [
    {
      width: "4cm",
      hAlign: "left",
      bold: true,
    },
    {
      width: "3cm",
      hAlign: "center",
      fontName: "Courier New",
      fontSize: 11,
    },
    {
      width: "3cm",
      hAlign: "center",
      fontName: "Courier New",
      fontSize: 11,
      fontColor: "0066CC",
    },
    {
      width: "3cm",
      hAlign: "center",
      fontName: "Courier New",
      fontSize: 11,
      fontColor: "00AA00",
      bold: true,
    },
  ],
});

// Header row
table.row(
  { text: "Feature", cellColor: "4472C4", fontColor: "FFFFFF" },
  { text: "Basic", cellColor: "4472C4", fontColor: "FFFFFF", bold: true },
  { text: "Pro", cellColor: "4472C4", fontColor: "FFFFFF", bold: true },
  { text: "Enterprise", cellColor: "4472C4", fontColor: "FFFFFF" },
);

// Simple rows automatically get column formatting
table.row("Price", "$10/mo", "$30/mo", "$100/mo");
table.row("Users", "1", "5", "Unlimited");
table.row("Storage", "10GB", "100GB", "1TB");
table.row("Support", "Email", "Priority", "24/7 Phone");

await doc.download("product-comparison.docx");
```

### Example 4: Complex Table with Rowspan/Colspan

```typescript
const doc = new DocXaur({ title: "Complex Table" });
const section = doc.addSection();

section.heading("Employee Schedule");

const table = section.table({
  columns: [
    { width: "3cm" },
    { width: "2.5cm" },
    { width: "2.5cm" },
    { width: "2.5cm" },
    { width: "2.5cm" },
  ],
});

// Header with colspan and custom styling
table.row(
  {
    text: "Employee",
    bold: true,
    cellColor: "4472C4",
    fontColor: "FFFFFF",
    rowspan: 2,
    vAlign: "center",
  },
  {
    text: "Work Schedule",
    bold: true,
    cellColor: "4472C4",
    fontColor: "FFFFFF",
    colspan: 4,
    hAlign: "center",
  },
);

// Days header (continuation of rowspan)
table.row(
  { text: "", rowspan: 0 },
  { text: "Mon", bold: true, cellColor: "D9E1F2" },
  { text: "Tue", bold: true, cellColor: "D9E1F2" },
  { text: "Wed", bold: true, cellColor: "D9E1F2" },
  { text: "Thu", bold: true, cellColor: "D9E1F2" },
);

// Employee data
table.row(
  { text: "John Doe", bold: true },
  { text: "9-5", fontSize: 10 },
  { text: "9-5", fontSize: 10 },
  { text: "OFF", cellColor: "FFCCCC", bold: true },
  { text: "9-5", fontSize: 10 },
);

table.row(
  { text: "Jane Smith", bold: true },
  { text: "10-6", fontSize: 10 },
  { text: "10-6", fontSize: 10 },
  { text: "10-6", fontSize: 10 },
  { text: "OFF", cellColor: "FFCCCC", bold: true },
);

await doc.download("employee-schedule.docx");
```

### Example 5: Multi-font Report with Column Defaults

```typescript
const doc = new DocXaur({ title: "Typography Showcase" });
const section = doc.addSection();

section.heading("Font Styles Demonstration");

// First column bold, second column uses different fonts
const table = section.table({
  columns: [
    { width: "4cm", hAlign: "left", bold: true },
    { width: "10cm", hAlign: "left" },
  ],
  borders: true,
});

table.row(
  { text: "Font Family", cellColor: "E0E0E0" },
  { text: "Sample Text", cellColor: "E0E0E0" },
);

// Override fontName per cell, bold inherited from column 1
table.row(
  { text: "Arial", fontName: "Arial" },
  { text: "The quick brown fox jumps over the lazy dog", fontName: "Arial" },
);

table.row(
  { text: "Times New Roman", fontName: "Times New Roman" },
  {
    text: "The quick brown fox jumps over the lazy dog",
    fontName: "Times New Roman",
  },
);

table.row(
  { text: "Calibri", fontName: "Calibri" },
  { text: "The quick brown fox jumps over the lazy dog", fontName: "Calibri" },
);

table.row(
  { text: "Courier New", fontName: "Courier New" },
  {
    text: "The quick brown fox jumps over the lazy dog",
    fontName: "Courier New",
  },
);

// Size variations table
section.lineBreak(2);
section.heading("Font Size Variations", 2);

const sizeTable = section.table({
  columns: [
    {
      width: "3cm",
      hAlign: "center",
      bold: true,
      cellColor: "F0F0F0",
    },
    { width: "11cm", hAlign: "left" },
  ],
});

sizeTable.row(
  { text: "Size", cellColor: "E0E0E0" },
  { text: "Sample", cellColor: "E0E0E0" },
);

sizeTable.row(
  { text: "8pt", fontSize: 8 },
  { text: "This is 8pt text", fontSize: 8 },
);

sizeTable.row(
  { text: "10pt", fontSize: 10 },
  { text: "This is 10pt text", fontSize: 10 },
);

sizeTable.row(
  { text: "12pt", fontSize: 12 },
  { text: "This is 12pt text", fontSize: 12 },
);

sizeTable.row(
  { text: "14pt", fontSize: 14 },
  { text: "This is 14pt text", fontSize: 14 },
);

sizeTable.row(
  { text: "16pt", fontSize: 16 },
  { text: "This is 16pt text", fontSize: 16 },
);

await doc.download("typography-showcase.docx");
```

---

## Tips & Best Practices

### 1. Always Use Fresh Islands

```typescript
// ✅ CORRECT - Fresh Island
// islands/MyDocumentGenerator.tsx
import { DocXaur } from "jsr:@yourscope/docxaur";

export default function MyDocumentGenerator() {
  const generate = async () => {
    const doc = new DocXaur();
    // ...
  };
  return <button onClick={generate}>Generate</button>;
}

// ❌ WRONG - Server-side route
// routes/api/generate.ts
import { DocXaur } from "jsr:@yourscope/docxaur";
// This will throw an error!
```

### 2. Use Semantic Method Chaining

```typescript
// Readable and maintainable
section
  .heading("Chapter 1")
  .paragraph().text("Introduction paragraph")
  .lineBreak()
  .paragraph().text("Second paragraph");
```

### 3. Store Images in Static Folder

```
your-fresh-app/
├── static/
│   └── images/
│       ├── logo.png
│       └── banner.jpg
└── islands/
    └── DocGenerator.tsx
```

```typescript
await section.image("/images/logo.png");
```

### 4. Use Hex Colors Without

```typescript
// ✅ CORRECT
section.paragraph().text("Red text", { color: "FF0000" });
table.row({ text: "Cell", cellColor: "E0E0E0", fontColor: "FF0000" });

// ❌ WRONG
section.paragraph().text("Red text", { color: "#FF0000" });
table.row({ text: "Cell", cellColor: "#E0E0E0" });
```

### 5. Apply Styles at Cell Level

```typescript
// ✅ CORRECT - Each cell has its own style
table.row(
  { text: "Name", bold: true, cellColor: "4472C4", fontColor: "FFFFFF" },
  { text: "Age", fontSize: 12, cellColor: "E0E0E0" },
  { text: "City", fontName: "Arial", italic: true },
);

// You can mix simple strings with styled cells
table.row(
  "John Doe",
  { text: "25", bold: true, hAlign: "center" },
  { text: "New York", fontColor: "666666" },
);
```

### 6. Handle Async Image Loading

```typescript
// Wait for all images before download
await section.image("/images/img1.png");
await section.image("/images/img2.png");
await doc.download();
```

### 7. Use Column Defaults Wisely

```typescript
// Set column defaults that apply to all cells in that column
const table = section.table({
  columns: [
    {
      width: "5cm",
      hAlign: "left",
      fontName: "Arial",
      bold: true,
    },
    {
      width: "3cm",
      hAlign: "right",
      fontSize: 12,
      fontColor: "0066CC",
    },
    {
      width: "3cm",
      hAlign: "center",
      cellColor: "E0E0E0",
    },
  ],
});

// Simple rows automatically inherit all column defaults
table.row("Product A", "$100", "Active");
// Product A: left-aligned, Arial, bold
// $100: right-aligned, size 12, blue color
// Active: center-aligned, gray background

// Override column defaults when needed
table.row(
  "Product B",
  { text: "$200", bold: true, fontColor: "FF0000" }, // Override color, add bold
  { text: "Inactive", cellColor: "FFCCCC" }, // Override background
);
```

---

## Common Use Cases

### Meeting Minutes

```typescript
const doc = new DocXaur({ title: "Meeting Minutes" });
const section = doc.addSection();

section.heading("Team Meeting - January 15, 2025");
section.paragraph().text("Attendees: ").text("John, Jane, Bob", { bold: true });
section.lineBreak();

section.heading("Agenda Items", 2);
section.paragraph().text("1. Project status update");
section.paragraph().text("2. Budget review");
section.paragraph().text("3. Next sprint planning");
```

### Certificate

```typescript
const doc = new DocXaur({ title: "Certificate" });
const section = doc.addSection();

section.lineBreak(3);
section.paragraph({ align: "center", size: 28, bold: true })
  .text("CERTIFICATE OF COMPLETION");

section.lineBreak(2);
section.paragraph({ align: "center", size: 14 })
  .text("This is to certify that");

section.lineBreak();
section.paragraph({ align: "center", size: 20, bold: true })
  .text("John Doe");

section.lineBreak();
section.paragraph({ align: "center", size: 14 })
  .text("has successfully completed");

section.lineBreak();
section.paragraph({ align: "center", size: 16, bold: true })
  .text("Advanced TypeScript Course");
```

### Product Catalog

```typescript
const products = [
  { name: "Widget A", price: "$29.99", stock: "In Stock", color: "00AA00" },
  { name: "Widget B", price: "$39.99", stock: "Low Stock", color: "FF6600" },
  { name: "Widget C", price: "$49.99", stock: "Out of Stock", color: "FF0000" },
];

const doc = new DocXaur({ title: "Product Catalog" });
const section = doc.addSection();

section.heading("Product Catalog 2025");

const table = section.table({
  columns: [
    { width: "6cm", hAlign: "left" },
    { width: "3cm", hAlign: "right", bold: true },
    { width: "4cm", hAlign: "center" },
  ],
});

// Header
table.row(
  { text: "Product", bold: true, cellColor: "4472C4", fontColor: "FFFFFF" },
  { text: "Price", cellColor: "4472C4", fontColor: "FFFFFF" },
  {
    text: "Availability",
    bold: true,
    cellColor: "4472C4",
    fontColor: "FFFFFF",
  },
);

// Data rows with conditional styling
// Price column automatically bold from column defaults
products.forEach((product) => {
  table.row(
    product.name,
    product.price,
    { text: product.stock, fontColor: product.color, bold: true },
  );
});

await doc.download("catalog-2025.docx");
```

### Grade Report

```typescript
const doc = new DocXaur({ title: "Grade Report" });
const section = doc.addSection();

section.heading("Student Grade Report", 1, { align: "center" });
section.paragraph({ align: "center" })
  .text("Student: ")
  .text("Jane Smith", { bold: true })
  .lineBreak()
  .text("Semester: ")
  .text("Fall 2024", { bold: true });

section.lineBreak(2);

const table = section.table({
  columns: [
    { width: "6cm", hAlign: "left" },
    { width: "2cm", hAlign: "center", fontSize: 14, bold: true },
    { width: "2cm", hAlign: "center" },
    { width: "3cm", hAlign: "center" },
  ],
});

// Header
table.row(
  { text: "Course", bold: true, cellColor: "4472C4", fontColor: "FFFFFF" },
  { text: "Grade", cellColor: "4472C4", fontColor: "FFFFFF" },
  { text: "Credits", bold: true, cellColor: "4472C4", fontColor: "FFFFFF" },
  { text: "Status", bold: true, cellColor: "4472C4", fontColor: "FFFFFF" },
);

// Grades with color coding
// Grade column automatically bold and size 14 from column defaults
table.row(
  "Mathematics 101",
  { text: "A", fontColor: "00AA00" },
  "4",
  { text: "Pass", cellColor: "CCFFCC" },
);

table.row(
  "Physics 201",
  { text: "B+", fontColor: "0066CC" },
  "3",
  { text: "Pass", cellColor: "CCFFCC" },
);

table.row(
  "English 101",
  { text: "A-", fontColor: "00AA00" },
  "3",
  { text: "Pass", cellColor: "CCFFCC" },
);

// GPA row
table.row(
  { text: "GPA", bold: true, cellColor: "E0E0E0", colspan: 2 },
  { text: "3.75", cellColor: "E0E0E0" },
  { text: "Excellent", bold: true, cellColor: "E0E0E0" },
);

await doc.download("grade-report.docx");
```

---

## Troubleshooting

### Error: "This library only works in Fresh Islands"

**Solution:** Move your DocXaur code to a Fresh Island component (in the
`islands/` folder).

### Images Not Loading

**Checklist:**

- ✅ Image is in `static/` folder or accessible URL
- ✅ Using absolute path (`/images/logo.png`) or full URL
- ✅ Image format is PNG, JPG, JPEG, GIF, or BMP
- ✅ CORS is enabled (for external URLs)

### Table Not Displaying Correctly

**Checklist:**

- ✅ Column widths are defined for all columns
- ✅ Number of cells matches number of columns (unless using colspan)
- ✅ Using correct units (cm, pt, mm, in, %)
- ✅ Properties like `fontSize`, `fontColor`, `fontName` are spelled correctly

### Cell Styling Not Applied

**Common mistakes:**

```typescript
// ❌ WRONG - Using wrong property names
table.row(
  { text: "Hello", size: 12 }, // Should be fontSize
);

// ❌ WRONG - Using color with #
table.row(
  { text: "Hello", fontColor: "#FF0000" }, // Remove #
);

// ✅ CORRECT
table.row(
  { text: "Hello", fontSize: 12, fontColor: "FF0000" },
);
```

### Colspan/Rowspan Issues

**Remember:**

- `colspan`: Number of columns to span (2, 3, etc.)
- `rowspan`: Number of rows to span (2, 3, etc.)
- `rowspan: 0`: Marks continuation of rowspan from above

```typescript
// Header spanning 3 columns
table.row(
  { text: "Title", colspan: 3, hAlign: "center" },
);

// Next row needs 3 cells
table.row("Cell 1", "Cell 2", "Cell 3");

// Rowspan example
table.row(
  { text: "Spanning 2 rows", rowspan: 2, vAlign: "center" },
  "Row 1, Col 2",
  "Row 1, Col 3",
);

// Second row - first cell is continuation
table.row(
  { text: "", rowspan: 0 }, // Continuation marker
  "Row 2, Col 2",
  "Row 2, Col 3",
);
```

---

## API Reference Summary

### Classes

- `DocXaur` - Main document class
- `Section` - Document section
- `Paragraph` - Paragraph builder
- `Table` - Table builder

### TableColumn Properties

- `width` (required) - Column width
- `hAlign` - Default horizontal alignment
- `vAlign` - Default vertical alignment
- `fontName` - Default font family
- `fontSize` - Default font size
- `fontColor` - Default text color
- `cellColor` - Default background color
- `bold` - Default bold styling
- `italic` - Default italic styling
- `underline` - Default underline styling

### TableCellData Properties

All column properties can be overridden at cell level, plus:

- `text` (required) - Cell content
- `colspan` - Column span
- `rowspan` - Row span
- `height` - Cell height

### Key Methods

**DocXaur:**

- `addSection()` - Add section
- `download()` - Download document
- `toBlob()` - Get document as Blob

**Section:**

- `heading()` - Add heading
- `paragraph()` - Create paragraph
- `image()` - Add image
- `table()` - Create table
- `lineBreak()` - Add line breaks
- `pageBreak()` - Add page breaks

**Paragraph:**

- `text()` - Add text with optional styling
- `tab()` - Add tab character
- `lineBreak()` - Add line break within paragraph
- `pageBreak()` - Add page break within paragraph
- `apply()` - Conditional operations

**Table:**

- `row()` - Add table row with cell-level styling
- `apply()` - Conditional operations

---

## Cell Styling Reference

### Complete TableCellData Example

```typescript
table.row(
  {
    // Required
    text: "Sample Cell",

    // Font Properties
    fontName: "Arial", // Font family
    fontSize: 12, // Size in points
    fontColor: "FF0000", // Text color (hex without #)
    bold: true, // Bold text
    italic: true, // Italic text
    underline: true, // Underlined text

    // Alignment
    hAlign: "center", // left | center | right
    vAlign: "top", // top | center | bottom

    // Cell Properties
    cellColor: "E0E0E0", // Background color (hex without #)
    height: "1cm", // Cell height

    // Spanning
    colspan: 2, // Span 2 columns
    rowspan: 3, // Span 3 rows (use 0 for continuation)
  },
);
```

### Color Palette Examples

```typescript
// Professional colors
const colors = {
  // Blues
  darkBlue: "4472C4",
  lightBlue: "D9E1F2",

  // Grays
  darkGray: "808080",
  lightGray: "E0E0E0",

  // Status colors
  success: "00AA00",
  warning: "FF6600",
  error: "FF0000",

  // Highlights
  yellow: "FFFF00",
  cyan: "00FFFF",
  magenta: "FF00FF",
};

// Usage in table
table.row(
  { text: "Header", cellColor: colors.darkBlue, fontColor: "FFFFFF" },
  { text: "Success", fontColor: colors.success, bold: true },
  { text: "Warning", fontColor: colors.warning, bold: true },
  { text: "Error", fontColor: colors.error, bold: true },
);
```

---

## Performance Tips

### 1. Reuse Table Definitions

```typescript
// Define reusable cell styles
const headerStyle = {
  bold: true,
  cellColor: "4472C4",
  fontColor: "FFFFFF",
};

const totalStyle = {
  bold: true,
  cellColor: "E0E0E0",
};

// Use in multiple tables
table1.row(
  { text: "Name", ...headerStyle },
  { text: "Value", ...headerStyle },
);

table2.row(
  { text: "Product", ...headerStyle },
  { text: "Price", ...headerStyle },
);
```

### 2. Batch Image Loading

```typescript
// Load all images at once
const section = doc.addSection();

await Promise.all([
  section.image("/images/logo.png"),
  section.image("/images/chart1.png"),
  section.image("/images/chart2.png"),
]);
```

### 3. Use Simple Strings When Possible

```typescript
// Simple data doesn't need objects
table.row("John", "25", "New York");

// Only use objects when styling is needed
table.row(
  { text: "Jane", bold: true },
  "30",
  { text: "Los Angeles", fontColor: "0066CC" },
);
```

---

## Migration Guide

### From Old API to New API

**Old way (deprecated):**

```typescript
// ❌ Row options don't exist anymore
table.row(
  { height: "1cm", vAlign: "top" },
  "Cell 1",
  "Cell 2",
);
```

**New way:**

```typescript
// ✅ Apply options to individual cells
table.row(
  { text: "Cell 1", height: "1cm", vAlign: "top" },
  { text: "Cell 2", height: "1cm", vAlign: "top" },
);

// Or mix simple and styled cells
table.row(
  { text: "Cell 1", height: "1cm" },
  "Cell 2", // Uses defaults
);
```

---

## Advanced Techniques

### Dynamic Table Generation

```typescript
const data = [
  { name: "Alice", score: 95, grade: "A" },
  { name: "Bob", score: 87, grade: "B" },
  { name: "Charlie", score: 78, grade: "C" },
];

const table = section.table({
  columns: [
    { width: "5cm", align: "left" },
    { width: "3cm", align: "center" },
    { width: "3cm", align: "center" },
  ],
});

// Header
table.row(
  { text: "Name", bold: true, cellColor: "4472C4", fontColor: "FFFFFF" },
  { text: "Score", bold: true, cellColor: "4472C4", fontColor: "FFFFFF" },
  { text: "Grade", bold: true, cellColor: "4472C4", fontColor: "FFFFFF" },
);

// Dynamic rows with conditional styling
data.forEach((student) => {
  const gradeColor = student.score >= 90
    ? "00AA00"
    : student.score >= 80
    ? "0066CC"
    : "FF6600";

  table.row(
    student.name,
    { text: student.score.toString(), bold: true },
    { text: student.grade, fontColor: gradeColor, bold: true },
  );
});
```

### Conditional Cell Styling

```typescript
const values = [100, -50, 200, -30, 150];

table.row(
  { text: "Value", bold: true, cellColor: "E0E0E0" },
  { text: "Status", bold: true, cellColor: "E0E0E0" },
);

values.forEach((value) => {
  const isPositive = value >= 0;

  table.row(
    {
      text: value.toString(),
      fontColor: isPositive ? "00AA00" : "FF0000",
      bold: true,
    },
    {
      text: isPositive ? "Profit" : "Loss",
      cellColor: isPositive ? "CCFFCC" : "FFCCCC",
      bold: true,
    },
  );
});
```

---

## License

MIT License

## Support

For issues and questions, please visit:
[GitHub Issues](https://github.com/yourrepo/docxaur/issues)

---

## Quick Reference Card

### Document Creation

```typescript
const doc = new DocXaur({ title: "My Doc" });
const section = doc.addSection();
await doc.download("file.docx");
```

### Text Content

```typescript
section.heading("Title", 1);
section.paragraph().text("Hello").text("World", { bold: true });
section.lineBreak();
section.pageBreak();
```

### Images

```typescript
await section.image("/images/logo.png", { width: "5cm" });
```

### Tables

```typescript
const table = section.table({
  columns: [
    { width: "5cm", align: "left" },
    { width: "3cm", align: "right" },
  ],
});

table.row(
  { text: "Header", bold: true, cellColor: "4472C4", fontColor: "FFFFFF" },
  { text: "Value", bold: true, cellColor: "4472C4", fontColor: "FFFFFF" },
);

table.row("Simple text", { text: "Styled", fontSize: 12, bold: true });
```

### Common Cell Properties

```typescript
{
  text: "Content",
  fontName: "Arial",
  fontSize: 12,
  fontColor: "FF0000",
  cellColor: "E0E0E0",
  bold: true,
  italic: true,
  underline: true,
  hAlign: "center",
  vAlign: "center",
  colspan: 2,
  rowspan: 2,
  height: "1cm"
}
```
