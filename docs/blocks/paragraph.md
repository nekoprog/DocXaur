# Paragraph

Represents `<w:p>` with inline text runs and styles. Chain methods for readable composition.

## Basic

```ts
section.paragraph().text("Hello ").text("World", { bold: true });
```

## Options

- `align`: "left" | "center" | "right" | "justify"
- `bold`, `italic`, `underline`: boolean (per run)
- `fontSize`: number (pt)
- `fontColor`: hex without `#` (e.g., "FF0000")
- `fontName`: string

## Line & Page Breaks

```ts
section.paragraph()
  .text("Line 1")
  .lineBreak()
  .text("Line 2")
  .pageBreak()
  .text("New Page");
```

## Tabs

```ts
section.paragraph().text("Name:").tab().text("Jane Doe");
```

## Deprecated

```ts
// Prefer explicit method calls
section.paragraph().apply(p => p.text("Text")); // warns (deprecated)
```
