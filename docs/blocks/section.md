# Section

A section groups blocks and emits page settings (`<w:sectPr>`). Add headings, paragraphs, images, tables, then finalize the section’s properties.

## Create & Configure

```ts
const section = doc.addSection({
  pageSize: { width: "21cm", height: "29.7cm", orientation: "portrait" },
  margins:  { top: "2.54cm", right: "2.54cm", bottom: "2.54cm", left: "2.54cm" },
});
```

## Heading

```ts
section.heading("Executive Summary", 1);                // H1
section.heading("Background", 2, { align: "center" }); // H2, centered
```

> Sizes (pt): H1=24, H2=20, H3=18, H4=16, H5=14, H6=12

## Breaks

```ts
section.paragraph().text("Intro").lineBreak(2);
section.paragraph().text("End of page").pageBreak();
```
