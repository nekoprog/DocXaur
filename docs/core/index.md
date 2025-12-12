# Core Overview

DocXaur’s core provides the foundation for building and packaging `.docx` files:

- `DocXaur` — Main class orchestrating sections, relationships, packaging
- `relationships` — Ensures image relationships exist in `word/_rels/document.xml.rels`
- `utils` — Unit conversions, XML escaping, and image fetching

## Minimal Flow

```ts
const doc = new DocXaur();
const section = doc.addSection();

section.paragraph({ align: "center" }).text("Hello DocXaur!");

await doc.download("hello.docx");
```
