# DocXaur Documentation

Welcome to DocXaur’s extended documentation. This site complements the API docs
that JSR generates from our JSDoc comments.

- **Core**: Packaging, relationships, and helpers
- **Blocks**: Paragraphs, Images, Tables (future: headers, footers, bookmarks)

## Quick Links

- [Core Overview](./core/index.md)
- [Blocks Overview](./blocks/index.md)
- [Paragraph](./blocks/paragraph.md)
- [Image](./blocks/image.md)
- [Table](./blocks/table.md)

## Quick Start

```ts
import { DocXaur } from "jsr:@your-scope/docxaur";

const doc = new DocXaur({ title: "My Document" });
const section = doc.addSection();

section.heading("Hello World", 1);
section.paragraph().text("This is a paragraph.");

await doc.download("my-document.docx");
```

> Works in **Fresh Islands** (browser) only.
