# Image

Adds an inline DrawingML picture referencing a relationship entry in `document.xml.rels`, backed by a binary in `word/media/`.

## Basic

```ts
await section.image("/images/logo.png", { width: "5cm", align: "left" });
```

## Options

- `width`: string (cm/pt/mm/in/px)
- `height`: string (cm/pt/mm/in/px) — defaults to square if omitted
- `align`: "left" | "center" | "right" | "justify"

## External URLs & Static Files

```ts
await section.image("https://example.com/photo.jpg", { width: "12cm" });
await section.image("/images/banner.png",             { width: "10cm" });
```

> Images must be accessible (CORS for external URLs). `jpeg` and `jpg` are supported.
