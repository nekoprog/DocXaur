# Relationships (Images)

DocXaur ensures every `<a:blip r:embed="rIdX">` in `word/document.xml` has a matching `<Relationship … Type="…/image">` entry in `word/_rels/document.xml.rels`, targeting `media/image{id}.{ext}`.

## What the Helper Does

- Reads or initializes `document.xml.rels`
- Appends **missing** image relationships with correct `Id` and `Target`
- Never duplicates existing relationships

## When It Runs

DocXaur generates `word/document.xml` **first** (tables register images during build), then:

1. Produces raw relationships XML  
2. Applies the guard (`ensureImageRelationships`)  
3. Writes the fixed `.rels` into the ZIP

This guarantees Word can resolve all images.

## Troubleshooting

- Ensure images exist under `word/media/*`
- Each `r:embed="rId…"` has a matching entry in `.rels`
- `[Content_Types].xml` includes your image extensions (`jpeg`, `jpg`, etc.)
