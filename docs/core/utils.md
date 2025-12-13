# Utils

Utilities used across DocXaur:

## Units

- `cmToTwips(cm)` → word twips
- `cmToEmu(cm)` → OOXML EMU
- `ptToHalfPoints(pt)` → Word half-points
- `parseNumberTwips("10cm" | "20pt" | "1in" | "50%")` → twips
- `parseImageSize("10cm" | "120px" | "1in")` → EMU

## XML

- `escapeXML(text)` → safe XML text

## Images

- `fetchImageAsBase64(url)` → `{ data, extension }`\
  Use `https://…` or `"/images/…"` (Fresh static); throws on failures.

## Example

```ts
import { cmToEmu, escapeXML } from "jsr:@fytz/docxaur";

const widthEmu = cmToEmu(10); // -> for DrawingML extent
const safeText = escapeXML('3 < 5 & "quotes"'); // -> XML-safe text
```
