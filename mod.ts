/**
 * DocXaur - A semantic, consistent DOCX generation library for Deno Fresh Islands
 *
 * ⚠️ IMPORTANT: This library ONLY works in Fresh Islands (browser environment)
 * It will NOT work in server-side routes, middleware, or Deno runtime
 *
 * Design Principles:
 * 1. Fresh Islands Only (browser APIs required)
 * 2. Semantic API (intuitive, minimal nesting)
 * 3. Consistent (similar patterns across all features)
 */

// Import from CDN that works in browser
// import JSZip from "https://cdn.jsdelivr.net/npm/jszip@3.10.1/+esm";

// Import zip-js from JSR
import {
  BlobReader,
  BlobWriter,
  TextReader,
  ZipWriter,
} from "jsr:@zip-js/zip-js";

// ============================================================================
// ENVIRONMENT CHECK
// ============================================================================

function checkBrowserEnvironment(): void {
  if (typeof window === "undefined" || typeof document === "undefined") {
    throw new Error(
      "❌ DocXaur Error: This library only works in Fresh Islands (browser environment).\n" +
        "\n" +
        "You're trying to use it in a server-side context.\n" +
        "\n" +
        "Solution: Move your code to a Fresh Island component.\n" +
        "Example: Create 'islands/DocxGenerator.tsx' and use DocXaur there.\n" +
        "\n" +
        "See documentation: https://github.com/nekoprog/DocXaur",
    );
  }
}

// ============================================================================
// TYPES & INTERFACES
// ============================================================================

export interface DocumentOptions {
  title?: string;
  creator?: string;
  description?: string;
  subject?: string;
  keywords?: string;
  fontName?: string;
  fontSize?: number;
}

export interface SectionOptions {
  pageSize?: PageSize;
  margins?: Margins;
}

export interface PageSize {
  width: string;
  height: string;
  orientation?: "portrait" | "landscape";
}

export interface Margins {
  top: string;
  right: string;
  bottom: string;
  left: string;
}

export interface TextStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  size?: number;
  color?: string;
  font?: string;
}

export interface ParagraphOptions extends TextStyle {
  align?: "left" | "center" | "right" | "justify";
  spacing?: {
    before?: number;
    after?: number;
    line?: number;
  };
  breakBefore?: number;
  breakAfter?: number;
}

export interface ImageOptions {
  width?: string;
  height?: string;
  align?: "left" | "center" | "right" | "justify";
}

export interface TableColumn {
  width: string;
  hAlign?: "left" | "center" | "right" | "justify";
  vAlign?: "top" | "center" | "bottom";
  fontName?: string;
  fontSize?: number;
  fontColor?: string;
  cellColor?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
}

export interface TableOptions {
  columns: TableColumn[];
  width?: string;
  align?: "left" | "center" | "right" | "justify";
  borders?: boolean;
  indent?: string;
}

export interface TableCellData {
  text: string;
  fontName?: string;
  fontSize?: number;
  fontColor?: string;
  cellColor?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  hAlign?: "left" | "center" | "right" | "justify";
  vAlign?: "top" | "center" | "bottom";
  colspan?: number;
  rowspan?: number;
  height?: string;
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

function cmToTwips(cm: number): number {
  return Math.round(cm * 567);
}

function cmToEmu(cm: number): number {
  return Math.round(cm * 360000);
}

function ptToHalfPoints(pt: number): number {
  return Math.round(pt * 2);
}

function parseNumberTwips(width: string): number {
  const match = width.match(/^([\d.]+)(cm|pt|mm|in|%)$/);
  if (!match) return 1000;

  const value = parseFloat(match[1]);
  const unit = match[2];

  switch (unit) {
    case "cm":
      return cmToTwips(value);
    case "mm":
      return Math.round(value * 56.7);
    case "pt":
      return Math.round(value * 20);
    case "in":
      return Math.round(value * 1440);
    case "%":
      return Math.round(value * 50);
    default:
      return 1000;
  }
}

function base64ToUint8Array(base64: string): Uint8Array {
  const binary = atob(base64);
  const len = binary.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

function parseImageSize(size: string): number {
  const match = size.match(/^([\d.]+)(cm|pt|mm|in|px)$/);
  if (!match) return cmToEmu(5);

  const value = parseFloat(match[1]);
  const unit = match[2];

  switch (unit) {
    case "cm":
      return cmToEmu(value);
    case "mm":
      return cmToEmu(value / 10);
    case "in":
      return Math.round(value * 914400);
    case "pt":
      return Math.round(value * 12700);
    case "px":
      return Math.round(value * 9525);
    default:
      return cmToEmu(5);
  }
}

function escapeXML(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

async function fetchImageAsBase64(
  url: string,
): Promise<{ data: string; extension: string }> {
  try {
    checkBrowserEnvironment();

    if (
      !url.startsWith("http://") && !url.startsWith("https://") &&
      !url.startsWith("/")
    ) {
      throw new Error(
        `Invalid image URL: "${url}"\n` +
          "In Fresh Islands, only HTTP/HTTPS URLs or absolute paths (e.g., /images/logo.png) are supported.\n" +
          "Place images in your Fresh 'static/' folder and reference them as '/images/filename.png'",
      );
    }

    const response = await fetch(url);
    if (!response.ok) {
      throw new Error(
        `Failed to fetch image: ${response.status} ${response.statusText}`,
      );
    }

    const arrayBuffer = await response.arrayBuffer();
    const bytes = new Uint8Array(arrayBuffer);
    let binary = "";
    for (let i = 0; i < bytes.length; i++) {
      binary += String.fromCharCode(bytes[i]);
    }
    const base64 = btoa(binary);

    const extension =
      url.toLowerCase().match(/\.(png|jpg|jpeg|gif|bmp)$/)?.[1] || "png";

    return { data: base64, extension };
  } catch (error) {
    throw new Error(
      `Failed to load image from "${url}":\n${error}\n\n` +
        "Tip: Make sure the image URL is publicly accessible and supports CORS.",
    );
  }
}

// ============================================================================
// CORE CLASSES
// ============================================================================

// export class DocXaur {
//   private sections: Section[] = [];
//   private options: DocumentOptions;
//   private images: Map<string, { data: string; extension: string; id: number }> =
//     new Map();
//   private imageCounter = 1;
//   private fontName: string;
//   private fontSize: number;

//   constructor(options: DocumentOptions = {}) {
//     checkBrowserEnvironment();

//     this.fontName = options.fontName || "Calibri";
//     this.fontSize = options.fontSize || 11;

//     this.options = {
//       title: options.title || "Document",
//       creator: options.creator || "DocXaur",
//       description: options.description || "",
//       subject: options.subject || "",
//       keywords: options.keywords || "",
//       fontName: this.fontName,
//       fontSize: this.fontSize,
//     };
//   }

//   getDefaultFont(): string {
//     return this.fontName;
//   }

//   getDefaultSize(): number {
//     return this.fontSize;
//   }

//   addSection(options?: SectionOptions): Section {
//     const section = new Section(options, this);
//     this.sections.push(section);
//     return section;
//   }

//   async registerImage(url: string): Promise<number> {
//     if (this.images.has(url)) {
//       return this.images.get(url)!.id;
//     }

//     const imageData = await fetchImageAsBase64(url);
//     const id = this.imageCounter++;
//     this.images.set(url, { ...imageData, id });
//     return id;
//   }

//   async toBlob(): Promise<Blob> {
//     checkBrowserEnvironment();

//     const zip = new JSZip();

//     zip.file("[Content_Types].xml", this.generateContentTypes());

//     const rels = zip.folder("_rels");
//     rels?.file(".rels", this.generateRootRels());

//     const docRels = zip.folder("word/_rels");
//     docRels?.file("document.xml.rels", this.generateDocRels());

//     const word = zip.folder("word");
//     word?.file("document.xml", this.generateDocument());
//     word?.file("styles.xml", this.generateStyles());
//     word?.file("fontTable.xml", this.generateFontTable());
//     word?.file("settings.xml", this.generateSettings());

//     const media = word?.folder("media");
//     for (const [path, imgData] of this.images) {
//       const filename = `image${imgData.id}.${imgData.extension}`;
//       media?.file(filename, imgData.data, { base64: true });
//     }

//     const blob = await zip.generateAsync({ type: "blob" });
//     return blob;
//   }

//   async download(filename: string = "document.docx"): Promise<void> {
//     checkBrowserEnvironment();

//     const blob = await this.toBlob();
//     const url = URL.createObjectURL(blob);
//     const link = document.createElement("a");
//     link.href = url;
//     link.download = filename;
//     document.body.appendChild(link);
//     link.click();
//     document.body.removeChild(link);
//     URL.revokeObjectURL(url);
//   }

export class DocXaur {
  private sections: Section[] = [];
  private options: DocumentOptions;
  private images: Map<string, { data: string; extension: string; id: number }> =
    new Map();
  private imageCounter = 1;
  private fontName: string;
  private fontSize: number;

  constructor(options: DocumentOptions = {}) {
    checkBrowserEnvironment();
    this.fontName = options.fontName ?? "Calibri";
    this.fontSize = options.fontSize ?? 11;
    this.options = {
      title: options.title ?? "Document",
      creator: options.creator ?? "DocXaur",
      description: options.description ?? "",
      subject: options.subject ?? "",
      keywords: options.keywords ?? "",
      fontName: this.fontName,
      fontSize: this.fontSize,
    };
  }

  addSection(options?: SectionOptions): Section {
    const section = new Section(options, this);
    this.sections.push(section);
    return section;
  }

  async registerImage(url: string): Promise<number> {
    if (this.images.has(url)) {
      return this.images.get(url)!.id;
    }
    const imageData = await fetchImageAsBase64(url);
    const id = this.imageCounter++;
    this.images.set(url, { ...imageData, id });
    return id;
  }

  // ✅ NEW: Create DOCX ZIP using zip-js
  private async createDocxZip(): Promise<Blob> {
    const blobWriter = new BlobWriter();
    const zipWriter = new ZipWriter(blobWriter);

    // Core files
    await zipWriter.add(
      "[Content_Types].xml",
      new TextReader(this.generateContentTypes()),
    );
    await zipWriter.add("_rels/.rels", new TextReader(this.generateRootRels()));
    await zipWriter.add(
      "word/_rels/document.xml.rels",
      new TextReader(this.generateDocRels()),
    );
    await zipWriter.add(
      "word/document.xml",
      new TextReader(this.generateDocument()),
    );
    await zipWriter.add(
      "word/styles.xml",
      new TextReader(this.generateStyles()),
    );
    await zipWriter.add(
      "word/fontTable.xml",
      new TextReader(this.generateFontTable()),
    );
    await zipWriter.add(
      "word/settings.xml",
      new TextReader(this.generateSettings()),
    );

    // Images
    for (const [path, imgData] of this.images) {
      const filename = `word/media/image${imgData.id}.${imgData.extension}`;
      await zipWriter.add(
        filename,
        new BlobReader(new Blob([base64ToUint8Array(imgData.data)])),
      );
    }

    await zipWriter.close();
    return await blobWriter.getData();
  }

  async toBlob(): Promise<Blob> {
    checkBrowserEnvironment();
    return await this.createDocxZip();
  }

  async download(filename: string = "document.docx"): Promise<void> {
    checkBrowserEnvironment();
    const blob = await this.toBlob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }

  private generateContentTypes(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="jpg" ContentType="image/jpeg"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="gif" ContentType="image/gif"/>
  <Default Extension="bmp" ContentType="image/bmp"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;
  }

  private generateRootRels(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
  }

  private generateDocRels(): string {
    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
`;

    let relId = 4;
    for (const [path, imgData] of this.images) {
      xml +=
        `  <Relationship Id="rId${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image${imgData.id}.${imgData.extension}"/>\n`;
      relId++;
    }

    xml += `</Relationships>`;
    return xml;
  }

  getImageRelId(imageId: number): string {
    return `rId${imageId + 3}`;
  }

  private generateDocument(): string {
    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
`;

    for (const section of this.sections) {
      xml += section.toXML();
    }

    if (this.sections.length > 0) {
      const lastSection = this.sections[this.sections.length - 1];
      xml += lastSection.getSectionPropertiesXML();
    }

    xml += `  </w:body>
</w:document>`;

    return xml;
  }

  private generateStyles(): string {
    const fontName = this.fontName;
    const fontSize = ptToHalfPoints(this.fontSize);

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="${fontName}" w:hAnsi="${fontName}" w:cs="${fontName}"/>
        <w:sz w:val="${fontSize}"/>
        <w:szCs w:val="${fontSize}"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
</w:styles>`;
  }

  private generateFontTable(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:font w:name="Arial">
    <w:panose1 w:val="020B0604020202020204"/>
    <w:charset w:val="00"/>
    <w:family w:val="swiss"/>
    <w:pitch w:val="variable"/>
  </w:font>
  <w:font w:name="Times New Roman">
    <w:panose1 w:val="02020603050405020304"/>
    <w:charset w:val="00"/>
    <w:family w:val="roman"/>
    <w:pitch w:val="variable"/>
  </w:font>
  <w:font w:name="Calibri">
    <w:panose1 w:val="020F0502020204030204"/>
    <w:charset w:val="00"/>
    <w:family w:val="swiss"/>
    <w:pitch w:val="variable"/>
  </w:font>
</w:fonts>`;
  }

  private generateSettings(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:zoom w:percent="100"/>
</w:settings>`;
  }
}

export class Section {
  private elements: Element[] = [];
  private options: SectionOptions;
  private doc: DocXaur;

  constructor(options: SectionOptions = {}, doc: DocXaur) {
    this.options = {
      pageSize: options.pageSize ||
        { width: "21cm", height: "29.7cm", orientation: "portrait" }, // A4 Portrait
      margins: options.margins ||
        { top: "2.54cm", right: "2.54cm", bottom: "2.54cm", left: "2.54cm" },
    };
    this.doc = doc;
  }

  paragraph(options?: ParagraphOptions): Paragraph {
    const paragraph = new Paragraph(options);
    this.elements.push(paragraph);
    return paragraph;
  }

  heading(
    content: string,
    level: 1 | 2 | 3 | 4 | 5 | 6 = 1,
    options?: ParagraphOptions,
  ): this {
    const sizes = [24, 20, 18, 16, 14, 12];
    const paragraph = new Paragraph(options);
    paragraph.text(content, { bold: true, size: sizes[level - 1], ...options });
    this.elements.push(paragraph);
    return this;
  }

  async image(url: string, options?: ImageOptions): Promise<this> {
    const imageId = await this.doc.registerImage(url);
    this.elements.push(new Image(imageId, this.doc, options));
    return this;
  }

  table(options: TableOptions): Table {
    const table = new Table(options);
    this.elements.push(table);
    return table;
  }

  // lineBreak(count: number = 1): this {
  //   for (let i = 0; i < count; i++) {
  //     this.elements.push(new LineBreak());
  //   }
  //   return this;
  // }

  // pageBreak(count: number = 1): this {
  //   for (let i = 0; i < count; i++) {
  //     this.elements.push(new PageBreak());
  //   }
  //   return this;
  // }

  toXML(): string {
    return this.elements.map((el) => el.toXML()).join("\n");
  }

  getSectionPropertiesXML(): string {
    const pageSize = this.options.pageSize!;
    const margins = this.options.margins!;

    const width = parseNumberTwips(pageSize.width);
    const height = parseNumberTwips(pageSize.height);
    const orient = pageSize.orientation === "landscape"
      ? "landscape"
      : "portrait";

    return `    <w:sectPr>
      <w:pgSz w:w="${width}" w:h="${height}" w:orient="${orient}"/>
      <w:pgMar w:top="${parseNumberTwips(margins.top)}"
               w:right="${parseNumberTwips(margins.right)}"
               w:bottom="${parseNumberTwips(margins.bottom)}"
               w:left="${parseNumberTwips(margins.left)}"
               w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
`;
  }
}

// ============================================================================
// ELEMENT CLASSES
// ============================================================================

abstract class Element {
  abstract toXML(): string;
}

interface TextRun {
  text: string;
  style?: TextStyle;
}

export class Paragraph extends Element {
  private runs: TextRun[] = [];
  private options: ParagraphOptions;

  constructor(options: ParagraphOptions = {}) {
    super();
    this.options = options;
  }

  text(text: string, style?: TextStyle): this {
    this.runs.push({ text, style });
    return this;
  }

  tab(): this {
    this.runs.push({ text: "\t" });
    return this;
  }

  lineBreak(count: number = 1): this {
    for (let i = 0; i < count; i++) {
      this.runs.push({ text: "\n" });
    }
    return this;
  }

  pageBreak(count: number = 1): this {
    for (let i = 0; i < count; i++) {
      this.runs.push({ text: "[PAGE_BREAK]" });
    }
    return this;
  }

  apply(...operations: ((builder: this) => this)[]): this {
    for (const operation of operations) {
      operation(this);
    }
    return this;
  }

  toXML(): string {
    const align = this.options.align || "left";
    const breaksBefore = this.options.breakBefore || 0;
    const breaksAfter = this.options.breakAfter || 0;

    let xml = "    <w:p>\n";

    xml += "      <w:pPr>\n";

    if (align !== "left") {
      const jc = align === "justify" ? "both" : align;
      xml += `        <w:jc w:val="${jc}"/>\n`;
    }

    const spacing = this.options.spacing;
    const before = spacing?.before ? ptToHalfPoints(spacing.before) * 20 : 0;
    const after = spacing?.after ? ptToHalfPoints(spacing.after) * 20 : 0;
    const line = spacing?.line ? Math.round(spacing.line * 240) : 240;

    xml +=
      `        <w:spacing w:after="${after}" w:before="${before}" w:line="${line}" w:lineRule="auto"/>\n`;

    xml += "      </w:pPr>\n";

    for (let i = 0; i < breaksBefore; i++) {
      xml += "      <w:r><w:br/></w:r>\n";
    }

    for (const run of this.runs) {
      if (run.text === "\t") {
        xml += "      <w:r><w:tab/></w:r>\n";
      } else if (run.text === "\n") {
        xml += "      <w:r><w:br/></w:r>\n";
      } else if (run.text === "[PAGE_BREAK]") {
        xml += '      <w:r><w:br w:type="page"/></w:r>\n';
      } else {
        xml += "      <w:r>\n";

        const style = run.style;
        if (style && this.hasRunProperties(style)) {
          xml += "        <w:rPr>\n";
          if (style.bold) xml += "          <w:b/>\n";
          if (style.italic) xml += "          <w:i/>\n";
          if (style.underline) xml += '          <w:u w:val="single"/>\n';
          if (style.size) {
            xml += `          <w:sz w:val="${ptToHalfPoints(style.size)}"/>\n`;
          }
          if (style.color) {
            xml += `          <w:color w:val="${style.color}"/>\n`;
          }
          if (style.font) {
            xml +=
              `          <w:rFonts w:ascii="${style.font}" w:hAnsi="${style.font}"/>\n`;
          }
          xml += "        </w:rPr>\n";
        }

        xml += `        <w:t xml:space="preserve">${
          escapeXML(run.text)
        }</w:t>\n`;
        xml += "      </w:r>\n";
      }
    }

    for (let i = 0; i < breaksAfter; i++) {
      xml += "      <w:r><w:br/></w:r>\n";
    }

    xml += "    </w:p>\n";

    return xml;
  }

  private hasRunProperties(style: TextStyle): boolean {
    return !!(style.bold || style.italic || style.underline ||
      style.size || style.color || style.font);
  }
}

// export class LineBreak extends Element {
//   toXML(): string {
//     return `    <w:p>
//       <w:r>
//         <w:br/>
//       </w:r>
//     </w:p>
// `;
//   }
// }

// export class PageBreak extends Element {
//   toXML(): string {
//     return `    <w:p>
//       <w:r>
//         <w:br w:type="page"/>
//       </w:r>
//     </w:p>
// `;
//   }
// }

export class Image extends Element {
  private static imageCounter = 1;

  constructor(
    private imageId: number,
    private doc: DocXaur,
    private options: ImageOptions = {},
  ) {
    super();
  }

  toXML(): string {
    const width = this.options.width
      ? parseImageSize(this.options.width)
      : cmToEmu(10);
    const height = this.options.height
      ? parseImageSize(this.options.height)
      : width;
    const align = this.options.align || "center";
    const relId = this.doc.getImageRelId(this.imageId);
    const drawingId = Image.imageCounter++;

    let xml = "    <w:p>\n";

    if (align !== "left") {
      xml += "      <w:pPr>\n";
      xml += `        <w:jc w:val="${align}"/>\n`;
      xml += "      </w:pPr>\n";
    }

    xml += `      <w:r>
        <w:drawing>
          <wp:inline distT="0" distB="0" distL="0" distR="0">
            <wp:extent cx="${width}" cy="${height}"/>
            <wp:effectExtent l="0" t="0" r="0" b="0"/>
            <wp:docPr id="${drawingId}" name="Picture ${drawingId}"/>
            <wp:cNvGraphicFramePr>
              <a:graphicFrameLocks noChangeAspect="1"/>
            </wp:cNvGraphicFramePr>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic>
                  <pic:nvPicPr>
                    <pic:cNvPr id="${drawingId}" name="Picture ${drawingId}"/>
                    <pic:cNvPicPr/>
                  </pic:nvPicPr>
                  <pic:blipFill>
                    <a:blip r:embed="${relId}"/>
                    <a:stretch>
                      <a:fillRect/>
                    </a:stretch>
                  </pic:blipFill>
                  <pic:spPr>
                    <a:xfrm>
                      <a:off x="0" y="0"/>
                      <a:ext cx="${width}" cy="${height}"/>
                    </a:xfrm>
                    <a:prstGeom prst="rect">
                      <a:avLst/>
                    </a:prstGeom>
                  </pic:spPr>
                </pic:pic>
              </a:graphicData>
            </a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
`;

    return xml;
  }
}

export class Table extends Element {
  private rows: TableRow[] = [];
  private options: TableOptions;

  constructor(options: TableOptions) {
    super();
    this.options = options;
    if (this.options.borders === undefined) {
      this.options.borders = true;
    }
  }

  row(...cells: (string | TableCellData)[]): this {
    const row = new TableRow(this.options);

    cells.forEach((cell, index) => {
      const colOptions = this.options.columns[index];

      if (typeof cell === "string") {
        row.cell({
          text: cell,
          hAlign: colOptions?.hAlign || "center",
          vAlign: colOptions?.vAlign || "center",
          fontName: colOptions?.fontName,
          fontSize: colOptions?.fontSize,
          fontColor: colOptions?.fontColor,
          cellColor: colOptions?.cellColor,
          bold: colOptions?.bold,
          italic: colOptions?.italic,
          underline: colOptions?.underline,
        });
      } else {
        row.cell({
          hAlign: colOptions?.hAlign || "center",
          vAlign: colOptions?.vAlign || "center",
          fontName: colOptions?.fontName,
          fontSize: colOptions?.fontSize,
          fontColor: colOptions?.fontColor,
          cellColor: colOptions?.cellColor,
          bold: colOptions?.bold,
          italic: colOptions?.italic,
          underline: colOptions?.underline,
          ...cell, // Cell properties override column defaults
        });
      }
    });

    this.rows.push(row);
    return this;
  }

  apply(...operations: ((builder: this) => this)[]): this {
    for (const op of operations) {
      op(this);
    }
    return this;
  }

  toXML(): string {
    const align = this.options.align || "center";

    let xml = "    <w:tbl>\n";
    xml += "      <w:tblPr>\n";

    if (this.options.indent) {
      const indentTwips = parseNumberTwips(this.options.indent);
      xml += `        <w:tblInd w:w="${indentTwips}" w:type="dxa"/>\n`;
    }

    if (this.options.width) {
      const width = parseNumberTwips(this.options.width);
      xml += `        <w:tblW w:w="${width}" w:type="dxa"/>\n`;
    }

    xml += `        <w:jc w:val="${align}"/>\n`;

    if (this.options.borders) {
      xml += `        <w:tblBorders>
          <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
          <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
          <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
          <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
          <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
          <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        </w:tblBorders>\n`;
    }

    xml += "      </w:tblPr>\n";

    xml += "      <w:tblGrid>\n";
    for (const col of this.options.columns) {
      const width = parseNumberTwips(col.width);
      xml += `        <w:gridCol w:w="${width}"/>\n`;
    }
    xml += "      </w:tblGrid>\n";

    for (const row of this.rows) {
      xml += row.toXML();
    }

    xml += "    </w:tbl>\n";
    return xml;
  }
}

class TableRow {
  private cells: TableCell[] = [];

  constructor(private tableOptions: TableOptions) {}

  cell(data: TableCellData): this {
    this.cells.push(new TableCell(data));
    return this;
  }

  toXML(): string {
    let xml = "      <w:tr>\n";

    // Calculate max height from all cells
    let maxHeight = 170; // default
    for (const cell of this.cells) {
      const cellHeight = cell.getHeight();
      if (cellHeight > maxHeight) {
        maxHeight = cellHeight;
      }
    }

    xml += "        <w:trPr>\n";
    xml += `          <w:trHeight w:val="${maxHeight}" w:hRule="atLeast"/>\n`;
    xml += "        </w:trPr>\n";

    for (let i = 0; i < this.cells.length; i++) {
      xml += this.cells[i].toXML(i, this.tableOptions);
    }

    xml += "      </w:tr>\n";
    return xml;
  }
}

class TableCell {
  constructor(private data: TableCellData) {}

  getHeight(): number {
    if (this.data.height) {
      return parseNumberTwips(this.data.height);
    }
    return 170; // default ~0.3cm
  }

  toXML(colIndex: number, tableOptions: TableOptions): string {
    const vAlign = this.data.vAlign || "center";
    const align = this.data.hAlign || "center";

    let xml = "        <w:tc>\n";

    xml += "          <w:tcPr>\n";
    const colWidth = parseNumberTwips(tableOptions.columns[colIndex].width);
    xml += `            <w:tcW w:w="${colWidth}" w:type="dxa"/>\n`;
    xml += `            <w:vAlign w:val="${vAlign}"/>\n`;

    if (this.data.colspan && this.data.colspan > 1) {
      xml += `            <w:gridSpan w:val="${this.data.colspan}"/>\n`;
    }

    if (this.data.rowspan && this.data.rowspan > 1) {
      xml += `            <w:vMerge w:val="restart"/>\n`;
    } else if (this.data.rowspan === 0) {
      xml += `            <w:vMerge/>\n`;
    }

    if (this.data.cellColor) {
      xml +=
        `            <w:shd w:val="clear" w:color="auto" w:fill="${this.data.cellColor}"/>\n`;
    }

    xml += "          </w:tcPr>\n";

    xml += "          <w:p>\n";
    xml += "            <w:pPr>\n";
    const jc = align === "justify" ? "both" : align;
    xml += `              <w:jc w:val="${jc}"/>\n`;
    xml += "            </w:pPr>\n";
    xml += "            <w:r>\n";

    if (
      this.data.bold || this.data.italic || this.data.underline ||
      this.data.fontSize || this.data.fontColor || this.data.fontName
    ) {
      xml += "              <w:rPr>\n";
      if (this.data.bold) xml += "                <w:b/>\n";
      if (this.data.italic) xml += "                <w:i/>\n";
      if (this.data.underline) {
        xml += '                <w:u w:val="single"/>\n';
      }
      if (this.data.fontSize) {
        xml += `                <w:sz w:val="${
          ptToHalfPoints(this.data.fontSize)
        }"/>\n`;
      }
      if (this.data.fontColor) {
        xml += `                <w:color w:val="${this.data.fontColor}"/>\n`;
      }
      if (this.data.fontName) {
        xml +=
          `                <w:rFonts w:ascii="${this.data.fontName}" w:hAnsi="${this.data.fontName}"/>\n`;
      }
      xml += "              </w:rPr>\n";
    }

    xml += `              <w:t xml:space="preserve">${
      escapeXML(this.data.text)
    }</w:t>\n`;
    xml += "            </w:r>\n";
    xml += "          </w:p>\n";
    xml += "        </w:tc>\n";

    return xml;
  }
}
