/**
 * DocXaur: semantic DOCX generator for Fresh islands.
 * @module
 *
 * This module exports the main `DocXaur` builder and the public option
 * interfaces used by the rest of the library.
 */
import { BlobReader, BlobWriter, TextReader, ZipWriter } from "@zip-js/zip-js";
import {
  base64ToArrayBuffer,
  fetchImageAsBase64,
  parseNumberTwips,
  ptToHalfPoints,
} from "./utils.ts";
import { ensureImageRelationships } from "./relationships.ts";
import { Section } from "../blocks/section.ts";

/** Document-level configuration options. */
export interface DocumentOptions {
  /** Title of the document. */
  title?: string;
  /** Creator/author. */
  creator?: string;
  /** Short description. */
  description?: string;
  /** Subject. */
  subject?: string;
  /** Keywords for the document. */
  keywords?: string;
  /** Default font family for the document. */
  fontName?: string;
  /** Default font size (points). */
  fontSize?: number;
}

/** Options for constructing a Section. */
export interface SectionOptions {
  /** Page size for the section. */
  pageSize?: PageSize;
  /** Margins for the section. */
  margins?: Margins;
}

/** Page size specification. */
export interface PageSize {
  width: string;
  height: string;
  orientation?: "portrait" | "landscape";
}

/** Page margins specification. */
export interface Margins {
  top: string;
  right: string;
  bottom: string;
  left: string;
}

/** Options when creating or embedding images. */
export interface ImageOptions {
  width?: string;
  height?: string;
  align?: "left" | "center" | "right" | "justify";
}

/** Paragraph styling options. */
export interface ParagraphOptions {
  align?: "left" | "center" | "right" | "justify";
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number;
  fontColor?: string;
  fontName?: string;
  spacing?: { before?: number; after?: number; line?: number };
  breakBefore?: number;
  breakAfter?: number;
}

/** Table column configuration. */
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
  marginTop?: string;
  marginRight?: string;
  marginBottom?: string;
  marginLeft?: string;
}

/** Options for Table creation. */
export interface TableOptions {
  columns: TableColumn[];
  width?: string;
  align?: "left" | "center" | "right" | "justify";
  borders?: boolean;
  indent?: string;
  marginTop?: string;
  marginRight?: string;
  marginBottom?: string;
  marginLeft?: string;
}

/** Table cell data shape. */
export interface TableCellData {
  text?: string;
  image?: { url: string; width?: string; height?: string };
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
  marginTop?: string;
  marginRight?: string;
  marginBottom?: string;
  marginLeft?: string;
}

type ImageMapValue = { data: string; extension: string; id: number };

/**
 * The main DocXaur builder class.
 *
 * Use `new DocXaur(options)` to create a document, add sections via
 * `.addSection(...)`, then call `.toBlob()` or `.download()` to obtain the
 * generated .docx file.
 */
export class DocXaur {
  private sections: Section[] = [];
  private options: DocumentOptions;
  private images: Map<string, ImageMapValue> = new Map();
  private imageCounter = 1;
  private fontName: string;
  private fontSize: number;

  constructor(options: DocumentOptions = {}) {
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
    if (this.images.has(url)) return this.images.get(url)!.id;
    const imageData = await fetchImageAsBase64(url);
    const id = this.imageCounter++;
    this.images.set(url, { ...imageData, id });
    return id;
  }

  getImageRelId(imageId: number): string {
    return `rId${imageId + 3}`;
  }

  private async createDocxZip(): Promise<Blob> {
    const blobWriter = new BlobWriter();
    const zipWriter = new ZipWriter(blobWriter);

    await zipWriter.add(
      "[Content_Types].xml",
      new TextReader(this.generateContentTypes()),
    );
    await zipWriter.add("_rels/.rels", new TextReader(this.generateRootRels()));
    await zipWriter.add(
      "docProps/core.xml",
      new TextReader(this.generateCoreProperties()),
    );
    await zipWriter.add(
      "docProps/app.xml",
      new TextReader(this.generateAppProperties()),
    );

    const docXml = await this.generateDocumentAsync();

    const rawDocRels = this.generateDocRels();
    const imageRels = Array.from(this.images.entries()).map(([url, d]) => ({
      rid: `rId${d.id + 3}`,
      target: `media/image${d.id}.${d.extension}`,
    }));
    const fixedDocRels = ensureImageRelationships(rawDocRels, imageRels);

    await zipWriter.add(
      "word/_rels/document.xml.rels",
      new TextReader(fixedDocRels),
    );
    await zipWriter.add("word/document.xml", new TextReader(docXml));
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

    for (const [, imgData] of this.images) {
      const filename = `word/media/image${imgData.id}.${imgData.extension}`;
      const arrayBuffer = base64ToArrayBuffer(imgData.data);
      await zipWriter.add(filename, new BlobReader(new Blob([arrayBuffer])));
    }

    await zipWriter.close();
    return await blobWriter.getData();
  }

  async toBlob(): Promise<Blob> {
    return await this.createDocxZip();
  }

  async download(filename: string = "document.docx"): Promise<void> {
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
  <Default Extension="xml"  ContentType="application/xml"/>
  <Default Extension="png"  ContentType="image/png"/>
  <Default Extension="jpg"  ContentType="image/jpeg"/>
  <Default Extension="jpeg" ContentType="image/jpeg"/>
  <Default Extension="gif"  ContentType="image/gif"/>
  <Default Extension="bmp"  ContentType="image/bmp"/>
  <Override PartName="/word/document.xml"  ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
  <Override PartName="/word/settings.xml"  ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/docProps/core.xml"  ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml"   ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
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
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"    Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"  Target="settings.xml"/>
`;
    const sorted = Array.from(this.images.entries()).sort((a, b) =>
      a[1].id - b[1].id
    );
    for (const [, d] of sorted) {
      const relIdNum = d.id + 3;
      xml +=
        `  <Relationship Id="rId${relIdNum}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image${d.id}.${d.extension}"/>
`;
    }
    xml += `</Relationships>`;
    return xml;
  }

  private async generateDocumentAsync(): Promise<string> {
    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
`;
    for (const section of this.sections) xml += await section.toXMLAsync();
    if (this.sections.length > 0) {
      const last = this.sections[this.sections.length - 1];
      xml += last.getSectionPropertiesXML();
    }
    xml += `  </w:body>
</w:document>`;
    return xml;
  }

  private generateStyles(): string {
    const fontSize = ptToHalfPoints(this.fontSize);
    const fontName = this.fontName;
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="${fontName}" w:hAnsi="${fontName}" w:cs="${fontName}"/>
        <w:sz   w:val="${fontSize}"/>
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
  <w:font w:name="Arial"/>
  <w:font w:name="Times New Roman"/>
  <w:font w:name="Calibri"/>
</w:fonts>`;
  }

  private generateSettings(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:zoom w:percent="100"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
</w:settings>`;
  }

  private generateCoreProperties(): string {
    const now = new Date().toISOString();
    const esc = (s: string) => this.escapeXml(s ?? "");
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <cp:coreProperties
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:dcmitype="http://purl.org/dc/dcmitype/"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <dc:title>${esc(this.options.title ?? "")}</dc:title>
    <dc:creator>${esc(this.options.creator ?? "")}</dc:creator>
    <dc:description>${esc(this.options.description ?? "")}</dc:description>
    <dc:subject>${esc(this.options.subject ?? "")}</dc:subject>
    <cp:keywords>${esc(this.options.keywords ?? "")}</cp:keywords>
    <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
    <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
  </cp:coreProperties>`;
  }

  private generateAppProperties(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <TotalTime>0</TotalTime>
  <Application>DocXaur</Application>
  <AppVersion>1.0</AppVersion>
</Properties>`;
  }

  private escapeXml(text: string): string {
    return text
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }
}
