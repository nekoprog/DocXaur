/**
 * DocXaur: semantic DOCX generator for Fresh islands.
 * Exposes the main API and performs packaging (ZIP).
 * @module
 */

import { BlobReader, BlobWriter, TextReader, ZipWriter } from "@zip-js/zip-js";
import {
  base64ToUint8Array,
  fetchImageAsBase64,
  parseNumberTwips,
  ptToHalfPoints,
} from "./utils.ts";
import { ensureImageRelationships } from "./relationships.ts";
import { Section } from "../blocks/section.ts";

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

export interface ImageOptions {
  width?: string;
  height?: string;
  align?: "left" | "center" | "right" | "justify";
}

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

/** Main class for building a document and downloading or producing a Blob. */
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

  /** Add a new section to the document. */
  addSection(options?: SectionOptions): Section {
    const section = new Section(options, this);
    this.sections.push(section);
    return section;
  }

  /** Register an image by URL; returns a numeric image id. */
  async registerImage(url: string): Promise<number> {
    if (this.images.has(url)) return this.images.get(url)!.id;
    const imageData = await fetchImageAsBase64(url);
    const id = this.imageCounter++;
    this.images.set(url, { ...imageData, id });
    return id;
  }

  /** Compute the relationship id (rId) for a given image id. */
  getImageRelId(imageId: number): string {
    // rId1-3 reserved for styles, fontTable, settings → images start at rId4
    return `rId${imageId + 3}`;
  }

  /** Generate a DOCX ZIP and return as Blob. */
  private async createDocxZip(): Promise<Blob> {
    const blobWriter = new BlobWriter();
    const zipWriter = new ZipWriter(blobWriter);

    // Base package parts
    await zipWriter.add(
      "[Content_Types].xml",
      new TextReader(this.generateContentTypes()),
    );
    await zipWriter.add("_rels/.rels", new TextReader(this.generateRootRels()));

    // 1) Build document first → allows tables to register images
    const docXml = await this.generateDocumentAsync();

    // 2) Build relationships with guard (all images are known)
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

    // Other parts
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

    // 3) Media binaries
    for (const [, imgData] of this.images) {
      const filename = `word/media/image${imgData.id}.${imgData.extension}`;
      await zipWriter.add(
        filename,
        new BlobReader(new Blob([base64ToUint8Array(imgData.data)])),
      );
    }

    await zipWriter.close();
    return await blobWriter.getData();
  }

  /** Return a Blob of the DOCX. */
  async toBlob(): Promise<Blob> {
    return await this.createDocxZip();
  }

  /** Trigger a browser download. */
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

  /** Internal: `[Content_Types].xml` */
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

  /** Internal: `_rels/.rels` */
  private generateRootRels(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
  }

  /** Internal: `word/_rels/document.xml.rels` (base + images) */
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
        `  <Relationship Id="rId${relIdNum}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image${d.id}.${d.extension}"/>\n`;
    }
    xml += `</Relationships>`;
    return xml;
  }

  /** Internal: assemble `word/document.xml` */
  private async generateDocumentAsync(): Promise<string> {
    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
`;
    for (const section of this.sections) {
      xml += await section.toXMLAsync();
    }
    if (this.sections.length > 0) {
      const last = this.sections[this.sections.length - 1];
      xml += last.getSectionPropertiesXML();
    }
    xml += `  </w:body>
</w:document>`;
    return xml;
  }

  /** Internal: `word/styles.xml` (defaults). */
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

  /** Internal: `word/fontTable.xml` */
  private generateFontTable(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:font w:name="Arial"/>
  <w:font w:name="Times New Roman"/>
  <w:font w:name="Calibri"/>
</w:fonts>`;
  }

  /** Internal: `word/settings.xml` */
  private generateSettings(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:zoom w:percent="100"/>
</w:settings>`;
  }
}
