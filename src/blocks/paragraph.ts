/**
 * Paragraph builder and `section.paragraph()` augmentation.
 * @module
 */

import { escapeXML, ptToHalfPoints } from "../core/utils.ts";
import type { ParagraphOptions } from "../core/docxaur.ts";
import { Element, Section } from "./section.ts";

/** Inline text run. */
interface TextRun {
  text: string;
  style?: ParagraphOptions;
}
type ParagraphOperation = () => void;

/** Paragraph element. */
export class Paragraph extends Element {
  private runs: TextRun[] = [];
  private options: ParagraphOptions;
  private operations: ParagraphOperation[] = [];
  private isBuilt = false;

  constructor(options: ParagraphOptions = {}) {
    super();
    this.options = options;
  }

  /** Add text with optional inline styling. */
  text(text: string, style?: ParagraphOptions): this {
    this.operations.push(() => this.runs.push({ text, style }));
    return this;
  }

  /** Insert a tab character. */
  tab(): this {
    this.operations.push(() => this.runs.push({ text: "\t" }));
    return this;
  }

  /** Add line breaks. */
  lineBreak(count: number = 1): this {
    this.operations.push(() => {
      for (let i = 0; i < count; i++) this.runs.push({ text: "\n" });
    });
    return this;
  }

  /** Add page breaks. */
  pageBreak(count: number = 1): this {
    this.operations.push(() => {
      for (let i = 0; i < count; i++) this.runs.push({ text: "[PAGE_BREAK]" });
    });
    return this;
  }

  /**
   * @deprecated Use direct method calls (`text()`, `lineBreak()`, etc.). This method will be removed.
   */
  apply(...operations: ((builder: this) => this)[]): this {
    console.warn(
      "Paragraph.apply() is deprecated. Use direct method calls instead.",
    );
    for (const op of operations) op(this);
    return this;
  }

  private hasRunProperties(style: ParagraphOptions): boolean {
    return !!(style.bold || style.italic || style.underline || style.fontSize ||
      style.fontColor || style.fontName);
  }

  private build(): void {
    if (this.isBuilt) return;
    this.isBuilt = true;
    for (const op of this.operations) op();
  }

  /** Emit OOXML for the paragraph. */
  toXML(): string {
    this.build();
    const align = this.options.align ?? "left";
    const breaksBefore = this.options.breakBefore ?? 0;
    const breaksAfter = this.options.breakAfter ?? 0;

    let xml = "  <w:p>\n";
    xml += "    <w:pPr>\n";
    if (align !== "left") {
      const jc = align === "justify" ? "both" : align;
      xml += `      <w:jc w:val="${jc}"/>\n`;
    }
    const spacing = this.options.spacing;
    const before = spacing?.before ? ptToHalfPoints(spacing.before) * 20 : 0;
    const after = spacing?.after ? ptToHalfPoints(spacing.after) * 20 : 0;
    const line = spacing?.line ? Math.round(spacing.line * 240) : 240;
    xml +=
      `      <w:spacing w:after="${after}" w:before="${before}" w:line="${line}" w:lineRule="auto"/>\n`;
    xml += "    </w:pPr>\n";

    for (let i = 0; i < breaksBefore; i++) xml += "    <w:r><w:br/></w:r>\n";

    for (const run of this.runs) {
      if (run.text === "\t") {
        xml += "    <w:r><w:tab/></w:r>\n";
      } else if (run.text === "\n") {
        xml += "    <w:r><w:br/></w:r>\n";
      } else if (run.text === "[PAGE_BREAK]") {
        xml += '    <w:r><w:br w:type="page"/></w:r>\n';
      } else {
        xml += "    <w:r>\n";
        const style = run.style;
        if (style && this.hasRunProperties(style)) {
          xml += "      <w:rPr>\n";
          if (style.bold) xml += "        <w:b/>\n";
          if (style.italic) xml += "        <w:i/>\n";
          if (style.underline) xml += '        <w:u w:val="single"/>\n';
          if (style.fontSize) {
            xml += `        <w:sz w:val="${
              ptToHalfPoints(style.fontSize)
            }"/>\n`;
          }
          if (style.fontColor) {
            xml += `        <w:color w:val="${style.fontColor}"/>\n`;
          }
          if (style.fontName) {
            xml +=
              `        <w:rFonts w:ascii="${style.fontName}" w:hAnsi="${style.fontName}"/>\n`;
          }
          xml += "      </w:rPr>\n";
        }
        xml += `      <w:t xml:space="preserve">${escapeXML(run.text)}</w:t>\n`;
        xml += "    </w:r>\n";
      }
    }

    for (let i = 0; i < breaksAfter; i++) xml += "    <w:r><w:br/></w:r>\n";
    xml += "  </w:p>";
    return xml;
  }
}

/** Module augmentation: add `paragraph()` and `heading()` to Section. */
declare module "./section.ts" {
  interface Section {
    paragraph(options?: ParagraphOptions): Paragraph;
    heading(
      content: string,
      level?: 1 | 2 | 3 | 4 | 5 | 6,
      options?: ParagraphOptions,
    ): Section;
  }
}
(Section.prototype as any).paragraph = function paragraph(
  this: Section,
  options?: ParagraphOptions,
): Paragraph {
  const p = new Paragraph(options);
  this._push(p);
  return p;
};
(Section.prototype as any).heading = function heading(
  this: Section,
  content: string,
  level: 1 | 2 | 3 | 4 | 5 | 6 = 1,
  options?: ParagraphOptions,
): Section {
  const sizes = [24, 20, 18, 16, 14, 12];
  const p = new Paragraph(options);
  p.text(content, { bold: true, fontSize: sizes[level - 1], ...options });
  this._push(p);
  return this;
};
