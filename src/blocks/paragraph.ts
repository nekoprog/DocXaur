/**
 * Paragraph builder block.
 *
 * Use an instance to accumulate runs (text, breaks, tabs, page breaks) and
 * then call `toXML()` to get the WordprocessingML fragment for the paragraph.
 *
 * Example:
 *   section.paragraph().text("Hello").text(" world", { bold: true });
 */
import { escapeXML, ptToHalfPoints } from "../core/utils.ts";
import type { ParagraphOptions } from "../core/docxaur.ts";
import { Element } from "./element.ts";

interface TextRun {
  text: string;
  style?: ParagraphOptions;
}

type ParagraphOperation = () => void;

export class Paragraph extends Element {
  private runs: TextRun[] = [];
  private options: ParagraphOptions;
  private operations: ParagraphOperation[] = [];
  private isBuilt = false;

  constructor(options: ParagraphOptions = {}) {
    super();
    this.options = options;
  }

  text(text: string, style?: ParagraphOptions): this {
    this.operations.push(() => this.runs.push({ text, style }));
    return this;
  }

  tab(): this {
    this.operations.push(() => this.runs.push({ text: "\t" }));
    return this;
  }

  lineBreak(count = 1): this {
    this.operations.push(() => {
      for (let i = 0; i < count; i++) {
        this.runs.push({ text: "\n" });
      }
    });
    return this;
  }

  pageBreak(count = 1): this {
    this.operations.push(() => {
      for (let i = 0; i < count; i++) {
        this.runs.push({ text: "[PAGE_BREAK]" });
      }
    });
    return this;
  }

  /** @deprecated Prefer explicit method calls instead of apply(). */
  apply(...operations: ((builder: this) => this)[]): this {
    console.warn(
      "Paragraph.apply() is deprecated. Use direct method calls instead.",
    );
    for (const op of operations) {
      op(this);
    }
    return this;
  }

  private hasRunProperties(style: ParagraphOptions): boolean {
    return !!(
      style.bold ||
      style.italic ||
      style.underline ||
      style.fontSize ||
      style.fontColor ||
      style.fontName
    );
  }

  private build(): void {
    if (this.isBuilt) return;
    this.isBuilt = true;
    for (const op of this.operations) {
      op();
    }
  }

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
    for (let i = 0; i < breaksBefore; i++) {
      xml += "    <w:r><w:br/></w:r>\n";
    }
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
    for (let i = 0; i < breaksAfter; i++) {
      xml += "    <w:r><w:br/></w:r>\n";
    }
    xml += "  </w:p>";
    return xml;
  }
}
