/**
 * Paragraph builder block.
 *
 * Use an instance to accumulate runs (text, breaks, tabs, page breaks) and
 * then call `toXML()` to get the WordprocessingML fragment for the paragraph.
 */

import { escapeXML, ptToHalfPoints } from "../core/utils.ts";
import type { ParagraphOptions } from "../core/docxaur.ts";
import { Element } from "./element.ts";
import {
  createShapeRun,
  SHAPE_CIRCLE,
  SHAPE_DIAMOND,
  SHAPE_HEART,
  SHAPE_HEXAGON,
  SHAPE_PENTAGON,
  SHAPE_RECT,
  SHAPE_ROUNDED_RECT,
  SHAPE_STAR_5,
  SHAPE_TRIANGLE,
  type ShapeOptions,
  type ShapeType,
} from "./shapes.ts";

interface TextRun {
  text: string;
  style?: ParagraphOptions;
  isShape?: boolean;
}

type ParagraphOperation = () => void;

/**
 * Paragraph builder block with text, breaks, and shape support.
 *
 * Represents `<w:p>` with inline styling, line breaks, page breaks,
 * and geometric shapes. Chain methods for readable composition.
 */
export class Paragraph extends Element {
  private runs: TextRun[] = [];
  private options: ParagraphOptions;
  private operations: ParagraphOperation[] = [];
  private isBuilt = false;

  /**
   * Creates a new paragraph.
   *
   * @param {ParagraphOptions} [options] - Paragraph styling
   */
  constructor(options: ParagraphOptions = {}) {
    super();
    this.options = options;
  }

  /**
   * Adds text with optional inline styling.
   *
   * @param {string} text - Text content
   * @param {ParagraphOptions} [style] - Run styling
   * @returns {this}
   */
  text(text: string, style?: ParagraphOptions): this {
    this.operations.push(() => this.runs.push({ text, style }));
    return this;
  }

  /**
   * Adds a tab character.
   *
   * @returns {this}
   */
  tab(): this {
    this.operations.push(() => this.runs.push({ text: "\t" }));
    return this;
  }

  /**
   * Adds line breaks.
   *
   * @param {number} [count] - Number of line breaks (default: 1)
   * @returns {this}
   */
  lineBreak(count = 1): this {
    this.operations.push(() => {
      for (let i = 0; i < count; i++) {
        this.runs.push({ text: "\n" });
      }
    });
    return this;
  }

  /**
   * Adds page breaks.
   *
   * @param {number} [count] - Number of page breaks (default: 1)
   * @returns {this}
   */
  pageBreak(count = 1): this {
    this.operations.push(() => {
      for (let i = 0; i < count; i++) {
        this.runs.push({ text: "[PAGE_BREAK]" });
      }
    });
    return this;
  }

  /**
   * Inserts a rectangle shape.
   *
   * @param {ShapeOptions} [options] - Shape configuration
   * @returns {this}
   */
  rectangle(options?: ShapeOptions): this {
    this.operations.push(() => {
      this.runs.push(createShapeRun(SHAPE_RECT, options));
    });
    return this;
  }

  /**
   * Inserts a rounded rectangle shape.
   *
   * @param {ShapeOptions} [options] - Shape configuration
   * @returns {this}
   */
  roundedRectangle(options?: ShapeOptions): this {
    this.operations.push(() => {
      this.runs.push(createShapeRun(SHAPE_ROUNDED_RECT, options));
    });
    return this;
  }

  /**
   * Inserts a circle or ellipse shape.
   *
   * @param {ShapeOptions} [options] - Shape configuration
   * @returns {this}
   */
  circle(options?: ShapeOptions): this {
    this.operations.push(() => {
      this.runs.push(createShapeRun(SHAPE_CIRCLE, options));
    });
    return this;
  }

  /**
   * Inserts a diamond shape.
   *
   * @param {ShapeOptions} [options] - Shape configuration
   * @returns {this}
   */
  diamond(options?: ShapeOptions): this {
    this.operations.push(() => {
      this.runs.push(createShapeRun(SHAPE_DIAMOND, options));
    });
    return this;
  }

  /**
   * Inserts a triangle shape.
   *
   * @param {ShapeOptions} [options] - Shape configuration
   * @returns {this}
   */
  triangle(options?: ShapeOptions): this {
    this.operations.push(() => {
      this.runs.push(createShapeRun(SHAPE_TRIANGLE, options));
    });
    return this;
  }

  /**
   * Inserts a pentagon shape.
   *
   * @param {ShapeOptions} [options] - Shape configuration
   * @returns {this}
   */
  pentagon(options?: ShapeOptions): this {
    this.operations.push(() => {
      this.runs.push(createShapeRun(SHAPE_PENTAGON, options));
    });
    return this;
  }

  /**
   * Inserts a hexagon shape.
   *
   * @param {ShapeOptions} [options] - Shape configuration
   * @returns {this}
   */
  hexagon(options?: ShapeOptions): this {
    this.operations.push(() => {
      this.runs.push(createShapeRun(SHAPE_HEXAGON, options));
    });
    return this;
  }

  /**
   * Inserts a five-point star shape.
   *
   * @param {ShapeOptions} [options] - Shape configuration
   * @returns {this}
   */
  star5(options?: ShapeOptions): this {
    this.operations.push(() => {
      this.runs.push(createShapeRun(SHAPE_STAR_5, options));
    });
    return this;
  }

  /**
   * Inserts a heart shape.
   *
   * @param {ShapeOptions} [options] - Shape configuration
   * @returns {this}
   */
  heart(options?: ShapeOptions): this {
    this.operations.push(() => {
      this.runs.push(createShapeRun(SHAPE_HEART, options));
    });
    return this;
  }

  /**
   * Inserts a generic shape by preset.
   *
   * @param {ShapeType} shapeType - Shape preset identifier
   * @param {ShapeOptions} [options] - Shape configuration
   * @returns {this}
   */
  shape(shapeType: ShapeType, options?: ShapeOptions): this {
    this.operations.push(() => {
      this.runs.push(createShapeRun(shapeType, options));
    });
    return this;
  }

  /**
   * Applies conditional operations.
   *
   * @deprecated Use direct method calls instead.
   * @param {...Function[]} operations - Operations to apply
   * @returns {this}
   */
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

  /**
   * Generates OOXML for this paragraph.
   *
   * Produces a `<w:p>` element with alignment, spacing, and content runs.
   * Handles text, breaks, and shapes inline.
   *
   * @returns {string} WordprocessingML paragraph element
   */
  toXML(): string {
    this.build();
    const hAlign = this.options.hAlign ?? "left";
    const breaksBefore = this.options.breakBefore ?? 0;
    const breaksAfter = this.options.breakAfter ?? 0;

    let xml = "  <w:p>\n";
    xml += "    <w:pPr>\n";
    if (hAlign !== "left") {
      const jc = hAlign === "justify" ? "both" : hAlign;
      xml += `      <w:jc w:val="${jc}"/>\n`;
    }
    const spacing = this.options.spacing;
    const before = spacing?.before ? Math.round(spacing.before * 20) : 0;
    const after = spacing?.after ? Math.round(spacing.after * 20) : 0;
    const line = spacing?.line ? Math.round(spacing.line * 240) : 240;
    xml +=
      `      <w:spacing w:after="${after}" w:before="${before}" w:line="${line}" w:lineRule="auto"/>\n`;

    if (this.options.baselineAlignment) {
      xml +=
        `      <w:textAlignment w:val="${this.options.baselineAlignment}"/>\n`;
    }

    xml += "    </w:pPr>\n";

    for (let i = 0; i < breaksBefore; i++) {
      xml += "    <w:r><w:br/></w:r>\n";
    }

    for (const run of this.runs) {
      if ((run as any).isShape) {
        xml += run.text;
      } else if (run.text === "\t") {
        xml += "    <w:r><w:tab/></w:r>\n";
      } else if (run.text === "\n") {
        xml += "    <w:r><w:br/></w:r>\n";
      } else if (run.text === "[PAGE_BREAK]") {
        xml += '    <w:r><w:br w:type="page"/></w:r>\n';
      } else {
        xml += "    <w:r>\n";
        const style = run.style;
        if (
          style && (
            style.bold || style.italic || style.underline ||
            style.fontSize || style.fontColor || style.fontName ||
            style.baselineShift !== undefined
          )
        ) {
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

          if (typeof style?.baselineShift === "number") {
            const vPt = style.baselineShift; // points
            xml += `        <w:position w:val="${ptToHalfPoints(vPt)}"/>\n`; // half-points
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
    xml += "  <w:p>".replace("<w:p>", "</w:p>");
    return xml;
  }
}
