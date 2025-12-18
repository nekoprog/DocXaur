/**
 * Section — holds a sequence of Elements and provides convenience methods
 * for creating blocks (paragraph, image, table, heading, etc.) in this section.
 *
 * The `Section` is responsible for maintaining section options (page size,
 * margins) and for producing section-level properties when building the doc.
 */
import { parseNumberTwips } from "../core/utils.ts";
import type {
  DocXaur,
  ImageOptions,
  ParagraphOptions,
  SectionOptions,
  TableOptions,
} from "../core/docxaur.ts";
import { Paragraph } from "./paragraph.ts";
import { Image } from "./image.ts";
import { Table } from "./table.ts";
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

type TableBridge = Element & { _table?: Table };

/**
 * Section — holds a sequence of Elements and provides convenience methods
 * for creating blocks in this section and producing the section XML.
 */
export class Section {
  private elements: Element[] = [];
  private options: Required<SectionOptions>;
  private doc: DocXaur;

  constructor(options: SectionOptions = {}, doc: DocXaur) {
    this.options = {
      pageSize: options.pageSize ?? {
        width: "21cm",
        height: "29.7cm",
        orientation: "portrait",
      },
      margins: options.margins ?? {
        top: "2.54cm",
        right: "2.54cm",
        bottom: "2.54cm",
        left: "2.54cm",
      },
    };
    this.doc = doc;
  }

  _push(el: Element): void {
    this.elements.push(el);
  }

  _doc(): DocXaur {
    return this.doc;
  }

  async toXMLAsync(): Promise<string> {
    let xml = "";
    for (const el of this.elements) {
      const bridge = el as TableBridge;
      if (bridge._table) {
        await bridge._table.buildRows(this);
      }
      xml += el.toXML() + "\n";
    }
    return xml;
  }

  getSectionPropertiesXML(): string {
    const pageSize = this.options.pageSize;
    const margins = this.options.margins;
    const widthTwips = parseNumberTwips(pageSize.width);
    const heightTwips = parseNumberTwips(pageSize.height);
    const orient = pageSize.orientation === "landscape"
      ? "landscape"
      : "portrait";
    return `  <w:sectPr>
     <w:pgSz w:w="${widthTwips}" w:h="${heightTwips}" w:orient="${orient}"/>
     <w:pgMar w:top="${parseNumberTwips(margins.top)}"
              w:right="${parseNumberTwips(margins.right)}"
              w:bottom="${parseNumberTwips(margins.bottom)}"
              w:left="${parseNumberTwips(margins.left)}"
              w:header="720" w:footer="720" w:gutter="0"/>
   </w:sectPr>
 `;
  }

  paragraph(options?: ParagraphOptions): Paragraph {
    const p = new Paragraph(options);
    this._push(p);
    return p;
  }

  heading(
    content: string,
    level: 1 | 2 | 3 | 4 | 5 | 6 = 1,
    options?: ParagraphOptions,
  ): this {
    const sizes = [24, 20, 18, 16, 14, 12];
    const p = new Paragraph(options);
    p.text(content, { bold: true, fontSize: sizes[level - 1], ...options });
    this._push(p);
    return this;
  }

  async image(url: string, options?: ImageOptions): Promise<this> {
    const id = await this._doc().registerImage(url);
    this._push(new Image(id, this, options));
    return this;
  }

  /**
   * Inserts a rectangle shape.
   *
   * @param {ShapeOptions} [options] - Shape configuration
   * @returns {Paragraph}
   */
  rectangle(options?: ShapeOptions): Paragraph {
    const p = new Paragraph();
    p.rectangle(options);
    this._push(p);
    return p;
  }

  /**
   * Inserts a rounded rectangle shape.
   *
   * @param {ShapeOptions} [options] - Shape configuration
   * @returns {Paragraph}
   */
  roundedRectangle(options?: ShapeOptions): Paragraph {
    const p = new Paragraph();
    p.roundedRectangle(options);
    this._push(p);
    return p;
  }

  /**
   * Inserts a circle or ellipse shape.
   *
   * @param {ShapeOptions} [options] - Shape configuration
   * @returns {Paragraph}
   */
  circle(options?: ShapeOptions): Paragraph {
    const p = new Paragraph();
    p.circle(options);
    this._push(p);
    return p;
  }

  /**
   * Inserts a diamond shape.
   *
   * @param {ShapeOptions} [options] - Shape configuration
   * @returns {Paragraph}
   */
  diamond(options?: ShapeOptions): Paragraph {
    const p = new Paragraph();
    p.diamond(options);
    this._push(p);
    return p;
  }

  /**
   * Inserts a triangle shape.
   *
   * @param {ShapeOptions} [options] - Shape configuration
   * @returns {Paragraph}
   */
  triangle(options?: ShapeOptions): Paragraph {
    const p = new Paragraph();
    p.triangle(options);
    this._push(p);
    return p;
  }

  /**
   * Inserts a pentagon shape.
   *
   * @param {ShapeOptions} [options] - Shape configuration
   * @returns {Paragraph}
   */
  pentagon(options?: ShapeOptions): Paragraph {
    const p = new Paragraph();
    p.pentagon(options);
    this._push(p);
    return p;
  }

  /**
   * Inserts a hexagon shape.
   *
   * @param {ShapeOptions} [options] - Shape configuration
   * @returns {Paragraph}
   */
  hexagon(options?: ShapeOptions): Paragraph {
    const p = new Paragraph();
    p.hexagon(options);
    this._push(p);
    return p;
  }

  /**
   * Inserts a five-point star shape.
   *
   * @param {ShapeOptions} [options] - Shape configuration
   * @returns {Paragraph}
   */
  star5(options?: ShapeOptions): Paragraph {
    const p = new Paragraph();
    p.star5(options);
    this._push(p);
    return p;
  }

  /**
   * Inserts a heart shape.
   *
   * @param {ShapeOptions} [options] - Shape configuration
   * @returns {Paragraph}
   */
  heart(options?: ShapeOptions): Paragraph {
    const p = new Paragraph();
    p.heart(options);
    this._push(p);
    return p;
  }

  /**
   * Inserts a generic shape by preset.
   *
   * @param {ShapeType} shapeType - Shape preset identifier
   * @param {ShapeOptions} [options] - Shape configuration
   * @returns {Paragraph}
   */
  shape(shapeType: ShapeType, options?: ShapeOptions): Paragraph {
    const p = new Paragraph();
    p.shape(shapeType, options);
    this._push(p);
    return p;
  }

  table(options: TableOptions): Table {
    const t = new Table(options);

    const bridgeSection = this;
    const bridge: TableBridge = new (class extends Element {
      toXML(): string {
        return t._toXMLWithSection(bridgeSection);
      }
    })();
    bridge._table = t;

    this._push(bridge);
    return t;
  }
}
