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

type TableBridge = Element & { _table?: Table };

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
        if (bridge._table._buildPromise) {
          await bridge._table._buildPromise;
        }
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

  // === Block methods ===
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

  table(options: TableOptions): Table {
    const t = new Table(options);
    t._buildPromise = t.buildRows(this);

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
