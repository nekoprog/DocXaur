/**
 * Base Section: holds elements and common APIs.
 * @module
 */

import { parseNumberTwips } from "../core/utils.ts";
import type { DocXaur, SectionOptions } from "../core/docxaur.ts";

export abstract class Element {
  abstract toXML(): string;
}

/** A document section containing paragraphs, images, tables, etc. */
export class Section {
  private elements: Element[] = [];
  private options: SectionOptions;
  private doc: DocXaur;

  constructor(options: SectionOptions = {}, doc: DocXaur) {
    this.options = {
      pageSize: options.pageSize ??
        { width: "21cm", height: "29.7cm", orientation: "portrait" },
      margins: options.margins ??
        { top: "2.54cm", right: "2.54cm", bottom: "2.54cm", left: "2.54cm" },
    };
    this.doc = doc;
  }

  /** Internal: push an element to this section. */
  _push(el: Element): void {
    this.elements.push(el);
  }

  /** Internal: access to DocXaur. */
  _doc(): DocXaur {
    return this.doc;
  }

  /** Build any deferred elements and return XML for all children. */
  async toXMLAsync(): Promise<string> {
    let xml = "";
    for (const el of this.elements) xml += el.toXML() + "\n";
    return xml;
  }

  /** Emit section properties (page size + margins). */
  getSectionPropertiesXML(): string {
    const pageSize = this.options.pageSize!;
    const margins = this.options.margins!;
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
}
