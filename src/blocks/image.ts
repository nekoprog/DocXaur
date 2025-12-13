/**
 * Image block.
 * @module
 */
import { cmToEmu, parseImageSize } from "../core/utils.ts";
import type { ImageOptions } from "../core/docxaur.ts";
import { Element, Section } from "./section.ts";

export class Image extends Element {
  private static drawingCounter = 1;

  static nextId(): number {
    return Image.drawingCounter++;
  }

  constructor(
    private imageId: number,
    private section: Section,
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
    const align = this.options.align ?? "center";
    const relId = this.section._doc().getImageRelId(this.imageId);
    const drawId = Image.nextId();
    let xml = "  <w:p>\n";
    if (align !== "left") {
      xml += "    <w:pPr>\n";
      xml += `      <w:jc w:val="${align}"/>\n`;
      xml += "    </w:pPr>\n";
    }
    xml += `    <w:r>
      <w:drawing>
        <wp:inline distT="0" distB="0" distL="0" distR="0">
          <wp:extent cx="${width}" cy="${height}"/>
          <wp:effectExtent l="0" t="0" r="0" b="0"/>
          <wp:docPr id="${drawId}" name="Picture ${drawId}"/>
          <wp:cNvGraphicFramePr>
            <a:graphicFrameLocks noChangeAspect="1"/>
          </wp:cNvGraphicFramePr>
          <a:graphic>
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:pic>
                <pic:nvPicPr>
                  <pic:cNvPr id="${drawId}" name="Picture ${drawId}"/>
                  <pic:cNvPicPr/>
                </pic:nvPicPr>
                <pic:blipFill>
                  <a:blip r:embed="${relId}"/>
                  <a:stretch><a:fillRect/></a:stretch>
                </pic:blipFill>
                <pic:spPr>
                  <a:xfrm><a:off x="0" y="0"/><a:ext cx="${width}" cy="${height}"/></a:xfrm>
                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                </pic:spPr>
              </pic:pic>
            </a:graphicData>
          </a:graphic>
        </wp:inline>
      </w:drawing>
    </w:r>
  </w:p>`;
    return xml;
  }
}
