/**
 * Table block.
 * @module
 */
import { cmToEmu, parseImageSize, parseNumberTwips } from "../core/utils.ts";
import type {
  TableCellData,
  TableColumn,
  TableOptions,
} from "../core/docxaur.ts";
import { Element, Section } from "./section.ts";
import { Image } from "./image.ts";

class TableRow {
  private cells: TableCell[] = [];

  constructor(private tableOptions: TableOptions) {}

  cell(data: TableCellData): this {
    this.cells.push(new TableCell(data));
    return this;
  }

  async initCells(section: Section): Promise<void> {
    await Promise.all(this.cells.map((cell) => cell.init(section)));
  }

  toXML(section: Section): string {
    let xml = "  <w:tr>\n";
    let maxHeight = 170;
    for (const cell of this.cells) {
      const h = cell.getHeight();
      if (h > maxHeight) maxHeight = h;
    }
    xml += `    <w:trPr>
      <w:trHeight w:val="${maxHeight}" w:hRule="atLeast"/>
    </w:trPr>
`;
    for (let i = 0; i < this.cells.length; i++) {
      xml += this.cells[i].toXML(i, this.tableOptions, section);
    }
    xml += "  </w:tr>\n";
    return xml;
  }
}

class TableCell {
  private imageId?: number;

  constructor(private data: TableCellData) {}

  async init(section: Section): Promise<void> {
    if (this.data.image) {
      this.imageId = await section._doc().registerImage(this.data.image.url);
    }
  }

  getHeight(): number {
    if (this.data.height) return parseNumberTwips(this.data.height);
    return 170;
  }

  toXML(
    colIndex: number,
    tableOptions: TableOptions,
    section: Section,
  ): string {
    const vAlign = this.data.vAlign ?? "center";
    const align = this.data.hAlign ?? "center";
    let xml = "    <w:tc>\n";
    xml += "      <w:tcPr>\n";
    const colWidth = parseNumberTwips(tableOptions.columns[colIndex].width);
    xml += `        <w:tcW w:w="${colWidth}" w:type="dxa"/>\n`;
    xml += `        <w:vAlign w:val="${vAlign}"/>\n`;
    if (this.data.colspan && this.data.colspan > 1) {
      xml += `        <w:gridSpan w:val="${this.data.colspan}"/>\n`;
    }
    if (this.data.rowspan && this.data.rowspan > 1) {
      xml += `        <w:vMerge w:val="restart"/>\n`;
    } else if (this.data.rowspan === 0) {
      xml += `        <w:vMerge/>\n`;
    }
    if (this.data.cellColor) {
      xml +=
        `        <w:shd w:val="clear" w:color="auto" w:fill="${this.data.cellColor}"/>\n`;
    }
    const col = tableOptions.columns[colIndex];
    const mt = this.data.marginTop ?? col?.marginTop ?? tableOptions.marginTop;
    const mr = this.data.marginRight ?? col?.marginRight ??
      tableOptions.marginRight;
    const mb = this.data.marginBottom ?? col?.marginBottom ??
      tableOptions.marginBottom;
    const ml = this.data.marginLeft ?? col?.marginLeft ??
      tableOptions.marginLeft;
    if (
      mt !== undefined || mr !== undefined || mb !== undefined ||
      ml !== undefined
    ) {
      xml += "        <w:tcMar>\n";
      if (mt !== undefined) {
        xml += `          <w:top   w:w="${
          parseNumberTwips(mt)
        }" w:type="dxa"/>\n`;
      }
      if (mr !== undefined) {
        xml += `          <w:end   w:w="${
          parseNumberTwips(mr)
        }" w:type="dxa"/>\n`;
      }
      if (mb !== undefined) {
        xml += `          <w:bottom w:w="${
          parseNumberTwips(mb)
        }" w:type="dxa"/>\n`;
      }
      if (ml !== undefined) {
        xml += `          <w:start w:w="${
          parseNumberTwips(ml)
        }" w:type="dxa"/>\n`;
      }
      xml += "        </w:tcMar>\n";
    }
    xml += "      </w:tcPr>\n";

    if (this.data.image && this.imageId !== undefined) {
      const img = this.data.image;
      const width = img.width ? parseImageSize(img.width) : cmToEmu(5);
      const height = img.height ? parseImageSize(img.height) : width;
      const hAlign = this.data.hAlign ?? "center";
      const relId = section._doc().getImageRelId(this.imageId);
      const drawId = Image.nextId();
      xml += "      <w:p>\n";
      if (hAlign !== "left") {
        xml += "        <w:pPr>\n";
        xml += `          <w:jc w:val="${hAlign}"/>\n`;
        xml += "        </w:pPr>\n";
      }
      xml += `        <w:r>
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
      </w:p>
`;
    } else {
      const jc = align === "justify" ? "both" : align;
      xml += "      <w:p>\n";
      xml += "        <w:pPr>\n";
      xml += `          <w:jc w:val="${jc}"/>\n`;
      xml += "        </w:pPr>\n";
      xml += `        <w:r><w:t xml:space="preserve">${
        this.data.text ?? ""
      }</w:t></w:r>\n`;
      xml += "      </w:p>\n";
    }
    xml += "    </w:tc>\n";
    return xml;
  }
}

export class Table extends Element {
  private rowDefs: Array<(string | TableCellData)[]> = [];
  private rows: TableRow[] = [];
  private options: TableOptions;
  private isBuilt = false;

  constructor(options: TableOptions) {
    super();
    this.options = options;
    if (this.options.borders === undefined) {
      this.options.borders = true;
    }
  }

  row(...cells: (string | TableCellData)[]): this {
    this.rowDefs.push(cells);
    return this;
  }

  /** @deprecated Use direct `.row()` calls. */
  apply(...ops: ((builder: this) => this)[]): this {
    console.warn(
      "Table.apply() is deprecated. Use direct .row() calls instead.",
    );
    for (const op of ops) {
      op(this);
    }
    return this;
  }

  async buildRows(section: Section): Promise<void> {
    if (this.isBuilt) return;
    this.isBuilt = true;
    for (const defs of this.rowDefs) {
      const row = new TableRow(this.options);
      defs.forEach((cell, i) => {
        const colOptions = this.options.columns[i];
        if (typeof cell === "string") {
          row.cell({
            text: cell,
            hAlign: colOptions?.hAlign ?? "center",
            vAlign: colOptions?.vAlign ?? "center",
            fontName: colOptions?.fontName,
            fontSize: colOptions?.fontSize,
            fontColor: colOptions?.fontColor,
            cellColor: colOptions?.cellColor,
            bold: colOptions?.bold,
            italic: colOptions?.italic,
            underline: colOptions?.underline,
            marginTop: colOptions?.marginTop,
            marginRight: colOptions?.marginRight,
            marginBottom: colOptions?.marginBottom,
            marginLeft: colOptions?.marginLeft,
          });
        } else {
          row.cell({
            hAlign: colOptions?.hAlign ?? "center",
            vAlign: colOptions?.vAlign ?? "center",
            fontName: colOptions?.fontName,
            fontSize: colOptions?.fontSize,
            fontColor: colOptions?.fontColor,
            cellColor: colOptions?.cellColor,
            bold: colOptions?.bold,
            italic: colOptions?.italic,
            underline: colOptions?.underline,
            marginTop: colOptions?.marginTop,
            marginRight: colOptions?.marginRight,
            marginBottom: colOptions?.marginBottom,
            marginLeft: colOptions?.marginLeft,
            ...cell,
          });
        }
      });
      this.rows.push(row);
    }
    await Promise.all(this.rows.map((row) => row.initCells(section)));
  }

  toXML(): string {
    throw new Error(
      "Table.toXML() requires a Section context—use section.table(...)",
    );
  }

  _toXMLWithSection(section: Section): string {
    const align = this.options.align ?? "center";
    let xml = "  <w:tbl>\n";
    xml += "    <w:tblPr>\n";
    if (this.options.indent) {
      const indentTwips = parseNumberTwips(this.options.indent);
      xml += `      <w:tblInd w:w="${indentTwips}" w:type="dxa"/>\n`;
    }
    if (this.options.width) {
      const width = parseNumberTwips(this.options.width);
      xml += `      <w:tblW w:w="${width}" w:type="dxa"/>\n`;
    }
    xml += `      <w:jc w:val="${align}"/>\n`;
    if (this.options.borders) {
      xml += `      <w:tblBorders>
        <w:top    w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:left   w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:right  w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
      </w:tblBorders>
`;
    }
    xml += "    </w:tblPr>\n";
    xml += "    <w:tblGrid>\n";
    for (const col of this.options.columns) {
      const w = parseNumberTwips(col.width);
      xml += `      <w:gridCol w:w="${w}"/>\n`;
    }
    xml += "    </w:tblGrid>\n";
    for (const row of this.rows) {
      xml += row.toXML(section);
    }
    xml += "  </w:tbl>\n";
    return xml;
  }
}
