/**
 * Table block implementation with rich text formatting, percentage-based widths,
 * and cell-level styling.
 *
 * Construct tables by defining columns in `TableOptions` and adding rows via
 * `.row(...)`. Use `section.table(options)` to obtain a `Table` tied to a
 * section — Table requires a Section context to resolve images and sizing.
 *
 * @module
 */

import {
  cmToEmu,
  escapeXML,
  parseImageSize,
  parseNumberTwips,
  ptToHalfPoints,
} from "../core/utils.ts";
import type {
  TableCellData,
  TableColumn,
  TableOptions,
} from "../core/docxaur.ts";
import { Element } from "./element.ts";
import type { Section } from "./section.ts";
import { Image } from "./image.ts";
import { buildShapeXML, type ShapeOptions, type ShapeType } from "./shapes.ts";

/**
 * Row segment — text cell, line break, page break, or shape.
 *
 * @typedef {Object} RowSegment
 */
type RowSegment =
  | { lineBreak?: number }
  | { pageBreak?: number }
  | ({ shape: string } & ShapeOptions)
  | EnhancedTableCellData;

/**
 * Text run style properties.
 *
 * @typedef {Object} TextRunStyle
 * @property {string} text - Run text content
 * @property {boolean} [bold] - Bold formatting
 * @property {boolean} [italic] - Italic formatting
 * @property {boolean} [underline] - Underline formatting
 * @property {number} [fontSize] - Font size in points
 * @property {string} [fontColor] - Font color hex
 * @property {string} [fontName] - Font family name
 */
interface TextRunStyle {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number;
  fontColor?: string;
  fontName?: string;
}

/**
 * Table cell run segment — text, line break, or page break.
 *
 * @typedef {Object} CellRunSegment
 */
type CellRunSegment = TextRunStyle | { lineBreak: number } | {
  pageBreak: number;
};

/**
 * Enhanced table cell data with rich formatting.
 *
 * Extends base TableCellData to provide multiple formatted text segments,
 * line and page break support, and per-cell font formatting that overrides
 * column defaults.
 *
 * @typedef {Object} EnhancedTableCellData
 * @property {string} [text] - Single text run
 * @property {CellRunSegment[]} [runs] - Array of formatted text segments, line breaks, and page breaks
 * @property {string} [fontName] - Font family
 * @property {number} [fontSize] - Font size in points
 * @property {string} [fontColor] - Font color hex
 * @property {boolean} [bold] - Bold formatting
 * @property {boolean} [italic] - Italic formatting
 * @property {boolean} [underline] - Underline formatting
 * @property {string} [hAlign] - Horizontal alignment
 * @property {string} [vAlign] - Vertical alignment
 * @property {string} [cellColor] - Cell background color hex
 * @property {number} [colspan] - Column span count
 * @property {number} [rowspan] - Row span count
 * @property {string} [height] - Cell height
 * @property {string} [marginTop] - Top margin
 * @property {string} [marginRight] - Right margin
 * @property {string} [marginBottom] - Bottom margin
 * @property {string} [marginLeft] - Left margin
 */
export interface EnhancedTableCellData extends TableCellData {
  runs?: CellRunSegment[];
}

/**
 * Row-level options for header and repeat settings.
 *
 * @typedef {Object} TableRowOptions
 * @property {boolean} [header] - Mark row as table header
 * @property {boolean} [repeat] - Repeat header on page breaks
 */
interface TableRowOptions {
  header?: boolean;
  repeat?: boolean;
}

/**
 * Parses width specification to OOXML table cell width value.
 *
 * Handles both percentage-based widths and fixed widths.
 * Percentage widths use type="pct" with value * 50 per OOXML specification.
 *
 * @private
 * @param {string} width - Width specification
 * @returns {Object} Parsed width with value and type
 */
function parseTableCellWidth(width: string): { value: number; type: string } {
  const isPercentage = width.endsWith("%");

  if (isPercentage) {
    const percentValue = parseInt(width, 10);
    return { value: percentValue * 50, type: "pct" };
  }

  return { value: parseNumberTwips(width), type: "dxa" };
}

/**
 * Formatted text run within a table cell.
 *
 * Generates OOXML `<w:r>` element with optional run properties including
 * font styling, color, and typeface.
 */
class TableCellRun {
  isShape: boolean = false;

  /**
   * Creates a new text run.
   *
   * @param {string} text - Run text content
   * @param {Object} [style] - Run styling
   */
  constructor(
    public text: string,
    public style?: {
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      fontSize?: number;
      fontColor?: string;
      fontName?: string;
    },
  ) {}

  /**
   * Generates OOXML for this text run.
   *
   * Produces a `<w:r>` element with embedded `<w:rPr>` properties if styling
   * is specified, followed by the text content in `<w:t>`.
   *
   * @returns {string} WordprocessingML run element
   */
  toXML(): string {
    if (this.isShape) {
      return this.text;
    }

    let xml = "        <w:r>\n";

    if (this.style && Object.values(this.style).some((v) => v !== undefined)) {
      xml += "          <w:rPr>\n";
      if (this.style.bold) xml += "            <w:b/>\n";
      if (this.style.italic) xml += "            <w:i/>\n";
      if (this.style.underline) xml += '            <w:u w:val="single"/>\n';
      if (this.style.fontSize) {
        xml += `            <w:sz w:val="${
          ptToHalfPoints(this.style.fontSize)
        }"/>\n`;
      }
      if (this.style.fontColor) {
        xml += `            <w:color w:val="${this.style.fontColor}"/>\n`;
      }
      if (this.style.fontName) {
        xml +=
          `            <w:rFonts w:ascii="${this.style.fontName}" w:hAnsi="${this.style.fontName}"/>\n`;
      }
      xml += "          </w:rPr>\n";
    }

    xml += `          <w:t xml:space="preserve">${
      escapeXML(this.text)
    }</w:t>\n`;
    xml += "        </w:r>\n";
    return xml;
  }
}

/**
 * Single table row containing cells.
 *
 * Manages row-level properties and cell initialization. Calculates row height
 * based on cell contents.
 */
class TableRow {
  private cells: TableCell[] = [];
  private header: boolean;
  private repeat: boolean;

  /**
   * Creates a new table row.
   *
   * @param {TableOptions} tableOptions - Parent table options
   * @param {TableRowOptions} [options] - Row-level options
   */
  constructor(
    private tableOptions: TableOptions,
    options?: TableRowOptions,
  ) {
    this.header = options?.header ?? false;
    this.repeat = options?.repeat ?? false;
  }

  /**
   * Adds a cell to this row.
   *
   * @param {EnhancedTableCellData} data - Cell content and styling
   * @returns {this}
   */
  cell(data: EnhancedTableCellData): this {
    this.cells.push(new TableCell(data));
    return this;
  }

  /**
   * Initializes all cells.
   *
   * Performs async operations like image registration.
   * Must be called before generating XML.
   *
   * @param {Section} section - Parent section context
   * @returns {Promise<void>}
   */
  async initCells(section: Section): Promise<void> {
    await Promise.all(this.cells.map((cell) => cell.init(section)));
  }

  /**
   * Generates OOXML for this table row.
   *
   * Produces a `<w:tr>` element with row properties and cell elements.
   * Includes `<w:tblHeader/>` when row is marked as header with repeat enabled.
   *
   * @param {Section} section - Parent section context
   * @returns {string} WordprocessingML table row element
   */
  toXML(section: Section): string {
    let xml = "  <w:tr>\n";
    xml += "    <w:trPr>\n";

    if (this.header && this.repeat) {
      xml += "      <w:tblHeader/>\n";
    }

    let maxHeight = 170;
    for (const cell of this.cells) {
      const h = cell.getHeight();
      if (h > maxHeight) maxHeight = h;
    }
    xml += `      <w:trHeight w:val="${maxHeight}" w:hRule="atLeast"/>\n`;
    xml += "    </w:trPr>\n";
    for (let i = 0; i < this.cells.length; i++) {
      xml += this.cells[i].toXML(i, this.tableOptions, section);
    }
    xml += "  </w:tr>\n";
    return xml;
  }
}

/**
 * Single table cell.
 *
 * Handles text content with multiple runs, images, cell-level styling,
 * merging properties, and margins. Supports percentage and fixed-width columns.
 */
class TableCell {
  private imageId?: number;

  /**
   * Creates a new table cell.
   *
   * @param {EnhancedTableCellData} data - Cell content and formatting
   */
  constructor(private data: EnhancedTableCellData) {}

  /**
   * Initializes the cell.
   *
   * Registers embedded images and prepares content for rendering.
   *
   * @param {Section} section - Parent section context
   * @returns {Promise<void>}
   */
  async init(section: Section): Promise<void> {
    if (this.data.image) {
      this.imageId = await section._doc().registerImage(this.data.image.url);
    }
  }

  /**
   * Gets cell height in twips.
   *
   * Returns explicit cell height or default minimum height.
   *
   * @returns {number} Height in twips
   */
  getHeight(): number {
    if (this.data.height) return parseNumberTwips(this.data.height);
    return 170;
  }

  /**
   * Builds array of text runs from cell data.
   *
   * Processes custom runs with text, line breaks, page breaks, and shapes.
   * Respects `runs` property for multiple formatted segments or falls back to
   * single text run with cell-level formatting.
   *
   * @private
   * @returns {TableCellRun[]} Array of formatted text runs
   */
  private buildRuns(): TableCellRun[] {
    const runs: TableCellRun[] = [];

    if (this.data.runs && this.data.runs.length > 0) {
      for (const segment of this.data.runs) {
        if ("lineBreak" in segment) {
          for (let i = 0; i < segment.lineBreak; i++) {
            runs.push(new TableCellRun("\n"));
          }
        } else if ("pageBreak" in segment) {
          for (let i = 0; i < segment.pageBreak; i++) {
            runs.push(new TableCellRun("[PAGE_BREAK]"));
          }
        } else if ("shape" in segment) {
          const shapeSegment = segment as { shape: string } & ShapeOptions;
          const shapeTypeMap: Record<string, ShapeType> = {
            rect: { preset: "rect", name: "Rectangle" },
            roundRect: { preset: "roundRect", name: "Rounded Rectangle" },
            ellipse: { preset: "ellipse", name: "Circle" },
            diamond: { preset: "diamond", name: "Diamond" },
            triangle: { preset: "triangle", name: "Triangle" },
            pentagon: { preset: "pentagon", name: "Pentagon" },
            hexagon: { preset: "hexagon", name: "Hexagon" },
            star5: { preset: "star5", name: "Star (5-point)" },
            heart: { preset: "heart", name: "Heart" },
          };

          const shapeType = shapeTypeMap[shapeSegment.shape] ??
            shapeTypeMap.rect;
          const shapeXML = buildShapeXML(shapeType, shapeSegment);
          runs.push(new TableCellRun(shapeXML, undefined));
          runs[runs.length - 1].isShape = true;
        } else {
          runs.push(
            new TableCellRun((segment as TextRunStyle).text, {
              bold: (segment as TextRunStyle).bold,
              italic: (segment as TextRunStyle).italic,
              underline: (segment as TextRunStyle).underline,
              fontSize: (segment as TextRunStyle).fontSize,
              fontColor: (segment as TextRunStyle).fontColor,
              fontName: (segment as TextRunStyle).fontName,
            }),
          );
        }
      }
    } else if (this.data.text) {
      runs.push(
        new TableCellRun(this.data.text, {
          bold: this.data.bold,
          italic: this.data.italic,
          underline: this.data.underline,
          fontSize: this.data.fontSize,
          fontColor: this.data.fontColor,
          fontName: this.data.fontName,
        }),
      );
    }

    return runs;
  }

  /**
   * Generates OOXML paragraph for cell text content.
   *
   * Produces a `<w:p>` element containing text runs with alignment
   * and formatting properties. Handles line breaks and page breaks.
   *
   * @private
   * @returns {string} WordprocessingML paragraph element
   */
  private buildParagraphXML(): string {
    const jc = this.data.hAlign === "justify"
      ? "both"
      : this.data.hAlign ?? "center";

    let xml = "      <w:p>\n";
    xml += "        <w:pPr>\n";
    xml += `          <w:jc w:val="${jc}"/>\n`;
    xml += "        </w:pPr>\n";

    const runs = this.buildRuns();
    for (const run of runs) {
      if (run.text === "\n") {
        xml += "        <w:r><w:br/></w:r>\n";
      } else if (run.text === "[PAGE_BREAK]") {
        xml += '        <w:r><w:br w:type="page"/></w:r>\n';
      } else {
        xml += run.toXML();
      }
    }

    xml += "      </w:p>\n";
    return xml;
  }

  /**
   * Generates OOXML for this table cell.
   *
   * Produces a `<w:tc>` element with cell properties (width, alignment, merging),
   * background color, margins, and content (text or image).
   *
   * Supports percentage-based widths and fixed widths. Cell font properties
   * override column defaults.
   *
   * @param {number} colIndex - Column index for width lookup
   * @param {TableOptions} tableOptions - Parent table options
   * @param {Section} section - Parent section context
   * @returns {string} WordprocessingML table cell element
   */
  toXML(
    colIndex: number,
    tableOptions: TableOptions,
    section: Section,
  ): string {
    const vAlign = this.data.vAlign ?? "center";
    const align = this.data.hAlign ?? "center";
    let xml = "    <w:tc>\n";
    xml += "      <w:tcPr>\n";

    const colWidth = tableOptions.columns[colIndex].width;
    const widthSpec = parseTableCellWidth(colWidth);

    xml +=
      `        <w:tcW w:w="${widthSpec.value}" w:type="${widthSpec.type}"/>\n`;

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
      xml += this.buildParagraphXML();
    }
    xml += "    </w:tc>\n";
    return xml;
  }
}

/**
 * Table block.
 *
 * Construct tables by defining columns in `TableOptions` and adding rows via
 * `.row(...)`. Use `section.table(options)` to obtain a `Table` tied to a
 * section — Table requires a Section context to resolve images and sizing.
 *
 * Supports percentage-based column widths, multiple text runs per cell with
 * independent formatting, line breaks, page breaks, and cell-level style
 * properties that override column defaults.
 */
export class Table extends Element {
  private rowDefs: Array<{
    cells: RowSegment[];
    options?: TableRowOptions;
  }> = [];
  private rows: TableRow[] = [];
  private options: TableOptions;
  private isBuilt = false;

  /**
   * Creates a new table.
   *
   * @param {TableOptions} options - Table configuration
   */
  constructor(options: TableOptions) {
    super();
    this.options = options;
    if (this.options.borders === undefined) {
      this.options.borders = true;
    }
  }

  /**
   * Adds a row to the table.
   *
   * Cell content can be simple strings or `EnhancedTableCellData` objects
   * with rich formatting. Special segments support line breaks, page breaks,
   * and shapes:
   *
   * - `{ lineBreak: number }` — insert line break(s)
   * - `{ pageBreak: number }` — insert page break(s)
   * - `{ shape: string; width?: string; height?: string; text?: string; ... }` — insert shape with sizing and text
   *
   * First parameter can be row options object with header and repeat flags,
   * followed by cell content. If options object is omitted, all arguments
   * are treated as cell content.
   *
   * @param {...(string | EnhancedTableCellData | TableRowOptions | RowSegment)[]} args - Row options object optionally followed by cell data
   * @returns {this}
   */
  row(
    ...args: (string | EnhancedTableCellData | TableRowOptions | RowSegment)[]
  ): this {
    let options: TableRowOptions | undefined;
    let cells: RowSegment[] = [];

    if (
      args.length > 0 &&
      typeof args[0] === "object" &&
      !Array.isArray(args[0]) &&
      (("header" in args[0]) || ("repeat" in args[0]))
    ) {
      options = args[0] as TableRowOptions;
      cells = args.slice(1) as RowSegment[];
    } else {
      cells = args as RowSegment[];
    }

    this.rowDefs.push({ cells, options });
    return this;
  }

  /**
   * Builds all rows.
   *
   * Initializes cells, registers images, and prepares content for rendering.
   * Called internally before generating OOXML.
   *
   * @param {Section} section - Parent section context
   * @returns {Promise<void>}
   */
  async buildRows(section: Section): Promise<void> {
    if (this.isBuilt) return;
    this.isBuilt = true;
    for (const rowDef of this.rowDefs) {
      const row = new TableRow(this.options, rowDef.options);
      rowDef.cells.forEach((cell, i) => {
        if ("lineBreak" in cell) {
          row.cell({
            runs: [{ lineBreak: cell.lineBreak ?? 1 }],
          });
        } else if ("pageBreak" in cell) {
          row.cell({
            runs: [{ pageBreak: cell.pageBreak ?? 1 }],
          });
        } else if ("shape" in cell) {
          const shapeCell = cell as { shape: string } & ShapeOptions;
          const shapeStr = shapeCell.shape;
          const colOptions = this.options.columns[i];

          const shapeTypeMap: Record<string, ShapeType> = {
            rect: { preset: "rect", name: "Rectangle" },
            roundRect: { preset: "roundRect", name: "Rounded Rectangle" },
            ellipse: { preset: "ellipse", name: "Circle" },
            diamond: { preset: "diamond", name: "Diamond" },
            triangle: { preset: "triangle", name: "Triangle" },
            pentagon: { preset: "pentagon", name: "Pentagon" },
            hexagon: { preset: "hexagon", name: "Hexagon" },
            star5: { preset: "star5", name: "Star (5-point)" },
            heart: { preset: "heart", name: "Heart" },
          };

          const shapeType = shapeTypeMap[shapeStr] ?? shapeTypeMap.rect;
          const shapeXML = buildShapeXML(shapeType, shapeCell);

          row.cell({
            text: "",
            hAlign: shapeCell.align ?? colOptions?.hAlign ?? "center",
            vAlign: colOptions?.vAlign ?? "center",
          });
        } else {
          const col = this.options.columns[i];
          if (typeof cell === "string") {
            row.cell({
              text: cell,
              hAlign: col?.hAlign ?? "center",
              vAlign: col?.vAlign ?? "center",
              fontName: col?.fontName,
              fontSize: col?.fontSize,
              fontColor: col?.fontColor,
              cellColor: col?.cellColor,
              bold: col?.bold,
              italic: col?.italic,
              underline: col?.underline,
              marginTop: col?.marginTop,
              marginRight: col?.marginRight,
              marginBottom: col?.marginBottom,
              marginLeft: col?.marginLeft,
            });
          } else {
            const cellData = cell as EnhancedTableCellData;
            row.cell({
              hAlign: cellData.hAlign ?? col?.hAlign ?? "center",
              vAlign: cellData.vAlign ?? col?.vAlign ?? "center",
              fontName: cellData.fontName ?? col?.fontName,
              fontSize: cellData.fontSize ?? col?.fontSize,
              fontColor: cellData.fontColor ?? col?.fontColor,
              cellColor: cellData.cellColor ?? col?.cellColor,
              bold: cellData.bold ?? col?.bold,
              italic: cellData.italic ?? col?.italic,
              underline: cellData.underline ?? col?.underline,
              marginTop: cellData.marginTop ?? col?.marginTop,
              marginRight: cellData.marginRight ?? col?.marginRight,
              marginBottom: cellData.marginBottom ?? col?.marginBottom,
              marginLeft: cellData.marginLeft ?? col?.marginLeft,
              ...cellData,
            });
          }
        }
      });
      this.rows.push(row);
    }
    await Promise.all(this.rows.map((row) => row.initCells(section)));
  }

  /**
   * Generates OOXML for this table.
   *
   * @throws {Error} Direct call not supported. Use via section.table() context.
   * @returns {string}
   */
  toXML(): string {
    throw new Error(
      "Table.toXML() requires a Section context—use section.table(...)",
    );
  }

  /**
   * Generates complete OOXML table element with section context.
   *
   * Produces a `<w:tbl>` element with table properties, grid definition,
   * and rows. Supports both percentage and fixed-width columns.
   *
   * @param {Section} section - Parent section context
   * @returns {string} WordprocessingML table element
   */
  _toXMLWithSection(section: Section): string {
    const align = this.options.align ?? "center";
    let xml = "  <w:tbl>\n";
    xml += "    <w:tblPr>\n";

    if (this.options.width) {
      const widthSpec = parseTableCellWidth(this.options.width);
      xml +=
        `      <w:tblW w:w="${widthSpec.value}" w:type="${widthSpec.type}"/>\n`;
    }

    if (this.options.indent) {
      const indentTwips = parseNumberTwips(this.options.indent);
      xml += `      <w:tblInd w:w="${indentTwips}" w:type="dxa"/>\n`;
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
      const widthSpec = parseTableCellWidth(col.width);
      if (widthSpec.type === "dxa") {
        xml += `      <w:gridCol w:w="${widthSpec.value}"/>\n`;
      } else {
        xml += `      <w:gridCol w:w="1440"/>\n`;
      }
    }
    xml += "    </w:tblGrid>\n";
    for (const row of this.rows) {
      xml += row.toXML(section);
    }
    xml += "  </w:tbl>\n";
    return xml;
  }
}
