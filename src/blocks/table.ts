/**
 * Table block implementation with rich text formatting, percentage-based widths,
 * cell-level styling, and shape support.
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

/**
 * Represents a line break marker within cell content.
 *
 * @private
 */
class LineBreak {
  constructor(public count: number = 1) {}
}

/**
 * Represents a page break marker within cell content.
 *
 * @private
 */
class PageBreak {
  constructor(public count: number = 1) {}
}

/**
 * Represents a shape marker within cell content.
 *
 * @private
 */
class ShapeMarker {
  constructor(
    public shapeType: ShapeType,
    public options?: ShapeOptions,
  ) {}
}

/**
 * Creates a line break marker for use in table cell runs.
 *
 * @private
 * @param {number} [count] - Number of line breaks (default: 1)
 * @returns {LineBreak} Line break marker
 */
function lineBreakMarker(count: number = 1): LineBreak {
  return new LineBreak(count);
}

/**
 * Creates a page break marker for use in table cell runs.
 *
 * @private
 * @param {number} [count] - Number of page breaks (default: 1)
 * @returns {PageBreak} Page break marker
 */
function pageBreakMarker(count: number = 1): PageBreak {
  return new PageBreak(count);
}

/**
 * Creates a shape marker for use in table cell runs.
 *
 * @private
 * @param {ShapeType} shapeType - Shape preset identifier
 * @param {ShapeOptions} [options] - Shape configuration
 * @returns {ShapeMarker} Shape marker
 */
function shapeMarker(
  shapeType: ShapeType,
  options?: ShapeOptions,
): ShapeMarker {
  return new ShapeMarker(shapeType, options);
}

/**
 * Style properties for a text run within a cell.
 *
 * @typedef {Object} TextRunStyle
 * @property {string} text - Run text content
 * @property {boolean} [bold] - Bold formatting
 * @property {boolean} [italic] - Italic formatting
 * @property {boolean} [underline] - Underline formatting
 * @property {number} [fontSize] - Font size in points
 * @property {string} [fontColor] - Font color (hex without #)
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
 * Table cell run segment — text, line break, page break, or shape.
 *
 * @typedef {Object} CellRunSegment
 */
type CellRunSegment = TextRunStyle | LineBreak | PageBreak | ShapeMarker;

/**
 * Table cell data with support for rich text runs, breaks, and shapes.
 *
 * Extends base TableCellData to provide multiple formatted text segments,
 * line and page break support, shape insertion, and per-cell font formatting
 * that overrides column defaults.
 *
 * When using `runs` for multiple formatted text segments, alignment can be
 * controlled via `hAlign` and `vAlign` properties on the cell data object,
 * or inherited from column defaults if not specified.
 *
 * @typedef {Object} EnhancedTableCellData
 * @property {string} [text] - Single text run (used if runs not provided)
 * @property {(TextRunStyle | LineBreak | PageBreak | ShapeMarker)[]} [runs] - Array of formatted text segments, line breaks, page breaks, and shapes
 * @property {string} [fontName] - Font family (overrides column default)
 * @property {number} [fontSize] - Font size in points (overrides column default)
 * @property {string} [fontColor] - Font color hex (overrides column default)
 * @property {boolean} [bold] - Bold formatting (overrides column default)
 * @property {boolean} [italic] - Italic formatting (overrides column default)
 * @property {boolean} [underline] - Underline formatting (overrides column default)
 * @property {string} [hAlign] - Horizontal alignment: "left" | "center" | "right" | "justify" (overrides column default)
 * @property {string} [vAlign] - Vertical alignment: "top" | "center" | "bottom" (overrides column default)
 * @property {string} [cellColor] - Cell background color hex
 * @property {number} [colspan] - Column span count
 * @property {number} [rowspan] - Row span count (0 for continuation)
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
 * @property {boolean} [repeat] - Repeat header on page breaks (default: false)
 */
interface TableRowOptions {
  header?: boolean;
  repeat?: boolean;
}

/**
 * Parses width specification to OOXML table cell width value.
 *
 * Handles both percentage-based widths (e.g., "28.7%") and fixed widths
 * (e.g., "5cm", "100pt"). Percentage widths use type="pct" with value * 50
 * per OOXML specification.
 *
 * @private
 * @param {string} width - Width specification
 * @returns {Object} Parsed width with value and type
 * @returns {number} return.value - Width value for OOXML
 * @returns {string} return.type - OOXML width type ("pct" or "dxa")
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
 * Represents a formatted text run within a table cell.
 *
 * Generates OOXML `<w:r>` element with optional run properties including
 * font styling, color, and typeface.
 */
class TableCellRun {
  /**
   * Creates a new text run.
   *
   * @param {string} text - Run text content
   * @param {Object} [style] - Run styling
   * @param {boolean} [style.bold] - Bold formatting
   * @param {boolean} [style.italic] - Italic formatting
   * @param {boolean} [style.underline] - Underline formatting
   * @param {number} [style.fontSize] - Font size in points
   * @param {string} [style.fontColor] - Font color (hex without #)
   * @param {string} [style.fontName] - Font family
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
 * Represents a single table row containing cells.
 *
 * Manages row-level properties and cell initialization. Calculates row height
 * based on cell contents.
 */
class TableRow {
  private cells: TableCell[] = [];
  private isHeader: boolean;
  private repeatHeader: boolean;

  /**
   * Creates a new table row.
   *
   * @param {TableOptions} tableOptions - Parent table options
   * @param {TableRowOptions} [options] - Row-level options for header and repeat
   */
  constructor(
    private tableOptions: TableOptions,
    options?: TableRowOptions,
  ) {
    this.isHeader = options?.header ?? false;
    this.repeatHeader = options?.repeat ?? false;
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
   * Performs async operations like image registration. Must be called
   * before generating XML.
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

    if (this.isHeader && this.repeatHeader) {
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
 * Represents a single table cell.
 *
 * Handles text content with multiple runs, images, shapes, cell-level styling,
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
   * Processes custom runs with line breaks, page breaks, and shapes.
   * Respects `runs` property for multiple formatted segments or falls back
   * to single text run with cell-level formatting.
   *
   * @private
   * @returns {(TableCellRun | ShapeMarker | LineBreak | PageBreak)[]} Array of formatted runs and markers
   */
  private buildRuns(): (TableCellRun | ShapeMarker | LineBreak | PageBreak)[] {
    const runs: (TableCellRun | ShapeMarker | LineBreak | PageBreak)[] = [];

    if (this.data.runs && this.data.runs.length > 0) {
      for (const segment of this.data.runs) {
        if (segment instanceof LineBreak) {
          runs.push(segment);
        } else if (segment instanceof PageBreak) {
          runs.push(segment);
        } else if (segment instanceof ShapeMarker) {
          runs.push(segment);
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
   * Creates a line break marker.
   *
   * @private
   * @param {number} [count] - Number of line breaks
   * @returns {LineBreak} Line break marker
   */
  static lineBreak(count: number = 1): LineBreak {
    return lineBreakMarker(count);
  }

  /**
   * Creates a page break marker.
   *
   * @private
   * @param {number} [count] - Number of page breaks
   * @returns {PageBreak} Page break marker
   */
  static pageBreak(count: number = 1): PageBreak {
    return pageBreakMarker(count);
  }

  /**
   * Creates a shape marker.
   *
   * @private
   * @param {ShapeType} shapeType - Shape preset identifier
   * @param {ShapeOptions} [options] - Shape configuration
   * @returns {ShapeMarker} Shape marker
   */
  static shape(
    shapeType: ShapeType,
    options?: ShapeOptions,
  ): ShapeMarker {
    return shapeMarker(shapeType, options);
  }

  /**
   * Generates OOXML paragraph for cell content.
   *
   * Produces a `<w:p>` element containing text runs, shapes, and breaks
   * with alignment and formatting properties.
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
      if (run instanceof LineBreak) {
        for (let i = 0; i < run.count; i++) {
          xml += "        <w:r><w:br/></w:r>\n";
        }
      } else if (run instanceof PageBreak) {
        for (let i = 0; i < run.count; i++) {
          xml += '        <w:r><w:br w:type="page"/></w:r>\n';
        }
      } else if (run instanceof ShapeMarker) {
        const shapeRun = createShapeRun(run.shapeType, run.options);
        xml += (shapeRun as any).text;
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
   * background color, margins, and content (text, shapes, or image).
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
 * Table block with shape support.
 *
 * Construct tables by defining columns in `TableOptions` and adding rows via
 * `.row(...)`. Use `section.table(options)` to obtain a `Table` tied to a
 * section — Table requires a Section context to resolve images and sizing.
 *
 * Supports percentage-based column widths, multiple text runs per cell with
 * independent formatting, line breaks, page breaks, shapes, and cell-level
 * style properties that override column defaults.
 */
export class Table extends Element {
  private rowDefs: Array<{
    cells: (string | EnhancedTableCellData)[];
    options?: TableRowOptions;
  }> = [];
  private rows: TableRow[] = [];
  private options: TableOptions;
  private isBuilt = false;

  /**
   * Creates a new table.
   *
   * @param {TableOptions} options - Table configuration
   * @param {TableColumn[]} options.columns - Column definitions with width
   * @param {string} [options.width] - Table width (percentage or fixed)
   * @param {string} [options.align] - Table alignment (left, center, right, justify)
   * @param {boolean} [options.borders] - Show table borders (default: true)
   * @param {string} [options.indent] - Table indentation
   * @param {string} [options.marginTop] - Top margin for all cells
   * @param {string} [options.marginRight] - Right margin for all cells
   * @param {string} [options.marginBottom] - Bottom margin for all cells
   * @param {string} [options.marginLeft] - Left margin for all cells
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
   * Rows can contain simple string cells or `EnhancedTableCellData` objects
   * with rich formatting, multiple runs including shapes, and line/page breaks.
   *
   * First parameter can be row options object with header and repeat flags,
   * followed by cell content. If options object is omitted, all arguments
   * are treated as cell content.
   *
   * @param {...(string | EnhancedTableCellData | TableRowOptions)[]} args - Row options object optionally followed by cell data
   * @returns {this}
   */
  row(...args: (string | EnhancedTableCellData | TableRowOptions)[]): this {
    let options: TableRowOptions | undefined;
    let cells: (string | EnhancedTableCellData)[] = [];

    if (
      args.length > 0 &&
      typeof args[0] === "object" &&
      !Array.isArray(args[0]) &&
      (("header" in args[0]) || ("repeat" in args[0]))
    ) {
      options = args[0] as TableRowOptions;
      cells = args.slice(1) as (string | EnhancedTableCellData)[];
    } else {
      cells = args as (string | EnhancedTableCellData)[];
    }

    this.rowDefs.push({ cells, options });
    return this;
  }

  /**
   * Applies conditional operations to the table.
   *
   * @deprecated Use direct `.row()` calls instead.
   * @param {...Function[]} ops - Operations to apply
   * @returns {this}
   */
  apply(...ops: ((builder: this) => this)[]): this {
    console.warn(
      "Table.apply() is deprecated. Use direct .row() calls instead.",
    );
    for (const op of ops) {
      op(this);
    }
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
            hAlign: cell.hAlign ?? colOptions?.hAlign ?? "center",
            vAlign: cell.vAlign ?? colOptions?.vAlign ?? "center",
            fontName: cell.fontName ?? colOptions?.fontName,
            fontSize: cell.fontSize ?? colOptions?.fontSize,
            fontColor: cell.fontColor ?? colOptions?.fontColor,
            cellColor: cell.cellColor ?? colOptions?.cellColor,
            bold: cell.bold ?? colOptions?.bold,
            italic: cell.italic ?? colOptions?.italic,
            underline: cell.underline ?? colOptions?.underline,
            marginTop: cell.marginTop ?? colOptions?.marginTop,
            marginRight: cell.marginRight ?? colOptions?.marginRight,
            marginBottom: cell.marginBottom ?? colOptions?.marginBottom,
            marginLeft: cell.marginLeft ?? colOptions?.marginLeft,
            ...cell,
          });
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
