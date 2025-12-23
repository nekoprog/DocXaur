/**
 * Shape block implementation with DrawingML support.
 *
 * Provides geometric shape presets and configuration for inline insertion
 * into paragraphs, sections, and tables.
 *
 * @module
 */

import { cmToEmu, escapeXML as escRunText } from "../core/utils.ts";

/**
 * Shape type identifier with preset name.
 *
 * @typedef {Object} ShapeType
 * @property {string} preset - OOXML preset shape name
 * @property {string} name - Human-readable shape name
 */
export interface ShapeType {
  preset: string;
  name: string;
}

/**
 * Rectangle shape preset.
 *
 * @type {ShapeType}
 */
export const SHAPE_RECT: ShapeType = {
  preset: "rect",
  name: "Rectangle",
};

/**
 * Rounded rectangle shape preset.
 *
 * @type {ShapeType}
 */
export const SHAPE_ROUNDED_RECT: ShapeType = {
  preset: "roundRect",
  name: "Rounded Rectangle",
};

/**
 * Circle or ellipse shape preset.
 *
 * @type {ShapeType}
 */
export const SHAPE_CIRCLE: ShapeType = {
  preset: "ellipse",
  name: "Circle",
};

/**
 * Diamond shape preset.
 *
 * @type {ShapeType}
 */
export const SHAPE_DIAMOND: ShapeType = {
  preset: "diamond",
  name: "Diamond",
};

/**
 * Triangle shape preset.
 *
 * @type {ShapeType}
 */
export const SHAPE_TRIANGLE: ShapeType = {
  preset: "triangle",
  name: "Triangle",
};

/**
 * Pentagon shape preset.
 *
 * @type {ShapeType}
 */
export const SHAPE_PENTAGON: ShapeType = {
  preset: "pentagon",
  name: "Pentagon",
};

/**
 * Hexagon shape preset.
 *
 * @type {ShapeType}
 */
export const SHAPE_HEXAGON: ShapeType = {
  preset: "hexagon",
  name: "Hexagon",
};

/**
 * Five-point star shape preset.
 *
 * @type {ShapeType}
 */
export const SHAPE_STAR_5: ShapeType = {
  preset: "star5",
  name: "Star (5-point)",
};

/**
 * Heart shape preset.
 *
 * @type {ShapeType}
 */
export const SHAPE_HEART: ShapeType = {
  preset: "heart",
  name: "Heart",
};

/**
 * Shape gradient stop configuration.
 *
 * @typedef {Object} GradientStop
 * @property {number} position - Position in gradient (0-100000)
 * @property {string} color - Stop color hex
 */
export interface GradientStop {
  position: number;
  color: string;
}

/**
 * Shape fill configuration.
 *
 * @typedef {Object} ShapeFill
 * @property {string} [color] - Solid fill color hex or "none"
 * @property {GradientStop[]} [gradient] - Gradient fill stops
 * @property {string} [gradientAngle] - Gradient angle in degrees
 */
export interface ShapeFill {
  color?: string;
  gradient?: GradientStop[];
  gradientAngle?: string;
}

/**
 * Shape line/border configuration.
 *
 * @typedef {Object} ShapeLine
 * @property {string} [color] - Line color hex or "none"
 * @property {number} [width] - Line width in pt
 * @property {string} [dash] - Dash style
 */
export interface ShapeLine {
  color?: string;
  width?: number;
  dash?: string;
}

/**
 * Shape size configuration.
 *
 * @typedef {Object} ShapeSize
 * @property {string} [width] - Width dimension
 * @property {string} [height] - Height dimension
 * @property {number} [rotate] - Shape rotation in degrees
 */
export interface ShapeSize {
  width?: string;
  height?: string;
  rotate?: number;
}

/**
 * Text box configuration for shape.
 *
 * @typedef {Object} ShapeTextBox
 * @property {string} [text] - Text content
 * @property {boolean} [bold] - Bold formatting
 * @property {boolean} [italic] - Italic formatting
 * @property {boolean} [underline] - Underline formatting
 * @property {number} [fontSize] - Font size in points
 * @property {string} [fontColor] - Text color hex
 * @property {string} [fontName] - Font family
 * @property {string} [hAlign] - Text alignment
 * @property {string} [marginTop] - Internal Top Margin
 * @property {string} [marginBottom] - Internal Bottom Margin
 * @property {string} [marginLeft] - Internal Left Margin
 * @property {string} [marginRight] - Internal Right Margin
 */
export interface ShapeTextBox {
  text?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number;
  fontColor?: string;
  fontName?: string;
  hAlign?: string;
  marginTop?: string;
  marginBottom?: string;
  marginLeft?: string;
  marginRight?: string;
}

/**
 * Shape options for insertion.
 *
 * @typedef {Object} ShapeOptions
 * @property {ShapeSize} [size] - Shape dimensions
 * @property {ShapeFill} [fill] - Fill properties
 * @property {ShapeLine} [line] - Border/line properties
 * @property {ShapeTextBox} [textBox] - Text box content and styling
 * @property {string} [hAlign] - Horizontal alignment
 * @property {"anchor" | "inline"} [position] - Shape positioning mode
 */
export interface ShapeOptions {
  size?: ShapeSize;
  fill?: ShapeFill;
  line?: ShapeLine;
  textBox?: ShapeTextBox;
  hAlign?: "left" | "center" | "right";
  position?: "anchor" | "inline";
}

let shapeCounter = 1;

/**
 * Parses dimension string to EMU units.
 *
 * @private
 * @param {string} dim - Dimension string
 * @returns {number} Dimension in EMU
 */
function parseShapeDim(dim: string): number {
  const match = dim.match(/^([\d.]+)(cm|pt|mm|in|px)$/);
  if (!match) return cmToEmu(2);

  const value = parseFloat(match[1]);
  const unit = match[2];

  switch (unit) {
    case "cm":
      return cmToEmu(value);
    case "mm":
      return cmToEmu(value / 10);
    case "in":
      return Math.round(value * 914400);
    case "pt":
      return Math.round(value * 12700);
    case "px":
      return Math.round(value * 9525);
    default:
      return cmToEmu(2);
  }
}

/**
 * Generates OOXML for shape fill properties.
 *
 * @private
 * @param {ShapeFill} [fill] - Fill configuration
 * @returns {string} OOXML solidFill or gradient element
 */
function buildShapeFillXML(fill?: ShapeFill): string {
  if (!fill) {
    return '        <a:solidFill><a:srgbClr val="000000"/></a:solidFill>\n';
  }

  if (fill.color === "none") {
    return "        <a:noFill/>\n";
  }

  if (fill.gradient && fill.gradient.length > 0) {
    let xml = "        <a:gradFill>\n";
    xml += "          <a:gsLst>\n";
    for (const stop of fill.gradient) {
      xml +=
        `            <a:gs pos="${stop.position}"><a:srgbClr val="${stop.color}"/></a:gs>\n`;
    }
    xml += "          </a:gsLst>\n";
    xml += `          <a:lin ang="${
      parseInt(fill.gradientAngle || "0") * 60000
    }" scaled="1"/>\n`;
    xml += "        </a:gradFill>\n";
    return xml;
  }

  const color = fill.color || "000000";
  return `        <a:solidFill><a:srgbClr val="${color}"/></a:solidFill>\n`;
}

/**
 * Generates OOXML for shape line properties.
 *
 * @private
 * @param {ShapeLine} [line] - Line configuration
 * @returns {string} OOXML ln element
 */
function buildShapeLineXML(line?: ShapeLine): string {
  if (!line || (!line.color && !line.width)) {
    return '        <a:ln w="19050"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>\n';
  }

  if (line.color === "none") {
    return '        <a:ln w="19050"><a:noFill/><a:prstDash val="solid"/><a:round/></a:ln>\n';
  }

  const color = line.color || "000000";
  const width = Math.round((line.width ?? 1) * 12700);
  const dash = line.dash || "solid";

  return `        <a:ln w="${width}"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:prstDash val="${dash}"/><a:round/></a:ln>\n`;
}

/**
 * Generates OOXML for text box body content.
 *
 * @private
 * @param {ShapeTextBox} textBox - Text box configuration
 * @returns {string} OOXML text body element
 */
function buildTextBoxXML(textBox: ShapeTextBox): string {
  if (!textBox.text) return "";

  const hAlign = textBox.hAlign || "left";
  const jc = hAlign === "justify" ? "both" : hAlign;

  let xml = "              <wps:txbx>\n";
  xml += "                <w:txbxContent>\n";
  xml += "                  <w:p>\n";
  xml += "                    <w:pPr>\n";
  if (hAlign !== "left") {
    xml += `                      <w:jc w:val="${jc}"/>\n`;
  }
  xml += "                    </w:pPr>\n";
  xml += "                    <w:r>\n";
  xml += "                      <w:rPr>\n";
  if (textBox.bold) xml += "                        <w:b/>\n";
  if (textBox.italic) xml += "                        <w:i/>\n";
  if (textBox.underline) {
    xml += '                        <w:u w:val="single"/>\n';
  }
  if (textBox.fontSize) {
    xml += `                        <w:sz w:val="${textBox.fontSize * 2}"/>\n`;
  }
  if (textBox.fontColor) {
    xml += `                        <w:color w:val="${textBox.fontColor}"/>\n`;
  }
  if (textBox.fontName) {
    xml +=
      `                        <w:rFonts w:ascii="${textBox.fontName}" w:hAnsi="${textBox.fontName}"/>\n`;
  }
  xml += "                      </w:rPr>\n";
  xml += `                      <w:t>${escRunText(textBox.text)}</w:t>\n`;
  xml += "                    </w:r>\n";
  xml += "                  </w:p>\n";
  xml += "                </w:txbxContent>\n";
  xml += "              </wps:txbx>\n";
  return xml;
}

/**
 * Generates OOXML DrawingML shape element.
 *
 * Produces a complete shape element with fill, line, and optional text box.
 * Supports both anchor-positioned and inline shapes.
 *
 * @param {ShapeType} shapeType - Shape preset
 * @param {ShapeOptions} [options] - Shape styling and positioning
 * @returns {string} OOXML shape element
 */
export function buildShapeXML(
  shapeType: ShapeType,
  options?: ShapeOptions,
): string {
  const width = options?.size?.width
    ? parseShapeDim(options.size.width)
    : cmToEmu(2);
  const height = options?.size?.height
    ? parseShapeDim(options.size.height)
    : cmToEmu(2);

  const rot = options?.size?.rotate !== undefined
    ? Math.round(options.size.rotate * 60000)
    : undefined;

  const fillXML = buildShapeFillXML(options?.fill);
  const lineXML = buildShapeLineXML(options?.line);
  const textBoxXML = options?.textBox ? buildTextBoxXML(options.textBox) : "";

  const shapeId = shapeCounter++;
  const hAlign = options?.hAlign || "center";
  const position = options?.position || "anchor";

  const lIns = options?.textBox?.marginLeft
    ? parseShapeDim(options.textBox.marginLeft)
    : undefined;
  const rIns = options?.textBox?.marginRight
    ? parseShapeDim(options.textBox.marginRight)
    : undefined;
  const tIns = options?.textBox?.marginTop
    ? parseShapeDim(options.textBox.marginTop)
    : undefined;
  const bIns = options?.textBox?.marginBottom
    ? parseShapeDim(options.textBox.marginBottom)
    : undefined;

  const bodyPrAttrs = [
    'rot="0"',
    'vert="horz"',
    'anchor="ctr"',
    'anchorCtr="0"',
    'rtlCol="0"',
    lIns !== undefined ? `lIns="${lIns}"` : null,
    tIns !== undefined ? `tIns="${tIns}"` : null,
    rIns !== undefined ? `rIns="${rIns}"` : null,
    bIns !== undefined ? `bIns="${bIns}"` : null,
  ]
    .filter(Boolean)
    .join(" ");

  const xfrmOpen = rot !== undefined ? `<a:xfrm rot="${rot}">` : `<a:xfrm>`;

  if (position === "inline") {
    return `    <w:r>
       <w:drawing>
         <wp:inline distT="0" distB="0" distL="0" distR="0">
           <wp:extent cx="${width}" cy="${height}"/>
           <wp:effectExtent l="0" t="0" r="0" b="0"/>
           <wp:docPr id="${shapeId}" name="Shape ${shapeId}"/>
           <wp:cNvGraphicFramePr/>
           <a:graphic>
             <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
               <wps:wsp>
                 <wps:cNvSpPr txBox="1"/>
                 <wps:spPr>
                   ${xfrmOpen}<a:off x="0" y="0"/><a:ext cx="${width}" cy="${height}"/></a:xfrm>
                   <a:prstGeom prst="${shapeType.preset}"><a:avLst/></a:prstGeom>
 ${fillXML}${lineXML}                </wps:spPr>
 ${textBoxXML}                <wps:bodyPr ${bodyPrAttrs}/>
               </wps:wsp>
             </a:graphicData>
           </a:graphic>
         </wp:inline>
       </w:drawing>
     </w:r>
     `;
  }

  return `    <w:r>
     <w:drawing>
       <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
         <wp:simplePos x="0" y="0"/>
         <wp:positionH relativeFrom="column"><wp:align>${hAlign}</wp:align></wp:positionH>
         <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
         <wp:extent cx="${width}" cy="${height}"/>
         <wp:effectExtent l="0" t="0" r="0" b="0"/>
         <wp:wrapNone/>
         <wp:docPr id="${shapeId}" name="Shape ${shapeId}"/>
         <wp:cNvGraphicFramePr/>
         <a:graphic>
           <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
             <wps:wsp>
               <wps:cNvSpPr txBox="1"/>
               <wps:spPr>
                 ${xfrmOpen}<a:off x="0" y="0"/><a:ext cx="${width}" cy="${height}"/></a:xfrm>
                 <a:prstGeom prst="${shapeType.preset}"><a:avLst/></a:prstGeom>
 ${fillXML}${lineXML}              </wps:spPr>
 ${textBoxXML}              <wps:bodyPr ${bodyPrAttrs}/>
             </wps:wsp>
           </a:graphicData>
         </a:graphic>
       </wp:anchor>
     </w:drawing>
   </w:r>
   `;
}

/**
 * Creates shape insertion for paragraph runs.
 *
 * Processes shape options and returns marked run for embedding in paragraph.
 *
 * @param {ShapeType} shapeType - Shape to create
 * @param {ShapeOptions} [options] - Shape configuration
 * @returns {Object} Shape run marker for paragraph
 */
export function createShapeRun(
  shapeType: ShapeType,
  options?: ShapeOptions,
): any {
  const shapeXML = buildShapeXML(shapeType, options);
  return { text: shapeXML, isShape: true };
}
