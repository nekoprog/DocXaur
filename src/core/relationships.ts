/**
 * Relationship helpers: ensure all image relationships exist in `word/_rels/document.xml.rels`.
 * @module
 */

const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const IMAGE_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

/**
 * Append missing image `<Relationship …>` entries to `document.xml.rels`.
 * @param relsXml Existing relationships XML.
 * @param images Array of `{ rid, target }` where `target` is `media/<file>`.
 * @returns Fixed relationships XML.
 */
export function ensureImageRelationships(
  relsXml: string | undefined,
  images: Array<{ rid: string; target: string }>,
): string {
  const xmlHeader = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
  const openTag = `<Relationships xmlns="${REL_NS}">`;
  const closeTag = `</Relationships>`;
  const base = (relsXml && relsXml.trim().length)
    ? relsXml
    : `${xmlHeader}${openTag}${closeTag}`;

  const existingIds = new Set(
    (base.match(/Id="(rId\d+)"/g) ?? []).map((s) =>
      s.replace(/.*Id="(rId\d+)".*/, "$1")
    ),
  );
  const additions = images
    .filter((img) => !existingIds.has(img.rid))
    .map((img) =>
      `<Relationship Id="${img.rid}" Type="${IMAGE_REL}" Target="${img.target}"/>`
    )
    .join("");

  return base.includes(closeTag)
    ? base.replace(closeTag, additions + closeTag)
    : `${xmlHeader}${openTag}${additions}${closeTag}`;
}
