/**
 * Relationship helpers for the generated .docx package.
 *
 * This module ensures that the `word/_rels/document.xml.rels` part contains
 * relationship entries for images that are embedded in the document.
 *
 * The implementation will merge new image relationships into an existing
 * relationships XML string or create a new relationships document if none
 * is provided.
 */
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const IMAGE_REL =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

/**
 * Represents a single image relationship entry used in the document relationships.
 *
 * - `rid` is the relationship id (e.g. `"rId5"`).
 * - `target` is the target path inside the package (e.g. `"media/image1.png"`).
 */
export interface ImageRelationship {
  /** Relationship id (e.g. "rId5"). */
  rid: string;
  /** Target path in the .docx package (e.g. "media/image1.png"). */
  target: string;
}

/**
 * Ensure the relationships XML contains the provided image relationships.
 *
 * This function:
 * - Accepts the existing relationships XML (`relsXml`) or `undefined`.
 * - Adds any image relationship entries from `images` that are not already
 *   present in the existing XML.
 * - Returns a well-formed relationships XML string.
 *
 * @param relsXml Existing relationships XML or undefined.
 * @param images Array of image relationships to ensure are present.
 * @returns Merged/created relationships XML including the provided images.
 */
export function ensureImageRelationships(
  relsXml: string | undefined,
  images: ImageRelationship[],
): string {
  const xmlHeader = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`;
  const openTag = `<Relationships xmlns="${REL_NS}">`;
  const closeTag = `</Relationships>`;
  const base = relsXml && relsXml.trim().length
    ? relsXml
    : `${xmlHeader}${openTag}${closeTag}`;

  const existingIds = new Set(
    (base.match(/Id="(rId\d+)"/g) ?? []).map((s) =>
      s.replace(/.*Id="(rId\d+)".*/, "$1")
    ),
  );
  const additions = images
    .filter((img) => !existingIds.has(img.rid))
    .map(
      (img) =>
        `<Relationship Id="${img.rid}" Type="${IMAGE_REL}" Target="${img.target}"/>`,
    )
    .join("");

  return base.includes(closeTag)
    ? base.replace(closeTag, additions + closeTag)
    : `${xmlHeader}${openTag}${additions}${closeTag}`;
}
