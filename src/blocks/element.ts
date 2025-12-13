/**
 * Base Element class for all document blocks.
 *
 * Concrete block implementations (e.g. Paragraph, Image, Table) extend
 * this abstract class and implement `toXML()` to produce the WordprocessingML
 * fragment for that block.
 *
 * This class is intentionally small and separated into its own module to
 * avoid circular import/initialization issues between `section.ts` and the
 * various block implementations.
 */
export abstract class Element {
  abstract toXML(): string;
}
