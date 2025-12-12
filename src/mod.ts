/**
 * DocXaur public entry (blocks).
 * @module
 *
 * JSR will generate documentation from these JSDoc-style comments:
 * - Symbol docs: add JSDoc above each export
 * - Module docs: add `@module` at top of each file
 * Learn more: https://jsr.io/docs/writing-docs
 */

// Side-effect imports: augment Section with paragraph/heading, image, table
import "./blocks/paragraph.ts";
import "./blocks/image.ts";
import "./blocks/table.ts";

// Public API re-exports
export * from "./core/docxaur.ts";
export * from "./core/relationships.ts";
export * from "./core/utils.ts";

export * from "./blocks/section.ts";
export * from "./blocks/paragraph.ts";
export * from "./blocks/image.ts";
export * from "./blocks/table.ts";
