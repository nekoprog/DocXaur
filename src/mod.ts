/**
 * DocXaur public entrypoint (re-exports).
 *
 * This module exports the public API of DocXaur:
 * - Core builder and types from `core/docxaur.ts`
 * - Utility helpers from `core/*`
 * - Block types (Section, Paragraph, Image, Table) from `blocks/*`
 *
 * Exported symbols are documented in their respective modules. Re-exporting
 * them here provides a single import surface for consumers:
 *
 *   import { DocXaur, Section, Paragraph } from "your/path/to/DocXaur/src/mod.ts";
 *
 * Note: keep this file minimal — per-symbol docs live in the original modules.
 */
export * from "./core/docxaur.ts";
export * from "./core/relationships.ts";
export * from "./core/utils.ts";

export * from "./blocks/section.ts";
export * from "./blocks/paragraph.ts";
export * from "./blocks/shapes.ts";
export * from "./blocks/image.ts";
export * from "./blocks/table.ts";
