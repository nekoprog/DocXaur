/**
 * Core utilities for DocXaur.
 * @module
 */

import { extension } from "@std/media-types";

/** Convert centimeters to twips (1 twip = 1/20 pt). */
export function cmToTwips(cm: number): number {
  return Math.round(cm * 567);
}

/** Convert centimeters to EMU (English Metric Units). */
export function cmToEmu(cm: number): number {
  return Math.round(cm * 360000);
}

/** Convert points to half-points (Word sizing unit). */
export function ptToHalfPoints(pt: number): number {
  return Math.round(pt * 2);
}

/** Parse a width/height like `10cm`, `12pt`, `25mm`, `1in`, `50%` into twips. */
export function parseNumberTwips(width: string): number {
  const match = width.match(/^([\d.]+)(cm|pt|mm|in|%)$/);
  if (!match) return 1000;
  const value = parseFloat(match[1]);
  const unit = match[2];
  switch (unit) {
    case "cm":
      return cmToTwips(value);
    case "mm":
      return Math.round(value * 56.7);
    case "pt":
      return Math.round(value * 20);
    case "in":
      return Math.round(value * 1440);
    case "%":
      return Math.round(value * 50);
    default:
      return 1000;
  }
}

/** Convert base64-encoded binary to Uint8Array. */
export function base64ToUint8Array(base64: string): Uint8Array {
  const binary = atob(base64);
  const len = binary.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

/** Convert size strings (`cm`, `pt`, `mm`, `in`, `px`) into EMU. */
export function parseImageSize(size: string): number {
  const match = size.match(/^([\d.]+)(cm|pt|mm|in|px)$/);
  if (!match) return cmToEmu(5);
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
      return cmToEmu(5);
  }
}

/** Escape XML text content. */
export function escapeXML(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/**
 * Fetch an image (HTTP/HTTPS or absolute path from Fresh `static/`) and return base64 + extension.
 * Throws if fetch fails.
 * @param url Publicly accessible image URL or absolute `/images/...` path.
 * @returns `{ data, extension }` where `data` is base64 and `extension` comes from Content-Type.
 */
export async function fetchImageAsBase64(
  url: string,
): Promise<{ data: string; extension: string }> {
  if (
    !url.startsWith("http://") && !url.startsWith("https://") &&
    !url.startsWith("/")
  ) {
    throw new Error(
      `Invalid image URL: "${url}". Only HTTP/HTTPS or absolute '/images/...'.`,
    );
  }

  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(
      `Failed to fetch image: ${response.status} ${response.statusText}`,
    );
  }

  const arrayBuffer = await response.arrayBuffer();
  const bytes = new Uint8Array(arrayBuffer);
  let binary = "";
  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  const base64 = btoa(binary);

  const contentType =
    response.headers.get("Content-Type")?.split(";")[0].trim() ?? "";
  const ext = extension(contentType) ?? "png";
  return { data: base64, extension: ext };
}
