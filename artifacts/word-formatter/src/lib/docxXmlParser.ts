/**
 * Direct DOCX XML parser using JSZip.
 *
 * KEY FIXES vs previous version:
 *  1. getParaText now SKIPS drawing/pict/mc:AlternateContent → no more text-box
 *     content being merged into a single line, no more cover-page duplication.
 *  2. Paragraphs that contain a <w:drawing> with an image blip are extracted
 *     as image blocks (reads from word/media/ in the ZIP).
 *  3. Text-box drawings (wps:txbx) are skipped entirely – their text already
 *     appears as normal body paragraphs in the OOXML stream.
 *  4. Relationships map (word/_rels/document.xml.rels) used for image paths.
 */
import JSZip from "jszip";
import { ParsedBlock } from "./tableDetector";

// ─── Namespace-agnostic DOM helpers ──────────────────────────────────────────

/** Direct children with the given local name. */
function ch(el: Element, name: string): Element[] {
  const out: Element[] = [];
  for (const n of Array.from(el.childNodes)) {
    if (n.nodeType === 1 && (n as Element).localName === name) out.push(n as Element);
  }
  return out;
}

/** w: attribute value (handles any prefix or no-prefix). */
function wa(el: Element, attr: string): string {
  return (
    el.getAttributeNS(
      "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
      attr
    ) ??
    el.getAttribute(`w:${attr}`) ??
    el.getAttribute(attr) ??
    ""
  );
}

// ─── Text extraction ──────────────────────────────────────────────────────────

/**
 * Extract text from a paragraph's runs.
 * CRITICAL: does NOT recurse into drawing / pict / mc:AlternateContent
 * so that floating text-box content is never merged here.
 */
function getParaText(p: Element): string {
  const SKIP = new Set([
    "drawing", "pict", "object",
    "AlternateContent",          // mc:AlternateContent
    "rPr", "pPr",                // property elements carry no text
    "fldChar", "instrText",      // field machinery
    "bookmarkStart", "bookmarkEnd",
    "proofErr", "rPrChange", "pPrChange",
    "del",                       // deleted (tracked changes) text
  ]);

  const parts: string[] = [];

  function walk(node: Element): void {
    for (const child of Array.from(node.childNodes)) {
      if (child.nodeType !== 1) continue;
      const el = child as Element;
      const name = el.localName;

      if (SKIP.has(name)) continue;

      if (name === "t") {
        parts.push(el.textContent ?? "");
      } else if (name === "tab") {
        parts.push(" ");
      } else if (name === "br") {
        // soft line-break → space, not a new paragraph
        parts.push(" ");
      } else {
        walk(el);
      }
    }
  }

  walk(p);
  return parts.join("").replace(/ {2,}/g, " ").trim();
}

/** Text from all paragraphs inside a table cell (joined with space). */
function getCellText(tc: Element): string {
  return ch(tc, "p")
    .map(getParaText)
    .filter(Boolean)
    .join(" ")
    .trim();
}

// ─── Paragraph properties ─────────────────────────────────────────────────────

function getParaStyle(p: Element): string {
  const pPr = ch(p, "pPr")[0];
  if (!pPr) return "";
  const pStyle = ch(pPr, "pStyle")[0];
  return pStyle ? wa(pStyle, "val").toLowerCase() : "";
}

function getParaAlignment(p: Element): "right" | "left" | "center" | "justify" | undefined {
  const pPr = ch(p, "pPr")[0];
  if (!pPr) return undefined;
  const jc = ch(pPr, "jc")[0];
  if (!jc) return undefined;
  const v = wa(jc, "val");
  if (v === "center") return "center";
  if (v === "left") return "left";
  if (v === "right") return "right";
  if (v === "both" || v.includes("distribute") || v === "thaiDistribute") return "justify";
  return undefined;
}

function getHeadingLevel(p: Element): 1 | 2 | null {
  const style = getParaStyle(p);
  if (!style) {
    // Fall back to outline level
    const pPr = ch(p, "pPr")[0];
    if (pPr) {
      const ol = ch(pPr, "outlineLvl")[0];
      if (ol) {
        const lvl = parseInt(wa(ol, "val") ?? "9");
        if (!isNaN(lvl)) {
          if (lvl === 0) return 1;
          if (lvl <= 2) return 2;
        }
      }
    }
    return null;
  }

  if (
    style === "heading1" || style === "1" || style === "title" ||
    style === "subtitle" || style.startsWith("heading1")
  ) return 1;

  if (
    style === "heading2" || style === "2" || style.startsWith("heading2")
  ) return 2;

  if (style.startsWith("heading")) {
    const m = style.match(/\d+/);
    if (m) {
      const n = parseInt(m[0]);
      if (n === 1) return 1;
      if (n <= 3) return 2;
    }
    return 1;
  }

  return null;
}

// ─── Image extraction ─────────────────────────────────────────────────────────

interface RelMap { [rId: string]: string }

async function loadRelationships(zip: JSZip): Promise<RelMap> {
  const map: RelMap = {};
  const relsFile = zip.file("word/_rels/document.xml.rels");
  if (!relsFile) return map;

  const xml = await relsFile.async("text");
  const doc = new DOMParser().parseFromString(xml, "text/xml");

  for (const rel of Array.from(doc.getElementsByTagName("Relationship"))) {
    const id = rel.getAttribute("Id") ?? "";
    const target = rel.getAttribute("Target") ?? "";
    const type = rel.getAttribute("Type") ?? "";
    if (!id || !target) continue;
    // Only care about image relationships
    if (type.includes("/image")) {
      const path = target.startsWith("/")
        ? target.slice(1)
        : `word/${target}`;
      map[id] = path;
    }
  }
  return map;
}

/**
 * Walk an element looking for the first a:blip r:embed attribute.
 * Returns the relationship ID, or null if none found.
 */
function findBlipEmbed(el: Element): string | null {
  // a:blip may carry the embed ID
  if (el.localName === "blip") {
    const embed =
      el.getAttributeNS(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "embed"
      ) ??
      el.getAttribute("r:embed") ??
      el.getAttribute("embed");
    if (embed) return embed;
  }

  for (const child of Array.from(el.childNodes)) {
    if (child.nodeType !== 1) continue;
    const found = findBlipEmbed(child as Element);
    if (found) return found;
  }
  return null;
}

/** Returns true if the paragraph contains a drawing/pict element (image OR shape). */
function hasParagraphDrawing(p: Element): boolean {
  for (const r of ch(p, "r")) {
    if (ch(r, "drawing").length > 0 || ch(r, "pict").length > 0) return true;
  }
  return false;
}

/** Returns true if the drawing subtree contains a text-box (wps:txbx / v:textbox). */
function isTextBoxDrawing(drawingEl: Element): boolean {
  function walk(el: Element): boolean {
    for (const child of Array.from(el.childNodes)) {
      if (child.nodeType !== 1) continue;
      const c = child as Element;
      if (c.localName === "txbx" || c.localName === "textbox" || c.localName === "txbxContent") return true;
      if (walk(c)) return true;
    }
    return false;
  }
  return walk(drawingEl);
}

/** EMU → pixels (96 DPI) */
function emuToPx(emu: number): number {
  return Math.round(emu / 9525);
}

/**
 * Try to get image dimensions from the drawing element (EMU → px).
 * Looks for a:ext cx/cy attributes.
 */
function getDrawingSize(drawingEl: Element): { w: number; h: number } {
  function findExt(el: Element): { w: number; h: number } | null {
    if (el.localName === "ext") {
      const cx = parseInt(el.getAttribute("cx") ?? "0");
      const cy = parseInt(el.getAttribute("cy") ?? "0");
      if (cx > 0 && cy > 0) return { w: emuToPx(cx), h: emuToPx(cy) };
    }
    for (const child of Array.from(el.childNodes)) {
      if (child.nodeType !== 1) continue;
      const r = findExt(child as Element);
      if (r) return r;
    }
    return null;
  }
  return findExt(drawingEl) ?? { w: 400, h: 300 };
}

const MIME_MAP: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  svg: "image/svg+xml",
  tif: "image/tiff",
  tiff: "image/tiff",
};

/**
 * Try to extract image block(s) from a paragraph that has <w:drawing>.
 * Returns true if at least one image was added.
 */
async function tryExtractImages(
  p: Element,
  zip: JSZip,
  rels: RelMap,
  blocks: ParsedBlock[]
): Promise<boolean> {
  let found = false;

  for (const r of ch(p, "r")) {
    const drawings = [...ch(r, "drawing"), ...ch(r, "pict")];
    for (const drawing of drawings) {
      // Skip text boxes
      if (isTextBoxDrawing(drawing)) continue;

      const rId = findBlipEmbed(drawing);
      if (!rId) continue;

      const mediaPath = rels[rId];
      if (!mediaPath) continue;

      const mediaFile = zip.file(mediaPath);
      if (!mediaFile) continue;

      const ext = mediaPath.split(".").pop()?.toLowerCase() ?? "png";

      // Skip WMF/EMF – these are Windows metafiles, show placeholder instead
      if (ext === "wmf" || ext === "emf") {
        blocks.push({ type: "paragraph", text: "[ رسم / شكل ]", alignment: "center" });
        found = true;
        continue;
      }

      const mime = MIME_MAP[ext] ?? "image/png";

      try {
        const base64 = await mediaFile.async("base64");
        const { w, h } = getDrawingSize(drawing);

        // Clamp dimensions to reasonable maximums
        const maxW = 500;
        const scale = w > maxW ? maxW / w : 1;

        blocks.push({
          type: "image",
          imageData: `data:${mime};base64,${base64}`,
          imageAlt: "صورة/رسم بياني",
          imageWidth: Math.round(w * scale),
          imageHeight: Math.round(h * scale),
        });
        found = true;
      } catch {
        // Skip unreadable media
      }
    }
  }

  return found;
}

// ─── Paragraph processor ──────────────────────────────────────────────────────

async function processXmlParagraph(
  p: Element,
  blocks: ParsedBlock[],
  zip: JSZip,
  rels: RelMap
): Promise<void> {
  if (hasParagraphDrawing(p)) {
    // Try to extract as image; ignore text-box drawings
    const added = await tryExtractImages(p, zip, rels, blocks);
    // If nothing was added (pure decorative shape), add nothing
    // Do NOT also add the text, because the paragraph body is empty (the content lives in the drawing)
    if (!added) {
      // Check if there is also real text OUTSIDE the drawing
      const text = getParaText(p);
      if (text) blocks.push({ type: "paragraph", text, alignment: getParaAlignment(p) });
    }
    return;
  }

  const text = getParaText(p);
  const alignment = getParaAlignment(p);

  if (!text) {
    blocks.push({ type: "empty" });
    return;
  }

  const heading = getHeadingLevel(p);
  if (heading === 1) {
    blocks.push({ type: "heading1", text, alignment });
  } else if (heading === 2) {
    blocks.push({ type: "heading2", text, alignment });
  } else {
    blocks.push({ type: "paragraph", text, alignment });
  }
}

// ─── Table processor ──────────────────────────────────────────────────────────

function processXmlTable(tbl: Element, blocks: ParsedBlock[]): void {
  const rows = ch(tbl, "tr");
  if (!rows.length) return;

  const allRows: string[][] = rows
    .map((tr) =>
      ch(tr, "tc")
        .filter((tc) => {
          // Skip vertical-merge continuation cells
          const tcPr = ch(tc, "tcPr")[0];
          if (!tcPr) return true;
          const vm = ch(tcPr, "vMerge")[0];
          if (!vm) return true;
          return wa(vm, "val") === "restart";
        })
        .map((tc) => getCellText(tc))
    )
    .filter((row) => row.some((c) => c.length > 0));

  if (!allRows.length) return;

  const maxCols = Math.max(...allRows.map((r) => r.length));

  if (maxCols < 2) {
    // Single-column table → paragraphs
    for (const row of allRows) {
      for (const cell of row) {
        if (cell.trim()) blocks.push({ type: "paragraph", text: cell });
      }
    }
    return;
  }

  const normalized = allRows.map((r) =>
    Array.from({ length: maxCols }, (_, i) => (r[i] ?? "").replace(/\n/g, " ").trim())
  );

  blocks.push({ type: "empty" });
  blocks.push({ type: "table", table: { headers: normalized[0], rows: normalized.slice(1) } });
  blocks.push({ type: "empty" });
}

// ─── SDT (Structured Document Tag, e.g. TOC) processor ───────────────────────

async function processXmlSdt(
  sdt: Element,
  blocks: ParsedBlock[],
  zip: JSZip,
  rels: RelMap
): Promise<void> {
  const content = ch(sdt, "sdtContent")[0];
  if (content) await processXmlChildren(content, blocks, zip, rels);
}

// ─── Container children processor ────────────────────────────────────────────

async function processXmlChildren(
  container: Element,
  blocks: ParsedBlock[],
  zip: JSZip,
  rels: RelMap
): Promise<void> {
  for (const node of Array.from(container.childNodes)) {
    if (node.nodeType !== 1) continue;
    const el = node as Element;
    const name = el.localName;

    if (name === "p") {
      await processXmlParagraph(el, blocks, zip, rels);
    } else if (name === "tbl") {
      processXmlTable(el, blocks);
    } else if (name === "sdt") {
      await processXmlSdt(el, blocks, zip, rels);
    }
    // sectPr, bookmarkStart, bookmarkEnd, mc:AlternateContent → ignored
  }
}

// ─── Post-processing ──────────────────────────────────────────────────────────

function cleanBlocks(raw: ParsedBlock[]): ParsedBlock[] {
  // Remove trailing empties
  while (raw.length && raw[raw.length - 1].type === "empty") raw.pop();

  // Collapse 3+ consecutive empties → 2
  const out: ParsedBlock[] = [];
  let empties = 0;
  for (const b of raw) {
    if (b.type === "empty") {
      empties++;
      if (empties <= 2) out.push(b);
    } else {
      empties = 0;
      out.push(b);
    }
  }
  return out;
}

// ─── Main exports ─────────────────────────────────────────────────────────────

/**
 * Parse a DOCX ArrayBuffer directly from OOXML.
 * Returns null on failure so the caller can fall back to mammoth.
 */
export async function parseDocxXml(
  arrayBuffer: ArrayBuffer
): Promise<ParsedBlock[] | null> {
  try {
    const zip = await JSZip.loadAsync(arrayBuffer);

    const docFile = zip.file("word/document.xml");
    if (!docFile) return null;

    const xmlText = await docFile.async("text");
    const xmlDoc = new DOMParser().parseFromString(xmlText, "text/xml");

    if (xmlDoc.querySelector("parsererror")) return null;

    // Find <w:body>
    let body: Element | null = null;
    const byTagName = xmlDoc.getElementsByTagName("w:body");
    if (byTagName.length) {
      body = byTagName[0];
    } else {
      const byNs = xmlDoc.getElementsByTagNameNS(
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "body"
      );
      if (byNs.length) body = byNs[0];
    }
    if (!body) return null;

    // Load image relationships
    const rels = await loadRelationships(zip);

    const blocks: ParsedBlock[] = [];
    await processXmlChildren(body, blocks, zip, rels);

    const cleaned = cleanBlocks(blocks);
    return cleaned.length ? cleaned : null;
  } catch {
    return null;
  }
}

/**
 * Extract raw text from a DOCX for display in the textarea.
 */
export async function extractRawTextFromDocx(
  arrayBuffer: ArrayBuffer
): Promise<string> {
  try {
    const zip = await JSZip.loadAsync(arrayBuffer);
    const docFile = zip.file("word/document.xml");
    if (!docFile) return "";

    const xmlText = await docFile.async("text");
    const xmlDoc = new DOMParser().parseFromString(xmlText, "text/xml");

    const lines: string[] = [];
    for (const t of Array.from(xmlDoc.getElementsByTagName("w:t"))) {
      const text = (t.textContent ?? "").trim();
      if (text) lines.push(text);
    }
    return lines.join("\n");
  } catch {
    return "";
  }
}
