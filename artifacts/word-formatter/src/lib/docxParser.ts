/**
 * DOCX → ParsedBlock[]
 *
 * Strategy (in order of preference):
 *  1. Direct XML parsing via JSZip (parseDocxXml) – exact structure, no HTML artifacts
 *  2. Mammoth HTML conversion – for images and complex formatting
 *  3. Raw text fallback – last resort
 */
import mammoth from "mammoth";
import { ParsedBlock, TableData, parseText } from "./tableDetector";
import { normalizeArabicText, hasPresentationForms } from "./normalizeArabic";
import { parseDocxXml, extractRawTextFromDocx } from "./docxXmlParser";

// ─── Mammoth HTML helpers (used only as secondary fallback) ────────────────────

function getCellTextMammoth(cell: Element): string {
  return (cell.textContent ?? "").replace(/\s+/g, " ").trim();
}

function getAlignmentMammoth(el: Element): "right" | "left" | "center" | "justify" | undefined {
  const style = (el.getAttribute("style") ?? "").toLowerCase();
  const cls = (el.getAttribute("class") ?? "").toLowerCase();
  if (
    style.includes("text-align:center") ||
    style.includes("text-align: center") ||
    cls.includes("center")
  )
    return "center";
  if (style.includes("text-align:right") || style.includes("text-align: right") || cls.includes("right"))
    return "right";
  if (style.includes("text-align:left") || style.includes("text-align: left") || cls.includes("left"))
    return "left";
  if (style.includes("text-align:justify") || style.includes("text-align: justify"))
    return "justify";
  return undefined;
}

function parseHtmlTable(tableEl: Element): TableData | null {
  const rows = Array.from(tableEl.querySelectorAll("tr"));
  if (rows.length === 0) return null;

  const allRows: string[][] = rows
    .map((row) => {
      const cells = Array.from(row.querySelectorAll("td, th"));
      return cells.map((c) => getCellTextMammoth(c));
    })
    .filter((row) => row.some((cell) => cell.length > 0));

  if (allRows.length === 0) return null;

  const maxCols = Math.max(...allRows.map((r) => r.length));
  if (maxCols < 2) return null;

  const normalised = allRows.map((r) =>
    Array.from({ length: maxCols }, (_, i) => r[i] ?? "")
  );

  return { headers: normalised[0], rows: normalised.slice(1) };
}

function tableAsParagraphs(tableEl: Element, blocks: ParsedBlock[]): void {
  const rows = Array.from(tableEl.querySelectorAll("tr"));
  for (const row of rows) {
    const cells = Array.from(row.querySelectorAll("td, th"))
      .map((c) => getCellTextMammoth(c))
      .filter((t) => t.length > 0);
    if (cells.length === 0) continue;
    blocks.push({
      type: "paragraph",
      text: cells.length === 1 ? cells[0] : cells.join("  |  "),
    });
  }
}

function processMammothElement(el: Element, blocks: ParsedBlock[]): void {
  const tag = el.tagName.toUpperCase();

  switch (tag) {
    case "IMG": {
      const src = el.getAttribute("src") ?? "";
      if (src.startsWith("data:image")) {
        blocks.push({ type: "image", imageData: src, imageAlt: el.getAttribute("alt") ?? "صورة" });
      }
      break;
    }

    case "TABLE": {
      const tableData = parseHtmlTable(el);
      if (tableData) {
        blocks.push({ type: "empty" });
        blocks.push({ type: "table", table: tableData });
        blocks.push({ type: "empty" });
      } else {
        tableAsParagraphs(el, blocks);
      }
      break;
    }

    case "H1":
    case "H2": {
      const text = (el.textContent ?? "").trim();
      if (text) blocks.push({ type: "heading1", text, alignment: getAlignmentMammoth(el) });
      break;
    }

    case "H3":
    case "H4":
    case "H5":
    case "H6": {
      const text = (el.textContent ?? "").trim();
      if (text) blocks.push({ type: "heading2", text, alignment: getAlignmentMammoth(el) });
      break;
    }

    case "P": {
      // Images inside paragraph
      const imgs = Array.from(el.querySelectorAll("img"));
      if (imgs.length > 0) {
        for (const img of imgs) {
          const src = img.getAttribute("src") ?? "";
          if (src.startsWith("data:image")) {
            blocks.push({ type: "image", imageData: src, imageAlt: img.getAttribute("alt") ?? "صورة" });
          }
        }
        break;
      }

      // Nested table
      const nestedTables = Array.from(el.children).filter(
        (c) => c.tagName.toUpperCase() === "TABLE"
      );
      if (nestedTables.length > 0) {
        for (const tbl of nestedTables) processMammothElement(tbl, blocks);
        break;
      }

      const text = (el.textContent ?? "").replace(/\s+/g, " ").trim();
      const align = getAlignmentMammoth(el);
      if (!text) {
        blocks.push({ type: "empty" });
      } else {
        blocks.push({ type: "paragraph", text, alignment: align });
      }
      break;
    }

    case "UL":
    case "OL": {
      for (const li of Array.from(el.querySelectorAll(":scope > li"))) {
        const text = (li.textContent ?? "").trim();
        if (text) blocks.push({ type: "paragraph", text: `• ${text}` });
      }
      break;
    }

    case "FIGURE":
    default: {
      for (const child of Array.from(el.children)) {
        processMammothElement(child, blocks);
      }
    }
  }
}

// ─── Main export ──────────────────────────────────────────────────────────────

export async function parseDocxToBlocks(arrayBuffer: ArrayBuffer): Promise<{
  blocks: ParsedBlock[];
  rawText: string;
  hadEncodingIssues: boolean;
}> {
  // ── Step 1: Try direct XML parser (most accurate) ──────────────────────────
  const xmlBlocks = await parseDocxXml(arrayBuffer);
  const rawTextFromXml = await extractRawTextFromDocx(arrayBuffer);

  const hadEncodingIssues = hasPresentationForms(rawTextFromXml);

  if (xmlBlocks && xmlBlocks.some((b) => b.type !== "empty" && (b.text || b.table))) {
    // XML parser succeeded - this is the preferred path
    return {
      blocks: xmlBlocks,
      rawText: rawTextFromXml || normalizeArabicText(
        (await mammoth.extractRawText({ arrayBuffer })).value
      ),
      hadEncodingIssues,
    };
  }

  // ── Step 2: Fall back to mammoth HTML conversion ───────────────────────────
  let rawText = rawTextFromXml;
  if (!rawText) {
    try {
      rawText = (await mammoth.extractRawText({ arrayBuffer })).value;
    } catch {
      rawText = "";
    }
  }

  if (!rawText.trim()) {
    throw new Error("لم يتم العثور على نص قابل للاستخراج في الملف");
  }

  let htmlValue = "";
  try {
    const htmlResult = await mammoth.convertToHtml(
      { arrayBuffer },
      {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        convertImage: (mammoth as any).images.imgElement((image: any) =>
          image.read("base64").then((base64Data: string) => ({
            src: `data:${image.contentType};base64,${base64Data}`,
          }))
        ),
      }
    );
    htmlValue = htmlResult.value;
  } catch {
    try {
      htmlValue = (await mammoth.convertToHtml({ arrayBuffer })).value;
    } catch {
      const normalizedRaw = normalizeArabicText(rawText);
      return {
        blocks: parseText(normalizedRaw),
        rawText: normalizedRaw,
        hadEncodingIssues,
      };
    }
  }

  const parser = new DOMParser();
  const doc = parser.parseFromString(htmlValue, "text/html");

  const blocks: ParsedBlock[] = [];
  for (const child of Array.from(doc.body.children)) {
    processMammothElement(child, blocks);
  }

  while (blocks.length > 0 && blocks[blocks.length - 1].type === "empty") {
    blocks.pop();
  }

  const hasContent = blocks.some(
    (b) => b.type !== "empty" && (b.text || b.table || b.imageData)
  );
  if (!hasContent) {
    const normalizedRaw = normalizeArabicText(rawText);
    return { blocks: parseText(normalizedRaw), rawText: normalizedRaw, hadEncodingIssues };
  }

  return { blocks, rawText: normalizeArabicText(rawText), hadEncodingIssues };
}
