import html2canvas from "html2canvas";
import jsPDF from "jspdf";
import { saveAs } from "file-saver";
import * as XLSX from "xlsx";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  AlignmentType,
  SectionType,
  convertInchesToTwip,
  BorderStyle,
  ShadingType,
  TableLayoutType,
  PageBorderDisplay,
  PageBorderOffsetFrom,
} from "docx";

// ─── Font constants ────────────────────────────────────────────────────────────
const ARABIC_FONT = "Simplified Arabic";
const ENGLISH_FONT = "Times New Roman";

function isArabicDominant(text: string): boolean {
  const ar = (text.match(/[\u0600-\u06FF\uFB50-\uFDFF\uFE70-\uFEFF]/g) || []).length;
  const en = (text.match(/[a-zA-Z]/g) || []).length;
  return ar >= en;
}

// ─── HTML → PDF ──────────────────────────────────────────────────────────────

export async function exportHtmlAsPdf(html: string, fileName: string): Promise<void> {
  const container = document.createElement("div");
  container.style.cssText = [
    "position:fixed",
    "left:-9999px",
    "top:0",
    "width:794px",
    "min-height:200px",
    "padding:56px 64px",
    "box-sizing:border-box",
    "background:white",
    "font-family:'Times New Roman',serif",
    "font-size:14pt",
    "line-height:1.6",
    "color:#000",
  ].join(";");
  container.innerHTML = html;
  document.body.appendChild(container);

  try {
    const canvas = await html2canvas(container, {
      scale: 2,
      useCORS: true,
      allowTaint: true,
      logging: false,
      backgroundColor: "#ffffff",
    });

    const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
    const pageWidthMm = pdf.internal.pageSize.getWidth();
    const pageHeightMm = pdf.internal.pageSize.getHeight();
    const marginMm = 10;
    const contentWidthMm = pageWidthMm - marginMm * 2;
    const canvasWidthPx = canvas.width;
    const canvasHeightPx = canvas.height;
    const pxPerMm = canvasWidthPx / contentWidthMm;
    const contentHeightMm = canvasHeightPx / pxPerMm;
    const pageContentHeightMm = pageHeightMm - marginMm * 2;

    let sourceYMm = 0;
    let firstPage = true;

    while (sourceYMm < contentHeightMm) {
      if (!firstPage) pdf.addPage();
      firstPage = false;

      const sliceHeightMm = Math.min(pageContentHeightMm, contentHeightMm - sourceYMm);
      const sourceYPx = Math.round(sourceYMm * pxPerMm);
      const sliceHeightPx = Math.round(sliceHeightMm * pxPerMm);

      const sliceCanvas = document.createElement("canvas");
      sliceCanvas.width = canvasWidthPx;
      sliceCanvas.height = sliceHeightPx;
      const ctx = sliceCanvas.getContext("2d")!;
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(0, 0, sliceCanvas.width, sliceCanvas.height);
      ctx.drawImage(canvas, 0, sourceYPx, canvasWidthPx, sliceHeightPx, 0, 0, canvasWidthPx, sliceHeightPx);

      const imgData = sliceCanvas.toDataURL("image/jpeg", 0.95);
      pdf.addImage(imgData, "JPEG", marginMm, marginMm, contentWidthMm, sliceHeightMm);
      sourceYMm += pageContentHeightMm;
    }

    pdf.save(`${fileName}.pdf`);
  } finally {
    document.body.removeChild(container);
  }
}

// ─── Rich inline text (preserves bold, italic, underline, color, size) ────────

interface RunStyle {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string;
  size?: number; // in pt
  strike?: boolean;
}

function extractInlineStyle(el: Element, inherited: RunStyle): RunStyle {
  const style = el.getAttribute("style") ?? "";
  const tag = el.tagName.toUpperCase();
  const result: RunStyle = { ...inherited };

  if (
    tag === "B" || tag === "STRONG" ||
    /font-weight\s*:\s*(bold|[789]\d\d)/i.test(style)
  ) result.bold = true;

  if (tag === "I" || tag === "EM" || /font-style\s*:\s*italic/i.test(style))
    result.italic = true;

  if (tag === "U" || /text-decoration[^;]*:\s*[^;]*underline/i.test(style))
    result.underline = true;

  if (tag === "S" || tag === "STRIKE" || tag === "DEL" ||
      /text-decoration[^;]*:\s*[^;]*line-through/i.test(style))
    result.strike = true;

  // Color: support #hex and rgb()
  const colorM = style.match(/(?:^|;)\s*color\s*:\s*(#[0-9a-fA-F]{3,6}|rgb\([^)]+\))/i);
  if (colorM) {
    const raw = colorM[1].trim();
    if (raw.startsWith("#")) {
      const hex = raw.slice(1);
      result.color = (hex.length === 3
        ? hex.split("").map((c) => c + c).join("")
        : hex
      ).toUpperCase().slice(0, 6);
    } else {
      const rgbM = raw.match(/rgb\((\d+),\s*(\d+),\s*(\d+)\)/i);
      if (rgbM) {
        result.color = [rgbM[1], rgbM[2], rgbM[3]]
          .map((n) => parseInt(n).toString(16).padStart(2, "0"))
          .join("")
          .toUpperCase();
      }
    }
  }

  // Font size
  const szM = style.match(/font-size\s*:\s*(\d+(?:\.\d+)?)(pt|px|em|rem)/i);
  if (szM) {
    const v = parseFloat(szM[1]);
    const u = szM[2].toLowerCase();
    if (u === "pt") result.size = v;
    else if (u === "px") result.size = Math.round(v * 0.75);
    else result.size = Math.round(v * 12);
  }

  return result;
}

function makeRunFromText(text: string, style: RunStyle): TextRun | null {
  const cleaned = text.replace(/\u00a0/g, " ");
  if (!cleaned) return null;

  const arabic = isArabicDominant(cleaned);
  const pt = style.size ?? 14;

  return new TextRun({
    text: cleaned,
    font: arabic
      ? { name: ARABIC_FONT, cs: ARABIC_FONT, eastAsia: ARABIC_FONT }
      : { name: ENGLISH_FONT, cs: ENGLISH_FONT },
    size: pt * 2,
    bold: style.bold ?? false,
    italics: style.italic ?? false,
    underline: style.underline ? { type: "single" as never } : undefined,
    strike: style.strike ?? false,
    color: style.color,
    rightToLeft: arabic,
  });
}

/** Recursively walk DOM nodes and collect TextRuns with full inline style */
function collectRuns(node: Node, style: RunStyle, runs: TextRun[]): void {
  if (node.nodeType === Node.TEXT_NODE) {
    const text = node.textContent ?? "";
    if (!text) return;
    const run = makeRunFromText(text, style);
    if (run) runs.push(run);
    return;
  }

  if (node.nodeType !== Node.ELEMENT_NODE) return;

  const el = node as Element;
  const tag = el.tagName.toUpperCase();

  // Block-level elements inside inline context → just recurse
  const BLOCK_TAGS = new Set([
    "DIV", "P", "H1", "H2", "H3", "H4", "H5", "H6",
    "UL", "OL", "LI", "TABLE", "THEAD", "TBODY", "TFOOT", "TR", "TD", "TH",
    "BLOCKQUOTE", "PRE", "HR", "BR",
  ]);

  if (tag === "BR") {
    runs.push(new TextRun({ break: 1 }));
    return;
  }

  if (BLOCK_TAGS.has(tag)) {
    for (const child of Array.from(el.childNodes)) collectRuns(child, style, runs);
    return;
  }

  const newStyle = extractInlineStyle(el, style);
  for (const child of Array.from(el.childNodes)) collectRuns(child, newStyle, runs);
}

// ─── Alignment detection ──────────────────────────────────────────────────────

function getHtmlAlignment(
  el: Element
): "right" | "left" | "center" | "justify" | undefined {
  const style = (el.getAttribute("style") ?? "").toLowerCase();
  const align = el.getAttribute("align") ?? "";

  if (style.includes("text-align:center") || style.includes("text-align: center") || align === "center") return "center";
  if (style.includes("text-align:right") || style.includes("text-align: right") || align === "right") return "right";
  if (style.includes("text-align:left") || style.includes("text-align: left") || align === "left") return "left";
  if (style.includes("text-align:justify") || style.includes("text-align: justify")) return "justify";
  return undefined;
}

function resolveAlignmentFromStyle(
  alignment: "right" | "left" | "center" | "justify" | undefined,
  arabic: boolean
): AlignmentType {
  if (alignment === "center") return AlignmentType.CENTER;
  if (alignment === "left") return AlignmentType.LEFT;
  if (alignment === "right") return AlignmentType.RIGHT;
  if (alignment === "justify")
    return arabic ? AlignmentType.THAI_DISTRIBUTE : AlignmentType.DISTRIBUTE;
  return arabic ? AlignmentType.RIGHT : AlignmentType.LEFT;
}

// ─── Paragraph builder ────────────────────────────────────────────────────────

function buildParagraphFromEl(
  el: Element,
  baseStyle: RunStyle = {},
  extraOpts: { spacing?: number } = {}
): Paragraph {
  const runs: TextRun[] = [];
  const style = extractInlineStyle(el, baseStyle);
  for (const child of Array.from(el.childNodes)) collectRuns(child, style, runs);

  if (runs.length === 0) runs.push(new TextRun({ text: "" }));

  const text = el.textContent ?? "";
  const arabic = isArabicDominant(text);
  const alignment = getHtmlAlignment(el);

  return new Paragraph({
    children: runs,
    bidirectional: arabic,
    alignment: resolveAlignmentFromStyle(alignment, arabic),
    spacing: { line: extraOpts.spacing ?? 360, lineRule: "auto" as never },
  });
}

// ─── Table from HTML ──────────────────────────────────────────────────────────

function buildTableFromHtml(tableEl: Element): Table | null {
  const rows = Array.from(tableEl.querySelectorAll("tr"));
  if (rows.length === 0) return null;

  const borderAttr = tableEl.getAttribute("border");
  const hasBorder = borderAttr && borderAttr !== "0";

  const borders = hasBorder
    ? {
        top: { style: BorderStyle.SINGLE, size: 4, color: "2C6FAC" },
        bottom: { style: BorderStyle.SINGLE, size: 4, color: "2C6FAC" },
        left: { style: BorderStyle.SINGLE, size: 4, color: "2C6FAC" },
        right: { style: BorderStyle.SINGLE, size: 4, color: "2C6FAC" },
        insideH: { style: BorderStyle.SINGLE, size: 2, color: "A0C4E8" },
        insideV: { style: BorderStyle.SINGLE, size: 2, color: "A0C4E8" },
      }
    : {
        top: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" },
        bottom: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" },
        left: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" },
        right: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" },
        insideH: { style: BorderStyle.SINGLE, size: 1, color: "DDDDDD" },
        insideV: { style: BorderStyle.SINGLE, size: 1, color: "DDDDDD" },
      };

  const headerText = rows[0]
    ? Array.from(rows[0].querySelectorAll("td, th"))
        .map((c) => c.textContent ?? "")
        .join(" ")
    : "";
  const isArabic = isArabicDominant(headerText);

  const docRows = rows.map((tr, ri) => {
    const cells = Array.from(tr.querySelectorAll("td, th"));
    const isHeaderRow = ri === 0 && Array.from(tr.querySelectorAll("th")).length > 0;

    const docCells = cells.map((td) => {
      const isHeader = isHeaderRow || td.tagName.toUpperCase() === "TH";
      const cellRuns: TextRun[] = [];
      for (const child of Array.from(td.childNodes)) {
        collectRuns(child, { bold: isHeader, size: 12 }, cellRuns);
      }
      if (cellRuns.length === 0) cellRuns.push(new TextRun({ text: "" }));

      const cellText = td.textContent ?? "";
      const cellArabic = isArabicDominant(cellText);
      const align = getHtmlAlignment(td);

      return new TableCell({
        children: [
          new Paragraph({
            children: cellRuns,
            bidirectional: cellArabic,
            alignment: resolveAlignmentFromStyle(align, cellArabic),
            spacing: { line: 276 },
          }),
        ],
        shading: isHeader
          ? { fill: "D6E8F8", type: ShadingType.CLEAR, color: "auto" }
          : undefined,
        margins: {
          top: convertInchesToTwip(0.05),
          bottom: convertInchesToTwip(0.05),
          left: convertInchesToTwip(0.1),
          right: convertInchesToTwip(0.1),
        },
      });
    });

    return new TableRow({
      children: docCells,
      tableHeader: ri === 0 && Array.from(tr.querySelectorAll("th")).length > 0,
    });
  });

  if (docRows.length === 0) return null;

  return new Table({
    layout: TableLayoutType.AUTOFIT,
    width: { size: 100, type: WidthType.PERCENTAGE },
    visuallyRightToLeft: isArabic,
    rows: docRows,
    borders,
  });
}

// ─── HTML → docx elements ─────────────────────────────────────────────────────

function processHtmlNode(el: Element, elements: (Paragraph | Table)[]): void {
  const tag = el.tagName.toUpperCase();

  switch (tag) {
    case "H1":
      elements.push(buildParagraphFromEl(el, { bold: true, size: 18 }, { spacing: 400 }));
      break;
    case "H2":
      elements.push(buildParagraphFromEl(el, { bold: true, size: 16 }, { spacing: 380 }));
      break;
    case "H3":
      elements.push(buildParagraphFromEl(el, { bold: true, size: 14 }, { spacing: 360 }));
      break;
    case "H4":
    case "H5":
    case "H6":
      elements.push(buildParagraphFromEl(el, { bold: true, size: 13 }, { spacing: 340 }));
      break;

    case "P": {
      const text = (el.textContent ?? "").trim();
      if (!text) {
        elements.push(new Paragraph({ children: [], spacing: { line: 240 } }));
      } else {
        elements.push(buildParagraphFromEl(el, { size: 14 }, { spacing: 360 }));
      }
      break;
    }

    case "BLOCKQUOTE": {
      const text = (el.textContent ?? "").trim();
      if (text) {
        elements.push(
          buildParagraphFromEl(el, { italic: true, size: 13, color: "555555" }, { spacing: 320 })
        );
      }
      break;
    }

    case "TABLE": {
      const table = buildTableFromHtml(el);
      if (table) {
        elements.push(new Paragraph({ children: [], spacing: { line: 200 } }));
        elements.push(table);
        elements.push(new Paragraph({ children: [], spacing: { line: 200 } }));
      }
      break;
    }

    case "UL":
    case "OL": {
      const lis = Array.from(el.querySelectorAll(":scope > li"));
      lis.forEach((li, i) => {
        const marker = tag === "OL" ? `${i + 1}. ` : "• ";
        const runs: TextRun[] = [];
        const arabic = isArabicDominant(li.textContent ?? "");
        runs.push(
          new TextRun({
            text: marker,
            font: arabic
              ? { name: ARABIC_FONT, cs: ARABIC_FONT }
              : { name: ENGLISH_FONT },
            size: 28,
            rightToLeft: arabic,
          })
        );
        for (const child of Array.from(li.childNodes)) collectRuns(child, { size: 14 }, runs);
        elements.push(
          new Paragraph({
            children: runs,
            bidirectional: arabic,
            alignment: arabic ? AlignmentType.RIGHT : AlignmentType.LEFT,
            spacing: { line: 300 },
            indent: { left: convertInchesToTwip(0.25) },
          })
        );
      });
      break;
    }

    case "HR":
      elements.push(new Paragraph({ children: [], spacing: { line: 120 } }));
      elements.push(new Paragraph({ children: [], spacing: { line: 120 } }));
      break;

    case "BR":
      elements.push(new Paragraph({ children: [], spacing: { line: 180 } }));
      break;

    case "PRE": {
      const text = (el.textContent ?? "").trim();
      if (text) {
        for (const line of text.split("\n")) {
          elements.push(
            new Paragraph({
              children: [
                new TextRun({
                  text: line || " ",
                  font: { name: "Courier New" },
                  size: 22,
                  color: "333333",
                }),
              ],
              spacing: { line: 280 },
            })
          );
        }
      }
      break;
    }

    case "DIV":
    case "SECTION":
    case "ARTICLE":
    case "MAIN":
    case "HEADER":
    case "FOOTER":
    case "ASIDE":
    case "NAV":
    case "FIGURE":
    case "FIGCAPTION":
    case "BODY": {
      for (const child of Array.from(el.children)) processHtmlNode(child, elements);
      break;
    }

    default: {
      // Try to handle unknown elements with text content
      const text = (el.textContent ?? "").trim();
      if (text) {
        elements.push(buildParagraphFromEl(el, { size: 14 }, { spacing: 320 }));
      }
    }
  }
}

// ─── HTML → Word (.docx) ─────────────────────────────────────────────────────

export async function exportHtmlAsDocx(html: string, fileName: string): Promise<void> {
  const parser = new DOMParser();
  const doc = parser.parseFromString(`<html><body>${html}</body></html>`, "text/html");

  const elements: (Paragraph | Table)[] = [];
  for (const child of Array.from(doc.body.children)) {
    processHtmlNode(child, elements);
  }

  // Remove trailing empties
  while (
    elements.length > 0 &&
    elements[elements.length - 1] instanceof Paragraph &&
    (elements[elements.length - 1] as Paragraph).options?.children?.length === 0
  ) {
    elements.pop();
  }

  if (elements.length === 0) {
    elements.push(new Paragraph({ children: [new TextRun({ text: "" })] }));
  }

  const wordDoc = new Document({
    styles: {
      default: {
        document: {
          run: { font: { name: ARABIC_FONT, cs: ARABIC_FONT }, size: 28 },
          paragraph: { spacing: { line: 360 } },
        },
      },
      paragraphStyles: [
        {
          id: "Normal",
          name: "Normal",
          quickFormat: true,
          run: { font: { name: ARABIC_FONT, cs: ARABIC_FONT }, size: 28 },
          paragraph: { spacing: { line: 360 } },
        },
      ],
    },
    sections: [
      {
        properties: {
          type: SectionType.CONTINUOUS,
          page: {
            margin: {
              top: convertInchesToTwip(1.0),
              bottom: convertInchesToTwip(1.0),
              left: convertInchesToTwip(1.25),
              right: convertInchesToTwip(1.25),
            },
          },
        },
        children: elements,
      },
    ],
  });

  const blob = await Packer.toBlob(wordDoc);
  saveAs(blob, `${fileName}.docx`);
}

// ─── HTML → Excel ─────────────────────────────────────────────────────────────

export function exportHtmlAsExcel(html: string, fileName: string): void {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, "text/html");
  const tables = Array.from(doc.querySelectorAll("table"));

  const wb = XLSX.utils.book_new();

  if (tables.length === 0) {
    const lines = (doc.body.textContent || "")
      .split("\n")
      .map((l) => l.trim())
      .filter((l) => l.length > 0);
    const aoa = lines.map((l) => [l]);
    const ws = XLSX.utils.aoa_to_sheet(aoa.length > 0 ? aoa : [[""]]);
    XLSX.utils.book_append_sheet(wb, ws, "المحتوى");
  } else {
    tables.forEach((table, idx) => {
      const ws = XLSX.utils.table_to_sheet(table, { raw: false });
      XLSX.utils.book_append_sheet(wb, ws, `جدول ${idx + 1}`);
    });
  }

  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  saveAs(new Blob([wbout], { type: "application/octet-stream" }), `${fileName}.xlsx`);
}
