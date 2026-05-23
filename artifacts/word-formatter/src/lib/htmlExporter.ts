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

const ARABIC_FONT = "Simplified Arabic";
const ENGLISH_FONT = "Times New Roman";

function isArabicDominant(text: string): boolean {
  const ar = (text.match(/[\u0600-\u06FF\uFB50-\uFDFF\uFE70-\uFEFF]/g) || []).length;
  const en = (text.match(/[a-zA-Z]/g) || []).length;
  return ar >= en;
}

// ─── Computed-style helpers ───────────────────────────────────────────────────

function rgbToHex(rgb: string): string | undefined {
  if (!rgb || rgb === "transparent" || rgb === "rgba(0, 0, 0, 0)" || rgb === "") return undefined;
  const m = rgb.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
  if (!m) return undefined;
  const hex = [m[1], m[2], m[3]]
    .map((n) => parseInt(n).toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase();
  return hex;
}

function pxToPt(px: number): number {
  return Math.round(px * 0.75);
}

interface ComputedStyle2 {
  color?: string;
  bgColor?: string;
  fontSize: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  textAlign: string;
  hasBorderRight: boolean;
  borderRightColor?: string;
}

function readComputedStyle(el: Element): ComputedStyle2 {
  const cs = window.getComputedStyle(el);
  const fontSize = pxToPt(parseFloat(cs.fontSize) || 16);
  const fontWeight = cs.fontWeight;
  const bold =
    fontWeight === "bold" || fontWeight === "bolder" || parseInt(fontWeight) >= 600;
  const italic = cs.fontStyle === "italic" || cs.fontStyle === "oblique";
  const textDeco = cs.textDecoration || "";
  const underline = textDeco.includes("underline");
  const strike = textDeco.includes("line-through");
  const textAlign = cs.textAlign;
  const rawBg = cs.backgroundColor;
  const bgColor = rawBg && rawBg !== "rgba(0, 0, 0, 0)" ? rgbToHex(rawBg) : undefined;
  const color = rgbToHex(cs.color);
  const bRW = parseFloat(cs.borderRightWidth) || 0;
  const hasBorderRight = bRW > 0 && cs.borderRightStyle !== "none";
  const borderRightColor = hasBorderRight ? rgbToHex(cs.borderRightColor) : undefined;

  return {
    color: color === "000000" ? undefined : color,
    bgColor,
    fontSize: Math.max(8, fontSize),
    bold,
    italic,
    underline,
    strike,
    textAlign,
    hasBorderRight,
    borderRightColor,
  };
}

// ─── Inject HTML+CSS into live DOM for rendering ──────────────────────────────

interface RenderContext {
  wrapper: HTMLDivElement;
  cleanup: () => void;
}

async function mountHtmlForRender(html: string): Promise<RenderContext> {
  const isFullDoc = /<html[\s>]/i.test(html);
  const srcDoc = isFullDoc
    ? html
    : `<!DOCTYPE html><html><head></head><body>${html}</body></html>`;

  const parser = new DOMParser();
  const parsed = parser.parseFromString(srcDoc, "text/html");

  const injectedEls: Element[] = [];

  for (const styleEl of Array.from(parsed.querySelectorAll("style"))) {
    const s = document.createElement("style");
    s.textContent = styleEl.textContent;
    document.head.appendChild(s);
    injectedEls.push(s);
  }

  const wrapper = document.createElement("div");
  wrapper.style.cssText =
    "position:fixed;left:-9999px;top:0;width:794px;visibility:hidden;z-index:-9999;direction:rtl";
  wrapper.innerHTML = parsed.body.innerHTML;
  document.body.appendChild(wrapper);

  await new Promise<void>((resolve) => requestAnimationFrame(() => setTimeout(resolve, 60)));

  const cleanup = () => {
    if (document.body.contains(wrapper)) document.body.removeChild(wrapper);
    for (const el of injectedEls) {
      if (el.parentNode) el.parentNode.removeChild(el);
    }
  };

  return { wrapper, cleanup };
}

// ─── Run collector ────────────────────────────────────────────────────────────

interface RunOpts {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  size?: number;
}

function makeRun(text: string, opts: RunOpts = {}): TextRun | null {
  const cleaned = text.replace(/\u00a0/g, " ");
  if (!cleaned) return null;
  const arabic = isArabicDominant(cleaned);
  const pt = opts.size ?? 14;
  return new TextRun({
    text: cleaned,
    font: arabic
      ? { name: ARABIC_FONT, cs: ARABIC_FONT, eastAsia: ARABIC_FONT }
      : { name: ENGLISH_FONT, cs: ENGLISH_FONT },
    size: pt * 2,
    bold: opts.bold ?? false,
    italics: opts.italic ?? false,
    underline: opts.underline ? { type: "single" as never } : undefined,
    strike: opts.strike ?? false,
    color: opts.color,
    rightToLeft: arabic,
  });
}

function collectRunsFromNode(node: Node, inherited: RunOpts, runs: TextRun[]): void {
  if (node.nodeType === Node.TEXT_NODE) {
    const text = node.textContent ?? "";
    if (!text.trim()) {
      if (text.includes(" ") || text.includes("\n")) {
        const run = makeRun(" ", inherited);
        if (run) runs.push(run);
      }
      return;
    }
    const run = makeRun(text, inherited);
    if (run) runs.push(run);
    return;
  }
  if (node.nodeType !== Node.ELEMENT_NODE) return;

  const el = node as Element;
  const tag = el.tagName.toUpperCase();

  if (tag === "BR") {
    runs.push(new TextRun({ break: 1 }));
    return;
  }
  if (["STYLE", "SCRIPT", "NOSCRIPT"].includes(tag)) return;

  const BLOCK_TAGS = new Set([
    "DIV","P","H1","H2","H3","H4","H5","H6","UL","OL","LI",
    "TABLE","THEAD","TBODY","TFOOT","TR","TD","TH","BLOCKQUOTE","PRE","HR",
    "SECTION","ARTICLE","HEADER","FOOTER","ASIDE","MAIN","FIGURE","NAV",
  ]);

  const cs = readComputedStyle(el);
  const opts: RunOpts = {
    bold: cs.bold || inherited.bold,
    italic: cs.italic || inherited.italic,
    underline: cs.underline || inherited.underline,
    strike: cs.strike || inherited.strike,
    color: cs.color ?? inherited.color,
    size: cs.fontSize ?? inherited.size,
  };

  if (BLOCK_TAGS.has(tag)) {
    for (const child of Array.from(el.childNodes)) collectRunsFromNode(child, opts, runs);
    return;
  }

  for (const child of Array.from(el.childNodes)) collectRunsFromNode(child, opts, runs);
}

// ─── Paragraph builder using computed styles ──────────────────────────────────

function buildComputedParagraph(
  el: Element,
  overrideStyle?: Partial<RunOpts>
): Paragraph {
  const cs = readComputedStyle(el);
  const base: RunOpts = {
    bold: cs.bold,
    italic: cs.italic,
    underline: cs.underline,
    strike: cs.strike,
    color: cs.color,
    size: cs.fontSize,
    ...overrideStyle,
  };

  const runs: TextRun[] = [];
  for (const child of Array.from(el.childNodes)) collectRunsFromNode(child, base, runs);
  if (runs.length === 0) runs.push(new TextRun({ text: "" }));

  const text = el.textContent ?? "";
  const arabic = isArabicDominant(text);

  let alignment = AlignmentType.RIGHT;
  if (cs.textAlign === "center") alignment = AlignmentType.CENTER;
  else if (cs.textAlign === "left") alignment = AlignmentType.LEFT;
  else if (cs.textAlign === "justify" || cs.textAlign === "justify-all")
    alignment = arabic ? AlignmentType.THAI_DISTRIBUTE : AlignmentType.DISTRIBUTE;
  else if (!arabic) alignment = AlignmentType.LEFT;

  const spacing = Math.max(240, Math.min(480, cs.fontSize * 30));

  const paragraphBorder = cs.hasBorderRight
    ? {
        right: {
          style: BorderStyle.THICK,
          size: 12,
          color: cs.borderRightColor ?? "F6B352",
          space: 8,
        },
      }
    : undefined;

  const shading =
    cs.bgColor && cs.bgColor !== "FFFFFF"
      ? { fill: cs.bgColor, type: ShadingType.CLEAR, color: "auto" }
      : undefined;

  return new Paragraph({
    children: runs,
    bidirectional: arabic,
    alignment,
    spacing: { line: spacing, lineRule: "auto" as never },
    border: paragraphBorder,
    shading,
  });
}

// ─── Table from computed DOM ──────────────────────────────────────────────────

function buildComputedTable(tableEl: Element): Table | null {
  const rows = Array.from(tableEl.querySelectorAll("tr"));
  if (rows.length === 0) return null;

  const headerText = rows[0]
    ? Array.from(rows[0].querySelectorAll("td, th"))
        .map((c) => c.textContent ?? "")
        .join(" ")
    : "";
  const isArabic = isArabicDominant(headerText);
  const tableCs = readComputedStyle(tableEl);
  const hasBorder = tableCs.hasBorderRight || tableEl.getAttribute("border");

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

  const docRows = rows.map((tr, ri) => {
    const cells = Array.from(tr.querySelectorAll("td, th"));
    const isHeaderRow = ri === 0 && tr.querySelectorAll("th").length > 0;

    const docCells = cells.map((td) => {
      const isHeader = isHeaderRow || td.tagName.toUpperCase() === "TH";
      const cellCs = readComputedStyle(td);
      const cellRuns: TextRun[] = [];
      for (const child of Array.from(td.childNodes)) {
        collectRunsFromNode(
          child,
          { bold: isHeader, size: cellCs.fontSize ?? 12, color: cellCs.color },
          cellRuns
        );
      }
      if (cellRuns.length === 0) cellRuns.push(new TextRun({ text: "" }));

      const cellText = td.textContent ?? "";
      const cellArabic = isArabicDominant(cellText);
      let cellAlign = cellArabic ? AlignmentType.RIGHT : AlignmentType.LEFT;
      if (cellCs.textAlign === "center") cellAlign = AlignmentType.CENTER;

      const cellBg = cellCs.bgColor && cellCs.bgColor !== "FFFFFF" ? cellCs.bgColor : undefined;
      const cellShading = cellBg
        ? { fill: cellBg, type: ShadingType.CLEAR, color: "auto" }
        : isHeader
        ? { fill: "D6E8F8", type: ShadingType.CLEAR, color: "auto" }
        : undefined;

      return new TableCell({
        children: [
          new Paragraph({
            children: cellRuns,
            bidirectional: cellArabic,
            alignment: cellAlign,
            spacing: { line: 276 },
          }),
        ],
        shading: cellShading,
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
      tableHeader: isHeaderRow,
    });
  });

  return new Table({
    layout: TableLayoutType.AUTOFIT,
    width: { size: 100, type: WidthType.PERCENTAGE },
    visuallyRightToLeft: isArabic,
    rows: docRows,
    borders,
  });
}

// ─── Grid/flex container → 2-column DOCX table ───────────────────────────────

function buildGridAsTable(el: Element): Table | null {
  const children = Array.from(el.children).filter((c) => {
    const cs = window.getComputedStyle(c);
    return cs.display !== "none" && c.tagName.toUpperCase() !== "STYLE";
  });
  if (children.length < 2) return null;

  const rows: TableRow[] = [];
  for (let i = 0; i < children.length; i += 2) {
    const left = children[i];
    const right = children[i + 1];

    const makeCell = (cellEl: Element | undefined): TableCell => {
      if (!cellEl) {
        return new TableCell({
          children: [new Paragraph({ children: [] })],
          width: { size: 50, type: WidthType.PERCENTAGE },
        });
      }
      const cs = readComputedStyle(cellEl);
      const shading = cs.bgColor && cs.bgColor !== "FFFFFF"
        ? { fill: cs.bgColor, type: ShadingType.CLEAR, color: "auto" }
        : undefined;

      const cellElements: (Paragraph | Table)[] = [];
      collectElementsFromNode(cellEl, cellElements);

      return new TableCell({
        children:
          cellElements.length > 0
            ? cellElements
            : [new Paragraph({ children: [] })],
        shading,
        width: { size: 50, type: WidthType.PERCENTAGE },
        margins: {
          top: convertInchesToTwip(0.1),
          bottom: convertInchesToTwip(0.1),
          left: convertInchesToTwip(0.1),
          right: convertInchesToTwip(0.1),
        },
      });
    };

    rows.push(
      new TableRow({
        children: [makeCell(left), makeCell(right)],
      })
    );
  }

  if (rows.length === 0) return null;

  return new Table({
    layout: TableLayoutType.FIXED,
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows,
    borders: {
      top: { style: BorderStyle.NONE },
      bottom: { style: BorderStyle.NONE },
      left: { style: BorderStyle.NONE },
      right: { style: BorderStyle.NONE },
      insideH: { style: BorderStyle.NONE },
      insideV: { style: BorderStyle.NONE },
    },
  });
}

// ─── Main node processor ──────────────────────────────────────────────────────

const SKIP_TAGS = new Set(["STYLE", "SCRIPT", "NOSCRIPT", "HEAD", "META", "LINK", "TITLE"]);
const BLOCK_CONTAINER_TAGS = new Set([
  "DIV","SECTION","ARTICLE","MAIN","HEADER","FOOTER","ASIDE","NAV",
  "FIGURE","FIGCAPTION","BODY","FORM","FIELDSET","DETAILS","SUMMARY",
]);

function isGridContainer(el: Element): boolean {
  const cs = window.getComputedStyle(el);
  const display = cs.display;
  return (
    (display === "grid" || display === "inline-grid" ||
     display === "flex" || display === "inline-flex") &&
    el.children.length >= 2 &&
    Array.from(el.children).every((c) => {
      const ccs = window.getComputedStyle(c);
      return ccs.display !== "none";
    })
  );
}

function collectElementsFromNode(
  el: Element,
  elements: (Paragraph | Table)[]
): void {
  const tag = el.tagName.toUpperCase();

  if (SKIP_TAGS.has(tag)) return;

  switch (tag) {
    case "H1":
    case "H2":
    case "H3":
    case "H4":
    case "H5":
    case "H6": {
      const para = buildComputedParagraph(el);
      elements.push(para);
      elements.push(new Paragraph({ children: [], spacing: { line: 160 } }));
      break;
    }

    case "P": {
      const text = (el.textContent ?? "").trim();
      if (!text) {
        elements.push(new Paragraph({ children: [], spacing: { line: 240 } }));
      } else {
        elements.push(buildComputedParagraph(el));
      }
      break;
    }

    case "BLOCKQUOTE":
      elements.push(buildComputedParagraph(el, { italic: true }));
      break;

    case "TABLE": {
      const table = buildComputedTable(el);
      if (table) {
        elements.push(new Paragraph({ children: [], spacing: { line: 160 } }));
        elements.push(table);
        elements.push(new Paragraph({ children: [], spacing: { line: 160 } }));
      }
      break;
    }

    case "UL":
    case "OL": {
      const lis = Array.from(el.querySelectorAll(":scope > li"));
      lis.forEach((li, i) => {
        const liCs = readComputedStyle(li);
        const arabic = isArabicDominant(li.textContent ?? "");
        const marker = tag === "OL" ? `${i + 1}. ` : "• ";
        const runs: TextRun[] = [
          new TextRun({
            text: marker,
            font: arabic ? { name: ARABIC_FONT, cs: ARABIC_FONT } : { name: ENGLISH_FONT },
            size: (liCs.fontSize ?? 14) * 2,
            rightToLeft: arabic,
            color: liCs.color,
          }),
        ];
        for (const child of Array.from(li.childNodes)) {
          collectRunsFromNode(child, { size: liCs.fontSize ?? 14, color: liCs.color, bold: liCs.bold }, runs);
        }
        const shading = liCs.bgColor && liCs.bgColor !== "FFFFFF"
          ? { fill: liCs.bgColor, type: ShadingType.CLEAR, color: "auto" }
          : undefined;
        elements.push(
          new Paragraph({
            children: runs,
            bidirectional: arabic,
            alignment: arabic ? AlignmentType.RIGHT : AlignmentType.LEFT,
            spacing: { line: 300 },
            indent: { left: convertInchesToTwip(0.2) },
            shading,
          })
        );
      });
      break;
    }

    case "HR":
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

    default: {
      if (BLOCK_CONTAINER_TAGS.has(tag)) {
        // Check if it's a grid/flex container → convert to table
        if (isGridContainer(el)) {
          const gridTable = buildGridAsTable(el);
          if (gridTable) {
            elements.push(new Paragraph({ children: [], spacing: { line: 120 } }));
            elements.push(gridTable);
            elements.push(new Paragraph({ children: [], spacing: { line: 120 } }));
            break;
          }
        }

        // Check if it has text directly (no block children) → paragraph
        const blockChildren = Array.from(el.children).filter((c) => {
          const t = c.tagName.toUpperCase();
          return !SKIP_TAGS.has(t);
        });

        if (blockChildren.length === 0 && (el.textContent ?? "").trim()) {
          elements.push(buildComputedParagraph(el));
          break;
        }

        for (const child of Array.from(el.children)) {
          collectElementsFromNode(child, elements);
        }
        break;
      }

      // Inline/unknown with text → paragraph
      const text = (el.textContent ?? "").trim();
      if (text) {
        elements.push(buildComputedParagraph(el));
      }
    }
  }
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

  const isFullDoc = /<html[\s>]/i.test(html);
  const srcDoc = isFullDoc
    ? html
    : `<!DOCTYPE html><html><head></head><body>${html}</body></html>`;
  const parser = new DOMParser();
  const parsed = parser.parseFromString(srcDoc, "text/html");

  const injectedStyles: HTMLStyleElement[] = [];
  for (const styleEl of Array.from(parsed.querySelectorAll("style"))) {
    const s = document.createElement("style");
    s.textContent = styleEl.textContent;
    document.head.appendChild(s);
    injectedStyles.push(s);
  }

  container.innerHTML = parsed.body.innerHTML;
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
    for (const s of injectedStyles) {
      if (s.parentNode) s.parentNode.removeChild(s);
    }
  }
}

// ─── HTML → Word (.docx) ─────────────────────────────────────────────────────

export async function exportHtmlAsDocx(html: string, fileName: string): Promise<void> {
  const { wrapper, cleanup } = await mountHtmlForRender(html);

  try {
    const elements: (Paragraph | Table)[] = [];
    for (const child of Array.from(wrapper.children)) {
      collectElementsFromNode(child, elements);
    }

    while (
      elements.length > 0 &&
      elements[elements.length - 1] instanceof Paragraph
    ) {
      const last = elements[elements.length - 1] as Paragraph;
      const kids = last.options?.children ?? [];
      if (kids.length === 0) elements.pop();
      else break;
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
  } finally {
    cleanup();
  }
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
