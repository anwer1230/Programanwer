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
  Footer,
  SectionType,
  convertInchesToTwip,
  BorderStyle,
  ShadingType,
  TableLayoutType,
  PageNumber,
  NumberFormat,
  PageBorderDisplay,
  PageBorderOffsetFrom,
  ImageRun,
} from "docx";
import { saveAs } from "file-saver";
import { ParsedBlock } from "./tableDetector";
import { BorderPreset } from "./borderPresets";

export const ARABIC_FONT = "Simplified Arabic";
export const ENGLISH_FONT = "Times New Roman";

export interface DocxRenderOptions {
  arabicFont?: string;
  englishFont?: string;
  fontSize?: number | null;
  margins?: { top: number; bottom: number; left: number; right: number } | null;
}

let _opts: DocxRenderOptions = {};
function _af() { return _opts.arabicFont ?? ARABIC_FONT; }
function _ef() { return _opts.englishFont ?? ENGLISH_FONT; }

export function isArabicDominant(text: string): boolean {
  const arabicChars = (
    text.match(/[\u0600-\u06FF\uFB50-\uFDFF\uFE70-\uFEFF]/g) || []
  ).length;
  const latinChars = (text.match(/[a-zA-Z]/g) || []).length;
  return arabicChars >= latinChars;
}

export interface TextSegment {
  text: string;
  arabic: boolean;
}

export function splitIntoSegments(text: string): TextSegment[] {
  const segments: TextSegment[] = [];
  const regex =
    /([\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF\s\d،؛؟.!,:\-()[\]"']+|[^\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]+)/g;
  let match: RegExpExecArray | null;
  while ((match = regex.exec(text)) !== null) {
    const seg = match[0];
    if (!seg) continue;
    const arabic = /[\u0600-\u06FF]/.test(seg);
    if (segments.length > 0 && segments[segments.length - 1].arabic === arabic) {
      segments[segments.length - 1].text += seg;
    } else {
      segments.push({ text: seg, arabic });
    }
  }
  return segments.filter((s) => s.text.trim().length > 0 || s.text.includes(" "));
}

export function makeTextRuns(
  text: string,
  options: { size?: number; bold?: boolean; color?: string; italic?: boolean } = {}
): TextRun[] {
  const { size = 14, bold = false, color, italic = false } = options;
  const segments = splitIntoSegments(text);
  if (segments.length === 0) return [new TextRun({ text: "" })];

  const finalSize = (_opts.fontSize ?? size);
  return segments.map(
    (seg) =>
      new TextRun({
        text: seg.text,
        font: seg.arabic
          ? { name: _af(), cs: _af(), eastAsia: _af() }
          : { name: _ef(), cs: _ef() },
        size: finalSize * 2,
        bold,
        italics: italic,
        color: color,
        rightToLeft: seg.arabic,
      })
  );
}

export function resolveAlignment(
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

export function makeParagraph(
  text: string,
  options: {
    size?: number;
    bold?: boolean;
    italic?: boolean;
    centered?: boolean;
    firstLineIndent?: boolean;
    alignment?: "right" | "left" | "center" | "justify";
    color?: string;
    lineSpacing?: number;
  } = {}
): Paragraph {
  const {
    size = 14,
    bold = false,
    italic = false,
    centered = false,
    firstLineIndent = false,
    alignment,
    color,
    lineSpacing = 360,
  } = options;
  const arabic = isArabicDominant(text);

  const applyIndent =
    firstLineIndent && text.trim().length >= 80 && alignment !== "center" && !centered;

  const docAlignment = centered
    ? AlignmentType.CENTER
    : resolveAlignment(alignment, arabic);

  return new Paragraph({
    children: makeTextRuns(text, { size, bold, color, italic }),
    bidirectional: arabic,
    alignment: docAlignment,
    spacing: { line: lineSpacing, lineRule: "auto" as never },
    indent: applyIndent ? { firstLine: convertInchesToTwip(0.5) } : undefined,
  });
}

// ─── Table cell ───────────────────────────────────────────────────────────────

export function makeTableCell(
  text: string,
  isHeader = false,
  width?: number
): TableCell {
  const arabic = isArabicDominant(text || " ");
  return new TableCell({
    children: [
      new Paragraph({
        children: makeTextRuns(text || "", { size: _opts.fontSize ?? 12, bold: isHeader }),
        bidirectional: arabic,
        alignment: arabic ? AlignmentType.RIGHT : AlignmentType.CENTER,
        spacing: { line: 276 },
      }),
    ],
    shading: isHeader
      ? { fill: "D6E8F8", type: ShadingType.CLEAR, color: "auto" }
      : undefined,
    margins: {
      top: convertInchesToTwip(0.06),
      bottom: convertInchesToTwip(0.06),
      left: convertInchesToTwip(0.1),
      right: convertInchesToTwip(0.1),
    },
    width: width ? { size: width, type: WidthType.DXA } : undefined,
  });
}

// ─── Smart column width calculation ──────────────────────────────────────────

/**
 * Guesses column widths based on content.
 * If the last column looks like page numbers (≤5 chars), it gets 15%.
 * Otherwise equal distribution.
 * Returns widths in twips (total ≈ 8640 for 6-inch content area).
 */
function calcColumnWidths(headers: string[], rows: string[][]): number[] | null {
  const colCount = headers.length;
  if (colCount === 1) return null;

  const TOTAL = 8640; // twips for ~6-inch content width

  if (colCount === 2) {
    const lastColMaxLen = Math.max(
      headers[1].length,
      ...rows.map((r) => (r[1] ?? "").length)
    );
    if (lastColMaxLen <= 6) {
      // Page number column: 15% narrow
      const narrow = Math.round(TOTAL * 0.15);
      return [TOTAL - narrow, narrow];
    }
    // 60/40 split
    return [Math.round(TOTAL * 0.6), TOTAL - Math.round(TOTAL * 0.6)];
  }

  if (colCount === 3) {
    const lastColMaxLen = Math.max(
      headers[colCount - 1].length,
      ...rows.map((r) => (r[colCount - 1] ?? "").length)
    );
    if (lastColMaxLen <= 6) {
      const narrow = Math.round(TOTAL * 0.12);
      const rest = TOTAL - narrow;
      return [Math.round(rest * 0.35), Math.round(rest * 0.65), narrow];
    }
  }

  // Equal distribution
  const w = Math.floor(TOTAL / colCount);
  const widths = Array(colCount).fill(w);
  widths[0] += TOTAL - w * colCount; // adjust first column for rounding
  return widths;
}

export function makeDocTable(headers: string[], rows: string[][]): Table {
  const colCount = headers.length;
  const colWidths = calcColumnWidths(headers, rows);
  const isArabic = isArabicDominant(headers.join(" ") || rows[0]?.join(" ") || "");

  const headerRow = new TableRow({
    children: headers.map((h, i) =>
      makeTableCell(h, true, colWidths ? colWidths[i] : undefined)
    ),
    tableHeader: true,
  });

  const dataRows = rows.map(
    (row) =>
      new TableRow({
        children: Array.from({ length: colCount }, (_, i) =>
          makeTableCell(row[i] ?? "", false, colWidths ? colWidths[i] : undefined)
        ),
      })
  );

  return new Table({
    layout: colWidths ? TableLayoutType.FIXED : TableLayoutType.AUTOFIT,
    width: { size: 100, type: WidthType.PERCENTAGE },
    visuallyRightToLeft: isArabic,
    rows: [headerRow, ...dataRows],
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: "2C6FAC" },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: "2C6FAC" },
      left: { style: BorderStyle.SINGLE, size: 4, color: "2C6FAC" },
      right: { style: BorderStyle.SINGLE, size: 4, color: "2C6FAC" },
      insideH: { style: BorderStyle.SINGLE, size: 2, color: "A0C4E8" },
      insideV: { style: BorderStyle.SINGLE, size: 2, color: "A0C4E8" },
    },
  });
}

// ─── Page setup ───────────────────────────────────────────────────────────────

function makePageMargins() {
  if (_opts.margins) {
    const cmToTwip = (cm: number) => Math.round(cm * 1440 / 2.54);
    return {
      top: cmToTwip(_opts.margins.top),
      bottom: cmToTwip(_opts.margins.bottom),
      left: cmToTwip(_opts.margins.left),
      right: cmToTwip(_opts.margins.right),
    };
  }
  return {
    top: convertInchesToTwip(1.0),
    bottom: convertInchesToTwip(1.0),
    left: convertInchesToTwip(1.25),
    right: convertInchesToTwip(1.25),
  };
}

function buildPageBorders(preset: BorderPreset | null) {
  if (!preset || preset.id === "none") return undefined;
  const side = {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    style: preset.style as any,
    size: preset.size,
    color: preset.color,
    space: preset.space,
  };
  return {
    pageBorders: {
      display: PageBorderDisplay.ALL_PAGES,
      offsetFrom: PageBorderOffsetFrom.PAGE,
    },
    pageBorderTop: side,
    pageBorderBottom: side,
    pageBorderLeft: side,
    pageBorderRight: side,
  };
}

function dataUriToUint8Array(dataUri: string): {
  data: Uint8Array;
  type: "png" | "jpg" | "gif" | "bmp";
} {
  const [header, base64] = dataUri.split(",");
  const mimeType = header.split(":")[1]?.split(";")[0] ?? "image/png";
  const binaryStr = atob(base64);
  const bytes = new Uint8Array(binaryStr.length);
  for (let i = 0; i < binaryStr.length; i++) {
    bytes[i] = binaryStr.charCodeAt(i);
  }
  let type: "png" | "jpg" | "gif" | "bmp" = "png";
  if (mimeType.includes("jpeg") || mimeType.includes("jpg")) type = "jpg";
  else if (mimeType.includes("gif")) type = "gif";
  else if (mimeType.includes("bmp")) type = "bmp";
  return { data: bytes, type };
}

// ─── Block → docx element ─────────────────────────────────────────────────────

export function buildBlockElements(blocks: ParsedBlock[]): (Paragraph | Table)[] {
  const elements: (Paragraph | Table)[] = [];

  for (const block of blocks) {
    if (block.type === "empty") {
      elements.push(new Paragraph({ children: [], spacing: { line: 240 } }));
    } else if (block.type === "heading1" && block.text) {
      elements.push(
        makeParagraph(block.text, { size: 16, bold: true, alignment: block.alignment })
      );
    } else if (block.type === "heading2" && block.text) {
      elements.push(
        makeParagraph(block.text, { size: 14, bold: true, alignment: block.alignment })
      );
    } else if (block.type === "paragraph" && block.text) {
      elements.push(
        makeParagraph(block.text, {
          size: 14,
          firstLineIndent: true,
          alignment: block.alignment,
        })
      );
    } else if (block.type === "table" && block.table) {
      elements.push(new Paragraph({ children: [], spacing: { line: 200 } }));
      elements.push(makeDocTable(block.table.headers, block.table.rows));
      elements.push(new Paragraph({ children: [], spacing: { line: 200 } }));
    } else if (block.type === "image" && block.imageData) {
      try {
        const { data, type } = dataUriToUint8Array(block.imageData);
        const width = block.imageWidth ?? 400;
        const height = block.imageHeight ?? 300;
        elements.push(
          new Paragraph({
            children: [
              new ImageRun({
                data,
                transformation: { width, height },
                type,
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { line: 240 },
          })
        );
      } catch {
        // skip invalid images
      }
    }
  }

  return elements;
}

// ─── Footer ───────────────────────────────────────────────────────────────────

function makeFooter(useArabic: boolean) {
  return new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            children: [PageNumber.CURRENT],
            font: useArabic
              ? { name: _af(), cs: _af() }
              : { name: _ef(), cs: _ef() },
            size: 22,
          }),
        ],
      }),
    ],
  });
}

function splitBlocksAtPage(blocks: ParsedBlock[], pageCount: number): number {
  const avgBlocksPerPage = 22;
  return Math.min(pageCount * avgBlocksPerPage, blocks.length);
}

// ─── Document styles ──────────────────────────────────────────────────────────

function makeDocumentStyles() {
  const sz = (_opts.fontSize ?? 14) * 2;
  return {
    default: {
      document: {
        run: { font: { name: _af(), cs: _af() }, size: sz },
        paragraph: { spacing: { line: 360 } },
      },
    },
    paragraphStyles: [
      {
        id: "Normal",
        name: "Normal",
        quickFormat: true,
        run: { font: { name: _af(), cs: _af() }, size: sz },
        paragraph: { spacing: { line: 360 } },
      },
    ],
  };
}

// ─── Main export ──────────────────────────────────────────────────────────────

export async function generateDocx(
  blocks: ParsedBlock[],
  fileName = "المستند_المنسّق",
  prelimPages = 0,
  borderPreset: BorderPreset | null = null,
  renderOpts: DocxRenderOptions = {}
): Promise<void> {
  _opts = renderOpts;
  const borders = buildPageBorders(borderPreset);

  let sections;

  if (prelimPages > 0) {
    const splitAt = splitBlocksAtPage(blocks, prelimPages);
    const prelimBlocks = blocks.slice(0, splitAt);
    const mainBlocks = blocks.slice(splitAt);

    const prelim = buildBlockElements(prelimBlocks);
    const main = buildBlockElements(mainBlocks);

    sections = [
      {
        properties: {
          type: SectionType.NEXT_PAGE,
          page: {
            margin: makePageMargins(),
            pageNumbers: { start: 1, formatType: NumberFormat.UPPER_ROMAN },
            borders,
          },
        },
        footers: { default: makeFooter(true) },
        children: prelim.length > 0 ? prelim : [new Paragraph({ children: [] })],
      },
      {
        properties: {
          type: SectionType.NEXT_PAGE,
          page: {
            margin: makePageMargins(),
            pageNumbers: { start: 1, formatType: NumberFormat.DECIMAL },
            borders,
          },
        },
        footers: { default: makeFooter(false) },
        children: main.length > 0 ? main : [new Paragraph({ children: [] })],
      },
    ];
  } else {
    const allElements = buildBlockElements(blocks);
    sections = [
      {
        properties: {
          type: SectionType.CONTINUOUS,
          page: {
            margin: makePageMargins(),
            pageNumbers: { start: 1, formatType: NumberFormat.DECIMAL },
            borders,
          },
        },
        footers: { default: makeFooter(false) },
        children:
          allElements.length > 0 ? allElements : [new Paragraph({ children: [] })],
      },
    ];
  }

  const doc = new Document({
    styles: makeDocumentStyles(),
    sections,
  });

  const blob = await Packer.toBlob(doc);
  saveAs(blob, `${fileName}.docx`);
}
