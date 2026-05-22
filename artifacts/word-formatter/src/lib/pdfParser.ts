import * as pdfjsLib from "pdfjs-dist";

// Use unpkg CDN worker matching the installed version — avoids Vite bundling issues
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(pdfjsLib as any).GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjsLib.version}/build/pdf.worker.min.mjs`;

interface TextItem {
  str: string;
  transform: number[]; // [scaleX, skewX, skewY, scaleY, x, y]
  width: number;
  height: number;
  dir: string;
}

const PRES_FORMS_RE = /[\uFB50-\uFDFF\uFE70-\uFEFF]/;
const ARABIC_ANY_RE = /[\u0600-\u06FF]/;

function nfkcNormalize(text: string): string {
  if (!PRES_FORMS_RE.test(text) && !ARABIC_ANY_RE.test(text)) return text;
  return text.normalize("NFKC");
}

/**
 * Build a single line string from text items that have already been sorted
 * in reading order. Uses spatial gaps to determine word boundaries.
 * For RTL items sorted descending-X: prev has larger X, curr has smaller X.
 * For LTR items sorted ascending-X: curr has larger X.
 */
function buildLineText(items: TextItem[], isRtl: boolean): string {
  let line = "";
  for (let j = 0; j < items.length; j++) {
    const str = nfkcNormalize(items[j].str);
    if (!str) continue;
    if (j === 0) {
      line += str;
      continue;
    }
    const prev = items[j - 1];
    const curr = items[j];
    const prevX = prev.transform[4];
    const currX = curr.transform[4];
    const prevW = Math.abs(prev.width);

    let gap: number;
    if (isRtl) {
      // sorted descending: prevX > currX. gap = space between left edge of prev and right edge of curr
      gap = prevX - prevW - currX;
    } else {
      gap = currX - (prevX + prevW);
    }

    // Positive gap > 2 pt → word boundary
    line += (gap > 2 ? " " : "") + str;
  }
  return line.replace(/\s+/g, " ").trim();
}

export async function parsePdfToText(arrayBuffer: ArrayBuffer): Promise<string> {
  const loadingTask = pdfjsLib.getDocument({ data: new Uint8Array(arrayBuffer) });
  const pdf = await loadingTask.promise;

  const pageTexts: string[] = [];

  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    const page = await pdf.getPage(pageNum);
    const textContent = await page.getTextContent();

    // Group items by rounded Y coordinate (line grouping)
    const lineMap = new Map<number, TextItem[]>();
    for (const item of textContent.items) {
      if (!("str" in item)) continue;
      const textItem = item as TextItem;
      if (!textItem.str.trim() && textItem.width < 1) continue;

      const y = Math.round(textItem.transform[5] / 2) * 2;
      if (!lineMap.has(y)) lineMap.set(y, []);
      lineMap.get(y)!.push(textItem);
    }

    // Sort lines top-to-bottom (PDF Y increases upward → sort descending)
    const sortedYs = Array.from(lineMap.keys()).sort((a, b) => b - a);

    const lines: string[] = [];
    let prevY: number | null = null;

    for (const y of sortedYs) {
      const items = lineMap.get(y)!;

      // Detect if this line is predominantly RTL (Arabic)
      const rtlCount = items.filter((i) => i.dir === "rtl").length;
      const arabicCount = items.filter((i) =>
        ARABIC_ANY_RE.test(i.str) || PRES_FORMS_RE.test(i.str)
      ).length;
      const isRtl = rtlCount > items.length / 2 || arabicCount > items.length / 2;

      if (isRtl) {
        // RTL: sort right-to-left (descending X = reading order for Arabic)
        items.sort((a, b) => b.transform[4] - a.transform[4]);
      } else {
        // LTR: sort left-to-right (ascending X)
        items.sort((a, b) => a.transform[4] - b.transform[4]);
      }

      const lineText = buildLineText(items, isRtl);
      if (!lineText) continue;

      // Insert blank line when there's a large vertical gap (paragraph break)
      if (prevY !== null && Math.abs(prevY - y) > 20) {
        lines.push("");
      }

      lines.push(lineText);
      prevY = y;
    }

    const pageText = lines.join("\n").trim();
    if (pageText) pageTexts.push(pageText);
  }

  return pageTexts.join("\n\n");
}
