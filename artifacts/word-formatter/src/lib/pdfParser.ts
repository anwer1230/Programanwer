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

// Extract all text from a PDF, reconstructing paragraphs from layout
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

      // Round Y to nearest 2pt to group same-line items
      const y = Math.round(textItem.transform[5] / 2) * 2;
      if (!lineMap.has(y)) lineMap.set(y, []);
      lineMap.get(y)!.push(textItem);
    }

    // Sort lines top-to-bottom (PDF Y increases upward, so sort descending)
    const sortedYs = Array.from(lineMap.keys()).sort((a, b) => b - a);

    const lines: string[] = [];
    let prevY: number | null = null;

    for (const y of sortedYs) {
      const items = lineMap.get(y)!;

      // Sort items left-to-right by X coordinate for LTR reading
      // For Arabic RTL PDFs, items may already be in visual order
      items.sort((a, b) => a.transform[4] - b.transform[4]);

      const lineText = items.map((i) => i.str).join(" ").replace(/\s+/g, " ").trim();
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
