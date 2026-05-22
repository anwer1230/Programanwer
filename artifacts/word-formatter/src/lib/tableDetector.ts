export interface TableData {
  headers: string[];
  rows: string[][];
}

export interface ParsedBlock {
  type: "paragraph" | "heading1" | "heading2" | "table" | "empty" | "image";
  text?: string;
  table?: TableData;
  imageData?: string;
  imageAlt?: string;
  imageWidth?: number;
  imageHeight?: number;
  alignment?: "right" | "left" | "center" | "justify";
  isBold?: boolean;
}

// ─── Heading Detection ────────────────────────────────────────────────────────

function detectHeadingLevel(text: string): "heading1" | "heading2" | null {
  const t = text.trim();
  if (!t) return null;

  if (
    /^(الفصل\s|الباب\s|القسم\s|الملحق\s|المبحث\s|المطلب\s|الخاتمة|التوصيات|المراجع|الملخص|المقدمة|الاستنتاجات|النتائج)/.test(t) ||
    /^(أولاً|ثانياً|ثالثاً|رابعاً|خامساً|سادساً|سابعاً|ثامناً|تاسعاً|عاشراً)/.test(t)
  ) {
    return "heading1";
  }

  if (
    /^(Chapter|Section|Part|Appendix|Introduction|Conclusion|Summary|References|Abstract)\s/i.test(t)
  ) {
    return "heading1";
  }

  if (/^(\d+\.\d+[\s:.\-])/.test(t) || /^(\d+\.\d+\.\d+)/.test(t)) {
    return "heading2";
  }

  if (/^([أ-ي][.)]\s)/.test(t)) return "heading2";

  if (/^(\d+\s*[-–—:]\s*[^.،,]{1,80}$)/.test(t)) return "heading1";
  if (/^(\d+\.\s+[^.،,]{1,80}$)/.test(t)) return "heading1";

  return null;
}

// ─── Table Parsing ────────────────────────────────────────────────────────────

function parseTableFromPipes(lines: string[]): TableData | null {
  const dataLines = lines.filter(
    (l) => !/^\s*[|:]?[\s\-=:]+[|:]/.test(l) && l.trim().length > 0
  );
  if (dataLines.length < 2) return null;

  const rows = dataLines.map((line) => {
    const stripped = line.replace(/^\|/, "").replace(/\|$/, "");
    return stripped
      .split("|")
      .map((cell) => cell.trim())
      .filter((cell) => cell !== "---" && cell !== "===");
  });

  const colCount = rows[0].length;
  if (colCount < 2) return null;
  if (!rows.every((r) => r.length >= colCount - 1 && r.length <= colCount + 1))
    return null;

  return {
    headers: rows[0],
    rows: rows.slice(1).filter((r) => r.some((c) => c.length > 0)),
  };
}

function parseTableFromTabs(lines: string[]): TableData | null {
  if (lines.length < 2) return null;
  const rows = lines.map((l) => l.split("\t").map((c) => c.trim()));
  const colCount = rows[0].length;
  if (colCount < 2) return null;
  if (!rows.every((r) => r.length === colCount)) return null;
  return {
    headers: rows[0],
    rows: rows.slice(1).filter((r) => r.some((c) => c.length > 0)),
  };
}

function parseTableFromSpaces(lines: string[]): TableData | null {
  if (lines.length < 2) return null;

  const rows = lines.map((l) =>
    l
      .split(/\s{2,}/)
      .map((c) => c.trim())
      .filter((c) => c.length > 0)
  );
  const colCount = rows[0].length;
  if (colCount < 2) return null;

  if (!rows.every((r) => r.length >= colCount - 1)) return null;

  return {
    headers: rows[0],
    rows: rows.slice(1).filter((r) => r.some((c) => c.length > 0)),
  };
}

function classifyLine(line: string): "pipe" | "tab" | "space" | null {
  const t = line.trim();
  if (!t) return null;
  if (t.includes("|")) return "pipe";
  if (t.includes("\t")) return "tab";
  const tokens = t.split(/\s{2,}/).filter((s) => s.trim().length > 0);
  if (tokens.length >= 2) return "space";
  return null;
}

function collectTableGroup(
  lines: string[],
  start: number
): { lines: string[]; end: number; type: "pipe" | "tab" | "space" } | null {
  const firstClass = classifyLine(lines[start]);
  if (!firstClass) return null;

  const group: string[] = [];
  let j = start;

  while (j < lines.length) {
    const line = lines[j];
    const cls = classifyLine(line);

    if (!cls && line.trim() === "" && group.length > 0) {
      if (j + 1 < lines.length && classifyLine(lines[j + 1]) === firstClass) {
        j++;
        continue;
      }
      break;
    }

    if (cls && cls !== firstClass && group.length >= 2) break;
    if (!cls) break;

    group.push(line);
    j++;
  }

  if (group.length < 2) return null;
  return { lines: group, end: j, type: firstClass };
}

// ─── Main Parser ──────────────────────────────────────────────────────────────

export function parseText(rawText: string): ParsedBlock[] {
  const lines = rawText.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
  const blocks: ParsedBlock[] = [];

  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    const trimmed = line.trim();

    if (!trimmed) {
      blocks.push({ type: "empty" });
      i++;
      continue;
    }

    const group = collectTableGroup(lines, i);
    if (group) {
      let tableData: TableData | null = null;

      if (group.type === "pipe") {
        tableData = parseTableFromPipes(group.lines);
      } else if (group.type === "tab") {
        tableData = parseTableFromTabs(group.lines);
      } else {
        tableData = parseTableFromSpaces(group.lines);
        if (!tableData) tableData = parseTableFromPipes(group.lines);
      }

      if (tableData && tableData.rows.length > 0) {
        blocks.push({ type: "table", table: tableData });
        i = group.end;
        continue;
      }
    }

    const headingLevel = detectHeadingLevel(trimmed);
    if (headingLevel) {
      blocks.push({ type: headingLevel, text: trimmed });
      i++;
      continue;
    }

    blocks.push({ type: "paragraph", text: trimmed });
    i++;
  }

  return blocks;
}
