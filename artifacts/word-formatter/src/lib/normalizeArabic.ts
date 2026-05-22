/**
 * Arabic Text Normalization
 *
 * Some Word/PDF files store Arabic text as "Presentation Forms" — pre-shaped
 * glyphs in the Unicode ranges FB50-FDFF and FE70-FEFF — in VISUAL left-to-right
 * order (each character is a separate token separated by spaces, stored in RTL
 * visual sequence stored as LTR bytes). This causes Word's bidi engine to render
 * the text from the wrong side.
 *
 * This module:
 *   1. Detects Presentation Form characters.
 *   2. Converts them to base Arabic Unicode via NFKC normalization.
 *   3. Reverses the visual glyph order back to logical Unicode order.
 */

const PRES_FORMS_RE = /[\uFB50-\uFDFF\uFE70-\uFEFF]/;
const ARABIC_CHAR_RE = /^[\u0600-\u06FF]$/;
const ARABIC_ANY_RE = /[\u0600-\u06FF]/;

/** Returns true if the text contains Arabic Presentation Form characters */
export function hasPresentationForms(text: string): boolean {
  return PRES_FORMS_RE.test(text);
}

/**
 * Fix a single line that was stored in visual (LTR) glyph order.
 *
 * Example:
 *   Input  : "ن ا م ع ة  ن ط ل س"   ← individual chars, reversed order
 *   Output : "سلطنة عمان"             ← logical Unicode RTL order
 *
 * Mixed lines (Arabic + numbers/English) are handled by grouping tokens,
 * reversing group order, and within each Arabic group reversing character order.
 */
function fixVisualLine(line: string): string {
  const trimmed = line.trim();
  if (!trimmed || !ARABIC_ANY_RE.test(trimmed)) return line;

  const tokens = trimmed.split(/\s+/);
  if (tokens.length < 2) return line;

  // Count single Arabic characters (individual glyph tokens)
  const arabicSingleCount = tokens.filter(
    (t) => t.length === 1 && ARABIC_CHAR_RE.test(t)
  ).length;

  // Not glyph mode — fewer than 35% single Arabic chars
  if (arabicSingleCount < tokens.length * 0.35) return line;

  // ── Glyph mode: group consecutive Arabic chars vs other tokens ──────────
  type Group = { arabic: boolean; tokens: string[] };
  const groups: Group[] = [];

  for (const token of tokens) {
    const isArabicSingle = ARABIC_CHAR_RE.test(token);
    const last = groups[groups.length - 1];
    if (last && last.arabic === isArabicSingle) {
      last.tokens.push(token);
    } else {
      groups.push({ arabic: isArabicSingle, tokens: [token] });
    }
  }

  // Reverse group order (visual → logical) and join Arabic chars without space
  const result = groups.reverse().map((g) => {
    if (g.arabic) {
      // Reverse char order within the group and join as a word (no spaces)
      return g.tokens.reverse().join("");
    }
    // Non-Arabic group: keep as-is
    return g.tokens.join(" ");
  });

  return result.join(" ");
}

/**
 * Normalize Arabic text:
 *  - If no Presentation Forms detected: just NFC normalize and return.
 *  - Otherwise: NFKC (converts glyphs to base chars) + visual-order fix.
 */
export function normalizeArabicText(text: string): string {
  if (!hasPresentationForms(text)) {
    return text.normalize("NFC");
  }

  // NFKC converts e.g. ﻦ (FE86) → ن, ﺔ (FE94) → ة, etc.
  const nfkc = text.normalize("NFKC");

  // Fix line-by-line visual order
  return nfkc.split("\n").map(fixVisualLine).join("\n");
}

/**
 * Normalize a single short string (cell text, heading, etc.)
 */
export function normalizeCell(text: string): string {
  if (!hasPresentationForms(text)) return text.normalize("NFC");
  return fixVisualLine(text.normalize("NFKC"));
}
