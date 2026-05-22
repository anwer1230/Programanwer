/**
 * docxFontPatcher.ts
 *
 * Opens the DOCX ZIP directly, patches ONLY the selected properties
 * (font names, font size, page margins, page borders) and re-packs it.
 *
 * Everything else — tables, images, alignment, colours, page-breaks,
 * heading levels, numbering — is preserved byte-for-byte.
 */
import JSZip from "jszip";
import { saveAs } from "file-saver";
import { BorderPreset } from "./borderPresets";

// 1 cm → twips  (1 inch = 2.54 cm = 1440 twips)
const CM_TO_TWIP = 1440 / 2.54;

export interface DocxMargins {
  top: number;    // cm
  bottom: number; // cm
  left: number;   // cm
  right: number;  // cm
}

export interface DocxPatchOptions {
  arabicFont: string;
  englishFont: string;
  fontSize: number | null;       // pt  — null = keep original
  margins: DocxMargins | null;   // null = keep original
  borderPreset: BorderPreset | null;
}

// ─── Main entry point ─────────────────────────────────────────────────────────

export async function patchDocxFonts(
  arrayBuffer: ArrayBuffer,
  outputFileName: string,
  options: DocxPatchOptions
): Promise<void> {
  const { arabicFont, englishFont, fontSize, margins, borderPreset } = options;

  const zip = await JSZip.loadAsync(arrayBuffer);

  // All XML files that may carry font / size declarations
  const fontTargets = Object.keys(zip.files).filter((name) =>
    /^word\/(document|styles|header\d*|footer\d*|numbering|endnotes|footnotes)\.xml$/.test(name)
  );

  for (const filePath of fontTargets) {
    const file = zip.file(filePath);
    if (!file) continue;

    let content = await file.async("string");

    // 1. Patch rFonts (explicit font names)
    content = patchRFonts(content, arabicFont, englishFont);

    // 2. Patch font size if requested
    if (fontSize !== null) {
      content = patchFontSize(content, fontSize);
    }

    // 3. Patch margins + borders only in document.xml
    if (filePath === "word/document.xml") {
      if (margins) {
        content = patchMargins(content, margins);
      }
      if (borderPreset && borderPreset.id !== "none") {
        content = patchPageBorders(content, borderPreset);
      }
    }

    zip.file(filePath, content);
  }

  const uint8 = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  const blob = new Blob([uint8], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
  saveAs(blob, `${outputFileName}.docx`);
}

// ─── Font name patching ───────────────────────────────────────────────────────

/**
 * Replaces every <w:rFonts .../> attribute set:
 *   w:ascii / w:hAnsi → englishFont  (Latin)
 *   w:cs              → arabicFont   (complex-script / Arabic)
 *   w:eastAsia        → arabicFont   (east-Asian slot, used by Arabic docs)
 *
 * Handles both self-closing (<w:rFonts ... />) and non-self-closing forms.
 * Also removes *Theme variants (w:asciiTheme etc.) so explicit names win.
 */
function patchRFonts(xml: string, arabicFont: string, englishFont: string): string {
  const patch = (attrs: string): string => replaceRFontsAttrs(attrs, arabicFont, englishFont);

  xml = xml.replace(/<w:rFonts([^/]*?)\/>/g, (_, a) => `<w:rFonts${patch(a)}/>`);
  xml = xml.replace(/<w:rFonts([^>]*?)>/g,   (_, a) => `<w:rFonts${patch(a)}>`);

  return xml;
}

function replaceRFontsAttrs(attrs: string, arabicFont: string, englishFont: string): string {
  // Remove theme-font overrides so explicit names always win
  attrs = attrs.replace(/\s+w:(ascii|hAnsi|eastAsia|cs)Theme="[^"]*"/g, "");

  // Set explicit names
  if (/w:cs="/.test(attrs)) {
    attrs = attrs.replace(/(w:cs=")[^"]*(")/g, `$1${arabicFont}$2`);
  } else {
    // No w:cs present → inject it (ensures Arabic runs get the new font)
    attrs = ` w:cs="${arabicFont}"` + attrs;
  }

  if (/w:ascii="/.test(attrs)) {
    attrs = attrs.replace(/(w:ascii=")[^"]*(")/g, `$1${englishFont}$2`);
  }
  if (/w:hAnsi="/.test(attrs)) {
    attrs = attrs.replace(/(w:hAnsi=")[^"]*(")/g, `$1${englishFont}$2`);
  }
  if (/w:eastAsia="/.test(attrs)) {
    attrs = attrs.replace(/(w:eastAsia=")[^"]*(")/g, `$1${arabicFont}$2`);
  }

  return attrs;
}

// ─── Font size patching ───────────────────────────────────────────────────────

/**
 * Replaces ALL <w:sz> and <w:szCs> values in the XML.
 * Values are stored in half-points (1 pt = 2 half-points).
 */
function patchFontSize(xml: string, fontSizePt: number): string {
  const hp = String(Math.round(fontSizePt * 2));
  xml = xml.replace(/(<w:sz\b[^>]*?w:val=")[^"]*(")/g,   `$1${hp}$2`);
  xml = xml.replace(/(<w:szCs\b[^>]*?w:val=")[^"]*(")/g, `$1${hp}$2`);
  return xml;
}

// ─── Margin patching ──────────────────────────────────────────────────────────

/**
 * Replaces w:top / w:bottom / w:left / w:right on every <w:pgMar> element.
 * Keeps header, footer, gutter values intact.
 */
function patchMargins(xml: string, margins: DocxMargins): string {
  const top    = Math.round(margins.top    * CM_TO_TWIP);
  const bottom = Math.round(margins.bottom * CM_TO_TWIP);
  const left   = Math.round(margins.left   * CM_TO_TWIP);
  const right  = Math.round(margins.right  * CM_TO_TWIP);

  xml = xml.replace(/<w:pgMar([^/]*?)\/>/g, (_, attrs) => {
    let a = attrs;
    a = a.replace(/(w:top=")[^"]*(")/g,    `$1${top}$2`);
    a = a.replace(/(w:bottom=")[^"]*(")/g, `$1${bottom}$2`);
    a = a.replace(/(w:left=")[^"]*(")/g,   `$1${left}$2`);
    a = a.replace(/(w:right=")[^"]*(")/g,  `$1${right}$2`);

    // If attributes were missing, inject them
    if (!/w:top="/.test(a))    a += ` w:top="${top}"`;
    if (!/w:bottom="/.test(a)) a += ` w:bottom="${bottom}"`;
    if (!/w:left="/.test(a))   a += ` w:left="${left}"`;
    if (!/w:right="/.test(a))  a += ` w:right="${right}"`;
    return `<w:pgMar${a}/>`;
  });

  return xml;
}

// ─── Page border patching ─────────────────────────────────────────────────────

function patchPageBorders(xml: string, preset: BorderPreset): string {
  const side = `w:val="${preset.style}" w:sz="${preset.size}" w:space="${preset.space}" w:color="${preset.color}"`;
  const pgBordersXml =
    `<w:pgBorders w:offsetFrom="page">` +
    `<w:top ${side}/>` +
    `<w:left ${side}/>` +
    `<w:bottom ${side}/>` +
    `<w:right ${side}/>` +
    `</w:pgBorders>`;

  if (/<w:pgBorders[\s>]/.test(xml)) {
    xml = xml.replace(/<w:pgBorders[\s\S]*?<\/w:pgBorders>/g, pgBordersXml);
  } else {
    xml = xml.replace(/<\/w:sectPr>/g, `${pgBordersXml}</w:sectPr>`);
  }

  return xml;
}
