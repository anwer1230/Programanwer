/**
 * docxFontPatcher.ts
 *
 * Instead of parsing → reconstructing (which loses ALL formatting),
 * this module opens the DOCX ZIP directly, patches ONLY the font
 * references in the XML, and re-packs it.
 *
 * Result: identical document with only fonts (and optionally size /
 * page borders) changed. Every table, image, alignment, colour, page-
 * break, heading style — everything — is preserved exactly.
 */
import JSZip from "jszip";
import { saveAs } from "file-saver";
import { BorderPreset } from "./borderPresets";

export interface DocxPatchOptions {
  arabicFont: string;
  englishFont: string;
  fontSize: number | null;
  borderPreset: BorderPreset | null;
}

// ─── Main entry point ─────────────────────────────────────────────────────────

export async function patchDocxFonts(
  arrayBuffer: ArrayBuffer,
  outputFileName: string,
  options: DocxPatchOptions
): Promise<void> {
  const { arabicFont, englishFont, fontSize, borderPreset } = options;

  const zip = await JSZip.loadAsync(arrayBuffer);

  const xmlTargets = Object.keys(zip.files).filter((name) =>
    /^word\/(document|styles|header\d*|footer\d*|numbering|endnotes|footnotes)\.xml$/.test(name)
  );

  for (const filePath of xmlTargets) {
    const file = zip.file(filePath);
    if (!file) continue;

    let content = await file.async("string");
    content = patchFonts(content, arabicFont, englishFont, fontSize);

    if (borderPreset && borderPreset.id !== "none" && filePath === "word/document.xml") {
      content = patchPageBorders(content, borderPreset);
    }

    zip.file(filePath, content);
  }

  const blob = await zip.generateAsync({ type: "blob", compression: "DEFLATE" });
  saveAs(blob, `${outputFileName}.docx`);
}

// ─── Font patching ────────────────────────────────────────────────────────────

/**
 * Replace every <w:rFonts .../> attribute set:
 *   w:ascii  / w:hAnsi  → englishFont   (Latin glyphs)
 *   w:cs                → arabicFont    (complex-script / Arabic)
 *   w:eastAsia          → arabicFont    (east-Asian slot used by Arabic docs)
 *
 * Optionally replaces w:sz / w:szCs (half-points) across all runs.
 */
function patchFonts(
  xml: string,
  arabicFont: string,
  englishFont: string,
  fontSize: number | null
): string {
  // Self-closing  <w:rFonts ... />
  xml = xml.replace(/<w:rFonts([^/]*?)\/>/g, (_, attrs) => {
    return `<w:rFonts${replaceRFontsAttrs(attrs, arabicFont, englishFont)}/>`;
  });

  // Non-self-closing  <w:rFonts ...>  (rare but possible)
  xml = xml.replace(/<w:rFonts([^>]*?)>/g, (_, attrs) => {
    return `<w:rFonts${replaceRFontsAttrs(attrs, arabicFont, englishFont)}>`;
  });

  // Font size override
  if (fontSize !== null) {
    const hp = String(Math.round(fontSize * 2)); // half-points
    xml = xml.replace(/(<w:sz\s+w:val=")[^"]*(")/g, `$1${hp}$2`);
    xml = xml.replace(/(<w:szCs\s+w:val=")[^"]*(")/g, `$1${hp}$2`);
  }

  return xml;
}

function replaceRFontsAttrs(
  attrs: string,
  arabicFont: string,
  englishFont: string
): string {
  // w:cs  → Arabic/complex-script font
  attrs = attrs.replace(/(w:cs=")[^"]*(")/g, `$1${arabicFont}$2`);
  // w:ascii → Latin font
  attrs = attrs.replace(/(w:ascii=")[^"]*(")/g, `$1${englishFont}$2`);
  // w:hAnsi → Latin font (high-ANSI)
  attrs = attrs.replace(/(w:hAnsi=")[^"]*(")/g, `$1${englishFont}$2`);
  // w:eastAsia → Arabic font (Arabic docs often land here)
  attrs = attrs.replace(/(w:eastAsia=")[^"]*(")/g, `$1${arabicFont}$2`);
  return attrs;
}

// ─── Page border patching ─────────────────────────────────────────────────────

function patchPageBorders(xml: string, preset: BorderPreset): string {
  const side = `w:val="${preset.style}" w:sz="${preset.size}" w:space="${preset.space}" w:color="${preset.color}"`;
  const pgBordersXml = [
    `<w:pgBorders w:offsetFrom="page">`,
    `  <w:top ${side}/>`,
    `  <w:left ${side}/>`,
    `  <w:bottom ${side}/>`,
    `  <w:right ${side}/>`,
    `</w:pgBorders>`,
  ].join("");

  if (/<w:pgBorders[\s>]/.test(xml)) {
    // Replace existing border block
    xml = xml.replace(/<w:pgBorders[\s\S]*?<\/w:pgBorders>/g, pgBordersXml);
  } else {
    // Inject before every </w:sectPr>
    xml = xml.replace(/<\/w:sectPr>/g, `${pgBordersXml}</w:sectPr>`);
  }

  return xml;
}
