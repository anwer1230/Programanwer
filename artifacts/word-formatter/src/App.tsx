import { useState, useRef } from "react";
import { parseText, ParsedBlock } from "./lib/tableDetector";
import { generateDocx } from "./lib/docGenerator";
import { BORDER_PRESETS, BorderPreset } from "./lib/borderPresets";
import { parseDocxToBlocks } from "./lib/docxParser";
import { parsePdfToText } from "./lib/pdfParser";
import { exportHtmlAsPdf, exportHtmlAsDocx, exportHtmlAsExcel } from "./lib/htmlExporter";
import {
  FileText,
  Download,
  RefreshCw,
  CheckCircle,
  AlertCircle,
  Table,
  AlignRight,
  Heading,
  Upload,
  Code,
  Image as ImageIcon,
} from "lucide-react";

type Status = "idle" | "processing" | "done" | "error";
type UploadStatus = "idle" | "reading" | "ready" | "error";
type FileKind = "docx" | "pdf" | null;
type ExportFormat = "pdf" | "docx" | "excel";
type ActiveTab = "formatter" | "html";

// ─── Block Preview ─────────────────────────────────────────────────────────────
function BlockPreview({ block, index }: { block: ParsedBlock; index: number }) {
  if (block.type === "empty") return <div className="h-2" key={index} />;

  if (block.type === "image" && block.imageData) {
    return (
      <div className="mb-3 rounded-lg border border-purple-200 bg-purple-50 overflow-hidden">
        <div className="flex items-center gap-1 px-2 py-1 bg-purple-100 text-purple-700 text-xs font-medium">
          <ImageIcon size={12} />
          <span>صورة / رسم بياني</span>
        </div>
        <div className="p-2 flex justify-center">
          <img
            src={block.imageData}
            alt={block.imageAlt || "صورة"}
            className="max-w-full max-h-64 object-contain rounded"
          />
        </div>
      </div>
    );
  }

  if (block.type === "table" && block.table) {
    return (
      <div className="mb-3 overflow-x-auto rounded border border-blue-200 bg-blue-50">
        <div className="flex items-center gap-1 px-2 py-1 bg-blue-100 text-blue-700 text-xs font-medium">
          <Table size={12} />
          <span>جدول مُكتشَف</span>
          <span className="mr-auto opacity-60">
            {block.table.headers.length} أعمدة × {block.table.rows.length} صف
          </span>
        </div>
        <table className="w-full text-sm border-collapse">
          <thead>
            <tr>
              {block.table.headers.map((h, i) => (
                <th key={i} className="border border-blue-300 bg-blue-200 px-2 py-1 text-right font-semibold text-blue-900">
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {block.table.rows.map((row, ri) => (
              <tr key={ri} className={ri % 2 === 0 ? "bg-white" : "bg-blue-50/50"}>
                {row.map((cell, ci) => (
                  <td key={ci} className="border border-blue-200 px-2 py-1 text-right text-gray-800">
                    {cell}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  }

  if (block.type === "heading1") {
    return (
      <div className="flex items-start gap-2 mb-2">
        <div className="mt-1 text-emerald-600 shrink-0"><Heading size={14} /></div>
        <p className="font-bold text-base text-gray-900">{block.text}</p>
      </div>
    );
  }

  if (block.type === "heading2") {
    return (
      <div className="flex items-start gap-2 mb-2">
        <div className="mt-1 text-emerald-400 shrink-0"><Heading size={12} /></div>
        <p className="font-semibold text-sm text-gray-800">{block.text}</p>
      </div>
    );
  }

  return (
    <div className="flex items-start gap-2 mb-2">
      <div className="mt-1 text-gray-400 shrink-0"><AlignRight size={12} /></div>
      <p className="text-sm text-gray-700 leading-relaxed">{block.text}</p>
    </div>
  );
}

function StatsBar({ blocks }: { blocks: ParsedBlock[] }) {
  const tables = blocks.filter((b) => b.type === "table").length;
  const images = blocks.filter((b) => b.type === "image").length;
  const h1 = blocks.filter((b) => b.type === "heading1").length;
  const h2 = blocks.filter((b) => b.type === "heading2").length;
  const paras = blocks.filter((b) => b.type === "paragraph").length;
  return (
    <div className="flex flex-wrap gap-3 text-xs text-muted-foreground px-1">
      <span className="flex items-center gap-1"><span className="w-2 h-2 rounded-full bg-emerald-500 inline-block" />{h1} عنوان رئيسي</span>
      <span className="flex items-center gap-1"><span className="w-2 h-2 rounded-full bg-emerald-300 inline-block" />{h2} عنوان فرعي</span>
      <span className="flex items-center gap-1"><span className="w-2 h-2 rounded-full bg-blue-500 inline-block" />{tables} جدول</span>
      {images > 0 && <span className="flex items-center gap-1"><span className="w-2 h-2 rounded-full bg-purple-500 inline-block" />{images} صورة</span>}
      <span className="flex items-center gap-1"><span className="w-2 h-2 rounded-full bg-gray-400 inline-block" />{paras} فقرة</span>
    </div>
  );
}

// ─── Border Card ───────────────────────────────────────────────────────────────
function BorderPreview({ preset, selected, onClick }: {
  preset: BorderPreset; selected: boolean; onClick: () => void;
}) {
  const isNone = preset.id === "none";
  return (
    <button
      onClick={onClick}
      className={`relative flex flex-col items-center gap-2 p-2 rounded-xl border-2 transition-all cursor-pointer text-center
        ${selected ? "border-primary bg-primary/8 shadow-md" : "border-border bg-white hover:border-primary/40 hover:bg-muted/30"}`}
    >
      <div
        className="w-12 h-16 rounded bg-white shadow-sm relative flex items-center justify-center overflow-hidden"
        style={{ border: isNone ? "1px dashed #d1d5db" : preset.cssPreview }}
      >
        <div className="flex flex-col gap-1 w-7">
          <div className="h-0.5 bg-gray-200 rounded w-full" />
          <div className="h-0.5 bg-gray-200 rounded w-5/6" />
          <div className="h-0.5 bg-gray-200 rounded w-full" />
          <div className="h-0.5 bg-gray-200 rounded w-4/6" />
          <div className="h-0.5 bg-gray-200 rounded w-full" />
        </div>
        {!isNone && (
          <div className="absolute -top-1 -right-1 text-xs leading-none">{preset.emoji}</div>
        )}
      </div>
      <div>
        <p className="text-xs font-semibold text-foreground leading-tight">{preset.label}</p>
        <p className="text-[10px] text-muted-foreground leading-tight mt-0.5">{preset.description}</p>
      </div>
      {selected && (
        <div className="absolute top-1.5 left-1.5 w-4 h-4 rounded-full bg-primary flex items-center justify-center">
          <CheckCircle size={10} className="text-white" />
        </div>
      )}
    </button>
  );
}

// ─── File type badge ───────────────────────────────────────────────────────────
function FileTypeBadge({ kind }: { kind: FileKind }) {
  if (!kind) return null;
  return (
    <span className={`inline-flex items-center gap-1 text-[10px] font-bold px-2 py-0.5 rounded-full uppercase tracking-wide
      ${kind === "pdf" ? "bg-red-100 text-red-700" : "bg-blue-100 text-blue-700"}`}>
      {kind === "pdf" ? "📄 PDF" : "📝 DOCX"}
    </span>
  );
}

// ─── HTML Exporter Tab ─────────────────────────────────────────────────────────
function HtmlExporterSection() {
  const [htmlCode, setHtmlCode] = useState("");
  const [exportFormat, setExportFormat] = useState<ExportFormat>("pdf");
  const [exporting, setExporting] = useState(false);
  const [exportStatus, setExportStatus] = useState<"idle" | "done" | "error">("idle");
  const [exportError, setExportError] = useState("");
  const [htmlFileName, setHtmlFileName] = useState("المستند_المُصدَّر");

  const handleExport = async () => {
    if (!htmlCode.trim()) return;
    setExporting(true);
    setExportStatus("idle");
    setExportError("");
    try {
      if (exportFormat === "pdf") {
        await exportHtmlAsPdf(htmlCode, htmlFileName);
      } else if (exportFormat === "docx") {
        await exportHtmlAsDocx(htmlCode, htmlFileName);
      } else {
        exportHtmlAsExcel(htmlCode, htmlFileName);
      }
      setExportStatus("done");
      setTimeout(() => setExportStatus("idle"), 5000);
    } catch (e: unknown) {
      setExportStatus("error");
      setExportError(e instanceof Error ? e.message : "حدث خطأ أثناء التصدير");
    } finally {
      setExporting(false);
    }
  };

  const formatLabels: Record<ExportFormat, string> = {
    pdf: "PDF",
    docx: "Word (.docx)",
    excel: "Excel (.xlsx)",
  };

  const formatIcons: Record<ExportFormat, string> = {
    pdf: "📄",
    docx: "📝",
    excel: "📊",
  };

  const exampleHtml = `<h1>عنوان المستند</h1>
<p>هذا نص تجريبي باللغة العربية. يمكنك كتابة أي محتوى HTML هنا.</p>
<h2>This is an English Heading</h2>
<p>English text goes left to right automatically.</p>
<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;width:100%">
  <tr><th>العمود 1</th><th>العمود 2</th><th>Column 3</th></tr>
  <tr><td>بيانات أ</td><td>بيانات ب</td><td>Data C</td></tr>
  <tr><td>بيانات د</td><td>بيانات هـ</td><td>Data F</td></tr>
</table>
<p>يمكنك إضافة قوائم أيضاً:</p>
<ul>
  <li>العنصر الأول</li>
  <li>العنصر الثاني</li>
  <li>Item Three</li>
</ul>`;

  return (
    <div className="space-y-5">
      {/* Settings */}
      <div className="bg-white rounded-xl border border-border shadow-sm p-4">
        <h2 className="text-sm font-semibold text-foreground mb-3 flex items-center gap-2">
          <Code size={15} className="text-primary" />
          إعدادات التصدير
        </h2>

        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mb-4">
          <div>
            <label className="block text-xs text-muted-foreground mb-1">اسم الملف</label>
            <input
              type="text"
              value={htmlFileName}
              onChange={(e) => setHtmlFileName(e.target.value)}
              className="w-full border border-input rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ring bg-background"
              placeholder="اسم الملف"
            />
          </div>
          <div>
            <label className="block text-xs text-muted-foreground mb-1">صيغة التصدير</label>
            <div className="flex gap-2">
              {(["pdf", "docx", "excel"] as ExportFormat[]).map((fmt) => (
                <button
                  key={fmt}
                  onClick={() => setExportFormat(fmt)}
                  className={`flex-1 flex flex-col items-center gap-0.5 py-2 px-1 rounded-lg text-xs font-medium border transition
                    ${exportFormat === fmt
                      ? "border-primary bg-primary text-white shadow-sm"
                      : "border-border bg-white text-foreground hover:bg-muted/30"
                    }`}
                >
                  <span className="text-base">{formatIcons[fmt]}</span>
                  <span>{fmt === "pdf" ? "PDF" : fmt === "docx" ? "Word" : "Excel"}</span>
                </button>
              ))}
            </div>
          </div>
        </div>

        <div className="bg-muted/30 rounded-lg p-3 text-xs text-muted-foreground flex flex-wrap gap-4">
          <span>📄 PDF — يحافظ على التصميم والألوان كاملاً</span>
          <span>📝 Word — يدعم العربية والإنجليزية مع الجداول</span>
          <span>📊 Excel — يستخرج الجداول إلى أوراق عمل</span>
        </div>
      </div>

      {/* HTML Editor */}
      <div className="bg-white rounded-xl border border-border shadow-sm">
        <div className="flex items-center justify-between px-4 pt-3 pb-2 border-b border-border">
          <h2 className="text-sm font-semibold text-foreground flex items-center gap-2">
            <Code size={14} className="text-primary" />
            كود HTML
          </h2>
          <button
            onClick={() => setHtmlCode(exampleHtml)}
            className="text-xs text-primary hover:underline"
          >
            إدراج مثال
          </button>
        </div>

        <div className="p-4">
          <p className="text-xs text-muted-foreground mb-2">
            💡 الصق كود HTML — سيتم عرضه في المعاينة وتصديره بالصيغة المختارة.
            النص العربي يبدأ من اليمين تلقائياً، والإنجليزي من اليسار.
          </p>
          <textarea
            value={htmlCode}
            onChange={(e) => setHtmlCode(e.target.value)}
            placeholder={`<h1>عنوان المستند</h1>\n<p>نص عربي من اليمين</p>\n<p dir="ltr">English text from left</p>\n<table>...</table>`}
            className="w-full h-56 border border-input rounded-lg p-3 text-sm font-mono resize-y focus:outline-none focus:ring-2 focus:ring-ring bg-background leading-relaxed"
            dir="ltr"
            style={{ fontFamily: "monospace", textAlign: "left" }}
          />
        </div>

        <div className="px-4 pb-4">
          <button
            onClick={handleExport}
            disabled={!htmlCode.trim() || exporting}
            className="flex items-center gap-2 bg-primary text-white px-6 py-2.5 rounded-lg text-sm font-semibold hover:opacity-90 disabled:opacity-50 transition shadow-sm"
          >
            {exporting ? (
              <><RefreshCw size={15} className="animate-spin" />جارٍ التصدير...</>
            ) : (
              <><Download size={15} />تصدير كـ {formatLabels[exportFormat]}</>
            )}
          </button>

          {exportStatus === "done" && (
            <div className="mt-3 flex items-center gap-2 bg-emerald-50 border border-emerald-200 text-emerald-700 rounded-lg px-3 py-2 text-sm">
              <CheckCircle size={15} />
              <span>تم التصدير بنجاح! تحقق من مجلد التنزيلات.</span>
            </div>
          )}
          {exportStatus === "error" && (
            <div className="mt-3 flex items-center gap-2 bg-red-50 border border-red-200 text-red-700 rounded-lg px-3 py-2 text-sm">
              <AlertCircle size={15} />
              <span>{exportError || "حدث خطأ أثناء التصدير"}</span>
            </div>
          )}
        </div>
      </div>

      {/* Live HTML Preview */}
      {htmlCode.trim() && (
        <div className="bg-white rounded-xl border border-border shadow-sm overflow-hidden">
          <div className="px-4 py-3 border-b border-border flex items-center justify-between">
            <h2 className="text-sm font-semibold text-foreground">معاينة HTML</h2>
            <span className="text-xs text-muted-foreground bg-muted/40 px-2 py-0.5 rounded">
              هذا ما ستظهر عليه النسخة المُصدَّرة
            </span>
          </div>
          <div
            className="p-6 min-h-24 max-h-[500px] overflow-auto"
            style={{ direction: "auto" }}
            dangerouslySetInnerHTML={{ __html: htmlCode }}
          />
        </div>
      )}

      {/* Guide */}
      <div className="bg-white rounded-xl border border-border shadow-sm p-4">
        <h2 className="text-sm font-semibold text-foreground mb-3">أمثلة على العناصر المدعومة</h2>
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs text-muted-foreground">
          <div className="bg-slate-50 rounded-lg p-3 border border-slate-100">
            <p className="font-semibold text-foreground mb-1">📑 العناوين</p>
            <code className="text-[10px] leading-relaxed block ltr">
              {`<h1>عنوان رئيسي</h1>\n<h2>عنوان فرعي</h2>`}
            </code>
          </div>
          <div className="bg-blue-50 rounded-lg p-3 border border-blue-100">
            <p className="font-semibold text-blue-800 mb-1">📊 الجداول</p>
            <code className="text-[10px] leading-relaxed block ltr">
              {`<table>\n  <tr><th>عمود</th></tr>\n  <tr><td>بيانات</td></tr>\n</table>`}
            </code>
          </div>
          <div className="bg-emerald-50 rounded-lg p-3 border border-emerald-100">
            <p className="font-semibold text-emerald-800 mb-1">🔤 اتجاه النص</p>
            <code className="text-[10px] leading-relaxed block ltr">
              {`<p>نص عربي ← يمين</p>\n<p dir="ltr">English → left</p>`}
            </code>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── Document Formatter Tab ────────────────────────────────────────────────────
function DocumentFormatterSection() {
  const [text, setText] = useState("");
  const [fileName, setFileName] = useState("المستند_المنسّق");
  const [blocks, setBlocks] = useState<ParsedBlock[]>([]);
  const [status, setStatus] = useState<Status>("idle");
  const [errorMsg, setErrorMsg] = useState("");
  const [prelim, setPrelim] = useState(5);
  const [selectedBorder, setSelectedBorder] = useState<BorderPreset>(BORDER_PRESETS[0]);

  const [uploadStatus, setUploadStatus] = useState<UploadStatus>("idle");
  const [uploadedFileName, setUploadedFileName] = useState("");
  const [uploadError, setUploadError] = useState("");
  const [uploadedFileKind, setUploadedFileKind] = useState<FileKind>(null);
  const [hadEncodingIssues, setHadEncodingIssues] = useState(false);
  const [structuredBlocks, setStructuredBlocks] = useState<ParsedBlock[] | null>(null);

  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleAnalyze = () => {
    if (!text.trim()) return;
    const parsed = parseText(text);
    setBlocks(parsed);
    setStructuredBlocks(null);
  };

  const handleGenerate = async () => {
    const hasContent = text.trim() || (structuredBlocks && structuredBlocks.length > 0);
    if (!hasContent) return;
    setStatus("processing");
    setErrorMsg("");
    try {
      const parsed = structuredBlocks ?? (blocks.length > 0 ? blocks : parseText(text));
      await generateDocx(
        parsed,
        fileName,
        prelim,
        selectedBorder.id === "none" ? null : selectedBorder
      );
      setStatus("done");
      setTimeout(() => setStatus("idle"), 4000);
    } catch (e: unknown) {
      setStatus("error");
      setErrorMsg(e instanceof Error ? e.message : "حدث خطأ غير معروف");
    }
  };

  const handleClear = () => {
    setText("");
    setBlocks([]);
    setStructuredBlocks(null);
    setStatus("idle");
    setErrorMsg("");
    setUploadStatus("idle");
    setUploadedFileName("");
    setUploadError("");
    setUploadedFileKind(null);
    textareaRef.current?.focus();
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    e.target.value = "";

    const isDocx = file.name.toLowerCase().endsWith(".docx");
    const isPdf = file.name.toLowerCase().endsWith(".pdf");

    if (!isDocx && !isPdf) {
      setUploadStatus("error");
      setUploadError("الرجاء اختيار ملف .docx أو .pdf فقط");
      return;
    }

    setUploadStatus("reading");
    setUploadError("");
    setUploadedFileName(file.name);
    setUploadedFileKind(isDocx ? "docx" : "pdf");
    setStructuredBlocks(null);
    setBlocks([]);
    setHadEncodingIssues(false);

    try {
      const arrayBuffer = await file.arrayBuffer();

      if (isDocx) {
        const { blocks: parsedBlocks, rawText, hadEncodingIssues: enc } = await parseDocxToBlocks(arrayBuffer);
        setText(rawText);
        setStructuredBlocks(parsedBlocks);
        setBlocks(parsedBlocks);
        setHadEncodingIssues(enc);
        setFileName(file.name.replace(/\.docx$/i, "") + "_منسّق");
        setUploadStatus("ready");
      } else {
        const extracted = await parsePdfToText(arrayBuffer);
        if (!extracted.trim()) {
          setUploadStatus("error");
          setUploadError("لم يتم العثور على نص قابل للاستخراج في ملف PDF");
          return;
        }
        setText(extracted);
        const parsed = parseText(extracted);
        setBlocks(parsed);
        setFileName(file.name.replace(/\.pdf$/i, "") + "_منسّق");
        setUploadStatus("ready");
      }
    } catch (err: unknown) {
      setUploadStatus("error");
      const msg = err instanceof Error ? err.message : "";
      setUploadError(
        isDocx
          ? `تعذّر قراءة ملف وورد. ${msg}`
          : `تعذّر قراءة ملف PDF. تأكد أن الملف غير محمي بكلمة مرور. ${msg}`
      );
    }
  };

  const hasContent = text.trim().length > 0 || (structuredBlocks && structuredBlocks.length > 0);
  const showPreview = blocks.length > 0;
  const tableCount = (structuredBlocks ?? blocks).filter((b) => b.type === "table").length;
  const imageCount = (structuredBlocks ?? blocks).filter((b) => b.type === "image").length;

  return (
    <div className="space-y-5">
      {/* Settings */}
      <div className="bg-white rounded-xl border border-border p-4 shadow-sm">
        <h2 className="text-sm font-semibold text-foreground mb-3">إعدادات المستند</h2>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <div>
            <label className="block text-xs text-muted-foreground mb-1">اسم الملف</label>
            <input type="text" value={fileName} onChange={(e) => setFileName(e.target.value)}
              className="w-full border border-input rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ring bg-background" placeholder="اسم الملف" />
          </div>
          <div>
            <label className="block text-xs text-muted-foreground mb-1">عدد الصفحات التمهيدية (تُرقَّم بأرقام رومانية)</label>
            <input type="number" min={0} max={30} value={prelim} onChange={(e) => setPrelim(Number(e.target.value))}
              className="w-full border border-input rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-ring bg-background" />
          </div>
        </div>
        <div className="mt-3 grid grid-cols-3 gap-3 text-xs text-muted-foreground bg-muted/40 rounded-lg p-3">
          <span>📌 الخط: Simplified Arabic / Times New Roman</span>
          <span>📏 الهوامش: 1.25 | 1.0 بوصة</span>
          <span>📐 تباعد الأسطر: 1.5</span>
        </div>
      </div>

      {/* Page Borders */}
      <div className="bg-white rounded-xl border border-border p-4 shadow-sm">
        <div className="flex items-center justify-between mb-3">
          <h2 className="text-sm font-semibold text-foreground flex items-center gap-2">
            <span className="text-base">🖼️</span>حدود الصفحة
          </h2>
          {selectedBorder.id !== "none" && (
            <span className="text-xs text-primary font-medium flex items-center gap-1">
              <span>{selectedBorder.emoji}</span><span>{selectedBorder.label}</span>
            </span>
          )}
        </div>
        <div className="grid grid-cols-4 sm:grid-cols-6 lg:grid-cols-12 gap-2">
          {BORDER_PRESETS.map((preset) => (
            <BorderPreview key={preset.id} preset={preset}
              selected={selectedBorder.id === preset.id}
              onClick={() => setSelectedBorder(preset)} />
          ))}
        </div>
        {selectedBorder.id !== "none" && (
          <div className="mt-3 rounded-lg px-3 py-2 text-xs flex items-center gap-2"
            style={{ background: selectedBorder.previewColor + "15", border: `1px solid ${selectedBorder.previewColor}40`, color: selectedBorder.previewColor }}>
            <span className="text-base">{selectedBorder.emoji}</span>
            <span>سيُضاف إطار <strong>{selectedBorder.label}</strong> حول كل صفحات المستند</span>
          </div>
        )}
      </div>

      {/* File Upload */}
      <div className="bg-white rounded-xl border border-border shadow-sm p-4">
        <h2 className="text-sm font-semibold text-foreground mb-3 flex items-center gap-2">
          <Upload size={15} className="text-primary" />
          رفع ملف وورد أو PDF وإعادة تنسيقه
        </h2>

        <input ref={fileInputRef} type="file" accept=".docx,.pdf" className="hidden" onChange={handleFileUpload} />

        <div
          onClick={() => fileInputRef.current?.click()}
          className={`relative cursor-pointer rounded-xl border-2 border-dashed transition-all px-6 py-7
            flex flex-col items-center justify-center gap-3 text-center select-none
            ${uploadStatus === "ready" ? "border-emerald-400 bg-emerald-50 hover:bg-emerald-100"
              : uploadStatus === "error" ? "border-red-300 bg-red-50 hover:bg-red-100"
              : uploadStatus === "reading" ? "border-blue-300 bg-blue-50 cursor-wait"
              : "border-muted-foreground/25 bg-muted/20 hover:border-primary/50 hover:bg-primary/5"}`}
        >
          {uploadStatus === "reading" ? (
            <>
              <RefreshCw size={30} className="text-blue-500 animate-spin" />
              <p className="text-sm font-medium text-blue-700">جارٍ قراءة الملف وتحليل محتواه...</p>
            </>
          ) : uploadStatus === "ready" ? (
            <>
              <CheckCircle size={30} className="text-emerald-500" />
              <div className="space-y-1">
                <div className="flex items-center justify-center gap-2">
                  <p className="text-sm font-semibold text-emerald-700">تم تحميل الملف بنجاح</p>
                  <FileTypeBadge kind={uploadedFileKind} />
                </div>
                <p className="text-xs text-emerald-600">{uploadedFileName}</p>
                {tableCount > 0 && (
                  <p className="text-xs font-medium text-blue-600">
                    ✓ تم اكتشاف {tableCount} جدول وحفظ بنيتها
                  </p>
                )}
                {imageCount > 0 && (
                  <p className="text-xs font-medium text-purple-600">
                    🖼️ تم استخراج {imageCount} صورة/رسم بياني
                  </p>
                )}
                <p className="text-xs text-muted-foreground">
                  اضغط "إنشاء وتحميل" للحصول على النسخة المنسّقة
                </p>
              </div>
              <p className="text-xs text-emerald-600 underline underline-offset-2">اضغط لاختيار ملف آخر</p>
            </>
          ) : uploadStatus === "error" ? (
            <>
              <AlertCircle size={30} className="text-red-500" />
              <div>
                <p className="text-sm font-semibold text-red-700">فشل تحميل الملف</p>
                <p className="text-xs text-red-600 mt-0.5">{uploadError}</p>
              </div>
              <p className="text-xs text-red-500 underline underline-offset-2">اضغط للمحاولة مجدداً</p>
            </>
          ) : (
            <>
              <div className="flex gap-4 items-center">
                <div className="w-12 h-12 rounded-2xl bg-blue-100 flex items-center justify-center">
                  <span className="text-2xl">📝</span>
                </div>
                <span className="text-2xl text-muted-foreground">أو</span>
                <div className="w-12 h-12 rounded-2xl bg-red-100 flex items-center justify-center">
                  <span className="text-2xl">📄</span>
                </div>
              </div>
              <div>
                <p className="text-sm font-semibold text-foreground">اسحب ملفك هنا أو اضغط للاختيار</p>
                <p className="text-xs text-muted-foreground mt-1">
                  يدعم{" "}
                  <span className="font-mono bg-blue-50 text-blue-700 px-1 rounded">.docx</span>
                  {" "}(يحفظ الجداول والرسوم كما هي) و{" "}
                  <span className="font-mono bg-red-50 text-red-700 px-1 rounded">.pdf</span>
                  {" "}(يستخرج النص ويُنسّقه)
                </p>
              </div>
            </>
          )}
        </div>
      </div>

      {/* Text Editor */}
      <div className="bg-white rounded-xl border border-border shadow-sm">
        <div className="flex items-center justify-between px-4 pt-3 pb-2 border-b border-border">
          <h2 className="text-sm font-semibold text-foreground">
            النص المراد تنسيقه
            {uploadedFileKind && uploadStatus === "ready" && (
              <span className="mr-2 inline-flex"><FileTypeBadge kind={uploadedFileKind} /></span>
            )}
          </h2>
          {(text || structuredBlocks) && (
            <button onClick={handleClear} className="text-xs text-muted-foreground hover:text-foreground flex items-center gap-1">
              <RefreshCw size={12} /> مسح
            </button>
          )}
        </div>

        <div className="p-4">
          {uploadedFileKind === "docx" && structuredBlocks && (
            <div className="mb-2 text-xs bg-blue-50 border border-blue-200 text-blue-700 rounded-lg px-3 py-2 flex items-center gap-2">
              <Table size={12} />
              <span>
                تم استخراج بنية الملف كاملةً ({structuredBlocks.filter(b => b.type === "table").length} جدول
                {structuredBlocks.filter(b => b.type === "image").length > 0
                  ? ` + ${structuredBlocks.filter(b => b.type === "image").length} صورة`
                  : ""} محفوظ).
                النص أدناه للاطلاع فقط — التنسيق سيستخدم البنية الأصلية.
              </span>
            </div>
          )}
          {hadEncodingIssues && (
            <div className="mb-2 text-xs bg-amber-50 border border-amber-300 text-amber-800 rounded-lg px-3 py-2 flex items-start gap-2">
              <span className="text-base shrink-0 mt-0.5">⚠️</span>
              <span>
                <strong>تحذير:</strong> يحتوي الملف على نص عربي مُشفَّر بصيغة قديمة (Presentation Forms). تمّ تطبيع النص تلقائياً، لكن قد يكون ترتيب بعض الكلمات غير مثالي.{" "}
                <strong>يُوصى باستخدام نسخة PDF بدلاً من DOCX للحصول على أفضل النتائج.</strong>
              </span>
            </div>
          )}
          {uploadedFileKind !== "docx" && (
            <p className="text-xs text-muted-foreground mb-2">
              💡 للجداول: افصل الأعمدة بـ <code className="bg-muted px-1 rounded">|</code> أو بعدة مسافات أو بـ Tab
            </p>
          )}
          <textarea
            ref={textareaRef}
            value={text}
            onChange={(e) => { setText(e.target.value); setBlocks([]); setStructuredBlocks(null); }}
            placeholder={`الصق نصك هنا...\n\nمثال على جدول:\nالاسم | العمر | التخصص\nأحمد | 25 | هندسة\nسارة | 23 | طب`}
            className="w-full h-64 border border-input rounded-lg p-3 text-sm resize-y focus:outline-none focus:ring-2 focus:ring-ring bg-background leading-relaxed font-mono"
            dir="auto"
          />
        </div>

        <div className="px-4 pb-4 flex flex-wrap gap-2">
          <button onClick={handleAnalyze} disabled={!text.trim()}
            className="flex items-center gap-2 bg-secondary text-secondary-foreground px-4 py-2 rounded-lg text-sm font-medium hover:opacity-90 disabled:opacity-40 transition">
            <Table size={15} />
            تحليل النص ومعاينة الجداول
          </button>

          <button onClick={handleGenerate} disabled={!hasContent || status === "processing"}
            className="flex items-center gap-2 bg-primary text-primary-foreground px-5 py-2 rounded-lg text-sm font-semibold hover:opacity-90 disabled:opacity-50 transition shadow-sm">
            {status === "processing" ? (
              <><RefreshCw size={15} className="animate-spin" />جارٍ الإنشاء...</>
            ) : (
              <><Download size={15} />إنشاء وتحميل ملف وورد</>
            )}
          </button>
        </div>
      </div>

      {/* Status */}
      {status === "done" && (
        <div className="flex items-center gap-2 bg-emerald-50 border border-emerald-200 text-emerald-700 rounded-xl px-4 py-3 text-sm">
          <CheckCircle size={18} /><span>تم إنشاء الملف بنجاح! تحقق من مجلد التنزيلات.</span>
        </div>
      )}
      {status === "error" && (
        <div className="flex items-center gap-2 bg-red-50 border border-red-200 text-red-700 rounded-xl px-4 py-3 text-sm">
          <AlertCircle size={18} /><span>{errorMsg || "حدث خطأ أثناء إنشاء الملف"}</span>
        </div>
      )}

      {/* Preview */}
      {showPreview && (
        <div className="bg-white rounded-xl border border-border shadow-sm">
          <div className="flex items-center justify-between px-4 py-3 border-b border-border">
            <h2 className="text-sm font-semibold text-foreground">معاينة هيكل المستند</h2>
            <StatsBar blocks={blocks} />
          </div>
          <div className="p-4 max-h-96 overflow-y-auto space-y-1">
            {blocks.map((block, i) => (
              <BlockPreview key={i} block={block} index={i} />
            ))}
          </div>
        </div>
      )}

      {/* Guide */}
      <div className="bg-white rounded-xl border border-border shadow-sm p-4">
        <h2 className="text-sm font-semibold text-foreground mb-3">دليل الاستخدام</h2>
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 text-xs text-muted-foreground">
          <div className="bg-blue-50 rounded-lg p-3 border border-blue-100">
            <p className="font-semibold text-blue-800 mb-1">📝 ملف وورد (.docx)</p>
            <p>يستخرج النص والجداول والصور والرسوم مباشرة — كل العناصر تُحفظ كما هي في موضعها</p>
          </div>
          <div className="bg-red-50 rounded-lg p-3 border border-red-100">
            <p className="font-semibold text-red-800 mb-1">📄 ملف PDF</p>
            <p>يستخرج النص صفحةً بصفحة ويُعيد بناء الفقرات، ثم يكتشف الجداول تلقائياً</p>
          </div>
          <div className="bg-muted/40 rounded-lg p-3">
            <p className="font-semibold text-foreground mb-1">✍️ نص مباشر</p>
            <pre className="font-mono leading-relaxed whitespace-pre-wrap">{`الاسم | الدرجة | النتيجة\nأحمد | 90 | ممتاز\nسارة | 85 | جيد جداً`}</pre>
          </div>
        </div>
      </div>
    </div>
  );
}

// ─── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  const [activeTab, setActiveTab] = useState<ActiveTab>("formatter");

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-blue-50" dir="rtl">
      {/* Header */}
      <header className="bg-white border-b border-border shadow-sm">
        <div className="max-w-5xl mx-auto px-4 py-4 flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-primary flex items-center justify-center shadow-sm">
            <FileText size={20} className="text-white" />
          </div>
          <div>
            <h1 className="text-lg font-bold text-foreground">منسّق مستندات وورد</h1>
            <p className="text-xs text-muted-foreground">
              تنسيق المستندات • محوّل HTML • دعم الرسوم البيانية والجداول
            </p>
          </div>
        </div>
      </header>

      {/* Tab Navigation */}
      <div className="max-w-5xl mx-auto px-4 pt-5">
        <div className="flex gap-1 bg-white rounded-xl border border-border p-1 shadow-sm">
          <button
            onClick={() => setActiveTab("formatter")}
            className={`flex-1 flex items-center justify-center gap-2 py-2.5 rounded-lg text-sm font-medium transition
              ${activeTab === "formatter"
                ? "bg-primary text-white shadow-sm"
                : "text-muted-foreground hover:bg-muted/30 hover:text-foreground"
              }`}
          >
            <FileText size={14} />
            منسّق المستندات
          </button>
          <button
            onClick={() => setActiveTab("html")}
            className={`flex-1 flex items-center justify-center gap-2 py-2.5 rounded-lg text-sm font-medium transition
              ${activeTab === "html"
                ? "bg-primary text-white shadow-sm"
                : "text-muted-foreground hover:bg-muted/30 hover:text-foreground"
              }`}
          >
            <Code size={14} />
            محوّل HTML
          </button>
        </div>
      </div>

      <main className="max-w-5xl mx-auto px-4 py-5">
        {activeTab === "formatter" ? (
          <DocumentFormatterSection />
        ) : (
          <HtmlExporterSection />
        )}
      </main>
    </div>
  );
}
