/**
 * HTML → DOCX 构建器
 * 把 HTML（来自 Word 转换结果、Markdown 渲染结果或用户上传的网页）
 * 映射成标准 OOXML 文档：标题、段落、列表、表格、代码块、图片、超链接。
 */
import {
  AlignmentType, BorderStyle, Document, ExternalHyperlink, HeadingLevel,
  ImageRun, LevelFormat, Packer, Paragraph, ShadingType, Table, TableCell,
  TableRow, TextRun, UnderlineType, WidthType,
} from "docx";

const MONO = "Consolas";
const BODY_FONT = { ascii: "Calibri", hAnsi: "Calibri", eastAsia: "等线" };
const HEAD_FONT = { ascii: "Calibri", hAnsi: "Calibri", eastAsia: "等线" };
const TWIP = 20; // 1pt = 20 twips

const HEADING_LEVELS = [
  HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3,
  HeadingLevel.HEADING_4, HeadingLevel.HEADING_5, HeadingLevel.HEADING_6,
];

/* ---------------- 内联内容 ---------------- */

function inlineChildren(el, style = {}, skipBlocks = false) {
  const out = [];
  for (const node of el.childNodes) {
    if (node.nodeType === 3) {
      const text = node.nodeValue.replace(/[\t\r\n]+/g, " ").replace(/ {2,}/g, " ");
      if (!text.trim()) continue;
      out.push(new TextRun({ text, ...style }));
      continue;
    }
    if (node.nodeType !== 1) continue;
    const tag = node.tagName.toLowerCase();
    if (skipBlocks && ["ul", "ol", "table", "pre", "figure"].includes(tag)) continue;
    if (tag === "br") {
      out.push(new TextRun({ text: "", break: 1, ...style }));
      continue;
    }
    if (tag === "img") continue;
    if (tag === "a") {
      const href = (node.getAttribute("href") || "").trim();
      const kids = inlineChildren(node, {
        ...style, color: "1155CC", underline: { type: UnderlineType.SINGLE },
      });
      if (kids.length && /^(https?:\/\/|mailto:|#)/i.test(href)) {
        out.push(new ExternalHyperlink({ link: href, children: kids }));
      } else {
        out.push(...kids);
      }
      continue;
    }
    const next = { ...style };
    if (tag === "strong" || tag === "b") next.bold = true;
    if (tag === "em" || tag === "i") next.italics = true;
    if (tag === "u") next.underline = { type: UnderlineType.SINGLE };
    if (tag === "s" || tag === "del" || tag === "strike") next.strike = true;
    if (tag === "code" || tag === "kbd" || tag === "samp") {
      next.font = MONO;
      next.size = 19;
      next.shading = { type: ShadingType.CLEAR, color: "auto", fill: "F1F2F4" };
    }
    if (tag === "mark") next.highlight = "yellow";
    if (tag === "sup") next.superScript = true;
    if (tag === "sub") next.subScript = true;
    out.push(...inlineChildren(node, next, skipBlocks));
  }
  return out;
}

function plainText(el) {
  return (el.textContent || "").replace(/\s+/g, " ").trim();
}

/* ---------------- 图片 ---------------- */

function dataUriToBytes(src) {
  const m = /^data:(image\/[a-z0-9.+-]+)?(;base64)?,([\s\S]*)$/i.exec(src);
  if (!m) return null;
  const mime = (m[1] || "image/png").toLowerCase();
  const type = mime.includes("jpeg") || mime.includes("jpg") ? "jpg"
    : mime.includes("gif") ? "gif"
    : mime.includes("bmp") ? "bmp"
    : mime.includes("svg") ? "svg" : "png";
  try {
    if (m[2]) {
      const bin = atob(m[3]);
      const bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      return { bytes, type };
    }
    return { bytes: new TextEncoder().encode(decodeURIComponent(m[3])), type };
  } catch {
    return null;
  }
}

async function measure(blob) {
  try {
    if (typeof createImageBitmap === "function") {
      const bmp = await createImageBitmap(blob);
      const size = { w: bmp.width, h: bmp.height };
      if (bmp.close) bmp.close();
      return size;
    }
  } catch { /* 回退到 Image */ }
  return new Promise((resolve) => {
    const img = new Image();
    const url = URL.createObjectURL(blob);
    img.onload = () => { resolve({ w: img.naturalWidth, h: img.naturalHeight }); URL.revokeObjectURL(url); };
    img.onerror = () => { resolve({ w: 480, h: 320 }); URL.revokeObjectURL(url); };
    img.src = url;
  });
}

async function imageParagraph(node) {
  const src = (node.getAttribute("src") || "").trim();
  const parsed = dataUriToBytes(src);
  if (!parsed || parsed.type === "svg") return null;
  const blob = new Blob([parsed.bytes]);
  const { w, h } = await measure(blob);
  const scale = Math.min(1, 600 / Math.max(w, h || 1));
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: 200 },
    children: [new ImageRun({
      data: parsed.bytes,
      type: parsed.type,
      transformation: { width: Math.round(w * scale), height: Math.round(h * scale) },
    })],
  });
}

/* ---------------- 块级内容 ---------------- */

function codeBlock(node) {
  const code = (node.textContent || "").replace(/^\n+|\n+$/g, "").replace(/\t/g, "  ");
  const lines = code.split("\n");
  const runs = [];
  lines.forEach((line, i) => {
    if (i) runs.push(new TextRun({ text: "", break: 1 }));
    runs.push(new TextRun({ text: line || " ", font: MONO, size: 18, color: "24292F" }));
  });
  return new Paragraph({
    children: runs.length ? runs : [new TextRun({ text: " ", font: MONO, size: 18 })],
    shading: { type: ShadingType.CLEAR, color: "auto", fill: "F5F6F8" },
    spacing: { before: 160, after: 200, line: 240 },
    indent: { left: 240, right: 160 },
    border: {
      top: { style: BorderStyle.SINGLE, size: 4, color: "E4E7EC", space: 6 },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: "E4E7EC", space: 6 },
      left: { style: BorderStyle.SINGLE, size: 4, color: "E4E7EC", space: 6 },
      right: { style: BorderStyle.SINGLE, size: 4, color: "E4E7EC", space: 6 },
    },
  });
}

async function listBlocks(listEl, ordered, level) {
  const out = [];
  const items = [...listEl.children].filter((c) => c.tagName.toLowerCase() === "li");
  for (const li of items) {
    const runs = inlineChildren(li, {}, true);
    const maker = ordered
      ? { numbering: { reference: "ds-ordered", level } }
      : { bullet: { level } };
    out.push(new Paragraph({
      children: runs.length ? runs : [new TextRun({ text: " " })],
      ...maker,
      spacing: { after: 60, line: 300 },
    }));
    for (const child of li.children) {
      const tag = child.tagName.toLowerCase();
      if (tag === "ul" || tag === "ol") {
        out.push(...await listBlocks(child, tag === "ol", Math.min(level + 1, 4)));
      }
    }
  }
  return out;
}

function tableBlock(node) {
  const rows = [...node.querySelectorAll("tr")];
  if (!rows.length) return null;
  const colCount = Math.max(...rows.map((r) => r.querySelectorAll("td,th").length), 1);
  const cellW = Math.floor(100 / colCount);
  const border = { style: BorderStyle.SINGLE, size: 4, color: "D5D9E2" };
  const tableRows = rows.map((tr) => new TableRow({
    tableHeader: tr.querySelector("th") ? true : undefined,
    children: [...tr.querySelectorAll("td,th")].map((cell) => new TableCell({
      width: { size: cellW, type: WidthType.PERCENTAGE },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      shading: cell.tagName.toLowerCase() === "th"
        ? { type: ShadingType.CLEAR, color: "auto", fill: "F2F4F8" } : undefined,
      children: (() => {
        const isHead = cell.tagName.toLowerCase() === "th";
        const runs = isHead
          ? [new TextRun({ text: plainText(cell), bold: true, size: 20 })]
          : inlineChildren(cell);
        return [new Paragraph({
          spacing: { before: 0, after: 0, line: 280 },
          children: runs.length ? runs : [new TextRun({ text: " " })],
        })];
      })(),
    })),
  }));
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: border, bottom: border, left: border, right: border,
      insideHorizontal: border, insideVertical: border,
    },
    rows: tableRows,
  });
}

function divider() {
  return new Paragraph({
    spacing: { before: 200, after: 240 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "D5D9E2", space: 1 } },
    children: [new TextRun({ text: "" })],
  });
}

async function blocksFrom(parent) {
  const blocks = [];
  let loose = [];
  const flush = () => {
    if (loose.length) {
      blocks.push(new Paragraph({ children: loose, spacing: { after: 160, line: 320 } }));
      loose = [];
    }
  };

  for (const node of parent.childNodes) {
    if (node.nodeType === 3) {
      const t = node.nodeValue.replace(/[\t\r\n]+/g, " ");
      if (t.trim()) loose.push(new TextRun({ text: t }));
      continue;
    }
    if (node.nodeType !== 1) continue;
    const tag = node.tagName.toLowerCase();

    if (tag === "p" || tag === "div" || tag === "section" || tag === "article") {
      flush();
      const runs = inlineChildren(node);
      const hasBlocks = [...node.children].some((c) =>
        ["p", "div", "ul", "ol", "table", "pre", "blockquote", "h1", "h2", "h3", "h4", "h5", "h6"].includes(c.tagName.toLowerCase()));
      if (runs.length || !hasBlocks) {
        const align = node.getAttribute("align");
        const opts = { spacing: { after: 160, line: 320 } };
        if (align === "center") opts.alignment = AlignmentType.CENTER;
        if (align === "right") opts.alignment = AlignmentType.RIGHT;
        blocks.push(new Paragraph({ children: runs.length ? runs : [new TextRun({ text: "" })], ...opts }));
      }
      if (hasBlocks) blocks.push(...await blocksFrom(node));
      continue;
    }

    if (/^h[1-6]$/.test(tag)) {
      flush();
      const lvl = Number(tag[1]) - 1;
      blocks.push(new Paragraph({
        heading: HEADING_LEVELS[lvl],
        children: inlineChildren(node).length ? inlineChildren(node) : [new TextRun({ text: plainText(node) })],
        spacing: { before: lvl === 0 ? 240 : 200, after: 120 },
      }));
      continue;
    }

    if (tag === "ul" || tag === "ol") {
      flush();
      blocks.push(...await listBlocks(node, tag === "ol", 0));
      continue;
    }

    if (tag === "table") {
      flush();
      const t = tableBlock(node);
      if (t) blocks.push(t);
      blocks.push(new Paragraph({ children: [new TextRun({ text: "", size: 8 })], spacing: { after: 80 } }));
      continue;
    }

    if (tag === "pre") { flush(); blocks.push(codeBlock(node)); continue; }

    if (tag === "blockquote") {
      flush();
      const quoteStyle = {
        indent: { left: 360 },
        border: { left: { style: BorderStyle.SINGLE, size: 12, color: "D5D9E2", space: 12 } },
        spacing: { before: 60, after: 140, line: 320 },
      };
      const paras = [];
      for (const child of node.children) {
        const t = child.tagName.toLowerCase();
        if (t === "p" || t === "div") {
          paras.push(new Paragraph({
            children: inlineChildren(child, { italics: true, color: "475467" }),
            ...quoteStyle,
          }));
        } else if (/^h[1-6]$/.test(t)) {
          paras.push(new Paragraph({
            children: inlineChildren(child, { bold: true, color: "1F2430" }),
            ...quoteStyle,
          }));
        } else {
          paras.push(new Paragraph({
            children: inlineChildren(node, { italics: true, color: "475467" }),
            ...quoteStyle,
          }));
          break;
        }
      }
      if (!paras.length) {
        paras.push(new Paragraph({
          children: inlineChildren(node, { italics: true, color: "475467" }),
          ...quoteStyle,
        }));
      }
      blocks.push(...paras);
      continue;
    }

    if (tag === "hr") { flush(); blocks.push(divider()); continue; }

    if (tag === "img") {
      flush();
      const p = await imageParagraph(node);
      if (p) blocks.push(p);
      else blocks.push(new Paragraph({ children: [new TextRun({ text: "[图片]", italics: true, color: "98A2B3" })] }));
      continue;
    }

    if (tag === "figure") { flush(); blocks.push(...await blocksFrom(node)); continue; }

    if (["script", "style", "head", "noscript", "iframe"].includes(tag)) continue;

    // 其它容器：继续向下递归
    const inner = [...node.children];
    if (inner.length) {
      flush();
      blocks.push(...await blocksFrom(node));
    } else {
      const t = inlineChildren(node);
      if (t.length) loose.push(...t);
    }
  }
  flush();
  return blocks;
}

/* ---------------- 文档组装 ---------------- */

const baseStyles = {
  default: {
    document: {
      run: { font: BODY_FONT, size: 22, color: "1F2430" },
      paragraph: { spacing: { line: 320, after: 160 } },
    },
  },
  paragraphStyles: [
    { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { font: HEAD_FONT, size: 40, bold: true, color: "111827" },
      paragraph: { spacing: { before: 320, after: 160 }, outlineLevel: 0 } },
    { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { font: HEAD_FONT, size: 32, bold: true, color: "1F2430" },
      paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 1 } },
    { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { font: HEAD_FONT, size: 26, bold: true, color: "374151" },
      paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 2 } },
    { id: "Heading4", name: "Heading 4", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { font: HEAD_FONT, size: 23, bold: true, color: "4B5563" },
      paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 3 } },
  ],
  characterStyles: [],
  numbering: {},
};

const numbering = {
  config: [{
    reference: "ds-ordered",
    levels: [0, 1, 2, 3, 4].map((level) => ({
      level,
      format: LevelFormat.DECIMAL,
      text: `%${level + 1}.`,
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 360 + level * 360, hanging: 260 } } },
    })),
  }],
};

export async function htmlToDocxBlob(html, { title = "文档" } = {}) {
  const doc = new DOMParser().parseFromString(html, "text/html");
  const blocks = await blocksFrom(doc.body);
  const document = new Document({
    creator: "DocShift",
    title,
    description: "由 DocShift 从 HTML 生成",
    styles: baseStyles,
    numbering,
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4，单位 twip
          margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 },
        },
      },
      children: blocks.length ? blocks : [new Paragraph({ text: "" })],
    }],
  });
  return Packer.toBlob(document);
}

export const __twip = TWIP;
