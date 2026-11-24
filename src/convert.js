/**
 * 文档转换核心：docx / html / markdown / 纯文本
 * 所有处理都在浏览器本地完成，不产生任何网络请求。
 */
import mammoth from "mammoth/mammoth.browser.min.js";
import { marked } from "marked";
import TurndownService from "turndown/lib/turndown.browser.es.js";
import { gfm } from "turndown-plugin-gfm/lib/turndown-plugin-gfm.browser.es.js";
import { htmlToDocxBlob } from "./docx-build.js";

marked.use({ gfm: true, breaks: false, pedantic: false });

/* ---------------- HTML 包装 ---------------- */

const DOC_CSS = `
:root { color-scheme: light; }
* { box-sizing: border-box; }
body {
  margin: 0; padding: 40px 28px 72px;
  font: 15px/1.78 -apple-system, "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", "Segoe UI", sans-serif;
  color: #1f2430; background: #fff; -webkit-font-smoothing: antialiased;
}
.doc { max-width: 760px; margin: 0 auto; }
h1, h2, h3, h4, h5, h6 { color: #111827; line-height: 1.35; letter-spacing: -.3px; margin: 1.6em 0 .6em; }
h1 { font-size: 30px; margin-top: 0; }
h2 { font-size: 23px; }
h3 { font-size: 19px; }
h4 { font-size: 17px; }
h5, h6 { font-size: 15px; }
p { margin: 0 0 1em; }
a { color: #4f46e5; text-decoration: none; border-bottom: 1px solid rgba(79,70,229,.32); }
img { max-width: 100%; height: auto; border-radius: 8px; margin: 8px 0; }
ul, ol { margin: 0 0 1em; padding-left: 1.5em; }
li { margin: .25em 0; }
table { border-collapse: collapse; width: 100%; margin: 1em 0; font-size: 13.5px; }
th, td { border: 1px solid #dfe3ea; padding: 8px 11px; text-align: left; vertical-align: top; }
th { background: #f5f7fa; font-weight: 600; }
code { font: 13px/1.6 ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; background: #f2f4f8; padding: 2px 6px; border-radius: 5px; }
pre { background: #f7f8fb; border: 1px solid #e4e7ec; border-radius: 10px; padding: 14px 16px; overflow-x: auto; }
pre code { background: none; padding: 0; font-size: 13px; }
blockquote { margin: 1em 0; padding: 2px 0 2px 16px; border-left: 3px solid #dfe3ea; color: #4b5563; font-style: italic; }
hr { border: 0; border-top: 1px solid #e4e7ec; margin: 2em 0; }
.doc-meta { max-width: 760px; margin: 0 auto 28px; padding-bottom: 14px; border-bottom: 1px solid #eceef3;
  font-size: 12px; color: #98a2b3; display: flex; gap: 10px; align-items: center; }
`;

export function wrapDocumentHtml(bodyHtml, title = "转换结果") {
  const safeTitle = String(title).replace(/[<>&]/g, (c) => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;" }[c]));
  return `<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${safeTitle}</title>
<style>${DOC_CSS}</style>
</head><body><div class="doc">${bodyHtml}</div></body></html>`;
}

export function wrapPreviewHtml(bodyHtml) {
  return `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>${DOC_CSS}</style></head><body><div class="doc">${bodyHtml}</div></body></html>`;
}

/* ---------------- docx 读取 ---------------- */

export async function docxToHtmlFragment(arrayBuffer) {
  const result = await mammoth.convertToHtml(
    { arrayBuffer },
    {
      styleMap: [
        "p[style-name='Title'] => h1:fresh",
        "p[style-name='Subtitle'] => p.subtitle:fresh",
      ],
      includeDefaultStyleMap: true,
      ignoreEmptyParagraphs: true,
    },
  );
  const warnings = (result.messages || [])
    .filter((m) => m.type === "warning")
    .map((m) => m.message)
    .slice(0, 6);
  return { html: result.value || "", warnings };
}

export async function docxToPlainText(arrayBuffer) {
  const result = await mammoth.extractRawText({ arrayBuffer });
  return (result.value || "").replace(/\n{3,}/g, "\n\n").trim();
}

/* ---------------- HTML ↔ Markdown ---------------- */

export function makeTurndown() {
  const td = new TurndownService({
    headingStyle: "atx",
    hr: "---",
    bulletListMarker: "-",
    codeBlockStyle: "fenced",
    fence: "```",
    emDelimiter: "*",
    strongDelimiter: "**",
    linkStyle: "inlined",
  });
  td.use(gfm);

  // 表格：Word 导出的表格没有 th，GFM 插件会原样保留 HTML，这里统一转成管道表格
  td.addRule("dsTable", {
    filter: (node) => node.nodeName === "TABLE",
    replacement: (_content, node) => {
      const rows = Array.from(node.rows || []);
      if (!rows.length) return "";
      const clean = (cell) => {
        const md = td.turndown(cell).replace(/\n+/g, " ").replace(/\|/g, "\\|").trim();
        return md || " ";
      };
      const cols = Math.max(...rows.map((r) => r.cells.length));
      const hasTh = Array.from(rows[0].cells || []).some((c) => c.tagName === "TH");
      const headRow = hasTh ? rows[0] : null;
      const bodyRows = hasTh ? rows.slice(1) : rows;
      const head = headRow
        ? Array.from(headRow.cells).map(clean)
        : (bodyRows.length > 1 ? Array.from(bodyRows[0].cells).map(clean) : new Array(cols).fill(" "));
      const body = headRow ? bodyRows : bodyRows.slice(1);
      const pad = (arr) => {
        const out = arr.slice(0, cols);
        while (out.length < cols) out.push(" ");
        return out;
      };
      const line = (cells) => `| ${pad(cells).join(" | ")} |`;
      const sep = `| ${new Array(cols).fill("---").join(" | ")} |`;
      return `\n\n${[line(head), sep, ...body.map((r) => line(Array.from(r.cells).map(clean)))].join("\n")}\n\n`;
    },
  });

  td.addRule("escapeExtra", {
    filter: (node) => node.nodeName === "PRE" && !node.querySelector("code"),
    replacement: (content, node) => `\n\n\`\`\`\n${(node.textContent || "").replace(/\n+$/, "")}\n\`\`\`\n\n`,
  });
  return td;
}

/** 统一列表缩进（turndown 默认会在标记后补三个空格），并压掉多余空行 */
function tidyMarkdown(markdown) {
  return markdown
    .replace(/^(\s*)([-*+]) {2,}/gm, "$1$2 ")
    .replace(/^(\s*)(\d+)\. {2,}/gm, "$1$2. ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

/** 把 base64 内嵌图片换成短占位，避免 Markdown 里塞几百 KB 的 data URI */
function shortenInlineImages(markdown) {
  let count = 0;
  const md = markdown.replace(/!\[([^\]]*)\]\(\s*data:[^)]+\)/g, (_m, alt) => {
    count += 1;
    const name = alt ? alt.replace(/[\\/:*?"<>|\s]/g, "-") : `image-${count}`;
    return `![${alt || "图片"}](images/${name}.png)`;
  });
  return { markdown: md, count };
}

export function htmlToMarkdown(html) {
  const td = makeTurndown();
  const md = tidyMarkdown(td.turndown(html));
  return shortenInlineImages(md);
}

export function markdownToHtmlFragment(markdown) {
  return marked.parse(markdown, { async: false });
}

/* ---------------- HTML → 纯文本 ---------------- */

const BLOCK_TAGS = new Set([
  "p", "div", "section", "article", "header", "footer", "h1", "h2", "h3", "h4", "h5", "h6",
  "li", "tr", "pre", "blockquote", "hr", "br", "table", "dd", "dt",
]);

export function htmlToPlainText(html) {
  const doc = new DOMParser().parseFromString(html, "text/html");
  let out = "";
  (function walk(node) {
    if (node.nodeType === 3) {
      out += (node.nodeValue || "").replace(/[\t\r\n]+/g, " ").replace(/ {2,}/g, " ");
      return;
    }
    if (node.nodeType !== 1) return;
    const tag = node.tagName.toLowerCase();
    if (["script", "style", "noscript"].includes(tag)) return;
    const isBlock = BLOCK_TAGS.has(tag);
    if (isBlock && !out.endsWith("\n")) out += "\n";
    if (tag === "li") out += "- ";
    if (tag === "td" || tag === "th") {
      if (!out.endsWith("\n") && !out.endsWith("| ")) out += "| ";
    }
    for (const child of node.childNodes) walk(child);
    if (tag === "td" || tag === "th") out += " ";
    if (isBlock) out += "\n";
  })(doc.body);
  return out
    .replace(/[ \t]+/g, " ")
    .replace(/ ?\n ?/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

/* ---------------- Markdown → docx ---------------- */

export async function markdownToDocxBlob(markdown, opts = {}) {
  const html = markdownToHtmlFragment(markdown);
  return htmlToDocxBlob(html, opts);
}

export async function htmlToDocx(html, opts = {}) {
  return htmlToDocxBlob(html, opts);
}

/* ---------------- 文本编码 ---------------- */

export function decodeSmart(arrayBuffer) {
  const bytes = new Uint8Array(arrayBuffer);
  const hasBomUtf8 = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf;
  const hasBomLe = bytes[0] === 0xff && bytes[1] === 0xfe;
  const hasBomBe = bytes[0] === 0xfe && bytes[1] === 0xff;

  if (hasBomUtf8) {
    return { text: new TextDecoder("utf-8").decode(bytes.subarray(3)), encoding: "UTF-8（含 BOM）", confident: true };
  }
  if (hasBomLe) {
    return { text: new TextDecoder("utf-16le").decode(bytes.subarray(2)), encoding: "UTF-16 LE", confident: true };
  }
  if (hasBomBe) {
    try {
      return { text: new TextDecoder("utf-16be").decode(bytes.subarray(2)), encoding: "UTF-16 BE", confident: true };
    } catch { /* 继续向下探测 */ }
  }

  const tryDecode = (label, fatal) => {
    try {
      return new TextDecoder(label, { fatal }).decode(bytes);
    } catch {
      return null;
    }
  };

  const utf8 = tryDecode("utf-8", false);
  const strictUtf8 = tryDecode("utf-8", true);
  if (strictUtf8 !== null) {
    // 纯 ASCII 也走这里；如果出现替换字符说明解码不干净
    if (!utf8.includes("\uFFFD")) {
      return { text: strictUtf8, encoding: "UTF-8", confident: true };
    }
  }

  const gbk = tryDecode("gb18030", true);
  if (gbk !== null) return { text: gbk, encoding: "GBK / GB18030", confident: true };

  const big5 = tryDecode("big5", true);
  if (big5 !== null && !big5.includes("\uFFFD")) return { text: big5, encoding: "Big5", confident: false };

  return { text: utf8 || "", encoding: "未知（按 UTF-8 尽力解码）", confident: false };
}

export function textToBlob(text, withBom = false) {
  const body = new TextEncoder().encode(text);
  if (!withBom) return new Blob([body], { type: "text/plain;charset=utf-8" });
  const bom = new Uint8Array([0xef, 0xbb, 0xbf]);
  const merged = new Uint8Array(bom.length + body.length);
  merged.set(bom, 0);
  merged.set(body, bom.length);
  return new Blob([merged], { type: "text/plain;charset=utf-8" });
}

export const HELPERS = { shortenInlineImages };
