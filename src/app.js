/**
 * DocShift — 交互层
 */
import * as C from "./convert.js";

const REPO_URL = "https://github.com/"; // 部署后替换为实际仓库地址

const $ = (id) => document.getElementById(id);

/* ---------------- 转换清单 ---------------- */

const stripExt = (name = "document") => name.replace(/\.(docx|doc|md|markdown|html?|txt)$/i, "") || "document";

const CONVERSIONS = [
  {
    id: "docx-html", group: "Word 文档", label: "Word", to: "HTML", note: "标题 / 表格 / 图片",
    accept: [".docx"], modes: ["file"], outExt: "html", mime: "text/html",
    async run({ buffer, name }) {
      const r = await C.docxToHtmlFragment(buffer);
      const body = r.html || "<p></p>";
      const doc = C.wrapDocumentHtml(body, stripExt(name));
      return { text: doc, lang: "html", previewBody: body, filename: `${stripExt(name)}.html`, mime: "text/html", warnings: r.warnings };
    },
  },
  {
    id: "docx-md", group: "Word 文档", label: "Word", to: "Markdown", note: "适合笔记 / 文档站",
    accept: [".docx"], modes: ["file"], outExt: "md", mime: "text/markdown",
    async run({ buffer, name }) {
      const r = await C.docxToHtmlFragment(buffer);
      const { markdown, count } = C.htmlToMarkdown(r.html);
      const warnings = [...r.warnings];
      if (count) warnings.push(`共有 ${count} 张内嵌图片已改为 ${count} 个相对路径占位（Markdown 内联 base64 会让文件膨胀）。需要保留图片请改用「Word → HTML」。`);
      return { text: markdown, lang: "markdown", filename: `${stripExt(name)}.md`, mime: "text/markdown", warnings };
    },
  },
  {
    id: "docx-txt", group: "Word 文档", label: "Word", to: "纯文本", note: "只保留文字",
    accept: [".docx"], modes: ["file"], outExt: "txt", mime: "text/plain",
    async run({ buffer, name }) {
      const text = await C.docxToPlainText(buffer);
      return { text, lang: "text", filename: `${stripExt(name)}.txt`, mime: "text/plain", warnings: [] };
    },
  },
  {
    id: "md-html", group: "Markdown", label: "Markdown", to: "HTML", note: "生成可直接分发的网页",
    accept: [".md", ".markdown", ".txt"], modes: ["file", "paste"], outExt: "html", mime: "text/html",
    async run({ text, name }) {
      const body = C.markdownToHtmlFragment(text);
      return { text: C.wrapDocumentHtml(body, stripExt(name)), lang: "html", previewBody: body, filename: `${stripExt(name)}.html`, mime: "text/html", warnings: [] };
    },
  },
  {
    id: "md-docx", group: "Markdown", label: "Markdown", to: "Word", note: "标准 .docx，可二次编辑",
    accept: [".md", ".markdown", ".txt"], modes: ["file", "paste"], outExt: "docx", mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    async run({ text, name }) {
      const blob = await C.markdownToDocxBlob(text, { title: stripExt(name) });
      return { binary: true, blob, filename: `${stripExt(name)}.docx`, outExt: "docx", label: "Word 文档", note: "含标题、列表、表格、代码块" };
    },
  },
  {
    id: "md-txt", group: "Markdown", label: "Markdown", to: "纯文本", note: "去掉所有标记",
    accept: [".md", ".markdown", ".txt"], modes: ["file", "paste"], outExt: "txt", mime: "text/plain",
    async run({ text, name }) {
      const plain = C.htmlToPlainText(C.markdownToHtmlFragment(text));
      return { text: plain, lang: "text", filename: `${stripExt(name)}.txt`, mime: "text/plain", warnings: [] };
    },
  },
  {
    id: "html-md", group: "HTML", label: "HTML", to: "Markdown", note: "网页 / 富文本回迁",
    accept: [".html", ".htm"], modes: ["file", "paste"], outExt: "md", mime: "text/markdown",
    async run({ text, name }) {
      const { markdown, count } = C.htmlToMarkdown(text);
      const warnings = count ? [`已把 ${count} 张内嵌图片改为相对路径占位。`] : [];
      return { text: markdown, lang: "markdown", filename: `${stripExt(name)}.md`, mime: "text/markdown", warnings };
    },
  },
  {
    id: "html-docx", group: "HTML", label: "HTML", to: "Word", note: "网页整理成文档",
    accept: [".html", ".htm"], modes: ["file", "paste"], outExt: "docx", mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    async run({ text, name }) {
      const blob = await C.htmlToDocx(text, { title: stripExt(name) });
      return { binary: true, blob, filename: `${stripExt(name)}.docx`, outExt: "docx", label: "Word 文档", note: "含标题、段落、列表、表格" };
    },
  },
];

/* ---------------- 状态 ---------------- */

let current = CONVERSIONS[0];
let source = null;          // { kind:'file'|'text', file, buffer, text, name, size }
let result = null;
let runToken = 0;
let debounceTimer = 0;

/* ---------------- 初始化 ---------------- */

function initChrome() {
  if (REPO_URL && REPO_URL !== "https://github.com/") {
    $("ghLink").href = REPO_URL;
    $("ghLink2").href = REPO_URL;
  } else {
    $("ghLink").classList.add("hidden");
    $("ghLink2").classList.add("hidden");
  }
}

function renderConvList() {
  const groups = new Map();
  for (const c of CONVERSIONS) {
    if (!groups.has(c.group)) groups.set(c.group, []);
    groups.get(c.group).push(c);
  }
  const host = $("convList");
  host.textContent = "";
  for (const [group, items] of groups) {
    const g = document.createElement("div");
    g.className = "conv-group";
    const label = document.createElement("div");
    label.className = "conv-group-label";
    label.textContent = group;
    g.appendChild(label);
    const row = document.createElement("div");
    row.className = "conv-group-items";
    for (const c of items) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "conv";
      btn.setAttribute("role", "radio");
      btn.setAttribute("aria-checked", String(c === current));
      btn.dataset.id = c.id;

      const path = document.createElement("span");
      path.className = "conv-path";
      path.textContent = `${c.label} `;
      const arrow = document.createElement("span");
      arrow.textContent = "→";
      arrow.style.color = "var(--ink-3)";
      const to = document.createElement("span");
      to.className = "conv-to";
      to.textContent = c.to;
      path.append(arrow, document.createTextNode(" "), to);

      const note = document.createElement("span");
      note.className = "conv-note";
      note.textContent = c.note;

      btn.append(path, note);
      btn.addEventListener("click", () => selectConversion(c));
      row.appendChild(btn);
    }
    g.appendChild(row);
    host.appendChild(g);
  }
}

function selectConversion(c) {
  current = c;
  for (const btn of $("convList").querySelectorAll(".conv")) {
    btn.setAttribute("aria-checked", String(btn.dataset.id === c.id));
  }
  resetSource();
  $("acceptHint").textContent = c.accept.join(" / ");
  $("fileInput").accept = c.accept.join(",");
  $("dropSub").textContent = `支持 ${c.accept.join(" / ")}，单个文件`;
  const canPaste = c.modes.includes("paste");
  $("pasteWrap").classList.toggle("hidden", !canPaste);
  if (!canPaste) $("pasteArea").value = "";
}

/* ---------------- 来源管理 ---------------- */

function resetSource() {
  source = null;
  result = null;
  $("srcMeta").classList.add("hidden");
  $("srcMeta").textContent = "";
  $("fileInput").value = "";
  $("convertBtn").disabled = true;
  $("clearBtn").classList.add("hidden");
  hideError();
  clearResult();
}

function clearResult() {
  $("result").classList.add("hidden");
  $("empty").classList.remove("hidden");
  $("previewBox").textContent = "";
  $("sourcePre").textContent = "";
  $("warns").classList.add("hidden");
  $("warns").textContent = "";
  $("tabs").classList.add("hidden");
  $("resMeta").textContent = "";
}

function fmtBytes(n) {
  if (!n && n !== 0) return "";
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / 1024 / 1024).toFixed(2)} MB`;
}

function renderSourceMeta() {
  if (!source) return;
  const box = $("srcMeta");
  box.textContent = "";
  box.classList.remove("hidden");
  const chip = document.createElement("span");
  chip.className = "chip-file";
  chip.textContent = source.kind === "file" ? source.name : "粘贴内容";
  const tag = document.createElement("span");
  tag.className = "chip-tag";
  tag.textContent = source.kind === "file"
    ? `${fmtBytes(source.size)} · ${(source.name.split(".").pop() || "").toUpperCase()}`
    : `${source.text.length} 字符`;
  box.append(chip, tag);
  $("convertBtn").disabled = false;
  $("clearBtn").classList.remove("hidden");
}

async function acceptFile(file) {
  if (!file) return;
  hideError();
  const ext = `.${(file.name.split(".").pop() || "").toLowerCase()}`;
  if (!current.accept.includes(ext)) {
    showError(`当前转换只接受 ${current.accept.join(" / ")}，你选的是 ${ext}。请切换转换类型或换一个文件。`);
    return;
  }
  if (file.size > 40 * 1024 * 1024) {
    showError("文件超过 40 MB。浏览器端转换会占用大量内存，建议先压缩图片或拆分文档。");
    return;
  }
  try {
    const buffer = await file.arrayBuffer();
    let text = null;
    if (current.modes.includes("paste")) {
      text = C.decodeSmart(buffer).text;
    }
    source = { kind: "file", file, buffer, text, name: file.name, size: file.size };
    $("pasteArea").value = "";
    renderSourceMeta();
    scheduleRun(0);
  } catch (e) {
    showError(`读取文件失败：${e.message || e}`);
  }
}

function acceptPaste(text) {
  const trimmed = text.trim();
  if (!trimmed) {
    if (source && source.kind === "text") resetSource();
    return;
  }
  source = { kind: "text", text, name: "粘贴内容", size: text.length };
  $("fileInput").value = "";
  renderSourceMeta();
  scheduleRun(700);
}

/* ---------------- 执行转换 ---------------- */

function setBusy(busy, label) {
  const btn = $("convertBtn");
  btn.disabled = busy || !source;
  btn.querySelector("span").textContent = busy ? (label || "转换中…") : (result ? "重新转换" : "开始转换");
}

function scheduleRun(delay) {
  clearTimeout(debounceTimer);
  debounceTimer = setTimeout(runConversion, delay || 0);
}

async function runConversion() {
  if (!source || !current) return;
  const token = ++runToken;
  setBusy(true);
  hideError();
  const started = performance.now();
  try {
    const input = source.kind === "file"
      ? { buffer: source.buffer, name: source.name }
      : { text: source.text, name: source.name };
    const out = await current.run(input);
    if (token !== runToken) return;
    result = out;
    renderResult(out, performance.now() - started);
  } catch (e) {
    if (token !== runToken) return;
    console.error(e);
    showError(`转换失败：${e?.message || e}`);
    clearResult();
  } finally {
    if (token === runToken) setBusy(false);
  }
}

/* ---------------- 渲染结果 ---------------- */

function chipArrow() {
  const s = document.createElement("span");
  s.className = "chip-arrow";
  s.textContent = "→";
  return s;
}

function renderResult(out, ms) {
  $("empty").classList.add("hidden");
  $("result").classList.remove("hidden");

  // 元信息
  const meta = $("resMeta");
  meta.textContent = "";
  const from = document.createElement("span");
  from.className = "chip-file";
  from.textContent = source.kind === "file" ? source.name : "粘贴内容";
  const toTag = document.createElement("span");
  toTag.className = "chip-tag";
  toTag.textContent = (out.outExt || (out.filename || "").split(".").pop() || "").toUpperCase();
  const size = document.createElement("span");
  const blobSize = out.blob ? out.blob.size : (out.text ? new Blob([out.text]).size : 0);
  size.textContent = fmtBytes(blobSize);
  const time = document.createElement("span");
  time.textContent = `${Math.round(ms)} ms`;
  meta.append(from, chipArrow(), toTag, size, document.createTextNode("·"), time);

  // 警告
  if (out.warnings && out.warnings.length) {
    $("warns").classList.remove("hidden");
    $("warns").textContent = out.warnings.join("  ");
  } else {
    $("warns").classList.add("hidden");
    $("warns").textContent = "";
  }

  // 预览
  const preview = $("previewBox");
  preview.textContent = "";
  const sourceBox = $("sourceBox");

  if (out.binary) {
    const card = document.createElement("div");
    card.className = "bin-card";
    const icon = document.createElement("div");
    icon.className = "bin-icon";
    icon.textContent = (out.outExt || "FILE").toUpperCase().slice(0, 4);
    const text = document.createElement("div");
    const t = document.createElement("div");
    t.className = "bin-t";
    t.textContent = out.label || "已生成文件";
    const s = document.createElement("div");
    s.className = "bin-s";
    s.textContent = `${out.filename} · ${fmtBytes(out.blob.size)} · 点击下方按钮保存`;
    text.append(t, s);
    card.append(icon, text);
    preview.appendChild(card);
    $("tabs").classList.add("hidden");
    sourceBox.classList.add("hidden");
    $("copyBtn").classList.add("hidden");
  } else {
    // 文本预览
    let previewHtml = null;
    if (out.lang === "html") previewHtml = C.wrapPreviewHtml(out.previewBody ?? out.text);
    else if (out.lang === "markdown") previewHtml = C.wrapPreviewHtml(C.markdownToHtmlFragment(out.text));
    else previewHtml = null;

    if (previewHtml) {
      const frame = document.createElement("iframe");
      frame.setAttribute("sandbox", "");
      frame.setAttribute("referrerpolicy", "no-referrer");
      frame.srcdoc = previewHtml;
      preview.appendChild(frame);
    } else {
      const pre = document.createElement("pre");
      pre.className = "enc-preview";
      pre.style.maxHeight = "420px";
      pre.style.height = "420px";
      pre.textContent = out.text;
      preview.appendChild(pre);
    }

    $("sourcePre").textContent = out.text || "";
    $("tabs").classList.remove("hidden");
    $("copyBtn").classList.remove("hidden");
    switchTab("preview");
  }

  $("dlLabel").textContent = (out.outExt || "文件").toUpperCase();
}

function switchTab(name) {
  const isSource = name === "source";
  for (const tab of $("tabs").querySelectorAll(".tab")) {
    tab.classList.toggle("active", tab.dataset.tab === name);
  }
  $("previewBox").classList.toggle("hidden", isSource);
  $("sourceBox").classList.toggle("hidden", !isSource);
}

/* ---------------- 下载 / 复制 ---------------- */

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}

function ensureResultBlob() {
  if (!result) return null;
  if (result.blob) return result.blob;
  return new Blob([result.text || ""], { type: result.mime || "text/plain;charset=utf-8" });
}

/* ---------------- 错误提示 ---------------- */

function showError(msg) {
  $("errBox").classList.remove("hidden");
  $("errBox").textContent = msg;
}

function hideError() {
  $("errBox").classList.add("hidden");
  $("errBox").textContent = "";
}

/* ---------------- 编码工具 ---------------- */

let encState = null;

async function handleEncodingFile(file) {
  if (!file) return;
  const buffer = await file.arrayBuffer();
  const decoded = C.decodeSmart(buffer);
  encState = { ...decoded, name: file.name, size: file.size };

  $("encOut").classList.remove("hidden");
  const meta = $("encMeta");
  meta.textContent = "";
  const chip = document.createElement("span");
  chip.className = "chip-file";
  chip.textContent = file.name;
  const tag = document.createElement("span");
  tag.className = encState.confident ? "chip-ok" : "chip-tag";
  tag.textContent = encState.confident ? `识别为 ${encState.encoding}` : `推测 ${encState.encoding}（请核对）`;
  const info = document.createElement("span");
  info.textContent = `${fmtBytes(file.size)} · ${encState.text.length} 字符`;
  meta.append(chip, tag, info);

  $("encPreview").textContent = encState.text.slice(0, 4000) + (encState.text.length > 4000 ? "\n\n…（已截断预览）" : "");
}

/* ---------------- 绑定 ---------------- */

function bind() {
  $("convertBtn").addEventListener("click", () => runConversion());
  $("clearBtn").addEventListener("click", () => resetSource());

  $("drop").addEventListener("click", () => $("fileInput").click());
  $("drop").addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); $("fileInput").click(); }
  });
  $("fileInput").addEventListener("change", (e) => {
    const file = e.target.files && e.target.files[0];
    if (file) acceptFile(file);
  });

  for (const evt of ["dragenter", "dragover"]) {
    $("drop").addEventListener(evt, (e) => { e.preventDefault(); $("drop").classList.add("dragging"); });
  }
  for (const evt of ["dragleave", "drop"]) {
    $("drop").addEventListener(evt, (e) => { e.preventDefault(); $("drop").classList.remove("dragging"); });
  }
  $("drop").addEventListener("drop", (e) => {
    const file = e.dataTransfer?.files?.[0];
    if (file) acceptFile(file);
  });
  window.addEventListener("dragover", (e) => e.preventDefault());
  window.addEventListener("drop", (e) => {
    if (!$("drop").contains(e.target)) e.preventDefault();
  });

  $("pasteArea").addEventListener("input", (e) => acceptPaste(e.target.value));
  $("pasteClear").addEventListener("click", () => {
    $("pasteArea").value = "";
    if (source && source.kind === "text") resetSource();
  });

  for (const tab of $("tabs").querySelectorAll(".tab")) {
    tab.addEventListener("click", () => switchTab(tab.dataset.tab));
  }

  $("dlBtn").addEventListener("click", () => {
    const blob = ensureResultBlob();
    if (!blob || !result) return;
    downloadBlob(blob, result.filename || `docshift.${result.outExt || "txt"}`);
  });

  $("copyBtn").addEventListener("click", async () => {
    if (!result?.text) return;
    try {
      await navigator.clipboard.writeText(result.text);
      $("copyBtn").textContent = "已复制";
      setTimeout(() => { $("copyBtn").textContent = "复制源码"; }, 1400);
    } catch {
      showError("浏览器拒绝了剪贴板访问，请用「下载」获取文件。");
    }
  });

  // 编码工具
  const encDrop = $("encDrop");
  encDrop.addEventListener("click", () => $("encInput").click());
  encDrop.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); $("encInput").click(); }
  });
  $("encInput").addEventListener("change", (e) => {
    const f = e.target.files && e.target.files[0];
    if (f) handleEncodingFile(f);
  });
  for (const evt of ["dragenter", "dragover"]) {
    encDrop.addEventListener(evt, (e) => { e.preventDefault(); encDrop.classList.add("dragging"); });
  }
  for (const evt of ["dragleave", "drop"]) {
    encDrop.addEventListener(evt, (e) => { e.preventDefault(); encDrop.classList.remove("dragging"); });
  }
  encDrop.addEventListener("drop", (e) => {
    const f = e.dataTransfer?.files?.[0];
    if (f) handleEncodingFile(f);
  });

  $("encSave").addEventListener("click", () => {
    if (!encState) return;
    downloadBlob(C.textToBlob(encState.text, false), `${stripExt(encState.name)}-utf8.txt`);
  });
  $("encSaveBom").addEventListener("click", () => {
    if (!encState) return;
    downloadBlob(C.textToBlob(encState.text, true), `${stripExt(encState.name)}-utf8-bom.txt`);
  });
  $("encCopy").addEventListener("click", async () => {
    if (!encState) return;
    try {
      await navigator.clipboard.writeText(encState.text);
      $("encCopy").textContent = "已复制";
      setTimeout(() => { $("encCopy").textContent = "复制文本"; }, 1400);
    } catch { /* 忽略 */ }
  });
}

/* ---------------- 启动 ---------------- */

initChrome();
renderConvList();
bind();
selectConversion(CONVERSIONS[0]);
