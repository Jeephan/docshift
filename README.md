# DocShift · 文档格式转换台

在浏览器里完成文档格式转换的小工具：**Word / Markdown / HTML / 纯文本互转**，附带文本乱码修复。

没有后端、没有登录、没有上传队列 —— 文件从你的硬盘进内存，转换完直接下载，全程不会离开这台电脑。

![DocShift 首页](docs/screenshot.png)

## 为什么做成纯前端

绝大多数在线转换站要把文件传到服务器，再排队等结果。文档转换里其实有一大半是**纯计算**：解析 OOXML、渲染 Markdown、生成 HTML。浏览器完全扛得住，于是就有了这个方案：

- **零成本**：一个静态站点，托管在 Cloudflare Pages 免费额度上，跑多久都是 0 元
- **零隐私风险**：页面里没有任何接收文件的服务端接口，断网也能用
- **零维护负担**：没有服务器、没有数据库，不用担心磁盘写满或 worker 挂掉

## 功能

| 转换 | 说明 |
| --- | --- |
| Word (.docx) → HTML | 保留标题层级、列表、表格、图片、超链接 |
| Word (.docx) → Markdown | 适合迁到 Notion / Obsidian / GitHub Wiki |
| Word (.docx) → 纯文本 | 只留文字，方便检索或喂给翻译工具 |
| Markdown → HTML | 输出可独立分发的单页 HTML（样式已内联） |
| Markdown → Word (.docx) | 标准 OOXML，Word / WPS / Pages 都能打开并二次编辑 |
| Markdown → 纯文本 | 去掉所有标记 |
| HTML → Markdown | 网页或富文本回迁到 Markdown |
| HTML → Word (.docx) | 把网页内容整理成可编辑文档 |
| 文本编码修复 | 自动识别 UTF-8 / GBK / UTF-16（含 BOM），一键转成 UTF-8 |

同时还提供：拖拽上传、直接粘贴文本、预览 / 源码切换、一键下载与复制。

## 本地运行

需要 Node.js 18+。

```bash
npm install
npm run dev      # 构建并在 http://127.0.0.1:4173 打开
npm run build    # 只构建，产物在 dist/
```

`dist/` 是完全的静态目录，丢进任何静态托管都能直接跑。

## 部署

### Cloudflare Pages（推荐，免费）

在 Cloudflare 控制台 → Workers & Pages → Create → Pages → Connect to Git，选中本仓库后填：

- Framework preset: `None`
- Build command: `npm run build`
- Build output directory: `dist`
- Node version: 18 或以上（环境变量 `NODE_VERSION=20` 可选）

之后每次 `git push` 都会自动构建发布。

命令行方式（不接 Git）：

```bash
npm run build
npx wrangler pages deploy dist --project-name=docshift
```

### GitHub Pages

```bash
npm run build
# 把 dist/ 推到 gh-pages 分支，或在仓库设置里选择 /docs 目录
```

因为构建产物不带任何绝对路径，子路径部署（如 `/docshift/`）也能正常工作。

### 内网 / 私有部署

```bash
npm run build
# 把 dist/ 拷到 Nginx 的 html 目录即可
```

## 技术栈

| 库 | 用途 | 协议 |
| --- | --- | --- |
| [mammoth](https://github.com/mwilliamson/mammoth.js) | 解析 .docx，输出语义化 HTML | BSD-2 |
| [marked](https://github.com/markedjs/marked) | Markdown → HTML | MIT |
| [turndown](https://github.com/mixmark-io/turndown) + [turndown-plugin-gfm](https://github.com/mixmark-io/turndown-plugin-gfm) | HTML → Markdown，支持表格 / 删除线 / 任务列表 | MIT |
| [docx](https://github.com/dolanmiu/docx) | 生成标准 OOXML 文档 | MIT |
| [esbuild](https://esbuild.github.io/) | 打包与压缩 | MIT |

所有依赖都打包进 `dist/app.js`，不依赖任何外部 CDN，部署后可完全离线运行。

## 工作原理

1. **读取**：`file.arrayBuffer()` 把文件读进内存，不经过任何网络。
2. **解析**：`.docx` 本质是一个 zip，mammoth 解出 `document.xml` 并转成带语义的 HTML（标题、列表、表格、图片为内嵌 base64）。
3. **转换**：HTML 与 Markdown 之间用 turndown / marked 互转；生成 Word 时用 docx 库逐块构造 OOXML（标题样式、编号列表、表格边框、等宽代码块）。
4. **导出**：`Blob` + `URL.createObjectURL` 触发浏览器下载，对象 URL 在几秒后自动释放。

## 目录结构

```
docshift/
├── index.html            # 页面结构（含无障碍标签与暗色模式适配）
├── src/
│   ├── app.js            # 交互逻辑：转换清单、拖拽、预览、下载
│   ├── convert.js        # 转换核心：docx / html / markdown / 编码
│   ├── docx-build.js     # HTML → OOXML 构建器
│   ├── styles.css        # 样式（含 prefers-color-scheme 暗色）
│   └── shims/buffer.js   # docx 库 Node 分支的占位模块
├── build.mjs             # esbuild 构建脚本
├── public/favicon.svg
└── docs/                 # 截图
```

## 已知限制

- `.docx` 的**版式信息**（页眉页脚、页码、分栏、文本框、艺术字、复杂公式）不会保留，只保留语义结构。这是所有 docx→HTML 方案的共同限制。
- Word → Markdown 时，内嵌图片会替换成 `images/xxx.png` 相对路径占位，避免 Markdown 文件被 base64 撑到几十 MB；需要图片请走 Word → HTML。
- 建议单个文件小于 40 MB。超大文件或含大量高清图片时会占用较多内存。
- 当前版本只做文档类格式；PDF、图片、音视频转换不在范围内。

## 贡献

Issue / PR 都欢迎。改动转换逻辑时建议先在 `src/convert.js` 里加一个最小用例，确认 `npm run build` 通过再提 PR。

## License

[MIT](LICENSE)
