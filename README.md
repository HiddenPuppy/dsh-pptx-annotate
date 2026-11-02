# dsh-pptx-annotate

> **在预览面板里逐元素批注 PPT，一键把整轮修改意见交给 Agent。**
> Annotate PPTX element by element inside the DSH preview panel, then hand the whole review to the agent in one click.

[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![DSH plugin](https://img.shields.io/badge/DSH-plugin-4A6CF7)](https://github.com/topics/dsh-plugin)
[![DSH version](https://img.shields.io/badge/DSH-0.2.1--rc.x-lightgrey)](https://github.com/deepseek-ai/deepseek-harness)
[![stars](https://img.shields.io/github/stars/HiddenPuppy/dsh-pptx-annotate)](https://github.com/HiddenPuppy/dsh-pptx-annotate/stargazers)

The built-in preview lets you *look* at a `.pptx`. This plugin lets you **talk to it**: click any
element on any slide, write what should change, and hand the whole review to the DSH agent in one
click. The agent gets `page 2 · #3 · "the heading" · what to change` instead of a paragraph of
prose describing where to look. Registered at `priority: 'extension'`, so it takes over `.pptx`
by default and the built-in PDF preview stays one dropdown away.

一个 DSH（DeepSeek Harness）文档预览器插件。它和内置的 Office 预览**并列**注册在同一个位置，
但把「看 PPT」变成了「改 PPT」：点任意元素写意见 → 一键提交 → Agent 逐条改 → 面板自动刷新。

![预览 → 批注](docs/demo.gif)

<sub>点任意元素 → 写意见 → 一键交给 Agent（上方 GIF 为真实界面录制的两帧：无批注 / 已批注）</sub>

> 想先看设计拆解与功能对比：[简介](docs/intro.zh.md)

---

## 它解决什么问题

DSH 内置的 Office 预览是**只读**的：`.pptx` 被转成 PDF 渲染，你能看，但不能对某一页的某个标题说话。

于是改一份 PPT 的真实流程是：看一眼 → 记住「第 2 页那个大标题」→ 切到对话 → 用文字描述位置 →
Agent 猜你指的是哪个元素 → 改错 → 再来一轮。

这个插件把「描述位置」这一步去掉了。页面上每个形状都是可点的热区，意见直接挂在元素上，
提交时变成一条结构化消息，Agent 拿到的是「第 2 页 · #3 · 三个工具，三种完全不同的结局：字号太大」。

## 和内置预览的差别

| 能力 | DSH 内置 Office 预览 | dsh-pptx-annotate |
|---|---|---|
| 打开 `.pptx` | ✅ 转 PDF 只读 | ✅ 直接渲染每页 |
| 选中单个元素 | ❌ 只有一整页 | ✅ 解析形状树，含组合展开与占位符继承 |
| 逐元素批注 | ❌ | ✅ 点元素 → 写意见 → 编号标记 |
| 批注回到对话 | ❌ 手打描述 | ✅ 一键提交为结构化 prompt |
| 改完回看 | ❌ 手动刷新 | ✅ 轮询文件版本，自动换图 |
| 目录 / 缩略图 | ❌ | ✅ 面板够宽时自动出现，可滚动跳转 |
| 提交进度 | ❌ | ✅ 输入框旁显示「N 条批注处理中…」 |
| 随时退回 | — | ✅ 工具栏下拉切回内置预览 |

**关键不是「多了几个功能」，而是它把 Agent 的修改能力接进了一个可视化界面。**
没有它，Agent 也能改 PPT，但你得用文字把位置说清楚；有了它，位置由点击决定，歧义消失。

## 一次批注的闭环

```text
① 右侧预览里点元素          → 元素高亮 + 编号
② 输入意见                  → 「标题太大，缩到 40pt」
③ 点「发送给 Agent」        → 生成结构化条目，投进当前会话
                              左侧输入框出现「N 条批注处理中…」
④ Agent 逐条改              → 每条回一行 ✅ 第 X 页 · 元素：改了什么
⑤ 文件一变，插件轮询到      → 预览自动换图，进度提示消失
```

提交的消息长这样，Agent 拿到的是可定位的清单而不是一段散文：

```text
【PPT 批注】`deck.pptx` · 共 3 条

- 第 2 页 · #3 · 三个工具，三种完全不同的结局：标题太大，缩到 40pt
- 第 3 页 · #8 · 表格：换成三列卡片
- 第 1 页 · #6 · 2026.10 · 内部技术分享：字号太小

处理约定（按这个顺序做，不要扩大范围）：
1. 先改后说。单点小改直接动手，不要先给方案、不要长篇分析；改完再汇报。
2. 只改批注点到的元素，其他元素一律不动。
3. 每改完一条，立刻用一行 ✅ 第 X 页 · 元素：改了什么 回报。
4. 汇报之后另起一段，给 1–3 条「可选优化建议」，标明是建议、尚未执行。
```

## 安装

插件是标准的 DSH bundle（一条 Loader 行同时带来 Host 与 Client 两面）。

```bash
# 方式一：dsh 自带的插件安装命令（会写入 profile 的 bundles）
dsh plugin --profile web add /absolute/path/to/dsh-pptx-annotate

# 方式二：手工 link（桌面版的 desktop profile 只能这样装：
#        公开 CLI 会拒绝管理 desktop profile）
cd ~/.dsh/profiles/desktop
ln -sfn /absolute/path/to/dsh-pptx-annotate node_modules/dsh-pptx-annotate
# 再把 "dsh-pptx-annotate": "link:/absolute/path/to/dsh-pptx-annotate"
# 加进 dependencies，并把包名加进 dsh.profile.bundles
```

改完 `package.json` 需要重启应用；之后只改 `lib/client.js` 会热替换，改 `lib/index.js` 需要触发一次 profile 重组
（随便存一下 `cordis.patch.yml`）。

卸载：从 `dsh.profile.bundles` 移除包名、删掉软链即可，内置预览不受影响。

## 它是怎么工作的

```
dsh-pptx-annotate/
├─ lib/index.js      Host 半边：两条鉴权路由
│                      GET /api/pptx-annotate/deck     形状树 → 每页元素盒（归一化 0..1）
│                      GET /api/pptx-annotate/slide    该页 PNG（LibreOffice Kit 渲染，按文件版本缓存）
│                      GET /api/pptx-annotate/version  仅 stat，用于判断「Agent 改完了没有」
├─ lib/client.js     Client 半边：注册 documentPreviews（pptx, priority=extension）
│                      → 挂 body 到 sidebar.right.tab.document
│                      → 挂进度条到 conversation.composer.dock
├─ src/zip.js        纯 JS 读 ZIP 中央目录（只支持 stored/deflate）
├─ src/xml.js        轻量 XML 读取（PPTX 是机器生成的良构 XML）
├─ src/pptx-geom.js  形状树 → 归一化盒；组合变换展开；占位符从 layout/master 继承位置
└─ src/render.js     子进程调用捆绑的 LibreOffice Kit 渲染 PNG
```

几个刻意的设计选择：

- **不打包任何依赖**：`client.js` 手写成 DSH 的 `window.__ModuleLoader__.load({ factory })` 格式，
  只 `require("react")` 这一个平台种子；Host 侧只用 `node:zlib` / `node:child_process`。
- **几何在 Host 算、UI 在 Client 画**：浏览器只拿到「第几页、哪些盒、什么文字」，
  不需要理解 OOXML，也不需要把 pptx 传进浏览器。
- **命中规则是「文字优先、面积最小优先」**：点大标题时不会选中它背后的装饰框。
- **刷新靠文件版本而不是超时**：图片 URL 带 `&v=<mtime>`，文件一变 URL 必变，绕开浏览器缓存。

## 测试

```bash
npm test     # 35 个单元 + 集成用例（node:test）
npm run smoke  # 真实 client bundle 的浏览器端到端，900px / 500px 两个宽度
```

覆盖：ZIP/XML 解析、形状几何（组合展开、占位符继承）、命中优先级、prompt 序列化、
Host 路由（含鉴权失败路径）、真实 PNG 渲染、组件挂载与点击流、提交路径的失败处理。

## 已知限制

- **只接管 `.pptx`**，`.ppt` / `.docx` 仍走内置预览。
- **首次打开约 10 秒**（LibreOffice 冷启动），之后同文件毫秒级。
- **批注不落盘**，发送后清空——刻意为之，避免下次打开看到过期意见。
- **文本框溢出**：几何来自 shape 的声明尺寸，文字跑出文本框时热区会略小于视觉范围。
- **进度提示的结束信号是文件变化**：若你只提问不改稿，提示会到 10 分钟超时。
- 依赖 DSH `0.2.1-rc.*` 的内部 slot 契约，升级大版本可能需要跟着调。

## Roadmap

- [ ] 批注持久化（`.dsh-annotations.json`），跨会话保留、可勾选已解决
- [ ] 逐条「接受 / 忽略」，把 Agent 的改动做成可审阅的 diff
- [ ] 支持 `.pptx` 之外：`.docx` 段落级批注
- [ ] 用会话 turn 状态替代文件轮询，进度提示更准

## 相关

- DSH 插件话题：[github.com/topics/dsh-plugin](https://github.com/topics/dsh-plugin)
- 社区讨论帖：[deepseek-ai/deepseek-harness#11564](https://github.com/deepseek-ai/deepseek-harness/discussions/11564)
- 同类项目集合：[awesome-dsh-plugin](https://github.com/awesome-dsh-plugin/awesome-dsh-plugin)

## License

MIT
