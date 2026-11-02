<!-- 提交到 awesome-dsh-plugin 的条目（Docs & Rendering 分类，按作者名字母序插在 HiddenPuppy 位置） -->

- [HiddenPuppy/dsh-pptx-annotate](https://github.com/HiddenPuppy/dsh-pptx-annotate) - Element-level PPTX annotation in the DSH preview panel: click any shape on any slide, write the change, and submit the whole review as one structured prompt addressed by page and element; the render refreshes by itself when the agent saves the file.

<!--
提交前检查（对应 awesome 列表的收录要求）：
1. `dsh plugin --profile web add /path/to/dsh-pptx-annotate` 能装上 → 已验证包声明了 dsh.bundle.patch
2. 描述里每个能力都能在源码里指出来：
   - "element-level" → src/pptx-geom.js 解析形状树
   - "submit as one structured prompt" → lib/client.js formatPrompt + conversation.send
   - "refreshes by itself" → /api/pptx-annotate/version 轮询 + &v=<mtime> 缓存击穿
3. 分类：Docs & Rendering（与 dsh-md-annotator 同组）
-->
