window.__ModuleLoader__.load({
	id: "dsh-pptx-annotate",
	factory: (require) => {
		var module = { exports: {} };
		var exports = module.exports;
		Object.defineProperty(exports, Symbol.toStringTag, { value: "Module" });
		const React = require("react");
		const h = React.createElement;

		const ID = "dsh-pptx-annotate";
		const inject = ["slots", "documentPreviews", "sessions"];
		const TEXT_KINDS = { title: 1, text: 1, table: 1 };

		const CSS = `
.pa-root{container-type:inline-size;display:flex;flex-direction:column;height:100%;min-width:0;min-height:0;background:var(--dsw-alias-bg-base,#fff);color:var(--dsw-alias-text-primary,#101828);font-size:13px;line-height:1.5}
.pa-bar{display:flex;align-items:center;gap:8px;padding:8px 12px;border-bottom:1px solid var(--dsw-alias-border-subtle,#e4e8f0);flex:0 0 auto;min-width:0}
.pa-name{flex:1 1 auto;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:var(--dsw-alias-text-secondary,#667085);font-size:12px}
.pa-count{flex:0 0 auto;font-size:11px;color:#2F6FED;background:rgba(47,111,237,.1);border-radius:9px;padding:1px 8px}
.pa-btn{flex:0 0 auto;border:1px solid var(--dsw-alias-border-subtle,#e4e8f0);background:transparent;color:inherit;border-radius:6px;padding:4px 10px;font:inherit;font-size:12px;cursor:pointer}
.pa-btn:hover{background:rgba(16,24,40,.04)}
.pa-btn.primary{background:#2F6FED;border-color:#2F6FED;color:#fff}
.pa-btn.primary:disabled{opacity:.45;cursor:default}
.pa-body{display:flex;flex:1 1 auto;min-width:0;min-height:0}
.pa-outline{display:none;flex:0 0 88px;flex-direction:column;gap:6px;overflow-y:auto;overflow-x:hidden;padding:8px 6px;border-right:1px solid var(--dsw-alias-border-subtle,#e4e8f0);background:#fafbfd}
@container (min-width:520px){.pa-outline{display:flex}}
.pa-thumb{margin:0;border:1px solid transparent;border-radius:4px;padding:2px;cursor:pointer;background:transparent;flex:0 0 auto;font:inherit}
.pa-thumb.on{border-color:#2F6FED;background:rgba(47,111,237,.07)}
.pa-thumb img{display:block;width:100%;height:auto;border-radius:2px;background:#fff}
.pa-thumb b{display:block;text-align:center;font-size:10px;font-weight:500;color:#667085;line-height:15px}
.pa-main{flex:1 1 auto;min-width:0;overflow-y:auto;overflow-x:hidden;padding:12px;background:#eef1f6;scroll-behavior:smooth}
.pa-slide{position:relative;width:100%;max-width:1000px;margin:0 auto 14px;background:#fff;border-radius:2px;box-shadow:0 2px 10px rgba(16,24,40,.12);overflow:hidden}
.pa-slide img{position:absolute;inset:0;width:100%;height:100%;object-fit:contain;display:block;user-select:none}
.pa-hot{position:absolute;border:1.5px solid transparent;border-radius:2px;cursor:pointer}
.pa-hot:hover{border-color:rgba(47,111,237,.65);background:rgba(47,111,237,.08)}
.pa-hot.on{border-color:#2F6FED;background:rgba(47,111,237,.13);box-shadow:0 0 0 2px rgba(47,111,237,.22)}
.pa-badge{position:absolute;top:-8px;left:-8px;min-width:16px;height:16px;border-radius:8px;background:#2F6FED;color:#fff;font-size:10px;line-height:16px;text-align:center;padding:0 3px;display:none;font-weight:600}
.pa-hot.on .pa-badge{display:block}
.pa-notes{flex:0 0 auto;max-height:38%;overflow-y:auto;border-top:1px solid var(--dsw-alias-border-subtle,#e4e8f0);padding:8px 12px 12px;min-width:0}
.pa-note{border:1px solid var(--dsw-alias-border-subtle,#e4e8f0);border-radius:8px;padding:6px 8px;margin-top:8px;background:rgba(16,24,40,.02)}
.pa-note.on{border-color:#2F6FED}
.pa-meta{display:flex;justify-content:space-between;gap:8px;font-size:11px;color:var(--dsw-alias-text-secondary,#667085);min-width:0}
.pa-meta em{font-style:normal;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;min-width:0}
.pa-meta b{color:#2F6FED;font-weight:600}
.pa-del{flex:0 0 auto;cursor:pointer;color:#B42318}
.pa-note textarea{width:100%;box-sizing:border-box;margin-top:4px;border:1px solid var(--dsw-alias-border-subtle,#e4e8f0);border-radius:6px;padding:5px 6px;font:inherit;font-size:12px;resize:vertical;min-height:44px;background:#fff;color:inherit}
.pa-empty{color:var(--dsw-alias-text-secondary,#667085);font-size:12px;text-align:center;padding:18px 12px}
.pa-state{padding:24px 16px;text-align:center;color:var(--dsw-alias-text-secondary,#667085);font-size:12px;line-height:1.7}
.pa-err{color:#B42318}
.pa-activity{display:inline-flex;align-items:center;gap:6px;font-size:12px;color:#2F6FED;white-space:nowrap;padding:0 4px}
.pa-dot{width:7px;height:7px;border-radius:50%;background:#2F6FED;animation:pa-pulse 1.2s ease-in-out infinite}
@keyframes pa-pulse{0%,100%{opacity:.35;transform:scale(.85)}50%{opacity:1;transform:scale(1.15)}}
`;

		// Plugin-owned stylesheet: tagged so HMR can drop it when this bundle is replaced.
		if (typeof document !== "undefined") {
			const tagId = "dsh-pptx-annotate/preview.css";
			if (document.querySelector('style[data-plugin-css="' + tagId + '"]') === null) {
				const tag = document.createElement("style");
				tag.dataset.plugin = "dsh-pptx-annotate";
				tag.dataset.pluginCss = tagId;
				tag.textContent = CSS;
				document.head.appendChild(tag);
			}
		}

		// Which sessions have annotations in flight: the composer chip reads this.
		const pending = new Map();
		const listeners = new Set();
		const pendingTimers = new Map();

		/**
		 * Mark a session as having annotations being processed.
		 * @param sessionId - owning session
		 * @param count - number of annotations submitted
		 */
		function markPending(sessionId, count) {
			pending.set(sessionId, count);
			for (const notify of listeners) notify();
			if (pendingTimers.has(sessionId)) clearTimeout(pendingTimers.get(sessionId));
			pendingTimers.set(sessionId, setTimeout(() => clearPending(sessionId), 10 * 60 * 1000));
		}

		/**
		 * Drop the processing marker for a session.
		 * @param sessionId - owning session
		 */
		function clearPending(sessionId) {
			if (!pending.delete(sessionId)) return;
			const timer = pendingTimers.get(sessionId);
			if (timer !== undefined) {
				clearTimeout(timer);
				pendingTimers.delete(sessionId);
			}
			for (const notify of listeners) notify();
		}

		/**
		 * Subscribe to the pending map.
		 * @param listener - callback
		 * @returns unsubscribe
		 */
		function subscribePending(listener) {
			listeners.add(listener);
			return () => listeners.delete(listener);
		}

		/** Composer chip: how many annotations this session is still processing. */
		function AnnotateActivity(props) {
			const sessionId = props.sessionId;
			const [count, setCount] = React.useState(() => pending.get(sessionId) || 0);
			React.useEffect(() => subscribePending(() => setCount(pending.get(sessionId) || 0)), [sessionId]);
			if (!count) return null;
			return h(
				"div",
				{ className: "pa-activity", "data-pptx-annotate-activity": true },
				h("span", { className: "pa-dot" }),
				count + " 条批注处理中…",
			);
		}

		/** Collapse whitespace. */
		function tidy(text) {
			return String(text == null ? "" : text).replace(/\s+/g, " ").trim();
		}

		/** Human label for one element. */
		function labelOf(shape, index) {
			const text = tidy(shape && shape.text);
			if (text) return text.length > 36 ? text.slice(0, 36) + "…" : text;
			const names = {
				picture: "图片",
				table: "表格",
				chart: "图表",
				frame: "图形框",
				connector: "连接线",
				title: "标题",
				text: "文本框",
				shape: "形状",
			};
			return names[shape && shape.kind] || "元素 " + (index + 1);
		}

		/**
		 * Pick the element under a point: text-bearing shapes win over decoration,
		 * and the smallest such shape wins, so a heading beats the box behind it.
		 */
		function pickAt(shapes, x, y, slack) {
			const pad = slack == null ? 0.004 : slack;
			const hits = [];
			for (let i = 0; i < shapes.length; i += 1) {
				const s = shapes[i];
				if (x >= s.x - pad && x <= s.x + s.w + pad && y >= s.y - pad && y <= s.y + s.h + pad) {
					hits.push({ shape: s, index: i, area: s.w * s.h });
				}
			}
			if (!hits.length) return null;
			const textHits = hits.filter((hit) => TEXT_KINDS[hit.shape.kind] || tidy(hit.shape.text));
			const pool = textHits.length ? textHits : hits;
			pool.sort((a, b) => a.area - b.area);
			return { shape: pool[0].shape, index: pool[0].index };
		}

		/** Build the prompt sent to the agent. */
		function formatPrompt(file, notes) {
			const rows = notes
				.filter((n) => tidy(n.note))
				.sort((a, b) => a.slide - b.slide || a.order - b.order)
				.map((n, i) => `- 第 ${n.slide} 页 · ${n.elementId} · ${n.label}：${tidy(n.note)}`);
			if (!rows.length) return "";
			return [
				"【PPT 批注】`" + file + "` · 共 " + rows.length + " 条",
				"",
				...rows,
				"",
				"处理约定（按这个顺序做，不要扩大范围）：",
				"1. 先改后说。单点小改（文字、字号、颜色、位置）直接动手，不要先给方案、不要长篇分析；改完再汇报。",
				"2. 只改批注点到的元素，其他元素一律不动。",
				"3. 每改完一条，立刻用一行 `✅ 第 X 页 · 元素：改了什么` 回报；全部改完只做一次渲染复核。",
				"4. 汇报之后另起一段，给 1–3 条「可选优化建议」，标明是建议、尚未执行。",
			].join("\n");
		}

		function Slide(props) {
			const { slide, path, notes, onPick, ratio } = props;
			const url = "/api/pptx-annotate/slide?path=" + encodeURIComponent(path) + "&n=" + slide.index + "&v=" + props.stamp;
			return h(
				"div",
				{
					className: "pa-slide",
					"data-slide": slide.index,
					style: { aspectRatio: ratio },
				},
				h("img", { src: url, alt: "第 " + slide.index + " 页", draggable: false, loading: "lazy" }),
				...slide.shapes.map((shape, index) => {
					const key = slide.index + "-" + index;
					const note = notes.find((n) => n.key === key);
					return h(
						"div",
						{
							key,
							className: "pa-hot" + (note ? " on" : ""),
							style: {
								left: shape.x * 100 + "%",
								top: shape.y * 100 + "%",
								width: shape.w * 100 + "%",
								height: shape.h * 100 + "%",
							},
							title: labelOf(shape, index),
							onClick: (event) => {
								event.stopPropagation();
								onPick(slide.index, index, shape);
							},
						},
						note ? h("span", { className: "pa-badge" }, String(note.order)) : null,
					);
				}),
			);
		}

		function Outline(props) {
			const { slides, active, path, onGo } = props;
			return h(
				"nav",
				{ className: "pa-outline", "data-pa-outline": true },
				...slides.map((slide) =>
					h(
						"button",
						{
							key: slide.index,
							type: "button",
							className: "pa-thumb" + (slide.index === active ? " on" : ""),
							onClick: () => onGo(slide.index),
							title: "第 " + slide.index + " 页",
						},
						h("img", {
							src: "/api/pptx-annotate/slide?path=" + encodeURIComponent(path) + "&n=" + slide.index,
							alt: "",
							loading: "lazy",
						}),
						h("b", null, String(slide.index)),
					),
				),
			);
		}

		function AnnotateBody(props) {
			const { content, resourceAddress, useTabInfo, useResource, onSubmit, sessionId } = props;
			const request = content && content.kind === "renderer" ? content : undefined;
			const revision = request && request.revision;
			const tab = useTabInfo().tab;
			const meta = useResource(resourceAddress);
			const absolutePath = meta && meta.value && meta.value.absolutePath;

			const [deck, setDeck] = React.useState(null);
			const [phase, setPhase] = React.useState(absolutePath ? "loading" : "waiting");
			const [error, setError] = React.useState("");
			const [notes, setNotes] = React.useState([]);
			const [activeKey, setActiveKey] = React.useState("");
			const [activeSlide, setActiveSlide] = React.useState(1);
			const [sent, setSent] = React.useState("");
			const [busy, setBusy] = React.useState(false);
			const [stamp, setStamp] = React.useState(0);
			const [reloadSeq, setReloadSeq] = React.useState(0);
			const [watching, setWatching] = React.useState(false);
			const mainRef = React.useRef(null);
			const versionRef = React.useRef(0);

			React.useEffect(() => {
				if (!absolutePath) {
					setPhase("waiting");
					return undefined;
				}
				let cancelled = false;
				setPhase("loading");
				setError("");
				fetch("/api/pptx-annotate/deck?path=" + encodeURIComponent(absolutePath), {
					credentials: "same-origin",
					headers: { accept: "application/json" },
				})
					.then((response) =>
						response.json().then((body) => {
							if (!response.ok) throw new Error((body && body.error && body.error.message) || "HTTP " + response.status);
							return body;
						}),
					)
					.then((body) => {
						if (cancelled) return;
						setDeck(body);
						versionRef.current = body.version || 0;
						setStamp(body.version || 0);
						setPhase("ready");
						if (request && request.loaded) request.loaded(1);
					})
					.catch((cause) => {
						if (cancelled) return;
						setError(cause && cause.message ? cause.message : String(cause));
						setPhase("error");
						if (request && request.failed) request.failed();
					});
				return () => {
					cancelled = true;
				};
			}, [absolutePath, revision, reloadSeq]);

			React.useEffect(() => {
				const node = mainRef.current;
				if (!node || !deck || typeof IntersectionObserver === "undefined") return undefined;
				const observer = new IntersectionObserver(
					(entries) => {
						let best = null;
						for (const entry of entries) {
							if (!entry.isIntersecting) continue;
							if (!best || entry.intersectionRatio > best.intersectionRatio) best = entry;
						}
						if (best) {
							const n = Number(best.target.getAttribute("data-slide"));
							if (n) setActiveSlide(n);
						}
					},
					{ root: node, threshold: [0.25, 0.5, 0.75] },
				);
				const nodes = node.querySelectorAll("[data-slide]");
				for (let i = 0; i < nodes.length; i += 1) observer.observe(nodes[i]);
				return () => observer.disconnect();
			}, [deck]);

			React.useEffect(() => {
				if (!watching || !absolutePath) return undefined;
				let ticks = 0;
				const timer = setInterval(() => {
					ticks += 1;
					if (ticks > 300) {
						setWatching(false);
						return;
					}
					fetch("/api/pptx-annotate/version?path=" + encodeURIComponent(absolutePath), {
						credentials: "same-origin",
						headers: { accept: "application/json" },
					})
						.then((r) => r.json())
						.then((body) => {
							if (body && body.version && body.version !== versionRef.current) {
								versionRef.current = body.version;
								setWatching(false);
								setReloadSeq((n) => n + 1);
								setSent("已更新 · 已载入最新效果");
								if (sessionId !== undefined) clearPending(sessionId);
							}
						})
						.catch(() => {});
				}, 2000);
				return () => clearInterval(timer);
			}, [watching, absolutePath]);

			const file = absolutePath ? absolutePath.split("/").pop() : "";

			const removeNote = (key) => {
				setNotes((current) => current.filter((n) => n.key !== key));
				setActiveKey((current) => (current === key ? "" : current));
			};

			const pick = (slideIndex, index, shape) => {
				const key = slideIndex + "-" + index;
				const existing = notes.find((n) => n.key === key);
				if (existing) {
					// Second click on the same element: clear an empty note, keep a written one.
					if (tidy(existing.note)) {
						setActiveKey(key);
						return;
					}
					removeNote(key);
					return;
				}
				setActiveKey(key);
				setSent("");
				setNotes((current) =>
					current.concat({
						key,
						slide: slideIndex,
						elementId: "#" + (index + 1),
						label: labelOf(shape, index),
						note: "",
						order: current.length + 1,
					}),
				);
			};

			const setNote = (key, text) => {
				setNotes((current) => current.map((n) => (n.key === key ? Object.assign({}, n, { note: text }) : n)));
			};

			const goTo = (n) => {
				const node = mainRef.current;
				if (!node) return;
				const target = node.querySelector('[data-slide="' + n + '"]');
				if (target && target.scrollIntoView) target.scrollIntoView({ block: "start" });
			};

			const submit = () => {
				const prompt = formatPrompt(file, notes);
				if (!prompt || !onSubmit) return;
				setBusy(true);
				setSent("");
				const watchdog = setTimeout(() => {
					setBusy(false);
					setSent("已提交，等待 Agent 响应…");
				}, 12000);
				new Promise((resolve) => {
					resolve(onSubmit(prompt));
				})
					.then(() => {
						const eta = notes.length <= 1 ? "预计 10–30 秒" : notes.length <= 4 ? "预计 30–60 秒" : "预计 1–2 分钟";
						setNotes([]);
						setActiveKey("");
						setSent("已提交 · " + eta + "，完成后自动刷新");
						setWatching(true);
						if (sessionId !== undefined) markPending(sessionId, filled);
					})
					.catch((cause) => setSent("发送失败：" + (cause && cause.message ? cause.message : String(cause))))
					.finally(() => {
						clearTimeout(watchdog);
						setBusy(false);
					});
			};

			if (phase === "waiting") return h("div", { className: "pa-state" }, "等待文件路径…");
			if (phase === "loading") {
				return h("div", { className: "pa-state" }, "正在解析并渲染幻灯片…", h("br"), "首次打开需要几秒钟");
			}
			if (phase === "error") {
				return h(
					"div",
					{ className: "pa-state" },
					h("div", { className: "pa-err" }, "预览失败：" + error),
					h(
						"button",
						{
							className: "pa-btn",
							style: { marginTop: 12 },
							onClick: () => {
								if (request && request.reload) request.reload();
								else if (absolutePath) setPhase("loading");
							},
						},
						"重试",
					),
				);
			}

			const filled = notes.filter((n) => tidy(n.note)).length;
			const ratio = deck.width + " / " + deck.height;
			return h(
				"div",
				{ className: "pa-root", "data-pptx-annotate": true },
				h(
					"div",
					{ className: "pa-bar" },
					h("span", { className: "pa-name", title: absolutePath }, file),
					notes.length ? h("span", { className: "pa-count" }, notes.length + " 条批注") : null,
					sent ? h("span", { className: "pa-name", style: { flex: "0 0 auto" } }, sent) : null,
					h("button", { className: "pa-btn", onClick: () => goTo(activeSlide), title: "回到当前页" }, "第 " + activeSlide + " 页"),
					h("button", { className: "pa-btn primary", disabled: busy || !filled, onClick: submit }, busy ? "发送中…" : "发送给 Agent"),
				),
				h(
					"div",
					{ className: "pa-body" },
					h(Outline, { slides: deck.slides, active: activeSlide, path: absolutePath, onGo: goTo }),
					h(
						"div",
						{ className: "pa-main", ref: mainRef, "data-pa-main": true },
						...deck.slides.map((slide) =>
							h(Slide, { key: slide.index + "-" + stamp, slide, path: absolutePath, notes, onPick: pick, ratio, stamp }),
						),
					),
				),
				h(
					"div",
					{ className: "pa-notes" },
					notes.length === 0
						? h("div", { className: "pa-empty" }, "点击任意元素添加批注 · 再次点击可取消")
						: h(
								"div",
								null,
								...notes
									.slice()
									.sort((a, b) => a.slide - b.slide || a.order - b.order)
									.map((note) =>
										h(
											"div",
											{ key: note.key, className: "pa-note" + (note.key === activeKey ? " on" : "") },
											h(
												"div",
												{ className: "pa-meta" },
												h("em", null, h("b", null, "P" + note.slide), " ", note.label),
												h("span", { className: "pa-del", onClick: () => removeNote(note.key) }, "删除"),
											),
											h("textarea", {
												value: note.note,
												placeholder: "这处要怎么改？例如：标题太大，缩到 28pt",
												onChange: (event) => setNote(note.key, event.target.value),
												onFocus: () => setActiveKey(note.key),
											}),
										),
									),
							),
				),
			);
		}

		/**
		 * Register the pptx preview implementation and its body.
		 * @param ctx - client root context carrying slots, previews and sessions.
		 */
		function apply(ctx) {
			ctx.effect(() => ctx.documentPreviews.register({
				id: ID,
				extensions: ["pptx"],
				binaryExtensions: ["pptx"],
				priority: "extension",
				title: () => "PPT 批注",
				loading: "renderer",
				wrap: false,
			}), "pptx-annotate: preview metadata");

			try {
				ctx.effect(() => ctx.slots.inject("conversation.composer.dock", () => ctx.slots.register({
					name: "conversation.composer.dock",
					id: "pptx-annotate-activity",
					order: 40,
					inject: (sessionId) => ({ sessionId }),
				}, AnnotateActivity)), "pptx-annotate: composer activity");
			} catch (error) {
				console.warn("pptx-annotate: composer activity seat unavailable", error);
			}

			ctx.effect(() => ctx.slots.inject("sidebar.right.tab.document", () => ctx.slots.register({
				name: "sidebar.right.tab.document",
				key: ID,
				inject: (sessionId) => ({
					sessionId,
					onSubmit: (text) => {
						const scoped = ctx.sessions.scope(sessionId);
						if (scoped === undefined) throw new Error("会话 " + String(sessionId) + " 尚未绑定，无法发送");
						const conversation = typeof scoped.get === "function" ? scoped.get("conversation") : scoped.conversation;
						if (conversation === undefined || typeof conversation.send !== "function") {
							throw new Error("conversation 服务不可用，无法发送");
						}
						return conversation.send(text);
					},
				}),
			}, AnnotateBody)), "pptx-annotate: preview body");
		}

		exports.apply = apply;
		exports.inject = inject;
		exports.AnnotateBody = AnnotateBody;
		exports.AnnotateActivity = AnnotateActivity;
		exports.formatPrompt = formatPrompt;
		exports.pickAt = pickAt;
		exports.labelOf = labelOf;
		return module.exports;
	}
});
