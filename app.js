(() => {
  function resolveUrl(rel) {
    try {
      return new URL(rel, document.baseURI).href;
    } catch (e) {
      return rel;
    }
  }
  if (window.pdfjsLib) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = resolveUrl("vendor/pdf.worker.min.js");
  }
  const $ = (id) => document.getElementById(id);

  const fileBase = $("fileBase");
  const fileComp = $("fileComp");
  const nameBase = $("nameBase");
  const nameComp = $("nameComp");
  const btnCompare = $("btnCompare");
  const btnDemo = $("btnDemo");
  const btnSwapFiles = $("btnSwapFiles");
  const btnReselect = $("btnReselect");
  const uploadPanel = $("uploadPanel");
  const resultPanel = $("resultPanel");
  const leftPane = $("leftPane");
  const rightPane = $("rightPane");
  const docsWrap = $("docsWrap");
  const svg = $("connectors");
  const changeList = $("changeList");
  const loading = $("loading");
  const toast = $("toast");
  const syncScroll = $("syncScroll");
  const swapDocs = $("swapDocs");
  const showLinks = $("showLinks");

  let token = "local";
  let data = null;
  let session = {
    demo: false,
    left: { paras: [], ext: ".docx", name: "", blob: null },
    right: { paras: [], ext: ".docx", name: "", blob: null },
  };
  let filterKind = "all";
  let selectedId = null;
  let syncing = false;
  let files = { base: null, comp: null };
  let anchors = [{ l: 0, r: 0 }];

  const KIND_LABEL = {
    delete: "删除",
    insert: "新增",
    modify: "修改",
    move: "移动",
  };

  function showToast(msg) {
    toast.textContent = msg;
    toast.hidden = false;
    clearTimeout(showToast._t);
    showToast._t = setTimeout(() => {
      toast.hidden = true;
    }, 3200);
  }

  function setLoading(on, msg) {
    loading.hidden = !on;
    if (msg) $("loadingText").textContent = msg;
  }

  function updateCompareEnabled() {
    btnCompare.disabled = !(files.base && files.comp);
  }

  function bindDrop(el, which) {
    function setOver(e) {
      e.preventDefault();
      e.stopPropagation();
      if (e.dataTransfer) e.dataTransfer.dropEffect = "copy";
      el.classList.add("over");
    }
    el.addEventListener("dragenter", setOver);
    el.addEventListener("dragover", setOver);
    el.addEventListener("dragleave", (e) => {
      if (!el.contains(e.relatedTarget)) el.classList.remove("over");
    });
    el.addEventListener("drop", (e) => {
      e.preventDefault();
      e.stopPropagation();
      el.classList.remove("over");
      const list = (e.dataTransfer && e.dataTransfer.files) || [];
      const f = list[0];
      if (f) setFile(which, f);
    });
  }

  function markDrops() {
    $("dropBase").classList.toggle("ready", !!files.base);
    $("dropComp").classList.toggle("ready", !!files.comp);
  }

  function setFile(which, f) {
    files[which] = f;
    (which === "base" ? nameBase : nameComp).textContent = f.name;
    markDrops();
    updateCompareEnabled();
  }

  bindDrop($("dropBase"), "base");
  bindDrop($("dropComp"), "comp");
  fileBase.addEventListener("change", () => {
    if (fileBase.files[0]) setFile("base", fileBase.files[0]);
  });
  fileComp.addEventListener("change", () => {
    if (fileComp.files[0]) setFile("comp", fileComp.files[0]);
  });

  btnSwapFiles.addEventListener("click", () => {
    const a = files.base;
    files.base = files.comp;
    files.comp = a;
    nameBase.textContent = files.base ? files.base.name : "点击选择，或把文件拖到这里";
    nameComp.textContent = files.comp ? files.comp.name : "点击选择，或把文件拖到这里";
    markDrops();
  });

  btnReselect.addEventListener("click", () => {
    resultPanel.hidden = true;
    uploadPanel.hidden = false;
    btnReselect.hidden = true;
    $("btnExportBoth").hidden = true;
    swapDocs.checked = false;
    svg.innerHTML = "";
    leftPane.closest(".pane").classList.remove("is-dirty");
    rightPane.closest(".pane").classList.remove("is-dirty");
  });

  function paraHtml(p, side) {
    const box = p.box ? ` box-${p.box}` : "";
    const cids = (p.change_ids || []).join(" ");
    const pair = p.pair == null ? "" : ` data-pair="${p.pair}"`;
    const kind = p.kind ? ` data-kind="${p.kind}"` : "";
    return `<p class="para${box}" id="para-${side}-${p.index}" data-index="${p.index}"${pair}${kind} data-cids="${cids}" contenteditable="true" spellcheck="false">${p.html || ""}</p>`;
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, (c) => ({
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#39;",
    }[c]));
  }

  function preview(text, n) {
    const t = (text || "").replace(/\s+/g, " ").trim();
    if (!t) return "（无）";
    const cut = t.length > n ? t.slice(0, n) + "…" : t;
    return escapeHtml(cut);
  }

  function visibleChanges() {
    if (!data) return [];
    if (filterKind === "all") return data.changes;
    return data.changes.filter((c) => c.kind === filterKind);
  }

  function renderSidebar() {
    const list = visibleChanges();
    changeList.innerHTML = list
      .map((c, i) => {
        return `<article class="card${c.id === selectedId ? " on" : ""}" data-id="${c.id}">
          <div><span class="idx">${i + 1}.</span><span class="kind ${c.kind}">${KIND_LABEL[c.kind]}</span></div>
          <div class="row"><span class="lab">原文</span>${preview(c.old_text, 80)}</div>
          <div class="row"><span class="lab">差异</span>${preview(c.new_text, 80)}</div>
        </article>`;
      })
      .join("");
  }

  function renderResult(payload) {
    data = payload;
    token = payload.token;
    $("paneLeftName").textContent = payload.left_name;
    $("paneRightName").textContent = payload.right_name;
    $("similarity").textContent = payload.similarity.toFixed(2) + "%";
    $("cntAll").textContent = payload.counts.all;
    $("cntDelete").textContent = payload.counts.delete;
    $("cntInsert").textContent = payload.counts.insert;
    $("cntModify").textContent = payload.counts.modify;
    $("cntMove").textContent = payload.counts.move;
    leftPane.innerHTML = payload.left.map((p) => paraHtml(p, "L")).join("");
    rightPane.innerHTML = payload.right.map((p) => paraHtml(p, "R")).join("");
    selectedId = null;
    renderSidebar();
    uploadPanel.hidden = true;
    resultPanel.hidden = false;
    btnReselect.hidden = false;
    $("btnExportBoth").hidden = false;
    requestAnimationFrame(() => {
      rebuildAnchors();
      drawConnectors();
      setTimeout(() => {
        rebuildAnchors();
        drawConnectors();
      }, 80);
    });
  }

  function contentTop(el, pane) {
    return el.getBoundingClientRect().top - pane.getBoundingClientRect().top + pane.scrollTop;
  }

  function rebuildAnchors() {
    const pts = [{ l: 0, r: 0 }];
    let lastL = -1;
    let lastR = -1;
    leftPane.querySelectorAll(".para").forEach((el) => {
      const kind = el.dataset.kind || "";
      if (kind === "move" || kind === "delete") return;
      if (el.dataset.pair == null || el.dataset.pair === "") return;
      const other = document.getElementById("para-R-" + el.dataset.pair);
      if (!other || other.dataset.kind === "move") return;
      const l = contentTop(el, leftPane);
      const r = contentTop(other, rightPane);
      if (l + 2 < lastL || r + 2 < lastR) return;
      lastL = l;
      lastR = r;
      pts.push({ l, r });
    });
    pts.push({ l: leftPane.scrollHeight, r: rightPane.scrollHeight });
    anchors = pts;
  }

  function mapAlongAnchors(y, fromKey, toKey) {
    const pts = anchors;
    if (!pts || pts.length < 2) return y;
    let i = 0;
    while (i < pts.length - 2 && pts[i + 1][fromKey] <= y) i += 1;
    const a = pts[i];
    const b = pts[i + 1];
    const span = b[fromKey] - a[fromKey];
    const t = span <= 1 ? 0 : (y - a[fromKey]) / span;
    return a[toKey] + t * (b[toKey] - a[toKey]);
  }

  function colorFor(kind) {
    return {
      delete: "#e53935",
      insert: "#2e7d32",
      modify: "#ef6c00",
      move: "#6a1b9a",
    }[kind] || "#1565c0";
  }

  function findPair(cid) {
    const a = document.getElementById(cid + "-L") || document.querySelector(`#leftPane [data-cid="${cid}"]`);
    const b = document.getElementById(cid + "-R") || document.querySelector(`#rightPane [data-cid="${cid}"]`);
    const ch = data && data.changes.find((x) => x.id === cid);
    const pa =
      a ||
      (ch && ch.left_para != null
        ? document.getElementById("para-L-" + ch.left_para)
        : null);
    const pb =
      b ||
      (ch && ch.right_para != null
        ? document.getElementById("para-R-" + ch.right_para)
        : null);
    return [pa, pb, ch];
  }

  function drawConnectors() {
    svg.innerHTML = "";
    if (!data || !showLinks.checked) return;
    const wrap = docsWrap.getBoundingClientRect();
    svg.setAttribute("viewBox", `0 0 ${wrap.width} ${wrap.height}`);
    svg.setAttribute("width", wrap.width);
    svg.setAttribute("height", wrap.height);

    const paneHeadEl = docsWrap.querySelector(".pane-head");
    const paneHeadH = paneHeadEl ? paneHeadEl.getBoundingClientRect().height : 0;

    const candidates = [];
    const kindsWanted =
      filterKind === "all"
        ? new Set(["modify", "move"])
        : new Set([filterKind]);

    for (const ch of data.changes) {
      if (!kindsWanted.has(ch.kind) && ch.id !== selectedId) continue;
      if (ch.kind === "delete" || ch.kind === "insert") {
        if (ch.id !== selectedId) continue;
      }
      const [elL, elR] = findPair(ch.id);
      if (!elL || !elR) continue;
      const rL = elL.getBoundingClientRect();
      const rR = elR.getBoundingClientRect();
      const visible =
        rL.bottom > wrap.top + paneHeadH &&
        rL.top < wrap.bottom &&
        rR.bottom > wrap.top + paneHeadH &&
        rR.top < wrap.bottom;
      if (!visible && ch.id !== selectedId) continue;
      const y1 = rL.top + rL.height / 2 - wrap.top;
      const y2 = rR.top + rR.height / 2 - wrap.top;
      const x1 = rL.right - wrap.left;
      const x2 = rR.left - wrap.left;
      if (x2 <= x1 + 8) continue;
      candidates.push({ ch, x1, y1, x2, y2, mid: Math.abs((y1 + y2) / 2 - wrap.height / 2) });
    }
    candidates.sort((a, b) => a.mid - b.mid);
    const keep = candidates.filter((c) => c.ch.id === selectedId);
    for (const c of candidates) {
      if (keep.length >= 36) break;
      if (!keep.includes(c)) keep.push(c);
    }
    const ns = "http://www.w3.org/2000/svg";
    for (const c of keep) {
      const path = document.createElementNS(ns, "path");
      const cx1 = c.x1 + Math.min(80, (c.x2 - c.x1) * 0.35);
      const cx2 = c.x2 - Math.min(80, (c.x2 - c.x1) * 0.35);
      path.setAttribute(
        "d",
        `M ${c.x1} ${c.y1} C ${cx1} ${c.y1}, ${cx2} ${c.y2}, ${c.x2} ${c.y2}`
      );
      path.setAttribute("stroke", colorFor(c.ch.kind));
      if (c.ch.id === selectedId) path.classList.add("sel");
      svg.appendChild(path);
    }
  }

  function clearActive() {
    document.querySelectorAll(".hl.active, .para.flash").forEach((el) => {
      el.classList.remove("active", "flash");
    });
  }

  function jumpTo(cid) {
    selectedId = cid;
    clearActive();
    const [elL, elR] = findPair(cid);
    syncing = true;
    scrollLeader = null;
    const focus = 0.28;
    if (elL) {
      elL.classList.add(elL.classList.contains("para") ? "flash" : "active");
      const y = contentTop(elL, leftPane);
      leftPane.scrollTop = Math.max(0, y - leftPane.clientHeight * focus);
    }
    if (elR) {
      elR.classList.add(elR.classList.contains("para") ? "flash" : "active");
      const y = contentTop(elR, rightPane);
      rightPane.scrollTop = Math.max(0, y - rightPane.clientHeight * focus);
    }
    renderSidebar();
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        syncing = false;
        drawConnectors();
      });
    });
  }

  changeList.addEventListener("click", (e) => {
    const card = e.target.closest(".card");
    if (card) jumpTo(card.dataset.id);
  });

  docsWrap.addEventListener("click", (e) => {
    if (e.target.closest(".para")) return;
    const hl = e.target.closest("[data-cid]");
    if (hl) jumpTo(hl.getAttribute("data-cid"));
  });

  $("filters").addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-kind]");
    if (!btn) return;
    filterKind = btn.dataset.kind;
    $("filters").querySelectorAll("button").forEach((b) => b.classList.toggle("on", b === btn));
    renderSidebar();
    drawConnectors();
  });

  function onScroll(from, to, fromSide) {
    if (!syncScroll.checked || syncing) {
      drawConnectors();
      return;
    }
    if (scrollLeader && from !== scrollLeader) {
      drawConnectors();
      return;
    }
    rebuildAnchors();
    syncing = true;
    const fromKey = fromSide === "L" ? "l" : "r";
    const toKey = fromSide === "L" ? "r" : "l";
    const focus = from.clientHeight * 0.28;
    const srcY = from.scrollTop + focus;
    const dstY = mapAlongAnchors(srcY, fromKey, toKey);
    const nextTop = Math.max(0, dstY - focus);
    to.scrollTop = nextTop;
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        syncing = false;
        drawConnectors();
      });
    });
  }
  let scrollLeader = null;
  function markLeader(el) {
    scrollLeader = el;
  }
  leftPane.addEventListener("wheel", () => markLeader(leftPane), { passive: true });
  rightPane.addEventListener("wheel", () => markLeader(rightPane), { passive: true });
  leftPane.addEventListener("pointerdown", () => markLeader(leftPane));
  rightPane.addEventListener("pointerdown", () => markLeader(rightPane));
  leftPane.addEventListener("scroll", () => onScroll(leftPane, rightPane, "L"), { passive: true });
  rightPane.addEventListener("scroll", () => onScroll(rightPane, leftPane, "R"), { passive: true });
  window.addEventListener("resize", () => {
    rebuildAnchors();
    drawConnectors();
  });
  function collectParas(pane) {
    return [...pane.querySelectorAll(".para")]
      .map((p) => (p.innerText || "").replace(/\u00a0/g, " ").replace(/\n+$/, ""))
      .filter((t) => t.trim());
  }

  function unwrapHighlights(para) {
    para.querySelectorAll("span.hl").forEach((sp) => {
      sp.replaceWith(document.createTextNode(sp.textContent || ""));
    });
    para.classList.remove("box-delete", "box-insert", "box-modify", "box-move");
    para.removeAttribute("data-cids");
    para.normalize();
  }

  function normalizeEditable(pane) {
    pane.querySelectorAll(":scope > p, :scope > div").forEach((el) => {
      el.classList.add("para");
      el.setAttribute("contenteditable", "true");
      el.setAttribute("spellcheck", "false");
    });
    pane.closest(".pane").classList.add("is-dirty");
    rebuildAnchors();
    drawConnectors();
  }

  function bindEditor(pane) {
    pane.addEventListener("input", (e) => {
      const para = e.target.closest && e.target.closest(".para");
      if (para) unwrapHighlights(para);
      normalizeEditable(pane);
    });
    pane.addEventListener("paste", (e) => {
      e.preventDefault();
      var clip = e.clipboardData || window.clipboardData;
      var text = "";
      if (clip) {
        text = clip.getData ? clip.getData("text/plain") || clip.getData("Text") || "" : "";
      }
      var ok = false;
      try {
        ok = document.execCommand("insertText", false, text);
      } catch (err) {
        ok = false;
      }
      if (ok) return;
      var sel = window.getSelection();
      if (!sel || !sel.rangeCount) return;
      var range = sel.getRangeAt(0);
      range.deleteContents();
      range.insertNode(document.createTextNode(text));
      range.collapse(false);
      sel.removeAllRanges();
      sel.addRange(range);
    });
  }
  bindEditor(leftPane);
  bindEditor(rightPane);

  function downloadBlob(blob, filename) {
    if (window.navigator && window.navigator.msSaveOrOpenBlob) {
      window.navigator.msSaveOrOpenBlob(blob, filename);
      return;
    }
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.rel = "noopener";
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    setTimeout(function () {
      if (a.parentNode) a.parentNode.removeChild(a);
      URL.revokeObjectURL(url);
    }, 40000);
  }

  async function exportOne(side, paragraphs, mode) {
    const rec = side === "right" ? session.right : session.left;
    const stem = (rec.name || "文档").replace(/\.[^.]+$/, "") || "文档";
    const tag = mode === "save" ? "已修改" : "导出";
    const orig = rec.paras || [];
    if (rec.blob && DocDiff.parasUnchanged(orig, paragraphs)) {
      const ext = rec.ext || ".docx";
      downloadBlob(rec.blob, stem + "-" + tag + ext);
      return ext;
    }
    const blob = await DocDiff.buildDocx(paragraphs);
    downloadBlob(blob, stem + "-" + tag + ".docx");
    if ((rec.ext || "") === ".pdf") {
      showToast("浏览器里改过的 PDF 已另存为 Word");
    }
    return ".docx";
  }

  async function postExport(opts) {
    if (!data) {
      showToast("请先完成对比");
      return;
    }
    setLoading(true, opts.mode === "save" ? "正在保存…" : "正在导出…");
    try {
      if (opts.side === "both") {
        const zip = new JSZip();
        const leftParas = opts.left_paragraphs || opts.paragraphs;
        const rightParas = opts.right_paragraphs;
        const lName = (session.left.name || "基准").replace(/\.[^.]+$/, "") + "-导出.docx";
        const rName = (session.right.name || "对照").replace(/\.[^.]+$/, "") + "-导出.docx";
        if (session.left.blob && DocDiff.parasUnchanged(session.left.paras, leftParas)) {
          zip.file(
            (session.left.name || "基准").replace(/\.[^.]+$/, "") + "-导出" + (session.left.ext || ".docx"),
            session.left.blob
          );
        } else {
          zip.file(lName, await DocDiff.buildDocx(leftParas));
        }
        if (session.right.blob && DocDiff.parasUnchanged(session.right.paras, rightParas)) {
          zip.file(
            (session.right.name || "对照").replace(/\.[^.]+$/, "") + "-导出" + (session.right.ext || ".docx"),
            session.right.blob
          );
        } else {
          zip.file(rName, await DocDiff.buildDocx(rightParas));
        }
        downloadBlob(
          await zip.generateAsync({ type: "blob", mimeType: "application/zip" }),
          "文档对比-导出.zip"
        );
      } else {
        await exportOne(opts.side, opts.paragraphs, opts.mode);
      }
      showToast(opts.mode === "save" ? "已保存" : "已导出");
      if (opts.side === "left" || opts.side === "both") {
        leftPane.closest(".pane").classList.remove("is-dirty");
      }
      if (opts.side === "right" || opts.side === "both") {
        rightPane.closest(".pane").classList.remove("is-dirty");
      }
    } catch (err) {
      showToast(err.message || String(err));
    } finally {
      setLoading(false);
    }
  }

  $("btnSaveLeft").addEventListener("click", () => {
    postExport({ token, side: "left", mode: "save", paragraphs: collectParas(leftPane) });
  });
  $("btnSaveRight").addEventListener("click", () => {
    postExport({ token, side: "right", mode: "save", paragraphs: collectParas(rightPane) });
  });
  $("btnExportLeft").addEventListener("click", () => {
    postExport({ token, side: "left", mode: "export", paragraphs: collectParas(leftPane) });
  });
  $("btnExportRight").addEventListener("click", () => {
    postExport({ token, side: "right", mode: "export", paragraphs: collectParas(rightPane) });
  });
  $("btnExportBoth").addEventListener("click", () => {
    postExport({
      token,
      side: "both",
      mode: "export",
      left_paragraphs: collectParas(leftPane),
      right_paragraphs: collectParas(rightPane),
    });
  });

  async function runCompare(left, right, demo) {
    setLoading(true, "正在比对，请稍候…");
    try {
      const payload = DocDiff.compareTexts(left.paras, right.paras, left.name, right.name);
      payload.token = "local";
      session = { demo: !!demo, left, right };
      renderResult(payload);
    } catch (err) {
      showToast(err.message || String(err));
    } finally {
      setLoading(false);
    }
  }

  btnCompare.addEventListener("click", async () => {
    if (!files.base || !files.comp) return;
    setLoading(true, "正在读取文件…");
    try {
      const left = await DocDiff.extractFile(files.base);
      const right = await DocDiff.extractFile(files.comp);
      if (!left.paras.length && !right.paras.length) {
        throw new Error("没有读到文字。扫描版 PDF 无法识别，请另存为可复制文字的 Word 再比");
      }
      await runCompare(left, right, false);
    } catch (err) {
      setLoading(false);
      showToast(err.message || String(err));
    }
  });

  btnDemo.addEventListener("click", () => {
    runCompare(
      { paras: DocDiff.DEMO_LEFT, ext: ".docx", name: "示例-基准.docx", blob: null },
      { paras: DocDiff.DEMO_RIGHT, ext: ".docx", name: "示例-对照.docx", blob: null },
      true
    );
  });

  if (!window.DocDiff || !window.JSZip) {
    showToast("脚本未加载完整，请用 Microsoft Edge 或 Chrome 打开，并允许运行 JavaScript");
  }

  swapDocs.addEventListener("change", async () => {
    if (!data) {
      swapDocs.checked = false;
      return;
    }
    const tmp = session.left;
    session.left = session.right;
    session.right = tmp;
    await runCompare(session.left, session.right, session.demo);
  });
})();
