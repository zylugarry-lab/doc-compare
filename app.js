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
  const btnOpenEdit = $("btnOpenEdit");
  const btnNewBlank = $("btnNewBlank");
  const btnSwapFiles = $("btnSwapFiles");
  const btnReselect = $("btnReselect");
  const uploadPanel = $("uploadPanel");
  const resultPanel = $("resultPanel");
  const editPanel = $("editPanel");
  const editPane = $("editPane");
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
  let uiMode = "compare";
  let editSession = { paras: [], ext: ".docx", name: "", blob: null };
  let activeEditPane = null;
  let activePara = null;

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
    btnOpenEdit.disabled = !(files.base || files.comp);
  }

  function setMode(mode) {
    uiMode = mode === "edit" ? "edit" : "compare";
    document.body.classList.toggle("app-mode-edit", uiMode === "edit");
    document.body.classList.toggle("app-mode-compare", uiMode !== "edit");
    $("tabCompare").classList.toggle("on", uiMode === "compare");
    $("tabEdit").classList.toggle("on", uiMode === "edit");
    $("tabCompare").setAttribute("aria-selected", uiMode === "compare" ? "true" : "false");
    $("tabEdit").setAttribute("aria-selected", uiMode === "edit" ? "true" : "false");
    btnCompare.hidden = uiMode !== "compare";
    btnDemo.hidden = uiMode !== "compare";
    btnOpenEdit.hidden = uiMode !== "edit";
    btnNewBlank.hidden = uiMode !== "edit";
    $("leadText").textContent =
      uiMode === "edit"
        ? "打开一份 Word / PDF，在网页里直接改字，保存为 Word。回车分段，段首退格合并。文件只在这台电脑处理，不会上传。"
        : "左侧放基准，右侧放对照。支持 .docx、.doc、RTF 和 PDF。对比后可直接改字，也可切到「在线修改」打开一份文件编辑。若选文件窗口里看不到 Word，把右下角改成「所有文件」。";
    $("labelBase").textContent = uiMode === "edit" ? "要修改的文档" : "基准文档";
    $("hintBase").textContent = uiMode === "edit" ? "打开后可在页面里改字并保存" : "差异将相对此文档计算";
    updateCompareEnabled();
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

  function goHome() {
    resultPanel.hidden = true;
    editPanel.hidden = true;
    uploadPanel.hidden = false;
    btnReselect.hidden = true;
    $("btnExportBoth").hidden = true;
    $("btnRecompare").hidden = true;
    $("btnEditSaveTop").hidden = true;
    swapDocs.checked = false;
    svg.innerHTML = "";
    leftPane.closest(".pane").classList.remove("is-dirty");
    rightPane.closest(".pane").classList.remove("is-dirty");
  }
  btnReselect.addEventListener("click", goHome);

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
        const applyLabel =
          c.kind === "insert" ? "写入基准" : c.kind === "delete" ? "从基准删去" : "采用到基准";
        return `<article class="card${c.id === selectedId ? " on" : ""}" data-id="${c.id}">
          <div><span class="idx">${i + 1}.</span><span class="kind ${c.kind}">${KIND_LABEL[c.kind]}</span></div>
          <div class="row"><span class="lab">原文</span>${preview(c.old_text, 80)}</div>
          <div class="row"><span class="lab">差异</span>${preview(c.new_text, 80)}</div>
          <div class="card-actions"><button type="button" class="btn-apply" data-apply="${c.id}">${applyLabel}</button></div>
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
    editPanel.hidden = true;
    resultPanel.hidden = false;
    btnReselect.hidden = false;
    $("btnExportBoth").hidden = false;
    $("btnRecompare").hidden = false;
    $("btnEditSaveTop").hidden = true;
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
    const applyBtn = e.target.closest("[data-apply]");
    if (applyBtn) {
      e.preventDefault();
      e.stopPropagation();
      applyChangeToLeft(applyBtn.getAttribute("data-apply"));
      return;
    }
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

  function paraPlain(el) {
    return (el.innerText || "").replace(/\u00a0/g, " ").replace(/\n+$/, "");
  }

  function unwrapHighlights(para) {
    para.querySelectorAll("span.hl").forEach((sp) => {
      sp.replaceWith(document.createTextNode(sp.textContent || ""));
    });
    para.classList.remove("box-delete", "box-insert", "box-modify", "box-move");
    para.removeAttribute("data-cids");
    para.normalize();
  }

  function markDirty(pane) {
    const wrap = pane.closest ? pane.closest(".pane") : null;
    if (wrap) wrap.classList.add("is-dirty");
    if (pane === editPane) $("editMeta").textContent = collectParas(editPane).length + " 段";
    if (pane === leftPane || pane === rightPane) {
      rebuildAnchors();
      drawConnectors();
    }
  }

  function normalizeEditable(pane) {
    pane.querySelectorAll(":scope > p, :scope > div").forEach((el) => {
      el.classList.add("para");
      el.setAttribute("contenteditable", "true");
      el.setAttribute("spellcheck", "false");
    });
    markDirty(pane);
  }

  function makeParaEl(text) {
    const p = document.createElement("p");
    p.className = "para";
    p.setAttribute("contenteditable", "true");
    p.setAttribute("spellcheck", "false");
    p.textContent = text || "";
    return p;
  }

  function caretOffset(para) {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return 0;
    const range = sel.getRangeAt(0);
    const pre = document.createRange();
    pre.selectNodeContents(para);
    try {
      pre.setEnd(range.startContainer, range.startOffset);
    } catch (err) {
      return 0;
    }
    return pre.toString().length;
  }

  function placeCaret(el, offset) {
    const sel = window.getSelection();
    if (!sel) return;
    const range = document.createRange();
    let remaining = Math.max(0, offset || 0);
    const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
    let node = walker.nextNode();
    if (!node) {
      el.appendChild(document.createTextNode(""));
      node = el.firstChild;
      range.setStart(node, 0);
      range.collapse(true);
      sel.removeAllRanges();
      sel.addRange(range);
      return;
    }
    while (node) {
      if (remaining <= node.length) {
        range.setStart(node, remaining);
        range.collapse(true);
        sel.removeAllRanges();
        sel.addRange(range);
        return;
      }
      remaining -= node.length;
      const next = walker.nextNode();
      if (!next) {
        range.setStart(node, node.length);
        range.collapse(true);
        sel.removeAllRanges();
        sel.addRange(range);
        return;
      }
      node = next;
    }
  }

  function splitPara(para) {
    unwrapHighlights(para);
    const offset = caretOffset(para);
    const text = paraPlain(para);
    const left = text.slice(0, offset);
    const right = text.slice(offset);
    para.textContent = left;
    const next = makeParaEl(right);
    para.after(next);
    next.focus();
    placeCaret(next, 0);
    markDirty(para.parentNode);
  }

  function mergePara(para) {
    const prev = para.previousElementSibling;
    if (!prev || (prev.className || "").indexOf("para") === -1) return false;
    unwrapHighlights(prev);
    unwrapHighlights(para);
    const keep = paraPlain(prev);
    prev.textContent = keep + paraPlain(para);
    para.parentNode.removeChild(para);
    prev.focus();
    placeCaret(prev, keep.length);
    markDirty(prev.parentNode);
    activePara = prev;
    return true;
  }

  function insertParaIn(pane) {
    const cur =
      activePara && pane.contains(activePara)
        ? activePara
        : pane.querySelector(".para:last-child");
    const p = makeParaEl("");
    if (cur) cur.after(p);
    else pane.appendChild(p);
    p.focus();
    placeCaret(p, 0);
    activePara = p;
    activeEditPane = pane;
    markDirty(pane);
  }

  function deleteParaIn(pane) {
    const cur = activePara && pane.contains(activePara) ? activePara : null;
    if (!cur) {
      showToast("请先点选要删除的段落");
      return;
    }
    const paras = pane.querySelectorAll(".para");
    if (paras.length <= 1) {
      cur.textContent = "";
      markDirty(pane);
      return;
    }
    const next = cur.nextElementSibling || cur.previousElementSibling;
    cur.parentNode.removeChild(cur);
    if (next && (next.className || "").indexOf("para") !== -1) {
      next.focus();
      placeCaret(next, paraPlain(next).length);
      activePara = next;
    }
    markDirty(pane);
  }

  function applyChangeToLeft(cid) {
    if (!data) return;
    const ch = data.changes.filter(function (x) {
      return x.id === cid;
    })[0];
    if (!ch) return;
    jumpTo(cid);
    if (ch.kind === "modify" || ch.kind === "move") {
      const elL = document.getElementById("para-L-" + ch.left_para);
      const elR =
        ch.right_para != null ? document.getElementById("para-R-" + ch.right_para) : null;
      if (!elL) return;
      unwrapHighlights(elL);
      elL.textContent = elR ? paraPlain(elR) : ch.new_text || "";
    } else if (ch.kind === "insert") {
      const elR =
        ch.right_para != null ? document.getElementById("para-R-" + ch.right_para) : null;
      const text = elR ? paraPlain(elR) : ch.new_text || "";
      const p = makeParaEl(text);
      let anchor = null;
      if (ch.right_para != null && ch.right_para > 0) {
        const prevR = document.getElementById("para-R-" + (ch.right_para - 1));
        if (prevR && prevR.getAttribute("data-pair")) {
          anchor = document.getElementById("para-L-" + prevR.getAttribute("data-pair"));
        }
      }
      if (anchor) anchor.after(p);
      else leftPane.appendChild(p);
      p.focus();
      activePara = p;
      activeEditPane = leftPane;
    } else if (ch.kind === "delete") {
      const elL = document.getElementById("para-L-" + ch.left_para);
      if (elL) {
        if (leftPane.querySelectorAll(".para").length <= 1) elL.textContent = "";
        else elL.parentNode.removeChild(elL);
      }
    }
    markDirty(leftPane);
    showToast("已写入基准，可再点「按正文重比」");
  }

  function bindEditor(pane) {
    pane.addEventListener("focusin", (e) => {
      const para = e.target.closest && e.target.closest(".para");
      if (para) {
        activeEditPane = pane;
        activePara = para;
      }
    });
    pane.addEventListener("input", (e) => {
      const para = e.target.closest && e.target.closest(".para");
      if (para) unwrapHighlights(para);
      normalizeEditable(pane);
    });
    pane.addEventListener("keydown", (e) => {
      const para = e.target.closest && e.target.closest(".para");
      if (!para) return;
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        splitPara(para);
      } else if (e.key === "Backspace") {
        const sel = window.getSelection();
        if (!sel || !sel.rangeCount || !sel.isCollapsed) return;
        if (caretOffset(para) !== 0) return;
        e.preventDefault();
        mergePara(para);
      }
    });
    pane.addEventListener("paste", (e) => {
      e.preventDefault();
      var clip = e.clipboardData || window.clipboardData;
      var text = "";
      if (clip) {
        text = clip.getData ? clip.getData("text/plain") || clip.getData("Text") || "" : "";
      }
      var blocks = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
      if (blocks.length > 1) {
        const para = e.target.closest && e.target.closest(".para");
        if (para) {
          unwrapHighlights(para);
          const offset = caretOffset(para);
          const cur = paraPlain(para);
          para.textContent = cur.slice(0, offset) + blocks[0];
          let last = para;
          for (let i = 1; i < blocks.length; i++) {
            const np = makeParaEl(blocks[i]);
            last.after(np);
            last = np;
          }
          last.focus();
          placeCaret(last, paraPlain(last).length);
          markDirty(pane);
          return;
        }
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
  bindEditor(editPane);

  function updateEditMeta() {
    $("editFileName").textContent = editSession.name || "未命名.docx";
    $("editMeta").textContent = (editPane.querySelectorAll(".para").length || 0) + " 段";
  }

  function openEditor(rec) {
    editSession = rec;
    const paras = rec.paras && rec.paras.length ? rec.paras : [""];
    editPane.innerHTML = "";
    paras.forEach((t) => {
      editPane.appendChild(makeParaEl(t));
    });
    uploadPanel.hidden = true;
    resultPanel.hidden = true;
    editPanel.hidden = false;
    btnReselect.hidden = false;
    $("btnEditSaveTop").hidden = false;
    $("btnExportBoth").hidden = true;
    $("btnRecompare").hidden = true;
    updateEditMeta();
    const first = editPane.querySelector(".para");
    if (first) {
      first.focus();
      placeCaret(first, paraPlain(first).length);
      activePara = first;
      activeEditPane = editPane;
    }
  }

  async function saveEditor() {
    const paragraphs = collectParas(editPane);
    if (!paragraphs.length) {
      showToast("没有可保存的文字");
      return;
    }
    setLoading(true, "正在保存…");
    try {
      const blob = await DocDiff.buildDocx(paragraphs);
      const stem = (editSession.name || "未命名").replace(/\.[^.]+$/, "") || "未命名";
      downloadBlob(blob, stem + "-已修改.docx");
      showToast("已保存为 Word");
    } catch (err) {
      showToast(err.message || String(err));
    } finally {
      setLoading(false);
    }
  }

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

  $("tabCompare").addEventListener("click", () => setMode("compare"));
  $("tabEdit").addEventListener("click", () => setMode("edit"));
  btnOpenEdit.addEventListener("click", async () => {
    const f = files.base || files.comp;
    if (!f) return;
    setLoading(true, "正在读取文件…");
    try {
      const rec = await DocDiff.extractFile(f);
      if (!rec.paras.length) {
        throw new Error("没有读到文字。扫描版 PDF 无法识别，请另存为可复制文字的 Word");
      }
      openEditor(rec);
    } catch (err) {
      showToast(err.message || String(err));
    } finally {
      setLoading(false);
    }
  });
  btnNewBlank.addEventListener("click", () => {
    openEditor({ paras: [""], ext: ".docx", name: "未命名.docx", blob: null });
  });
  $("btnEditInsert").addEventListener("click", () => insertParaIn(editPane));
  $("btnEditDelete").addEventListener("click", () => deleteParaIn(editPane));
  $("btnEditSave").addEventListener("click", () => saveEditor());
  $("btnEditSaveTop").addEventListener("click", () => saveEditor());
  $("btnInsertLeft").addEventListener("click", () => insertParaIn(leftPane));
  $("btnDeleteLeft").addEventListener("click", () => deleteParaIn(leftPane));
  $("btnInsertRight").addEventListener("click", () => insertParaIn(rightPane));
  $("btnDeleteRight").addEventListener("click", () => deleteParaIn(rightPane));
  $("btnRecompare").addEventListener("click", async () => {
    if (!session.left || !session.right) return;
    const left = {
      paras: collectParas(leftPane),
      ext: session.left.ext,
      name: session.left.name,
      blob: session.left.blob,
    };
    const right = {
      paras: collectParas(rightPane),
      ext: session.right.ext,
      name: session.right.name,
      blob: session.right.blob,
    };
    await runCompare(left, right, session.demo);
  });
  document.addEventListener("keydown", (e) => {
    if (!(e.ctrlKey || e.metaKey) || String(e.key).toLowerCase() !== "s") return;
    if (editPanel.hidden && resultPanel.hidden) return;
    e.preventDefault();
    if (!editPanel.hidden) saveEditor();
    else if (activeEditPane === rightPane) {
      postExport({ token, side: "right", mode: "save", paragraphs: collectParas(rightPane) });
    } else {
      postExport({ token, side: "left", mode: "save", paragraphs: collectParas(leftPane) });
    }
  });

  setMode("compare");

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
