/* Plot to Docx: Local Web App Frontend
 * Key design invariant:
 *   Step 2 (Organize & Group) mutates DOM elements IN PLACE for reordering,
 *   drag-and-drop, checkboxes, renaming, and page breaks so scrolling and
 *   eyes never suffer a full-page reload flash.
 */

(() => {
  "use strict";

  const PAGE_CHUNK_SIZE = 200;
  const AUTO_SELECT_THRESHOLD = 100;

  // ----------------------------------------------------------------------- //
  // Application State                                                       //
  // ----------------------------------------------------------------------- //
  const state = {
    currentStep: "select",
    hasWin32Com: true,
    hasDialogs: true,

    // Step 1
    sourceFolder: "",
    lastFailedFolder: "",
    treeMode: "lazy", // "lazy" | "scan"
    dirCache: new Map(), // folderPath -> { path, name, parent, folders, files, loaded }
    expandedFolders: new Set(),
    collapsedTreeFolders: new Set(),
    visibleLimits: new Map(), // folderPath -> number of files shown
    scanId: null,
    scanNextIndex: 0,
    scanPollTimer: null,
    scanFiles: [],
    selectedSet: new Set(),
    fileInfoMap: new Map(), // path -> { name, path, folder, ext, size_str, mtime }
    plotRelPaths: new Map(), // path -> rel_path
    selectFilter: "",

    // Step 2
    groupingMode: "manual", // "manual" | "auto" | "flat"
    sections: [], // [{ id, name, page_break, collapsed, plots: [path, ...], subsections: [...] }]
    secCounter: 1,
    checkedOrgSet: new Set(),
    showThumbs: false,
    orgSearch: "",

    // Step 3 & Config
    docOptions: {
      orientation: "landscape",
      num_cols: 2,
      max_width_cm: null,
      max_height_cm: null,
      export_docx: true,
      export_html: false,
      html_embed_images: false,
      export_pdf: false,
      last_output_path: null,
    },
    configPath: null,
    configName: "",
    fileStatusMap: {}, // path -> { exists, size_str, mtime_str }
    genJobId: null,
    genPollTimer: null,
  };

  // Drag-and-drop and UI transient state
  let dragContext = null; // { type: "plot"|"section", paths?: string[], sourceSecId?: string, secId?: string }
  let thumbObserver = null;
  let lastFocusedBeforeModal = null;

  const SVG_ICONS = {
    sun: '<svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="square" aria-hidden="true"><circle cx="8" cy="8" r="3"/><path d="M8 1.5v1.5M8 13v1.5M1.5 8H3M13 8h1.5M3.4 3.4l1.1 1.1M11.5 11.5l1.1 1.1M12.6 3.4l-1.1 1.1M4.5 11.5l-1.1 1.1"/></svg>',
    moon: '<svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="square" aria-hidden="true"><path d="M13 9.8A5.3 5.3 0 0 1 6.2 3a5.5 5.5 0 1 0 6.8 6.8Z"/></svg>',
    folder: '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="square" aria-hidden="true"><path d="M1.5 3.5h4.5l1.5 1.5h7v7.5h-13z"/></svg>',
    image: '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="square" aria-hidden="true"><rect x="2" y="2.5" width="12" height="11"/><path d="M2.5 11l3.5-3.5 2.5 2.5 2-2 3 3"/><circle cx="10.8" cy="5.8" r="0.9" fill="currentColor" stroke="none"/></svg>',
    eye: '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="square" aria-hidden="true"><path d="M1.5 8s2.5-4.2 6.5-4.2S14.5 8 14.5 8s-2.5 4.2-6.5 4.2S1.5 8 1.5 8Z"/><circle cx="8" cy="8" r="1.8"/></svg>',
    chevronRight: '<svg width="12" height="12" viewBox="0 0 12 12" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="square" aria-hidden="true"><path d="M4.5 2.5L8 6l-3.5 3.5"/></svg>',
    chevronDown: '<svg width="12" height="12" viewBox="0 0 12 12" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="square" aria-hidden="true"><path d="M2.5 4.5L6 8l3.5-3.5"/></svg>',
    up: '<svg width="12" height="12" viewBox="0 0 12 12" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="square" aria-hidden="true"><path d="M2.5 7.5L6 4l3.5 3.5"/></svg>',
    down: '<svg width="12" height="12" viewBox="0 0 12 12" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="square" aria-hidden="true"><path d="M2.5 4.5L6 8l3.5-3.5"/></svg>',
    close: '<svg width="12" height="12" viewBox="0 0 12 12" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="square" aria-hidden="true"><path d="M2.5 2.5l7 7M9.5 2.5l-7 7"/></svg>',
    grip: '<svg width="12" height="14" viewBox="0 0 12 14" fill="currentColor" aria-hidden="true"><rect x="2" y="2" width="2" height="2"/><rect x="8" y="2" width="2" height="2"/><rect x="2" y="6" width="2" height="2"/><rect x="8" y="6" width="2" height="2"/><rect x="2" y="10" width="2" height="2"/><rect x="8" y="10" width="2" height="2"/></svg>',
  };

  function setIconHtml(el, name) {
    el.innerHTML = SVG_ICONS[name] || "";
  }

  function createIconSpan(name, className = "tree-glyph") {
    const span = document.createElement("span");
    span.className = className;
    span.setAttribute("aria-hidden", "true");
    setIconHtml(span, name);
    return span;
  }

  // ----------------------------------------------------------------------- //
  // API Helpers                                                             //
  // ----------------------------------------------------------------------- //
  async function apiGet(url) {
    const res = await fetch(url, { cache: "no-store" });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || `Request failed (${res.status})`);
    return data;
  }

  async function apiPost(url, body = {}) {
    const res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
      cache: "no-store",
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || `Request failed (${res.status})`);
    return data;
  }

  function showToast(message, kind = "info", timeoutMs = 3800) {
    const container = document.getElementById("toast-container");
    const el = document.createElement("div");
    el.className = `toast toast-${kind}`;
    el.textContent = message;
    container.appendChild(el);
    setTimeout(() => {
      el.remove();
    }, timeoutMs);
  }

  function openModal({ title, bodyEl, buttons = [] }) {
    lastFocusedBeforeModal = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    const backdrop = document.getElementById("modal-backdrop");
    document.getElementById("modal-title").textContent = title;
    const body = document.getElementById("modal-body");
    const footer = document.getElementById("modal-footer");
    body.innerHTML = "";
    footer.innerHTML = "";
    if (typeof bodyEl === "string") {
      body.textContent = bodyEl;
    } else {
      body.appendChild(bodyEl);
    }
    for (const btnCfg of buttons) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = `btn ${btnCfg.className || "btn-secondary"}`;
      btn.textContent = btnCfg.label;
      btn.addEventListener("click", () => btnCfg.onClick?.(closeModal));
      footer.appendChild(btn);
    }
    backdrop.classList.remove("hidden");
    const firstInput = body.querySelector("input, select, textarea");
    const firstFooterBtn = footer.querySelector("button");
    const focusTarget = firstInput || firstFooterBtn || document.getElementById("modal-close");
    if (focusTarget) setTimeout(() => focusTarget.focus(), 20);
  }

  function closeModal() {
    const backdrop = document.getElementById("modal-backdrop");
    if (backdrop.classList.contains("hidden")) return;
    backdrop.classList.add("hidden");
    if (lastFocusedBeforeModal && document.contains(lastFocusedBeforeModal)) {
      lastFocusedBeforeModal.focus();
    }
    lastFocusedBeforeModal = null;
  }

  function promptPathModal({ title, label, defaultValue = "", placeholder = "", submitLabel = "Confirm" }) {
    return new Promise((resolve) => {
      let settled = false;
      const finish = (val) => {
        if (settled) return;
        settled = true;
        closeModal();
        resolve(val);
      };

      const wrap = document.createElement("div");
      const lbl = document.createElement("label");
      lbl.className = "field-label";
      lbl.htmlFor = "modal-path-input";
      lbl.textContent = label;
      const input = document.createElement("input");
      input.id = "modal-path-input";
      input.type = "text";
      input.className = "input mono-text";
      input.style.width = "100%";
      input.value = defaultValue;
      input.placeholder = placeholder;
      input.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
          e.preventDefault();
          const v = input.value.trim();
          if (v) finish(v);
        }
      });
      wrap.append(lbl, input);

      openModal({
        title,
        bodyEl: wrap,
        buttons: [
          { label: "Cancel", className: "btn-secondary", onClick: () => finish(null) },
          {
            label: submitLabel,
            className: "btn-primary",
            onClick: () => {
              const v = input.value.trim();
              if (!v) return;
              finish(v);
            },
          },
        ],
      });

      const backdrop = document.getElementById("modal-backdrop");
      const closeBtn = document.getElementById("modal-close");
      const onCancelWatch = () => {
        if (backdrop.classList.contains("hidden") && !settled) {
          settled = true;
          resolve(null);
        }
      };
      closeBtn.addEventListener("click", onCancelWatch, { once: true });
      setTimeout(() => { input.focus(); input.select(); }, 30);
    });
  }

  function basename(p) {
    if (!p) return "";
    const idx = Math.max(p.lastIndexOf("\\"), p.lastIndexOf("/"));
    return idx >= 0 ? p.slice(idx + 1) : p;
  }

  function dirname(p) {
    if (!p) return "";
    const norm = p.replace(/[\\/]+$/, "");
    const idx = Math.max(norm.lastIndexOf("\\"), norm.lastIndexOf("/"));
    return idx > 0 ? norm.slice(0, idx) : norm;
  }

  function thumbUrl(filePath, size = 120) {
    return `/api/thumb?path=${encodeURIComponent(filePath)}&size=${size}`;
  }

  function showImagePreview(filePath) {
    const wrap = document.createElement("div");
    const statusNote = document.createElement("div");
    statusNote.className = "muted";
    statusNote.style.marginBottom = "8px";
    statusNote.textContent = "Loading preview...";

    const img = document.createElement("img");
    img.className = "preview-img";
    img.width = 480;
    img.height = 320;
    img.alt = basename(filePath);
    img.addEventListener("load", () => {
      statusNote.remove();
    });
    img.addEventListener("error", () => {
      img.remove();
      statusNote.className = "preview-error";
      statusNote.textContent = "Could not load image preview. The file may have been moved or deleted.";
    });
    img.src = thumbUrl(filePath, 512);

    const caption = document.createElement("div");
    caption.className = "muted mono-text";
    caption.style.wordBreak = "break-all";
    caption.textContent = filePath;
    wrap.append(statusNote, img, caption);

    openModal({
      title: basename(filePath),
      bodyEl: wrap,
      buttons: [{ label: "Close", className: "btn-primary", onClick: closeModal }],
    });
  }

  // ----------------------------------------------------------------------- //
  // Theme & Lifecycle Heartbeat                                             //
  // ----------------------------------------------------------------------- //
  function applyThemeUi(theme) {
    document.documentElement.setAttribute("data-theme", theme);
    const btn = document.getElementById("btn-theme-toggle");
    setIconHtml(btn, theme === "dark" ? "sun" : "moon");
    btn.setAttribute("aria-label", theme === "dark" ? "Switch to light theme" : "Switch to dark theme");
  }

  function initTheme() {
    const prefLight = window.matchMedia && window.matchMedia("(prefers-color-scheme: light)").matches;
    const saved = localStorage.getItem("plot2docx_theme") || (prefLight ? "light" : "dark");
    applyThemeUi(saved);
  }

  function toggleTheme() {
    const curr = document.documentElement.getAttribute("data-theme") || "dark";
    const next = curr === "dark" ? "light" : "dark";
    localStorage.setItem("plot2docx_theme", next);
    applyThemeUi(next);
  }

  function initHeartbeat() {
    apiPost("/api/ping", {}).catch(() => {});
    setInterval(() => {
      apiPost("/api/ping", {}).catch(() => {});
    }, 5000);

    const sendBye = () => {
      try {
        if (navigator.sendBeacon) {
          navigator.sendBeacon("/api/bye", "{}");
        }
      } catch (_) {}
    };
    window.addEventListener("pagehide", sendBye);
  }

  // ----------------------------------------------------------------------- //
  // Navigation & Badges                                                     //
  // ----------------------------------------------------------------------- //
  function updateHeaderBadges() {
    document.getElementById("badge-selected-count").textContent = String(state.selectedSet.size);
    const activeSecCount = state.sections.filter((s) => collectSectionPlots(s, 0).length > 0).length || state.sections.length;
    document.getElementById("badge-section-count").textContent = String(activeSecCount);
  }

  function switchStep(step) {
    state.currentStep = step;
    document.querySelectorAll(".step-pill").forEach((btn) => {
      const isActive = btn.dataset.step === step;
      btn.classList.toggle("active", isActive);
      if (isActive) btn.setAttribute("aria-current", "step");
      else btn.removeAttribute("aria-current");
    });
    document.querySelectorAll(".step-view").forEach((view) => {
      view.classList.toggle("active", view.id === `view-${step}`);
    });

    const subtitles = {
      select: "Browse a local or UNC folder, select plots, then organize and export.",
      organize: "Drag handles or use arrow buttons to reorder plots and sections without flicker.",
      generate: "Customize page layout, verify files, and generate Word (.docx) & lossless high-res PDF.",
    };
    document.getElementById("step-subtitle").textContent = subtitles[step] || "";

    if (step === "organize") {
      reconcileSectionsWithSelection();
      renderSectionsFull();
    } else if (step === "generate") {
      reconcileSectionsWithSelection();
      renderGenerateSummary();
    }
    updateHeaderBadges();
  }

  function cancelActiveScanIfAny() {
    stopScanPolling();
    if (state.scanId) {
      const prevId = state.scanId;
      state.scanId = null;
      apiPost("/api/scan/cancel", { scan_id: prevId }).catch(() => {});
    }
  }

  // ----------------------------------------------------------------------- //
  // Step 1: Lazy Folder Browser & Guarded Deep Scan                         //
  // ----------------------------------------------------------------------- //
  async function openFolderLazy(folderPath) {
    const target = (folderPath || "").trim();
    if (!target) {
      showToast("Enter or browse to a folder path first.", "warn");
      return;
    }

    cancelActiveScanIfAny();
    document.getElementById("scan-banner").classList.add("hidden");

    const emptyEl = document.getElementById("tree-empty-state");
    const loadingEl = document.getElementById("tree-loading-state");
    const errorEl = document.getElementById("tree-error-state");
    const container = document.getElementById("folder-tree-container");
    emptyEl.classList.add("hidden");
    errorEl.classList.add("hidden");
    container.classList.add("hidden");
    document.getElementById("tree-loading-detail").textContent = target;
    loadingEl.classList.remove("hidden");

    try {
      const listing = await apiGet(`/api/fs/list?path=${encodeURIComponent(target)}`);
      loadingEl.classList.add("hidden");
      state.lastFailedFolder = "";
      state.treeMode = "lazy";
      state.sourceFolder = listing.path;
      document.getElementById("folder-path-input").value = listing.path;

      state.dirCache.clear();
      state.expandedFolders.clear();
      state.collapsedTreeFolders.clear();
      state.visibleLimits.clear();
      state.scanFiles = [];

      cacheDirectoryListing(listing);
      state.expandedFolders.add(listing.path);

      if (!state.configPath && state.selectedSet.size === 0) {
        if (listing.files.length > 0 && listing.files.length <= AUTO_SELECT_THRESHOLD) {
          for (const f of listing.files) state.selectedSet.add(f.path);
        } else if (listing.files.length > AUTO_SELECT_THRESHOLD) {
          showToast(
            `Loaded ${listing.files.length} plots in this folder (> ${AUTO_SELECT_THRESHOLD}). Not auto-selected to keep things fast.`,
            "info",
            5000
          );
        }
      }

      renderStep1Tree();
      updateHeaderBadges();
    } catch (err) {
      loadingEl.classList.add("hidden");
      state.lastFailedFolder = target;
      document.getElementById("tree-error-message").textContent = err.message;
      errorEl.classList.remove("hidden");
      showToast(err.message, "error");
    }
  }

  function cacheDirectoryListing(listing) {
    state.dirCache.set(listing.path, {
      ...listing,
      loaded: true,
    });
    for (const f of listing.files) {
      state.fileInfoMap.set(f.path, f);
    }
  }

  async function toggleLazyFolder(folderPath, triggerBtn = null) {
    if (state.expandedFolders.has(folderPath)) {
      state.expandedFolders.delete(folderPath);
      renderStep1Tree();
      return;
    }

    const cached = state.dirCache.get(folderPath);
    if (!cached || !cached.loaded) {
      if (triggerBtn) triggerBtn.disabled = true;
      try {
        const listing = await apiGet(`/api/fs/list?path=${encodeURIComponent(folderPath)}`);
        cacheDirectoryListing(listing);
      } catch (err) {
        if (triggerBtn) triggerBtn.disabled = false;
        showToast(err.message, "error");
        return;
      }
    }
    state.expandedFolders.add(folderPath);
    renderStep1Tree();
  }

  function stopScanPolling() {
    if (state.scanPollTimer) {
      clearTimeout(state.scanPollTimer);
      state.scanPollTimer = null;
    }
  }

  async function startDeepScan(folderPath) {
    const target = (folderPath || "").trim();
    if (!target) {
      showToast("Enter or browse to a folder path first.", "warn");
      return;
    }

    cancelActiveScanIfAny();
    document.getElementById("tree-loading-state").classList.add("hidden");
    document.getElementById("tree-error-state").classList.add("hidden");
    try {
      const res = await apiPost("/api/scan/start", { path: target, soft_limit: 500 });
      state.treeMode = "scan";
      state.sourceFolder = target;
      state.scanId = res.scan_id;
      state.scanNextIndex = 0;
      state.scanFiles = [];
      state.dirCache.clear();
      state.expandedFolders.clear();
      state.collapsedTreeFolders.clear();
      state.visibleLimits.clear();

      const banner = document.getElementById("scan-banner");
      banner.classList.remove("hidden");
      document.getElementById("scan-spinner").classList.remove("paused");
      document.getElementById("btn-scan-continue").classList.add("hidden");
      document.getElementById("btn-scan-keep").classList.add("hidden");
      document.getElementById("btn-scan-cancel").classList.remove("hidden");
      document.getElementById("scan-status-title").textContent = "Deep scanning folder...";
      document.getElementById("scan-status-detail").textContent = target;

      pollScanStatus();
    } catch (err) {
      showToast(err.message, "error");
    }
  }

  async function pollScanStatus() {
    if (!state.scanId) return;
    try {
      const st = await apiGet(
        `/api/scan/status?id=${encodeURIComponent(state.scanId)}&since=${state.scanNextIndex}`
      );
      state.sourceFolder = st.root || state.sourceFolder;
      document.getElementById("folder-path-input").value = state.sourceFolder;

      if (st.new_files && st.new_files.length > 0) {
        for (const f of st.new_files) {
          state.scanFiles.push(f);
          state.fileInfoMap.set(f.path, f);
        }
        state.scanNextIndex = st.next_index;
        renderStep1Tree();
      }

      document.getElementById("scan-status-detail").textContent =
        `Found ${st.total_files} plot(s) across ${st.scanned_dirs} folder(s)...`;

      if (st.status === "running") {
        state.scanPollTimer = setTimeout(pollScanStatus, 250);
      } else if (st.status === "paused_limit") {
        document.getElementById("scan-spinner").classList.add("paused");
        document.getElementById("scan-status-title").textContent =
          `Paused at ${st.total_files} images (large folder guardrail)`;
        document.getElementById("scan-status-detail").textContent =
          `Scanned ${st.scanned_dirs} of ${st.total_dirs} discovered folders. Choose how to proceed:`;
        document.getElementById("btn-scan-continue").classList.remove("hidden");
        document.getElementById("btn-scan-keep").classList.remove("hidden");
      } else if (st.status === "done") {
        document.getElementById("scan-banner").classList.add("hidden");
        if (!state.configPath && state.selectedSet.size === 0) {
          if (state.scanFiles.length > 0 && state.scanFiles.length <= AUTO_SELECT_THRESHOLD) {
            for (const f of state.scanFiles) state.selectedSet.add(f.path);
            showToast(`Scan complete: ${state.scanFiles.length} plots auto-selected.`, "success");
          } else if (state.scanFiles.length > AUTO_SELECT_THRESHOLD) {
            showToast(
              `Scan complete: ${state.scanFiles.length} plots found (> ${AUTO_SELECT_THRESHOLD}). Not auto-selected; pick what you need.`,
              "info",
              5500
            );
          } else {
            showToast("Scan complete: no image files found.", "warn");
          }
        } else {
          showToast(`Scan complete: ${state.scanFiles.length} plots indexed.`, "success");
        }
        renderStep1Tree();
        updateHeaderBadges();
      } else if (st.status === "cancelled") {
        document.getElementById("scan-banner").classList.add("hidden");
        showToast(`Scan cancelled (${state.scanFiles.length} plots kept loaded).`, "warn");
        renderStep1Tree();
      } else if (st.status === "error") {
        document.getElementById("scan-banner").classList.add("hidden");
        showToast(`Scan error: ${st.error}`, "error");
      }
    } catch (err) {
      document.getElementById("scan-banner").classList.add("hidden");
      showToast(err.message, "error");
    }
  }

  function getLoadedFilesList() {
    if (state.treeMode === "scan") {
      return state.scanFiles;
    }
    const list = [];
    for (const dir of state.dirCache.values()) {
      for (const f of dir.files || []) list.push(f);
    }
    return list;
  }

  function getFilteredLoadedFiles() {
    const q = state.selectFilter.trim().toLowerCase();
    const all = getLoadedFilesList();
    if (!q) return all;
    return all.filter((f) => f.name.toLowerCase().includes(q) || f.path.toLowerCase().includes(q));
  }

  function getSubtreeLoadedFilePaths(folderPath) {
    const results = [];
    const stack = [folderPath];
    const visited = new Set();
    while (stack.length > 0) {
      const cur = stack.pop();
      if (visited.has(cur)) continue;
      visited.add(cur);
      const entry = state.dirCache.get(cur);
      if (!entry) continue;
      for (const f of entry.files || []) results.push(f.path);
      for (const sub of entry.folders || []) stack.push(sub.path);
    }
    return results;
  }

  function buildNestedTreeFromFiles(rootPath, files) {
    const cleanRoot = (rootPath || "").replace(/[\\/]+$/, "");
    const sep = cleanRoot.includes("\\") ? "\\" : "/";
    const rootNode = {
      path: cleanRoot || ".",
      name: cleanRoot || "Plots",
      children: new Map(),
      files: [],
    };
    const rootLower = cleanRoot.toLowerCase();

    function getFolderNode(folderPath) {
      const cleanFolder = (folderPath || "").replace(/[\\/]+$/, "");
      if (!cleanFolder || cleanFolder.toLowerCase() === rootLower) {
        return rootNode;
      }
      const lower = cleanFolder.toLowerCase();
      let parts = [];
      if (rootLower && (lower.startsWith(rootLower + "\\") || lower.startsWith(rootLower + "/"))) {
        parts = cleanFolder.slice(cleanRoot.length).replace(/^[\\/]+/, "").split(/[\\/]+/).filter(Boolean);
      } else {
        parts = [cleanFolder];
      }
      let curr = rootNode;
      for (const part of parts) {
        const key = part.toLowerCase();
        if (!curr.children.has(key)) {
          const nextPath = curr === rootNode && parts.length === 1 && cleanFolder === part
            ? cleanFolder
            : `${curr.path}${sep}${part}`;
          curr.children.set(key, {
            path: nextPath,
            name: part,
            children: new Map(),
            files: [],
          });
        }
        curr = curr.children.get(key);
      }
      return curr;
    }

    for (const f of files) {
      getFolderNode(f.folder).files.push(f);
    }
    return rootNode;
  }

  function collectNodeFilePaths(node) {
    const out = node.files.map((f) => f.path);
    for (const child of node.children.values()) {
      out.push(...collectNodeFilePaths(child));
    }
    return out;
  }

  function renderExplorerNode(node, isRoot, isFilterActive) {
    const wrap = document.createElement("div");
    wrap.className = isRoot ? "tree-node tree-root" : "tree-node";

    const isCollapsed = !isFilterActive && state.collapsedTreeFolders.has(node.path);
    const subtreePaths = collectNodeFilePaths(node);
    const labelText = isRoot ? node.path : node.name;

    const row = document.createElement("div");
    row.className = "tree-folder-row";

    const expander = document.createElement("button");
    expander.type = "button";
    expander.className = "tree-expander";
    expander.setAttribute("aria-expanded", String(!isCollapsed));
    expander.setAttribute("aria-label", `${isCollapsed ? "Expand" : "Collapse"} ${labelText}`);
    setIconHtml(expander, isCollapsed ? "chevronRight" : "chevronDown");

    const toggleExpand = () => {
      if (isFilterActive) return;
      if (state.collapsedTreeFolders.has(node.path)) state.collapsedTreeFolders.delete(node.path);
      else state.collapsedTreeFolders.add(node.path);
      renderStep1Tree();
    };
    expander.addEventListener("click", toggleExpand);

    const chk = buildFolderCheckbox(node.path, subtreePaths, labelText);
    const icon = createIconSpan("folder");

    const nameBtn = document.createElement("button");
    nameBtn.type = "button";
    nameBtn.className = "tree-name-btn";
    nameBtn.textContent = labelText;
    nameBtn.addEventListener("click", toggleExpand);

    const meta = document.createElement("span");
    meta.className = "tree-meta mono-text";
    meta.textContent = `${node.children.size} subfolder(s), ${subtreePaths.length} plot(s)`;

    row.append(expander, chk, icon, nameBtn, meta);
    wrap.appendChild(row);

    if (!isCollapsed) {
      const childrenWrap = document.createElement("div");
      childrenWrap.className = "tree-children";

      const sortedChildren = Array.from(node.children.values()).sort((a, b) =>
        a.name.localeCompare(b.name, undefined, { sensitivity: "base" })
      );
      for (const child of sortedChildren) {
        childrenWrap.appendChild(renderExplorerNode(child, false, isFilterActive));
      }

      const limit = state.visibleLimits.get(node.path) || PAGE_CHUNK_SIZE;
      for (const f of node.files.slice(0, limit)) {
        childrenWrap.appendChild(createFileRowEl(f));
      }
      if (node.files.length > limit) {
        childrenWrap.appendChild(buildChunkControls(node.path, node.files.length, limit));
      }
      wrap.appendChild(childrenWrap);
    }
    return wrap;
  }

  function renderStep1Tree() {
    const emptyEl = document.getElementById("tree-empty-state");
    const loadingEl = document.getElementById("tree-loading-state");
    const errorEl = document.getElementById("tree-error-state");
    const container = document.getElementById("folder-tree-container");
    const statsEl = document.getElementById("loaded-stats-lbl");

    loadingEl.classList.add("hidden");
    errorEl.classList.add("hidden");

    const filterQ = state.selectFilter.trim().toLowerCase();
    document.getElementById("btn-select-filtered").disabled = !filterQ;

    const allLoaded = getLoadedFilesList();
    const hasRoot = Boolean(state.sourceFolder) && (state.dirCache.size > 0 || state.scanFiles.length > 0);

    if (!hasRoot) {
      emptyEl.classList.remove("hidden");
      container.classList.add("hidden");
      statsEl.textContent = "No folder opened yet.";
      return;
    }

    emptyEl.classList.add("hidden");
    container.classList.remove("hidden");
    statsEl.textContent = `${allLoaded.length} loaded | ${state.selectedSet.size} selected`;

    const prevScroll = container.scrollTop;
    const frag = document.createDocumentFragment();

    if (filterQ || state.treeMode === "scan") {
      const files = getFilteredLoadedFiles();
      if (files.length === 0) {
        const msg = document.createElement("div");
        msg.className = "section-empty-note";
        msg.textContent = filterQ ? "No loaded plots match your filter." : "No image files found.";
        frag.appendChild(msg);
      } else {
        const treeRoot = buildNestedTreeFromFiles(state.sourceFolder, files);
        frag.appendChild(renderExplorerNode(treeRoot, true, Boolean(filterQ)));
      }
    } else {
      const rootEntry = state.dirCache.get(state.sourceFolder);
      if (rootEntry) {
        frag.appendChild(buildLazyFolderNode(rootEntry, true));
      }
    }

    container.innerHTML = "";
    container.appendChild(frag);
    container.scrollTop = prevScroll;
  }

  function createFileRowEl(fileItem) {
    const row = document.createElement("div");
    row.className = "tree-file-row";

    const chk = document.createElement("input");
    chk.type = "checkbox";
    chk.checked = state.selectedSet.has(fileItem.path);
    chk.setAttribute("aria-label", `Select ${fileItem.name}`);
    chk.addEventListener("change", () => {
      if (chk.checked) {
        state.selectedSet.add(fileItem.path);
      } else {
        state.selectedSet.delete(fileItem.path);
      }
      updateHeaderBadges();
      renderStep1Tree();
    });

    const icon = createIconSpan("image");

    const name = document.createElement("span");
    name.className = "tree-name";
    name.textContent = fileItem.name;
    name.title = fileItem.path;

    const meta = document.createElement("span");
    meta.className = "tree-meta mono-text";
    meta.textContent = `${fileItem.ext} · ${fileItem.size_str}`;

    const previewBtn = document.createElement("button");
    previewBtn.type = "button";
    previewBtn.className = "btn btn-sm btn-ghost";
    setIconHtml(previewBtn, "eye");
    previewBtn.title = "Preview plot";
    previewBtn.setAttribute("aria-label", `Preview ${fileItem.name}`);
    previewBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      showImagePreview(fileItem.path);
    });

    row.append(chk, icon, name, meta, previewBtn);
    return row;
  }

  function buildFolderCheckbox(folderPath, filePaths, labelName) {
    const chk = document.createElement("input");
    chk.type = "checkbox";
    chk.setAttribute("aria-label", `Select all in ${labelName}`);
    const total = filePaths.length;
    const checkedCount = filePaths.filter((p) => state.selectedSet.has(p)).length;
    chk.checked = total > 0 && checkedCount === total;
    chk.indeterminate = checkedCount > 0 && checkedCount < total;
    chk.disabled = total === 0 && state.treeMode === "lazy" && !state.expandedFolders.has(folderPath);

    chk.addEventListener("change", () => {
      for (const p of filePaths) {
        if (chk.checked) state.selectedSet.add(p);
        else state.selectedSet.delete(p);
      }
      updateHeaderBadges();
      renderStep1Tree();
    });
    return chk;
  }

  function buildLazyFolderNode(dirEntry, isRoot = false) {
    const wrap = document.createElement("div");
    wrap.className = isRoot ? "tree-node tree-root" : "tree-node";

    const isExpanded = state.expandedFolders.has(dirEntry.path);
    const subtreeFiles = getSubtreeLoadedFilePaths(dirEntry.path);
    const labelText = isRoot ? dirEntry.path : dirEntry.name;

    const row = document.createElement("div");
    row.className = "tree-folder-row";

    const expander = document.createElement("button");
    expander.type = "button";
    expander.className = "tree-expander";
    expander.setAttribute("aria-expanded", String(isExpanded));
    expander.setAttribute("aria-label", `${isExpanded ? "Collapse" : "Expand"} ${labelText}`);
    setIconHtml(expander, isExpanded ? "chevronDown" : "chevronRight");

    const chk = buildFolderCheckbox(dirEntry.path, subtreeFiles, labelText);

    const icon = createIconSpan("folder");

    const nameBtn = document.createElement("button");
    nameBtn.type = "button";
    nameBtn.className = "tree-name-btn";
    nameBtn.textContent = labelText;

    expander.addEventListener("click", () => toggleLazyFolder(dirEntry.path, expander));
    nameBtn.addEventListener("click", () => toggleLazyFolder(dirEntry.path, nameBtn));

    const meta = document.createElement("span");
    meta.className = "tree-meta mono-text";
    meta.textContent = isExpanded
      ? `${dirEntry.folders.length} subfolder(s), ${dirEntry.files.length} plot(s)`
      : "Click to expand";

    row.append(expander, chk, icon, nameBtn, meta);
    wrap.appendChild(row);

    if (isExpanded) {
      const childrenWrap = document.createElement("div");
      childrenWrap.className = "tree-children";

      for (const sub of dirEntry.folders) {
        const cachedSub = state.dirCache.get(sub.path) || {
          path: sub.path,
          name: sub.name,
          folders: [],
          files: [],
          loaded: false,
        };
        childrenWrap.appendChild(buildLazyFolderNode(cachedSub, false));
      }

      const limit = state.visibleLimits.get(dirEntry.path) || PAGE_CHUNK_SIZE;
      const shownFiles = dirEntry.files.slice(0, limit);
      for (const f of shownFiles) {
        childrenWrap.appendChild(createFileRowEl(f));
      }
      if (dirEntry.files.length > limit) {
        childrenWrap.appendChild(buildChunkControls(dirEntry.path, dirEntry.files.length, limit));
      }
      if (dirEntry.folders.length === 0 && dirEntry.files.length === 0) {
        const emptyNote = document.createElement("div");
        emptyNote.className = "section-empty-note";
        emptyNote.textContent = "No images or subfolders in this folder.";
        childrenWrap.appendChild(emptyNote);
      }
      wrap.appendChild(childrenWrap);
    }
    return wrap;
  }

  function buildChunkControls(folderPath, totalCount, currentLimit) {
    const box = document.createElement("div");
    box.className = "tree-show-more";

    const moreBtn = document.createElement("button");
    moreBtn.type = "button";
    moreBtn.className = "btn btn-sm btn-secondary";
    const nextChunk = Math.min(totalCount - currentLimit, PAGE_CHUNK_SIZE);
    moreBtn.textContent = `Show ${nextChunk} more (${totalCount - currentLimit} remaining)`;
    moreBtn.addEventListener("click", () => {
      state.visibleLimits.set(folderPath, currentLimit + PAGE_CHUNK_SIZE);
      renderStep1Tree();
    });

    const allBtn = document.createElement("button");
    allBtn.type = "button";
    allBtn.className = "btn btn-sm btn-ghost";
    allBtn.textContent = `Show all (${totalCount})`;
    allBtn.addEventListener("click", () => {
      state.visibleLimits.set(folderPath, totalCount);
      renderStep1Tree();
    });

    box.append(moreBtn, allBtn);
    return box;
  }

  // ----------------------------------------------------------------------- //
  // Step 2: Section Reconciliation & Zero-Flicker In-Place Organizer        //
  // ----------------------------------------------------------------------- //
  const MAX_SUBSECTION_DEPTH = 2;

  function walkSections(list, fn, depth = 0) {
    for (const sec of list || []) {
      fn(sec, depth);
      if (depth < MAX_SUBSECTION_DEPTH && sec.subsections?.length) {
        walkSections(sec.subsections, fn, depth + 1);
      }
    }
  }

  function collectSectionPlots(sec, depth = 0) {
    const out = [...(sec.plots || [])];
    if (depth < MAX_SUBSECTION_DEPTH) {
      for (const sub of sec.subsections || []) {
        out.push(...collectSectionPlots(sub, depth + 1));
      }
    }
    return out;
  }

  function findSectionEntry(secId, list = state.sections, parent = null, depth = 0) {
    for (const sec of list || []) {
      if (sec.id === secId) return { sec, parent, list, depth };
      if (depth < MAX_SUBSECTION_DEPTH && sec.subsections?.length) {
        const hit = findSectionEntry(secId, sec.subsections, sec, depth + 1);
        if (hit) return hit;
      }
    }
    return null;
  }

  function reconcileSectionsWithSelection() {
    const selectedOrder = Array.from(state.selectedSet);
    const selectedLookup = new Set(selectedOrder);

    // Prune checked organizer set to active selection
    for (const p of Array.from(state.checkedOrgSet)) {
      if (!selectedLookup.has(p)) state.checkedOrgSet.delete(p);
    }

    if (selectedOrder.length === 0) {
      if (state.groupingMode !== "manual") {
        state.sections = [];
      } else {
        walkSections(state.sections, (s) => { s.plots = []; });
        state.sections = state.sections.filter((s) => collectSectionPlots(s, 0).length > 0);
      }
      return;
    }

    if (state.groupingMode === "auto") {
      const prevPb = new Map(state.sections.map((s) => [s.name, s.page_break]));
      const groups = new Map();
      for (const p of selectedOrder) {
        const folderName = basename(dirname(p)) || "General";
        if (!groups.has(folderName)) groups.set(folderName, []);
        groups.get(folderName).push(p);
      }
      state.sections = [];
      let idx = 1;
      for (const [folderName, plots] of groups.entries()) {
        state.sections.push({
          id: `sec_${idx++}`,
          name: folderName,
          page_break: prevPb.has(folderName) ? prevPb.get(folderName) : true,
          collapsed: false,
          plots,
          subsections: [],
        });
      }
      state.secCounter = idx;
      return;
    }

    if (state.groupingMode === "flat") {
      const existing = state.sections[0]?.plots || [];
      const kept = existing.filter((p) => selectedLookup.has(p));
      const keptSet = new Set(kept);
      const added = selectedOrder.filter((p) => !keptSet.has(p));
      state.sections = [{
        id: "sec_flat",
        name: "All Plots (No Sections)",
        page_break: false,
        collapsed: false,
        plots: [...kept, ...added],
        subsections: [],
      }];
      return;
    }

    // Manual mode: preserve existing sections, subsections, renames, page breaks & plot order
    const present = new Set();
    walkSections(state.sections, (sec) => {
      if (!sec.subsections) sec.subsections = [];
      sec.plots = (sec.plots || []).filter((p) => {
        if (selectedLookup.has(p) && !present.has(p)) {
          present.add(p);
          return true;
        }
        return false;
      });
    });
    const newlySelected = selectedOrder.filter((p) => !present.has(p));
    if (newlySelected.length > 0) {
      if (state.sections.length === 0) {
        state.sections.push({
          id: `sec_${state.secCounter++}`,
          name: "General Plots",
          page_break: true,
          collapsed: false,
          plots: newlySelected,
          subsections: [],
        });
      } else {
        state.sections[state.sections.length - 1].plots.push(...newlySelected);
      }
    }
  }

  function setupThumbObserver() {
    if (thumbObserver) thumbObserver.disconnect();
    if (!state.showThumbs) return;
    const container = document.getElementById("sections-container");
    if (typeof IntersectionObserver === "undefined") {
      container.querySelectorAll("img.plot-thumb").forEach((img) => {
        if (!img.getAttribute("src") && img.dataset.src) {
          img.src = img.dataset.src;
        }
      });
      return;
    }
    thumbObserver = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          if (!entry.isIntersecting) continue;
          const img = entry.target;
          if (img.dataset.src && !img.getAttribute("src")) {
            img.src = img.dataset.src;
          }
          thumbObserver.unobserve(img);
        }
      },
      { root: container, rootMargin: "200px" }
    );
    container.querySelectorAll("img.plot-thumb").forEach((img) => {
      if (!img.getAttribute("src") && img.dataset.src) {
        thumbObserver.observe(img);
      }
    });
  }

  function renderSectionsFull() {
    const container = document.getElementById("sections-container");
    const emptyEl = document.getElementById("organize-empty-state");
    const prevScroll = container.scrollTop;

    document.getElementById("btn-add-section").disabled = state.groupingMode !== "manual";
    document.querySelectorAll("#grouping-segmented .seg-btn").forEach((b) => {
      const isActive = b.dataset.mode === state.groupingMode;
      b.classList.toggle("active", isActive);
      b.setAttribute("aria-pressed", String(isActive));
    });

    const totalPlots = state.sections.reduce((acc, s) => acc + collectSectionPlots(s, 0).length, 0);
    if (totalPlots === 0 && state.sections.length === 0) {
      container.classList.add("hidden");
      emptyEl.classList.remove("hidden");
      updateBulkBarUI();
      return;
    }

    container.classList.remove("hidden");
    emptyEl.classList.add("hidden");
    container.classList.toggle("show-thumbs", state.showThumbs);

    const frag = document.createDocumentFragment();
    for (const sec of state.sections) {
      frag.appendChild(createSectionCardElement(sec, 0));
    }
    container.innerHTML = "";
    container.appendChild(frag);
    container.scrollTop = prevScroll;

    syncStateAndDomInPlace();
    setupThumbObserver();
  }

  function createSectionCardElement(sec, depth = 0) {
    const isFlat = state.groupingMode === "flat";
    const isManual = state.groupingMode === "manual";

    const card = document.createElement("div");
    card.className = `section-card section-depth-${depth}` + (sec.collapsed ? " collapsed" : "");
    card.dataset.secId = sec.id;
    card.dataset.depth = String(depth);

    // Header
    const header = document.createElement("div");
    header.className = "section-header";

    const left = document.createElement("div");
    left.className = "section-header-left";

    if (!isFlat) {
      const handle = document.createElement("span");
      handle.className = "drag-handle";
      setIconHtml(handle, "grip");
      handle.title = "Drag to reorder section";
      handle.setAttribute("aria-hidden", "true");
      handle.draggable = true;
      handle.addEventListener("dragstart", (e) => onSectionDragStart(e, sec.id, card));
      handle.addEventListener("dragend", () => onDragEndCleanup());
      left.appendChild(handle);
    }

    const collapseBtn = document.createElement("button");
    collapseBtn.type = "button";
    collapseBtn.className = "btn btn-sm btn-ghost sec-collapse-btn";
    setIconHtml(collapseBtn, sec.collapsed ? "chevronRight" : "chevronDown");
    collapseBtn.title = "Collapse / Expand section";
    collapseBtn.setAttribute("aria-expanded", String(!sec.collapsed));
    collapseBtn.setAttribute("aria-label", `${sec.collapsed ? "Expand" : "Collapse"} section ${sec.name}`);
    collapseBtn.addEventListener("click", () => {
      sec.collapsed = !sec.collapsed;
      card.classList.toggle("collapsed", sec.collapsed);
      setIconHtml(collapseBtn, sec.collapsed ? "chevronRight" : "chevronDown");
      collapseBtn.setAttribute("aria-expanded", String(!sec.collapsed));
      collapseBtn.setAttribute("aria-label", `${sec.collapsed ? "Expand" : "Collapse"} section ${sec.name}`);
    });
    left.appendChild(collapseBtn);

    const titleSpan = document.createElement("span");
    titleSpan.className = "sec-title-text";
    left.appendChild(titleSpan);

    const countSpan = document.createElement("span");
    countSpan.className = "sec-count-badge mono-text";
    left.appendChild(countSpan);

    const selSecBtn = document.createElement("button");
    selSecBtn.type = "button";
    selSecBtn.className = "btn btn-sm btn-secondary sec-select-btn";
    selSecBtn.textContent = "Select Section";
    selSecBtn.addEventListener("click", () => toggleSectionPlotsChecked(sec.id));
    left.appendChild(selSecBtn);

    const right = document.createElement("div");
    right.className = "section-header-right";

    if (!isFlat) {
      const pbLabel = document.createElement("label");
      pbLabel.className = "sec-pb-label";
      const pbChk = document.createElement("input");
      pbChk.type = "checkbox";
      pbChk.checked = Boolean(sec.page_break);
      pbChk.addEventListener("change", () => {
        sec.page_break = pbChk.checked;
      });
      const pbTxt = document.createElement("span");
      pbTxt.textContent = "Page Break Before";
      pbLabel.append(pbChk, pbTxt);
      right.appendChild(pbLabel);

      if (isManual && depth < MAX_SUBSECTION_DEPTH) {
        const subBtn = document.createElement("button");
        subBtn.type = "button";
        subBtn.className = "btn btn-sm btn-secondary sec-add-sub-btn";
        subBtn.textContent = "+ Subsection";
        subBtn.addEventListener("click", () => addSubsectionModal(sec, depth + 1));
        right.appendChild(subBtn);
      }

      if (isManual) {
        const renBtn = document.createElement("button");
        renBtn.type = "button";
        renBtn.className = "btn btn-sm btn-secondary";
        renBtn.textContent = "Rename";
        renBtn.addEventListener("click", () => renameSectionModal(sec));
        right.appendChild(renBtn);
      }

      const upBtn = document.createElement("button");
      upBtn.type = "button";
      upBtn.className = "btn btn-sm btn-secondary sec-up-btn";
      setIconHtml(upBtn, "up");
      upBtn.title = "Move section up";
      upBtn.setAttribute("aria-label", `Move section ${sec.name} up`);
      upBtn.addEventListener("click", () => moveSectionInPlace(sec.id, -1));

      const downBtn = document.createElement("button");
      downBtn.type = "button";
      downBtn.className = "btn btn-sm btn-secondary sec-down-btn";
      setIconHtml(downBtn, "down");
      downBtn.title = "Move section down";
      downBtn.setAttribute("aria-label", `Move section ${sec.name} down`);
      downBtn.addEventListener("click", () => moveSectionInPlace(sec.id, 1));

      right.append(upBtn, downBtn);

      if (isManual) {
        const delBtn = document.createElement("button");
        delBtn.type = "button";
        delBtn.className = "btn btn-sm btn-danger sec-del-btn";
        setIconHtml(delBtn, "close");
        delBtn.title = depth === 0 ? "Delete section" : "Delete subsection";
        delBtn.setAttribute("aria-label", `Delete ${depth === 0 ? "section" : "subsection"} ${sec.name}`);
        delBtn.addEventListener("click", () => deleteSectionInPlace(sec.id));
        right.appendChild(delBtn);
      }
    }

    header.append(left, right);
    card.appendChild(header);

    // Section-level drag target for reordering sibling sections
    card.addEventListener("dragover", (e) => onSectionCardDragOver(e, card));
    card.addEventListener("dragleave", () => {
      card.classList.remove("drag-over-before", "drag-over-after");
    });
    card.addEventListener("drop", (e) => onSectionCardDrop(e, card, sec.id));

    // Plots body
    const plotsBox = document.createElement("div");
    plotsBox.className = "section-plots";
    plotsBox.dataset.secId = sec.id;

    plotsBox.addEventListener("dragover", (e) => onSectionPlotsDragOver(e, plotsBox));
    plotsBox.addEventListener("dragleave", () => plotsBox.classList.remove("drag-over-empty"));
    plotsBox.addEventListener("drop", (e) => onSectionPlotsDrop(e, plotsBox, sec.id));

    for (const p of sec.plots || []) {
      plotsBox.appendChild(createPlotRowElement(p));
    }

    card.appendChild(plotsBox);

    if (!isFlat && depth < MAX_SUBSECTION_DEPTH) {
      const subsBox = document.createElement("div");
      subsBox.className = "section-subsections";
      subsBox.dataset.parentSecId = sec.id;
      for (const sub of sec.subsections || []) {
        subsBox.appendChild(createSectionCardElement(sub, depth + 1));
      }
      card.appendChild(subsBox);
    }
    return card;
  }

  function createPlotRowElement(plotPath) {
    const plotName = basename(plotPath);
    const row = document.createElement("div");
    row.className = "plot-row";
    row.dataset.path = plotPath;
    row.draggable = true;

    if (state.checkedOrgSet.has(plotPath)) {
      row.classList.add("is-selected");
    }

    const handle = document.createElement("span");
    handle.className = "drag-handle";
    setIconHtml(handle, "grip");
    handle.title = "Drag to reorder";
    handle.setAttribute("aria-hidden", "true");

    const chk = document.createElement("input");
    chk.type = "checkbox";
    chk.className = "plot-chk";
    chk.checked = state.checkedOrgSet.has(plotPath);
    chk.setAttribute("aria-label", `Select ${plotName}`);
    chk.addEventListener("change", () => {
      if (chk.checked) {
        state.checkedOrgSet.add(plotPath);
      } else {
        state.checkedOrgSet.delete(plotPath);
      }
      row.classList.toggle("is-selected", chk.checked);
      updateBulkBarUI();
      updateSectionSelectButtons();
    });

    const thumbBtn = document.createElement("button");
    thumbBtn.type = "button";
    thumbBtn.className = "thumb-btn";
    thumbBtn.setAttribute("aria-label", `Preview ${plotName}`);
    thumbBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      showImagePreview(plotPath);
    });

    const thumb = document.createElement("img");
    thumb.className = "plot-thumb";
    thumb.width = 44;
    thumb.height = 32;
    thumb.alt = plotName;
    thumb.dataset.src = thumbUrl(plotPath, 100);
    thumbBtn.appendChild(thumb);

    const figBadge = document.createElement("span");
    figBadge.className = "fig-badge mono-text";

    const info = document.createElement("div");
    info.className = "plot-info";
    const nameEl = document.createElement("span");
    nameEl.className = "plot-name";
    nameEl.textContent = plotName;

    const pathEl = document.createElement("span");
    pathEl.className = "plot-path mono-text";
    pathEl.textContent = plotPath;
    pathEl.title = plotPath;

    info.append(nameEl, pathEl);

    const statusBadge = document.createElement("span");
    statusBadge.className = "tag tag-warn plot-missing-tag hidden";
    statusBadge.textContent = "Missing";
    const st = state.fileStatusMap[plotPath];
    if (st && st.exists === false) {
      statusBadge.classList.remove("hidden");
    }

    const actions = document.createElement("div");
    actions.className = "plot-actions";

    const previewBtn = document.createElement("button");
    previewBtn.type = "button";
    previewBtn.className = "btn btn-sm btn-ghost";
    setIconHtml(previewBtn, "eye");
    previewBtn.title = "Preview plot";
    previewBtn.setAttribute("aria-label", `Preview ${plotName}`);
    previewBtn.addEventListener("click", () => showImagePreview(plotPath));

    const upBtn = document.createElement("button");
    upBtn.type = "button";
    upBtn.className = "btn btn-sm btn-secondary plot-up-btn";
    setIconHtml(upBtn, "up");
    upBtn.title = "Move up";
    upBtn.setAttribute("aria-label", `Move ${plotName} up`);
    upBtn.addEventListener("click", () => movePlotAdjacentInPlace(row, -1));

    const downBtn = document.createElement("button");
    downBtn.type = "button";
    downBtn.className = "btn btn-sm btn-secondary plot-down-btn";
    setIconHtml(downBtn, "down");
    downBtn.title = "Move down";
    downBtn.setAttribute("aria-label", `Move ${plotName} down`);
    downBtn.addEventListener("click", () => movePlotAdjacentInPlace(row, 1));

    actions.append(previewBtn, upBtn, downBtn);

    row.append(handle, chk, thumbBtn, figBadge, info, statusBadge, actions);

    // Drag events
    row.addEventListener("dragstart", (e) => onPlotDragStart(e, row));
    row.addEventListener("dragend", () => onDragEndCleanup());
    row.addEventListener("dragover", (e) => onPlotRowDragOver(e, row));
    row.addEventListener("dragleave", () => row.classList.remove("drop-before", "drop-after"));
    row.addEventListener("drop", (e) => onPlotRowDrop(e, row));

    return row;
  }

  // Sync state.sections from the authoritative DOM order + renumber everything IN PLACE
  function syncStateAndDomInPlace() {
    const container = document.getElementById("sections-container");
    const secMap = new Map();
    walkSections(state.sections, (s) => secMap.set(s.id, s));
    const isFlat = state.groupingMode === "flat";
    const isManual = state.groupingMode === "manual";

    function syncLevel(cards, depth, numPrefix, secNumForFigs, secFigCounter) {
      const result = [];
      cards.forEach((card, idx) => {
        const secId = card.dataset.secId;
        const sec = secMap.get(secId);
        if (!sec) return;

        const levelNum = depth === 0 ? String(idx + 1) : `${numPrefix}.${idx + 1}`;
        const figSecNum = depth === 0 ? (isFlat ? null : idx + 1) : secNumForFigs;
        const figCounter = depth === 0 ? [0] : secFigCounter;

        const plotsBox = card.querySelector(":scope > .section-plots");
        const rows = Array.from(plotsBox.querySelectorAll(":scope > .plot-row"));
        sec.plots = rows.map((r) => r.dataset.path);

        // Empty placeholder inside plotsBox
        let emptyNote = plotsBox.querySelector(":scope > .section-empty-note");
        if (rows.length === 0) {
          if (!emptyNote) {
            emptyNote = document.createElement("div");
            emptyNote.className = "section-empty-note";
            emptyNote.textContent = "Drop plots here or move plots into this section.";
            plotsBox.appendChild(emptyNote);
          }
        } else if (emptyNote) {
          emptyNote.remove();
        }

        // Renumber Figure X.Y badges (Level-1 section-scoped) & row up/down buttons in place
        rows.forEach((row, pIdx) => {
          figCounter[0] += 1;
          const badge = row.querySelector(".fig-badge");
          if (badge) {
            badge.textContent = isFlat
              ? `Figure ${figCounter[0]}`
              : `Figure ${figSecNum}.${figCounter[0]}`;
          }
          const upBtn = row.querySelector(".plot-up-btn");
          const downBtn = row.querySelector(".plot-down-btn");
          if (upBtn) upBtn.disabled = pIdx === 0;
          if (downBtn) downBtn.disabled = pIdx === rows.length - 1;
        });

        const subsBox = card.querySelector(":scope > .section-subsections");
        if (subsBox && depth < MAX_SUBSECTION_DEPTH) {
          const childCards = Array.from(subsBox.querySelectorAll(":scope > .section-card"));
          sec.subsections = syncLevel(childCards, depth + 1, levelNum, figSecNum, figCounter);
        } else {
          sec.subsections = [];
        }

        // Update section header text & count in place (direct + total if nested)
        const titleEl = card.querySelector(":scope > .section-header .sec-title-text");
        if (titleEl) {
          if (isFlat) {
            titleEl.textContent = "All Plots (No Sections)";
          } else if (depth === 0) {
            titleEl.textContent = `Section ${levelNum}: ${sec.name}`;
          } else {
            titleEl.textContent = `${levelNum} ${sec.name}`;
          }
        }
        const totalInBranch = collectSectionPlots(sec, depth).length;
        const countEl = card.querySelector(":scope > .section-header .sec-count-badge");
        if (countEl) {
          if (sec.subsections.length > 0 && totalInBranch !== rows.length) {
            countEl.textContent = `(${rows.length} direct, ${totalInBranch} total)`;
          } else {
            countEl.textContent = `(${rows.length} plot${rows.length === 1 ? "" : "s"})`;
          }
        }

        // Section up/down/delete button visibility
        const headerEl = card.querySelector(":scope > .section-header");
        const upSec = headerEl?.querySelector(".sec-up-btn");
        const downSec = headerEl?.querySelector(".sec-down-btn");
        const delSec = headerEl?.querySelector(".sec-del-btn");
        if (upSec) upSec.disabled = idx === 0;
        if (downSec) downSec.disabled = idx === cards.length - 1;
        if (delSec) delSec.disabled = !isManual || (depth === 0 && cards.length <= 1);

        result.push(sec);
      });
      return result;
    }

    const topCards = Array.from(container.querySelectorAll(":scope > .section-card"));
    state.sections = syncLevel(topCards, 0, "", null, [0]);
    updateSectionSelectButtons();
    updateBulkBarUI();
    updateHeaderBadges();
    applyOrgSearchHighlight();
  }

  function updateSectionSelectButtons() {
    const container = document.getElementById("sections-container");
    walkSections(state.sections, (sec) => {
      const card = container.querySelector(`.section-card[data-sec-id="${CSS.escape(sec.id)}"]`);
      if (!card) return;
      const btn = card.querySelector(":scope > .section-header .sec-select-btn");
      if (!btn) return;
      const allChecked = sec.plots.length > 0 && sec.plots.every((p) => state.checkedOrgSet.has(p));
      btn.textContent = allChecked ? "Deselect Section" : "Select Section";
    });
  }

  function listSectionTargets(list = state.sections, depth = 0, prefix = "") {
    const out = [];
    (list || []).forEach((sec, idx) => {
      const labelNum = depth === 0 ? `${idx + 1}.` : `${prefix}.${idx + 1}`;
      out.push({ sec, depth, labelNum });
      if (depth < MAX_SUBSECTION_DEPTH && sec.subsections?.length) {
        out.push(...listSectionTargets(sec.subsections, depth + 1, depth === 0 ? String(idx + 1) : labelNum));
      }
    });
    return out;
  }

  function updateBulkBarUI() {
    const count = state.checkedOrgSet.size;
    document.getElementById("org-checked-count").textContent = `${count} plot(s) checked`;

    const targets = listSectionTargets();
    const hasChecked = count > 0;
    const isManual = state.groupingMode === "manual";
    const multiSec = targets.length > 1;

    document.getElementById("btn-bulk-up").disabled = !hasChecked;
    document.getElementById("btn-bulk-down").disabled = !hasChecked;
    document.getElementById("btn-bulk-remove").disabled = !hasChecked;

    const targetSelect = document.getElementById("bulk-target-section");
    const prevVal = targetSelect.value;
    targetSelect.innerHTML = "";
    for (const { sec, depth, labelNum } of targets) {
      const opt = document.createElement("option");
      opt.value = sec.id;
      const indent = "\u00a0\u00a0".repeat(depth);
      opt.textContent = `${indent}${labelNum} ${sec.name || "Section"}`;
      targetSelect.appendChild(opt);
    }
    if (prevVal && targets.some((t) => t.sec.id === prevVal)) {
      targetSelect.value = prevVal;
    }

    const canMoveSec = hasChecked && isManual && multiSec;
    targetSelect.disabled = !canMoveSec;
    document.getElementById("btn-bulk-move").disabled = !canMoveSec;
  }

  function applyOrgSearchHighlight() {
    const q = state.orgSearch.trim().toLowerCase();
    document.querySelectorAll("#sections-container .plot-row").forEach((row) => {
      if (!q) {
        row.classList.remove("is-search-match", "is-search-dim");
      } else {
        const match = row.dataset.path.toLowerCase().includes(q);
        row.classList.toggle("is-search-match", match);
        row.classList.toggle("is-search-dim", !match);
      }
    });
  }

  // In-place operations
  function toggleSectionPlotsChecked(secId) {
    const hit = findSectionEntry(secId);
    if (!hit) return;
    const sec = hit.sec;
    const allChecked = sec.plots.length > 0 && sec.plots.every((p) => state.checkedOrgSet.has(p));
    for (const p of sec.plots) {
      if (allChecked) state.checkedOrgSet.delete(p);
      else state.checkedOrgSet.add(p);
    }
    const card = document.querySelector(`.section-card[data-sec-id="${CSS.escape(secId)}"]`);
    if (card) {
      card.querySelectorAll(":scope > .section-plots > .plot-row").forEach((row) => {
        const isChk = state.checkedOrgSet.has(row.dataset.path);
        row.classList.toggle("is-selected", isChk);
        const input = row.querySelector(".plot-chk");
        if (input) input.checked = isChk;
      });
    }
    updateSectionSelectButtons();
    updateBulkBarUI();
  }

  function setAllOrgChecked(checked) {
    if (checked) {
      walkSections(state.sections, (sec) => {
        for (const p of sec.plots) state.checkedOrgSet.add(p);
      });
    } else {
      state.checkedOrgSet.clear();
    }
    document.querySelectorAll("#sections-container .plot-row").forEach((row) => {
      const isChk = state.checkedOrgSet.has(row.dataset.path);
      row.classList.toggle("is-selected", isChk);
      const input = row.querySelector(".plot-chk");
      if (input) input.checked = isChk;
    });
    updateSectionSelectButtons();
    updateBulkBarUI();
  }

  function movePlotAdjacentInPlace(row, direction) {
    const parent = row.parentElement;
    if (!parent) return;
    if (direction < 0) {
      const prev = row.previousElementSibling;
      if (prev && prev.classList.contains("plot-row")) {
        parent.insertBefore(row, prev);
      }
    } else {
      const next = row.nextElementSibling;
      if (next && next.classList.contains("plot-row")) {
        parent.insertBefore(next, row);
      }
    }
    syncStateAndDomInPlace();
  }

  function bulkReorderInPlace(direction) {
    if (state.checkedOrgSet.size === 0) return;
    const container = document.getElementById("sections-container");
    container.querySelectorAll(".section-plots").forEach((plotsBox) => {
      const rows = Array.from(plotsBox.querySelectorAll(":scope > .plot-row"));
      if (direction < 0) {
        for (let i = 1; i < rows.length; i++) {
          const cur = rows[i];
          const prev = cur.previousElementSibling;
          if (
            prev &&
            prev.classList.contains("plot-row") &&
            state.checkedOrgSet.has(cur.dataset.path) &&
            !state.checkedOrgSet.has(prev.dataset.path)
          ) {
            plotsBox.insertBefore(cur, prev);
          }
        }
      } else {
        for (let i = rows.length - 2; i >= 0; i--) {
          const cur = rows[i];
          const next = cur.nextElementSibling;
          if (
            next &&
            next.classList.contains("plot-row") &&
            state.checkedOrgSet.has(cur.dataset.path) &&
            !state.checkedOrgSet.has(next.dataset.path)
          ) {
            plotsBox.insertBefore(next, cur);
          }
        }
      }
    });
    syncStateAndDomInPlace();
  }

  function bulkMoveCheckedToSection(targetSecId) {
    if (state.checkedOrgSet.size === 0 || state.groupingMode !== "manual") return;
    const container = document.getElementById("sections-container");
    const targetBox = container.querySelector(`.section-plots[data-sec-id="${CSS.escape(targetSecId)}"]`);
    if (!targetBox) return;

    const checkedRows = Array.from(container.querySelectorAll(".plot-row")).filter((r) =>
      state.checkedOrgSet.has(r.dataset.path)
    );
    for (const row of checkedRows) {
      targetBox.appendChild(row);
    }
    syncStateAndDomInPlace();
  }

  function bulkRemoveCheckedInPlace() {
    if (state.checkedOrgSet.size === 0) return;
    const removedCount = state.checkedOrgSet.size;
    for (const p of Array.from(state.checkedOrgSet)) {
      state.selectedSet.delete(p);
      state.checkedOrgSet.delete(p);
    }
    document.querySelectorAll("#sections-container .plot-row").forEach((row) => {
      if (!state.selectedSet.has(row.dataset.path)) {
        row.remove();
      }
    });
    syncStateAndDomInPlace();
    if (state.selectedSet.size === 0 || state.groupingMode !== "manual") {
      reconcileSectionsWithSelection();
      if (state.sections.length === 0) {
        renderSectionsFull();
      }
    }
    renderStep1Tree();
    showToast(`Removed ${removedCount} plot(s) from the report.`, "info");
  }

  function moveSectionInPlace(secId, direction) {
    const container = document.getElementById("sections-container");
    const card = container.querySelector(`.section-card[data-sec-id="${CSS.escape(secId)}"]`);
    if (!card || !card.parentElement) return;
    const listParent = card.parentElement;
    if (direction < 0 && card.previousElementSibling) {
      listParent.insertBefore(card, card.previousElementSibling);
    } else if (direction > 0 && card.nextElementSibling) {
      listParent.insertBefore(card.nextElementSibling, card);
    }
    syncStateAndDomInPlace();
  }

  function addSectionModal() {
    if (state.groupingMode !== "manual") return;
    const input = document.createElement("input");
    input.type = "text";
    input.className = "input";
    input.style.width = "100%";
    input.value = `Section ${state.secCounter}`;

    const commitAdd = () => {
      const name = input.value.trim();
      if (!name) return;
      const newSec = {
        id: `sec_${state.secCounter++}`,
        name,
        page_break: true,
        collapsed: false,
        plots: [],
        subsections: [],
      };
      state.sections.push(newSec);
      const container = document.getElementById("sections-container");
      document.getElementById("organize-empty-state").classList.add("hidden");
      container.classList.remove("hidden");
      container.appendChild(createSectionCardElement(newSec, 0));
      syncStateAndDomInPlace();
      closeModal();
    };

    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        commitAdd();
      }
    });

    openModal({
      title: "Add New Section",
      bodyEl: input,
      buttons: [
        { label: "Cancel", className: "btn-secondary", onClick: closeModal },
        { label: "Add Section", className: "btn-primary", onClick: commitAdd },
      ],
    });
    setTimeout(() => { input.focus(); input.select(); }, 30);
  }

  function addSubsectionModal(parentSec, nextDepth) {
    if (state.groupingMode !== "manual" || nextDepth > MAX_SUBSECTION_DEPTH) return;
    const input = document.createElement("input");
    input.type = "text";
    input.className = "input";
    input.style.width = "100%";
    input.value = `Subsection ${(parentSec.subsections?.length || 0) + 1}`;

    const commitAddSub = () => {
      const name = input.value.trim();
      if (!name) return;
      if (!parentSec.subsections) parentSec.subsections = [];
      const newSub = {
        id: `sec_${state.secCounter++}`,
        name,
        page_break: false,
        collapsed: false,
        plots: [],
        subsections: [],
      };
      parentSec.subsections.push(newSub);
      parentSec.collapsed = false;
      const container = document.getElementById("sections-container");
      const parentCard = container.querySelector(`.section-card[data-sec-id="${CSS.escape(parentSec.id)}"]`);
      const subsBox = parentCard?.querySelector(":scope > .section-subsections");
      if (parentCard && subsBox) {
        parentCard.classList.remove("collapsed");
        subsBox.appendChild(createSectionCardElement(newSub, nextDepth));
        syncStateAndDomInPlace();
      } else {
        renderSectionsFull();
      }
      closeModal();
    };

    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        commitAddSub();
      }
    });

    openModal({
      title: `Add Subsection under "${parentSec.name}"`,
      bodyEl: input,
      buttons: [
        { label: "Cancel", className: "btn-secondary", onClick: closeModal },
        { label: "Add Subsection", className: "btn-primary", onClick: commitAddSub },
      ],
    });
    setTimeout(() => { input.focus(); input.select(); }, 30);
  }

  function renameSectionModal(sec) {
    const input = document.createElement("input");
    input.type = "text";
    input.className = "input";
    input.style.width = "100%";
    input.value = sec.name;

    const commitRename = () => {
      const newName = input.value.trim();
      if (!newName) return;
      sec.name = newName;
      syncStateAndDomInPlace();
      closeModal();
    };

    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        commitRename();
      }
    });

    openModal({
      title: "Rename Section",
      bodyEl: input,
      buttons: [
        { label: "Cancel", className: "btn-secondary", onClick: closeModal },
        { label: "Save Name", className: "btn-primary", onClick: commitRename },
      ],
    });
    setTimeout(() => { input.focus(); input.select(); }, 30);
  }

  function deleteSectionInPlace(secId) {
    if (state.groupingMode !== "manual") return;
    const hit = findSectionEntry(secId);
    if (!hit) return;
    const container = document.getElementById("sections-container");
    const card = container.querySelector(`.section-card[data-sec-id="${CSS.escape(secId)}"]`);
    if (!card) return;

    let targetBox = null;
    if (hit.depth > 0 && hit.parent) {
      targetBox = container.querySelector(`.section-plots[data-sec-id="${CSS.escape(hit.parent.id)}"]`);
    } else if (state.sections.length > 1) {
      const targetSec = state.sections.find((s) => s.id !== secId);
      if (targetSec) {
        targetBox = container.querySelector(`.section-plots[data-sec-id="${CSS.escape(targetSec.id)}"]`);
      }
    } else {
      return;
    }

    if (targetBox) {
      card.querySelectorAll(".plot-row").forEach((r) => {
        targetBox.appendChild(r);
      });
      card.remove();
    }
    syncStateAndDomInPlace();
  }

  // ----------------------------------------------------------------------- //
  // Drag and Drop Handlers (Plot & Section)                                 //
  // ----------------------------------------------------------------------- //
  function clearDropIndicators() {
    document.querySelectorAll(".drop-before, .drop-after, .drag-over-before, .drag-over-after, .drag-over-empty").forEach((el) => {
      el.classList.remove("drop-before", "drop-after", "drag-over-before", "drag-over-after", "drag-over-empty");
    });
  }

  function onDragEndCleanup() {
    document.querySelectorAll(".dragging").forEach((el) => el.classList.remove("dragging"));
    clearDropIndicators();
    dragContext = null;
  }

  function onPlotDragStart(e, row) {
    const plotPath = row.dataset.path;
    const isMulti = state.checkedOrgSet.has(plotPath) && state.checkedOrgSet.size > 1;
    let paths = [plotPath];
    if (isMulti) {
      const scopeEl = state.groupingMode === "manual"
        ? document.getElementById("sections-container")
        : (row.closest(".section-plots") || document.getElementById("sections-container"));
      paths = Array.from(scopeEl.querySelectorAll(".plot-row"))
        .map((r) => r.dataset.path)
        .filter((p) => state.checkedOrgSet.has(p));
    }

    dragContext = { type: "plot", paths };
    e.dataTransfer.effectAllowed = "move";
    e.dataTransfer.setData("text/plain", plotPath);

    const pathSet = new Set(paths);
    document.querySelectorAll("#sections-container .plot-row").forEach((r) => {
      if (pathSet.has(r.dataset.path)) r.classList.add("dragging");
    });
  }

  function onPlotRowDragOver(e, targetRow) {
    if (!dragContext || dragContext.type !== "plot") return;
    if (dragContext.paths.includes(targetRow.dataset.path)) return;

    const targetSecId = targetRow.closest(".section-plots")?.dataset.secId;
    const firstDraggedRow = document.querySelector(`.plot-row[data-path="${CSS.escape(dragContext.paths[0])}"]`);
    const sourceSecId = firstDraggedRow?.closest(".section-plots")?.dataset.secId;
    if (state.groupingMode !== "manual" && targetSecId !== sourceSecId) {
      return; // In auto/flat mode keep plots inside their folder section
    }

    e.preventDefault();
    e.stopPropagation();
    const rect = targetRow.getBoundingClientRect();
    const before = (e.clientY - rect.top) < rect.height / 2;
    clearDropIndicators();
    targetRow.classList.add(before ? "drop-before" : "drop-after");
  }

  function onPlotRowDrop(e, targetRow) {
    if (!dragContext || dragContext.type !== "plot") return;
    if (dragContext.paths.includes(targetRow.dataset.path)) return;

    const targetBox = targetRow.closest(".section-plots");
    if (!targetBox) return;

    const firstDraggedRow = document.querySelector(`.plot-row[data-path="${CSS.escape(dragContext.paths[0])}"]`);
    const sourceSecId = firstDraggedRow?.closest(".section-plots")?.dataset.secId;
    if (state.groupingMode !== "manual" && targetBox.dataset.secId !== sourceSecId) {
      onDragEndCleanup();
      return;
    }

    e.preventDefault();
    e.stopPropagation();
    const rect = targetRow.getBoundingClientRect();
    const before = (e.clientY - rect.top) < rect.height / 2;
    const refNode = before ? targetRow : targetRow.nextElementSibling;

    const container = document.getElementById("sections-container");
    const draggedRows = dragContext.paths
      .map((p) => container.querySelector(`.plot-row[data-path="${CSS.escape(p)}"]`))
      .filter(Boolean);

    for (const dr of draggedRows) {
      targetBox.insertBefore(dr, refNode);
    }

    onDragEndCleanup();
    syncStateAndDomInPlace();
  }

  function onSectionPlotsDragOver(e, plotsBox) {
    if (!dragContext || dragContext.type !== "plot") return;
    if (state.groupingMode !== "manual") {
      const firstDraggedRow = document.querySelector(`.plot-row[data-path="${CSS.escape(dragContext.paths[0])}"]`);
      if (firstDraggedRow?.closest(".section-plots") !== plotsBox) return;
    }
    if (e.target === plotsBox || e.target.classList.contains("section-empty-note")) {
      e.preventDefault();
      clearDropIndicators();
      plotsBox.classList.add("drag-over-empty");
    }
  }

  function onSectionPlotsDrop(e, plotsBox) {
    if (!dragContext || dragContext.type !== "plot") return;
    if (state.groupingMode !== "manual") {
      const firstDraggedRow = document.querySelector(`.plot-row[data-path="${CSS.escape(dragContext.paths[0])}"]`);
      if (firstDraggedRow?.closest(".section-plots") !== plotsBox) {
        onDragEndCleanup();
        return;
      }
    }
    e.preventDefault();
    const container = document.getElementById("sections-container");
    const draggedRows = dragContext.paths
      .map((p) => container.querySelector(`.plot-row[data-path="${CSS.escape(p)}"]`))
      .filter(Boolean);
    for (const dr of draggedRows) {
      plotsBox.appendChild(dr);
    }
    onDragEndCleanup();
    syncStateAndDomInPlace();
  }

  function onSectionDragStart(e, secId, card) {
    dragContext = { type: "section", secId };
    e.dataTransfer.effectAllowed = "move";
    e.dataTransfer.setData("text/plain", secId);
    card.classList.add("dragging");
  }

  function onSectionCardDragOver(e, targetCard) {
    if (!dragContext || dragContext.type !== "section") return;
    if (dragContext.secId === targetCard.dataset.secId) return;
    const container = document.getElementById("sections-container");
    const sourceCard = container.querySelector(`.section-card[data-sec-id="${CSS.escape(dragContext.secId)}"]`);
    if (!sourceCard || sourceCard.parentElement !== targetCard.parentElement) return;
    e.preventDefault();
    e.stopPropagation();
    const rect = targetCard.getBoundingClientRect();
    const before = (e.clientY - rect.top) < rect.height / 2;
    clearDropIndicators();
    targetCard.classList.add(before ? "drag-over-before" : "drag-over-after");
  }

  function onSectionCardDrop(e, targetCard, targetSecId) {
    if (!dragContext || dragContext.type !== "section") return;
    if (dragContext.secId === targetSecId) return;
    const container = document.getElementById("sections-container");
    const sourceCard = container.querySelector(`.section-card[data-sec-id="${CSS.escape(dragContext.secId)}"]`);
    if (!sourceCard || sourceCard.parentElement !== targetCard.parentElement) {
      onDragEndCleanup();
      return;
    }
    e.preventDefault();
    e.stopPropagation();
    const rect = targetCard.getBoundingClientRect();
    const before = (e.clientY - rect.top) < rect.height / 2;
    targetCard.parentElement.insertBefore(sourceCard, before ? targetCard : targetCard.nextElementSibling);
    onDragEndCleanup();
    syncStateAndDomInPlace();
  }

  // ----------------------------------------------------------------------- //
  // Step 3: Summary, Validation & Document Generation                       //
  // ----------------------------------------------------------------------- //
  function syncDocOptionsFromInputs() {
    const maxW = document.getElementById("opt-max-width").value.trim();
    const maxH = document.getElementById("opt-max-height").value.trim();
    state.docOptions.max_width_cm = maxW ? parseFloat(maxW) : null;
    state.docOptions.max_height_cm = maxH ? parseFloat(maxH) : null;
    state.docOptions.export_docx = document.getElementById("opt-export-docx").checked;
    state.docOptions.export_html = document.getElementById("opt-export-html").checked;
    state.docOptions.html_embed_images = document.getElementById("opt-html-embed").checked;
    state.docOptions.export_pdf = state.docOptions.export_docx
      ? document.getElementById("opt-export-pdf").checked
      : false;
  }

  function applyDocOptionsToInputs() {
    document.querySelectorAll("#orient-segmented .seg-btn").forEach((b) => {
      const isActive = b.dataset.value === state.docOptions.orientation;
      b.classList.toggle("active", isActive);
      b.setAttribute("aria-pressed", String(isActive));
    });
    document.querySelectorAll("#cols-segmented .seg-btn").forEach((b) => {
      const isActive = Number(b.dataset.value) === Number(state.docOptions.num_cols);
      b.classList.toggle("active", isActive);
      b.setAttribute("aria-pressed", String(isActive));
    });
    document.getElementById("opt-max-width").value =
      state.docOptions.max_width_cm != null ? String(state.docOptions.max_width_cm) : "";
    document.getElementById("opt-max-height").value =
      state.docOptions.max_height_cm != null ? String(state.docOptions.max_height_cm) : "";

    const docxChk = document.getElementById("opt-export-docx");
    const htmlChk = document.getElementById("opt-export-html");
    const embedChk = document.getElementById("opt-html-embed");
    const embedRow = document.getElementById("html-embed-row");
    const pdfChk = document.getElementById("opt-export-pdf");
    const pdfNote = document.getElementById("pdf-requires-docx-note");

    docxChk.checked = state.docOptions.export_docx !== false;
    htmlChk.checked = Boolean(state.docOptions.export_html);
    embedChk.checked = Boolean(state.docOptions.html_embed_images);
    embedRow.classList.toggle("hidden", !htmlChk.checked);

    if (!docxChk.checked) {
      state.docOptions.export_pdf = false;
      pdfChk.checked = false;
      pdfChk.disabled = true;
      pdfNote.classList.remove("hidden");
    } else {
      pdfChk.disabled = false;
      pdfChk.checked = Boolean(state.docOptions.export_pdf);
      pdfNote.classList.add("hidden");
    }

    const genBtn = document.getElementById("btn-generate-as");
    if (docxChk.checked && htmlChk.checked) {
      genBtn.textContent = "Choose Output & Generate (.docx + .html)";
    } else if (htmlChk.checked) {
      genBtn.textContent = "Choose Output & Generate (.html)";
    } else {
      genBtn.textContent = "Choose Output & Generate (.docx)";
    }

    const quickBtn = document.getElementById("btn-quick-regen");
    const hint = document.getElementById("last-output-hint");
    if (state.docOptions.last_output_path) {
      quickBtn.classList.remove("hidden");
      hint.classList.remove("hidden");
      hint.textContent = `Last output: ${state.docOptions.last_output_path}`;
    } else {
      quickBtn.classList.add("hidden");
      hint.classList.add("hidden");
    }

    const totalPlots = state.sections.reduce((acc, s) => acc + collectSectionPlots(s, 0).length, 0);
    const hasFormat = docxChk.checked || htmlChk.checked;
    const isRunning = Boolean(state.genPollTimer);
    genBtn.disabled = isRunning || totalPlots === 0 || !hasFormat;
    quickBtn.disabled = isRunning || totalPlots === 0 || !hasFormat;
  }

  function renderGenerateSummary() {
    applyDocOptionsToInputs();
    const nonEmpty = state.sections.filter((s) => collectSectionPlots(s, 0).length > 0);
    const allPlots = nonEmpty.flatMap((s) => collectSectionPlots(s, 0));
    const missingCount = allPlots.filter((p) => state.fileStatusMap[p]?.exists === false).length;

    const statsEl = document.getElementById("gen-summary-stats");
    statsEl.innerHTML = "";
    const addPill = (text, cls = "tag-accent") => {
      const span = document.createElement("span");
      span.className = `tag ${cls}`;
      span.textContent = text;
      statsEl.appendChild(span);
    };
    addPill(`${nonEmpty.length} section(s)`);
    addPill(`${allPlots.length} plot(s)`);
    if (missingCount > 0) addPill(`${missingCount} missing`, "tag-warn");
    else if (allPlots.length > 0) addPill("Files checked", "tag-ok");

    const missingAlert = document.getElementById("gen-missing-alert");
    if (missingCount > 0) {
      missingAlert.classList.remove("hidden");
      missingAlert.textContent = `${missingCount} plot file(s) could not be found on disk. Use "Re-check Files" or "Change Source Folder" if plots were moved.`;
    } else {
      missingAlert.classList.add("hidden");
    }

    const listEl = document.getElementById("gen-section-list");
    listEl.innerHTML = "";
    if (nonEmpty.length === 0) {
      const empty = document.createElement("div");
      empty.className = "empty-state";
      const msg = document.createElement("p");
      msg.textContent = "No plots selected yet. Select and organize plots first.";
      const goBtn = document.createElement("button");
      goBtn.type = "button";
      goBtn.className = "btn btn-sm btn-primary";
      goBtn.textContent = "Go to Select Plots";
      goBtn.addEventListener("click", () => switchStep("select"));
      empty.append(msg, goBtn);
      listEl.appendChild(empty);
      return;
    }

    function appendSummaryNodes(nodes, depth, numPrefix) {
      let idx = 0;
      for (const sec of nodes) {
        if (collectSectionPlots(sec, depth).length === 0) continue;
        idx += 1;
        const labelNum = depth === 0 ? String(idx) : `${numPrefix}.${idx}`;
        const item = document.createElement("div");
        item.className = "gen-sec-item";
        if (depth > 0) item.style.marginLeft = `${depth * 14}px`;

        const head = document.createElement("div");
        head.className = "gen-sec-head";
        const left = document.createElement("span");
        if (state.groupingMode === "flat") {
          left.textContent = "All Plots (Flat)";
        } else if (depth === 0) {
          left.textContent = `Section ${labelNum}: ${sec.name}`;
        } else {
          left.textContent = `${labelNum} ${sec.name}`;
        }
        const right = document.createElement("span");
        right.className = "muted mono-text";
        right.textContent = `${sec.plots.length} plot(s)` + (sec.page_break && (depth > 0 || idx > 1) ? " · Page Break" : "");
        head.append(left, right);

        const preview = document.createElement("div");
        preview.className = "muted mono-text";
        preview.style.fontSize = "11.5px";
        if (sec.plots.length > 0) {
          const names = sec.plots.slice(0, 3).map(basename).join(", ");
          preview.textContent = names + (sec.plots.length > 3 ? ` (+${sec.plots.length - 3} more)` : "");
        } else {
          preview.textContent = "Contains subsections only";
        }
        item.append(head, preview);
        listEl.appendChild(item);

        if (depth < MAX_SUBSECTION_DEPTH && sec.subsections?.length) {
          appendSummaryNodes(sec.subsections, depth + 1, labelNum);
        }
      }
    }
    appendSummaryNodes(nonEmpty, 0, "");
  }

  async function validateCurrentPlots(silent = false) {
    const paths = Array.from(state.selectedSet);
    if (paths.length === 0) return;
    const recheckBtns = [
      document.getElementById("btn-recheck-files"),
      document.getElementById("btn-gen-recheck"),
    ];
    for (const b of recheckBtns) {
      b.disabled = true;
      b.textContent = "Checking...";
    }
    try {
      const res = await apiPost("/api/files/validate", { paths });
      state.fileStatusMap = res.file_status || {};
      updateConfigBanner();
      document.querySelectorAll("#sections-container .plot-row").forEach((row) => {
        const tag = row.querySelector(".plot-missing-tag");
        if (!tag) return;
        const ok = state.fileStatusMap[row.dataset.path]?.exists !== false;
        tag.classList.toggle("hidden", ok);
      });
      if (state.currentStep === "generate") {
        renderGenerateSummary();
      }
      if (!silent) {
        const missing = paths.filter((p) => state.fileStatusMap[p]?.exists === false).length;
        if (missing > 0) showToast(`${missing} plot(s) are missing on disk.`, "warn");
        else showToast(`All ${paths.length} plot files verified on disk.`, "success");
      }
    } catch (err) {
      if (!silent) showToast(err.message, "error");
    } finally {
      for (const b of recheckBtns) {
        b.disabled = false;
        b.textContent = "Re-check Files";
      }
    }
  }

  function serializeSectionsForGenerate(list, depth = 0) {
    const isFlat = state.groupingMode === "flat";
    const out = [];
    for (const s of list || []) {
      if (collectSectionPlots(s, depth).length === 0) continue;
      out.push({
        title: isFlat ? "" : s.name,
        page_break: isFlat ? false : Boolean(s.page_break),
        plots: s.plots,
        subsections: !isFlat && depth < MAX_SUBSECTION_DEPTH
          ? serializeSectionsForGenerate(s.subsections || [], depth + 1)
          : [],
      });
    }
    return out;
  }

  async function runGeneration(saveFilePath) {
    syncDocOptionsFromInputs();
    if (!state.docOptions.export_docx && !state.docOptions.export_html) {
      showToast("Select at least one output format (.docx or .html).", "warn");
      return;
    }
    const payloadSections = serializeSectionsForGenerate(state.sections, 0);
    if (payloadSections.length === 0) {
      showToast("No plots to generate. Select and organize plots first.", "warn");
      return;
    }

    if (state.docOptions.max_width_cm !== null && (Number.isNaN(state.docOptions.max_width_cm) || state.docOptions.max_width_cm <= 0)) {
      showToast("Max Width must be a positive number.", "error");
      return;
    }
    if (state.docOptions.max_height_cm !== null && (Number.isNaN(state.docOptions.max_height_cm) || state.docOptions.max_height_cm <= 0)) {
      showToast("Max Height must be a positive number.", "error");
      return;
    }

    const defaultExt = state.docOptions.export_docx ? ".docx" : ".html";
    let chosenPath = saveFilePath;
    if (!chosenPath) {
      const initialDir = state.docOptions.last_output_path ? dirname(state.docOptions.last_output_path) : state.sourceFolder;
      if (!state.hasDialogs) {
        const fallbackDefault = initialDir ? `${initialDir}\\Plot_Report${defaultExt}` : `Plot_Report${defaultExt}`;
        chosenPath = await promptPathModal({
          title: defaultExt === ".html" ? "Save Generated HTML Report" : "Save Generated Report",
          label: `Output ${defaultExt} path`,
          defaultValue: fallbackDefault,
          placeholder: `e.g. C:\\Projects\\Plot_Report${defaultExt}`,
          submitLabel: "Generate Report",
        });
        if (!chosenPath) return;
      } else {
        try {
          const dlg = await apiPost("/api/dialog/save-docx", { initial_dir: initialDir, default_ext: defaultExt });
          if (dlg.cancelled || !dlg.path) return;
          chosenPath = dlg.path;
        } catch (err) {
          showToast(err.message, "error");
          return;
        }
      }
    }

    state.docOptions.last_output_path = chosenPath;
    applyDocOptionsToInputs();

    const progBox = document.getElementById("gen-progress-box");
    const resBox = document.getElementById("gen-result-box");
    progBox.classList.remove("hidden");
    resBox.classList.add("hidden");
    document.getElementById("gen-progress-fill").style.width = "5%";
    document.getElementById("gen-progress-text").textContent = "Starting generation...";
    document.getElementById("btn-generate-as").disabled = true;
    document.getElementById("btn-quick-regen").disabled = true;

    try {
      const start = await apiPost("/api/generate/start", {
        sections: payloadSections,
        save_file: chosenPath,
        orientation: state.docOptions.orientation,
        num_cols: state.docOptions.num_cols,
        max_width_cm: state.docOptions.max_width_cm,
        max_height_cm: state.docOptions.max_height_cm,
        export_docx: state.docOptions.export_docx,
        export_html: state.docOptions.export_html,
        html_embed_images: state.docOptions.html_embed_images,
        export_pdf: state.docOptions.export_pdf,
      });
      state.genJobId = start.job_id;
      pollGenerationStatus();
    } catch (err) {
      document.getElementById("btn-generate-as").disabled = false;
      document.getElementById("btn-quick-regen").disabled = false;
      showToast(err.message, "error");
    }
  }

  async function pollGenerationStatus() {
    if (!state.genJobId) return;
    try {
      const st = await apiGet(`/api/generate/status?id=${encodeURIComponent(state.genJobId)}`);
      document.getElementById("gen-progress-fill").style.width = `${Math.round((st.progress || 0) * 100)}%`;
      document.getElementById("gen-progress-text").textContent = st.message || "Working...";

      if (st.status === "running") {
        state.genPollTimer = setTimeout(pollGenerationStatus, 250);
        return;
      }
      state.genPollTimer = null;

      applyDocOptionsToInputs();

      const resBox = document.getElementById("gen-result-box");
      resBox.classList.remove("hidden", "is-success", "is-error");
      resBox.innerHTML = "";

      if (st.status === "done" && st.result) {
        resBox.classList.add("is-success");
        const title = document.createElement("strong");
        title.textContent = `Generated report with ${st.result.count} plot(s)!`;
        resBox.appendChild(title);

        if (st.result.skipped && st.result.skipped.length > 0) {
          const skipEl = document.createElement("div");
          skipEl.className = "tag tag-warn";
          skipEl.textContent = `${st.result.skipped.length} plot(s) were skipped`;
          resBox.appendChild(skipEl);
        }

        if (st.result.docx_path) {
          const docxLine = document.createElement("div");
          docxLine.className = "muted mono-text";
          docxLine.textContent = `Word: ${st.result.docx_path}`;
          resBox.appendChild(docxLine);
        }

        if (st.result.html_path) {
          const htmlLine = document.createElement("div");
          htmlLine.className = "muted mono-text";
          htmlLine.textContent = `HTML: ${st.result.html_path}`;
          resBox.appendChild(htmlLine);
        }

        if (st.result.pdf_path) {
          const pdfLine = document.createElement("div");
          pdfLine.className = "muted mono-text";
          pdfLine.textContent = `PDF: ${st.result.pdf_path}`;
          resBox.appendChild(pdfLine);
        }

        const actions = document.createElement("div");
        actions.className = "result-actions";

        if (st.result.docx_path) {
          const openDocx = document.createElement("button");
          openDocx.type = "button";
          openDocx.className = "btn btn-sm btn-primary";
          openDocx.textContent = "Open Word Doc";
          openDocx.addEventListener("click", () => apiPost("/api/open-file", { path: st.result.docx_path }).catch((e) => showToast(e.message, "error")));
          actions.appendChild(openDocx);
        }

        if (st.result.html_path) {
          const openHtml = document.createElement("button");
          openHtml.type = "button";
          openHtml.className = "btn btn-sm btn-primary";
          openHtml.textContent = "Open HTML Report";
          openHtml.addEventListener("click", () => apiPost("/api/open-file", { path: st.result.html_path }).catch((e) => showToast(e.message, "error")));
          actions.appendChild(openHtml);
        }

        if (st.result.pdf_path) {
          const openPdf = document.createElement("button");
          openPdf.type = "button";
          openPdf.className = "btn btn-sm btn-success";
          openPdf.textContent = "Open PDF";
          openPdf.addEventListener("click", () => apiPost("/api/open-file", { path: st.result.pdf_path }).catch((e) => showToast(e.message, "error")));
          actions.appendChild(openPdf);
        }

        const primaryOutput = st.result.docx_path || st.result.html_path || st.result.pdf_path;
        if (primaryOutput) {
          const openDir = document.createElement("button");
          openDir.type = "button";
          openDir.className = "btn btn-sm btn-secondary";
          openDir.textContent = "Open Folder";
          openDir.addEventListener("click", () => apiPost("/api/open-file", { path: dirname(primaryOutput) }).catch((e) => showToast(e.message, "error")));
          actions.appendChild(openDir);
        }

        resBox.appendChild(actions);
        showToast("Document generated successfully!", "success");
      } else {
        resBox.classList.add("is-error");
        const title = document.createElement("strong");
        title.textContent = "Generation failed";
        const errPre = document.createElement("pre");
        errPre.style.whiteSpace = "pre-wrap";
        errPre.style.margin = "0";
        errPre.textContent = st.error || "Unknown error.";
        resBox.append(title, errPre);
        showToast("Document generation failed.", "error");
      }
    } catch (err) {
      state.genPollTimer = null;
      applyDocOptionsToInputs();
      showToast(err.message, "error");
    }
  }

  // ----------------------------------------------------------------------- //
  // Save / Load / Remap Config                                              //
  // ----------------------------------------------------------------------- //
  function serializeSectionsForConfig(list, depth = 0) {
    return (list || []).map((s) => ({
      id: s.id,
      name: s.name,
      page_break: s.page_break,
      plots: (s.plots || []).map((p) => ({
        path: p,
        rel_path: state.plotRelPaths.get(p) || null,
      })),
      subsections: depth < MAX_SUBSECTION_DEPTH
        ? serializeSectionsForConfig(s.subsections || [], depth + 1)
        : [],
    }));
  }

  function deserializeSectionsFromConfig(list, depth = 0, maxIdRef = { max: 0 }) {
    const out = [];
    for (const sec of list || []) {
      const plotPaths = [];
      for (const p of sec.plots || []) {
        const plotPath = typeof p === "string" ? p : p.path;
        if (!plotPath) continue;
        plotPaths.push(plotPath);
        if (typeof p === "object" && p.rel_path) {
          state.plotRelPaths.set(plotPath, p.rel_path);
        }
        state.selectedSet.add(plotPath);
      }
      const match = /^sec_(\d+)$/.exec(sec.id || "");
      if (match) maxIdRef.max = Math.max(maxIdRef.max, Number(match[1]));
      out.push({
        id: sec.id || `sec_${++maxIdRef.max}`,
        name: sec.name || "Section",
        page_break: "page_break" in sec ? Boolean(sec.page_break) : depth === 0,
        collapsed: false,
        plots: plotPaths,
        subsections: depth < MAX_SUBSECTION_DEPTH
          ? deserializeSectionsFromConfig(sec.subsections || [], depth + 1, maxIdRef)
          : [],
      });
    }
    return out;
  }

  function updateConfigBanner() {
    const banner = document.getElementById("config-banner");
    if (!state.configPath && !state.configName) {
      banner.classList.add("hidden");
      return;
    }
    banner.classList.remove("hidden");
    document.getElementById("config-banner-name").textContent = state.configName || basename(state.configPath);
    document.getElementById("config-banner-path").textContent = state.configPath || "";

    const paths = Array.from(state.selectedSet);
    const missing = paths.filter((p) => state.fileStatusMap[p]?.exists === false).length;
    const missBadge = document.getElementById("config-missing-badge");
    const okBadge = document.getElementById("config-ok-badge");
    if (missing > 0) {
      missBadge.textContent = `${missing} missing`;
      missBadge.classList.remove("hidden");
      okBadge.classList.add("hidden");
    } else if (paths.length > 0) {
      missBadge.classList.add("hidden");
      okBadge.classList.remove("hidden");
    } else {
      missBadge.classList.add("hidden");
      okBadge.classList.add("hidden");
    }
  }

  async function saveCurrentConfig() {
    reconcileSectionsWithSelection();
    syncDocOptionsFromInputs();
    try {
      const initialDir = state.configPath ? dirname(state.configPath) : state.sourceFolder;
      const initialFile = state.configPath ? basename(state.configPath) : "";
      let savePath = "";
      if (!state.hasDialogs) {
        const defaultVal = state.configPath || (initialDir ? `${initialDir}\\report.plot2docx.json` : "report.plot2docx.json");
        savePath = await promptPathModal({
          title: "Save Plot2Docx Configuration",
          label: "Config file path (.plot2docx.json)",
          defaultValue: defaultVal,
          placeholder: "e.g. C:\\Projects\\report.plot2docx.json",
          submitLabel: "Save Config",
        });
        if (!savePath) return;
      } else {
        const dlg = await apiPost("/api/dialog/save-config", { initial_dir: initialDir, initial_file: initialFile });
        if (dlg.cancelled || !dlg.path) return;
        savePath = dlg.path;
      }

      const payloadState = {
        name: state.configName || basename(savePath).replace(/\.plot2docx\.json$/i, "").replace(/\.json$/i, ""),
        source_folder: state.sourceFolder,
        grouping_mode: state.groupingMode,
        sections: serializeSectionsForConfig(state.sections, 0),
        doc_options: state.docOptions,
      };
      const res = await apiPost("/api/config/save", { path: savePath, state: payloadState });
      state.configPath = res.path;
      state.configName = res.config.name;
      state.fileStatusMap = res.file_status || {};
      walkSections(res.config.sections || [], (sec) => {
        for (const item of sec.plots || []) {
          if (item.path && item.rel_path) state.plotRelPaths.set(item.path, item.rel_path);
        }
      });
      updateConfigBanner();
      showToast(`Configuration saved to ${basename(res.path)}`, "success");
    } catch (err) {
      showToast(err.message, "error");
    }
  }

  async function loadConfigFromPath(configPath) {
    cancelActiveScanIfAny();
    document.getElementById("scan-banner").classList.add("hidden");
    document.getElementById("gen-progress-box").classList.add("hidden");
    document.getElementById("gen-result-box").classList.add("hidden");
    try {
      const res = await apiPost("/api/config/load", { path: configPath });
      const cfg = res.config;

      state.configPath = res.path;
      state.configName = cfg.name || basename(res.path);
      state.sourceFolder = cfg.source_folder || "";
      state.groupingMode = cfg.grouping_mode || "manual";
      state.docOptions = { ...state.docOptions, ...(cfg.doc_options || {}) };
      state.fileStatusMap = res.file_status || {};

      state.selectedSet.clear();
      state.checkedOrgSet.clear();
      state.plotRelPaths.clear();

      const maxIdRef = { max: 0 };
      state.sections = deserializeSectionsFromConfig(cfg.sections || [], 0, maxIdRef);
      state.secCounter = maxIdRef.max + 1;

      if (state.sourceFolder) {
        document.getElementById("folder-path-input").value = state.sourceFolder;
      }

      updateConfigBanner();
      applyDocOptionsToInputs();
      updateHeaderBadges();
      switchStep("organize");

      const total = state.selectedSet.size;
      const missing = Array.from(state.selectedSet).filter((p) => state.fileStatusMap[p]?.exists === false).length;
      if (missing > 0) {
        showToast(`Loaded config (${total} plots, ${missing} missing on disk).`, "warn", 5000);
      } else {
        showToast(`Loaded config "${state.configName}" (${total} plots ready).`, "success");
      }
    } catch (err) {
      showToast(err.message, "error");
    }
  }

  async function openConfigViaDialog() {
    try {
      const initialDir = state.configPath ? dirname(state.configPath) : state.sourceFolder;
      let chosenPath = "";
      if (!state.hasDialogs) {
        chosenPath = await promptPathModal({
          title: "Load Plot2Docx Configuration",
          label: "Config file path (.plot2docx.json)",
          defaultValue: state.configPath || "",
          placeholder: "e.g. C:\\Projects\\report.plot2docx.json",
          submitLabel: "Load Config",
        });
        if (!chosenPath) return;
      } else {
        const dlg = await apiPost("/api/dialog/open-config", { initial_dir: initialDir });
        if (dlg.cancelled || !dlg.path) return;
        chosenPath = dlg.path;
      }
      await loadConfigFromPath(chosenPath);
    } catch (err) {
      showToast(err.message, "error");
    }
  }

  async function remapSourceFolderViaDialog() {
    try {
      let newSource = "";
      if (!state.hasDialogs) {
        newSource = await promptPathModal({
          title: "Change Source Folder",
          label: "New base plots folder path",
          defaultValue: state.sourceFolder || "",
          placeholder: "e.g. C:\\Projects\\Plots_v2",
          submitLabel: "Remap Plots",
        });
        if (!newSource) return;
      } else {
        const dlg = await apiPost("/api/dialog/browse-folder", { initial_dir: state.sourceFolder });
        if (dlg.cancelled || !dlg.path) return;
        newSource = dlg.path;
      }

      const res = await apiPost("/api/config/remap", {
        old_source_folder: state.sourceFolder,
        sections: serializeSectionsForConfig(state.sections, 0),
        new_source_folder: newSource,
      });

      state.sourceFolder = res.source_folder;
      document.getElementById("folder-path-input").value = state.sourceFolder;
      state.selectedSet.clear();
      state.checkedOrgSet.clear();
      state.plotRelPaths.clear();
      state.fileStatusMap = res.file_status || {};

      const maxIdRef = { max: state.secCounter };
      state.sections = deserializeSectionsFromConfig(res.sections || [], 0, maxIdRef);

      updateConfigBanner();
      updateHeaderBadges();
      if (state.currentStep === "select") renderStep1Tree();
      if (state.currentStep === "organize") renderSectionsFull();
      if (state.currentStep === "generate") renderGenerateSummary();
      showToast(`Remapped to new folder: ${res.matched_count} matched, ${res.missing_count} missing.`, res.missing_count ? "warn" : "success");
    } catch (err) {
      showToast(err.message, "error");
    }
  }

  function closeRecentMenu() {
    const menu = document.getElementById("recent-menu");
    const caret = document.getElementById("btn-recent-menu");
    menu.classList.add("hidden");
    caret.setAttribute("aria-expanded", "false");
  }

  async function toggleRecentDropdown() {
    const menu = document.getElementById("recent-menu");
    const caret = document.getElementById("btn-recent-menu");
    if (!menu.classList.contains("hidden")) {
      closeRecentMenu();
      return;
    }
    try {
      const res = await apiGet("/api/config/recent");
      const list = res.recent || [];
      menu.innerHTML = "";
      if (list.length === 0) {
        const empty = document.createElement("div");
        empty.className = "section-empty-note";
        empty.textContent = "No recent configurations yet.";
        menu.appendChild(empty);
      } else {
        for (const item of list) {
          const btn = document.createElement("button");
          btn.type = "button";
          btn.className = "dropdown-item";
          btn.setAttribute("role", "menuitem");
          const title = document.createElement("strong");
          title.textContent = `${item.name} (${item.plot_count} plots)`;
          const path = document.createElement("span");
          path.className = "dropdown-item-path mono-text";
          path.textContent = item.path;
          btn.append(title, path);
          btn.addEventListener("click", () => {
            closeRecentMenu();
            loadConfigFromPath(item.path);
          });
          menu.appendChild(btn);
        }
      }
      menu.classList.remove("hidden");
      caret.setAttribute("aria-expanded", "true");
    } catch (err) {
      showToast(err.message, "error");
    }
  }

  // ----------------------------------------------------------------------- //
  // Event Wiring & Bootstrap                                                //
  // ----------------------------------------------------------------------- //
  function wireEvents() {
    document.getElementById("btn-theme-toggle").addEventListener("click", toggleTheme);
    document.getElementById("modal-close").addEventListener("click", closeModal);
    document.getElementById("modal-backdrop").addEventListener("click", (e) => {
      if (e.target.id === "modal-backdrop") closeModal();
    });

    document.addEventListener("keydown", (e) => {
      if (e.key !== "Escape") return;
      const backdrop = document.getElementById("modal-backdrop");
      if (!backdrop.classList.contains("hidden")) {
        closeModal();
        return;
      }
      const menu = document.getElementById("recent-menu");
      if (!menu.classList.contains("hidden")) {
        closeRecentMenu();
        document.getElementById("btn-recent-menu").focus();
      }
    });

    // Step navigation
    document.querySelectorAll(".step-pill").forEach((btn) => {
      btn.addEventListener("click", () => switchStep(btn.dataset.step));
    });
    document.getElementById("btn-goto-organize").addEventListener("click", () => switchStep("organize"));
    document.getElementById("btn-goto-generate").addEventListener("click", () => switchStep("generate"));
    document.getElementById("btn-back-to-select").addEventListener("click", () => switchStep("select"));

    // Config toolbar & banner
    document.getElementById("btn-save-config").addEventListener("click", saveCurrentConfig);
    document.getElementById("btn-load-config").addEventListener("click", openConfigViaDialog);
    document.getElementById("btn-recent-menu").addEventListener("click", (e) => {
      e.stopPropagation();
      toggleRecentDropdown();
    });
    document.getElementById("recent-menu").addEventListener("click", (e) => {
      e.stopPropagation();
    });
    document.addEventListener("click", () => {
      closeRecentMenu();
    });
    document.getElementById("btn-recheck-files").addEventListener("click", () => validateCurrentPlots(false));
    document.getElementById("btn-gen-recheck").addEventListener("click", () => validateCurrentPlots(false));
    document.getElementById("btn-remap-source").addEventListener("click", remapSourceFolderViaDialog);
    document.getElementById("btn-clear-config").addEventListener("click", () => {
      state.configPath = null;
      state.configName = "";
      updateConfigBanner();
    });

    // Step 1 controls
    document.getElementById("btn-browse-folder").addEventListener("click", async () => {
      try {
        const cur = document.getElementById("folder-path-input").value.trim();
        let chosen = "";
        if (!state.hasDialogs) {
          chosen = await promptPathModal({
            title: "Select Plots Folder",
            label: "Plots folder or UNC path",
            defaultValue: cur,
            placeholder: "e.g. C:\\Projects\\Plots or \\\\Server\\Share\\Plots",
            submitLabel: "Open Folder",
          });
          if (!chosen) return;
        } else {
          const res = await apiPost("/api/dialog/browse-folder", { initial_dir: cur });
          if (res.cancelled || !res.path) return;
          chosen = res.path;
        }
        document.getElementById("folder-path-input").value = chosen;
        await openFolderLazy(chosen);
      } catch (err) {
        showToast(err.message, "error");
      }
    });
    document.getElementById("btn-open-folder").addEventListener("click", () => {
      openFolderLazy(document.getElementById("folder-path-input").value);
    });
    document.getElementById("btn-tree-retry").addEventListener("click", () => {
      openFolderLazy(state.lastFailedFolder || document.getElementById("folder-path-input").value);
    });
    document.getElementById("btn-deep-scan").addEventListener("click", () => {
      startDeepScan(document.getElementById("folder-path-input").value);
    });
    document.getElementById("folder-path-input").addEventListener("keydown", (e) => {
      if (e.key === "Enter") openFolderLazy(e.target.value);
    });
    document.getElementById("select-filter-input").addEventListener("input", (e) => {
      state.selectFilter = e.target.value;
      renderStep1Tree();
    });
    document.getElementById("btn-select-filtered").addEventListener("click", () => {
      for (const f of getFilteredLoadedFiles()) state.selectedSet.add(f.path);
      updateHeaderBadges();
      renderStep1Tree();
    });
    document.getElementById("btn-select-all-loaded").addEventListener("click", () => {
      for (const f of getLoadedFilesList()) state.selectedSet.add(f.path);
      updateHeaderBadges();
      renderStep1Tree();
    });
    document.getElementById("btn-deselect-all-loaded").addEventListener("click", () => {
      state.selectedSet.clear();
      updateHeaderBadges();
      renderStep1Tree();
    });

    // Scan banner actions
    document.getElementById("btn-scan-continue").addEventListener("click", async () => {
      if (!state.scanId) return;
      try {
        document.getElementById("btn-scan-continue").classList.add("hidden");
        document.getElementById("btn-scan-keep").classList.add("hidden");
        document.getElementById("scan-spinner").classList.remove("paused");
        document.getElementById("scan-status-title").textContent = "Continuing full scan...";
        await apiPost("/api/scan/continue", { scan_id: state.scanId, new_limit: 0 });
        pollScanStatus();
      } catch (err) {
        showToast(err.message, "error");
      }
    });
    document.getElementById("btn-scan-keep").addEventListener("click", async () => {
      if (!state.scanId) return;
      try {
        await apiPost("/api/scan/continue", { scan_id: state.scanId, keep_loaded: true });
        pollScanStatus();
      } catch (err) {
        showToast(err.message, "error");
      }
    });
    document.getElementById("btn-scan-cancel").addEventListener("click", async () => {
      if (!state.scanId) return;
      stopScanPolling();
      const scanId = state.scanId;
      state.scanId = null;
      try {
        await apiPost("/api/scan/cancel", { scan_id: scanId });
        document.getElementById("scan-banner").classList.add("hidden");
        showToast("Scan cancelled.", "warn");
        renderStep1Tree();
      } catch (err) {
        document.getElementById("scan-banner").classList.add("hidden");
        showToast(err.message, "error");
      }
    });

    // Step 2 controls
    document.querySelectorAll("#grouping-segmented .seg-btn").forEach((btn) => {
      btn.addEventListener("click", () => {
        const nextMode = btn.dataset.mode;
        if (state.groupingMode === nextMode) return;
        state.groupingMode = nextMode;
        if (nextMode === "manual" && state.sections.length === 1 && state.sections[0].id === "sec_flat") {
          state.sections[0].id = `sec_${state.secCounter++}`;
          state.sections[0].name = "General Plots";
          state.sections[0].page_break = true;
        }
        reconcileSectionsWithSelection();
        renderSectionsFull();
      });
    });

    document.getElementById("btn-add-section").addEventListener("click", addSectionModal);
    document.getElementById("chk-show-thumbs").addEventListener("change", (e) => {
      state.showThumbs = e.target.checked;
      const container = document.getElementById("sections-container");
      container.classList.toggle("show-thumbs", state.showThumbs);
      setupThumbObserver();
    });
    document.getElementById("org-search-input").addEventListener("input", (e) => {
      state.orgSearch = e.target.value;
      applyOrgSearchHighlight();
    });
    document.getElementById("btn-collapse-all").addEventListener("click", () => {
      walkSections(state.sections, (sec) => { sec.collapsed = true; });
      document.querySelectorAll("#sections-container .section-card").forEach((card) => {
        card.classList.add("collapsed");
        const btn = card.querySelector(":scope > .section-header .sec-collapse-btn");
        if (btn) {
          setIconHtml(btn, "chevronRight");
          btn.setAttribute("aria-expanded", "false");
        }
      });
    });
    document.getElementById("btn-expand-all").addEventListener("click", () => {
      walkSections(state.sections, (sec) => { sec.collapsed = false; });
      document.querySelectorAll("#sections-container .section-card").forEach((card) => {
        card.classList.remove("collapsed");
        const btn = card.querySelector(":scope > .section-header .sec-collapse-btn");
        if (btn) {
          setIconHtml(btn, "chevronDown");
          btn.setAttribute("aria-expanded", "true");
        }
      });
    });

    document.getElementById("btn-org-select-all").addEventListener("click", () => setAllOrgChecked(true));
    document.getElementById("btn-org-deselect-all").addEventListener("click", () => setAllOrgChecked(false));
    document.getElementById("btn-bulk-up").addEventListener("click", () => bulkReorderInPlace(-1));
    document.getElementById("btn-bulk-down").addEventListener("click", () => bulkReorderInPlace(1));
    document.getElementById("btn-bulk-move").addEventListener("click", () => {
      const targetId = document.getElementById("bulk-target-section").value;
      if (targetId) bulkMoveCheckedToSection(targetId);
    });
    document.getElementById("btn-bulk-remove").addEventListener("click", bulkRemoveCheckedInPlace);

    // Step 3 controls
    document.querySelectorAll("#orient-segmented .seg-btn").forEach((btn) => {
      btn.addEventListener("click", () => {
        state.docOptions.orientation = btn.dataset.value;
        applyDocOptionsToInputs();
      });
    });
    document.querySelectorAll("#cols-segmented .seg-btn").forEach((btn) => {
      btn.addEventListener("click", () => {
        state.docOptions.num_cols = Number(btn.dataset.value);
        applyDocOptionsToInputs();
      });
    });
    document.getElementById("opt-max-width").addEventListener("change", syncDocOptionsFromInputs);
    document.getElementById("opt-max-height").addEventListener("change", syncDocOptionsFromInputs);
    document.getElementById("opt-export-docx").addEventListener("change", () => {
      syncDocOptionsFromInputs();
      applyDocOptionsToInputs();
    });
    document.getElementById("opt-export-html").addEventListener("change", () => {
      syncDocOptionsFromInputs();
      applyDocOptionsToInputs();
    });
    document.getElementById("opt-html-embed").addEventListener("change", syncDocOptionsFromInputs);
    document.getElementById("opt-export-pdf").addEventListener("change", syncDocOptionsFromInputs);
    document.getElementById("btn-generate-as").addEventListener("click", () => runGeneration(null));
    document.getElementById("btn-quick-regen").addEventListener("click", () => runGeneration(state.docOptions.last_output_path));
  }

  async function init() {
    initTheme();
    wireEvents();
    initHeartbeat();
    try {
      const health = await apiGet("/api/health");
      state.hasWin32Com = Boolean(health.has_win32com);
      state.hasDialogs = Boolean(health.has_dialogs);
      if (!state.hasWin32Com) {
        document.getElementById("win32-warning").classList.remove("hidden");
      }
    } catch (_) {}
    updateHeaderBadges();
  }

  document.addEventListener("DOMContentLoaded", init);
})();
