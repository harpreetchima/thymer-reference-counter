class Plugin extends AppPlugin {
  onLoad() {
    this._version = '0.1.0';
    this._pluginName = 'Reference Counter';

    this._panelStates = new Map();
    this._eventHandlerIds = [];
    this._commandHandles = [];

    this._countCache = new Map();
    this._recordExistsCache = new Map();
    this._recordNameCache = new Map();

    this._storageKeyEnabled = 'thymer_reference_counter_enabled_v1';
    this._storageKeyHoverOnly = 'thymer_reference_counter_hover_only_v1';

    const cfg = this.getConfiguration?.() || {};
    const custom = cfg.custom || {};

    this._defaultEnabled = custom.enabledByDefault !== false;
    this._defaultHoverOnly = custom.hoverOnlyByDefault === true;

    this._enabled = this.loadBoolSetting(this._storageKeyEnabled, this._defaultEnabled);
    this._hoverOnly = this.loadBoolSetting(this._storageKeyHoverOnly, this._defaultHoverOnly);

    this._countMode = custom.countMode === 'records' ? 'records' : 'lines';
    this._minCount = this.coercePositiveInt(custom.minCount, 1);
    this._showZero = custom.showZero === true;
    this._showSelf = custom.showSelf === true;
    this._maxResults = this.coercePositiveInt(custom.maxResults, 250);
    this._cacheTtlMs = this.coercePositiveInt(custom.cacheTtlMs, 120000);
    this._refreshDebounceMs = this.coercePositiveInt(custom.refreshDebounceMs, 180);
    this._opacity = this.coerceOpacity(custom.opacity, 0.55);
    this._fontScaleClass = this.normalizeFontScale(custom.fontScale);

    this.injectCss();
    this.registerCommands();
    this.registerEventHandlers();

    const seen = new Set();
    const panels = this.ui.getPanels?.() || [];
    for (const panel of panels) {
      const panelId = panel?.getId?.() || null;
      if (!panelId || seen.has(panelId)) continue;
      seen.add(panelId);
      this.handlePanelChanged(panel, 'initial-panels');
    }

    const active = this.ui.getActivePanel?.() || null;
    if (active) this.handlePanelChanged(active, 'initial-active');

    setTimeout(() => {
      this.refreshAllPanels({ force: true, reason: 'initial-delayed' });
    }, 350);
  }

  onUnload() {
    for (const id of this._eventHandlerIds || []) {
      try {
        this.events.off(id);
      } catch (e) {
        // ignore
      }
    }
    this._eventHandlerIds = [];

    for (const handle of this._commandHandles || []) {
      try {
        handle?.remove?.();
      } catch (e) {
        // ignore
      }
    }
    this._commandHandles = [];

    for (const panelId of Array.from(this._panelStates.keys())) {
      this.disposePanelState(panelId);
    }

    this._panelStates.clear();
    this._countCache.clear();
    this._recordExistsCache.clear();
    this._recordNameCache.clear();
  }

  registerCommands() {
    this._commandHandles.push(
      this.ui.addCommandPaletteCommand({
        label: 'Reference Counter: Toggle inline counters',
        icon: 'hash',
        onSelected: () => this.toggleEnabled()
      })
    );

    this._commandHandles.push(
      this.ui.addCommandPaletteCommand({
        label: 'Reference Counter: Toggle hover-only counters',
        icon: 'eye',
        onSelected: () => this.toggleHoverOnly()
      })
    );

    this._commandHandles.push(
      this.ui.addCommandPaletteCommand({
        label: 'Reference Counter: Refresh active page',
        icon: 'refresh',
        onSelected: () => {
          const panel = this.ui.getActivePanel?.() || null;
          if (panel) this.handlePanelChanged(panel, 'cmd-refresh-active');
          this.refreshAllPanels({ force: true, reason: 'cmd-refresh-all' });
        }
      })
    );

    this._commandHandles.push(
      this.ui.addCommandPaletteCommand({
        label: 'Reference Counter: Clear count cache',
        icon: 'trash',
        onSelected: () => {
          this.clearCountCache();
          this.refreshAllPanels({ force: true, reason: 'cmd-clear-cache' });
          this.showToast('Reference Counter', 'Local count cache cleared.');
        }
      })
    );
  }

  registerEventHandlers() {
    this._eventHandlerIds.push(
      this.events.on('panel.navigated', (ev) => this.handlePanelChanged(ev?.panel || null, 'panel.navigated'))
    );
    this._eventHandlerIds.push(
      this.events.on('panel.focused', (ev) => this.handlePanelChanged(ev?.panel || null, 'panel.focused'))
    );
    this._eventHandlerIds.push(
      this.events.on('panel.closed', (ev) => this.handlePanelClosed(ev?.panel || null))
    );

    this._eventHandlerIds.push(
      this.events.on('lineitem.updated', (ev) => this.handleLineItemUpdated(ev))
    );
    this._eventHandlerIds.push(
      this.events.on('lineitem.deleted', () => this.handleReferenceMutation('lineitem.deleted'))
    );
    this._eventHandlerIds.push(
      this.events.on('record.updated', () => this.handleReferenceMutation('record.updated'))
    );

    this._eventHandlerIds.push(
      this.events.on('reload', () => {
        this.clearCountCache();
        this.refreshAllPanels({ force: true, reason: 'reload' });
      })
    );
  }

  toggleEnabled() {
    this._enabled = !this._enabled;
    this.saveBoolSetting(this._storageKeyEnabled, this._enabled);

    if (!this._enabled) {
      for (const state of this._panelStates.values()) {
        const panelEl = state?.panel?.getElement?.() || state?.panelEl || null;
        if (panelEl) this.clearBadgesInElement(panelEl);
      }
      this.showToast('Reference Counter', 'Inline counters disabled.');
      return;
    }

    this.refreshAllPanels({ force: true, reason: 'toggle-enabled' });
    this.showToast('Reference Counter', 'Inline counters enabled.');
  }

  toggleHoverOnly() {
    this._hoverOnly = !this._hoverOnly;
    this.saveBoolSetting(this._storageKeyHoverOnly, this._hoverOnly);
    this.refreshAllPanels({ force: true, reason: 'toggle-hover-only' });
    this.showToast(
      'Reference Counter',
      this._hoverOnly ? 'Hover-only counters enabled.' : 'Hover-only counters disabled.'
    );
  }

  showToast(title, message) {
    try {
      this.ui.addToaster({
        title,
        message,
        dismissible: true,
        autoDestroyTime: 1800
      });
    } catch (e) {
      // ignore
    }
  }

  loadBoolSetting(key, fallback) {
    try {
      const v = localStorage.getItem(key);
      if (v === '1') return true;
      if (v === '0') return false;
    } catch (e) {
      // ignore
    }
    return fallback === true;
  }

  saveBoolSetting(key, value) {
    try {
      localStorage.setItem(key, value ? '1' : '0');
    } catch (e) {
      // ignore
    }
  }

  coercePositiveInt(value, fallback) {
    const n = Number(value);
    if (!Number.isFinite(n) || n <= 0) return fallback;
    return Math.floor(n);
  }

  coerceOpacity(value, fallback) {
    const n = Number(value);
    if (!Number.isFinite(n)) return fallback;
    if (n < 0.1) return 0.1;
    if (n > 1) return 1;
    return n;
  }

  normalizeFontScale(value) {
    const raw = typeof value === 'string' ? value.trim().toLowerCase() : '';
    if (raw === 'xsmall') return 'trc-size-xsmall';
    if (raw === 'small') return 'trc-size-small';
    if (raw === 'large') return 'trc-size-large';
    return 'trc-size-medium';
  }

  injectCss() {
    this.ui.injectCSS(`
      .trc-refcount-badge-wrap {
        display: inline-flex;
        align-items: flex-start;
        vertical-align: super;
        line-height: 1;
        margin-left: 2px;
        user-select: none;
      }

      .trc-refcount-badge {
        border: 0;
        background: transparent;
        color: var(--text-muted, var(--text-default, var(--text, inherit)));
        font-weight: 650;
        line-height: 1;
        padding: 0 1px;
        border-radius: var(--radius-normal, 4px);
        opacity: ${this._opacity};
        cursor: pointer;
      }

      .trc-refcount-badge:hover {
        opacity: 1;
        color: var(--ed-link-color, var(--link-color, var(--accent, inherit)));
        text-decoration: underline;
      }

      .trc-refcount-badge-wrap.trc-size-xsmall .trc-refcount-badge { font-size: 9px; }
      .trc-refcount-badge-wrap.trc-size-small .trc-refcount-badge { font-size: 10px; }
      .trc-refcount-badge-wrap.trc-size-medium .trc-refcount-badge { font-size: 11px; }
      .trc-refcount-badge-wrap.trc-size-large .trc-refcount-badge { font-size: 12px; }

      .trc-refcount-badge-wrap.trc-hover-only .trc-refcount-badge {
        opacity: 0;
        pointer-events: none;
      }

      .line-div:hover .trc-refcount-badge-wrap.trc-hover-only .trc-refcount-badge,
      .line-check-div:hover .trc-refcount-badge-wrap.trc-hover-only .trc-refcount-badge,
      .trc-refcount-badge-wrap.trc-hover-only:hover .trc-refcount-badge {
        opacity: ${this._opacity};
        pointer-events: auto;
      }
    `);
  }

  handlePanelChanged(panel, reason) {
    const panelId = panel?.getId?.() || null;
    if (!panelId) return;

    const panelEl = panel?.getElement?.() || null;
    if (!panelEl) {
      this.disposePanelState(panelId);
      return;
    }

    if (this.shouldSuppressInPanel(panel, panelEl)) {
      this.disposePanelState(panelId);
      return;
    }

    const record = panel?.getActiveRecord?.() || null;
    const recordGuid = record?.guid || null;
    if (!recordGuid) {
      this.disposePanelState(panelId);
      return;
    }

    const state = this.getOrCreatePanelState(panel);
    state.panel = panel;
    state.panelEl = panelEl;
    state.recordGuid = recordGuid;

    this.attachObserver(state, panelEl);
    this.scheduleScan(state, { force: reason !== 'panel.focused', reason: reason || 'panel-changed' });
  }

  shouldSuppressInPanel(panel, panelEl) {
    const navType = panel?.getNavigation?.()?.type || '';
    if (navType === 'custom' || navType === 'custom_panel') return true;

    const root = this.findEditorRoot(panelEl);
    return !root;
  }

  getOrCreatePanelState(panel) {
    const panelId = panel?.getId?.() || 'unknown';
    let state = this._panelStates.get(panelId) || null;
    if (state) return state;

    state = {
      panelId,
      panel,
      panelEl: panel?.getElement?.() || null,
      recordGuid: null,
      observer: null,
      observerEl: null,
      scanTimer: null,
      scanSeq: 0,
      ignoreMutationsUntil: 0
    };

    this._panelStates.set(panelId, state);
    return state;
  }

  attachObserver(state, panelEl) {
    if (!state || !panelEl) return;
    if (state.observer && state.observerEl === panelEl) return;

    if (state.observer) {
      try {
        state.observer.disconnect();
      } catch (e) {
        // ignore
      }
      state.observer = null;
      state.observerEl = null;
    }

    const observer = new MutationObserver(() => {
      if (state.ignoreMutationsUntil > Date.now()) return;
      this.scheduleScan(state, { force: false, reason: 'dom-mutation' });
    });

    observer.observe(panelEl, { childList: true, subtree: true });
    state.observer = observer;
    state.observerEl = panelEl;
  }

  handlePanelClosed(panel) {
    const panelId = panel?.getId?.() || null;
    if (!panelId) return;
    this.disposePanelState(panelId);
  }

  disposePanelState(panelId) {
    const state = this._panelStates.get(panelId) || null;
    if (!state) return;

    if (state.scanTimer) {
      clearTimeout(state.scanTimer);
      state.scanTimer = null;
    }

    if (state.observer) {
      try {
        state.observer.disconnect();
      } catch (e) {
        // ignore
      }
      state.observer = null;
      state.observerEl = null;
    }

    const panelEl = state?.panel?.getElement?.() || state.panelEl || null;
    if (panelEl) this.clearBadgesInElement(panelEl);

    this._panelStates.delete(panelId);
  }

  scheduleScan(state, { force, reason }) {
    if (!state) return;

    if (state.scanTimer) {
      clearTimeout(state.scanTimer);
      state.scanTimer = null;
    }

    const delay = force ? 0 : this._refreshDebounceMs;
    state.scanTimer = setTimeout(() => {
      state.scanTimer = null;
      this.scanPanel(state.panelId, reason || 'scheduled').catch(() => {
        // ignore
      });
    }, delay);
  }

  refreshAllPanels({ force, reason }) {
    for (const state of this._panelStates.values()) {
      this.scheduleScan(state, { force: force === true, reason: reason || 'all-panels' });
    }
  }

  async scanPanel(panelId, reason) {
    const state = this._panelStates.get(panelId) || null;
    if (!state) return;

    const panel = state.panel || null;
    const panelEl = panel?.getElement?.() || state.panelEl || null;
    if (!panel || !panelEl || !panelEl.isConnected) return;

    const editorRoot = this.findEditorRoot(panelEl);
    if (!editorRoot) {
      this.clearBadgesInElement(panelEl);
      return;
    }

    if (!this._enabled) {
      this.clearBadgesInElement(editorRoot);
      return;
    }

    const seq = (state.scanSeq || 0) + 1;
    state.scanSeq = seq;

    state.ignoreMutationsUntil = Date.now() + 80;
    this.clearBadgesInElement(editorRoot);

    const refs = this.collectReferenceTargets(editorRoot);
    if (refs.length === 0) return;

    const uniqueGuids = Array.from(new Set(refs.map((x) => x.guid)));
    const counts = new Map();

    await Promise.all(
      uniqueGuids.map(async (guid) => {
        const info = await this.getCountInfoForGuid(guid);
        counts.set(guid, info);
      })
    );

    if (!this._panelStates.has(panelId)) return;
    if (state.scanSeq !== seq) return;

    state.ignoreMutationsUntil = Date.now() + 120;
    for (const ref of refs) {
      const info = counts.get(ref.guid) || { count: 0, capped: false };
      if (!this._showZero && info.count <= 0) continue;
      if (info.count < this._minCount) continue;
      this.insertBadge(ref.el, ref.guid, info, panel);
    }
  }

  findEditorRoot(panelEl) {
    if (!panelEl) return null;
    const checks = ['.page-content', '.editor-wrapper', '.editor-panel', '#editor'];
    for (const selector of checks) {
      if (panelEl.matches?.(selector)) return panelEl;
      const child = panelEl.querySelector?.(selector) || null;
      if (child) return child;
    }
    return null;
  }

  clearBadgesInElement(rootEl) {
    if (!rootEl?.querySelectorAll) return;
    const badges = rootEl.querySelectorAll('.trc-refcount-badge-wrap');
    for (const el of badges) {
      try {
        el.remove();
      } catch (e) {
        // ignore
      }
    }
  }

  collectReferenceTargets(editorRoot) {
    const selectors = [
      '.line-div [data-ref-guid]',
      '.line-div [data-record-guid]',
      '.line-div [data-link-guid]',
      '.line-div [data-guid]',
      '.line-div [data-root-id]',
      '.line-check-div [data-ref-guid]',
      '.line-check-div [data-record-guid]',
      '.line-check-div [data-link-guid]',
      '.line-check-div [data-guid]',
      '.line-check-div [data-root-id]',
      '.line-div .line-ref',
      '.line-div .segment-ref',
      '.line-div .record-link',
      '.line-div [class*="line-ref"]',
      '.line-div [class*="segment-ref"]',
      '.line-div [class*="record-link"]',
      '.line-div [class*="ref-open"]'
    ].join(', ');

    const nodes = editorRoot.querySelectorAll(selectors);
    const out = [];
    const seen = new Set();

    for (const el of nodes) {
      if (seen.has(el)) continue;
      seen.add(el);

      if (!this.isLikelyReferenceElement(el)) continue;
      const guid = this.extractRecordGuidFromElement(el);
      if (!guid) continue;

      out.push({ el, guid });
    }

    return out;
  }

  isLikelyReferenceElement(el) {
    if (!el || el.nodeType !== 1) return false;
    if (el.closest('.trc-refcount-badge-wrap')) return false;
    if (!el.closest('.line-div, .line-check-div')) return false;

    const cls = (el.className || '').toString().toLowerCase();
    if (!cls && !el.hasAttribute('data-guid') && !el.hasAttribute('data-record-guid') && !el.hasAttribute('data-ref-guid')) {
      return false;
    }

    if (
      cls.includes('trc-refcount') ||
      cls.includes('listitem-indentline') ||
      el.classList.contains('line-div') ||
      el.classList.contains('line-check-div') ||
      el.classList.contains('listitem')
    ) {
      return false;
    }

    const hasHint =
      cls.includes('ref') ||
      cls.includes('link') ||
      el.hasAttribute('data-ref-guid') ||
      el.hasAttribute('data-record-guid') ||
      el.hasAttribute('data-link-guid') ||
      (el.tagName === 'A' && (el.hasAttribute('data-guid') || el.hasAttribute('data-root-id')));

    if (!hasHint) return false;

    const text = (el.textContent || '').trim();
    if (text.length > 240) return false;
    if (el.childElementCount > 6) return false;

    return true;
  }

  extractRecordGuidFromElement(el) {
    const candidates = this.extractGuidCandidates(el);
    for (const raw of candidates) {
      const guid = typeof raw === 'string' ? raw.trim() : '';
      if (!this.looksLikeGuid(guid)) continue;
      if (!this.isExistingRecordGuid(guid)) continue;
      return guid;
    }
    return null;
  }

  extractGuidCandidates(el) {
    const out = [];
    const seen = new Set();
    const dataKeys = ['refGuid', 'recordGuid', 'linkGuid', 'guid', 'rootId', 'linkUid', 'uid'];
    const attrKeys = [
      'data-ref-guid',
      'data-record-guid',
      'data-link-guid',
      'data-guid',
      'data-root-id',
      'data-link-uid',
      'data-uid'
    ];

    const push = (value) => {
      if (typeof value !== 'string') return;
      const v = value.trim();
      if (!v || seen.has(v)) return;
      seen.add(v);
      out.push(v);
    };

    let node = el;
    for (let depth = 0; node && depth < 4; depth += 1) {
      for (const key of dataKeys) {
        try {
          push(node.dataset?.[key]);
        } catch (e) {
          // ignore
        }
      }

      for (const attr of attrKeys) {
        try {
          push(node.getAttribute?.(attr));
        } catch (e) {
          // ignore
        }
      }

      if (node.tagName === 'A') {
        const href = node.getAttribute?.('href') || '';
        const m = href.match(/([A-Za-z0-9_-]{12,})/g) || [];
        for (const value of m) push(value);
      }

      node = node.parentElement;
    }

    return out;
  }

  looksLikeGuid(value) {
    if (typeof value !== 'string') return false;
    const v = value.trim();
    if (v.length < 12 || v.length > 64) return false;
    if (!/^[A-Za-z0-9_-]+$/.test(v)) return false;
    return /[A-Z0-9]/.test(v);
  }

  isExistingRecordGuid(guid) {
    if (this._recordExistsCache.has(guid)) {
      return this._recordExistsCache.get(guid) === true;
    }

    let exists = false;
    try {
      exists = !!this.data.getRecord?.(guid);
    } catch (e) {
      exists = false;
    }

    this._recordExistsCache.set(guid, exists);
    return exists;
  }

  getOrLoadRecordName(guid) {
    if (this._recordNameCache.has(guid)) {
      return this._recordNameCache.get(guid) || guid;
    }

    let name = guid;
    try {
      const rec = this.data.getRecord?.(guid) || null;
      const n = rec?.getName?.() || '';
      if (typeof n === 'string' && n.trim()) name = n.trim();
    } catch (e) {
      // ignore
    }

    this._recordNameCache.set(guid, name);
    return name;
  }

  insertBadge(targetEl, guid, info, panel) {
    if (!targetEl?.isConnected) return;
    if (targetEl.closest('.trc-refcount-badge-wrap')) return;

    const existing = targetEl.nextElementSibling;
    if (existing?.classList?.contains('trc-refcount-badge-wrap')) {
      try {
        existing.remove();
      } catch (e) {
        // ignore
      }
    }

    const wrap = document.createElement('span');
    wrap.className = 'trc-refcount-badge-wrap';
    wrap.classList.add(this._fontScaleClass);
    if (this._hoverOnly) wrap.classList.add('trc-hover-only');
    wrap.dataset.guid = guid;

    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'trc-refcount-badge button-none button-small button-minimal-hover text-details tooltip';
    btn.textContent = this.formatCountLabel(info);

    const recordName = this.getOrLoadRecordName(guid);
    const tooltip = `${info.count}${info.capped ? '+' : ''} references to ${recordName}`;
    btn.title = `${tooltip} (Ctrl/Cmd-click to open in a new panel)`;
    btn.dataset.tooltip = tooltip;
    btn.dataset.tooltipDir = 'top';
    btn.setAttribute('aria-label', tooltip);

    btn.addEventListener('click', (ev) => {
      ev.preventDefault();
      ev.stopPropagation();
      if (typeof ev.stopImmediatePropagation === 'function') ev.stopImmediatePropagation();
      void this.openTargetRecord(panel, guid, ev);
    });

    wrap.appendChild(btn);

    try {
      targetEl.insertAdjacentElement('afterend', wrap);
    } catch (e) {
      // ignore
    }
  }

  formatCountLabel(info) {
    const count = Number(info?.count) || 0;
    if (count >= 1000) {
      const compact = Math.round((count / 1000) * 10) / 10;
      return `${compact}k${info?.capped ? '+' : ''}`;
    }
    return `${count}${info?.capped ? '+' : ''}`;
  }

  async openTargetRecord(panel, guid, ev) {
    if (!panel) return;

    const workspaceGuid = this.getWorkspaceGuid?.() || null;
    if (!workspaceGuid) return;

    const openInNew = ev?.metaKey === true || ev?.ctrlKey === true;
    if (openInNew) {
      try {
        const newPanel = await this.ui.createPanel({ afterPanel: panel });
        if (!newPanel) return;
        newPanel.navigateTo({
          type: 'edit_panel',
          rootId: guid,
          subId: null,
          workspaceGuid
        });
        this.ui.setActivePanel(newPanel);
      } catch (e) {
        // ignore
      }
      return;
    }

    panel.navigateTo({
      type: 'edit_panel',
      rootId: guid,
      subId: null,
      workspaceGuid
    });
    this.ui.setActivePanel(panel);
  }

  async getCountInfoForGuid(guid) {
    const cached = this.getCachedCountInfo(guid);
    if (cached) return cached;

    const pending = this._countCache.get(guid)?.promise || null;
    if (pending) return pending;

    const promise = this.loadCountInfo(guid)
      .then((info) => {
        this.setCachedCountInfo(guid, info);
        return info;
      })
      .catch(() => {
        const fallback = { count: 0, capped: false };
        this.setCachedCountInfo(guid, fallback);
        return fallback;
      })
      .finally(() => {
        const current = this._countCache.get(guid);
        if (current && current.promise) {
          delete current.promise;
        }
      });

    this._countCache.set(guid, {
      count: 0,
      capped: false,
      updatedAt: Date.now(),
      promise
    });

    return promise;
  }

  getCachedCountInfo(guid) {
    const entry = this._countCache.get(guid) || null;
    if (!entry || typeof entry.count !== 'number') return null;
    if ((Date.now() - (entry.updatedAt || 0)) > this._cacheTtlMs) return null;
    return { count: entry.count, capped: entry.capped === true };
  }

  setCachedCountInfo(guid, info) {
    this._countCache.set(guid, {
      count: Number(info?.count) || 0,
      capped: info?.capped === true,
      updatedAt: Date.now()
    });
  }

  async loadCountInfo(guid) {
    if (!this.isExistingRecordGuid(guid)) return { count: 0, capped: false };

    if (this._countMode === 'records') {
      const record = this.data.getRecord?.(guid) || null;
      if (!record?.getBackReferenceRecords) return { count: 0, capped: false };

      const refs = await record.getBackReferenceRecords();
      let records = Array.isArray(refs) ? refs : [];
      if (!this._showSelf) records = records.filter((x) => x?.guid !== guid);
      return { count: records.length, capped: false };
    }

    const query = `@linkto = "${guid}"`;
    const result = await this.data.searchByQuery(query, this._maxResults);
    if (result?.error) return { count: 0, capped: false };

    const lines = Array.isArray(result?.lines) ? result.lines : [];
    const unique = new Set();
    let fallbackIdx = 0;

    for (const line of lines) {
      if (!line) continue;
      const sourceRecordGuid = line?.record?.guid || '';
      if (!this._showSelf && sourceRecordGuid === guid) continue;

      const key = line.guid || `${sourceRecordGuid}:${fallbackIdx++}`;
      unique.add(key);
    }

    return {
      count: unique.size,
      capped: lines.length >= this._maxResults
    };
  }

  clearCountCache() {
    this._countCache.clear();
    this._recordExistsCache.clear();
    this._recordNameCache.clear();
  }

  handleReferenceMutation(reason) {
    this.clearCountCache();
    this.refreshAllPanels({ force: false, reason: reason || 'reference-mutation' });
  }

  handleLineItemUpdated(ev) {
    let touchedAny = false;

    if (ev?.hasSegments?.() && typeof ev.getSegments === 'function') {
      const segments = ev.getSegments() || [];
      for (const seg of segments) {
        if (seg?.type !== 'ref') continue;

        let guid = '';
        if (typeof seg.text === 'string') guid = seg.text;
        else if (seg.text && typeof seg.text.guid === 'string') guid = seg.text.guid;
        guid = (guid || '').trim();

        if (!guid) continue;
        this._countCache.delete(guid);
        touchedAny = true;
      }
    }

    if (!touchedAny) {
      this.clearCountCache();
    }

    this.refreshAllPanels({ force: false, reason: 'lineitem.updated' });
  }
}
