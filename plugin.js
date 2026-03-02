class Plugin extends AppPlugin {
  onLoad() {
    this._version = '0.1.0';
    this._pluginName = 'Reference Counter';
    this._isUnloading = false;
    this._initialRefreshTimer = null;
    this.ensureRuntimeState();

    this._panelStates = new Map();
    this._eventHandlerIds = [];
    this._commandHandles = [];

    this._countCache = new Map();
    this._recordExistsCache = new Map();
    this._recordNameCache = new Map();
    this._lineRefGuids = new Map();
    this._sharedIgnoreMetaKey = 'plugin.refs.v1.ignore';

    this._storageKeyEnabled = 'thymer_reference_counter_enabled_v1';
    this._storageKeyHoverOnly = 'thymer_reference_counter_hover_only_v1';

    const cfg = this.getConfiguration?.() || {};
    const custom = cfg.custom || {};

    this._defaultEnabled = custom.enabledByDefault !== false;
    this._defaultHoverOnly = custom.hoverOnlyByDefault === true;

    this._enabled = this.loadBoolSetting(this._storageKeyEnabled, this._defaultEnabled);
    this._hoverOnly = this.loadBoolSetting(this._storageKeyHoverOnly, this._defaultHoverOnly);

    this._countMode = this.normalizeCountMode(custom.countMode);
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

    this._initialRefreshTimer = setTimeout(() => {
      this._initialRefreshTimer = null;
      if (this._isUnloading) return;
      this.refreshAllPanels({ force: true, reason: 'initial-delayed' });
    }, 350);
  }

  onUnload() {
    this._isUnloading = true;
    this.ensureRuntimeState();

    if (this._initialRefreshTimer) {
      clearTimeout(this._initialRefreshTimer);
      this._initialRefreshTimer = null;
    }

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

    for (const panelId of Array.from(this._panelStates?.keys() || [])) {
      this.disposePanelState(panelId);
    }

    this._panelStates?.clear();
    this._countCache?.clear();
    this._recordExistsCache?.clear();
    this._recordNameCache?.clear();
    this._lineRefGuids?.clear();
  }

  ensureRuntimeState() {
    if (!(this._panelStates instanceof Map)) this._panelStates = new Map();
    if (!Array.isArray(this._eventHandlerIds)) this._eventHandlerIds = [];
    if (!Array.isArray(this._commandHandles)) this._commandHandles = [];
    if (!(this._countCache instanceof Map)) this._countCache = new Map();
    if (!(this._recordExistsCache instanceof Map)) this._recordExistsCache = new Map();
    if (!(this._recordNameCache instanceof Map)) this._recordNameCache = new Map();
    if (!(this._lineRefGuids instanceof Map)) this._lineRefGuids = new Map();
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
      this.events.on('lineitem.deleted', (ev) => this.handleLineItemDeleted(ev))
    );
    this._eventHandlerIds.push(
      this.events.on('record.updated', (ev) => this.handleRecordUpdated(ev))
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
      for (const state of this._panelStates?.values() || []) {
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

  normalizeCountMode(value) {
    const raw = typeof value === 'string' ? value.trim().toLowerCase() : '';
    if (raw === 'lines') return 'lines';
    if (raw === 'records') return 'records';
    if (raw === 'combined') return 'combined';
    return 'combined';
  }

  didSharedIgnoreChange(ev) {
    const mp = ev?.metaProperties;
    if (!mp || typeof mp !== 'object') return false;

    if (Object.prototype.hasOwnProperty.call(mp, this._sharedIgnoreMetaKey)) return true;

    const nested = mp?.plugin?.refs?.v1;
    if (nested && Object.prototype.hasOwnProperty.call(nested, 'ignore')) return true;

    return false;
  }

  normalizeSharedIgnoreValue(value) {
    if (value === true || value === 1) return true;
    if (typeof value === 'string') {
      const v = value.trim().toLowerCase();
      if (v === '1' || v === 'true' || v === 'yes' || v === 'on') return true;
    }
    return false;
  }

  readSharedIgnoreFromProps(props) {
    if (!props || typeof props !== 'object') return false;

    const direct = props?.[this._sharedIgnoreMetaKey];
    if (this.normalizeSharedIgnoreValue(direct)) return true;

    const underscore = props?.plugin_refs_v1_ignore;
    if (this.normalizeSharedIgnoreValue(underscore)) return true;

    const nested = props?.plugin?.refs?.v1?.ignore;
    if (this.normalizeSharedIgnoreValue(nested)) return true;

    return false;
  }

  isLineSharedIgnored(line) {
    if (!line) return false;
    return this.readSharedIgnoreFromProps(line?.props || null);
  }

  injectCss() {
    this.ui.injectCSS(`
      .trc-ref-anchor {
        position: relative;
      }

      .trc-refcount-badge-wrap {
        position: absolute;
        left: auto;
        right: 2.5px;
        top: 0;
        transform: translate(8px, -0.34em);
        display: inline-flex;
        align-items: center;
        line-height: 1;
        user-select: none;
        pointer-events: none;
        z-index: 2;
      }

      .trc-refcount-badge {
        color: var(--text-muted, var(--text-default, var(--text, inherit)));
        font-family: var(--ed-variable-width-font, var(--font-sans, inherit));
        font-weight: 650;
        line-height: 1;
        padding: 0;
        opacity: ${this._opacity};
        pointer-events: none;
      }

      .trc-refcount-badge-wrap.trc-size-xsmall .trc-refcount-badge { font-size: 9px; }
      .trc-refcount-badge-wrap.trc-size-small .trc-refcount-badge { font-size: 10px; }
      .trc-refcount-badge-wrap.trc-size-medium .trc-refcount-badge { font-size: 11px; }
      .trc-refcount-badge-wrap.trc-size-large .trc-refcount-badge { font-size: 12px; }

      .trc-refcount-badge-wrap.trc-hover-only .trc-refcount-badge {
        opacity: 0;
      }

      .line-div:hover .trc-refcount-badge-wrap.trc-hover-only .trc-refcount-badge,
      .line-check-div:hover .trc-refcount-badge-wrap.trc-hover-only .trc-refcount-badge,
      .trc-refcount-badge-wrap.trc-hover-only:hover .trc-refcount-badge {
        opacity: ${this._opacity};
      }
    `);
  }

  handlePanelChanged(panel, reason) {
    if (this._isUnloading) return;
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
    this.ensureRuntimeState();
    const panelId = panel?.getId?.() || 'unknown';
    let state = this._panelStates?.get(panelId) || null;
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

    this._panelStates?.set(panelId, state);
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

    const observer = new MutationObserver((mutations) => {
      if (state.ignoreMutationsUntil > Date.now()) return;
      if (!this.hasReferenceRelevantMutation(mutations)) return;
      this.scheduleScan(state, { force: false, reason: 'dom-mutation' });
    });

    observer.observe(panelEl, {
      childList: true,
      subtree: true,
      attributes: true,
      attributeFilter: ['class', 'data-guid', 'data-ref-guid', 'data-record-guid', 'data-link-guid', 'data-root-id']
    });
    state.observer = observer;
    state.observerEl = panelEl;
  }

  hasReferenceRelevantMutation(mutations) {
    if (!Array.isArray(mutations) || mutations.length === 0) return false;

    for (const m of mutations) {
      if (!m) continue;

      if (m.type === 'attributes') {
        if (this.nodeHasReferenceHint(m.target)) return true;
        continue;
      }

      if (m.type !== 'childList') continue;

      const added = m.addedNodes || [];
      for (const node of added) {
        if (this.nodeHasReferenceHint(node)) return true;
      }

      const removed = m.removedNodes || [];
      for (const node of removed) {
        if (this.nodeHasReferenceHint(node)) return true;
      }
    }

    return false;
  }

  nodeHasReferenceHint(node) {
    if (!(node instanceof Element)) return false;
    if (node.classList?.contains('trc-refcount-badge-wrap')) return false;

    const selector = '.lineitem-ref, .lineitem-ref-title, .lineitem-lineref, .line-ref, .segment-ref, .record-link, [data-ref-guid], [data-record-guid], [data-link-guid], [data-guid], [data-root-id]';
    if (node.matches?.(selector)) return true;
    if (node.querySelector?.(selector)) return true;

    return false;
  }

  handlePanelClosed(panel) {
    const panelId = panel?.getId?.() || null;
    if (!panelId) return;
    this.disposePanelState(panelId);
  }

  disposePanelState(panelId) {
    const state = this._panelStates?.get(panelId) || null;
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

    this._panelStates?.delete(panelId);
  }

  scheduleScan(state, { force, reason }) {
    if (!state || this._isUnloading) return;

    if (state.scanTimer) {
      if (!force) return;
      clearTimeout(state.scanTimer);
      state.scanTimer = null;
    }

    const delay = force ? 0 : this._refreshDebounceMs;
    state.scanTimer = setTimeout(() => {
      state.scanTimer = null;
      if (this._isUnloading) return;
      this.scanPanel(state.panelId, reason || 'scheduled').catch(() => {
        // ignore
      });
    }, delay);
  }

  refreshAllPanels({ force, reason }) {
    if (this._isUnloading) return;
    for (const state of this._panelStates?.values() || []) {
      this.scheduleScan(state, { force: force === true, reason: reason || 'all-panels' });
    }
  }

  async scanPanel(panelId, reason) {
    if (this._isUnloading) return;
    const state = this._panelStates?.get(panelId) || null;
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

    const refs = this.collectReferenceTargets(editorRoot);
    const activeTargets = new Set(refs.map((x) => x.anchor || x.el));
    if (refs.length === 0) {
      state.ignoreMutationsUntil = Date.now() + 140;
      this.removeOrphanBadges(editorRoot, activeTargets);
      return;
    }

    const uniqueGuids = Array.from(new Set(refs.map((x) => x.guid)));
    const counts = new Map();

    await Promise.all(
      uniqueGuids.map(async (guid) => {
        const info = await this.getCountInfoForGuid(guid);
        counts.set(guid, info);
      })
    );

    if (this._isUnloading) return;
    if (!this._panelStates?.has(panelId)) return;
    if (state.scanSeq !== seq) return;

    state.ignoreMutationsUntil = Date.now() + 180;
    for (const ref of refs) {
      const info = counts.get(ref.guid) || { count: 0, capped: false };
      if ((!this._showZero && info.count <= 0) || info.count < this._minCount) {
        this.removeBadgeFromTarget(ref.anchor || ref.el);
        continue;
      }
      this.upsertBadge(ref.anchor || ref.el, ref.guid, info);
    }

    this.removeOrphanBadges(editorRoot, activeTargets);
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

    const anchors = rootEl.querySelectorAll('.trc-ref-anchor');
    for (const anchor of anchors) {
      if (!anchor.querySelector('.trc-refcount-badge-wrap')) {
        anchor.style.removeProperty('margin-right');
        anchor.removeAttribute('data-trc-original-inline-margin-right');
        anchor.classList.remove('trc-ref-anchor');
      }
    }
  }

  removeBadgeFromTarget(targetEl) {
    if (!targetEl?.querySelectorAll) return;

    const directBadges = targetEl.querySelectorAll(':scope > .trc-refcount-badge-wrap');
    for (const badge of directBadges) {
      try {
        badge.remove();
      } catch (e) {
        // ignore
      }
    }

    let sibling = targetEl.nextElementSibling;
    while (sibling?.classList?.contains('trc-refcount-badge-wrap')) {
      const next = sibling.nextElementSibling;
      try {
        sibling.remove();
      } catch (e) {
        // ignore
      }
      sibling = next;
    }

    if (!targetEl.querySelector(':scope > .trc-refcount-badge-wrap')) {
      targetEl.style.removeProperty('margin-right');
      targetEl.removeAttribute('data-trc-original-inline-margin-right');
      targetEl.classList.remove('trc-ref-anchor');
    }
  }

  removeOrphanBadges(editorRoot, activeTargets) {
    if (!editorRoot?.querySelectorAll) return;

    const badges = editorRoot.querySelectorAll('.trc-refcount-badge-wrap');
    for (const badge of badges) {
      const parent = badge.parentElement;
      const anchored = parent && activeTargets.has(parent);
      if (anchored) continue;

      try {
        badge.remove();
      } catch (e) {
        // ignore
      }
    }

    const anchors = editorRoot.querySelectorAll('.trc-ref-anchor');
    for (const anchor of anchors) {
      if (!activeTargets.has(anchor) || !anchor.querySelector('.trc-refcount-badge-wrap')) {
        anchor.style.removeProperty('margin-right');
        anchor.removeAttribute('data-trc-original-inline-margin-right');
        anchor.classList.remove('trc-ref-anchor');
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
    const seenTargets = new Set();

    for (const node of nodes) {
      if (seen.has(node)) continue;
      seen.add(node);

      const el = this.resolveReferenceTargetElement(node);
      if (!el) continue;
      if (seenTargets.has(el)) continue;
      seenTargets.add(el);

      if (!this.isLikelyReferenceElement(el)) continue;
      const guid = this.extractRecordGuidFromElement(el);
      if (!guid) continue;

      const anchor = this.resolveBadgeAnchorElement(el);
      if (!anchor) continue;

      out.push({ el, guid, anchor });
    }

    return out;
  }

  resolveBadgeAnchorElement(refEl) {
    if (!refEl || refEl.nodeType !== 1) return null;

    const directArrow = refEl.querySelector(':scope > .lineitem-lineref');
    if (directArrow) return directArrow;

    const anyArrow = refEl.querySelector('.lineitem-lineref');
    if (anyArrow) return anyArrow;

    return refEl;
  }

  resolveReferenceTargetElement(node) {
    if (!node || node.nodeType !== 1) return null;
    if (node.closest('.trc-refcount-badge-wrap')) return null;

    const refRoot = node.closest('.lineitem-ref');
    if (refRoot?.classList?.contains('lineitem-ref')) {
      return refRoot;
    }

    if (node.classList?.contains('lineitem-lineref')) {
      return null;
    }

    let el = node;
    for (let depth = 0; el && depth < 4; depth += 1) {
      if (
        el.hasAttribute('data-ref-guid') ||
        el.hasAttribute('data-record-guid') ||
        el.hasAttribute('data-link-guid') ||
        el.hasAttribute('data-guid') ||
        el.hasAttribute('data-root-id')
      ) {
        return el;
      }
      el = el.parentElement;
    }

    return node;
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
      cls.includes('lineref') ||
      cls.includes('link-menu-opener') ||
      el.classList.contains('line-div') ||
      el.classList.contains('line-check-div') ||
      el.classList.contains('listitem')
    ) {
      return false;
    }

    // Explicit reference nodes should not be rejected by generic size heuristics
    // (long quoted titles can exceed text-length thresholds).
    const explicitRef =
      el.classList.contains('lineitem-ref') ||
      el.classList.contains('record-link') ||
      el.classList.contains('line-ref') ||
      el.classList.contains('segment-ref') ||
      el.hasAttribute('data-ref-guid') ||
      el.hasAttribute('data-record-guid') ||
      el.hasAttribute('data-link-guid');
    if (explicitRef) return true;

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

  upsertBadge(targetEl, guid, info) {
    if (!targetEl?.isConnected) return;
    if (targetEl.closest('.trc-refcount-badge-wrap')) return;

    targetEl.classList.add('trc-ref-anchor');
    targetEl.style.removeProperty('margin-right');
    targetEl.removeAttribute('data-trc-original-inline-margin-right');

    let legacy = targetEl.nextElementSibling;
    while (legacy?.classList?.contains('trc-refcount-badge-wrap')) {
      const next = legacy.nextElementSibling;
      try {
        legacy.remove();
      } catch (e) {
        // ignore
      }
      legacy = next;
    }

    let wrap = targetEl.querySelector(':scope > .trc-refcount-badge-wrap');
    if (!wrap) {
      wrap = document.createElement('span');
      wrap.className = 'trc-refcount-badge-wrap';
      try {
        targetEl.appendChild(wrap);
      } catch (e) {
        return;
      }
    }

    wrap.classList.remove('trc-size-xsmall', 'trc-size-small', 'trc-size-medium', 'trc-size-large');
    wrap.classList.add(this._fontScaleClass);
    wrap.classList.toggle('trc-hover-only', this._hoverOnly);
    wrap.dataset.guid = guid;

    let label = wrap.querySelector(':scope > .trc-refcount-badge');
    if (!label) {
      label = document.createElement('span');
      label.className = 'trc-refcount-badge text-details';
      label.setAttribute('aria-hidden', 'true');
      wrap.appendChild(label);
    }

    label.textContent = this.formatCountLabel(info);

    const recordName = this.getOrLoadRecordName(guid);
    const tooltip = `${info.count}${info.capped ? '+' : ''} references to ${recordName}`;
    wrap.setAttribute('title', tooltip);
  }

  formatCountLabel(info) {
    const count = Number(info?.count) || 0;
    if (count >= 1000) {
      const compact = Math.round((count / 1000) * 10) / 10;
      return `${compact}k${info?.capped ? '+' : ''}`;
    }
    return `${count}${info?.capped ? '+' : ''}`;
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
      const recordsInfo = await this.loadRecordBackReferenceCount(guid);
      return { count: recordsInfo.count, capped: false };
    }

    const linesInfo = await this.loadLineReferenceCount(guid);
    if (this._countMode === 'lines') {
      return { count: linesInfo.count, capped: linesInfo.capped };
    }

    const propertyInfo = await this.loadPropertyReferenceRecordCount(guid);
    const combinedCount = this.combineLineAndPropertyCounts(linesInfo, propertyInfo);
    return {
      count: combinedCount,
      capped: linesInfo.capped
    };
  }

  async loadLineReferenceCount(guid) {
    const query = `@linkto = "${guid}"`;
    const result = await this.data.searchByQuery(query, this._maxResults);
    if (result?.error) {
      return {
        count: 0,
        capped: false,
        sourceRecordGuids: new Set()
      };
    }

    const lines = Array.isArray(result?.lines) ? result.lines : [];
    const uniqueLineRefs = new Set();
    const sourceRecordGuids = new Set();
    let fallbackIdx = 0;

    for (const line of lines) {
      if (!line) continue;
      if (this.isLineSharedIgnored(line)) continue;
      const sourceRecordGuid = line?.record?.guid || '';
      if (!this._showSelf && sourceRecordGuid === guid) continue;

      const key = line.guid || `${sourceRecordGuid}:${fallbackIdx++}`;
      uniqueLineRefs.add(key);
      if (sourceRecordGuid) sourceRecordGuids.add(sourceRecordGuid);
    }

    return {
      count: uniqueLineRefs.size,
      capped: lines.length >= this._maxResults,
      sourceRecordGuids
    };
  }

  async loadRecordBackReferenceCount(guid) {
    const record = this.data.getRecord?.(guid) || null;
    if (!record?.getBackReferenceRecords) {
      return {
        count: 0,
        recordGuids: new Set()
      };
    }

    const refs = await record.getBackReferenceRecords();
    let records = Array.isArray(refs) ? refs : [];
    if (!this._showSelf) records = records.filter((x) => x?.guid !== guid);

    const recordGuids = new Set();
    for (const rec of records) {
      const recGuid = rec?.guid || '';
      if (!recGuid) continue;
      recordGuids.add(recGuid);
    }

    return {
      count: recordGuids.size,
      recordGuids
    };
  }

  async loadPropertyReferenceRecordCount(targetGuid) {
    if (!targetGuid) {
      return {
        count: 0,
        recordGuids: new Set()
      };
    }

    const allRecords = this.data.getAllRecords?.() || [];
    const recordGuids = new Set();

    for (const src of allRecords || []) {
      const srcGuid = src?.guid || '';
      if (!srcGuid) continue;
      if (!this._showSelf && srcGuid === targetGuid) continue;

      const props = src.getAllProperties?.() || [];
      let hit = false;
      for (const prop of props || []) {
        if (!this.propertyReferencesGuid(prop, targetGuid)) continue;
        hit = true;
        break;
      }

      if (hit) recordGuids.add(srcGuid);
    }

    return {
      count: recordGuids.size,
      recordGuids
    };
  }

  propertyReferencesGuid(prop, targetGuid) {
    if (!prop || !targetGuid) return false;

    const values = this.getPropertyCandidateValues(prop);
    for (const v of values) {
      if (v === targetGuid) return true;
    }
    return false;
  }

  getPropertyCandidateValues(prop) {
    const out = [];
    const seen = new Set();

    const push = (v) => {
      if (typeof v !== 'string') return;
      const t = v.trim();
      if (!t) return;
      if (seen.has(t)) return;
      seen.add(t);
      out.push(t);
    };

    let raw = [];
    try {
      raw.push(prop.text?.());
    } catch (e) {
      // ignore
    }
    try {
      raw.push(prop.choice?.());
    } catch (e) {
      // ignore
    }

    for (const r of raw) {
      for (const v of this.expandPossibleListString(r)) {
        push(v);
      }
    }

    return out;
  }

  expandPossibleListString(v) {
    if (typeof v !== 'string') return [];
    const t = v.trim();
    if (!t) return [];

    if (t.startsWith('[') && t.endsWith(']')) {
      try {
        const parsed = JSON.parse(t);
        if (Array.isArray(parsed)) {
          return parsed
            .filter((x) => typeof x === 'string')
            .map((x) => x.trim())
            .filter(Boolean);
        }
      } catch (e) {
        // fall through
      }
    }

    if (t.includes(',')) {
      return t
        .split(',')
        .map((x) => x.trim())
        .filter(Boolean);
    }

    if (t.includes('\n')) {
      return t
        .split(/\r?\n/)
        .map((x) => x.trim())
        .filter(Boolean);
    }

    return [t];
  }

  combineLineAndPropertyCounts(linesInfo, propertyInfo) {
    if (!linesInfo || !propertyInfo) return 0;
    if (linesInfo.capped) return linesInfo.count;

    let propertyOnlyRecords = 0;
    for (const recordGuid of propertyInfo.recordGuids || []) {
      if (linesInfo.sourceRecordGuids?.has(recordGuid)) continue;
      propertyOnlyRecords += 1;
    }

    return linesInfo.count + propertyOnlyRecords;
  }

  clearCountCache() {
    this._countCache?.clear();
    this._recordExistsCache?.clear();
    this._recordNameCache?.clear();
  }

  handleLineItemUpdated(ev) {
    const ignoreChanged = this.didSharedIgnoreChange(ev);
    const hasSegments = ev?.hasSegments?.() && typeof ev.getSegments === 'function';
    if (!hasSegments && !ignoreChanged) return;

    const lineGuid = typeof ev.lineItemGuid === 'string' ? ev.lineItemGuid : null;
    const previousRefs = lineGuid ? (this._lineRefGuids?.get(lineGuid) || new Set()) : new Set();
    let currentRefs = previousRefs;
    if (hasSegments) {
      const segments = ev.getSegments() || [];
      currentRefs = this.extractRefGuidsFromSegments(segments);
    }

    if (lineGuid && hasSegments) {
      this._lineRefGuids?.set(lineGuid, currentRefs);
    }

    const refsChanged = hasSegments ? (!lineGuid || !this.areSetsEqual(previousRefs, currentRefs)) : false;
    if (refsChanged || ignoreChanged) {
      const affected = new Set([...previousRefs, ...currentRefs]);
      for (const guid of affected) {
        this._countCache?.delete(guid);
        this._recordNameCache?.delete(guid);
      }
    }

    if (currentRefs.size === 0 && previousRefs.size === 0) return;

    this.refreshAllPanels({
      force: false,
      reason: refsChanged ? 'lineitem.ref-change' : ignoreChanged ? 'lineitem.ignore-change' : 'lineitem.ref-redraw'
    });
  }

  handleLineItemDeleted(ev) {
    const lineGuid = typeof ev?.lineItemGuid === 'string' ? ev.lineItemGuid : null;
    if (!lineGuid) return;

    const previousRefs = this._lineRefGuids?.get(lineGuid) || null;
    this._lineRefGuids?.delete(lineGuid);
    if (!previousRefs || previousRefs.size === 0) return;

    for (const guid of previousRefs) {
      this._countCache?.delete(guid);
      this._recordNameCache?.delete(guid);
    }

    this.refreshAllPanels({ force: false, reason: 'lineitem.deleted' });
  }

  handleRecordUpdated(ev) {
    const recordGuid = typeof ev?.recordGuid === 'string' ? ev.recordGuid : null;
    if (!recordGuid) return;
    this._recordNameCache?.delete(recordGuid);
  }

  extractRefGuidsFromSegments(segments) {
    const out = new Set();
    if (!Array.isArray(segments)) return out;

    for (const seg of segments) {
      if (seg?.type !== 'ref') continue;

      let guid = '';
      if (typeof seg.text === 'string') guid = seg.text;
      else if (seg.text && typeof seg.text.guid === 'string') guid = seg.text.guid;

      guid = (guid || '').trim();
      if (!guid) continue;
      out.add(guid);
    }

    return out;
  }

  areSetsEqual(a, b) {
    if (a === b) return true;
    if (!a || !b) return false;
    if (a.size !== b.size) return false;
    for (const value of a) {
      if (!b.has(value)) return false;
    }
    return true;
  }
}
