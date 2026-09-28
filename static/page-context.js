import { uploadFormData } from './upload.js';

// A page stays alive while detached. Only destruction aborts its tasks.
export function createPageContext(root, route, motion) {
  const state = { ui: { fields: {}, scroll: [0, 0], innerScroll: [] } };
  const lifetime = new AbortController();
  const globalListeners = [];
  const deactivateHooks = [];
  const destroyHooks = [];
  const timers = new Set();
  const intervals = new Set();
  const requests = new Set();
  const objectUrls = new Set();
  const downloadItems = new Map();
  let active = false;
  let destroyed = false;
  let activation = 0;
  let mounted = false;
  let focusTarget = null;
  let downloadSequence = 0;
  let downloadVersion = 0;
  const nativeFetch = window.fetch.bind(window);
  const nativeSetTimeout = window.setTimeout.bind(window);
  const nativeClearTimeout = window.clearTimeout.bind(window);

  const recovery = document.createElement('aside');
  recovery.className = 'download-recovery';
  recovery.setAttribute('aria-label', '已生成文件');
  recovery.hidden = true;

  function renderDownloads() {
    recovery.replaceChildren();
    recovery.hidden = downloadItems.size === 0;
    if (!downloadItems.size) return;
    const note = document.createElement('p');
    note.textContent = '文件已就绪。如下载未开始，可点击下方链接。';
    recovery.append(note);
    for (const item of downloadItems.values()) {
      const link = document.createElement('a');
      link.href = item.url;
      if (item.filename) link.download = item.filename;
      link.textContent = item.filename || '下载生成文件';
      recovery.append(link);
    }
  }

  function flushDownloads() {
    if (!active || destroyed) return;
    for (const item of downloadItems.values()) {
      if (item.attempted) continue;
      item.attempted = true;
      const link = document.createElement('a');
      link.href = item.url;
      if (item.filename) link.download = item.filename;
      link.hidden = true;
      root.append(link);
      link.click();
      link.remove();
    }
  }

  class PageRequest extends window.XMLHttpRequest {
    constructor() {
      super();
      requests.add(this);
      this.addEventListener('loadend', () => requests.delete(this), { once: true });
    }
    send(body) {
      if (destroyed) throw new DOMException('页面已销毁', 'AbortError');
      super.send(body);
    }
  }

  const context = {
    root, state, route, motion, title: route.title,
    get active() { return active; },
    get activation() { return activation; },
    isCurrentActivation(value) { return active && activation === value && !destroyed; },
    XMLHttpRequest: PageRequest,
    async fetch(input, options = {}) {
      const controller = new AbortController();
      const abort = () => controller.abort();
      const signals = [lifetime.signal, options.signal].filter(Boolean);
      signals.forEach(signal => {
        if (signal.aborted) abort();
        else signal.addEventListener('abort', abort, { once: true });
      });
      try { return await nativeFetch(input, { ...options, signal: controller.signal }); }
      finally { signals.forEach(signal => signal.removeEventListener('abort', abort)); }
    },
    uploadFormData(url, body, onProgress) { return uploadFormData(url, body, onProgress, PageRequest); },
    setTimeout(callback, delay, ...args) {
      if (destroyed) return null;
      const timer = nativeSetTimeout(() => {
        timers.delete(timer);
        if (!destroyed) callback(...args);
      }, delay);
      timers.add(timer);
      return timer;
    },
    clearTimeout(timer) { nativeClearTimeout(timer); timers.delete(timer); },
    visualInterval(callback, delay) {
      const entry = { callback, delay, timer: null };
      if (active) entry.timer = window.setInterval(callback, delay);
      intervals.add(entry);
      return entry;
    },
    clearVisualInterval(entry) {
      if (!entry) return;
      window.clearInterval(entry.timer);
      intervals.delete(entry);
    },
    listenGlobal(target, type, handler, options) {
      // Capture handlers must preserve stopImmediatePropagation and event ordering.
      const guarded = event => { if (active && !destroyed) handler(event); };
      const binding = { target, type, guarded, options };
      globalListeners.push(binding);
      if (active) target.addEventListener(type, guarded, options);
    },
    onDeactivate(callback) { deactivateHooks.push(callback); },
    onDestroy(callback) { destroyHooks.push(callback); },
    setScrollLock(locked) { if (active) document.body.style.overflow = locked ? 'hidden' : ''; },
    focus(element, options = { preventScroll: true }) { if (active) element?.focus(options); },
    scrollTo(options) {
      if (active) window.scrollTo({ ...options, behavior: motion.quiet() ? 'auto' : options.behavior });
    },
    scrollIntoView(element, options) {
      if (active) element.scrollIntoView({ ...options, behavior: motion.quiet() ? 'auto' : options.behavior });
    },
    createObjectURL(blob) { const url = URL.createObjectURL(blob); objectUrls.add(url); return url; },
    revokeObjectURL(url) { URL.revokeObjectURL(url); objectUrls.delete(url); },
    downloads: {
      get version() { return downloadVersion; },
      request({ url, blob, filename, key, version = downloadVersion }) {
        if (destroyed || version !== downloadVersion) return;
        const id = key || url || `download-${++downloadSequence}`;
        const previous = downloadItems.get(id);
        if (previous?.owned) context.revokeObjectURL(previous.url);
        downloadItems.set(id, { url: blob ? context.createObjectURL(blob) : url, filename, owned: Boolean(blob), attempted: false });
        renderDownloads();
        flushDownloads();
      },
      clear(key) {
        downloadVersion += 1;
        for (const [id, item] of downloadItems) {
          if (key && id !== key) continue;
          if (item.owned) context.revokeObjectURL(item.url);
          downloadItems.delete(id);
        }
        renderDownloads();
      },
    },
    lifecycle(initialize) {
      return {
        mount() {
          if (mounted || destroyed) return;
          initialize();
          (root.querySelector('.download-slot') || root).append(recovery);
          root.addEventListener('input', event => {
            const element = event.target;
            if (element.id && element.type !== 'file') state.ui.fields[element.id] = element.value;
          }, { signal: lifetime.signal });
          mounted = true;
        },
        activate(nextRoute) {
          if (destroyed || active) return;
          context.route = nextRoute;
          activation += 1;
          active = true;
          root.inert = false;
          globalListeners.forEach(({ target, type, guarded, options }) => target.addEventListener(type, guarded, options));
          intervals.forEach(entry => { entry.timer = window.setInterval(entry.callback, entry.delay); });
          motion.init(root);
          for (const [element, left, top] of state.ui.innerScroll) {
            element.scrollLeft = left;
            element.scrollTop = top;
          }
          const heading = focusTarget?.isConnected ? focusTarget : root.querySelector('h1');
          if (heading) { if (!heading.hasAttribute('tabindex')) heading.tabIndex = -1; heading.focus({ preventScroll: true }); }
          flushDownloads();
        },
        deactivate() {
          if (!active) return;
          state.ui.scroll = [window.scrollX, window.scrollY];
          state.ui.innerScroll = [...root.querySelectorAll('*')]
            .filter(element => element.scrollLeft || element.scrollTop)
            .map(element => [element, element.scrollLeft, element.scrollTop]);
          focusTarget = root.contains(document.activeElement) && !document.activeElement.closest('.viewer, .preview-dialog')
            ? document.activeElement : null;
          active = false;
          activation += 1;
          globalListeners.forEach(({ target, type, guarded, options }) => target.removeEventListener(type, guarded, options));
          intervals.forEach(entry => { window.clearInterval(entry.timer); entry.timer = null; });
          deactivateHooks.forEach(callback => callback());
          root.querySelectorAll('.viewer').forEach(viewer => {
            viewer.classList.remove('is-open'); viewer.inert = true; viewer.setAttribute('aria-hidden', 'true');
          });
          root.querySelectorAll('.preview-dialog').forEach(dialog => { dialog.hidden = true; });
          document.body.style.overflow = '';
          motion.deactivate(root);
          root.inert = true;
        },
        destroy() {
          this.deactivate();
          destroyed = true;
          lifetime.abort();
          requests.forEach(request => request.abort());
          requests.clear();
          timers.forEach(nativeClearTimeout);
          timers.clear();
          intervals.forEach(entry => window.clearInterval(entry.timer));
          intervals.clear();
          destroyHooks.forEach(callback => callback());
          context.downloads.clear();
          objectUrls.forEach(url => URL.revokeObjectURL(url));
          objectUrls.clear();
          root.replaceChildren();
        },
      };
    },
  };
  return context;
}
