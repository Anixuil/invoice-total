import { createMotion } from './motion.js?v=20260924-fade-tabs';
import { createPageContext } from './page-context.js';

// Keep lazy page assets on the same release as the application shell.
const assetVersion = new URL(import.meta.url).searchParams.get('v') || 'dev';
const pageAsset = (page, extension) => `/static/pages/${page}.${extension}?v=${encodeURIComponent(assetVersion)}`;

const routes = new Map([
  ['/', { page: 'invoice', title: '发票合计 · 本地核算工具' }],
  ['/jira', { page: 'jira', title: 'Jira 数据处理 · 本地工具' }],
  ['/weekly-report', { page: 'weekly-report', title: '周报整合 · 本地工具' }],
  ['/reimbursement', { page: 'reimbursement', title: '报销单生成' }],
  ['/image-to-ppt', { page: 'image-to-ppt', title: '图片转 PPT · 本地工具' }],
]);
const outlet = document.getElementById('pageOutlet');
const status = document.getElementById('routeStatus');
const message = document.getElementById('routeMessage');
const retry = document.getElementById('routeRetry');
const motion = createMotion();
const assets = new Map();
const styles = new Map();
const instances = new Map();
let current = null;
let navigation = 0;
let failedUrl = null;
let configPromise = null;
let historyRecovery = null;
let resolveRecovery = null;
let recoveryUrl = '';

history.scrollRestoration = 'manual';
if (!Number.isInteger(history.state?.appIndex)) {
  history.replaceState({ ...history.state, appIndex: 0 }, '', location.href);
}

function parseRoute(value) {
  const url = new URL(value, location.href);
  const definition = routes.get(url.pathname);
  if (url.origin !== location.origin || !definition) return null;
  const mode = definition.page === 'jira' && url.searchParams.get('mode') === 'daily' ? 'daily' : 'weekly';
  return { ...definition, url, key: definition.page === 'jira' ? `jira:${mode}` : definition.page };
}

function loadConfig() {
  if (!configPromise) {
    configPromise = fetch('/api/ui-config', { cache: 'no-store' }).then(async response => {
      if (!response.ok) throw new Error('无法读取功能配置');
      const config = await response.json();
      document.getElementById('imagePptNavigation').hidden = !config.image_to_ppt_available;
      return config;
    }).catch(error => { configPromise = null; throw error; });
  }
  return configPromise;
}

function stylesheet(page) {
  if (styles.has(page)) return styles.get(page);
  const promise = new Promise((resolve, reject) => {
    const link = document.createElement('link');
    link.rel = 'stylesheet';
    link.href = pageAsset(page, 'css');
    link.onload = resolve;
    link.onerror = () => { link.remove(); styles.delete(page); reject(new Error('页面样式加载失败，请重试')); };
    document.head.insertBefore(link, document.querySelector('link[href*="app.css"]'));
  });
  styles.set(page, promise);
  return promise;
}

async function getInstance(route) {
  if (route.page === 'image-to-ppt' && !(await loadConfig()).image_to_ppt_available) {
    throw new Error('图片转 PPT 功能未部署');
  }
  if (instances.has(route.key)) return instances.get(route.key);
  if (!assets.has(route.page)) {
    assets.set(route.page, Promise.all([
      fetch(pageAsset(route.page, 'html')).then(response => {
        if (!response.ok) throw new Error(`页面加载失败 (${response.status})`);
        return response.text();
      }),
      import(pageAsset(route.page, 'js')),
      stylesheet(route.page),
    ]).catch(error => { assets.delete(route.page); throw error; }));
  }
  const [template, module] = await assets.get(route.page);
  // Concurrent navigation requests can share the same module load.
  if (instances.has(route.key)) return instances.get(route.key);
  const root = document.createElement('section');
  root.className = 'page-root';
  root.dataset.page = route.page;
  root.innerHTML = template;
  const context = createPageContext(root, route, motion);
  const page = module.createPage(context);
  try { page.mount(); }
  catch (error) { page.destroy(); throw error; }
  const instance = { root, context, page, visited: false };
  instances.set(route.key, instance);
  return instance;
}

function navigationState(route) {
  document.querySelectorAll('.top-nav a').forEach(link => {
    const selected = link.pathname === route.url.pathname;
    link.classList.toggle('is-active', selected);
    if (selected) link.setAttribute('aria-current', 'page');
    else link.removeAttribute('aria-current');
  });
  const nav = document.querySelector('.top-nav');
  const selected = nav.querySelector('.is-active');
  if (selected && nav.scrollWidth > nav.clientWidth) {
    const offset = selected.getBoundingClientRect().left - nav.getBoundingClientRect().left;
    nav.scrollLeft += offset - (nav.clientWidth - selected.offsetWidth) / 2;
  }
}

function restoreHistory() {
  if (!current) return;
  const index = history.state?.appIndex;
  const delta = Number.isInteger(index) ? current.index - index : 0;
  if (delta) {
    recoveryUrl = current.route.url.href;
    historyRecovery = new Promise(resolve => { resolveRecovery = resolve; });
    history.go(delta);
  } else {
    history.replaceState({ ...history.state, appIndex: current.index }, '', current.route.url);
  }
}

async function navigate(value, mode = 'push') {
  const ticket = ++navigation;
  if (historyRecovery) await historyRecovery;
  if (ticket !== navigation) return;
  const route = parseRoute(value);
  if (!route) return;
  if (current) { motion.cleanup(current.instance.root); current.instance.root.inert = false; }
  status.hidden = true;
  retry.hidden = true;
  const previous = current;
  let next = null;
  try {
    next = await getInstance(route);
    if (ticket !== navigation) return;
    const changed = previous?.instance !== next;
    if (changed && previous) {
      previous.instance.root.inert = true;
      await motion.leave(previous.instance.root);
      if (ticket !== navigation) return;
    }
    // History and visible content commit together, with no await in between.
    if (changed) {
      previous?.instance.page.deactivate();
      document.body.dataset.page = route.page;
      outlet.replaceChildren(next.root);
      next.page.activate(route);
    } else {
      next.context.route = route;
    }
    let index = history.state?.appIndex ?? 0;
    if (mode === 'push' && location.href !== route.url.href) {
      index += 1;
      history.pushState({ appIndex: index }, '', route.url);
    }
    current = { instance: next, route, index };
    document.title = next.context.title;
    navigationState(route);
    status.hidden = true;
    failedUrl = null;
    if (changed) {
      const [left, top] = next.context.state.ui.scroll;
      window.scrollTo({ left, top, behavior: 'instant' });
      // Swap layout while transparent, then fade in without moving the content.
      motion.fadeIn(next.root);
      next.visited = true;
    }
    if (route.url.hash) {
      let id;
      try { id = decodeURIComponent(route.url.hash.slice(1)); } catch { id = route.url.hash.slice(1); }
      next.root.querySelector(`#${CSS.escape(id)}`)?.scrollIntoView({ behavior: 'instant' });
    }
  } catch (error) {
    if (ticket !== navigation) return;
    if (previous) {
      if (next && next !== previous.instance) next.page.deactivate();
      document.body.dataset.page = previous.route.page;
      outlet.replaceChildren(previous.instance.root);
      previous.instance.page.activate(previous.route);
      motion.cleanup(previous.instance.root);
      previous.instance.root.inert = false;
      current = previous;
      document.title = previous.instance.context.title;
      navigationState(previous.route);
      if (mode === 'pop') restoreHistory();
    }
    failedUrl = route.url.href;
    status.hidden = false;
    message.textContent = error.message || '页面加载失败';
    retry.hidden = false;
  }
}

document.addEventListener('click', event => {
  const link = event.target.closest?.('a[href]');
  if (link?.classList.contains('skip-link')) return;
  if (!link || event.defaultPrevented || event.button !== 0 || event.metaKey || event.ctrlKey ||
      event.shiftKey || event.altKey || link.hasAttribute('download') ||
      (link.target && link.target !== '_self') || !parseRoute(link.href)) return;
  event.preventDefault();
  void navigate(link.href);
});
window.addEventListener('popstate', () => {
  if (resolveRecovery) {
    const restored = location.href === recoveryUrl;
    const resolve = resolveRecovery;
    resolveRecovery = null;
    historyRecovery = null;
    resolve();
    if (restored) return;
  }
  void navigate(location.href, 'pop');
});
retry.addEventListener('click', () => { if (failedUrl) void navigate(failedUrl); });
window.addEventListener('pagehide', event => {
  if (event.persisted) return;
  navigation += 1;
  instances.forEach(instance => instance.page.destroy());
});
// A failed capability lookup does not block the other four tools.
void loadConfig().catch(() => {});
void navigate(location.href, 'initial');
