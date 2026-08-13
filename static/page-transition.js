(() => {
  const body = document.body;
  const reduceMotion = window.matchMedia?.('(prefers-reduced-motion: reduce)').matches;
  const duration = reduceMotion ? 0 : 340;
  let navigationPending = false;
  const routeKey = value => {
    const url = new URL(value, window.location.href);
    return `${url.pathname}${url.search}${url.hash}`;
  };
  let currentRoutePath = routeKey(window.location.href);
  const routeCache = new Map();

  const reveal = () => window.requestAnimationFrame(() => body.classList.add('page-ready'));
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', reveal, { once: true });
  } else {
    reveal();
  }

  const wait = ms => new Promise(resolve => window.setTimeout(resolve, ms));

  function bindNavigation(root = document) {
    root.querySelectorAll('.top-nav a, .mode-switch a').forEach(link => {
      if (link.dataset.routeBound === 'true') return;
      link.dataset.routeBound = 'true';
      link.addEventListener('click', event => {
        if (
        navigationPending ||
          event.defaultPrevented ||
          (event.button !== undefined && event.button !== 0) ||
          event.metaKey || event.ctrlKey || event.shiftKey || event.altKey ||
          link.target === '_blank' ||
          link.origin !== window.location.origin ||
          `${link.pathname}${link.search}${link.hash}` === currentRoutePath
        ) return;
        event.preventDefault();
        navigate(link.pathname + link.search + link.hash, true);
      });
    });
  }

  function updateActiveNavigation(path) {
    const activePath = routeKey(path).split(/[?#]/, 1)[0];
    document.querySelectorAll('.top-nav a').forEach(link => {
      link.classList.toggle('is-active', link.pathname === activePath);
    });
  }

  function syncModeSwitch(path, markup) {
    const current = document.querySelector('.mode-switch');
    const isJira = new URL(path, window.location.href).pathname === '/jira';
    if (!isJira) {
      current?.remove();
      return;
    }
    if (current || !markup) return;
    const shell = document.querySelector('main.shell');
    const routeContent = document.querySelector('#route-content');
    if (!shell || !routeContent) return;
    const fragment = document.createRange().createContextualFragment(markup);
    shell.insertBefore(fragment, routeContent);
  }

  async function navigate(path, pushHistory) {
    if (navigationPending || path === currentRoutePath) return;
    navigationPending = true;
    const currentContent = document.querySelector('#route-content');
    if (!currentContent) { navigationPending = false; window.location.assign(path); return; }
    currentContent.classList.add('route-shell');
    window.requestAnimationFrame(() => currentContent.classList.add('route-leaving'));

    const previousRoutePath = currentRoutePath;
    const previousStyle = document.querySelector('style[data-app-style]');
    const previousTitle = document.title;
    const previousModeSwitchHTML = document.querySelector('.mode-switch')?.outerHTML || '';
    const cachedRoute = routeCache.get(path);
    let nextShell = null;
    let replaced = false;
    let nextStyleElement = null;
    let loadedDocument = null;
    try {
      routeCache.set(previousRoutePath, {
        content: currentContent,
        styleName: previousStyle?.dataset.appStyle || '',
      styleText: previousStyle?.textContent || '',
      title: previousTitle,
      modeSwitchHTML: document.querySelector('.mode-switch')?.outerHTML || '',
      });
      document.body.style.overflow = '';
      document.querySelectorAll('.viewer.is-open').forEach(viewer => {
        viewer.classList.remove('is-open');
        viewer.setAttribute('aria-hidden', 'true');
      });
      const responsePromise = cachedRoute
        ? Promise.resolve(null)
        : fetch(path, { headers: { 'X-Requested-With': 'spa' } });
      const [, response] = await Promise.all([wait(duration), responsePromise]);
      let nextStyleName = '';
      let nextStyleText = '';
      if (cachedRoute) {
        nextShell = cachedRoute.content;
        nextStyleName = cachedRoute.styleName;
        nextStyleText = cachedRoute.styleText;
        document.title = cachedRoute.title;
        syncModeSwitch(path, cachedRoute.modeSwitchHTML);
      } else {
        if (!response.ok) throw new Error(`页面加载失败 (${response.status})`);
        const html = await response.text();
        loadedDocument = new DOMParser().parseFromString(html, 'text/html');
        nextShell = loadedDocument.querySelector('#route-content');
        const nextStyle = loadedDocument.querySelector('style[data-app-style]');
        if (!nextShell || !nextStyle) throw new Error('页面结构不完整');
        nextStyleName = nextStyle.dataset.appStyle;
        nextStyleText = nextStyle.textContent;
        document.title = loadedDocument.title;
        syncModeSwitch(path, loadedDocument.querySelector('.mode-switch')?.outerHTML || '');
      }

      document.querySelectorAll('style[data-app-style]').forEach(style => style.remove());
      nextStyleElement = document.createElement('style');
      nextStyleElement.dataset.appStyle = nextStyleName;
      nextStyleElement.textContent = nextStyleText;
      document.head.appendChild(nextStyleElement);

      nextShell.classList.remove('route-shell', 'route-entering', 'route-ready', 'route-leaving');
      // Jira 页面根据 window.location.search 初始化周报/每日模式，必须先同步路由再执行内联脚本。
      if (pushHistory) window.history.pushState({}, '', path);
      currentRoutePath = routeKey(path);
      currentContent.replaceWith(nextShell);
      replaced = true;
      window.scrollTo(0, 0);
      nextShell.classList.add('route-shell', 'route-entering');
      if (loadedDocument) {
        executeInlineScripts(loadedDocument);
      }
      bindNavigation(document.querySelector('main.shell'));
      updateActiveNavigation(path);
      routeCache.set(currentRoutePath, {
        content: nextShell,
        styleName: nextStyleName,
        styleText: nextStyleText,
        title: document.title,
        modeSwitchHTML: document.querySelector('.mode-switch')?.outerHTML || '',
      });
      window.requestAnimationFrame(() => nextShell.classList.add('route-ready'));
      window.setTimeout(() => {
        nextShell.classList.remove('route-shell', 'route-entering', 'route-ready');
      }, duration + 20);
    } catch (error) {
      if (replaced && nextShell?.isConnected) {
        nextShell.replaceWith(currentContent);
        nextShell.classList.remove('route-entering', 'route-leaving');
      }
      currentContent.classList.remove('route-leaving', 'route-shell');
      if (nextStyleElement?.isConnected) nextStyleElement.remove();
      if (previousStyle && !previousStyle.isConnected) document.head.appendChild(previousStyle);
      document.title = previousTitle;
      syncModeSwitch(previousRoutePath, previousModeSwitchHTML);
      if (routeKey(window.location.href) !== previousRoutePath) {
        window.history.replaceState({}, '', previousRoutePath);
      }
      currentRoutePath = previousRoutePath;
      window.console.error(error);
    } finally {
      navigationPending = false;
    }
  }

  function executeInlineScripts(documentFragment) {
    documentFragment.querySelectorAll('script:not([src])').forEach(script => {
      if (!script.textContent.trim()) return;
      // 各业务页的内联脚本只负责绑定当前页面 DOM；外部转场脚本不重复执行。
      new Function(script.textContent)();
    });
  }

  bindNavigation(document.querySelector('main.shell'));
  updateActiveNavigation(window.location.href);
  const initialContent = document.querySelector('#route-content');
  const initialStyle = document.querySelector('style[data-app-style]');
  if (initialContent && initialStyle) {
    routeCache.set(currentRoutePath, {
      content: initialContent,
      styleName: initialStyle.dataset.appStyle || '',
      styleText: initialStyle.textContent || '',
      title: document.title,
      modeSwitchHTML: document.querySelector('.mode-switch')?.outerHTML || '',
    });
  }
  window.addEventListener('popstate', () => navigate(routeKey(window.location.href), false));
})();
