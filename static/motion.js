export function createMotion() {
  const body = document.body;
  const motionPreference = window.matchMedia?.('(prefers-reduced-motion: reduce)');
  const animations = new Map();
  const closingDialogs = new Map();
  const visibility = new WeakMap();
  const feedbackStates = new WeakMap();
  const enhancedRoots = new WeakSet();
  let modalState = null;
  const panelSelector = '.summary-panel, .progress-panel, .weekly-job-progress, #selectedFile, .job-progress, .result';
  let keyboardInput = false;
  let motionObserver;
  const quiet = () => motionPreference?.matches || keyboardInput || document.hidden;
  const motionTokens = getComputedStyle(document.documentElement);
  const easeOut = motionTokens.getPropertyValue('--motion-ease-out').trim() || 'cubic-bezier(0.23, 1, 0.32, 1)';
  const enterDuration = parseFloat(motionTokens.getPropertyValue('--motion-enter')) || 180;
  const dialogDuration = parseFloat(motionTokens.getPropertyValue('--motion-dialog')) || 220;
  const exitDuration = parseFloat(motionTokens.getPropertyValue('--motion-select')) || 160;

  function cancel(element) {
    animations.get(element)?.cancel();
    animations.delete(element);
  }

  function play(element, frames, duration = enterDuration, delay = 0, hold = false) {
    if (!element) return;
    cancel(element);
    if (quiet() || !element.animate || !element.isConnected) return;
    try {
      const animation = element.animate(frames, { duration, delay, easing: easeOut, fill: hold ? 'both' : 'backwards' });
      animations.set(element, animation);
      const release = () => {
        if (animations.get(element) === animation) animations.delete(element);
      };
      animation.onfinish = hold ? null : release;
      animation.oncancel = release;
    } catch { /* Motion is optional; the final DOM state is already usable. */ }
  }

  function enter(element, axis = 'y', delay = 0, direction = 1) {
    play(element, [
      { opacity: 0, transform: element.matches('.view-panel') ? 'none' : axis === 'x' ? `translateX(${8 * direction}px)` : 'translateY(8px)' },
      { opacity: 1, transform: 'none' },
    ], enterDuration, delay);
  }

  function cleanup(root) {
    for (const element of animations.keys()) {
      if (root === element || root.contains(element)) cancel(element);
    }
    for (const [dialog, timer] of closingDialogs) {
      if (root === dialog || root.contains(dialog)) {
        clearTimeout(timer);
        closingDialogs.delete(dialog);
        dialog.hidden = true;
        dialog.classList.remove('motion-closing');
      }
    }
  }

  // Keep focus inside the visible dialog without making any ancestor inert.
  function modal(element, open, restore = true) {
    if (modalState && (!open || modalState.element !== element)) {
      const previous = modalState;
      modalState = null;
      previous.siblings.forEach(([node, inert]) => { node.inert = inert; });
      body.style.overflow = previous.overflow;
      if (restore && previous.trigger?.isConnected && !previous.trigger.closest('[inert]')) previous.trigger.focus({ preventScroll: true });
    }
    if (!open || !element.isConnected || modalState) return;
    const trigger = document.activeElement;
    const siblings = [];
    for (let node = element; node.parentElement && node !== body; node = node.parentElement) {
      [...node.parentElement.children].forEach(sibling => {
        if (sibling !== node && sibling instanceof HTMLElement && !['SCRIPT', 'STYLE', 'LINK'].includes(sibling.tagName)) {
          siblings.push([sibling, sibling.inert]);
          sibling.inert = true;
        }
      });
    }
    modalState = { element, siblings, trigger, overflow: body.style.overflow };
    body.style.overflow = 'hidden';
    element.tabIndex = -1;
    (element.querySelector('#viewerClose, #previewCloseButton, button:not(:disabled)') || element).focus({ preventScroll: true });
  }

  document.addEventListener('keydown', event => {
    if (!modalState) return;
    keyboardInput = true;
    body.classList.add('motion-keyboard');
    for (const element of animations.keys()) cancel(element);
    const element = modalState.element;
    if (event.key === 'Escape') {
      event.preventDefault();
      event.stopImmediatePropagation();
      element.querySelector('#viewerClose, #previewCloseButton')?.click();
    } else if (event.key === 'Tab') {
      const items = [...element.querySelectorAll('a[href], button, input, select, textarea, [tabindex]')]
        .filter(node => !node.disabled && node.tabIndex >= 0 && !node.closest('[hidden], [inert]') && node.getClientRects().length);
      const first = items[0] || element;
      const last = items.at(-1) || element;
      if (!items.length || document.activeElement === element || !element.contains(document.activeElement) || (event.shiftKey ? document.activeElement === first : document.activeElement === last)) {
        event.preventDefault();
        (event.shiftKey ? last : first).focus();
      }
    }
  }, true);
  document.addEventListener('focusin', event => {
    if (modalState && !modalState.element.contains(event.target)) modalState.element.focus({ preventScroll: true });
  });

  function dialog(element, open) {
    clearTimeout(closingDialogs.get(element));
    closingDialogs.delete(element);
    cleanup(element);
    element.classList.remove('motion-closing');
    element.inert = !open;
    element.setAttribute('aria-hidden', String(!open));
    if (open) {
      // A transformed route must never become the fixed overlay's containing block.
      for (let parent = element.parentElement; parent; parent = parent.parentElement) cancel(parent);
      element.hidden = false;
      modal(element, true);
      play(element, [{ opacity: 0 }, { opacity: 1 }], dialogDuration);
      play(element.querySelector('.preview-dialog__content'), [
        { transform: 'scale(.97)' }, { transform: 'none' },
      ], dialogDuration);
    } else if (quiet() || element.hidden) {
      element.hidden = true;
    } else {
      element.classList.add('motion-closing');
      // Cleanup of preview data stays in the caller and happens immediately.
      const timer = setTimeout(() => {
        closingDialogs.delete(element);
        element.hidden = true;
        element.classList.remove('motion-closing');
      }, exitDuration);
      closingDialogs.set(element, timer);
    }
    if (!open && modalState?.element === element) modal(element, false);
  }

  function init(root) {
    motionObserver?.disconnect();
    if (!root || typeof MutationObserver === 'undefined') return;
    const syncTabs = () => root.querySelectorAll('.tabs[role="tablist"]').forEach(group => {
      if (!group.hasAttribute('aria-label')) group.setAttribute('aria-label', '处理结果视图');
      group.querySelectorAll('.tab[data-view]').forEach(tab => {
        const panel = root.querySelector(`#${CSS.escape(tab.dataset.view)}`);
        const selected = tab.classList.contains('is-active');
        tab.id ||= `${root.dataset.page}-${tab.dataset.view}-tab`;
        tab.setAttribute('role', 'tab');
        tab.setAttribute('aria-selected', String(selected));
        tab.setAttribute('aria-controls', tab.dataset.view);
        tab.tabIndex = selected ? 0 : -1;
        if (panel) {
          panel.setAttribute('role', 'tabpanel');
          panel.setAttribute('aria-labelledby', tab.id);
          panel.tabIndex = 0;
        }
      });
    });
    syncTabs();
    if (!enhancedRoots.has(root)) {
      enhancedRoots.add(root);
      root.addEventListener('keydown', event => {
        const tab = event.target.closest('.tab[data-view]');
        if (!tab || !['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return;
        const items = [...tab.closest('.tabs').querySelectorAll('.tab')].filter(item => !item.hidden && !item.disabled);
        const index = items.indexOf(tab);
        const next = event.key === 'Home' ? 0 : event.key === 'End' ? items.length - 1 : (index + (event.key === 'ArrowRight' ? 1 : -1) + items.length) % items.length;
        event.preventDefault();
        items[next]?.click();
        items[next]?.focus();
      });
    }
    const visible = element => element.isConnected && !element.closest('[hidden]');
    const refresh = element => {
      const shown = visible(element);
      if (shown && visibility.get(element) === false) enter(element);
      if (!shown) cleanup(element);
      visibility.set(element, shown);
    };
    root.querySelectorAll(panelSelector).forEach(element => visibility.set(element, visible(element)));
    motionObserver = new MutationObserver(records => {
      if (records.some(record => record.target.matches?.('.tab, .view-panel'))) syncTabs();
      for (const record of records) {
        const target = record.target;
        if (record.type === 'attributes' && record.attributeName === 'hidden') {
          if (target.hidden) cleanup(target);
          if (target.matches(panelSelector)) refresh(target);
          target.querySelectorAll(panelSelector).forEach(refresh);
        }
        if (record.type === 'attributes' && record.attributeName === 'class' &&
            target.matches('.weekly-job-progress, .copy-button, .progress-step')) {
          const complete = target.matches('.is-complete, .is-copied, .is-done');
          if (complete && !feedbackStates.get(target) && visible(target)) {
            play(target.querySelector('.weekly-job-progress__mark') || target,
              [{ opacity: .55 }, { opacity: 1 }]);
          }
          feedbackStates.set(target, complete);
        }
        for (const node of record.addedNodes || []) {
          if (!(node instanceof Element)) continue;
          const panels = [...node.querySelectorAll(panelSelector)];
          if (node.matches(panelSelector)) panels.unshift(node);
          panels.forEach(panel => {
            visibility.set(panel, false);
            refresh(panel);
          });
        }
        for (const node of record.removedNodes || []) {
          if (node instanceof Element) cleanup(node);
        }
      }
    });
    motionObserver.observe(root, { subtree: true, childList: true, attributes: true, attributeFilter: ['hidden', 'class'] });
  }


  document.addEventListener('keydown', () => {
    keyboardInput = true;
    body.classList.add('motion-keyboard');
    for (const element of animations.keys()) cancel(element);
  }, true);
  document.addEventListener('pointerdown', () => {
    keyboardInput = false;
    body.classList.remove('motion-keyboard');
  }, true);
  document.addEventListener('toggle', event => {
    if (event.target.matches?.('details')) {
      const content = [...event.target.children].filter(child => child.tagName !== 'SUMMARY');
      content.forEach(element => event.target.open
        ? play(element, [{ opacity: 0 }, { opacity: 1 }]) : cleanup(element));
    }
  }, true);
  motionPreference?.addEventListener('change', () => {
    if (motionPreference.matches) cleanup(document.body);
  });
  document.addEventListener('visibilitychange', () => {
    body.dataset.background = String(document.hidden);
    if (document.hidden) for (const element of animations.keys()) cancel(element);
  });
  function fadeIn(root) {
    play(root, [{ opacity: 0 }, { opacity: 1 }], 140);
  }

  async function leave(root) {
    cleanup(root);
    if (quiet() || !root.animate) return;
    play(root, [{ opacity: 1 }, { opacity: 0 }], 80, 0, true);
    const animation = animations.get(root);
    if (animation?.finished) {
      let timer;
      await Promise.race([
        animation.finished.catch(() => {}),
        new Promise(resolve => { timer = setTimeout(resolve, 120); }),
      ]);
      clearTimeout(timer);
    }
  }
  function deactivate(root) {
    if (modalState && root.contains(modalState.element)) modal(modalState.element, false, false);
    motionObserver?.disconnect();
    cleanup(root);
  }
  function intro(root) {
    [...root.querySelectorAll('.hero, .upload-panel, .input-guide, #weeklyNewTaskToolbar, .upload, .details')]
      .filter(element => !element.closest('[hidden]') && !element.parentElement.closest('.upload-panel'))
      .slice(0, 4).forEach((element, index) => enter(element, 'y', index * 30));
  }
  function progress(bar, value) {
    if (!bar) return;
    const known = value !== null && value !== undefined && value !== '' && Number.isFinite(Number(value));
    const percent = known ? Math.max(0, Math.min(100, Number(value))) : null;
    bar.classList.toggle('is-indeterminate', !known);
    bar.style.transform = `scaleX(${known ? percent / 100 : 1})`;
    const track = bar.parentElement;
    track.setAttribute('role', 'progressbar');
    if (!track.hasAttribute('aria-label')) track.setAttribute('aria-label', '任务处理进度');
    track.setAttribute('aria-valuemin', '0');
    track.setAttribute('aria-valuemax', '100');
    if (known) { track.setAttribute('aria-valuenow', String(percent)); track.removeAttribute('aria-valuetext'); }
    else { track.removeAttribute('aria-valuenow'); track.setAttribute('aria-valuetext', '处理中'); }
    return percent;
  }
  return { enter, fadeIn, cancel, cleanup, dialog, modal, progress, init, leave, deactivate, intro, quiet };
}
