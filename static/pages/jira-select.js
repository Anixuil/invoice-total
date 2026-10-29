// Preserve native select values, form submission and existing input/change handlers.
export function enhanceSelects(root) {
  const controls = new Map();
  let opened = null;
  let serial = 0;
  function close(restore = false) {
    if (!opened) return;
    const { panel, button } = opened;
    opened = null;
    panel.remove();
    button.setAttribute('aria-expanded', 'false');
    if (restore && button.isConnected) button.focus();
  }
  function sync(select, button) {
    button.disabled = select.disabled;
    const text = [...select.selectedOptions].map(item => item.textContent).join('、') || '请选择';
    if (button.textContent !== text) button.textContent = text;
    button.title = text;
    button.setAttribute('aria-label', select.getAttribute('aria-label') || select.closest('label')?.firstChild?.textContent?.trim() || '选择选项');
  }
  function emit(select) {
    select.dispatchEvent(new Event('input', { bubbles: true }));
    select.dispatchEvent(new Event('change', { bubbles: true }));
  }
  function place() {
    if (!opened) return;
    const { button, panel } = opened;
    if (!button.isConnected || button.disabled) { close(); return; }
    const rect = button.getBoundingClientRect();
    const width = Math.min(Math.max(rect.width, 260), window.innerWidth - 16);
    panel.style.width = width + 'px';
    panel.style.left = Math.max(8, Math.min(rect.left, window.innerWidth - width - 8)) + 'px';
    const below = window.innerHeight - rect.bottom - 8;
    const above = rect.top - 8;
    const height = Math.min(320, Math.max(below, above));
    panel.style.maxHeight = Math.max(120, height) + 'px';
    panel.style.top = (below >= Math.min(320, above) ? rect.bottom + 4 : Math.max(8, rect.top - Math.min(panel.scrollHeight, height) - 4)) + 'px';
  }
  function renderOptions() {
    if (!opened) return;
    const { select, search, list, button } = opened;
    const term = search.value.trim().toLocaleLowerCase();
    list.replaceChildren();
    const matches = [...select.options].filter(item => !item.hidden && `${item.textContent} ${item.value}`.toLocaleLowerCase().includes(term));
    for (const item of matches) {
      const row = document.createElement('button');
      row.type = 'button'; row.className = 'jc-select-option';
      row.setAttribute('role', 'option'); row.setAttribute('aria-selected', String(item.selected));
      row.disabled = item.disabled || item.parentElement?.disabled;
      row.textContent = (select.multiple ? (item.selected ? '☑ ' : '☐ ') : (item.selected ? '✓ ' : '')) + item.textContent;
      row.addEventListener('click', event => {
        event.preventDefault();
        if (select.disabled) return;
        if (select.multiple) item.selected = !item.selected;
        else select.value = item.value;
        emit(select);
        if (!select.multiple) close(true);
        else if (select.isConnected) { sync(select, button); renderOptions(); search.focus(); }
        else close();
      });
      list.append(row);
    }
    if (!matches.length) { const empty = document.createElement('p'); empty.textContent = '没有匹配选项'; list.append(empty); }
    place();
  }
  function open(select, button) {
    if (opened?.select === select) { close(); return; }
    close();
    if (select.disabled) return;
    const panel = document.createElement('div'); panel.className = 'jc-select-popup';
    const search = document.createElement('input'); search.type = 'search'; search.placeholder = '输入名称或编号过滤'; search.setAttribute('aria-label', '搜索选项');
    const list = document.createElement('div'); list.className = 'jc-select-options'; list.id = `jc-select-list-${++serial}`;
    list.setAttribute('role', 'listbox');
    if (select.multiple) list.setAttribute('aria-multiselectable', 'true');
    button.setAttribute('aria-controls', list.id); button.setAttribute('aria-expanded', 'true');
    panel.append(search, list);
    if (select.multiple) {
      const clear = document.createElement('button'); clear.type = 'button'; clear.textContent = '清空已选'; clear.className = 'ghost-button';
      clear.addEventListener('click', () => { for (const option of select.options) option.selected = false; emit(select); sync(select, button); renderOptions(); });
      panel.append(clear);
    }
    (select.closest('dialog') || root.querySelector('.jc') || root).append(panel);
    opened = { select, button, panel, search, list };
    search.addEventListener('input', event => { event.stopPropagation(); renderOptions(); });
    panel.addEventListener('keydown', event => {
      if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); close(true); }
      if (event.key === 'Tab') close();
      if (['ArrowDown', 'ArrowUp'].includes(event.key)) {
        event.preventDefault();
        const buttons = [...list.querySelectorAll('button:not(:disabled)')];
        const index = buttons.indexOf(document.activeElement);
        buttons[(index + (event.key === 'ArrowDown' ? 1 : -1) + buttons.length) % buttons.length]?.focus();
      }
    });
    renderOptions(); search.focus();
  }
  function refresh() {
    for (const [select, button] of controls) {
      if (!root.contains(select)) { button.remove(); controls.delete(select); if (opened?.select === select) close(); }
    }
    for (const select of root.querySelectorAll('select')) {
      let button = controls.get(select);
      if (!button) {
        button = document.createElement('button'); button.type = 'button'; button.className = 'jc-select-trigger';
        button.setAttribute('aria-haspopup', 'listbox'); button.setAttribute('aria-expanded', 'false');
        select.classList.add('jc-native-select'); select.tabIndex = -1;
        select.insertAdjacentElement('afterend', button); controls.set(select, button);
        button.addEventListener('click', event => { event.preventDefault(); open(select, button); });
      }
      sync(select, button);
    }
    if (opened) { if (opened.select.disabled) close(); else renderOptions(); }
  }
  const observer = new MutationObserver(records => {
    if (records.some(record => record.target.closest?.('select') || [...record.addedNodes, ...record.removedNodes].some(node => node.nodeType === 1 && (node.matches('select') || node.querySelector('select'))))) refresh();
  });
  observer.observe(root, { subtree: true, childList: true, attributes: true, attributeFilter: ['disabled', 'selected', 'multiple'] });
  const outside = event => { if (opened && !opened.panel.contains(event.target) && !opened.button.contains(event.target)) close(); };
  const change = event => { if (event.target.matches('select')) refresh(); };
  root.addEventListener('change', change);
  document.addEventListener('pointerdown', outside);
  window.addEventListener('resize', place);
  const scroll = event => { if (opened && !opened.panel.contains(event.target)) place(); };
  document.addEventListener('scroll', scroll, true);
  refresh();
  return { close, destroy() { close(); observer.disconnect(); root.removeEventListener('change', change); document.removeEventListener('pointerdown', outside); window.removeEventListener('resize', place); document.removeEventListener('scroll', scroll, true); controls.clear(); } };
}
