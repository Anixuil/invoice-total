// Each factory owns its DOM, files and task state for one route instance.
export function createPage(context) {
  const { root, state } = context;
  const byId = id => root.querySelector(`#${CSS.escape(id)}`);
  function mount() {
    const input = byId('file');
    const drop = byId('drop');
    const name = byId('name');
    const button = byId('generate');
    const status = byId('status');
    const result = byId('result');
    const checks = byId('checks');
    const score = byId('score');
    const download = byId('download');
    let selected = null;
    let busy = false;

    function setLoading(active) {
      busy = active;
      root.querySelector('.output-placeholder strong').textContent = active ? '正在生成与核验…' : '等待生成';
      button.disabled = active || !selected;
      input.disabled = active;
      button.setAttribute('aria-busy', String(active));
      drop.setAttribute('aria-disabled', String(active));
      button.textContent = active ? '正在生成 PPT…' : '生成可编辑 PPT';
    }
    function set(file) {
      if (!file || busy) return;
      if (!/^image\/(png|jpeg|webp)$/.test(file.type)) {
        status.textContent = '请选择 PNG、JPG、JPEG 或 WEBP 图片。';
        status.className = 'status error';
        return;
      }
      if (file.size > 20 * 1024 * 1024) {
        status.textContent = '图片不能超过 20 MB。';
        status.className = 'status error';
        return;
      }
      context.downloads.clear();
      selected = file;
      name.textContent = file.name;
      status.textContent = '';
      status.className = 'status';
      button.disabled = false;
      result.hidden = true;
    }
    input.addEventListener('change', () => set(input.files[0]));
    drop.addEventListener('keydown', event => {
      if (!busy && (event.key === 'Enter' || event.key === ' ')) {
        event.preventDefault();
        input.click();
      }
    });
    ['dragenter', 'dragover'].forEach(type => drop.addEventListener(type, event => {
      event.preventDefault();
      if (!busy) drop.classList.add('drag');
    }));
    ['dragleave', 'drop'].forEach(type => drop.addEventListener(type, event => {
      event.preventDefault();
      drop.classList.remove('drag');
    }));
    drop.addEventListener('drop', event => set(event.dataTransfer.files[0]));
    button.addEventListener('click', async () => {
      if (!selected || busy) return;
      context.downloads.clear();
      setLoading(true);
      result.hidden = true;
      status.className = 'status';
      status.textContent = '正在生成 PPT 并执行核验…';
      const body = new FormData();
      body.append('file', selected);
      try {
        const response = await context.fetch('/api/image-to-ppt/generate', { method: 'POST', body });
        const data = await response.json();
        if (!response.ok) throw new Error(data.detail || '生成失败');
        state.result = data;
        const report = data.validation;
        score.textContent = `核验得分 ${report.score}/100`;
        score.className = `score ${report.ok ? '' : 'warn'}`;
        checks.innerHTML = report.checks.map(item => `<div class="check"><strong class="${item.passed ? 'pass' : 'fail'}">${item.passed ? '通过' : '未通过'} · ${escapeHtml(item.name)}</strong><span>${escapeHtml(item.detail)}</span></div>`).join('');
        download.href = data.download_url;
        download.hidden = !report.ok;
        result.hidden = false;
        status.textContent = report.ok ? 'PPT 已生成，核验通过。' : 'PPT 已生成，但核验未全部通过。';
      } catch (error) {
        status.textContent = error.message || '生成失败';
        status.className = 'status error';
      } finally {
        setLoading(false);
      }
    });
    function escapeHtml(value) {
      return String(value || '').replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char]);
    }
    Object.defineProperty(state, 'selected', { enumerable: true, get: () => selected, set: value => { selected = value; } });
    Object.defineProperty(state, 'busy', { enumerable: true, get: () => busy });
  }
  return context.lifecycle(mount);
}
