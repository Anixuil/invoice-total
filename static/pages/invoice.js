// Each factory owns its DOM, files and task state for one route instance.
export function createPage(context) {
  const { root, state } = context;
  const byId = id => root.querySelector(`#${CSS.escape(id)}`);
  const { fetch, XMLHttpRequest, setTimeout, clearTimeout } = context;
  function mount() {
    state.taskVersion = 0;
    let copyFeedbackTimer = null;
    let uploadBusy = false;
    const drop = byId('drop');
    const fileInput = byId('file');
    const folderInput = byId('folder');
    const progress = byId('progress');
    const workspace = byId('workspace');
    const sourceTree = byId('sourceTree');
    const previewProgress = byId('previewProgress');
    const previewProgressTitle = byId('previewProgressTitle');
    const previewProgressText = byId('previewProgressText');
    const previewProgressCount = byId('previewProgressCount');
    const previewProgressBar = byId('previewProgressBar');
    const previewSubprogressText = byId('previewSubprogressText');
    const previewSubprogressBar = byId('previewSubprogressBar');
    const list = byId('list');
    const viewer = byId('viewer');
    const viewerImage = byId('viewerImage');
    const viewerTitle = byId('viewerTitle');
    const viewerOpen = byId('viewerOpen');
    const viewerClose = byId('viewerClose');
    const uppercaseAmount = byId('uppercaseAmount');
    const copyUppercase = byId('copyUppercase');
    const copyStatus = byId('copyStatus');
    
    drop.addEventListener('click', event => {
      if (!event.target.closest('button')) fileInput.click();
    });
    drop.querySelector('.select-button').addEventListener('click', event => {
      event.stopPropagation();
      fileInput.click();
    });
    drop.querySelector('.folder-button').addEventListener('click', event => {
      event.stopPropagation();
      folderInput.click();
    });
    drop.addEventListener('keydown', event => {
      if (event.target === drop && (event.key === 'Enter' || event.key === ' ')) {
        event.preventDefault();
        fileInput.click();
      }
    });
    fileInput.addEventListener('change', () => {
      if (fileInput.files.length) upload(fileInput.files);
    });
    folderInput.addEventListener('change', () => {
      const pdfFiles = Array.from(folderInput.files).filter(file => file.name.toLowerCase().endsWith('.pdf'));
      folderInput.value = '';
      if (!pdfFiles.length) {
        renderUploadError(new Error('所选目录及其子目录中没有 PDF 发票文件'));
        return;
      }
      if (pdfFiles.length > 50) {
        renderUploadError(new Error(`目录中包含 ${pdfFiles.length} 个 PDF，单次最多处理 50 个`));
        return;
      }
      upload(pdfFiles, true);
    });
    drop.addEventListener('dragover', event => {
      event.preventDefault();
      drop.classList.add('is-dragging');
    });
    drop.addEventListener('dragleave', event => {
      if (!drop.contains(event.relatedTarget)) drop.classList.remove('is-dragging');
    });
    drop.addEventListener('drop', event => {
      event.preventDefault();
      drop.classList.remove('is-dragging');
      if (event.dataTransfer.files.length) upload(event.dataTransfer.files);
    });
    
    byId('addFiles').addEventListener('click', () => fileInput.click());
    byId('clearResults').addEventListener('click', resetResults);
    copyUppercase.addEventListener('click', copyUppercaseAmount);
    list.addEventListener('click', event => {
      const button = event.target.closest('.preview-open');
      if (!button) return;
      const image = button.querySelector('img');
      openViewer(image.src, button.dataset.title || '发票票面预览');
    });
    viewerClose.addEventListener('click', closeViewer);
    viewer.addEventListener('click', event => {
      if (event.target === viewer) closeViewer();
    });
    context.listenGlobal(document, 'keydown', event => {
      if (event.key === 'Escape' && viewer.classList.contains('is-open')) closeViewer();
    });
    
    function formatBytes(bytes) {
      if (!bytes) return '0 KB';
      return bytes >= 1048576 ? `${(bytes / 1048576).toFixed(1)} MB` : `${Math.max(1, Math.round(bytes / 1024))} KB`;
    }
    
    async function upload(files, preserveRelativePath = false) {
      if (uploadBusy) return;
      const taskVersion = ++state.taskVersion;
      state.files = Array.from(files);
      files = state.files;
      setLoading(true, files.length);
      const formData = new FormData();
      for (const file of files) {
        const uploadName = preserveRelativePath && file.webkitRelativePath ? file.webkitRelativePath : file.name;
        formData.append('files', file, uploadName);
      }
      try {
        const totalBytes = Array.from(files).reduce((total, file) => total + file.size, 0);
        const data = await context.uploadFormData('/api/upload', formData, (loaded) => {
          if (taskVersion !== state.taskVersion) return;
          const percent = totalBytes ? Math.min(99, Math.round(loaded / totalBytes * 100)) : 0;
          byId('progressTitle').textContent = `正在上传 ${percent}%`;
          byId('progressText').textContent = `已上传 ${formatBytes(loaded)} / ${formatBytes(totalBytes)}`;
        });
        if (taskVersion !== state.taskVersion) return;
        render(data);
        setLoading(false);
        if (data.job_id) await pollPreviewJob(data.job_id, taskVersion);
      } catch (error) {
        if (taskVersion === state.taskVersion) renderUploadError(error);
      } finally {
        if (taskVersion === state.taskVersion) {
          setLoading(false);
          fileInput.value = '';
          folderInput.value = '';
        }
      }
    }
    
    function setLoading(active, count = 0) {
      uploadBusy = active;
      drop.setAttribute('aria-busy', String(active));
      drop.setAttribute('aria-disabled', String(active));
      drop.hidden = active;
      progress.hidden = !active;
      if (active) {
        context.motion.progress(progress.querySelector('.progress-bar'), null);
        byId('progressTitle').textContent = `正在核算 ${count} 个上传项`;
        byId('progressText').textContent = '正在读取 PDF、目录内容和图片页二维码…';
      }
    }
    
    function render(data) {
      state.result = data;
      workspace.hidden = false;
      const validTotal = data.grand_total !== null && data.grand_total !== undefined;
      const invoiceCount = data.results.reduce((sum, result) => sum + (result.invoice_count || 0), 0);
      const imageCount = data.results.reduce((sum, result) =>
        sum + (result.routes?.qr?.values?.length || 0), 0);
      const unresolvedCount = data.results.reduce((sum, result) =>
        sum + (result.unresolved_image_pages?.length || 0), 0);
      const overallConfidence = aggregateConfidence(data.results);
    
      animateMoney(byId('grandAmount'), validTotal ? data.grand_total : null);
      updateUppercaseTotal(validTotal ? data.grand_total : null);
      byId('fileStat').textContent = `${data.ok_count}/${data.files}`;
      byId('invoiceStat').textContent = invoiceCount;
      byId('imageStat').textContent = imageCount;
      byId('attentionStat').textContent = data.results.filter(result => !result.ok || result.unresolved_image_pages?.length || !['high', 'medium'].includes(result.confidence)).length;
      byId('resultsCount').textContent = `${data.results.length} 个文件`;
      renderSourceTree(data.sources || []);
      updatePreviewProgress(data.preview);
      byId('summaryStatus').textContent = unresolvedCount
        ? `${unresolvedCount} 个图片页需要处理`
        : imageCount ? `${imageCount} 张图片发票已识别，金额已完整纳入` : '全部金额核算完成';
      const summaryConfidence = byId('summaryConfidence');
      summaryConfidence.textContent = overallConfidence.label;
      summaryConfidence.className = `summary-confidence is-${overallConfidence.level}`;
      root.querySelector('.status-dot').className = `status-dot${unresolvedCount || overallConfidence.level === 'low' ? ' is-warn' : ''}`;
    
      list.innerHTML = '';
      data.results.forEach(result => list.appendChild(resultCard(result)));
      requestAnimationFrame(() => {
        if (state.result === data) context.scrollIntoView(workspace, { behavior: 'smooth', block: 'start' });
      });
    }
    
    async function pollPreviewJob(jobId, taskVersion) {
      state.jobId = jobId;
      while (true) {
        await wait(550);
        if (taskVersion !== state.taskVersion) return;
        const response = await fetch(`/api/jobs/${encodeURIComponent(jobId)}`);
        const snapshot = await response.json();
        if (taskVersion !== state.taskVersion) return;
        if (!response.ok) throw new Error(snapshot.detail || response.statusText);
        updatePreviewProgress(snapshot.preview);
        if (snapshot.status === 'done' || snapshot.status === 'error') {
          if (snapshot.results) render(snapshot);
          return;
        }
      }
    }
    
    function updatePreviewProgress(progress) {
      state.progress = progress;
      if (!progress || progress.status === 'done') {
        previewProgress.hidden = true;
        return;
      }
      previewProgress.hidden = false;
      if (progress.status === 'error') {
        previewProgressTitle.textContent = '票面核对未完成';
        previewProgressText.textContent = progress.message || '金额结果仍可使用，请人工查看原始文件';
        previewProgressCount.textContent = '需人工处理';
        return;
      }
      const totalFiles = Number(progress.total_files || 0);
      const completedFiles = Number(progress.completed_files || 0);
      const totalPages = Number(progress.current_pages || 0);
      const currentPage = Number(progress.current_page || 0);
      const fileFraction = totalPages ? currentPage / totalPages : 0;
      const totalRatio = totalFiles ? Math.min(100, (completedFiles + fileFraction) / totalFiles * 100) : 0;
      const subRatio = totalPages ? Math.min(100, currentPage / totalPages * 100) : 0;
      previewProgressTitle.textContent = progress.status === 'running' ? '金额已完成，正在核对票面' : '正在准备票面核对';
      previewProgressText.textContent = progress.current_file
        ? `正在处理 ${progress.current_file}`
        : '金额已优先完成，正在异步生成票面预览';
      previewProgressCount.textContent = `${completedFiles}/${totalFiles} 个 PDF`;
      context.motion.progress(previewProgressBar, totalRatio);
      previewSubprogressText.textContent = progress.current_file
        ? `当前票面 ${currentPage}/${totalPages} 页`
        : '等待开始';
      context.motion.progress(previewSubprogressBar, subRatio);
    }
    
    function wait(ms) {
      return new Promise(resolve => setTimeout(resolve, ms));
    }
    
    function renderSourceTree(sources) {
      if (!sources.length) {
        sourceTree.hidden = true;
        sourceTree.innerHTML = '';
        return;
      }
      sourceTree.hidden = false;
      sourceTree.innerHTML = `
        <div class="source-tree-head">
          <strong>来源目录</strong>
          <span>确认压缩包内的文件是否都已探查</span>
        </div>
        <div class="source-list">${sources.map(sourceItemHtml).join('')}</div>`;
    }
    
    function sourceItemHtml(source) {
      const entries = source.entries || [];
      const parsed = entries.filter(entry => entry.kind === 'pdf' && entry.result_index !== undefined).length;
      const pdfCount = source.type === 'zip' ? Number(source.pdf_count || 0) : entries.filter(entry => entry.kind === 'pdf').length;
      const summary = source.type === 'zip'
        ? `${parsed}/${pdfCount} 个 PDF 已探查`
        : entries[0]?.ok ? '已探查' : '未完成';
      const rows = entries.length ? entries.map(sourceEntryHtml).join('') :
        '<div class="source-entry is-error"><span>!</span><span class="source-entry-path">未能读取压缩包目录</span></div>';
      return `
        <details class="source-item" open>
          <summary><span class="source-arrow">›</span><span>${source.type === 'zip' ? 'ZIP' : 'PDF'}</span><span>${esc(source.name)}</span><span class="source-badge">${summary}</span></summary>
          <div class="source-entry-list">${rows}</div>
        </details>`;
    }
    
    function sourceEntryHtml(entry) {
      const className = entry.kind === 'directory' ? 'is-directory' : entry.kind === 'ignored' ? 'is-ignored' : '';
      const depth = Math.max(0, String(entry.path || '').split('/').length - 1);
      let status = entry.kind === 'directory' ? '目录' : entry.kind === 'ignored' ? '忽略' : '未探查';
      if (entry.result_index !== undefined) {
        status = entry.ok ? entry.total === null || entry.total === undefined ? '待核对' : fmt(entry.total) : '解析失败';
      }
      const mark = entry.kind === 'directory' ? '▾' : entry.kind === 'ignored' ? '·' : entry.ok === false ? '!' : '□';
      return `<div class="source-entry ${className}" style="--entry-depth:${depth}"><span>${mark}</span><span class="source-entry-path">${esc(entry.path)}</span><span class="source-entry-status">${status}</span></div>`;
    }
    
    function resultCard(result) {
      const card = document.createElement('article');
      card.className = `result-card${result.ok ? '' : ' error-card'}`;
    
      if (!result.ok) {
        card.innerHTML = `
          <div class="result-main">
            <div class="result-top">
              <div>
                <div class="file-line"><span class="file-icon">${fileTypeLabel(result.file)}</span><span class="file-name">${esc(result.file || '未命名文件')}</span><span class="badge badge-error">未完成</span></div>
                <div class="error-title">${esc(result.error || '无法处理此文件')}</div>
              </div>
            </div>
            ${messagesHtml(result)}
          </div>
          ${previewsHtml(result)}`;
        return card;
      }
    
      const routes = routeEntries(result);
      const imageRecognized = (result.routes?.qr?.values?.length || 0) > 0 &&
        !(result.unresolved_image_pages || []).length;
      const imagePageCount = result.image_only_pages?.length || 0;
      const unresolvedImageCount = result.unresolved_image_pages?.length || 0;
      const imageSummary = unresolvedImageCount ? `${unresolvedImageCount} 张待处理` :
        imageRecognized ? `${result.routes.qr.values.length} 张已识别` :
        imagePageCount ? `${imagePageCount} 张图片页` : '无图片页';
      const status = resultStatus(result, imageRecognized);
      const availableCore = ['text', 'table', 'cn', 'items']
        .map(key => result.routes?.[key]?.sum)
        .filter(value => value !== null && value !== undefined);
      const agreement = agreementCount(availableCore);
      const reconcileText = result.reconcile?.available === false ? '不可用' :
        result.reconcile?.partial ? result.reconcile?.passed ? '部分通过' : '部分失败' :
        result.reconcile?.passed ? '已通过' : '建议查看';
      const reconcilePages = result.reconcile?.available
        ? ` · ${Number(result.reconcile.passed_pages?.length || 0)}/${Number(result.reconcile.checked_pages?.length || 0)} 页通过`
        : '';
    
      card.innerHTML = `
        <div class="result-main">
          <div class="result-top">
            <div>
              <div class="file-line">
                <span class="file-icon">PDF</span>
                <span class="file-name">${esc(result.file)}</span>
                <span class="badge ${status.className}">${status.label}</span>
                <span class="badge badge-info">${Number(result.invoice_count || 0)} 张发票</span>
              </div>
              <div class="result-sub">总体置信度：${overallConfidenceLabel(result.confidence)}${result.confidence_desc ? ` · ${esc(result.confidence_desc)}` : ''}</div>
            </div>
            <div class="file-total"><span>文件合计</span><strong>${fmt(result.total)}</strong></div>
          </div>
    
          <div class="check-grid">
            <div class="check-item"><span>文本核算</span><strong>${availableCore.length ? `${agreement}/${availableCore.length} 路一致` : '未使用'}</strong></div>
            <div class="check-item"><span>图片发票</span><strong>${imageSummary}</strong></div>
            <div class="check-item" title="${esc(result.reconcile?.desc || '')}"><span>表格勾稽</span><strong>${reconcileText}${reconcilePages}</strong></div>
          </div>
          ${messagesHtml(result)}
        </div>
    
        <details class="route-details">
          <summary><span>查看完整核算明细</span><span class="detail-arrow">⌄</span></summary>
          <div class="route-list">${routes.map(routeRowHtml).join('')}</div>
        </details>
        ${previewsHtml(result)}`;
      return card;
    }
    
    function resultStatus(result, imageRecognized) {
      if ((result.unresolved_image_pages || []).length || result.confidence === 'low') {
        return { label: '建议核对', className: 'badge-warn' };
      }
      if (result.confidence === 'high') return { label: '高置信度', className: 'badge-success' };
      if (result.confidence === 'medium') return { label: imageRecognized ? '图片已识别 · 中置信度' : '中置信度', className: 'badge-info' };
      return { label: '待判断', className: 'badge-warn' };
    }
    
    function aggregateConfidence(results) {
      const levels = results.map(result => result.ok ? result.confidence : 'low');
      if (!levels.length) return { level: 'low', label: '总体置信度待判断' };
      if (levels.some(level => level === 'low' || level === 'unknown')) return { level: 'low', label: '总体需要核对' };
      if (levels.some(level => level === 'medium')) return { level: 'medium', label: '总体中置信度' };
      return { level: 'high', label: '总体高置信度' };
    }
    
    function overallConfidenceLabel(value) {
      return value === 'high' ? '高' : value === 'medium' ? '中' : value === 'low' ? '低，需要核对' : '待判断';
    }
    
    function fileTypeLabel(fileName) {
      return String(fileName || '').toLowerCase().endsWith('.zip') ? 'ZIP' : 'PDF';
    }
    
    function routeEntries(result) {
      const labels = [
        ['text', '文本合计字段'],
        ['table', '表格结构合计'],
        ['cn', '中文大写金额'],
        ['items', '明细行加总'],
        ['qr', '图片页二维码'],
      ];
      return labels
        .filter(([key]) => result.routes?.[key])
        .map(([key, label]) => ({ key, label, ...result.routes[key] }));
    }
    
    function routeRowHtml(route) {
      const values = route.values || [];
      const firstValues = values.slice(0, 8);
      const restValues = values.slice(8);
      return `
        <div class="route-row">
          <div class="route-name"><strong>${esc(route.label)}</strong><span>${values.length ? `${values.length} 处金额` : '未提取到结果'}</span></div>
          <div class="route-sum">${fmt(route.sum)}</div>
          <div class="value-chips">
            ${firstValues.map(value => `<span class="value-chip">${fmt(value)}</span>`).join('') || '<span class="value-chip">—</span>'}
            ${restValues.length ? `<details class="values-more"><summary>展开其余 ${restValues.length} 个金额</summary><div class="value-chips">${restValues.map(value => `<span class="value-chip">${fmt(value)}</span>`).join('')}</div></details>` : ''}
          </div>
        </div>`;
    }
    
    function messagesHtml(result) {
      const notices = (result.notices || []).map(message =>
        `<div class="message message-info"><span class="message-mark">✓</span><span>${esc(message)}</span></div>`);
      const warnings = (result.warnings || []).map(message =>
        `<div class="message message-warn"><span class="message-mark">!</span><span>${esc(message)}</span></div>`);
      return [...notices, ...warnings].join('');
    }
    
    function previewsHtml(result) {
      const previews = result.previews || [];
      const pageAmounts = result.page_amounts || [];
      if (!previews.length) {
        if (result.preview_status === 'pending' || result.preview_status === 'running') {
          return '<div class="preview-section"><div class="preview-heading">票面核对</div><div class="message message-info"><span class="message-mark">…</span><span>金额已优先完成，票面缩略图正在后台生成。</span></div></div>';
        }
        return '';
      }
      const cards = previews.map(preview => {
        const pageResult = pageAmounts.find(item => Number(item.page) === Number(preview.page));
        const pageAmount = pageResult?.amount;
        const pageConfidence = pageResult?.confidence && pageResult.confidence !== 'unknown' ? ` · ${confidenceLabel(pageResult.confidence)}` : '';
        const amountText = pageAmount === null || pageAmount === undefined ? '待人工核对' : fmt(pageAmount);
        return `
        <button class="preview-open" type="button" data-title="${esc(result.file)} · 第 ${Number(preview.page)} 页 · ${esc(amountText)}">
          <img class="preview-thumb" src="${preview.data_url}" alt="第 ${Number(preview.page)} 页发票缩略图" loading="lazy">
          <span class="preview-copy"><strong>第 ${Number(preview.page)} 页票面</strong><span>点击查看完整票面</span><em class="preview-amount">票面金额 ${amountText}${pageConfidence}</em></span>
          <span class="preview-zoom" aria-hidden="true">⌕</span>
        </button>`;
      }).join('');
      const truncated = result.preview_truncated
        ? '<div class="message message-info"><span>页数较多，此处显示前 100 页。</span></div>' : '';
      return `<div class="preview-section"><div class="preview-heading">票面核对 · 缩略图与识别合计</div><div class="preview-grid">${cards}</div>${truncated}</div>`;
    }
    
    function confidenceLabel(value) {
      return value === 'high' ? '多路一致' : value === 'medium' ? '两路一致' : value === 'single' ? '单路识别' : value === 'low' ? '需核对' : '待核对';
    }
    
    function agreementCount(values) {
      if (!values.length) return 0;
      const counts = new Map();
      for (const value of values) {
        const key = Number(value).toFixed(2);
        counts.set(key, (counts.get(key) || 0) + 1);
      }
      return Math.max(...counts.values());
    }
    
    function animateMoney(element, target) {
      element.textContent = target === null ? '—' : fmt(target);
    }
    
    function renderUploadError(error) {
      workspace.hidden = false;
      sourceTree.hidden = true;
      sourceTree.innerHTML = '';
      previewProgress.hidden = true;
      byId('grandAmount').textContent = '—';
      updateUppercaseTotal(null);
      byId('fileStat').textContent = '0';
      byId('invoiceStat').textContent = '0';
      byId('imageStat').textContent = '0';
      byId('attentionStat').textContent = '1';
      byId('summaryStatus').textContent = '上传未完成';
      byId('summaryConfidence').textContent = '总体置信度待判断';
      byId('summaryConfidence').className = 'summary-confidence is-low';
      root.querySelector('.status-dot').className = 'status-dot is-error';
      byId('resultsCount').textContent = '1 个问题';
      list.innerHTML = `<article class="result-card error-card"><div class="result-main"><div class="file-line"><span class="file-icon">!</span><span class="file-name">上传失败</span><span class="badge badge-error">未完成</span></div><div class="error-title">${esc(error.message || '网络请求失败')}</div></div></article>`;
    }
    
    function resetResults() {
      state.taskVersion += 1;
      setLoading(false);
      state.files = []; state.result = null; state.jobId = null;
      context.downloads.clear();
      workspace.hidden = true;
      sourceTree.hidden = true;
      sourceTree.innerHTML = '';
      previewProgress.hidden = true;
      list.innerHTML = '';
      updateUppercaseTotal(null);
      context.scrollTo({ top: 0, behavior: 'smooth' });
    }
    
    function openViewer(src, title) {
      viewerImage.src = src;
      viewerOpen.href = src;
      viewerTitle.textContent = title;
      for (let parent = viewer.parentElement; parent; parent = parent.parentElement) context.motion?.cancel(parent);
      viewer.inert = false;
      viewer.classList.add('is-open');
      viewer.setAttribute('aria-hidden', 'false');
      context.motion.modal(viewer, true);
    }
    
    function closeViewer() {
      context.motion.modal(viewer, false);
      viewer.inert = true;
      viewer.classList.remove('is-open');
      viewer.setAttribute('aria-hidden', 'true');
      context.setScrollLock(false);
      // Retain the preview during its closing transition; the next open replaces it.
    }
    
    function fmt(value) {
      if (value === null || value === undefined) return '—';
      return '¥' + Number(value).toLocaleString('zh-CN', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      });
    }
    
    function updateUppercaseTotal(value) {
      const hasAmount = value !== null && value !== undefined && Number.isFinite(Number(value));
      if (hasAmount) {
        const uppercaseText = moneyToChineseUppercase(value).replace(/^人民币/, '');
        uppercaseAmount.innerHTML = [...uppercaseText].map(character =>
          '零壹贰叁肆伍陆柒捌玖'.includes(character)
            ? `<span class="amount-digit">${character}</span>`
            : character
        ).join('');
      } else {
        uppercaseAmount.textContent = '—';
      }
      copyUppercase.disabled = !hasAmount;
      copyUppercase.classList.remove('is-copied');
      copyUppercase.querySelector('span').textContent = '复制大写金额';
      copyStatus.textContent = '';
    }
    
    function moneyToChineseUppercase(value) {
      const amount = Number(value);
      const maxAmount = 9999999999999.99;
      if (!Number.isFinite(amount) || Math.abs(amount) > maxAmount) return '金额超出转换范围';
      if (amount === 0) return '人民币零元整';
    
      const digits = '零壹贰叁肆伍陆柒捌玖';
      const groupUnits = ['', '万', '亿', '万亿'];
      const totalCents = Math.round(Math.abs(amount) * 100);
      const integer = Math.floor(totalCents / 100);
      const jiao = Math.floor(totalCents / 10) % 10;
      const fen = totalCents % 10;
    
      function convertGroup(group) {
        const units = ['仟', '佰', '拾', ''];
        const divisors = [1000, 100, 10, 1];
        let text = '';
        let needsZero = false;
        divisors.forEach((divisor, index) => {
          const digit = Math.floor(group / divisor) % 10;
          if (digit) {
            if (needsZero && text) text += digits[0];
            text += digits[digit] + units[index];
            needsZero = false;
          } else if (text) {
            needsZero = true;
          }
        });
        return text;
      }
    
      function convertInteger(number) {
        if (!number) return digits[0];
        const groups = [];
        while (number > 0) {
          groups.push(number % 10000);
          number = Math.floor(number / 10000);
        }
    
        let text = '';
        let skippedGroup = false;
        for (let index = groups.length - 1; index >= 0; index -= 1) {
          const group = groups[index];
          if (!group) {
            if (text) skippedGroup = true;
            continue;
          }
          if (text && (skippedGroup || group < 1000) && !text.endsWith(digits[0])) text += digits[0];
          text += convertGroup(group) + groupUnits[index];
          skippedGroup = false;
        }
        return text;
      }
    
      let result = `人民币${amount < 0 ? '负' : ''}${convertInteger(integer)}元`;
      if (!jiao && !fen) return result + '整';
      if (jiao) result += digits[jiao] + '角';
      if (fen || jiao) result += (!jiao ? digits[0] : '') + digits[fen] + '分';
      return result;
    }
    
    async function copyUppercaseAmount() {
      const activation = context.activation;
      const text = uppercaseAmount.textContent;
      try {
        if (navigator.clipboard && window.isSecureContext) {
          await navigator.clipboard.writeText(text);
        } else {
          const textarea = document.createElement('textarea');
          textarea.value = text;
          textarea.setAttribute('readonly', '');
          textarea.style.position = 'fixed';
          textarea.style.opacity = '0';
          document.body.appendChild(textarea);
          textarea.select();
          const copied = document.execCommand('copy');
          textarea.remove();
          if (!copied) throw new Error('copy failed');
        }
        if (!context.isCurrentActivation(activation)) return;
        copyUppercase.classList.add('is-copied');
        copyUppercase.querySelector('span').textContent = '已复制';
        copyStatus.textContent = '大写金额已复制到剪贴板';
        clearTimeout(copyFeedbackTimer);
        copyFeedbackTimer = setTimeout(() => {
          copyUppercase.classList.remove('is-copied');
          copyUppercase.querySelector('span').textContent = '复制大写金额';
        }, 1800);
      } catch (error) {
        copyStatus.textContent = '复制失败，请手动选择大写金额';
        copyUppercase.querySelector('span').textContent = '复制失败';
      }
    }
    
    function esc(value) {
      return String(value ?? '').replace(/[&<>"']/g, char => ({
        '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
      })[char]);
    }
    
    context.onDeactivate(() => {
      clearTimeout(copyFeedbackTimer);
      closeViewer();
      copyUppercase.classList.remove('is-copied');
      copyUppercase.querySelector('span').textContent = '复制大写金额';
    });
  }
  return context.lifecycle(mount);
}
