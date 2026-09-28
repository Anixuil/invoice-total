// Each factory owns its DOM, files and task state for one route instance.
export function createPage(context) {
  const { root, state } = context;
  const byId = id => root.querySelector(`#${CSS.escape(id)}`);
  const { fetch, XMLHttpRequest, setTimeout, clearTimeout } = context;
  function mount() {
    const drop = byId('drop');
    const fileInput = byId('file');
    const zipFileInput = byId('zipFile');
    const directoryInput = byId('directoryFiles');
    const newTaskExcel = byId('newTaskExcel');
    const newTaskScreenshot = byId('newTaskScreenshot');
    const clipboardPasteTarget = byId('clipboardPasteTarget');
    const weeklyUpload = byId('weeklyUpload');
    const weeklyNewTaskToolbar = byId('weeklyNewTaskToolbar');
    const progress = byId('progress');
    const workspace = byId('workspace');
    const downloadButton = byId('downloadButton');
    const jiraMode = new URLSearchParams(context.route.url.search).get('mode') === 'daily' ? 'daily' : 'weekly';
    const isDaily = jiraMode === 'daily';
    let currentJobId = '';
    let currentResult = null;
    let newTaskMode = false;
    let newTaskSource = null;
    let newTaskScreenshotFile = null;
    let activeWeeklyStat = 'new-tasks';
    let pasteTargetStat = '';
    const weeklyStatisticInputs = {};
    const weeklyStatFiles = {};
    const weeklyStatImages = {};
    const weeklyStatResults = {};
    const weeklyStatJobs = {};
    const weeklyStatState = {};
    const WEEKLY_STAT_POLL_TIMEOUT_MS = 300000;
    const WEEKLY_STAT_WATCHDOG_MS = WEEKLY_STAT_POLL_TIMEOUT_MS + 15000;
    const weeklyStatFileNames = {
      'new-tasks': '本周新增任务数.xlsx', 'completed-tasks': '本周完成任务数.xlsx',
      'new-defects': '本周新增缺陷数.xlsx', 'fixed-defects': '本周已修复缺陷数.xlsx',
      'delayed-defects': '本周延期缺陷数.xlsx', 'pending-defects': '总挂起缺陷数.xlsx',
      'delayed-tasks': '本周延期任务数.xlsx', 'pending-tasks': '总挂起任务数.xlsx'
    };
    let generatedStatKey = '';
    let previewObjectUrl = '';
    let detailRecords = [];
    let detailFilter = { query: '', status: 'all', included: 'all' };
    
    function setWeeklyActionState(button, label, state = 'ready') {
      if (!button) return;
      if (!button.dataset.defaultLabel) button.dataset.defaultLabel = button.textContent;
      if (button._scanActive && state !== 'processing') {
        button._scanLabel = label;
        return;
      }
      button.textContent = label;
      button.classList.remove('is-ready', 'is-processing');
      button.classList.add(state === 'processing' ? 'is-processing' : 'is-ready');
    }
    
    function scanWeeklyUpload(button, completeLabel, duration = 800) {
      if (!button) return;
      cancelWeeklyScan(button, true);
      const scanId = (button._scanId || 0) + 1;
      button._scanId = scanId;
      button._scanActive = true;
      button._scanLabel = completeLabel;
      let percent = 8;
      button.textContent = `扫描中 ${percent}%`;
      button.classList.remove('is-ready');
      button.classList.add('is-processing');
      button._scanTimer = context.visualInterval(() => {
        if (button._scanId !== scanId) return;
        percent = Math.min(96, percent + 11);
        button.textContent = `扫描中 ${percent}%`;
      }, 110);
      button._scanFinishTimer = setTimeout(() => {
        if (button._scanId !== scanId) return;
        context.clearVisualInterval(button._scanTimer);
        button._scanTimer = null;
        button._scanActive = false;
        setWeeklyActionState(button, button._scanLabel || completeLabel);
      }, duration);
    }
    
    function scanWeeklyGeneration(button) {
      if (!button) return;
      if (!button.dataset.defaultLabel) button.dataset.defaultLabel = button.textContent;
      button.classList.remove('is-awaiting-paste');
      cancelWeeklyScan(button, true);
      const scanId = (button._scanId || 0) + 1;
      button._scanId = scanId;
      button._scanActive = true;
      button._scanLabel = '';
      button.textContent = '正在提交任务';
      button.classList.remove('is-ready');
      button.classList.add('is-processing');
      clearWeeklyJobProgress(button);
      updateWeeklyJobProgress(button, { stage: '正在提交任务', percent: 5, detail: '正在上传本项所需文件' });
    }
    
    function scanWeeklyScreenshot(button) {
      if (!button) return;
      if (!button.dataset.defaultLabel) button.dataset.defaultLabel = button.textContent;
      cancelWeeklyScan(button, true);
      const scanId = (button._scanId || 0) + 1;
      button._scanId = scanId;
      button._scanActive = true;
      button._scanLabel = '';
      let percent = 8;
      button.textContent = `校验截图 ${percent}%`;
      button.classList.remove('is-ready');
      button.classList.add('is-processing');
      button._scanTimer = context.visualInterval(() => {
        if (button._scanId !== scanId) return;
        percent = Math.min(94, percent + 6);
        button.textContent = `校验截图 ${percent}%`;
      }, 300);
      button._scanWatchdog = setTimeout(() => {
        if (button._scanId !== scanId || !button._scanActive) return;
        cancelWeeklyScan(button);
        const statKey = button.id === 'newTaskPasteButton' ? 'new-tasks' : button.dataset.statPaste || button.dataset.statUpload;
        if (statKey) showWeeklyInlineError(statKey, '截图校验超时，请重新 Ctrl+V 粘贴后重试');
      }, 90000);
    }
    
    function completeWeeklyGeneration(button, label) {
      if (!button) return;
      if (button._scanTimer) context.clearVisualInterval(button._scanTimer);
      if (button._scanWatchdog) clearTimeout(button._scanWatchdog);
      if (button._scanFinishTimer) clearTimeout(button._scanFinishTimer);
      button._scanTimer = null;
      button._scanWatchdog = null;
      button._scanFinishTimer = null;
      button._scanActive = false;
      updateWeeklyJobProgress(button, { stage: '生成完成', percent: 100, detail: '文件已生成，可点击下载' }, true);
      setWeeklyActionState(button, label);
    }
    
    function cancelWeeklyScan(button, preserveLabel = false) {
      if (!button) return;
      if (button._scanTimer) context.clearVisualInterval(button._scanTimer);
      if (button._scanFinishTimer) clearTimeout(button._scanFinishTimer);
      if (button._scanWatchdog) clearTimeout(button._scanWatchdog);
      button._scanId = (button._scanId || 0) + 1;
      button._scanTimer = null;
      button._scanFinishTimer = null;
      button._scanWatchdog = null;
      button._scanActive = false;
      button._scanLabel = '';
      if (!preserveLabel) {
        button.textContent = button.dataset.defaultLabel || '重新上传';
        button.classList.remove('is-ready', 'is-processing', 'is-awaiting-paste');
      }
    }
    
    function clearWeeklyInput(button, statKey) {
      const inputKey = statKey === 'new-tasks' ? (button.id === 'newTaskExcelButton' ? 'new-tasks-file' : 'new-tasks-image') : statKey;
      const isWorkbook = button.id === 'newTaskExcelButton' || button.dataset.statUpload === 'completed-tasks';
      if (statKey !== 'completed-tasks') delete weeklyStatisticInputs[inputKey];
      if (isWorkbook) delete weeklyStatFiles[statKey];
      else delete weeklyStatImages[statKey];
      if (weeklyStatState[statKey]) {
        weeklyStatState[statKey][isWorkbook ? 'excel' : 'screenshot'] = null;
      }
      if (pasteTargetStat === statKey) pasteTargetStat = '';
      if (statKey === 'new-tasks') {
        if (button.id === 'newTaskExcelButton') newTaskSource = null;
        if (button.id === 'newTaskPasteButton') newTaskScreenshotFile = null;
      }
      button._weeklyUploadFile = null;
      button._weeklyInputVersion = (button._weeklyInputVersion || 0) + 1;
      cancelWeeklyScan(button);
      clearWeeklyJobProgress(button);
      button.textContent = button.id === 'newTaskExcelButton' || button.dataset.statUpload === 'completed-tasks'
        ? '上传文件'
        : '粘贴截图';
      button.classList.remove('is-ready', 'is-processing', 'is-awaiting-paste');
      if (isWorkbook) newTaskExcel.value = '';
      else newTaskScreenshot.value = '';
      byId('newTaskStatus').textContent = '已删除，可重新上传';
    }
    
    function isClearClick(event, button) {
      if (button.id === 'weeklySummaryButton') return false;
      const rect = button.getBoundingClientRect();
      const inClearArea = event.clientX >= rect.right - 80 || event.offsetX >= button.clientWidth - 80;
      return (button.classList.contains('is-ready') || button.classList.contains('is-processing')) && inClearArea;
    }
    
    function saveWeeklyStatInput(statKey, inputType, file) {
      if (!weeklyStatState[statKey]) weeklyStatState[statKey] = { excel: null, screenshot: null };
      weeklyStatState[statKey][inputType] = file;
    }
    
    function getWeeklyStatInput(statKey, inputType) {
      return weeklyStatState[statKey]?.[inputType] || null;
    }
    
    function retainPastedScreenshot(file) {
      try {
        const transfer = new DataTransfer();
        transfer.items.add(file);
        newTaskScreenshot.files = transfer.files;
      } catch (_) {
        // The per-row state remains the fallback for browsers that disallow it.
      }
    }
    
    function clearGeneratedAction(button, statKey = '') {
      const oldJobId = weeklyStatJobs[statKey] || button?.dataset.weeklyJobId;
      if (oldJobId) context.downloads.clear(`/api/jira/weekly-new-tasks/jobs/${encodeURIComponent(oldJobId)}/export`);
      if (statKey) delete weeklyStatResults[statKey];
      if (statKey) delete weeklyStatJobs[statKey];
      if (button) delete button.dataset.weeklyJobId;
      if (generatedStatKey === statKey) { currentJobId = ''; currentResult = null; generatedStatKey = ''; }
      button._weeklyInputVersion = (button._weeklyInputVersion || 0) + 1;
      cancelWeeklyScan(button);
      clearWeeklyJobProgress(button);
      button.textContent = button.dataset.defaultLabel || '生成文件';
      button.classList.remove('is-ready', 'is-processing', 'is-awaiting-paste');
      byId('newTaskStatus').textContent = '结果已删除，可重新生成';
    }
    
    function screenshotActionButton(statKey) {
      if (statKey === 'new-tasks') return byId('newTaskPasteButton');
      return root.querySelector(`[data-stat-paste="${statKey}"]`) || root.querySelector(`[data-stat-upload="${statKey}"]`);
    }
    
    function updateWeeklyGenerationProgress(button, progress, completed = false) {
      if (!button) return;
      if (!button._scanActive && !completed) return;
      const percent = completed ? 100 : progress?.percent;
      const stage = progress?.stage || '正在处理';
      if (completed) {
        const statKey = button.id === 'newTaskGenerateButton' ? 'new-tasks' : button.dataset.statGenerate;
        const filename = statKey && weeklyStatJobs[statKey]
          ? weeklyStatFileNames[statKey] || '统计结果.xlsx'
          : stage;
        completeWeeklyGeneration(button, filename);
        return;
      }
      button.textContent = stage;
      updateWeeklyJobProgress(button, { ...progress, stage, percent });
    }
    
    function updateWeeklyJobProgress(button, progress, completed = false) {
      const row = button?.closest('.weekly-stat-row');
      if (!row) return;
      let panel = row.querySelector('.weekly-job-progress');
      if (!panel) {
        panel = document.createElement('div');
        panel.className = 'weekly-job-progress';
        panel.setAttribute('role', 'status');
        panel.setAttribute('aria-live', 'polite');
        panel.innerHTML = '<span class="weekly-job-progress__mark"></span><div class="weekly-job-progress__body"><span class="weekly-job-progress__stage"></span><span class="weekly-job-progress__detail"></span></div><span class="weekly-job-progress__percent"></span><div class="weekly-job-progress__track"><div class="weekly-job-progress__bar"></div></div>';
        row.appendChild(panel);
      }
      const percent = context.motion.progress(panel.querySelector('.weekly-job-progress__bar'), progress?.percent);
      panel.classList.toggle('is-complete', completed);
      panel.classList.remove('is-error');
      panel.querySelector('.weekly-job-progress__mark').textContent = completed ? '✓' : '…';
      panel.querySelector('.weekly-job-progress__stage').textContent = progress?.stage || '正在处理';
      panel.querySelector('.weekly-job-progress__detail').textContent = progress?.detail || '正在等待服务端返回处理进度';
      panel.querySelector('.weekly-job-progress__percent').textContent = percent === null ? '处理中' : `${percent}%`;
      byId('newTaskStatus').textContent = `${progress?.stage || '正在处理'}：${progress?.detail || '正在等待服务端返回处理进度'}`;
    }
    
    function clearWeeklyJobProgress(button) {
      button?.closest('.weekly-stat-row')?.querySelector('.weekly-job-progress')?.remove();
    }
    
    function showWeeklyJobFailure(button, message) {
      updateWeeklyJobProgress(button, { stage: '生成失败', percent: 100, detail: message });
      const panel = button?.closest('.weekly-stat-row')?.querySelector('.weekly-job-progress');
      panel?.classList.add('is-error');
      if (panel) panel.querySelector('.weekly-job-progress__mark').textContent = '!';
    }
    
    function statGenerateButton(statKey) {
      return statKey === 'new-tasks' ? byId('newTaskGenerateButton') : root.querySelector(`[data-stat-generate="${statKey}"]`);
    }
    
    function showWeeklyInlineError(statKey, message) {
      const row = statGenerateButton(statKey)?.closest('.weekly-stat-row');
      if (!row) return;
      let target = row.querySelector('.weekly-inline-error');
      if (!target) { target = document.createElement('div'); target.className = 'weekly-inline-error'; row.appendChild(target); }
      target.textContent = message;
    }
    
    function clearWeeklyInlineError(statKey) {
      statGenerateButton(statKey)?.closest('.weekly-stat-row')?.querySelector('.weekly-inline-error')?.remove();
    }
    
    function closePreview() {
      state.previewVersion += 1;
      const dialog = byId('previewDialog');
      if (dialog.contains(document.activeElement)) document.activeElement.blur();
      if (context.motion) context.motion.dialog(dialog, false);
      else dialog.hidden = true;
      byId('previewBody').innerHTML = '';
      if (previewObjectUrl) { context.revokeObjectURL(previewObjectUrl); previewObjectUrl = ''; }
    }
    
    function previewScreenshot(file) {
      closePreview();
      if (!file) return;
      previewObjectUrl = context.createObjectURL(file);
      byId('previewTitle').textContent = file.name || '粘贴截图预览';
      byId('previewBody').innerHTML = `<img class="preview-dialog__image" src="${previewObjectUrl}" alt="已上传截图预览">`;
      const dialog = byId('previewDialog');
      if (context.motion) context.motion.dialog(dialog, true);
      else dialog.hidden = false;
    }
    
    async function previewWorkbook(file) {
      closePreview();
      const previewVersion = ++state.previewVersion;
      if (!file) return;
      byId('previewTitle').textContent = file.name;
      byId('previewBody').textContent = '正在读取 Excel 预览...';
      const dialog = byId('previewDialog');
      if (context.motion) context.motion.dialog(dialog, true);
      else dialog.hidden = false;
      const body = new FormData(); body.append('file', file);
      try {
        const response = await fetch('/api/jira/workbook-preview', { method: 'POST', body });
        const data = await response.json();
        if (!response.ok) throw new Error(data.detail || response.statusText);
        if (previewVersion !== state.previewVersion || !context.active) return;
        byId('previewBody').innerHTML = `<p class="upload-copy">${esc(data.sheet)} · 共 ${data.total_rows} 条记录，显示前 ${data.rows.length} 条</p><div class="table-scroll"><table><thead><tr>${data.headers.map(header => `<th>${esc(header)}</th>`).join('')}</tr></thead><tbody>${data.rows.map(row => `<tr>${row.map(value => `<td>${esc(value)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
      } catch (error) { if (previewVersion !== state.previewVersion || !context.active) return; byId('previewBody').textContent = error.message || '无法读取 Excel 预览'; }
    }
    
    byId('previewCloseButton').addEventListener('click', closePreview);
    byId('previewDialog').addEventListener('click', event => { if (event.target === event.currentTarget) closePreview(); });
    
    function clipboardImageFile(event) {
      const clipboard = event.clipboardData;
      const file = Array.from(clipboard?.files || []).find(item => item.type.startsWith('image/'));
      if (file) return file;
      const item = Array.from(clipboard?.items || []).find(item => item.type.startsWith('image/'));
      return item?.getAsFile() || null;
    }
    
    // Lock the target row before any of the legacy button handlers run. The
    // original page reused one hidden file input for every row, which could leave
    // a paste validation pointed at the first statistic after changing rows.
    context.listenGlobal(document, 'click', event => {
      const button = event.target.closest('#newTaskExcelButton, #newTaskPasteButton, #newTaskGenerateButton, [data-stat-upload], [data-stat-paste], [data-stat-generate]');
      if (!button) return;
      const isGenerateButton = button.id === 'newTaskGenerateButton' || button.dataset.statGenerate;
      const hasFailedGeneration = Boolean(button.closest('.weekly-stat-row')?.querySelector('.weekly-job-progress.is-error'));
      // Clear clicks must be consumed before any of the legacy target handlers can
      // interpret the same button as a preview or download action.
      if (isClearClick(event, button)) {
        event.preventDefault();
        if (button.id === 'newTaskExcelButton' || button.id === 'newTaskPasteButton') clearWeeklyInput(button, 'new-tasks');
        else if (button.id === 'newTaskGenerateButton') clearGeneratedAction(button, 'new-tasks');
        else if (button.dataset.statUpload) clearWeeklyInput(button, button.dataset.statUpload);
        else if (button.dataset.statPaste) clearWeeklyInput(button, button.dataset.statPaste);
        else if (button.dataset.statGenerate) clearGeneratedAction(button, button.dataset.statGenerate);
        event.stopImmediatePropagation();
        return;
      }
      if (isGenerateButton && button.classList.contains('is-processing') && !hasFailedGeneration) {
        event.preventDefault();
        event.stopImmediatePropagation();
        return;
      }
      if (isGenerateButton) {
        // Handle generation here with DOM file fallbacks. The page historically
        // had several generation listeners; this capture path is the single
        // authoritative entry point and prevents a stale listener from winning.
        const statKey = button.id === 'newTaskGenerateButton' ? 'new-tasks' : button.dataset.statGenerate;
        const needsWorkbook = statKey === 'new-tasks' || statKey === 'completed-tasks';
        const screenshotButton = screenshotActionButton(statKey);
        const excelButton = statKey === 'new-tasks' ? byId('newTaskExcelButton') : root.querySelector(`[data-stat-upload="${statKey}"]`);
        const screenshot = screenshotButton?._weeklyUploadFile || getWeeklyStatInput(statKey, 'screenshot') || newTaskScreenshot.files[0] || weeklyStatImages[statKey];
        const excel = excelButton?._weeklyUploadFile || getWeeklyStatInput(statKey, 'excel') || newTaskExcel.files[0] || weeklyStatFiles[statKey];
        event.preventDefault();
        event.stopImmediatePropagation();
        if (weeklyStatJobs[statKey]) { openWeeklyStatFile(statKey); return; }
        clearWeeklyInlineError(statKey);
        if (!screenshot || (needsWorkbook && !excel)) {
          showWeeklyInlineError(statKey, needsWorkbook ? '请先上传 Excel 并粘贴截图' : '请先粘贴截图');
          return;
        }
        scanWeeklyGeneration(button);
        if (needsWorkbook) void uploadNewTask(excel, screenshot, statKey, button);
        else void uploadScreenshotStat(statKey, screenshot, button);
        return;
        /*
        const statKey = button.id === 'newTaskGenerateButton' ? 'new-tasks' : button.dataset.statGenerate;
        const needsWorkbook = statKey === 'new-tasks' || statKey === 'completed-tasks';
        const screenshotButton = screenshotActionButton(statKey);
        const excelButton = statKey === 'new-tasks'
          ? byId('newTaskExcelButton')
          : root.querySelector(`[data-stat-upload="${statKey}"]`);
        const screenshot = screenshotButton?._weeklyUploadFile || getWeeklyStatInput(statKey, 'screenshot')
          || (statKey === 'new-tasks' ? weeklyStatisticInputs['new-tasks-image'] || newTaskScreenshotFile : weeklyStatImages[statKey]);
        const excel = excelButton?._weeklyUploadFile || getWeeklyStatInput(statKey, 'excel')
          || (statKey === 'new-tasks' ? weeklyStatisticInputs['new-tasks-file'] || newTaskSource : weeklyStatFiles[statKey]);
    
        event.preventDefault();
        event.stopImmediatePropagation();
        if (weeklyStatJobs[statKey]) {
          openWeeklyStatFile(statKey);
          return;
        }
        clearWeeklyInlineError(statKey);
        if (!screenshot || (needsWorkbook && !excel)) {
          showWeeklyInlineError(statKey, needsWorkbook ? '请先上传 Excel 并粘贴截图' : '请先粘贴或选择截图');
          return;
        }
        scanWeeklyGeneration(button);
        if (needsWorkbook) uploadNewTask(excel, screenshot, statKey, button);
        else uploadScreenshotStat(statKey, screenshot, button);
        return;
        */
      }
      const isScreenshotButton = button.id === 'newTaskPasteButton' || button.dataset.statPaste || (button.dataset.statUpload && button.dataset.statUpload !== 'completed-tasks');
      // Screenshot controls only accept Ctrl+V. Handle the row selection here so
      // the legacy handlers cannot switch the paste target to a different row.
      if (isScreenshotButton && button.classList.contains('is-ready')) return;
      if (isScreenshotButton) {
        event.preventDefault();
        event.stopImmediatePropagation();
        newTaskMode = true;
        activeWeeklyStat = button.id === 'newTaskPasteButton'
          ? 'new-tasks'
          : (button.dataset.statPaste || button.dataset.statUpload);
        pasteTargetStat = activeWeeklyStat;
        button.textContent = '请按 Ctrl+V 粘贴';
        button.classList.add('is-awaiting-paste');
        clipboardPasteTarget.textContent = '';
        context.focus(clipboardPasteTarget, { preventScroll: true });
        byId('newTaskStatus').textContent = '请直接按 Ctrl+V 粘贴截图';
        return;
      }
      if (button.id === 'newTaskExcelButton') {
        activeWeeklyStat = 'new-tasks';
        pasteTargetStat = '';
      } else if (button.id === 'newTaskPasteButton') {
        activeWeeklyStat = 'new-tasks';
        pasteTargetStat = 'new-tasks';
        button.textContent = 'Ctrl+V 或再次点击选择文件';
        button.classList.add('is-awaiting-paste');
      } else if (button.dataset.statUpload) {
        activeWeeklyStat = button.dataset.statUpload;
        pasteTargetStat = activeWeeklyStat === 'completed-tasks' ? '' : activeWeeklyStat;
        if (activeWeeklyStat !== 'completed-tasks') {
          button.textContent = 'Ctrl+V 或再次点击选择文件';
          button.classList.add('is-awaiting-paste');
        }
      } else if (button.dataset.statPaste) {
        activeWeeklyStat = button.dataset.statPaste;
        pasteTargetStat = activeWeeklyStat;
        button.textContent = '请按 Ctrl+V 粘贴';
        button.classList.add('is-awaiting-paste');
      }
      newTaskMode = true;
    }, true);
    
    root.querySelectorAll('.weekly-action-button').forEach(button => button.addEventListener('click', event => {
      if (button.id === 'weeklySummaryButton') return;
      if (!isClearClick(event, button)) return;
      if (button.id === 'newTaskExcelButton') clearWeeklyInput(button, 'new-tasks');
      else if (button.id === 'newTaskPasteButton') clearWeeklyInput(button, 'new-tasks');
      else if (button.id === 'newTaskGenerateButton') clearGeneratedAction(button, 'new-tasks');
      else if (button.dataset.statUpload) clearWeeklyInput(button, button.dataset.statUpload);
      else if (button.dataset.statPaste) clearWeeklyInput(button, button.dataset.statPaste);
      else if (button.dataset.statGenerate) clearGeneratedAction(button, button.dataset.statGenerate);
      event.stopImmediatePropagation();
    }, true));
    
    root.querySelectorAll('.weekly-action-button').forEach(button => button.addEventListener('click', event => {
      if (!button.classList.contains('is-ready') || isClearClick(event, button)) return;
      const statKey = button.dataset.statUpload || button.dataset.statPaste;
      if (button.id === 'newTaskExcelButton') previewWorkbook(button._weeklyUploadFile || newTaskSource || weeklyStatisticInputs['new-tasks-file'] || newTaskExcel.files[0]);
      else if (button.id === 'newTaskPasteButton') previewScreenshot(button._weeklyUploadFile || newTaskScreenshotFile || weeklyStatisticInputs['new-tasks-image'] || newTaskScreenshot.files[0]);
      else if (button.dataset.statUpload && statKey === 'completed-tasks') previewWorkbook(button._weeklyUploadFile || weeklyStatFiles[statKey] || getWeeklyStatInput(statKey, 'excel') || newTaskExcel.files[0]);
      else if (statKey) previewScreenshot(button._weeklyUploadFile || weeklyStatImages[statKey] || getWeeklyStatInput(statKey, 'screenshot') || newTaskScreenshot.files[0]);
      else return;
      event.stopImmediatePropagation();
    }, true));
    
    function openWeeklyStatFile(statKey) {
      const jobId = statGenerateButton(statKey)?.dataset.weeklyJobId || weeklyStatJobs[statKey];
      downloadWeeklyStatFile(jobId);
    }
    
    function downloadWeeklyStatFile(jobId) {
      if (jobId) context.downloads.request({ url: `/api/jira/weekly-new-tasks/jobs/${encodeURIComponent(jobId)}/export` });
    }
    
    byId('newTaskGenerateButton').addEventListener('click', event => {
      if (isClearClick(event, event.currentTarget) || !weeklyStatJobs['new-tasks']) return;
      event.stopImmediatePropagation();
      openWeeklyStatFile('new-tasks');
    }, true);
    root.querySelectorAll('[data-stat-generate]').forEach(button => button.addEventListener('click', event => {
      const statKey = button.dataset.statGenerate;
      if (isClearClick(event, button) || !weeklyStatJobs[statKey]) return;
      event.stopImmediatePropagation();
      openWeeklyStatFile(statKey);
    }, true));
    
    byId('newTaskPasteButton').addEventListener('click', event => {
      if (isClearClick(event, event.currentTarget)) return;
      newTaskMode = true; activeWeeklyStat = 'new-tasks'; pasteTargetStat = 'new-tasks';
      byId('newTaskStatus').textContent = '请直接按 Ctrl+V 粘贴截图';
      event.stopImmediatePropagation();
    }, true);
    root.querySelectorAll('[data-stat-upload], [data-stat-paste]').forEach(button => button.addEventListener('click', event => {
      const statKey = button.dataset.statUpload || button.dataset.statPaste;
      if (statKey === 'completed-tasks' && button.dataset.statUpload) return;
      if (isClearClick(event, button)) return;
      newTaskMode = true; activeWeeklyStat = statKey; pasteTargetStat = statKey;
      byId('newTaskStatus').textContent = '请直接按 Ctrl+V 粘贴截图';
      event.stopImmediatePropagation();
    }, true));
    
    byId('weeklyModeLink').classList.toggle('is-active', !isDaily);
    byId('dailyModeLink').classList.toggle('is-active', isDaily);
    root.querySelector('.mode-switch').dataset.mode = jiraMode;
    byId('weeklyModeLink').setAttribute('aria-selected', String(!isDaily));
    byId('dailyModeLink').setAttribute('aria-selected', String(isDaily));
    if (isDaily) {
      context.title = '每日 Jira 处理 · 本地工具';
      byId('heroEyebrow').textContent = '每日 Jira 数据工具';
      byId('heroTitle').textContent = '每日 Jira 处理';
      byId('heroCopy').textContent = '按开发人员、经办人、报告人的优先级归属任务，转换状态后合并每日工作内容。';
      byId('uploadHint').textContent = '开发人员为空时使用经办人，经办人也为空时使用报告人，并在明细中标注回退原因。';
      byId('pendingLabel').textContent = '开发人员缺失';
      byId('mismatchLabel').textContent = '人员不同';
      byId('resultsCopy').textContent = '按每日人员优先级生成汇总、归属明细和回退备注。';
      byId('anomalyTab').textContent = '人员备注';
    } 
    else {
      weeklyUpload.hidden = true;
      drop.hidden = true;
      weeklyNewTaskToolbar.hidden = false;
      byId('uploadTitle').textContent = '上传周报 ZIP 或选择数据目录';
      byId('uploadHint').textContent = '自动识别六类 Jira 导出文件，忽略 ~$ 临时文件并生成当前周统计。';
      byId('selectButton').hidden = true;
    }
    
    drop.addEventListener('click', event => { if (isDaily && !event.target.closest('.primary-button')) fileInput.click(); });
    byId('selectButton').addEventListener('click', event => { event.stopPropagation(); fileInput.click(); });
    byId('zipButton').addEventListener('click', () => zipFileInput.click());
    byId('directoryButton').addEventListener('click', () => directoryInput.click());
    byId('newTaskButton').addEventListener('click', () => { newTaskMode = true; activeWeeklyStat = 'new-tasks'; byId('newTaskStatus').textContent = '当前统计：本周新增任务数'; });
    byId('newTaskExcelButton').addEventListener('click', event => { if (isClearClick(event, event.currentTarget)) { clearWeeklyInput(event.currentTarget, 'new-tasks'); return; } activeWeeklyStat = 'new-tasks'; pasteTargetStat = ''; newTaskExcel.click(); });
    byId('newTaskPasteButton').addEventListener('click', event => { if (isClearClick(event, event.currentTarget)) { clearWeeklyInput(event.currentTarget, 'new-tasks'); return; } newTaskMode = true; activeWeeklyStat = 'new-tasks'; newTaskScreenshot.click(); });
    byId('newTaskGenerateButton').addEventListener('click', () => {
      // The upload controls retain their files in weeklyStatisticInputs so a user
      // can work on several rows in any order. Read that canonical state here.
      const jiraFile = weeklyStatisticInputs['new-tasks-file'] || newTaskSource;
      const screenshot = weeklyStatisticInputs['new-tasks-image'] || newTaskScreenshotFile;
      if (jiraFile && screenshot) {
        newTaskSource = jiraFile;
        newTaskScreenshotFile = screenshot;
        setWeeklyActionState(byId('newTaskGenerateButton'), '正在生成', 'processing');
        uploadNewTask(jiraFile, screenshot);
      } else {
        byId('newTaskStatus').textContent = '请先上传 Excel 并选择截图';
      }
    });
    root.querySelectorAll('[data-stat]').forEach(button => button.addEventListener('click', () => { newTaskMode = true; activeWeeklyStat = button.dataset.stat; byId('newTaskStatus').textContent = `当前统计项：${button.textContent}`; }));
    root.querySelectorAll('[data-stat-upload]').forEach(button => button.addEventListener('click', event => { activeWeeklyStat = button.dataset.statUpload; if (isClearClick(event, button)) { clearWeeklyInput(button, activeWeeklyStat); return; } newTaskMode = true; if (activeWeeklyStat === 'completed-tasks') pasteTargetStat = ''; (activeWeeklyStat === 'completed-tasks' ? newTaskExcel : newTaskScreenshot).click(); }));
    root.querySelectorAll('[data-stat-paste]').forEach(button => button.addEventListener('click', () => { newTaskMode = true; activeWeeklyStat = button.dataset.statPaste; newTaskScreenshot.click(); }));
    root.querySelectorAll('[data-stat-generate]').forEach(button => button.addEventListener('click', () => {
      activeWeeklyStat = button.dataset.statGenerate;
      const needsWorkbook = activeWeeklyStat === 'completed-tasks';
      const ready = needsWorkbook ? weeklyStatFiles[activeWeeklyStat] && weeklyStatImages[activeWeeklyStat] : weeklyStatImages[activeWeeklyStat];
      if (!ready) byId('newTaskStatus').textContent = needsWorkbook ? '请先上传 Excel 并粘贴截图' : '请先粘贴截图';
    }));
    drop.addEventListener('keydown', event => { if (event.target === drop && (event.key === 'Enter' || event.key === ' ')) { event.preventDefault(); (isDaily ? fileInput : zipFileInput).click(); } });
    fileInput.addEventListener('change', () => { if (fileInput.files.length) upload(fileInput.files[0]); });
    zipFileInput.addEventListener('change', () => { if (zipFileInput.files.length) uploadWeekly(Array.from(zipFileInput.files)); });
    directoryInput.addEventListener('change', () => { if (directoryInput.files.length) uploadWeekly(Array.from(directoryInput.files)); });
    newTaskExcel.addEventListener('change', () => { if (newTaskExcel.files.length) { pasteTargetStat = ''; const inputKey = activeWeeklyStat === 'new-tasks' ? 'new-tasks-file' : activeWeeklyStat; weeklyStatisticInputs[inputKey] = newTaskExcel.files[0]; if (activeWeeklyStat === 'new-tasks') { newTaskSource = newTaskExcel.files[0]; setWeeklyActionState(byId('newTaskExcelButton'), '已选文件'); } else { weeklyStatFiles[activeWeeklyStat] = newTaskExcel.files[0]; setWeeklyActionState(root.querySelector(`[data-stat-upload="${activeWeeklyStat}"]`), '已选文件'); } byId('newTaskStatus').textContent = '文件已保存，可继续处理其他统计项'; } });
    newTaskScreenshot.addEventListener('change', () => { if (newTaskScreenshot.files.length) { const inputKey = activeWeeklyStat === 'new-tasks' ? 'new-tasks-image' : activeWeeklyStat; weeklyStatisticInputs[inputKey] = newTaskScreenshot.files[0]; if (activeWeeklyStat === 'new-tasks') { newTaskScreenshotFile = newTaskScreenshot.files[0]; setWeeklyActionState(byId('newTaskPasteButton'), '已选截图'); } else { weeklyStatImages[activeWeeklyStat] = newTaskScreenshot.files[0]; setWeeklyActionState(screenshotActionButton(activeWeeklyStat), '已选截图'); } byId('newTaskStatus').textContent = '截图已保存，可继续处理其他统计项'; } });
    let lastImagePasteAt = 0;
    context.listenGlobal(document, 'paste', async event => {
      if (!newTaskMode || !pasteTargetStat) return;
      const imageFile = clipboardImageFile(event);
      if (!imageFile) {
        showWeeklyInlineError(pasteTargetStat, '未检测到截图，请先复制图片后再按 Ctrl+V');
        return;
      }
      lastImagePasteAt = Date.now();
      event.preventDefault();
      event.stopImmediatePropagation();
      const statKey = pasteTargetStat;
      const scanButton = screenshotActionButton(statKey);
      scanWeeklyScreenshot(scanButton);
      clearWeeklyInlineError(statKey);
      activeWeeklyStat = statKey;
      const inputKey = activeWeeklyStat === 'new-tasks' ? 'new-tasks-image' : activeWeeklyStat;
      weeklyStatisticInputs[inputKey] = imageFile;
      retainPastedScreenshot(weeklyStatisticInputs[inputKey]);
      if (activeWeeklyStat === 'new-tasks') { newTaskScreenshotFile = weeklyStatisticInputs[inputKey]; }
      else { weeklyStatImages[activeWeeklyStat] = weeklyStatisticInputs[inputKey]; }
      saveWeeklyStatInput(statKey, 'screenshot', weeklyStatisticInputs[inputKey]);
      if (scanButton) scanButton._weeklyUploadFile = weeklyStatisticInputs[inputKey];
      completeWeeklyGeneration(scanButton, '已粘贴截图');
      byId('newTaskStatus').textContent = '截图已粘贴，可继续处理其他统计项';
    });
    context.listenGlobal(document, 'keydown', event => {
      const activation = context.activation;
      if (!(event.ctrlKey || event.metaKey) || event.key.toLowerCase() !== 'v' || !newTaskMode || !pasteTargetStat) return;
      const attemptAt = Date.now();
      setTimeout(async () => {
        if (!context.isCurrentActivation(activation)) return;
        if (lastImagePasteAt >= attemptAt || !navigator.clipboard?.read) return;
        try {
          const items = await navigator.clipboard.read();
          if (!context.isCurrentActivation(activation)) return;
          for (const item of items) {
            const imageType = item.types.find(type => type.startsWith('image/'));
            if (!imageType || lastImagePasteAt >= attemptAt) continue;
            const blob = await item.getType(imageType);
            if (!context.isCurrentActivation(activation)) return;
            const transfer = new DataTransfer();
            transfer.items.add(new File([blob], 'clipboard.png', { type: imageType }));
            clipboardPasteTarget.dispatchEvent(new ClipboardEvent('paste', {
              bubbles: true, cancelable: true, clipboardData: transfer
            }));
            return;
          }
        } catch (_) {
          // Direct paste remains available where clipboard-read permission is denied.
        }
      }, 120);
    }, true);
    // Keep the files used for generation in one state object. These listeners run
    // after the UI handlers above, so changing rows cannot make a visible upload
    // disappear from the corresponding generate action.
    newTaskExcel.addEventListener('change', () => {
      if (!newTaskExcel.files.length) return;
      saveWeeklyStatInput(activeWeeklyStat, 'excel', newTaskExcel.files[0]);
      const button = activeWeeklyStat === 'new-tasks'
        ? byId('newTaskExcelButton')
        : root.querySelector(`[data-stat-upload="${activeWeeklyStat}"]`);
      if (button) {
        button._weeklyUploadFile = newTaskExcel.files[0];
        scanWeeklyUpload(button, '已选文件');
      }
    }, true);
    newTaskScreenshot.addEventListener('change', () => {
      if (!newTaskScreenshot.files.length) return;
      saveWeeklyStatInput(activeWeeklyStat, 'screenshot', newTaskScreenshot.files[0]);
      const button = screenshotActionButton(activeWeeklyStat);
      if (button) {
        button._weeklyUploadFile = newTaskScreenshot.files[0];
        scanWeeklyUpload(button, '已选截图');
      }
    });
    drop.addEventListener('dragover', event => { event.preventDefault(); drop.classList.add('is-dragging'); });
    drop.addEventListener('dragleave', event => { if (!drop.contains(event.relatedTarget)) drop.classList.remove('is-dragging'); });
    drop.addEventListener('drop', event => { event.preventDefault(); drop.classList.remove('is-dragging'); if (!event.dataTransfer.files.length) return; const files = Array.from(event.dataTransfer.files); if (isDaily) upload(files[0]); else uploadWeekly(files); });
    byId('clearButton').addEventListener('click', reset);
    byId('newTaskGenerateButton').addEventListener('click', event => {
      if (isClearClick(event, event.currentTarget)) {
        currentJobId = '';
        currentResult = null;
        setWeeklyActionState(event.currentTarget, '生成文件');
        byId('newTaskStatus').textContent = '结果已移除，可重新生成';
        event.stopImmediatePropagation();
        return;
      }
      if (generatedStatKey !== 'new-tasks' || !currentJobId || !currentResult) return;
      event.stopImmediatePropagation();
      context.downloads.request({ url: `/api/jira/weekly-new-tasks/jobs/${encodeURIComponent(currentJobId)}/export` });
    }, true);
    downloadButton.addEventListener('click', () => { if (!currentJobId) return; const endpoint = newTaskMode ? `/api/jira/weekly-new-tasks/jobs/${encodeURIComponent(currentJobId)}/export` : isDaily ? `/api/jira/jobs/${encodeURIComponent(currentJobId)}/export` : `/api/jira/weekly-statistics/jobs/${encodeURIComponent(currentJobId)}/export`; context.downloads.request({ url: endpoint }); });
    root.querySelectorAll('.tab').forEach(tab => tab.addEventListener('click', () => showView(tab.dataset.view, tab)));
    root.querySelectorAll('[data-stat-generate]').forEach(button => button.addEventListener('click', event => {
      const statKey = button.dataset.statGenerate;
      if (generatedStatKey !== statKey || !currentJobId || !currentResult) return;
      if (isClearClick(event, button)) {
        currentJobId = ''; currentResult = null; generatedStatKey = '';
        setWeeklyActionState(button, '生成文件');
        event.stopImmediatePropagation();
        return;
      }
      event.stopImmediatePropagation();
      context.downloads.request({ url: `/api/jira/weekly-new-tasks/jobs/${encodeURIComponent(currentJobId)}/export` });
    }, true));
    root.querySelectorAll('[data-stat-generate]').forEach(button => button.addEventListener('click', () => {
      const statKey = button.dataset.statGenerate;
      if (statKey === 'completed-tasks' && weeklyStatFiles[statKey] && weeklyStatImages[statKey]) {
        setWeeklyActionState(button, '正在生成', 'processing');
        uploadNewTask(weeklyStatFiles[statKey], weeklyStatImages[statKey], statKey, button);
      } else if (statKey !== 'completed-tasks' && weeklyStatImages[statKey]) {
        setWeeklyActionState(button, '正在生成', 'processing');
        uploadScreenshotStat(statKey, weeklyStatImages[statKey], button);
      }
    }));
    // Generate from the same per-row state that drives the "uploaded" label.
    // This capture handler prevents the legacy bubble handlers below from
    // overwriting the action with a stale "please paste" message.
    root.querySelectorAll('[data-stat-generate], #newTaskGenerateButton').forEach(button => button.addEventListener('click', event => {
      const statKey = button.id === 'newTaskGenerateButton' ? 'new-tasks' : button.dataset.statGenerate;
      if (isClearClick(event, button) || weeklyStatJobs[statKey]) return;
      const screenshotButton = screenshotActionButton(statKey);
      const excelButton = statKey === 'new-tasks'
        ? byId('newTaskExcelButton')
        : root.querySelector(`[data-stat-upload="${statKey}"]`);
      const screenshot = screenshotButton?._weeklyUploadFile || getWeeklyStatInput(statKey, 'screenshot') || newTaskScreenshot.files[0]
        || (statKey === 'new-tasks' ? weeklyStatisticInputs['new-tasks-image'] || newTaskScreenshotFile : weeklyStatImages[statKey]);
      const excel = excelButton?._weeklyUploadFile || getWeeklyStatInput(statKey, 'excel')
        || (statKey === 'new-tasks' ? weeklyStatisticInputs['new-tasks-file'] || newTaskSource : weeklyStatFiles[statKey]);
      const needsWorkbook = statKey === 'new-tasks' || statKey === 'completed-tasks';
      event.stopImmediatePropagation();
      clearWeeklyInlineError(statKey);
      if (!screenshot || (needsWorkbook && !excel)) {
        showWeeklyInlineError(statKey, needsWorkbook ? '请先上传 Excel 并粘贴截图' : '请先粘贴截图');
        return;
      }
      scanWeeklyGeneration(button);
      if (needsWorkbook) uploadNewTask(excel, screenshot, statKey, button);
      else uploadScreenshotStat(statKey, screenshot, button);
    }, true));
    
    byId('weeklySummaryClearButton').addEventListener('click', () => {
      const button = byId('weeklySummaryButton');
      if (button.disabled) return;
      context.downloads.clear('weekly-summary');
      cancelWeeklyScan(button);
      button.textContent = '整合生成';
      byId('weeklySummaryClearButton').hidden = true;
      byId('newTaskStatus').textContent = '整合结果已移除，可重新生成';
    });

    byId('weeklySummaryButton').addEventListener('click', async () => {
      const button = byId('weeklySummaryButton');
      if (button.disabled) return;
      const completedJobs = {};
      const completedStatKeys = [];
      root.querySelectorAll('[data-stat-generate], #newTaskGenerateButton').forEach(action => {
        const statKey = action.id === 'newTaskGenerateButton' ? 'new-tasks' : action.dataset.statGenerate;
        const jobId = action.dataset.weeklyJobId || weeklyStatJobs[statKey];
        const completedPanel = action.closest('.weekly-stat-row')?.querySelector('.weekly-job-progress.is-complete');
        if (jobId) completedJobs[statKey] = jobId;
        else if (completedPanel || action.classList.contains('is-ready')) completedStatKeys.push(statKey);
      });
      if (!Object.keys(completedJobs).length && !completedStatKeys.length) {
        byId('newTaskStatus').textContent = '请先生成至少一项统计文件，再进行整合';
        return;
      }
      button.disabled = true;
      byId('weeklySummaryClearButton').hidden = true;
      scanWeeklyGeneration(button);
      byId('newTaskStatus').textContent = '正在整合生成，请稍候';
      context.downloads.clear('weekly-summary');
      const downloadVersion = context.downloads.version;
      try {
        const response = await fetch('/api/jira/weekly-summary/export', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ job_ids: completedJobs, stat_keys: completedStatKeys }) });
        if (!response.ok) { const data = await response.json(); throw new Error(data.detail || response.statusText); }
        const blob = await response.blob();
        if (downloadVersion !== context.downloads.version) return;
        context.downloads.request({ blob, filename: 'Jira周报统计汇总.xlsx', key: 'weekly-summary', version: downloadVersion });
        completeWeeklyGeneration(button, '重新生成');
        byId('weeklySummaryClearButton').hidden = false;
        byId('newTaskStatus').textContent = '整合生成完成，可下载文件、重新生成或清除整合结果';
      } catch (error) { cancelWeeklyScan(button); setWeeklyActionState(button, '整合生成'); byId('newTaskStatus').textContent = error.message || '整合生成失败'; }
      finally {
        if (button._scanActive) cancelWeeklyScan(button);
        button.disabled = false;
      }
    });
    
    async function upload(file) {
      newTaskMode = false;
      if (!file.name.toLowerCase().endsWith('.xlsx')) { showError('请选择 Jira 导出的 .xlsx 文件'); return; }
      currentJobId = ''; currentResult = null; downloadButton.disabled = true;
      root.querySelectorAll('.view-panel').forEach(view => { view.innerHTML = ''; });
      drop.hidden = true; progress.hidden = false; workspace.hidden = true; setProgress({ percent: 2, stage: '上传文件', detail: '正在上传 Excel 文件' });
      const body = new FormData(); body.append('file', file);
      try {
        const response = await fetch(`/api/jira/import?mode=${jiraMode}`, { method: 'POST', body });
        const data = await response.json();
        if (!response.ok) throw new Error(data.detail || response.statusText);
        currentJobId = data.job_id;
        await poll(data.job_id);
      } catch (error) { showError(error.message || '上传失败'); }
      finally { fileInput.value = ''; drop.hidden = false; progress.hidden = true; }
    }
    
    async function uploadWeekly(files) {
      newTaskMode = false;
      const valid = files.filter(file => file.name.toLowerCase().endsWith('.xlsx') || file.name.toLowerCase().endsWith('.zip'));
      if (!valid.length) { showError('请选择周报 ZIP 或包含 Excel 的目录'); return; }
      currentJobId = ''; currentResult = null; downloadButton.disabled = true;
      byId('weeklyFiles').textContent = `已选择 ${valid.length} 个文件：${valid.map(file => file.name).join('、')}`;
      drop.hidden = true; weeklyUpload.hidden = true; progress.hidden = false; workspace.hidden = true;
      setProgress({ percent: 2, stage: '上传文件', detail: '正在上传周报 ZIP/目录数据' });
      const body = new FormData(); valid.forEach(file => body.append('files', file, file.webkitRelativePath || file.name));
      try {
        const response = await fetch('/api/jira/weekly-statistics/import', { method: 'POST', body });
        const data = await response.json();
        if (!response.ok) throw new Error(data.detail || response.statusText);
        currentJobId = data.job_id;
        await pollWeekly(data.job_id);
      } catch (error) { showError(error.message || '周报统计失败'); }
      finally { zipFileInput.value = ''; directoryInput.value = ''; drop.hidden = false; weeklyUpload.hidden = false; progress.hidden = true; }
    }
    
    async function pollWeekly(jobId) {
      while (true) {
        await wait(450);
        const response = await fetch(`/api/jira/weekly-statistics/jobs/${encodeURIComponent(jobId)}`);
        const data = await response.json();
        if (!response.ok) throw new Error(data.detail || response.statusText);
        setProgress(data.progress || {});
        if (data.status === 'done') { currentResult = data.result; renderWeekly(data.source, data.result); return; }
        if (data.status === 'error') throw new Error(data.error || '周报统计失败');
      }
    }
    
    async function uploadNewTask(jiraFile, screenshot, statKey = 'new-tasks', actionButton = byId('newTaskGenerateButton')) {
      const generationVersion = actionButton._weeklyInputVersion || 0;
      newTaskMode = true; currentJobId = ''; currentResult = null; downloadButton.disabled = true;
      drop.hidden = true; weeklyUpload.hidden = true; progress.hidden = false; workspace.hidden = true;
      setProgress({ percent: 5, stage: '上传文件', detail: '正在上传 Jira Excel 和统计截图' });
      const body = new FormData(); body.append('jira_file', jiraFile); body.append('screenshot', screenshot);
      try {
        const { response, data } = await postFormJson(`/api/jira/weekly-new-tasks/import?stat_key=${encodeURIComponent(statKey)}`, body, 90000);
        if (!response.ok) throw new Error(data.detail || response.statusText);
        const jobId = data.job_id;
        currentJobId = jobId;
        const pollingDeadline = Date.now() + 300000;
        while (true) {
          if (Date.now() > pollingDeadline) throw new Error('统计处理超过 5 分钟仍未完成，请检查截图和 Excel 后重试');
          await wait(450);
          const { response: statusResponse, data: status } = await getJson(`/api/jira/weekly-new-tasks/jobs/${encodeURIComponent(jobId)}`, 30000);
          if (!statusResponse.ok) throw new Error(status.detail || statusResponse.statusText);
          const completed = status.status === 'done';
          setProgress(status.progress || {});
          if (completed) {
            clearWeeklyInlineError(statKey);
            weeklyStatJobs[statKey] = jobId;
            actionButton.dataset.weeklyJobId = jobId;
            completeWeeklyGeneration(actionButton, weeklyStatFileNames[statKey] || '统计结果.xlsx');
            generatedStatKey = statKey;
            currentResult = status.result;
            weeklyStatResults[statKey] = status.result.summary || [];
            if (statKey === 'new-tasks') renderNewTask(status.result);
            break;
          }
          updateWeeklyGenerationProgress(actionButton, status.progress);
          if (status.status === 'error') throw new Error(status.error || '本周新增任务统计失败');
        }
      } catch (error) { if (actionButton._weeklyInputVersion === generationVersion) { workspace.hidden = true; const message = error.message || '统计生成失败'; showWeeklyInlineError(statKey, message); showWeeklyJobFailure(actionButton, message); cancelWeeklyScan(actionButton); } }
      finally { newTaskExcel.value = ''; newTaskScreenshot.value = ''; newTaskSource = null; newTaskScreenshotFile = null; newTaskMode = false; drop.hidden = true; weeklyUpload.hidden = true; progress.hidden = true; weeklyNewTaskToolbar.hidden = false; }
    }
    
    async function uploadScreenshotStat(statKey, screenshot, actionButton) {
      const generationVersion = actionButton._weeklyInputVersion || 0;
      newTaskMode = true; currentJobId = ''; currentResult = null; generatedStatKey = '';
      const body = new FormData(); body.append('screenshot', screenshot);
      try {
        const { response, data } = await postFormJson(`/api/jira/weekly-screenshot-stats/${encodeURIComponent(statKey)}/import`, body, 90000);
        if (!response.ok) throw new Error(data.detail || response.statusText);
        const jobId = data.job_id;
        currentJobId = jobId;
        const pollingDeadline = Date.now() + 300000;
        while (true) {
          if (Date.now() > pollingDeadline) throw new Error('截图统计处理超过 5 分钟仍未完成，请重试');
          await wait(450);
          const { response: statusResponse, data: status } = await getJson(`/api/jira/weekly-new-tasks/jobs/${encodeURIComponent(jobId)}`, 30000);
          if (!statusResponse.ok) throw new Error(status.detail || statusResponse.statusText);
          const completed = status.status === 'done';
          if (completed) {
            clearWeeklyInlineError(statKey);
            currentResult = status.result;
            weeklyStatJobs[statKey] = jobId;
            actionButton.dataset.weeklyJobId = jobId;
            weeklyStatResults[statKey] = status.result.summary || [];
            generatedStatKey = statKey;
            completeWeeklyGeneration(actionButton, weeklyStatFileNames[statKey] || '统计结果.xlsx');
            break;
          }
          updateWeeklyGenerationProgress(actionButton, status.progress);
          if (status.status === 'error') throw new Error(status.error || '统计生成失败');
        }
      } catch (error) { if (actionButton._weeklyInputVersion === generationVersion) { const message = error.message || '统计生成失败'; showWeeklyInlineError(statKey, message); showWeeklyJobFailure(actionButton, message); cancelWeeklyScan(actionButton); } }
    }
    
    function renderNewTask(result) {
      workspace.hidden = true;
      downloadButton.disabled = true;
      byId('sourceName').textContent = '本周新增任务统计';
      byId('summaryNotice').textContent = `已检查 ${result.mismatches?.length || 0} 条开发人员与经办人不一致的任务，并按开发人员修正数量。`;
      byId('totalStat').textContent = result.summary?.reduce((sum, item) => sum + Number(item['本周新增任务数'] || 0), 0) || 0;
      byId('includedStat').textContent = result.summary?.length || 0;
      byId('pendingStat').textContent = result.mismatches?.length || 0;
      byId('mismatchStat').textContent = '—';
      byId('resultsCopy').textContent = '按截图合计数量调整人员归属后的本周新增任务数。';
      const summary = byId('summaryView');
      const total = (result.summary || []).reduce((sum, item) => sum + Number(item['本周新增任务数'] || 0), 0);
      summary.innerHTML = `<div class="table-scroll"><table class="summary-table"><thead><tr><th>姓名</th><th>本周新增任务数</th></tr></thead><tbody>${(result.summary || []).map(item => `<tr><td><strong>${esc(item['姓名'])}</strong></td><td>${item['本周新增任务数']}</td></tr>`).join('')}<tr><td><strong>合计</strong></td><td><strong>${total}</strong></td></tr></tbody></table></div>`;
      const detail = byId('detailView');
      detail.innerHTML = `<div class="table-scroll"><table><thead><tr><th>问题关键字</th><th>原经办人</th><th>调整后开发人员</th></tr></thead><tbody>${(result.mismatches || []).map(item => `<tr><td>${esc(item.issue_key)}</td><td>${esc(item.assignee)}</td><td>${esc(item.developer)}</td></tr>`).join('')}</tbody></table></div>`;
      showView('summaryView', root.querySelector('.tab[data-view="summaryView"]'));
    }
    
    function renderWeekly(source, result) {
      workspace.hidden = false; downloadButton.disabled = false;
      byId('sourceName').textContent = source || '周报 ZIP/目录';
      byId('summaryNotice').textContent = `报告日期：${result.report_date || '—'}；已处理 ${result.stats?.total_rows || 0} 条 Jira 记录，异常 ${result.stats?.anomaly_count || 0} 条。`;
      byId('resultsCopy').textContent = '当前周 23 人统计、延期/挂起明细和异常清单。';
      byId('totalStat').textContent = result.stats?.total_rows || 0;
      byId('includedStat').textContent = result.summary?.length || 0;
      byId('pendingStat').textContent = result.stats?.anomaly_count || 0;
      byId('mismatchStat').textContent = result.report_date || '—';
      renderWeeklySummary(result.summary || []); renderWeeklyDetails(result.sections || {}); renderAnomalies(result.anomalies || []); renderMapping({来源文件: {header: Object.values(result.source_files || {}).join('、'), column: '', found: true}});
      showView('summaryView', root.querySelector('.tab[data-view="summaryView"]'));
    }
    
    function renderWeeklySummary(items) {
      const target = byId('summaryView');
      const headers = ['姓名','本周新增任务数','本周完成任务数','本周新增缺陷数','本周已修复缺陷数','本周延期缺陷数','总挂起缺陷数','本周延期任务数','总挂起任务数'];
      const keys = ['new_task_count','completed_task_count','new_defect_count','fixed_defect_count','delayed_defect_count','pending_defect_count','delayed_task_count','pending_task_count'];
      target.innerHTML = `<div class="table-scroll"><table class="summary-table"><thead><tr>${headers.map(header => `<th>${header}</th>`).join('')}</tr></thead><tbody>${items.map(item => `<tr><td><strong>${esc(item['姓名'])}</strong></td>${keys.map((key, index) => `<td class="${index >= 4 && item[key] ? 'is-warning' : ''}">${item[key] || (index >= 4 ? '' : 0)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
    }
    
    function renderWeeklyDetails(sections) {
      const labels = { delayed_defects: '本周延期缺陷数', pending_defects: '总挂起缺陷数', delayed_tasks: '本周延期任务数', pending_tasks: '总挂起任务数' };
      const target = byId('detailView');
      target.innerHTML = Object.entries(labels).map(([key, label]) => `<h3>${label}</h3><div class="table-scroll"><table><thead><tr><th>问题关键字</th><th>状态</th><th>概要</th><th>经办人</th><th>开发人员</th><th>交付日期</th><th>解决结果</th><th>归属备注</th></tr></thead><tbody>${(sections[key] || []).map(row => `<tr><td>${esc(row.issue_key)}</td><td>${esc(row.status)}</td><td>${esc(row.summary)}</td><td>${esc(row.assignee)}</td><td>${esc(row.developer)}</td><td>${esc(row.delivery_date || '')}</td><td>${esc(row.resolution || '')}</td><td>${esc(row.owner_note || '')}</td></tr>`).join('')}</tbody></table></div>`).join('');
    }
    
    async function poll(jobId) {
      while (true) {
        await wait(450);
        const response = await fetch(`/api/jira/jobs/${encodeURIComponent(jobId)}`);
        const data = await response.json();
        if (!response.ok) throw new Error(data.detail || response.statusText);
        setProgress(data.progress);
        if (data.status === 'done') { currentResult = data.result; render(data.source, data.result); return; }
        if (data.status === 'error') throw new Error(data.error || 'Jira 数据处理失败');
      }
    }
    
    function setProgress(progressInfo = {}) {
      progressInfo ||= {};
      state.progress = progressInfo;
      byId('progressTitle').textContent = progressInfo.stage || '正在处理 Jira 数据';
      byId('progressText').textContent = progressInfo.detail || '请稍候';
      context.motion.progress(byId('progressBar'), progressInfo.percent);
    }
    
    function render(source, result) {
      workspace.hidden = false;
      downloadButton.disabled = false;
      const stats = result.stats || {};
      byId('sourceName').textContent = source || result.source || 'Jira Excel';
      byId('totalStat').textContent = stats.total_rows ?? 0;
      byId('includedStat').textContent = stats.included_rows ?? 0;
      byId('pendingStat').textContent = stats.pending_developer ?? 0;
      byId('mismatchStat').textContent = stats.mismatched_developer ?? 0;
      const pending = Number(stats.pending_developer || 0);
      const mismatch = Number(stats.mismatched_developer || 0);
      const excluded = Number(stats.excluded_rows || 0);
      const notice = byId('summaryNotice');
      const hasWarning = Boolean(pending || mismatch || excluded);
      notice.className = `notice${hasWarning ? ' is-warn' : ''}`;
      if (isDaily) {
        notice.textContent = `已按人员优先级纳入 ${stats.included_rows || 0} 条任务；其中 ${pending} 条开发人员缺失并已自动回退${excluded ? `，${excluded} 条因三个人员字段均为空而未纳入` : ''}。`;
      } else {
        notice.textContent = pending || mismatch ? `已按人员归属纳入 ${stats.included_rows || 0} 条任务；其中 ${pending} 条开发人员缺失并已自动回退、${mismatch} 条人员不一致，请在异常清单核对。` : `全部 ${stats.total_rows || 0} 条任务均已按开发人员生成汇总。`;
      }
      renderSummary(result.summaries || []); renderDetails(result.records || []); renderAnomalies(result.anomalies || []); renderMapping(result.field_mapping || {});
      showView('summaryView', root.querySelector('.tab[data-view="summaryView"]'));
      requestAnimationFrame(() => context.scrollIntoView(workspace, { behavior: 'smooth', block: 'start' }));
    }
    
    function renderSummary(items) {
      const target = byId('summaryView');
      if (!items.length) { target.innerHTML = '<div class="empty">没有可展示的人员汇总</div>'; return; }
      const headers = isDaily ? '<th>汇总人员</th><th>合并内容</th>' : '<th>人员</th><th>任务数</th><th>已完成</th><th>进行中</th><th>开放</th><th>合并任务内容</th>';
      target.innerHTML = `<div class="table-filter"><label class="table-filter-label" for="summaryFilterInput">筛选</label><input id="summaryFilterInput" class="table-filter-input" type="search" placeholder="搜索人员或任务内容"><span id="summaryFilterCount" class="table-filter-count"></span></div><div class="table-scroll"><table class="summary-table"><thead><tr>${headers}</tr></thead><tbody></tbody></table></div>`;
      bindTableFilter('summaryFilterInput', 'summaryFilterCount', '#summaryView tbody', items, item => `${item.name || ''} ${item.merged_text || ''}`, item => isDaily
        ? `<tr><td title="${esc(item.name)}"><strong>${esc(item.name)}</strong></td><td title="${esc(item.merged_text || '')}">${item.merged_text ? esc(item.merged_text) : '<span class="mapping-key">暂无任务</span>'}</td></tr>`
        : `<tr><td title="${esc(item.name)}"><strong>${esc(item.name)}</strong></td><td>${item.task_count}</td><td>${item.completed_count}</td><td>${item.in_progress_count}</td><td>${item.open_count}</td><td title="${esc(item.merged_text || '')}">${item.merged_text ? esc(item.merged_text) : '<span class="mapping-key">暂无纳入汇总的任务</span>'}</td></tr>`);
    }
    
    function bindTableFilter(inputId, countId, bodySelector, items, getSearchText, renderRow) {
      const input = byId(inputId);
      const count = byId(countId);
      const body = root.querySelector(bodySelector);
      if (!input || !count || !body) return;
      const apply = () => {
        const query = input.value.trim().toLowerCase();
        const filtered = items.filter(item => !query || getSearchText(item).toLowerCase().includes(query));
        body.innerHTML = filtered.map(renderRow).join('');
        count.textContent = `显示 ${filtered.length} / ${items.length} 条`;
      };
      input.addEventListener('input', apply);
      apply();
    }
    
    function renderDetails(records) {
      const target = byId('detailView');
      detailRecords = records;
      detailFilter = { query: '', status: 'all', included: 'all' };
      if (!records.length) { target.innerHTML = '<div class="empty">没有任务明细</div>'; return; }
      const detailHeaders = isDaily ? '<th>行</th><th>问题关键字</th><th>概要</th><th>原始状态</th><th>转换状态</th><th>开发人员</th><th>经办人</th><th>报告人</th><th>汇总人员</th><th>汇总依据</th><th>备注</th><th>人员情况</th><th>汇总</th>' : '<th>行</th><th>问题关键字</th><th>概要</th><th>原始状态</th><th>展示状态</th><th>经办人</th><th>报告人</th><th>开发人员</th><th>汇总人员</th><th>汇总依据</th><th>备注</th><th>模块</th><th>计划完成日期</th><th>汇总</th>';
      target.innerHTML = `<div class="detail-toolbar" role="search" aria-label="任务明细筛选"><label class="detail-filter-label" for="detailFilterInput">筛选</label><input id="detailFilterInput" class="detail-filter-input" type="search" placeholder="搜索关键字、概要或人员"><select id="detailStatusFilter" class="detail-filter-select" aria-label="展示状态"><option value="all">全部状态</option><option value="已完成">已完成</option><option value="进行中">进行中</option><option value="开放">开放</option></select><select id="detailIncludedFilter" class="detail-filter-select" aria-label="汇总状态"><option value="all">全部汇总状态</option><option value="included">已纳入</option><option value="excluded">未纳入</option></select><span id="detailFilterCount" class="detail-filter-count"></span></div><div class="detail-table-scroll"><table><thead><tr>${detailHeaders}</tr></thead><tbody></tbody></table></div>`;
      const input = byId('detailFilterInput');
      const statusSelect = byId('detailStatusFilter');
      const includedSelect = byId('detailIncludedFilter');
      const applyFilter = () => {
        detailFilter = { query: input.value.trim().toLowerCase(), status: statusSelect.value, included: includedSelect.value };
        renderDetailRows();
      };
      input.addEventListener('input', applyFilter);
      statusSelect.addEventListener('change', applyFilter);
      includedSelect.addEventListener('change', applyFilter);
      renderDetailRows();
    }
    
    function renderDetailRows() {
      const target = root.querySelector('#detailView .detail-table-scroll tbody');
      const count = byId('detailFilterCount');
      if (!target || !count) return;
      const filtered = detailRecords.filter(record => {
        const searchable = [record.issue_key, record.summary, record.status_raw, record.status, record.assignee, record.reporter, record.developer, record.summary_person, record.summary_source, record.summary_note, record.module, record.planned_completion]
          .map(value => String(value || '')).join(' ').toLowerCase();
        const matchesQuery = !detailFilter.query || searchable.includes(detailFilter.query);
        const matchesStatus = detailFilter.status === 'all' || record.status === detailFilter.status;
        const matchesIncluded = detailFilter.included === 'all' || (detailFilter.included === 'included' ? record.included : !record.included);
        return matchesQuery && matchesStatus && matchesIncluded;
      });
      target.innerHTML = filtered.map(detailRowHtml).join('');
      count.textContent = `显示 ${filtered.length} / ${detailRecords.length} 条`;
    }
    
    function detailRowHtml(record) {
        const summary = record.summary || '—';
        const rawStatus = record.status_raw || '—';
        const status = record.status || '—';
        const assignee = record.assignee || '—';
        const reporter = record.reporter || '—';
        const developer = record.developer || '—';
        const summaryPerson = record.summary_person || assignee;
        const summarySource = record.summary_source || '经办人';
        const summaryNote = record.summary_note || '—';
        const module = record.module || '—';
        const planned = record.planned_completion || '—';
        const included = record.included ? '已纳入' : '未纳入';
      if (isDaily) {
        const consistency = record.consistency || '—';
        return `<tr><td>${record.source_row}</td><td title="${esc(record.issue_key)}"><strong>${esc(record.issue_key)}</strong></td><td class="detail-summary" title="${esc(summary)}">${esc(summary)}</td><td>${esc(rawStatus)}</td><td><span class="badge ${status === '已完成' ? 'badge-success' : status === '进行中' ? 'badge-info' : 'badge-warn'}">${esc(status)}</span></td><td>${esc(developer)}</td><td>${esc(assignee)}</td><td>${esc(reporter)}</td><td>${esc(summaryPerson)}</td><td>${esc(summarySource)}</td><td title="${esc(summaryNote)}">${esc(summaryNote)}</td><td><span class="badge ${record.included ? 'badge-success' : 'badge-warn'}">${esc(consistency)}</span></td><td><span class="badge ${record.included ? 'badge-success' : 'badge-warn'}">${included}</span></td></tr>`;
      }
      return `<tr><td title="${record.source_row}">${record.source_row}</td><td title="${esc(record.issue_key)}"><strong>${esc(record.issue_key)}</strong></td><td class="detail-summary" title="${esc(summary)}">${esc(summary)}</td><td title="${esc(rawStatus)}">${esc(rawStatus)}</td><td title="${esc(status)}"><span class="badge ${status === '已完成' ? 'badge-success' : status === '进行中' ? 'badge-info' : 'badge-warn'}">${esc(status)}</span></td><td title="${esc(assignee)}">${esc(assignee)}</td><td title="${esc(reporter)}">${esc(reporter)}</td><td title="${esc(developer)}">${esc(developer)}</td><td title="${esc(summaryPerson)}">${esc(summaryPerson)}</td><td title="${esc(summarySource)}">${esc(summarySource)}</td><td title="${esc(summaryNote)}">${esc(summaryNote)}</td><td title="${esc(module)}">${esc(module)}</td><td title="${esc(planned)}">${esc(planned)}</td><td title="${included}"><span class="badge ${record.included ? 'badge-success' : 'badge-warn'}">${included}</span></td></tr>`;
    }
    
    function renderAnomalies(items) {
      const target = byId('anomalyView');
      if (!items.length) { target.innerHTML = '<div class="empty">暂无异常，所有记录均符合当前规则</div>'; return; }
      target.innerHTML = `<div class="table-filter"><label class="table-filter-label" for="anomalyFilterInput">筛选</label><input id="anomalyFilterInput" class="table-filter-input" type="search" placeholder="搜索关键字、异常类型或说明"><span id="anomalyFilterCount" class="table-filter-count"></span></div><div class="table-scroll"><table><thead><tr><th>来源行</th><th>问题关键字</th><th>异常类型</th><th>说明</th></tr></thead><tbody></tbody></table></div>`;
      bindTableFilter('anomalyFilterInput', 'anomalyFilterCount', '#anomalyView tbody', items, item => `${item.issue_key || ''} ${item.label || ''} ${item.detail || ''}`, item => `<tr><td>${item.row}</td><td title="${esc(item.issue_key)}">${esc(item.issue_key)}</td><td title="${esc(item.label)}"><span class="badge badge-warn">${esc(item.label)}</span></td><td title="${esc(item.detail)}">${esc(item.detail)}</td></tr>`);
    }
    
    function renderMapping(mapping) {
      const rows = Object.entries(mapping);
      const target = byId('mappingView');
      if (!rows.length) { target.innerHTML = '<div class="empty">没有字段映射信息</div>'; return; }
      target.innerHTML = `<div class="table-filter"><label class="table-filter-label" for="mappingFilterInput">筛选</label><input id="mappingFilterInput" class="table-filter-input" type="search" placeholder="搜索标准字段、表头或列"><span id="mappingFilterCount" class="table-filter-count"></span></div><div class="table-scroll"><table><thead><tr><th>标准字段</th><th>原始表头</th><th>列</th><th>识别结果</th></tr></thead><tbody></tbody></table></div>`;
      bindTableFilter('mappingFilterInput', 'mappingFilterCount', '#mappingView tbody', rows, ([key, value]) => `${key} ${value.header || ''} ${value.column || ''} ${value.found ? '已识别' : '未识别'}`, ([key, value]) => `<tr><td title="${esc(key)}">${esc(key)}</td><td title="${esc(value.header || '—')}">${esc(value.header || '—')}</td><td title="${esc(value.column || '—')}">${esc(value.column || '—')}</td><td title="${value.found ? '已识别' : '未识别'}"><span class="badge ${value.found ? 'badge-success' : 'badge-warn'}">${value.found ? '已识别' : '未识别'}</span></td></tr>`);
    }
    
    function showView(id, activeTab) {
      const next = byId(id);
      if (!next) return;
      const current = root.querySelector('.view-panel.is-active-view');
      root.querySelectorAll('.tab').forEach(tab => tab.classList.toggle('is-active', tab === activeTab));
      if (current === next && !next.hidden) return;
      root.querySelectorAll('.view-panel').forEach(view => {
        context.motion?.cleanup(view);
        view.hidden = view !== next;
        view.classList.remove('is-entering', 'is-leaving');
        view.classList.toggle('is-active-view', view === next);
      });
      context.motion?.enter(next, 'x');
    }
    
    function showError(message) { currentJobId = ''; currentResult = null; downloadButton.disabled = true; workspace.hidden = false; byId('sourceName').textContent = '处理失败'; byId('summaryNotice').className = 'notice is-warn'; byId('summaryNotice').textContent = message; byId('totalStat').textContent = '—'; byId('includedStat').textContent = '—'; byId('pendingStat').textContent = '—'; byId('mismatchStat').textContent = '—'; }
    function reset() {
      context.downloads.clear();
      currentJobId = ''; currentResult = null; downloadButton.disabled = true; workspace.hidden = true;
      root.querySelectorAll('.view-panel').forEach(view => { context.motion?.cleanup(view); view.hidden = true; view.classList.remove('is-active-view', 'is-entering', 'is-leaving'); view.innerHTML = ''; });
      const summary = byId('summaryView');
      summary.hidden = false; summary.classList.add('is-active-view');
      root.querySelectorAll('.tab').forEach(tab => tab.classList.toggle('is-active', tab.dataset.view === 'summaryView'));
      context.scrollTo({ top: 0, behavior: 'smooth' });
    }
    function wait(ms) { return new Promise(resolve => setTimeout(resolve, ms)); }
    
    async function fetchWithTimeout(url, options = {}, timeoutMs = 30000) {
      const controller = new AbortController();
      const timer = setTimeout(() => controller.abort(), timeoutMs);
      try {
        return await fetch(url, { ...options, signal: controller.signal });
      } finally {
        clearTimeout(timer);
      }
    }
    function postFormJson(url, body, timeoutMs = 90000) {
      return new Promise((resolve, reject) => {
        const xhr = new XMLHttpRequest();
        let settled = false;
        const finish = (fn, value) => { if (settled) return; settled = true; fn(value); };
        xhr.open('POST', url, true);
        xhr.timeout = timeoutMs;
        const handleResponse = () => {
          if (xhr.readyState !== 4) return;
          let data = {};
          try { data = xhr.responseText ? JSON.parse(xhr.responseText) : {}; } catch (_) { data = { detail: xhr.responseText || '服务器返回了无法读取的结果' }; }
          finish(resolve, { response: { ok: xhr.status >= 200 && xhr.status < 300, status: xhr.status, statusText: xhr.statusText }, data });
        };
        xhr.onload = handleResponse;
        xhr.onreadystatechange = handleResponse;
        xhr.onerror = () => finish(reject, new Error('上传请求失败，请检查本地服务是否正常运行'));
        xhr.ontimeout = () => finish(reject, new Error('上传处理超时，请重新粘贴截图后重试'));
        xhr.send(body);
      });
    }
    function getJson(url, timeoutMs = 30000) {
      return new Promise((resolve, reject) => {
        const xhr = new XMLHttpRequest();
        let settled = false;
        const finish = (fn, value) => { if (settled) return; settled = true; fn(value); };
        xhr.open('GET', url, true);
        xhr.timeout = timeoutMs;
        const handleResponse = () => {
          if (xhr.readyState !== 4) return;
          let data = {};
          try { data = xhr.responseText ? JSON.parse(xhr.responseText) : {}; } catch (_) { data = { detail: xhr.responseText || '服务器返回了无法读取的结果' }; }
          finish(resolve, { response: { ok: xhr.status >= 200 && xhr.status < 300, status: xhr.status, statusText: xhr.statusText }, data });
        };
        xhr.onload = handleResponse;
        xhr.onreadystatechange = handleResponse;
        xhr.onerror = () => finish(reject, new Error('读取处理进度失败，请检查本地服务是否正常运行'));
        xhr.ontimeout = () => finish(reject, new Error('读取处理进度超时，请重试'));
        xhr.send();
      });
    }
    function esc(value) { return String(value ?? '').replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char]); }
    Object.defineProperty(state, "currentJobId", { enumerable: true, get: () => currentJobId, set: value => { currentJobId = value; } });
    Object.defineProperty(state, "currentResult", { enumerable: true, get: () => currentResult, set: value => { currentResult = value; } });
    Object.defineProperty(state, "newTaskMode", { enumerable: true, get: () => newTaskMode, set: value => { newTaskMode = value; } });
    Object.defineProperty(state, "newTaskSource", { enumerable: true, get: () => newTaskSource, set: value => { newTaskSource = value; } });
    Object.defineProperty(state, "newTaskScreenshotFile", { enumerable: true, get: () => newTaskScreenshotFile, set: value => { newTaskScreenshotFile = value; } });
    Object.defineProperty(state, "activeWeeklyStat", { enumerable: true, get: () => activeWeeklyStat, set: value => { activeWeeklyStat = value; } });
    Object.defineProperty(state, "pasteTargetStat", { enumerable: true, get: () => pasteTargetStat, set: value => { pasteTargetStat = value; } });
    Object.defineProperty(state, "generatedStatKey", { enumerable: true, get: () => generatedStatKey, set: value => { generatedStatKey = value; } });
    Object.defineProperty(state, "detailRecords", { enumerable: true, get: () => detailRecords, set: value => { detailRecords = value; } });
    Object.defineProperty(state, "detailFilter", { enumerable: true, get: () => detailFilter, set: value => { detailFilter = value; } });
    Object.defineProperty(state, "lastImagePasteAt", { enumerable: true, get: () => lastImagePasteAt, set: value => { lastImagePasteAt = value; } });
    state.previewVersion = 0;
    Object.assign(state, { weeklyStatisticInputs, weeklyStatFiles, weeklyStatImages, weeklyStatResults, weeklyStatJobs, weeklyStatState });
    context.onDeactivate(() => {
      closePreview();
      root.querySelectorAll('.is-awaiting-paste').forEach(button => button.classList.remove('is-awaiting-paste'));
    });
  }
  return context.lifecycle(mount);
}
