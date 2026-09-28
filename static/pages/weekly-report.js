// Each factory owns its DOM, files and task state for one route instance.
export function createPage(context) {
  const { root, state } = context;
  const byId = id => root.querySelector(`#${CSS.escape(id)}`);
  const { fetch, XMLHttpRequest, setTimeout, clearTimeout } = context;
  function mount() {
    const drop=byId('drop');
    const fileInput=byId('file');
    const folderInput=byId('folder');
    const selectedPanel=byId('selectedFile');
    const progress=byId('progress');
    const workspace=byId('workspace');
    const downloadButtons={ppt:byId('downloadPptButton'),docx:byId('downloadDocxButton'),xlsx:byId('downloadXlsxButton'),zip:byId('downloadZipButton')};
    let selectedFile=null,currentJobId='',currentResult=null,busy=false;
    
    drop.addEventListener('click',event=>{if(!event.target.closest('button'))fileInput.click();});
    byId('zipButton').addEventListener('click',event=>{event.stopPropagation();fileInput.click();});
    byId('folderButton').addEventListener('click',event=>{event.stopPropagation();folderInput.click();});
    drop.addEventListener('keydown',event=>{if(event.target===drop&&(event.key==='Enter'||event.key===' ')){event.preventDefault();fileInput.click();}});
    drop.addEventListener('dragover',event=>{event.preventDefault();drop.classList.add('is-dragging');});
    drop.addEventListener('dragleave',event=>{if(!drop.contains(event.relatedTarget))drop.classList.remove('is-dragging');});
    drop.addEventListener('drop',event=>{event.preventDefault();drop.classList.remove('is-dragging');chooseFiles([...event.dataTransfer.files]);});
    fileInput.addEventListener('change',()=>{chooseFiles([...fileInput.files]);fileInput.value='';});
    folderInput.addEventListener('change',()=>{chooseFiles([...folderInput.files]);folderInput.value='';});
    byId('removeButton').addEventListener('click',()=>choose(null));
    byId('startButton').addEventListener('click',upload);
    byId('clearButton').addEventListener('click',reset);
    Object.entries(downloadButtons).forEach(([kind,button])=>button.addEventListener('click',()=>{if(currentJobId)context.downloads.request({ url: `/api/weekly-report/jobs/${encodeURIComponent(currentJobId)}/export${kind==='zip'?'':`/${kind}`}` });}));
    root.querySelectorAll('.tab').forEach(tab=>tab.addEventListener('click',()=>showView(tab.dataset.view,tab)));
    
    function isWeeklySource(file){
      const name=(file.webkitRelativePath||file.name).replace(/\\/g,'/').split('/').pop();
      return !name.startsWith('~$')&&/\.(pptx|docx)$/i.test(name);
    }
    function chooseFiles(files){
      if(busy)return;
      byId('uploadError').hidden=true;
      if(!files.length){selectedFile=null;selectedPanel.hidden=true;return;}
      const isArchive=files.length===1&&files[0].name.toLowerCase().endsWith('.zip');
      if(!isArchive&&!files.some(isWeeklySource)){showUploadError('请选择包含 PPTX / DOCX 的文件夹或 ZIP 文件');return;}
      selectedFile=files;selectedPanel.hidden=false;
      const size=files.reduce((total,file)=>total+file.size,0);
      const pptCount=files.filter(isWeeklySource).length;
      byId('selectedName').textContent=isArchive?files[0].name:`文件夹（${files.length} 个目录项，其中 ${pptCount} 个 PPTX / DOCX）`;
      byId('selectedMeta').textContent=`${formatSize(size)} · 文件仅在本机临时处理`;
    }
    function choose(file){chooseFiles(file?[file]:[]);}
    
    function uploadFiles(url,body,onProgress){
      return new Promise((resolve,reject)=>{
        const request=new XMLHttpRequest();request.open('POST',url);request.responseType='text';
        request.upload.addEventListener('progress',event=>{if(event.lengthComputable)onProgress?.(event.loaded,event.total);});
        request.addEventListener('load',()=>{let data={};try{data=request.responseText?JSON.parse(request.responseText):{};}catch{reject(new Error('服务器返回了无法解析的响应'));return;}if(request.status>=200&&request.status<300)resolve(data);else reject(new Error(data.detail||request.statusText||'上传失败'));});
        request.addEventListener('error',()=>reject(new Error(`文件“${body.get('file')?.name||'未知文件'}”上传失败：连接中断或文件无法读取。请确认本地服务正在运行，重新选择文件夹或 ZIP 后重试。`)));request.send(body);
      });
    }
    
    async function upload(){
      if(!selectedFile||busy)return;
      busy=true;byId('startButton').disabled=true;byId('startButton').setAttribute('aria-busy','true');
      currentJobId='';currentResult=null;workspace.hidden=true;drop.hidden=true;selectedPanel.hidden=true;progress.hidden=false;
      const files=selectedFile;setProgress({stage:'上传文件',percent:null,detail:`正在上传 ${files.length===1?files[0].name:`${files.length} 个文件`}`});
      const body=new FormData();const isArchive=files.length===1&&files[0].name.toLowerCase().endsWith('.zip');
      try{
        let data;
        if(isArchive){
          body.append('file',files[0],files[0].name);
          data=await uploadFiles('/api/weekly-report/import',body,(loaded,total)=>{
            const uploadPercent=total?Math.min(10,Math.round(loaded/total*10)):null;
            setProgress({stage:'上传文件',percent:uploadPercent,detail:`正在上传 ${formatSize(loaded)} / ${formatSize(total)}`});
          });
        }else{
          const pptFiles=files.filter(isWeeklySource);
          const start=await fetch('/api/weekly-report/folder/start',{method:'POST'}).then(async response=>{const value=await response.json();if(!response.ok)throw new Error(value.detail||response.statusText);return value;});
          for(let index=0;index<pptFiles.length;index++){
            const file=pptFiles[index],fileBody=new FormData();fileBody.append('file',file,file.name);
            const relativePath=file.webkitRelativePath||file.name;
            await uploadFiles(`/api/weekly-report/folder/${encodeURIComponent(start.job_id)}/file?relative_path=${encodeURIComponent(relativePath)}`,fileBody,(loaded,total)=>{
              const completed=index+(total?loaded/total:0);const percent=Math.min(9,Math.round(completed/pptFiles.length*9));
              setProgress({stage:'上传文件',percent,detail:`正在上传 ${index+1} / ${pptFiles.length}：${file.name}`});
            });
          }
          data=await fetch(`/api/weekly-report/folder/${encodeURIComponent(start.job_id)}/finish`,{method:'POST'}).then(async response=>{const value=await response.json();if(!response.ok)throw new Error(value.detail||response.statusText);return value;});
        }
        currentJobId=data.job_id;await poll(data.job_id);
      }catch(error){showError(error.message||'周报整合失败');}
      finally{busy=false;byId('startButton').disabled=false;byId('startButton').setAttribute('aria-busy','false');progress.hidden=true;drop.hidden=false;selectedPanel.hidden=!selectedFile;}
    }
    
    async function poll(jobId){
      while(true){await wait(500);const response=await fetch(`/api/weekly-report/jobs/${encodeURIComponent(jobId)}`);const data=await response.json();if(!response.ok)throw new Error(data.detail||response.statusText);setProgress(data.progress||{});if(data.status==='done'){render(data.source,data.result);return;}if(data.status==='error')throw new Error(data.error||'周报整合失败');}
    }
    
    function setProgress(info){
      state.progress = info;
      const percent=context.motion.progress(byId('progressBar'),info.percent);byId('progressTitle').textContent=info.stage||'正在处理周报';byId('progressText').textContent=info.detail||'请稍候';byId('progressValue').textContent=percent===null?'处理中':`${percent}%`;
      const active=percent===null?0:percent<12?1:percent<88?2:percent<96?3:4;root.querySelectorAll('.progress-step').forEach(step=>{const index=Number(step.dataset.step);step.classList.toggle('is-active',index===active);step.classList.toggle('is-done',index<active||percent===100);});
    }
    
    function render(source,result){
      currentResult=result;workspace.hidden=false;setDownloadState(false);const stats=result.stats||{};
      byId('sourceName').textContent=source||'部门项目周报.zip';byId('outputName').textContent=`输出：${result.output_stem}.pptx · 部门周例会${String(result.output_stem).replace('项目周报','')}.docx · ${result.output_stem}-审核报告.xlsx`;
      byId('projectStat').textContent=stats.project_count||0;byId('matchedStat').textContent=stats.matched_projects||0;byId('errorStat').textContent=stats.error_count||0;byId('warningStat').textContent=stats.warning_count||0;
      const qa=result.qa||{};const qaText=qa.status?`生成质量 ${qa.score??'—'} 分，${qa.status}`:'';const notice=byId('summaryNotice');notice.className=`notice${stats.error_count||stats.warning_count||qa.stable===false?' is-warn':''}`;notice.textContent=stats.error_count||stats.warning_count?`整合文件已生成；审核发现 ${stats.error_count||0} 个错误、${stats.warning_count||0} 个待确认项。${qaText}`:`${stats.project_count||0} 个模板项目均已对应。${qaText}`;
      renderAssembly(result.projects||[]);renderAudit(result.issues||[]);renderQa(qa);renderMeeting(result.meeting_projects||[]);renderSources(result.sources||[]);showView('assemblyView',root.querySelector('.tab[data-view="assemblyView"]'));requestAnimationFrame(()=>context.scrollIntoView(workspace, {behavior:'smooth',block:'start'}));
    }
    
    function renderAssembly(items){
      tableView('assemblyView','assemblyFilter','assemblyCount',items,item=>`${item.title} ${item.reporter} ${item.status} ${item.source_files.join(' ')}`,'搜索项目、汇报人或源文件','<tr><th>状态</th><th>项目</th><th>汇报人</th><th>源文件</th><th>源页码</th><th>内容页数</th></tr>',item=>`<tr><td><span class="badge ${badgeClass(item.status)}">${esc(item.status)}</span></td><td title="${esc(item.title)}"><strong>${esc(item.title)}</strong></td><td>${esc(item.reporter)}</td><td title="${esc(item.source_files.join('、')||'—')}">${esc(item.source_files.join('、')||'—')}</td><td title="${esc(item.slides.map(slide=>`${slide.file} 第${slide.slide}页`).join('；')||'—')}">${esc(item.slides.map(slide=>slide.slide).join('、')||'—')}</td><td>${item.slides.length}</td></tr>`);
    }
    
    function renderAudit(items){
      const target=byId('auditView');if(!items.length){target.innerHTML='<div class="empty">没有发现审核问题</div>';return;}
      target.innerHTML=`<div class="table-filter"><label class="filter-label" for="auditFilter">筛选</label><input id="auditFilter" class="filter-input" type="search" placeholder="搜索文件、项目、位置或描述"><select id="severityFilter" class="filter-select"><option value="all">全部级别</option><option value="error">错误</option><option value="warning">待确认</option><option value="info">信息</option></select><span id="auditCount" class="filter-count"></span></div><div class="table-scroll"><table><thead><tr><th>级别</th><th>文件</th><th>页码</th><th>具体位置</th><th>项目</th><th>问题</th><th>描述</th><th>处理建议</th></tr></thead><tbody></tbody></table></div>`;
      const apply=()=>{const query=byId('auditFilter').value.trim().toLowerCase();const severity=byId('severityFilter').value;const filtered=items.filter(item=>(severity==='all'||item.severity===severity)&&(!query||`${item.file} ${item.location} ${item.project} ${item.label} ${item.detail}`.toLowerCase().includes(query)));target.querySelector('tbody').innerHTML=filtered.map(item=>`<tr><td><span class="badge ${severityClass(item.severity)}">${severityText(item.severity)}</span></td><td title="${esc(item.file||'—')}">${esc(item.file||'—')}</td><td>${item.slide||'—'}</td><td title="${esc(item.location)}">${esc(item.location)}</td><td title="${esc(item.project||'—')}">${esc(item.project||'—')}</td><td title="${esc(item.label)}">${esc(item.label)}</td><td title="${esc(item.detail)}">${esc(item.detail)}</td><td title="${esc(item.suggestion)}">${esc(item.suggestion)}</td></tr>`).join('');byId('auditCount').textContent=`显示 ${filtered.length} / ${items.length} 条`;};
      byId('auditFilter').addEventListener('input',apply);byId('severityFilter').addEventListener('change',apply);apply();
    }
    
    function renderQa(qa){
      const items=qa.rounds||[];const target=byId('qaView');
      if(!items.length){target.innerHTML='<div class="empty">尚无生成质量核验记录</div>';return;}
      target.innerHTML=`<div class="table-filter"><span class="filter-label">最终状态</span><span class="badge ${qa.stable?'badge-success':'badge-warn'}">${esc(qa.status||'待核验')}</span><span class="filter-count">综合稳定度 ${esc(qa.score??'—')} 分</span></div><div class="table-scroll"><table><thead><tr><th>轮次</th><th>内容准确度</th><th>模板样式准确度</th><th>版面稳定度</th><th>自动修复</th><th>问题数</th><th>连续稳定</th></tr></thead><tbody>${items.map(item=>`<tr><td>第 ${item.round} 轮</td><td>${item.content_score}%</td><td>${item.style_score}%</td><td>${item.layout_score}%</td><td>${item.repairs}</td><td>${item.issues}</td><td><span class="badge ${item.stable?'badge-success':'badge-info'}">${item.stable?'已稳定':'继续核验'}</span></td></tr>`).join('')}</tbody></table></div>`;
    }
    
    function renderMeeting(items){
      const normalizedContent=value=>{
        const content=String(value??'').replace(/\r\n?/g,'\n').trim();
        return !content||/^(?:无|暂无|未填写)[。．.!！]?$/.test(content)?'无':content;
      };
      const meetingItems=items.map(item=>{
        const sourceSections=Array.isArray(item.supplementary_sections)?item.supplementary_sections:[{title:'本周问题',kind:'issue',content:item.issues}];
        const groupedSections=new Map();
        sourceSections.forEach(section=>{
          const sourceTitle=String(section.title??'').trim().replace(/[：:]\s*$/,'');
          const kind=section.kind==='issue'||/问题|风险|阻塞/.test(sourceTitle)?'issue':'normal';
          const title=kind==='issue'?'本周问题':sourceTitle||'补充内容';
          const key=`${kind}:${title}`;
          if(!groupedSections.has(key))groupedSections.set(key,{title,blocks:[]});
          const blocks=groupedSections.get(key).blocks;
          const content=normalizedContent(section.content);
          if(content!=='无'&&!blocks.includes(content))blocks.push(content);
        });
        const sections=Array.from(groupedSections.values(),section=>({title:section.title,content:section.blocks.join('\n\n')||'无'}));
        return {...item,
          current:normalizedContent(item.current),next:normalizedContent(item.next),
          supplementaryText:sections.map(section=>`${section.title}\n${section.content}`).join('\n\n')||'无',
          supplementaryHtml:sections.map(section=>`<div><strong>${esc(section.title)}</strong><br>${esc(section.content).replace(/\n/g,'<br>')}</div>`).join('<br>')||'无'
        };
      });
      tableView('meetingView','meetingFilter','meetingCount',meetingItems,item=>`${item.title} ${item.reporter} ${item.current} ${item.next} ${item.supplementaryText}`,'搜索项目或周例会内容','<tr><th>项目</th><th>汇报人</th><th>本周进展</th><th>下周计划</th><th>补充内容</th></tr>',item=>`<tr><td title="${esc(item.title)}">${esc(item.title)}</td><td>${esc(item.reporter)}</td><td title="${esc(item.current)}">${esc(item.current)}</td><td title="${esc(item.next)}">${esc(item.next)}</td><td title="${esc(item.supplementaryText)}">${item.supplementaryHtml}</td></tr>`);
    }
    
    function renderSources(items){
      tableView('sourceView','sourceFilter','sourceCount',items,item=>`${item.path} ${item.kind} ${item.status}`,'搜索目录、文件名或状态','<tr><th>路径</th><th>类型</th><th>大小</th><th>处理状态</th></tr>',item=>`<tr><td class="path-cell" title="${esc(item.path)}">${esc(item.path)}</td><td>${kindText(item.kind)}</td><td>${formatSize(item.size)}</td><td><span class="badge ${['ppt','docx'].includes(item.kind)?'badge-success':'badge-info'}">${esc(item.status||'已扫描')}</span></td></tr>`);
    }
    
    function tableView(viewId,inputId,countId,items,searchText,placeholder,head,rowHtml){
      const target=byId(viewId);if(!items.length){target.innerHTML='<div class="empty">没有可展示的数据</div>';return;}target.innerHTML=`<div class="table-filter"><label class="filter-label" for="${inputId}">筛选</label><input id="${inputId}" class="filter-input" type="search" placeholder="${placeholder}"><span id="${countId}" class="filter-count"></span></div><div class="table-scroll"><table><thead>${head}</thead><tbody></tbody></table></div>`;const input=byId(inputId),body=target.querySelector('tbody'),count=byId(countId);const apply=()=>{const query=input.value.trim().toLowerCase();const filtered=items.filter(item=>!query||searchText(item).toLowerCase().includes(query));body.innerHTML=filtered.map(rowHtml).join('');count.textContent=`显示 ${filtered.length} / ${items.length} 条`;};input.addEventListener('input',apply);apply();
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
    
    function showUploadError(message){selectedFile=null;selectedPanel.hidden=true;byId('uploadError').textContent=message;byId('uploadError').hidden=false;}
    function setDownloadState(disabled){Object.values(downloadButtons).forEach(button=>button.disabled=disabled);}
    function showError(message){currentJobId='';currentResult=null;setDownloadState(true);workspace.hidden=false;byId('sourceName').textContent='处理失败';byId('outputName').textContent='未生成文件';const notice=byId('summaryNotice');notice.className='notice is-warn';notice.textContent=message;['projectStat','matchedStat','errorStat','warningStat'].forEach(id=>byId(id).textContent='—');root.querySelectorAll('.view-panel').forEach(view=>{context.motion?.cleanup(view);view.innerHTML='';view.hidden=true;view.classList.remove('is-active-view','is-entering','is-leaving');});const first=byId('assemblyView');first.hidden=false;first.classList.add('is-active-view');first.innerHTML='<div class="empty">请检查压缩包后重新处理</div>';}
    function reset(){context.downloads.clear();selectedFile=null;currentJobId='';currentResult=null;workspace.hidden=true;setDownloadState(true);selectedPanel.hidden=true;root.querySelectorAll('.view-panel').forEach(view=>{context.motion?.cleanup(view);view.innerHTML='';view.hidden=true;view.classList.remove('is-active-view','is-entering','is-leaving');});const first=byId('assemblyView');first.hidden=false;first.classList.add('is-active-view');}
    function badgeClass(status){return status==='通过'?'badge-success':status==='待确认'?'badge-warn':'badge-error';}
    function severityClass(value){return value==='error'?'badge-error':value==='warning'?'badge-warn':'badge-info';}
    function severityText(value){return value==='error'?'错误':value==='warning'?'待确认':'信息';}
    function kindText(kind){return({archive:'ZIP',directory:'目录',ppt:'PPTX',docx:'DOCX',historical:'历史成品',ignored:'已忽略'})[kind]||kind||'文件';}
    function formatSize(value){const size=Number(value||0);if(!size)return'—';return size>=1048576?`${(size/1048576).toFixed(1)} MB`:`${Math.max(1,Math.round(size/1024))} KB`;}
    function wait(ms){return new Promise(resolve=>setTimeout(resolve,ms));}
    function esc(value){return String(value??'').replace(/[&<>"']/g,char=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[char]);}
    Object.defineProperty(state, "selectedFile", { enumerable: true, get: () => selectedFile, set: value => { selectedFile = value; } });
    Object.defineProperty(state, "currentJobId", { enumerable: true, get: () => currentJobId, set: value => { currentJobId = value; } });
    Object.defineProperty(state, "currentResult", { enumerable: true, get: () => currentResult, set: value => { currentResult = value; } });
  }
  return context.lifecycle(mount);
}
