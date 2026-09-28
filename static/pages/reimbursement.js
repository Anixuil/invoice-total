// Each factory owns its DOM, files and task state for one route instance.
export function createPage(context) {
  const { root, state } = context;
  const byId = id => root.querySelector(`#${CSS.escape(id)}`);
  const { fetch, XMLHttpRequest, setTimeout, clearTimeout } = context;
  function mount() {
    const input=root.querySelector('#file'),zone=root.querySelector('#drop'),name=root.querySelector('#name'),button=root.querySelector('#generate'),buttonLabel=root.querySelector('#generateLabel'),status=root.querySelector('#status'),result=root.querySelector('#result'),progressBox=root.querySelector('#progress'),progressStage=root.querySelector('#progressStage'),progressPercent=root.querySelector('#progressPercent'),progressBar=root.querySelector('#progressBar'),progressDetail=root.querySelector('#progressDetail');
      let selected,busy=false;
    
      function setProgress(stage,percent,detail,progressState=''){
        state.progress = { stage, percent, detail, state: progressState };
        const value=context.motion.progress(progressBar,percent);
        progressBox.hidden=false;
        progressBox.className=`job-progress${progressState==='error'?' is-error':''}`;
        progressStage.textContent=stage||'处理中';
        progressPercent.textContent=value===null?'处理中':`${value}%`;
        progressDetail.textContent=detail||'';

      }
    
      function setLoading(active){
        busy=active;
        button.disabled=active||!selected;
        input.disabled=active;
        zone.setAttribute('aria-disabled',String(active));
        button.classList.toggle('is-loading',active);
        button.setAttribute('aria-busy',String(active));
        buttonLabel.textContent=active?'正在生成':'生成 PDF';
      }
    
      function set(file){
        if(!file||busy)return;
        if(!file.name.toLowerCase().endsWith('.docx')){
          status.textContent='请选择 DOCX 格式的报销信息文件。';
          status.className='status error';
          return;
        }
        context.downloads.clear();
        selected=file;
        name.textContent=file.name;
        button.disabled=false;
        status.textContent='';
        status.className='status';
        result.hidden=true;
        progressBox.hidden=true;
      }
    
      function uploadJob(data){
        return new Promise((resolve,reject)=>{
          const request=new XMLHttpRequest();
          request.open('POST','/api/reimbursement/jobs');
          request.responseType='json';
          request.upload.addEventListener('progress',event=>{
            const percent=event.lengthComputable?Math.round(event.loaded/event.total*10):null;
            const total=event.lengthComputable?event.total:selected.size;
            setProgress('上传报销信息',percent,`已上传 ${formatBytes(event.loaded)} / ${formatBytes(total)}`);
          });
          request.addEventListener('load',()=>{
            const body=request.response||{};
            if(request.status>=200&&request.status<300)resolve(body);
            else reject(new Error(body.detail||request.statusText||'上传失败'));
          });
          request.addEventListener('error',()=>reject(new Error('网络连接中断，上传未完成')));
          request.send(data);
        });
      }
    
      async function waitForJob(jobId){
        while(true){
          const response=await fetch(`/api/reimbursement/jobs/${jobId}`,{cache:'no-store'});
          const body=await response.json().catch(()=>({}));
          if(!response.ok)throw new Error(body.detail||'无法读取生成进度');
          setProgress(body.progress?.stage,body.progress?.percent,body.progress?.detail,body.status==='error'?'error':'');
          if(body.status==='done')return body;
          if(body.status==='error')throw new Error(body.error||'生成失败');
          await new Promise(resolve=>setTimeout(resolve,400));
        }
      }
    
      function formatBytes(bytes){
        return bytes>=1048576?`${(bytes/1048576).toFixed(1)} MB`:`${Math.max(1,Math.round(bytes/1024))} KB`;
      }
    
      zone.addEventListener('keydown',event=>{if((event.key==='Enter'||event.key===' ')&&!busy){event.preventDefault();input.click();}});
      input.addEventListener('change',()=>set(input.files[0]));
      ['dragenter','dragover'].forEach(type=>zone.addEventListener(type,event=>{event.preventDefault();if(!busy)zone.classList.add('drag')}));
      ['dragleave','drop'].forEach(type=>zone.addEventListener(type,event=>{event.preventDefault();zone.classList.remove('drag')}));
      zone.addEventListener('drop',event=>set(event.dataTransfer.files[0]));
      button.addEventListener('click',async()=>{
        if(!selected||busy)return;
        context.downloads.clear();
        const downloadVersion=context.downloads.version;
        setLoading(true);
        status.textContent='';
        status.className='status';
        result.hidden=true;
        setProgress('上传报销信息',null,'正在上传 DOCX 文件');
        const data=new FormData();
        data.append('file',selected);
        try{
          const created=await uploadJob(data);
          if(!created.job_id)throw new Error('服务器未返回生成任务编号');
          state.jobId=created.job_id;
          const completed=await waitForJob(created.job_id);
          state.result=completed;
          const response=await fetch(completed.download_url,{cache:'no-store'});
          if(!response.ok){const body=await response.json().catch(()=>({}));throw new Error(body.detail||'下载生成文件失败')}
          const blob=await response.blob();
          const disposition=response.headers.get('content-disposition')||'',filenameMatch=disposition.match(/filename\*=utf-8''([^;]+)/i);
          context.downloads.request({ blob, key: 'result', version: downloadVersion, filename: filenameMatch?decodeURIComponent(filenameMatch[1]):completed.count>1?'报销单.zip':'报销单.pdf' });
          setProgress('生成完成',100,`${completed.count||1} 张报销单已通过整体核验`);
          result.textContent=completed.count>1?`已生成 ${completed.count} 张报销单，文件已就绪，可通过下方链接再次下载。`:'PDF 已生成，整体核验通过，文件已就绪，可通过下方链接再次下载。';
          result.hidden=false;
        }catch(error){
          setProgress('生成失败',null,error.message,'error');
          status.textContent=error.message;
          status.className='status error';
        }finally{
          setLoading(false);
        }
      });
    Object.defineProperty(state, "selected", { enumerable: true, get: () => selected, set: value => { selected = value; } });
    Object.defineProperty(state, "busy", { enumerable: true, get: () => busy, set: value => { busy = value; } });
  }
  return context.lifecycle(mount);
}
