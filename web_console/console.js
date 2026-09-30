'use strict';
(() => {
  const $ = id => document.getElementById(id);
  let state=null,csrf='',cursor=0,initial=true,connected=false,pending=false,exiting=false;
  let logLines=[],catalogKey='',resultKey='',polling=false;
  const labels={idle:'就绪',running:'运行中',stopping:'正在停止',finishing:'正在收尾',completed:'已完成',stopped:'已停止',error:'需要处理'};
  function notice(message,error=false){$('notice').textContent=message;$('notice').classList.toggle('error',error);$('notice').hidden=false;}
  function availability(){
    const blocked=!connected||pending||!!state?.busy||exiting;
    ['run-fields','settings-fields','start','use-cache','shutdown'].forEach(id=>$(id).disabled=blocked);
    $('cache').disabled=blocked;$('stop').disabled=!connected||pending||!state?.busy||state.status==='stopping';
  }
  function modeChanged(){
    const mode=$('mode').value,stage2=mode==='stage2';
    $('dates').hidden=stage2;$('stage1-inputs').hidden=stage2;$('stage2-inputs').hidden=!stage2;
    $('force-wrap').hidden=mode==='stage1';$('date-from').required=$('date-to').required=!stage2;
    $('start').textContent=stage2?'开始工单核对':mode==='all'?'开始一站式运行':'开始邮件解析';
    $('mode-help').textContent=stage2?'只查询所选文件中符合条件的业务明细。':mode==='all'?'沿用桌面一站式流程：解析后继续查询。需要先人工复核时，请选择阶段一。':'仅处理所选日期范围内的邮件；不会扫描全部邮件。';
  }
  async function request(path,payload={},upload=null){
    const controller=new AbortController(),timer=setTimeout(()=>controller.abort(),45000);
    try{
      const response=await fetch('/api/run/'+path,{method:'POST',cache:'no-store',signal:controller.signal,
        headers:{'X-Run-Token':csrf,'Content-Type':upload?'application/octet-stream':'application/json'},body:upload||JSON.stringify(payload)});
      const data=await response.json();if(!response.ok||!data.ok)throw Error(data.error||'操作失败');return data;
    }finally{clearTimeout(timer);}
  }
  async function action(fn){
    if(pending)return;pending=true;availability();
    try{const result=await fn();if(result?.message)notice(result.message);await poll();}
    catch(error){notice(error.name==='AbortError'?'操作响应超时，请查看任务状态后再决定是否重试。':error.message,true);}
    finally{pending=false;availability();}
  }
  function options(id,items,placeholder){
    const select=$(id),value=select.value;select.replaceChildren(new Option(placeholder,''));
    items.forEach(item=>select.add(new Option(item.label||item.name,item.token)));
    if(items.some(item=>item.token===value))select.value=value;
  }
  function render(data){
    state=data;csrf=data.csrf;connected=true;$('build').textContent=data.build;
    if(initial){const saved=data.session;$('mode').value=['stage1','stage2','all'].includes(saved.mode)?saved.mode:'stage1';
      $('date-from').value=saved.date_from;$('date-to').value=saved.date_to;$('force-live').checked=saved.force_live_query;
      $('email-address').value=data.credentials.email_address;$('wo-user').value=data.credentials.workorder_username;modeChanged();initial=false;}
    $('email-password').placeholder=data.credentials.email_password_set?'已配置，留空保持不变':'尚未配置，请填写';
    $('wo-password').placeholder=data.credentials.workorder_password_set?'已配置，留空保持不变':'尚未配置，请填写';
    for(const role of ['agent','internal','project'])$('file-'+role).textContent=data.imports[role]||'尚未导入';
    $('status').textContent=labels[data.status]||data.status;$('status').className='badge '+data.status;$('message').textContent=data.message;
    const[current,total]=data.progress;$('progress').value=total>0?Math.min(100,current/total*100):data.status==='completed'?100:0;
    $('progress-text').textContent=total>0?`已处理 ${current} / ${total}`:data.busy?'正在执行，详细进度见下方日志':data.status==='idle'?'未开始':labels[data.status];
    if(data.logs.length){logLines.push(...data.logs.map(item=>item.text));logLines=logLines.slice(-2000);
      const log=$('log'),top=log.scrollTop;log.textContent=logLines.join('\n');log.scrollTop=$('follow').checked?log.scrollHeight:top;}
    cursor=data.cursor;
    const nextCatalog=JSON.stringify(data.catalog);
    if(nextCatalog!==catalogKey){options('stage2-file',data.catalog.stage2,'请选择已确认导出的文件');options('cache',data.catalog.caches,'请选择阶段一日期缓存');catalogKey=nextCatalog;
      if(data.imports.stage2){const item=data.catalog.stage2.find(x=>x.name===data.imports.stage2);if(item&&!$('stage2-file').value)$('stage2-file').value=item.token;}}
    const nextResults=JSON.stringify(data.results);
    if(resultKey!==nextResults){$('results').replaceChildren();data.results.forEach(item=>{const a=document.createElement('a');a.href=item.url;a.className='button';a.textContent='下载 '+item.name;a.setAttribute('download','');$('results').append(a);});resultKey=nextResults;}
    availability();
  }
  async function poll(){
    if(polling||exiting)return;polling=true;const controller=new AbortController(),timer=setTimeout(()=>controller.abort(),10000);
    try{const response=await fetch('/api/run/state?after='+cursor,{cache:'no-store',signal:controller.signal});if(!response.ok)throw Error('本机服务未连接');render(await response.json());}
    catch(error){connected=false;$('status').textContent='连接中断';$('message').textContent='无法连接本机服务，请通过“启动网页版.bat”启动。连接恢复后自动继续显示；不会自动重启任务。';availability();}
    finally{clearTimeout(timer);polling=false;}
  }
  $('mode').addEventListener('change',modeChanged);
  $('run-form').addEventListener('submit',event=>{event.preventDefault();action(()=>request('start',{mode:$('mode').value,date_from:$('date-from').value,date_to:$('date-to').value,input_token:$('stage2-file').value,force_live_query:$('force-live').checked}));});
  $('stop').addEventListener('click',()=>action(()=>request('stop')));
  $('settings-form').addEventListener('submit',event=>{event.preventDefault();action(async()=>{const result=await request('settings',{email_address:$('email-address').value,email_password:$('email-password').value,workorder_username:$('wo-user').value,workorder_password:$('wo-password').value});$('email-password').value=$('wo-password').value='';return result;});});
  document.querySelectorAll('input[data-role]').forEach(input=>input.addEventListener('change',()=>{const file=input.files[0];if(!file)return;
    if(file.size>20*1024*1024){notice('文件超过 20 MB，请拆分参考表后导入。',true);input.value='';return;}
    action(async()=>{notice('正在校验并导入文件…');try{return await request('upload?role='+input.dataset.role+'&name='+encodeURIComponent(file.name),{},file);}finally{input.value='';}});
  }));
  $('refresh-files').addEventListener('click',()=>action(()=>request('refresh')));
  $('use-cache').addEventListener('click',()=>action(async()=>{if(!$('cache').value)throw Error('请先选择缓存日期');
    const result=await request('cache',{token:$('cache').value});$('cache-message').textContent=result.message;
    const response=await fetch('/api/run/state?after='+cursor,{cache:'no-store'});const data=await response.json();
    $('date-from').value=data.session.date_from;$('date-to').value=data.session.date_to;render(data);return result;
  }));
  $('shutdown').addEventListener('click',()=>{if(!confirm('退出本机服务后，运行控制台和复核页面将断开。确定退出吗？'))return;action(async()=>{const result=await request('shutdown');exiting=true;notice('本机服务已请求退出。下次双击“启动网页版.bat”即可重新打开。');return result;});});
  poll();setInterval(()=>poll(),1500);
})();
