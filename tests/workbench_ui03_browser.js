async page => {
  const checks=[],errors=[],requests=[];
  const check=(value,label)=>{if(!value)throw Error(label);checks.push(label)};
  const fields={company:'Example Trading LLC',agent:'示例代理',country:'法国',program:'法国包装法',request:'注册'};
  const makeDetail=id=>({id,detail_number:id,status:'ready',fields:{...fields,company:'Example '+id+' LLC'},evidence:[],events:[],confidence:'high',source:'匿名测试'});
  const details=Array.from({length:5},(_,i)=>makeDetail('D'+(i+1)));
  const mail={id:'M1',mail_number:'M1',subject:'测试邮件 — 五条业务',sender:'fixture@example.test',recipient:'review@example.test',date:'2026-09-30 09:00:00',body:'原文可复制 Example Evidence LLC',raw_body:'原文可复制 Example Evidence LLC',attachments:'申请表.xlsx',details,company_count:5,project_count:1,detail_count:5,status:'ready'};
  const other={...mail,id:'M2',mail_number:'M2',subject:'第二封邮件',details:[makeDetail('X1')]};
  const weee={...makeDetail('W1'),fields:{...fields,program:'德国WEEE',country:'德国'},weee:{enabled:true,status:'pending',confirmed:false,items:[{item_id:'WITEM1',brand:'Example',category:'小型设备',category_class:'5',category_class_name:'小型设备',weee_confirmed:false}]}};
  const wm={...mail,id:'MW',mail_number:'MW',subject:'德国 WEEE 测试',details:[weee]};
  const state={ok:true,mails:[mail,other,wm],filtered_mails:[{id:'F1',subject:'已过滤测试邮件',date:mail.date,body:'过滤原文',reason:'测试',sender:mail.sender,attachments:''}],missing_mails:[other],projects:[{name:'法国包装法',country:'法国'}],counts:{imap_read_total:4},history:{enabled:true,completed_count:0,unfinished_count:7},paths:{}};
  let failEdit=false;
  page.on('pageerror',error=>errors.push(error.message));
  await page.route('**/api/**',async route=>{
    const request=route.request(),url=new URL(request.url()),body=request.postDataJSON?.();
    requests.push({path:url.pathname,body});
    let result={ok:true};
    if(url.pathname==='/api/state')result=JSON.parse(JSON.stringify(state));
    else if(url.pathname==='/api/action'){
      if(body.action==='edit'&&failEdit){await route.fulfill({status:422,contentType:'application/json',body:JSON.stringify({ok:false,error:'模拟保存失败'})});return}
      const m=state.mails.find(m=>m.id===body.mail_id),d=m?.details.find(d=>d.id===body.record_id);
      if(body.action==='edit'&&d){d.fields={...d.fields,...body.fields};d.status='review'}
      if(body.action==='confirm_detail'&&d)d.status='confirmed';
      if(body.action==='return'&&d)d.status='returned';
      if(body.action==='delete_project'&&m)m.details=m.details.filter(d=>d.id!==body.record_id);
      if(body.action==='delete_mail')state.mails=state.mails.filter(m=>m.id!==body.mail_id);
      if(body.action==='add_project'){m.details.push({...makeDetail('NEW'),fields:body.fields});result.record_id='NEW'}
      if(body.action==='bulk_confirm'){result.skipped_count=0;result.confirmed_details=0;for(const id of body.mail_ids){const target=state.mails.find(m=>m.id===id);target?.details.forEach(d=>{d.status='confirmed';result.confirmed_details++})}}
      if(body.action==='confirm_weee_item'&&d){d.weee.confirmed=true;d.weee.status='confirmed';result.weee_status='confirmed';result.weee_item_confirmed=true}
      result.message='测试保存成功';
    }else if(url.pathname==='/api/export')result.message='测试导出成功';
    else if(url.pathname==='/api/attachment-preview')result={ok:true,kind:'workbook',filename:'申请表.xlsx',sheets:[{name:'EPR申请表',row_count:2,column_count:2,rows:[{row_number:1,cells:['公司英文名称','Example Evidence LLC']},{row_number:2,cells:['项目','法国包装法']}],next_offset:null}]};
    else if(!['/api/health','/api/run/state','/api/version'].includes(url.pathname))throw Error('未预期接口 '+url.pathname);
    await route.fulfill({status:200,contentType:'application/json',body:JSON.stringify(result)});
  });
  await page.goto('http://127.0.0.1:8765/?ui=03-test');
  await page.evaluate(()=>{localStorage.clear()});
  await page.reload();
  await page.setViewportSize({width:1600,height:1000});
  await page.locator('.business-actions').first().waitFor();
  const action=(id,kind)=>page.locator('[data-business-id="'+id+'"] [data-business-action="'+kind+'"]');
  check(await page.locator('#agent').count()===0,'独立确认代理入口已移除');
  check(await page.locator('#detail .detail-head [data-business-action]').count()===0,'邮件头不混入业务按钮');
  check(await action('D5','edit').count()===1&&await action('D5','confirm').count()===1,'每条业务独立修正确认');
  check(await page.locator('[data-fill-evidence]').count()===0,'移除填入选中文字');
  check(await page.locator('#f-reason').count()===0,'移除修改原因');
  check(await page.locator('.work').evaluate(e=>e.getBoundingClientRect().y)<180,'工作区顶端小于180px');
  const width=selector=>page.locator(selector).evaluate(e=>e.getBoundingClientRect().width);
  const drag=async(type,delta)=>{const r=await page.locator('[data-splitter="'+type+'"]').boundingBox();await page.mouse.move(r.x+r.width/2,r.y+120);await page.mouse.down();await page.mouse.move(r.x+r.width/2+delta,r.y+120,{steps:8});await page.mouse.up()};
  const qw=await width('.queue');await drag('queue',60);check(Math.abs(await width('.queue')-qw-60)<3,'左拖动改变队列宽度');
  const ew=await width('.evidence');await drag('evidence',60);check(Math.abs(await width('.evidence')-ew+60)<3,'右拖动改变证据宽度');
  const saved=[await width('.queue'),await width('.evidence')];await page.reload();await page.locator('.business-actions').first().waitFor();
  check(Math.abs(await width('.queue')-saved[0])<3&&Math.abs(await width('.evidence')-saved[1])<3,'刷新恢复两栏宽度');
  await page.locator('[data-splitter="queue"]').press('ArrowRight');
  const keyboard=await width('.queue');await page.reload();await page.locator('.business-actions').first().waitFor();
  check(Math.abs(await width('.queue')-keyboard)<3,'键盘调整也保存');
  await drag('queue',900);check(await width('.detail')>=299,'极限拖动不挤没业务栏');
  await page.locator('[data-splitter="queue"]').dblclick();
  await page.locator('#nav-toggle').click();
  check(await page.evaluate(()=>document.querySelector('.side').getBoundingClientRect().right<=document.querySelector('.queue').getBoundingClientRect().left),'导航展开不遮挡队列');
  await page.locator('#nav-toggle').click();
  await action('D4','confirm').click();
  await page.waitForFunction(()=>app.data.mails[0].details.find(d=>d.id==='D4').status==='confirmed');
  check(requests.filter(r=>r.body?.action==='confirm_detail').at(-1).body.record_id==='D4','确认目标是所点击的第四条');
  check(await page.evaluate(()=>app.detailId)==='D4','确认后保持当前业务');
  const scroll=await page.locator('#detail .detail-body').evaluate(e=>e.scrollTop);
  await page.waitForTimeout(2000);
  check(await page.evaluate(()=>app.detailId)==='D4'&&Math.abs(await page.locator('#detail .detail-body').evaluate(e=>e.scrollTop)-scroll)<5,'后台刷新保留选中项与滚动位置');
  await action('D3','edit').click();
  check(await width('#modal-overlay .modal')===380,'修正窗口380px紧凑宽度');
  check(await page.locator('#modal-overlay .modal').evaluate(e=>e.getBoundingClientRect().height)<470,'修正窗口高度紧凑');
  const box=await page.locator('#modal-overlay .modal').boundingBox(),head=await page.locator('#modal-overlay .modal-head').boundingBox();
  await page.mouse.move(head.x+60,head.y+20);await page.mouse.down();await page.mouse.move(head.x-160,head.y+30,{steps:6});await page.mouse.up();
  check((await page.locator('#modal-overlay .modal').boundingBox()).x<box.x-100,'修正窗口标题栏可拖动');
  check(await page.evaluate(()=>getComputedStyle(document.querySelector('#modal-overlay')).pointerEvents)==='none','弹窗外证据可以操作');
  const selection=await page.locator('#evidence [data-evidence-source="body"]').evaluate(e=>{const r=document.createRange();r.selectNodeContents(e);const s=window.getSelection();s.removeAllRanges();s.addRange(r);return s.toString()});
  check(selection.includes('Example Evidence LLC'),'原始邮件正文可选中文字');
  await page.locator('#f-company').fill('Fixture 4025f229 LLC');
  await page.locator('#f-agent').fill('新代理');
  // Changing the selected card while the editor is open must not change its save target.
  await page.evaluate(()=>{app.detailId='D1';renderDetailKeepingScroll()});
  failEdit=true;await page.locator('#modal-save').click();
  await page.locator('#modal-error').filter({hasText:'保存未完成'}).waitFor();
  check(await page.locator('#modal-overlay').isVisible()&&await page.locator('#f-company').inputValue()==='Fixture 4025f229 LLC','保存失败保留窗口和输入');
  failEdit=false;await page.locator('#modal-save').click();await page.locator('#modal-overlay').waitFor({state:'hidden'});
  const edit=requests.filter(r=>r.body?.action==='edit').at(-1).body;
  check(edit.record_id==='D3'&&edit.fields.company==='Fixture 4025f229 LLC'&&edit.fields.agent==='新代理','修正代理与公司一次保存且不串业务');
  await action('D5','return').click();
  check(requests.filter(r=>r.body?.action==='return').at(-1).body.record_id==='D5','退回复核作用于所在业务');
  await page.evaluate(()=>{window.confirm=()=>true});
  await action('D2','delete').click();
  await page.waitForFunction(()=>!app.data.mails[0].details.some(d=>d.id==='D2'));
  check(requests.filter(r=>r.body?.action==='delete_project').at(-1).body.record_id==='D2','删除单条业务不删除邮件');
  await page.locator('#add-project').click();await page.locator('#f-program').fill('法国包装法');await page.locator('#f-country').fill('法国');await page.locator('#f-request').fill('注册');
  await page.locator('#modal-save').click();await page.locator('#modal-overlay').waitFor({state:'hidden'});
  check(requests.some(r=>r.body?.action==='add_project'&&r.body.mail_id==='M1'),'邮件级新增业务接口保留');
  await page.locator('#confirm').click();
  await page.waitForFunction(()=>app.data.mails[0].details.every(d=>d.status==='confirmed'));
  check(requests.filter(r=>r.body?.action==='bulk_confirm').at(-1).body.mail_ids.join()==='M1','整理完成使用整封邮件范围');
  await page.locator('#date-start').fill('2026-09-30');await page.locator('#date-start').dispatchEvent('change');
  await page.locator('#date-end').fill('2026-09-30');await page.locator('#date-end').dispatchEvent('change');
  await page.locator('#export').click();
  check(requests.filter(r=>r.path==='/api/export').at(-1).body.date_start==='2026-09-30','导出保留当前邮件日期范围');
  await page.locator('[data-view="missing"]').click();
  check(await page.locator('.business-actions').filter({hasText:'同步工单'}).count()===1,'漏单同步和自动重查保留在业务旁');
  await page.locator('[data-view="filtered"]').click();
  check(await page.locator('#restore-review').isVisible(),'已过滤邮件恢复入口保留');
  await page.locator('[data-view="weee"]').click();
  check(await page.locator('[data-confirm-weee-item]').count()===1,'德国WEEE单品牌品类确认保留');
  check(await page.locator('[data-delete-weee-item]').count()===1,'德国WEEE单项删除保留');
  check(await page.locator('.weee-detail-head [data-business-action="edit"]').count()===1,'WEEE业务也可修正代理等字段');
  await page.locator('[data-view="inbox"]').click();
  for(const viewport of [{width:1440,height:900},{width:1280,height:800},{width:1024,height:900},{width:768,height:900},{width:390,height:844}]){
    await page.setViewportSize(viewport);await page.waitForTimeout(80);
    check(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth+1),'无页面横向溢出 '+viewport.width);
    check(await page.locator('#detail').evaluate(e=>e.getBoundingClientRect().width)>250,'业务面板可见 '+viewport.width);
  }
  await page.setViewportSize({width:1600,height:1000});
  await page.screenshot({path:'B:/Codex_EcoPV_workspace/output/workbench-ui03-production-fixture.png'});
  check(errors.length===0,'页面脚本无异常 '+JSON.stringify(errors));
  return {passed:checks.length,checks,requestCount:requests.length,productionWrites:0};
}
