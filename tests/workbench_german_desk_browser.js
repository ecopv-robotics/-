async page => {
  const checks=[],errors=[],requests=[];
  const check=(value,label)=>{if(!value)throw Error(label);checks.push(label)};
  const detail=(id,status='ready',program='法国包装法')=>({id,detail_number:id,status,fields:{company:'Example '+id+' LLC',agent:'Example Agent',country:program.startsWith('德国')?'德国':'法国',program,request:'注册'},events:[],evidence:[],confidence:'high'});
  const weee=id=>({...detail(id,'ready','德国WEEE'),weee:{enabled:true,status:'pending',confirmed:false,items:[{item_id:id+'I',brand:'Example',category:'小型设备',category_class:'5',weee_confirmed:false,source:'test source',extraction_method:'test method',candidates:['test candidate']}]}});
  const mail=(id,details)=>({id,mail_number:id,subject:'Anonymous '+id,sender:'fixture@example.test',recipient:'review@example.test',date:'2026-09-30 09:00:00',body:'Fixture 61a2ff37 LLC',raw_body:'Fixture 61a2ff37 LLC',attachments:'',details,status:'ready',company_count:details.length,project_count:details.length,detail_count:details.length});
  const state={ok:true,mails:[
    mail('REVIEW',[detail('R1','review'),detail('R2'),detail('R3','confirmed')]),
    mail('READY',[detail('P1')]),mail('DONE',[detail('C1','confirmed')]),
    mail('WEEE',[weee('W1')]),mail('BATTERY',[detail('B1','ready','德国电池法')]),
    mail('MIXED',[weee('W2'),detail('F1')]),
    {...mail('OLD',[detail('O1')]),date:'2026-09-29 09:00:00'}
  ],filtered_mails:[],missing_mails:[],projects:[],counts:{imap_read_total:7},history:{enabled:true},paths:{}};
  page.on('pageerror',error=>errors.push(error.message));
  await page.unroute('**/api/**');
  await page.route('**/api/**',async route=>{
    const request=route.request(),path=new URL(request.url()).pathname;
    const body=request.method()==='POST'?request.postDataJSON():null;
    requests.push({path,body});
    let result={ok:true};
    if(path==='/api/state')result=JSON.parse(JSON.stringify(state));
    else if(path==='/api/action'){
      const m=state.mails.find(m=>m.id===body.mail_id),d=m?.details.find(d=>d.id===body.record_id);
      if(body.action==='confirm_detail'&&d)d.status='confirmed';
      if(body.action==='save_battery_draft'&&d)d.battery={enabled:true,items:body.battery_items,requires_review:true,confirmed:false,status:'pending'};
      if(body.action==='confirm_battery_item'&&d){d.battery={enabled:true,items:body.battery_items.map(i=>({...i,battery_confirmed:true})),requires_review:true,confirmed:true,status:'confirmed'};d.status='confirmed';result.battery=d.battery}
      if(body.action==='save_weee_draft'&&d)d.weee.items=body.weee_items;
      if(body.action==='delete_weee_item'&&d)d.weee.items=body.weee_items.filter(item=>item.item_id!==body.weee_item_id);
      if(body.action==='confirm_weee_item'&&d){d.weee.confirmed=true;d.weee.status='confirmed';d.weee.items.forEach(i=>i.weee_confirmed=true);result.weee_status='confirmed';result.weee_item_confirmed=true}
      result.message='Anonymous test saved';
    }
    await route.fulfill({status:200,contentType:'application/json',body:JSON.stringify(result)});
  });
  await page.goto('http://127.0.0.1:8765/?test=german-r10');
  await page.evaluate(()=>localStorage.clear());await page.reload();
  await page.setViewportSize({width:1600,height:1000});
  await page.waitForFunction(()=>typeof app!=='undefined'&&app.data?.mails?.length===7);
  await page.locator('#date-start').fill('2026-09-30');await page.locator('#date-start').dispatchEvent('change');
  await page.locator('#date-end').fill('2026-09-30');await page.locator('#date-end').dispatchEvent('change');
  const select=async(view,id,filter='all')=>page.evaluate(({view,id,filter})=>{app.view=view;app.filter=filter;app.mailId=id;app.detailId=app.data.mails.find(m=>m.id===id)?.details[0]?.id;render()},{view,id,filter});
  const ids=()=>page.evaluate(()=>queueMails().map(m=>m.id).sort().join(','));
  await select('inbox','REVIEW');
  check(await ids()==='READY,REVIEW','普通询单只含当天未完成且非德国专项邮件');
  check(!await page.locator('[data-filter="confirmed"]').isVisible(),'询单已完成筛选已隐藏');
  check((await page.locator('[data-filter="review"]').innerText()).includes('1'),'待复核按整封邮件去重计数');
  check((await page.locator('[data-filter="ready"]').innerText()).includes('1'),'待确认按整封邮件去重计数');
  await page.locator('[data-filter="review"]').click();check(await ids()==='REVIEW','混合状态邮件仅在待复核');
  await page.locator('[data-filter="ready"]').click();check(await ids()==='READY','待确认与待复核互斥');
  await select('weee','BATTERY');
  check(await ids()==='BATTERY,MIXED,WEEE','德国专项包含WEEE、电池法和混合项目');
  check(await page.locator('[data-battery-detail="B1"] [data-confirm-battery-item]').count()===1,'电池法使用独立品类确认');
  check(await page.locator('#detail [data-weee-class]').count()===0,'电池法不套用WEEE六分类');
  check(await page.locator('#detail [data-business-action="delete"]').count()===0,'德国专项移除删除业务按钮');
  check((await page.locator('#evidence').innerText()).includes('Fixture 61a2ff37 LLC'),'电池法保留原始邮件证据');
  for(const [key,value] of [['brand','Example Battery'],['category','Original battery type'],['category_class','Manual battery class']])await page.locator('[data-battery-detail="B1"] [data-battery-field="'+key+'"]').fill(value);
  await page.locator('[data-battery-detail="B1"] [data-confirm-battery-item]').click();
  await page.waitForFunction(()=>app.data.mails.find(m=>m.id==='BATTERY').details[0].status==='confirmed');
  check(requests.some(r=>r.body?.action==='confirm_battery_item'&&r.body.record_id==='B1'),'电池法确认提交正确业务编号');
  await select('weee','MIXED');
  check(await page.locator('[data-weee-detail="W2"]').count()===1&&await page.locator('[data-detail="F1"]').count()===1,'混合邮件保留其他国家项目');
  await page.locator('[data-business-id="F1"] [data-business-action="confirm"]').click();
  await page.waitForFunction(()=>app.data.mails.find(m=>m.id==='MIXED').details.find(d=>d.id==='F1').status==='confirmed');
  check(requests.some(r=>r.body?.record_id==='F1'&&r.body.action==='confirm_detail'),'混合邮件中的其他项目可确认');
  await select('weee','WEEE');
  const contents=await page.locator('#detail').innerText();
  for(const text of ['提取方式','候选：','证据：','字段流程','删除此业务','品类明细','每个品牌/原始品类'])check(!contents.includes(text),'精简说明 '+text);
  const row=page.locator('[data-weee-detail="W1"] .weee-item-form');
  const tops=await row.locator('input,select').evaluateAll(nodes=>nodes.map(n=>Math.round(n.getBoundingClientRect().top)));
  check(tops.length===3&&tops[0]===tops[1]&&tops[2]>tops[1],'德国分类单独一行');
  const buttonTops=await row.locator('.weee-item-actions button').evaluateAll(nodes=>nodes.map(n=>Math.round(n.getBoundingClientRect().top)));
  check(new Set(buttonTops).size===1&&buttonTops.length===3,'保存、删除项、新增项三个按钮同一行');
  await row.locator('[data-weee-brand]').fill('EditedExample');
  await page.waitForTimeout(650);
  check(requests.some(r=>r.body?.action==='save_weee_draft'&&r.body.weee_items?.[0]?.brand==='EditedExample'),'输入自动保存草稿仍有效');
  await page.locator('[data-add-weee-item="W1"]').click();
  check(await page.locator('[data-weee-detail="W1"] .weee-item-form').count()===2,'可新增品牌品类项');
  await page.evaluate(()=>window.confirm=()=>true);
  await page.locator('[data-delete-weee-item="W1"]').last().click();
  await page.waitForFunction(()=>document.querySelectorAll('[data-weee-detail="W1"] .weee-item-form').length===1);
  check(await page.locator('[data-weee-detail="W1"] .weee-item-form').count()===1,'可删除单个品类项');
  await page.locator('[data-confirm-weee-item="W1"]').click();
  await page.waitForFunction(()=>app.data.mails.find(m=>m.id==='WEEE').details[0].weee.confirmed);
  check(requests.some(r=>r.body?.action==='confirm_weee_item'&&r.body.record_id==='W1'),'WEEE单项确认提交正确业务编号');
  check((await page.locator('[data-weee-detail="W1"] .weee-detail-head').innerText()).includes('已确认'),'WEEE确认状态正常显示');
  const width=selector=>page.locator(selector).evaluate(e=>e.getBoundingClientRect().width);
  const drag=async(type,delta)=>{const r=await page.locator('[data-splitter="'+type+'"]').boundingBox();await page.mouse.move(r.x+r.width/2,r.y+100);await page.mouse.down();await page.mouse.move(r.x+r.width/2+delta,r.y+100,{steps:8});await page.mouse.up()};
  const qw=await width('.queue');await drag('queue',45);check(Math.abs(await width('.queue')-qw-45)<3,'队列分隔条仍可左右拖动');
  const ew=await width('.evidence');await drag('evidence',45);check(Math.abs(await width('.evidence')-ew+45)<3,'证据分隔条仍可左右拖动');
  const selection=await page.locator('#evidence [data-evidence-source="body"]').evaluate(e=>{const r=document.createRange();r.selectNodeContents(e);const s=getSelection();s.removeAllRanges();s.addRange(r);return s.toString()});
  check(selection.includes('Fixture 61a2ff37 LLC'),'原始正文仍可选择复制');
  await page.evaluate(()=>getSelection().removeAllRanges());
  await select('inbox','READY');
  await page.locator('[data-business-id="P1"] [data-business-action="confirm"]').click();
  await page.waitForFunction(()=>!queueMails().some(m=>m.id==='READY'));
  check(await ids()==='REVIEW','整封确认后立即从询单待处理移除');
  check(await page.evaluate(()=>app.data.mails.some(m=>m.id==='READY')),'已完成业务仍保存在数据中');
  check(errors.length===0,'页面无脚本异常 '+JSON.stringify(errors));
  return {passed:checks.length,checks,productionWrites:0};
}
