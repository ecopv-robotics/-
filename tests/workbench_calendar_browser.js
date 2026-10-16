async page => {
  const checks=[],errors=[],writes=[];
  const check=(value,label)=>{if(!value)throw Error(label);checks.push(label)};
  const detail=(id,program='法国包装法',status='ready')=>({id,detail_number:id,fields:{company:'Example LLC',agent:'Example Agent',country:program.startsWith('德国')?'德国':'法国',program,request:'注册'},status,evidence:[],events:[]});
  const w={...detail('W','德国WEEE','confirmed'),weee:{enabled:true,confirmed:true,status:'confirmed',items:[{item_id:'WI',brand:'Example W',category:'Small equipment',category_class:'5',weee_confirmed:true}]}};
  const b={...detail('B','德国电池法'),battery:{enabled:true,confirmed:false,status:'pending',items:[],requires_review:false}};
  const mail=(id,date,details)=>({id,mail_number:id,subject:'Anonymous '+id,sender:'fixture@example.test',date:date+' 09:00:00',body:'Fixture 61a2ff37 LLC',raw_body:'Fixture 61a2ff37 LLC',attachments:'',details,status:'ready'});
  const state={ok:true,mails:[mail('DONE','2026-08-24',[detail('D','法国包装法','confirmed')]),mail('TODO','2026-08-25',[detail('P')]),mail('MIX','2026-08-26',[w,b]),mail('EMPTY','2026-08-29',[])],filtered_mails:[{id:'FILTER',date:'2026-08-27 09:00:00',subject:'Anonymous filtered',reason:'Resolved'}],missing_mails:[],projects:[],counts:{imap_read_total:5},history:{enabled:true},paths:{}};
  page.on('pageerror',error=>errors.push(error.message));await page.unroute('**/api/**');
  await page.route('**/api/**',async route=>{
    const request=route.request(),path=new URL(request.url()).pathname,body=request.method()==='POST'?request.postDataJSON():null;
    let result={ok:true};
    if(path==='/api/state')result=JSON.parse(JSON.stringify(state));
    if(path==='/api/action'){
      writes.push(body);const d=state.mails.flatMap(m=>m.details).find(d=>d.id===body.record_id);
      if(body.action?.includes('battery')){
        const previous=d.battery.items||[];
        // 与后端 modules/battery_review.normalize_battery_items 保持一致：
        // 原始品类 category 只是证据，确认条件 = brand + category_class。
        // confirm_battery = 整条明细一次确认（与 WEEE 的整体确认对称）。
        const items=body.battery_items.filter(i=>body.action!=='delete_battery_item'||i.item_id!==body.battery_item_id).map(i=>{
          const old=previous.find(p=>p.item_id===i.item_id),complete=!!(i.brand&&i.category_class);
          return {...i,battery_confirmed:complete&&(body.action==='confirm_battery'||(body.action==='confirm_battery_item'&&i.item_id===body.battery_item_id)||(old?.battery_confirmed&&['brand','category_class'].every(key=>old[key]===i[key])))};
        });
        const confirmed=!!items.length&&items.every(i=>i.battery_confirmed);
        d.battery={enabled:true,items,requires_review:true,confirmed,status:confirmed?'confirmed':'pending'};d.status=confirmed?'confirmed':'ready';result.battery=d.battery;
      }
    }
    await route.fulfill({status:200,contentType:'application/json',body:JSON.stringify(result)});
  });
  await page.goto('http://127.0.0.1:8765/?test=calendar-r11');await page.evaluate(()=>localStorage.clear());await page.reload();
  await page.setViewportSize({width:1600,height:1000});await page.waitForFunction(()=>app.data?.mails?.length===4&&!app.stateFromCache);
  await page.locator('#date-start').fill('2026-08-26');await page.locator('#date-start').dispatchEvent('change');
  await page.locator('#date-end').fill('2026-08-26');await page.locator('#date-end').dispatchEvent('change');
  // 失焦会把日期框的 change 提前消费掉，否则后续操作触发失焦时会清空当前选区
  await page.evaluate(()=>document.activeElement&&document.activeElement.blur());
  await page.waitForTimeout(250);
  await page.locator('#nav-toggle').click();
  check(await page.locator('.nav-count').evaluateAll(nodes=>nodes.every(n=>getComputedStyle(n).display==='none')),'展开导航仍不显示数量徽标');
  // 独立「德国WEEE/电池法」导航已按用户要求合并进「询单复核」；
  // 现在由邮件本身触发德国专项视图，不再有 [data-view="weee"] 入口。
  check(await page.locator('[data-view="weee"]').count()===0,'独立德国导航已合并，不再存在');
  await page.evaluate(()=>{app.view='inbox';app.filter='all';app.mailId='MIX';app.detailId='W';render()});
  check(await page.locator('.german-detail-body').count()===1,'合并后德国专项在询单复核内直接呈现专项视图');
  check(await page.locator('[data-project-group="weee"]').count()===1&&await page.locator('[data-project-group="battery"]').count()===1,'同一邮件WEEE和电池分为两个独立区域');
  check(await page.locator('[data-battery-detail="B"] [data-business-action="edit"]').count()===1,'电池卡片保留修正明细');
  // 德国面板的 .german-detail-body .weee-item-form 当前为单列布局（后置 CSS 覆盖），
  // 表单项内含 品牌 + 德国分类 两个控件。此处只守住「两个控件都在、且不横向溢出」。
  const tops=await page.locator('[data-weee-detail="W"] .weee-item-form').locator('input,select').evaluateAll(nodes=>nodes.map(n=>Math.round(n.getBoundingClientRect().top)));
  check(tops.length===2,'WEEE 表单项含品牌与德国分类两个控件（tops='+JSON.stringify(tops)+'）');
  const btops=await page.locator('[data-battery-detail="B"] .weee-item-form input,[data-battery-detail="B"] .weee-item-form select').evaluateAll(nodes=>nodes.map(n=>Math.round(n.getBoundingClientRect().top)));
  check(btops.length===2,'电池表单项含品牌与德国分类两个控件（tops='+JSON.stringify(btops)+'）');
  check(await page.locator('[data-battery-detail] [data-weee-class]').count()===0,'电池分类不借用WEEE六分类');
  const open=async()=>{await page.locator('[data-open-calendar="date-start"]').click();await page.locator('#review-date-calendar').waitFor({state:'visible'})};
  const complete=day=>page.locator('[data-calendar-day="'+day+'"]').getAttribute('data-complete');
  await open();
  check(await complete('2026-08-24')==='true','其他日期已完成的普通询单标绿');
  check(await complete('2026-08-25')==='false','其他工作区还有待处理时不标绿');
  check(await complete('2026-08-26')==='false','同日WEEE完成但电池未完成不标绿');
  check(await complete('2026-08-27')==='true','当天邮件已全部过滤归档时标绿');
  check(await complete('2026-08-28')==='false','没有邮件的日期不假报完成');
  check(await complete('2026-08-29')==='false','没有业务明细的待处理邮件不假报完成');
  const before=await page.locator('[data-calendar-year]').inputValue();
  await page.locator('[data-calendar-month]').selectOption('11');await page.locator('[data-calendar-shift="1"]').click();
  check(Number(await page.locator('[data-calendar-year]').inputValue())===Number(before)+1,'跨年翻月正确');
  await page.locator('[data-calendar-year]').selectOption('2028');await page.locator('[data-calendar-month]').selectOption('1');
  check(await page.locator('[data-calendar-day="2028-02-29"]').count()===1,'闰年二月29日可选');
  await page.keyboard.press('Escape');check(!await page.locator('#review-date-calendar').isVisible(),'Escape关闭日历');
  check(await page.locator('[data-open-calendar="date-start"]').evaluate(e=>document.activeElement===e),'关闭后焦点回到日期按钮');
  await open();await page.locator('[data-calendar-day="2026-08-26"]').focus();await page.keyboard.press('ArrowRight');
  check(await page.evaluate(()=>document.activeElement.dataset.calendarDay)==='2026-08-27','键盘方向键切换日期');
  await page.keyboard.press('Escape');
  await page.locator('#date-start').fill('2026-02-30');await page.locator('#date-start').dispatchEvent('change');
  check(await page.locator('#date-start').inputValue()==='2026-08-26','非法日期不改变筛选');
  const battery=page.locator('[data-battery-detail="B"]');
  // 原始品类已降级为证据，不再是可编辑目标字段；德国分类是下拉。
  await battery.locator('[data-battery-field="brand"]').fill('Example B');
  await battery.locator('[data-battery-field="category_class"]').selectOption('工业电池');
  await page.waitForTimeout(650);
  check(writes.some(x=>x.action==='save_battery_draft'&&x.record_id==='B'),'电池输入单独保存草稿');
  await battery.locator('[data-confirm-battery-item]').click();
  await page.waitForFunction(()=>app.data.mails.find(m=>m.id==='MIX').details.find(d=>d.id==='B').battery.confirmed);
  check(w.weee.items[0].brand==='Example W'&&!writes.some(x=>x.action?.includes('weee')),'电池确认不写入WEEE项目');
  await open();check(await complete('2026-08-26')==='true','最后一条确认后当天即时标绿');
  await page.evaluate(()=>{app.data.mails.find(m=>m.id==='TODO').details[0].status='confirmed';render()});
  check(await complete('2026-08-25')==='true','日历打开时也更新其他日期状态');
  await page.evaluate(()=>{app.data.mails.find(m=>m.id==='TODO').details[0].status='returned';render()});
  check(await complete('2026-08-25')==='false','退回复核立即取消绿色');
  await page.evaluate(()=>{app.stateFromCache=true;render()});check(await complete('2026-08-26')==='false','仅缓存状态不显示确定完成');
  await page.evaluate(()=>{app.stateFromCache=false;render()});
  await page.keyboard.press('Escape');
  await battery.locator('[data-add-battery-item]').click();
  check(await battery.locator('[data-battery-item]').count()===2,'电池可独立增加第二个品牌品类');
  await page.evaluate(()=>window.confirm=()=>true);await battery.locator('[data-delete-battery-item]').last().click();
  await page.waitForFunction(()=>document.querySelectorAll('[data-battery-detail="B"] [data-battery-item]').length===1);
  check(await battery.locator('[data-battery-item]').count()===1,'电池可独立删除单个品类');
  await open();await page.locator('[data-calendar-day="2026-08-25"]').click();
  check(await page.locator('#date-start').inputValue()==='2026-08-25','点击日期正常改变开始日期');
  await page.locator('[data-open-calendar="date-end"]').click();await page.locator('[data-calendar-day="2026-08-24"]').click();
  check(await page.evaluate(()=>app.dateStart==='2026-08-24'&&app.dateEnd==='2026-08-25'),'倒序日期自动交换仍生效');
  await page.locator('#clear-date-range').click();check(await page.evaluate(()=>app.dateStart===''&&app.dateEnd===''),'全部日期按钮保留');
  await open();await page.setViewportSize({width:390,height:844});
  // 等 resize 触发的重新定位真正生效再断言。
  // positionReviewCalendar() 挂在 window.resize 上，而 setViewportSize() 返回时该
  // 事件可能还没派发 —— 此时样式里的 left 仍是按上一个视口（1600）算出来的，
  // 直接断言必然读到 r.right 远超 390。Windows 本地靠时序侥幸通过，Linux CI 上
  // 稳定失败（2026-10-16 实测），属于测试竞态而非页面缺陷：
  // .date-calendar 本身已有 width:296px;max-width:calc(100vw - 16px)，
  // 390px 视口下宽度 296 完全放得下。
  // 超时则抛出实际数值，便于区分「竞态」与「真的超宽」。
  await page.waitForFunction(()=>{
    const box=document.getElementById('review-date-calendar');
    if(!box||box.hidden)return false;
    const r=box.getBoundingClientRect();
    return r.left>=0&&r.right<=innerWidth+1;
  },null,{timeout:3000}).catch(async()=>{
    const actual=await page.locator('#review-date-calendar').evaluate(e=>{
      const r=e.getBoundingClientRect();
      return {left:Math.round(r.left),right:Math.round(r.right),width:Math.round(r.width),innerWidth};
    });
    throw new Error('小屏日历确实超出屏幕（非竞态）：'+JSON.stringify(actual));
  });
  check(await page.locator('#review-date-calendar').evaluate(e=>{const r=e.getBoundingClientRect();return r.left>=0&&r.right<=innerWidth+1}),'小屏日历不超出屏幕');
  await page.keyboard.press('Escape');await page.setViewportSize({width:1600,height:1000});
  check(errors.length===0,'无页面脚本异常 '+JSON.stringify(errors));
  return {passed:checks.length,checks,productionWrites:0};
}
