async page=>{
  const checks=[],errors=[],calls=[];
  const assert=(ok,label)=>{if(!ok)throw Error(label);checks.push(label)};
  const fields={company:'Alpha Example LLC',agent:'正文独有代理',country:'标题独有国家',program:'产品说明.pdf',request:'注册'};
  const sheet=(company,offset=0,name='EPR申请表')=>({sheet_name:name,row_count:80,column_count:2,offset,next_offset:offset<40?40:null,complete:true,
    rows:Array.from({length:40},(_,i)=>({row_number:offset+i+1,cells:offset+i===0?['EPR 合规注册申请表','']:offset+i===2?['公司英文名称\nCompany name (in English)',company]:['字段 '+(offset+i+1),'原文 '+(offset+i+1)]}))});
  const detail={id:'D1',detail_number:'D1',fields,status:'ready',evidence:[],events:[],attachment_evidence:[{filename:'三家公司.zip',sheets:[{...sheet(fields.company),member_path:'Alpha/EPR申请表.xlsx'}]}]};
  const file={filename:'三家公司.zip',token:'fixture-archive.zip',attachment_index:0};
  const pdf={filename:'产品说明.pdf',token:'fixture-doc.pdf',attachment_index:1};
  const mail={id:'M-TABS',mail_number:'M-TABS',subject:'标题独有国家 / Alpha Example LLC',date:'2026-09-30 08:00:00',sender:'source@example.test',recipient:'review@example.test',body:'正文独有代理。公司 Alpha Example LLC。',raw_body:'正文独有代理。公司 Alpha Example LLC。',attachments:'三家公司.zip；产品说明.pdf',attachment_files:[file,pdf],details:[detail],status:'ready'};
  const state={ok:true,mails:[mail],filtered_mails:[],missing_mails:[],projects:[],counts:{},history:{},paths:{}};
  const members=[
    {path:'Alpha/EPR申请表.xlsx',member_token:'chain1.alpha',filename:'EPR申请表.xlsx',previewable:true,downloadable:true},
    {path:'Beta/EPR申请表.xlsx',member_token:'chain1.beta',filename:'EPR申请表.xlsx',previewable:true,downloadable:true},
    {path:'nested.rar → Gamma/EPR申请表.xlsx',member_token:'chain1.nested-gamma',filename:'EPR申请表.xlsx',previewable:true,downloadable:true},
    {path:'附件/产品.pdf',member_token:'chain1.pdf',filename:'产品.pdf',previewable:false,downloadable:true},
    {path:'加密文件.xlsx',member_token:'chain1.encrypted',filename:'加密文件.xlsx',previewable:false,downloadable:false,reason:'成员已加密，当前无法读取'}
  ];
  let failBeta=true;
  page.on('pageerror',e=>errors.push(e.message));
  await page.route('**/api/**',async route=>{
    const req=route.request(),u=new URL(req.url());calls.push({path:u.pathname,member:u.searchParams.get('member'),offset:u.searchParams.get('offset'),method:req.method()});
    if(req.method()!=='GET')throw Error('证据浏览不应发送写请求');
    let result={ok:true};
    if(u.pathname==='/api/state')result=state;
    else if(u.pathname==='/api/attachment-preview'){
      const member=u.searchParams.get('member'),offset=Number(u.searchParams.get('offset')||0);
      if(u.searchParams.get('token')==='fixture-direct.xlsm'||u.searchParams.get('token')==='fixture-legacy.xls')result={ok:true,archive:false,sheets:[{sheet_name:'登记清单',row_count:1,column_count:2,rows:[{row_number:1,cells:['客户','普通表格原始值']}]}]};
      else if(member==='chain1.register')result={ok:true,archive:false,sheets:[{sheet_name:'客户登记',row_count:1,column_count:2,rows:[{row_number:1,cells:['客户','压缩包内登记原始值']}]}]};
      else if(!member)result={ok:true,archive:true,filename:file.filename,members,member_count:members.length,workbook_count:members.filter(m=>m.previewable).length};
      else if(member==='chain1.beta'&&failBeta){await route.fulfill({status:422,contentType:'application/json',body:JSON.stringify({ok:false,error:'模拟此文件读取失败'})});return}
      else{const company=member==='chain1.alpha'?'Alpha Example LLC':member==='chain1.beta'?'Beta Example LLC':'Gamma Example LLC';result={ok:true,archive:false,filename:'EPR申请表.xlsx',sheets:[sheet(company,offset),sheet(company,offset,'产品信息')]}}
    }else if(!['/api/health','/api/run/state'].includes(u.pathname))throw Error('未预期读取 '+u.pathname);
    await route.fulfill({status:200,contentType:'application/json',body:JSON.stringify(result)});
  });
  await page.goto('http://127.0.0.1:8765/?v=evidence-tabs-test');
  await page.evaluate(()=>localStorage.clear());await page.reload();
  await page.setViewportSize({width:1600,height:1000});
  await page.waitForFunction(()=>document.querySelectorAll('.source-book-tabs button').length===3);
  await page.waitForFunction(()=>!document.querySelector('.source-loading'));
  assert(await page.locator('.source-tabs button').count()===4,'四个证据来源页签');
  assert(await page.locator('.source-book-tabs button').count()===3,'压缩包三家公司分别显示且缓存不重复');
  assert(await page.locator('.source-tabs [data-source-tab="epr"]').textContent()==='EPR申请表及其他注册表 3','统一表格页签名称和数量正确');
  assert(await page.locator('.source-table').textContent().then(t=>t.includes('Alpha Example LLC')),'默认显示第一家真实申请表');
  assert(!calls.some(c=>c.member==='chain1.beta'||c.member==='chain1.nested-gamma'),'不提前读取其他公司的表内容');
  const stateReads=calls.filter(c=>c.path==='/api/state').length;
  await page.locator('.source-book-tabs button').nth(2).click();
  await page.waitForFunction(()=>document.querySelector('.source-table')?.textContent.includes('Gamma Example LLC'));
  assert((await page.locator('.source-file-copy b').textContent()).includes('Gamma'),'文件头与第三份申请表匹配');
  assert(await page.locator('.source-download').getAttribute('href').then(h=>new URL('http://local'+h).searchParams.get('member')==='chain1.nested-gamma'),'嵌套RAR成员下载使用后端成员标识');
  assert(calls.some(c=>c.member==='chain1.nested-gamma'),'嵌套RAR成员通过原预览接口读取');
  assert(await page.evaluate(()=>app.data.mails[0].details[0].fields.company)==='Alpha Example LLC','切换证据不修改业务公司');
  assert(await page.locator('.source-sheet-tabs button').count()===2,'工作簿全部工作表可切换');
  await page.locator('[data-source-sheet="产品信息"]').click();
  assert((await page.locator('.source-sheet-head').textContent()).includes('产品信息'),'切换到产品信息工作表');
  await page.locator('[data-source-more]').click();
  await page.waitForFunction(()=>document.querySelectorAll('.source-table tr').length===80);
  assert(await page.locator('.source-table tr').count()===80,'加载后续行保留全部80行');
  assert(await page.locator('[data-source-more]').count()===0,'读取到末尾不再显示加载按钮');
  const copied=await page.locator('.source-table td').nth(5).evaluate(e=>{const r=document.createRange();r.selectNodeContents(e);const s=window.getSelection();s.removeAllRanges();s.addRange(r);return s.toString()});
  assert(copied.length>0,'原表单元格可原生选字');
  await page.locator('[data-source-tab="body"]').click();
  assert((await page.locator('.source-original').textContent()).includes('正文独有代理'),'正文页显示原文');
  await page.locator('[data-source-tab="subject"]').click();
  assert((await page.locator('.source-original').textContent()).includes('标题独有国家'),'标题页显示原文');
  await page.locator('[data-source-tab="attachments"]').click();
  assert(await page.locator('.source-attachment').count()===2,'附件页显示所有邮件附件');
  assert(!(await page.locator('#evidence').textContent()).includes('其他附件文本与表格'),'不再显示重复的其他附件文本与表格折叠栏');
  await page.locator('.source-archive summary').first().click();
  assert(await page.locator('.source-archive-row').count()===5,'附件页显示压缩包全部成员');
  assert(await page.locator('.source-archive-row').filter({hasText:'加密文件'}).textContent().then(t=>t.includes('已加密')),'不可读取的成员有真实原因');
  assert(calls.filter(c=>c.path==='/api/state').length===stateReads,'切换来源和申请表不重新拉取整份邮件状态');
  await page.locator('[data-source-tab="epr"]').click();
  await page.locator('.source-book-tabs button').nth(1).click();
  await page.locator('.source-warning').filter({hasText:'模拟此文件读取失败'}).waitFor();
  assert(await page.locator('.source-book-tabs button').count()===3,'单个文件失败不丢失其他申请表');
  failBeta=false;await page.locator('[data-source-retry]').click();
  await page.waitForFunction(()=>document.querySelector('.source-table')?.textContent.includes('Beta Example LLC'));
  assert(await page.locator('.source-table').textContent().then(t=>t.includes('Beta Example LLC')),'失败文件可独立重试');
  await page.locator('[data-detail="D1"] [data-field="company"]').click();
  await page.waitForFunction(()=>document.querySelector('.source-table')?.textContent.includes('Alpha Example LLC'));
  assert(await page.locator('[data-source-tab="epr"]').getAttribute('class')==='active','点击公司字段优先定位申请表');
  assert(await page.locator('.source-table .evidence-hit').count()>0,'只在真实原文命中处标红');
  await page.locator('[data-detail="D1"] [data-field="agent"]').click();
  assert(await page.locator('[data-source-tab="body"]').getAttribute('class')==='active','没有表内证据则定位正文');
  await page.locator('[data-detail="D1"] [data-field="country"]').click();
  assert(await page.locator('[data-source-tab="subject"]').getAttribute('class')==='active','正文没有时定位标题');
  await page.locator('[data-detail="D1"] [data-field="program"]').click();
  assert(await page.locator('[data-source-tab="attachments"]').getAttribute('class')==='active','最后回落到附件名称');
  await page.locator('[data-source-tab="epr"]').click();await page.locator('.source-book-tabs button').nth(2).click();
  await page.evaluate(()=>{app.highlightField=null});
  await page.locator('[data-source-tab="epr"]').click();
  // Non-EPR workbooks, direct and archived, share the same tab and count.
  mail.attachment_files.push({filename:'客户登记.xlsm',token:'fixture-direct.xlsm',attachment_index:2},{filename:'历史登记.xls',token:'fixture-legacy.xls',attachment_index:3});
  members.push({path:'其他资料/客户登记.xlsx',member_token:'chain1.register',filename:'客户登记.xlsx',previewable:true,downloadable:true});
  await page.reload();
  await page.waitForFunction(()=>document.querySelectorAll('.source-book-tabs button').length===6);
  assert(await page.locator('[data-source-tab="epr"]').textContent()==='EPR申请表及其他注册表 6','数量包含直接上传与压缩包内全部六份表格');
  for(const filename of ['客户登记.xlsm','历史登记.xls']){
    await page.locator('.source-book-tabs button[title="'+filename+'"]').click();
    await page.waitForFunction(()=>document.querySelector('.source-table')?.textContent.includes('普通表格原始值'));
    assert((await page.locator('.source-file-copy b').textContent())===filename,'直接上传普通表格可预览 '+filename);
    assert(new URL('http://local'+await page.locator('.source-download').getAttribute('href')).pathname==='/api/attachment','直接上传表格保留原件下载 '+filename);
  }
  await page.locator('.source-book-tabs button[title="其他资料/客户登记.xlsx"]').click();
  await page.waitForFunction(()=>document.querySelector('.source-table')?.textContent.includes('压缩包内登记原始值'));
  assert((await page.locator('.source-file-copy b').textContent()).includes('客户登记.xlsx'),'压缩包内非EPR表同样可预览');
  assert(await page.locator('.source-sheet').getAttribute('data-evidence-source')==='attachment-content','普通表格合并展示但不冒充EPR申请表证据');
  assert(new URL('http://local'+await page.locator('.source-download').getAttribute('href')).searchParams.get('member')==='chain1.register','普通表格下载使用正确压缩包成员标识');
  assert(await page.locator('.source-book-tabs button').count()===6,'读取各表后仍完整保留全部表格');
  for(const width of [1440,1280,1024,390]){
    await page.setViewportSize({width,height:900});await page.waitForTimeout(60);
    assert(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth+1),'页签表格无页面横向溢出 '+width);
  }
  await page.setViewportSize({width:1600,height:1000});
  await page.screenshot({path:'B:/Codex_EcoPV_workspace/output/workbench-evidence-tabs-fixture.png'});
  assert(errors.length===0,'页面脚本无异常');
  assert(calls.every(c=>c.method==='GET'),'证据浏览无任何写接口请求');
  return {passed:checks.length,checks,catalogRequests:calls.filter(c=>c.path==='/api/attachment-preview'&&!c.member).length,productionWrites:0};
}
