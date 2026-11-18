import {test,expect} from '@playwright/test';
import fs from 'node:fs';
import path from 'node:path';
import {pathToFileURL} from 'node:url';
import {loadCourse,root} from './plan.mjs';
const {lessons}=loadCourse();
const oracle=JSON.parse(fs.readFileSync(path.join(root,'tests/excel-oracle.json'),'utf8'));
const projectOracle=JSON.parse(fs.readFileSync(path.join(root,'tests/project-oracle.json'),'utf8'));
const url=process.env.COURSE_URL||'http://127.0.0.1:8765/START-HERE.html';
const previewURL=pathToFileURL(path.join(process.env.COURSE_ROOT||root,'START-HERE.html')).href;
const key='excel-zero-to-reporting-public-v2',oldKey='cert-skills-linear-v1';
async function record(page,...value){
  return page.evaluate(({key,write,value})=>new Promise((resolve,reject)=>{
    const opening=indexedDB.open(key,1);
    opening.onupgradeneeded=()=>opening.result.createObjectStore('records');
    opening.onerror=()=>reject(opening.error);
    opening.onsuccess=()=>{
      const db=opening.result,transaction=db.transaction('records',write?'readwrite':'readonly'),store=transaction.objectStore('records');
      const request=write?store.put(value,'progress'):store.get('progress');
      let result;request.onsuccess=()=>{result=request.result;};
      transaction.oncomplete=()=>{db.close();resolve(write?value:result);};
      transaction.onabort=()=>{db.close();reject(transaction.error);};
    };
  }),{key,write:value.length>0,value:value[0]});
}
async function seedRecord(page,raw){
  let seeded=false;
  await page.route('**/app.js',async route=>{if(!seeded){seeded=true;await record(page,raw);}await route.continue();});
}
const display=(v,f)=>typeof v!=='number'?String(v):f==='decimal2'?v.toFixed(2):f==='percent'?(v*100).toFixed(1)+'%':f==='date'?new Date(Date.UTC(1899,11,30)+v*86400000).toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric',timeZone:'UTC'}):String(Number(v.toPrecision(12)));
async function settled(page){await expect(page.locator('#save-status')).not.toHaveText('Saving progress...');}
async function open(page,id){await page.goto(url+(id?'#'+id:''));await expect(page.locator('[aria-current="page"]')).toHaveCount(1);await settled(page);}
async function compare(page,l,matrix){await expect(page.locator('[data-result]')).toHaveText(matrix.flat().map(v=>display(v,l.example.display)));}
async function compareSavedControls(page,correct){
  const saved=JSON.parse(await record(page));expect(saved.completed).toEqual(['s01']);expect(saved.answers.s01).toBe(correct);
  await expect(page.locator('#progress')).toHaveAttribute('value','1');await expect(page.locator('#complete')).toBeChecked();
  await expect(page.locator(`input[name="answer"][value="${correct}"]`)).toBeChecked();await expect(page.locator('#quiz-feedback')).toContainText('Correct.');
}
async function compareWorked(page,pairs){
  await page.locator('#worked-check summary').click();
  for(const [id,name]of pairs){const observed=projectOracle.checks.find(c=>c.name===name);expect(observed?.passed).toBe(true);await expect(page.locator(`[data-worked-case="${id}"] td`)).toHaveText(String(observed.actual));}
}
async function importBackup(page,testInfo,name,backup,accept=true){
  const file=testInfo.outputPath(name);fs.writeFileSync(file,typeof backup==='string'?backup:JSON.stringify(backup));
  page.once('dialog',d=>accept?d.accept():d.dismiss());await page.locator('#import-progress').setInputFiles(file);
}
test('all lessons: native results, edits, reset, feedback, navigation and persistence',async({page},testInfo)=>{
  test.setTimeout(180000);const errors=[];page.on('pageerror',e=>errors.push(e.message));await open(page);
  await expect(page.locator('.topic-group h3').first()).toHaveText('Start small');await expect(page.locator('#progress')).toHaveAttribute('value','0');
  for(let i=0;i<lessons.length;i++){
    const l=lessons[i],native=oracle.cases.find(c=>c.id===l.id);expect(native?.passed).toBe(true);
    await expect(page.locator('#lesson-title')).toHaveText(l.title);await expect(page.locator('#tutorial-practice li')).toHaveCount(l.practice.length);await compare(page,l,native.base);
    await page.locator(`[data-input="${l.example.change.cell}"]`).fill(String(l.example.change.value));await page.locator('#run-example').click();await expect(page.locator('#example-status')).toContainText('Calculated');await compare(page,l,native.changed);
    await page.locator('#reset-example').click();await compare(page,l,native.base);
    await page.locator('#quiz-form button').click();await expect(page.locator('#quiz-feedback')).toContainText('Choose one');
    await page.locator(`input[name="answer"][value="${(l.quiz.correct+1)%l.quiz.options.length}"]`).check();await page.locator('#quiz-form button').click();await expect(page.locator('#quiz-feedback')).toContainText('Not quite.');
    await page.locator(`input[name="answer"][value="${l.quiz.correct}"]`).check();await page.locator('#quiz-form button').click();await expect(page.locator('#quiz-feedback')).toContainText('Correct.');await expect(page.locator('#complete')).not.toBeChecked();
    await page.locator('#worked-check summary').click();await expect(page.locator('#worked-content')).toBeVisible();
    await page.locator('#complete').check();await expect(page.locator('#progress')).toHaveAttribute('value',String(i+1));
    if([0,35,56].includes(i)){await page.locator('#lesson-title').scrollIntoViewIfNeeded();await page.screenshot({path:testInfo.outputPath(l.id+'.png')});}
    if(i<lessons.length-1){await page.locator(i%2?'#next':'#next-top').click();await expect(page.locator('#lesson-title')).toBeFocused();}
  }
  await expect(page.locator('#next')).toBeDisabled();await page.reload();await expect(page.locator('#progress')).toHaveAttribute('value',String(lessons.length));await expect(page.locator('#quiz-progress')).toContainText(`${lessons.length} of ${lessons.length}`);
  await page.locator('#previous-top').click();await expect(page.locator('#lesson-title')).toHaveText(lessons.at(-2).title);await page.goBack();await expect(page.locator('#lesson-title')).toHaveText(lessons.at(-1).title);expect(errors).toEqual([]);
});
test('genuine backup download, restore, cancellation, foreign record and legacy isolation',async({page},testInfo)=>{
  const legacy=JSON.stringify({current:'s02',completed:['s01'],positions:{s02:225},answers:{s01:1}});
  const prototypeKey='excel-zero-to-reporting-public-v1';
  await page.addInitScript(({oldKey,prototypeKey,legacy})=>{if(localStorage.getItem(oldKey)===null)localStorage.setItem(oldKey,legacy);if(localStorage.getItem(prototypeKey)===null)localStorage.setItem(prototypeKey,legacy);},{oldKey,prototypeKey,legacy});
  await open(page);await page.locator('#complete').check();await page.locator('input[name="answer"][value="1"]').check();await page.locator('#quiz-form button').click();
  const downloadPromise=page.waitForEvent('download');await page.locator('#export-progress').click();const download=await downloadPromise,file=testInfo.outputPath('downloaded-progress.json');await download.saveAs(file);const backup=JSON.parse(fs.readFileSync(file,'utf8'));expect(backup.format).toBe('excel-zero-to-reporting-public');expect(backup.state.completed).toEqual(['s01']);
  await page.locator('#complete').uncheck();await importBackup(page,testInfo,'cancel.json',backup,false);await expect(page.locator('#backup-status')).toContainText('cancelled');await expect(page.locator('#complete')).not.toBeChecked();
  await importBackup(page,testInfo,'restore.json',backup,true);await expect(page.locator('#backup-status')).toHaveText('Progress restored and saved in this browser. Excel workbooks are unchanged.');
  await page.reload();await expect(page.locator('#complete')).toBeChecked();
  const before=await record(page);
  await importBackup(page,testInfo,'foreign.json',{format:'cert-skills-progress',version:2,courseVersion:1,state:backup.state});await expect(page.locator('#backup-status')).toContainText('not restored');
  await importBackup(page,testInfo,'invalid.json',{...backup,state:{...backup.state,answers:{s01:99}}});await expect(page.locator('#backup-status')).toContainText('invalid');
  await importBackup(page,testInfo,'malformed.json','{broken');await expect(page.locator('#backup-status')).toContainText('invalid JSON');
  expect(await record(page)).toBe(before);expect(await page.evaluate(k=>localStorage.getItem(k),oldKey)).toBe(legacy);expect(await page.evaluate(k=>localStorage.getItem(k),prototypeKey)).toBe(legacy);
});
test('unreadable stored bytes survive until explicit confirmed recovery',async({page},testInfo)=>{
  const raw='{not valid progress';await seedRecord(page,raw);await open(page);
  await expect(page.locator('#damaged-progress')).toBeVisible();await page.locator('#complete').check();expect(await record(page)).toBe(raw);
  const dl=page.waitForEvent('download');await page.locator('#export-damaged').click();const file=testInfo.outputPath('unreadable.txt');await(await dl).saveAs(file);expect(fs.readFileSync(file,'utf8')).toBe(raw);
  page.once('dialog',d=>d.dismiss());await page.locator('#reset-progress').click();expect(await record(page)).toBe(raw);
  page.once('dialog',d=>d.accept());await page.locator('#reset-progress').click();await expect(page.locator('#progress')).toHaveAttribute('value','0');expect(JSON.parse(await record(page)).completed).toEqual([]);
});
for(const mode of ['read-blocked','quota-denied'])test(`storage ${mode}: usable lesson and real export`,async({page},testInfo)=>{
  await page.addInitScript(mode=>{(mode==='read-blocked'?IDBFactory.prototype:IDBObjectStore.prototype)[mode==='read-blocked'?'open':'put']=()=>{throw new DOMException('Test storage refusal','SecurityError');};},mode);await open(page);await page.locator('#complete').check();await expect(page.locator('#save-status')).toContainText('memory only');
  const dl=page.waitForEvent('download');await page.locator('#export-progress').click();const file=testInfo.outputPath('memory-backup.json');await(await dl).saveAs(file);expect(JSON.parse(fs.readFileSync(file,'utf8')).state.completed).toEqual(['s01']);
});
test('search, mobile keyboard, theme and print',async({page},testInfo)=>{
  await page.setViewportSize({width:390,height:844});await open(page);await page.locator('#menu').click();await expect(page.locator('#topic-search')).toBeFocused();await page.locator('#topic-search').fill('weighted');await expect(page.locator('.session-link')).toHaveCount(1);await page.keyboard.press('Escape');await expect(page.locator('#menu')).toBeFocused();await page.locator('#menu').click();await page.locator('.session-link').click();await expect(page.locator('#lesson-title')).toHaveText('Excel Weighted Averages');
  expect(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth)).toBe(true);await page.screenshot({path:testInfo.outputPath('mobile.png')});await page.locator('#theme').click();await expect(page.locator('html')).toHaveAttribute('data-theme','dark');await page.screenshot({path:testInfo.outputPath('mobile-dark.png')});
  await page.emulateMedia({media:'print'});await expect(page.locator('#quiz-question')).toBeVisible();await page.pdf({path:testInfo.outputPath('printed-lesson.pdf'),format:'A4'});
});
test('unsupported and invalid inputs do not show plausible stale output',async({page})=>{
  await open(page,'s08');await page.locator('[data-input="D2"]').fill('*');await page.locator('#run-example').click();await expect(page.locator('#example-status')).toContainText('literal criteria');await expect(page.locator('[data-result]')).toHaveText('—');
  await open(page,'s05');await page.locator('[data-input="B2"]').fill('0');await page.locator('#run-example').click();await expect(page.locator('#example-status')).toContainText('denominator is zero');
  await page.locator('[data-input="B2"]').fill('');await page.locator('#run-example').click();await expect(page.locator('#example-status')).toContainText('Enter a number');await page.locator('#reset-example').click();await expect(page.locator('#example-status')).toContainText('Calculated');
});
test('offline file preview downloads original inputs and offers copy fallback',async({page,context},testInfo)=>{
  await context.setOffline(true);await page.goto(previewURL+'#s51');await expect(page.locator('#lesson-title')).toHaveText('Excel Refreshable Asset Report');
  await page.evaluate(()=>{document.execCommand=()=>false;});await page.locator('#copy-data').click();await expect(page.locator('#copy-fallback')).toBeVisible();await expect(page.locator('#copy-text')).toHaveValue(/AssetID\tOwner/);
  for(const link of await page.locator('#resources a').all()){
    const download=page.waitForEvent('download');await page.locator('#resources').evaluate(e=>{e.open=true;});await link.click();const d=await download;const target=testInfo.outputPath(d.suggestedFilename());await d.saveAs(target);expect(fs.statSync(target).size).toBeGreaterThan(0);
  }
});
test('worked project checks and the actual schedule link agree with supplied sales',async({page})=>{
  await open(page,'s45');await page.locator('#worked-check summary').click();await expect(page.locator('#worked-content')).toContainText('RM10200');await expect(page.locator('#worked-content')).toContainText('RM12500');
  await page.locator('#help summary').click();await page.getByRole('link',{name:'Complete study schedule',exact:true}).click();await expect(page).toHaveTitle('Your Excel study schedule');await expect(page.getByRole('heading',{name:'Your complete study schedule'})).toBeVisible();
  await page.getByRole('link',{name:'Excel Refreshable Asset Report',exact:true}).click();await expect(page.locator('#lesson-title')).toHaveText('Excel Refreshable Asset Report');
});
test('the native-result browser verifier rejects a deliberately broken calculator',async({page})=>{
  await open(page,'s01');const l=lessons.find(l=>l.id==='s01'),native=oracle.cases.find(c=>c.id==='s01');await compare(page,l,native.base);
  await page.evaluate(()=>{window.EXCEL_DEMOS.calculate=()=>[[0]];});await page.locator('#run-example').click();await expect(page.locator('[data-result]')).toHaveText('0');
  await expect(compare(page,l,native.base)).rejects.toThrow();
});
test('independent worked text outputs and band boundaries agree with real Excel',async({page})=>{
  await open(page,'x05');await compareWorked(page,[['document-prefix','document prefix'],['document-year','document year'],['document-sequence','document sequence preserves zero'],['document-length','document character count']]);
  await page.locator('a[data-session="s06"]').click();await compareWorked(page,[['band-0.59','band 0.59'],['band-0.6','band 0.6'],['band-0.79','band 0.79'],['band-0.8','band 0.8']]);
});
test('wrong and omitted independent answers are rejected by the factual verifier',async({page})=>{
  const facts=[['document-prefix','document prefix'],['document-year','document year'],['document-sequence','document sequence preserves zero'],['document-length','document character count']];
  await open(page,'x05');await compareWorked(page,facts);
  await page.evaluate(()=>{window.EXCEL_WORKED_OUTPUTS.x05.find(c=>c.id==='document-length').value=99;});
  await page.locator('a[data-session="s06"]').click();await page.locator('a[data-session="x05"]').click();await expect(compareWorked(page,facts)).rejects.toThrow();
  await page.evaluate(()=>{window.EXCEL_WORKED_OUTPUTS.x05=window.EXCEL_WORKED_OUTPUTS.x05.filter(c=>c.id!=='document-length');});
  await page.locator('a[data-session="s06"]').click();await page.locator('a[data-session="x05"]').click();await expect(compareWorked(page,facts)).rejects.toThrow();
});
test('recovery probe is a usable two-line CSV, not an escaped display string',async({page})=>{
  await open(page,'s49');await expect(page.locator('#lesson-text pre')).toHaveText('Date,Product,Department,Units,UnitPrice\n2026-04-01,AP,IT,1,123');
});
test('restored progress and lesson survive reload when History API is unavailable',async({page},testInfo)=>{
  await page.addInitScript(()=>{history.replaceState=()=>{throw new DOMException('History API unavailable for this file context.','SecurityError');};});
  await open(page,'s03');
  const restored={format:'excel-zero-to-reporting-public',version:1,state:{current:'s01',completed:['s01'],answers:{s01:lessons[0].quiz.correct}}};
  await importBackup(page,testInfo,'history-denied.json',restored,true);
  await expect(page.locator('#backup-status')).toHaveText('Progress restored and saved in this browser. Excel workbooks are unchanged.');
  await expect(page).toHaveURL(/#s01$/);await expect(page.locator('#complete')).toBeChecked();
  await page.reload();await expect(page.locator('#lesson-title')).toHaveText(lessons[0].title);await expect(page.locator('#complete')).toBeChecked();
  await expect(page.locator(`input[name="answer"][value="${lessons[0].quiz.correct}"]`)).toBeChecked();await expect(page.locator('#quiz-progress')).toContainText('1 of 60');
});
test('document Back history restores saved practice and answers, not obsolete form values',async({page})=>{
  const correct=lessons[0].quiz.correct,wrong=(correct+1)%lessons[0].quiz.options.length;
  await open(page,'s01');await page.locator(`input[name="answer"][value="${wrong}"]`).check();
  await page.goto('about:blank');await open(page,'s01');
  await page.locator('#complete').check();await page.locator(`input[name="answer"][value="${correct}"]`).check();await page.locator('#quiz-form button').click();
  await settled(page);await page.goBack({waitUntil:'commit'});await page.goBack({waitUntil:'commit'});await expect(page.locator('#lesson-title')).toHaveText(lessons[0].title);
  await compareSavedControls(page,correct);
});
test('file preview preserves old saved bytes and exports current in-memory practice honestly',async({page},testInfo)=>{
  const original=JSON.stringify({current:'s01',completed:[],answers:{}});
  await page.addInitScript(({key,original})=>{if(localStorage.getItem(key)===null)localStorage.setItem(key,original);},{key,original});
  await page.goto(previewURL+'#s01');await expect(page.locator('#save-status')).toContainText('File preview: automatic saving is off');
  await expect(page.locator('#save-status')).toBeInViewport();
  await page.locator('#complete').check();expect(await page.evaluate(k=>localStorage.getItem(k),key)).toBe(original);
  const downloading=page.waitForEvent('download');await page.locator('#export-progress').click();const download=await downloading,file=testInfo.outputPath('file-preview-backup.json');await download.saveAs(file);
  const backup=JSON.parse(fs.readFileSync(file,'utf8'));expect(backup.state.completed).toEqual(['s01']);
  await page.locator('#complete').uncheck();await importBackup(page,testInfo,'preview-restore.json',backup);
  await expect(page.locator('#complete')).toBeChecked();await expect(page.locator('#backup-status')).toContainText('Restored in memory only');expect(await page.evaluate(k=>localStorage.getItem(k),key)).toBe(original);
});
test.describe('native back-forward cache',()=>{
  for(const broken of [false,true])test(broken?'the saved-control verifier rejects a stale cached document':'cached documents re-read the current saved practice and answer',async({page})=>{
    await page.addInitScript(broken=>{
      window.addEventListener('pageshow',event=>{window.__courseWasCached=event.persisted;});
      if(broken){const add=window.addEventListener;window.addEventListener=function(type,...args){if(type!=='pageshow'&&type!=='focus')return add.call(this,type,...args);};}
    },broken);
    const correct=lessons[0].quiz.correct,wrong=(correct+1)%lessons[0].quiz.options.length;
    await open(page,'s01');await page.locator(`input[name="answer"][value="${wrong}"]`).check();
    await page.goto(new URL('SCHEDULE.html',url).href);await open(page,'s01');
    await page.locator('#complete').check();await page.locator(`input[name="answer"][value="${correct}"]`).check();await page.locator('#quiz-form button').click();
    await settled(page);await page.goBack({waitUntil:'commit'});await page.goBack({waitUntil:'commit'});
    await expect.poll(()=>page.evaluate(()=>window.__courseWasCached)).toBe(true);
    if(broken)await expect(compareSavedControls(page,correct)).rejects.toThrow();else await compareSavedControls(page,correct);
  });
});

for(const broken of [false,true])test(broken?'the progress verifier rejects a stale-record reader in a second tab':'second-tab navigation preserves saved practice and checked answers',async({page,context})=>{
  await open(page,'s01');const other=await context.newPage();await open(other,'s01');
  if(broken)await other.evaluate(async()=>{const stale=await window.EXCEL_PROGRESS_STORE.read(),update=window.EXCEL_PROGRESS_STORE.update;window.EXCEL_PROGRESS_STORE.update=edit=>update(()=>edit(stale));});
  await page.locator('#complete').check();await page.locator(`input[name="answer"][value="${lessons[0].quiz.correct}"]`).check();await page.locator('#quiz-form button').click();await settled(page);
  await other.locator('#next-top').click();await expect(other.locator('#lesson-title')).toHaveText(lessons[1].title);await settled(other);
  const reopened=await context.newPage();await open(reopened,'s01');
  if(broken){await expect(compareSavedControls(reopened,lessons[0].quiz.correct)).rejects.toThrow();await expect(reopened.locator('#progress')).toHaveAttribute('value','0');}
  else await compareSavedControls(reopened,lessons[0].quiz.correct);
});

test('simultaneously queued practice saves retain both lessons after the real shared transaction is released',async({page,context})=>{
  await open(page,'s01');const other=await context.newPage();await open(other,'s02');
  await page.evaluate(key=>new Promise((ready,reject)=>{
    const opening=indexedDB.open(key,1);opening.onerror=()=>reject(opening.error);
    opening.onsuccess=()=>{const db=opening.result,transaction=db.transaction('records','readwrite'),store=transaction.objectStore('records');let held=true;
      window.releaseCourseTransaction=()=>{held=false;};
      const keepAlive=()=>{store.get('progress').onsuccess=()=>{ready();if(held)keepAlive();};};keepAlive();
      transaction.oncomplete=()=>db.close();transaction.onabort=()=>{db.close();reject(transaction.error);};
    };
  }),key);
  try{
    await Promise.all([page.locator('#complete').check(),other.locator('#complete').check()]);
    await expect(page.locator('#save-status')).toHaveText('Saving progress...');await expect(other.locator('#save-status')).toHaveText('Saving progress...');
  }finally{await page.evaluate(()=>window.releaseCourseTransaction());}
  await settled(page);await settled(other);await page.reload();await settled(page);
  await expect(page.locator('#complete')).toBeChecked();await expect(page.locator('#progress')).toHaveAttribute('value','2');
  await other.reload();await settled(other);await expect(other.locator('#complete')).toBeChecked();
});

test('uncheck, replacement, reset and a refused write do not resurrect or erase another tab record',async({page,context},testInfo)=>{
  await open(page,'s01');const other=await context.newPage();await open(other,'s02');
  await page.locator('#complete').check();await settled(page);await other.locator('#complete').check();await settled(other);
  await page.locator('#complete').uncheck();await settled(page);await other.locator('#next-top').click();await settled(other);
  await page.reload();await settled(page);await expect(page.locator('#progress')).toHaveAttribute('value','1');await expect(page.locator('#complete')).not.toBeChecked();
  const backup={format:'excel-zero-to-reporting-public',version:1,state:{current:'s03',completed:['s03'],answers:{}}};
  await importBackup(page,testInfo,'replacement.json',backup);await expect(page.locator('#backup-status')).toContainText('restored and saved');
  await other.locator('#next-top').click();await settled(other);await page.reload();await settled(page);await expect(page.locator('#progress')).toHaveAttribute('value','1');await expect(page.locator('#complete')).toBeChecked();
  await record(page,'{broken recovery fixture');await other.reload();await expect(other.locator('#damaged-progress')).toBeVisible();
  other.once('dialog',d=>d.accept());await other.locator('#reset-progress').click();await expect(other.locator('#backup-status')).toHaveText('Fresh progress saved.');await settled(other);
  await page.locator('#next-top').click();await settled(page);await expect(page.locator('#progress')).toHaveAttribute('value','0');
  await page.locator('#complete').check();await settled(page);
  await other.evaluate(()=>{IDBObjectStore.prototype.put=()=>{throw new DOMException('Refused test write','QuotaExceededError');};});
  await other.locator('#complete').check();await expect(other.locator('#save-status')).toContainText('memory only');
  const downloading=other.waitForEvent('download');await other.locator('#export-progress').click();const target=testInfo.outputPath('refused-write.json');await(await downloading).saveAs(target);
  expect(JSON.parse(fs.readFileSync(target,'utf8')).state.completed.sort()).toEqual(['s01','s04']);
  await page.reload();await settled(page);await expect(page.locator('#progress')).toHaveAttribute('value','1');await expect(page.locator('#complete')).toBeChecked();
});

test('unavailable transactional storage preserves prototype bytes and exports memory-only practice',async({page},testInfo)=>{
  const original=JSON.stringify({current:'s01',completed:[],answers:{}});
  await page.addInitScript(({key,original})=>{Object.defineProperty(window,'indexedDB',{value:undefined});localStorage.setItem(key,original);},{key,original});
  await open(page,'s01');await page.locator('#complete').check();await expect(page.locator('#save-status')).toContainText('memory only');
  const downloading=page.waitForEvent('download');await page.locator('#export-progress').click();const target=testInfo.outputPath('no-locks.json');await(await downloading).saveAs(target);
  expect(JSON.parse(fs.readFileSync(target,'utf8')).state.completed).toEqual(['s01']);expect(await page.evaluate(key=>localStorage.getItem(key),key)).toBe(original);
});
