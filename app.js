(async() => {
  'use strict';
  const lessons=window.EXCEL_TUTORIAL,demos=window.EXCEL_DEMOS,checks=window.EXCEL_CHECKS,storage=window.EXCEL_PROGRESS_STORE,$=id=>document.getElementById(id);
  if(!storage||!Array.isArray(lessons)||!lessons.length||!demos||!checks||new Set(lessons.map(l=>l.id)).size!==lessons.length||lessons.some(l=>typeof checks[l.id]!=='string'||!checks[l.id].trim()))return;
  const byId=new Map(lessons.map(l=>[l.id,l])),key='excel-zero-to-reporting-public-v2',courseId='excel-zero-to-reporting-public',version=1,filePreview=location.protocol==='file:';
  const modules={'00':'Start small','01':'Cells and formulas','02':'Counts and conditions','03':'Formatting and rounding','04':'Dates and durations','05':'Rules and errors','06':'Lookups and reconciliation','07':'Reliable formulas and Tables','08':'Pivots and readable reports','09':'Import, clean and refresh','10':'Your independent asset report','11':'Later: advanced tools and MOS'};
  const fresh=()=>({current:lessons[0].id,completed:[],answers:{}});
  function validate(value){
    if(!value||typeof value!=='object'||Array.isArray(value)||!byId.has(value.current)||!Array.isArray(value.completed)||new Set(value.completed).size!==value.completed.length||value.completed.some(id=>!byId.has(id))||!value.answers||typeof value.answers!=='object'||Array.isArray(value.answers))throw Error('The progress record does not match this public course.');
    const answers={};
    for(const [id,answer]of Object.entries(value.answers)){
      if(!byId.has(id)||!Number.isInteger(answer)||answer<0||answer>=byId.get(id).quiz.options.length)throw Error('A saved question answer is invalid.');
      answers[id]=answer;
    }
    return {current:value.current,completed:[...value.completed],answers};
  }
  let state=fresh(),damaged=null,storageBlocked=false,current=lessons[0],cells={},answerDirty=false,pending=0,saving=Promise.resolve(),revision=0,refreshId=0,updates=null;
  let saveFailure='Browser storage is unavailable. Progress is in memory only. Save a JSON backup before closing.';
  function decode(raw){
    if(raw===undefined)return fresh();
    try{const next=validate(JSON.parse(raw));damaged=null;return next;}
    catch(error){damaged=typeof raw==='string'?raw:JSON.stringify(raw);throw error;}
  }
  async function readSaved(){
    try{return decode(await storage.read());}
    catch(_){if(damaged===null)storageBlocked=true;return null;}
  }
  function saveNotice(){
    $('damaged-progress').hidden=damaged===null;
    if(damaged!==null){$('save-status').textContent='Saved progress is unreadable and has not been overwritten. Save the unreadable record or restore a valid backup. Your Excel files are unaffected.';$('progress-link').textContent='Recover progress';}
    else if(filePreview){$('save-status').className='warning';$('save-status').textContent='File preview: automatic saving is off because file:// storage is unreliable. Start START-COURSE.cmd, or run python -m tools.serve, for reliable local progress. Save a JSON backup before leaving this preview; restore it in the launched course.';$('progress-link').textContent='Save backup';}
    else if(storageBlocked){$('save-status').textContent=saveFailure;$('progress-link').textContent='Save backup';}
    else{$('save-status').textContent=pending?'Saving progress...':'Progress saved in this browser. Save your Excel workbook separately.';$('progress-link').textContent='Progress';}
  }
  function syncProgress(){
    state.current=current.id;progress();rail();$('complete').checked=state.completed.includes(current.id);
    if(!answerDirty){for(const radio of $('quiz-options').querySelectorAll('input'))radio.checked=state.answers[current.id]===Number(radio.value);feedback();}
    saveNotice();
  }
  async function refreshProgress(){
    if(pending)return;
    const attempt=++refreshId,before=revision;
    if(!filePreview&&!storageBlocked&&damaged===null){
      const latest=await readSaved();
      if(attempt!==refreshId||before!==revision||pending)return;
      if(latest)state=latest;
    }
    syncProgress();
  }
  function connectUpdates(){
    if(filePreview||storageBlocked||updates||typeof BroadcastChannel!=='function')return;
    try{updates=new BroadcastChannel(key);updates.onmessage=()=>refreshProgress();}catch(_){updates=null;}
  }
  function leaving(event){event.preventDefault();event.returnValue='';}
  function persist(edit,replace=false){
    revision++;if(!pending)window.addEventListener('beforeunload',leaving);pending++;saveNotice();
    saving=saving.then(async()=>{
      if(filePreview||storageBlocked||(damaged!==null&&!replace)){state=edit(state);return false;}
      let next;
      try{
        await storage.update(raw=>{next=edit(replace?state:decode(raw));return JSON.stringify(next);});
        state=next;
      }catch(_){state=next||edit(state);if(damaged===null){storageBlocked=true;saveFailure='Browser storage cannot save progress. Progress is in memory only. Save a JSON backup before closing.';}return false;}
      try{updates?.postMessage(null);}catch(_){updates?.close();updates=null;}
      return true;
    }).finally(()=>{pending--;if(!pending){window.removeEventListener('beforeunload',leaving);syncProgress();}else saveNotice();});
    return saving;
  }
  function progress(){
    const correct=Object.entries(state.answers).filter(([id,a])=>byId.get(id).quiz.correct===a).length;
    $('progress').max=lessons.length;$('progress').value=state.completed.length;
    $('progress-text').textContent=`${state.completed.length} of ${lessons.length} Excel practice checks marked complete`;
    $('quiz-progress').textContent=`${correct} of ${lessons.length} tutorial questions checked correctly. Workbook practice and questions are tracked separately.`;
  }
  function rail(){
    const query=$('topic-search').value.trim().toLowerCase(),fragment=document.createDocumentFragment();let count=0;
    for(const [module,title]of Object.entries(modules).sort(([a],[b])=>a.localeCompare(b))){
      const matches=lessons.filter(l=>l.module===module&&`${l.title} ${title}`.toLowerCase().includes(query));if(!matches.length)continue;
      const group=document.createElement('div');group.className='topic-group';const h=document.createElement('h3');h.textContent=title;group.append(h);
      for(const l of matches){const a=document.createElement('a');a.className='session-link';a.href='#'+l.id;a.dataset.session=l.id;a.textContent=l.title;if(l.id===state.current)a.setAttribute('aria-current','page');
        if(state.completed.includes(l.id)){const tag=document.createElement('span');tag.className='done-tag';tag.textContent='✓';tag.setAttribute('aria-label','Excel practice marked complete');a.append(tag);}
        a.addEventListener('click',event=>{event.preventDefault();go(l.id);});group.append(a);count++;
      }fragment.append(group);
    }
    $('session-list').replaceChildren(fragment);$('search-status').textContent=query?`${count} matching topics${count?'':'. Clear the search to see all lessons.'}`:'';
  }
  function valueText(value,format){
    if(typeof value!=='number')return String(value);
    if(format==='decimal2')return value.toFixed(2);if(format==='percent')return (100*value).toFixed(1)+'%';
    if(format==='date')return new Date(Date.UTC(1899,11,30)+value*86400000).toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric',timeZone:'UTC'});
    return String(Number(value.toPrecision(12)));
  }
  function renderExample(result=null){
    const e=current.example,output=demos.address(e.output),outputs={};if(result)result.forEach((row,r)=>{row.forEach((v,c)=>{outputs[demos.name(output.column+c,output.row+r)]=v;});});
    const positions=Object.keys({...e.cells,...outputs}).map(demos.address),maxRow=Math.max(output.row,...positions.map(p=>p.row)),maxColumn=Math.max(output.column,...positions.map(p=>p.column));
    const table=$('example-sheet');table.replaceChildren();const caption=document.createElement('caption');caption.textContent='Guided example — highlighted cells are editable';table.append(caption);
    const head=document.createElement('thead'),header=document.createElement('tr');
    for(let c=0;c<=maxColumn;c++){const th=document.createElement('th');th.scope='col';th.textContent=c?demos.name(c,1).replace(/1$/,''):'';header.append(th);}head.append(header);table.append(head);
    const body=document.createElement('tbody');
    for(let r=1;r<=maxRow;r++){const row=document.createElement('tr'),label=document.createElement('th');label.scope='row';label.textContent=r;row.append(label);
      for(let c=1;c<=maxColumn;c++){const address=demos.name(c,r),td=document.createElement('td');td.dataset.cell=address;
        if(e.editable.includes(address)){const input=document.createElement('input');input.autocomplete='off';input.dataset.input=address;input.setAttribute('aria-label',`Example cell ${address}`);input.type=typeof e.cells[address]==='number'?'number':'text';if(input.type==='number')input.step='any';input.value=cells[address]??'';
          input.addEventListener('input',()=>{$('example-status').textContent='Inputs changed. Run example to calculate. Example edits reset when you leave the lesson.';for(const output of table.querySelectorAll('.output'))output.textContent='—';});td.append(input);
        }else if(Object.hasOwn(outputs,address)){td.className='output';td.dataset.result=address;td.textContent=valueText(outputs[address],e.display);}else if(address===e.output){td.className='output';td.textContent='—';}else td.textContent=e.cells[address]??'';row.append(td);
      }body.append(row);
    }table.append(body);
  }
  function run(){
    try{for(const input of $('example-sheet').querySelectorAll('input')){if(input.type==='number'){const n=Number(input.value);if(input.value.trim()===''||!Number.isFinite(n)||Math.abs(n)>1e9)throw Error(`Enter a number between -1000000000 and 1000000000 in ${input.dataset.input}.`);cells[input.dataset.input]=n;}else cells[input.dataset.input]=input.value;}
      const result=demos.calculate(current.example,cells);renderExample(result);$('example-status').className='';$('example-status').textContent=`Calculated in ${current.example.output}${result.length>1||result[0].length>1?' and its spilled cells':''}. Change an input to explore the rule.`;
    }catch(error){$('example-status').className='error';$('example-status').textContent=error.message;for(const out of $('example-sheet').querySelectorAll('.output'))out.textContent='—';}
  }
  function quiz(){
    const q=current.quiz;$('quiz-question').textContent=q.question;$('quiz-options').replaceChildren(...q.options.map((text,i)=>{const label=document.createElement('label');label.className='quiz-option';const radio=document.createElement('input');radio.type='radio';radio.name='answer';radio.value=i;radio.checked=state.answers[current.id]===i;
      radio.addEventListener('change',()=>{answerDirty=true;$('quiz-feedback').className='';$('quiz-feedback').textContent='Selection changed. Choose Check answer to save and check it.';});label.append(radio,document.createTextNode(text));return label;}));
    answerDirty=false;feedback();
  }
  function feedback(){
    const q=current.quiz,answer=state.answers[current.id];$('quiz-feedback').className=answer===undefined?'':answer===q.correct?'correct':'error';$('quiz-feedback').textContent=answer===undefined?'Choose an answer, then check it.':`${answer===q.correct?'Correct.':'Not quite. Try again.'} ${q.explain}`;
  }
  function show(id,focus=true){
    const lesson=byId.get(id);if(!lesson)return;current=lesson;state.current=id;
    $('session-label').textContent=`${modules[lesson.module]} · Lesson ${lessons.indexOf(lesson)+1} · ${lesson.minutes} min starting estimate`;
    $('lesson-title').textContent=lesson.title;document.title=lesson.title+' — Excel Tutorial';$('goal').textContent=lesson.module==='11'?'Optional advanced practice after your independent core report.':'Read, change the example, check the question, then complete the independent Excel task.';
    $('lesson-text').innerHTML=lesson.html;$('tutorial-practice').replaceChildren(...lesson.practice.map(text=>{const li=document.createElement('li');li.textContent=text;return li;}));
    $('worked-check').open=false;const check=document.createElement('p');check.textContent=checks[id];$('worked-content').replaceChildren(check);
    const samples=window.EXCEL_WORKED_OUTPUTS?.[id];
    if(samples?.length){
      const table=document.createElement('table');table.className='sheet';
      const caption=document.createElement('caption');caption.textContent='Expected results for independent practice';table.append(caption);
      const head=document.createElement('thead'),row=document.createElement('tr');
      for(const label of ['Input or formula','Expected result']){const th=document.createElement('th');th.scope='col';th.textContent=label;row.append(th);}head.append(row);table.append(head);
      const body=document.createElement('tbody');
      for(const sample of samples){const tr=document.createElement('tr');tr.dataset.workedCase=sample.id;const input=document.createElement('th');input.scope='row';input.textContent=sample.input;const value=document.createElement('td');value.textContent=String(sample.value);tr.append(input,value);body.append(tr);}
      table.append(body);const wrap=document.createElement('div');wrap.className='sheet-wrap';wrap.append(table);$('worked-content').append(wrap);
    }
    $('example-explain').textContent=lesson.example.explain;$('example-formula').value=lesson.example.formula;cells={...lesson.example.cells};renderExample();run();quiz();$('copy-fallback').hidden=true;$('copy-status').textContent='';$('complete').checked=state.completed.includes(id);
    const index=lessons.indexOf(lesson);for(const suffix of ['','-top']){$('previous'+suffix).disabled=index===0;$('next'+suffix).disabled=index===lessons.length-1;$('next'+suffix).textContent=index===lessons.length-1?'End of tutorial':'Next ❯';$('position'+suffix).textContent=`${index+1} / ${lessons.length}`;}
    progress();rail();persist(next=>{next.current=id;return next;});$('sidebar').classList.remove('open');$('menu').setAttribute('aria-expanded','false');if(focus){$('lesson-title').focus({preventScroll:true});window.scrollTo({top:0,behavior:'instant'});}
  }
  function go(id){if(!byId.has(id))return;if(location.hash==='#'+id)show(id);else location.hash=id;}
  function download(text,filename,type){const url=URL.createObjectURL(new Blob([text],{type})),a=document.createElement('a');a.href=url;a.download=filename;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);}
  async function use(next){
    damaged=null;
    // Explicit replacement is confirmed by the caller and uses one write transaction.
    const saved=await persist(()=>validate(next),true);
    show(next.current);
    if(location.hash&&location.hash!=='#'+next.current)go(next.current);
    return saved;
  }
  function copy(text){
    const source=document.activeElement,y=window.scrollY;$('copy-text').value=text;$('copy-fallback').hidden=false;$('copy-text').focus();$('copy-text').select();let copied=false;try{copied=document.execCommand('copy');}catch(_){}
    if(copied){$('copy-fallback').hidden=true;source?.focus({preventScroll:true});window.scrollTo({top:y,behavior:'instant'});$('copy-status').textContent=text.startsWith('=')?`Copied formula. Paste into ${current.example.output} in Excel.`:'Copied starting data. Paste at A1. Set ID columns to Text first if leading zeros matter.';}else $('copy-status').textContent='Automatic copying was blocked. Press Ctrl+C to copy the selected text below.';
  }
  $('run-example').addEventListener('click',run);$('reset-example').addEventListener('click',()=>{cells={...current.example.cells};renderExample();run();});
  $('copy-formula').addEventListener('click',()=>copy(current.example.formula));
  $('copy-data').addEventListener('click',()=>{const e=current.example,p=Object.keys(e.cells).map(demos.address),rows=Math.max(...p.map(x=>x.row)),cols=Math.max(...p.map(x=>x.column));copy(Array.from({length:rows},(_,r)=>Array.from({length:cols},(_,c)=>String(e.cells[demos.name(c+1,r+1)]??'')).join('\t')).join('\n'));});
  $('quiz-form').addEventListener('submit',event=>{event.preventDefault();const chosen=$('quiz-form').querySelector('input:checked');if(!chosen){answerDirty=true;$('quiz-feedback').className='error';$('quiz-feedback').textContent='Choose one answer before checking.';return;}const id=current.id,answer=Number(chosen.value);answerDirty=false;persist(next=>{next.current=id;next.answers[id]=answer;return next;});});
  $('complete').addEventListener('change',()=>{const id=current.id,checked=$('complete').checked;persist(next=>{next.current=id;next.completed=checked?[...new Set([...next.completed,id])]:next.completed.filter(value=>value!==id);return next;});});
  for(const suffix of ['','-top']){$('previous'+suffix).addEventListener('click',()=>go(lessons[lessons.indexOf(current)-1]?.id));$('next'+suffix).addEventListener('click',()=>go(lessons[lessons.indexOf(current)+1]?.id));}
  $('resume').addEventListener('click',()=>go((lessons.find(l=>!state.completed.includes(l.id))||lessons.at(-1)).id));
  window.addEventListener('pageshow',async event=>{if(event.persisted){$('save-status').textContent='Loading saved progress...';const before=revision;if(!filePreview&&!storageBlocked&&damaged===null){const latest=await readSaved();if(latest&&before===revision&&!pending)state=latest;}connectUpdates();show(byId.has(location.hash.slice(1))?location.hash.slice(1):state.current,false);}});
  window.addEventListener('pagehide',()=>{if(!pending){updates?.close();updates=null;storage.close();}});
  window.addEventListener('focus',()=>refreshProgress());
  window.addEventListener('hashchange',()=>{if(byId.has(location.hash.slice(1)))show(location.hash.slice(1));});$('topic-search').addEventListener('input',rail);
  $('menu').addEventListener('click',()=>{const open=$('sidebar').classList.toggle('open');$('menu').setAttribute('aria-expanded',String(open));if(open)$('topic-search').focus();});document.addEventListener('keydown',event=>{if(event.key==='Escape'&&$('sidebar').classList.contains('open')){$('sidebar').classList.remove('open');$('menu').setAttribute('aria-expanded','false');$('menu').focus();}});
  $('theme').addEventListener('click',()=>{const dark=document.documentElement.dataset.theme!=='dark';document.documentElement.dataset.theme=dark?'dark':'light';$('theme').setAttribute('aria-pressed',String(dark));$('theme').textContent=dark?'Light theme':'Dark theme';});
  $('export-progress').addEventListener('click',async()=>{await saving;await refreshProgress();download(JSON.stringify({format:courseId,version,state,exportedAt:new Date().toISOString()},null,2),'excel-public-progress-'+new Date().toISOString().slice(0,10)+'.json','application/json;charset=utf-8');$('backup-status').textContent='Backup sent to browser downloads. Excel workbooks are not included.';});
  $('import-progress').addEventListener('change',async event=>{const input=event.target,file=input.files[0];if(!file)return;input.disabled=true;$('backup-status').textContent='Reading backup...';
    try{if(file.size>128*1024)throw Error('Choose a backup smaller than 128 KiB.');const backup=JSON.parse(await file.text());if(backup.format!==courseId||backup.version!==version)throw Error('Choose a backup for this public edition. Old local-course backups remain usable only in the old course.');const next=validate(backup.state);
      if(!window.confirm(`Restore ${next.completed.length} practice ticks and ${Object.keys(next.answers).length} checked answers? This replaces only this public course's progress. Excel files and the old local course are unchanged.`)){$('backup-status').textContent='Restore cancelled. Current progress is unchanged.';return;}
      const saved=await use(next);$('backup-status').textContent=saved?'Progress restored and saved in this browser. Excel workbooks are unchanged.':'Restored in memory only. Save a JSON backup before closing.';
    }catch(error){$('backup-status').textContent=`Backup not restored: ${error instanceof SyntaxError?'invalid JSON.':error.message} Current progress is unchanged.`;}finally{input.value='';input.disabled=false;}
  });
  $('export-damaged').addEventListener('click',()=>{if(damaged!==null)download(damaged,'excel-public-progress-unreadable.txt','text/plain;charset=utf-8');});
  $('reset-progress').addEventListener('click',async()=>{if(window.confirm('Start fresh? Save the unreadable record first if needed. This replaces only public-course progress. Excel files and old local progress remain unchanged.')){const saved=await use(fresh());$('backup-status').textContent=saved?'Fresh progress saved.':'Fresh progress is in memory only. Save a JSON backup.';}});
  $('complete').disabled=true;$('quiz-form').querySelector('button').disabled=true;
  if(!filePreview)state=await readSaved()||state;
  connectUpdates();$('complete').disabled=false;$('quiz-form').querySelector('button').disabled=false;
  $('export-progress').disabled=false;$('import-progress').disabled=false;
  if(filePreview)$('loading-warning').after($('save-status'));
  $('loading-warning').hidden=true;show(byId.has(location.hash.slice(1))?location.hash.slice(1):state.current,false);
})();
