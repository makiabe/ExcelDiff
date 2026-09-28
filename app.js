(() => {
  'use strict';
  const $=s=>document.querySelector(s),$$=s=>[...document.querySelectorAll(s)];
  const C=globalThis.ExcelDiffCodec,E=globalThis.ExcelDiffEngine;
  const state={files:[null,null],status:[globalThis.ExcelDiffI18n.t('app.001'),globalThis.ExcelDiffI18n.t('app.002')],books:[],result:null,activeSheet:0,page:0,revision:0,loading:false,target:null};
  const el={input:$('#fileInput'),drop:$('#dropZone'),chips:$('#fileChips'),message:$('#message'),compare:$('#compareBtn'),results:$('#results'),summary:$('#resultSummary'),tabs:$('#sheetTabs'),table:$('#resultTableWrap'),diagnostics:$('#diagnostics'),warnings:$('#resultWarnings'),pagination:$('#pagination')};
  const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const fmt=n=>n<1024?`${n} B`:n<1048576?`${(n/1024).toFixed(1)} KB`:`${(n/1048576).toFixed(1)} MB`;
  const statusLabel=s=>({added:globalThis.ExcelDiffI18n.t('app.003'),deleted:globalThis.ExcelDiffI18n.t('app.004'),changed:globalThis.ExcelDiffI18n.t('app.005'),same:globalThis.ExcelDiffI18n.t('app.006')}[s]||s);
  const typeLabel=t=>({n:globalThis.ExcelDiffI18n.t('app.007'),s:globalThis.ExcelDiffI18n.t('app.008'),b:globalThis.ExcelDiffI18n.t('app.009'),e:globalThis.ExcelDiffI18n.t('app.010'),d:globalThis.ExcelDiffI18n.t('app.011'),blank:globalThis.ExcelDiffI18n.t('app.012')}[t]||t);
  const styleLabel=k=>({font:globalThis.ExcelDiffI18n.t('app.013'),fill:globalThis.ExcelDiffI18n.t('app.014'),border:globalThis.ExcelDiffI18n.t('app.015'),alignment:globalThis.ExcelDiffI18n.t('app.016'),numFmt:globalThis.ExcelDiffI18n.t('app.017'),protection:globalThis.ExcelDiffI18n.t('app.018'),flags:globalThis.ExcelDiffI18n.t('app.019'),rich:globalThis.ExcelDiffI18n.t('app.020')}[k]||k);
  const setMessage=(s,kind='')=>{el.message.textContent=s;el.message.className='message'+(kind?' '+kind:'');};
  function options(){return {ignoreWhitespace:$('#ignoreWhitespace').checked,ignoreCase:$('#ignoreCase').checked,compareFormulas:$('#compareFormulas').checked,compareStyles:$('#compareStyles').checked,compareTypes:$('#compareTypes').checked,excludeColumns:$('#excludeColumns').value,primaryKey:$('#primaryKey').value};}
  function invalidate(message){state.revision++;state.result=null;state.books=[];el.results.hidden=true;el.compare.disabled=!state.files.every(Boolean)||state.loading;if(message)setMessage(message);}
  function renderFiles(){
    el.chips.innerHTML=state.files.map((f,i)=>`<div class="file-slot ${f?'selected':'empty'}" data-slot="${i}"><span class="file-slot-label">${i?globalThis.ExcelDiffI18n.t('app.021'):globalThis.ExcelDiffI18n.t('app.022')}</span><span class="file-slot-name" title="${esc(f?.name||globalThis.ExcelDiffI18n.t('app.023'))}">${esc(f?.name||globalThis.ExcelDiffI18n.t('app.024'))}</span><span class="file-slot-meta">${f?fmt(f.size)+' · ':''}${esc(state.status[i])}</span><div class="slot-actions"><button type="button" data-pick="${i}">${f?globalThis.ExcelDiffI18n.t('app.025'):globalThis.ExcelDiffI18n.t('app.026')}</button>${f?globalThis.ExcelDiffI18n.t('app.027',{p0:i,p1:i?'B':'A'}):''}</div></div>`).join('');
    el.compare.disabled=!state.files.every(Boolean)||state.loading;
    $('#swapFiles').disabled=!state.files.every(Boolean)||state.loading;
  }
  function chooseFiles(list,target=null){
    const incoming=[...list];if(!incoming.length)return;
    if(incoming.some(f=>!/^.+\.(xlsx|xls|csv)$/i.test(f.name))){invalidate();setMessage(globalThis.ExcelDiffI18n.t('app.028'),'error');return;}
    if(incoming.length>2||(target!==null&&incoming.length!==1)){setMessage(target===null?globalThis.ExcelDiffI18n.t('app.029'):globalThis.ExcelDiffI18n.t('app.030'),'error');return;}
    if(incoming.length===2){state.files=[incoming[0],incoming[1]];state.status=[globalThis.ExcelDiffI18n.t('app.031'),globalThis.ExcelDiffI18n.t('app.032')];}
    else{const i=target!==null?target:!state.files[0]?0:1;state.files[i]=incoming[0];state.status[i]=globalThis.ExcelDiffI18n.t('app.033');}
    state.status=state.files.map(f=>f?globalThis.ExcelDiffI18n.t('app.034'):globalThis.ExcelDiffI18n.t('app.035'));invalidate();renderFiles();setMessage(state.files.every(Boolean)?globalThis.ExcelDiffI18n.t('app.036'):globalThis.ExcelDiffI18n.t('app.037'));
  }
  function pick(i=null){state.target=i;el.input.multiple=i===null;el.input.click();}
  el.input.addEventListener('change',e=>{chooseFiles(e.target.files,state.target);e.target.value='';state.target=null;});
  el.drop.addEventListener('click',e=>{if(!e.target.closest('.file-slot,button'))pick();});
  el.chips.addEventListener('click',e=>{const p=e.target.closest('[data-pick]'),r=e.target.closest('[data-remove]');if(p){e.stopPropagation();pick(+p.dataset.pick);}if(r){e.stopPropagation();const i=+r.dataset.remove;state.files[i]=null;state.status[i]=globalThis.ExcelDiffI18n.t('app.038');invalidate(globalThis.ExcelDiffI18n.t('app.039'));renderFiles();}});
  ['dragenter','dragover'].forEach(event=>el.drop.addEventListener(event,e=>{e.preventDefault();el.drop.classList.add('drop-active');}));
  ['dragleave','drop'].forEach(event=>el.drop.addEventListener(event,e=>{e.preventDefault();el.drop.classList.remove('drop-active');}));
  el.drop.addEventListener('drop',e=>{const slot=e.target.closest('[data-slot]');chooseFiles(e.dataTransfer.files,slot?+slot.dataset.slot:null);});
  window.addEventListener('dragover',e=>{if(e.dataTransfer?.types.includes('Files'))e.preventDefault();});
  window.addEventListener('drop',e=>{if(e.dataTransfer?.types.includes('Files'))e.preventDefault();});
  $('#swapFiles').addEventListener('click',()=>{state.files.reverse();state.status.reverse();invalidate(globalThis.ExcelDiffI18n.t('app.040'));renderFiles();});
  $('#clearFiles').addEventListener('click',()=>{state.files=[null,null];state.status=[globalThis.ExcelDiffI18n.t('app.041'),globalThis.ExcelDiffI18n.t('app.042')];invalidate(globalThis.ExcelDiffI18n.t('app.043'));renderFiles();});
  for(const id of ['ignoreWhitespace','ignoreCase','compareFormulas','compareStyles','compareTypes','excludeColumns','primaryKey'])$('#'+id).addEventListener('input',()=>invalidate(globalThis.ExcelDiffI18n.t('app.044')));
  $$('input[name="view"]').forEach(n=>n.addEventListener('change',()=>{state.page=0;if(state.result)renderTable();}));
  el.compare.addEventListener('click',compare);
  async function compare(){
    if(state.loading||!state.files.every(Boolean))return;
    const rev=++state.revision,files=[...state.files],o=options();state.loading=true;state.result=null;el.results.hidden=true;state.status=[globalThis.ExcelDiffI18n.t('app.045'),globalThis.ExcelDiffI18n.t('app.046')];renderFiles();setMessage(globalThis.ExcelDiffI18n.t('app.047'));
    try {
      const books=await Promise.all(files.map(async(f,i)=>{try{const book=await C.read(f);if(rev===state.revision){state.status[i]=globalThis.ExcelDiffI18n.t('app.048',{p0:book.sheets.length});renderFiles();}return book;}catch(e){if(rev===state.revision){state.status[i]=globalThis.ExcelDiffI18n.t('app.049');renderFiles();}throw new Error(`${i?'B':'A'}「${f.name}」: ${e.message}`);}}));
      if(rev!==state.revision)return;
      const result=E.compareBooks(books[0],books[1],o);
      if(rev!==state.revision)return;
      state.books=books;state.result=result;state.activeSheet=Math.max(0,result.sheets.findIndex(s=>s.diffCount>0));state.page=0;renderResult();el.results.hidden=false;
      const st=result.stats;
      if(result.totalDiffs)setMessage(globalThis.ExcelDiffI18n.t('app.050',{p0:st.changedCells,p1:st.addedRows,p2:st.deletedRows,p3:st.structural}),'ok');
      else setMessage(globalThis.ExcelDiffI18n.t('app.052',{p0:st.ignoredCells?globalThis.ExcelDiffI18n.t('app.051',{p0:st.ignoredCells}):''}),'notice');
      el.results.scrollIntoView({behavior:'smooth',block:'start'});
    }catch(e){if(rev===state.revision){state.result=null;el.results.hidden=true;setMessage(globalThis.ExcelDiffI18n.t('app.053')+e.message,'error');}}
    finally{state.loading=false;renderFiles();}
  }
  function summary(r){const s=r.stats;return globalThis.ExcelDiffI18n.t('app.054',{p0:s.changedCells,p1:s.changedRows,p2:s.addedRows,p3:s.deletedRows,p4:s.structural});}
  function scopeText(r){const o=r.options;return globalThis.ExcelDiffI18n.t('app.058',{p0:o.primaryKey?globalThis.ExcelDiffI18n.t('app.055',{p0:o.primaryKey}):globalThis.ExcelDiffI18n.t('app.056'),p1:o.ignoreWhitespace?'ON':'OFF',p2:o.ignoreCase?'ON':'OFF',p3:o.compareFormulas?'ON':'OFF',p4:o.compareStyles?'ON':'OFF',p5:o.excludeColumns||globalThis.ExcelDiffI18n.t('app.057')});}
  function renderResult(){
    const r=state.result;el.summary.textContent=`A: ${r.a} → B: ${r.b} ｜ ${summary(r)}`;
    el.tabs.innerHTML=r.sheets.map((s,i)=>`<button type="button" class="sheet-tab ${i===state.activeSheet?'active':''}" data-sheet="${i}">${esc(s.name)} (${s.diffCount})</button>`).join('');
    el.diagnostics.innerHTML=globalThis.ExcelDiffI18n.t('app.060',{p0:esc(scopeText(r)),p1:r.stats.comparedCells.toLocaleString(globalThis.ExcelDiffI18n.locale),p2:r.stats.ignoredCells,p3:r.sheets.map(s=>globalThis.ExcelDiffI18n.t('app.059',{p0:esc(s.name),p1:s.maxRowA,p2:s.maxColA,p3:s.maxRowB,p4:s.maxColB})).join('')});
    el.warnings.innerHTML=r.warnings.length?globalThis.ExcelDiffI18n.t('app.061',{p0:r.warnings.map(w=>`<p>${esc(w)}</p>`).join('')}):'';el.warnings.hidden=!r.warnings.length;renderTable();
  }
  el.tabs.addEventListener('click',e=>{const b=e.target.closest('[data-sheet]');if(!b)return;state.activeSheet=+b.dataset.sheet;state.page=0;renderResult();});
  function display(c){if(!c)return '';const v=c.v===null||c.v===undefined?globalThis.ExcelDiffI18n.t('app.062'):c.t==='b'?(c.v?'TRUE':'FALSE'):String(c.v);return c.f!==null&&c.f!==undefined?`=${c.f}${c.cached&&c.v!==null?'  ['+v+']':''}`:v;}
  function shortStyle(ch){return ch.styleFields.map(styleLabel).join('・');}
  function cellHtml(row,i,c){
    const a=row.a[i],b=row.b[i],ch=row.cells.find(x=>x.c===c);
    if(row.status==='added')return `<span class="cell-new">${esc(display(b))}</span>`;
    if(row.status==='deleted')return `<span class="cell-old">${esc(display(a))}</span>`;
    if(!ch)return esc(display(b));
    const labels=[];if(ch.typeChanged)labels.push(globalThis.ExcelDiffI18n.t('app.063',{p0:typeLabel(a.t),p1:typeLabel(b.t)}));if(ch.formulaChanged)labels.push(globalThis.ExcelDiffI18n.t('app.064'));if(ch.styleChanged)labels.push(shortStyle(ch)+globalThis.ExcelDiffI18n.t('app.065'));
    return `<div class="cell-old" title="${esc(a.addr)}">${esc(display(a))}</div><div class="cell-new" title="${esc(b.addr)}">${esc(display(b))}</div><div class="cell-meta">${esc(a.addr)} → ${esc(b.addr)}${labels.length?' ｜ '+esc(labels.join(' / ')):''}</div>${ch.styleChanged?globalThis.ExcelDiffI18n.t('app.066',{p0:esc(JSON.stringify(Object.fromEntries(ch.styleFields.map(k=>[styleLabel(k),{A:k==='rich'?a.rich:a.style?.[k],B:k==='rich'?b.rich:b.style?.[k]}])),null,2))}):''}`;
  }
  function structureHtml(s){return s.structural.length?globalThis.ExcelDiffI18n.t('app.067',{p0:s.structural.map(x=>`<details><summary>${esc(x.item)}</summary><div>A: <pre>${esc(JSON.stringify(x.before))}</pre></div><div>B: <pre>${esc(JSON.stringify(x.after))}</pre></div></details>`).join('')}):'';}
  function tableHtml(s,rows,view='diff'){
    if(!rows.length)return `<div class="empty-result">${s.diffCount?globalThis.ExcelDiffI18n.t('app.068'):globalThis.ExcelDiffI18n.t('app.069')}</div>`;
    return globalThis.ExcelDiffI18n.t('app.070',{p0:view==='cells'?'cells-only':'',p1:s.columns.map(c=>`<th>${C.colName(c)}</th>`).join(''),p2:rows.map(row=>`<tr class="${row.status}"><td><span class="status-badge status-${row.status}">${statusLabel(row.status)}</span></td><td>${row.rowA??'—'} → ${row.rowB??'—'}</td>${s.columns.map((c,i)=>`<td class="${row.cells.some(x=>x.c===c)?'diff-cell':''}">${cellHtml(row,i,c)}</td>`).join('')}</tr>`).join('')});
  }
  function renderTable(){
    const s=state.result.sheets[state.activeSheet],view=$('input[name="view"]:checked').value,rows=view==='diff'?s.rows.filter(r=>r.status!=='same'):s.rows,limit=150,totalPages=Math.max(1,Math.ceil(rows.length/limit));state.page=Math.min(state.page,totalPages-1);
    el.table.innerHTML=structureHtml(s)+tableHtml(s,rows.slice(state.page*limit,(state.page+1)*limit),view);
    el.pagination.innerHTML=globalThis.ExcelDiffI18n.t('app.072',{p0:rows.length,p1:rows.length?state.page*limit+1:0,p2:Math.min(rows.length,(state.page+1)*limit),p3:totalPages>1?globalThis.ExcelDiffI18n.t('app.071',{p0:state.page===0?'disabled':'',p1:state.page+1,p2:totalPages,p3:state.page===totalPages-1?'disabled':''}):''});
  }
  el.pagination.addEventListener('click',e=>{const b=e.target.closest('[data-page]');if(b){state.page+=+b.dataset.page;renderTable();}});
  function reportHtml(){
    const r=state.result;return globalThis.ExcelDiffI18n.t('app.073',{p0:esc(r.a),p1:esc(r.b),p2:esc(summary(r)),p3:esc(scopeText(r)),p4:esc(r.comparedAt),p5:r.version,p6:r.warnings.length?`<div class="warnings">${r.warnings.map(w=>`<p>${esc(w)}</p>`).join('')}</div>`:'',p7:r.sheets.map(s=>`<h2>${esc(s.name)} (${s.diffCount})</h2>${structureHtml(s)}${tableHtml(s,s.rows.filter(row=>row.status!=='same'))}`).join('')});
  }
  function reportSheets(){
    const r=state.result,info=[[globalThis.ExcelDiffI18n.t('app.074'),globalThis.ExcelDiffI18n.t('app.075')],[globalThis.ExcelDiffI18n.t('app.076'),r.version],[globalThis.ExcelDiffI18n.t('app.077'),r.a],[globalThis.ExcelDiffI18n.t('app.078'),r.b],[globalThis.ExcelDiffI18n.t('app.079'),r.comparedAt],[globalThis.ExcelDiffI18n.t('app.080'),summary(r)],[globalThis.ExcelDiffI18n.t('app.081'),scopeText(r)],[globalThis.ExcelDiffI18n.t('app.082'),r.stats.ignoredCells],...r.warnings.map(w=>[globalThis.ExcelDiffI18n.t('app.083'),w])];
    const cells=[[globalThis.ExcelDiffI18n.t('app.084'),globalThis.ExcelDiffI18n.t('app.085'),globalThis.ExcelDiffI18n.t('app.086'),globalThis.ExcelDiffI18n.t('app.087'),globalThis.ExcelDiffI18n.t('app.088'),globalThis.ExcelDiffI18n.t('app.089'),globalThis.ExcelDiffI18n.t('app.090'),globalThis.ExcelDiffI18n.t('app.091'),globalThis.ExcelDiffI18n.t('app.092'),globalThis.ExcelDiffI18n.t('app.093'),globalThis.ExcelDiffI18n.t('app.094'),globalThis.ExcelDiffI18n.t('app.095')]],rows=[[globalThis.ExcelDiffI18n.t('app.096'),globalThis.ExcelDiffI18n.t('app.097'),globalThis.ExcelDiffI18n.t('app.098'),globalThis.ExcelDiffI18n.t('app.099'),globalThis.ExcelDiffI18n.t('app.100'),globalThis.ExcelDiffI18n.t('app.101'),globalThis.ExcelDiffI18n.t('app.102')]],structures=[[globalThis.ExcelDiffI18n.t('app.103'),globalThis.ExcelDiffI18n.t('app.104'),'A','B']];
    for(const s of r.sheets){for(const row of s.rows){for(const ch of row.cells)cells.push([s.name,ch.a.addr,ch.b.addr,[ch.valueChanged?globalThis.ExcelDiffI18n.t('app.105'):'',ch.typeChanged?globalThis.ExcelDiffI18n.t('app.106'):'',ch.formulaChanged?globalThis.ExcelDiffI18n.t('app.107'):'',ch.styleChanged?shortStyle(ch):''].filter(Boolean).join(' / '),ch.a.v??'',ch.b.v??'',typeLabel(ch.a.t),typeLabel(ch.b.t),ch.a.f!==null?'='+ch.a.f:'',ch.b.f!==null?'='+ch.b.f:'',ch.styleChanged?JSON.stringify(ch.a.style):'',ch.styleChanged?JSON.stringify(ch.b.style):'']);if(row.status==='added'||row.status==='deleted')s.columns.forEach((c,i)=>rows.push([s.name,statusLabel(row.status),row.rowA??'',row.rowB??'',C.colName(c),row.status==='deleted'?display(row.a[i]):'',row.status==='added'?display(row.b[i]):'']));}for(const ch of s.structural)structures.push([s.name,ch.item,JSON.stringify(ch.before),JSON.stringify(ch.after)]);}
    return [{name:globalThis.ExcelDiffI18n.t('app.108'),rows:info},{name:globalThis.ExcelDiffI18n.t('app.109'),rows:cells},{name:globalThis.ExcelDiffI18n.t('app.110'),rows},{name:globalThis.ExcelDiffI18n.t('app.111'),rows:structures}];
  }
  function download(blob,name){const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=name;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),30000);}
  $('#exportExcel').addEventListener('click',()=>{if(state.result)try{download(C.writeReport(reportSheets()),'excel-diff-report.xlsx');}catch(e){setMessage(globalThis.ExcelDiffI18n.t('app.112')+e.message,'error');}});
  $('#exportHtml').addEventListener('click',()=>{if(state.result)download(new Blob([reportHtml()],{type:'text/html;charset=utf-8'}),'excel-diff-report.html');});
  $('#exportPdf').addEventListener('click',()=>{if(!state.result)return;const frame=document.createElement('iframe');frame.style.cssText='position:fixed;left:-10000px;width:1200px;height:800px;border:0';frame.title=globalThis.ExcelDiffI18n.t('app.113');frame.onload=()=>{setTimeout(()=>{frame.contentWindow.focus();frame.contentWindow.print();},150);};frame.srcdoc=reportHtml();document.body.append(frame);setTimeout(()=>frame.remove(),120000);setMessage(globalThis.ExcelDiffI18n.t('app.114'));});
  renderFiles();setMessage(globalThis.ExcelDiffI18n.t('app.115'));
})();
