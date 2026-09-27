(() => {
  'use strict';
  const $=s=>document.querySelector(s),$$=s=>[...document.querySelectorAll(s)];
  const C=globalThis.ExcelDiffCodec,E=globalThis.ExcelDiffEngine;
  const state={files:[null,null],status:['未選択','未選択'],books:[],result:null,activeSheet:0,page:0,revision:0,loading:false,target:null};
  const el={input:$('#fileInput'),drop:$('#dropZone'),chips:$('#fileChips'),message:$('#message'),compare:$('#compareBtn'),results:$('#results'),summary:$('#resultSummary'),tabs:$('#sheetTabs'),table:$('#resultTableWrap'),diagnostics:$('#diagnostics'),warnings:$('#resultWarnings'),pagination:$('#pagination')};
  const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const fmt=n=>n<1024?`${n} B`:n<1048576?`${(n/1024).toFixed(1)} KB`:`${(n/1048576).toFixed(1)} MB`;
  const statusLabel=s=>({added:'追加',deleted:'削除',changed:'変更',same:'同一'}[s]||s);
  const typeLabel=t=>({n:'数値',s:'文字列',b:'論理値',e:'エラー',d:'日付',blank:'空セル'}[t]||t);
  const styleLabel=k=>({font:'フォント・文字色',fill:'背景色',border:'罫線',alignment:'配置',numFmt:'表示形式',protection:'保護',flags:'その他書式',rich:'部分文字書式'}[k]||k);
  const setMessage=(s,kind='')=>{el.message.textContent=s;el.message.className='message'+(kind?' '+kind:'');};
  function options(){return {ignoreWhitespace:$('#ignoreWhitespace').checked,ignoreCase:$('#ignoreCase').checked,compareFormulas:$('#compareFormulas').checked,compareStyles:$('#compareStyles').checked,compareTypes:$('#compareTypes').checked,excludeColumns:$('#excludeColumns').value,primaryKey:$('#primaryKey').value};}
  function invalidate(message){state.revision++;state.result=null;state.books=[];el.results.hidden=true;el.compare.disabled=!state.files.every(Boolean)||state.loading;if(message)setMessage(message);}
  function renderFiles(){
    el.chips.innerHTML=state.files.map((f,i)=>`<div class="file-slot ${f?'selected':'empty'}" data-slot="${i}"><span class="file-slot-label">${i?'B・比較先（変更後）':'A・比較元（変更前）'}</span><span class="file-slot-name" title="${esc(f?.name||'未選択')}">${esc(f?.name||'未選択')}</span><span class="file-slot-meta">${f?fmt(f.size)+' · ':''}${esc(state.status[i])}</span><div class="slot-actions"><button type="button" data-pick="${i}">${f?'選び直す':'ファイルを選択'}</button>${f?`<button type="button" data-remove="${i}" aria-label="${i?'B':'A'}のファイルを解除">解除</button>`:''}</div></div>`).join('');
    el.compare.disabled=!state.files.every(Boolean)||state.loading;
    $('#swapFiles').disabled=!state.files.every(Boolean)||state.loading;
  }
  function chooseFiles(list,target=null){
    const incoming=[...list];if(!incoming.length)return;
    if(incoming.some(f=>!/^.+\.(xlsx|xls|csv)$/i.test(f.name))){invalidate();setMessage('XLSX / XLS / CSV以外のファイルが含まれています。選択し直してください。','error');return;}
    if(incoming.length>2||(target!==null&&incoming.length!==1)){setMessage(target===null?'比較できるファイルは2つまでです。':'この枠には1つのファイルを選択してください。','error');return;}
    if(incoming.length===2){state.files=[incoming[0],incoming[1]];state.status=['選択済み（未比較）','選択済み（未比較）'];}
    else{const i=target!==null?target:!state.files[0]?0:1;state.files[i]=incoming[0];state.status[i]='選択済み（未比較）';}
    state.status=state.files.map(f=>f?'選択済み（未比較）':'未選択');invalidate();renderFiles();setMessage(state.files.every(Boolean)?'AとBを選択しました。ファイル名を確認して「比較する」を押してください。':'もう1つのファイルを選択してください。');
  }
  function pick(i=null){state.target=i;el.input.multiple=i===null;el.input.click();}
  el.input.addEventListener('change',e=>{chooseFiles(e.target.files,state.target);e.target.value='';state.target=null;});
  el.drop.addEventListener('click',e=>{if(!e.target.closest('.file-slot,button'))pick();});
  el.chips.addEventListener('click',e=>{const p=e.target.closest('[data-pick]'),r=e.target.closest('[data-remove]');if(p){e.stopPropagation();pick(+p.dataset.pick);}if(r){e.stopPropagation();const i=+r.dataset.remove;state.files[i]=null;state.status[i]='未選択';invalidate('ファイルを解除しました。');renderFiles();}});
  ['dragenter','dragover'].forEach(event=>el.drop.addEventListener(event,e=>{e.preventDefault();el.drop.classList.add('drop-active');}));
  ['dragleave','drop'].forEach(event=>el.drop.addEventListener(event,e=>{e.preventDefault();el.drop.classList.remove('drop-active');}));
  el.drop.addEventListener('drop',e=>{const slot=e.target.closest('[data-slot]');chooseFiles(e.dataTransfer.files,slot?+slot.dataset.slot:null);});
  window.addEventListener('dragover',e=>{if(e.dataTransfer?.types.includes('Files'))e.preventDefault();});
  window.addEventListener('drop',e=>{if(e.dataTransfer?.types.includes('Files'))e.preventDefault();});
  $('#swapFiles').addEventListener('click',()=>{state.files.reverse();state.status.reverse();invalidate('比較元と比較先を入れ替えました。もう一度比較してください。');renderFiles();});
  $('#clearFiles').addEventListener('click',()=>{state.files=[null,null];state.status=['未選択','未選択'];invalidate('選択をクリアしました。');renderFiles();});
  for(const id of ['ignoreWhitespace','ignoreCase','compareFormulas','compareStyles','compareTypes','excludeColumns','primaryKey'])$('#'+id).addEventListener('input',()=>invalidate('比較設定を変更しました。もう一度「比較する」を押してください。'));
  $$('input[name="view"]').forEach(n=>n.addEventListener('change',()=>{state.page=0;if(state.result)renderTable();}));
  el.compare.addEventListener('click',compare);
  async function compare(){
    if(state.loading||!state.files.every(Boolean))return;
    const rev=++state.revision,files=[...state.files],o=options();state.loading=true;state.result=null;el.results.hidden=true;state.status=['読み込み中…','読み込み中…'];renderFiles();setMessage('ファイルを読み込んで比較しています…');
    try {
      const books=await Promise.all(files.map(async(f,i)=>{try{const book=await C.read(f);if(rev===state.revision){state.status[i]=`${book.sheets.length}シート読み込み済み`;renderFiles();}return book;}catch(e){if(rev===state.revision){state.status[i]='読み込み失敗';renderFiles();}throw new Error(`${i?'B':'A'}「${f.name}」: ${e.message}`);}}));
      if(rev!==state.revision)return;
      const result=E.compareBooks(books[0],books[1],o);
      if(rev!==state.revision)return;
      state.books=books;state.result=result;state.activeSheet=Math.max(0,result.sheets.findIndex(s=>s.diffCount>0));state.page=0;renderResult();el.results.hidden=false;
      const st=result.stats;
      if(result.totalDiffs)setMessage(`比較完了：変更 ${st.changedCells}セル / 追加 ${st.addedRows}行 / 削除 ${st.deletedRows}行 / 構造・書式設定 ${st.structural}件`,'ok');
      else setMessage(`現在の設定・対応項目では差分0件です。${st.ignoredCells?`設定で無視された差分は${st.ignoredCells}セルあります。`:''}下の比較範囲と注意事項を確認してください。`,'notice');
      el.results.scrollIntoView({behavior:'smooth',block:'start'});
    }catch(e){if(rev===state.revision){state.result=null;el.results.hidden=true;setMessage('比較を完了できませんでした。'+e.message,'error');}}
    finally{state.loading=false;renderFiles();}
  }
  function summary(r){const s=r.stats;return `変更 ${s.changedCells}セル（${s.changedRows}行） / 追加 ${s.addedRows}行 / 削除 ${s.deletedRows}行 / 構造・書式設定 ${s.structural}件`;}
  function scopeText(r){const o=r.options;return `先頭行を含む全行を比較 / ${o.primaryKey?`主キー: ${o.primaryKey}`:'同じセル位置で比較'} / 空白無視: ${o.ignoreWhitespace?'ON':'OFF'} / 大文字小文字無視: ${o.ignoreCase?'ON':'OFF'} / 数式: ${o.compareFormulas?'ON':'OFF'} / 書式: ${o.compareStyles?'ON':'OFF'} / 除外列: ${o.excludeColumns||'なし'}`;}
  function renderResult(){
    const r=state.result;el.summary.textContent=`A: ${r.a} → B: ${r.b} ｜ ${summary(r)}`;
    el.tabs.innerHTML=r.sheets.map((s,i)=>`<button type="button" class="sheet-tab ${i===state.activeSheet?'active':''}" data-sheet="${i}">${esc(s.name)} (${s.diffCount})</button>`).join('');
    el.diagnostics.innerHTML=`<p>${esc(scopeText(r))}</p><p>比較したセルの組数: ${r.stats.comparedCells.toLocaleString()} / 設定で無視したセル差分: ${r.stats.ignoredCells}</p><details><summary>読み込み範囲を確認</summary>${r.sheets.map(s=>`<p><strong>${esc(s.name)}</strong> ｜ A: ${s.maxRowA}行・${s.maxColA}列 ／ B: ${s.maxRowB}行・${s.maxColB}列</p>`).join('')}</details>`;
    el.warnings.innerHTML=r.warnings.length?`<strong>確認事項</strong>${r.warnings.map(w=>`<p>${esc(w)}</p>`).join('')}`:'';el.warnings.hidden=!r.warnings.length;renderTable();
  }
  el.tabs.addEventListener('click',e=>{const b=e.target.closest('[data-sheet]');if(!b)return;state.activeSheet=+b.dataset.sheet;state.page=0;renderResult();});
  function display(c){if(!c)return '';const v=c.v===null||c.v===undefined?'（空）':c.t==='b'?(c.v?'TRUE':'FALSE'):String(c.v);return c.f!==null&&c.f!==undefined?`=${c.f}${c.cached&&c.v!==null?'  ['+v+']':''}`:v;}
  function shortStyle(ch){return ch.styleFields.map(styleLabel).join('・');}
  function cellHtml(row,i,c){
    const a=row.a[i],b=row.b[i],ch=row.cells.find(x=>x.c===c);
    if(row.status==='added')return `<span class="cell-new">${esc(display(b))}</span>`;
    if(row.status==='deleted')return `<span class="cell-old">${esc(display(a))}</span>`;
    if(!ch)return esc(display(b));
    const labels=[];if(ch.typeChanged)labels.push(`型: ${typeLabel(a.t)} → ${typeLabel(b.t)}`);if(ch.formulaChanged)labels.push('数式変更');if(ch.styleChanged)labels.push(shortStyle(ch)+'変更');
    return `<div class="cell-old" title="${esc(a.addr)}">${esc(display(a))}</div><div class="cell-new" title="${esc(b.addr)}">${esc(display(b))}</div><div class="cell-meta">${esc(a.addr)} → ${esc(b.addr)}${labels.length?' ｜ '+esc(labels.join(' / ')):''}</div>${ch.styleChanged?`<details class="style-detail"><summary>書式の差分</summary><pre>${esc(JSON.stringify(Object.fromEntries(ch.styleFields.map(k=>[styleLabel(k),{A:k==='rich'?a.rich:a.style?.[k],B:k==='rich'?b.rich:b.style?.[k]}])),null,2))}</pre></details>`:''}`;
  }
  function structureHtml(s){return s.structural.length?`<div class="structural-diffs"><strong>構造・書式設定の変更</strong>${s.structural.map(x=>`<details><summary>${esc(x.item)}</summary><div>A: <pre>${esc(JSON.stringify(x.before))}</pre></div><div>B: <pre>${esc(JSON.stringify(x.after))}</pre></div></details>`).join('')}</div>`:'';}
  function tableHtml(s,rows,view='diff'){
    if(!rows.length)return `<div class="empty-result">${s.diffCount?'このシートにはセル行以外の変更があります。':'このシートは現在の設定・対応項目では差分0件です。'}</div>`;
    return `<table class="diff-table ${view==='cells'?'cells-only':''}"><thead><tr><th>状態</th><th>A行 → B行</th>${s.columns.map(c=>`<th>${C.colName(c)}</th>`).join('')}</tr></thead><tbody>${rows.map(row=>`<tr class="${row.status}"><td><span class="status-badge status-${row.status}">${statusLabel(row.status)}</span></td><td>${row.rowA??'—'} → ${row.rowB??'—'}</td>${s.columns.map((c,i)=>`<td class="${row.cells.some(x=>x.c===c)?'diff-cell':''}">${cellHtml(row,i,c)}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
  }
  function renderTable(){
    const s=state.result.sheets[state.activeSheet],view=$('input[name="view"]:checked').value,rows=view==='diff'?s.rows.filter(r=>r.status!=='same'):s.rows,limit=150,totalPages=Math.max(1,Math.ceil(rows.length/limit));state.page=Math.min(state.page,totalPages-1);
    el.table.innerHTML=structureHtml(s)+tableHtml(s,rows.slice(state.page*limit,(state.page+1)*limit),view);
    el.pagination.innerHTML=`<span>${rows.length}行中 ${rows.length?state.page*limit+1:0}–${Math.min(rows.length,(state.page+1)*limit)}行を表示（レポートには全差分を出力）</span>${totalPages>1?`<button type="button" data-page="-1" ${state.page===0?'disabled':''}>前へ</button><span>${state.page+1} / ${totalPages}</span><button type="button" data-page="1" ${state.page===totalPages-1?'disabled':''}>次へ</button>`:''}`;
  }
  el.pagination.addEventListener('click',e=>{const b=e.target.closest('[data-page]');if(b){state.page+=+b.dataset.page;renderTable();}});
  function reportHtml(){
    const r=state.result;return `<!doctype html><html lang="ja"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Excel Diff 差分レポート</title><style>body{font-family:system-ui,sans-serif;color:#20362a;margin:28px;font-size:12px}table{border-collapse:collapse;width:100%;margin:15px 0}th,td{border:1px solid #d9e1dd;padding:7px;text-align:left;vertical-align:top;white-space:pre-wrap;overflow-wrap:anywhere}th{background:#eaf4ee}.cell-old{color:#a52e2e;text-decoration:line-through}.cell-new{color:#167444}.cell-meta{font-size:10px;color:#53685e}.diff-cell{background:#fff3ce}.added{background:#eaf9ee}.deleted{background:#ffeded}pre{white-space:pre-wrap;overflow-wrap:anywhere}h2{margin-top:28px}details{margin:8px 0}.warnings{padding:12px;background:#fff7df}@page{size:landscape;margin:12mm}@media print{body{margin:0;font-size:9px}details:not([open])>*:not(summary){display:block}thead{display:table-header-group}tr{break-inside:avoid}h2{break-after:avoid}}</style></head><body><h1>Excel Diff 差分レポート</h1><p>A: ${esc(r.a)} → B: ${esc(r.b)}</p><p>${esc(summary(r))}</p><p>${esc(scopeText(r))}</p><p>作成: ${esc(r.comparedAt)} / v${r.version}</p><p>保存値と数式を比較しています。数式の再計算は行いません。</p>${r.warnings.length?`<div class="warnings">${r.warnings.map(w=>`<p>${esc(w)}</p>`).join('')}</div>`:''}${r.sheets.map(s=>`<h2>${esc(s.name)} (${s.diffCount})</h2>${structureHtml(s)}${tableHtml(s,s.rows.filter(row=>row.status!=='same'))}`).join('')}<p>比較対象外: 条件付き書式の評価結果、画像・図形、コメント、リンク先、マクロ、印刷設定、テーブルスタイル。</p></body></html>`;
  }
  function reportSheets(){
    const r=state.result,info=[['項目','内容'],['バージョン',r.version],['A 比較元',r.a],['B 比較先',r.b],['比較日時',r.comparedAt],['概要',summary(r)],['設定',scopeText(r)],['無視したセル差分',r.stats.ignoredCells],...r.warnings.map(w=>['確認事項',w])];
    const cells=[['シート','Aセル','Bセル','差分種別','A保存値','B保存値','A型','B型','A数式','B数式','A書式','B書式']],rows=[['シート','状態','A行','B行','列','A値','B値']],structures=[['シート','項目','A','B']];
    for(const s of r.sheets){for(const row of s.rows){for(const ch of row.cells)cells.push([s.name,ch.a.addr,ch.b.addr,[ch.valueChanged?'値':'',ch.typeChanged?'型':'',ch.formulaChanged?'数式':'',ch.styleChanged?shortStyle(ch):''].filter(Boolean).join(' / '),ch.a.v??'',ch.b.v??'',typeLabel(ch.a.t),typeLabel(ch.b.t),ch.a.f!==null?'='+ch.a.f:'',ch.b.f!==null?'='+ch.b.f:'',ch.styleChanged?JSON.stringify(ch.a.style):'',ch.styleChanged?JSON.stringify(ch.b.style):'']);if(row.status==='added'||row.status==='deleted')s.columns.forEach((c,i)=>rows.push([s.name,statusLabel(row.status),row.rowA??'',row.rowB??'',C.colName(c),row.status==='deleted'?display(row.a[i]):'',row.status==='added'?display(row.b[i]):'']));}for(const ch of s.structural)structures.push([s.name,ch.item,JSON.stringify(ch.before),JSON.stringify(ch.after)]);}
    return [{name:'比較概要',rows:info},{name:'セル差分',rows:cells},{name:'行追加削除',rows},{name:'構造差分',rows:structures}];
  }
  function download(blob,name){const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=name;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),30000);}
  $('#exportExcel').addEventListener('click',()=>{if(state.result)try{download(C.writeReport(reportSheets()),'excel-diff-report.xlsx');}catch(e){setMessage('Excel出力に失敗しました: '+e.message,'error');}});
  $('#exportHtml').addEventListener('click',()=>{if(state.result)download(new Blob([reportHtml()],{type:'text/html;charset=utf-8'}),'excel-diff-report.html');});
  $('#exportPdf').addEventListener('click',()=>{if(!state.result)return;const frame=document.createElement('iframe');frame.style.cssText='position:fixed;left:-10000px;width:1200px;height:800px;border:0';frame.title='PDF印刷用レポート';frame.onload=()=>{setTimeout(()=>{frame.contentWindow.focus();frame.contentWindow.print();},150);};frame.srcdoc=reportHtml();document.body.append(frame);setTimeout(()=>frame.remove(),120000);setMessage('印刷画面で「PDFに保存」を選択してください。全シートの差分レポートを出力します。');});
  renderFiles();setMessage('A・比較元とB・比較先を選択してください。');
})();
