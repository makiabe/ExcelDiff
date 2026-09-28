/* Pure comparison layer. UI filters never alter detection or exports. */
(() => {
  'use strict';
  const C=globalThis.ExcelDiffCodec;
  const stable=v=>JSON.stringify(v??null);
  const DEFAULTS={ignoreWhitespace:false,ignoreCase:false,compareFormulas:true,compareStyles:true,compareTypes:true,excludeColumns:'',primaryKey:''};
  const blank=(sheet,r,c)=>({v:null,t:'blank',f:null,rich:'',formulaMeta:'',addr:C.colName(c)+r,style:sheet?.rowStyles?.get(r)||[...(sheet?.colStyles||[])].reverse().find(x=>c>=x.min&&c<=x.max)?.style||sheet?.defaultStyle||{}});
  const cell=(sheet,r,c)=>sheet?.rows.get(r)?.get(c)||blank(sheet,r,c);
  const isEmpty=x=>x.v===null||x.v===undefined||x.v==='';
  function norm(v,o) {let s=String(v??'');if(o.ignoreWhitespace)s=s.replace(/\s/g,'');if(o.ignoreCase)s=s.toLowerCase();return s;}
  function valueKey(x,o) {if(isEmpty(x))return 'blank:';return (o.compareTypes ? x.t+':' : '')+norm(x.v,o);}
  function changes(a,b,o,styleMode='full') {
    const fields=styleMode==='numberOnly'?['numFmt']:['font','fill','border','alignment','numFmt','protection','flags'];
    const styleFields=fields.filter(k=>stable(a.style?.[k])!==stable(b.style?.[k]));
    if(styleMode==='full'&&(a.rich||'')!==(b.rich||''))styleFields.push('rich');
    return {valueChanged:valueKey(a,o)!==valueKey(b,o),typeChanged:o.compareTypes&&!isEmpty(a)&&!isEmpty(b)&&a.t!==b.t,
      formulaChanged:!!o.compareFormulas&&((a.f??null)!==(b.f??null)||(a.formulaMeta||'')!==(b.formulaMeta||'')),
      styleChanged:!!o.compareStyles&&styleFields.length>0,styleFields};
  }
  const hasChange=x=>x.valueChanged||x.formulaChanged||x.styleChanged;
  function resolveColumn(token,sheet) {
    const names=[...(sheet.rows.get(1)||[])].filter(([c,v])=>String(v.v??'').trim().toLowerCase()===token.toLowerCase());
    if(names.length>1)throw new Error(globalThis.ExcelDiffI18n.t('diff-engine.001',{p0:sheet.name,p1:token}));
    let idx=names.length?names[0][0]:/^\d+$/.test(token)?+token-1:/^[A-Za-z]+$/.test(token)?C.colIndex(token):-1;
    if(idx<0||idx>=sheet.maxCol)throw new Error(globalThis.ExcelDiffI18n.t('diff-engine.002',{p0:sheet.name,p1:token}));
    return idx;
  }
  function tokens(s) {return String(s||'').split(/[,、]/).map(t=>t.trim()).filter(Boolean);}
  function columnsOf(sa,sb) {const cols=new Set();for(const s of [sa,sb])if(s)for(const row of s.rows.values())for(const c of row.keys())cols.add(c);return [...cols].sort((x,y)=>x-y);}
  function resolveExclude(input,sa,sb) {
    const result=new Set();for(const token of tokens(input)) {
      let found=false;for(const s of [sa,sb])if(s) {try{result.add(resolveColumn(token,s));found=true;}catch(e){if(!e.message.includes(globalThis.ExcelDiffI18n.t('diff-engine.003')))throw e;}}
      if(!found)throw new Error(globalThis.ExcelDiffI18n.t('diff-engine.004',{p0:token}));
    }return result;
  }
  function keyMap(sheet,indices,o) {
    const map=new Map();
    for(const [r]of sheet.rows) {
      if(r===1)continue;
      const parts=indices.map(c=>cell(sheet,r,c));
      if(parts.some(v=>norm(v.v,o)===''))throw new Error(globalThis.ExcelDiffI18n.t('diff-engine.005',{p0:sheet.name,p1:r}));
      const key=JSON.stringify(parts.map(v=>valueKey(v,o)));
      if(map.has(key))throw new Error(globalThis.ExcelDiffI18n.t('diff-engine.006',{p0:sheet.name,p1:map.get(key),p2:r}));
      map.set(key,r);
    }
    return map;
  }
  function compareSheet(sa,sb,options,kindA,kindB) {
    const o={...options},name=sb?.name||sa?.name,allCols=columnsOf(sa,sb),exclude=resolveExclude(o.excludeColumns,sa,sb),columns=allCols.filter(c=>!exclude.has(c));
    if(allCols.length&&!columns.length)throw new Error(globalThis.ExcelDiffI18n.t('diff-engine.007',{p0:name}));
    if(allCols.length * Math.max(sa?.rows.size||0,sb?.rows.size||0) > 2000000) throw new Error(globalThis.ExcelDiffI18n.t('diff-engine.008',{p0:name}));
    const rows=[],structural=[],warnings=[],stats={changedCells:0,changedRows:0,addedRows:0,deletedRows:0,structural:0,comparedCells:0,ignoredCells:0};
    const styleMode=kindA==='xls'||kindB==='xls'?'numberOnly':'full';
    if(kindA==='csv'||kindB==='csv') {o.compareStyles=false;o.compareTypes=false;}
    const event=(item,before,after)=>{if(stable(before)!==stable(after))structural.push({item,before,after});};
    if(!sa||!sb)structural.push({item:globalThis.ExcelDiffI18n.t('diff-engine.009'),before:sa?globalThis.ExcelDiffI18n.t('diff-engine.010'):globalThis.ExcelDiffI18n.t('diff-engine.011'),after:sb?globalThis.ExcelDiffI18n.t('diff-engine.012'):globalThis.ExcelDiffI18n.t('diff-engine.013')});
    else {
      event(globalThis.ExcelDiffI18n.t('diff-engine.014'),sa.name,sb.name);
      if(kindA!=='csv'&&kindB!=='csv') {
        event(globalThis.ExcelDiffI18n.t('diff-engine.015'),sa.metadata?.visibility,sb.metadata?.visibility);
        if(o.compareStyles)for(const [key,label]of Object.entries({merges:globalThis.ExcelDiffI18n.t('diff-engine.016'),rowDimensions:globalThis.ExcelDiffI18n.t('diff-engine.017'),columnDimensions:globalThis.ExcelDiffI18n.t('diff-engine.018'),rowStyles:globalThis.ExcelDiffI18n.t('diff-engine.019'),columnStyles:globalThis.ExcelDiffI18n.t('diff-engine.020'),defaults:globalThis.ExcelDiffI18n.t('diff-engine.021')}))event(label,sa.metadata?.[key],sb.metadata?.[key]);
      }
    }
    const pairs=[];
    if(sa&&sb&&o.primaryKey.trim()) {
      const ts=tokens(o.primaryKey),ia=ts.map(k=>resolveColumn(k,sa)),ib=ts.map(k=>resolveColumn(k,sb));
      // The header is compared, never discarded. Row 1 is only excluded from key lookup.
      if(sa.rows.has(1)||sb.rows.has(1))pairs.push([1,1,globalThis.ExcelDiffI18n.t('diff-engine.022')]);
      const ma=keyMap(sa,ia,o),mb=keyMap(sb,ib,o);
      for(const key of new Set([...ma.keys(),...mb.keys()]))pairs.push([ma.get(key)??null,mb.get(key)??null,key]);
    } else {
      const numbers=new Set([...(sa?.rows.keys()||[]),...(sb?.rows.keys()||[])]);
      for(const r of [...numbers].sort((x,y)=>x-y))pairs.push([sa&&r<=sa.maxRow?r:null,sb&&r<=sb.maxRow?r:null,String(r)]);
    }
    for(const [ra,rb,key]of pairs) {
      if(ra===null||rb===null) {
        const status=ra===null?'added':'deleted';stats[ra===null?'addedRows':'deletedRows']++;
        rows.push({status,rowA:ra,rowB:rb,key,cells:[],a:columns.map(c=>cell(sa,ra,c)),b:columns.map(c=>cell(sb,rb,c))});continue;
      }
      const cells=[];
      for(const c of allCols) {
        const a=cell(sa,ra,c),b=cell(sb,rb,c),ch=changes(a,b,o,styleMode);
        const raw=changes(a,b,{...o,ignoreWhitespace:false,ignoreCase:false,compareFormulas:true,compareStyles:o.compareStyles,compareTypes:o.compareTypes},styleMode);
        if(exclude.has(c)) {if(hasChange(raw))stats.ignoredCells++;continue;}
        stats.comparedCells++;
        if(hasChange(ch))cells.push({c,a,b,...ch});else if(hasChange(raw))stats.ignoredCells++;
      }
      stats.changedCells+=cells.length;if(cells.length)stats.changedRows++;
      rows.push({status:cells.length?'changed':'same',rowA:ra,rowB:rb,key,cells,a:columns.map(c=>cell(sa,ra,c)),b:columns.map(c=>cell(sb,rb,c))});
    }
    stats.structural=structural.length;
    return {name,nameA:sa?.name||'',nameB:sb?.name||'',columns,rows,structural,stats,warnings:[...(sa?.warnings||[]),...(sb?.warnings||[])],maxRowA:sa?.maxRow||0,maxRowB:sb?.maxRow||0,maxColA:sa?.maxCol||0,maxColB:sb?.maxCol||0,diffCount:stats.changedCells+stats.addedRows+stats.deletedRows+stats.structural};
  }
  function compareBooks(a,b,options={}) {
    const o={...DEFAULTS,...options},pairs=[];
    if(a.sheets.length===1&&b.sheets.length===1) {
      const sa=a.sheets[0],sb=b.sheets[0];
      // CSV has no meaningful sheet name. Do not report a synthetic rename.
      pairs.push([sa,b.kind==='csv'?{...sb,name:sa.name}:a.kind==='csv'?sb:sb]);
      if(a.kind==='csv')pairs[0][0]={...sa,name:sb.name};
    } else {
      const ma=new Map(a.sheets.map(s=>[s.name,s])),mb=new Map(b.sheets.map(s=>[s.name,s]));
      for(const name of new Set([...ma.keys(),...mb.keys()]))pairs.push([ma.get(name),mb.get(name)]);
    }
    const sheets=pairs.map(([sa,sb])=>compareSheet(sa,sb,o,a.kind,b.kind));
    const warnings=[...a.warnings,...b.warnings,...sheets.flatMap(s=>s.warnings)];
    if((a.kind==='csv'||b.kind==='csv')&&(o.compareStyles||o.compareTypes))warnings.push(globalThis.ExcelDiffI18n.t('diff-engine.023'));
    if(a.kind==='xlsx'&&b.kind==='xlsx') {
      for(const [key,label]of [['date1904',globalThis.ExcelDiffI18n.t('diff-engine.024')],['theme',globalThis.ExcelDiffI18n.t('diff-engine.025')]]) {
        if(key==='theme'&&!o.compareStyles)continue;
        if(stable(a.metadata[key])!==stable(b.metadata[key])){sheets[0].structural.push({item:label,before:a.metadata[key],after:b.metadata[key]});sheets[0].stats.structural++;sheets[0].diffCount++;}
      }
    }
    if(a.bytes&&b.bytes&&a.bytes.length===b.bytes.length&&a.bytes.every((v,i)=>v===b.bytes[i]))warnings.push(globalThis.ExcelDiffI18n.t('diff-engine.026'));
    const stats={changedCells:0,changedRows:0,addedRows:0,deletedRows:0,structural:0,comparedCells:0,ignoredCells:0};
    for(const s of sheets)for(const k of Object.keys(stats))stats[k]+=s.stats[k];
    return {a:a.name,b:b.name,sheets,options:o,stats,totalDiffs:sheets.reduce((n,s)=>n+s.diffCount,0),warnings:[...new Set(warnings)],comparedAt:new Date().toISOString(),version:'1.2.0'};
  }
  globalThis.ExcelDiffEngine={compareBooks,compareSheet,changes,resolveColumn,DEFAULTS,cell,valueKey};
})();
