/* Excel Diff 1.1.0 — local XLSX/CSV reader and report writer.
 * XLSX: ZIP + Open XML; no network access and no formula execution.
 * Reject unsupported/corrupt structures rather than returning an empty workbook.
 */
(() => {
  'use strict';
  const enc = new TextEncoder();
  const LIMITS = { file: 50 * 1024 * 1024, part: 80 * 1024 * 1024, expanded: 240 * 1024 * 1024, cells: 500000 };
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const children = (n, name) => [...(n?.children || [])].filter(x => !name || x.localName === name);
  const child = (n, name) => children(n, name)[0];
  const nodes = (n, name) => [...n.getElementsByTagNameNS('*', name)];
  const a = (n, name, fallback = '') => n?.getAttribute(name) ?? fallback;
  const text = n => n?.textContent ?? '';
  const xmlEscape = s => String(s ?? '').replace(/_x([0-9a-f]{4})_/gi, '_x005F_x$1_').replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, c => '_x' + c.charCodeAt(0).toString(16).padStart(4, '0') + '_').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
  const unescapeExcel = s => String(s).replace(/_x([0-9a-f]{4})_/gi, (_, hex) => String.fromCharCode(parseInt(hex, 16)));
  function colName(c) { let s = ''; for (c++; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + (c - 1) % 26) + s; return s; }
  function colIndex(s) { let c = 0; for (const ch of s.toUpperCase()) c = c * 26 + ch.charCodeAt(0) - 64; return c - 1; }
  function address(s) { const m = /^\$?([A-Z]{1,3})\$?([1-9]\d*)$/i.exec(s); if (!m) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.001',{p0:s})); const c = colIndex(m[1]), r = +m[2]; if (c > 16383 || r > 1048576) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.002',{p0:s})); return {r, c}; }
  const crcTable = Uint32Array.from({length:256}, (_, n) => { for(let k = 0; k < 8; k++) n = (n & 1) ? 0xedb88320 ^ (n >>> 1) : n >>> 1; return n >>> 0; });
  function crc32(bytes) { let c = 0xffffffff; for (const b of bytes) c = crcTable[(c ^ b) & 255] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; }
  function decodeXml(bytes) { if (bytes[0] === 255 && bytes[1] === 254) return new TextDecoder('utf-16le', {fatal:true}).decode(bytes); if (bytes[0] === 254 && bytes[1] === 255) return new TextDecoder('utf-16be', {fatal:true}).decode(bytes); return new TextDecoder('utf-8', {fatal:true}).decode(bytes); }
  function parseXml(bytes, path) { const s = decodeXml(bytes); if (/<!DOCTYPE|<!ENTITY/i.test(s)) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.003',{p0:path})); const doc = new DOMParser().parseFromString(s, 'application/xml'); if (nodes(doc, 'parsererror').length) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.004',{p0:path})); return doc; }
  function resolvePath(base, target) { if (/^[a-z][a-z0-9+.-]*:/i.test(target)) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.005')); const parts = (target.startsWith('/') ? target.slice(1) : base.slice(0, base.lastIndexOf('/') + 1) + target).split('/'); const out = []; for (const p of parts) { if (p === '..') { if (!out.length) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.006')); out.pop(); } else if (p && p !== '.') out.push(p); } return out.join('/'); }
  class ZipReader {
    constructor(buffer) {
      this.bytes = new Uint8Array(buffer); this.view = new DataView(buffer); this.entries = new Map(); this.expanded = 0;
      const b = this.bytes, v = this.view; let e = -1;
      for (let p = b.length - 22; p >= Math.max(0, b.length - 65557); p--) if (v.getUint32(p, true) === 0x06054b50 && p + 22 + v.getUint16(p + 20, true) === b.length) {e = p; break;}
      if (e < 0) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.007'));
      const count = v.getUint16(e + 10, true); let p = v.getUint32(e + 16, true);
      if (count === 65535 || p === 0xffffffff || v.getUint16(e + 4, true) !== 0) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.008'));
      for (let i = 0; i < count; i++) {
        if (p + 46 > b.length || v.getUint32(p, true) !== 0x02014b50) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.009'));
        const flag = v.getUint16(p + 8, true), method = v.getUint16(p + 10, true), crc = v.getUint32(p + 16, true), size = v.getUint32(p + 20, true), rawSize = v.getUint32(p + 24, true), n = v.getUint16(p + 28, true), x = v.getUint16(p + 30, true), cm = v.getUint16(p + 32, true), offset = v.getUint32(p + 42, true);
        const name = new TextDecoder('utf-8', {fatal:true}).decode(b.subarray(p + 46, p + 46 + n));
        if (this.entries.has(name)) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.010'));
        this.entries.set(name, {flag, method, crc, size, rawSize, offset}); p += 46 + n + x + cm;
      }
    }
    async read(name, optional = false) {
      const e = this.entries.get(name); if (!e) { if (optional) return null; throw new Error(globalThis.ExcelDiffI18n.t('file-codec.011',{p0:name})); }
      if (e.flag & 1) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.012'));
      if (e.rawSize > LIMITS.part || this.expanded + e.rawSize > LIMITS.expanded) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.013'));
      const v = this.view, p = e.offset;
      if (p + 30 > this.bytes.length || v.getUint32(p, true) !== 0x04034b50) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.014'));
      const start = p + 30 + v.getUint16(p + 26, true) + v.getUint16(p + 28, true);
      if (start + e.size > this.bytes.length) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.015'));
      const input = this.bytes.subarray(start, start + e.size); let data;
      if (e.method === 0) data = input;
      else if (e.method === 8) {
        let ds; try { ds = new DecompressionStream('deflate-raw'); } catch { throw new Error(globalThis.ExcelDiffI18n.t('file-codec.016')); }
        const reader = new Blob([input]).stream().pipeThrough(ds).getReader(), chunks = []; let size = 0;
        while (true) { const {value, done} = await reader.read(); if (done) break; size += value.length; if (size > e.rawSize || size > LIMITS.part) { await reader.cancel(); throw new Error(globalThis.ExcelDiffI18n.t('file-codec.017')); } chunks.push(value); }
        data = new Uint8Array(size); let pos = 0; for (const c of chunks) { data.set(c, pos); pos += c.length; }
      } else throw new Error(globalThis.ExcelDiffI18n.t('file-codec.018',{p0:e.method}));
      if (data.length !== e.rawSize || crc32(data) !== e.crc) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.019',{p0:name}));
      this.expanded += data.length; return data;
    }
    async xml(path, optional = false) { const bytes = await this.read(path, optional); return bytes ? parseXml(bytes, path) : null; }
  }
  // Canonicalize definitions, not file-local style IDs. Attribute order is irrelevant.
  function canonical(n) { if (!n) return ''; const attrs = [...n.attributes].filter(x => x.prefix !== 'xmlns' && x.name !== 'xmlns').map(x => [x.localName, x.value]).sort((x,y) => x[0].localeCompare(y[0])); const kids = children(n).map(canonical); return JSON.stringify([n.localName, attrs, kids, kids.length ? '' : text(n)]); }
  const BUILTIN = {0:'General',1:'0',2:'0.00',3:'#,##0',4:'#,##0.00',9:'0%',10:'0.00%',11:'0.00E+00',12:'# ?/?',13:'# ??/??',14:'mm-dd-yy',15:'d-mmm-yy',16:'d-mmm',17:'mmm-yy',18:'h:mm AM/PM',19:'h:mm:ss AM/PM',20:'h:mm',21:'h:mm:ss',22:'m/d/yy h:mm',49:'@'};
  function readStyles(doc) {
    if (!doc) return [{font:'',fill:'',border:'',alignment:'',numFmt:'General',protection:'',flags:''}];
    const root = doc.documentElement;
    const fonts = children(child(root, 'fonts')), fills = children(child(root, 'fills')), borders = children(child(root, 'borders'));
    const base = children(child(root, 'cellStyleXfs')), xfs = children(child(root, 'cellXfs'));
    const formats = new Map(children(child(root, 'numFmts')).map(n => [+a(n,'numFmtId'), a(n,'formatCode')]));
    if (!xfs.length) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.020'));
    return xfs.map(n => {
      const parent = base[+a(n,'xfId','0')];
      const part = (id, array, apply) => {
        const src = a(n, apply) === '0' || a(n, apply) === 'false' ? (parent || n) : n;
        const i = +(a(src,id) || a(parent,id,'0')); const node = array[i]; if (!node && i !== 0) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.021',{p0:id,p1:i})); return canonical(node);
      };
      const inherit = (name, apply) => canonical((a(n,apply) === '0' ? null : child(n,name)) || child(parent,name));
      const numSrc = a(n,'applyNumberFormat') === '0' ? (parent || n) : n;
      const numId = +(a(numSrc,'numFmtId') || a(parent,'numFmtId','0'));
      return {font:part('fontId',fonts,'applyFont'), fill:part('fillId',fills,'applyFill'), border:part('borderId',borders,'applyBorder'), alignment:inherit('alignment','applyAlignment'), numFmt:formats.get(numId) ?? BUILTIN[numId] ?? globalThis.ExcelDiffI18n.t('file-codec.022',{p0:numId}), protection:inherit('protection','applyProtection'), flags:JSON.stringify({quotePrefix:a(n,'quotePrefix','0'),pivotButton:a(n,'pivotButton','0')})};
    });
  }
  function richString(n) { return unescapeExcel(children(n).map(x => x.localName === 't' ? text(x) : x.localName === 'r' ? text(child(x,'t')) : '').join('')); }
  function makeSheet(name) { return {name, rows:new Map(), maxRow:0, maxCol:0, defaultStyle:{}, rowStyles:new Map(), colStyles:[], metadata:{}, warnings:[]}; }
  function putCell(sheet, r, c, cell) { if (!sheet.rows.has(r)) sheet.rows.set(r, new Map()); const row = sheet.rows.get(r); if (row.has(c)) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.023',{p0:sheet.name,p1:colName(c),p2:r})); row.set(c, {...cell, addr:colName(c)+r}); sheet.maxRow = Math.max(sheet.maxRow,r); sheet.maxCol = Math.max(sheet.maxCol,c+1); }
  function shiftFormula(formula, dr, dc) {
    // Quoted strings and quoted sheet names must not be rewritten.
    return formula.split(/("(?:[^"\n]|"")*"|'(?:[^']|'')*')/g).map((part,i) => {
      if (i % 2) return part;
      part = part.replace(/(^|[^A-Za-z0-9_.])([$]?)([A-Za-z]{1,3})([$]?)([1-9]\d*)(?![A-Za-z0-9_!(])/g, (all, pre, ac, col, ar, row) => {
        let c = colIndex(col), r = +row; if (c > 16383 || r > 1048576) return all;
        c += ac ? 0 : dc; r += ar ? 0 : dr; return pre + (c < 0 || c > 16383 || r < 1 || r > 1048576 ? '#REF!' : ac+colName(c)+ar+r);
      });
      part = part.replace(/(^|[^A-Za-z0-9_.])([$]?)([A-Za-z]{1,3}):([$]?)([A-Za-z]{1,3})(?![A-Za-z0-9_])/g, (all,pre,a1,c1,a2,c2) => { const x=colIndex(c1)+(a1?0:dc),y=colIndex(c2)+(a2?0:dc); return x<0||y<0||x>16383||y>16383?pre+'#REF!':pre+a1+colName(x)+':'+a2+colName(y); });
      part = part.replace(/(^|[^A-Za-z0-9_.])([$]?)([1-9]\d*):([$]?)([1-9]\d*)(?![A-Za-z0-9_])/g, (all,pre,a1,r1,a2,r2) => { const x=+r1+(a1?0:dr),y=+r2+(a2?0:dr); return x<1||y<1||x>1048576||y>1048576?pre+'#REF!':pre+a1+x+':'+a2+y; });
      return part;
    }).join('');
  }
  async function readXlsx(buffer, name) {
    const zip = new ZipReader(buffer), rels = await zip.xml('_rels/.rels');
    const office = nodes(rels,'Relationship').find(n => a(n,'Type').endsWith('/officeDocument'));
    if (!office || a(office,'TargetMode') === 'External') throw new Error(globalThis.ExcelDiffI18n.t('file-codec.024'));
    const path = resolvePath('',a(office,'Target')), workbook = await zip.xml(path);
    if (workbook.documentElement.localName !== 'workbook') throw new Error(globalThis.ExcelDiffI18n.t('file-codec.025'));
    const rp = path.slice(0,path.lastIndexOf('/')+1)+'_rels/'+path.slice(path.lastIndexOf('/')+1)+'.rels';
    const relations = nodes(await zip.xml(rp),'Relationship');
    const relPath = suffix => {const n=relations.find(n=>a(n,'Type').endsWith('/'+suffix)); return n && a(n,'TargetMode') !== 'External' ? resolvePath(path,a(n,'Target')) : null;};
    const sp=relPath('styles'), styles=readStyles(sp ? await zip.xml(sp) : null);
    const ss=relPath('sharedStrings'), sdoc=ss ? await zip.xml(ss) : null;
    const shared=sdoc ? children(sdoc.documentElement,'si').map(n=>({v:richString(n),rich:children(n,'r').length ? canonical(n) : ''})) : [];
    const tp=relPath('theme'), theme=tp ? canonical((await zip.xml(tp)).documentElement) : '';
    const sheets=[], warnings=[]; let totalCells=0;
    for (const info of nodes(workbook,'sheet')) {
      const rid = [...info.attributes].find(x=>x.localName==='id')?.value;
      const rel=relations.find(n=>a(n,'Id')===rid);
      if (!rel || a(rel,'TargetMode')==='External' || !a(rel,'Type').endsWith('/worksheet')) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.026',{p0:a(info,'name')}));
      const sheet=makeSheet(a(info,'name')), doc=await zip.xml(resolvePath(path,a(rel,'Target'))), root=doc.documentElement;
      if (sheets.some(s=>s.name===sheet.name)) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.027'));
      sheet.defaultStyle=styles[0];
      const cols=nodes(doc,'col');
      for (const col of cols) if (a(col,'style')) { const style=styles[+a(col,'style')]; if (!style) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.028')); sheet.colStyles.push({min:+a(col,'min')-1,max:+a(col,'max')-1,style}); }
      const rowDims=[], sharedFormula=new Map(), pending=[]; let previousRow=0, noCache=0;
      for (const row of children(child(root,'sheetData'),'row')) {
        const r=+(a(row,'r') || previousRow+1); previousRow=r; let cPrev=-1;
        if(a(row,'s') && a(row,'customFormat','1') !== '0') { const style=styles[+a(row,'s')]; if(!style) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.029')); sheet.rowStyles.set(r,style); }
        const dims={}; for (const k of ['ht','hidden','outlineLevel','collapsed']) if(a(row,k) && a(row,k)!=='0') dims[k]=a(row,k);
        if(Object.keys(dims).length) rowDims.push([r,dims]);
        for (const n of children(row,'c')) {
          const pos=a(n,'r') ? address(a(n,'r')) : {r,c:cPrev+1}; cPrev=pos.c;
          if(pos.r!==r) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.030'));
          if(++totalCells>LIMITS.cells) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.031'));
          const vn=child(n,'v'), fn=child(n,'f'), type=a(n,'t','n'), raw=text(vn); let v=null,t='blank',rich='';
          if(type==='s') { const entry=shared[Number(raw)]; if(!vn || !/^\d+$/.test(raw) || !entry) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.032',{p0:sheet.name,p1:a(n,'r')})); v=entry.v;t='s';rich=entry.rich; }
          else if(type==='inlineStr') { const si=child(n,'is'); v=richString(si);t='s';rich=children(si,'r').length?canonical(si):''; }
          else if(type==='str') { v=unescapeExcel(raw);t='s'; }
          else if(vn && raw!=='') {
            if(type==='b') { if(raw!=='0'&&raw!=='1') throw new Error(globalThis.ExcelDiffI18n.t('file-codec.033')); v=raw==='1';t='b'; }
            else if(type==='e') {v=raw;t='e';}
            else if(type==='d') { const d=new Date(raw); if(!Number.isFinite(d.getTime())) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.034'));v=d.toISOString();t='d'; }
            else if(type==='n') {v=Number(raw);t='n';if(!Number.isFinite(v)) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.035'));}
            else throw new Error(globalThis.ExcelDiffI18n.t('file-codec.036',{p0:type}));
          }
          const inherited=sheet.rowStyles.get(r)||[...sheet.colStyles].reverse().find(x=>pos.c>=x.min&&pos.c<=x.max)?.style||styles[0];
          const style=a(n,'s') ? styles[+a(n,'s')] : inherited;
          if(!style) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.037'));
          const cell={v,t,f:fn?text(fn):null,style,rich,formulaMeta:'',cached:!!vn};
          if(fn) {
            const ft=a(fn,'t');
            if(ft==='shared') { const id=a(fn,'si'); if(text(fn)) sharedFormula.set(id,{r,c:pos.c,f:text(fn)}); else pending.push({r,c:pos.c,id,cell}); }
            else if(ft==='array') cell.formulaMeta=JSON.stringify({type:'array',ref:a(fn,'ref')});
            else if(ft==='dataTable') cell.formulaMeta=canonical(fn);
            else if(ft && ft!=='normal') throw new Error(globalThis.ExcelDiffI18n.t('file-codec.038',{p0:ft}));
            if(!vn) noCache++;
          }
          putCell(sheet,r,pos.c,cell);
        }
      }
      for(const item of pending) { const base=sharedFormula.get(item.id); if(!base) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.039')); sheet.rows.get(item.r).get(item.c).f=shiftFormula(base.f,item.r-base.r,item.c-base.c); }
      const colDims=cols.map(n=>{const out={};for(const k of ['min','max','width','hidden','outlineLevel','collapsed']) if(a(n,k))out[k]=a(n,k);return out;}).filter(x=>x.width||x.hidden==='1'||x.outlineLevel||x.collapsed==='1');
      sheet.metadata={visibility:a(info,'state','visible'), merges:nodes(doc,'mergeCell').map(n=>a(n,'ref')).sort(), rowDimensions:rowDims, columnDimensions:colDims,
        rowStyles:[...sheet.rowStyles], columnStyles:sheet.colStyles, defaults:canonical(child(root,'sheetFormatPr'))};
      if(nodes(doc,'conditionalFormatting').length) sheet.warnings.push(globalThis.ExcelDiffI18n.t('file-codec.040',{p0:sheet.name}));
      if(nodes(doc,'drawing').length||nodes(doc,'legacyDrawing').length||nodes(doc,'tablePart').length) sheet.warnings.push(globalThis.ExcelDiffI18n.t('file-codec.041',{p0:sheet.name}));
      if(nodes(doc,'hyperlink').length) sheet.warnings.push(globalThis.ExcelDiffI18n.t('file-codec.042',{p0:sheet.name}));
      if(noCache) sheet.warnings.push(globalThis.ExcelDiffI18n.t('file-codec.043',{p0:sheet.name,p1:noCache}));
      sheets.push(sheet);
    }
    if(!sheets.length) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.044'));
    return {name,kind:'xlsx',sheets,warnings,metadata:{date1904:a(child(workbook.documentElement,'workbookPr'),'date1904','0'),theme}};
  }
  function parseCsv(s, delimiter=',') {
    const rows=[]; let row=[],cell='',quoted=false,closed=false,atStart=true;
    if(s.startsWith('sep=') && /^sep=.\r?\n/.test(s)) {delimiter=s[4];s=s.replace(/^sep=.\r?\n/,'');}
    // Tab-delimited exports are accepted only when the first logical line has no comma.
    if(delimiter===',' && s.split(/\r?\n/,1)[0].includes('\t')&&!s.split(/\r?\n/,1)[0].includes(',')) delimiter='\t';
    for(let i=0;i<s.length;i++) {const ch=s[i];
      if(quoted) { if(ch==='"') {if(s[i+1]==='"'){cell+='"';i++;}else{quoted=false;closed=true;}} else cell+=ch; continue; }
      if(ch==='"') {if(!atStart) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.045'));quoted=true;atStart=false;continue;}
      if(ch===delimiter) {row.push(cell);cell='';closed=false;atStart=true;continue;}
      if(ch==='\r'||ch==='\n') {if(ch==='\r'&&s[i+1]==='\n')i++;row.push(cell);rows.push(row);row=[];cell='';closed=false;atStart=true;continue;}
      if(closed) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.046'));cell+=ch;atStart=false;
    }
    if(quoted) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.047'));
    if(row.length||cell!==''||!atStart) {row.push(cell);rows.push(row);}
    return rows;
  }
  function readCsv(buffer,name) {
    const bytes=new Uint8Array(buffer); let s,encoding='UTF-8';
    if(bytes[0]===255&&bytes[1]===254){s=new TextDecoder('utf-16le',{fatal:true}).decode(bytes);encoding='UTF-16LE';}
    else if(bytes[0]===254&&bytes[1]===255){s=new TextDecoder('utf-16be',{fatal:true}).decode(bytes);encoding='UTF-16BE';}
    else {try{s=new TextDecoder('utf-8',{fatal:true}).decode(bytes);}catch{s=new TextDecoder('shift_jis',{fatal:true}).decode(bytes);encoding=globalThis.ExcelDiffI18n.t('file-codec.048');}}
    if(s.includes('\u0000'))throw new Error(globalThis.ExcelDiffI18n.t('file-codec.049'));
    const sheet=makeSheet('Sheet1'),rows=parseCsv(s.replace(/^\uFEFF/,''));let count=0;
    rows.forEach((row,i)=>row.forEach((v,c)=>{if(++count>LIMITS.cells)throw new Error(globalThis.ExcelDiffI18n.t('file-codec.050'));putCell(sheet,i+1,c,{v,t:'s',f:null,style:{},rich:'',formulaMeta:'',cached:true});}));
    return {name,kind:'csv',encoding,sheets:[sheet],warnings:encoding.includes(globalThis.ExcelDiffI18n.t('file-codec.051'))?[globalThis.ExcelDiffI18n.t('file-codec.052',{p0:name})]:[],metadata:{}};
  }
  let legacyPromise;
  async function readLegacy(buffer,name) {
    // Optional compatibility path. XLSX and CSV never need this external dependency.
    if(!globalThis.XLSX) {
      legacyPromise ||= new Promise((resolve,reject)=>{const s=document.createElement('script'); const fail=()=>{s.remove();reject(new Error(globalThis.ExcelDiffI18n.t('file-codec.053')));}; const timer=setTimeout(fail,15000);s.src='https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';s.onload=()=>{clearTimeout(timer);resolve();};s.onerror=()=>{clearTimeout(timer);fail();};document.head.append(s);}).catch(e=>{legacyPromise=null;throw e;});
      await legacyPromise;
    }
    const wb=globalThis.XLSX.read(buffer,{type:'array',cellFormula:true,cellNF:true,cellStyles:true,sheetStubs:true,raw:true,WTF:true});
    const sheets=wb.SheetNames.map(name=>{const sheet=makeSheet(name),ws=wb.Sheets[name];for(const [addr,c] of Object.entries(ws)){if(addr.startsWith('!'))continue;const pos=address(addr);putCell(sheet,pos.r,pos.c,{v:c.v??null,t:c.t==='z'?'blank':c.t||'blank',f:c.f??null,style:{numFmt:c.z||'General'},rich:'',formulaMeta:c.F||'',cached:c.v!==undefined});}return sheet;});
    if(!sheets.length)throw new Error(globalThis.ExcelDiffI18n.t('file-codec.054'));
    return {name,kind:'xls',sheets,warnings:[globalThis.ExcelDiffI18n.t('file-codec.055',{p0:name})],metadata:{}};
  }
  async function read(file) {
    if(file.size>LIMITS.file)throw new Error(globalThis.ExcelDiffI18n.t('file-codec.056',{p0:file.name}));
    if(file.size===0)throw new Error(globalThis.ExcelDiffI18n.t('file-codec.057',{p0:file.name}));
    const buffer=await file.arrayBuffer(),name=file.name;let book;
    if(/\.xlsx$/i.test(name))book=await readXlsx(buffer,name);
    else if(/\.csv$/i.test(name))book=readCsv(buffer,name);
    else if(/\.xls$/i.test(name))book=await readLegacy(buffer,name);
    else throw new Error(globalThis.ExcelDiffI18n.t('file-codec.058'));
    book.bytes=new Uint8Array(buffer); return book;
  }
  // Write uncompressed ZIP so the report is independent of third-party libraries.
  function zipStore(files) {
    const local=[],central=[];let offset=0,centralSize=0;
    for(const [name,content]of Object.entries(files)) {
      const n=enc.encode(name),data=typeof content==='string'?enc.encode(content):content,crc=crc32(data);
      const h=new Uint8Array(30+n.length),v=new DataView(h.buffer);v.setUint32(0,0x04034b50,true);v.setUint16(4,20,true);v.setUint16(6,0x800,true);v.setUint16(12,0x21,true);v.setUint32(14,crc,true);v.setUint32(18,data.length,true);v.setUint32(22,data.length,true);v.setUint16(26,n.length,true);h.set(n,30);local.push(h,data);
      const c=new Uint8Array(46+n.length),d=new DataView(c.buffer);d.setUint32(0,0x02014b50,true);d.setUint16(4,20,true);d.setUint16(6,20,true);d.setUint16(8,0x800,true);d.setUint16(14,0x21,true);d.setUint32(16,crc,true);d.setUint32(20,data.length,true);d.setUint32(24,data.length,true);d.setUint16(28,n.length,true);d.setUint32(42,offset,true);c.set(n,46);central.push(c);offset+=h.length+data.length;centralSize+=c.length;
    }
    const end=new Uint8Array(22),v=new DataView(end.buffer);v.setUint32(0,0x06054b50,true);v.setUint16(8,central.length,true);v.setUint16(10,central.length,true);v.setUint32(12,centralSize,true);v.setUint32(16,offset,true);return new Blob([...local,...central,end],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  }
  function writeReport(sheets) {
    const relNS='http://schemas.openxmlformats.org/package/2006/relationships', files={};
    files['[Content_Types].xml']=`<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets.map((_,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`;
    files['_rels/.rels']=`<Relationships xmlns="${relNS}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
    files['xl/workbook.xml']=`<workbook xmlns="${NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets.map((s,i)=>`<sheet name="${xmlEscape(s.name)}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join('')}</sheets></workbook>`;
    files['xl/_rels/workbook.xml.rels']=`<Relationships xmlns="${relNS}">${sheets.map((_,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join('')}<Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
    files['xl/styles.xml']=`<styleSheet xmlns="${NS}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF177742"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"><alignment vertical="center" wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
    sheets.forEach((sheet,i)=>{ for(const row of sheet.rows) for(const v of row) if(String(v??'').length>32767) throw new Error(globalThis.ExcelDiffI18n.t('file-codec.059')); const width=Math.max(1,...sheet.rows.map(r=>r.length));files[`xl/worksheets/sheet${i+1}.xml`]=`<worksheet xmlns="${NS}"><dimension ref="A1:${colName(width-1)}${Math.max(1,sheet.rows.length)}"/><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>${Array.from({length:width},(_,c)=>`<col min="${c+1}" max="${c+1}" width="${c<3?22:36}" customWidth="1"/>`).join('')}</cols><sheetData>${sheet.rows.map((row,r)=>`<row r="${r+1}"${r===0?' ht="30" customHeight="1"':''}>${row.map((v,c)=>{const ref=colName(c)+(r+1);return typeof v==='number'&&Number.isFinite(v)?`<c r="${ref}" s="${r===0?1:0}"><v>${v}</v></c>`:`<c r="${ref}" s="${r===0?1:0}" t="inlineStr"><is><t xml:space="preserve">${xmlEscape(v)}</t></is></c>`;}).join('')}</row>`).join('')}</sheetData></worksheet>`;});
    return zipStore(files);
  }
  globalThis.ExcelDiffCodec={read,readXlsx,readCsv,parseCsv,writeReport,colName,colIndex,address,shiftFormula,crc32,canonical,LIMITS};
})();
