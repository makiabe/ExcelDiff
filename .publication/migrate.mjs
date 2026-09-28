// One-time extraction for the approved localization change. Removed before publication.
import fs from 'node:fs';
import {parse} from 'acorn';
const root=process.cwd()+'/';
fs.mkdirSync(root+'locales',{recursive:true});fs.mkdirSync(root+'src',{recursive:true});
fs.writeFileSync(root+'.publication/index-original.html',fs.readFileSync(root+'index.html'));
const dict={};
for(const file of ['app.js','diff-engine.js','file-codec.js']) {
 const source=fs.readFileSync(root+file,'utf8');
 const ast=parse(source,{ecmaVersion:'latest'});let counter=0;
 function translate(node) {
   const children=[];
   for(const [k,v]of Object.entries(node)) {
    if(['start','end'].includes(k))continue;
    if(Array.isArray(v)) for(const c of v) {if(c?.type)children.push(c);}
    else if(v?.type)children.push(v);
   }
   const replacements=children.map(c=>({node:c,text:translate(c)})).filter(x=>x.text!==source.slice(x.node.start,x.node.end));
   let modified=source.slice(node.start,node.end);
   for(const r of replacements.sort((a,b)=>b.node.start-a.node.start)) modified=modified.slice(0,r.node.start-node.start)+r.text+modified.slice(r.node.end-node.start);
   let value;
   if(node.type==='Literal'&&typeof node.value==='string'&&/[\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Han}]/u.test(node.value))value=node.value;
   if(node.type==='TemplateLiteral'&&node.quasis.some(q=>/[\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Han}]/u.test(q.value.cooked||'')))value=node.quasis.map((q,i)=>(q.value.cooked||'')+(i<node.expressions.length?`{p${i}}`:'')).join('');
   if(value!==undefined) {
    const key=file.replace('.js','')+'.'+String(++counter).padStart(3,'0');dict[key]=value;
    if(node.type==='Literal')return `globalThis.ExcelDiffI18n.t('${key}')`;
    return `globalThis.ExcelDiffI18n.t('${key}',{${node.expressions.map((e,i)=>`p${i}:${translateCached(e,replacements)}`).join(',')}})`;
   }
   return modified;
 }
 function translateCached(e,arr){const r=arr.find(x=>x.node===e);return r?r.text:source.slice(e.start,e.end);}
 let result=translate(ast);
 if(file==='app.js')result=result.replaceAll('.toLocaleString()','.toLocaleString(globalThis.ExcelDiffI18n.locale)');
 if(file==='diff-engine.js')result=result.replace("version:'1.1.0'","version:'1.2.0'");
 fs.writeFileSync(root+file,result);
}
if(Object.keys(dict).length!==200)throw Error('Unexpected source messages; stop rather than overwrite a changed source.');
fs.writeFileSync(root+'locales/ja.json',JSON.stringify(dict,null,2)+'\n');
