(function(root){
'use strict';
const esc=s=>s.replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
function parse(source){
 if(typeof source!=='string'||/[\r\n]/.test(source)||source.length>10000)throw Error('请选择单个段落（最多 10000 字符）。');
 let heading=0, body=source, changed=false;
 const h=/^(#{1,6})\s+(.+)$/.exec(source); if(h){heading=h[1].length;body=h[2];changed=true;}
 function inline(s,depth=0){
  if(depth>16)return {html:esc(s),text:s};
  let html='',text='';
  for(let i=0;i<s.length;){
   if(s[i]==='\\'&&i+1<s.length&&/[\\*~=]/.test(s[i+1])){html+=esc(s[i+1]);text+=s[i+1];i+=2;changed=true;continue;}
   const mark=['***','**','~~','==','*'].find(m=>s.startsWith(m,i));
   let end=-1;
   if(mark){for(let j=i+mark.length;j<s.length;j++){if(s[j]==='\\'){j++;continue;}if(s.startsWith(mark,j)){end=j;break;}}}
   if(mark&&end>i+mark.length&&!/^\s|\s$/.test(s.slice(i+mark.length,end))){
    const inner=inline(s.slice(i+mark.length,end),depth+1);
    const tags={'***':['<strong><em>','</em></strong>'],'**':['<strong>','</strong>'],'*':['<em>','</em>'],'~~':['<s>','</s>'],'==':['<span style="background-color:#ffff00">','</span>']}[mark];
    html+=tags[0]+inner.html+tags[1];text+=inner.text;i=end+mark.length;changed=true;
   }else{html+=esc(s[i]);text+=s[i++];}
  }
  return {html,text};
 }
 const r=inline(body),tag=heading?'h'+heading:'p';
 return {source,text:r.text,html:'<'+tag+'>'+r.html+'</'+tag+'>',heading,changed};
}
// Compare effective character formatting, not OneNote's transient span layout.
function fingerprint(html){
 if(typeof DOMParser==='undefined')return html;
 const doc=new DOMParser().parseFromString(html,'text/html'),out=[];
 function walk(node,style={}){
  if(node.nodeType===3){for(const ch of node.textContent)out.push([ch,style]);return;}
  if(node.nodeType!==1)return;
  const next={...style},tag=node.tagName.toLowerCase(),css=node.style;
  if(['b','strong'].includes(tag))next.bold=true;
  if(['i','em'].includes(tag))next.italic=true;
  if(['s','del','strike'].includes(tag))next.strike=true;
  if(css.fontWeight)next.bold=css.fontWeight==='bold'||Number(css.fontWeight)>=600;
  if(css.fontStyle)next.italic=css.fontStyle==='italic';
  if(css.textDecoration.includes('line-through'))next.strike=true;
  if(css.backgroundColor)next.background=css.backgroundColor;
  if(css.color)next.color=css.color;
  if(css.fontSize)next.size=css.fontSize;
  if(css.fontFamily)next.font=css.fontFamily;
  if(/^h[1-6]$/.test(tag))next.heading=tag;
  // Stable property order, independent of the source span nesting.
  const normalized=Object.fromEntries(Object.entries(next).filter(([,v])=>v).sort(([a],[b])=>a.localeCompare(b)));
  for(const child of node.childNodes)walk(child,normalized);
 }
 walk(doc.body);return JSON.stringify(out);
}
const stripEnd=s=>s.replace(/\r?\n$|\r$/,'');
const key=c=>JSON.stringify([c.notebook,c.page,c.paragraph]);
class SourceStore{
 constructor(storage){this.storage=storage;this.name='notemark.sources.v1';this.records={};try{const value=JSON.parse(storage.getItem(this.name)||'{}');if(value&&typeof value==='object'&&!Array.isArray(value))this.records=Object.fromEntries(Object.entries(value).filter(([,v])=>v&&typeof v.source==='string'&&typeof v.at==='number'));}catch{}}
 prepare(context,source){this.storage.setItem('notemark.pending.v1',JSON.stringify({context,source,at:Date.now()}));}
 put(context,source,text,html){const records={...this.records,[key(context)]:{source,text,html,at:Date.now()}};const entries=Object.entries(records).sort((a,b)=>b[1].at-a[1].at).slice(0,500);this.storage.setItem(this.name,JSON.stringify(Object.fromEntries(entries)));this.records=Object.fromEntries(entries);}
 get(context,text,html){const r=this.records[key(context)];if(!r)return null;if(r.text!==text||fingerprint(r.html)!==fingerprint(html))throw Error('此段落渲染后已被编辑，不能用旧原文覆盖。');return r.source;}
 clear(){this.storage.removeItem(this.name);this.storage.removeItem('notemark.pending.v1');this.records={};}
}
const api={parse,esc,stripEnd,SourceStore,fingerprint};if(typeof module!=='undefined')module.exports=api;else root.NoteMarkCore=api;
})(globalThis);
