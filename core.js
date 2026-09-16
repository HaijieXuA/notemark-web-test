(function(root){
'use strict';
const esc=s=>s.replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
// Strict, versioned style configuration. Unknown fields never silently disappear.
function normalizeConfig(value){
 const object=(v,label)=>{if(!v||typeof v!=='object'||Array.isArray(v))throw Error(label+' 必须是对象');};
 object(value,'配置');
 for(const k of Object.keys(value))if(!['version','name','headings'].includes(k))throw Error('不支持的配置字段：'+k);
 if(value.version!==1)throw Error('配置 version 必须为 1；其他格式需要先适配');
 if(value.name!==undefined&&(typeof value.name!=='string'||value.name.length>80))throw Error('name 必须是最多 80 字符的文本');
 object(value.headings,'headings');
 const headings={};
 for(const [level,style] of Object.entries(value.headings)){
  if(!/^h[1-6]$/.test(level))throw Error('不支持的标题级别：'+level);
  object(style,level);const out={};
  for(const [key,v] of Object.entries(style)){
   if(!['fontFamily','fontSize','color','bold','italic'].includes(key))throw Error(level+' 不支持字段：'+key);
   if(key==='fontFamily'&&(typeof v!=='string'||!v.trim()||v.length>100||/[;:{}<>"\\]/.test(v)))throw Error(level+' 字体名称无效');
   if(key==='fontSize'&&(typeof v!=='number'||!Number.isFinite(v)||v<6||v>96))throw Error(level+' 字号必须为 6–96 的数字（磅）');
   if(key==='color'&&(typeof v!=='string'||!/^#[0-9a-f]{6}$/i.test(v)))throw Error(level+' 颜色需为 #RRGGBB');
   if(['bold','italic'].includes(key)&&typeof v!=='boolean')throw Error(level+' '+key+' 必须为布尔值');
   out[key]=v;
  }
  headings[level]=out;
 }
 return {version:1,name:value.name||'自定义标题样式',headings};
}
function parseConfig(text){
 if(typeof text!=='string'||text.length>65536)throw Error('配置文件不能超过 64 KB');
 let value;try{value=JSON.parse(text.replace(/^\uFEFF/,''));}catch{throw Error('JSON 格式不正确');}
 return normalizeConfig(value);
}
function headingStyle(config,heading){
 const style=config?.headings?.['h'+heading];if(!style)return '';
 const css=[];
 for(const [key,value] of Object.entries(style)){
  if(key==='fontFamily')css.push('font-family:'+value);
  if(key==='fontSize')css.push('font-size:'+value+'pt');
  if(key==='color')css.push('color:'+value);
  if(key==='bold')css.push('font-weight:'+(value?'bold':'normal'));
  if(key==='italic')css.push('font-style:'+(value?'italic':'normal'));
 }
 return css.join(';');
}
function parse(source,config=null){
 if(config)config=normalizeConfig(config);
 if(typeof source!=='string'||/[\r\n]/.test(source)||source.length>10000)throw Error('请选择单个段落（最多 10000 字符）。');
 let heading=0, body=source, changed=false,list=null;
 const lm=/^(?:(\d{1,9})[.)]|([-+*]))[ \t\u00a0]+(\S.*)$/.exec(source);
 if(lm){list={type:lm[1]?'ol':'ul',start:lm[1]?Number(lm[1]):1};body=lm[3];changed=true;}
 if(/^[ \t]+(?:\d+[.)]|[-+*])[ \t]+/.test(source))throw Error('暂不转换嵌套列表，请先使用一级列表。');
 const h=list?null:/^(#{1,6})\s+(.+)$/.exec(source); if(h){heading=h[1].length;body=h[2];changed=true;}
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
 const css=headingStyle(config,heading);
 return {source,text:r.text,html:'<'+tag+'>'+(css?'<span style="'+esc(css)+'">'+r.html+'</span>':r.html)+'</'+tag+'>',heading,changed,list};
}
// Built-in Haru adaptation from the user's haru-8.css. No external CSS/fonts loaded.
function renderHaru(source,nativeList=null){
 const parsed=parse(source);if(!parsed.list&&nativeList){if(parsed.heading)throw Error('请先退出列表再转换标题。');parsed.list=nativeList;parsed.changed=true;}const sizes={1:22.5,2:18.75,3:15,4:13.5,5:12};
 let html=parsed.html.replace(/<strong>/g,'<strong style="color:#5b32b4">').replace(/<em>/g,'<em style="color:#4169e1;text-decoration:underline">');
 const heading=parsed.heading,tag=heading?'h'+heading:'p';
 const font=heading?"Roboto Slab":"Glow Sans";
 const styles=['font-family:'+font,'color:'+(heading===1?'#461289':'#0c0c0c')];
 if(!heading||sizes[heading])styles.push('font-size:'+(heading?sizes[heading]:12)+'pt');
 if(heading)styles.push('font-weight:bold');
 // Explicit runs survive OneNote's single-family normalization better than a fallback list.
 if(heading)html=html.split(/(<[^>]+>|&(?:#\d+|#x[\da-f]+|[a-z]+);)/gi).map(part=>part.startsWith('<')||part.startsWith('&')?part:part.replace(/[\p{Script=Han}\u3000-\u303f\uff01-\uff60]+/gu,text=>'<span style="font-family:SimSun">'+text+'</span>')).join('');

 html=html.replace('<'+tag+'>','<'+tag+(heading===1?' style="text-align:center"':'')+'><span style="'+styles.join(';')+'">').replace('</'+tag+'>','</span></'+tag+'>');
 if(heading===1)html=html.replace('<span style="','<u><span style="').replace('</span></h1>','</span></u></h1>');
 let prefix='';
 if(heading===2||heading===3){
  const color=heading===2?'#801eff':'#4169e1',size=heading===2?15.5:12.5;
  prefix='\u202f\u00a0 \u202f\u202f\u202f ';
  const decoration='<span data-notemark-decoration="true" style="color:#ffffff;background-color:'+color+';font-family:Omgnore,sans-serif;font-size:'+size+'pt">'+prefix.slice(0,4)+'</span><span>'+prefix.slice(4)+'</span>';
  html=html.replace('<h'+heading+'>','<h'+heading+'>'+decoration);
 }
 if(parsed.list){
  const {type,start}=parsed.list;
  if(!['ol','ul'].includes(type)||!Number.isInteger(start)||start<0||start>999999998)throw Error('列表编号无效。');
  prefix=(type==='ol'?start+'.':'•')+'\u00a0\u00a0 ';
  const marker='<span data-notemark-list="'+type+'" style="font-family:Glow Sans;font-size:11pt;color:#6602e8">'+prefix+'</span>';
  html=html.replace('font-size:12pt','font-size:12.5pt;line-height:17.25pt').replace('<p>','<p style="margin-left:21.75pt;text-indent:-17.25pt;margin-top:0;margin-bottom:5.25pt">'+marker);
 }

 return {...parsed,text:prefix+parsed.text,html};
}
// Serialize only supported inline formatting; never silently discard other content.
function markdownFromRuns(runs,heading=0){
 const escapeText=s=>s.replace(/[\\*~=]/g,'\\$&');
 let body='';const merged=[];
 for(const r of runs){const marks=(r.bold?'**':'')+(r.italic?'*':'');const key=JSON.stringify([marks,!!r.strike,!!r.highlight]);
  if(merged.length&&merged[merged.length-1].key===key)merged[merged.length-1].text+=r.text;
  else merged.push({...r,key,marks});}
 for(const r of merged){const m=/^(\s*)([\s\S]*?)(\s*)$/.exec(r.text),inner=escapeText(m[2]);
  const open=(r.highlight?'==':'')+(r.strike?'~~':'')+r.marks;
  const close=r.marks+(r.strike?'~~':'')+(r.highlight?'==':'');
  body+=m[1]+(inner?open+inner+close:'')+m[3];}
 const source=(heading?'#'.repeat(heading)+' ':'')+body;
 const expected=runs.map(r=>r.text).join('');
 if(parse(source).text!==expected)throw Error('当前格式组合无法可靠生成 Markdown，正文未修改。');
 return source;
}
function fromHtml(html,nativeList=null,visualHint=null){
 const doc=new DOMParser().parseFromString(html,'text/html'),runs=[];let heading=0;
 if(doc.querySelector('table,img,a,code,pre,br,script,style'))throw Error('本段含暂不支持还原的内容，正文未修改。');
 const lists=doc.querySelectorAll('ol,ul'),items=doc.querySelectorAll('li');
 if(lists.length>1||items.length>1)throw Error('请只还原单个一级列表项。');
 let list=nativeList;
 if(lists.length){const el=lists[0],type=el.tagName.toLowerCase();list={type,start:Number(el.getAttribute('start')||1)};}
 const blocks=doc.querySelectorAll('p,h1,h2,h3,h4,h5,h6');
 if(blocks.length>1)throw Error('请只选择一个段落。');
 function walk(n,style={}){
  if(n.nodeType===3){runs.push({...style,text:n.textContent});return;}
  if(n.nodeType!==1)return;
  const tag=n.tagName.toLowerCase(),css=n.style,next={...style};
  if(/^h[1-6]$/.test(tag))heading=Number(tag[1]);
  if(['b','strong'].includes(tag))next.bold=true;
  if(['i','em'].includes(tag))next.italic=true;
  if(['s','del','strike'].includes(tag))next.strike=true;
  if(tag==='mark')next.highlight=true;
  if(css.fontWeight)next.bold=css.fontWeight==='bold'||Number(css.fontWeight)>=600;
  if(css.fontStyle)next.italic=css.fontStyle==='italic';
  if(css.textDecoration.includes('line-through'))next.strike=true;
  if(css.backgroundColor)next.background=css.backgroundColor;
  if(css.fontSize)next.fontSize=css.fontSize;
  if(css.color)next.color=css.color;
  if(css.backgroundColor&&css.backgroundColor!=='transparent'&&css.backgroundColor!=='rgba(0, 0, 0, 0)')next.highlight=true;
  for(const child of n.childNodes)walk(child,next);
 }
 walk(doc.body);
 let visibleList=null;
 const markerText=runs.map(r=>r.text).join('').match(/^(\d{1,9}\.|•)\u00a0{2}[\u00a0 ]/);
 if(markerText){
  const first=runs.find(r=>r.text.length);
  if(first&&(first.color==='#6602e8'||first.color==='rgb(102, 2, 232)'||(visualHint&&visualHint.type===(markerText[1]==='•'?'ul':'ol')&&visualHint.start===(markerText[1]==='•'?1:parseInt(markerText[1],10))))){
   visibleList={type:markerText[1]==='•'?'ul':'ol',start:markerText[1]==='•'?1:parseInt(markerText[1],10)};
   let left=markerText[0].length;while(left&&runs.length){const n=Math.min(left,runs[0].text.length);runs[0].text=runs[0].text.slice(n);left-=n;if(!runs[0].text)runs.shift();}
  }
 }
 const prefixes=['\u202f\u00a0 \u202f\u202f\u202f ','\u202f\u00a0 \u202f\u202f\u202f\u200b'];
 const all=runs.map(r=>r.text).join(''),prefix=prefixes.find(p=>all.startsWith(p));
 if((heading===2||heading===3)&&prefix){
  const expected=heading===2?'rgb(128, 30, 255)':'rgb(65, 105, 225)';
  let checked=0;for(const r of runs){if(checked>=4)break;if(r.background!==expected&&r.background!==(heading===2?'#801eff':'#4169e1')&&!(!r.background&&r.fontSize===(heading===2?'15.5pt':'12.5pt')))throw Error('无法确认标题装饰格式，请使用已保存的原文还原。');checked+=r.text.length;}
  let left=prefix.length;while(left&&runs.length){const n=Math.min(left,runs[0].text.length);runs[0].text=runs[0].text.slice(n);left-=n;if(!runs[0].text)runs.shift();}
 }
 list=visibleList||list;const markdown=markdownFromRuns(runs,heading);return list?(list.type==='ol'?list.start+'. ':'- ')+markdown:markdown;
}
// Compare effective character formatting, not OneNote's transient span layout.
function fingerprint(html,normalizeList=false){
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
  if(css.fontSize)next.fontSize=css.fontSize;
  if(css.color)next.color=css.color;
  if(css.fontSize)next.size=css.fontSize;
  if(css.fontFamily)next.font=css.fontFamily;
  if(/^h[1-6]$/.test(tag))next.heading=tag;
  // Stable property order, independent of the source span nesting.
  const normalized=Object.fromEntries(Object.entries(next).filter(([,v])=>v).sort(([a],[b])=>a.localeCompare(b)));
  for(const child of node.childNodes)walk(child,normalized);
 }
 walk(doc.body);if(normalizeList){const m=/^(?:\d{1,9}\.|•)\u00a0{2}[\u00a0 ]/.exec(out.map(x=>x[0]).join(''));if(m)out[m[0].length-1][0]=' ';}return JSON.stringify(out);
}
const normalizeListText=s=>s.replace(/^((?:\d{1,9}\.|•)\u00a0{2})[\u00a0 ]/,'$1 ');
const stripEnd=s=>s.replace(/\r?\n$|\r$/,'');
const key=c=>JSON.stringify([c.notebook,c.page,c.paragraph]);
class SourceStore{
 constructor(storage){this.storage=storage;this.name='notemark.sources.v1';this.records={};try{const value=JSON.parse(storage.getItem(this.name)||'{}');if(value&&typeof value==='object'&&!Array.isArray(value))this.records=Object.fromEntries(Object.entries(value).filter(([,v])=>v&&typeof v.source==='string'&&typeof v.at==='number'));}catch{}}
 prepare(context,source){this.storage.setItem('notemark.pending.v1',JSON.stringify({context,source,at:Date.now()}));}
 put(context,source,text,html){const records={...this.records,[key(context)]:{source,text,html,list:context.list||null,at:Date.now()}};const entries=Object.entries(records).sort((a,b)=>b[1].at-a[1].at).slice(0,500);this.storage.setItem(this.name,JSON.stringify(Object.fromEntries(entries)));this.records=Object.fromEntries(entries);}
 get(context,text,html){const r=this.records[key(context)];if(!r)return null;if(r.list&&(context.list||context.visualList)&&JSON.stringify(r.list)!==JSON.stringify(context.list||context.visualList))throw Error('列表编号已变化，请按当前格式重新生成 Markdown。');if((r.list?normalizeListText(r.text)!==normalizeListText(text):r.text!==text)||fingerprint(r.html,!!r.list)!==fingerprint(html,!!r.list))throw Error('此段落渲染后已被编辑，不能用旧原文覆盖。');return r.source;}
 clear(){this.storage.removeItem(this.name);this.storage.removeItem('notemark.pending.v1');this.records={};}
}
const api={renderHaru,normalizeConfig,parseConfig,fromHtml,markdownFromRuns,parse,esc,stripEnd,SourceStore,fingerprint};if(typeof module!=='undefined')module.exports=api;else root.NoteMarkCore=api;
})(globalThis);
