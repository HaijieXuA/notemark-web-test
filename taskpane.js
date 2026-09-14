'use strict';
const C=NoteMarkCore,$=id=>document.getElementById(id),controls=['convert','restore','inspect'];
let ready=false,busy=false,store;
const hostPending=new Map();
function hostHighlight(text,ranges){const id=crypto.randomUUID();return new Promise((resolve,reject)=>{const timer=setTimeout(()=>{hostPending.delete(id);reject(Error('高亮步骤未响应，原文备份已保留'));},5000);hostPending.set(id,{resolve,reject,timer});parent.postMessage({channel:'notemark.v2',hostCommand:'highlight',id,text,ranges},'https://onenote.officeapps.live.com');});}
function highlights(html){const doc=new DOMParser().parseFromString(html,'text/html'),ranges=[];let position=0;function walk(n,active=false){if(n.nodeType===3){if(active&&n.textContent)ranges.push({start:position,length:[...n.textContent].length,text:n.textContent});position+=[...n.textContent].length;return;}if(n.nodeType!==1)return;for(const child of n.childNodes)walk(child,active||!!n.style.backgroundColor);}walk(doc.body);return ranges;}
try{store=new C.SourceStore(localStorage);}catch(e){$('status').textContent='浏览器存储不可用：'+e.message;}
function report(v){$('log').textContent=(JSON.stringify(v,null,2)+'\n'+$('log').textContent).slice(0,12000);}
function status(s){$('status').textContent=s;}
async function snapshot(hint,includeHtml=true){return OneNote.run(async ctx=>{
 const app=ctx.application,book=app.getActiveNotebook(),page=app.getActivePage();
 let p=app.getActiveParagraph(),textLoaded=false;
 book.load('id');page.load('id,title');p.load('id,type');await ctx.sync();
 if(!$('allpages').checked&&page.title.trim()!=='NoteMark Web Test')throw Error('默认仅在 NoteMark Web Test 页操作。');
 if(p.isNullObject){
  const selected=C.stripEnd(await (hint??read()));
  const ps=app.getActiveOutline().paragraphs;
  ps.load('items/id,items/type,items/richText/text');await ctx.sync();
  const matches=ps.items.filter(x=>x.type==='RichText'&&C.stripEnd(x.richText.text)===selected);
  if(matches.length!==1)throw Error('不能唯一定位当前段落，请将光标放在一行文字中重试。');
  p=matches[0];textLoaded=true;
 }
 if(p.type!=='RichText')throw Error('当前段落不是文字。');
 let html;
 if(includeHtml){p.richText.load('text');html=p.richText.getHtml();await ctx.sync();}
 else if(!textLoaded){p.richText.load('text');await ctx.sync();}
 return {notebook:book.id,page:page.id,paragraph:p.id,text:C.stripEnd(p.richText.text),html:html?.value};
});}
const read=()=>new Promise((resolve,reject)=>Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,r=>r.status===Office.AsyncResultStatus.Succeeded?resolve(r.value):reject(r.error)));
const write=html=>new Promise((resolve,reject)=>Office.context.document.setSelectedDataAsync(html,{coercionType:Office.CoercionType.Html},r=>r.status===Office.AsyncResultStatus.Succeeded?resolve():reject(r.error)));
async function run(mode,fromBridge=false){
 if(!ready||busy)throw Error('正在处理，请稍后重试。');busy=true;controls.forEach(id=>$(id).disabled=true);
 const start=performance.now();let wrote=false;
 try{
  const selectedRequest=read();
  const [before,selected]=await Promise.all([snapshot(selectedRequest,mode!=='convert'),selectedRequest]);
  if(mode==='inspect'){report({context:before,selected});status('已读取当前段落。');return {};}
  const raw=C.stripEnd(selected);
  if(!raw||raw!==before.text)throw Error('请完整选中当前段落，避免覆盖部分文字。');
  let html,source,ranges=[];
  if(mode==='restore'){source=store.get(before,before.text,before.html);if(source===null)throw Error('此段落没有匹配的原文记录。');html='<p>'+C.esc(source)+'</p>';}
  else{source=raw;const parsed=C.parse(source);if(!parsed.changed){status('当前行没有可转换的标记。');return {unchanged:true,text:raw};}html=parsed.html;ranges=highlights(html);if(ranges.length&&!fromBridge)throw Error('高亮需要配套网页扩展；本次未修改正文。');}
  if(mode!=='restore')store.prepare(before,source);
  if(await read()!==selected)throw Error('选区已变化，已取消。');
  // Do not retry writes: a delayed successful write must never be duplicated.
  const tw=performance.now();wrote=true;await write(html);const writeMs=performance.now()-tw;
  const expected=mode==='restore'?source:C.parse(source).text;
  if(ranges.length)await hostHighlight(expected,ranges);
  const after=await snapshot(expected,mode!=='restore');
  if(after.text!==expected)throw Error('接口已返回，但正文未通过核对。请检查正文；未自动重试。');
  if(mode!=='restore')store.put(after,source,after.text,after.html);
  const result={mode,text:expected,writeMs,totalMs:performance.now()-start,paragraph:after.paragraph};report(result);status('已完成，耗时 '+Math.round(result.totalMs)+' ms。');return result;
 }catch(e){e.wrote=wrote;status(e.message||String(e));report({error:e.message||String(e),code:e.code});throw e;}
 finally{busy=false;controls.forEach(id=>$(id).disabled=!ready);}
}
for(const id of controls)$(id).onclick=()=>run(id).catch(()=>{});
$('backup').onclick=()=>{try{const r=JSON.parse(localStorage.getItem('notemark.pending.v1')||'null');$('backuptext').hidden=false;$('backuptext').value=r?.source||'没有备份记录';}catch(e){status('无法读取原文备份：'+e.message);}};
$('clear').onclick=()=>{if(!store)return;store.clear();$('backuptext').value='';status('本地原文记录已清空。');};
// The paired extension runs only in the OneNote editor. Do not accept other origins or nested senders.
window.addEventListener('message',async e=>{
 if(e.origin!=='https://onenote.officeapps.live.com'||e.source!==parent)return;
 if(e.data?.channel==='notemark.host'){const p=hostPending.get(e.data.id);if(!p)return;clearTimeout(p.timer);hostPending.delete(e.data.id);e.data.ok?p.resolve():p.reject(Error(e.data.error));return;}
 if(e.data?.channel!=='notemark.v2'||!$('bridge').checked)return;
 const {id,mode}=e.data;if(typeof id!=='string'||!['convert','restore','ping'].includes(mode))return;
 try{const result=mode==='ping'?{ready:ready&&!busy,autoEnter:$('autoenter').checked}:await run(mode,true);e.source.postMessage({channel:'notemark.v2',id,ok:true,result},e.origin);}
 catch(err){e.source.postMessage({channel:'notemark.v2',id,ok:false,error:err.message||String(err),wrote:err.wrote},e.origin);}
});
if(typeof Office==='undefined')status('Office.js 未加载。');
else Office.onReady(info=>{ready=info.host==='OneNote'&&!!store;controls.forEach(id=>$(id).disabled=!ready);status(ready?'已连接 OneNote。':'请从 OneNote 加载项菜单打开。');report(info);});
