'use strict';
const C=NoteMarkCore,$=id=>document.getElementById(id),controls=['convert','restore','generate','inspect'];
let ready=false,busy=false,store,lastBridge=0;
function updateBridge(){const connected=Date.now()-lastBridge<4000;
 $('connection').textContent=connected?'扩展已连接':'扩展未连接：回车、⌘, 和高亮不可用';
 $('autoenter').disabled=!connected||!$('bridge').checked;
 // Preserve the selected default while the extension is connecting.
}
setInterval(updateBridge,1000);
function sendSettings(){parent.postMessage({channel:'notemark.v2',hostCommand:'settings',ready:ready&&!busy&&$('bridge').checked,autoEnter:$('autoenter').checked},'https://onenote.officeapps.live.com');}
$('bridge').onchange=()=>{updateBridge();sendSettings();};
$('autoenter').onchange=sendSettings;
$('bridge').checked=true;$('autoenter').checked=true;
const allPagesKey='notemark.settings.allpages.v1';
$('allpages').checked=true;
try{$('allpages').checked=localStorage.getItem(allPagesKey)!=='false';}catch{}
$('allpages').onchange=()=>{try{localStorage.setItem(allPagesKey,String($('allpages').checked));}catch{status('当前选择已生效，但浏览器无法保存此设置。');}};
const hostPending=new Map();
function hostHighlight(text,ranges){const id=crypto.randomUUID();return new Promise((resolve,reject)=>{const timer=setTimeout(()=>{hostPending.delete(id);reject(Error('高亮步骤未响应，原文备份已保留'));},5000);hostPending.set(id,{resolve,reject,timer});parent.postMessage({channel:'notemark.v2',hostCommand:'highlight',id,text,ranges},'https://onenote.officeapps.live.com');});}
function hostRemoveList(text){const id=crypto.randomUUID();return new Promise((resolve,reject)=>{const timer=setTimeout(()=>{hostPending.delete(id);reject(Error('清除原生列表编号超时，请检查当前行'));},5000);hostPending.set(id,{resolve,reject,timer});parent.postMessage({channel:'notemark.v2',hostCommand:'removeList',id,text},'https://onenote.officeapps.live.com');});}
function highlights(html){const doc=new DOMParser().parseFromString(html,'text/html'),ranges=[];let position=0;function walk(n,active=false){if(n.nodeType===3){if(active&&n.textContent)ranges.push({start:position,length:[...n.textContent].length,text:n.textContent});position+=[...n.textContent].length;return;}if(n.nodeType!==1)return;for(const child of n.childNodes)walk(child,active||(!n.getAttribute?.('data-notemark-decoration')&&!!n.style.backgroundColor));}walk(doc.body);return ranges;}
try{store=new C.SourceStore(localStorage);}catch(e){$('status').textContent='浏览器存储不可用：'+e.message;}
function report(v){$('log').textContent=(JSON.stringify(v,null,2)+'\n'+$('log').textContent).slice(0,12000);}
function status(s){$('status').textContent=s;}
// The web editor may expose extra empty layout paragraphs absent from Office.js.
// Match the complete ordered sequence; never skip text or guess between blank targets.
function locateDomParagraph(dom,items){
 const fail=()=>{throw Error('不能唯一对应当前段落，请重新点击文字后再试。');};
 if(!dom||!Number.isInteger(dom.index)||!Array.isArray(dom.texts)||dom.index<0||dom.index>=dom.texts.length||dom.texts.some(t=>typeof t!=='string'))return fail();
 const a=dom.texts.map(C.stripEnd);
 const same=(x,y)=>x===y||x+'\u00a0'===y||x===y+'\u00a0';
 const target=a[dom.index];
 // A unique complete target in the active outline does not depend on unrelated
 // nested lists, images or stale display rows elsewhere in the outline.
 if(target.trim()&&a.filter(t=>same(t,target)).length===1){
  const candidates=items.map((p,i)=>p.type==='RichText'&&same(C.stripEnd(p.richText.text),target)?i:-1).filter(i=>i>=0);
  if(candidates.length===1)return candidates[0];
 }
 // Repeated/blank text still needs complete ordered structural evidence.
 if(items.some(p=>p.type!=='RichText'))return fail();
 const b=items.map(p=>C.stripEnd(p.richText.text));
 if(a.length===b.length){if(a.some((t,i)=>!same(t,b[i])))return fail();return dom.index;}
 if(a.length<b.length||(a.length+1)*(b.length+1)>1000000)return fail();
 const empty=t=>/^[ \u00a0]*$/.test(t),n=a.length,m=b.length;
 const forward=Array.from({length:n+1},()=>new Uint8Array(m+1));
 const backward=Array.from({length:n+1},()=>new Uint8Array(m+1));
 forward[0][0]=1;backward[n][m]=1;
 for(let i=0;i<n;i++)for(let j=0;j<=m;j++)if(forward[i][j]){
  if(empty(a[i]))forward[i+1][j]=1;
  if(j<m&&same(a[i],b[j]))forward[i+1][j+1]=1;
 }
 if(!forward[n][m])return fail();
 for(let i=n-1;i>=0;i--)for(let j=m;j>=0;j--)backward[i][j]=Number((empty(a[i])&&backward[i+1][j])||(j<m&&same(a[i],b[j])&&backward[i+1][j+1]));
 const i=dom.index,candidates=[];
 for(let j=0;j<=m;j++)if(forward[i][j]){
  if(empty(a[i])&&backward[i+1][j])return fail();
  if(j<m&&same(a[i],b[j])&&backward[i+1][j+1])candidates.push(j);
 }
 if(candidates.length!==1)return fail();return candidates[0];
}
async function snapshot(hint,includeHtml=true,anchor=null,afterWrite=false,capture=false,dom=null){return OneNote.run(async ctx=>{
 const app=ctx.application,book=app.getActiveNotebook(),page=app.getActivePage();
 let p=anchor?null:app.getActiveParagraph(),textLoaded=false,location;
 const outline=app.getActiveOutline();
 book.load('id');page.load('id,title');if(p)p.load('id,type');await ctx.sync();
 if(!$('allpages').checked&&page.title.trim()!=='NoteMark Web Test')throw Error('已限制为测试页；请开启“允许在其他页面使用”。');
 if(anchor||capture){
  outline.load('id');const ps=outline.paragraphs;ps.load('items/id,items/type,items/richText/text');await ctx.sync();
  if(anchor){
   if(book.id!==anchor.notebook||page.id!==anchor.page||outline.id!==anchor.outline)throw Error('当前笔记块已变化，已停止定位。');
   const extra=afterWrite&&ps.items.length===anchor.count+1&&ps.items[anchor.index+1]?.type==='RichText'&&C.stripEnd(ps.items[anchor.index+1].richText.text).trim()==='';
   if(ps.items.length!==anchor.count&&!extra)throw Error('写入后段落数量变化（'+anchor.count+'→'+ps.items.length+'），已停止定位。');
   if(extra&&(!anchor.ids||ps.items.filter((x,i)=>i!==anchor.index+1).some((x,i)=>i!==anchor.index&&x.id!==anchor.ids[i])))throw Error('相邻段落已变化，已停止定位。');
   if(!afterWrite&&anchor.ids&&ps.items.some((x,i)=>i!==anchor.index&&x.id!==anchor.ids[i]))throw Error('相邻段落已变化，已停止定位。');
   p=afterWrite?ps.items[anchor.index]:ps.items.find(x=>x.id===anchor.paragraph);
   if(!p||(!afterWrite&&ps.items[anchor.index]?.id!==p.id))throw Error('原段落已移动，已取消。');
   if(extra){
    const blank=ps.items[anchor.index+1];
    if(anchor.ids.includes(blank.id)||C.stripEnd(p.richText.text)!==hint)throw Error('新增空段落核对失败，已停止清理。');
    blank.paragraphs.load('items/id');await ctx.sync();
    if(blank.paragraphs.items.length)throw Error('新增段落包含子内容，已停止清理。');
    // Only the new, empty sibling from this write may be removed; existing blanks stay.
    blank.delete();await ctx.sync();
    // OneNote caches collection items within a request context after delete.
    // Re-read in a fresh context without enabling another cleanup attempt.
    const clean=await snapshot(hint,includeHtml,{...anchor,paragraph:p.id},false);
    return clean;
   }
   textLoaded=true;location=anchor;
  }else{
   if(p.isNullObject){
    p=ps.items[locateDomParagraph(dom,ps.items)];
   }
   const index=ps.items.findIndex(x=>x.id===p.id);if(index<0)throw Error('无法记录当前段落位置。');
   location={notebook:book.id,page:page.id,outline:outline.id,paragraph:p.id,index,count:ps.items.length,ids:ps.items.map(x=>x.id)};
  }
 }
 if(p.isNullObject){
  const selected=C.stripEnd(await (hint??read()));
  const ps=app.getActiveOutline().paragraphs;
  ps.load('items/id,items/type,items/richText/text');await ctx.sync();
  const matches=ps.items.filter(x=>x.type==='RichText'&&C.stripEnd(x.richText.text)===selected);
  if(matches.length!==1)throw Error('不能唯一定位当前段落，请将光标放在一行文字中重试。');
  p=matches[0];textLoaded=true;
 }
 if(p.type!=='RichText')throw Error('当前段落不是文字。');
 const listInfo=p.getParagraphInfo?.();let html;
 if(includeHtml){p.richText.load('text');html=p.richText.getHtml();await ctx.sync();}
 else if(!textLoaded){p.richText.load('text');await ctx.sync();}else if(listInfo)await ctx.sync();
 const info=listInfo?.value;let list=anchor?.list||dom?.list||null;if(info&&info.listType!=='None'){if(info.listType==='Number'&&info.numberType!=='Arabic')throw Error('暂不转换这种编号样式。');list={type:info.listType==='Number'?'ol':'ul',start:Math.max(1,Number(info.index)||1)};}
 const visualList=anchor?.visualList||dom?.visualList||null;if(location&&capture){location.list=list;location.visualList=visualList;}
 return {list,visualList,location,notebook:book.id,page:page.id,paragraph:p.id,text:C.stripEnd(p.richText.text),html:html?.value};
});}
const read=()=>new Promise((resolve,reject)=>Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,r=>r.status===Office.AsyncResultStatus.Succeeded?resolve(r.value):reject(r.error)));
const write=html=>new Promise((resolve,reject)=>Office.context.document.setSelectedDataAsync(html,{coercionType:Office.CoercionType.Html},r=>r.status===Office.AsyncResultStatus.Succeeded?resolve():reject(r.error)));
async function run(mode,fromBridge=false,anchor=null){
 if(!ready||busy)throw Error('正在处理，请稍后重试。');busy=true;controls.forEach(id=>$(id).disabled=true);
 const start=performance.now();let wrote=false;
 try{
  const selectedRequest=read();
  const [before,selected]=await Promise.all([snapshot(selectedRequest,mode!=='convert',anchor),selectedRequest]);
  if(mode==='inspect'){report({context:before,selected});status('已读取当前段落。');return {};}
  const readMs=performance.now()-start;
  const raw=C.stripEnd(selected);
  if(!raw||raw!==before.text)throw Error('请完整选中当前段落，避免覆盖部分文字。');
  let html,source,ranges=[],generated=false;
  if(mode==='toggle'){
   const exact=store.get(before,before.text,before.html,{allowStale:true});
   mode=exact!==null?'restore':(!before.list&&!before.visualList&&C.parse(raw).changed?'convert':'restore');
  }
  const restoring=mode==='restore'||mode==='generate';
  if(restoring){
   source=mode==='generate'?null:store.get(before,before.text,before.html,{allowStale:true});
   if(source===null){source=C.fromHtml(before.html,before.list,before.visualList);generated=true;}
   html='<p>'+C.esc(source)+'</p>';
  }
  else{source=before.list&&!C.parse(raw).list?(before.list.type==='ol'?before.list.start+'. ':'- ')+raw:raw;const parsed=C.renderHaru(source,before.list);if(!parsed.changed){status('当前行没有可转换的标记。');return {unchanged:true,text:raw};}html=parsed.html;ranges=highlights(html);if(ranges.length&&!fromBridge)throw Error('高亮需要配套网页扩展；本次未修改正文。');}
  if(restoring&&before.list&&!fromBridge)throw Error('请用 ⌘, 或重新生成按钮还原列表项。');
  store.prepare(before,restoring?before.text:source);
  if(await read()!==selected)throw Error('选区已变化，已取消。');
  // Do not retry writes: a delayed successful write must never be duplicated.
  const tw=performance.now();wrote=true;await write(html);const writeMs=performance.now()-tw;
  const expected=restoring?source:C.renderHaru(source,before.list).text;
  if(ranges.length)await hostHighlight(expected,ranges);
  if(before.list)await hostRemoveList(expected);
  const after=await snapshot(expected,!restoring,anchor,true);
  if(after.text!==expected)throw Error('接口已返回，但正文未通过核对。请检查正文；未自动重试。');
  if(!restoring)store.put({...after,list:C.renderHaru(source,before.list).list},source,after.text,after.html);
  const result={list:restoring?null:C.renderHaru(source,before.list).list,mode,generated,text:expected,readMs,writeMs,verifyMs:performance.now()-tw-writeMs,totalMs:performance.now()-start,paragraph:after.paragraph};report(result);status((generated?'已根据当前格式生成 Markdown（接口未提供的高亮无法恢复）。':'已完成。')+' 耗时 '+Math.round(result.totalMs)+' ms。');return result;
 }catch(e){e.wrote=wrote;status(e.message||String(e));report({error:e.message||String(e),code:e.code});throw e;}
 finally{busy=false;controls.forEach(id=>$(id).disabled=!ready);}
}
async function continueList(anchor){
 if(!ready||busy)throw Error('正在处理，请稍后重试。');
 const prefix=anchor?.next;if(typeof prefix!=='string'||! /^(?:\d{1,9}\.|-)\u00a0$/.test(prefix))throw Error('无效的列表续项。');
 busy=true;
 try{
  const before=await snapshot('',false,anchor);if(before.text.trim()||C.stripEnd(await read()).trim())throw Error('下一行已有输入，取消自动续项。');
  await write('<p><span style="font-family:Glow Sans;font-size:12.5pt">'+C.esc(prefix)+'</span></p>');
  const after=await snapshot(prefix,false,anchor,true);if(after.text!==prefix)throw Error('续项文字核对失败，请检查下一行。');
  return {text:prefix};
 }finally{busy=false;}
}
function generateCurrent(){
 if(!$('bridge').checked||Date.now()-lastBridge>=4000)return run('generate');
 const id=crypto.randomUUID();return new Promise((resolve,reject)=>{
  const timer=setTimeout(()=>{hostPending.delete(id);reject(Error('重新生成未响应，请重新点击正文段落后重试。'));},15000);
  hostPending.set(id,{resolve,reject,timer});
  parent.postMessage({channel:'notemark.v2',hostCommand:'generate',id},'https://onenote.officeapps.live.com');
 });
}
for(const id of controls)$(id).onclick=()=> (id==='generate'?generateCurrent():run(id)).catch(e=>status(e.message||String(e)));
$('backup').onclick=()=>{try{const r=JSON.parse(localStorage.getItem('notemark.pending.v1')||'null');$('backuptext').hidden=false;$('backuptext').value=r?.source||'没有备份记录';}catch(e){status('无法读取原文备份：'+e.message);}};
$('clear').onclick=()=>{if(!store||!confirm('清空后无法精确恢复旧原文。确定清空本地记录？'))return;store.clear();$('backuptext').value='';status('本地原文记录已清空。');};
// The paired extension runs only in the OneNote editor. Do not accept other origins or nested senders.
window.addEventListener('message',async e=>{
 if(e.origin!=='https://onenote.officeapps.live.com'||e.source!==parent)return;
 if(e.data?.channel==='notemark.host'){const p=hostPending.get(e.data.id);if(!p)return;clearTimeout(p.timer);hostPending.delete(e.data.id);e.data.ok?p.resolve():p.reject(Error(e.data.error));return;}
 if(e.data?.channel!=='notemark.v2')return;
 if(e.data.mode==='ping'){lastBridge=Date.now();updateBridge();}
 if(e.data.mode!=='ping'&&!$('bridge').checked)return;
 const {id,mode}=e.data;if(typeof id!=='string'||!['convert','restore','toggle','generate','ping','anchor','continue'].includes(mode))return;
 try{const result=mode==='ping'?{ready:ready&&!busy&&$('bridge').checked,autoEnter:$('autoenter').checked}:mode==='anchor'?await snapshot(undefined,false,null,false,true,e.data.anchor):mode==='continue'?await continueList(e.data.anchor):await run(mode,true,e.data.anchor);e.source.postMessage({channel:'notemark.v2',id,ok:true,result},e.origin);}
 catch(err){e.source.postMessage({channel:'notemark.v2',id,ok:false,error:err.message||String(err),wrote:err.wrote},e.origin);}
});
if(typeof Office==='undefined')status('Office.js 未加载。');
else Office.onReady(info=>{ready=info.host==='OneNote'&&!!store;controls.forEach(id=>$(id).disabled=!ready);status(ready?'已连接 OneNote。':'请从 OneNote 加载项菜单打开。');report(info);});
