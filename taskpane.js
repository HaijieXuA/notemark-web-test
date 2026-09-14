'use strict';
const C=NoteMarkCore,$=id=>document.getElementById(id),controls=['convert','restore','inspect'];
let ready=false,busy=false,store;
try{store=new C.SourceStore(localStorage);}catch(e){$('status').textContent='浏览器存储不可用：'+e.message;}
function report(v){$('log').textContent=(JSON.stringify(v,null,2)+'\n'+$('log').textContent).slice(0,12000);}
function status(s){$('status').textContent=s;}
async function snapshot(){return OneNote.run(async ctx=>{
 const app=ctx.application,book=app.getActiveNotebook(),page=app.getActivePage(),p=app.getActiveParagraph();
 book.load('id');page.load('id,title');p.load('id,type');await ctx.sync();
 if(!$('allpages').checked&&page.title.trim()!=='NoteMark Web Test')throw Error('默认仅在 NoteMark Web Test 页操作。');
 if(p.type!=='RichText')throw Error('当前段落不是文字。');
 p.richText.load('text');const html=p.richText.getHtml();await ctx.sync();
 return {notebook:book.id,page:page.id,paragraph:p.id,text:C.stripEnd(p.richText.text),html:html.value};
});}
const read=()=>new Promise((resolve,reject)=>Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,r=>r.status===Office.AsyncResultStatus.Succeeded?resolve(r.value):reject(r.error)));
const write=html=>new Promise((resolve,reject)=>Office.context.document.setSelectedDataAsync(html,{coercionType:Office.CoercionType.Html},r=>r.status===Office.AsyncResultStatus.Succeeded?resolve():reject(r.error)));
async function run(mode){
 if(!ready||busy)throw Error('正在处理，请稍后重试。');busy=true;controls.forEach(id=>$(id).disabled=true);
 const start=performance.now();
 try{
  const [before,selected]=await Promise.all([snapshot(),read()]);
  if(mode==='inspect'){report({context:before,selected});status('已读取当前段落。');return {};}
  const raw=C.stripEnd(selected);
  if(!raw||raw!==before.text)throw Error('请完整选中当前段落，避免覆盖部分文字。');
  let html,source;
  if(mode==='restore'){source=store.get(before,before.text,before.html);if(source===null)throw Error('此段落没有匹配的原文记录。');html='<p>'+C.esc(source)+'</p>';}
  else{source=raw;const parsed=C.parse(source);if(!parsed.changed){status('当前行没有可转换的标记。');return {unchanged:true};}html=parsed.html;}
  if(mode!=='restore')store.prepare(before,source);
  if(await read()!==selected)throw Error('选区已变化，已取消。');
  // Do not retry writes: a delayed successful write must never be duplicated.
  const tw=performance.now();await write(html);const writeMs=performance.now()-tw;
  const after=await snapshot();
  const expected=mode==='restore'?source:C.parse(source).text;
  if(after.text!==expected)throw Error('接口已返回，但正文未通过核对。请检查正文；未自动重试。');
  if(mode!=='restore')store.put(after,source,after.text,after.html);
  const result={mode,writeMs,totalMs:performance.now()-start,paragraph:after.paragraph};report(result);status('已完成，耗时 '+Math.round(result.totalMs)+' ms。');return result;
 }catch(e){status(e.message||String(e));report({error:e.message||String(e),code:e.code});throw e;}
 finally{busy=false;controls.forEach(id=>$(id).disabled=!ready);}
}
for(const id of controls)$(id).onclick=()=>run(id).catch(()=>{});
$('clear').onclick=()=>{store.clear();status('本地原文记录已清空。');};
// The paired extension runs only in the OneNote editor. Do not accept other origins or nested senders.
window.addEventListener('message',async e=>{
 if(e.origin!=='https://onenote.officeapps.live.com'||e.source!==parent||e.data?.channel!=='notemark.v2'||!$('bridge').checked)return;
 const {id,mode}=e.data;if(typeof id!=='string'||!['convert','restore','ping'].includes(mode))return;
 try{const result=mode==='ping'?{ready}:await run(mode);e.source.postMessage({channel:'notemark.v2',id,ok:true,result},e.origin);}
 catch(err){e.source.postMessage({channel:'notemark.v2',id,ok:false,error:err.message||String(err)},e.origin);}
});
if(typeof Office==='undefined')status('Office.js 未加载。');
else Office.onReady(info=>{ready=info.host==='OneNote'&&!!store;controls.forEach(id=>$(id).disabled=!ready);status(ready?'已连接 OneNote。':'请从 OneNote 加载项菜单打开。');report(info);});
