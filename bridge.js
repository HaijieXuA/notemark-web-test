// Experimental bridge. All document writes go through the paired Office.js pane.
(()=>{
'use strict';
window.__NoteMarkBridge?.stop();
let busy=false,composing=false,lastComposition=0,availableUntil=0,autoEnter=false;
const origin='https://haijiexua.github.io',pending=new Map(),seen=new Set();
const editor=()=>document.getElementById('WACViewPanel_EditingElement');
const frame=()=>[...document.querySelectorAll('iframe')].find(f=>{try{const u=new URL(f.src);return u.origin===origin&&u.pathname==='/notemark-web-test/taskpane.html';}catch{return false;}});
const tick=()=>new Promise(r=>setTimeout(r,0));
function key(k,n,mods={}){const el=editor();if(!el)throw Error('编辑区未就绪');el.focus();el.dispatchEvent(new KeyboardEvent('keydown',{key:k,code:k,keyCode:n,which:n,bubbles:true,cancelable:true,...mods}));}
function notice(s){let el=document.getElementById('notemark-status');if(!el){el=document.createElement('div');el.id='notemark-status';el.style.cssText='position:fixed;bottom:15px;right:20px;z-index:2147483647;background:#34213f;color:white;padding:10px 15px;border-radius:8px;max-width:360px;font:14px system-ui';document.body.append(el);}el.textContent=s;clearTimeout(el.timer);el.timer=setTimeout(()=>el.remove(),5000);}
function request(mode){const f=frame();if(!f)return Promise.reject(Error('请打开 NoteMark 窗格'));const id=crypto.randomUUID();return new Promise((resolve,reject)=>{const timer=setTimeout(()=>{pending.delete(id);availableUntil=0;reject(Error('加载项响应超时，请检查正文后重试'));},10000);pending.set(id,{resolve,reject,timer,source:f.contentWindow});f.contentWindow.postMessage({channel:'notemark.v2',id,mode},origin);});}
function text(){return [...editor().querySelectorAll('.TextRun')].map(n=>n.textContent).join('');}
function focusParagraph(expected){
 const matches=[...document.querySelectorAll('p.Paragraph')].filter(p=>[...p.querySelectorAll('.TextRun')].map(n=>n.textContent).join('')===expected);
 if(matches.length!==1)throw Error('无法唯一定位转换后的段落，已停止移动光标');
 const p=matches[0],r=p.getBoundingClientRect();
 if(r.width===0||r.height===0)throw Error('当前段落不可见');
 for(const type of ['mousedown','mouseup','click'])p.dispatchEvent(new MouseEvent(type,{bubbles:true,cancelable:true,view:window,clientX:r.right-2,clientY:r.top+r.height/2,button:0,buttons:type==='mousedown'?1:0}));
 if(!editor()||text()!==expected)throw Error('正文焦点核对失败');
}
async function highlight(expected,ranges){
 focusParagraph(expected);
 if(ranges.length){document.getElementById('Home')?.click();await tick();}
 for(const range of ranges){
  const b=document.getElementById('Highlighter')?.querySelector('button');
  if(!b||!b.getAttribute('aria-label')?.includes('Yellow'))throw Error('请先将 OneNote 高亮颜色设置为黄色');
  key('Home',36);
  const s=getSelection();
  // Selection.modify performs the browser's default caret movement, unlike synthetic ArrowRight.
  for(let i=0;i<range.start;i++)s.modify('move','forward','character');
  for(let i=0;i<range.length;i++)s.modify('extend','forward','character');
  if(s.toString()!==range.text)throw Error('高亮选区核对失败');
  await tick();b.click();await tick();
 }
 return {text:expected};
}
async function message(e){
 if(e.origin!==origin||e.source!==frame()?.contentWindow||e.data?.channel!=='notemark.v2')return;
 if(e.data.hostCommand==='highlight'){
  const {id,text:expected,ranges}=e.data;
  if(seen.has(id))return;seen.add(id);if(seen.size>100)seen.delete(seen.values().next().value);
  try{if(!busy)throw Error('没有活动转换');await highlight(expected,ranges);e.source.postMessage({channel:'notemark.host',id,ok:true},origin);}
  catch(err){e.source.postMessage({channel:'notemark.host',id,ok:false,error:err.message},origin);}return;
 }
 const p=pending.get(e.data.id);if(!p||e.source!==p.source)return;
 clearTimeout(p.timer);pending.delete(e.data.id);
 e.data.ok?p.resolve(e.data.result):p.reject(Object.assign(Error(e.data.error),{wrote:e.data.wrote}));
}
function atEnd(){const s=getSelection();if(!s?.isCollapsed||!s.rangeCount)return false;const p=editor().querySelector('p');if(!p)return false;try{const r=s.getRangeAt(0).cloneRange();r.setStart(p,0);return r.toString().length>=text().length;}catch{return false;}}
async function convert(mode,newline){
 busy=true;
 try{
  key('Home',36);key('End',35,{shiftKey:true});
  const result=await request(mode);
  focusParagraph(result.text);key('End',35);
  if(newline){key('Enter',13);document.getElementById('ClearFormatting')?.click();}
  notice(result.unchanged?'没有可转换标记':'NoteMark 已转换');
 }catch(e){notice(e.message);if(e.wrote===false&&editor()){key('End',35);if(newline)key('Enter',13);}}
 finally{busy=false;}
}
function start(){composing=true;availableUntil=0;}
function end(){composing=false;lastComposition=performance.now();}
function onKey(e){
 if(!e.isTrusted||!editor()?.contains(e.target))return;
 // Auto mode remains opt-in. During a request, keep ordinary typing available.
 if(busy)return;
 if(composing||e.isComposing||e.keyCode===229||performance.now()-lastComposition<100||e.repeat)return;
 const restore=e.metaKey&&!e.ctrlKey&&!e.altKey&&!e.shiftKey&&e.key===',';
 const manual=e.metaKey&&e.altKey&&!e.ctrlKey&&e.key==='Enter';
 const enter=e.key==='Enter'&&!e.metaKey&&!e.ctrlKey&&!e.altKey&&!e.shiftKey;
 if(!restore&&!manual&&!enter)return;
 if(enter&&!autoEnter)return;
 if(performance.now()>availableUntil){if(manual||restore)notice('请打开 NoteMark 窗格并启用扩展选项');return;}
 if(enter&&(!atEnd()||!/(^#{1,6}\s)|\*|~~|==/.test(text())))return;
 e.preventDefault();e.stopImmediatePropagation();convert(restore?'restore':'convert',enter);
}
const heartbeat=setInterval(()=>{if(!busy&&!composing&&frame())request('ping').then(r=>{availableUntil=r.ready?performance.now()+2500:0;autoEnter=!!r.autoEnter;}).catch(()=>{availableUntil=0;});},1500);
window.addEventListener('message',message);document.addEventListener('keydown',onKey,true);document.addEventListener('compositionstart',start,true);document.addEventListener('compositionend',end,true);
window.__NoteMarkBridge={stop(){clearInterval(heartbeat);window.removeEventListener('message',message);document.removeEventListener('keydown',onKey,true);document.removeEventListener('compositionstart',start,true);document.removeEventListener('compositionend',end,true);for(const p of pending.values()){clearTimeout(p.timer);p.reject(Error('扩展停止'));}pending.clear();}};
})();
