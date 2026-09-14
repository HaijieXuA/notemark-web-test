/* Minimal API experiment: one committed line, no keyboard interception. */
'use strict';
const log = document.getElementById('log');
const buttons = [...document.querySelectorAll('button')];
function report(value) { log.textContent += JSON.stringify(value, null, 2) + '\n'; }
function escapeHtml(s) { return s.replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
function render(s) {
  return '<p>' + escapeHtml(s).replace(/\*\*([^*]+)\*\*|\*([^*]+)\*|~~([^~]+)~~|==([^=]+)==/g,
    (_, bold, italic, strike, highlight) => bold !== undefined ? '<strong>'+bold+'</strong>' : italic !== undefined ? '<em>'+italic+'</em>' : strike !== undefined ? '<span style="text-decoration:line-through">'+strike+'</span>' : '<span style="background-color:#ffff00">'+highlight+'</span>') + '</p>';
}
async function pageInfo() {
  return OneNote.run(async context => {
    const page = context.application.getActivePage();
    page.load('id,title');
    await context.sync();
    if (page.title.trim() !== 'NoteMark Web Test') throw Error('请打开 NoteMark Web Test 测试页。');
    return {id:page.id, title:page.title};
  });
}
function readSelection() {
  return new Promise((resolve,reject) => Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
    r => r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error)));
}
function writeSelection(html) {
  return new Promise((resolve,reject) => Office.context.document.setSelectedDataAsync(html,{coercionType:Office.CoercionType.Html},
    r => r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(r.error)));
}
async function run(convert) {
  buttons.forEach(b=>b.disabled=true);
  try {
    const t0=performance.now(), page=await pageInfo(), t1=performance.now();
    const source=await readSelection(), t2=performance.now();
    report({page:page.title, source, guardMs:t1-t0, readMs:t2-t1});
    if (!convert) return;
    if (typeof source!=='string' || !/^API-\d+ /.test(source) || /[\r\n]/.test(source) || source.length>500) throw Error('只转换以 API-数字 开头、长度小于 500 的单行测试选区。');
    const html=render(source);
    // Re-check page and selected text after asynchronous reads before mutating.
    const again=await pageInfo();
    if(again.id!==page.id || await readSelection()!==source) throw Error('选区已变化，已取消。');
    const tw=performance.now();
    await writeSelection(html);
    report({result:'HTML 写入接口已成功返回；请核对正文、一次撤销及刷新保存。',writeMs:performance.now()-tw,totalMs:performance.now()-t0,html});
  } catch(e) { report({error:e.message || String(e),code:e.code}); }
  finally { buttons.forEach(b=>b.disabled=false); }
}
if (typeof Office === 'undefined') document.getElementById('status').textContent='Office.js 未加载，请检查微软脚本连接。';
else Office.onReady(info=>{
  report({host:info.host,platform:info.platform});
  document.getElementById('status').textContent=info.host==='OneNote'?'已连接 OneNote 官方 API':'请从 OneNote 的加载项菜单打开本页。';
  if(info.host==='OneNote') buttons.forEach(b=>b.disabled=false);
});
document.getElementById('inspect').onclick=()=>run(false);
document.getElementById('convert').onclick=()=>run(true);
