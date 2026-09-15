'use strict';
const $=id=>document.getElementById(id);
let workbook=null, jobs=[], mode='individual', busy=false;
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
function status(message,error=false){$('status').textContent=message;$('status').classList.toggle('error',error);}
function render(){
  const allPeople=Schedule.people(jobs), query=$('search').value.trim().toLocaleLowerCase();
  const hasSchedule=jobs.length>0;
  $('metric-jobs').textContent=hasSchedule?jobs.length+' รายการ':'—';
  $('metric-people').textContent=hasSchedule?allPeople.length+' คน':'—';
  $('metric-month').textContent=$('sheet').value||(workbook?'กรุณาเลือกแท็บเดือน':'เลือกไฟล์ Excel');
  const entries=mode==='individual'?allPeople:jobs.map(j=>[j.hospital,[j]]);
  const shown=entries.filter(([name,list])=>(name+' '+list.map(j=>[j.hospital,j.province,j.date,j.leader,...j.inspectors].join(' ')).join(' ')).toLocaleLowerCase().includes(query));
  $('workspace').hidden=!jobs.length;$('empty').hidden=!!jobs.length;
  $('stats').textContent=`ผู้ตรวจ ${allPeople.length} คน • ตาราง ${jobs.length} รายการ • แสดง ${shown.length} การ์ด`;
  $('cards').replaceChildren();
  for(const [name,list] of shown){
    const card=document.createElement('article');card.className='card';card.dataset.name=name;
    const colors=['#a75a40','#397653','#a04e79','#b37b28','#306da0'];
    card.style.setProperty('--accent',colors[Math.min(list.length,5)-1]);
    card.innerHTML=`<div class="card-head"><small>ตารางตรวจประเมิน · ${esc($('period').value)}</small><h2>${esc(name)}</h2><small>${mode==='individual'?`${list.length} รายการตรวจ`:'การแจ้งเตือนทีมผู้ตรวจประเมิน'}</small></div><div class="card-body">${list.map(j=>`<div class="job"><div class="date">${esc(j.date)}${j.role?`<span class="role">${esc(j.role)}</span>`:''}</div><div class="hospital">${esc(j.hospital)}</div><div class="muted">${esc(j.province)}</div><p class="muted">หัวหน้าทีม: ${esc(j.leader||'ยังไม่ระบุ')}<br>ผู้ตรวจ: ${esc(j.inspectors.join(', ')||'ยังไม่ระบุ')}</p></div>`).join('')}</div><div class="card-actions no-print" data-html2canvas-ignore="true"><button data-action="image">บันทึก PNG</button><button data-action="print" class="secondary">พิมพ์ / PDF</button></div>`;
    $('cards').append(card);
  }
  if(!shown.length&&jobs.length) $('cards').textContent='ไม่พบข้อมูลที่ตรงกับคำค้นหา';
  $('zip').disabled=!shown.length;$('print').disabled=!shown.length;
}
function loadSheet(){
  jobs=[];$('period').value='';$('search').value='';
  if(!$('sheet').value){status('เลือกเดือน / แท็บใน Excel เพื่อสร้างการ์ด');render();return;}
  try{
    const rows=XLSX.utils.sheet_to_json(workbook.Sheets[$('sheet').value],{header:1,defval:null,raw:true});
    const result=Schedule.parseRows(rows,$('sheet').value);jobs=result.jobs;$('period').value=result.period;
    status(`แท็บ ${$('sheet').value} • นำเข้า ${jobs.length} รายการสำเร็จ${result.skipped?` • ข้าม ${result.skipped} แถวที่ไม่มีวันนัดหรือเลื่อน/ยกเลิก`:''} • โปรดตรวจสอบวันและชื่อก่อนส่ง`);
  }catch(error){jobs=[];status(error.message,true);}
  render();
}
$('file').addEventListener('change',async event=>{
  const file=event.target.files[0];if(!file)return;
  $('file').disabled=true;$('sheet').disabled=true;
  try{
    if(!globalThis.XLSX)throw new Error('โหลดตัวอ่าน Excel ไม่สำเร็จ กรุณารีเฟรชหน้าเว็บ');
    if(file.size>20*1024*1024)throw new Error('กรุณาใช้ไฟล์ขนาดไม่เกิน 20 MB');
    status('กำลังอ่านไฟล์…');
    workbook=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:true});
    $('file-info').textContent=`ไฟล์: ${file.name} • ${workbook.SheetNames.length} แท็บ`;
    $('sheet').replaceChildren(new Option('— เลือกเดือน / แท็บที่ต้องการ —',''),...workbook.SheetNames.map(name=>new Option(name,name)));
    if(workbook.SheetNames.length===1)$('sheet').value=workbook.SheetNames[0];
    $('sheet').disabled=false;loadSheet();
  }catch(error){workbook=null;jobs=[];$('period').value='';$('file-info').textContent='ยังไม่ได้เลือกไฟล์ที่อ่านได้';$('sheet').replaceChildren(new Option('เลือกไฟล์ Excel ก่อน',''));$('sheet').disabled=true;status('อ่านไฟล์ไม่สำเร็จ: '+error.message,true);render();}
  finally{event.target.value='';$('file').disabled=false;}
});
$('sheet').addEventListener('change',loadSheet);
$('period').addEventListener('input',render);$('search').addEventListener('input',render);
for(const tab of ['individual','group'])$(tab).addEventListener('click',()=>{mode=tab;for(const t of ['individual','group'])$(t).setAttribute('aria-pressed',String(t===tab));render();});
$('clear').addEventListener('click',()=>{workbook=null;jobs=[];$('file-info').textContent='ยังไม่ได้เลือกไฟล์ • เลือกได้จากทุกโฟลเดอร์ในเครื่องหรือไดรฟ์ที่เข้าถึงได้';$('sheet').replaceChildren(new Option('เลือกไฟล์ Excel ก่อน',''));$('sheet').disabled=true;$('period').value='';$('search').value='';status('ล้างข้อมูลแล้ว เลือกไฟล์ Excel เพื่อเริ่มใหม่');render();});
function download(blob,name){const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),10000);}
function filename(name){return name.replace(/[<>:"/\\|?*\x00-\x1f]/g,'-').slice(0,120);}
async function capture(card){
  if(!globalThis.html2canvas)throw new Error('โหลดตัวบันทึกภาพไม่สำเร็จ กรุณารีเฟรชหน้าเว็บ');
  const canvas=await html2canvas(card,{scale:2,backgroundColor:'#ffffff',logging:false});
  return new Promise((resolve,reject)=>canvas.toBlob(blob=>blob?resolve(blob):reject(new Error('สร้างภาพไม่สำเร็จ')),'image/png'));
}
async function exporting(action){if(busy)return;busy=true;const controls=[...document.querySelectorAll('button,input,select')],states=controls.map(c=>c.disabled);controls.forEach(c=>c.disabled=true);try{await action();status('ดาวน์โหลดสำเร็จ ตรวจสอบไฟล์ในโฟลเดอร์ดาวน์โหลด');}catch(error){status('ส่งออกไม่สำเร็จ: '+error.message,true);}finally{controls.forEach((c,i)=>c.disabled=states[i]);busy=false;}}
$('cards').addEventListener('click',event=>{const button=event.target.closest('button');if(!button)return;const card=button.closest('.card');if(button.dataset.action==='image')exporting(async()=>download(await capture(card),filename(card.dataset.name)+'.png'));else{card.classList.add('print-target');document.body.classList.add('printing-one');window.print();}});
window.addEventListener('afterprint',()=>{document.body.classList.remove('printing-one');document.querySelectorAll('.print-target').forEach(c=>c.classList.remove('print-target'));});
$('print').addEventListener('click',()=>window.print());
$('zip').addEventListener('click',()=>exporting(async()=>{
  if(!globalThis.JSZip)throw new Error('โหลดตัวสร้าง ZIP ไม่สำเร็จ กรุณารีเฟรชหน้าเว็บ');
  const zip=new JSZip(),cards=[...$('cards').querySelectorAll('.card')];
  for(let i=0;i<cards.length;i++){status(`กำลังสร้างภาพ ${i+1}/${cards.length}…`);zip.file(`${i+1}_${filename(cards[i].dataset.name)}.png`,await capture(cards[i]));}
  download(await zip.generateAsync({type:'blob'}),`ตารางผู้ตรวจ_${filename($('period').value)}.zip`);
}));
