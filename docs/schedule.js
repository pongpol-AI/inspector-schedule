(function (root) {
  'use strict';
  const months = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน','กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  const abbr = ['มค','กพ','มีค','เมย','พค','มิย','กค','สค','กย','ตค','พย','ธค'];
  const clean = v => v == null ? '' : String(v).trim();
  const compact = v => clean(v).replace(/[\s.]/g, '');
  function dateText(value) {
    if (value instanceof Date && !isNaN(value)) return `${value.getDate()} ${months[value.getMonth()]} ${value.getFullYear()+543}`;
    return clean(value);
  }
  function dateKey(text) {
    const value = compact(text);
    const month = months.findIndex((m,i) => value.includes(m) || value.includes(abbr[i]));
    const day = value.match(/^\d{1,2}/);
    const year = value.match(/(\d{4}|\d{2})$/);
    return `${year ? year[0].slice(-2) : '99'}-${String(month < 0 ? 99 : month+1).padStart(2,'0')}-${String(day ? day[0] : 99).padStart(2,'0')}`;
  }
  function parseRows(rows, sheetName='') {
    const hospitalNames = ['โรงพยาบาล','หน่วยงาน','สถานที่ตรวจ'];
    const h = rows.slice(0,20).findIndex(row => row.some(v => hospitalNames.includes(compact(v))));
    if (h < 0) throw new Error('ไม่พบหัวตาราง โรงพยาบาล / หน่วยงาน ภายใน 20 แถวแรก');
    const headers = rows[h].map(compact);
    const find = names => headers.findIndex(v => names.includes(v));
    const hospital = find(hospitalNames), province = find(['จังหวัด']);
    const date = find(['การนัดหมาย','วันที่ตรวจ','วันที่','วันตรวจ']);
    const leader = headers.findIndex(v => v.includes('หัวหน้า'));
    const inspectors = headers.map((v,i) => /^ผู้ตรวจ/.test(v) && i!==leader ? i : -1).filter(i=>i>=0);
    if (date < 0 || (leader < 0 && !inspectors.length)) throw new Error('ต้องมีคอลัมน์วันที่นัด และหัวหน้าทีมหรือผู้ตรวจ');
    let skipped=0;
    const jobs=[];
    rows.slice(h+1).forEach((row,i) => {
      if (!clean(row[hospital])) return;
      const when=dateText(row[date]);
      if (!when || /เลื่อน|ยกเลิก/.test(when)) { skipped++; return; }
      jobs.push({row:h+i+2,hospital:clean(row[hospital]),province:province<0?'':clean(row[province]),date:when,
        leader:leader<0?'':clean(row[leader]),inspectors:[...new Set(inspectors.map(c=>clean(row[c])).filter(Boolean))]});
    });
    if (!jobs.length) throw new Error('ไม่พบรายการที่มีหน่วยงานและวันนัดหมายในแผ่นงานนี้');
    jobs.sort((a,b)=>dateKey(a.date).localeCompare(dateKey(b.date)));
    const title = rows.slice(0,h).flat().map(clean).filter(Boolean).join(' ');
    const month = months.find(m=>title.includes(m)||sheetName.includes(m));
    const year = (title+' '+sheetName).match(/25\d{2}|\b\d{2}\b/);
    return {jobs,skipped,period:month ? `${month} ${year ? (year[0].length===2?'25':'')+year[0]:''}`.trim():sheetName};
  }
  function people(jobs) {
    const map=new Map();
    jobs.forEach(job => {
      [...new Set([job.leader,...job.inspectors].filter(Boolean))].forEach(name=>{
        if(!map.has(name)) map.set(name,[]);
        map.get(name).push({...job,role:name===job.leader?'หัวหน้าทีม':'ผู้ตรวจประเมิน'});
      });
    });
    return [...map].sort((a,b)=>b[1].length-a[1].length||a[0].localeCompare(b[0],'th'));
  }
  const api={parseRows,people,dateKey};
  if(typeof module!=='undefined') module.exports=api;
  else root.Schedule=api;
})(globalThis);
