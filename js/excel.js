// Load SheetJS library
const sheetScript=document.createElement('script');
sheetScript.src='https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
document.head.appendChild(sheetScript);

const EXCEL_TEMPLATES={
  resources:{
    headers:['type','name','cpuCap','ramCap','storCap','note'],
    rows:[
      ['cluster','Cluster-1',128,512,'','Cluster produksi'],
      ['cluster','Cluster-2',64,256,'','Cluster staging'],
      ['storage_core','SAN-Core-A','','',50,'Storage core aplikasi'],
      ['storage_support','SAN-Support-A','','',30,'Storage support'],
    ],
    notes:[
      '* type: cluster / storage_core / storage_support',
      '* cpuCap & ramCap diisi untuk cluster, storCap untuk storage',
      '* storCap dalam TB, ramCap dalam GB, cpuCap dalam cores',
    ]
  },
  history:{
    headers:['resName','date','label','cpu','ram','stor'],
    rows:[
      ['Cluster-1','2025-01-06','Minggu 1 Jan',65.5,78.2,''],
      ['Cluster-1','2025-01-13','Minggu 2 Jan',67.1,80.0,''],
      ['SAN-Core-A','2025-01-06','Minggu 1 Jan','','',55.3],
    ],
    notes:[
      '* resName harus sama persis dengan nama resource yang sudah ada',
      '* date format: YYYY-MM-DD (contoh: 2025-01-06)',
      '* cpu & ram diisi untuk cluster, stor untuk storage',
      '* Kosongkan kolom yang tidak berlaku (jangan dihapus kolomnya)',
    ]
  },
  projections:{
    headers:['resName','metric','year','projected'],
    rows:[
      ['Cluster-1','cpu',2026,204],
      ['Cluster-1','ram',2026,380],
      ['SAN-Core-A','stor',2026,35],
    ],
    notes:[
      '* resName harus sama persis dengan nama resource yang sudah ada',
      '* metric: cpu / ram / stor',
      '* year: tahun proyeksi (contoh: 2026)',
      '* projected: nilai ABSOLUT yang diproyeksikan akan digunakan',
      '  - CPU dalam cores (contoh: 204 artinya 204 cores)',
      '  - RAM dalam GB (contoh: 380 artinya 380 GB)',
      '  - Storage dalam TB (contoh: 35 artinya 35 TB)',
      '* Nilai projected HARUS lebih kecil dari kapasitas total resource',
    ]
  }
};

function downloadTemplate(type){
  const tmpl=EXCEL_TEMPLATES[type];
  const wb=XLSX.utils.book_new();
  const wsData=[tmpl.headers,...tmpl.rows];
  const ws=XLSX.utils.aoa_to_sheet(wsData);
  const colWidths=tmpl.headers.map((h,i)=>({wch:Math.max(h.length,...tmpl.rows.map(r=>String(r[i]||'').length))+4}));
  ws['!cols']=colWidths;
  const notesData=[['PANDUAN PENGISIAN'],...tmpl.notes.map(n=>[n]),[''],['Hapus baris contoh ini sebelum upload!']];
  const wsNotes=XLSX.utils.aoa_to_sheet(notesData);
  wsNotes['!cols']=[{wch:60}];
  XLSX.utils.book_append_sheet(wb,ws,'Data');
  XLSX.utils.book_append_sheet(wb,wsNotes,'Panduan');
  XLSX.writeFile(wb,`template_${type}.xlsx`);
}

function safeFloat(v){
  if(v===''||v==null) return '';
  const n=parseFloat(String(v).replace(',','.'));
  return isNaN(n)?'':n;
}

function excelDateToString(val){
  if(typeof val==='number'){
    const date=new Date(Math.round((val-25569)*86400*1000));
    const y=date.getUTCFullYear();
    const m=String(date.getUTCMonth()+1).padStart(2,'0');
    const d=String(date.getUTCDate()).padStart(2,'0');
    return `${y}-${m}-${d}`;
  }
  if(typeof val==='string') return val.trim();
  return '';
}

function renderPreview(containerId,rows,columns){
  if(!rows.length) return;
  const wrap=document.getElementById(containerId); if(!wrap) return;
  wrap.innerHTML=`
    <div style="font-size:11px;color:var(--accent3);margin-bottom:6px;font-family:'Space Mono',monospace">${rows.length} baris terdeteksi — Preview 5 baris pertama:</div>
    <div style="overflow-x:auto"><table>
      <thead><tr>${columns.map(c=>`<th>${c}</th>`).join('')}</tr></thead>
      <tbody>${rows.slice(0,5).map(r=>`<tr>${columns.map(c=>`<td style="font-size:11px">${r[c]!==undefined&&r[c]!==''?r[c]:'—'}</td>`).join('')}</tr>`).join('')}</tbody>
    </table></div>`;
}

async function uploadExcel(type){
  const fileInput=document.getElementById(`csv-${type}`);
  const msgId=type==='resources'?'csv-res-msg':type==='history'?'csv-hist-msg':'csv-proj-msg';
  const previewId=type==='resources'?'csv-res-preview':type==='history'?'csv-hist-preview':'csv-proj-preview';

  if(!fileInput.files.length) return fmsg(msgId,'✗ Pilih file Excel terlebih dahulu.','var(--danger)');
  if(typeof XLSX==='undefined') return fmsg(msgId,'✗ Library Excel belum siap, coba lagi.','var(--danger)');

  const buf=await fileInput.files[0].arrayBuffer();
  const wb=XLSX.read(buf,{type:'array',cellDates:false});
  const ws=wb.Sheets[wb.SheetNames[0]];
  const raw=XLSX.utils.sheet_to_json(ws,{defval:''});
  if(!raw.length) return fmsg(msgId,'✗ File Excel kosong atau format tidak valid.','var(--danger)');

  const rows=raw.map(r=>{
    const obj={};
    Object.keys(r).forEach(k=>{ obj[k.toLowerCase().trim()]=r[k]; });
    return obj;
  });

  if(type==='resources'){
    renderPreview(previewId,rows,['type','name','cpucap','ramcap','storcap','note']);
    const newRes=rows.map(r=>({
      id:Date.now()+Math.random(),
      type:String(r.type||'').trim(),
      name:String(r.name||'').trim(),
      cpuCap:safeFloat(r.cpucap),
      ramCap:safeFloat(r.ramcap),
      storCap:safeFloat(r.storcap),
      note:String(r.note||'').trim()
    })).filter(r=>r.name&&r.type);
    if(!newRes.length) return fmsg(msgId,'✗ Tidak ada data valid. Pastikan kolom type dan name terisi.','var(--danger)');
    resources=[...resources,...newRes];
    await syncData('saveResources',resources);
    renderResTable();
    fmsg(msgId,`✓ ${newRes.length} resource berhasil diimport!`,'var(--accent3)');
  }
  else if(type==='history'){
    renderPreview(previewId,rows,['resname','date','label','cpu','ram','stor']);
    const newHist=rows.map(r=>{
      const res=resources.find(x=>x.name.toLowerCase()===String(r.resname||'').toLowerCase().trim());
      if(!res) return null;
      return {
        id:Date.now()+Math.random(),
        resId:res.id,resName:res.name,
        date:excelDateToString(r.date),
        label:String(r.label||'').trim(),
        cpu:safeFloat(r.cpu),
        ram:safeFloat(r.ram),
        stor:safeFloat(r.stor),
      };
    }).filter(r=>r&&r.date);
    if(!newHist.length) return fmsg(msgId,'✗ Tidak ada data valid. Pastikan resName sesuai nama resource yang sudah ada.','var(--danger)');
    history=[...history,...newHist].sort((a,b)=>String(a.date).localeCompare(String(b.date)));
    await syncData('saveHistory',history);
    renderHistTable();
    fmsg(msgId,`✓ ${newHist.length} data utilisasi berhasil diimport!`,'var(--accent3)');
  }
  else if(type==='projections'){
    renderPreview(previewId,rows,['resname','metric','year','projected']);
    const newProj=rows.map(r=>{
      const res=resources.find(x=>x.name.toLowerCase()===String(r.resname||'').toLowerCase().trim());
      if(!res) return null;
      const projVal=safeFloat(r.projected);
      if(projVal===''||projVal<=0) return null;
      return {
        id:Date.now()+Math.random(),
        resId:res.id,resName:res.name,
        metric:String(r.metric||'cpu').trim(),
        year:parseInt(r.year)||2026,
        projected:projVal
      };
    }).filter(r=>r);
    if(!newProj.length) return fmsg(msgId,'✗ Tidak ada data valid. Pastikan resName sesuai nama resource yang sudah ada dan kolom projected terisi.','var(--danger)');
    newProj.forEach(p=>{
      const idx=projections.findIndex(x=>x.resId==p.resId&&x.metric===p.metric&&x.year==p.year);
      if(idx>=0) projections[idx]=p; else projections.push(p);
    });
    await syncData('saveProjections',projections);
    renderProjTable();
    fmsg(msgId,`✓ ${newProj.length} proyeksi berhasil diimport!`,'var(--accent3)');
  }
}
