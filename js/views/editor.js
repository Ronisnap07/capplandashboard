// ===== RESOURCES =====
function onTypeChange(){
  const t=document.getElementById('f-type').value;
  document.getElementById('grp-cpu').style.display=t==='cluster'?'':'none';
  document.getElementById('grp-ram').style.display=t==='cluster'?'':'none';
  document.getElementById('grp-stor').style.display=t!=='cluster'?'':'none';
}

async function addResource(){
  const t=document.getElementById('f-type').value;
  const name=document.getElementById('f-name').value.trim();
  if(!name) return fmsg('res-msg','✗ Nama wajib diisi.','var(--danger)');
  let r={id:Date.now(),type:t,name,note:document.getElementById('f-note').value};
  if(t==='cluster'){
    const cpu=parseFloat(document.getElementById('f-cpu').value),ram=parseFloat(document.getElementById('f-ram').value);
    if(!cpu||cpu<=0) return fmsg('res-msg','✗ Kapasitas CPU wajib diisi.','var(--danger)');
    if(!ram||ram<=0) return fmsg('res-msg','✗ Kapasitas RAM wajib diisi.','var(--danger)');
    r.cpuCap=cpu; r.ramCap=ram;
  } else {
    const stor=parseFloat(document.getElementById('f-stor').value);
    if(!stor||stor<=0) return fmsg('res-msg','✗ Kapasitas Storage wajib diisi.','var(--danger)');
    r.storCap=stor;
  }
  resources.push(r);
  await syncData('saveResources',resources);
  renderResTable(); clearResForm();
  fmsg('res-msg',`✓ Resource "${name}" berhasil ditambahkan!`,'var(--accent3)');
}

async function delRes(id){
  if(!confirm('Hapus resource ini?')) return;
  resources=resources.filter(r=>r.id!=id);
  history=history.filter(h=>h.resId!=id);
  projections=projections.filter(p=>p.resId!=id);
  await syncData('saveResources',resources);
  await syncData('saveHistory',history);
  await syncData('saveProjections',projections);
  renderResTable();
}

function renderResTable(){
  const tbody=document.getElementById('res-tbody'); if(!tbody) return;
  if(!resources.length){
    tbody.innerHTML=`<tr><td colspan="8" style="text-align:center;padding:24px;color:var(--text-dim)">Belum ada resource.</td></tr>`;
    return;
  }
  const tl={cluster:'Host Cluster',storage_core:'Storage Core',storage_support:'Storage Support'};
  const tc={cluster:'ab-cluster',storage_core:'ab-core',storage_support:'ab-support'};
  tbody.innerHTML=resources.map((r,i)=>`<tr>
    <td style="color:var(--text-dim);font-size:10px">${i+1}</td>
    <td><b>${r.name}</b></td>
    <td><span class="ab ${tc[r.type]}">${tl[r.type]}</span></td>
    <td>${r.cpuCap?r.cpuCap+' cores':'—'}</td>
    <td>${r.ramCap?r.ramCap+' GB':'—'}</td>
    <td>${r.storCap?r.storCap+' TB':'—'}</td>
    <td style="color:var(--text-dim)">${r.note||'—'}</td>
    <td><button class="act-btn del" onclick="delRes(${r.id})">✕</button></td>
  </tr>`).join('');
}

function clearResForm(){
  ['f-name','f-cpu','f-ram','f-stor','f-note'].forEach(id=>{ const el=document.getElementById(id); if(el) el.value=''; });
}

// ===== HISTORY =====
function renderHistorySelects(){
  const sel=document.getElementById('h-res'); if(!sel) return;
  sel.innerHTML=resources.map(r=>`<option value="${r.id}">${r.name}</option>`).join('');
  onHistResChange();
  if(!document.getElementById('h-date').value) document.getElementById('h-date').valueAsDate=new Date();
}

function onHistResChange(){
  const id=document.getElementById('h-res')?.value;
  const r=resources.find(x=>x.id==id); if(!r) return;
  document.getElementById('grp-h-cpu').style.display=r.type==='cluster'?'':'none';
  document.getElementById('grp-h-ram').style.display=r.type==='cluster'?'':'none';
  document.getElementById('grp-h-stor').style.display=r.type!=='cluster'?'':'none';
}

async function addHistory(){
  const resId=document.getElementById('h-res').value;
  const r=resources.find(x=>x.id==resId); if(!r) return fmsg('hist-msg','✗ Pilih resource.','var(--danger)');
  const date=document.getElementById('h-date').value; if(!date) return fmsg('hist-msg','✗ Tanggal wajib diisi.','var(--danger)');
  let entry={id:Date.now(),resId:r.id,resName:r.name,date,label:document.getElementById('h-label').value.trim()};
  if(r.type==='cluster'){
    const cpu=parseFloat(document.getElementById('h-cpu').value),ram=parseFloat(document.getElementById('h-ram').value);
    if(isNaN(cpu)&&isNaN(ram)) return fmsg('hist-msg','✗ Isi minimal CPU atau RAM.','var(--danger)');
    entry.cpu=isNaN(cpu)?'':cpu; entry.ram=isNaN(ram)?'':ram; entry.stor='';
  } else {
    const stor=parseFloat(document.getElementById('h-stor').value);
    if(isNaN(stor)) return fmsg('hist-msg','✗ Isi utilisasi storage.','var(--danger)');
    entry.cpu=''; entry.ram=''; entry.stor=stor;
  }
  history.push(entry);
  history.sort((a,b)=>String(a.date).localeCompare(String(b.date)));
  await syncData('saveHistory',history);
  renderHistTable();
  fmsg('hist-msg','✓ Data berhasil disimpan!','var(--accent3)');
}

function renderHistFilter(){
  const wrap=document.getElementById('hist-filter'); if(!wrap) return;
  wrap.innerHTML=`<button class="btn ${histFilterId===null?'btn-primary':'btn-sec'}" onclick="setHistFilter(null)">Semua</button>`
    +resources.map(r=>`<button class="btn ${histFilterId==r.id?'btn-primary':'btn-sec'}" onclick="setHistFilter(${r.id})">${r.name}</button>`).join('');
}

function setHistFilter(id){ histFilterId=id; renderHistFilter(); renderHistTable(); }

function renderHistTable(){
  const tbody=document.getElementById('hist-tbody'); if(!tbody) return;
  const data=histFilterId!==null?history.filter(h=>h.resId==histFilterId):[...history];
  const sorted=[...data].reverse();
  if(!sorted.length){
    tbody.innerHTML=`<tr><td colspan="8" style="text-align:center;padding:20px;color:var(--text-dim)">Belum ada riwayat.</td></tr>`;
    return;
  }
  tbody.innerHTML=sorted.map((h,i)=>`<tr>
    <td style="color:var(--text-dim);font-size:10px">${i+1}</td>
    <td>${h.resName||'—'}</td>
    <td style="font-family:'Space Mono',monospace;font-size:11px">${h.date}</td>
    <td style="color:var(--text-dim);font-size:11px">${h.label||'—'}</td>
    <td>${h.cpu!==''&&h.cpu!=null?h.cpu+'%':'—'}</td>
    <td>${h.ram!==''&&h.ram!=null?h.ram+'%':'—'}</td>
    <td>${h.stor!==''&&h.stor!=null?h.stor+'%':'—'}</td>
    <td><button class="act-btn del" onclick="delHist(${h.id})">✕</button></td>
  </tr>`).join('');
}

async function delHist(id){
  if(!confirm('Hapus?')) return;
  history=history.filter(h=>h.id!=id);
  await syncData('saveHistory',history);
  renderHistTable();
}

function clearHistForm(){
  ['h-cpu','h-ram','h-stor','h-label'].forEach(id=>{ const el=document.getElementById(id); if(el) el.value=''; });
}

// ===== PROJECTION EDITOR =====
function renderProjUI(){
  const sel=document.getElementById('proj-res-sel'); if(!sel) return;
  sel.innerHTML=resources.map(r=>`<button class="proj-res-btn ${selProjRes==r.id?'active':''}" onclick="selectProjRes(${r.id})">${r.name}</button>`).join('');
  renderProjMetrics();
  renderProjAbsInput();
  renderProjTable();
}

function selectProjRes(id){ selProjRes=id; selProjMetric=null; renderProjUI(); }

function renderProjMetrics(){
  const wrap=document.getElementById('proj-metric-sel'); if(!wrap) return;
  if(!selProjRes){ wrap.innerHTML=''; return; }
  const r=resources.find(x=>x.id==selProjRes);
  const metrics=r&&r.type==='cluster'?[{k:'cpu',l:'CPU (cores)'},{k:'ram',l:'RAM (GB)'}]:[{k:'stor',l:'Storage (TB)'}];
  wrap.innerHTML=metrics.map(m=>`<button class="proj-res-btn ${selProjMetric===m.k?'active':''}" onclick="selectProjMetric('${m.k}')">${m.l}</button>`).join('');
}

function selectProjMetric(k){ selProjMetric=k; renderProjAbsInput(); }

function renderProjAbsInput(){
  const wrap=document.getElementById('proj-abs-input'); if(!wrap) return;
  if(!selProjRes||!selProjMetric){ wrap.style.display='none'; return; }
  const r=resources.find(x=>x.id==selProjRes);
  const year=parseInt(document.getElementById('proj-year').value)||2026;
  const cap=getCapacity(r,selProjMetric);
  const unit=getUnit(selProjMetric);
  const actual=getActualAbsolute(selProjRes,selProjMetric);
  const actualPct=getLatest(selProjRes,selProjMetric);
  const existing=projections.find(p=>p.resId==selProjRes&&p.metric===selProjMetric&&p.year==year);
  const thr=getThr(r);
  const thrVal=selProjMetric==='cpu'?thr.cpu:selProjMetric==='ram'?thr.ram:thr.stor;
  const thrAbs=cap&&thrVal?parseFloat(((thrVal/100)*cap).toFixed(2)):null;

  document.getElementById('proj-abs-label').textContent=`Nilai Proyeksi ${year} (${unit})`;
  document.getElementById('proj-abs-hint').textContent=`Masukkan dalam ${unit}. Kapasitas total: ${cap?cap+' '+unit:'—'}${thrAbs?' | Batas threshold: '+thrAbs+' '+unit:''}`;
  document.getElementById('pm-abs').value=existing?existing.projected:'';
  document.getElementById('pm-abs').placeholder=actual?`Realisasi saat ini: ${actual} ${unit}`:`Contoh: ${cap?Math.round(cap*0.6):'200'} ${unit}`;

  let infoHtml=`<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:8px">`;
  infoHtml+=`<div><div style="font-size:10px;color:var(--text-dim);font-family:'Space Mono',monospace;letter-spacing:1px">KAPASITAS TOTAL</div><div style="font-size:16px;font-family:'Space Mono',monospace;color:var(--accent)">${cap?cap+' '+unit:'—'}</div></div>`;
  infoHtml+=`<div><div style="font-size:10px;color:var(--text-dim);font-family:'Space Mono',monospace;letter-spacing:1px">REALISASI SAAT INI</div><div style="font-size:16px;font-family:'Space Mono',monospace;color:var(--accent3)">${actual?actual+' '+unit:'—'}${actualPct!=null?' ('+actualPct.toFixed(1)+'%)':''}</div></div>`;
  infoHtml+=`<div><div style="font-size:10px;color:var(--text-dim);font-family:'Space Mono',monospace;letter-spacing:1px">BATAS THRESHOLD</div><div style="font-size:16px;font-family:'Space Mono',monospace;color:var(--warn)">${thrAbs?thrAbs+' '+unit+' ('+thrVal+'%)':'—'}</div></div>`;
  if(existing){ infoHtml+=`<div><div style="font-size:10px;color:var(--text-dim);font-family:'Space Mono',monospace;letter-spacing:1px">PROYEKSI ${year} TERSIMPAN</div><div style="font-size:16px;font-family:'Space Mono',monospace;color:var(--purple)">${existing.projected} ${unit}</div></div>`; }
  infoHtml+=`</div>`;
  document.getElementById('proj-current-info').innerHTML=infoHtml;
  wrap.style.display='block';
}

function renderProjTable(){
  const tbody=document.getElementById('proj-tbody'); if(!tbody) return;
  if(!projections.length){
    tbody.innerHTML=`<tr><td colspan="10" style="text-align:center;padding:20px;color:var(--text-dim)">Belum ada proyeksi.</td></tr>`;
    return;
  }
  const rows=projections.map((p,i)=>{
    const r=resources.find(x=>x.id==p.resId);
    const cap=r?getCapacity(r,p.metric):null;
    const unit=getUnit(p.metric);
    const actual=getActualAbsolute(p.resId,p.metric);
    const thr=r?getThr(r):{};
    const thrVal=p.metric==='cpu'?thr.cpu:p.metric==='ram'?thr.ram:thr.stor;
    const thrAbs=cap&&thrVal?((thrVal/100)*cap):null;
    const projPct=cap&&p.projected?((p.projected/cap)*100).toFixed(1):null;
    let acc='—'; let accColor='var(--text-dim)';
    if(actual!=null&&p.projected){
      const err=Math.abs(actual-p.projected)/p.projected*100;
      acc=err.toFixed(1)+'%';
      accColor=err<=10?'var(--accent3)':err<=20?'var(--warn)':'var(--danger)';
    }
    let warnHtml='';
    if(actual!=null&&thrAbs!=null){
      const pctToThr=(actual/thrAbs)*100;
      if(pctToThr>=85&&pctToThr<100) warnHtml=`<span class="ab ab-warn" style="margin-left:4px">⚠ Mendekati</span>`;
      else if(pctToThr>=100) warnHtml=`<span class="ab ab-danger" style="margin-left:4px">✕ Melebihi</span>`;
    }
    return `<tr>
      <td style="color:var(--text-dim);font-size:10px">${i+1}</td>
      <td><b>${p.resName||'—'}</b></td>
      <td><span style="font-family:'Space Mono',monospace;font-size:11px">${p.metric.toUpperCase()}</span></td>
      <td style="font-family:'Space Mono',monospace">${p.year}</td>
      <td style="font-family:'Space Mono',monospace;color:var(--purple)">${p.projected} ${unit}</td>
      <td style="font-family:'Space Mono',monospace;color:var(--text-dim)">${cap?cap+' '+unit:'—'}</td>
      <td style="font-family:'Space Mono',monospace">${projPct?projPct+'%':'—'}</td>
      <td style="font-family:'Space Mono',monospace;color:var(--accent3)">${actual!=null?actual+' '+unit:'—'}${warnHtml}</td>
      <td style="font-family:'Space Mono',monospace;color:${accColor}">${acc}</td>
      <td><button class="act-btn del" onclick="delProj(${p.id})">✕</button></td>
    </tr>`;
  }).join('');
  tbody.innerHTML=rows;
}

async function saveProjection(){
  if(!selProjRes||!selProjMetric) return fmsg('proj-msg','✗ Pilih resource dan metrik.','var(--danger)');
  const year=parseInt(document.getElementById('proj-year').value)||2026;
  const absVal=parseFloat(document.getElementById('pm-abs').value);
  if(isNaN(absVal)||absVal<=0) return fmsg('proj-msg','✗ Nilai proyeksi harus lebih dari 0.','var(--danger)');
  const r=resources.find(x=>x.id==selProjRes);
  const cap=getCapacity(r,selProjMetric);
  if(cap&&absVal>cap) return fmsg('proj-msg',`✗ Nilai proyeksi (${absVal}) melebihi kapasitas total (${cap}).`,'var(--danger)');
  const idx=projections.findIndex(p=>p.resId==selProjRes&&p.metric===selProjMetric&&p.year==year);
  const entry={id:idx>=0?projections[idx].id:Date.now(),resId:selProjRes,resName:r?r.name:'',metric:selProjMetric,year,projected:absVal};
  if(idx>=0) projections[idx]=entry; else projections.push(entry);
  await syncData('saveProjections',projections);
  fmsg('proj-msg','✓ Proyeksi berhasil disimpan!','var(--accent3)');
  renderProjTable(); renderProjAbsInput();
}

function clearProjInputs(){ const el=document.getElementById('pm-abs'); if(el) el.value=''; }

async function delProj(id){
  if(!confirm('Hapus proyeksi ini?')) return;
  projections=projections.filter(p=>p.id!=id);
  await syncData('saveProjections',projections);
  renderProjTable();
}

// ===== THRESHOLD =====
const THR_META=[
  {key:'cluster_cpu',   label:'CPU (Host Cluster)',  color:'var(--orange)'},
  {key:'cluster_ram',   label:'RAM (Host Cluster)',  color:'var(--purple)'},
  {key:'storage_core',  label:'Storage Core',        color:'var(--accent)'},
  {key:'storage_support',label:'Storage Support',    color:'var(--accent3)'}
];

function renderThresholds(){
  const g=document.getElementById('thr-grid'); if(!g) return;
  g.innerHTML=THR_META.map(m=>`<div class="thr-item">
    <label>${m.label}</label>
    <div class="thr-val">
      <input type="number" id="thr-${m.key}" value="${thresholds[m.key]}" min="1" max="100">
      <span class="thr-badge" id="thr-b-${m.key}" style="color:${m.color};background:${m.color}22;border:1px solid ${m.color}44">${thresholds[m.key]}%</span>
    </div>
  </div>`).join('');
  THR_META.forEach(m=>{
    document.getElementById('thr-'+m.key).addEventListener('input',function(){
      document.getElementById('thr-b-'+m.key).textContent=this.value+'%';
    });
  });
}

async function saveThresholds(){
  THR_META.forEach(m=>{ thresholds[m.key]=parseFloat(document.getElementById('thr-'+m.key).value)||DEFAULT_THR[m.key]; });
  const data=Object.entries(thresholds).map(([key,value])=>({key,value}));
  await syncData('saveThresholds',data);
  renderThresholds();
  fmsg('thr-msg','✓ Threshold disimpan!','var(--accent3)');
}

function resetThresholds(){
  thresholds={...DEFAULT_THR};
  const data=Object.entries(thresholds).map(([key,value])=>({key,value}));
  syncData('saveThresholds',data);
  renderThresholds();
  fmsg('thr-msg','✓ Direset ke default.','var(--warn)');
}
