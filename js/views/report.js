const RPT_MONTHS = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
let _rptLastRows = [], _rptLastYearMonth = '';

function renderReport(){
  const wrap = document.getElementById('report-content'); if(!wrap) return;
  const now  = new Date();
  const yearOpts = [];
  for(let y = 2024; y <= 2030; y++) yearOpts.push(y);

  wrap.innerHTML = `
    <div class="card" style="margin-bottom:22px">
      <div style="display:flex;gap:12px;align-items:flex-end;flex-wrap:wrap">
        <div class="fg" style="flex:0 0 auto">
          <label>Tahun</label>
          <select id="rpt-year" style="width:110px">
            ${yearOpts.map(y=>`<option value="${y}" ${y===now.getFullYear()?'selected':''}>${y}</option>`).join('')}
          </select>
        </div>
        <div class="fg" style="flex:0 0 auto">
          <label>Bulan</label>
          <select id="rpt-month" style="width:160px">
            ${RPT_MONTHS.map((n,i)=>`<option value="${i+1}" ${i+1===now.getMonth()+1?'selected':''}>${n}</option>`).join('')}
          </select>
        </div>
        <button class="btn btn-primary" onclick="generateReport()">⚡ Generate</button>
        <button class="btn btn-sec rpt-no-print" onclick="exportReportExcel()">⬇ Export Excel</button>
        <button class="btn btn-sec rpt-no-print" onclick="exportReportPdf()">⬇ Export PDF</button>
      </div>
    </div>
    <div id="report-result"></div>`;
}

function generateReport(){
  const year     = parseInt(document.getElementById('rpt-year').value);
  const month    = parseInt(document.getElementById('rpt-month').value);
  const yearMonth = `${year}-${String(month).padStart(2,'0')}`;
  const prevYM   = rptPrevMonth(yearMonth);
  const label    = `${RPT_MONTHS[month-1]} ${year}`;

  const wrap = document.getElementById('report-result'); if(!wrap) return;
  if(!resources.length){ wrap.innerHTML=`<div class="empty-state"><p>Belum ada resource.</p></div>`; return; }

  // ── Kumpulkan baris data ──────────────────────────────────────────
  const rows = [];
  resources.forEach(r=>{
    const thr = getThr(r);
    const metrics = r.type==='cluster'
      ? [{m:'cpu',l:'CPU',t:thr.cpu},{m:'ram',l:'RAM',t:thr.ram}]
      : [{m:'stor',l:'Storage',t:thr.stor}];

    metrics.forEach(({m,l,t})=>{
      const cap  = getCapacity(r,m);
      const unit = getUnit(m);

      // Rata-rata utilisasi bulan ini & bulan lalu
      const avgUtil  = rptMonthAvg(r.id, m, yearMonth);
      const prevUtil = rptMonthAvg(r.id, m, prevYM);
      const delta    = avgUtil!=null && prevUtil!=null ? avgUtil-prevUtil : null;

      // Proyeksi linear bulan ini
      const proj = projections.find(p=>p.resId==r.id&&p.metric===m&&p.year==year);
      let projMonthAbs=null, projMonthPct=null;
      if(proj && cap){
        const startAbs = rptYearStartAbs(r.id, m, year);
        const mps = calcMonthlyProjections(startAbs||0, proj.projected, year);
        const mp  = mps.find(x=>x.month===month);
        if(mp){
          projMonthAbs = mp.projected;
          projMonthPct = ((projMonthAbs/cap)*100).toFixed(1);
        }
      }

      // Realisasi absolut bulan ini
      const actualAbs = avgUtil!=null && cap ? parseFloat(((avgUtil/100)*cap).toFixed(2)) : null;

      // Akurasi
      let accuracy=null, accColor='var(--text-dim)';
      if(actualAbs!=null && projMonthAbs!=null && projMonthAbs>0){
        const err = Math.abs(actualAbs-projMonthAbs)/projMonthAbs*100;
        accuracy  = Math.max(0, 100-err);
        accColor  = accuracy>=90?'var(--accent3)':accuracy>=80?'var(--warn)':'var(--danger)';
      }

      // Growth dari semua histori
      const allE   = history.filter(h=>h.resId==r.id&&h[m]!==''&&h[m]!=null)
                            .sort((a,b)=>String(a.date).localeCompare(String(b.date)));
      const allV   = allE.map(h=>parseFloat(h[m]));
      const growth = calcGrowth(allV);
      const eta    = t ? calcETA(allV,t) : null;
      const status = statusOf(avgUtil, t);

      rows.push({r, m, l, t, unit, cap, avgUtil, delta, projMonthAbs, projMonthPct,
                 actualAbs, accuracy, accColor, status, growth, eta, allE});
    });
  });

  _rptLastRows = rows; _rptLastYearMonth = yearMonth;

  // ── Summary ───────────────────────────────────────────────────────
  const withData = rows.filter(x=>x.avgUtil!=null).length;
  const warn     = rows.filter(x=>x.status==='warn').length;
  const crit     = rows.filter(x=>x.status==='danger'||x.status==='over').length;
  const accRows  = rows.filter(x=>x.accuracy!=null);
  const avgAcc   = accRows.length ? accRows.reduce((s,x)=>s+x.accuracy,0)/accRows.length : null;
  const avgAccC  = avgAcc==null?'var(--text-dim)':avgAcc>=90?'var(--accent3)':avgAcc>=80?'var(--warn)':'var(--danger)';

  // ── ETA urgency ───────────────────────────────────────────────────
  const urgentRows = rows.map(x=>{
    const em = rptEtaMonths(x.eta);
    if(em===null) return null;
    const urgency = em===0?'over':em<=6?'kritis':em<=12?'waspada':em<=18?'perhatian':null;
    if(!urgency) return null;
    return {...x, etaMonths:em, urgency};
  }).filter(Boolean);

  let html = `
    <div class="sec-title">Laporan Bulanan — ${label}</div>
    <div class="sum-grid" style="margin-bottom:22px">
      ${[
        {l:'Periode',   v:label,              c:'var(--accent)',  s:'',              fs:'13px'},
        {l:'Dipantau',  v:rows.length,         c:'',              s:`${withData} ada data`},
        {l:'Warning',   v:warn,                c:'var(--warn)',   s:'metrik'},
        {l:'Kritis/Over',v:crit,               c:'var(--danger)', s:'metrik'},
        {l:'Rata-rata Akurasi', v:avgAcc!=null?avgAcc.toFixed(1)+'%':'—', c:avgAccC, s:`${accRows.length} terukur`},
        {l:'Perlu Tindak Lanjut', v:urgentRows.length, c:urgentRows.length>0?'var(--danger)':'var(--accent3)', s:'ETA ≤ 18 periode'},
      ].map(x=>`<div class="sum-card">
        <div class="sum-label">${x.l}</div>
        <div class="sum-val" style="color:${x.c||'var(--text-bright)'};${x.fs?'font-size:'+x.fs:''}">${x.v}</div>
        <div class="sum-sub">${x.s}</div>
      </div>`).join('')}
    </div>

    <!-- Tabel Utama -->
    <div class="tbl-wrap" style="margin-bottom:28px">
      <div style="overflow-x:auto">
        <table id="rpt-main-table">
          <thead><tr>
            <th>#</th><th>Resource</th><th>Tipe</th><th>Metrik</th>
            <th>P90 Utilisasi</th><th>vs Bln Lalu</th>
            <th>Growth Bln</th>
            <th>Proyeksi Bln Ini</th><th>Realisasi</th><th>Selisih</th>
            <th>Akurasi</th><th>ETA Threshold</th><th>Status</th>
          </tr></thead>
          <tbody>
          ${rows.map((x,i)=>{
            const uC  = SC[x.status];
            const dC  = x.delta==null?'var(--text-dim)':x.delta>0?'var(--danger)':x.delta<0?'var(--accent3)':'var(--text-dim)';
            const gmo = fmtG(x.growth.monthly);
            const diff = x.actualAbs!=null&&x.projMonthAbs!=null ? x.actualAbs-x.projMonthAbs : null;
            const dfC  = diff==null?'var(--text-dim)':diff>0?'var(--danger)':'var(--accent3)';
            const TC   = {cluster:'ab-cluster',storage_core:'ab-core',storage_support:'ab-support'};
            const TL   = {cluster:'Cluster',storage_core:'Core',storage_support:'Support'};
            const monthEntries = x.allE.filter(h=>String(h.date).startsWith(yearMonth));
            const weeklyHtml = rptWeeklyRows(monthEntries, x.m, x.unit, x.cap);
            return '<tr style="cursor:pointer" onclick="rptToggleWeekly(\'rpt-wr-'+i+'\')">'
              +'<td style="color:var(--text-dim);font-size:10px">'+String(i+1)+'</td>'
              +'<td><b>'+x.r.name+'</b> <span style="font-size:9px;color:var(--accent)">▶</span></td>'
              +'<td><span class="ab '+TC[x.r.type]+'">'+TL[x.r.type]+'</span></td>'
              +'<td style="font-family:\'Space Mono\',monospace;font-size:11px">'+x.l+'</td>'
              +'<td style="font-family:\'Space Mono\',monospace;color:'+uC+'">'+(x.avgUtil!=null?x.avgUtil.toFixed(1)+'%':'—')+'</td>'
              +'<td style="font-family:\'Space Mono\',monospace;font-size:11px;color:'+dC+'">'+(x.delta!=null?(x.delta>0?'+':'')+x.delta.toFixed(2)+'%':'—')+'</td>'
              +'<td style="font-family:\'Space Mono\',monospace;font-size:11px;color:'+gmo.c+'">'+gmo.t+'</td>'
              +'<td style="font-family:\'Space Mono\',monospace;font-size:11px;color:var(--purple)">'+(x.projMonthAbs!=null?x.projMonthAbs+' '+x.unit+(x.projMonthPct?' ('+x.projMonthPct+'%)':''):'—')+'</td>'
              +'<td style="font-family:\'Space Mono\',monospace;font-size:11px;color:var(--accent3)">'+(x.actualAbs!=null?x.actualAbs+' '+x.unit:'—')+'</td>'
              +'<td style="font-family:\'Space Mono\',monospace;font-size:11px;color:'+dfC+'">'+(diff!=null?(diff>0?'+':'')+diff.toFixed(2)+' '+x.unit:'—')+'</td>'
              +'<td>'+(x.accuracy!=null
                ?'<div style="display:flex;align-items:center;gap:6px"><div style="width:48px;height:4px;background:var(--bg3);border-radius:2px;overflow:hidden"><div style="height:100%;width:'+x.accuracy+'%;background:'+x.accColor+';border-radius:2px"></div></div><span style="font-family:\'Space Mono\',monospace;font-size:10px;color:'+x.accColor+'">'+x.accuracy.toFixed(1)+'%</span></div>'
                :'<span style="font-size:10px;color:var(--text-dim)">—</span>')+'</td>'
              +'<td style="font-size:11px;color:var(--warn)">'+(x.eta||'—')+'</td>'
              +'<td><span class="ab '+SAC[x.status]+'">'+SL[x.status]+'</span></td>'
              +'</tr>'
              +'<tr id="rpt-wr-'+i+'" style="display:none">'
              +'<td colspan="13" style="padding:0;border-top:none">'
              +weeklyHtml
              +'</td></tr>';
          }).join('')}
          </tbody>
        </table>
      </div>
    </div>

    <!-- Rekomendasi Tindak Lanjut -->
    ${rptBuildUrgentHtml(urgentRows)}

    <!-- Detail Pertumbuhan per Resource -->
    <div class="sec-title">Tren Utilisasi — ${label}</div>
    <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(340px,1fr));gap:16px;margin-bottom:8px" id="rpt-charts-grid">
    ${resources.filter(r=>{
      const ms = r.type==='cluster'?['cpu','ram']:['stor'];
      return ms.some(m=>history.some(h=>h.resId==r.id&&h[m]!==''&&h[m]!=null));
    }).map(r=>{
      const metrics = r.type==='cluster'
        ?[{m:'cpu',l:'CPU',c:'#fb923c'},{m:'ram',l:'RAM',c:'#c084fc'}]
        :[{m:'stor',l:'Storage',c:'#00d4ff'}];
      return `<div class="card" style="padding:16px">
        <div style="font-family:'Space Mono',monospace;font-size:11px;color:var(--text-bright);margin-bottom:12px;font-weight:700">${r.name}</div>
        ${metrics.map(({m,l,c})=>{
          const e = history.filter(h=>h.resId==r.id&&h[m]!==''&&h[m]!=null)
                          .sort((a,b)=>String(a.date).localeCompare(String(b.date))).slice(-10);
          if(!e.length) return '';
          return `<div style="font-size:9px;color:${c};font-family:'Space Mono',monospace;letter-spacing:1px;margin-bottom:4px">${l}</div>
                  <canvas id="rptc-${r.id}-${m}" height="80" style="margin-bottom:10px"></canvas>`;
        }).join('')}
      </div>`;
    }).join('')}
    </div>`;

  wrap.innerHTML = html;

  // Draw tren charts
  setTimeout(()=>{
    resources.forEach(r=>{
      const metrics = r.type==='cluster'
        ?[{m:'cpu',l:'CPU',c:'#fb923c'},{m:'ram',l:'RAM',c:'#c084fc'}]
        :[{m:'stor',l:'Storage',c:'#00d4ff'}];
      const thr = getThr(r);
      metrics.forEach(({m,l,c})=>{
        const cEl = document.getElementById(`rptc-${r.id}-${m}`); if(!cEl) return;
        const entries = history.filter(h=>h.resId==r.id&&h[m]!==''&&h[m]!=null)
          .sort((a,b)=>String(a.date).localeCompare(String(b.date))).slice(-10);
        if(entries.length < 2) return;
        const thrVal = m==='cpu'?thr.cpu:m==='ram'?thr.ram:thr.stor;
        new Chart(cEl, {
          type:'line',
          data:{
            labels: entries.map(h=>h.label||h.date),
            datasets:[{
              data: entries.map(h=>parseFloat(h[m])),
              borderColor:c, backgroundColor:c+'22',
              borderWidth:2, tension:.4, pointRadius:3, fill:true
            }]
          },
          options:{
            responsive:true,
            plugins:{legend:{display:false}},
            scales:{
              x:{ticks:{color:'#64748b',font:{size:9}},grid:{color:'#1e2d45'}},
              y:{ticks:{color:'#64748b',font:{size:9}},grid:{color:'#1e2d45'},min:0,max:110}
            }
          },
          plugins:[{id:'thr',afterDraw(chart){
            if(!thrVal) return;
            const{ctx,chartArea:{left,right,top,bottom},scales:{y}}=chart;
            const yp=y.getPixelForValue(thrVal); if(yp<top||yp>bottom) return;
            ctx.save();ctx.setLineDash([4,3]);ctx.strokeStyle='rgba(255,204,0,0.5)';
            ctx.lineWidth=1;ctx.beginPath();ctx.moveTo(left,yp);ctx.lineTo(right,yp);ctx.stroke();ctx.restore();
          }}]
        });
      });
    });
  }, 80);
}

// ───── Helpers ──────────────────────────────────────────────────────
function rptMonthAvg(resId, metric, yearMonth){
  const e = history.filter(h=>
    h.resId==resId && h[metric]!==''&&h[metric]!=null &&
    String(h.date).startsWith(yearMonth)
  );
  if(!e.length) return null;
  const sorted = e.map(h=>parseFloat(h[metric])).sort((a,b)=>a-b);
  const idx = Math.ceil(sorted.length * 0.9) - 1;
  return sorted[Math.max(0, idx)];
}

function rptPrevMonth(yearMonth){
  const [y,m] = yearMonth.split('-').map(Number);
  const d = new Date(y, m-2, 1);
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
}

function rptYearStartAbs(resId, metric, year){
  const r = resources.find(x=>x.id==resId); if(!r) return null;
  const cap = getCapacity(r, metric); if(!cap) return null;
  const e = history.filter(h=>
    h.resId==resId && h[metric]!==''&&h[metric]!=null &&
    parseInt(String(h.date).substring(0,4)) <= year
  ).sort((a,b)=>String(a.date).localeCompare(String(b.date)));
  if(!e.length) return null;
  return parseFloat(((parseFloat(e[e.length-1][metric])/100)*cap).toFixed(2));
}

function rptToggleWeekly(id){
  const row = document.getElementById(id); if(!row) return;
  const isOpen = row.style.display !== 'none';
  row.style.display = isOpen ? 'none' : 'table-row';
  const mainRow = row.previousElementSibling;
  if(mainRow){
    const arrow = mainRow.querySelector('span[style*="color:var(--accent)"]');
    if(arrow) arrow.textContent = isOpen ? '▶' : '▼';
  }
}

function rptWeeklyRows(entries, metric, unit, cap){
  if(!entries.length) return '<div style="padding:10px 16px;font-size:11px;color:var(--text-dim)">Tidak ada data dalam bulan ini.</div>';
  let rows = '';
  entries.forEach(function(h, idx){
    const val = h[metric]!==''&&h[metric]!=null ? parseFloat(h[metric]) : null;
    const abs = val!=null && cap ? parseFloat(((val/100)*cap).toFixed(2)) : null;
    const uC  = val==null?'var(--text-dim)':val>=90?'var(--danger)':val>=80?'var(--warn)':'var(--accent3)';
    rows += '<tr style="background:var(--bg2)">'
      +'<td style="font-size:10px;color:var(--text-dim);padding-left:32px">'+(idx+1)+'</td>'
      +'<td style="font-family:\'Space Mono\',monospace;font-size:11px;color:var(--accent)">'+(h.label||'—')+'</td>'
      +'<td style="font-family:\'Space Mono\',monospace;font-size:11px;color:var(--text-dim)">'+h.date+'</td>'
      +'<td colspan="10" style="font-family:\'Space Mono\',monospace;font-size:11px;color:'+uC+'">'
      +(val!=null ? val.toFixed(1)+'%'+(abs!=null?' ('+abs+' '+unit+')':'') : '—')
      +'</td></tr>';
  });
  return '<table style="width:100%;border:none"><thead><tr>'
    +'<th style="font-size:10px;padding-left:32px;color:var(--text-dim)">#</th>'
    +'<th style="font-size:10px;color:var(--text-dim)">Label</th>'
    +'<th style="font-size:10px;color:var(--text-dim)">Tanggal</th>'
    +'<th colspan="10" style="font-size:10px;color:var(--text-dim)">Utilisasi</th>'
    +'</tr></thead><tbody>'+rows+'</tbody></table>';
}

function rptBuildUrgentHtml(urgentRows){
  if(!urgentRows.length) return '';
  const UC = {over:'var(--danger)',kritis:'var(--danger)',waspada:'var(--warn)',perhatian:'#fb923c'};
  const UL = {over:'SUDAH MELEWATI THRESHOLD',kritis:'KRITIS — Segera Ditangani',waspada:'WASPADA — Rencanakan Segera',perhatian:'PERHATIAN — Masukkan ke Roadmap'};
  const UB = {over:'#ff3a4422',kritis:'#ff3a4415',waspada:'#ffcc0015',perhatian:'#fb923c15'};
  let cards = '';
  urgentRows.forEach(function(x){
    const uc = UC[x.urgency], ul = UL[x.urgency], ub = UB[x.urgency];
    const recs = rptGetRecs(x.r.type, x.m, x.urgency);
    const liHtml = recs.map(function(r){ return '<li style="font-size:12px;color:var(--text-bright);line-height:1.5">'+r+'</li>'; }).join('');
    cards += '<div style="border:1px solid '+uc+'44;border-left:3px solid '+uc+';background:'+ub+';border-radius:6px;padding:14px 16px;margin-bottom:10px">'
      +'<div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;flex-wrap:wrap">'
      +'<span style="font-family:\'Space Mono\',monospace;font-weight:700;font-size:11px;color:'+uc+'">'+ul+'</span>'
      +'<span style="font-size:10px;background:'+uc+'33;color:'+uc+';padding:2px 8px;border-radius:10px;font-family:\'Space Mono\',monospace">'+x.r.name+' / '+x.l+'</span>'
      +'<span style="font-size:10px;color:var(--text-dim);margin-left:auto">ETA: '+x.eta+'</span>'
      +'</div>'
      +'<ul style="margin:0;padding-left:18px;display:flex;flex-direction:column;gap:5px">'+liHtml+'</ul>'
      +'</div>';
  });
  return '<div class="sec-title" style="color:var(--danger)">&#9888; Rekomendasi Tindak Lanjut</div>'
        +'<div style="margin-bottom:28px">'+cards+'</div>';
}

function rptEtaMonths(etaStr){
  if(!etaStr) return null;
  if(etaStr==='Sudah melewati threshold') return 0;
  const m = String(etaStr).match(/~?(\d+)/);
  return m ? parseInt(m[1]) : null;
}

function rptGetRecs(type, metric, urgency){
  const isOver    = urgency==='over';
  const isKritis  = urgency==='kritis';
  const isWaspada = urgency==='waspada';

  if(metric==='cpu'){
    if(isOver||isKritis) return [
      'Lakukan VM/workload rebalancing dan evaluasi placement segera.',
      'Identifikasi workload dengan CPU usage tertinggi — pertimbangkan optimasi atau migrasi.',
      'Ajukan permintaan penambahan node komputasi ke tim pengadaan (scale-out).',
      'Aktifkan CPU overcommitment policy sementara sambil menunggu hardware baru.',
    ];
    if(isWaspada) return [
      'Mulai proses procurement node baru — estimasi lead time hardware 3–6 bulan.',
      'Lakukan capacity review bersama tim aplikasi: inventarisir workload yang akan onboard.',
      'Evaluasi potensi efisiensi melalui right-sizing VM yang over-provisioned.',
    ];
    return [
      'Masukkan kebutuhan penambahan kapasitas CPU ke roadmap investasi infrastruktur.',
      'Update proyeksi capacity planning tahunan berdasarkan tren pertumbuhan saat ini.',
      'Tingkatkan frekuensi monitoring utilisasi menjadi setiap minggu.',
    ];
  }
  if(metric==='ram'){
    if(isOver||isKritis) return [
      'Audit alokasi memory per VM — identifikasi yang over-provisioned dan lakukan reclaim.',
      'Aktifkan memory balloon driver / transparent huge pages sebagai mitigasi sementara.',
      'Ajukan permintaan penambahan RAM atau node baru ke tim pengadaan segera.',
      'Batasi onboarding workload baru ke cluster ini hingga kapasitas ditambah.',
    ];
    if(isWaspada) return [
      'Mulai proses procurement RAM atau node tambahan — koordinasi dengan vendor.',
      'Review roadmap aplikasi yang akan onboard: estimasi kebutuhan memory-nya.',
      'Optimalkan penggunaan memory dengan audit VM yang jarang aktif.',
    ];
    return [
      'Masukkan kebutuhan upgrade RAM ke roadmap investasi infrastruktur tahunan.',
      'Monitor trend memory usage setiap bulan dan bandingkan dengan proyeksi.',
      'Dokumentasikan estimasi kebutuhan ke tim perencanaan anggaran.',
    ];
  }
  // storage
  if(isOver||isKritis) return [
    'Koordinasi segera dengan vendor untuk ekspansi volume atau pengadaan storage baru.',
    'Lakukan data archiving/tiering: pindahkan data dingin ke storage tier lebih rendah.',
    'Aktifkan deduplication dan kompresi jika belum diaktifkan untuk membebaskan kapasitas.',
    'Audit snapshot dan backup — hapus snapshot lama yang tidak diperlukan.',
  ];
  if(isWaspada) return [
    'Mulai proses tender atau pengadaan storage baru — koordinasi dengan tim pengadaan.',
    'Review kebijakan retensi data: tetapkan dan jalankan auto-deletion data melebihi umur tertentu.',
    'Evaluasi opsi storage tiering atau cloud storage sebagai extension kapasitas.',
  ];
  return [
    'Masukkan kebutuhan ekspansi storage ke roadmap investasi infrastruktur.',
    'Evaluasi pertumbuhan data secara berkala dan update proyeksi tahunan.',
    'Review kebijakan backup dan snapshot agar tidak mengonsumsi kapasitas berlebih.',
  ];
}

// ───── Export Excel ──────────────────────────────────────────────────
function exportReportExcel(){
  if(typeof XLSX==='undefined'){ alert('Library Excel belum siap, coba lagi.'); return; }
  const year  = document.getElementById('rpt-year')?.value;
  const month = document.getElementById('rpt-month')?.value;
  const table = document.getElementById('rpt-main-table');
  if(!table){ alert('Generate report dulu sebelum export.'); return; }

  const wb = XLSX.utils.book_new();

  // Sheet 1 — Ringkasan
  const ws = XLSX.utils.table_to_sheet(table);
  ws['!cols'] = Array(13).fill({wch:18});
  XLSX.utils.book_append_sheet(wb, ws, 'Ringkasan');

  // Sheet 2 — Detail Mingguan
  const detailData = [['Resource','Tipe','Metrik','Tanggal','Label','Utilisasi (%)','Nilai Absolut','Satuan']];
  const TC = {cluster:'Cluster',storage_core:'Storage Core',storage_support:'Storage Support'};
  _rptLastRows.forEach(function(x){
    const monthEntries = x.allE.filter(function(h){ return String(h.date).startsWith(_rptLastYearMonth); });
    monthEntries.forEach(function(h){
      const val = h[x.m]!==''&&h[x.m]!=null ? parseFloat(h[x.m]) : null;
      const abs = val!=null && x.cap ? parseFloat(((val/100)*x.cap).toFixed(2)) : '';
      detailData.push([x.r.name, TC[x.r.type]||x.r.type, x.l, h.date, h.label||'', val!=null?val:'', abs, x.unit]);
    });
  });
  const ws2 = XLSX.utils.aoa_to_sheet(detailData);
  ws2['!cols'] = [{wch:20},{wch:16},{wch:10},{wch:14},{wch:18},{wch:14},{wch:14},{wch:8}];
  XLSX.utils.book_append_sheet(wb, ws2, 'Detail Mingguan');

  XLSX.writeFile(wb, 'report_'+year+'-'+String(month).padStart(2,'0')+'.xlsx');
}

function exportReportPdf(){
  const result = document.getElementById('report-result');
  if(!result||!result.innerHTML.trim()){ alert('Generate report dulu sebelum export PDF.'); return; }
  window.print();
}

window.addEventListener('beforeprint', function(){
  document.querySelectorAll('[id^="rpt-wr-"]').forEach(function(r){
    r.style.display='table-row';
  });
});

window.addEventListener('afterprint', function(){
  document.querySelectorAll('[id^="rpt-wr-"]').forEach(function(r){
    r.style.display='none';
  });
});
