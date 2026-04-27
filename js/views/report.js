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
        <button class="btn btn-sec rpt-no-print" disabled style="opacity:0.4;cursor:not-allowed" title="Sementara dinonaktifkan">⬇ Export Excel</button>
        <button class="btn btn-sec rpt-no-print" onclick="exportReportWord()">⬇ Export Word</button>
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

// ───── Export Word (Multi-Template) ─────────────────────────────────
function exportReportWord(){
  if(!_rptLastRows.length){ alert('Generate report dulu sebelum export Word.'); return; }
  if(typeof JSZip==='undefined'){ alert('Library Word belum siap, coba lagi.'); return; }
  var yearMonth = _rptLastYearMonth;
  var year  = parseInt(document.getElementById('rpt-year').value);
  var month = parseInt(document.getElementById('rpt-month').value);
  var label = RPT_MONTHS[month-1]+' '+year;
  var todayStr = new Date().toLocaleDateString('id-ID',{day:'numeric',month:'long',year:'numeric'});
  _buildMultiTemplateDoc(year, month, label, todayStr, yearMonth)
    .then(function(blob){
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      var _mo=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
      var _ym=yearMonth.split('-'); var _fn='Laporan_Monitoring_Capplan_'+(_mo[parseInt(_ym[1],10)-1]||_ym[1])+'_'+_ym[0];
      a.download = _fn+'.docx';
      document.body.appendChild(a); a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(a.href);
    })
    .catch(function(e){ console.error('Export Word error:',e); alert('Gagal export Word: '+e.message); });
}

function _esc(s){
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function _extractBodyXml(docXml, stripSectPr){
  var s = docXml.indexOf('<w:body>') + '<w:body>'.length;
  var e = docXml.lastIndexOf('</w:body>');
  var body = docXml.substring(s, e);
  if(stripSectPr){
    var sp = body.lastIndexOf('<w:sectPr');
    if(sp >= 0) body = body.substring(0, sp);
  }
  return body;
}

function _extractSectPr(docXml){
  var s = docXml.lastIndexOf('<w:sectPr');
  var e = docXml.lastIndexOf('</w:body>');
  if(s >= 0 && e >= 0 && s < e) return docXml.substring(s, e);
  return '';
}

function _fmtDateShort(dateStr){
  if(!dateStr) return '—';
  var d = new Date(dateStr);
  if(isNaN(d)) return String(dateStr);
  var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var dd = String(d.getUTCDate()).padStart(2,'0');
  return dd+' '+months[d.getUTCMonth()]+' '+d.getUTCFullYear();
}

function _fillPlaceholders(xml, data){
  /* 1. Remove Word spell-check markers and inline bookmarks */
  xml = xml.replace(/<w:proofErr[^>]*\/>/g, '');
  xml = xml.replace(/<w:bookmarkStart[^>]*\/>/g, '');
  xml = xml.replace(/<w:bookmarkEnd[^>]*\/>/g, '');
  /* 1b. Also handle non-self-closing bookmark tags */
  xml = xml.replace(/<w:bookmarkStart[^>]*><\/w:bookmarkStart>/g, '');
  xml = xml.replace(/<w:bookmarkEnd[^>]*><\/w:bookmarkEnd>/g, '');
  /* 2. 4-run split: {{ [run:name] [run:}] [run:}] — }} split across two separate runs */
  xml = xml.replace(
    /\{\{<\/w:t><\/w:r><w:r[^>]*>(?:<w:rPr>[\s\S]*?<\/w:rPr>)?<w:t[^>]*>([^<{}]+)<\/w:t><\/w:r><w:r[^>]*>(?:<w:rPr>[\s\S]*?<\/w:rPr>)?<w:t[^>]*>\}<\/w:t><\/w:r><w:r[^>]*>(?:<w:rPr>[\s\S]*?<\/w:rPr>)?<w:t[^>]*>\}/g,
    function(match, key){
      var val = data[key.trim()];
      return val !== undefined ? _esc(String(val)) : match;
    }
  );
  /* 3. 3-run split: {{ [run:name] [run:}}] */
  xml = xml.replace(
    /\{\{<\/w:t><\/w:r><w:r[^>]*>(?:<w:rPr>[\s\S]*?<\/w:rPr>)?<w:t[^>]*>([^<{}]+)<\/w:t><\/w:r><w:r[^>]*>(?:<w:rPr>[\s\S]*?<\/w:rPr>)?<w:t[^>]*>\}\}/g,
    function(match, key){
      var val = data[key.trim()];
      return val !== undefined ? _esc(String(val)) : match;
    }
  );
  /* 4. Intact {{key}} placeholders */
  xml = xml.replace(/\{\{([^}]+)\}\}/g, function(match, key){
    var val = data[key.trim()];
    return val !== undefined ? _esc(String(val)) : match;
  });
  return xml;
}

/* Find the true end of a <w:tr> element, accounting for nested <w:tr> (from nested tables) */
function _findTrEndDepth(xml, start){
  var depth = 0, pos = start;
  while(pos < xml.length){
    var openA = xml.indexOf('<w:tr ', pos);
    var openB = xml.indexOf('<w:tr>', pos);
    var open  = openA < 0 ? openB : (openB < 0 ? openA : Math.min(openA, openB));
    var close = xml.indexOf('</w:tr>', pos);
    if(close < 0) return -1;
    if(open >= 0 && open < close){ depth++; pos = open + 5; }
    else { depth--; pos = close + 7; if(depth === 0) return pos; }
  }
  return -1;
}

function _expandWeeklyRows(xml, x, monthEntries){
  /* find the <w:tr> that contains ALL weekly placeholders (cols are nested tables within cells) */
  var trStart = -1, pos = 0;
  while(pos < xml.length){
    var openA = xml.indexOf('<w:tr ', pos);
    var openB = xml.indexOf('<w:tr>', pos);
    var idx   = openA < 0 ? openB : (openB < 0 ? openA : Math.min(openA, openB));
    if(idx < 0) break;
    var end = _findTrEndDepth(xml, idx);
    if(end < 0) break;
    var trXml = xml.substring(idx, end);
    if(trXml.indexOf('index') >= 0 && trXml.indexOf('period') >= 0){
      trStart = idx; break;
    }
    pos = end;
  }
  if(trStart < 0) return xml;
  var trEnd = _findTrEndDepth(xml, trStart);
  if(trEnd < 0) return xml;
  var tplRow = xml.substring(trStart, trEnd);

  var rows = '';
  if(monthEntries.length){
    monthEntries.forEach(function(h, wi){
      var val = h[x.m] !== '' && h[x.m] != null ? parseFloat(h[x.m]) : null;
      var abs = val != null && x.cap ? parseFloat(((val/100)*x.cap).toFixed(2))+' '+x.unit : '—';
      var pct = val != null ? val.toFixed(1)+'%' : '—';
      var wSt = val==null?'No Data':val>100?'Over Cap':(x.t&&val>x.t?'Kritis':(x.t&&val>x.t*0.85?'Warning':'Normal'));
      rows += _fillPlaceholders(tplRow, {
        'index': String(wi+1), 'period': h.label||'—', 'date': _fmtDateShort(h.date),
        'capacity_used': abs,
        'capacity_total': x.cap!=null ? x.cap+' '+x.unit : '—',
        'capacity_utilization': pct, 'status': wSt
      });
    });
  } else {
    rows = _fillPlaceholders(tplRow, {
      'index':'—','period':'Tidak ada data','date':'—',
      'capacity_used':'—','capacity_total':'—','capacity_utilization':'—','status':'—'
    });
  }
  return xml.substring(0, trStart) + rows + xml.substring(trEnd);
}

function _fillContentSection(contentDocXml, x, idx, yearMonth, label, globalData){
  var TL  = {cluster:'CLUSTER',storage_core:'STORAGE CORE',storage_support:'STORAGE SUPPORT'};
  var SL2 = {ok:'Normal',warn:'Warning',danger:'Kritis',over:'Over Cap',na:'No Data'};
  var gmo     = fmtG(x.growth.monthly);
  var monthEntries = Array.isArray(yearMonth)
    ? x.allE.filter(function(h){ var d=String(h.date); return yearMonth.some(function(ym){ return d.startsWith(ym); }); })
    : x.allE.filter(function(h){ return String(h.date).startsWith(yearMonth); });
  var capStr  = x.cap!=null ? x.cap+' '+x.unit : '—';
  var actualStr= x.actualAbs!=null ? x.actualAbs+' '+x.unit : '—';
  var projStr  = x.projMonthAbs!=null ? x.projMonthAbs+' '+x.unit : '—';
  var projPct  = x.projMonthPct!=null ? x.projMonthPct+'%' : '—';
  var accStr   = x.accuracy!=null ? x.accuracy.toFixed(1)+'%' : '—';
  var p90Str   = x.avgUtil!=null ? x.avgUtil.toFixed(1)+'%' : '—';
  var deltaStr = x.delta!=null ? (x.delta>0?'+':'')+x.delta.toFixed(2)+'%' : '—';
  var data = {
    'source_name':             x.r.name+' \u2013 Metrik : '+x.l,
    'source_category':         TL[x.r.type]||x.r.type,
    'metric.name':             x.l,
    'capacity.total':          capStr,
    'capacity.used':           actualStr,
    'capacity.used_pct':       p90Str,
    'capacity.projection':     projStr,
    'capacity.projection_pct': projPct,
    'capacity.accuracy':       accStr,
    'capacity.growth_current': gmo.t,
    'capacity.growth_previous':deltaStr,
    'capacity.p90':            p90Str,
    'capacity.status':         SL2[x.status]||x.status,
    'periode_selection':       globalData.periode_selection,
    'sum_of_source':           globalData.sum_of_source,
    'average_acurate_source':  globalData.average_acurate_source,
    'need_to_action':          globalData.need_to_action
  };
  /* Remove the source_category/metric.name subtitle paragraph from template */
  var xml = contentDocXml;
  var scPos = xml.indexOf('source_category');
  if(scPos >= 0){
    var scPA = xml.lastIndexOf('<w:p ', scPos);
    var scPB = xml.lastIndexOf('<w:p>', scPos);
    var scPS = Math.max(scPA, scPB);
    var scPE = xml.indexOf('</w:p>', scPos) + 6;
    if(scPS >= 0 && scPE > scPS) xml = xml.substring(0, scPS) + xml.substring(scPE);
  }
  /* expand weekly rows first (before global placeholder fill) */
  xml = _expandWeeklyRows(xml, x, monthEntries);
  /* fill all remaining placeholders */
  return _fillPlaceholders(xml, data);
}

function _buildMultiTemplateDoc(year, month, label, todayStr, yearMonth, rows, yearMonths){
  rows       = rows       || _rptLastRows;
  yearMonths = yearMonths || [yearMonth];
  var accRows     = rows.filter(function(x){ return x.accuracy!=null; });
  var avgAcc      = accRows.length ? (accRows.reduce(function(s,x){ return s+x.accuracy; },0)/accRows.length).toFixed(1)+'%' : '—';
  var urgentCount = rows.filter(function(x){ var em=rptEtaMonths(x.eta); return em!==null&&em<=18; }).length;
  var globalData  = {
    'periode_selection':      label.toUpperCase(),
    'date_generate':          todayStr,
    'sum_of_source':          String(_rptLastRows.length),
    'average_acurate_source': avgAcc,
    'need_to_action':         String(urgentCount)
  };
  var PAGE_BREAK = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';

  return Promise.all([
    JSZip.loadAsync(TPL_COVER_B64,    {base64:true}),
    JSZip.loadAsync(TPL_LEMPEN_B64,   {base64:true}),
    JSZip.loadAsync(TPL_DAFTARISI_B64,{base64:true}),
    JSZip.loadAsync(TPL_CONTENT_B64,  {base64:true})
  ]).then(function(zips){
    var coverZip   = zips[0];
    var contentZip = zips[3];
    return Promise.all([
      zips[0].file('word/document.xml').async('text'),
      zips[1].file('word/document.xml').async('text'),
      zips[2].file('word/document.xml').async('text'),
      zips[3].file('word/document.xml').async('text'),
      contentZip.file('word/header1.xml').async('text'),
      contentZip.file('word/_rels/document.xml.rels').async('text'),
      coverZip.file('word/_rels/document.xml.rels').async('text'),
      contentZip.file('[Content_Types].xml').async('text'),
      coverZip.file('word/header2.xml').async('text')
    ]).then(function(docs){
      var coverDoc    = docs[0];
      var lempenDoc   = docs[1];
      var daftarDoc   = docs[2];
      var contentDoc  = docs[3];
      var headerXml   = docs[4];
      var contentRels = docs[5];
      var coverRels   = docs[6];
      var contentTypes= docs[7];
      var emptyHdrXml = docs[8];   /* cover.docx header2.xml — already an empty header */

      /* inject periode_selection + date into the letterhead header */
      contentZip.file('word/header1.xml', _fillPlaceholders(headerXml, globalData));

      /* use cover's actual empty header (avoids namespace declaration issues) */
      contentZip.file('word/headerCover.xml', emptyHdrXml);

      /* register headerCover.xml and image rId in content's rels */
      var updatedRels = contentRels.replace('</Relationships>',
        '<Relationship Id="rIdCoverImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>'
        + '<Relationship Id="rIdCoverHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="headerCover.xml"/>'
        + '</Relationships>');
      contentZip.file('word/_rels/document.xml.rels', updatedRels);

      /* register headerCover.xml content type */
      var updatedTypes = contentTypes.replace('</Types>',
        '<Override PartName="/word/headerCover.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
        + '</Types>');
      contentZip.file('[Content_Types].xml', updatedTypes);

      /* find which rId cover uses for its embedded image */
      var coverImgMatch = coverRels.match(/Id="(rId\d+)"[^>]*Target="media\/[^"]+\.(?:png|jpg|jpeg|gif|emf|wmf)"/i);
      var coverImgRId   = coverImgMatch ? coverImgMatch[1] : 'rId7';

      /* extract and fill cover body; remap cover's image rId to rIdCoverImg */
      var rawCoverBody = _extractBodyXml(_fillPlaceholders(coverDoc, globalData), true);
      var coverBody = rawCoverBody
        .replace(new RegExp('r:embed="' + coverImgRId + '"', 'g'), 'r:embed="rIdCoverImg"')
        .replace(new RegExp('r:id="'    + coverImgRId + '"', 'g'), 'r:id="rIdCoverImg"');

      var lempenBody = _extractBodyXml(_fillPlaceholders(lempenDoc, globalData), true);
      var daftarBody = _extractBodyXml(_fillPlaceholders(daftarDoc, globalData), true);

      /* use content.docx sectPr as master (has rId7 = letterhead header) */
      var masterSect = _extractSectPr(contentDoc);

      /* build cover-section sectPr: no header, same page size/margins as content.
         In OOXML, intermediate sectPr MUST be inside <w:p><w:pPr>…</w:pPr></w:p> */
      var pgSzMatch  = masterSect.match(/<w:pgSz\b[^>]*\/>/);
      var pgMarMatch = masterSect.match(/<w:pgMar\b[^>]*\/>/);
      var coverSectPr = '<w:p><w:pPr><w:sectPr>'
        + '<w:headerReference w:type="default" r:id="rIdCoverHdr"/>'
        + '<w:headerReference w:type="first"   r:id="rIdCoverHdr"/>'
        + '<w:headerReference w:type="even"    r:id="rIdCoverHdr"/>'
        + '<w:type w:val="nextPage"/>'
        + (pgSzMatch  ? pgSzMatch[0]  : '')
        + (pgMarMatch ? pgMarMatch[0] : '')
        + '</w:sectPr></w:pPr></w:p>';

      /* split content body: ringkasan+description (once) | resource section (repeats per resource)
         source_name is split across runs by Word spellcheck, so search for 'source_name' text.
         Split point: start of the <w:p> that contains 'source_name' */
      var contentBody = _extractBodyXml(contentDoc, true);
      var ringkasanBody, resourceDetailTpl;
      var snPos = contentBody.indexOf('source_name');
      if(snPos >= 0){
        /* walk back to find the opening <w:p that encloses source_name */
        var pA = contentBody.lastIndexOf('<w:p ', snPos);
        var pB = contentBody.lastIndexOf('<w:p>', snPos);
        var splitAt = Math.max(pA, pB);
        if(splitAt < 0) splitAt = snPos;
        ringkasanBody     = _fillPlaceholders(contentBody.substring(0, splitAt), globalData);
        resourceDetailTpl = contentBody.substring(splitAt);
      } else {
        var detailIdx = contentBody.indexOf('DETAIL KAPASITAS');
        var splitAt2  = detailIdx >= 0 ? contentBody.indexOf('</w:p>', detailIdx) + 6 : 0;
        ringkasanBody     = _fillPlaceholders(contentBody.substring(0, splitAt2), globalData);
        resourceDetailTpl = contentBody.substring(splitAt2);
      }

      /* fill resource detail template once per resource */
      var resourceBodies = rows.map(function(x, i){
        return _fillContentSection(resourceDetailTpl, x, i, yearMonths, label, globalData);
      });

      /* combine: cover(no-header sectPr) | lempen | daftar | ringkasan | detail×N | masterSectPr */
      var combined = coverBody + coverSectPr
        + lempenBody + PAGE_BREAK
        + daftarBody + PAGE_BREAK
        + ringkasanBody
        + (resourceBodies.length ? PAGE_BREAK + resourceBodies.join(PAGE_BREAK) : '')
        + masterSect;

      /* inject missing namespace declarations used by cover body (DrawingML a:, pic:)
         that content.docx root element doesn't already declare */
      var docXml = contentDoc;
      if(docXml.indexOf('xmlns:a=') < 0){
        docXml = docXml.replace('<w:document ',
          '<w:document'
          + ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
          + ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
          + ' ');
      }

      /* replace body using string methods (more reliable than regex on large strings) */
      var bStart = docXml.indexOf('<w:body>') + '<w:body>'.length;
      var bEnd   = docXml.lastIndexOf('</w:body>');
      var finalDocXml = docXml.substring(0, bStart) + combined + docXml.substring(bEnd);
      contentZip.file('word/document.xml', finalDocXml);

      return contentZip.generateAsync({
        type:'blob',
        mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      });
    });
  });
}

/* LEGACY — kept for reference, no longer called */
function _buildWordDoc(logoSrc){
  var year      = parseInt(document.getElementById('rpt-year')?.value);
  var month     = parseInt(document.getElementById('rpt-month')?.value);
  var label     = RPT_MONTHS[month-1]+' '+year;
  var yearMonth = _rptLastYearMonth;
  var now       = new Date();
  var todayStr  = now.toLocaleDateString('id-ID',{day:'numeric',month:'long',year:'numeric'});

  var withData    = _rptLastRows.filter(function(x){ return x.avgUtil!=null; }).length;
  var warn        = _rptLastRows.filter(function(x){ return x.status==='warn'; }).length;
  var crit        = _rptLastRows.filter(function(x){ return x.status==='danger'||x.status==='over'; }).length;
  var accRows     = _rptLastRows.filter(function(x){ return x.accuracy!=null; });
  var avgAcc      = accRows.length?(accRows.reduce(function(s,x){ return s+x.accuracy; },0)/accRows.length).toFixed(1)+'%':'—';
  var urgentCount = _rptLastRows.filter(function(x){ var em=rptEtaMonths(x.eta); return em!==null&&em<=18; }).length;

  var TL  = {cluster:'CLUSTER',storage_core:'STORAGE CORE',storage_support:'STORAGE SUPPORT'};
  var SL2 = {ok:'Normal',warn:'Warning',danger:'Kritis',over:'Over Cap',na:'No Data'};

  /* colours */
  var C_BLUE   = '#003087';
  var C_LBLUE  = '#dce8f8'; /* light blue alternating row */
  var C_DGRAY  = '#2d3a4a'; /* dark header for weekly table */
  var C_LGRAY  = '#f0f4f8'; /* light alternating for weekly */
  var C_WHITE  = '#ffffff';

  var logoTag = logoSrc
    ? '<img src="'+logoSrc+'" style="width:72pt;height:auto;max-height:48pt;display:block;margin:0 auto">'
    : '<strong style="font-size:18pt;color:'+C_BLUE+'">PNM</strong><br><span style="font-size:7pt;color:'+C_BLUE+'">Permodalan Nasional Madani</span>';

  /* letterhead */
  function lh(halaman, docTitle){
    var dt = docTitle||'DOKUMEN MONITORING CAPACITY PLANNING';
    return '<table border="1" style="width:100%;border-collapse:collapse;margin-bottom:12pt;font-size:10pt">'
      +'<tr>'
        +'<td rowspan="4" style="width:80pt;text-align:center;vertical-align:middle;padding:6pt;border:1pt solid #000">'+logoTag+'</td>'
        +'<td colspan="2" style="text-align:right;font-weight:bold;padding:5pt 8pt;font-size:10.5pt;border:1pt solid #000">DIVISI STRATEGI DAN PERENCANAAN<br>TEKNOLOGI INFORMASI</td>'
      +'</tr>'
      +'<tr><td colspan="2" style="text-align:center;font-weight:bold;padding:5pt 8pt;font-size:10pt;border:1pt solid #000">'+dt+'<br>PERIODE '+label.toUpperCase()+'</td></tr>'
      +'<tr>'
        +'<td style="padding:3pt 8pt;font-size:9pt;border:1pt solid #000"><b>Tanggal</b>&nbsp;: '+todayStr+'</td>'
        +'<td style="padding:3pt 8pt;font-size:9pt;border:1pt solid #000"><b>No. Dokumen</b>&nbsp;: &nbsp;-</td>'
      +'</tr>'
      +'<tr>'
        +'<td style="padding:3pt 8pt;font-size:9pt;border:1pt solid #000"><b>Halaman</b>&nbsp;: '+halaman+'</td>'
        +'<td style="padding:3pt 8pt;font-size:9pt;border:1pt solid #000"><b>Revisi</b>&nbsp;: &nbsp;0.0</td>'
      +'</tr>'
    +'</table>';
  }

  var confidential = '<p style="margin-top:28pt;font-size:8pt;color:#666;font-style:italic;border-top:1pt solid #ccc;padding-top:4pt">Dokumen ini bersifat <em>confidential</em>.<br>Proyeksi dan estimasi dapat berubah sesuai kondisi aktual.</p>';
  function pg(){ return '<p style="page-break-after:always">&nbsp;</p>'; }
  /* level 1 = h1 (Heading 1), level 2 = h2 (Heading 2), level 3 = h3 (Heading 3) */
  function sec(t, italic, level){
    var l   = level||1;
    var tag = 'h'+l;
    var fs  = l===1?'13pt':l===2?'12pt':'11pt';
    var inner = italic ? '<em>'+t+'</em>' : t;
    return '<'+tag+' style="font-size:'+fs+';color:'+C_BLUE+';border-bottom:1.5pt solid '+C_BLUE
      +';padding-bottom:3pt;margin:14pt 0 8pt'+(italic?';font-style:italic':'')+'">'+inner+'</'+tag+'>';
  }

  /* ── Resource sections ── */
  var resHtml = '';
  _rptLastRows.forEach(function(x,i){
    var typeLabel    = TL[x.r.type]||x.r.type;
    var monthEntries = x.allE.filter(function(h){ return String(h.date).startsWith(yearMonth); });
    var gmo          = fmtG(x.growth.monthly);
    var actualStr    = x.actualAbs!=null ? x.actualAbs+' '+x.unit : '—';
    var projStr      = x.projMonthAbs!=null ? x.projMonthAbs+' '+x.unit+(x.projMonthPct?' ('+x.projMonthPct+'%)':'') : '—';
    var accStr       = x.accuracy!=null ? x.accuracy.toFixed(1)+'%' : '—';
    var p90Str       = x.avgUtil!=null ? x.avgUtil.toFixed(1)+'%' : '—';
    var capStr       = x.cap!=null ? x.cap+' '+x.unit : '—';
    var deltaStr     = x.delta!=null?(x.delta>0?'+':'')+x.delta.toFixed(2)+'%':'—';
    var statusStr    = SL2[x.status]||x.status;

    /* colour status text */
    var stCol = x.status==='ok'?'#16a34a':x.status==='warn'?'#d97706':x.status==='danger'||x.status==='over'?'#dc2626':'#555';

    /* weekly rows */
    var wRows = '';
    monthEntries.forEach(function(h,wi){
      var val  = h[x.m]!==''&&h[x.m]!=null ? parseFloat(h[x.m]) : null;
      var abs  = val!=null&&x.cap ? parseFloat(((val/100)*x.cap).toFixed(2))+' '+x.unit : '—';
      var pct  = val!=null ? val.toFixed(1)+'%' : '—';
      var wSt  = val==null?'No Data':val>100?'Over Cap':x.t&&val>x.t?'Kritis':x.t&&val>x.t*0.85?'Warning':'Normal';
      var wStC = wSt==='Normal'?'#16a34a':wSt==='Warning'?'#d97706':wSt==='No Data'?'#888':'#dc2626';
      var bg   = wi%2===0 ? C_WHITE : C_LGRAY;
      wRows += '<tr style="background:'+bg+'">'
        +'<td style="text-align:center;padding:4pt;border:1pt solid #c0ccd8">'+(wi+1)+'</td>'
        +'<td style="padding:4pt;border:1pt solid #c0ccd8;color:'+C_BLUE+';font-weight:bold">'+(h.label||'—')+'</td>'
        +'<td style="padding:4pt;border:1pt solid #c0ccd8">'+h.date+'</td>'
        +'<td style="text-align:center;padding:4pt;border:1pt solid #c0ccd8;font-weight:bold">'+abs+'</td>'
        +'<td style="text-align:center;padding:4pt;border:1pt solid #c0ccd8">'+(x.cap!=null?x.cap+' '+x.unit:'—')+'</td>'
        +'<td style="text-align:center;padding:4pt;border:1pt solid #c0ccd8">'+pct+'</td>'
        +'<td style="text-align:center;padding:4pt;border:1pt solid #c0ccd8;color:#1a56db">'+projStr+'</td>'
        +'<td style="text-align:center;padding:4pt;border:1pt solid #c0ccd8;font-weight:bold;color:'+wStC+'">'+wSt+'</td>'
        +'</tr>';
    });

    resHtml += '<h3 style="font-size:11pt;color:'+C_BLUE+';margin:14pt 0 4pt;border-bottom:1pt solid #b8cce4;padding-bottom:2pt">'
      +(i+1)+'. '+x.r.name
      +'&nbsp;&nbsp;<span style="background:'+C_BLUE+';color:#fff;padding:1pt 7pt;font-size:8pt;font-weight:bold">'+typeLabel+'</span>'
      +'&nbsp;<span style="border:1pt solid '+C_BLUE+';color:'+C_BLUE+';padding:1pt 7pt;font-size:8pt">Metrik: '+x.l+'</span>'
      +'</h3>'

      /* metrics summary table */
      +'<table border="1" style="width:100%;border-collapse:collapse;font-size:9pt;margin-bottom:6pt;table-layout:fixed">'
        +'<colgroup>'
          +'<col style="width:13%"><col style="width:15%"><col style="width:15%"><col style="width:11%">'
          +'<col style="width:11%"><col style="width:11%"><col style="width:12%"><col style="width:12%">'
        +'</colgroup>'
        +'<tr style="background:'+C_BLUE+';color:'+C_WHITE+'">'
          +'<th style="padding:5pt 3pt;text-align:center;border:1pt solid #001a5c">Total<br>Kapasitas</th>'
          +'<th style="padding:5pt 3pt;text-align:center;border:1pt solid #001a5c">Realisasi<br>(Used)</th>'
          +'<th style="padding:5pt 3pt;text-align:center;border:1pt solid #001a5c">Proyeksi<br>Bln Ini</th>'
          +'<th style="padding:5pt 3pt;text-align:center;border:1pt solid #001a5c">Akurasi<br>Proyeksi</th>'
          +'<th style="padding:5pt 3pt;text-align:center;border:1pt solid #001a5c">Growth<br>Bln Ini</th>'
          +'<th style="padding:5pt 3pt;text-align:center;border:1pt solid #001a5c">Growth<br>Bln Lalu</th>'
          +'<th style="padding:5pt 3pt;text-align:center;border:1pt solid #001a5c">P90<br>Utilisasi</th>'
          +'<th style="padding:5pt 3pt;text-align:center;border:1pt solid #001a5c">Status</th>'
        +'</tr>'
        +'<tr style="background:'+C_LBLUE+'">'
          +'<td style="text-align:center;padding:5pt 4pt;border:1pt solid #b8cce4">'+capStr+'</td>'
          +'<td style="text-align:center;padding:5pt 4pt;border:1pt solid #b8cce4"><strong style="color:'+C_BLUE+'">'+actualStr+'</strong><br><span style="font-size:8pt;color:#555">('+p90Str+')</span></td>'
          +'<td style="text-align:center;padding:5pt 4pt;border:1pt solid #b8cce4;color:#1a56db;font-weight:bold">'+projStr+'</td>'
          +'<td style="text-align:center;padding:5pt 4pt;border:1pt solid #b8cce4">'+accStr+'</td>'
          +'<td style="text-align:center;padding:5pt 4pt;border:1pt solid #b8cce4">'+gmo.t+'</td>'
          +'<td style="text-align:center;padding:5pt 4pt;border:1pt solid #b8cce4">'+deltaStr+'</td>'
          +'<td style="text-align:center;padding:5pt 4pt;border:1pt solid #b8cce4">'+p90Str+'</td>'
          +'<td style="text-align:center;padding:5pt 4pt;border:1pt solid #b8cce4;font-weight:bold;color:'+stCol+'">'+statusStr+'</td>'
        +'</tr>'
      +'</table>'

      /* weekly breakdown */
      +(monthEntries.length
        ? '<h4 style="font-size:9pt;font-weight:bold;margin:8pt 0 3pt;color:#333">Detail Pemakaian Mingguan:</h4>'
          +'<table border="1" style="width:100%;border-collapse:collapse;font-size:9pt;margin-bottom:4pt;table-layout:fixed">'
            +'<colgroup>'
              +'<col style="width:4%"><col style="width:15%"><col style="width:13%"><col style="width:14%">'
              +'<col style="width:15%"><col style="width:11%"><col style="width:16%"><col style="width:12%">'
            +'</colgroup>'
            +'<tr style="background:'+C_DGRAY+';color:'+C_WHITE+'">'
              +'<th style="padding:4pt;text-align:center;border:1pt solid #1a2533">#</th>'
              +'<th style="padding:4pt;border:1pt solid #1a2533">PERIODE</th>'
              +'<th style="padding:4pt;border:1pt solid #1a2533">TANGGAL</th>'
              +'<th style="padding:4pt;text-align:center;border:1pt solid #1a2533">KAP. TERPAKAI</th>'
              +'<th style="padding:4pt;text-align:center;border:1pt solid #1a2533">TOTAL KAPASITAS</th>'
              +'<th style="padding:4pt;text-align:center;border:1pt solid #1a2533">% UTILISASI</th>'
              +'<th style="padding:4pt;text-align:center;border:1pt solid #1a2533">PROYEKSI BLN INI</th>'
              +'<th style="padding:4pt;text-align:center;border:1pt solid #1a2533">STATUS</th>'
            +'</tr>'
            +wRows
          +'</table>'
        : '<p style="font-size:9pt;color:#888;margin-bottom:4pt">Tidak ada data mingguan pada periode ini.</p>');
  });

  /* ── Assemble full document ── */
  var css = '<style>'
    +'body{font-family:Arial,sans-serif;font-size:11pt;color:#000;margin:0}'
    +'table{border-collapse:collapse;word-wrap:break-word}td,th{vertical-align:middle;overflow-wrap:break-word}'
    +'p{margin:0 0 6pt}'
    +'h1{font-size:13pt;color:'+C_BLUE+';margin:14pt 0 8pt}'
    +'h2{font-size:12pt;color:'+C_BLUE+';margin:12pt 0 6pt}'
    +'h3{font-size:11pt;color:'+C_BLUE+';margin:10pt 0 4pt}'
    +'h4{font-size:9pt;color:#333;margin:8pt 0 3pt}'
    +'ul{margin:6pt 0 6pt 20pt}li{margin-bottom:4pt}'
    +'</style>';

  var body = ''
    /* ── Cover ── */
    +(logoSrc ? '<p style="margin-bottom:8pt"><img src="'+logoSrc+'" style="width:100pt;height:auto"></p>' : '<p style="margin-bottom:8pt"><strong style="font-size:24pt;color:'+C_BLUE+';letter-spacing:2pt">PNM</strong><br><span style="font-size:9pt;color:'+C_BLUE+'">Permodalan Nasional Madani</span></p>')
    +'<hr style="border:0;border-top:3pt solid '+C_BLUE+';margin-bottom:36pt">'
    +'<p style="font-size:24pt;font-weight:bold;line-height:1.3;margin-bottom:14pt">DOKUMEN MONITORING CAPACITY<br>PLANNING</p>'
    +'<p style="font-size:13pt;font-weight:bold;margin-bottom:8pt">PERIODE '+label.toUpperCase()+'</p>'
    +'<p style="font-size:10.5pt;margin-bottom:4pt">BAGIAN STRATEGI TEKNOLOGI INFORMASI</p>'
    +'<p style="font-size:10.5pt;margin-bottom:70pt">DIVISI STRATEGI DAN PERENCANAAN TEKNOLOGI INFORMASI</p>'
    +'<p style="font-size:12pt;font-weight:bold;text-align:center">DIVISI STRATEGI DAN PERENCANAAN TEKNOLOGI INFORMASI</p>'
    +'<p style="font-size:12pt;font-weight:bold;text-align:center;margin-bottom:40pt">PT PERMODALAN NASIONAL MADANI</p>'
    +pg()

    /* ── Lembar Pengesahan (page 2) ── */
    +sec('LEMBAR PENGESAHAN', false, 1)
    +'<table style="width:100%;border:none;margin-top:24pt">'
      +'<tr>'
        +'<td style="width:50%;text-align:center;vertical-align:top;border:none;padding:0 8pt">'
          +'<p style="font-weight:bold;margin-bottom:54pt">DIBUAT OLEH</p>'
          +'<p style="font-weight:bold;text-decoration:underline">Abdur Roni</p>'
          +'<p style="font-size:10pt;margin-top:4pt">Officer \u2013 Divisi Strategi dan Perencanaan<br>Teknologi Informasi</p>'
        +'</td>'
        +'<td style="width:50%;text-align:center;vertical-align:top;border:none;padding:0 8pt">'
          +'<p style="font-weight:bold;margin-bottom:54pt">DIPERIKSA OLEH</p>'
          +'<p style="font-weight:bold;text-decoration:underline">M. Yusup Hamdani</p>'
          +'<p style="font-size:10pt;margin-top:4pt">Pj. Kepala Bagian \u2013 Divisi Strategi dan<br>Perencanaan Teknologi Informasi</p>'
        +'</td>'
      +'</tr>'
    +'</table>'
    +'<table style="width:100%;border:none;margin-top:24pt">'
      +'<tr><td style="text-align:center;border:none">'
        +'<p style="font-weight:bold;margin-bottom:54pt">DISETUJUI OLEH</p>'
        +'<p style="font-weight:bold;text-decoration:underline">Satria Pujakesuma</p>'
        +'<p style="font-size:10pt;margin-top:4pt">Kepala Divisi Strategi dan Perencanaan<br>Teknologi Informasi</p>'
      +'</td></tr>'
    +'</table>'
    +pg()

    /* ── Daftar Isi (page 3) ── */
    +sec('DAFTAR ISI', false, 1)
    +(function(){
      var tocEntries = [
        {label:'Lembar Pengesahan', page:'2'},
        {label:'Daftar Isi', page:'3'},
        {label:'Laporan Bulanan \u2014 '+label, page:'4'},
        {label:'\u00a0\u00a0\u00a0Ringkasan Eksekutif', page:'4'},
        {label:'\u00a0\u00a0\u00a0Detail Kapasitas &amp; Breakdown Mingguan', page:'5'}
      ];
      var rows = tocEntries.map(function(e){
        return '<tr>'
          +'<td style="padding:5pt 4pt;border:none;font-size:10pt">'+e.label+'</td>'
          +'<td style="padding:5pt 4pt;border:none;font-size:10pt;text-align:right;white-space:nowrap;color:#555">'
            +'<span style="letter-spacing:2pt">.................</span>&nbsp;'+e.page
          +'</td>'
        +'</tr>';
      });
      return '<table style="width:100%;border:none;border-collapse:collapse;margin-top:8pt">'+rows.join('')+'</table>';
    })()
    +pg()

    /* ── Laporan Bulanan ── */
    +'<p style="color:'+C_BLUE+';font-size:13pt;font-weight:bold;margin-bottom:3pt">LAPORAN BULANAN \u2014 '+label.toUpperCase()+'</p>'
    +'<p style="font-size:9pt;color:#555;margin-bottom:14pt">Periode: '+label+' | Tahun: '+year+'</p>'

    +sec('Ringkasan Eksekutif', false, 2)
    +'<table border="1" style="width:100%;border-collapse:collapse;font-size:10pt;margin-bottom:14pt;table-layout:fixed">'
      +'<colgroup><col style="width:42%"><col style="width:58%"></colgroup>'
      +'<tr style="background:'+C_BLUE+';color:'+C_WHITE+'">'
        +'<th style="padding:6pt 10pt;text-align:left;border:1pt solid #001a5c">PARAMETER</th>'
        +'<th style="padding:6pt 10pt;text-align:left;border:1pt solid #001a5c">NILAI</th>'
      +'</tr>'
      +'<tr style="background:'+C_WHITE+'"><td style="padding:6pt 10pt;border:1pt solid #b8cce4"><strong>PERIODE</strong></td><td style="padding:6pt 10pt;border:1pt solid #b8cce4">'+label+'</td></tr>'
      +'<tr style="background:'+C_LBLUE+'"><td style="padding:6pt 10pt;border:1pt solid #b8cce4"><strong>DIPANTAU</strong></td><td style="padding:6pt 10pt;border:1pt solid #b8cce4">'+_rptLastRows.length+' ('+withData+' ada data)</td></tr>'
      +'<tr style="background:'+C_WHITE+'"><td style="padding:6pt 10pt;border:1pt solid #b8cce4"><strong>WARNING</strong></td><td style="padding:6pt 10pt;border:1pt solid #b8cce4">'+warn+' metrik</td></tr>'
      +'<tr style="background:'+C_LBLUE+'"><td style="padding:6pt 10pt;border:1pt solid #b8cce4"><strong>KRITIS/OVER</strong></td><td style="padding:6pt 10pt;border:1pt solid #b8cce4">'+crit+' metrik</td></tr>'
      +'<tr style="background:'+C_WHITE+'"><td style="padding:6pt 10pt;border:1pt solid #b8cce4"><strong>RATA-RATA AKURASI</strong></td><td style="padding:6pt 10pt;border:1pt solid #b8cce4">'+avgAcc+' ('+accRows.length+' terukur)</td></tr>'
      +'<tr style="background:'+C_LBLUE+'"><td style="padding:6pt 10pt;border:1pt solid #b8cce4"><strong>PERLU TINDAK LANJUT</strong></td><td style="padding:6pt 10pt;border:1pt solid #b8cce4">'+urgentCount+' (ETA \u2264 18 periode)</td></tr>'
    +'</table>'

    +sec('DETAIL KAPASITAS &amp; BREAKDOWN MINGGUAN PER RESOURCE', false, 2)
    +'<p style="font-size:9pt;color:#555;margin-bottom:10pt">Setiap resource menampilkan informasi lengkap: total kapasitas, realisasi pemakaian, proyeksi, akurasi proyeksi, pertumbuhan bulan ini &amp; bulan lalu, P90 utilisasi \u2014 diikuti detail pemakaian per minggu.</p>'
    +resHtml;

  var full = '<!DOCTYPE html><html><head><meta charset="UTF-8">'+css+'</head><body>'+body+'</body></html>';
  /* top margin increased to 3000 twip (~5.3cm) to accommodate the Word header */
  return htmlDocx.asBlob(full, {orientation:'portrait', margins:{top:3000,right:1418,bottom:1440,left:1418}});
}

// ───── Word Header/Footer Injection (dari template) ──────────────────
function _injectWordHeader(blob, logoSrc){
  var year     = parseInt(document.getElementById('rpt-year')?.value);
  var month    = parseInt(document.getElementById('rpt-month')?.value);
  var label    = RPT_MONTHS[month-1]+' '+year;
  var todayStr = new Date().toLocaleDateString('id-ID',{day:'numeric',month:'long',year:'numeric'});
  /* Load generated DOCX + embedded template in parallel */
  return Promise.all([
    JSZip.loadAsync(blob),
    JSZip.loadAsync(WORD_TEMPLATE_B64, {base64: true})
  ]).then(function(zips){
    var genZip = zips[0], tplZip = zips[1];

    /* Read all needed parts from template */
    return Promise.all([
      tplZip.file('word/header1.xml').async('text'),
      tplZip.file('word/footer1.xml').async('text'),
      tplZip.file('word/footer2.xml').async('text'),
      tplZip.file('word/media/image2.png').async('uint8array'),
      tplZip.file('word/_rels/header1.xml.rels').async('text')
    ]).then(function(parts){
      var hdrXml     = parts[0];
      var ftrXml     = parts[1];
      var ftrFirstXml= parts[2];
      var logoBytes  = parts[3];
      var hdrRels    = parts[4];

      /* ── Substitute dynamic values in header1.xml ── */
      /* 1. Period label */
      hdrXml = hdrXml.replace(/PERIODE\s+\w+ \d{4}/g, 'PERIODE '+label.toUpperCase());
      /* 2. Date — replace "9</w:t>...<w:t...>April 2026" with current date in single run */
      hdrXml = hdrXml.replace(/>:\s+\d+<\/w:t>[\s\S]*?>(?:\w+ )?\d{4}<\/w:t>/,
        '>:  '+todayStr+'</w:t>');

      /* ── Inject files into generated DOCX ── */
      var emptyRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        +'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>';

      genZip.file('word/header1.xml', hdrXml);
      genZip.file('word/footer1.xml', ftrXml);
      genZip.file('word/footer2.xml', ftrFirstXml);
      genZip.file('word/media/image2.png', logoBytes);
      genZip.file('word/_rels/header1.xml.rels', hdrRels);
      genZip.file('word/_rels/footer1.xml.rels', emptyRels);
      genZip.file('word/_rels/footer2.xml.rels', emptyRels);

      /* ── Update [Content_Types].xml ── */
      return genZip.file('[Content_Types].xml').async('text').then(function(ct){
        if(ct.indexOf('Extension="png"')<0){
          ct = ct.replace('</Types>','<Default Extension="png" ContentType="image/png"/></Types>');
        }
        if(ct.indexOf('header1.xml')<0){
          ct = ct.replace('</Types>',
            '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
            +'<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
            +'<Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
            +'</Types>');
        }
        genZip.file('[Content_Types].xml', ct);
        return genZip.file('word/_rels/document.xml.rels').async('text');

      }).then(function(docRels){
        var maxId = 0;
        (docRels.match(/Id="rId(\d+)"/g)||[]).forEach(function(m){
          var n=parseInt(m.match(/\d+/)[0]); if(n>maxId) maxId=n;
        });
        var ridH1='rId'+(maxId+1), ridF1='rId'+(maxId+2), ridF2='rId'+(maxId+3);
        docRels = docRels.replace('</Relationships>',
          '<Relationship Id="'+ridH1+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
          +'<Relationship Id="'+ridF1+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>'
          +'<Relationship Id="'+ridF2+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer2.xml"/>'
          +'</Relationships>');
        genZip.file('word/_rels/document.xml.rels', docRels);
        return genZip.file('word/document.xml').async('text').then(function(d){
          return {d:d, ridH1:ridH1, ridF1:ridF1, ridF2:ridF2};
        });

      }).then(function(o){
        var docXml = o.d;
        /* No titlePg: header1 is default, no "first" header → cover page also gets the header.
           Footer: default + first (footer2 = same confidential text on all pages) */
        var refs = '<w:headerReference w:type="default" r:id="'+o.ridH1+'"/>'
          +'<w:footerReference w:type="default" r:id="'+o.ridF1+'"/>'
          +'<w:footerReference w:type="first" r:id="'+o.ridF2+'"/>'
          +'<w:titlePg/>';
        if(docXml.indexOf('<w:sectPr')>=0){
          docXml = docXml.replace(/<w:sectPr(\s[^>]*)?>/, function(m){ return m+refs; });
        } else {
          docXml = docXml.replace('</w:body>','<w:sectPr>'+refs+'</w:sectPr></w:body>');
        }
        genZip.file('word/document.xml', docXml);
        return genZip.generateAsync({
          type:'blob',
          mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        });
      });
    });
  }).catch(function(e){
    console.warn('Template inject failed, returning plain DOCX:', e);
    return blob;
  });
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
  if(!_rptLastRows.length){ alert('Generate report dulu sebelum export PDF.'); return; }

  const year  = parseInt(document.getElementById('rpt-year')?.value);
  const month = parseInt(document.getElementById('rpt-month')?.value);
  const label = RPT_MONTHS[month-1]+' '+year;
  const yearMonth = _rptLastYearMonth;
  const now = new Date();
  const todayStr = now.toLocaleDateString('id-ID',{day:'numeric',month:'long',year:'numeric'});

  const withData = _rptLastRows.filter(x=>x.avgUtil!=null).length;
  const warn     = _rptLastRows.filter(x=>x.status==='warn').length;
  const crit     = _rptLastRows.filter(x=>x.status==='danger'||x.status==='over').length;
  const accRows  = _rptLastRows.filter(x=>x.accuracy!=null);
  const avgAcc   = accRows.length?(accRows.reduce((s,x)=>s+x.accuracy,0)/accRows.length).toFixed(1)+'%':'—';
  const urgentCount = _rptLastRows.filter(x=>{ const em=rptEtaMonths(x.eta); return em!==null&&em<=18; }).length;

  const TL = {cluster:'CLUSTER',storage_core:'STORAGE CORE',storage_support:'STORAGE SUPPORT'};
  const SL2 = {ok:'Normal',warn:'Warning',danger:'Kritis',over:'Over Cap',na:'No Data'};

  function makeHeader(halaman, docTitle){
    docTitle = docTitle||'DOKUMEN MONITORING CAPACITY PLANNING';
    return `<table class="lh">
      <tr>
        <td rowspan="4" class="lh-logo"><b style="font-size:20px;color:#003087;letter-spacing:1px">PNM</b><br><span style="font-size:7px;color:#003087">Permodalan Nasional Madani</span></td>
        <td colspan="2" class="lh-div">DIVISI STRATEGI DAN PERENCANAAN<br>TEKNOLOGI INFORMASI</td>
      </tr>
      <tr><td colspan="2" class="lh-doctitle">${docTitle}<br>PERIODE ${label.toUpperCase()}</td></tr>
      <tr><td class="lh-inf"><b>Tanggal</b> : ${todayStr}</td><td class="lh-inf"><b>No. Dokumen</b> : &nbsp;-</td></tr>
      <tr><td class="lh-inf"><b>Halaman</b> : ${halaman}</td><td class="lh-inf"><b>Revisi</b> : &nbsp;0.0</td></tr>
    </table>`;
  }

  let resourcesHtml = '';
  _rptLastRows.forEach(function(x,i){
    const typeLabel = TL[x.r.type]||x.r.type;
    const monthEntries = x.allE.filter(h=>String(h.date).startsWith(yearMonth));
    const gmo = fmtG(x.growth.monthly);
    const actualStr  = x.actualAbs!=null ? x.actualAbs+' '+x.unit : '—';
    const projStr    = x.projMonthAbs!=null ? x.projMonthAbs+' '+x.unit+(x.projMonthPct?' ('+x.projMonthPct+'%)':'') : '—';
    const accStr     = x.accuracy!=null ? x.accuracy.toFixed(1)+'%' : '—';
    const p90Str     = x.avgUtil!=null ? x.avgUtil.toFixed(1)+'%' : '—';
    const capStr     = x.cap!=null ? x.cap+' '+x.unit : '—';
    const deltaStr   = x.delta!=null?(x.delta>0?'+':'')+x.delta.toFixed(2)+'%':'—';
    const statusStr  = SL2[x.status]||x.status;

    let weeklyRows = '';
    monthEntries.forEach(function(h, wi){
      const val = h[x.m]!==''&&h[x.m]!=null ? parseFloat(h[x.m]) : null;
      const abs = val!=null&&x.cap ? parseFloat(((val/100)*x.cap).toFixed(2))+' '+x.unit : '—';
      const pct = val!=null ? val.toFixed(1)+'%' : '—';
      const wSt = val==null?'No Data':val>100?'Over Cap':x.t&&val>x.t?'Kritis':x.t&&val>x.t*0.85?'Warning':'Normal';
      weeklyRows += '<tr>'
        +'<td>'+(wi+1)+'</td>'
        +'<td style="color:#003087">'+(h.label||'—')+'</td>'
        +'<td>'+(h.date||'—')+'</td>'
        +'<td><b>'+abs+'</b></td>'
        +'<td>'+(x.cap!=null?x.cap+' '+x.unit:'—')+'</td>'
        +'<td>'+pct+'</td>'
        +'<td style="color:#1a56db">'+projStr+'</td>'
        +'<td><b>'+wSt+'</b></td>'
        +'</tr>';
    });

    resourcesHtml += '<div class="res-section">'
      +'<div class="res-hdr">'
        +'<span class="res-num">'+(i+1)+'.</span>'
        +'<span class="res-name">'+x.r.name+'</span>'
        +'<span class="badge badge-type">'+typeLabel+'</span>'
        +'<span class="badge badge-metric">Metrik: '+x.l+'</span>'
      +'</div>'
      +'<table class="mt"><thead><tr>'
        +'<th>Total<br>Kapasitas</th><th>Realisasi<br>(Used)</th><th>Proyeksi Bln<br>Ini</th>'
        +'<th>Akurasi<br>Proyeksi</th><th>Growth Bln<br>Ini</th><th>Growth Bln<br>Lalu</th>'
        +'<th>P90 Utilisasi</th><th>Status</th>'
      +'</tr></thead><tbody><tr>'
        +'<td>'+capStr+'</td>'
        +'<td><b>'+actualStr+'</b><br><span class="sub">('+p90Str+')</span></td>'
        +'<td style="color:#1a56db"><b>'+projStr+'</b></td>'
        +'<td>'+accStr+'</td>'
        +'<td>'+gmo.t+'</td>'
        +'<td>'+deltaStr+'</td>'
        +'<td>'+p90Str+'</td>'
        +'<td><b>'+statusStr+'</b></td>'
      +'</tr></tbody></table>'
      +(monthEntries.length
        ? '<div class="wk-title">Detail Pemakaian Mingguan:</div>'
          +'<table class="wt"><thead><tr>'
          +'<th>#</th><th>PERIODE</th><th>TANGGAL</th><th>KAP. TERPAKAI</th>'
          +'<th>TOTAL KAPASITAS</th><th>% UTILISASI</th><th>PROYEKSI BLN INI</th><th>STATUS</th>'
          +'</tr></thead><tbody>'+weeklyRows+'</tbody></table>'
        : '<p style="font-size:10px;color:#666;margin:6px 0">Tidak ada data mingguan pada periode ini.</p>')
      +'</div>';
  });

  const html = '<!DOCTYPE html><html lang="id"><head><meta charset="UTF-8">'
    +'<title>Laporan Capacity Planning — '+label+'</title>'
    +'<style>'
    +'*{margin:0;padding:0;box-sizing:border-box}'
    +'body{font-family:Arial,sans-serif;font-size:11px;color:#000;background:#fff}'
    +'.page{width:210mm;min-height:297mm;padding:14mm 18mm 20mm;margin:0 auto;position:relative;page-break-after:always}'
    +'.page:last-child{page-break-after:auto}'
    +'.footer{position:absolute;bottom:8mm;left:18mm;right:18mm;font-size:8px;color:#666;font-style:italic;border-top:1px solid #ccc;padding-top:3px}'
    /* Letterhead */
    +'.lh{width:100%;border-collapse:collapse;margin-bottom:14px;border:1.5px solid #000}'
    +'.lh td{border:1px solid #000;padding:4px 8px;vertical-align:middle}'
    +'.lh-logo{width:85px;text-align:center}'
    +'.lh-div{font-weight:bold;font-size:11px;text-align:right}'
    +'.lh-doctitle{text-align:center;font-weight:bold;font-size:10px}'
    +'.lh-inf{font-size:10px}'
    /* Cover */
    +'.cover{text-align:center;padding-top:60px}'
    +'.cover-line{width:80%;height:3px;background:#003087;margin:6px auto}'
    /* Summary table */
    +'.st{width:100%;border-collapse:collapse;margin-bottom:14px}'
    +'.st th{background:#003087;color:#fff;padding:6px 10px;font-size:11px;border:1px solid #000;text-align:left}'
    +'.st td{padding:6px 10px;font-size:11px;border:1px solid #ccc}'
    +'.st tr:nth-child(even) td{background:#f5f7fa}'
    /* Section title */
    +'.sec{font-size:12px;font-weight:bold;margin:14px 0 8px;padding-bottom:4px;border-bottom:2px solid #003087;color:#003087}'
    +'.sec-italic{font-style:italic}'
    /* Resource */
    +'.res-section{margin-bottom:16px;page-break-inside:avoid}'
    +'.res-hdr{display:flex;align-items:center;gap:8px;padding:5px 8px;background:#eef2ff;border-left:3px solid #003087;margin-bottom:6px}'
    +'.res-num{font-weight:bold;font-size:12px}'
    +'.res-name{font-weight:bold;font-size:12px;color:#003087;flex:1}'
    +'.badge{font-size:9px;padding:2px 8px;border-radius:10px;font-weight:bold;white-space:nowrap}'
    +'.badge-type{background:#003087;color:#fff}'
    +'.badge-metric{background:#dbeafe;color:#003087;border:1px solid #93c5fd}'
    +'.mt{width:100%;border-collapse:collapse;font-size:10px;margin-bottom:6px}'
    +'.mt th{background:#003087;color:#fff;padding:5px 6px;text-align:center;border:1px solid #000;font-size:9px;line-height:1.3}'
    +'.mt td{padding:6px 7px;text-align:center;border:1px solid #ccc;background:#fafafa}'
    +'.sub{font-size:9px;color:#666}'
    +'.wk-title{font-size:10px;font-weight:bold;margin:6px 0 3px;color:#374151}'
    +'.wt{width:100%;border-collapse:collapse;font-size:10px}'
    +'.wt th{background:#374151;color:#fff;padding:4px 6px;text-align:center;border:1px solid #000;font-size:9px}'
    +'.wt td{padding:4px 6px;border:1px solid #ddd;text-align:center}'
    +'.wt tr:nth-child(even) td{background:#f5f5f5}'
    /* Signature */
    +'.sig-grid{display:grid;grid-template-columns:1fr 1fr;gap:40px;margin:36px 0}'
    +'.sig-center{display:flex;justify-content:center;margin-top:16px}'
    +'.sig-box{text-align:center}'
    +'.sig-title{font-weight:bold;font-size:11px;margin-bottom:55px}'
    +'.sig-name{font-weight:bold;font-size:11px;text-decoration:underline}'
    +'.sig-role{font-size:10px;color:#333;margin-top:4px;line-height:1.5}'
    /* Disclaimer */
    +'.dl-text{font-size:11px;line-height:1.7;margin-bottom:10px;text-align:justify}'
    +'.dl-list{margin:8px 0 10px 20px}'
    +'.dl-list li{font-size:11px;line-height:1.7;margin-bottom:4px;text-align:justify}'
    +'@media print{body{-webkit-print-color-adjust:exact;print-color-adjust:exact}.page{margin:0;padding:10mm 15mm 18mm}}'
    +'</style></head><body>'

    /* ── COVER ── */
    +'<div class="page cover">'
    +'<div style="text-align:left;margin-bottom:30px"><b style="font-size:26px;color:#003087;letter-spacing:2px">PNM</b><br><span style="font-size:9px;color:#003087">Permodalan Nasional Madani</span></div>'
    +'<div class="cover-line"></div>'
    +'<div style="margin-top:60px"><div style="font-size:26px;font-weight:bold;line-height:1.3;margin-bottom:20px">DOKUMEN MONITORING CAPACITY<br>PLANNING</div>'
    +'<div style="font-size:14px;font-weight:bold;margin-bottom:6px">PERIODE '+label.toUpperCase()+'</div>'
    +'<div style="font-size:11px;color:#333;margin-bottom:6px">BAGIAN STRATEGI TEKNOLOGI INFORMASI</div>'
    +'<div style="font-size:11px;color:#333">DIVISI STRATEGI DAN PERENCANAAN TEKNOLOGI INFORMASI</div>'
    +'</div>'
    +'<div style="position:absolute;bottom:40mm;left:50%;transform:translateX(-50%);text-align:center;white-space:nowrap">'
    +'<div style="font-size:12px;font-weight:bold">DIVISI STRATEGI DAN PERENCANAAN TEKNOLOGI INFORMASI</div>'
    +'<div style="font-size:12px;font-weight:bold">PT PERMODALAN NASIONAL MADANI</div>'
    +'</div>'
    +'<div style="position:absolute;bottom:12mm;left:18mm;font-size:8px;color:#666;font-style:italic">Dokumen ini bersifat confidential.<br>Proyeksi dan estimasi dapat berubah sesuai kondisi aktual.</div>'
    +'</div>'

    /* ── LEMBAR PENGESAHAN ── */
    +'<div class="page">'
    +makeHeader('2')
    +'<div class="sec">LEMBAR PENGESAHAN</div>'
    +'<div class="sig-grid">'
    +'<div class="sig-box"><div class="sig-title">DIBUAT OLEH</div>'
    +'<div class="sig-name">Abdur Roni</div>'
    +'<div class="sig-role">Officer – Divisi Strategi dan Perencanaan<br>Teknologi Informasi</div></div>'
    +'<div class="sig-box"><div class="sig-title">DIPERIKSA OLEH</div>'
    +'<div class="sig-name">M. Yusup Hamdani</div>'
    +'<div class="sig-role">Pj. Kepala Bagian – Divisi Strategi dan<br>Perencanaan Teknologi Informasi</div></div>'
    +'</div>'
    +'<div class="sig-center"><div class="sig-box"><div class="sig-title">DISETUJUI OLEH</div>'
    +'<div class="sig-name">Satria Pujakesuma</div>'
    +'<div class="sig-role">Kepala Divisi Strategi dan Perencanaan<br>Teknologi Informasi</div></div></div>'
    +'<div class="footer">Dokumen ini bersifat confidential.<br>Proyeksi dan estimasi dapat berubah sesuai kondisi aktual.</div>'
    +'</div>'

    /* ── DISCLAIMER ── */
    +'<div class="page">'
    +makeHeader('3')
    +'<div class="sec sec-italic">DISCLAIMER</div>'
    +'<p class="dl-text">Laporan <em>capacity planning</em> ini disusun berdasarkan data historis, analisis tren, dan proyeksi bisnis yang tersedia pada saat penyusunan. Seluruh perhitungan dan proyeksi dalam laporan ini bersifat estimasi dan perkiraan.</p>'
    +'<p class="dl-text">Perlu dipahami bahwa:</p>'
    +'<ul class="dl-list">'
    +'<li>Proyeksi kapasitas didasarkan pada asumsi pertumbuhan dan kondisi bisnis yang dapat berubah sewaktu-waktu.</li>'
    +'<li>Perbedaan antara proyeksi dengan realisasi aktual merupakan hal yang wajar dan tidak dapat sepenuhnya diprediksi.</li>'
    +'<li>Faktor eksternal seperti perubahan strategi bisnis, kondisi pasar, teknologi baru, atau kejadian tak terduga dapat mempengaruhi akurasi proyeksi.</li>'
    +'<li>Rekomendasi dan estimasi biaya dapat berubah sesuai kondisi pasar dan ketersediaan vendor.</li>'
    +'<li><em>Monitoring</em> dan <em>review</em> berkala sangat penting untuk menyesuaikan perencanaan dengan kondisi aktual.</li>'
    +'</ul>'
    +'<p class="dl-text">Laporan ini harus digunakan sebagai panduan perencanaan dan bukan sebagai jaminan absolut. Evaluasi dan <em>adjustment</em> berkelanjutan sangat disarankan.</p>'
    +'<div class="footer">Dokumen ini bersifat confidential.<br>Proyeksi dan estimasi dapat berubah sesuai kondisi aktual.</div>'
    +'</div>'

    /* ── LAPORAN BULANAN ── */
    +'<div class="page">'
    +makeHeader('4','DOKUMEN LAPORAN CAPACITY PLANNING')
    +'<div style="color:#003087;font-size:13px;font-weight:bold;margin-bottom:3px">LAPORAN BULANAN — '+label.toUpperCase()+'</div>'
    +'<div style="font-size:10px;color:#555;margin-bottom:14px">Periode: '+label+' | Tahun: '+year+'</div>'
    +'<div class="sec">Ringkasan Eksekutif</div>'
    +'<table class="st"><thead><tr><th>PARAMETER</th><th>NILAI</th></tr></thead><tbody>'
    +'<tr><td><b>PERIODE</b></td><td>'+label+'</td></tr>'
    +'<tr><td><b>DIPANTAU</b></td><td>'+_rptLastRows.length+' ('+withData+' ada data)</td></tr>'
    +'<tr><td><b>WARNING</b></td><td>'+warn+' metrik</td></tr>'
    +'<tr><td><b>KRITIS/OVER</b></td><td>'+crit+' metrik</td></tr>'
    +'<tr><td><b>RATA-RATA AKURASI</b></td><td>'+avgAcc+' ('+accRows.length+' terukur)</td></tr>'
    +'<tr><td><b>PERLU TINDAK LANJUT</b></td><td>'+urgentCount+' (ETA \u2264 18 periode)</td></tr>'
    +'</tbody></table>'
    +'<div class="sec">DETAIL KAPASITAS &amp; BREAKDOWN MINGGUAN PER RESOURCE</div>'
    +'<p style="font-size:10px;color:#555;margin-bottom:10px">Setiap resource menampilkan informasi lengkap: total kapasitas, realisasi pemakaian, proyeksi, akurasi proyeksi, pertumbuhan bulan ini &amp; bulan lalu, P90 utilisasi — diikuti detail pemakaian per minggu.</p>'
    +resourcesHtml
    +'<div class="footer">Dokumen ini bersifat confidential.<br>Proyeksi dan estimasi dapat berubah sesuai kondisi aktual.</div>'
    +'</div>'

    +'</body></html>';

  const w = window.open('','_blank','width=950,height=750,scrollbars=yes');
  if(!w){ alert('Pop-up diblokir browser. Izinkan pop-up untuk halaman ini lalu coba lagi.'); return; }
  w.document.write(html);
  w.document.close();
  setTimeout(function(){ w.print(); }, 600);
}
