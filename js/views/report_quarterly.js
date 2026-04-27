const RPTQ_QUARTERS = [
  null,
  {name:'Triwulan I',   short:'Q1', months:[1,2,3],   label:'Jan - Mar'},
  {name:'Triwulan II',  short:'Q2', months:[4,5,6],   label:'Apr - Jun'},
  {name:'Triwulan III', short:'Q3', months:[7,8,9],   label:'Jul - Sep'},
  {name:'Triwulan IV',  short:'Q4', months:[10,11,12],label:'Okt - Des'}
];

let _rptQLastRows = [], _rptQLastYearMonths = [], _rptQLastLabel = '';

function rptqYearMonths(year, quarter){
  return RPTQ_QUARTERS[quarter].months.map(function(m){
    return year + '-' + String(m).padStart(2,'0');
  });
}

function rptqPrevYearMonths(year, quarter){
  if(quarter === 1) return rptqYearMonths(year - 1, 4);
  return rptqYearMonths(year, quarter - 1);
}

function rptqAvg(resId, metric, yearMonths){
  var vals = yearMonths.map(function(ym){ return rptMonthAvg(resId, metric, ym); })
                       .filter(function(v){ return v != null; });
  return vals.length ? vals.reduce(function(s,v){ return s+v; }, 0) / vals.length : null;
}

// ───── Render UI ────────────────────────────────────────────────────
function renderReportQuarterly(){
  var wrap = document.getElementById('report-content-quarterly'); if(!wrap) return;
  var now  = new Date();
  var yearOpts = [];
  for(var y = 2024; y <= 2030; y++) yearOpts.push(y);
  var curQ = Math.ceil((now.getMonth() + 1) / 3);

  wrap.innerHTML =
    '<div class="card" style="margin-bottom:22px">'
    + '<div style="display:flex;gap:12px;align-items:flex-end;flex-wrap:wrap">'
    + '<div class="fg" style="flex:0 0 auto"><label>Tahun</label>'
    + '<select id="rptq-year" style="width:110px">'
    + yearOpts.map(function(y){ return '<option value="'+y+'"'+(y===now.getFullYear()?' selected':'')+'>'+y+'</option>'; }).join('')
    + '</select></div>'
    + '<div class="fg" style="flex:0 0 auto"><label>Triwulan</label>'
    + '<select id="rptq-quarter" style="width:200px">'
    + [1,2,3,4].map(function(q){
        var qi = RPTQ_QUARTERS[q];
        return '<option value="'+q+'"'+(q===curQ?' selected':'')+'>'
          +qi.name+' ('+qi.label+')</option>';
      }).join('')
    + '</select></div>'
    + '<button class="btn btn-primary" onclick="generateReportQuarterly()">⚡ Generate</button>'
    + '<button class="btn btn-sec rpt-no-print" onclick="exportReportQuarterlyWord()">⬇ Export Word</button>'
    + '</div></div>'
    + '<div id="rptq-result"></div>';
}

// ───── Generate ─────────────────────────────────────────────────────
function generateReportQuarterly(){
  var year    = parseInt(document.getElementById('rptq-year').value);
  var quarter = parseInt(document.getElementById('rptq-quarter').value);
  var qi      = RPTQ_QUARTERS[quarter];
  var yearMonths  = rptqYearMonths(year, quarter);
  var prevYMs     = rptqPrevYearMonths(year, quarter);
  var label   = qi.name + ' ' + year + ' (' + qi.label + ')';

  var wrap = document.getElementById('rptq-result'); if(!wrap) return;
  if(!resources.length){ wrap.innerHTML='<div class="empty-state"><p>Belum ada resource.</p></div>'; return; }

  var SC  = {ok:'var(--accent3)',warn:'var(--warn)',danger:'var(--danger)',over:'var(--danger)',na:'var(--text-dim)'};
  var SAC = {ok:'ab-ok',warn:'ab-warn',danger:'ab-danger',over:'ab-danger',na:'ab-na'};
  var SL  = {ok:'Normal',warn:'Warning',danger:'Kritis',over:'Over Cap',na:'No Data'};
  var TC  = {cluster:'ab-cluster',storage_core:'ab-core',storage_support:'ab-support'};
  var TL  = {cluster:'Cluster',storage_core:'Core',storage_support:'Support'};

  var rows = [];
  resources.forEach(function(r){
    var thr = getThr(r);
    var metrics = r.type==='cluster'
      ? [{m:'cpu',l:'CPU',t:thr.cpu},{m:'ram',l:'RAM',t:thr.ram}]
      : [{m:'stor',l:'Storage',t:thr.stor}];

    metrics.forEach(function(mi){
      var m = mi.m, l = mi.l, t = mi.t;
      var cap  = getCapacity(r, m);
      var unit = getUnit(m);

      var avgUtil  = rptqAvg(r.id, m, yearMonths);
      var prevUtil = rptqAvg(r.id, m, prevYMs);
      var delta    = avgUtil != null && prevUtil != null ? avgUtil - prevUtil : null;

      var projVals = [];
      var proj = projections.find(function(p){ return p.resId==r.id && p.metric===m && p.year==year; });
      if(proj && cap){
        var startAbs = rptYearStartAbs(r.id, m, year);
        var mps = calcMonthlyProjections(startAbs || 0, proj.projected, year);
        qi.months.forEach(function(mo){
          var mp = mps.find(function(x){ return x.month === mo; });
          if(mp) projVals.push(mp.projected);
        });
      }
      var projMonthAbs = projVals.length ? projVals.reduce(function(s,v){return s+v;},0)/projVals.length : null;
      var projMonthPct = projMonthAbs != null && cap ? ((projMonthAbs/cap)*100).toFixed(1) : null;

      var actualAbs = avgUtil != null && cap ? parseFloat(((avgUtil/100)*cap).toFixed(2)) : null;

      var accVals = [];
      if(proj && cap){
        var startAbs2 = rptYearStartAbs(r.id, m, year);
        var mps2 = calcMonthlyProjections(startAbs2 || 0, proj.projected, year);
        qi.months.forEach(function(mo){
          var ym  = year + '-' + String(mo).padStart(2,'0');
          var mUtil = rptMonthAvg(r.id, m, ym);
          var mAct  = mUtil != null && cap ? (mUtil/100)*cap : null;
          var mp2   = mps2.find(function(x){ return x.month === mo; });
          if(mp2 && mp2.projected > 0 && mAct != null){
            accVals.push(Math.max(0, 100 - Math.abs(mAct - mp2.projected)/mp2.projected*100));
          }
        });
      }
      var accuracy = accVals.length ? accVals.reduce(function(s,v){return s+v;},0)/accVals.length : null;
      var accColor = accuracy==null?'var(--text-dim)':accuracy>=90?'var(--accent3)':accuracy>=80?'var(--warn)':'var(--danger)';

      var allE   = history.filter(function(h){ return h.resId==r.id&&h[m]!==''&&h[m]!=null; })
                          .sort(function(a,b){ return String(a.date).localeCompare(String(b.date)); });
      var allV   = allE.map(function(h){ return parseFloat(h[m]); });
      var growth = calcGrowth(allV);
      var eta    = t ? calcETA(allV, t) : null;
      var status = statusOf(avgUtil, t);

      rows.push({r:r, m:m, l:l, t:t, unit:unit, cap:cap, avgUtil:avgUtil, delta:delta,
                 projMonthAbs:projMonthAbs, projMonthPct:projMonthPct, actualAbs:actualAbs,
                 accuracy:accuracy, accColor:accColor, status:status, growth:growth, eta:eta, allE:allE});
    });
  });

  _rptQLastRows      = rows;
  _rptQLastYearMonths = yearMonths;
  _rptQLastLabel     = label;

  var withData    = rows.filter(function(x){ return x.avgUtil!=null; }).length;
  var warn        = rows.filter(function(x){ return x.status==='warn'; }).length;
  var crit        = rows.filter(function(x){ return x.status==='danger'||x.status==='over'; }).length;
  var accRowsF    = rows.filter(function(x){ return x.accuracy!=null; });
  var avgAcc      = accRowsF.length ? accRowsF.reduce(function(s,x){return s+x.accuracy;},0)/accRowsF.length : null;
  var avgAccC     = avgAcc==null?'var(--text-dim)':avgAcc>=90?'var(--accent3)':avgAcc>=80?'var(--warn)':'var(--danger)';
  var urgentRows  = rows.map(function(x){
    var em = rptEtaMonths(x.eta); if(em===null) return null;
    var urgency = em===0?'over':em<=6?'kritis':em<=12?'waspada':em<=18?'perhatian':null;
    if(!urgency) return null;
    return Object.assign({},x,{etaMonths:em,urgency:urgency});
  }).filter(Boolean);

  var html = '<div class="sec-title">Laporan 3 Bulanan — '+label+'</div>'
    + '<div class="sum-grid" style="margin-bottom:22px">'
    + [
        {l:'Periode',             v:label,     c:'var(--accent)', s:'', fs:'13px'},
        {l:'Dipantau',            v:rows.length, c:'', s:withData+' ada data'},
        {l:'Warning',             v:warn,      c:'var(--warn)',   s:'metrik'},
        {l:'Kritis/Over',         v:crit,      c:'var(--danger)', s:'metrik'},
        {l:'Rata-rata Akurasi',   v:avgAcc!=null?avgAcc.toFixed(1)+'%':'—', c:avgAccC, s:accRowsF.length+' terukur'},
        {l:'Perlu Tindak Lanjut', v:urgentRows.length, c:urgentRows.length>0?'var(--danger)':'var(--accent3)', s:'ETA ≤ 18 periode'}
      ].map(function(x){
        return '<div class="sum-card"><div class="sum-label">'+x.l+'</div>'
          +'<div class="sum-val" style="color:'+(x.c||'var(--text-bright)')+(x.fs?';font-size:'+x.fs:'')+'">'+x.v+'</div>'
          +'<div class="sum-sub">'+x.s+'</div></div>';
      }).join('')
    + '</div>'
    + '<div class="tbl-wrap" style="margin-bottom:28px"><div style="overflow-x:auto">'
    + '<table id="rptq-main-table"><thead><tr>'
    + '<th>#</th><th>Resource</th><th>Tipe</th><th>Metrik</th>'
    + '<th>Avg Util (3 Bln)</th><th>vs Triwulan Lalu</th><th>Growth</th>'
    + '<th>Proyeksi Rata-rata</th><th>Realisasi Rata-rata</th><th>Selisih</th>'
    + '<th>Akurasi</th><th>ETA Threshold</th><th>Status</th>'
    + '</tr></thead><tbody>'
    + rows.map(function(x, i){
        var uC   = SC[x.status];
        var dC   = x.delta==null?'var(--text-dim)':x.delta>0?'var(--danger)':x.delta<0?'var(--accent3)':'var(--text-dim)';
        var gmo  = fmtG(x.growth.monthly);
        var diff = x.actualAbs!=null&&x.projMonthAbs!=null ? x.actualAbs-x.projMonthAbs : null;
        var dfC  = diff==null?'var(--text-dim)':diff>0?'var(--danger)':'var(--accent3)';
        var qEntries = x.allE.filter(function(h){
          var d = String(h.date);
          return yearMonths.some(function(ym){ return d.startsWith(ym); });
        });
        var weeklyHtml = rptWeeklyRows(qEntries, x.m, x.unit, x.cap);
        return '<tr style="cursor:pointer" onclick="rptToggleWeekly(\'rptq-wr-'+i+'\')">'
          +'<td style="color:var(--text-dim);font-size:10px">'+(i+1)+'</td>'
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
            ?'<div style="display:flex;align-items:center;gap:6px">'
              +'<div style="width:48px;height:4px;background:var(--bg3);border-radius:2px;overflow:hidden">'
              +'<div style="height:100%;width:'+x.accuracy+'%;background:'+x.accColor+';border-radius:2px"></div></div>'
              +'<span style="font-family:\'Space Mono\',monospace;font-size:10px;color:'+x.accColor+'">'+x.accuracy.toFixed(1)+'%</span></div>'
            :'<span style="font-size:10px;color:var(--text-dim)">—</span>')+'</td>'
          +'<td style="font-size:11px;color:var(--warn)">'+(x.eta||'—')+'</td>'
          +'<td><span class="ab '+SAC[x.status]+'">'+SL[x.status]+'</span></td>'
          +'</tr>'
          +'<tr id="rptq-wr-'+i+'" style="display:none">'
          +'<td colspan="13" style="padding:0;border-top:none">'+weeklyHtml+'</td></tr>';
      }).join('')
    + '</tbody></table></div></div>'
    + rptBuildUrgentHtml(urgentRows);

  wrap.innerHTML = html;
}

// ───── Export Word ───────────────────────────────────────────────────
function exportReportQuarterlyWord(){
  if(!_rptQLastRows.length){ alert('Generate report 3 bulanan dulu sebelum export Word.'); return; }
  if(typeof JSZip==='undefined'){ alert('Library Word belum siap, coba lagi.'); return; }

  var year    = parseInt(document.getElementById('rptq-year').value);
  var quarter = parseInt(document.getElementById('rptq-quarter').value);
  var qi      = RPTQ_QUARTERS[quarter];
  var label   = _rptQLastLabel;
  var todayStr= new Date().toLocaleDateString('id-ID',{day:'numeric',month:'long',year:'numeric'});

  _buildMultiTemplateDoc(year, qi.months[0], label, todayStr,
    _rptQLastYearMonths[0], _rptQLastRows, _rptQLastYearMonths)
    .then(function(blob){
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = 'Laporan_Monitoring_Capplan_'+qi.short+'_'+year+'.docx';
      document.body.appendChild(a); a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(a.href);
    })
    .catch(function(e){ console.error('Export Word Q error:',e); alert('Gagal export Word: '+e.message); });
}
