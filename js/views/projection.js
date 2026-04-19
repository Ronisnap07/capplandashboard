function toggleMonthBreakdown(id){
  const el=document.getElementById(id);
  const arr=document.getElementById('arr-'+id);
  if(!el) return;
  const isOpen=el.style.display!=='none';
  el.style.display=isOpen?'none':'block';
  if(arr) arr.style.transform=isOpen?'rotate(0deg)':'rotate(90deg)';
}

function getActualByMonth(resId,metric,yearMonth){
  const r=resources.find(x=>x.id==resId); if(!r) return null;
  const cap=getCapacity(r,metric); if(!cap) return null;
  const entries=history.filter(h=>{
    if(h.resId!=resId) return false;
    if(h[metric]===''||h[metric]==null) return false;
    return String(h.date).startsWith(yearMonth);
  });
  if(!entries.length) return null;
  const avg=entries.reduce((s,h)=>s+parseFloat(h[metric]),0)/entries.length;
  return parseFloat(((avg/100)*cap).toFixed(2));
}

function calcMonthlyProjections(startVal,endVal,year){
  const months=[];
  const MNAMES=['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
  for(let m=0;m<12;m++){
    const projected=parseFloat((startVal+(endVal-startVal)*((m+1)/12)).toFixed(2));
    const monthStr=String(m+1).padStart(2,'0');
    months.push({month:m+1,label:MNAMES[m],yearMonth:`${year}-${monthStr}`,projected});
  }
  return months;
}

function renderProjection(){
  const wrap=document.getElementById('proj-content'); if(!wrap) return;
  if(!projections.length){
    wrap.innerHTML=`<div class="empty-state"><div style="font-size:32px">📈</div><p>Belum ada data proyeksi.</p></div>`;
    return;
  }

  const years=[...new Set(projections.map(p=>p.year))].sort();
  let html='';

  years.forEach(year=>{
    const yearProjs=projections.filter(p=>p.year==year);
    html+=`<div style="margin-bottom:32px"><div class="sec-title">Proyeksi Tahun ${year}</div>`;

    yearProjs.forEach(p=>{
      const r=resources.find(x=>x.id==p.resId);
      const cap=r?getCapacity(r,p.metric):null;
      const unit=getUnit(p.metric);
      const actual=getActualAbsolute(p.resId,p.metric);
      const actualPct=getLatest(p.resId,p.metric);
      const thr=r?getThr(r):{};
      const thrVal=p.metric==='cpu'?thr.cpu:p.metric==='ram'?thr.ram:thr.stor;
      const thrAbs=cap&&thrVal?parseFloat(((thrVal/100)*cap).toFixed(2)):null;
      const projPct=cap&&p.projected?parseFloat(((p.projected/cap)*100).toFixed(1)):null;
      const metricColor=p.metric==='cpu'?'var(--orange)':p.metric==='ram'?'var(--purple)':'var(--accent)';

      let accBadge='', accInfo='—', accColor='var(--text-dim)', accScore=null;
      if(actual!=null&&p.projected){
        const err=Math.abs(actual-p.projected)/p.projected*100;
        accScore=Math.max(0,100-err);
        const diff=actual-p.projected;
        accInfo=diff>0?`+${diff.toFixed(2)} ${unit} dari proyeksi`:`${diff.toFixed(2)} ${unit} dari proyeksi`;
        accColor=accScore>=90?'var(--accent3)':accScore>=80?'var(--warn)':'var(--danger)';
        accBadge=`<span style="font-family:'Space Mono',monospace;font-size:13px;font-weight:700;color:${accColor};background:${accColor}18;border:1px solid ${accColor}44;padding:3px 10px;border-radius:4px">${accScore.toFixed(1)}% akurat</span>`;
      } else if(actual==null){
        accBadge=`<span class="ab" style="color:var(--text-dim);border-color:var(--border)">Belum ada data</span>`;
      }

      let thrWarn='';
      if(thrAbs!=null&&p.projected){
        const pctToThr=(p.projected/thrAbs)*100;
        if(pctToThr>=85&&pctToThr<100) thrWarn=`<div style="margin-top:8px;padding:6px 8px;background:rgba(255,204,0,.08);border:1px solid rgba(255,204,0,.3);border-radius:4px;font-size:11px;color:var(--warn)">⚠ Proyeksi mendekati threshold (${pctToThr.toFixed(0)}% dari batas)</div>`;
        else if(pctToThr>=100) thrWarn=`<div style="margin-top:8px;padding:6px 8px;background:rgba(255,69,96,.08);border:1px solid rgba(255,69,96,.3);border-radius:4px;font-size:11px;color:var(--danger)">✕ Proyeksi MELEBIHI threshold!</div>`;
      }
      let realWarn='';
      if(actual!=null&&thrAbs!=null){
        const realPctToThr=(actual/thrAbs)*100;
        if(realPctToThr>=85&&realPctToThr<100) realWarn=`<div style="margin-top:6px;padding:6px 8px;background:rgba(255,107,53,.08);border:1px solid rgba(255,107,53,.3);border-radius:4px;font-size:11px;color:var(--orange)">⚠ Realisasi mendekati threshold — proyeksi perlu direvisi!</div>`;
        else if(realPctToThr>=100) realWarn=`<div style="margin-top:6px;padding:6px 8px;background:rgba(255,69,96,.12);border:1px solid rgba(255,69,96,.5);border-radius:4px;font-size:11px;color:var(--danger)"><b>✕ Realisasi MELEWATI threshold — capacity planning tidak akurat!</b></div>`;
      }

      const startVal=actual!=null?actual:0;
      const months=calcMonthlyProjections(startVal,p.projected,year);

      let monthRows='';
      let monthAccScores=[];
      months.forEach(m=>{
        const actualMonth=getActualByMonth(p.resId,p.metric,m.yearMonth);
        let mAccScore=null, mAccColor='var(--text-dim)', mActualStr='—', mDiffStr='—';
        if(actualMonth!=null){
          const err=Math.abs(actualMonth-m.projected)/m.projected*100;
          mAccScore=Math.max(0,100-err);
          mAccColor=mAccScore>=90?'var(--accent3)':mAccScore>=80?'var(--warn)':'var(--danger)';
          mActualStr=`${actualMonth} ${unit}`;
          const diff=actualMonth-m.projected;
          mDiffStr=(diff>0?'+':'')+diff.toFixed(2);
          monthAccScores.push(mAccScore);
        }
        const projPctM=cap?((m.projected/cap)*100).toFixed(1):null;
        const isThrWarn=thrAbs&&m.projected>=thrAbs*0.85&&m.projected<thrAbs;
        const isThrOver=thrAbs&&m.projected>=thrAbs;
        monthRows+=`<tr>
          <td style="font-family:'Space Mono',monospace;font-size:11px;color:var(--text-dim)">${m.label} ${year}</td>
          <td style="font-family:'Space Mono',monospace;font-size:11px">
            <span style="color:var(--purple)">${m.projected} ${unit}</span>
            ${projPctM?`<span style="color:var(--text-dim);font-size:10px"> (${projPctM}%)</span>`:''}
            ${isThrOver?`<span style="color:var(--danger);font-size:10px"> ⚠ >Thr</span>`:isThrWarn?`<span style="color:var(--warn);font-size:10px"> ⚠ ~Thr</span>`:''}
          </td>
          <td style="font-family:'Space Mono',monospace;font-size:11px;color:var(--accent3)">${mActualStr}</td>
          <td style="font-family:'Space Mono',monospace;font-size:11px;color:${mAccColor}">${mDiffStr!=='—'?mDiffStr+' '+unit:'—'}</td>
          <td>
            ${mAccScore!=null?`
              <div style="display:flex;align-items:center;gap:6px">
                <div style="flex:1;height:4px;background:var(--bg3);border-radius:2px;overflow:hidden;min-width:60px">
                  <div style="height:100%;width:${mAccScore}%;background:${mAccColor};border-radius:2px"></div>
                </div>
                <span style="font-family:'Space Mono',monospace;font-size:10px;color:${mAccColor};white-space:nowrap">${mAccScore.toFixed(1)}%</span>
              </div>`:'<span style="font-size:10px;color:var(--text-dim)">Belum ada data</span>'}
          </td>
        </tr>`;
      });

      const avgAcc=monthAccScores.length?monthAccScores.reduce((a,b)=>a+b,0)/monthAccScores.length:null;
      const avgAccColor=avgAcc!=null?(avgAcc>=90?'var(--accent3)':avgAcc>=80?'var(--warn)':'var(--danger)'):'var(--text-dim)';

      html+=`<div class="card" style="margin-bottom:14px;border-left:3px solid ${metricColor}">
        <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:12px;flex-wrap:wrap;gap:8px">
          <div>
            <div style="font-family:'Space Mono',monospace;font-size:13px;color:var(--text-bright)">${r?r.name:'—'} <span style="color:${metricColor}">/ ${p.metric.toUpperCase()}</span></div>
            <div style="font-size:11px;color:var(--text-dim);margin-top:2px">${cap?'Kapasitas: '+cap+' '+unit:'—'}</div>
          </div>
          <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
            ${avgAcc!=null?`<span style="font-size:11px;color:var(--text-dim)">Rata-rata akurasi:</span><span style="font-family:'Space Mono',monospace;font-size:13px;font-weight:700;color:${avgAccColor};background:${avgAccColor}18;border:1px solid ${avgAccColor}44;padding:3px 10px;border-radius:4px">${avgAcc.toFixed(1)}%</span>`:''}
            ${accBadge}
          </div>
        </div>
        <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:10px;margin-bottom:12px">
          <div style="background:var(--bg3);border-radius:6px;padding:10px">
            <div style="font-size:10px;color:var(--text-dim);font-family:'Space Mono',monospace;letter-spacing:1px;margin-bottom:4px">TARGET ${year}</div>
            <div style="font-size:18px;font-family:'Space Mono',monospace;color:var(--purple)">${p.projected} <span style="font-size:11px">${unit}</span></div>
            ${projPct?`<div style="font-size:11px;color:var(--text-dim)">${projPct}% dari kapasitas</div>`:''}
          </div>
          <div style="background:var(--bg3);border-radius:6px;padding:10px">
            <div style="font-size:10px;color:var(--text-dim);font-family:'Space Mono',monospace;letter-spacing:1px;margin-bottom:4px">TITIK AWAL (HISTORI)</div>
            <div style="font-size:18px;font-family:'Space Mono',monospace;color:var(--accent3)">${actual!=null?actual:'—'} <span style="font-size:11px">${actual!=null?unit:''}</span></div>
            ${actualPct!=null?`<div style="font-size:11px;color:var(--text-dim)">${actualPct.toFixed(1)}% dari kapasitas</div>`:'<div style="font-size:11px;color:var(--text-dim)">Belum ada histori</div>'}
          </div>
          <div style="background:var(--bg3);border-radius:6px;padding:10px">
            <div style="font-size:10px;color:var(--text-dim);font-family:'Space Mono',monospace;letter-spacing:1px;margin-bottom:4px">THRESHOLD</div>
            <div style="font-size:18px;font-family:'Space Mono',monospace;color:var(--warn)">${thrAbs!=null?thrAbs:'—'} <span style="font-size:11px">${thrAbs!=null?unit:''}</span></div>
            ${thrVal?`<div style="font-size:11px;color:var(--text-dim)">${thrVal}% dari kapasitas</div>`:''}
          </div>
          <div style="background:var(--bg3);border-radius:6px;padding:10px">
            <div style="font-size:10px;color:var(--text-dim);font-family:'Space Mono',monospace;letter-spacing:1px;margin-bottom:4px">KENAIKAN / BULAN</div>
            <div style="font-size:18px;font-family:'Space Mono',monospace;color:${metricColor}">${parseFloat(((p.projected-startVal)/12).toFixed(2))} <span style="font-size:11px">${unit}</span></div>
            <div style="font-size:11px;color:var(--text-dim)">linear / bulan</div>
          </div>
        </div>
        ${thrWarn}${realWarn}
        <div style="margin-top:14px;border:1px solid var(--border);border-radius:6px;overflow:hidden">
          <div onclick="toggleMonthBreakdown('mb-${p.resId}-${p.metric}-${year}')" style="display:flex;justify-content:space-between;align-items:center;padding:10px 14px;cursor:pointer;background:var(--bg3);user-select:none" onmouseover="this.style.background='rgba(0,212,255,0.06)'" onmouseout="this.style.background='var(--bg3)'">
            <div style="font-family:'Space Mono',monospace;font-size:9px;letter-spacing:2px;color:${metricColor};display:flex;align-items:center;gap:6px">
              <span style="width:10px;height:2px;background:${metricColor};display:inline-block"></span>
              BREAKDOWN BULANAN
            </div>
            <div style="display:flex;align-items:center;gap:8px">
              <span style="width:10px;height:2px;background:${metricColor};display:inline-block"></span>
              Detail
            </div>
          </div>
          <div id="mb-${p.resId}-${p.metric}-${year}" style="display:none;padding:0">
            <div style="overflow-x:auto">
              <table>
                <thead><tr><th>Bulan</th><th>Proyeksi Linear</th><th>Realisasi Aktual</th><th>Selisih</th><th>Akurasi</th></tr></thead>
                <tbody>${monthRows}</tbody>
              </table>
            </div>
          </div>
        </div>
      </div>`;
    });

    html+=`</div>`;
  });

  wrap.innerHTML=html;
}
