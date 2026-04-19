function calcGrowth(arr){
  const n=arr.length; if(n<2) return {};
  const last=arr[n-1];
  return {
    weekly:  n>=2  ? last-arr[n-2]  : null,
    monthly: n>=4  ? last-arr[n-4]  : null,
    quarterly:n>=12? last-arr[n-12] : null
  };
}

function fmtG(v){
  if(v==null) return {t:'—',c:'var(--text-dim)'};
  return {t:(v>0?'+':'')+v.toFixed(2)+'%', c:v>0?'var(--danger)':v<0?'var(--accent3)':'var(--text-dim)'};
}

function calcETA(vals,threshold){
  if(vals.length<2) return null;
  const last=vals[vals.length-1];
  if(last>=threshold) return 'Sudah melewati threshold';
  const diffs=[]; for(let i=1;i<vals.length;i++) diffs.push(vals[i]-vals[i-1]);
  const avg=diffs.reduce((a,b)=>a+b,0)/diffs.length;
  if(avg<=0) return 'Tidak ada pertumbuhan';
  return `~${Math.ceil((threshold-last)/avg)} periode lagi`;
}

function renderGrowth(){
  const gg=document.getElementById('growth-grid'); if(!gg) return;
  if(!resources.length){ gg.innerHTML=`<div class="empty-state"><p>Belum ada resource.</p></div>`; return; }

  let html='';
  resources.forEach(r=>{
    const metrics=r.type==='cluster'
      ?[{m:'cpu',l:'CPU',c:'#fb923c'},{m:'ram',l:'RAM',c:'#c084fc'}]
      :[{m:'stor',l:'Storage',c:'#00d4ff'}];
    metrics.forEach(({m,l,c})=>{
      const entries=history.filter(h=>h.resId==r.id&&h[m]!==''&&h[m]!=null)
        .sort((a,b)=>String(a.date).localeCompare(String(b.date)));
      const vals=entries.map(h=>parseFloat(h[m]));
      const g=calcGrowth(vals);
      const w=fmtG(g.weekly),mo=fmtG(g.monthly),q=fmtG(g.quarterly);
      html+=`<div class="growth-card" style="cursor:pointer" onclick="drillGrowth(${r.id})">
        <div style="font-weight:600;margin-bottom:10px">${r.name} — <span style="color:${c}">${l}</span></div>
        <div class="growth-row"><span class="growth-lbl">Mingguan</span><span class="growth-val" style="color:${w.c}">${w.t}</span></div>
        <div class="growth-row"><span class="growth-lbl">Bulanan</span><span class="growth-val" style="color:${mo.c}">${mo.t}</span></div>
        <div class="growth-row"><span class="growth-lbl">3 Bulanan</span><span class="growth-val" style="color:${q.c}">${q.t}</span></div>
        <div style="text-align:right;font-size:10px;color:var(--accent);margin-top:8px">→ Detail</div>
      </div>`;
    });
  });
  gg.innerHTML=`<div class="growth-grid">${html}</div>`;
}

function drillGrowth(resId){
  const r=resources.find(x=>x.id==resId); if(!r) return;
  const lv=document.getElementById('growth-list-view');
  const dv=document.getElementById('growth-drill-view');
  const wrap=document.getElementById('growth-drill-content');
  if(lv) lv.style.display='none'; if(dv) dv.style.display='';

  const metrics=r.type==='cluster'
    ?[{m:'cpu',l:'CPU',c:'#fb923c'},{m:'ram',l:'RAM',c:'#c084fc'}]
    :[{m:'stor',l:'Storage',c:'#00d4ff'}];
  const thr=getThr(r);
  let html=`<div style="font-family:'Space Mono',monospace;font-size:16px;font-weight:700;color:var(--text-bright);margin-bottom:16px">${r.name}</div>`;

  metrics.forEach(({m,l,c})=>{
    const entries=history.filter(h=>h.resId==resId&&h[m]!==''&&h[m]!=null)
      .sort((a,b)=>String(a.date).localeCompare(String(b.date)));
    const vals=entries.map(h=>parseFloat(h[m]));
    const labels=entries.map(h=>h.label||h.date);
    const thrVal=m==='cpu'?thr.cpu:m==='ram'?thr.ram:thr.stor;
    const latest=vals.length?vals[vals.length-1]:null;
    const g=calcGrowth(vals);
    const w=fmtG(g.weekly),mo=fmtG(g.monthly),q=fmtG(g.quarterly);
    const eta=thrVal?calcETA(vals,thrVal):null;
    const cid=`gdr-${resId}-${m}`;

    html+=`<div class="card" style="margin-bottom:16px;border-left:3px solid ${c}">
      <div style="font-family:'Space Mono',monospace;font-size:12px;font-weight:700;color:${c};margin-bottom:12px">${l}</div>
      <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(130px,1fr));gap:8px;margin-bottom:14px">
        <div style="background:var(--bg3);border-radius:6px;padding:10px">
          <div style="font-size:9px;color:var(--text-dim);font-family:'Space Mono',monospace;letter-spacing:1px;margin-bottom:3px">SAAT INI</div>
          <div style="font-size:20px;font-family:'Space Mono',monospace;font-weight:700;color:${c}">${latest!=null?latest.toFixed(1)+'%':'—'}</div>
        </div>
        <div style="background:var(--bg3);border-radius:6px;padding:10px">
          <div style="font-size:9px;color:var(--text-dim);font-family:'Space Mono',monospace;letter-spacing:1px;margin-bottom:3px">MINGGUAN</div>
          <div style="font-size:16px;font-family:'Space Mono',monospace;font-weight:700;color:${w.c}">${w.t}</div>
        </div>
        <div style="background:var(--bg3);border-radius:6px;padding:10px">
          <div style="font-size:9px;color:var(--text-dim);font-family:'Space Mono',monospace;letter-spacing:1px;margin-bottom:3px">BULANAN</div>
          <div style="font-size:16px;font-family:'Space Mono',monospace;font-weight:700;color:${mo.c}">${mo.t}</div>
        </div>
        <div style="background:var(--bg3);border-radius:6px;padding:10px">
          <div style="font-size:9px;color:var(--text-dim);font-family:'Space Mono',monospace;letter-spacing:1px;margin-bottom:3px">3 BULANAN</div>
          <div style="font-size:16px;font-family:'Space Mono',monospace;font-weight:700;color:${q.c}">${q.t}</div>
        </div>
        ${thrVal&&eta?`<div style="background:var(--bg3);border-radius:6px;padding:10px;border:1px solid rgba(255,204,0,.25)">
          <div style="font-size:9px;color:var(--warn);font-family:'Space Mono',monospace;letter-spacing:1px;margin-bottom:3px">ETA THRESHOLD ${thrVal}%</div>
          <div style="font-size:13px;font-family:'Space Mono',monospace;font-weight:700;color:var(--warn)">${eta}</div>
        </div>`:''}
      </div>
      ${entries.length>1?`<div style="margin-bottom:14px"><canvas id="${cid}" height="150"></canvas></div>`:''}
      ${entries.length?`
      <div style="font-size:9px;color:var(--text-dim);font-family:'Space Mono',monospace;letter-spacing:1px;margin-bottom:8px">RIWAYAT (${entries.length} entri)</div>
      <div class="tbl-wrap"><div style="overflow-x:auto"><table>
        <thead><tr><th>Tanggal</th><th>Label</th><th>Utilisasi</th><th>Perubahan</th><th>Status</th></tr></thead>
        <tbody>${[...entries].reverse().map((h,i,arr)=>{
          const v=parseFloat(h[m]);
          const prev=i<arr.length-1?parseFloat(arr[i+1][m]):null;
          const diff=prev!=null?v-prev:null;
          const dc=diff==null?'var(--text-dim)':diff>0?'var(--danger)':diff<0?'var(--accent3)':'var(--text-dim)';
          const s=thrVal?statusOf(v,thrVal):'ok';
          return `<tr>
            <td style="font-family:'Space Mono',monospace;font-size:11px">${h.date}</td>
            <td style="color:var(--text-dim);font-size:11px">${h.label||'—'}</td>
            <td style="font-family:'Space Mono',monospace;font-size:11px;color:${c}">${v.toFixed(1)}%</td>
            <td style="font-family:'Space Mono',monospace;font-size:11px;color:${dc}">${diff!=null?(diff>0?'+':'')+diff.toFixed(2)+'%':'—'}</td>
            <td><span class="ab ${SAC[s]}">${SL[s]}</span></td>
          </tr>`;
        }).join('')}</tbody>
      </table></div></div>`:'<div style="color:var(--text-dim);font-size:11px">Belum ada data histori.</div>'}
    </div>`;
  });

  wrap.innerHTML=html;

  setTimeout(()=>{
    metrics.forEach(({m,l,c})=>{
      const entries=history.filter(h=>h.resId==resId&&h[m]!==''&&h[m]!=null)
        .sort((a,b)=>String(a.date).localeCompare(String(b.date)));
      if(entries.length<2) return;
      const cid=`gdr-${resId}-${m}`;
      const thr=getThr(r);
      const thrVal=m==='cpu'?thr.cpu:m==='ram'?thr.ram:thr.stor;
      const cEl=document.getElementById(cid); if(!cEl) return;
      if(growthCharts[cid]) growthCharts[cid].destroy();
      growthCharts[cid]=new Chart(cEl,{
        type:'line',
        data:{labels:entries.map(h=>h.label||h.date),datasets:[{label:l+' (%)',data:entries.map(h=>parseFloat(h[m])),borderColor:c,backgroundColor:c+'20',borderWidth:2,tension:.4,pointRadius:3,fill:true}]},
        options:{responsive:true,plugins:{legend:{display:false}},scales:{
          x:{ticks:{color:'#64748b',font:{size:10}},grid:{color:'#1e2d45'}},
          y:{ticks:{color:'#64748b',font:{size:10}},grid:{color:'#1e2d45'},min:0,max:110}
        }},
        plugins:[{id:'thr',afterDraw(chart){
          if(!thrVal) return;
          const{ctx,chartArea:{left,right,top,bottom},scales:{y}}=chart;
          const yp=y.getPixelForValue(thrVal); if(yp<top||yp>bottom) return;
          ctx.save();ctx.setLineDash([5,4]);ctx.strokeStyle='rgba(255,204,0,0.6)';ctx.lineWidth=1.5;
          ctx.beginPath();ctx.moveTo(left,yp);ctx.lineTo(right,yp);ctx.stroke();
          ctx.fillStyle='#ffcc00';ctx.font='9px Space Mono';ctx.fillText('Threshold '+thrVal+'%',left+4,yp-4);
          ctx.restore();
        }}]
      });
    });
  },50);
}

function backToGrowth(){
  const lv=document.getElementById('growth-list-view');
  const dv=document.getElementById('growth-drill-view');
  if(lv) lv.style.display=''; if(dv) dv.style.display='none';
  Object.values(growthCharts).forEach(c=>c.destroy());
  growthCharts={};
}
