function drawGauge(id,val,thr,color){
  const c=document.getElementById(id); if(!c) return;
  const ctx=c.getContext('2d'); c.width=120; c.height=72;
  ctx.clearRect(0,0,120,72);
  const cx=60,cy=70,r=52;
  ctx.beginPath(); ctx.arc(cx,cy,r,Math.PI,2*Math.PI); ctx.strokeStyle='rgba(0,212,255,0.15)'; ctx.lineWidth=10; ctx.stroke();
  ctx.beginPath(); ctx.arc(cx,cy,r,Math.PI,Math.PI+(thr/100)*Math.PI); ctx.strokeStyle='rgba(255,204,0,0.15)'; ctx.lineWidth=10; ctx.stroke();
  const pct=Math.min(val??0,120)/100;
  ctx.beginPath(); ctx.arc(cx,cy,r,Math.PI,Math.PI+pct*Math.PI); ctx.strokeStyle=color; ctx.lineWidth=10; ctx.lineCap='round'; ctx.stroke();
  ctx.save(); ctx.translate(cx,cy); ctx.rotate(Math.PI+(thr/100)*Math.PI);
  ctx.beginPath(); ctx.moveTo(r-14,0); ctx.lineTo(r+4,0); ctx.strokeStyle='rgba(255,204,0,0.7)'; ctx.lineWidth=2; ctx.stroke();
  ctx.restore();
}

function renderOverview(){
  if(!resources.length){
    document.getElementById('gauge-grid').innerHTML=`<div class="empty-state" style="grid-column:1/-1"><p>Belum ada resource. Input di Editor.</p></div>`;
    return;
  }

  let html='';
  resources.forEach(r=>{
    const thr=getThr(r);
    const metrics=r.type==='cluster'
      ?[{m:'cpu',l:'CPU',t:thr.cpu},{m:'ram',l:'RAM',t:thr.ram}]
      :[{m:'stor',l:r.type==='storage_core'?'Core':'Support',t:thr.stor}];
    metrics.forEach(({m,l,t})=>{
      const val=getLatest(r.id,m); const s=statusOf(val,t); const col=SC[s];
      html+=`<div class="gauge-card s-${s}" onclick="showTab('detail');drillDown(${r.id})">
        <div class="gauge-title">${r.name}</div><div class="gauge-sub">${l}</div>
        <div class="gauge-wrap"><canvas id="g-${r.id}-${m}"></canvas><div class="gauge-val" style="color:${col}">${val!=null?val.toFixed(1)+'%':'—'}</div></div>
        <div class="gauge-status" style="color:${col}">${SL[s]}</div>
        <div class="gauge-thr">Threshold: ${t}%</div>
      </div>`;
    });
  });
  document.getElementById('gauge-grid').innerHTML=html;

  setTimeout(()=>{
    resources.forEach(r=>{
      const thr=getThr(r);
      if(r.type==='cluster'){
        [{m:'cpu',t:thr.cpu},{m:'ram',t:thr.ram}].forEach(({m,t})=>{
          const v=getLatest(r.id,m); drawGauge(`g-${r.id}-${m}`,v,t,SC[statusOf(v,t)]);
        });
      } else {
        const v=getLatest(r.id,'stor'); drawGauge(`g-${r.id}-stor`,v,thr.stor,SC[statusOf(v,thr.stor)]);
      }
    });
  },50);

  // Summary cards
  let total=0,ok=0,warn=0,danger=0;
  resources.forEach(r=>{
    const thr=getThr(r);
    (r.type==='cluster'?[{m:'cpu',t:thr.cpu},{m:'ram',t:thr.ram}]:[{m:'stor',t:thr.stor}]).forEach(({m,t})=>{
      total++;
      const s=statusOf(getLatest(r.id,m),t);
      if(s==='ok')ok++; else if(s==='warn')warn++; else if(s==='danger'||s==='over')danger++;
    });
  });
  document.getElementById('sum-grid').innerHTML=[
    {l:'Total Metrik',v:total,c:'',s:'dipantau'},
    {l:'Aman',v:ok,c:'var(--accent3)',s:'metrik'},
    {l:'Warning',v:warn,c:'var(--warn)',s:'metrik'},
    {l:'Kritis/Over',v:danger,c:'var(--danger)',s:'metrik'},
    {l:'Resources',v:resources.length,c:'var(--accent)',s:'terdaftar'},
    {l:'Data Histori',v:history.length,c:'',s:'entri'},
  ].map(x=>`<div class="sum-card"><div class="sum-label">${x.l}</div><div class="sum-val" style="color:${x.c||'var(--text-bright)'}">${x.v}</div><div class="sum-sub">${x.s}</div></div>`).join('');

  // Overview chart
  const labels=[],vals=[];
  resources.forEach(r=>{
    const thr=getThr(r);
    (r.type==='cluster'?[{m:'cpu',l:'CPU'},{m:'ram',l:'RAM'}]:[{m:'stor',l:'Storage'}]).forEach(({m,l})=>{
      const v=getLatest(r.id,m); labels.push(`${r.name} ${l}`); vals.push(v??0);
    });
  });
  if(overviewChart) overviewChart.destroy();
  overviewChart=new Chart(document.getElementById('chart-overview'),{
    type:'bar',
    data:{labels,datasets:[{label:'Utilisasi (%)',data:vals,backgroundColor:'rgba(0,212,255,0.5)',borderColor:'#00d4ff',borderWidth:1,borderRadius:3}]},
    options:{responsive:true,plugins:{legend:{display:false}},scales:{
      x:{ticks:{color:'#64748b',font:{size:10}},grid:{color:'#1e2d45'}},
      y:{ticks:{color:'#64748b',font:{size:10}},grid:{color:'#1e2d45'},min:0,max:120}
    }}
  });
}
