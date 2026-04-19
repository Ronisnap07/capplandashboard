function setDetailView(v){
  detailView=v;
  document.querySelectorAll('.toggle-btn').forEach((b,i)=>b.classList.toggle('active',i===(v==='list'?0:1)));
  renderDetailList();
}

function renderDetailList(){
  const wrap=document.getElementById('detail-list-content'); if(!wrap) return;
  if(!resources.length){ wrap.innerHTML=`<div class="empty-state"><p>Belum ada resource.</p></div>`; return; }

  if(detailView==='list'){
    wrap.innerHTML=`<div class="tbl-wrap"><div style="overflow-x:auto"><table>
      <thead><tr><th>#</th><th>Resource</th><th>Tipe</th><th>CPU%</th><th>RAM%</th><th>Storage%</th><th>Status</th><th>Aksi</th></tr></thead>
      <tbody>
      ${resources.map((r,i)=>{
        const thr=getThr(r);
        const cpu=getLatest(r.id,'cpu'),ram=getLatest(r.id,'ram'),stor=getLatest(r.id,'stor');
        const sc=statusOf(cpu,thr.cpu),sr=statusOf(ram,thr.ram),ss=statusOf(stor,thr.stor);
        const worst=r.type==='cluster'
          ?([sc,sr].find(s=>s==='over')||[sc,sr].find(s=>s==='danger')||[sc,sr].find(s=>s==='warn')||'ok')
          :ss;
        const tc={cluster:'ab-cluster',storage_core:'ab-core',storage_support:'ab-support'};
        const tl={cluster:'Cluster',storage_core:'Storage Core',storage_support:'Storage Support'};
        return `<tr>
          <td style="color:var(--text-dim);font-size:10px">${i+1}</td>
          <td><b>${r.name}</b></td>
          <td><span class="ab ${tc[r.type]}">${tl[r.type]}</span></td>
          <td style="font-family:'Space Mono',monospace">${cpu!=null?cpu.toFixed(1)+'%':'—'}</td>
          <td style="font-family:'Space Mono',monospace">${ram!=null?ram.toFixed(1)+'%':'—'}</td>
          <td style="font-family:'Space Mono',monospace">${stor!=null?stor.toFixed(1)+'%':'—'}</td>
          <td><span class="ab ${SAC[worst]}">${SL[worst]}</span></td>
          <td><button class="act-btn" onclick="drillDown(${r.id})">→ Detail</button></td>
        </tr>`;
      }).join('')}
      </tbody></table></div></div>`;
  } else {
    const types=[{k:'cluster',l:'🖥️ Host Cluster'},{k:'storage_core',l:'💾 Storage Core'},{k:'storage_support',l:'💾 Storage Support'}];
    wrap.innerHTML=types.map(tp=>{
      const res=resources.filter(r=>r.type===tp.k); if(!res.length) return '';
      return `<div style="margin-bottom:20px">
        <div class="sec-title">${tp.l}</div>
        <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(210px,1fr));gap:12px">
          ${res.map(r=>{
            const thr=getThr(r);
            const metrics=r.type==='cluster'
              ?[{m:'cpu',l:'CPU',t:thr.cpu},{m:'ram',l:'RAM',t:thr.ram}]
              :[{m:'stor',l:'Storage',t:thr.stor}];
            return `<div class="sum-card" style="cursor:pointer" onclick="drillDown(${r.id})">
              <div style="font-weight:600;margin-bottom:10px">${r.name}</div>
              ${metrics.map(({m,l,t})=>{
                const v=getLatest(r.id,m); const s=statusOf(v,t);
                return `<div style="margin-bottom:7px">
                  <div style="display:flex;justify-content:space-between;font-size:11px">
                    <span style="color:var(--text-dim)">${l}</span>
                    <span style="font-family:'Space Mono',monospace;color:${SC[s]}">${v!=null?v.toFixed(1)+'%':'—'}</span>
                  </div>
                  <div class="util-bar-track"><div class="util-bar-fill" style="width:${v!=null?Math.min(v,100):0}%;background:${SC[s]}"></div></div>
                </div>`;
              }).join('')}
              <div style="text-align:right;font-size:10px;color:var(--accent)">→ Detail</div>
            </div>`;
          }).join('')}
        </div>
      </div>`;
    }).join('');
  }
}

function drillDown(resId){
  const r=resources.find(x=>x.id==resId); if(!r) return;
  const lv=document.getElementById('detail-list-view');
  const dv=document.getElementById('detail-drill-view');
  if(lv) lv.style.display='none'; if(dv) dv.style.display='block';

  const thr=getThr(r);
  const metrics=r.type==='cluster'
    ?[{m:'cpu',l:'CPU',t:thr.cpu,c:'var(--orange)'},{m:'ram',l:'RAM',t:thr.ram,c:'var(--purple)'}]
    :[{m:'stor',l:'Storage',t:thr.stor,c:'var(--accent)'}];
  const entries=history.filter(h=>h.resId==resId).sort((a,b)=>String(a.date).localeCompare(String(b.date)));
  const labels=entries.map(h=>h.label||h.date);

  let html=`<div style="font-family:'Space Mono',monospace;font-size:14px;color:var(--accent);margin-bottom:14px">${r.name}</div>
  <div class="chart-grid">`;
  metrics.forEach(({m,l,t})=>{
    const v=getLatest(resId,m); const s=statusOf(v,t);
    html+=`<div class="chart-card"><div class="chart-title">${l} — <span style="color:${SC[s]}">${v!=null?v.toFixed(1)+'%':'No Data'} (${SL[s]})</span></div><canvas id="dc-${resId}-${m}" height="160"></canvas></div>`;
  });
  html+=`</div>
  <div class="sec-title">Riwayat Data</div>
  <div class="tbl-wrap"><div style="overflow-x:auto"><table>
    <thead><tr><th>Tanggal</th><th>Label</th>${metrics.map(x=>`<th>${x.l}%</th><th>Status</th>`).join('')}</tr></thead>
    <tbody>
    ${[...entries].reverse().map(h=>`<tr>
      <td style="font-family:'Space Mono',monospace;font-size:11px">${h.date}</td>
      <td style="color:var(--text-dim);font-size:11px">${h.label||'—'}</td>
      ${metrics.map(({m,t})=>{
        const v=h[m]!==''&&h[m]!=null?parseFloat(h[m]):null;
        const s=statusOf(v,t);
        return `<td style="font-family:'Space Mono',monospace">${v!=null?v.toFixed(1)+'%':'—'}</td>
        <td>${v!=null?`<span class="ab ${SAC[s]}">${SL[s]}</span>`:'—'}</td>`;
      }).join('')}
    </tr>`).join('')}
    </tbody></table></div></div>`;

  document.getElementById('drill-content').innerHTML=html;

  setTimeout(()=>{
    metrics.forEach(({m,l,t,c})=>{
      const vals=entries.map(h=>{ const v=h[m]; return v!==''&&v!=null?parseFloat(v):null; });
      const cid=`dc-${resId}-${m}`;
      if(drillCharts[cid]) drillCharts[cid].destroy();
      drillCharts[cid]=new Chart(document.getElementById(cid),{
        type:'line',
        data:{labels,datasets:[{label:l+' (%)',data:vals,borderColor:'#00d4ff',backgroundColor:'rgba(0,212,255,0.12)',borderWidth:2,tension:.4,pointRadius:3,fill:true,spanGaps:true}]},
        options:{responsive:true,plugins:{legend:{display:false}},scales:{
          x:{ticks:{color:'#64748b',font:{size:10}},grid:{color:'#1e2d45'}},
          y:{ticks:{color:'#64748b',font:{size:10}},grid:{color:'#1e2d45'},min:0,max:110}
        }},
        plugins:[{id:'thr',afterDraw(chart){
          const{ctx,chartArea:{left,right,top,bottom},scales:{y}}=chart;
          const yp=y.getPixelForValue(t); if(yp<top||yp>bottom) return;
          ctx.save();ctx.setLineDash([5,4]);ctx.strokeStyle='rgba(255,204,0,0.5)';ctx.lineWidth=1.5;
          ctx.beginPath();ctx.moveTo(left,yp);ctx.lineTo(right,yp);ctx.stroke();
          ctx.fillStyle='#ffcc00';ctx.font='9px Space Mono';ctx.fillText('Max '+t+'%',left+4,yp-4);
          ctx.restore();
        }}]
      });
    });
  },50);
}

function backToList(){
  const lv=document.getElementById('detail-list-view');
  const dv=document.getElementById('detail-drill-view');
  if(lv) lv.style.display=''; if(dv) dv.style.display='none';
  Object.values(drillCharts).forEach(c=>c.destroy());
  drillCharts={};
}
