function getThr(r){
  if(r.type==='cluster') return {cpu:thresholds.cluster_cpu,ram:thresholds.cluster_ram};
  if(r.type==='storage_core') return {stor:thresholds.storage_core};
  return {stor:thresholds.storage_support};
}

function getLatest(resId,metric){
  const e=history.filter(h=>h.resId==resId&&h[metric]!==''&&h[metric]!=null)
    .sort((a,b)=>String(a.date).localeCompare(String(b.date)));
  return e.length?parseFloat(e[e.length-1][metric]):null;
}

function statusOf(v,t){
  if(v==null) return 'na';
  if(v>100) return 'over';
  if(v>t) return 'danger';
  if(v>t*0.85) return 'warn';
  return 'ok';
}

const SC={ok:'var(--accent3)',warn:'var(--warn)',danger:'var(--danger)',over:'var(--purple)',na:'var(--text-dim)'};
const SL={ok:'Aman',warn:'Warning',danger:'Kritis',over:'Over Cap',na:'No Data'};
const SAC={ok:'ab-ok',warn:'ab-warn',danger:'ab-danger',over:'ab-over',na:''};

function fmsg(id,text,color){
  const el=document.getElementById(id); if(!el) return;
  el.textContent=text; el.style.color=color;
  setTimeout(()=>el.textContent='',3500);
}

function getCapacity(r,metric){
  if(metric==='cpu') return r.cpuCap||null;
  if(metric==='ram') return r.ramCap||null;
  if(metric==='stor') return r.storCap||null;
  return null;
}

function getUnit(metric){ return metric==='cpu'?'cores':metric==='ram'?'GB':'TB'; }

function getActualAbsolute(resId,metric){
  const r=resources.find(x=>x.id==resId); if(!r) return null;
  const cap=getCapacity(r,metric); if(!cap) return null;
  const latestPct=getLatest(resId,metric); if(latestPct==null) return null;
  return parseFloat(((latestPct/100)*cap).toFixed(2));
}
