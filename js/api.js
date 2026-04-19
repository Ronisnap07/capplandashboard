function setSyncInfo(type,text){
  const el=document.getElementById('sync-info');
  el.className='sync-info '+type;
  el.textContent=text;
}

async function apiCall(action,data){
  let url=API+'?action='+action;
  if(data!==undefined){
    url+='&pwd='+encodeURIComponent(EDITOR_PASSWORD);
    url+='&data='+encodeURIComponent(JSON.stringify(data));
  }
  const res=await fetch(url);
  return res.json();
}

async function loadAllData(){
  const d=await apiCall('getAll');
  resources   = d.resources   || [];
  history     = d.history     || [];
  projections = d.projections || [];
  if(d.thresholds && d.thresholds.length){
    thresholds={...DEFAULT_THR};
    d.thresholds.forEach(t=>{ if(t.key) thresholds[t.key]=parseFloat(t.value)||DEFAULT_THR[t.key]; });
  }
  const now=new Date().toLocaleTimeString('id-ID',{hour:'2-digit',minute:'2-digit'});
  setSyncInfo('synced','Update: '+now);
}

async function loadAndRender(){
  setSyncInfo('syncing','Memuat...');
  try {
    await loadAllData();
    renderCurrentTab();
  } catch(e){ setSyncInfo('error','Gagal memuat'); }
}

async function syncData(action,data){
  setSyncInfo('syncing','Menyimpan...');
  try {
    const r=await apiCall(action,data);
    if(r.error) throw new Error(r.error);
    setSyncInfo('synced','Tersimpan ✓');
  } catch(e){ setSyncInfo('error','Gagal: '+e.message); }
}
