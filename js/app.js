function init(){
  const params  = new URLSearchParams(window.location.search);
  isEditor      = params.get('mode') === 'editor';
  const page    = params.get('page');
  const isAuth  = sessionStorage.getItem('editor_auth') === '1';

  if(!isEditor){
    buildViewerNav();
    loadAndRender();
    return;
  }

  // Setup header editor
  document.getElementById('header-sub').textContent = '// Editor';
  document.getElementById('mode-badge').className   = 'badge badge-editor';
  document.getElementById('mode-badge').textContent = '✏️ EDITOR ▾';

  if(!isAuth){
    document.getElementById('login-screen').style.display = 'flex';
    return;
  }

  // Sudah login
  if(page === 'config'){
    buildConfigPage();
  } else {
    buildEditorNav();
    loadAndRender();
  }

  document.addEventListener('click', e => {
    if(!document.getElementById('editor-menu-wrap').contains(e.target))
      document.getElementById('editor-dropdown').classList.remove('open');
  });
}

// ===== LOGIN =====
function doLogin(){
  const pwd = document.getElementById('pwd-input').value;
  if(!pwd){ document.getElementById('login-err').textContent = 'Masukkan password.'; return; }
  if(pwd !== EDITOR_PASSWORD){ document.getElementById('login-err').textContent = '✗ Password salah!'; return; }

  sessionStorage.setItem('editor_auth', '1');
  document.getElementById('login-screen').style.display = 'none';

  const page = new URLSearchParams(window.location.search).get('page');
  if(page === 'config'){
    buildConfigPage();
  } else {
    buildEditorNav();
    loadAndRender();
  }

  document.addEventListener('click', e => {
    if(!document.getElementById('editor-menu-wrap').contains(e.target))
      document.getElementById('editor-dropdown').classList.remove('open');
  });
}

// ===== EDITOR DROPDOWN =====
function toggleEditorMenu(){
  if(!isEditor) return;
  document.getElementById('editor-dropdown').classList.toggle('open');
}

function openConfig(){
  document.getElementById('editor-dropdown').classList.remove('open');
  window.location.href = '?mode=editor&page=config';
}

function logoutEditor(){
  sessionStorage.removeItem('editor_auth');
  window.location.href = '?mode=editor';
}

// ===== CONFIG PAGE =====
let currentConfigTab = 'c-resources';

function buildConfigPage(){
  document.getElementById('main-nav').innerHTML = `
    <button class="tab active" onclick="window.location.href='?mode=editor'">← Dashboard</button>`;

  document.getElementById('main-content').innerHTML = `
    <nav class="cpg-tabs">
      <button class="cpg-btn active" data-ctab="c-resources"  onclick="showConfigTab('c-resources')">Resources</button>
      <button class="cpg-btn"        data-ctab="c-history"    onclick="showConfigTab('c-history')">Update Utilisasi</button>
      <button class="cpg-btn"        data-ctab="c-projection" onclick="showConfigTab('c-projection')">Proyeksi Tahunan</button>
      <button class="cpg-btn"        data-ctab="c-threshold"  onclick="showConfigTab('c-threshold')">Threshold</button>
      <button class="cpg-btn"        data-ctab="c-report"     onclick="showConfigTab('c-report')">📊 Report Bulanan</button>
      <button class="cpg-btn"        data-ctab="c-report-quarterly" onclick="showConfigTab('c-report-quarterly')">📊 Report 3 Bulanan</button>
    </nav>

    <!-- Resources -->
    <div id="cfg-c-resources" class="cpg-view active">
      <div class="sec-title">Upload Bulk — Resources</div>
      <div class="card">
        <p style="font-size:13px;color:var(--text-dim);margin-bottom:12px">Upload file Excel (.xlsx) dengan kolom: <code style="color:var(--accent);background:var(--bg3);padding:2px 6px;border-radius:3px">type, name, cpuCap, ramCap, storCap, note</code></p>
        <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:10px">
          <button class="btn btn-sec" onclick="downloadTemplate('resources')">⬇ Download Template</button>
        </div>
        <div class="fg">
          <label>Pilih File Excel</label>
          <input type="file" id="csv-resources" accept=".xlsx,.xls" style="background:var(--bg3);border:1px solid var(--border);color:var(--text);padding:8px;border-radius:5px;width:100%">
        </div>
        <div style="margin-top:10px;">
          <button class="btn btn-primary" onclick="uploadExcel('resources')">⬆ Upload & Simpan</button>
        </div>
        <div class="fmsg" id="csv-res-msg"></div>
        <div id="csv-res-preview" style="margin-top:12px"></div>
      </div>
      <div class="sec-title">Tambah Resource Manual</div>
      <div class="card">
        <div class="form-grid">
          <div class="fg"><label>Tipe Resource *</label><select id="f-type" onchange="onTypeChange()"><option value="cluster">Host Cluster</option><option value="storage_core">Storage Core</option><option value="storage_support">Storage Support</option></select></div>
          <div class="fg"><label>Nama Resource *</label><input type="text" id="f-name" placeholder="Cth: Cluster-1"></div>
          <div class="fg" id="grp-cpu"><label>Kapasitas CPU (cores) *</label><input type="number" id="f-cpu" placeholder="128" min="1" step="any"></div>
          <div class="fg" id="grp-ram"><label>Kapasitas RAM (GB) *</label><input type="number" id="f-ram" placeholder="512" min="1" step="any"></div>
          <div class="fg" id="grp-stor" style="display:none"><label>Kapasitas Storage (TB) *</label><input type="number" id="f-stor" placeholder="50" min="0.1" step="any"></div>
          <div class="fg"><label>Keterangan</label><input type="text" id="f-note" placeholder="Opsional"></div>
        </div>
        <div style="display:flex;gap:8px;">
          <button class="btn btn-primary" onclick="addResource()">+ Tambahkan</button>
          <button class="btn btn-sec" onclick="clearResForm()">Reset</button>
        </div>
        <div class="fmsg" id="res-msg"></div>
      </div>
      <div class="sec-title">Daftar Resource</div>
      <div class="tbl-wrap"><div style="overflow-x:auto"><table>
        <thead><tr><th>#</th><th>Nama</th><th>Tipe</th><th>CPU</th><th>RAM</th><th>Storage</th><th>Keterangan</th><th>Aksi</th></tr></thead>
        <tbody id="res-tbody"></tbody>
      </table></div></div>
    </div>

    <!-- History -->
    <div id="cfg-c-history" class="cpg-view">
      <div class="sec-title">Upload Bulk — Utilisasi</div>
      <div class="card">
        <p style="font-size:13px;color:var(--text-dim);margin-bottom:12px">Upload file Excel (.xlsx) dengan kolom: <code style="color:var(--accent);background:var(--bg3);padding:2px 6px;border-radius:3px">resName, date, label, cpu, ram, stor</code></p>
        <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:10px">
          <button class="btn btn-sec" onclick="downloadTemplate('history')">⬇ Download Template</button>
        </div>
        <div class="fg">
          <label>Pilih File Excel</label>
          <input type="file" id="csv-history" accept=".xlsx,.xls" style="background:var(--bg3);border:1px solid var(--border);color:var(--text);padding:8px;border-radius:5px;width:100%">
        </div>
        <div style="margin-top:10px;">
          <button class="btn btn-primary" onclick="uploadExcel('history')">⬆ Upload & Simpan</button>
        </div>
        <div class="fmsg" id="csv-hist-msg"></div>
        <div id="csv-hist-preview" style="margin-top:12px"></div>
      </div>
      <div class="sec-title">Update Utilisasi Manual</div>
      <div class="card">
        <div class="form-grid">
          <div class="fg"><label>Resource *</label><select id="h-res" onchange="onHistResChange()"></select></div>
          <div class="fg"><label>Tanggal *</label><input type="date" id="h-date"></div>
          <div class="fg" id="grp-h-cpu"><label>Utilisasi CPU (%)</label><input type="number" id="h-cpu" placeholder="72.5" min="0" max="100" step="0.1"></div>
          <div class="fg" id="grp-h-ram"><label>Utilisasi RAM (%)</label><input type="number" id="h-ram" placeholder="85.2" min="0" max="100" step="0.1"></div>
          <div class="fg" id="grp-h-stor" style="display:none"><label>Utilisasi Storage (%)</label><input type="number" id="h-stor" placeholder="65.0" min="0" max="100" step="0.1"></div>
          <div class="fg"><label>Label Periode</label><input type="text" id="h-label" placeholder="Minggu 1 Jan"><div class="hint">Opsional, untuk label chart</div></div>
        </div>
        <div style="display:flex;gap:8px;">
          <button class="btn btn-primary" onclick="addHistory()">+ Simpan</button>
          <button class="btn btn-sec" onclick="clearHistForm()">Reset</button>
        </div>
        <div class="fmsg" id="hist-msg"></div>
      </div>
      <div class="sec-title">Riwayat Utilisasi</div>
      <div style="display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap;" id="hist-filter"></div>
      <div class="tbl-wrap"><div style="overflow-x:auto"><table>
        <thead><tr><th>#</th><th>Resource</th><th>Tanggal</th><th>Label</th><th>CPU%</th><th>RAM%</th><th>Storage%</th><th>Aksi</th></tr></thead>
        <tbody id="hist-tbody"></tbody>
      </table></div></div>
    </div>

    <!-- Projection -->
    <div id="cfg-c-projection" class="cpg-view">
      <div class="sec-title">Upload Bulk — Proyeksi Tahunan</div>
      <div class="card">
        <p style="font-size:13px;color:var(--text-dim);margin-bottom:12px">Upload file Excel (.xlsx) dengan kolom: <code style="color:var(--accent);background:var(--bg3);padding:2px 6px;border-radius:3px">resName, metric, year, projected</code></p>
        <p style="font-size:12px;color:var(--text-dim);margin-bottom:12px">Nilai <b>projected</b> dalam satuan absolut: CPU = cores, RAM = GB, Storage = TB</p>
        <div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:10px">
          <button class="btn btn-sec" onclick="downloadTemplate('projections')">⬇ Download Template</button>
        </div>
        <div class="fg">
          <label>Pilih File Excel</label>
          <input type="file" id="csv-projections" accept=".xlsx,.xls" style="background:var(--bg3);border:1px solid var(--border);color:var(--text);padding:8px;border-radius:5px;width:100%">
        </div>
        <div style="margin-top:10px;">
          <button class="btn btn-primary" onclick="uploadExcel('projections')">⬆ Upload & Simpan</button>
        </div>
        <div class="fmsg" id="csv-proj-msg"></div>
        <div id="csv-proj-preview" style="margin-top:12px"></div>
      </div>
      <div class="sec-title">Input Proyeksi Manual</div>
      <div class="card">
        <p style="font-size:13px;color:var(--text-dim);margin-bottom:14px;">Input proyeksi kapasitas per tahun dalam satuan <b style="color:var(--accent)">absolut</b> (CPU: cores, RAM: GB, Storage: TB).</p>
        <div class="form-grid" style="margin-bottom:14px;">
          <div class="fg"><label>Tahun</label><input type="number" id="proj-year" value="2026" min="2020" max="2035"></div>
        </div>
        <div class="proj-res-sel" id="proj-res-sel"></div>
        <div id="proj-metric-sel" style="display:flex;gap:8px;margin-bottom:14px;flex-wrap:wrap;"></div>
        <div id="proj-abs-input" style="display:none;margin-top:10px;">
          <div id="proj-current-info" style="margin-bottom:12px;padding:12px;background:var(--bg3);border:1px solid var(--border);border-radius:6px;font-size:12px;"></div>
          <div class="form-grid" style="grid-template-columns:repeat(auto-fill,minmax(220px,1fr))">
            <div class="fg">
              <label id="proj-abs-label">Nilai Proyeksi (absolut)</label>
              <input type="number" id="pm-abs" min="0" step="1" placeholder="Contoh: 204">
              <div class="hint" id="proj-abs-hint">Masukkan nilai dalam cores/GB/TB</div>
            </div>
          </div>
        </div>
        <div id="proj-list" style="margin-top:16px;"></div>
        <div style="display:flex;gap:8px;margin-top:16px;">
          <button class="btn btn-primary" onclick="saveProjection()">💾 Simpan</button>
          <button class="btn btn-sec" onclick="clearProjInputs()">Reset</button>
        </div>
        <div class="fmsg" id="proj-msg"></div>
      </div>
      <div class="sec-title">Daftar Proyeksi Tersimpan</div>
      <div class="tbl-wrap"><div style="overflow-x:auto"><table>
        <thead><tr><th>#</th><th>Resource</th><th>Metrik</th><th>Tahun</th><th>Nilai Proyeksi</th><th>Kapasitas</th><th>% Kapasitas</th><th>Realisasi</th><th>Akurasi</th><th>Aksi</th></tr></thead>
        <tbody id="proj-tbody"></tbody>
      </table></div></div>
    </div>

    <!-- Threshold -->
    <div id="cfg-c-threshold" class="cpg-view">
      <div class="sec-title">Threshold Settings</div>
      <div class="card">
        <p style="font-size:13px;color:var(--text-dim);margin-bottom:16px;">Threshold default sudah disesuaikan. Ubah sesuai kebutuhan.</p>
        <div class="thr-grid" id="thr-grid"></div>
        <div style="margin-top:16px;display:flex;gap:8px;">
          <button class="btn btn-primary" onclick="saveThresholds()">💾 Simpan</button>
          <button class="btn btn-sec" onclick="resetThresholds()">Reset Default</button>
        </div>
        <div class="fmsg" id="thr-msg"></div>
      </div>
    </div>

    <!-- Report Bulanan -->
    <div id="cfg-c-report" class="cpg-view">
      <div id="report-content"></div>
    </div>
    <div id="cfg-c-report-quarterly" class="cpg-view">
      <div id="report-content-quarterly"></div>
    </div>`;

  currentConfigTab = 'c-resources';
  loadAndRenderConfig();
}

function showConfigTab(name){
  currentConfigTab = name;
  document.querySelectorAll('.cpg-btn').forEach(b => b.classList.toggle('active', b.dataset.ctab === name));
  document.querySelectorAll('.cpg-view').forEach(v => v.classList.remove('active'));
  document.getElementById('cfg-'+name).classList.add('active');
  renderCurrentConfigTab();
}

async function loadAndRenderConfig(){
  setSyncInfo('syncing','Memuat...');
  try {
    await loadAllData();
    renderCurrentConfigTab();
  } catch(e){ setSyncInfo('error','Gagal memuat'); }
}

function renderCurrentConfigTab(){
  if(currentConfigTab==='c-resources'){ renderResTable(); }
  else if(currentConfigTab==='c-history'){ renderHistorySelects(); renderHistFilter(); renderHistTable(); }
  else if(currentConfigTab==='c-projection') renderProjUI();
  else if(currentConfigTab==='c-threshold')  renderThresholds();
  else if(currentConfigTab==='c-report')     renderReport();
  else if(currentConfigTab==='c-report-quarterly') renderReportQuarterly();
}

// ===== VIEWER / EDITOR NAV =====
function buildViewerNav(){
  document.getElementById('main-nav').innerHTML=`
    <button class="tab active" data-tab="overview"   onclick="showTab('overview')">Overview</button>
    <button class="tab" data-tab="detail"             onclick="showTab('detail')">Detail Resource</button>
    <button class="tab" data-tab="growth"             onclick="showTab('growth')">Pertumbuhan</button>
    <button class="tab" data-tab="projection"         onclick="showTab('projection')">Proyeksi & Akurasi</button>`;
  document.getElementById('main-content').innerHTML = viewerTabsHTML();
  currentTab = 'overview';
}

function buildEditorNav(){
  document.getElementById('main-nav').innerHTML=`
    <button class="tab active" data-tab="overview"   onclick="showTab('overview')">Overview</button>
    <button class="tab" data-tab="detail"             onclick="showTab('detail')">Detail Resource</button>
    <button class="tab" data-tab="growth"             onclick="showTab('growth')">Pertumbuhan</button>
    <button class="tab" data-tab="projection"         onclick="showTab('projection')">Proyeksi & Akurasi</button>`;
  document.getElementById('main-content').innerHTML = viewerTabsHTML();
  currentTab = 'overview';
}

function viewerTabsHTML(){
  return `
    <div id="tab-overview" class="view active">
      <div class="sec-title">Status Utilisasi Saat Ini</div>
      <div class="gauge-grid" id="gauge-grid"></div>
      <div class="sec-title">Summary</div>
      <div class="sum-grid" id="sum-grid"></div>
      <div class="chart-grid">
        <div class="chart-card full">
          <div class="chart-title">Utilisasi Aktual — Semua Resource</div>
          <canvas id="chart-overview" height="90"></canvas>
        </div>
      </div>
    </div>
    <div id="tab-detail" class="view">
      <div id="detail-list-view">
        <div class="sec-title">Resource Detail</div>
        <div class="toggle-wrap">
          <button class="toggle-btn active" onclick="setDetailView('list')">List</button>
          <button class="toggle-btn" onclick="setDetailView('group')">Per Kategori</button>
        </div>
        <div id="detail-list-content"></div>
      </div>
      <div id="detail-drill-view" style="display:none;">
        <div class="detail-back" onclick="backToList()">← Kembali</div>
        <div id="drill-content"></div>
      </div>
    </div>
    <div id="tab-growth" class="view">
      <div id="growth-list-view">
        <div class="sec-title">Pertumbuhan Utilisasi</div>
        <div id="growth-grid"></div>
      </div>
      <div id="growth-drill-view" style="display:none;">
        <div class="detail-back" onclick="backToGrowth()">← Kembali</div>
        <div id="growth-drill-content"></div>
      </div>
    </div>
    <div id="tab-projection" class="view">
      <div class="sec-title">Proyeksi Tahunan & Akurasi Forecasting</div>
      <div id="proj-content"></div>
    </div>`;
}

// ===== TAB NAVIGATION =====
function showTab(name){
  document.querySelectorAll('.view').forEach(v=>v.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  document.getElementById('tab-'+name).classList.add('active');
  const btn=document.querySelector(`.tab[data-tab="${name}"]`);
  if(btn) btn.classList.add('active');
  currentTab=name;
  renderCurrentTab();
}

function renderCurrentTab(){
  if(currentTab==='overview')        renderOverview();
  else if(currentTab==='detail')     renderDetailList();
  else if(currentTab==='growth')     renderGrowth();
  else if(currentTab==='projection') renderProjection();
}

init();
