/* ==========================================================================
   BPCL RO-WISE TANK STOCK MONITORING & GOOGLE DRIVE SYNC ENGINE
   ========================================================================== */

(function() {
  'use strict';

  const FILES = {
    'Stock hourly.xlsx': 'hourly',
    'Onboarded tanks.xlsx': 'onboarded',
    'Last updated stock.xlsx': 'latest'
  };

  let data = { hourly: [], onboarded: [], latest: [] };
  let folderHandle = null;
  let refreshTimer = null;
  let selectedRO = null;
  let allRO = [];
  let fileTimestamps = {};

  const $ = id => document.getElementById(id);
  const norm = v => String(v ?? '').trim();
  const key = v => norm(v).toUpperCase();
  const num = v => {
    if (v === null || v === undefined || v === '') return null;
    let n = Number(String(v).replace(/,/g, ''));
    return Number.isFinite(n) ? n : null;
  };
  const fmt = (v, d = 0) => v == null ? '—' : Number(v).toLocaleString('en-IN', { maximumFractionDigits: d, minimumFractionDigits: d });
  const esc = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  // Parse Excel ArrayBuffer into JSON
  function parseWorkbook(file, type) {
    return file.arrayBuffer().then(buf => {
      const wb = XLSX.read(buf, { type: 'array', cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      return XLSX.utils.sheet_to_json(ws, { defval: null, raw: true }).map(src => {
        const r = {};
        Object.keys(src).forEach(k => r[String(k).trim()] = src[k]);
        if (type === 'hourly') {
          return { ...r, _ro: key(r['RO ID']), _tank: key(r['Tank ID']), _product: key(r['Product Name']) };
        }
        if (type === 'onboarded') {
          return { ...r, _ro: key(r.ROID || r['RO ID'] || r.RO_ID), _tank: key(r['Local tank id'] || r.Tank_ID), _product: key(r.Product_Name || r['Product Name']) };
        }
        return { ...r, _ro: key(r['RO ID'] || r.ROID), _product: key(r['Product Name'] || r.Product_Name) };
      });
    });
  }

  // Load files into data model
  async function loadFiles(fileMap, forceMapping = false) {
    try {
      if (!window.XLSX) throw Error('Excel library (xlsx.full.min.js) not ready.');
      let parsed = {};

      for (const name of ['Stock hourly.xlsx', 'Last updated stock.xlsx']) {
        if (!fileMap[name]) throw Error('Missing required file in folder: ' + name);
        parsed[FILES[name]] = await parseWorkbook(fileMap[name], FILES[name]);
      }

      let cached = null;
      try {
        cached = JSON.parse(localStorage.getItem('bpcl_tank_mapping_v1') || 'null');
      } catch (e) {}

      if (fileMap['Onboarded tanks.xlsx'] && (forceMapping || !cached)) {
        parsed.onboarded = await parseWorkbook(fileMap['Onboarded tanks.xlsx'], 'onboarded');
        try {
          localStorage.setItem('bpcl_tank_mapping_v1', JSON.stringify(parsed.onboarded));
        } catch (e) {}
      } else {
        parsed.onboarded = cached || data.onboarded || [];
      }

      data.hourly = parsed.hourly || [];
      data.latest = parsed.latest || [];
      data.onboarded = parsed.onboarded || [];

      // Cache current stock in localStorage for instant offline access
      try {
        localStorage.setItem('bpcl_tank_stock_cache', JSON.stringify({
          hourly: data.hourly.slice(0, 5000), // safe subset
          latest: data.latest,
          savedAt: Date.now()
        }));
      } catch(e){}

      renderDashboard();
      
      const statusEl = $('ts-data-status');
      if (statusEl) {
        statusEl.innerHTML = `<span class="pill">● Google Drive Live Sync</span> Loaded <b>${data.hourly.length.toLocaleString()}</b> hourly records · <b>${data.onboarded.length.toLocaleString()}</b> tanks · <b>${data.latest.length.toLocaleString()}</b> stock records`;
      }
      const updatedEl = $('ts-updated-at');
      if (updatedEl) {
        updatedEl.textContent = 'Updated: ' + new Date().toLocaleTimeString();
      }

      if (window.showToast) {
        window.showToast('Tank stock details updated successfully from Google Drive!', 'success');
      }
    } catch (e) {
      console.error('Tank stock load error:', e);
      const statusEl = $('ts-data-status');
      if (statusEl) statusEl.innerHTML = `<span class="pill off">Error</span> ${e.message}`;
      if (window.showToast) window.showToast('Error reading stock: ' + e.message, 'error');
    }
  }

  // Google Drive folder selection using File System Access API
  async function chooseGoogleDriveFolder() {
    if (!window.showDirectoryPicker) {
      alert('Local Google Drive folder direct access is supported in Chrome or Edge.\n\nPlease use "Upload Excel Files" button to select the workbooks, or open over HTTPS/localhost.');
      $('ts-file-input')?.click();
      return;
    }
    try {
      folderHandle = await window.showDirectoryPicker({ mode: 'read' });
      const folderName = folderHandle.name || 'Google Drive Folder';
      if ($('ts-folder-name')) $('ts-folder-name').textContent = folderName;
      await readDriveFolder(false, true);

      // Start live auto-polling every 30 seconds
      if (refreshTimer) clearInterval(refreshTimer);
      refreshTimer = setInterval(() => readDriveFolder(true), 30000);
    } catch (e) {
      if (e.name !== 'AbortError') {
        alert('Could not access Google Drive folder: ' + e.message);
      }
    }
  }

  // Read folder and detect changes
  async function readDriveFolder(silent = false, forceMapping = false) {
    if (!folderHandle) return;
    try {
      const map = {};
      let hasChanges = false;

      for (const name of ['Stock hourly.xlsx', 'Last updated stock.xlsx']) {
        const h = await folderHandle.getFileHandle(name);
        const file = await h.getFile();
        map[name] = file;
        if (!fileTimestamps[name] || fileTimestamps[name] !== file.lastModified) {
          fileTimestamps[name] = file.lastModified;
          hasChanges = true;
        }
      }

      let hasMapping = false;
      try { hasMapping = !!localStorage.getItem('bpcl_tank_mapping_v1'); } catch (e) {}

      try {
        const h = await folderHandle.getFileHandle('Onboarded tanks.xlsx');
        const file = await h.getFile();
        map['Onboarded tanks.xlsx'] = file;
        if (!fileTimestamps['Onboarded tanks.xlsx'] || fileTimestamps['Onboarded tanks.xlsx'] !== file.lastModified) {
          fileTimestamps['Onboarded tanks.xlsx'] = file.lastModified;
          hasChanges = true;
        }
      } catch(e) {}

      if (hasChanges || forceMapping || !data.latest.length) {
        await loadFiles(map, forceMapping);
      }
    } catch (e) {
      if (!silent) {
        alert('Could not read the Google Drive folder. Ensure "Stock hourly.xlsx", "Last updated stock.xlsx", and "Onboarded tanks.xlsx" are present.\n\n' + e.message);
      } else {
        console.warn('Auto drive check failed:', e);
      }
    }
  }

  // Group data by RO
  function groupByRO() {
    const m = new Map();
    const ensure = (id, row) => {
      if (!m.has(id)) {
        m.set(id, {
          id,
          name: row['Retail Outlet'] || row.RO_Name || row.ROName || 'Outlet ' + id,
          territory: row.Territory || 'Belgaum',
          area: row['Sales Area'] || row.Salesarea || row.SalesArea || '',
          vendor: row['RA Vendor'] || row.RA_Vendor || row.Vendor || '',
          state: row.State || 'Karnataka',
          ebms: 0,
          hsd: 0,
          total: 0,
          tanks: [],
          latest: [],
          online: 0,
          offline: 0
        });
      }
      return m.get(id);
    };

    data.onboarded.forEach(r => {
      if (!r._ro) return;
      let x = ensure(r._ro, r);
      x.tanks.push(r);
      let st = key(r.Status);
      if (st === 'ONLINE') x.online++;
      else if (st === 'OFFLINE') x.offline++;
    });

    data.latest.forEach(r => {
      if (!r._ro) return;
      let x = ensure(r._ro, r);
      x.latest.push(r);
      let p = key(r['Product Name'] || r.Product_Name);
      let stock = num(r.Stock) || 0;
      if (p.includes('EBMS')) x.ebms += stock;
      else if (p === 'HSD') x.hsd += stock;
      x.total += stock;
    });

    data.hourly.forEach(r => {
      if (!r._ro) return;
      ensure(r._ro, r);
    });

    return [...m.values()].sort((a, b) => a.name.localeCompare(b.name));
  }

  function fillFilters() {
    const fields = [
      ['ts-territory', 'territory'],
      ['ts-area', 'area'],
      ['ts-vendor', 'vendor'],
      ['ts-state', 'state']
    ];
    fields.forEach(([id, k]) => {
      const el = $(id);
      if (!el) return;
      const old = el.value;
      const vals = [...new Set(allRO.map(r => r[k]).filter(Boolean))].sort();
      el.innerHTML = `<option value="">All ${k.charAt(0).toUpperCase() + k.slice(1)}s</option>` +
        vals.map(v => `<option value="${esc(v)}" ${v === old ? 'selected' : ''}>${esc(v)}</option>`).join('');
    });
  }

  function renderDashboard() {
    allRO = groupByRO();
    fillFilters();
    renderKPIs();
    drawROTable();
  }

  function renderKPIs() {
    const kpisEl = $('ts-kpis');
    if (!kpisEl) return;
    const totalOutlets = allRO.length;
    const totalTanks = allRO.reduce((s, r) => s + r.tanks.length, 0);
    const onlineTanks = allRO.reduce((s, r) => s + r.online, 0);
    const offlineTanks = allRO.reduce((s, r) => s + r.offline, 0);
    const totalStock = allRO.reduce((s, r) => s + r.total, 0);

    kpisEl.innerHTML = `
      <div class="kpi">
        <div class="klabel">RETAIL OUTLETS</div>
        <div class="value">${fmt(totalOutlets)}</div>
        <div class="sub">Reporting Outlets</div>
      </div>
      <div class="kpi">
        <div class="klabel">ONBOARDED TANKS</div>
        <div class="value">${fmt(totalTanks)}</div>
        <div class="sub">Total Calibrated Tanks</div>
      </div>
      <div class="kpi">
        <div class="klabel">ONLINE TANKS</div>
        <div class="value text-green">${fmt(onlineTanks)}</div>
        <div class="sub">${totalTanks ? fmt((onlineTanks / totalTanks) * 100, 1) + '% active' : '—'}</div>
      </div>
      <div class="kpi">
        <div class="klabel">OFFLINE TANKS</div>
        <div class="value text-red">${fmt(offlineTanks)}</div>
        <div class="sub">Requiring attention</div>
      </div>
      <div class="kpi">
        <div class="klabel">TOTAL STOCK (KL)</div>
        <div class="value text-navy">${fmt(totalStock, 1)}</div>
        <div class="sub">Combined EBMS + HSD</div>
      </div>
    `;
  }

  function drawROTable() {
    const q = key($('ts-search')?.value);
    const territory = key($('ts-territory')?.value);
    const area = key($('ts-area')?.value);
    const vendor = key($('ts-vendor')?.value);
    const state = key($('ts-state')?.value);

    const filtered = allRO.filter(r => {
      if (q && !r.id.includes(q) && !key(r.name).includes(q)) return false;
      if (territory && key(r.territory) !== territory) return false;
      if (area && key(r.area) !== area) return false;
      if (vendor && key(r.vendor) !== vendor) return false;
      if (state && key(r.state) !== state) return false;
      return true;
    });

    const countEl = $('ts-row-count');
    if (countEl) countEl.textContent = `${filtered.length} of ${allRO.length} outlets`;

    const tbody = $('ts-ro-table');
    if (!tbody) return;

    if (!filtered.length) {
      tbody.innerHTML = `<tr><td colspan="9" class="empty">No matching outlets found. Select folder or adjust filters.</td></tr>`;
      return;
    }

    tbody.innerHTML = filtered.map(r => `
      <tr onclick="window.showRODetail('${esc(r.id)}')">
        <td><b>${esc(r.id)}</b></td>
        <td class="ro-name">${esc(r.name)}</td>
        <td>${esc(r.area || r.territory || '—')}</td>
        <td><span class="tag">${esc(r.vendor || '—')}</span></td>
        <td class="num product-ebms">${r.ebms ? fmt(r.ebms, 2) : '—'}</td>
        <td class="num product-hsd">${r.hsd ? fmt(r.hsd, 2) : '—'}</td>
        <td class="num"><b>${fmt(r.total, 2)}</b></td>
        <td class="num">${r.tanks.length || '—'}</td>
        <td>
          <button class="btn btn-sm" onclick="event.stopPropagation(); window.showRODetail('${esc(r.id)}')">
            View Tanks ➔
          </button>
        </td>
      </tr>
    `).join('');
  }

  // Show detailed tank view for an RO
  window.showRODetail = function(id) {
    selectedRO = allRO.find(r => r.id === id);
    if (!selectedRO) return;

    $('ts-dashboard-view')?.classList.add('hidden');
    $('ts-detail-view')?.classList.remove('hidden');

    $('ts-ro-title').textContent = `${selectedRO.name} (${selectedRO.id})`;
    $('ts-ro-subtitle').textContent = `${selectedRO.area || selectedRO.territory} · Vendor: ${selectedRO.vendor || 'N/A'} · State: ${selectedRO.state || 'Karnataka'}`;

    // Product summaries
    const prods = ['EBMS', 'HSD'];
    $('ts-product-summaries').innerHTML = prods.map(p => {
      let recs = selectedRO.latest.filter(r => key(r['Product Name'] || r.Product_Name) === p);
      let capacity = recs.reduce((s, r) => s + (num(r['On-boarded Tankage'] || r['On-Boarded Tankage (KL)']) || 0), 0);
      let stock = recs.reduce((s, r) => s + (num(r.Stock) || 0), 0);
      let onlineCap = recs.reduce((s, r) => s + (num(r['On-Line Tankage'] || r['Online Tankage']) || 0), 0);
      let ullage = Math.max(0, capacity - stock);
      let net = capacity ? (stock / capacity) * 100 : 0;

      return `
        <div class="summarybox">
          <h3>STOCK SUMMARY — <span class="product-${p.toLowerCase()}">${p}</span></h3>
          <table>
            <thead>
              <tr>
                <th>Product</th>
                <th class="num">Capacity (KL)</th>
                <th class="num">Stock (KL)</th>
                <th class="num">Ullage (KL)</th>
                <th class="num">Stock %</th>
                <th class="num">Ullage %</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td class="product-${p.toLowerCase()}"><b>${p}</b></td>
                <td class="num">${fmt(capacity, 2)}</td>
                <td class="num">${fmt(stock, 2)}</td>
                <td class="num">${fmt(ullage, 2)}</td>
                <td class="num"><b>${capacity ? fmt(net, 1) + '%' : '—'}</b></td>
                <td class="num">${capacity ? fmt(100 - net, 1) + '%' : '—'}</td>
              </tr>
            </tbody>
          </table>
          <div class="small" style="padding: 6px 14px 10px; color: #64748b;">
            On-line tankage: ${fmt(onlineCap, 2)} KL · ${recs.length} record(s)
          </div>
        </div>
      `;
    }).join('');

    // Tank cards
    let tanks = selectedRO.tanks.slice().sort((a, b) => (num(a['Local tank id']) || 0) - (num(b['Local tank id']) || 0));
    const tankCardsEl = $('ts-tank-cards');
    if (tankCardsEl) {
      tankCardsEl.innerHTML = tanks.length ? tanks.map(t => tankCard(t, selectedRO)).join('') : '<div class="empty">No onboarded tank records found for this RO.</div>';
      
      tankCardsEl.querySelectorAll('.tank[data-tank]').forEach(el => {
        el.addEventListener('click', () => showHourly(id, el.dataset.tank, el.dataset.product));
      });
    }
  };

  function tankCard(t, ro) {
    let id = key(t['Local tank id'] || t.Tank_ID);
    let p = key(t.Product_Name || t['Product Name']);
    let hr = data.hourly.find(h => h._ro === ro.id && h._tank === id && h._product === p);
    let hours = ['12 AM','1 AM','2 AM','3 AM','4 AM','5 AM','6 AM','7 AM','8 AM','9 AM','10 AM','11 AM','12 PM','1 PM','2 PM','3 PM','4 PM','5 PM','6 PM','7 PM','8 PM','9 PM','10 PM','11 PM'];
    let vals = hr ? hours.map(h => num(hr[h + ' %'])) : [];
    let valid = vals.filter(v => v != null);
    let current = valid.length ? valid[valid.length - 1] : null;
    let first = valid.length ? valid[0] : null;
    let capacity = num(t['On-Boarded Tankage (KL)'] || t['On-boarded Tankage']) ?? (hr ? num(hr['On-Boarded Tankage (KL)']) : null);
    let status = key(t.Status);
    let up = String(t['Uptime_tank(%)'] ?? '—');
    let color = p === 'EBMS' ? '#16a34a' : p === 'HSD' ? '#0284c7' : '#d97706';

    return `
      <div class="tank" data-tank="${esc(id)}" data-product="${esc(p)}">
        <div class="tanktop">
          <div class="tanktitle">Tank ${esc(id)} <small>(${fmt(capacity, 2)} KL)</small></div>
          <div>${status === 'ONLINE' ? '🟢' : '🔴'} <span class="tankproduct product-${p.toLowerCase()}">${esc(p)}</span></div>
        </div>
        <div class="tankmeta">
          <span>◈ ${esc(up)}% uptime</span>
          <span class="tag">${esc(t['Morning_Status'] || '—')} / ${esc(t['Evening_Status'] || '—')}</span>
        </div>
        <div class="gauge" style="--pct:${Math.max(0, Math.min(100, current ?? 0))};--gauge:${color}">
          <span>${current == null ? '—' : fmt(current, 1) + '%'}</span>
        </div>
        <div class="tankstats">
          <div><div class="statlabel">First reported</div><div class="statvalue">${first == null ? '—' : fmt(first, 1) + '%'}</div></div>
          <div style="text-align:right"><div class="statlabel">Latest hourly</div><div class="statvalue">${current == null ? '—' : fmt(current, 1) + '%'}</div></div>
          <div><div class="statlabel">Capacity</div><div class="statvalue">${fmt(capacity, 2)} KL</div></div>
          <div style="text-align:right"><div class="statlabel">Online hrs</div><div class="statvalue">${fmt(num(t.Online_hours))} / 24</div></div>
        </div>
        <div class="tankbottom">
          <div><div class="statlabel">Updated</div><div class="statvalue">${esc(t.Table_updated_at || 'Today')}</div></div>
          <div style="text-align:right"><span class="pill ${status === 'OFFLINE' ? 'off' : status === 'ONLINE' ? '' : 'partial'}">${esc(t.Status || 'Active')}</span></div>
        </div>
      </div>
    `;
  }

  function showHourly(roId, tankId, product) {
    let row = data.hourly.find(h => h._ro === roId && h._tank === key(tankId) && h._product === key(product));
    const hourlyPanel = $('ts-hourly-panel');
    if (!hourlyPanel) return;

    hourlyPanel.classList.remove('hidden');
    $('ts-hourly-title').textContent = `HOURLY STOCK TREND — TANK ${tankId} (${product})`;
    $('ts-hourly-subtitle').textContent = `24-Hour Stock Level Progression · RO: ${selectedRO ? selectedRO.name : roId}`;

    let hours = ['12 AM','1 AM','2 AM','3 AM','4 AM','5 AM','6 AM','7 AM','8 AM','9 AM','10 AM','11 AM','12 PM','1 PM','2 PM','3 PM','4 PM','5 PM','6 PM','7 PM','8 PM','9 PM','10 PM','11 PM'];
    let vals = hours.map(h => row ? num(row[h + ' %']) : null);
    let max = Math.max(100, ...vals.filter(v => v != null));

    $('ts-hourly-chart').innerHTML = vals.map((v, i) => `
      <div class="barwrap" title="${hours[i]}: ${v == null ? 'No data' : fmt(v, 2) + '%'}">
        <small>${v == null ? '' : fmt(v, 0)}</small>
        <div class="bar" style="height:${v == null ? 0 : Math.max(3, (v / max) * 140)}px;background:${key(product) === 'EBMS' ? '#16a34a' : '#0284c7'}"></div>
        <label>${hours[i].replace(' ', '')}</label>
      </div>
    `).join('');

    hourlyPanel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }

  // Back to summary
  function backToSummary() {
    selectedRO = null;
    $('ts-detail-view')?.classList.add('hidden');
    $('ts-dashboard-view')?.classList.remove('hidden');
  }

  // Preload local Excel files if present in workspace
  async function tryPreloadLocalFiles() {
    try {
      const map = {};
      const fileNames = ['Stock hourly.xlsx', 'Last updated stock.xlsx', 'Onboarded tanks.xlsx'];
      let foundCount = 0;

      for (const name of fileNames) {
        try {
          const resp = await fetch(name);
          if (resp.ok) {
            const blob = await resp.blob();
            map[name] = blob;
            foundCount++;
          }
        } catch(e) {}
      }

      if (foundCount >= 2) {
        await loadFiles(map);
      }
    } catch(e) {
      console.warn('Local prefetch error:', e);
    }
  }

  // Initialize Tank Stock View
  window.initTankStockDashboard = function() {
    const folderBtn = $('ts-folder-btn');
    const refreshBtn = $('ts-refresh-btn');
    const uploadBtn = $('ts-upload-btn');
    const fileInput = $('ts-file-input');
    const backBtn = $('ts-back-btn');
    const closeHourly = $('ts-close-hourly');

    if (folderBtn && !folderBtn._bound) {
      folderBtn._bound = true;
      folderBtn.addEventListener('click', chooseGoogleDriveFolder);
    }
    if (refreshBtn && !refreshBtn._bound) {
      refreshBtn._bound = true;
      refreshBtn.addEventListener('click', () => {
        if (folderHandle) readDriveFolder();
        else fileInput?.click();
      });
    }
    if (uploadBtn && !uploadBtn._bound) {
      uploadBtn._bound = true;
      uploadBtn.addEventListener('click', () => fileInput?.click());
    }
    if (fileInput && !fileInput._bound) {
      fileInput._bound = true;
      fileInput.addEventListener('change', async e => {
        let map = {};
        for (const f of e.target.files) {
          if (FILES[f.name]) map[f.name] = f;
        }
        await loadFiles(map, true);
        e.target.value = '';
      });
    }
    if (backBtn && !backBtn._bound) {
      backBtn._bound = true;
      backBtn.addEventListener('click', backToSummary);
    }
    if (closeHourly && !closeHourly._bound) {
      closeHourly._bound = true;
      closeHourly.addEventListener('click', () => $('ts-hourly-panel')?.classList.add('hidden'));
    }

    ['ts-search', 'ts-territory', 'ts-area', 'ts-vendor', 'ts-state'].forEach(id => {
      const el = $(id);
      if (el && !el._bound) {
        el._bound = true;
        el.addEventListener(id === 'ts-search' ? 'input' : 'change', drawROTable);
      }
    });

    // If data not loaded yet, try loading local workspace files
    if (!data.latest.length) {
      tryPreloadLocalFiles();
    }
  };

  // Run on DOM load
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', window.initTankStockDashboard);
  } else {
    window.initTankStockDashboard();
  }

})();
