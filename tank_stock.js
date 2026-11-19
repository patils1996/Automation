/* ==========================================================================
   BPCL RO-WISE TANK STOCK MONITORING & GOOGLE DRIVE SYNC ENGINE
   Exact Implementation matching BPCL_Tank_Stock_Dashboard (2).html
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

  // Scoped element selector to #tab-tank-stock to prevent collisions
  const $ = id => document.querySelector('#tab-tank-stock #' + id) || document.getElementById(id);
  const norm = v => String(v ?? '').trim();
  const key = v => norm(v).toUpperCase();
  const num = v => {
    if (v === null || v === undefined || v === '') return null;
    let n = Number(String(v).replace(/,/g, ''));
    return Number.isFinite(n) ? n : null;
  };
  const fmt = (v, d = 0) => v == null ? '—' : Number(v).toLocaleString('en-IN', { maximumFractionDigits: d, minimumFractionDigits: d });
  const pct = v => v == null ? '—' : fmt(v * 100, 1) + '%';
  function esc(s) {
    return String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }

  function parseWorkbook(file, type) {
    return file.arrayBuffer().then(buf => {
      const wb = XLSX.read(buf, { type: 'array', cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      return XLSX.utils.sheet_to_json(ws, { defval: null, raw: true }).map(src => {
        const r = {};
        Object.keys(src).forEach(k => r[String(k).trim()] = src[k]);
        if (type === 'hourly') return { ...r, _ro: key(r['RO ID']), _tank: key(r['Tank ID']), _product: key(r['Product Name']) };
        if (type === 'onboarded') return { ...r, _ro: key(r.ROID), _tank: key(r['Local tank id']), _product: key(r.Product_Name) };
        return { ...r, _ro: key(r['RO ID']), _product: key(r['Product Name']) };
      });
    });
  }

  async function loadFiles(fileMap, forceMapping = false) {
    try {
      if (!window.XLSX) throw Error('Excel reader could not load. Check internet access and reload the page.');
      let parsed = {};
      for (const name of ['Stock hourly.xlsx', 'Last updated stock.xlsx']) {
        if (!fileMap[name]) throw Error('Missing file: ' + name);
        parsed[FILES[name]] = await parseWorkbook(fileMap[name], FILES[name]);
      }
      let cached = null;
      try { cached = JSON.parse(localStorage.getItem('bpcl_tank_mapping_v1') || 'null'); } catch (e) {}
      if (fileMap['Onboarded tanks.xlsx'] && (forceMapping || !cached)) {
        parsed.onboarded = await parseWorkbook(fileMap['Onboarded tanks.xlsx'], 'onboarded');
        try { localStorage.setItem('bpcl_tank_mapping_v1', JSON.stringify(parsed.onboarded)); } catch (e) { console.warn('Tank mapping could not be cached', e); }
      } else {
        parsed.onboarded = cached || data.onboarded || [];
      }
      data.hourly = parsed.hourly;
      data.latest = parsed.latest;
      data.onboarded = parsed.onboarded;

      try {
        localStorage.setItem('bpcl_tank_stock_cache', JSON.stringify({
          hourly: data.hourly,
          latest: data.latest,
          savedAt: Date.now()
        }));
      } catch (e) {}

      renderDashboard();
      if ($('dataStatus')) {
        $('dataStatus').textContent = `Loaded ${data.hourly.length.toLocaleString()} hourly rows • ${data.onboarded.length.toLocaleString()} fixed tank mappings • ${data.latest.length.toLocaleString()} latest stock rows`;
      }
      if ($('updatedAt')) {
        $('updatedAt').textContent = 'Stock refreshed at ' + new Date().toLocaleString() + (forceMapping ? ' • Tank mapping updated' : '');
      }
      if (window.showToast) {
        window.showToast('Tank stock updated successfully!', 'success');
      }
    } catch (e) {
      if ($('dataStatus')) $('dataStatus').innerHTML = '<span class="error">' + e.message + '</span>';
      console.error(e);
      if (window.showToast) window.showToast(e.message, 'error');
    }
  }

  async function chooseFolder() {
    if (!window.showDirectoryPicker) {
      alert('Folder access is not supported in this browser context. Use Upload Excel files, or open this page in current Chrome/Edge over localhost or HTTPS.');
      $('tankStockFileInput')?.click();
      return;
    }
    try {
      folderHandle = await window.showDirectoryPicker({ mode: 'read' });
      await readFolder();
      if (refreshTimer) clearInterval(refreshTimer);
      refreshTimer = setInterval(() => readFolder(true), 30000); // 30s auto-refresh
    } catch (e) {
      if (e.name !== 'AbortError') alert('Could not access folder: ' + e.message);
    }
  }

  async function readFolder(silent = false, forceMapping = false) {
    if (!folderHandle) return;
    try {
      const map = {};
      let hasChanges = false;
      for (const name of ['Stock hourly.xlsx', 'Last updated stock.xlsx']) {
        const h = await folderHandle.getFileHandle(name);
        const f = await h.getFile();
        map[name] = f;
        if (!fileTimestamps[name] || fileTimestamps[name] !== f.lastModified) {
          fileTimestamps[name] = f.lastModified;
          hasChanges = true;
        }
      }
      let hasMapping = false;
      try { hasMapping = !!localStorage.getItem('bpcl_tank_mapping_v1'); } catch (e) {}
      if (forceMapping || !hasMapping) {
        const h = await folderHandle.getFileHandle('Onboarded tanks.xlsx');
        const f = await h.getFile();
        map['Onboarded tanks.xlsx'] = f;
        if (!fileTimestamps['Onboarded tanks.xlsx'] || fileTimestamps['Onboarded tanks.xlsx'] !== f.lastModified) {
          fileTimestamps['Onboarded tanks.xlsx'] = f.lastModified;
          hasChanges = true;
        }
      }
      if (hasChanges || forceMapping || !data.latest.length) {
        await loadFiles(map, forceMapping);
      }
    } catch (e) {
      if (!silent) alert('Could not read the selected folder. Confirm that the required Excel files are present. ' + e.message);
      else console.warn(e);
    }
  }

  function groupByRO() {
    const m = new Map();
    const ensure = (id, row) => {
      if (!m.has(id)) m.set(id, { id, name: row['Retail Outlet'] || row.RO_Name || 'Unknown RO', territory: row.Territory || '', area: row['Sales Area'] || row.Salesarea || row.Salesarea || '', vendor: row['RA Vendor'] || row.RA_Vendor || '', state: row.State || '', ebms: 0, hsd: 0, total: 0, tanks: [], latest: [], online: 0, offline: 0 });
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
      let p = key(r['Product Name']);
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
    const fields = [['territory', 'territory'], ['area', 'area'], ['vendor', 'vendor'], ['state', 'state']];
    fields.forEach(([id, k]) => {
      const el = $(id);
      if (!el) return;
      const old = el.value;
      const vals = [...new Set(allRO.map(r => r[k]).filter(Boolean))].sort();
      el.innerHTML = `<option value="">All ${k.charAt(0).toUpperCase() + k.slice(1)}s</option>` + vals.map(v => `<option value="${esc(v)}"${v === old ? ' selected' : ''}>${esc(v)}</option>`).join('');
    });
  }

  function renderDashboard() {
    allRO = groupByRO();
    fillFilters();
    let totalT = 0, on = 0, off = 0;
    allRO.forEach(r => { totalT += r.tanks.length; on += r.online; off += r.offline; });
    const kpisEl = $('kpis');
    if (kpisEl) {
      kpisEl.innerHTML = [
        ['RETAIL OUTLETS', allRO.length, ''],
        ['ONBOARDED TANKS', totalT, ''],
        ['ONLINE TANKS', on, totalT ? `${(on / totalT * 100).toFixed(1)}%` : ''],
        ['OFFLINE TANKS', off, ''],
        ['LATEST STOCK RECORDS', data.latest.length, '']
      ].map(([label, val, sub]) => `
        <div class="kpi">
          <div class="klabel">${label}</div>
          <div class="value">${Number(val || 0).toLocaleString('en-IN')}</div>
          ${sub ? `<div class="sub">${sub}</div>` : ''}
        </div>
      `).join('');
    }
    drawROTable();
  }

  function drawROTable() {
    const searchEl = $('search'), terrEl = $('territory'), areaEl = $('area'), vendEl = $('vendor'), stateEl = $('state');
    if (!searchEl) return;
    const q = key(searchEl.value), territory = terrEl ? terrEl.value : '', area = areaEl ? areaEl.value : '', vendor = vendEl ? vendEl.value : '', state = stateEl ? stateEl.value : '';
    let rows = allRO.filter(r => (!q || r.id.includes(q) || key(r.name).includes(q)) && (!territory || r.territory === territory) && (!area || r.area === area) && (!vendor || r.vendor === vendor) && (!state || r.state === state));
    if ($('rowCount')) $('rowCount').textContent = `${rows.length.toLocaleString()} outlets`;
    const tbody = $('roTable');
    if (!tbody) return;
    tbody.innerHTML = rows.map(r => {
      let st = r.offline === 0 && r.online > 0 ? 'ONLINE' : (r.online === 0 ? 'OFFLINE' : 'PARTIAL');
      return `<tr onclick="openRO('${r.id}')"><td>${r.id}</td><td class="ro-name">${esc(r.name)}</td><td>${esc(r.area || r.territory)}</td><td>${esc(r.vendor)}</td><td class="num product-ebms">${fmt(r.ebms, 2)}</td><td class="num product-hsd">${fmt(r.hsd, 2)}</td><td class="num"><b>${fmt(r.total, 2)}</b></td><td class="num">${r.tanks.length}</td><td><span class="pill ${st === 'OFFLINE' ? 'off' : st === 'ONLINE' ? '' : 'partial'}">${st}</span></td></tr>`;
    }).join('') || '<tr><td colspan="9" class="empty">No matching retail outlets found.</td></tr>';
  }

  window.openRO = function(id) {
    selectedRO = id;
    let ro = allRO.find(x => x.id === id);
    if (!ro) return;
    $('dashboardView')?.classList.add('hidden');
    $('detailView')?.classList.remove('hidden');
    if ($('roTitle')) $('roTitle').textContent = `${ro.name} (${ro.id})`;
    if ($('roSubtitle')) $('roSubtitle').textContent = `${ro.area || ro.territory} • RA Vendor: ${ro.vendor || '—'} • State: ${ro.state || '—'}`;
    let prods = ['EBMS', 'HSD'];
    if ($('productSummaries')) {
      $('productSummaries').innerHTML = prods.map(p => {
        let recs = ro.latest.filter(r => key(r['Product Name']) === p),
            capacity = recs.reduce((s, r) => s + (num(r['On-boarded Tankage']) || 0), 0),
            stock = recs.reduce((s, r) => s + (num(r.Stock) || 0), 0),
            onlineCap = recs.reduce((s, r) => s + (num(r['On-Line Tankage']) || 0), 0),
            ullage = Math.max(0, capacity - stock),
            net = capacity ? stock / capacity * 100 : 0;
        return `<div class="summarybox"><h3>STOCK SUMMARY — <span class="product-${p.toLowerCase()}">${p}</span></h3><table><thead><tr><th>Product</th><th class="num">Capacity (KL)</th><th class="num">Stock (KL)</th><th class="num">Ullage (KL)</th><th class="num">Stock %</th><th class="num">Ullage %</th></tr></thead><tbody><tr><td class="product-${p.toLowerCase()}">${p}</td><td class="num">${fmt(capacity, 2)}</td><td class="num">${fmt(stock, 2)}</td><td class="num">${fmt(ullage, 2)}</td><td class="num">${capacity ? fmt(net, 1) + '%' : '—'}</td><td class="num">${capacity ? fmt(100 - net, 1) + '%' : '—'}</td></tr></tbody></table><div class="small" style="padding:0 14px 12px">On-line tankage: ${fmt(onlineCap, 2)} KL • ${recs.length} product record(s)</div></div>`;
      }).join('');
    }
    let tanks = ro.tanks.slice().sort((a, b) => (num(a['Local tank id']) || 0) - (num(b['Local tank id']) || 0));
    if ($('tankCards')) {
      $('tankCards').innerHTML = tanks.length ? tanks.map(t => tankCard(t, ro)).join('') : '<div class="empty">No onboarded tank records for this RO.</div>';
      document.querySelectorAll('#tab-tank-stock .tank[data-tank]').forEach(el => el.addEventListener('click', () => showHourly(id, el.dataset.tank, el.dataset.product)));
    }
  };

  function tankCard(t, ro) {
    let id = key(t['Local tank id']), p = key(t.Product_Name), hr = data.hourly.find(h => h._ro === ro.id && h._tank === id && h._product === p);
    let hours = ['12 AM', '1 AM', '2 AM', '3 AM', '4 AM', '5 AM', '6 AM', '7 AM', '8 AM', '9 AM', '10 AM', '11 AM', '12 PM', '1 PM', '2 PM', '3 PM', '4 PM', '5 PM', '6 PM', '7 PM', '8 PM', '9 PM', '10 PM', '11 PM'];
    let vals = hr ? hours.map(h => num(hr[h + ' %'])) : [];
    let valid = vals.filter(v => v != null), current = valid.length ? valid[valid.length - 1] : null, first = valid.length ? valid[0] : null, capacity = num(t['On-Boarded Tankage (KL)']) ?? (hr ? num(hr['On-Boarded Tankage (KL)']): null);
    let status = key(t.Status), up = String(t['Uptime_tank(%)'] ?? '—');
    let color = p === 'EBMS' ? '#22bd42' : p === 'HSD' ? '#079fe8' : '#f2d000';
    return `<div class="tank" data-tank="${esc(id)}" data-product="${esc(p)}"><div class="tanktop"><div class="tanktitle">Tank ${esc(id)} <small>(${fmt(capacity, 2)} KL)</small></div><div>${status === 'ONLINE' ? '🔒' : '⚠️'} <span class="tankproduct product-${p.toLowerCase()}">${esc(p)}</span></div></div><div class="tankmeta"><span>◈ ${esc(up)} uptime</span><span class="tag">${esc(t['Morning_Status'] || '—')} / ${esc(t['Evening_Status'] || '—')}</span></div><div class="gauge" style="--pct:${Math.max(0, Math.min(100, current ?? 0))};--gauge:${color}"><span>${current == null ? '—' : fmt(current, 2) + '%'}</span></div><div class="tankstats"><div><div class="statlabel">First reported %</div><div class="statvalue">${first == null ? '—' : fmt(first, 2) + '%'}</div></div><div style="text-align:right"><div class="statlabel">Latest hourly %</div><div class="statvalue">${current == null ? '—' : fmt(current, 2) + '%'}</div></div><div><div class="statlabel">Onboarded tankage</div><div class="statvalue">${fmt(capacity, 2)} KL</div></div><div style="text-align:right"><div class="statlabel">Online hours</div><div class="statvalue">${fmt(num(t.Online_hours))} / 24</div></div></div><div class="tankbottom"><div><div class="statlabel">Table updated</div><div class="statvalue">${esc(t.Table_updated_at || '—')}</div></div><div style="text-align:right"><div class="statlabel">Status</div><span class="pill ${status === 'OFFLINE' ? 'off' : status === 'ONLINE' ? '' : 'partial'}">${esc(t.Status || 'Unknown')}</span></div></div></div>`;
  }

  function showHourly(roId, tankId, product) {
    let row = data.hourly.find(h => h._ro === roId && h._tank === key(tankId) && h._product === key(product));
    $('hourlyPanel')?.classList.remove('hidden');
    if ($('hourlyTitle')) $('hourlyTitle').textContent = `HOURLY STOCK TREND — TANK ${tankId} (${product})`;
    if ($('hourlySubtitle')) $('hourlySubtitle').textContent = 'Hourly stock percentage from Stock hourly.xlsx';
    let hours = ['12 AM', '1 AM', '2 AM', '3 AM', '4 AM', '5 AM', '6 AM', '7 AM', '8 AM', '9 AM', '10 AM', '11 AM', '12 PM', '1 PM', '2 PM', '3 PM', '4 PM', '5 PM', '6 PM', '7 PM', '8 PM', '9 PM', '10 PM', '11 PM'];
    let vals = hours.map(h => row ? num(row[h + ' %']) : null), max = Math.max(100, ...vals.filter(v => v != null));
    if ($('hourlyChart')) {
      $('hourlyChart').innerHTML = vals.map((v, i) => `<div class="barwrap" title="${hours[i]}: ${v == null ? 'No data' : fmt(v, 2) + '%'}"><small>${v == null ? '' : fmt(v, 0)}</small><div class="bar" style="height:${v == null ? 0 : Math.max(2, v / max * 145)}px;background:${key(product) === 'EBMS' ? '#22bd42' : '#079fe8'}"></div><label>${hours[i].replace(' ', '')}</label></div>`).join('');
    }
    $('hourlyPanel')?.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }

  function back() {
    selectedRO = null;
    $('detailView')?.classList.add('hidden');
    $('dashboardView')?.classList.remove('hidden');
  }

  // Pre-load local files if present
  async function tryPreloadLocalFiles() {
    try {
      const map = {};
      for (const name of ['Stock hourly.xlsx', 'Last updated stock.xlsx', 'Onboarded tanks.xlsx']) {
        try {
          const resp = await fetch(name);
          if (resp.ok) {
            map[name] = await resp.blob();
          }
        } catch (e) {}
      }
      if (map['Stock hourly.xlsx'] && map['Last updated stock.xlsx']) {
        await loadFiles(map);
      }
    } catch (e) {
      console.warn('Preload check:', e);
    }
  }

  function initEvents() {
    const folderBtn = $('folderBtn');
    const refreshBtn = $('refreshBtn');
    const mappingBtn = $('mappingBtn');
    const uploadBtn = $('uploadBtn');
    const fileInput = $('tankStockFileInput') || $('fileInput');
    const backBtn = $('backBtn');
    const closeHourly = $('closeHourly');
    const modalClose = $('modalClose');
    const chooseFolderNow = $('chooseFolderNow');

    if (folderBtn) folderBtn.onclick = () => chooseFolder();
    if (refreshBtn) refreshBtn.onclick = () => folderHandle ? readFolder() : fileInput?.click();
    if (mappingBtn) mappingBtn.onclick = () => {
      if (folderHandle) readFolder(false, true);
      else alert('Please select the Excel folder first. Then use Update tank mapping to refresh the fixed onboarded tank list.');
    };
    if (uploadBtn) uploadBtn.onclick = () => fileInput?.click();
    if (fileInput) fileInput.onchange = async e => {
      let map = {};
      for (const f of e.target.files) {
        if (FILES[f.name]) map[f.name] = f;
      }
      await loadFiles(map, true);
      e.target.value = '';
    };
    if (backBtn) backBtn.onclick = back;
    if (closeHourly) closeHourly.onclick = () => $('hourlyPanel')?.classList.add('hidden');
    if (modalClose) modalClose.onclick = () => $('tankModal')?.classList.add('hidden');
    if (chooseFolderNow) chooseFolderNow.onclick = () => {
      $('tankModal')?.classList.add('hidden');
      chooseFolder();
    };

    ['search', 'territory', 'area', 'vendor', 'state'].forEach(id => {
      const el = $(id);
      if (el) el.addEventListener(id === 'search' ? 'input' : 'change', drawROTable);
    });

    if (!window.showDirectoryPicker) {
      if ($('dataStatus')) $('dataStatus').textContent = 'Folder auto-loading needs Chrome/Edge with local folder access. You can still use Upload Excel files.';
    } else {
      if (!data.latest.length && $('tankModal')) {
        // Show modal when tab is opened
      }
    }

    tryPreloadLocalFiles();
  }

  window.initTankStockDashboard = function() {
    initEvents();
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initEvents);
  } else {
    initEvents();
  }

})();
