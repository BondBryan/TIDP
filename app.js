/* TIDP Builder - Bond Bryan Architects, Smarter Working Group
   Static single-page app. Library data comes from library.json (maintained in GitHub).
   No server, no macros: the browser generates the TIDP and exports Excel / Word files. */
(function () {
  'use strict';

  // ------------------------------------------------------------------ state
  const newProject = () => ({
    name: '', code: '', issuer: '', sector: '', works: 'new', outsourcing: false,
    storeys: [], volumes: [], stageDates: { 0: '', 1: '', 2: '', 3: '', 4: '', 5: '', 6: '', 7: '' },
    stageOn: { 0: true, 1: true, 2: true, 3: true, 4: true, 5: true, 6: true, 7: true },
    tenders: { '3+': true, '4a': true, '4b': true }
  });
  const newUi = () => ({ tab: 'Drawings', search: '', status: '', showOff: true });
  const S = {
    lib: null,
    project: newProject(),
    sel: {},                       // id -> true/false overrides of the sector default
    ui: newUi()
  };
  const FIXED_CODES = ['ZZ', 'XX']; // always created; shown ticked and locked on step 1
  const REV = 'P01';               // every TIDP produced here is the first issue, dated today
  const STAGE_COLOURS = ['D98C6B', 'E88EB9', '59A3A1', 'F9C829', '6DA97D', '8D8FB9', 'E5C683', '5092BF']; // indexed by RIBA stage 0-7
  const STAGES = [0, 1, 2, 3, 4, 5, 6, 7];
  const TENDERS = ['3+', '4a', '4b'];
  const STATUS_KEY = s => /outsourc/i.test(s) ? 'op' : /confirm/i.test(s) ? 'confirm' : /cdp/i.test(s) ? 'cdp' : 'internal';
  const STATUS_SHORT = s => ({ op: 'Outsourcing Partner', confirm: 'Scope confirmation', cdp: 'CDP item', internal: 'Internal team' })[STATUS_KEY(s)];
  /** Status as shown on step 2: without outsourcing everything but scope confirmation is assumed internal, so left blank. */
  const statusLabel = s => S.project.outsourcing ? (STATUS_KEY(s) === 'op' ? 'Potential outsourcing' : STATUS_SHORT(s)) : STATUS_KEY(s) === 'confirm' ? STATUS_SHORT(s) : '';
  const CONTENT_TABS = ['Drawings', 'Images', 'Lists', 'Models', 'Text', 'Video'];
  const INTROS = {
    1: ['Project set-up', 'Enter the project details once. Everything below drives the file references, the sector selection and the storey and volume duplication. Nothing here can break the Library.'],
    2: ['Deliverables', 'Pre-selected from the Library for your sector and works type. Untick what the appointment does not cover, tick what it adds. Titles, numbers and formats come from the Library and are not editable here, which keeps every project on the practice standard.'],
    3: ['Review & export', 'Check the totals then export.']
  };

  function today() { return new Date().toISOString().slice(0, 10); }
  const $ = id => document.getElementById(id);
  const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  // ------------------------------------------------------------------ load library
  try { localStorage.removeItem('tidp-draft'); } catch (e) { /* drafts from earlier versions are no longer used */ }
  fetch('library.json', { cache: 'no-store' })
    .then(r => { if (!r.ok) throw new Error(r.status + ' ' + r.statusText); return r.json(); })
    .then(lib => { S.lib = lib; initFromLibrary(); renderAll(); })
    .catch(e => {
      $('libInfo').textContent = 'Library not loaded';
      $('setupWarn').style.display = 'block';
      $('setupWarn').textContent = 'library.json could not be loaded (' + e.message + '). When opening the page from a local folder, serve it with a local web server or use the GitHub Pages address.';
    });

  function initFromLibrary() {
    const L = S.lib.lists;
    $('libInfo').textContent = S.lib.meta.documents + ' Library items · v' + S.lib.meta.version;
    $('libVer').textContent = 'v' + S.lib.meta.version + ' (' + S.lib.meta.extracted + ')';
    $('pSector').innerHTML = '<option value="">Select a sector or discipline…</option>' + L.sectors.map(s => `<option>${esc(s)}</option>`).join('');
    defaultLevels();
  }
  function defaultLevels() {
    const L = S.lib.lists;
    // default storeys: every level code except ZZ/XX, 00 + 01 + RF ticked (mirrors the Info tab defaults)
    S.project.storeys = L.levels.filter(c => !FIXED_CODES.includes(c)).map(c => ({ code: c, name: L.storeyNames[c] || c, on: ['00', '01', 'RF'].includes(c) }));
    S.project.volumes = L.volumes.filter(c => !FIXED_CODES.includes(c)).map(c => ({ code: c, name: '', on: false }));
  }

  // ------------------------------------------------------------------ selection logic
  function defaultOn(d) {
    const p = S.project;
    if (p.sector && d.sectors[p.sector] === false) return false;
    if (d.buildType === 'existing' && p.works === 'new') return false;
    if (d.buildType === 'new' && p.works === 'existing') return false;
    // items tied to stages are dropped when none of their stages is in the appointment
    const used = STAGES.filter(n => d.stages[n]);
    if (used.length && !used.some(n => p.stageOn[n])) return false;
    return true;
  }
  const isOn = d => (d.id in S.sel) ? S.sel[d.id] : defaultOn(d);
  /** Items switched on only for disciplines (Landscape, Interior Design) are hidden for every other sector or discipline. */
  function visible(d) {
    const s = S.project.sector;
    if (!s || d.sectors[s]) return true;
    const disciplines = Object.keys(S.lib.lists.roleBySector).filter(k => k !== 'default');
    const onFor = Object.keys(d.sectors).filter(k => d.sectors[k]);
    return !(onFor.length && onFor.every(k => disciplines.includes(k)));
  }
  function setOn(d, v) { if (v === defaultOn(d)) delete S.sel[d.id]; else S.sel[d.id] = v; }

  // ------------------------------------------------------------------ generation (mirrors the VBA GenerateTIDP)
  function role() { const m = S.lib.lists.roleBySector; return m[S.project.sector] || m.default; }
  function discipline() { const m = S.lib.lists.disciplineBySector || {}; return m[S.project.sector] || m.default || 'Architect'; }
  function storeyDesc(base, code) {
    let d = base; if (/plans$/i.test(d)) d = d.slice(0, -1);
    const name = (S.project.storeys.find(s => s.code === code) || {}).name || S.lib.lists.storeyNames[code] || code;
    return name + ' ' + d;
  }
  /** Stage values for a document, restricted to the ticked stages. When the Library's first-issue
      stage is not ticked, the first ticked stage the document runs through becomes its first issue. */
  function stageValues(d) {
    const on = S.project.stageOn, vals = STAGES.map(n => on[n] ? d.stages[n] : null);
    const first = STAGES.find(n => d.stages[n] === 'YYYY-MM-DD');
    if (first && !on[first]) { const i = vals.findIndex(v => v); if (i >= 0) vals[i] = 'YYYY-MM-DD'; }
    return vals.map((v, i) => {
      if (!v) return '';
      if (v === 'YYYY-MM-DD') { const dt = S.project.stageDates[STAGES[i]]; return dt ? parseDate(dt) : 'YYYY-MM-DD'; }
      return v;
    });
  }
  function parseDate(iso) { const [y, m, d] = iso.split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)); }

  /** Returns [{tab, series, rows:[row...]}] for the selected documents, storey/volume duplication applied. */
  function generate(includeInternal) {
    const p = S.project, out = [];
    const storeys = p.storeys.filter(s => s.on).map(s => s.code);
    const volumes = p.volumes.filter(v => v.on).map(v => v.code);
    const tabs = includeInternal ? [...CONTENT_TABS, 'Internal'] : CONTENT_TABS;
    for (const tab of tabs) {
      const docs = S.lib.documents.filter(d => d.tab === tab && visible(d) && isOn(d));
      if (!docs.length) continue;
      const groups = [];
      for (const d of docs) {
        let g = groups[groups.length - 1];
        if (!g || g.series !== d.series) { g = { series: d.series, rows: [] }; groups.push(g); }
        const combos = [];
        if (d.floorDup && storeys.length) {
          for (const f of storeys) {
            if (d.volDup && volumes.length) for (const v of volumes) combos.push([f, v, storeyDesc(d.description, f)]);
            else combos.push([f, d.volume, storeyDesc(d.description, f)]);
          }
        } else if (d.volDup && volumes.length) {
          for (const v of volumes) combos.push([d.level, v, d.description]);
        } else combos.push([d.level, d.volume, d.description]);
        const stages = stageValues(d);
        for (const [lvl, vol, desc] of combos) {
          g.rows.push({
            id: d.id, project: p.code || 'ABCDE', originator: d.originator, volume: vol, level: lvl, type: d.type, role: role(),
            number: d.number, description: desc, format: d.format, scale: d.scale, workPackage: d.workPackage,
            stages, tender: d.tender, status: d.status, comments: d.comments, series: d.series, tab
          });
        }
      }
      out.push({ tab, groups, count: groups.reduce((n, g) => n + g.rows.length, 0) });
    }
    return out;
  }
  const docRef = r => [r.project, r.originator, r.volume, r.level, r.type, r.role, r.number].filter(Boolean).join('-');
  const fileName = s => s.replace(/[\\/:*?"<>|]+/g, '-');
  function fileRef() { return `${S.project.code || 'ABCDE'}-BBA-XX-XX-L-${role()}-1001 Task Information Delivery Plan_${discipline()}`; }
  function noteRef() { return `${S.project.code || 'ABCDE'}-BBA-XX-XX-T-${role()}-0010 ResourcingNote_${discipline()}`; }

  // ------------------------------------------------------------------ rendering
  function renderAll() { renderProject(); renderDocs(); renderSummary(); }

  function renderProject() {
    const p = S.project;
    $('pName').value = p.name; $('pCode').value = p.code;
    $('pIssuer').value = p.issuer; $('pSector').value = p.sector; $('pWorks').value = p.works; $('pOutsourcing').checked = p.outsourcing;
    $('chipOut').classList.toggle('on', p.outsourcing);
    const locked = FIXED_CODES.map(c => `<label class="chip locked" title="${c} is always created"><input type="checkbox" checked disabled> ${c}</label>`).join('');
    $('storeys').innerHTML = locked + p.storeys.map((s, i) => `<label class="chip ${s.on ? 'on' : ''}"><input type="checkbox" data-storey="${i}" ${s.on ? 'checked' : ''}> ${esc(s.code)} <small>${esc(s.name)}</small></label>`).join('');
    $('volumes').innerHTML = locked + p.volumes.map((v, i) => `<label class="chip ${v.on ? 'on' : ''}"><input type="checkbox" data-volume="${i}" ${v.on ? 'checked' : ''}> ${esc(v.code)} <small>${esc(v.name)}</small></label>`).join('');
    $('stageDates').innerHTML = S.lib.lists.stages.map(s => {
      const on = p.stageOn[s.n];
      return `<div class="stage ${on ? '' : 'off'}" style="border-color:#${STAGE_COLOURS[s.n]}"><label><input type="checkbox" data-stageon="${s.n}" ${on ? 'checked' : ''}><b>Stage ${s.n}</b></label><span class="hint">${esc(s.name)}</span><input type="date" data-stage="${s.n}" value="${p.stageDates[s.n] || ''}" ${on ? '' : 'disabled'}></div>`;
    }).join('');
    $('tenders').innerHTML = TENDERS.map(t => `<label class="chip ${p.tenders[t] ? 'on' : ''}"><input type="checkbox" data-tender="${t}" ${p.tenders[t] ? 'checked' : ''}> Tender ${t}</label>`).join('');
  }

  function renderDocs() {
    const p = S.project, u = S.ui, L = S.lib;
    const tabs = [...CONTENT_TABS, 'Internal'];
    $('tabFilter').innerHTML = tabs.map(t => {
      const n = L.documents.filter(d => d.tab === t && visible(d) && isOn(d)).length, tot = L.documents.filter(d => d.tab === t && visible(d)).length;
      return `<button class="${u.tab === t ? 'active' : ''}" data-tab="${t}">${t}<span class="c">${n}/${tot}</span></button>`;
    }).join('');
    const w = $('setupWarn');
    if (!p.sector) { w.style.display = 'block'; w.textContent = 'No sector or discipline selected yet: every Library item is shown as selected. Choose the sector or discipline on step 1 to apply the Library defaults.'; }
    else w.style.display = 'none';

    const statuses = L.lists.statuses.filter(statusLabel);
    if (!statuses.includes(u.status)) u.status = '';
    $('statusFilter').innerHTML = '<option value="">All statuses</option>' + statuses.map(s => `<option value="${esc(s)}">${esc(statusLabel(s))}</option>`).join('');
    $('statusFilter').value = u.status;

    const q = u.search.trim().toLowerCase();
    const docs = L.documents.filter(d => d.tab === u.tab && visible(d))
      .filter(d => !q || (d.number + ' ' + d.description + ' ' + d.workPackage + ' ' + d.series + ' ' + d.format).toLowerCase().includes(q))
      .filter(d => !u.status || d.status === u.status)
      .filter(d => u.showOff || isOn(d));
    const storeyN = p.storeys.filter(s => s.on).length, volN = p.volumes.filter(v => v.on).length;
    let html = '', series = null;
    for (const d of docs) {
      if (d.series !== series) { series = d.series; html += `<tr class="series"><td colspan="12">${esc(series)}</td></tr>`; }
      const on = isOn(d), k = STATUS_KEY(d.status);
      const stg = STAGES.map(n => `<span class="stg ${d.stages[n] ? 'on' : ''} ${p.stageOn[n] ? '' : 'out'}" style="background:#${STAGE_COLOURS[n]}" title="Stage ${n}: ${esc(d.stages[n] || 'not required')}${p.stageOn[n] ? '' : ' (stage not in this appointment)'}">${n}</span>`).join('');
      const dup = [d.floorDup ? `× ${storeyN} storeys` : '', d.volDup && volN ? `× ${volN} volumes` : ''].filter(Boolean).join(', ');
      const flag = d.buildType === 'existing' ? ' <span class="pill" title="Existing building only">existing</span>' : '';
      html += `<tr class="${on ? '' : 'off'}" data-id="${esc(d.id)}">
        <td><input type="checkbox" data-sel="${esc(d.id)}" ${on ? 'checked' : ''}></td>
        <td class="num">${esc(d.type)}-${esc(d.number)}</td>
        <td class="desc">${esc(d.description)}${d.comments ? '<i class="cmt" aria-label="Has a comment">i</i>' : ''}${flag}</td>
        <td>${esc(d.format)}</td><td>${esc(d.scale)}</td><td>${esc(d.workPackage)}</td>
        <td style="white-space:nowrap">${stg}</td>
        <td class="${d.tender['3+'] ? 'tick' : 'tickx'}">✓</td><td class="${d.tender['4a'] ? 'tick' : 'tickx'}">✓</td><td class="${d.tender['4b'] ? 'tick' : 'tickx'}">✓</td>
        <td>${statusLabel(d.status) ? `<span class="st ${k}"></span>${esc(statusLabel(d.status))}` : ''}</td>
        <td class="dup">${dup}</td></tr>`;
    }
    $('docBody').innerHTML = html || '<tr><td colspan="12" class="hint">No items match.</td></tr>';
  }

  function renderSummary() {
    const p = S.project, gen = generate(true);
    const issued = gen.filter(g => g.tab !== 'Internal').reduce((n, g) => n + g.count, 0);
    const internal = (gen.find(g => g.tab === 'Internal') || { count: 0 }).count;
    const rows = gen.flatMap(g => g.groups.flatMap(x => x.rows));
    const op = rows.filter(r => STATUS_KEY(r.status) === 'op').length, conf = rows.filter(r => STATUS_KEY(r.status) === 'confirm').length;
    $('summary').innerHTML = [
      ['Documents issued', issued, 'across ' + gen.filter(g => g.tab !== 'Internal').length + ' sheets'],
      ['Internal items', internal, 'not in the issued TIDP; listed in the resourcing note'],
      ['Documents suitable for outsourcing', op, p.outsourcing ? 'in the resourcing note' : 'outsourcing not selected'],
      ['Scope to confirm', conf, 'Confirmation of BBA scope required', scopeValidated !== null && scopeValidated === scopeTicked()
        ? '<button class="btn small done" id="btnScope" title="Scope reviewed. Click to reopen and edit">✓ Validated</button>'
        : '<button class="btn secondary small" id="btnScope">Review now</button>'],
      ['Storeys × volumes', p.storeys.filter(s => s.on).length + ' × ' + (p.volumes.filter(v => v.on).length || '-'), 'duplication applied']
    ].map(([l, v, s, extra]) => `<div class="kpi"><b>${v}</b>${l}<br><span>${s}</span>${extra ? '<br>' + extra : ''}</div>`).join('');
    const issues = [];
    if (!p.sector) issues.push('No sector or discipline selected.');
    if (!p.code) issues.push('No project code: file references will read "ABCDE".');
    if (!p.name) issues.push('No project name.');
    if (!p.storeys.some(s => s.on)) issues.push('No storeys ticked: per-storey drawings will not be generated.');
    if (!STAGES.some(n => p.stageOn[n])) issues.push('No RIBA stage ticked on step 1.');
    $('exportWarn').innerHTML = issues.length ? `<div class="warn"><b>Before exporting:</b> ${issues.join(' ')}</div>` : '<div class="ok">Set-up complete. The export will contain ' + issued + ' issued documents.</div>';
    $('exportFiles').innerHTML = [fileName(fileRef()) + '.xlsx', fileName(noteRef()) + '.docx'].map(f => `<li>${esc(f)}</li>`).join('');
    $('btnExport').textContent = 'Export TIDP and resourcing note';
    const toReview = scopeNeedsReview() ? scopeDocs().filter(isOn).length : 0;
    $('exportHint').textContent = toReview
      ? `${toReview} scope confirmation item${toReview > 1 ? 's are' : ' is'} ticked: the export opens the scope review first, then export once validated.`
      : 'Your browser may ask once to allow this page to download multiple files.';
  }

  // ------------------------------------------------------------------ scope confirmation review (second route to the step 2 ticks)
  const scopeDocs = () => S.lib.documents.filter(d => STATUS_KEY(d.status) === 'confirm' && visible(d));
  // ticked scope items at the last Validate; any change to them means reviewing again before export
  let scopeValidated = null;
  const scopeTicked = () => scopeDocs().filter(isOn).map(d => d.id).join('|');
  const scopeNeedsReview = () => scopeDocs().some(isOn) && scopeTicked() !== scopeValidated;
  function openScope() {
    let html = '', tab = null;
    for (const d of scopeDocs()) {
      if (d.tab !== tab) { tab = d.tab; html += `<tr class="series"><td colspan="4">${esc(tab)}</td></tr>`; }
      html += `<tr class="${isOn(d) ? '' : 'off'}"><td><input type="checkbox" data-scope="${esc(d.id)}" ${isOn(d) ? 'checked' : ''}></td>
        <td class="num">${esc(d.type)}-${esc(d.number)}</td><td class="desc">${esc(d.description)}</td><td>${esc(d.comments)}</td></tr>`;
    }
    $('scopeBody').innerHTML = html;
    $('scopeDlg').showModal();
  }
  $('summary').addEventListener('click', e => { if (e.target.id === 'btnScope') openScope(); });
  $('scopeBody').addEventListener('change', e => { if (e.target.dataset.scope) e.target.closest('tr').classList.toggle('off', !e.target.checked); });
  $('scopeCancel').addEventListener('click', () => $('scopeDlg').close());
  $('scopeOk').addEventListener('click', () => {
    $('scopeBody').querySelectorAll('input[data-scope]').forEach(i => setOn(S.lib.documents.find(d => d.id === i.dataset.scope), i.checked));
    scopeValidated = scopeTicked();
    $('scopeDlg').close(); renderAll(); toast('Scope selection validated');
  });

  // ------------------------------------------------------------------ intro pop-ups
  const introSeen = {};
  function showIntro(n, force) {
    if (!INTROS[n] || (introSeen[n] && !force)) return;
    introSeen[n] = true;
    const dlg = $('intro');
    $('introTitle').textContent = INTROS[n][0]; $('introText').textContent = INTROS[n][1];
    if (dlg.open) dlg.close();
    if (dlg.showModal) dlg.showModal(); else alert(INTROS[n][1]);
  }
  document.querySelectorAll('[data-intro]').forEach(b => b.addEventListener('click', () => showIntro(b.dataset.intro, true)));
  showIntro(1);

  $('btnReset').addEventListener('click', () => {
    if (!confirm('Clear the project set-up and all deliverable ticks, and go back to the defaults?')) return;
    S.project = newProject(); S.sel = {}; S.ui = newUi(); projectErr = false; markProject(); scopeValidated = null;
    if (S.lib) defaultLevels();
    ['storeyCode', 'storeyName', 'volCode', 'volName', 'search'].forEach(id => $(id).value = '');
    $('statusFilter').value = ''; $('showOff').checked = true;
    Object.keys(introSeen).forEach(k => delete introSeen[k]);
    go(1);
  });

  // ------------------------------------------------------------------ events
  document.querySelectorAll('nav.steps button, [data-goto]').forEach(b => b.addEventListener('click', () => go(b.dataset.step || b.dataset.goto)));
  // step 1 minimum: these fields must be filled before the deliverables can be reviewed
  const REQUIRED = { name: 'pName', code: 'pCode', issuer: 'pIssuer', sector: 'pSector' };
  let projectErr = false;
  const missingFields = () => Object.keys(REQUIRED).filter(k => !String(S.project[k] || '').trim());
  function markProject() {
    const miss = missingFields();
    if (!miss.length) projectErr = false;
    $('projectErr').style.display = projectErr ? 'block' : 'none';
    $('cardProject').classList.toggle('missing', projectErr);
    for (const [k, id] of Object.entries(REQUIRED)) $(id).closest('.f').classList.toggle('missing', projectErr && miss.includes(k));
  }
  $('cardProject').addEventListener('input', markProject);
  $('cardProject').addEventListener('change', markProject);

  function go(n) {
    if (String(n) !== '1' && missingFields().length) {
      projectErr = true; markProject();
      $('projectErr').scrollIntoView({ block: 'center', behavior: 'smooth' });
      return;
    }
    document.querySelectorAll('nav.steps button').forEach(b => b.classList.toggle('active', b.dataset.step === String(n)));
    document.querySelectorAll('section.step').forEach(s => s.classList.toggle('active', s.id === 'step' + n));
    window.scrollTo(0, 0); if (S.lib) renderAll(); showIntro(n);
  }
  const bind = (id, key, ev = 'input') => $(id).addEventListener(ev, e => { S.project[key] = e.target.type === 'checkbox' ? e.target.checked : e.target.value; if (key === 'outsourcing') $('chipOut').classList.toggle('on', e.target.checked); });
  bind('pName', 'name'); bind('pCode', 'code'); bind('pIssuer', 'issuer');
  bind('pSector', 'sector', 'change'); bind('pWorks', 'works', 'change'); bind('pOutsourcing', 'outsourcing', 'change');
  $('pCode').addEventListener('input', e => { e.target.value = e.target.value.toUpperCase(); S.project.code = e.target.value; });
  $('storeys').addEventListener('change', e => { const i = e.target.dataset.storey; if (i != null) { S.project.storeys[i].on = e.target.checked; renderProject(); } });
  $('volumes').addEventListener('change', e => { const i = e.target.dataset.volume; if (i != null) { S.project.volumes[i].on = e.target.checked; renderProject(); } });
  $('stageDates').addEventListener('change', e => {
    const t = e.target;
    if (t.dataset.stage) S.project.stageDates[t.dataset.stage] = t.value;
    if (t.dataset.stageon) { S.project.stageOn[t.dataset.stageon] = t.checked; renderProject(); }
  });
  $('tenders').addEventListener('change', e => { const t = e.target.dataset.tender; if (t) { S.project.tenders[t] = e.target.checked; renderProject(); } });
  $('addStorey').addEventListener('click', () => {
    const code = $('storeyCode').value.trim().toUpperCase(), name = $('storeyName').value.trim();
    if (!code) return toast('Enter a level code');
    if (S.project.storeys.some(s => s.code === code)) return toast('Level ' + code + ' already exists');
    S.project.storeys.push({ code, name: name || code, on: true }); $('storeyCode').value = $('storeyName').value = ''; renderProject();
  });
  $('addVolume').addEventListener('click', () => {
    const code = $('volCode').value.trim().toUpperCase(), name = $('volName').value.trim();
    if (!code) return toast('Enter a volume code');
    if (S.project.volumes.some(v => v.code === code)) return toast('Volume ' + code + ' already exists');
    S.project.volumes.push({ code, name, on: true }); $('volCode').value = $('volName').value = ''; renderProject();
  });
  $('tabFilter').addEventListener('click', e => { const t = e.target.closest('button'); if (t) { S.ui.tab = t.dataset.tab; renderDocs(); } });
  $('search').addEventListener('input', e => { S.ui.search = e.target.value; renderDocs(); });
  $('statusFilter').addEventListener('change', e => { S.ui.status = e.target.value; renderDocs(); });
  $('showOff').addEventListener('change', e => { S.ui.showOff = e.target.checked; renderDocs(); });
  // hovering a row shows its Library comment
  const tip = $('tip');
  $('docBody').addEventListener('mousemove', e => {
    const tr = e.target.closest('tr[data-id]'), d = tr && S.lib.documents.find(x => x.id === tr.dataset.id);
    if (!d || !d.comments) { tip.style.display = 'none'; return; }
    if (tip.dataset.id !== d.id) { tip.dataset.id = d.id; tip.innerHTML = '<b>Comment</b>' + esc(d.comments); }
    tip.style.display = 'block';
    const x = e.clientX + 14, y = e.clientY + 16, w = tip.offsetWidth, h = tip.offsetHeight;
    tip.style.left = Math.max(8, x + w > innerWidth - 8 ? e.clientX - w - 14 : x) + 'px';
    tip.style.top = Math.max(8, y + h > innerHeight - 8 ? e.clientY - h - 12 : y) + 'px';
  });
  const hideTip = () => { tip.style.display = 'none'; };
  $('docBody').addEventListener('mouseleave', hideTip);
  addEventListener('scroll', hideTip, { passive: true });
  $('docBody').addEventListener('change', e => {
    const id = e.target.dataset.sel; if (!id) return;
    const d = S.lib.documents.find(x => x.id === id); setOn(d, e.target.checked);
    e.target.closest('tr').classList.toggle('off', !e.target.checked); renderDocs();
  });
  function shownDocs() { return [...document.querySelectorAll('#docBody input[data-sel]')].map(i => S.lib.documents.find(d => d.id === i.dataset.sel)); }
  $('selAll').addEventListener('click', () => { shownDocs().forEach(d => setOn(d, true)); renderDocs(); });
  $('selNone').addEventListener('click', () => { shownDocs().forEach(d => setOn(d, false)); renderDocs(); });
  $('selReset').addEventListener('click', () => { if (confirm('Discard all manual ticks and return to the Library defaults for this sector?')) { S.sel = {}; renderDocs(); } });

  // ------------------------------------------------------------------ export (Excel + resourcing note in one go)
  $('btnExport').addEventListener('click', async e => {
    if (scopeNeedsReview()) { toast('Review the scope confirmation items before exporting'); return openScope(); }
    const btn = e.currentTarget; btn.disabled = true;
    try {
      toast('Building the export…');
      const xlsx = await buildWorkbook(), note = await buildNote();
      download(xlsx, fileName(fileRef()) + '.xlsx');
      setTimeout(() => download(note, fileName(noteRef()) + '.docx'), 400);
      toast('TIDP and resourcing note exported');
    } catch (err) { console.error(err); alert('The export failed: ' + err.message); }
    finally { btn.disabled = false; }
  });

  // ------------------------------------------------------------------ Excel export (ExcelJS)
  async function buildWorkbook() {
    const p = S.project, gen = generate(false), wb = new ExcelJS.Workbook();
    wb.creator = 'Bond Bryan Architects - TIDP Builder'; wb.created = new Date();
    const thin = { style: 'thin', color: { argb: 'FFBFBFBF' } }, border = { top: thin, left: thin, bottom: thin, right: thin };
    const tenders = TENDERS.filter(t => p.tenders[t]), SN = 11 + STAGES.length, NCOL = SN + tenders.length; // stage columns 12..SN
    const HEAD7 = ['Project', 'Originator', 'Volume / system', 'Level', 'Type', 'Role', 'Number', '', '', '', '', ...STAGES.map(n => 'Stage ' + n), ...tenders.map((_, i) => i ? '' : 'Tender')];
    const HEAD8 = ['', '', '', '', '', '', '', 'Description', 'Format', 'Scale', 'Work Package', ...S.lib.lists.stages.map(s => s.name), ...tenders];
    const WIDTHS = [11, 9, 9, 7, 6, 6, 9, 52, 12, 15, 38, ...STAGES.map(() => 14), ...tenders.map(() => 5.5)];
    const date = parseDate(today());
    const logoId = await loadLogo(wb);

    // ---- Cover
    const cover = wb.addWorksheet('Cover', { pageSetup: { orientation: 'portrait', fitToPage: true } });
    cover.columns = [{ width: 24 }, { width: 60 }];
    if (logoId != null) cover.addImage(logoId, { tl: { col: 0, row: 0 }, ext: { width: 300, height: 51 } });
    cover.getCell('A5').value = 'TASK INFORMATION DELIVERY PLAN'; cover.getCell('A5').font = { bold: true, size: 16 };
    const info = [['Project', p.name], ['Project code', p.code], ['File reference', fileRef() + ' - ' + REV], ['Date', date], ['Revision', REV], ['Sector or discipline', p.sector], ['Works', { new: 'New build', existing: 'Existing building', mixed: 'New build and existing' }[p.works]], ['Issued by', p.issuer], ['Issuing party', 'Bond Bryan']];
    info.forEach(([k, v], i) => { const r = 7 + i; cover.getCell(r, 1).value = k; cover.getCell(r, 1).font = { bold: true }; cover.getCell(r, 2).value = v; if (v instanceof Date) cover.getCell(r, 2).numFmt = 'yyyy-mm-dd'; });
    let r = 7 + info.length + 1;
    cover.getCell(r, 1).value = 'Contents'; cover.getCell(r, 1).font = { bold: true, size: 12 }; r++;
    [['Sheet', 'Documents'], ...gen.map(g => [g.tab, g.count]), ['Data', gen.reduce((n, g) => n + g.count, 0)]].forEach((row, i) => {
      row.forEach((v, c) => { const cell = cover.getCell(r, c + 1); cell.value = v; cell.border = border; if (i === 0) { cell.font = { bold: true }; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } }; } });
      r++;
    });
    cover.headerFooter.oddFooter = '&C' + 'Bond Bryan';

    // ---- content sheets
    const dataRows = []; // {sheet, row, tab, series}
    for (const g of gen) {
      const ws = wb.addWorksheet(g.tab, { views: [{ state: 'frozen', ySplit: 8 }], pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, paperSize: 8 } });
      ws.columns = WIDTHS.map(w => ({ width: w }));
      ws.getCell('A1').value = 'TASK INFORMATION DELIVERY PLAN'; ws.getCell('A1').font = { bold: true, size: 14 }; ws.mergeCells('A1:E1');
      [['Project Name', p.name], ['File Ref', fileRef() + ' - ' + REV], ['Date', date]].forEach(([k, v], i) => {
        const row = 3 + i; ws.getCell(row, 1).value = k; ws.getCell(row, 1).font = { bold: true }; ws.mergeCells(row, 1, row, 2);
        ws.getCell(row, 3).value = v; if (v instanceof Date) ws.getCell(row, 3).numFmt = 'yyyy-mm-dd'; ws.mergeCells(row, 3, row, i < 2 ? 8 : 5);
      });
      HEAD7.forEach((h, c) => { const cell = ws.getCell(7, c + 1); cell.value = h || null; });
      HEAD8.forEach((h, c) => { const cell = ws.getCell(8, c + 1); cell.value = h || null; });
      for (let c = 1; c <= NCOL; c++) for (const row of [7, 8]) {
        const cell = ws.getCell(row, c);
        cell.font = { bold: true, size: 9 }; cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }; cell.border = border;
        if (c >= 12 && c <= SN && !p.stageOn[STAGES[c - 12]]) { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEDEDED' } }; cell.font = { bold: true, size: 9, color: { argb: 'FF8C8280' } }; }
        else if (c >= 12 && c <= SN) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + STAGE_COLOURS[STAGES[c - 12]] } };
        else cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
      }
      for (let c = 1; c <= 7; c++) ws.mergeCells(7, c, 8, c);
      if (tenders.length > 1) ws.mergeCells(7, SN + 1, 7, NCOL);
      ws.getRow(8).height = 60;

      let rn = 9;
      for (const grp of g.groups) {
        const hr = ws.getRow(rn);
        hr.getCell(1).value = grp.series; hr.font = { bold: true };
        for (let c = 1; c <= NCOL; c++) { hr.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } }; if (c >= 12) hr.getCell(c).value = '◦'; }
        ws.mergeCells(rn, 1, rn, 11); hr.height = 19; rn++;
        for (const x of grp.rows) {
          const row = ws.getRow(rn);
          const vals = [x.project, x.originator, x.volume, x.level, x.type, x.role, x.number, x.description, x.format, x.scale, x.workPackage, ...x.stages, ...tenders.map(t => x.tender[t] ? '✓' : '')];
          vals.forEach((v, c) => {
            const cell = row.getCell(c + 1); cell.value = v === '' ? null : v; cell.border = border; cell.font = { size: 9 };
            if (c === 6) cell.numFmt = '@';
            if (c >= 11 && c < SN) { cell.alignment = { horizontal: 'center' }; if (v instanceof Date) cell.numFmt = 'yyyy-mm-dd'; else if (v === 'YYYY-MM-DD') cell.font = { size: 9, color: { argb: 'FFFF0000' } }; }
            if (c >= SN) cell.alignment = { horizontal: 'center' };
          });
          row.height = 19;
          dataRows.push({ sheet: g.tab, row: rn, tab: g.tab, series: grp.series });
          rn++;
        }
      }
      ws.autoFilter = { from: { row: 8, column: 1 }, to: { row: rn - 1, column: NCOL } };
      ws.headerFooter.oddFooter = '&L' + (p.name || '') + '&C' + 'Bond Bryan' + '&R' + 'Page &P of &N';
    }

    // ---- Data sheet (formula links, so edits on the content sheets flow through)
    const data = wb.addWorksheet('Data', { views: [{ state: 'frozen', ySplit: 1 }] });
    data.columns = [{ width: 34 }, { width: 11 }, { width: 30 }, ...WIDTHS.map(w => ({ width: w }))];
    const heads = ['Document reference', 'Document type', 'Series', ...HEAD8.map((h, i) => h || HEAD7[i])];
    heads.forEach((h, c) => { const cell = data.getCell(1, c + 1); cell.value = h; cell.font = { bold: true }; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } }; });
    let dr = 2;
    const q = s => `'${s.replace(/'/g, "''")}'`;
    for (const d of dataRows) {
      const sh = q(d.sheet), row = data.getRow(dr);
      row.getCell(1).value = { formula: ['A', 'B', 'C', 'D', 'E', 'F', 'G'].map(c => `${sh}!${c}${d.row}`).join('&"-"&') };
      row.getCell(2).value = d.tab; row.getCell(3).value = d.series;
      for (let c = 1; c <= NCOL; c++) {
        const ref = `${sh}!${colLetter(c)}${d.row}`, cell = row.getCell(3 + c);
        cell.value = { formula: `IF(${ref}="","",${ref})` };
        if (c >= 12 && c <= SN) cell.numFmt = 'yyyy-mm-dd';
        if (c === 7) cell.numFmt = '@';
      }
      dr++;
    }
    data.autoFilter = dr > 2 ? { from: 'A1', to: { row: dr - 1, column: 3 + NCOL } } : undefined;
    wb.calcProperties.fullCalcOnLoad = true;
    return new Blob([await wb.xlsx.writeBuffer()], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }
  function colLetter(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; }
  async function logoBytes() {
    try { const r = await fetch('logo.png'); return r.ok ? await r.arrayBuffer() : null; } catch (e) { return null; }
  }
  async function loadLogo(wb) { const buf = await logoBytes(); return buf ? wb.addImage({ buffer: buf, extension: 'png' }) : null; }

  // ------------------------------------------------------------------ resourcing note (docx)
  async function buildNote() {
    const D = window.docx, p = S.project, gen = generate(true);
    const rows = gen.flatMap(g => g.groups.flatMap(x => x.rows));
    const cat = r => r.tab === 'Internal' ? 'internalUse' : STATUS_KEY(r.status);
    const count = k => rows.filter(r => cat(r) === k).length;
    const sections = [['confirm', 'Confirmation of BBA scope required', 'Items to confirm against the appointment before they are committed.'],
      ['op', 'Outsourcing Partner items', 'Candidates for the Outsourcing Partner. The Project Lead co-ordinates scope, programme and quality checks.'],
      ['cdp', 'Contractor Designed Portion items', 'Items expected from the contractor or specialist; BBA reviews only.'],
      ['internal', 'Internal Team Items', 'General documents produced by the Bond Bryan project team; the Project Lead allocates resource.'],
      ['internalUse', 'Internal usage only', 'Documents for Bond Bryan internal use only; they are not part of the issued TIDP.']];
    const P = (t, o = {}) => new D.Paragraph({ children: [new D.TextRun({ text: t, font: 'Arial', size: o.size || 20, bold: o.bold, color: o.color })], spacing: { after: o.after ?? 100 }, heading: o.h });
    const cellOpts = { margins: { top: 40, bottom: 40, left: 80, right: 80 } };
    const cell = (t, w, bold, fill) => new D.TableCell({ ...cellOpts, width: { size: w, type: D.WidthType.DXA }, shading: fill ? { fill } : undefined, children: [new D.Paragraph({ children: [new D.TextRun({ text: String(t ?? ''), font: 'Arial', size: 16, bold })] })] });
    const logo = await logoBytes();
    const children = [
      ...(logo ? [new D.Paragraph({ children: [new D.ImageRun({ data: new Uint8Array(logo), transformation: { width: 240, height: 41 } })], spacing: { after: 300 } })] : []),
      P('INTERNAL RESOURCING NOTE', { size: 32, bold: true }), P('Task Information Delivery Plan · Internal, not for issue', { color: '8C8280' }),
      P(`Project: ${p.name}   Code: ${p.code}   Revision: ${REV}   Date: ${today()}   Sector or discipline: ${p.sector}   Issued by: ${p.issuer}   Outsourcing Partner: ${p.outsourcing ? 'may be used' : 'not selected'}`, { after: 200 }),
      P(`Summary: ${rows.length} deliverables in total, of which ${count('confirm')} requiring scope confirmation, ${count('op')} documents suitable for outsourcing, ${count('internal')} internal team items and ${count('internalUse')} for internal usage only.`, { after: 300 })];
    for (const [k, title, intro] of sections) {
      const list = rows.filter(r => cat(r) === k); if (!list.length) continue;
      children.push(P(`${title} (${list.length})`, { size: 24, bold: true, after: 60 }), P(intro, { color: '444444', after: 120 }));
      const W = [1900, 1100, 3900, 1000, 1000, 3100];
      const trs = [new D.TableRow({ tableHeader: true, children: ['Reference', 'Sheet', 'Description', 'Format', 'First stage', 'Comment'].map((h, i) => cell(h, W[i], true, 'D9D9D9')) })];
      for (const r of list) {
        const first = r.stages.findIndex(s => s); const fs = first >= 0 ? 'Stage ' + STAGES[first] + (r.stages[first] instanceof Date ? ' · ' + r.stages[first].toISOString().slice(0, 10) : '') : '';
        trs.push(new D.TableRow({ children: [docRef(r), r.tab, r.description, r.format, fs, r.comments].map((v, i) => cell(v, W[i])) }));
      }
      children.push(new D.Table({ rows: trs, width: { size: 12000, type: D.WidthType.DXA } }), P('', { after: 200 }));
    }
    const doc = new D.Document({ sections: [{ properties: { page: { size: { orientation: D.PageOrientation.LANDSCAPE }, margin: { top: 900, bottom: 900, left: 900, right: 900 } } }, children }] });
    return D.Packer.toBlob(doc);
  }

  // ------------------------------------------------------------------ helpers
  function download(blob, name) { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000); }
  let toastT; function toast(msg) { const t = $('toast'); t.textContent = msg; t.style.display = 'block'; clearTimeout(toastT); toastT = setTimeout(() => t.style.display = 'none', 3500); }
  window.TIDP = { S, generate, buildWorkbook, buildNote, fileRef, noteRef }; // for tests
})();
