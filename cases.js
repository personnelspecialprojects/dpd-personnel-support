(window.PS_FILE_VERSIONS = window.PS_FILE_VERSIONS || {})["cases.js"] = "2026.10.05-3";
/* ============================================================
   cases.js — 2.0 employee processes ("cases").

   An activity type becomes an EMPLOYEE PROCESS when its CaseFields
   column holds a definition:
     { "employee": true, "processKey": "loa",
       "fields": [ { "key": "received", "label": "Date request received", "type": "date", "col": "received" },
                   { "key": "enteredWorkday", "label": "Entered in Workday", "type": "date" }, ... ] }

   Each case is a Work Item (so checklists, handoffs, history and
   completion all work as before) tied to an employee. Fields with
   "col" are stored in real Work Items columns (ReceivedDate, StartDate,
   EndDate, Supervisor, Outcome); every other field is stored in the
   CaseData column as JSON.

   A checklist step can be linked to a field ({ text, field: "enteredWorkday" }).
   Checking the step fills the field (today's date / Yes / your name);
   filling the field checks the step. The employee's record therefore
   updates when the checklist does.
   ============================================================ */

const Cases = (() => {
  const W = WORK_FIELDS;
  const T = TYPE_FIELDS;
  let wired = false;
  let currentTypeId = null;
  const cache = new Map();       // typeId -> cases[]
  let loading = false;

  /* ---------- process definitions ---------- */

  function def(type){
    if(!type) return null;
    const raw = type[T.caseFields];
    if(!raw || !String(raw).trim()) return null;
    try{
      const d = JSON.parse(raw);
      if(!d || !d.employee) return null;
      d.fields = Array.isArray(d.fields) ? d.fields.filter(f => f && f.key && f.label) : [];
      return d;
    }catch(_){
      console.warn("CaseFields isn't valid JSON on activity", type.Id);
      return null;
    }
  }

  function isEmployeeProcess(type){ return !!def(type); }

  function typeFor(item){ return Activity.types.find(t => t.Id === item[W.activityTypeId]) || null; }

  function fieldsFor(item){
    const d = def(typeFor(item));
    if(d) return d.fields;
    // Activity removed or redefined: fall back to whatever data the case holds.
    return Object.keys(readData(item)).filter(k => !k.startsWith("_")).map(k => ({ key: k, label: k, type: "text" }));
  }

  function fieldByKey(item, key){ return fieldsFor(item).find(f => f.key === key) || null; }

  /* Fields a checklist step can be linked to (they can be filled automatically). */
  function linkable(f){ return f && ["date", "yesno", "text"].includes(f.type); }

  /* ---------- reading and writing values ---------- */

  function readData(item){
    try{
      const d = JSON.parse(item[W.caseData] || "{}");
      return d && typeof d === "object" ? d : {};
    }catch(_){ return {}; }
  }

  function getValue(item, f){
    if(f.col && CASE_COLUMNS[f.col]){
      const v = item[W[CASE_COLUMNS[f.col].field]];
      if(CASE_COLUMNS[f.col].type === "date"){
        if(v) return isoDay(v);
        return readData(item)[`_txt_${f.key}`] || "";   // legacy text like "none" in a date column
      }
      return v || "";
    }
    const d = readData(item);
    return d[f.key] === undefined || d[f.key] === null ? "" : String(d[f.key]);
  }

  function isFilled(v){
    const s = String(v ?? "").trim();
    return s !== "" && !/^no$/i.test(s);
  }

  /* Build the SharePoint body for a set of field changes { key: value } on an item. */
  function bodyFor(item, changes){
    const body = {};
    const data = readData(item);
    let dataChanged = false;
    Object.entries(changes).forEach(([key, value]) => {
      const f = fieldByKey(item, key) || { key, type: "text" };
      const v = value === null || value === undefined ? "" : String(value).trim();
      if(f.col && CASE_COLUMNS[f.col]){
        const colName = W[CASE_COLUMNS[f.col].field];
        if(CASE_COLUMNS[f.col].type === "date"){
          const iso = /^\d{4}-\d{2}-\d{2}$/.test(v) ? dayToIso(v) : null;
          body[colName] = iso;
          const txtKey = `_txt_${key}`;
          if(!iso && v){ data[txtKey] = v; dataChanged = true; }
          else if(data[txtKey] !== undefined){ delete data[txtKey]; dataChanged = true; }
        }else{
          body[colName] = v;
        }
      }else{
        if(v === "") delete data[key]; else data[key] = v;
        dataChanged = true;
      }
    });
    if(dataChanged) body[W.caseData] = JSON.stringify(data);
    return body;
  }

  /* Value a field gets when its linked step is checked. */
  function valueForStep(f){
    if(!f) return "";
    if(f.type === "date") return isoDay(new Date());
    if(f.type === "yesno") return "Yes";
    if(f.type === "text") return App.user.name;
    return "";
  }

  /* After field changes, make linked steps agree: filled field = checked step. */
  function syncSteps(steps, valuesByKey){
    const changed = [];
    steps.forEach(s => {
      if(!s.field || !(s.field in valuesByKey)) return;
      const want = isFilled(valuesByKey[s.field]);
      if(!!s.done !== want){
        s.done = want;
        s.doneBy = want ? App.user.name : "";
        s.doneOn = want ? new Date().toISOString() : "";
        changed.push(`${want ? "Checked" : "Unchecked"}: ${s.text}`);
      }
    });
    return changed;
  }

  function display(f, v){
    if(!isFilled(v) && !/^no$/i.test(String(v || ""))) return "";
    if(f.type === "date") return showDay(v);
    if(f.type === "money"){
      const n = Number(String(v).replace(/[$,\s]/g, ""));
      return Number.isFinite(n) ? n.toLocaleString("en-US", { style: "currency", currency: "USD" }) : String(v);
    }
    return String(v);
  }

  /* ---------- case details panel inside the work item window ---------- */

  function employeeLine(item){
    const id = item[W.employeeId];
    if(!id && !item[W.employeeName]) return "";
    const emp = id ? Roster.find(id) : null;
    const name = emp ? emp[EMP_FIELDS.title] : item[W.employeeName];
    const bits = [
      emp && emp[EMP_FIELDS.rank], emp && emp[EMP_FIELDS.division],
      emp && emp[EMP_FIELDS.badge] ? `Badge ${emp[EMP_FIELDS.badge]}` : "",
      emp && emp[EMP_FIELDS.active] === false ? "Inactive in roster" : "",
      !emp && id ? "Not in the current roster" : ""
    ].filter(Boolean).map(escapeHtml).join(" · ");
    return `<div class="case-emp">
      <div><span class="meta-label">Employee</span>
        <button type="button" class="link-btn case-emp-link" data-emp="${escapeHtml(id || "")}">${escapeHtml(name || "(no name)")}</button>
        <span class="mine-id">${escapeHtml(id || "no ID")}</span></div>
      ${bits ? `<div class="muted small">${bits}</div>` : ""}
    </div>`;
  }

  function inputHtml(f, v, disabled){
    const id = `cf_${f.key}`;
    const dis = disabled ? "disabled" : "";
    const val = escapeHtml(v || "");
    if(f.type === "date"){
      const isDate = /^\d{4}-\d{2}-\d{2}$/.test(v || "");
      return isDate || !v
        ? `<input type="date" id="${id}" data-key="${f.key}" value="${val}" ${dis}>`
        : `<input type="text" id="${id}" data-key="${f.key}" value="${val}" ${dis} title="Imported text. Replace with a date (YYYY-MM-DD) if known.">`;
    }
    if(f.type === "yesno"){
      const opts = ["", "Yes", "No"];
      const extra = v && !opts.includes(v) ? [v] : [];
      return `<select id="${id}" data-key="${f.key}" ${dis}>${[...opts, ...extra].map(o =>
        `<option value="${escapeHtml(o)}"${o === (v || "") ? " selected" : ""}>${escapeHtml(o || "—")}</option>`).join("")}</select>`;
    }
    if(f.type === "longtext") return `<textarea id="${id}" data-key="${f.key}" rows="2" ${dis}>${val}</textarea>`;
    if(f.type === "choice"){
      const list = `dl_${f.key}`;
      return `<input type="text" id="${id}" data-key="${f.key}" value="${val}" list="${list}" ${dis}>
        <datalist id="${list}">${(f.options || []).map(o => `<option value="${escapeHtml(o)}"></option>`).join("")}</datalist>`;
    }
    const inputmode = f.type === "number" || f.type === "money" ? `inputmode="decimal"` : "";
    return `<input type="text" id="${id}" data-key="${f.key}" value="${val}" ${inputmode} ${dis}>`;
  }

  function sectionHtml(item, open){
    const fields = fieldsFor(item);
    const legacy = item[W.legacyKey] ? `<div class="muted small">Imported from the legacy spreadsheet (${escapeHtml(String(item[W.legacyKey]).split("|")[1] || "")}).</div>` : "";
    if(!fields.length && !item[W.employeeId]) return "";
    const linked = {};
    stepsOfItem(item).forEach((s, i) => { if(s.field) linked[s.field] = i + 1; });
    return `<section class="case-section">
      ${employeeLine(item)}
      ${legacy}
      ${(() => {
        const field = f => `
          <label class="case-field${f.type === "longtext" ? " wide" : ""}">
            <span class="case-label">${escapeHtml(f.label)}${linked[f.key] ? ` <span class="link-badge" title="Checking step ${linked[f.key]} fills this in">step ${linked[f.key]}</span>` : ""}</span>
            ${inputHtml(f, getValue(item, f), !open)}
          </label>`;
        const hl = fields.filter(f => f.highlight);
        const rest = fields.filter(f => !f.highlight);
        return (hl.length ? `<div class="case-highlight"><div class="case-grid">${hl.map(field).join("")}</div></div>` : "") +
          `<div class="case-grid">${rest.map(field).join("")}</div>`;
      })()}
      ${open && fields.length ? `<div class="case-actions"><button type="button" class="btn btn-navy btn-sm" id="caseSaveBtn">Save details</button><span class="muted small" id="caseSaveNote"></span></div>` : ""}
    </section>`;
  }

  function stepsOfItem(item){
    try{ const a = JSON.parse(item[W.steps] || "[]"); return Array.isArray(a) ? a : []; }catch(_){ return []; }
  }

  function wireSection(item, onSave){
    const btn = document.getElementById("caseSaveBtn");
    if(btn) btn.addEventListener("click", () => {
      const changes = {};
      fieldsFor(item).forEach(f => {
        const el = document.getElementById(`cf_${f.key}`);
        if(!el) return;
        const now = el.value.trim();
        if(now !== getValue(item, f)) changes[f.key] = now;
      });
      if(!Object.keys(changes).length){
        document.getElementById("caseSaveNote").textContent = "No changes to save.";
        return;
      }
      onSave(changes);
    });
    document.querySelectorAll("#workModalBody .case-field input, #workModalBody .case-field select, #workModalBody .case-field textarea")
      .forEach(el => el.addEventListener("input", () => {
        const note = document.getElementById("caseSaveNote");
        if(note) note.textContent = "Unsaved changes";
      }));
    document.querySelectorAll("#workModalBody .case-emp-link").forEach(b => b.addEventListener("click", () => {
      if(b.dataset.emp) Roster.showEmployee(b.dataset.emp);
    }));
  }

  /* Describe changes for history/audit, e.g. 'Start date: "" → Oct 1, 2026'. */
  function describeChanges(item, changes){
    return Object.entries(changes).map(([k, v]) => {
      const f = fieldByKey(item, k) || { key: k, label: k, type: "text" };
      const before = display(f, getValue(item, f)) || "(blank)";
      const after = display(f, v) || "(blank)";
      return `${f.label}: ${before} → ${after}`;
    }).join("; ");
  }

  /* ---------- starting a case for an employee ---------- */

  /* Initial values a new case gets from the roster. */
  function initialValues(type, emp){
    const d = def(type);
    const vals = {};
    if(!d) return vals;
    d.fields.forEach(f => {
      if(f.col === "received") vals[f.key] = isoDay(new Date());
      if(f.col === "supervisor" && emp && emp[EMP_FIELDS.supervisor]) vals[f.key] = emp[EMP_FIELDS.supervisor];
    });
    return vals;
  }

  /* ---------- Employee processes tab ---------- */

  function init(){
    if(!wired) wire();
    onView("processView", () => { renderProcessPicker(); if(currentTypeId) loadCases(currentTypeId); });
    onView("reportsView", loadLastBackup);
    const bb = document.getElementById("backupBtn");
    if(bb && !bb.dataset.wired){ bb.dataset.wired = "1"; bb.addEventListener("click", exportBackup); }
  }

  function wire(){
    wired = true;
    document.getElementById("procSelect").addEventListener("change", e => {
      currentTypeId = Number(e.target.value) || null;
      if(currentTypeId) loadCases(currentTypeId); else renderTable();
    });
    document.getElementById("procStatus").addEventListener("change", renderTable);
    document.getElementById("procSearch").addEventListener("input", renderTable);
    document.getElementById("procRefresh").addEventListener("click", () => { if(currentTypeId) loadCases(currentTypeId, true); });
    document.getElementById("procExportOne").addEventListener("click", exportCurrent);
    document.getElementById("procExportAll").addEventListener("click", exportAll);
    document.getElementById("procStartBtn").addEventListener("click", startFromTab);
    document.getElementById("procStartId").addEventListener("keydown", e => { if(e.key === "Enter"){ e.preventDefault(); startFromTab(); } });
  }

  function processes(){
    return Activity.sortedTypes(Activity.types.filter(t => isEmployeeProcess(t) && t[T.active] !== false));
  }

  function renderProcessPicker(){
    const sel = document.getElementById("procSelect");
    const list = processes();
    const keep = currentTypeId;
    sel.innerHTML = `<option value="">Choose a process…</option>` + list.map(t =>
      `<option value="${t.Id}">${escapeHtml(t.Title)}${Activity.teamOf(t) !== TEAM_BOTH ? ` (${escapeHtml(Activity.teamOf(t))})` : ""}</option>`).join("");
    if(keep && list.some(t => t.Id === keep)) sel.value = String(keep);
    else if(list.length && !keep){ currentTypeId = list[0].Id; sel.value = String(currentTypeId); }
    document.getElementById("procEmpty").hidden = list.length > 0;
    document.getElementById("procMain").hidden = list.length === 0;
  }

  async function loadCases(typeId, force){
    if(cache.has(typeId) && !force){ renderTable(); return; }
    if(loading) return;
    loading = true;
    document.getElementById("procBody").innerHTML = `<tr class="static"><td colspan="20" class="muted">Loading...</td></tr>`;
    try{
      const rows = await spGetAll(CONFIG.lists.workItems,
        filterQuery(`${W.activityTypeId} eq ${typeId}`, "$top=2000"));
      cache.set(typeId, rows);
    }catch(err){
      console.error(err);
      document.getElementById("procBody").innerHTML = `<tr class="static"><td colspan="20" class="error-text">Cases couldn't load: ${escapeHtml(err.message)}</td></tr>`;
      loading = false;
      return;
    }
    loading = false;
    renderTable();
  }

  /* Called by Work whenever a case is created or changed, so tables stay current. */
  function itemChanged(item){
    const list = cache.get(item[W.activityTypeId]);
    if(!list) return;
    const i = list.findIndex(x => x.Id === item.Id);
    if(i >= 0) list[i] = item; else list.unshift(item);
    if(document.getElementById("processView").classList.contains("active") && currentTypeId === item[W.activityTypeId]) renderTable();
  }

  function filteredCases(){
    const rows = cache.get(currentTypeId) || [];
    const st = document.getElementById("procStatus").value;
    const q = document.getElementById("procSearch").value.trim().toLowerCase();
    return rows.filter(r => {
      const s = r[W.status] || "Open";
      if(st === "open" && s !== "Open") return false;
      if(st === "closed" && s === "Open") return false;
      if(q){
        const hay = `${r[W.employeeId] || ""} ${r[W.employeeName] || ""} ${r[W.identifier] || ""} ${r[W.caseData] || ""} ${r[W.supervisor] || ""}`.toLowerCase();
        if(!hay.includes(q)) return false;
      }
      return true;
    }).sort((a, b) => new Date(b[W.receivedDate] || b[W.startedOn]) - new Date(a[W.receivedDate] || a[W.startedOn]));
  }

  function renderTable(){
    const type = Activity.types.find(t => t.Id === currentTypeId);
    const head = document.getElementById("procHead");
    const body = document.getElementById("procBody");
    if(!type){ head.innerHTML = ""; body.innerHTML = ""; return; }
    const shown = def(type).fields.filter(f => f.type !== "longtext");
    const fields = [...shown.filter(f => f.highlight), ...shown.filter(f => !f.highlight)];   // highlighted columns lead
    const rows = filteredCases();
    const all = cache.get(currentTypeId) || [];
    document.getElementById("procCount").textContent =
      `${rows.length} of ${all.length} cases · ${all.filter(r => (r[W.status] || "Open") === "Open").length} open`;
    head.innerHTML = `<tr><th>Employee</th>${fields.map(f => `<th class="${f.highlight ? "hl-col" : ""}">${escapeHtml(f.label)}</th>`).join("")}<th>Checklist</th><th>Status</th><th>Owner</th></tr>`;
    body.innerHTML = rows.map(r => {
      const st = stepsOfItem(r);
      const done = st.filter(s => s.done).length;
      const status = r[W.status] || "Open";
      return `<tr data-case="${r.Id}" tabindex="0" class="${status !== "Open" ? "case-closed" : ""}">
        <td class="nowrap"><div class="inq-title">${escapeHtml(r[W.employeeName] || "(no name)")}</div><div class="muted small">${escapeHtml(r[W.employeeId] || "")}</div></td>
        ${fields.map(f => `<td class="${f.type === "date" || f.type === "money" || f.type === "number" ? "nowrap" : ""}${f.highlight ? " hl-col" : ""}${f.type === "number" || f.type === "money" ? " num" : ""}">${escapeHtml(display(f, getValue(r, f)))}</td>`).join("")}
        <td class="nowrap">${st.length ? `${done}/${st.length}` : "—"}</td>
        <td class="nowrap"><span class="status-badge status-${status === "Open" ? "InProgress" : status === "Completed" ? "Completed" : "Merged"}">${escapeHtml(status)}</span></td>
        <td class="nowrap small">${escapeHtml(r[W.ownerName] || "")}</td>
      </tr>`;
    }).join("") || `<tr class="static"><td colspan="${fields.length + 4}" class="muted">No cases match.</td></tr>`;
    body.querySelectorAll("tr[data-case]").forEach(tr => {
      const go = () => Work.openItem(Number(tr.dataset.case));
      tr.addEventListener("click", go);
      tr.addEventListener("keydown", e => { if(e.key === "Enter") go(); });
    });
  }

  async function startFromTab(){
    const type = Activity.types.find(t => t.Id === currentTypeId);
    const input = document.getElementById("procStartId");
    const id = input.value.trim();
    const msg = document.getElementById("procStartMsg");
    msg.className = "save-msg"; msg.textContent = "";
    if(!type) return;
    if(!/^\d+$/.test(id)){ msg.className = "save-msg err"; msg.textContent = "Enter an employee number."; input.focus(); return; }
    const res = await startCase(type, id, {});
    if(res.needsConfirm){
      if(!confirm(res.needsConfirm)) return;
      await startCase(type, id, { force: true, notInRoster: true });
    }else if(res.duplicate){
      if(confirm(`${res.duplicate[W.employeeName] || id} already has an open ${type.Title}. Open it instead? (Cancel starts a new one.)`)){
        Work.openItem(res.duplicate.Id);
        input.value = "";
        return;
      }
      await startCase(type, id, { force: true, notInRoster: !Roster.find(id) });
    }
    input.value = "";
  }

  /* Start a case. Returns { item } | { duplicate } | { needsConfirm: message }. */
  async function startCase(type, empId, opts){
    await Roster.ready();
    const emp = Roster.find(empId);
    if(!emp && !opts.notInRoster){
      return { needsConfirm: `Employee ${empId} isn't in the roster${Roster.count() ? "" : " (no roster has been imported yet)"}. Start the case anyway?` };
    }
    const res = await Work.start(type, empId, { force: opts.force, employee: emp || { [EMP_FIELDS.employeeId]: empId } });
    if(res.item){
      toast(`Started ${type.Title} for ${res.item[W.employeeName] || empId}.`, {
        type: "success", actionLabel: "Open", onAction: () => Work.openItem(res.item.Id), duration: 8000
      });
    }
    return res;
  }

  /* ---------- Excel export (works like the old spreadsheet) ---------- */

  function splitName(full){
    const s = String(full || "");
    if(s.includes(",")){ const [l, f] = s.split(","); return [l.trim(), (f || "").trim()]; }
    const parts = s.trim().split(/\s+/);
    return parts.length > 1 ? [parts.slice(-1)[0], parts.slice(0, -1).join(" ")] : [s, ""];
  }

  function cellValue(f, v){
    if(f.type === "date" && /^\d{4}-\d{2}-\d{2}$/.test(v || "")){
      const [y, m, d] = v.split("-").map(Number);
      return new Date(y, m - 1, d);
    }
    if((f.type === "money" || f.type === "number") && v !== ""){
      const n = Number(String(v).replace(/[$,\s]/g, ""));
      return Number.isFinite(n) ? n : v;
    }
    return v || "";
  }

  function sheetRows(type, cases, full){
    const fields = def(type).fields;
    const tplSteps = Activity.templateSteps(type).filter(s => !s.field);   // linked steps already appear as fields
    const header = ["Employee ID", "Last Name", "First Name", ...fields.map(f => f.label),
      ...tplSteps.map(s => s.text), "Case Status", "Owner", "Case Started", "Case Closed",
      ...(full ? ["Checklist (all steps)", "Case #", "Imported From"] : [])];
    const rows = cases.slice().sort((a, b) => new Date(a[W.receivedDate] || a[W.startedOn]) - new Date(b[W.receivedDate] || b[W.startedOn]))
      .map(r => {
        const [last, first] = splitName(r[W.employeeName]);
        const steps = stepsOfItem(r);
        return [r[W.employeeId] || "", last, first,
          ...fields.map(f => cellValue(f, getValue(r, f))),
          ...tplSteps.map(t => {
            const s = steps.find(x => x.text === t.text);
            return s && s.done ? (s.doneOn ? new Date(s.doneOn) : "Yes") : "";
          }),
          r[W.status] || "Open", r[W.ownerName] || "",
          r[W.startedOn] ? new Date(r[W.startedOn]) : "", r[W.closedOn] ? new Date(r[W.closedOn]) : "",
          ...(full ? [
            steps.map((s, i) => `${i + 1}. ${s.done ? "[x]" : "[ ]"} ${s.text}${s.done && s.doneOn ? ` (${showDay(s.doneOn)}${s.doneBy ? ", " + s.doneBy : ""})` : ""}`).join("\n"),
            r.Id, r[W.legacyKey] ? String(r[W.legacyKey]).split("|")[1] || "" : ""
          ] : [])];
      });
    return [header, ...rows];
  }

  function addSheet(XLSXlib, wb, name, aoa){
    const ws = XLSXlib.utils.aoa_to_sheet(aoa, { cellDates: true, dateNF: "m/d/yyyy" });
    ws["!cols"] = aoa[0].map((h, i) => ({ wch: Math.min(40, Math.max(10, String(h).length + 2, ...aoa.slice(1, 50).map(r => String(r[i] instanceof Date ? "00/00/0000" : r[i] ?? "").length + 1))) }));
    ws["!autofilter"] = { ref: XLSXlib.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: aoa.length - 1, c: aoa[0].length - 1 } }) };
    let safe = name.replace(/[\\/?*\[\]:]/g, " ").slice(0, 31).trim() || "Sheet";
    let n = 2; const base = safe;
    while(wb.SheetNames.includes(safe)) safe = `${base.slice(0, 28)} ${n++}`;
    XLSXlib.utils.book_append_sheet(wb, ws, safe);
  }

  async function exportCurrent(){
    const type = Activity.types.find(t => t.Id === currentTypeId);
    if(!type){ toast("Choose a process first.", { type: "error" }); return; }
    try{
      const X = await loadSheetJS();
      const cases = filteredCases();
      const wb = X.utils.book_new();
      addSheet(X, wb, type.Title, sheetRows(type, cases));
      X.writeFile(wb, `${type.Title.replace(/[^\w\- ]+/g, "")} ${isoDay(new Date())}.xlsx`, { cellDates: true });
      audit(AUDIT_AREAS.data, `Exported ${type.Title} cases to Excel (${plural(cases.length, "case")})`);
    }catch(err){
      console.error(err);
      toast(`Export failed: ${err.message}`, { type: "error", duration: 0 });
    }
  }

  async function exportAll(){
    const btn = document.getElementById("procExportAll");
    btn.disabled = true; btn.textContent = "Preparing...";
    try{
      const X = await loadSheetJS();
      const wb = X.utils.book_new();
      let total = 0;
      for(const type of processes()){
        if(!cache.has(type.Id)){
          cache.set(type.Id, await spGetAll(CONFIG.lists.workItems, filterQuery(`${W.activityTypeId} eq ${type.Id}`, "$top=2000")));
        }
        const cases = cache.get(type.Id);
        total += cases.length;
        addSheet(X, wb, type.Title, sheetRows(type, cases));
      }
      if(!wb.SheetNames.length){ toast("There are no employee processes to export yet.", { type: "error" }); return; }
      X.writeFile(wb, `Personnel Support processes ${isoDay(new Date())}.xlsx`, { cellDates: true });
      audit(AUDIT_AREAS.data, `Exported all employee processes to Excel (${plural(wb.SheetNames.length, "tab")}, ${plural(total, "case")})`);
    }catch(err){
      console.error(err);
      toast(`Export failed: ${err.message}`, { type: "error", duration: 0 });
    }finally{
      btn.disabled = false; btn.textContent = "Export all processes";
    }
  }

  /* "Vacation balance (hrs) 140 · Sick balance (hrs) 312.25" for a case's highlighted fields. */
  function highlightSummary(item){
    return fieldsFor(item).filter(f => f.highlight)
      .map(f => { const v = display(f, getValue(item, f)); return v ? `${f.label.replace(/\s*\(hrs\)$/i, "")}: ${v}` : ""; })
      .filter(Boolean).join(" · ");
  }

  /* ---------- full backup (Reports tab) ----------
     One workbook with everything needed if the app were ever unavailable:
     About, one tab per employee process (legacy-style), Employees, other tracked work,
     the activity log, and every case's history and notes. */
  async function exportBackup(){
    const btn = document.getElementById("backupBtn");
    const status = document.getElementById("backupStatus");
    btn.disabled = true;
    const step = t => { status.className = "muted small"; status.textContent = t; };
    try{
      const X = await loadSheetJS();
      step("Reading employee processes...");
      const procs = processes();
      const procIds = new Set(procs.map(t => t.Id));
      const perProc = [];
      for(const type of procs){
        const cases = await spGetAll(CONFIG.lists.workItems, filterQuery(`${W.activityTypeId} eq ${type.Id}`, "$top=2000"));
        cache.set(type.Id, cases);
        perProc.push({ type, cases });
      }
      step("Reading the roster...");
      await Roster.ready(true);
      const emps = Roster.all.slice().sort((a, b) => String(a[EMP_FIELDS.title] || "").localeCompare(String(b[EMP_FIELDS.title] || "")));
      step("Reading other tracked work...");
      const otherTypes = Activity.types.filter(t => !procIds.has(t.Id) && Activity.templateSteps(t).length);
      const other = [];
      for(const t of otherTypes){
        other.push(...(await spGetAll(CONFIG.lists.workItems, filterQuery(`${W.activityTypeId} eq ${t.Id}`, "$top=2000"))));
      }
      step("Reading the activity log...");
      const L = LOG_FIELDS;
      const log = (await spGetAll(CONFIG.lists.activityLog, "$top=2000")).filter(e => !e[L.voided]);
      step("Reading case history and notes...");
      const WL = WORKLOG_FIELDS;
      const hist = await spGetAll(CONFIG.lists.workItemLog, "$top=2000");
      step("Building the workbook...");

      const wb = X.utils.book_new();
      const now = new Date();
      const totalCases = perProc.reduce((n, p) => n + p.cases.length, 0);
      addSheet(X, wb, "About", [
        ["Personnel Support backup"],
        ["Created", now],
        ["Created by", App.user.name],
        ["Portal version", window.PS_PAGE_VERSION || ""],
        ["SharePoint site", CONFIG.siteUrl],
        [],
        ["Contents", "Rows"],
        ...perProc.map(p => [`Process: ${p.type.Title}`, p.cases.length]),
        ["Employees (roster)", emps.length],
        ["Other tracked work", other.length],
        ["Activity log (one-step tasks)", log.length],
        ["Case history and notes", hist.length],
        [],
        ["Each process tab lists one case per row, like the original tracking workbook. Dates are calendar dates; checklist steps show the date they were done."],
        ["This file contains employee information. Store it where personnel records are kept, not on a personal drive."]
      ]);
      perProc.forEach(p => addSheet(X, wb, p.type.Title, sheetRows(p.type, p.cases, true)));

      const E = EMP_FIELDS;
      addSheet(X, wb, "Employees", [
        ["Emp#", "Badge", "Last Name", "First Name", "Rank", "Workgroup", "Supervisor", "Hire Date (AdjSvcDate)", "Active", "Source", "Last On Roster"],
        ...emps.map(e => [Roster.normId(e[E.employeeId]), e[E.badge] || "", e[E.lastName] || "", e[E.firstName] || "", e[E.rank] || "",
          e[E.division] || "", e[E.supervisor] || "", e[E.hireDate] ? new Date(e[E.hireDate]) : "",
          e[E.active] === false ? "No" : "Yes", e[E.source] || "", e[E.lastRosterDate] ? new Date(e[E.lastRosterDate]) : ""])
      ]);

      addSheet(X, wb, "Other tracked work", [
        ["Item #", "Activity", "Name / Reference", "Status", "Owner", "Started", "Closed", "Checklist"],
        ...other.map(r => [r.Id, r[W.activityName] || "", r[W.identifier] || r[W.title] || "", r[W.status] || "Open", r[W.ownerName] || "",
          r[W.startedOn] ? new Date(r[W.startedOn]) : "", r[W.closedOn] ? new Date(r[W.closedOn]) : "",
          stepsOfItem(r).map((s, i) => `${i + 1}. ${s.done ? "[x]" : "[ ]"} ${s.text}`).join("\n")])
      ]);

      addSheet(X, wb, "Activity log", [
        ["Logged", "Category", "Activity", "Employee # / Reference", "Quantity", "Staff Member", "Entry #"],
        ...log.sort((a, b) => new Date(a[L.loggedAt]) - new Date(b[L.loggedAt]))
          .map(e => [e[L.loggedAt] ? new Date(e[L.loggedAt]) : "", e[L.category] || "", e[L.activityName] || "", e[L.identifier] || "",
            Activity.qty(e), e[L.staffName] || "", e.Id])
      ]);

      addSheet(X, wb, "Case history and notes", [
        ["When", "Case #", "Staff Member", "Entry"],
        ...hist.sort((a, b) => (a[WL.workItemId] - b[WL.workItemId]) || (new Date(a[WL.loggedAt]) - new Date(b[WL.loggedAt])))
          .map(h => [h[WL.loggedAt] ? new Date(h[WL.loggedAt]) : "", h[WL.workItemId], h[WL.staffName] || "", h[WL.action] || ""])
      ]);

      const stamp = `${isoDay(now)} ${String(now.getHours()).padStart(2, "0")}${String(now.getMinutes()).padStart(2, "0")}`;
      X.writeFile(wb, `Personnel Support BACKUP ${stamp}.xlsx`, { cellDates: true });
      audit(AUDIT_AREAS.data, `Downloaded full backup (${procs.length} processes, ${totalCases} cases, ${emps.length} employees, ${log.length} activity entries)`);
      status.className = "small";
      status.textContent = `Backup downloaded: ${totalCases.toLocaleString()} cases, ${emps.length.toLocaleString()} employees, ${log.length.toLocaleString()} activity entries. Save it to the team's records folder.`;
      lastBackup = { when: now.toISOString(), who: App.user.name };
      renderLastBackup();
    }catch(err){
      console.error(err);
      status.className = "error-text small";
      status.textContent = `Backup didn't finish: ${err.message}`;
    }finally{
      btn.disabled = false;
    }
  }

  /* "Last full backup: Oct 5, 2026 3:10 PM by Jane Doe" from the audit log (last 90 days). */
  let lastBackup = null;
  async function loadLastBackup(){
    try{
      const A = AUDIT_FIELDS;
      const rows = await spGetAll(CONFIG.lists.audit, filterQuery(`${A.logTime} ge ${odataDate(addDays(new Date(), -90))}`, "$top=2000"));
      const hit = rows.filter(r => /^Downloaded full backup/.test(r[A.action] || ""))
        .sort((a, b) => new Date(b[A.logTime]) - new Date(a[A.logTime]))[0];
      lastBackup = hit ? { when: hit[A.logTime], who: hit[A.staffMember] } : null;
    }catch(err){ console.warn("Couldn't check the last backup:", err); }
    renderLastBackup();
  }

  function renderLastBackup(){
    const el = document.getElementById("backupLast");
    if(!el) return;
    if(!lastBackup){ el.innerHTML = `<span class="warn-text">No full backup in the last 90 days.</span>`; return; }
    const days = daysSince(lastBackup.when);
    el.innerHTML = `Last full backup: <strong>${escapeHtml(formatDate(lastBackup.when))}</strong> by ${escapeHtml(lastBackup.who || "")}` +
      (days >= 30 ? ` <span class="warn-text">(${days} days ago — time for a new one)</span>` : "");
  }

  function invalidate(typeId){ if(typeId) cache.delete(typeId); else cache.clear(); }

  return {
    init, def, isEmployeeProcess, fieldsFor, fieldByKey, linkable, getValue, readData, isFilled,
    bodyFor, valueForStep, syncSteps, display, sectionHtml, wireSection, describeChanges,
    initialValues, startCase, itemChanged, invalidate, splitName, stepsOfItem, highlightSummary,
    exportBackup, loadLastBackup,
    showProcess(typeId){ currentTypeId = typeId; switchTab("processView"); }
  };
})();
