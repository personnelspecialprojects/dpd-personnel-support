(window.PS_FILE_VERSIONS = window.PS_FILE_VERSIONS || {})["roster.js"] = "2026.10.05-2";
/* ============================================================
   roster.js — 2.0 employee roster.

   - Employees list, loaded once and cached for lookups by employee ID.
   - Employees tab: search, an employee's details, and every case
     (employee process) on file for them, with a Start button.
   - Monthly roster import (Admin): reads the SQL report export
     (.xlsx or .csv) IN THE BROWSER, maps its columns, and updates
     the Employees list: adds new people, updates changed details,
     and marks people no longer on the report inactive. Nobody is
     deleted, so case history always keeps its employee.
   ============================================================ */

const Roster = (() => {
  const E = EMP_FIELDS;
  const W = WORK_FIELDS;
  let employees = [];
  let byId = new Map();
  let loadPromise = null;
  let wired = false;
  let currentEmpId = null;

  /* ---------- loading & lookup ---------- */

  function normId(v){
    if(v === null || v === undefined) return "";
    let s = String(v).trim();
    if(/^\d+\.0+$/.test(s)) s = s.replace(/\.0+$/, "");   // 123456.0 from some exports
    return s.replace(/^0+(?=\d)/, "");
  }

  let byBadge = new Map();
  function normBadge(v){ return String(v ?? "").trim().toUpperCase(); }

  function index(){
    byId = new Map(employees.map(e => [normId(e[E.employeeId]), e]));
    byBadge = new Map(employees.filter(e => e[E.badge]).map(e => [normBadge(e[E.badge]), e]));
  }

  function ready(force){
    if(force) loadPromise = null;
    if(!loadPromise){
      loadPromise = spGetAll(CONFIG.lists.employees, "$top=2000")
        .then(rows => { employees = rows; index(); return employees; })
        .catch(err => { loadPromise = null; console.error("Roster didn't load:", err); employees = []; index(); throw err; });
    }
    return loadPromise.catch(() => employees);
  }

  /* Employee number first; badge numbers are accepted too (some files use the badge as the ID). */
  function find(id){ return byId.get(normId(id)) || byBadge.get(normBadge(id)) || null; }
  function count(){ return employees.length; }

  function fullName(first, last){
    first = String(first || "").trim(); last = String(last || "").trim();
    return last && first ? `${last}, ${first}` : (last || first);
  }

  /* ---------- Employees tab ---------- */

  function init(){
    if(!wired) wire();
    onView("employeeView", async () => {
      await ready();
      renderSearch();
      if(currentEmpId) renderDetail(currentEmpId);
    });
  }

  function wire(){
    wired = true;
    document.getElementById("empSearch").addEventListener("input", renderSearch);
    document.getElementById("empShowInactive").addEventListener("change", renderSearch);
    document.getElementById("empAddBtn").addEventListener("click", addManual);
    // roster import (Admin)
    const file = document.getElementById("rosterFile");
    if(file) file.addEventListener("change", e => { if(e.target.files[0]) readRosterFile(e.target.files[0]); });
  }

  function renderSearch(){
    const q = document.getElementById("empSearch").value.trim().toLowerCase();
    const showInactive = document.getElementById("empShowInactive").checked;
    const box = document.getElementById("empResults");
    document.getElementById("empCount").textContent =
      `${employees.filter(e => e[E.active] !== false).length.toLocaleString()} active in roster`;
    if(!employees.length){
      box.innerHTML = `<div class="empty-block"><p><strong>No roster yet.</strong></p><p>${App.isAdmin
        ? "Import the monthly roster report in Admin → Employee roster."
        : `Ask ${escapeHtml(CONFIG.adminContact)} to import the roster.`}</p></div>`;
      return;
    }
    if(q.length < 2){
      box.innerHTML = `<p class="muted small">Type at least 2 characters of a name or employee number.</p>`;
      return;
    }
    const hits = employees.filter(e => {
      if(!showInactive && e[E.active] === false) return false;
      return `${e[E.employeeId]} ${e[E.badge] || ""} ${e[E.title]} ${e[E.firstName]} ${e[E.lastName]}`.toLowerCase().includes(q);
    }).slice(0, 60);
    box.innerHTML = hits.length ? `<table class="data"><tbody>${hits.map(e => `
      <tr data-emp="${escapeHtml(e[E.employeeId])}" tabindex="0" class="${e[E.active] === false ? "case-closed" : ""}">
        <td class="nowrap mine-id">${escapeHtml(e[E.employeeId])}</td>
        <td><div class="inq-title">${escapeHtml(e[E.title] || "")}</div><div class="muted small">${escapeHtml([e[E.rank], e[E.division], e[E.badge] ? `Badge ${e[E.badge]}` : ""].filter(Boolean).join(" · "))}</div></td>
        <td class="small">${e[E.active] === false ? "Inactive" : ""}</td>
      </tr>`).join("")}</tbody></table>` : `<p class="muted">No one matches.</p>`;
    box.querySelectorAll("tr[data-emp]").forEach(tr => {
      const go = () => renderDetail(tr.dataset.emp);
      tr.addEventListener("click", go);
      tr.addEventListener("keydown", ev => { if(ev.key === "Enter") go(); });
    });
  }

  async function showEmployee(id){
    currentEmpId = normId(id);
    closeOverlay("workOverlay");
    switchTab("employeeView");
  }

  async function renderDetail(id){
    const emp = find(id);
    currentEmpId = emp ? normId(emp[E.employeeId]) : normId(id);
    const wrap = document.getElementById("empDetail");
    const procs = Activity.sortedTypes(Activity.types.filter(t => Cases.isEmployeeProcess(t) && t[TYPE_FIELDS.active] !== false));
    wrap.hidden = false;
    wrap.innerHTML = `
      <div class="emp-head">
        <div>
          <h2>${escapeHtml(emp ? emp[E.title] : "Not in roster")}</h2>
          <div class="muted">Emp# ${escapeHtml(emp ? normId(emp[E.employeeId]) : currentEmpId)}${emp ? escapeHtml([emp[E.badge] ? `Badge ${emp[E.badge]}` : "", emp[E.rank], emp[E.division]].filter(Boolean).map(x => " · " + x).join("")) : ""}</div>
          ${emp ? `<div class="muted small">Supervisor: ${escapeHtml(emp[E.supervisor] || "—")}${emp[E.hireDate] ? ` · Hire date (adj. service): ${escapeHtml(showDay(emp[E.hireDate]))}` : ""} · ${emp[E.active] === false ? "<strong>Inactive</strong>" : "Active"} · Source: ${escapeHtml(emp[E.source] || "")}</div>` : ""}
        </div>
        <div class="emp-start">
          <select id="empStartType" aria-label="Process to start">${procs.map(t => `<option value="${t.Id}">${escapeHtml(t.Title)}</option>`).join("")}</select>
          <button type="button" class="btn btn-navy btn-sm" id="empStartBtn" ${procs.length ? "" : "disabled"}>Start process</button>
        </div>
      </div>
      <div id="empCases"><p class="muted">Loading cases...</p></div>`;
    const startBtn = document.getElementById("empStartBtn");
    if(startBtn) startBtn.addEventListener("click", async () => {
      const type = Activity.types.find(t => t.Id === Number(document.getElementById("empStartType").value));
      if(!type) return;
      let res = await Cases.startCase(type, currentEmpId, {});
      if(res.needsConfirm && confirm(res.needsConfirm)) res = await Cases.startCase(type, currentEmpId, { notInRoster: true });
      if(res.duplicate){
        if(confirm(`There's already an open ${type.Title} for this employee. Open it? (Cancel starts another.)`)){ Work.openItem(res.duplicate.Id); return; }
        res = await Cases.startCase(type, currentEmpId, { force: true, notInRoster: !emp });
      }
      if(res.item) renderDetail(currentEmpId);
    });
    try{
      const cases = await spGetAll(CONFIG.lists.workItems, filterQuery(`${W.employeeId} eq ${odataString(currentEmpId)}`, "$top=500"));
      cases.sort((a, b) => new Date(b[W.receivedDate] || b[W.startedOn]) - new Date(a[W.receivedDate] || a[W.startedOn]));
      audit(AUDIT_AREAS.data, `Viewed employee record ${currentEmpId} (${plural(cases.length, "case")})`);
      document.getElementById("empCases").innerHTML = cases.length ? `
        <table class="data"><thead><tr><th>Process</th><th>Received</th><th>Dates</th><th>Outcome</th><th>Checklist</th><th>Status</th></tr></thead>
        <tbody>${cases.map(c => {
          const st = Cases.stepsOfItem(c);
          return `<tr data-case="${c.Id}" tabindex="0" class="${(c[W.status] || "Open") !== "Open" ? "case-closed" : ""}">
            <td class="inq-title">${escapeHtml(c[W.activityName] || "")}</td>
            <td class="nowrap">${escapeHtml(showDay(c[W.receivedDate]))}</td>
            <td class="nowrap small">${escapeHtml([showDay(c[W.startDate]), showDay(c[W.endDate])].filter(Boolean).join(" – "))}</td>
            <td>${escapeHtml(c[W.outcome] || "")}</td>
            <td class="nowrap">${st.length ? `${st.filter(s => s.done).length}/${st.length}` : "—"}</td>
            <td class="nowrap">${escapeHtml(c[W.status] || "Open")}</td>
          </tr>`;
        }).join("")}</tbody></table>` : `<p class="muted">No cases on file for this employee.</p>`;
      document.querySelectorAll("#empCases tr[data-case]").forEach(tr => {
        const go = () => Work.openItem(Number(tr.dataset.case));
        tr.addEventListener("click", go);
        tr.addEventListener("keydown", ev => { if(ev.key === "Enter") go(); });
      });
    }catch(err){
      console.error(err);
      document.getElementById("empCases").innerHTML = `<p class="error-text">Cases couldn't load: ${escapeHtml(err.message)}</p>`;
    }
  }

  /* Add someone missing from the roster (new hire not yet on the report, etc.). */
  async function addManual(){
    const id = normId(prompt("Employee number:") || "");
    if(!id) return;
    if(!/^\d+$/.test(id)){ toast("Employee numbers are digits only.", { type: "error" }); return; }
    if(find(id)){ renderDetail(id); toast("That employee is already in the roster.", { type: "info" }); return; }
    const last = (prompt("Last name:") || "").trim();
    const first = (prompt("First name:") || "").trim();
    if(!last && !first) return;
    try{
      const row = await spCreate(CONFIG.lists.employees, {
        [E.title]: fullName(first, last), [E.employeeId]: id, [E.firstName]: first, [E.lastName]: last,
        [E.active]: true, [E.source]: "Manual"
      });
      employees.push(row); index();
      audit(AUDIT_AREAS.employees, `Added employee ${id} (${fullName(first, last)}) manually`, { recordId: row.Id });
      document.getElementById("empSearch").value = id;
      renderSearch();
      renderDetail(id);
    }catch(err){
      console.error(err);
      toast(`Employee wasn't added: ${err.message}`, { type: "error" });
    }
  }

  /* ---------- monthly roster import (Admin) ---------- */

  /* Report columns, matched by name ignoring case and spaces (never by position).
     The SQL report's order is Emp#, Badge, FirstName, LastName, Rank, WorkingOrg, Workgroup, Supervisor, AdjSvcDate.
     WorkingOrg is deliberately not read: org codes are being retired. */
  const TARGETS = [
    { key: "employeeId", label: "Emp# (employee number)", required: true, names: ["emp#", "empno", "empnum", "employeenumber", "employee#", "empid", "employeeid"] },
    { key: "badge", label: "Badge", names: ["badge", "badge#", "badgeno", "badgenumber"] },
    { key: "firstName", label: "First name", names: ["firstname", "first"] },
    { key: "lastName", label: "Last name", names: ["lastname", "last"] },
    { key: "rank", label: "Rank", names: ["rank", "ranktitle"] },
    { key: "workgroup", label: "Workgroup", names: ["workgroup", "assignment"] },
    { key: "supervisor", label: "Supervisor", names: ["supervisor", "supervisorname"] },
    { key: "hireDate", label: "AdjSvcDate (hire date)", names: ["adjsvcdate", "adjustedservicedate", "hiredate"] }
  ];
  const squashHeader = v => String(v ?? "").toLowerCase().replace(/\s+/g, "");

  let parsed = null;   // { headers, rows, firstDataRow }

  async function readRosterFile(file){
    const out = document.getElementById("rosterImportArea");
    out.innerHTML = `<p class="muted">Reading ${escapeHtml(file.name)} on this computer...</p>`;
    try{
      const X = await loadSheetJS();
      const wb = X.read(await file.arrayBuffer(), { type: "array", cellDates: false });
      const ws = wb.Sheets[wb.SheetNames[0]];   // first sheet
      const aoa = X.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
      // Title rows may sit above the header: find the row that has the employee-number column.
      const empNames = TARGETS[0].names;
      const h = aoa.findIndex((r, i) => i < 30 && (r || []).some(c => empNames.includes(squashHeader(c))));
      if(h < 0) throw new Error("No Emp# column found. The report needs a column named Emp#, Emp #, Employee Number, or EmpID");
      const headers = aoa[h].map((c, i) => String(c ?? `Column ${i + 1}`).replace(/\s+/g, " ").trim());
      const rows = aoa.slice(h + 1).map((r, i) => ({ r: r || [], n: h + 2 + i }))
        .filter(x => x.r.some(c => c !== null && String(c).trim() !== ""));
      parsed = { headers, rows, fileName: file.name };
      renderMapping();
    }catch(err){
      console.error(err);
      out.innerHTML = `<p class="error-text">That file couldn't be read: ${escapeHtml(err.message)}</p>`;
    }
  }

  function renderMapping(){
    const out = document.getElementById("rosterImportArea");
    const opt = (i, sel) => `<option value="${i}"${sel ? " selected" : ""}>${escapeHtml(parsed.headers[i])}</option>`;
    const found = parsed.headers.map(squashHeader);
    out.innerHTML = `
      <p><strong>${escapeHtml(parsed.fileName)}</strong>: ${parsed.rows.length.toLocaleString()} rows. Columns found (WorkingOrg is ignored on purpose):</p>
      <div class="map-grid">
        ${TARGETS.map(t => {
          const idx = found.findIndex(h => t.names.includes(h));
          return `<label class="stack-label">${escapeHtml(t.label)}${t.required ? " *" : ""}
            <select data-target="${t.key}"><option value="">(not in this report)</option>${parsed.headers.map((_, i) => opt(i, i === idx)).join("")}</select>
          </label>`;
        }).join("")}
      </div>
      <label class="check small"><input type="checkbox" id="rosterInactivate" checked> Mark people who aren't on this report as inactive (they're never deleted)</label>
      <div class="toolbar" style="margin-top:10px;">
        <button type="button" class="btn btn-ghost btn-sm" id="rosterPreviewBtn">Preview changes</button>
      </div>
      <div id="rosterPreview"></div>`;
    document.getElementById("rosterPreviewBtn").addEventListener("click", preview);
  }

  function mapping(){
    const m = {};
    document.querySelectorAll("#rosterImportArea [data-target]").forEach(s => { if(s.value !== "") m[s.dataset.target] = Number(s.value); });
    return m;
  }

  function cell(row, i){
    if(i === undefined) return "";
    const v = row[i];
    if(v === null || v === undefined) return "";
    return String(v).replace(/\s+/g, " ").trim();
  }

  /* One report row -> Employees list values. Hire date becomes "YYYY-MM-DD" (saved at local noon). */
  function rosterRecord(row, m){
    const first = cell(row, m.firstName), last = cell(row, m.lastName);
    let hire = "", hireProblem = "";
    if(m.hireDate !== undefined && row[m.hireDate] !== null && row[m.hireDate] !== undefined && row[m.hireDate] !== ""){
      const d = Legacy.toDay(row[m.hireDate]);
      if(/^\d{4}-\d{2}-\d{2}$/.test(d.v)) hire = d.v; else hireProblem = String(d.v);
    }
    return {
      rec: {
        [E.employeeId]: normId(cell(row, m.employeeId)),
        [E.badge]: normBadge(cell(row, m.badge)),
        [E.firstName]: first, [E.lastName]: last, [E.title]: fullName(first, last),
        [E.rank]: cell(row, m.rank), [E.division]: cell(row, m.workgroup),
        [E.supervisor]: cell(row, m.supervisor)
      },
      hire, hireProblem
    };
  }

  let plan = null;

  async function preview(){
    const m = mapping();
    const box = document.getElementById("rosterPreview");
    if(m.employeeId === undefined){ box.innerHTML = `<p class="error-text">Choose which column holds the Emp#.</p>`; return; }
    box.innerHTML = `<p class="muted">Comparing with the current roster...</p>`;
    await ready(true);
    const firstRowFor = new Map();
    const dups = [];
    const create = [], update = [], same = [];
    const hireProblems = [];
    let noId = 0;
    parsed.rows.forEach(({ r, n }) => {
      const { rec, hire, hireProblem } = rosterRecord(r, m);
      const id = rec[E.employeeId];
      if(!id){ noId++; return; }
      if(firstRowFor.has(id)){ dups.push(`${id} (row ${n}; first seen row ${firstRowFor.get(id)})`); return; }   // first row wins
      firstRowFor.set(id, n);
      if(hireProblem) hireProblems.push(`row ${n}: AdjSvcDate "${hireProblem}"`);
      const cur = byId.get(id);
      if(!cur){ create.push({ rec, hire }); return; }
      const diff = {};
      Object.keys(rec).forEach(k => {
        if(k === E.employeeId) return;
        if(rec[k] === "") return;   // a blank cell never erases a saved value
        if(String(cur[k] ?? "") !== rec[k]) diff[k] = rec[k];
      });
      if(hire && isoDay(cur[E.hireDate] || "") !== hire) diff[E.hireDate] = dayToIso(hire);
      if(cur[E.active] === false) diff[E.active] = true;
      if(cur[E.source] !== "Roster") diff[E.source] = "Roster";
      if(Object.keys(diff).length) update.push({ cur, diff }); else same.push(cur);
    });
    const inactivate = document.getElementById("rosterInactivate").checked
      ? employees.filter(e => e[E.active] !== false && e[E.source] === "Roster" && !firstRowFor.has(normId(e[E.employeeId])))
      : [];
    plan = { create, update, inactivate, dups: dups.length };
    box.innerHTML = `
      <ul class="plan-list">
        <li><strong>${create.length.toLocaleString()}</strong> new employees to add</li>
        <li><strong>${update.length.toLocaleString()}</strong> existing employees with changes (name, badge, rank, workgroup, supervisor, hire date, or back to active)</li>
        <li><strong>${same.length.toLocaleString()}</strong> unchanged</li>
        <li><strong>${inactivate.length.toLocaleString()}</strong> no longer on the report → marked inactive</li>
        ${noId ? `<li class="warn-text">${noId} rows skipped: no Emp#</li>` : ""}
      </ul>
      ${dups.length ? `<details class="mine" open><summary class="warn-text">${dups.length} duplicate Emp# row${dups.length === 1 ? "" : "s"}: the first row was used, these were skipped</summary><ul class="small">${dups.slice(0, 100).map(d => `<li>${escapeHtml(d)}</li>`).join("")}</ul></details>` : ""}
      ${hireProblems.length ? `<details class="mine"><summary class="warn-text">${hireProblems.length} AdjSvcDate value${hireProblems.length === 1 ? "" : "s"} couldn't be read as a date (left unchanged)</summary><ul class="small">${hireProblems.slice(0, 100).map(d => `<li>${escapeHtml(d)}</li>`).join("")}</ul></details>` : ""}
      ${inactivate.length > Math.max(50, employees.length * 0.2) ? `<p class="warn-text"><strong>Check before importing:</strong> that's a lot of people to mark inactive. Make sure this is the full roster report, not a partial one.</p>` : ""}
      <button type="button" class="btn btn-navy" id="rosterRunBtn" ${create.length + update.length + inactivate.length ? "" : "disabled"}>Update the roster</button>
      <div id="rosterProgress"></div>`;
    document.getElementById("rosterRunBtn").addEventListener("click", runImport);
  }

  async function runImport(){
    const btn = document.getElementById("rosterRunBtn");
    btn.disabled = true;
    const prog = document.getElementById("rosterProgress");
    const now = new Date().toISOString();
    const jobs = [
      ...plan.create.map(({ rec, hire }) => () => spCreate(CONFIG.lists.employees, {
        ...rec, ...(hire ? { [E.hireDate]: dayToIso(hire) } : {}),
        [E.active]: true, [E.source]: "Roster", [E.lastRosterDate]: now })),
      ...plan.update.map(u => () => spUpdate(CONFIG.lists.employees, u.cur.Id, { ...u.diff, [E.lastRosterDate]: now })),
      ...plan.inactivate.map(e => () => spUpdate(CONFIG.lists.employees, e.Id, { [E.active]: false }))
    ];
    const show = (d, n) => { prog.innerHTML = `<div class="progress progress-lg"><span style="width:${Math.round(d / n * 100)}%"></span></div><p class="muted small">${d.toLocaleString()} of ${n.toLocaleString()} saved. Keep this tab open.</p>`; };
    show(0, jobs.length);
    const res = await runPool(jobs, job => job(), 8, show);
    await ready(true);
    prog.innerHTML = `<p class="${res.failed.length ? "warn-text" : ""}"><strong>Roster updated.</strong> ${res.ok.toLocaleString()} saved${res.failed.length ? `, ${res.failed.length} failed (${escapeHtml(res.failed[0].err.message)}). Run the import again to retry; it only changes what's still different.` : "."}</p>`;
    audit(AUDIT_AREAS.imports, `Imported roster from ${parsed.fileName}: ${plan.create.length} added, ${plan.update.length} updated, ${plan.inactivate.length} marked inactive${plan.dups ? `, ${plan.dups} duplicate Emp# rows skipped` : ""}${res.failed.length ? `, ${res.failed.length} failed` : ""}`);
    renderSearch();
  }

  /* Used by the legacy import: add people who appear in old records but not the roster. */
  async function ensureEmployee(id, first, last){
    id = normId(id);
    if(!id || find(id)) return find(id);
    const row = await spCreate(CONFIG.lists.employees, {
      [E.title]: fullName(first, last), [E.employeeId]: id, [E.firstName]: first || "", [E.lastName]: last || "",
      [E.active]: false, [E.source]: "Legacy import"
    });
    employees.push(row); byId.set(id, row);
    return row;
  }

  return { init, ready, find, count, normId, fullName, showEmployee, ensureEmployee,
    get all(){ return employees; } };
})();
