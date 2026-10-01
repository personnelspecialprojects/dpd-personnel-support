/* ============================================================
   audit.js — the Audit log tab. Shows the app-wide Audit Log list
   (inquiries, activity, tracked work, admin, access, data access).
   Rows are written by writeAudit()/audit() in core.js.
   ============================================================ */

const Audit = (() => {
  const A = AUDIT_FIELDS;
  let rows = [];
  let loadedDays = null;   // null until the tab is first opened
  let wired = false;

  function init(){
    if(!wired) wire();
    onView("auditView", () => { if(loadedDays === null) load(); else render(); });
  }

  function wire(){
    wired = true;
    const areaSel = document.getElementById("auditArea");
    areaSel.innerHTML = `<option value="">All areas</option>` +
      Object.values(AUDIT_AREAS).map(a => `<option value="${escapeHtml(a)}">${escapeHtml(a)}</option>`).join("");
    const days = document.getElementById("auditDays");
    if([...days.options].some(o => Number(o.value) === CONFIG.auditDefaultDays)) days.value = String(CONFIG.auditDefaultDays);
    areaSel.addEventListener("change", render);
    days.addEventListener("change", load);
    document.getElementById("auditSearchBox").addEventListener("input", render);
    document.getElementById("exportAuditCsvBtn").addEventListener("click", exportCsv);
  }

  function isViewActive(){
    const v = document.getElementById("auditView");
    return v && v.classList.contains("active");
  }

  async function load(){
    const days = Number(document.getElementById("auditDays").value) || CONFIG.auditDefaultDays;
    const body = document.getElementById("auditBody");
    if(loadedDays === null) body.innerHTML = `<tr class="static"><td colspan="5" class="muted">Loading...</td></tr>`;
    try{
      const since = addDays(startOfToday(), -(days - 1));
      const items = await spGetAll(CONFIG.lists.audit, filterQuery(`${A.logTime} ge ${odataDate(since)}`, "$top=2000"));
      rows = items.sort((a, b) => new Date(b[A.logTime]) - new Date(a[A.logTime]));
      loadedDays = days;
      render();
    }catch(err){
      console.error(err);
      body.innerHTML = `<tr class="static"><td colspan="5" class="error-text">The audit log couldn't load: ${escapeHtml(err.message)}. ${App.isAdmin ? "Run the SharePoint setup check in Admin." : `Let ${escapeHtml(CONFIG.adminContact)} know.`}</td></tr>`;
    }
  }

  /* Polling: only reload when someone is looking at the tab. */
  async function refresh(){
    if(loadedDays !== null && isViewActive()) await load();
  }

  /* New rows written in this session appear immediately. */
  function add(row){
    if(loadedDays === null || !row) return;
    rows.unshift(row);
    if(isViewActive()) render();
  }

  function areaOf(e){
    return e[A.area] || (e[A.ticketId] ? AUDIT_AREAS.inquiry : "");
  }

  function recordLabel(e){
    const area = areaOf(e);
    const id = e[A.recordId] ?? e[A.ticketId];
    if(id === null || id === undefined || id === "") return "";
    if(area === AUDIT_AREAS.inquiry) return `Inquiry #${Tickets.formatId(id)}`;
    if(area === AUDIT_AREAS.work) return `Item #${id}`;
    if(area === AUDIT_AREAS.activity) return `Entry #${id}`;
    if(area === AUDIT_AREAS.admin) return `#${id}`;
    return `#${id}`;
  }

  function filtered(){
    const area = document.getElementById("auditArea").value;
    const q = document.getElementById("auditSearchBox").value.trim().toLowerCase();
    return rows.filter(e => {
      if(area && areaOf(e) !== area) return false;
      if(q){
        const hay = `${recordLabel(e)} ${e[A.staffMember] || ""} ${e[A.action] || ""} ${e[A.details] || ""}`.toLowerCase();
        if(!hay.includes(q)) return false;
      }
      return true;
    });
  }

  function render(){
    const body = document.getElementById("auditBody");
    if(!body || loadedDays === null) return;
    const list = filtered();
    document.getElementById("auditEmptyState").hidden = list.length > 0;
    const shown = list.slice(0, 1000);
    body.innerHTML = shown.map(e => `
      <tr class="static">
        <td class="nowrap">${formatDate(e[A.logTime])}</td>
        <td class="nowrap"><span class="area-tag">${escapeHtml(areaOf(e))}</span></td>
        <td class="nowrap">${escapeHtml(recordLabel(e))}</td>
        <td>${escapeHtml(e[A.staffMember] || "")}</td>
        <td>${escapeHtml(e[A.action] || "")}${e[A.details] ? `<div class="muted small audit-details">${escapeHtml(e[A.details])}</div>` : ""}</td>
      </tr>`).join("");
    document.getElementById("auditPills").innerHTML =
      `<span class="pill"><b>${list.length}</b> ${list.length === 1 ? "entry" : "entries"}${list.length > shown.length ? ` (showing newest ${shown.length}; export for all)` : ""}</span>`;
  }

  function exportCsv(){
    if(loadedDays === null){ toast("Open the audit log first.", { type: "error" }); return; }
    const list = filtered();
    const out = [["Time", "Area", "Record", "Staff Member", "Staff Email", "Action", "Details",
      "Previous Status", "New Status", "Problem Snapshot", "Internal Notes Snapshot", "Customer Note Snapshot"]];
    list.forEach(e => out.push([
      formatDate(e[A.logTime]), areaOf(e), recordLabel(e), e[A.staffMember], e[A.staffEmail], e[A.action], e[A.details],
      e[A.previousStatus], e[A.newStatus], e[A.problemSnapshot], e[A.internalNotesSnapshot], e[A.customerNoteSnapshot]
    ]));
    downloadCsv(`audit-log-${toDateInput(new Date())}.csv`, out);
    audit(AUDIT_AREAS.data, `Exported the audit log (${plural(list.length, "row")}, last ${plural(loadedDays, "day")})`);
  }

  return { init, refresh, add, load };
})();
