(window.PS_FILE_VERSIONS = window.PS_FILE_VERSIONS || {})["work.js"] = "2026.10.05-2";
/* ============================================================
   work.js — the Tracked work tab: multi-step items that run over
   days or weeks (e.g. a job posting).

   An activity type with checklist steps shows "Start" on the board.
   Starting creates a Work Item holding its OWN copy of the steps, so
   editing the template later never changes items already underway,
   and steps can be added, edited, reordered or removed per item.

   Every change is read-modify-write against the latest copy from
   SharePoint, so two people checking different steps at the same
   time don't overwrite each other.

   Completing an item writes one normal Activity Log entry, so the
   Reports tab counts it like any other activity. Reopening voids it.
   ============================================================ */

const Work = (() => {
  const W = WORK_FIELDS;
  const WL = WORKLOG_FIELDS;
  const F = LOG_FIELDS;
  const T = TYPE_FIELDS;

  let items = [];            // open items + items closed in the last CONFIG.closedWorkDays
  let wired = false;
  let filter = "mine";       // mine | open | waiting | closed
  let typeFilter = null;     // activity type Id, set from the board's "N in progress" link
  let currentId = null;
  let history = [];
  let saving = false;
  let editingStepId = null;

  /* ---------- loading ---------- */

  async function init(){
    if(!wired) wire();
    await load();
    render();
    Activity.renderCounts();
    onView("workView", render);
  }

  async function refresh(){
    await load();
    render();
    // Refresh an open checklist only if the user isn't mid-edit in it.
    if(currentId && isModalOpen() && !editingStepId && !modalHasFocusInField() && !saving){
      renderModal();
    }
  }

  async function load(){
    const since = addDays(startOfToday(), -CONFIG.closedWorkDays);
    const [open, closed] = await Promise.all([
      spGetAll(CONFIG.lists.workItems, filterQuery(`${W.status} eq 'Open'`, "$top=2000")),
      spGetAll(CONFIG.lists.workItems, filterQuery(`${W.closedOn} ge ${odataDate(since)}`, "$top=2000"))
    ]);
    const byId = new Map();
    [...open, ...closed].forEach(i => byId.set(i.Id, i));
    items = [...byId.values()];
  }

  /* ---------- model helpers ---------- */

  function newStepId(){ return Math.random().toString(36).slice(2, 10); }

  function stepsOf(item){
    try{
      const arr = JSON.parse(item[W.steps] || "[]");
      return Array.isArray(arr) ? arr : [];
    }catch(_){
      return [];
    }
  }

  function info(item){
    const steps = stepsOf(item);
    const done = steps.filter(s => s.done).length;
    const next = steps.find(s => !s.done) || null;
    return { steps, done, total: steps.length, next, waiting: !!(next && next.external) };
  }

  function isOpen(item){ return (item[W.status] || "Open") === "Open"; }
  function isMe(email){ return !!email && localPart(email) === App.user.key; }

  function involvesMe(item){
    if(isMe(item[W.ownerEmail])) return true;
    return stepsOf(item).some(s => !s.done && isMe(s.assignee));
  }

  function openCountsByType(){
    const out = {};
    items.forEach(i => { if(isOpen(i)) out[i[W.activityTypeId]] = (out[i[W.activityTypeId]] || 0) + 1; });
    return out;
  }

  function itemForLog(logId){
    return items.find(i => i[W.completionLogId] === logId) || null;
  }

  function ageText(d){
    const n = daysSince(d);
    if(n === null) return "";
    if(n === 0) return "today";
    return plural(n, "day");
  }

  function teamOptions(selectedEmail, blankLabel){
    const opts = App.team
      .map(m => ({ email: String(m[TEAM_FIELDS.title] || "").toLowerCase(), name: teamMemberName(m[TEAM_FIELDS.title]) }))
      .filter(m => m.email)
      .sort((a, b) => a.name.localeCompare(b.name));
    const sel = localPart(selectedEmail);
    return (blankLabel !== undefined ? `<option value="">${escapeHtml(blankLabel)}</option>` : "") +
      opts.map(m => `<option value="${escapeHtml(m.email)}"${localPart(m.email) === sel && sel ? " selected" : ""}>${escapeHtml(m.name)}</option>`).join("");
  }

  /* ---------- start (called from the board) ---------- */

  async function start(type, identifier, opts = {}){
    const idf = String(identifier || "").trim();
    if(idf && !opts.force){
      const dup = items.find(i => isOpen(i) &&
        i[W.activityTypeId] === type.Id &&
        String(i[W.identifier] || "").toLowerCase() === idf.toLowerCase());
      if(dup) return { duplicate: dup };
    }
    const steps = Activity.templateSteps(type).map(s => ({
      id: newStepId(), text: s.text, external: !!s.external, ...(s.field ? { field: s.field } : {}),
      assignee: "", assigneeName: "", done: false, doneBy: "", doneOn: ""
    }));
    const now = new Date().toISOString();
    // 2.0: employee processes record who the case is about, and pre-fill from the roster.
    const emp = opts.employee || null;
    const empName = emp ? (emp[EMP_FIELDS.title] || "") : "";
    let caseBody = {};
    if(emp){
      const shell = { [W.activityTypeId]: type.Id, [W.caseData]: "{}" };
      const initial = Cases.initialValues(type, emp);
      caseBody = {
        [W.employeeId]: Roster.normId(emp[EMP_FIELDS.employeeId] || idf),
        [W.employeeName]: empName,
        ...Cases.bodyFor(shell, initial)
      };
      Cases.syncSteps(steps, initial);
    }
    const title = emp ? `${type.Title} — ${empName || idf}` : idf ? `${type.Title} — ${idf}` : `${type.Title} — ${formatDate(now)}`;
    const item = await spCreate(CONFIG.lists.workItems, {
      ...caseBody,
      [W.title]: title.slice(0, 255),
      [W.activityTypeId]: type.Id,
      [W.activityName]: type.Title,
      [W.category]: type[T.category] || "",
      [W.identifier]: idf,
      [W.status]: "Open",
      [W.ownerName]: App.user.name,
      [W.ownerEmail]: App.user.username,
      [W.startedBy]: App.user.name,
      [W.startedOn]: now,
      [W.lastActivityOn]: now,
      [W.steps]: JSON.stringify(steps)
    });
    items.push(item);
    await addLog(item.Id, emp ? `Started for ${empName || ""} (${idf})` : "Started");
    render();
    if(typeof Cases !== "undefined") Cases.itemChanged(item);
    return { item };
  }

  /* ---------- list view ---------- */

  function wire(){
    wired = true;
    document.querySelectorAll("#workFilters [data-filter]").forEach(b =>
      b.addEventListener("click", () => { filter = b.dataset.filter; render(); }));
    document.getElementById("workSearch").addEventListener("input", render);
    document.getElementById("workModalClose").addEventListener("click", closeModal);
  }

  function showType(typeId){
    typeFilter = typeId;
    filter = "open";
    switchTab("workView");
  }

  function filtered(){
    const q = document.getElementById("workSearch").value.trim().toLowerCase();
    return items.filter(i => {
      if(typeFilter && i[W.activityTypeId] !== typeFilter) return false;
      const open = isOpen(i);
      if(filter === "mine" && !(open && involvesMe(i))) return false;
      if(filter === "open" && !open) return false;
      if(filter === "waiting" && !(open && info(i).waiting)) return false;
      if(filter === "closed" && open) return false;
      if(q){
        const hay = `${i[W.title]} ${i[W.identifier] || ""} ${i[W.ownerName] || ""} ${i[W.activityName] || ""}`.toLowerCase();
        if(!hay.includes(q)) return false;
      }
      return true;
    }).sort((a, b) => filter === "closed"
      ? new Date(b[W.closedOn]) - new Date(a[W.closedOn])
      : new Date(a[W.startedOn]) - new Date(b[W.startedOn]));   // oldest first: they need attention
  }

  function render(){
    const body = document.getElementById("workBody");
    if(!body) return;

    const openItems = items.filter(isOpen);
    const counts = {
      mine: openItems.filter(involvesMe).length,
      open: openItems.length,
      waiting: openItems.filter(i => info(i).waiting).length,
      closed: items.length - openItems.length
    };
    document.querySelectorAll("#workFilters [data-filter]").forEach(b => {
      b.classList.toggle("active", b.dataset.filter === filter);
      b.setAttribute("aria-pressed", b.dataset.filter === filter ? "true" : "false");
      const c = b.querySelector(".seg-count");
      if(c) c.textContent = counts[b.dataset.filter];
    });
    const badge = document.getElementById("myWorkBadge");
    if(badge){ badge.textContent = counts.mine; badge.hidden = counts.mine === 0; }

    const chip = document.getElementById("workTypeChip");
    if(typeFilter){
      const t = Activity.types.find(x => x.Id === typeFilter);
      chip.hidden = false;
      chip.innerHTML = `Showing ${escapeHtml(t ? t.Title : "one activity")} only <button type="button" aria-label="Show all activities">&times;</button>`;
      chip.querySelector("button").addEventListener("click", () => { typeFilter = null; render(); });
    }else{
      chip.hidden = true;
    }

    const list = filtered();
    const empty = document.getElementById("workEmpty");
    empty.hidden = list.length > 0;
    empty.innerHTML = {
      mine: "<p>Nothing is assigned to you right now.</p><p class=\"small\">Checklist activities on the Dashboard have a <strong>Start</strong> button. Items you start, or steps handed to you, show up here.</p>",
      open: "<p>No checklist items are in progress.</p>",
      waiting: "<p>Nothing is waiting on another office.</p>",
      closed: `<p>Nothing was completed or cancelled in the last ${CONFIG.closedWorkDays} days.</p>`
    }[filter];
    document.getElementById("workTableWrap").hidden = list.length === 0;

    body.innerHTML = list.map(i => {
      const inf = info(i);
      const pct = inf.total ? Math.round(inf.done / inf.total * 100) : 0;
      const open = isOpen(i);
      let nextHtml;
      if(!open){
        nextHtml = `<span class="status-badge status-${i[W.status] === "Completed" ? "Completed" : "Merged"}">${escapeHtml(i[W.status])}</span>`;
      }else if(!inf.next){
        nextHtml = `<span class="muted">All steps checked. Ready to complete.</span>`;
      }else{
        nextHtml = `${escapeHtml(inf.next.text)}` +
          (inf.next.external ? ` <span class="tag-ext">Other office</span>` : "") +
          (inf.next.assignee ? ` <span class="tag-person">${isMe(inf.next.assignee) ? "You" : escapeHtml(inf.next.assigneeName || nameFromEmail(inf.next.assignee))}</span>` : "");
      }
      return `<tr data-work-id="${i.Id}" tabindex="0">
        <td><div class="work-title">${escapeHtml(i[W.employeeName] || i[W.identifier] || i[W.title] || "")}</div><div class="muted small">${escapeHtml(i[W.activityName] || "")}${i[W.employeeId] ? ` · ${escapeHtml(i[W.employeeId])}` : ""}</div></td>
        <td>${isMe(i[W.ownerEmail]) ? "<strong>You</strong>" : escapeHtml(i[W.ownerName] || "")}</td>
        <td><div class="progress" aria-label="${inf.done} of ${inf.total} steps done"><span style="width:${pct}%"></span></div><div class="muted small">${inf.done} of ${inf.total}</div></td>
        <td>${nextHtml}</td>
        <td class="nowrap">${open ? ageText(i[W.startedOn]) : formatDay(i[W.closedOn])}</td>
        <td class="nowrap muted small">${formatDay(i[W.lastActivityOn] || i[W.startedOn])}</td>
      </tr>`;
    }).join("");
    body.querySelectorAll("tr").forEach(tr => {
      const go = () => openItem(Number(tr.dataset.workId));
      tr.addEventListener("click", go);
      tr.addEventListener("keydown", e => { if(e.key === "Enter") go(); });
    });
  }

  /* ---------- item modal ---------- */

  function isModalOpen(){ return document.getElementById("workOverlay").classList.contains("active"); }

  function modalHasFocusInField(){
    const a = document.activeElement;
    return !!(a && document.getElementById("workOverlay").contains(a) &&
      (a.tagName === "INPUT" && a.type === "text" || a.tagName === "TEXTAREA" || a.tagName === "SELECT"));
  }

  async function openItem(id){
    let item = items.find(i => i.Id === id);
    if(!item){
      try{
        item = await spGetItem(CONFIG.lists.workItems, id);
        items.push(item);
      }catch(err){
        toast(`That item couldn't be opened: ${err.message}`, { type: "error" });
        return;
      }
    }
    currentId = id;
    editingStepId = null;
    history = [];
    renderModal();
    openOverlay("workOverlay");
    loadHistory(id);
  }

  function closeModal(){
    closeOverlay("workOverlay");
    currentId = null;
    editingStepId = null;
  }

  async function loadHistory(id){
    try{
      const rows = await spGetAll(CONFIG.lists.workItemLog, filterQuery(`${WL.workItemId} eq ${id}`, "$top=500"));
      if(currentId !== id) return;
      history = rows.sort((a, b) => new Date(b[WL.loggedAt]) - new Date(a[WL.loggedAt]));
      renderHistory();
    }catch(err){
      console.error("History didn't load:", err);
      const el = document.getElementById("wkHistory");
      if(el) el.innerHTML = `<p class="error-text small">History couldn't load: ${escapeHtml(err.message)}</p>`;
    }
  }

  async function addLog(itemId, action){
    // Mirror every item-history entry into the app-wide audit log.
    const it = items.find(i => i.Id === itemId);
    audit(AUDIT_AREAS.work, `${it ? (it[W.identifier] || it[W.title]) : `Item #${itemId}`}: ${action}`, {
      recordId: itemId,
      details: it ? `Activity: ${it[W.activityName] || ""}` : ""
    });
    try{
      const row = await spCreate(CONFIG.lists.workItemLog, {
        [WL.title]: `Work item ${itemId}`,
        [WL.workItemId]: itemId,
        [WL.action]: action,
        [WL.staffName]: App.user.name,
        [WL.staffEmail]: App.user.username,
        [WL.loggedAt]: new Date().toISOString()
      });
      if(itemId === currentId){ history.unshift(row); renderHistory(); }
      return row;
    }catch(err){
      console.error("Work item history entry failed:", err);
      toast(`The change saved, but its history entry didn't (${err.message}).`, { type: "error" });
      return null;
    }
  }

  function renderModal(){
    const item = items.find(i => i.Id === currentId);
    const wrap = document.getElementById("workModalBody");
    if(!item || !wrap) return;
    const inf = info(item);
    const open = isOpen(item);
    const pct = inf.total ? Math.round(inf.done / inf.total * 100) : 0;

    document.getElementById("workModalTitle").textContent = item[W.employeeName] || item[W.identifier] || item[W.title] || "Work item";
    document.getElementById("workModalSub").textContent = item[W.activityName] || "";

    wrap.innerHTML = `
      <div class="work-meta">
        <div><span class="meta-label">Owner</span>
          <span class="meta-val">${isMe(item[W.ownerEmail]) ? "You" : escapeHtml(item[W.ownerName] || "")}</span>
          ${open ? `<button type="button" class="link-btn" id="wkHandoffToggle">Hand off</button>` : ""}
        </div>
        <div><span class="meta-label">Started</span>
          <span class="meta-val">${formatDate(item[W.startedOn])} by ${escapeHtml(item[W.startedBy] || "")}</span>
        </div>
        <div><span class="meta-label">Status</span>
          <span class="meta-val">${open
            ? `Open for ${ageText(item[W.startedOn])}`
            : `${escapeHtml(item[W.status])} ${formatDate(item[W.closedOn])} by ${escapeHtml(item[W.closedBy] || "")}`}</span>
        </div>
      </div>

      <div class="handoff-panel" id="wkHandoffPanel" hidden>
        <label class="stack-label">Hand off to
          <select id="wkHandoffTo">${teamOptions("", "Choose a person")}</select>
        </label>
        <label class="stack-label grow">Note for them (optional)
          <input type="text" id="wkHandoffNote" placeholder="e.g. Waiting on HR to post; follow up Friday">
        </label>
        <button type="button" class="btn btn-navy btn-sm" id="wkHandoffBtn">Hand off</button>
      </div>

      ${Cases.sectionHtml(item, open)}

      <div class="work-progress">
        <div class="progress progress-lg"><span style="width:${pct}%"></span></div>
        <span class="muted small">${inf.done} of ${inf.total} steps done${inf.waiting && open ? " · waiting on another office" : ""}</span>
      </div>

      <ol class="steps" id="wkSteps">
        ${inf.steps.map((s, idx) => stepHtml(s, idx, inf.steps.length, open)).join("")}
      </ol>
      ${!inf.steps.length ? `<p class="muted small">This item has no steps. Add one below or mark it complete.</p>` : ""}

      ${open ? `
      <div class="add-step">
        <input type="text" id="wkNewStep" placeholder="Add a step to this item only">
        <label class="check small"><input type="checkbox" id="wkNewStepExt"> Other office</label>
        <button type="button" class="btn btn-ghost btn-sm" id="wkAddStepBtn">Add step</button>
      </div>` : ""}

      <div class="work-notes">
        <h3>Notes and history</h3>
        <div class="note-entry">
          <input type="text" id="wkNote" placeholder="Add a note, e.g. Called Civil Service, posting goes live Monday">
          <button type="button" class="btn btn-ghost btn-sm" id="wkNoteBtn">Add note</button>
        </div>
        <div id="wkHistory"></div>
      </div>`;

    const footer = document.getElementById("workModalFooter");
    footer.innerHTML = open
      ? `<button type="button" class="link-btn danger" id="wkCancelItem">Cancel this item</button>
         <span class="save-msg" id="wkSaveMsg"></span>
         <button type="button" class="btn btn-ghost" id="wkCloseBtn">Close</button>
         <button type="button" class="btn btn-navy" id="wkCompleteBtn">Mark complete</button>`
      : `<span class="save-msg" id="wkSaveMsg"></span>
         <button type="button" class="btn btn-ghost" id="wkCloseBtn">Close</button>
         <button type="button" class="btn btn-ghost" id="wkReopenBtn">Reopen</button>`;

    wireModal(item);
    renderHistory();
  }

  function stepHtml(s, idx, count, open){
    const who = s.assignee ? (isMe(s.assignee) ? "You" : (s.assigneeName || nameFromEmail(s.assignee))) : "";
    const meta = [
      s.external ? `<span class="tag-ext">Other office</span>` : "",
      who && !s.done ? `<span class="tag-person">${escapeHtml(who)}</span>` : "",
      s.done ? `<span class="muted small">Done by ${escapeHtml(s.doneBy || "")} ${formatDate(s.doneOn)}</span>` : ""
    ].filter(Boolean).join(" ");

    if(editingStepId === s.id && open){
      return `<li class="step editing" data-step-id="${s.id}">
        <div class="step-edit">
          <input type="text" class="wk-edit-text" value="${escapeHtml(s.text)}" aria-label="Step text">
          <div class="step-edit-row">
            <label class="check small"><input type="checkbox" class="wk-edit-ext" ${s.external ? "checked" : ""}> Done by another office</label>
            <label class="inline-label small">Assigned to
              <select class="wk-edit-assignee">${teamOptions(s.assignee, "Item owner")}</select>
            </label>
          </div>
          <div class="step-edit-row">
            <button type="button" class="btn btn-navy btn-sm" data-act="save">Save</button>
            <button type="button" class="btn btn-ghost btn-sm" data-act="cancel">Cancel</button>
            <span class="grow"></span>
            <button type="button" class="link-btn" data-act="up" ${idx === 0 ? "disabled" : ""}>Move up</button>
            <button type="button" class="link-btn" data-act="down" ${idx === count - 1 ? "disabled" : ""}>Move down</button>
            <button type="button" class="link-btn danger" data-act="delete">Remove step</button>
          </div>
        </div>
      </li>`;
    }
    return `<li class="step${s.done ? " done" : ""}${s.external ? " external" : ""}" data-step-id="${s.id}">
      <input type="checkbox" class="wk-check" ${s.done ? "checked" : ""} ${open ? "" : "disabled"} aria-label="${escapeHtml(s.text)}">
      <div class="step-body">
        <div class="step-text">${escapeHtml(s.text)}</div>
        ${meta ? `<div class="step-meta">${meta}</div>` : ""}
      </div>
      ${open ? `<button type="button" class="link-btn step-edit-btn" data-act="edit">Edit</button>` : ""}
    </li>`;
  }

  function renderHistory(){
    const el = document.getElementById("wkHistory");
    if(!el) return;
    if(!history.length){ el.innerHTML = `<p class="muted small">No history yet.</p>`; return; }
    el.innerHTML = history.map(h => `
      <div class="hist-item${/^Note: /.test(h[WL.action] || "") ? " note" : ""}">
        <div class="hist-when">${formatDate(h[WL.loggedAt])}</div>
        <div><strong>${escapeHtml(h[WL.staffName] || "")}</strong> ${escapeHtml(String(h[WL.action] || "").replace(/^Note: /, "noted: "))}</div>
      </div>`).join("");
  }

  function wireModal(item){
    const $ = id => document.getElementById(id);
    $("wkCloseBtn").addEventListener("click", closeModal);
    Cases.wireSection(item, saveCaseFields);

    if(!isOpen(item)){
      $("wkReopenBtn").addEventListener("click", reopen);
    }else{
      $("wkCompleteBtn").addEventListener("click", complete);
      $("wkCancelItem").addEventListener("click", cancelItem);
      $("wkHandoffToggle").addEventListener("click", () => {
        const p = $("wkHandoffPanel");
        p.hidden = !p.hidden;
        if(!p.hidden) $("wkHandoffTo").focus();
      });
      $("wkHandoffBtn").addEventListener("click", handoff);
      $("wkAddStepBtn").addEventListener("click", addStep);
      $("wkNewStep").addEventListener("keydown", e => { if(e.key === "Enter"){ e.preventDefault(); addStep(); } });
    }
    $("wkNoteBtn").addEventListener("click", addNote);
    $("wkNote").addEventListener("keydown", e => { if(e.key === "Enter"){ e.preventDefault(); addNote(); } });

    document.querySelectorAll("#wkSteps .step").forEach(li => {
      const stepId = li.dataset.stepId;
      const check = li.querySelector(".wk-check");
      if(check) check.addEventListener("change", () => toggleStep(stepId, check.checked));
      li.querySelectorAll("[data-act]").forEach(b => b.addEventListener("click", () => stepAction(stepId, b.dataset.act, li)));
      const txt = li.querySelector(".wk-edit-text");
      if(txt){
        txt.focus();
        txt.addEventListener("keydown", e => {
          if(e.key === "Enter"){ e.preventDefault(); stepAction(stepId, "save", li); }
          if(e.key === "Escape"){ e.preventDefault(); e.stopPropagation(); stepAction(stepId, "cancel", li); }
        });
      }
    });
  }

  function setBusy(on, text){
    saving = on;
    const box = document.querySelector("#workOverlay .modal");
    if(box) box.classList.toggle("busy", on);
    const msg = document.getElementById("wkSaveMsg");
    if(msg){ msg.className = "save-msg"; msg.textContent = on ? (text || "Saving...") : ""; }
  }

  /* Read the latest copy, apply the change, write it back. changeFn may
     return extra fields to write alongside the steps. */
  async function mutate(changeFn, action){
    if(saving){ renderModal(); return false; }   // snap any clicked checkbox back
    const id = currentId;
    setBusy(true);
    try{
      const fresh = await spGetItem(CONFIG.lists.workItems, id);
      const steps = stepsOf(fresh);
      const extra = changeFn(steps, fresh) || {};
      const now = new Date().toISOString();
      const body = { [W.steps]: JSON.stringify(steps), [W.lastActivityOn]: now, ...extra };
      await spUpdate(CONFIG.lists.workItems, id, body);
      Object.assign(fresh, body);
      const idx = items.findIndex(i => i.Id === id);
      if(idx >= 0) items[idx] = fresh; else items.push(fresh);
      if(action) await addLog(id, action);
      return true;
    }catch(err){
      console.error(err);
      toast(err.userMessage || `Not saved: ${err.message}`, { type: "error" });
      try{
        const fresh = await spGetItem(CONFIG.lists.workItems, id);
        const idx = items.findIndex(i => i.Id === id);
        if(idx >= 0) items[idx] = fresh;
      }catch(_){ /* keep what we have */ }
      return false;
    }finally{
      setBusy(false);
      render();
      const changed = items.find(i => i.Id === id);
      if(changed && typeof Cases !== "undefined") Cases.itemChanged(changed);
      if(currentId === id && isModalOpen()) renderModal();
      Activity.renderCounts();
    }
  }

  function missingStep(){
    const e = new Error("step missing");
    e.userMessage = "Someone else changed that step. The checklist has been refreshed.";
    return e;
  }

  function findStep(steps, stepId){
    const s = steps.find(x => x.id === stepId);
    if(!s) throw missingStep();
    return s;
  }

  async function toggleStep(stepId, checked){
    let text = "";
    let fieldNote = "";
    await mutate((steps, fresh) => {
      const s = findStep(steps, stepId);
      text = s.text;
      s.done = checked;
      s.doneBy = checked ? App.user.name : "";
      s.doneOn = checked ? new Date().toISOString() : "";
      // 2.0: a step linked to a case field updates the employee's record too.
      if(s.field){
        const f = Cases.fieldByKey(fresh, s.field);
        if(f){
          const cur = Cases.getValue(fresh, f);
          let next = cur;
          if(checked && !Cases.isFilled(cur)) next = Cases.valueForStep(f);
          if(!checked) next = "";
          if(next !== cur){
            fieldNote = ` (${Cases.describeChanges(fresh, { [s.field]: next })})`;
            return Cases.bodyFor(fresh, { [s.field]: next });
          }
        }
      }
    }, null).then(ok => { if(ok) addLog(currentId, `${checked ? "Checked" : "Unchecked"}: ${text}${fieldNote}`); });
  }

  /* 2.0: save edited case fields; linked steps follow the fields. */
  async function saveCaseFields(changes){
    const before = items.find(i => i.Id === currentId);
    const summary = before ? Cases.describeChanges(before, changes) : "";
    let stepNotes = [];
    const ok = await mutate((steps, fresh) => {
      const body = Cases.bodyFor(fresh, changes);
      stepNotes = Cases.syncSteps(steps, changes);
      return body;
    }, null);
    if(ok){
      addLog(currentId, `Updated case details: ${summary}${stepNotes.length ? `. ${stepNotes.join("; ")}` : ""}`);
      toast("Case details saved.", { type: "success" });
    }
  }

  async function stepAction(stepId, act, li){
    if(act === "edit"){ editingStepId = stepId; renderModal(); return; }
    if(act === "cancel"){ editingStepId = null; renderModal(); return; }
    if(act === "up" || act === "down"){
      let moved = "";
      await mutate(steps => {
        const i = steps.findIndex(x => x.id === stepId);
        if(i < 0) throw missingStep();
        const j = act === "up" ? i - 1 : i + 1;
        if(j < 0 || j >= steps.length) return;
        moved = steps[i].text;
        [steps[i], steps[j]] = [steps[j], steps[i]];
      }, null).then(ok => { if(ok && moved) addLog(currentId, `Moved step ${act}: ${moved}`); });
      return;
    }
    if(act === "delete"){
      const item = items.find(i => i.Id === currentId);
      const s = stepsOf(item).find(x => x.id === stepId);
      if(!confirm(`Remove the step "${s ? s.text : ""}" from this item? The activity's template isn't affected.`)) return;
      editingStepId = null;
      await mutate(steps => {
        const i = steps.findIndex(x => x.id === stepId);
        if(i < 0) throw missingStep();
        steps.splice(i, 1);
      }, `Removed step: ${s ? s.text : ""}`);
      return;
    }
    if(act === "save"){
      const text = li.querySelector(".wk-edit-text").value.trim();
      if(!text){ toast("A step needs some text.", { type: "error" }); return; }
      const external = li.querySelector(".wk-edit-ext").checked;
      const assignee = li.querySelector(".wk-edit-assignee").value;
      const changes = [];
      editingStepId = null;
      await mutate(steps => {
        const s = findStep(steps, stepId);
        if(s.text !== text) changes.push(`renamed "${s.text}" to "${text}"`);
        if(!!s.external !== external) changes.push(external ? `marked "${text}" as done by another office` : `marked "${text}" as done by the team`);
        if((s.assignee || "") !== assignee){
          changes.push(assignee ? `assigned "${text}" to ${teamMemberName(assignee)}` : `cleared the assignee on "${text}"`);
        }
        s.text = text;
        s.external = external;
        s.assignee = assignee;
        s.assigneeName = assignee ? teamMemberName(assignee) : "";
      }, null).then(ok => { if(ok && changes.length) addLog(currentId, capitalize(changes.join("; "))); });
    }
  }

  function capitalize(t){ return t.charAt(0).toUpperCase() + t.slice(1); }

  async function addStep(){
    const input = document.getElementById("wkNewStep");
    const text = input.value.trim();
    if(!text){ input.focus(); return; }
    const external = document.getElementById("wkNewStepExt").checked;
    const ok = await mutate(steps => {
      steps.push({ id: newStepId(), text, external, assignee: "", assigneeName: "", done: false, doneBy: "", doneOn: "" });
    }, `Added step: ${text}`);
    if(ok){ const again = document.getElementById("wkNewStep"); if(again) again.focus(); }
  }

  async function addNote(){
    const input = document.getElementById("wkNote");
    const text = input.value.trim();
    if(!text){ input.focus(); return; }
    input.disabled = true;
    const row = await addLog(currentId, `Note: ${text}`);
    input.disabled = false;
    if(row){ input.value = ""; input.focus(); }
    // a note counts as activity on the item
    try{
      const now = new Date().toISOString();
      await spUpdate(CONFIG.lists.workItems, currentId, { [W.lastActivityOn]: now });
      const it = items.find(i => i.Id === currentId);
      if(it) it[W.lastActivityOn] = now;
      render();
    }catch(_){ /* not important enough to surface */ }
  }

  async function handoff(){
    const sel = document.getElementById("wkHandoffTo");
    const email = sel.value;
    if(!email){ sel.focus(); return; }
    const note = document.getElementById("wkHandoffNote").value.trim();
    const item = items.find(i => i.Id === currentId);
    if(localPart(email) === localPart(item[W.ownerEmail])){ toast("They already own this item.", { type: "error" }); return; }
    const name = teamMemberName(email);
    const from = item[W.ownerName] || "the previous owner";
    const ok = await mutate(() => ({ [W.ownerName]: name, [W.ownerEmail]: email }),
      `Handed off from ${from} to ${name}${note ? `. Note: ${note}` : ""}`);
    if(ok) toast(`Handed off to ${name}. It now shows under their "Mine" filter.`, { type: "success" });
  }

  async function complete(){
    const item = items.find(i => i.Id === currentId);
    const inf = info(item);
    const left = inf.total - inf.done;
    if(left > 0 && !confirm(`${plural(left, "step")} ${left === 1 ? "isn't" : "aren't"} checked. Mark this item complete anyway?`)) return;

    // Write the Activity Log entry first so Reports count it; undo it if the item update fails.
    let logEntry = null;
    try{
      logEntry = await spCreate(CONFIG.lists.activityLog, {
        [F.title]: item[W.identifier] || item[W.activityName],
        [F.activityTypeId]: item[W.activityTypeId],
        [F.activityName]: item[W.activityName],
        [F.category]: item[W.category] || "",
        [F.identifier]: item[W.identifier] || "",
        [F.staffName]: App.user.name,
        [F.staffEmail]: App.user.username,
        [F.loggedAt]: new Date().toISOString(),
        [F.voided]: false
      });
    }catch(err){
      console.error(err);
      toast(`Not completed: ${err.message}`, { type: "error" });
      return;
    }
    const now = new Date().toISOString();
    const ok = await mutate(() => ({
      [W.status]: "Completed",
      [W.closedOn]: now,
      [W.closedBy]: App.user.name,
      [W.completionLogId]: logEntry.Id
    }), left > 0 ? `Marked complete with ${plural(left, "step")} unchecked` : "Marked complete");
    if(ok){
      Activity.ingest(logEntry);
      toast(`Completed. It counts in Reports as "${item[W.activityName]}".`, { type: "success" });
    }else{
      try{ await spUpdate(CONFIG.lists.activityLog, logEntry.Id, { [F.voided]: true, [F.voidedBy]: App.user.name, [F.voidedOn]: now }); }
      catch(e){ console.error("Couldn't roll back the completion entry", logEntry.Id, e); }
    }
  }

  async function reopen(){
    const item = items.find(i => i.Id === currentId);
    const logId = item[W.completionLogId];
    const ok = await mutate(() => ({
      [W.status]: "Open",
      [W.closedOn]: null,
      [W.closedBy]: "",
      [W.completionLogId]: null
    }), "Reopened");
    if(ok && logId){
      try{
        await spUpdate(CONFIG.lists.activityLog, logId, {
          [F.voided]: true, [F.voidedBy]: App.user.name, [F.voidedOn]: new Date().toISOString()
        });
        Activity.markVoided(logId);
      }catch(err){
        console.error(err);
        toast(`Reopened, but its completion still counts in Reports (${err.message}). Remove entry #${logId} in the Activity Log list.`, { type: "error", duration: 0 });
      }
    }
  }

  async function cancelItem(){
    const item = items.find(i => i.Id === currentId);
    const reason = prompt(`Cancel "${item[W.identifier] || item[W.title]}"? It won't count in Reports.\n\nReason (optional):`);
    if(reason === null) return;
    await mutate(() => ({
      [W.status]: "Cancelled",
      [W.closedOn]: new Date().toISOString(),
      [W.closedBy]: App.user.name
    }), `Cancelled${reason.trim() ? `: ${reason.trim()}` : ""}`);
  }

  return {
    init, refresh, start, openItem, showType, openCountsByType, itemForLog,
    get all(){ return items; },
    info, isOpen
  };
})();
