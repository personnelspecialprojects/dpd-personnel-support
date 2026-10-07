(window.PS_FILE_VERSIONS = window.PS_FILE_VERSIONS || {})["tickets.js"] = "2026.10.05-3";
/* ============================================================
   tickets.js — the Inquiries tab, Audit Log tab, and aging alert
   rules. Ported from the Secondary Employment Support Portal,
   rewired to the shared helpers in core.js.
   ============================================================ */

const Tickets = (() => {
  const R = REQ_FIELDS;
  const A = AUDIT_FIELDS;
  const AL = ALERT_FIELDS;

  let tickets = [];
  let rules = [];
  let currentId = null;
  let original = null;
  let problemUnlocked = false;
  let sortField = "Id", sortDir = -1;
  let selectedColor = PRESET_COLORS[0].hex;
  let wired = false;

  /* ---------- loading ---------- */

  async function init(){
    if(!wired) wire();
    const results = await Promise.allSettled([loadTickets(), loadRules()]);
    results.forEach(r => { if(r.status === "rejected") console.error(r.reason); });
    if(results[0].status === "rejected"){
      toast(`Inquiries couldn't load: ${results[0].reason.message}`, { type: "error", duration: 0 });
    }
    renderTickets();
  }

  async function refresh(){
    await Promise.all([loadTickets(), loadRules()]);
    renderTickets();
  }

  async function loadTickets(){
    tickets = await spGetAll(CONFIG.lists.requests, "$top=2000&$orderby=Id desc");
  }

  async function loadRules(){
    try{
      const items = await spGetAll(CONFIG.lists.alertRules, "$top=200");
      rules = items.filter(r => r[AL.active] !== false);
    }catch(err){
      console.error("Alert rules didn't load:", err);
    }
  }

  /* ---------- wiring ---------- */

  function wire(){
    wired = true;
    const on = (id, ev, fn) => document.getElementById(id).addEventListener(ev, fn);
    Object.values(LIST_UI).forEach(ui => {
      on(ui.filter, "change", renderTickets);
      on(ui.openOnly, "change", renderTickets);
      on(ui.search, "input", renderTickets);
      document.querySelectorAll(`#${ui.table} th[data-sort]`).forEach(th =>
        th.addEventListener("click", () => sortBy(th.dataset.sort)));
    });
    on("mQueueBtn", "click", async () => {
      const id = currentId;
      const t = tickets.find(x => x.Id === id);
      closeModal();
      if(t) await moveQueue(id, queueOf(t) === "special" ? "main" : "special");
    });
    on("modalClose", "click", closeModal);
    on("cancelBtn", "click", closeModal);
    on("saveBtn", "click", saveTicket);
    on("toggleMergeLink", "click", e => {
      e.preventDefault();
      const p = document.getElementById("mergePanel");
      p.hidden = !p.hidden;
    });
    on("mergeBtn", "click", mergeIntoTicket);
    on("mNoActionBtn", "click", async () => {
      const id = currentId;
      closeModal();
      await markNoAction(id);
    });
    on("problemLockBtn", "click", () => problemUnlocked ? lockProblem() : unlockProblem());

    on("openManualEntryBtn", "click", () => openManualEntry("main"));
    on("spOpenManualEntryBtn", "click", () => openManualEntry("special"));
    on("manualEntryClose", "click", () => closeOverlay("manualEntryOverlay"));
    on("meCancelBtn", "click", () => closeOverlay("manualEntryOverlay"));
    on("meSaveBtn", "click", createManualEntry);
    ["meMethodPhone", "meMethodEmail", "meMethodWalkin"].forEach(id => on(id, "change", updateMethodFields));



    // alert rules (UI lives in the Admin tab)
    on("ruleThresholdUnit", "change", () => {
      document.getElementById("excludeWeekendsWrap").hidden = document.getElementById("ruleThresholdUnit").value !== "Days";
    });
    on("createRuleBtn", "click", createRule);
    renderPalette();

    onView("dashView", renderTickets);
    onView("specialView", renderTickets);
  }

  /* ---------- helpers ---------- */

  function statusOf(t){ return t[R.status] || "New"; }

  function formatTicketId(id){
    const t = tickets.find(x => x.Id === Number(id));
    return (t && t[R.entryType] === "Manual") ? `M${id}` : `${id}`;
  }

  function deriveNameFromEmail(email){
    if(!email) return "";
    const parts = localPart(email).replace(/[0-9]+$/, "").split(/[._-]+/).filter(Boolean);
    return parts.map(s => s.charAt(0).toUpperCase() + s.slice(1)).join(" ");
  }

  function cleanEmailText(str){
    if(!str) return "";
    return str
      .replace(/\[cid:[^\]]*\]/g, "")
      .replace(/[ \t]+\n/g, "\n")
      .replace(/\n[ \t]+\n/g, "\n\n")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
  }

  function formatProblemHtml(raw){
    let escaped = escapeHtml(cleanEmailText(raw));
    escaped = escaped.replace(/\[(https?:\/\/[^\]\s]+)\]/g, '<a href="$1" target="_blank" rel="noopener">$1</a>');
    return escaped || '<em class="muted">No details provided.</em>';
  }

  /* Inquiry audit entries keep their status and note snapshots, and now go to the
     app-wide Audit Log with Area = Inquiry. */
  async function postAudit(entry){
    const row = await writeAudit(AUDIT_AREAS.inquiry, { [A.recordId]: entry[A.ticketId], ...entry });
    return !!row;
  }

  /* ---------- aging alerts ---------- */

  function ruleMoment(rule, start){
    const val = Number(rule[AL.thresholdValue]) || 0;
    const unit = rule[AL.thresholdUnit];
    if(unit === "Days" && rule[AL.excludeWeekends]){
      const d = new Date(start);
      let added = 0;
      while(added < Math.floor(val)){
        d.setDate(d.getDate() + 1);
        const day = d.getDay();
        if(day !== 0 && day !== 6) added++;
      }
      return d;
    }
    const ms = unit === "Days" ? val * 86400000 : val * 3600000;
    return new Date(new Date(start).getTime() + ms);
  }

  function alertColor(t){
    if(!t[R.receivedOn]) return null;
    const s = statusOf(t);
    if(s === "Completed" || s === "Merged" || s === NO_ACTION_STATUS) return null;
    const start = new Date(t[R.receivedOn]);
    const now = new Date();
    let worst = null;
    rules.forEach(rule => {
      const m = ruleMoment(rule, start);
      if(now >= m && (!worst || m > worst.m)) worst = { color: rule[AL.color], m };
    });
    return worst ? worst.color : null;
  }

  function ruleDescription(rule){
    const unit = rule[AL.thresholdUnit];
    const val = Number(rule[AL.thresholdValue]);
    if(unit === "Days") return `${val} ${rule[AL.excludeWeekends] ? "business " : ""}day${val === 1 ? "" : "s"}`;
    return `${val} hour${val === 1 ? "" : "s"}`;
  }

  function renderPalette(){
    const wrap = document.getElementById("colorPalette");
    wrap.innerHTML = "";
    PRESET_COLORS.forEach((c, i) => {
      const sw = document.createElement("button");
      sw.type = "button";
      sw.className = "color-swatch" + (i === 0 ? " selected" : "");
      sw.style.background = c.hex;
      sw.title = c.name;
      sw.setAttribute("aria-label", c.name);
      sw.addEventListener("click", () => {
        selectedColor = c.hex;
        wrap.querySelectorAll(".color-swatch").forEach(s => s.classList.remove("selected"));
        sw.classList.add("selected");
      });
      wrap.appendChild(sw);
    });
  }

  function renderRules(){
    const wrap = document.getElementById("rulesList");
    if(!rules.length){
      wrap.innerHTML = `<p class="muted small">No alert rules yet. Inquiries won't be color-flagged until you add one.</p>`;
      return;
    }
    const hours = r => (Number(r[AL.thresholdValue]) || 0) * (r[AL.thresholdUnit] === "Days" ? 24 : 1);
    wrap.innerHTML = rules.slice().sort((a, b) => hours(a) - hours(b)).map(r => `
      <div class="list-item">
        <span class="rule-dot" style="background:${escapeHtml(r[AL.color])}"></span>
        <span class="grow">Flag inquiries open longer than <strong>${ruleDescription(r)}</strong></span>
        <button type="button" class="link-btn danger" data-rule-id="${r.Id}">Delete</button>
      </div>`).join("");
    wrap.querySelectorAll("[data-rule-id]").forEach(b =>
      b.addEventListener("click", () => deleteRule(Number(b.dataset.ruleId))));
  }

  async function createRule(){
    const val = parseFloat(document.getElementById("ruleThresholdValue").value);
    const unit = document.getElementById("ruleThresholdUnit").value;
    const biz = unit === "Days" && document.getElementById("ruleExcludeWeekends").checked;
    if(!val || val <= 0){ toast("Enter a threshold greater than 0.", { type: "error" }); return; }
    const btn = document.getElementById("createRuleBtn");
    btn.disabled = true;
    try{
      await spCreate(CONFIG.lists.alertRules, {
        Title: `${selectedColor} at ${val} ${unit}`,
        [AL.color]: selectedColor,
        [AL.thresholdValue]: val,
        [AL.thresholdUnit]: unit,
        [AL.excludeWeekends]: biz,
        [AL.active]: true
      });
      await loadRules();
      renderRules();
      renderTickets();
      audit(AUDIT_AREAS.admin, `Created aging alert: ${PRESET_COLORS.find(c => c.hex === selectedColor)?.name || selectedColor} after ${val} ${biz ? "business " : ""}${unit.toLowerCase()}`);
      toast("Alert rule created.", { type: "success" });
    }catch(err){
      console.error(err);
      toast(`Alert rule wasn't created: ${err.message}`, { type: "error" });
    }finally{
      btn.disabled = false;
    }
  }

  async function deleteRule(id){
    if(!confirm("Delete this alert rule?")) return;
    const rule = rules.find(r => r.Id === id);
    try{
      await spDelete(CONFIG.lists.alertRules, id);
      audit(AUDIT_AREAS.admin, `Deleted aging alert: ${rule ? `${PRESET_COLORS.find(c => c.hex === rule[AL.color])?.name || rule[AL.color]} after ${ruleDescription(rule)}` : `#${id}`}`, { recordId: id });
      await loadRules();
      renderRules();
      renderTickets();
    }catch(err){
      console.error(err);
      toast(`Rule wasn't deleted: ${err.message}`, { type: "error" });
    }
  }

  /* ---------- list view ---------- */

  function sortBy(field){
    const map = { ID: "Id", Received: R.receivedOn, Title: R.title, Email: R.email, Status: R.status, CompletedBy: R.completedBy };
    const key = map[field] || field;
    if(sortField === key) sortDir *= -1; else { sortField = key; sortDir = 1; }
    renderTickets();
  }

  const STATUSES = ["New", "In Progress", "Completed", "Merged", NO_ACTION_STATUS];
  function isOpenStatus(s){ return s === "New" || s === "In Progress"; }

  /* The two inquiry lists share one implementation. "main" is the Dashboard panel;
     "special" is the Special Projects Inquiries tab. An inquiry belongs to one by its Queue value. */
  const LIST_UI = {
    main: { body: "ticketsBody", filter: "filterStatus", search: "searchBox", openOnly: "hideClosedCheckbox",
            empty: "emptyState", headCount: "inqOpenCount", badge: "openInquiryBadge", table: "ticketsTable" },
    special: { body: "spTicketsBody", filter: "spFilterStatus", search: "spSearchBox", openOnly: "spOpenOnly",
            empty: "spEmptyState", headCount: "spOpenCount", badge: "specialBadge", table: "spTicketsTable" }
  };

  function queueOf(t){
    return String(t[R.queue] || "").trim().toLowerCase() === SPECIAL_QUEUE.toLowerCase() ? "special" : "main";
  }

  function filtered(key){
    const ui = LIST_UI[key];
    const status = document.getElementById(ui.filter).value;
    const search = document.getElementById(ui.search).value.toLowerCase();
    const openOnly = document.getElementById(ui.openOnly).checked;
    return tickets.filter(t => {
      if(queueOf(t) !== key) return false;
      const s = statusOf(t);
      if(openOnly && !status && !isOpenStatus(s)) return false;   // a chosen status overrides "Open only"
      if(status && s !== status) return false;
      if(search){
        const hay = `${t.Title || ""} ${t[R.email] || ""} ${t[R.requesterName] || ""} ${formatTicketId(t.Id)}`.toLowerCase();
        if(!hay.includes(search)) return false;
      }
      return true;
    }).sort((a, b) => {
      let av = a[sortField], bv = b[sortField];
      if(av === undefined || av === null) av = "";
      if(bv === undefined || bv === null) bv = "";
      return av > bv ? sortDir : av < bv ? -sortDir : 0;
    });
  }

  /* Status dropdown shows live counts, e.g. "New (4)". Keeps the panel compact. */
  function renderStatusOptions(selId, counts){
    const sel = document.getElementById(selId);
    const current = sel.value;
    sel.innerHTML = `<option value="">All statuses</option>` +
      STATUSES.map(s => `<option value="${escapeHtml(s)}">${escapeHtml(s)} (${counts[s] || 0})</option>`).join("");
    sel.value = current;
  }

  function renderTickets(){
    renderList("main");
    renderList("special");
  }

  function renderList(key){
    const ui = LIST_UI[key];
    const body = document.getElementById(ui.body);
    if(!body) return;

    const counts = {};
    tickets.forEach(t => { if(queueOf(t) === key){ const s = statusOf(t); counts[s] = (counts[s] || 0) + 1; } });
    renderStatusOptions(ui.filter, counts);
    const open = (counts.New || 0) + (counts["In Progress"] || 0);
    const badge = document.getElementById(ui.badge);
    if(badge){ badge.textContent = open; badge.hidden = open === 0; }
    const head = document.getElementById(ui.headCount);
    if(head) head.textContent = `${open} open`;

    const list = filtered(key);
    const empty = document.getElementById(ui.empty);
    empty.hidden = list.length > 0;
    empty.innerHTML = document.getElementById(ui.openOnly).checked && !document.getElementById(ui.filter).value
      && !document.getElementById(ui.search).value
      ? (key === "special"
          ? `<p><strong>No open ${escapeHtml(SPECIAL_QUEUE)}.</strong></p><p class="small">Use the <strong>Special Projects</strong> button on a Dashboard inquiry to move it here.</p>`
          : "<p><strong>All caught up.</strong> No open inquiries.</p>")
      : "<p>No inquiries match these filters.</p>";

    body.innerHTML = list.map(t => {
      const s = statusOf(t);
      const color = alertColor(t);
      const who = t[R.requesterName] || t[R.email] || (t[R.phoneNumber] ? `Phone: ${t[R.phoneNumber]}` : "");
      const moveBtn = key === "special"
        ? `<button type="button" class="btn-noaction" data-move="${t.Id}" data-to="main" title="Move back to the Dashboard inquiries">Move back</button>`
        : `<button type="button" class="btn-noaction btn-special" data-move="${t.Id}" data-to="special" title="Move to the ${escapeHtml(SPECIAL_QUEUE)} tab">Special Projects</button>`;
      return `<tr data-ticket-id="${t.Id}" class="${color ? "alert-row" : ""}" ${color ? `style="--alert-color:${escapeHtml(color)}"` : ""} tabindex="0">
        <td class="nowrap muted">${formatTicketId(t.Id)}</td>
        <td>
          <div class="inq-title">${escapeHtml(t.Title || "(no subject)")}</div>
          <div class="muted small">${escapeHtml(who)}</div>
        </td>
        <td class="nowrap small" title="${escapeHtml(formatDate(t[R.receivedOn]))}">${escapeHtml(formatDay(t[R.receivedOn]))}<div class="muted">${t[R.receivedOn] ? formatTime(t[R.receivedOn]) : ""}</div></td>
        <td class="nowrap"><span class="status-badge status-${s.replace(/\s/g, "")}">${escapeHtml(s)}</span>
          ${!isOpenStatus(s) && t[R.completedBy] ? `<div class="muted small">${escapeHtml(t[R.completedBy])}</div>` : ""}</td>
        <td class="quick">${isOpenStatus(s)
          ? `<div class="quick-btns">${moveBtn}<button type="button" class="btn-noaction" data-noaction="${t.Id}" title="Close without a response. It doesn't count as a resolved inquiry.">No Action Required</button></div>`
          : ""}</td>
      </tr>`;
    }).join("");
    body.querySelectorAll("[data-noaction]").forEach(b => {
      b.addEventListener("click", e => {
        e.stopPropagation();          // don't open the ticket
        b.disabled = true;
        markNoAction(Number(b.dataset.noaction));
      });
      b.addEventListener("keydown", e => e.stopPropagation());
    });
    body.querySelectorAll("[data-move]").forEach(b => {
      b.addEventListener("click", e => {
        e.stopPropagation();
        b.disabled = true;
        moveQueue(Number(b.dataset.move), b.dataset.to);
      });
      b.addEventListener("keydown", e => e.stopPropagation());
    });
    body.querySelectorAll("tr").forEach(tr => {
      const go = () => openModal(Number(tr.dataset.ticketId));
      tr.addEventListener("click", go);
      tr.addEventListener("keydown", e => { if(e.key === "Enter") go(); });
    });
  }

  /* ---------- Special Projects Inquiries: move between tabs (one click, with Undo) ---------- */

  const QUEUE_LABEL = { main: "Inquiries", special: SPECIAL_QUEUE };

  async function moveQueue(id, toKey, isUndo){
    const t = tickets.find(x => x.Id === id);
    if(!t) return;
    const fromKey = queueOf(t);
    if(fromKey === toKey){ renderTickets(); return; }
    const value = toKey === "special" ? SPECIAL_QUEUE : "";
    try{
      await spUpdate(CONFIG.lists.requests, id, { [R.queue]: value });
      t[R.queue] = value;
      renderTickets();
      postAudit({
        Title: `Inquiry #${formatTicketId(id)}`,
        [A.ticketId]: id, [A.staffMember]: App.user.name, [A.logTime]: new Date().toISOString(),
        [A.action]: isUndo
          ? `Undid move: back to ${QUEUE_LABEL[toKey]}`
          : (toKey === "special" ? `Moved to ${SPECIAL_QUEUE}` : `Moved back to Inquiries (from ${SPECIAL_QUEUE})`),
        [A.previousStatus]: statusOf(t), [A.newStatus]: statusOf(t),
        [A.internalNotesSnapshot]: t[R.internalNotes] || "",
        [A.problemSnapshot]: t[R.problem] || ""
      });
      if(!isUndo){
        toast(`#${formatTicketId(id)} moved to ${QUEUE_LABEL[toKey]}.`, {
          type: "success",
          actionLabel: "Undo",
          onAction: () => moveQueue(id, fromKey, true),
          duration: CONFIG.undoSeconds * 1000
        });
      }else{
        toast(`#${formatTicketId(id)} is back in ${QUEUE_LABEL[toKey]}.`, { type: "info" });
      }
    }catch(err){
      console.error(err);
      toast(`#${formatTicketId(id)} wasn't moved: ${err.message}. If the Requests list has no Queue column yet, run the setup check in Admin.`,
        { type: "error", duration: 0 });
      renderTickets();
    }
  }

  /* ---------- ticket modal ---------- */

  function openModal(id){
    const t = tickets.find(x => x.Id === id);
    if(!t) return;
    currentId = id;
    const storedName = t[R.requesterName] || "";
    original = {
      status: statusOf(t),
      internalNotes: t[R.internalNotes] || "",
      problem: cleanEmailText(t[R.problem] || ""),
      requesterName: storedName
    };
    const $ = id => document.getElementById(id);
    $("mRequesterName").value = storedName || deriveNameFromEmail(t[R.email]);
    $("mRequesterNameHint").hidden = !!storedName;
    const inSpecial = queueOf(t) === "special";
    $("modalTitle").textContent = `Inquiry #${formatTicketId(id)}${inSpecial ? ` · ${SPECIAL_QUEUE}` : ""}`;
    $("mQueueBtn").textContent = inSpecial ? "Move back to Inquiries" : `Move to ${SPECIAL_QUEUE}`;
    $("mReq").textContent = t.Title || "";
    $("mEmail").textContent = t[R.email] || (t[R.phoneNumber] ? `Phone: ${t[R.phoneNumber]}` : "");
    $("mCreated").textContent = formatDate(t[R.receivedOn]);
    $("mProblem").innerHTML = formatProblemHtml(t[R.problem]);
    $("mProblemEdit").value = original.problem;
    lockProblem();
    $("mStatus").value = original.status;
    $("mInternal").value = original.internalNotes;
    $("mNoActionBtn").hidden = !(original.status === "New" || original.status === "In Progress");
    $("mergePanel").hidden = true;
    $("mergeSourceId").textContent = formatTicketId(id);
    populateMergeTargets(id);
    if(t[R.completed]){
      $("completedMetaRow").hidden = false;
      $("mCompletedMeta").textContent = `${t[R.completedBy] || ""} on ${formatDate(t[R.completedOn])}`;
    }else{
      $("completedMetaRow").hidden = true;
    }
    $("saveMsg").textContent = "";
    openOverlay("modalOverlay");
  }

  function closeModal(){
    closeOverlay("modalOverlay");
    currentId = null;
    lockProblem();
  }

  function lockProblem(){
    problemUnlocked = false;
    document.getElementById("mProblem").hidden = false;
    document.getElementById("mProblemEdit").hidden = true;
    document.getElementById("problemEditHint").hidden = true;
    const b = document.getElementById("problemLockBtn");
    b.textContent = "🔒";
    b.title = "Unlock to trim (not for adding notes)";
  }

  function unlockProblem(){
    problemUnlocked = true;
    document.getElementById("mProblem").hidden = true;
    document.getElementById("mProblemEdit").hidden = false;
    document.getElementById("problemEditHint").hidden = false;
    const b = document.getElementById("problemLockBtn");
    b.textContent = "🔓";
    b.title = "Unlocked — click to lock again";
    document.getElementById("mProblemEdit").focus();
  }

  function populateMergeTargets(excludeId){
    const sel = document.getElementById("mergeTargetSelect");
    sel.innerHTML = '<option value="">Select the original inquiry...</option>' +
      tickets.filter(t => t.Id !== excludeId && statusOf(t) !== "Merged" && statusOf(t) !== NO_ACTION_STATUS)
        .sort((a, b) => b.Id - a.Id)
        .map(t => `<option value="${t.Id}">#${formatTicketId(t.Id)}${queueOf(t) === "special" ? ` [${SPECIAL_QUEUE}]` : ""} — ${escapeHtml(t.Title || "(no subject)")} — ${escapeHtml(t[R.email] || t[R.requesterName] || "")}</option>`)
        .join("");
  }

  async function saveTicket(){
    const $ = id => document.getElementById(id);
    const saveBtn = $("saveBtn");
    const msg = $("saveMsg");
    saveBtn.disabled = true;
    msg.className = "save-msg";
    msg.textContent = "Saving...";

    const newStatus = $("mStatus").value;
    const newInternal = $("mInternal").value;
    const newName = $("mRequesterName").value.trim();
    const body = {
      [R.status]: newStatus,
      [R.internalNotes]: newInternal,
      [R.requesterName]: newName,
      [R.completed]: newStatus === "Completed" || newStatus === "Merged" || newStatus === NO_ACTION_STATUS
    };
    // Choosing No Action Required from the dropdown behaves like the one-click button.
    if(newStatus === NO_ACTION_STATUS && original.status !== NO_ACTION_STATUS){
      body[R.completedOn] = new Date().toISOString();
      body[R.completedBy] = App.user.name;
    }
    // Reopening a no-action inquiry clears its closure details.
    if(original.status === NO_ACTION_STATUS && (newStatus === "New" || newStatus === "In Progress")){
      body[R.completedOn] = null;
      body[R.completedBy] = "";
    }
    if(problemUnlocked) body[R.problem] = $("mProblemEdit").value;

    // Only stamp completion when the ticket is newly completed, so re-saving a
    // closed ticket doesn't move its completion date.
    const newlyCompleted = newStatus === "Completed" && original.status !== "Completed";
    if(newlyCompleted){
      body[R.completedOn] = new Date().toISOString();
      body[R.completedBy] = App.user.name;
    }

    try{
      await spUpdate(CONFIG.lists.requests, currentId, body);
      msg.className = "save-msg ok";
      msg.textContent = "Saved.";

      const changes = [];
      if(original.status !== newStatus) changes.push(`Status: ${original.status} → ${newStatus}`);
      if(original.internalNotes !== newInternal) changes.push("Notes updated");
      if(original.requesterName !== newName) changes.push("Requestor Name updated");
      if(problemUnlocked && original.problem !== $("mProblemEdit").value) changes.push("Problem field trimmed/edited");

      if(changes.length){
        await postAudit({
          Title: `Inquiry #${formatTicketId(currentId)}`,
          [A.ticketId]: currentId,
          [A.staffMember]: App.user.name,
          [A.logTime]: new Date().toISOString(),
          [A.action]: changes.join("; "),
          [A.previousStatus]: original.status,
          [A.newStatus]: newStatus,
          [A.internalNotesSnapshot]: newInternal,
          [A.problemSnapshot]: problemUnlocked ? $("mProblemEdit").value : original.problem
        });
      }
      await loadTickets();
      renderTickets();
      setTimeout(closeModal, 500);
    }catch(err){
      console.error(err);
      msg.className = "save-msg err";
      msg.textContent = `Not saved: ${err.message}`;
    }finally{
      saveBtn.disabled = false;
    }
  }

  async function mergeIntoTicket(){
    const targetId = parseInt(document.getElementById("mergeTargetSelect").value, 10);
    if(!targetId){ toast("Select the original inquiry to merge into.", { type: "error" }); return; }
    const target = tickets.find(x => x.Id === targetId);
    const source = tickets.find(x => x.Id === currentId);
    if(!target || !source) return;
    const btn = document.getElementById("mergeBtn");
    btn.disabled = true;
    btn.textContent = "Merging...";
    try{
      const now = new Date().toISOString();
      const staff = App.user.name;
      const targetNotes = `${target[R.internalNotes] ? target[R.internalNotes] + "\n\n" : ""}--- Merged from Inquiry #${formatTicketId(source.Id)} (${source.Title || "no subject"}) by ${staff} on ${formatDate(now)} ---\n${cleanEmailText(source[R.problem] || "")}`;
      await spUpdate(CONFIG.lists.requests, target.Id, { [R.internalNotes]: targetNotes });

      const sourceNotes = `${source[R.internalNotes] ? source[R.internalNotes] + "\n\n" : ""}Merged into Inquiry #${formatTicketId(target.Id)} by ${staff} on ${formatDate(now)}.`;
      await spUpdate(CONFIG.lists.requests, source.Id, {
        [R.status]: "Merged",
        [R.completed]: true,
        [R.completedOn]: now,
        [R.completedBy]: staff,
        [R.internalNotes]: sourceNotes
      });

      await postAudit({
        Title: `Inquiry #${formatTicketId(source.Id)}`,
        [A.ticketId]: source.Id, [A.staffMember]: staff, [A.logTime]: now,
        [A.action]: `Merged into Inquiry #${formatTicketId(target.Id)}`,
        [A.previousStatus]: original ? original.status : "", [A.newStatus]: "Merged",
        [A.internalNotesSnapshot]: sourceNotes,
        [A.problemSnapshot]: source[R.problem] || ""
      });
      await postAudit({
        Title: `Inquiry #${formatTicketId(target.Id)}`,
        [A.ticketId]: target.Id, [A.staffMember]: staff, [A.logTime]: now,
        [A.action]: `Received merged content from Inquiry #${formatTicketId(source.Id)}`,
        [A.previousStatus]: statusOf(target), [A.newStatus]: statusOf(target),
        [A.internalNotesSnapshot]: targetNotes,
        [A.problemSnapshot]: target[R.problem] || ""
      });
      await loadTickets();
      renderTickets();
      closeModal();
      toast(`Merged into #${formatTicketId(target.Id)}.`, { type: "success" });
    }catch(err){
      console.error(err);
      toast(`Merge didn't finish: ${err.message}`, { type: "error", duration: 0 });
    }finally{
      btn.disabled = false;
      btn.textContent = "Merge and close this inquiry";
    }
  }

  /* ---------- No Action Required (one click, with Undo) ---------- */

  async function markNoAction(id){
    const t = tickets.find(x => x.Id === id);
    if(!t) return;
    const prev = {
      status: statusOf(t),
      completed: !!t[R.completed],
      completedOn: t[R.completedOn] || null,
      completedBy: t[R.completedBy] || ""
    };
    if(prev.status === NO_ACTION_STATUS) return;
    const now = new Date().toISOString();
    const body = {
      [R.status]: NO_ACTION_STATUS,
      [R.completed]: true,
      [R.completedOn]: now,
      [R.completedBy]: App.user.name
    };
    try{
      await spUpdate(CONFIG.lists.requests, id, body);
      Object.assign(t, body);
      renderTickets();
      postAudit({
        Title: `Inquiry #${formatTicketId(id)}`,
        [A.ticketId]: id, [A.staffMember]: App.user.name, [A.logTime]: now,
        [A.action]: `Marked ${NO_ACTION_STATUS}`,
        [A.previousStatus]: prev.status, [A.newStatus]: NO_ACTION_STATUS,
        [A.internalNotesSnapshot]: t[R.internalNotes] || "",
        [A.problemSnapshot]: t[R.problem] || ""
      });
      toast(`#${formatTicketId(id)} marked ${NO_ACTION_STATUS}.`, {
        type: "success",
        actionLabel: "Undo",
        onAction: () => undoNoAction(id, prev),
        duration: CONFIG.undoSeconds * 1000
      });
    }catch(err){
      console.error(err);
      toast(`#${formatTicketId(id)} wasn't updated: ${err.message}`, { type: "error" });
      renderTickets();
    }
  }

  async function undoNoAction(id, prev){
    const t = tickets.find(x => x.Id === id);
    const body = {
      [R.status]: prev.status,
      [R.completed]: prev.completed,
      [R.completedOn]: prev.completedOn,
      [R.completedBy]: prev.completedBy
    };
    try{
      await spUpdate(CONFIG.lists.requests, id, body);
      if(t) Object.assign(t, body);
      renderTickets();
      await postAudit({
        Title: `Inquiry #${formatTicketId(id)}`,
        [A.ticketId]: id, [A.staffMember]: App.user.name, [A.logTime]: new Date().toISOString(),
        [A.action]: `Undid ${NO_ACTION_STATUS}`,
        [A.previousStatus]: NO_ACTION_STATUS, [A.newStatus]: prev.status,
        [A.internalNotesSnapshot]: t ? t[R.internalNotes] || "" : "",
        [A.problemSnapshot]: t ? t[R.problem] || "" : ""
      });
      toast(`#${formatTicketId(id)} is back in the queue.`, { type: "info" });
    }catch(err){
      console.error(err);
      toast(`Undo didn't work: ${err.message}. Open the inquiry and change its status.`, { type: "error", duration: 0 });
    }
  }

  /* ---------- manual entry ---------- */

  let manualQueue = "main";

  function openManualEntry(queueKey){
    manualQueue = queueKey === "special" ? "special" : "main";
    const $ = id => document.getElementById(id);
    $("meTitleHeading").textContent = manualQueue === "special" ? `New manual inquiry (${SPECIAL_QUEUE})` : "New manual inquiry";
    $("meMethodPhone").checked = true;
    ["meRequesterName", "mePhoneNumber", "meEmail", "meTitle", "meDescription", "meInternal"]
      .forEach(id => { $(id).value = ""; });
    $("meStatus").value = "New";
    $("meSaveMsg").textContent = "";
    updateMethodFields();
    openOverlay("manualEntryOverlay");
    $("meRequesterName").focus();
  }

  function updateMethodFields(){
    const isPhone = document.getElementById("meMethodPhone").checked;
    document.getElementById("mePhoneRow").hidden = !isPhone;
    document.getElementById("meEmailRow").hidden = isPhone;
  }

  async function createManualEntry(){
    const $ = id => document.getElementById(id);
    const method = [...document.getElementsByName("meMethod")].find(r => r.checked)?.value || "Phone";
    const isPhone = method === "Phone";
    const name = $("meRequesterName").value.trim();
    const phone = $("mePhoneNumber").value.trim();
    const email = $("meEmail").value.trim();
    const title = $("meTitle").value.trim();
    const description = $("meDescription").value.trim();
    const status = $("meStatus").value;
    const internal = $("meInternal").value;

    if(!name || !title || !description){ toast("Requestor name, problem title, and description are required.", { type: "error" }); return; }
    if(isPhone && !phone){ toast("Enter a phone number, or switch the support method to Email or Walk-in.", { type: "error" }); return; }
    if(!isPhone && !email){ toast(`Enter an email address for this ${method.toLowerCase()} contact.`, { type: "error" }); return; }

    const btn = $("meSaveBtn");
    const msg = $("meSaveMsg");
    btn.disabled = true;
    msg.className = "save-msg";
    msg.textContent = "Creating...";
    const now = new Date().toISOString();
    const body = {
      Title: title,
      [R.problem]: description,
      [R.email]: isPhone ? "" : email,
      [R.receivedOn]: now,
      [R.status]: status,
      [R.internalNotes]: internal,
      [R.entryType]: "Manual",
      [R.source]: method,
      [R.requesterName]: name,
      [R.phoneNumber]: isPhone ? phone : "",
      [R.completed]: status === "Completed"
    };
    if(manualQueue === "special") body[R.queue] = SPECIAL_QUEUE;
    if(status === "Completed"){
      body[R.completedOn] = now;
      body[R.completedBy] = App.user.name;
    }
    try{
      const created = await spCreate(CONFIG.lists.requests, body);
      await postAudit({
        Title: `Inquiry #M${created.Id}`,
        [A.ticketId]: created.Id, [A.staffMember]: App.user.name, [A.logTime]: now,
        [A.action]: `Manual inquiry created (${method})${manualQueue === "special" ? ` in ${SPECIAL_QUEUE}` : ""}${status === "Completed" ? " and closed" : ""}`,
        [A.previousStatus]: "", [A.newStatus]: status,
        [A.internalNotesSnapshot]: internal,
        [A.problemSnapshot]: description
      });
      await loadTickets();
      renderTickets();
      msg.className = "save-msg ok";
      msg.textContent = "Created.";
      setTimeout(() => closeOverlay("manualEntryOverlay"), 500);
    }catch(err){
      console.error(err);
      msg.className = "save-msg err";
      msg.textContent = `Not created: ${err.message}`;
    }finally{
      btn.disabled = false;
    }
  }

  return {
    init, refresh, renderRules, statusOf, queueOf, formatId: formatTicketId,
    get all(){ return tickets; }
  };
})();
