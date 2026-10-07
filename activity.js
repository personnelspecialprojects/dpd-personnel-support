(window.PS_FILE_VERSIONS = window.PS_FILE_VERSIONS || {})["activity.js"] = "2026.10.05-3";
/* ============================================================
   activity.js — the Log Activity tab.
   One row per active activity type. Type an ID, press Enter.
   Several IDs separated by spaces/commas/new lines log at once.
   ============================================================ */

const Activity = (() => {
  const T = TYPE_FIELDS;
  const F = LOG_FIELDS;
  let allTypes = [];
  let recent = [];          // every entry from the last CONFIG.recentDays days (whole team)
  let boardSignature = "";
  let currentTeam = null;    // which team's activities the board shows (SRU / PSU)
  const busy = new Set();

  /* ---------- loading ---------- */

  async function init(){
    initTeam();
    await Promise.all([loadTypes(), loadRecent()]);
    buildBoard();
    renderCounts();
    renderMine();
    onView("dashView", () => { renderCounts(); renderMine(); });
  }

  async function refresh(){
    await Promise.all([loadTypes(), loadRecent()]);
    buildBoard();          // only rebuilds if the activity list actually changed
    renderCounts();
    renderMine();
  }

  async function loadTypes(){
    allTypes = await spGetAll(CONFIG.lists.activityTypes, "$top=500");
  }

  async function reloadTypes(){
    await loadTypes();
    buildBoard();
    renderCounts();
  }

  async function loadRecent(){
    const since = addDays(startOfToday(), -(CONFIG.recentDays - 1));
    recent = await spGetAll(CONFIG.lists.activityLog,
      filterQuery(`${F.loggedAt} ge ${odataDate(since)}`, "$top=2000"));
  }

  /* ---------- type helpers ---------- */

  function isActive(t){ return t[T.active] !== false; }

  function inputMode(t){
    const v = String(t[T.inputType] || "Employee Number").toLowerCase();
    if(v.startsWith("click")) return "click";
    if(v.startsWith("ref")) return "reference";
    if(v.startsWith("count")) return "count";
    return "employee";
  }

  /* How many items an entry represents. Entries without a Quantity count as 1. */
  function qty(e){
    const n = Number(e && e[F.quantity]);
    return Number.isFinite(n) && n > 0 ? n : 1;
  }

  function inputLabel(t){
    if(t[T.inputLabel]) return t[T.inputLabel];
    const mode = inputMode(t);
    return mode === "employee" ? "Employee #" : mode === "count" ? "How many?" : "Reference";
  }

  /* Checklist template on an activity type. Stored as JSON [{text, external}], but a
     plain list (one step per line, "[other]" prefix for other-office steps) also works,
     so the column can be edited directly in SharePoint if needed. */
  function templateSteps(t){
    const raw = t && t[T.checklistSteps];
    if(!raw || !String(raw).trim()) return [];
    try{
      const arr = JSON.parse(raw);
      if(Array.isArray(arr)){
        return arr.map(x => ({ text: String((x && x.text) || "").trim(), external: !!(x && x.external),
            ...(x && x.field ? { field: String(x.field) } : {}) }))
          .filter(x => x.text);
      }
    }catch(_){ /* not JSON — fall through to line format */ }
    return String(raw).split(/\r?\n/).map(l => l.trim()).filter(Boolean).map(l => {
      const ext = /^\[(other|external)\]/i.test(l);
      return { text: l.replace(/^\[(other|external)\]\s*/i, ""), external: ext };
    });
  }

  function hasChecklist(t){
    if(inputMode(t) === "count") return false;
    // 2.0: employee processes always run as cases, even before steps are added.
    if(typeof Cases !== "undefined" && Cases.isEmployeeProcess(t)) return true;
    return templateSteps(t).length > 0;
  }

  function windowDays(t){
    const v = t[T.duplicateWindowDays];
    if(v === null || v === undefined || v === "") return CONFIG.defaultDuplicateWindowDays;
    return Number(v) || 0;
  }

  function sortNum(t){
    const n = Number(t[T.sortOrder]);
    return Number.isFinite(n) && t[T.sortOrder] !== null && t[T.sortOrder] !== "" ? n : 9999;
  }

  function sortedTypes(list){
    return list.slice().sort((a, b) =>
      sortNum(a) - sortNum(b) || String(a.Title || "").localeCompare(String(b.Title || "")));
  }

  /* Categories appear in the order of their lowest-numbered activity. */
  function groupByCategory(types){
    const groups = [];
    const index = {};
    sortedTypes(types).forEach(t => {
      const name = (t[T.category] || "General").trim() || "General";
      if(!(name in index)){ index[name] = groups.length; groups.push({ name, types: [] }); }
      groups[index[name]].types.push(t);
    });
    return groups;
  }

  function isMine(e){ return localPart(e[F.staffEmail]) === App.user.key; }

  /* ---------- team toggle (SRU / PSU) ---------- */

  /* An activity's team; blank or unrecognized counts as shared by both teams. */
  function teamOf(t){
    const v = t && t[T.team];
    return TEAMS.includes(v) ? v : TEAM_BOTH;
  }

  function showsOnTeam(t, team){
    const v = teamOf(t);
    return v === TEAM_BOTH || v === team;
  }

  function teamStorageKey(){ return `ps-dashboard-team:${App.user ? App.user.key : ""}`; }

  /* Last choice on this browser, else the person's team from the roster, else the first team. */
  function initTeam(){
    let saved = null;
    try{ saved = localStorage.getItem(teamStorageKey()); }catch(_){ /* storage blocked: fine */ }
    currentTeam = TEAMS.includes(saved) ? saved : (App.myTeam || TEAMS[0]);
    renderTeamToggle();
  }

  function renderTeamToggle(){
    const wrap = document.getElementById("teamToggle");
    if(!wrap) return;
    wrap.innerHTML = TEAMS.map(tm => `<button type="button" data-team="${escapeHtml(tm)}"
      class="${tm === currentTeam ? "active" : ""}" aria-pressed="${tm === currentTeam}">${escapeHtml(tm)}</button>`).join("");
    wrap.querySelectorAll("[data-team]").forEach(b => b.addEventListener("click", () => setTeam(b.dataset.team)));
  }

  function setTeam(team){
    if(!TEAMS.includes(team) || team === currentTeam) return;
    currentTeam = team;
    try{ localStorage.setItem(teamStorageKey(), team); }catch(_){}
    renderTeamToggle();
    buildBoard();
    renderCounts();
    const first = document.querySelector("#board .task-input");
    if(first) first.focus();
  }

  /* ---------- board ---------- */

  function buildBoard(){
    const anyActive = allTypes.some(isActive);
    const types = allTypes.filter(t => isActive(t) && showsOnTeam(t, currentTeam));
    const sig = currentTeam + JSON.stringify(sortedTypes(types).map(t =>
      [t.Id, t.Title, t[T.category], t[T.inputType], t[T.inputLabel], t[T.description], t[T.sortOrder], hasChecklist(t)]));
    if(sig === boardSignature) return;
    boardSignature = sig;

    const board = document.getElementById("board");
    // keep anything half-typed across a rebuild (e.g. an admin edits the list mid-shift)
    const saved = {};
    board.querySelectorAll(".task-input").forEach(i => { if(i.value) saved[i.dataset.typeId] = i.value; });
    const focusedId = document.activeElement && document.activeElement.classList.contains("task-input")
      ? document.activeElement.dataset.typeId : null;

    if(!types.length){
      board.innerHTML = `<div class="empty-block">
        <p><strong>${anyActive ? `No ${escapeHtml(currentTeam)} activities are set up yet.` : "No activities are set up yet."}</strong></p>
        <p>${App.isAdmin
          ? 'Add the tasks your team completes in the <a href="#" id="goAdminLink">Admin tab</a>. Each one becomes a row here.'
          : `Ask ${escapeHtml(CONFIG.adminContact)} to add your team's tasks.`}</p></div>`;
      const link = document.getElementById("goAdminLink");
      if(link) link.addEventListener("click", e => { e.preventDefault(); switchTab("adminView"); });
      return;
    }

    board.innerHTML = `
      <div class="board-colhead" aria-hidden="true">
        <span>Activity</span><span class="c-today" title="Logged today: by you / by the whole team">Today</span><span>Log it</span>
      </div>
      ${groupByCategory(types).map(g => `
        <section class="task-group">
          <h3 class="group-name">${escapeHtml(g.name)}</h3>
          ${g.types.map(rowHtml).join("")}
        </section>`).join("")}`;

    board.querySelectorAll(".task-log").forEach(btn =>
      btn.addEventListener("click", () => submit(Number(btn.dataset.typeId))));
    board.querySelectorAll(".task-input").forEach(input => {
      input.addEventListener("keydown", e => {
        if(e.key === "Enter"){ e.preventDefault(); submit(Number(input.dataset.typeId)); }
      });
      input.addEventListener("input", () => clearFeedback(Number(input.dataset.typeId), true));
      // Single-line inputs silently drop line breaks, which would glue a pasted
      // Excel column into one long number. Convert breaks/tabs to spaces instead.
      input.addEventListener("paste", e => {
        const text = e.clipboardData && e.clipboardData.getData("text");
        if(!text || !/[\r\n\t]/.test(text)) return;
        e.preventDefault();
        const clean = text.replace(/[\r\n\t]+/g, " ").trim();
        const start = input.selectionStart ?? input.value.length;
        const end = input.selectionEnd ?? input.value.length;
        const before = input.value.slice(0, start);
        const sep = before && !/\s$/.test(before) ? " " : "";
        input.value = before + sep + clean + input.value.slice(end);
        const caret = (before + sep + clean).length;
        input.setSelectionRange(caret, caret);
        clearFeedback(Number(input.dataset.typeId), true);
      });
      if(saved[input.dataset.typeId]) input.value = saved[input.dataset.typeId];
      if(focusedId === input.dataset.typeId) input.focus();
    });
  }

  function rowHtml(t){
    const mode = inputMode(t);
    const label = inputLabel(t);
    const checklist = hasChecklist(t);
    const empProc = typeof Cases !== "undefined" && Cases.isEmployeeProcess(t);
    const control = checklist
      ? (mode === "click" && !empProc
          ? `<button type="button" class="btn btn-navy task-log task-log-wide" data-type-id="${t.Id}">Start one</button>`
          : `<input class="task-input" data-type-id="${t.Id}" type="text" autocomplete="off" spellcheck="false"
                ${mode === "employee" || empProc ? 'inputmode="numeric"' : ""}
                placeholder="${escapeHtml(empProc ? "Employee #" : label)}" aria-label="${escapeHtml(`${t.Title}: ${empProc ? "Employee #" : label}`)}">
             <button type="button" class="btn btn-navy task-log" data-type-id="${t.Id}">Start</button>`)
      : mode === "click"
      ? `<button type="button" class="btn btn-navy task-log task-log-wide" data-type-id="${t.Id}">Log one</button>`
      : mode === "count"
      ? `<input class="task-input task-input-count" data-type-id="${t.Id}" type="text" inputmode="numeric" autocomplete="off"
            placeholder="${escapeHtml(label)}" aria-label="${escapeHtml(`${t.Title}: ${label}`)}">
         <button type="button" class="btn btn-navy task-log" data-type-id="${t.Id}">Log</button>`
      : `<input class="task-input" data-type-id="${t.Id}" type="text" autocomplete="off" spellcheck="false"
            ${mode === "employee" ? 'inputmode="numeric"' : ""}
            placeholder="${escapeHtml(label)}" aria-label="${escapeHtml(`${t.Title}: ${label}`)}">
         <button type="button" class="btn btn-navy task-log" data-type-id="${t.Id}">Log</button>`;
    return `
      <div class="task-row" data-row-for="${t.Id}">
        <div class="task-main">
          <div class="task-name">${escapeHtml(t.Title || "")}</div>
          ${t[T.description] ? `<div class="task-desc">${escapeHtml(t[T.description])}</div>` : ""}
          ${checklist ? `<div class="task-sub"><span class="checklist-tag">${empProc ? "Employee process" : "Checklist"} · ${plural(templateSteps(t).length, "step")}</span><span data-open-for="${t.Id}"></span></div>` : ""}
        </div>
        <div class="task-count" data-count-for="${t.Id}"></div>
        <div class="task-entry">${control}</div>
        <div class="task-feedback" data-feedback-for="${t.Id}" aria-live="polite"></div>
      </div>`;
  }

  /* ---------- counts & my entries ---------- */

  function renderCounts(){
    const today = startOfToday();
    const counts = {};
    let meTotal = 0, teamTotal = 0;
    recent.forEach(e => {
      if(e[F.voided]) return;
      if(new Date(e[F.loggedAt]) < today) return;
      const c = counts[e[F.activityTypeId]] || (counts[e[F.activityTypeId]] = { me: 0, team: 0 });
      const n = qty(e);
      c.team += n; teamTotal += n;
      if(isMine(e)){ c.me += n; meTotal += n; }
    });
    document.querySelectorAll("[data-count-for]").forEach(el => {
      const c = counts[el.dataset.countFor] || { me: 0, team: 0 };
      el.innerHTML = `<span class="count-me${c.me ? " has" : ""}" title="Logged by you today">${c.me}</span>` +
        `<span class="count-sep">/</span><span class="count-team" title="Logged by the team today">${c.team}</span>`;
    });
    const open = (typeof Work !== "undefined") ? Work.openCountsByType() : {};
    document.querySelectorAll("[data-open-for]").forEach(el => {
      const n = open[el.dataset.openFor] || 0;
      el.innerHTML = n
        ? `<button type="button" class="link-btn open-link" data-open-type="${el.dataset.openFor}">${n} in progress</button>`
        : "";
      const b = el.querySelector("button");
      if(b) b.addEventListener("click", () => Work.showType(Number(b.dataset.openType)));
    });
    const summary = document.getElementById("todaySummary");
    if(summary){
      summary.textContent = `You've logged ${plural(meTotal, "item")} today. Team total: ${teamTotal}.`;
    }
  }

  function renderMine(){
    const wrap = document.getElementById("myRecent");
    if(!wrap) return;
    const mine = recent.filter(isMine)
      .sort((a, b) => new Date(b[F.loggedAt]) - new Date(a[F.loggedAt]))
      .slice(0, 60);
    if(!mine.length){
      wrap.innerHTML = `<p class="muted small">Nothing logged in the last ${CONFIG.recentDays} days. Entries you log show up here so you can remove a mistake.</p>`;
      return;
    }
    let lastDay = "";
    wrap.innerHTML = mine.map(e => {
      const day = formatDay(e[F.loggedAt]);
      const header = day !== lastDay ? `<div class="mine-day">${escapeHtml(day)}</div>` : "";
      lastDay = day;
      const voided = !!e[F.voided];
      return `${header}
        <div class="mine-item${voided ? " voided" : ""}">
          <div class="mine-text">
            <div class="mine-name">${escapeHtml(e[F.activityName] || "")}</div>
            <div class="mine-meta">${qty(e) > 1 || e[F.quantity] ? `<span class="mine-id">${qty(e).toLocaleString()} items</span> ` : ""}${e[F.identifier] ? `<span class="mine-id">${escapeHtml(e[F.identifier])}</span> ` : ""}${formatTime(e[F.loggedAt])}</div>
          </div>
          ${voided
            ? `<span class="mine-removed">Removed</span>`
            : (typeof Work !== "undefined" && Work.itemForLog(e.Id))
              ? `<button type="button" class="link-btn mine-open" data-work-id="${Work.itemForLog(e.Id).Id}" title="Completed from a checklist. Reopen the item to undo.">Checklist</button>`
              : `<button type="button" class="link-btn mine-remove" data-entry-id="${e.Id}">Remove</button>`}
        </div>`;
    }).join("");
    wrap.querySelectorAll(".mine-open").forEach(btn => btn.addEventListener("click", () =>
      Work.openItem(Number(btn.dataset.workId))));
    wrap.querySelectorAll(".mine-remove").forEach(btn => btn.addEventListener("click", () => {
      const entry = recent.find(e => e.Id === Number(btn.dataset.entryId));
      if(!entry) return;
      const what = entry[F.identifier] ? `${entry[F.activityName]} for ${entry[F.identifier]}` : entry[F.activityName];
      if(confirm(`Remove "${what}" from your log? It will no longer count in reports.`)) voidEntries([entry]);
    }));
  }

  /* ---------- feedback area under each row ---------- */

  function feedbackEl(typeId){ return document.querySelector(`[data-feedback-for="${typeId}"]`); }

  function showFeedback(typeId, message, kind){
    const el = feedbackEl(typeId);
    if(!el) return;
    el.className = `task-feedback show ${kind || ""}`;
    el.textContent = message;
  }

  function clearFeedback(typeId, onlyErrors){
    const el = feedbackEl(typeId);
    if(!el) return;
    if(onlyErrors && !el.classList.contains("error")) return;
    el.className = "task-feedback";
    el.innerHTML = "";
  }

  function setRowBusy(typeId, on){
    const row = document.querySelector(`[data-row-for="${typeId}"]`);
    if(!row) return;
    row.classList.toggle("busy", on);
    row.querySelectorAll("button, input").forEach(el => { el.disabled = on; });
  }

  function flashRow(typeId){
    const row = document.querySelector(`[data-row-for="${typeId}"]`);
    if(!row) return;
    row.classList.remove("flash");
    void row.offsetWidth;   // restart the animation
    row.classList.add("flash");
  }

  /* ---------- submit ---------- */

  function parseIds(raw){
    return [...new Set(String(raw).split(/[\s,;]+/).map(s => s.trim()).filter(Boolean))];
  }

  async function submit(typeId){
    if(busy.has(typeId)) return;
    const type = allTypes.find(t => t.Id === typeId);
    if(!type) return;
    const mode = inputMode(type);
    const input = document.querySelector(`.task-input[data-type-id="${typeId}"]`);

    if(hasChecklist(type)){
      await startChecklist(type, input, false);
      return;
    }
    if(mode === "count"){
      await submitCount(type, input);
      return;
    }

    let ids = [""];
    if(mode !== "click"){
      ids = parseIds(input.value);
      if(!ids.length){
        showFeedback(typeId, `Enter ${inputLabel(type)} first.`, "error");
        input.focus();
        return;
      }
      if(mode === "employee"){
        const bad = ids.filter(x => !/^\d+$/.test(x));
        if(bad.length){
          showFeedback(typeId, `Employee numbers are digits only. Check: ${bad.join(", ")}`, "error");
          input.focus();
          return;
        }
      }
      if(ids.length > CONFIG.maxBatch){
        showFeedback(typeId, `That's ${ids.length} entries. Log up to ${CONFIG.maxBatch} at a time.`, "error");
        return;
      }
    }

    busy.add(typeId);
    setRowBusy(typeId, true);
    clearFeedback(typeId);
    let keepFocus = true;
    try{
      let clean = ids;
      let dupes = [];
      if(mode !== "click"){
        const found = await findDuplicates(type, ids);
        dupes = ids.filter(x => found.has(x)).map(x => ({ idf: x, prior: found.get(x) }));
        clean = ids.filter(x => !found.has(x));
      }

      const { created, failed } = await createEntries(type, clean);
      if(input) input.value = failed.map(f => f.idf).join(" ");

      if(created.length){
        announceLogged(type, created);
        auditLogged(type, created);
      }
      if(failed.length){
        showFeedback(typeId,
          `${plural(failed.length, "entry", "entries")} didn't save (${failed[0].err.message}). They're still in the box — press Log to try again.`,
          "error");
      } else if(dupes.length){
        showDuplicatePrompt(type, dupes);
        keepFocus = false;
      }
    }catch(err){
      console.error(err);
      showFeedback(typeId, `Not logged: ${err.message}`, "error");
    }finally{
      busy.delete(typeId);
      setRowBusy(typeId, false);
      if(keepFocus && input) input.focus();
    }
  }

  /* ---------- count activities: one entry worth N items ---------- */

  async function submitCount(type, input){
    const raw = input.value.trim().replace(/,/g, "");
    if(!raw){
      showFeedback(type.Id, "Enter how many you completed.", "error");
      input.focus();
      return;
    }
    if(!/^\d+$/.test(raw) || Number(raw) < 1){
      showFeedback(type.Id, "Enter a whole number, like 50 or 115.", "error");
      input.focus();
      return;
    }
    const n = Number(raw);
    if(n > CONFIG.maxCount){
      showFeedback(type.Id, `That's more than ${CONFIG.maxCount.toLocaleString()} in one entry. Check the number, or log it in parts.`, "error");
      return;
    }
    if(busy.has(type.Id)) return;
    busy.add(type.Id);
    setRowBusy(type.Id, true);
    clearFeedback(type.Id);
    try{
      const item = await spCreate(CONFIG.lists.activityLog, {
        [F.title]: `${n} × ${type.Title}`.slice(0, 255),
        [F.activityTypeId]: type.Id,
        [F.activityName]: type.Title,
        [F.category]: type[T.category] || "",
        [F.identifier]: "",
        [F.quantity]: n,
        [F.staffName]: App.user.name,
        [F.staffEmail]: App.user.username,
        [F.loggedAt]: new Date().toISOString(),
        [F.voided]: false
      });
      recent.unshift(item);
      input.value = "";
      flashRow(type.Id);
      renderCounts();
      renderMine();
      toast(`Logged ${n.toLocaleString()} for "${type.Title}".`, {
        type: "success",
        actionLabel: "Undo",
        onAction: () => voidEntries([item]),
        duration: CONFIG.undoSeconds * 1000
      });
      audit(AUDIT_AREAS.activity, `Logged "${type.Title}": ${n.toLocaleString()} completed`, {
        recordId: item.Id, details: `Activity Log entry #${item.Id}, quantity ${n}`
      });
    }catch(err){
      console.error(err);
      showFeedback(type.Id, `Not logged: ${err.message}. The number is still in the box; press Log to try again.`, "error");
    }finally{
      busy.delete(type.Id);
      setRowBusy(type.Id, false);
      input.focus();
    }
  }

  /* ---------- checklist activities: Start creates a work item ---------- */

  async function startChecklist(type, input, force, presetValue, notInRoster){
    const isEmpProcess = typeof Cases !== "undefined" && Cases.isEmployeeProcess(type);
    const mode = isEmpProcess ? "employee" : inputMode(type);   // employee processes always take an employee #
    let value = presetValue !== undefined ? presetValue : (input ? input.value.trim() : "");
    if(mode !== "click"){
      if(!value){
        showFeedback(type.Id, `Enter ${inputLabel(type)} first.`, "error");
        if(input) input.focus();
        return;
      }
      // References can be titles with spaces ("Payroll Clerk 26-114"), so only
      // employee-number checklists are checked for a single, digits-only value.
      if(mode === "employee"){
        if(parseIds(value).length > 1){
          showFeedback(type.Id, "Checklist items start one at a time. Enter a single employee number.", "error");
          return;
        }
        if(!/^\d+$/.test(value)){
          showFeedback(type.Id, "Employee numbers are digits only.", "error");
          return;
        }
      }
      value = value.replace(/\s+/g, " ").slice(0, 200);
    }
    if(busy.has(type.Id)) return;
    busy.add(type.Id);
    setRowBusy(type.Id, true);
    clearFeedback(type.Id);
    let keepFocus = true;
    try{
      let result;
      if(isEmpProcess){
        // 2.0: look the employee up in the roster and tie the case to them.
        await Roster.ready();
        const emp = Roster.find(value);
        if(!emp && !notInRoster){
          showNotInRoster(type, input, value, force);
          keepFocus = false;
          return;
        }
        result = await Work.start(type, value, { force, employee: emp || { [EMP_FIELDS.employeeId]: value } });
      }else{
        result = await Work.start(type, value, { force });
      }
      if(result.duplicate){
        showStartDuplicate(type, value, result.duplicate);
        keepFocus = false;
        return;
      }
      if(input) input.value = "";
      flashRow(type.Id);
      renderCounts();
      toast(`Started "${result.item[WORK_FIELDS.title]}". You're the owner.`, {
        type: "success",
        actionLabel: "Open checklist",
        onAction: () => Work.openItem(result.item.Id),
        duration: 8000
      });
    }catch(err){
      console.error(err);
      showFeedback(type.Id, `Not started: ${err.message}`, "error");
    }finally{
      busy.delete(type.Id);
      setRowBusy(type.Id, false);
      if(keepFocus && input) input.focus();
    }
  }

  function showNotInRoster(type, input, value, force){
    const el = feedbackEl(type.Id);
    if(!el) return;
    el.className = "task-feedback show warn";
    el.innerHTML = `
      <div class="dupe-text">Employee <strong>${escapeHtml(value)}</strong> isn't in the roster${Roster.count() ? "" : " (no roster has been imported yet)"}. Check the number, or start the case anyway.</div>
      <div class="dupe-actions">
        <button type="button" class="btn btn-sm btn-navy" data-act="go">Start anyway</button>
        <button type="button" class="btn btn-sm btn-ghost" data-act="skip">Cancel</button>
      </div>`;
    el.querySelector('[data-act="go"]').addEventListener("click", () => { clearFeedback(type.Id); startChecklist(type, input, force, value, true); });
    el.querySelector('[data-act="skip"]').addEventListener("click", () => { clearFeedback(type.Id); if(input) input.focus(); });
    el.querySelector('[data-act="go"]').focus();
  }

  function showStartDuplicate(type, value, existing){
    const el = feedbackEl(type.Id);
    if(!el) return;
    const W = WORK_FIELDS;
    const owner = localPart(existing[W.ownerEmail]) === App.user.key ? "you" : escapeHtml(existing[W.ownerName] || "someone");
    el.className = "task-feedback show warn";
    el.innerHTML = `
      <div class="dupe-text"><strong>${escapeHtml(value)}</strong> is already in progress, owned by ${owner}, started ${formatDate(existing[W.startedOn])}.</div>
      <div class="dupe-actions">
        <button type="button" class="btn btn-sm btn-navy" data-act="open">Open it</button>
        <button type="button" class="btn btn-sm btn-ghost" data-act="again">Start another</button>
        <button type="button" class="btn btn-sm btn-ghost" data-act="skip">Cancel</button>
      </div>`;
    const input = document.querySelector(`.task-input[data-type-id="${type.Id}"]`);
    el.querySelector('[data-act="open"]').addEventListener("click", () => {
      clearFeedback(type.Id);
      if(input) input.value = "";
      Work.openItem(existing.Id);
    });
    el.querySelector('[data-act="again"]').addEventListener("click", () => {
      clearFeedback(type.Id);
      startChecklist(type, input, true, value, typeof Roster !== "undefined" && !Roster.find(value));
    });
    el.querySelector('[data-act="skip"]').addEventListener("click", () => {
      clearFeedback(type.Id);
      if(input) input.focus();
    });
    el.querySelector('[data-act="open"]').focus();
  }

  /* Called by Work when an item completes or reopens, so counts stay current. */
  function ingest(entry){
    if(entry) recent.unshift(entry);
    renderCounts();
    renderMine();
  }

  function markVoided(id){
    const e = recent.find(r => r.Id === id);
    if(e){ e[F.voided] = true; e[F.voidedBy] = App.user.name; e[F.voidedOn] = new Date().toISOString(); }
    renderCounts();
    renderMine();
  }

  /* Returns Map(identifier -> most recent matching entry) within the type's window. */
  async function findDuplicates(type, ids){
    const found = new Map();
    const days = windowDays(type);
    if(days <= 0) return found;
    const since = addDays(new Date(), -days);

    const results = await Promise.all(ids.map(async idf => {
      try{
        // Identifier first: it's the most selective indexed column (keeps us under the 5,000-item threshold)
        const expr = `${F.identifier} eq ${odataString(idf)} and ${F.activityTypeId} eq ${type.Id} and ${F.loggedAt} ge ${odataDate(since)}`;
        const items = await spGetAll(CONFIG.lists.activityLog, filterQuery(expr, "$top=50"));
        return [idf, items.filter(i => !i[F.voided])];
      }catch(err){
        console.warn("Duplicate check used local data instead of SharePoint:", err);
        return [idf, recent.filter(i =>
          i[F.activityTypeId] === type.Id && i[F.identifier] === idf &&
          !i[F.voided] && new Date(i[F.loggedAt]) >= since)];
      }
    }));

    results.forEach(([idf, items]) => {
      if(!items.length) return;
      items.sort((a, b) => new Date(b[F.loggedAt]) - new Date(a[F.loggedAt]));
      found.set(idf, items[0]);
    });
    return found;
  }

  async function createEntries(type, ids){
    const created = [];
    const failed = [];
    const now = new Date().toISOString();
    const chunkSize = 5;
    for(let i = 0; i < ids.length; i += chunkSize){
      const chunk = ids.slice(i, i + chunkSize);
      await Promise.all(chunk.map(async idf => {
        try{
          const item = await spCreate(CONFIG.lists.activityLog, {
            [F.title]: idf || type.Title,
            [F.activityTypeId]: type.Id,
            [F.activityName]: type.Title,
            [F.category]: type[T.category] || "",
            [F.identifier]: idf,
            [F.staffName]: App.user.name,
            [F.staffEmail]: App.user.username,
            [F.loggedAt]: now,
            [F.voided]: false
          });
          created.push(item);
          recent.unshift(item);
        }catch(err){
          console.error("Activity entry failed to save:", err);
          failed.push({ idf, err });
        }
      }));
    }
    return { created, failed };
  }

  function announceLogged(type, created){
    flashRow(type.Id);
    renderCounts();
    renderMine();
    const ids = created.map(c => c[F.identifier]).filter(Boolean);
    let msg = `Logged "${type.Title}"`;
    if(ids.length === 1) msg += ` for ${ids[0]}`;
    else if(ids.length > 1) msg += ` for ${ids.length} entries`;
    toast(`${msg}.`, {
      type: "success",
      actionLabel: "Undo",
      onAction: () => voidEntries(created),
      duration: CONFIG.undoSeconds * 1000
    });
  }

  function showDuplicatePrompt(type, dupes){
    const el = feedbackEl(type.Id);
    if(!el) return;
    const days = windowDays(type);
    el.className = "task-feedback show warn";
    el.innerHTML = `
      <div class="dupe-text">
        ${dupes.length === 1 ? "This was" : "These were"} already logged for this activity in the last ${plural(days, "day")}:
        <ul>${dupes.map(d => `<li><strong>${escapeHtml(d.idf)}</strong> by ${
          isMine(d.prior) ? "you" : escapeHtml(d.prior[F.staffName] || "someone")} on ${formatDate(d.prior[F.loggedAt])}</li>`).join("")}</ul>
      </div>
      <div class="dupe-actions">
        <button type="button" class="btn btn-sm btn-navy" data-act="again">Log again anyway</button>
        <button type="button" class="btn btn-sm btn-ghost" data-act="skip">Skip</button>
      </div>`;
    el.querySelector('[data-act="skip"]').addEventListener("click", () => {
      clearFeedback(type.Id);
      const input = document.querySelector(`.task-input[data-type-id="${type.Id}"]`);
      if(input) input.focus();
    });
    const againBtn = el.querySelector('[data-act="again"]');
    againBtn.addEventListener("click", async () => {
      againBtn.disabled = true;
      setRowBusy(type.Id, true);
      try{
        const { created, failed } = await createEntries(type, dupes.map(d => d.idf));
        clearFeedback(type.Id);
        if(created.length){
          announceLogged(type, created);
          auditLogged(type, created, dupes);
        }
        if(failed.length) showFeedback(type.Id, `${plural(failed.length, "entry", "entries")} didn't save: ${failed[0].err.message}`, "error");
      }finally{
        setRowBusy(type.Id, false);
        const input = document.querySelector(`.task-input[data-type-id="${type.Id}"]`);
        if(input) input.focus();
      }
    });
    againBtn.focus();
  }

  /* ---------- audit entries ---------- */

  function idList(entries){
    const ids = entries.map(e => e[F.identifier]).filter(Boolean);
    return ids.length > 12 ? `${ids.slice(0, 12).join(", ")} and ${ids.length - 12} more` : ids.join(", ");
  }

  /* One audit row per submission (not per number), with every entry Id in Details. */
  function auditLogged(type, created, overriddenDupes){
    const ids = idList(created);
    let action = `Logged "${type.Title}"${ids ? ` for ${ids}` : ""}`;
    let details = `Activity Log entries: ${created.map(e => "#" + e.Id).join(", ")}`;
    if(overriddenDupes && overriddenDupes.length){
      action += " after a duplicate warning (logged again anyway)";
      details += `\nPreviously logged: ` + overriddenDupes.map(d =>
        `${d.idf} by ${d.prior[F.staffName] || "unknown"} on ${formatDate(d.prior[F.loggedAt])}`).join("; ");
    }
    audit(AUDIT_AREAS.activity, action, { recordId: created.length === 1 ? created[0].Id : undefined, details });
  }

  /* ---------- remove (void) ---------- */

  async function voidEntries(entries){
    const now = new Date().toISOString();
    let ok = 0;
    for(const e of entries){
      try{
        await spUpdate(CONFIG.lists.activityLog, e.Id, {
          [F.voided]: true,
          [F.voidedBy]: App.user.name,
          [F.voidedOn]: now
        });
        [e, recent.find(r => r.Id === e.Id)].forEach(x => {
          if(x){ x[F.voided] = true; x[F.voidedBy] = App.user.name; x[F.voidedOn] = now; }
        });
        ok++;
      }catch(err){
        console.error("Couldn't remove entry", e.Id, err);
      }
    }
    renderCounts();
    renderMine();
    const removed = entries.filter(e => e[F.voided]);
    if(removed.length){
      const names = [...new Set(removed.map(e => e[F.activityName]))].join(", ");
      const ids = idList(removed);
      const counted = removed.some(e => e[F.quantity]) ? ` (${removed.reduce((a, e) => a + qty(e), 0).toLocaleString()} items)` : "";
      audit(AUDIT_AREAS.activity, `Removed ${removed.length === 1 ? "entry" : `${removed.length} entries`}: "${names}"${ids ? ` for ${ids}` : ""}${counted}`, {
        recordId: removed.length === 1 ? removed[0].Id : undefined,
        details: `Activity Log entries: ${removed.map(e => "#" + e.Id).join(", ")}; originally logged ${removed.map(e => formatDate(e[F.loggedAt])).join(", ")}`
      });
    }
    if(ok === entries.length){
      toast(ok === 1 ? "Entry removed." : `${ok} entries removed.`, { type: "info" });
    }else{
      toast(`Removed ${ok} of ${entries.length}. The rest couldn't be removed; check your connection and try again from Your recent entries.`,
        { type: "error", duration: 0 });
    }
    return ok;
  }

  return {
    init, refresh, reloadTypes, voidEntries, renderCounts, ingest, markVoided,
    inputMode, inputLabel, windowDays, groupByCategory, sortedTypes, templateSteps, hasChecklist, qty,
    teamOf, get currentTeam(){ return currentTeam; },
    get types(){ return allTypes; }
  };
})();
