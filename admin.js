(window.PS_FILE_VERSIONS = window.PS_FILE_VERSIONS || {})["admin.js"] = "2026.10.02-2";
/* ============================================================
   admin.js — Admin tab (visible to Role = Admin only).
   Activities: add, edit, reorder, turn on/off. Activities are
   never deleted, so historical entries keep their meaning.
   ============================================================ */

const Admin = (() => {
  const T = TYPE_FIELDS;
  const M = TEAM_FIELDS;
  let wired = false;
  let editingId = null;
  let tplSteps = [];      // checklist template being edited in the activity modal

  function init(){
    if(!App.isAdmin) return;
    if(!wired) wire();
    onView("adminView", () => { renderTypes(); Tickets.renderRules(); renderTeam(); });
  }

  function wire(){
    wired = true;
    const on = (id, ev, fn) => document.getElementById(id).addEventListener(ev, fn);
    on("addTypeBtn", "click", () => openTypeModal(null));
    on("typeModalClose", "click", closeTypeModal);
    on("typeCancelBtn", "click", closeTypeModal);
    on("typeSaveBtn", "click", saveType);
    on("atInputType", "change", syncInputTypeFields);
    on("addStaffBtn", "click", addMember);
    on("newStaffEmail", "keydown", e => { if(e.key === "Enter"){ e.preventDefault(); addMember(); } });
    on("showInactiveTypes", "change", renderTypes);
    on("atAddStepBtn", "click", addTplStep);
    document.getElementById("atTeam").innerHTML =
      TEAMS.map(tm => `<option value="${escapeHtml(tm)}">${escapeHtml(tm)}</option>`).join("") +
      `<option value="${TEAM_BOTH}">Both teams</option>`;
    document.getElementById("newStaffTeam").innerHTML = `<option value="">No team</option>` +
      TEAMS.map(tm => `<option value="${escapeHtml(tm)}">${escapeHtml(tm)}</option>`).join("");
    on("atNewStep", "keydown", e => { if(e.key === "Enter"){ e.preventDefault(); addTplStep(); } });
  }

  /* ---------- activity types ---------- */

  function renderTypes(){
    const wrap = document.getElementById("typesTable");
    const showInactive = document.getElementById("showInactiveTypes").checked;
    const types = Activity.types.filter(t => showInactive || t[T.active] !== false);
    if(!types.length){
      wrap.innerHTML = `<div class="empty-block"><p>No activities yet. Select <strong>Add activity</strong> to create the first one; it appears on everyone's Dashboard right away.</p></div>`;
      return;
    }
    const groups = Activity.groupByCategory(types);
    wrap.innerHTML = `
      <div class="table-scroll"><table class="plain types">
        <thead><tr><th>Order</th><th>Activity</th><th>Team</th><th>Asks for</th><th>Checklist</th><th>Duplicate check</th><th>Status</th><th></th></tr></thead>
        <tbody>${groups.map(g => `
          <tr class="grp"><th colspan="8">${escapeHtml(g.name)}</th></tr>
          ${g.types.map(t => {
            const active = t[T.active] !== false;
            const mode = Activity.inputMode(t);
            const days = Activity.windowDays(t);
            return `<tr class="static${active ? "" : " inactive"}">
              <td class="num">${t[T.sortOrder] ?? ""}</td>
              <td><strong>${escapeHtml(t.Title || "")}</strong>${t[T.description] ? `<div class="muted small">${escapeHtml(t[T.description])}</div>` : ""}</td>
              <td><span class="team-chip">${escapeHtml(Activity.teamOf(t))}</span></td>
              <td>${mode === "click" ? "Nothing (one click)" : mode === "count" ? `A count (${escapeHtml(Activity.inputLabel(t))})` : escapeHtml(Activity.inputLabel(t))}</td>
              <td>${Activity.hasChecklist(t) ? plural(Activity.templateSteps(t).length, "step") : "—"}</td>
              <td>${Activity.hasChecklist(t) ? "Open items" : (mode === "click" || mode === "count") ? "—" : days > 0 ? `Last ${plural(days, "day")}` : "Off"}</td>
              <td>${active ? "Active" : "Off"}</td>
              <td class="actions">
                <button type="button" class="link-btn" data-edit="${t.Id}">Edit</button>
                <button type="button" class="link-btn" data-toggle="${t.Id}">${active ? "Turn off" : "Turn on"}</button>
              </td>
            </tr>`;
          }).join("")}`).join("")}
        </tbody>
      </table></div>`;
    wrap.querySelectorAll("[data-edit]").forEach(b => b.addEventListener("click", () => openTypeModal(Number(b.dataset.edit))));
    wrap.querySelectorAll("[data-toggle]").forEach(b => b.addEventListener("click", () => toggleType(Number(b.dataset.toggle))));
  }

  function openTypeModal(id){
    editingId = id;
    const t = id ? Activity.types.find(x => x.Id === id) : null;
    const $ = x => document.getElementById(x);
    const maxOrder = Math.max(0, ...Activity.types.map(x => Number(x[T.sortOrder]) || 0));
    $("typeModalTitle").textContent = t ? "Edit activity" : "Add activity";
    $("atName").value = t ? t.Title || "" : "";
    $("atCategory").value = t ? t[T.category] || "" : "";
    $("atTeam").value = t ? Activity.teamOf(t) : (Activity.currentTeam || TEAMS[0]);
    $("atInputType").value = t ? (INPUT_TYPES.includes(t[T.inputType]) ? t[T.inputType] : "Employee Number") : "Employee Number";
    $("atInputLabel").value = t ? t[T.inputLabel] || "" : "";
    $("atDupDays").value = t
      ? (t[T.duplicateWindowDays] ?? CONFIG.defaultDuplicateWindowDays)
      : CONFIG.defaultDuplicateWindowDays;
    $("atSortOrder").value = t ? (t[T.sortOrder] ?? "") : maxOrder + 10;
    $("atDescription").value = t ? t[T.description] || "" : "";
    $("typeSaveMsg").textContent = "";
    $("typeRenameHint").hidden = !t;
    tplSteps = t ? Activity.templateSteps(t).map(x => ({ ...x })) : [];
    $("atNewStep").value = "";
    $("atNewStepExt").checked = false;
    $("atStepsHint").hidden = !(t && tplSteps.length);
    renderTplSteps();

    const cats = [...new Set(Activity.types.map(x => (x[T.category] || "").trim()).filter(Boolean))].sort();
    $("categoryOptions").innerHTML = cats.map(c => `<option value="${escapeHtml(c)}"></option>`).join("");

    syncInputTypeFields();
    openOverlay("typeModalOverlay");
    $("atName").focus();
  }

  function closeTypeModal(){
    closeOverlay("typeModalOverlay");
    editingId = null;
  }

  function syncInputTypeFields(){
    const mode = document.getElementById("atInputType").value;
    const click = mode === "Click Only";
    const count = mode === "Count";
    document.getElementById("atInputLabelRow").hidden = click;
    // A count has no single item to check for duplicates, and is logged in one step (no checklist).
    document.getElementById("atDupRow").hidden = click || count || tplSteps.length > 0;
    document.getElementById("atChecklistRow").hidden = count;
    document.getElementById("atInputLabel").placeholder =
      mode === "Reference" ? "e.g. Pay period, Case #" : count ? "How many?" : "Employee #";
  }

  async function saveType(){
    const $ = x => document.getElementById(x);
    const name = $("atName").value.trim();
    const category = $("atCategory").value.trim();
    const inputType = $("atInputType").value;
    const label = $("atInputLabel").value.trim();
    const dupRaw = $("atDupDays").value.trim();
    const orderRaw = $("atSortOrder").value.trim();
    const msg = $("typeSaveMsg");

    if(!name){ msg.className = "save-msg err"; msg.textContent = "Give the activity a name."; $("atName").focus(); return; }
    const clash = Activity.types.find(x =>
      x.Id !== editingId && x[T.active] !== false && String(x.Title || "").trim().toLowerCase() === name.toLowerCase());
    if(clash){ msg.className = "save-msg err"; msg.textContent = "An active activity already has that name."; return; }
    tplSteps = tplSteps.map(x => ({ text: String(x.text || "").trim(), external: !!x.external })).filter(x => x.text);
    const pending = $("atNewStep").value.trim();
    if(pending){ tplSteps.push({ text: pending, external: $("atNewStepExt").checked }); $("atNewStep").value = ""; }
    const dup = dupRaw === "" ? CONFIG.defaultDuplicateWindowDays : Number(dupRaw);
    if(!Number.isFinite(dup) || dup < 0){ msg.className = "save-msg err"; msg.textContent = "Duplicate check must be 0 or more days."; return; }

    const body = {
      Title: name,
      [T.category]: category || "General",
      [T.team]: $("atTeam").value,
      [T.inputType]: inputType,
      [T.inputLabel]: inputType === "Click Only" ? "" : label,
      [T.duplicateWindowDays]: (inputType === "Click Only" || inputType === "Count") ? 0 : Math.round(dup),
      [T.sortOrder]: orderRaw === "" ? null : Number(orderRaw),
      [T.description]: $("atDescription").value.trim(),
      [T.checklistSteps]: tplSteps.length && inputType !== "Count"
        ? JSON.stringify(tplSteps.map(x => ({ text: x.text, external: !!x.external })))
        : ""
    };
    if(!editingId) body[T.active] = true;

    const btn = $("typeSaveBtn");
    btn.disabled = true;
    msg.className = "save-msg";
    msg.textContent = "Saving...";
    const wasEditing = editingId;
    const before = wasEditing ? Activity.types.find(x => x.Id === wasEditing) : null;
    try{
      let savedId = wasEditing;
      if(wasEditing) await spUpdate(CONFIG.lists.activityTypes, wasEditing, body);
      else savedId = (await spCreate(CONFIG.lists.activityTypes, body)).Id;
      await Activity.reloadTypes();
      renderTypes();
      closeTypeModal();
      auditTypeSave(before, body, savedId);
      toast(wasEditing ? "Activity updated." : `"${name}" added to the board.`, { type: "success" });
    }catch(err){
      console.error(err);
      msg.className = "save-msg err";
      msg.textContent = `Not saved: ${err.message}`;
    }finally{
      btn.disabled = false;
    }
  }

  /* ---------- checklist template editor ---------- */

  function renderTplSteps(){
    const ol = document.getElementById("atSteps");
    ol.innerHTML = tplSteps.map((st, i) => `
      <li class="tpl-step" data-i="${i}">
        <input type="text" class="tpl-text" value="${escapeHtml(st.text)}" aria-label="Step ${i + 1}">
        <label class="check small" title="Done by another office (e.g. Civil Service, HR)"><input type="checkbox" class="tpl-ext" ${st.external ? "checked" : ""}> Other office</label>
        <button type="button" class="icon-btn" data-act="up" ${i === 0 ? "disabled" : ""} aria-label="Move up">↑</button>
        <button type="button" class="icon-btn" data-act="down" ${i === tplSteps.length - 1 ? "disabled" : ""} aria-label="Move down">↓</button>
        <button type="button" class="icon-btn danger" data-act="del" aria-label="Remove step">&times;</button>
      </li>`).join("");
    ol.querySelectorAll(".tpl-step").forEach(li => {
      const i = Number(li.dataset.i);
      li.querySelector(".tpl-text").addEventListener("input", e => { tplSteps[i].text = e.target.value; });
      li.querySelector(".tpl-ext").addEventListener("change", e => { tplSteps[i].external = e.target.checked; });
      li.querySelectorAll("[data-act]").forEach(b => b.addEventListener("click", () => {
        const act = b.dataset.act;
        if(act === "del") tplSteps.splice(i, 1);
        if(act === "up" && i > 0) [tplSteps[i - 1], tplSteps[i]] = [tplSteps[i], tplSteps[i - 1]];
        if(act === "down" && i < tplSteps.length - 1) [tplSteps[i + 1], tplSteps[i]] = [tplSteps[i], tplSteps[i + 1]];
        renderTplSteps();
        syncInputTypeFields();
      }));
    });
  }

  function addTplStep(){
    const input = document.getElementById("atNewStep");
    const text = input.value.trim();
    if(!text){ input.focus(); return; }
    tplSteps.push({ text, external: document.getElementById("atNewStepExt").checked });
    input.value = "";
    document.getElementById("atNewStepExt").checked = false;
    renderTplSteps();
    syncInputTypeFields();
    input.focus();
  }

  /* Audit an activity create/edit, listing exactly what changed. */
  function auditTypeSave(before, body, id){
    const describeSteps = raw => {
      const st = Activity.templateSteps({ [T.checklistSteps]: raw });
      return st.length ? st.map((x, i) => `${i + 1}. ${x.text}${x.external ? " [other office]" : ""}`).join("\n") : "(none)";
    };
    const labels = {
      Title: "name", [T.category]: "category", [T.team]: "team", [T.inputType]: "staff enter", [T.inputLabel]: "box label",
      [T.duplicateWindowDays]: "duplicate check days", [T.sortOrder]: "order", [T.description]: "helper text"
    };
    if(!before){
      const steps = Activity.templateSteps(body).length;
      audit(AUDIT_AREAS.admin, `Added activity "${body.Title}" (${body[T.team]}, ${body[T.category]}, ${body[T.inputType]}${steps ? `, checklist with ${plural(steps, "step")}` : ""})`, {
        recordId: id, details: steps ? `Checklist:\n${describeSteps(body[T.checklistSteps])}` : ""
      });
      return;
    }
    const norm = v => (v === null || v === undefined) ? "" : String(v);
    const changes = Object.keys(labels)
      .filter(k => norm(before[k]) !== norm(body[k]))
      .map(k => `${labels[k]}: "${norm(before[k])}" → "${norm(body[k])}"`);
    const stepsChanged = norm(before[T.checklistSteps]) !== norm(body[T.checklistSteps]) &&
      describeSteps(before[T.checklistSteps]) !== describeSteps(body[T.checklistSteps]);
    if(stepsChanged) changes.push("checklist steps");
    if(!changes.length) return;
    audit(AUDIT_AREAS.admin, `Edited activity "${body.Title}": ${changes.join("; ")}`, {
      recordId: id,
      details: stepsChanged ? `Checklist before:\n${describeSteps(before[T.checklistSteps])}\n\nChecklist after:\n${describeSteps(body[T.checklistSteps])}` : ""
    });
  }

  async function toggleType(id){
    const t = Activity.types.find(x => x.Id === id);
    if(!t) return;
    const turningOff = t[T.active] !== false;
    if(turningOff && !confirm(`Turn off "${t.Title}"? It disappears from the board. Past entries stay in reports, and you can turn it back on any time.`)) return;
    try{
      await spUpdate(CONFIG.lists.activityTypes, id, { [T.active]: !turningOff });
      audit(AUDIT_AREAS.admin, `${turningOff ? "Turned off" : "Turned on"} activity "${t.Title}"`, { recordId: id });
      await Activity.reloadTypes();
      renderTypes();
    }catch(err){
      console.error(err);
      toast(`Couldn't change "${t.Title}": ${err.message}`, { type: "error" });
    }
  }

  /* ---------- team access ---------- */

  async function reloadTeam(){
    App.team = await spGetAll(CONFIG.lists.team, "$top=500");
    renderTeam();
  }

  function renderTeam(){
    const wrap = document.getElementById("staffList");
    const team = App.team.slice().sort((a, b) =>
      String(a[M.title] || "").localeCompare(String(b[M.title] || "")));
    if(!team.length){ wrap.innerHTML = `<p class="muted small">No one on the roster yet.</p>`; return; }
    wrap.innerHTML = team.map(m => {
      const isMe = localPart(m[M.title]) === App.user.key;
      const role = m[M.role] || "Staff";
      return `<div class="list-item">
        <span class="grow">${escapeHtml(m[M.title] || "")}${isMe ? ' <span class="muted small">(you)</span>' : ""}</span>
        <select class="member-team" data-team-for="${m.Id}" aria-label="Team for ${escapeHtml(m[M.title] || "")}">
          <option value="">No team</option>
          ${TEAMS.map(tm => `<option value="${escapeHtml(tm)}"${m[M.team] === tm ? " selected" : ""}>${escapeHtml(tm)}</option>`).join("")}
        </select>
        <span class="role-chip${role === "Admin" ? " admin" : ""}">${escapeHtml(role)}</span>
        ${isMe ? "" : `
          <button type="button" class="link-btn" data-role="${m.Id}">${role === "Admin" ? "Make staff" : "Make admin"}</button>
          <button type="button" class="link-btn danger" data-remove="${m.Id}">Remove</button>`}
      </div>`;
    }).join("");
    wrap.querySelectorAll("[data-role]").forEach(b => b.addEventListener("click", () => toggleRole(Number(b.dataset.role))));
    wrap.querySelectorAll("[data-team-for]").forEach(sel => sel.addEventListener("change", () => setMemberTeam(Number(sel.dataset.teamFor), sel.value, sel)));
    wrap.querySelectorAll("[data-remove]").forEach(b => b.addEventListener("click", () => removeMember(Number(b.dataset.remove))));
  }

  function setStaffMsg(text, kind){
    const el = document.getElementById("staffSaveMsg");
    el.className = `save-msg ${kind || ""}`;
    el.textContent = text;
  }

  async function addMember(){
    const input = document.getElementById("newStaffEmail");
    const email = input.value.trim().toLowerCase();
    const role = document.getElementById("newStaffRole").value;
    const team = document.getElementById("newStaffTeam").value;
    if(!email || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)){ setStaffMsg("Enter a full email address.", "err"); return; }
    if(App.team.some(m => localPart(m[M.title]) === localPart(email))){ setStaffMsg("That person is already on the roster.", "err"); return; }
    const btn = document.getElementById("addStaffBtn");
    btn.disabled = true;
    setStaffMsg("Adding...");
    try{
      const created = await spCreate(CONFIG.lists.team, { [M.title]: email, [M.role]: role, ...(team ? { [M.team]: team } : {}) });
      audit(AUDIT_AREAS.admin, `Gave portal access to ${email} as ${role}${team ? `, team ${team}` : ""}`, { recordId: created.Id });
      input.value = "";
      setStaffMsg(`Added ${email}.`, "ok");
      await reloadTeam();
    }catch(err){
      console.error(err);
      setStaffMsg(`Not added: ${err.message}`, "err");
    }finally{
      btn.disabled = false;
    }
  }

  async function setMemberTeam(id, team, sel){
    const m = App.team.find(x => x.Id === id);
    if(!m) return;
    const before = m[M.team] || "";
    sel.disabled = true;
    try{
      await spUpdate(CONFIG.lists.team, id, { [M.team]: team || null });
      m[M.team] = team || null;
      audit(AUDIT_AREAS.admin, `Set ${m[M.title]}'s team: ${before || "none"} → ${team || "none"}`, { recordId: id });
      toast(`${m[M.title]} is now ${team ? `on ${team}` : "not assigned to a team"}.`, { type: "success" });
    }catch(err){
      console.error(err);
      sel.value = before;
      toast(`Team wasn't changed: ${err.message}. If the Team Members list has no Team column yet, run the setup check.`, { type: "error" });
    }finally{
      sel.disabled = false;
    }
  }

  async function toggleRole(id){
    const m = App.team.find(x => x.Id === id);
    if(!m) return;
    const next = (m[M.role] || "Staff") === "Admin" ? "Staff" : "Admin";
    try{
      await spUpdate(CONFIG.lists.team, id, { [M.role]: next });
      audit(AUDIT_AREAS.admin, `Changed ${m[M.title]} from ${m[M.role] || "Staff"} to ${next}`, { recordId: id });
      await reloadTeam();
    }catch(err){
      console.error(err);
      toast(`Role wasn't changed: ${err.message}`, { type: "error" });
    }
  }

  async function removeMember(id){
    const m = App.team.find(x => x.Id === id);
    if(!confirm(`Remove ${m ? m[M.title] : "this person"} from the roster? They lose access right away. Their past entries stay in reports.`)) return;
    try{
      await spDelete(CONFIG.lists.team, id);
      audit(AUDIT_AREAS.admin, `Removed portal access for ${m ? m[M.title] : `roster entry #${id}`}`, { recordId: id });
      await reloadTeam();
    }catch(err){
      console.error(err);
      toast(`Not removed: ${err.message}`, { type: "error" });
    }
  }

  return { init };
})();
