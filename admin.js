(window.PS_FILE_VERSIONS = window.PS_FILE_VERSIONS || {})["admin.js"] = "2026.10.05-2";
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
  let caseDef = null;     // 2.0: employee-process definition being edited (null = not an employee process)

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
    on("atEmpProc", "change", () => {
      if(document.getElementById("atEmpProc").checked && !caseDef){
        caseDef = { employee: true, processKey: "", fields: [
          { key: "received", label: "Date request received", type: "date", col: "received" },
          { key: "supervisor", label: "Supervisor", type: "text", col: "supervisor" }
        ] };
      }
      renderCaseFields(); renderTplSteps(); syncInputTypeFields();
    });
    on("atAddFieldBtn", "click", addCaseField);
    on("atNewFieldLabel", "keydown", e => { if(e.key === "Enter"){ e.preventDefault(); addCaseField(); } });
    document.getElementById("atNewFieldType").innerHTML =
      Object.entries(CASE_FIELD_TYPES).map(([k, v]) => `<option value="${k}">${escapeHtml(v)}</option>`).join("");
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
    const d = t ? Cases.def(t) : null;
    caseDef = d ? JSON.parse(JSON.stringify(d)) : null;
    $("atEmpProc").checked = !!caseDef;
    $("atNewFieldLabel").value = "";
    renderCaseFields();
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
    const emp = !!(document.getElementById("atEmpProc") && document.getElementById("atEmpProc").checked);
    document.getElementById("atDupRow").hidden = emp || click || count || tplSteps.length > 0;
    document.getElementById("atInputTypeRow").hidden = emp;
    document.getElementById("atInputLabelRow").hidden = click || emp;
    document.getElementById("atEmpProcRow").hidden = count;
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
    const emp = empProcOn();
    const fieldKeys = new Set(emp ? caseDef.fields.map(f => f.key) : []);
    tplSteps = tplSteps.map(x => ({ text: String(x.text || "").trim(), external: !!x.external,
      ...(emp && x.field && fieldKeys.has(x.field) ? { field: x.field } : {}) })).filter(x => x.text);
    if(emp){
      caseDef.fields = caseDef.fields.map(f => ({ ...f, label: String(f.label || "").trim() })).filter(f => f.label);
      if(!caseDef.processKey) caseDef.processKey = slugKey(name, new Set(Activity.types.map(t => (Cases.def(t) || {}).processKey)));
      const cols = caseDef.fields.filter(f => f.col).map(f => f.col);
      if(new Set(cols).size !== cols.length){ msg.className = "save-msg err"; msg.textContent = "Two fields are stored in the same column. Change one of them."; return; }
    }
    const pending = $("atNewStep").value.trim();
    if(pending){ tplSteps.push({ text: pending, external: $("atNewStepExt").checked }); $("atNewStep").value = ""; }
    const dup = dupRaw === "" ? CONFIG.defaultDuplicateWindowDays : Number(dupRaw);
    if(!Number.isFinite(dup) || dup < 0){ msg.className = "save-msg err"; msg.textContent = "Duplicate check must be 0 or more days."; return; }

    const body = {
      Title: name,
      [T.category]: category || "General",
      [T.team]: $("atTeam").value,
      [T.inputType]: emp ? "Employee Number" : inputType,
      [T.inputLabel]: emp ? "Employee #" : inputType === "Click Only" ? "" : label,
      [T.duplicateWindowDays]: (emp || inputType === "Click Only" || inputType === "Count") ? 0 : Math.round(dup),
      [T.caseFields]: emp ? JSON.stringify(caseDef) : "",
      [T.sortOrder]: orderRaw === "" ? null : Number(orderRaw),
      [T.description]: $("atDescription").value.trim(),
      [T.checklistSteps]: tplSteps.length && (emp || inputType !== "Count")
        ? JSON.stringify(tplSteps.map(x => ({ text: x.text, external: !!x.external, ...(x.field ? { field: x.field } : {}) })))
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

  /* ---------- 2.0: employee-process field editor ---------- */

  function empProcOn(){ return !!(document.getElementById("atEmpProc").checked && caseDef); }

  function slugKey(label, taken){
    const words = String(label).toLowerCase().replace(/[^a-z0-9 ]+/g, " ").trim().split(/\s+/).filter(Boolean).slice(0, 5);
    let base = words.map((w, i) => i ? w[0].toUpperCase() + w.slice(1) : w).join("") || "field";
    if(/^\d/.test(base)) base = "f" + base;
    let k = base, n = 2;
    while(taken.has(k)) k = base + n++;
    return k;
  }

  /* Which real columns a field type may be stored in. */
  function colChoices(type){
    return Object.entries(CASE_COLUMNS).filter(([, c]) =>
      c.type === "date" ? type === "date" : (type === "text" || type === "choice")).map(([k, c]) => [k, c.label]);
  }

  function renderCaseFields(){
    const row = document.getElementById("atFieldsRow");
    const on = empProcOn();
    row.hidden = !on;
    if(!on) return;
    const list = document.getElementById("atFields");
    list.innerHTML = caseDef.fields.map((f, i) => {
      const cols = colChoices(f.type);
      return `<li class="field-item" data-i="${i}">
        <input type="text" class="fd-label" value="${escapeHtml(f.label)}" aria-label="Field name">
        <select class="fd-type" aria-label="Field type">${Object.entries(CASE_FIELD_TYPES).map(([k, v]) => `<option value="${k}"${f.type === k ? " selected" : ""}>${escapeHtml(v)}</option>`).join("")}</select>
        <select class="fd-col" aria-label="Stored in" title="Common fields can be stored in their own SharePoint column so they can be filtered there">
          <option value="">Stored with the case</option>
          ${cols.map(([k, l]) => `<option value="${k}"${f.col === k ? " selected" : ""}>Column: ${escapeHtml(l)}</option>`).join("")}
        </select>
        <input type="text" class="fd-options" value="${escapeHtml((f.options || []).join(", "))}" placeholder="Choices, separated by commas" ${f.type === "choice" ? "" : "hidden"}>
        <button type="button" class="icon-btn" data-act="up" ${i === 0 ? "disabled" : ""} aria-label="Move up">↑</button>
        <button type="button" class="icon-btn" data-act="down" ${i === caseDef.fields.length - 1 ? "disabled" : ""} aria-label="Move down">↓</button>
        <button type="button" class="icon-btn danger" data-act="del" aria-label="Remove field">&times;</button>
      </li>`;
    }).join("");
    list.querySelectorAll(".field-item").forEach(li => {
      const i = Number(li.dataset.i);
      const f = caseDef.fields[i];
      li.querySelector(".fd-label").addEventListener("input", e => { f.label = e.target.value; });
      li.querySelector(".fd-type").addEventListener("change", e => {
        f.type = e.target.value;
        if(f.col && !colChoices(f.type).some(([k]) => k === f.col)) delete f.col;
        renderCaseFields(); renderTplSteps();
      });
      li.querySelector(".fd-col").addEventListener("change", e => { if(e.target.value) f.col = e.target.value; else delete f.col; });
      li.querySelector(".fd-options").addEventListener("input", e => {
        f.options = e.target.value.split(",").map(x => x.trim()).filter(Boolean);
      });
      li.querySelectorAll("[data-act]").forEach(b => b.addEventListener("click", () => {
        const act = b.dataset.act;
        if(act === "del"){
          if(!confirm(`Remove the field "${f.label}"? Values already saved on cases stay in SharePoint but won't show.`)) return;
          caseDef.fields.splice(i, 1);
          tplSteps.forEach(s => { if(s.field === f.key) delete s.field; });
        }
        if(act === "up" && i > 0) [caseDef.fields[i - 1], caseDef.fields[i]] = [caseDef.fields[i], caseDef.fields[i - 1]];
        if(act === "down" && i < caseDef.fields.length - 1) [caseDef.fields[i + 1], caseDef.fields[i]] = [caseDef.fields[i], caseDef.fields[i + 1]];
        renderCaseFields(); renderTplSteps();
      }));
    });
  }

  function addCaseField(){
    const input = document.getElementById("atNewFieldLabel");
    const label = input.value.trim();
    if(!label){ input.focus(); return; }
    const type = document.getElementById("atNewFieldType").value;
    caseDef.fields.push({ key: slugKey(label, new Set(caseDef.fields.map(f => f.key))), label, type });
    input.value = "";
    renderCaseFields(); renderTplSteps();
    input.focus();
  }

  /* ---------- checklist template editor ---------- */

  function renderTplSteps(){
    const ol = document.getElementById("atSteps");
    ol.innerHTML = tplSteps.map((st, i) => `
      <li class="tpl-step" data-i="${i}">
        <input type="text" class="tpl-text" value="${escapeHtml(st.text)}" aria-label="Step ${i + 1}">
        <label class="check small" title="Done by another office (e.g. Civil Service, HR)"><input type="checkbox" class="tpl-ext" ${st.external ? "checked" : ""}> Other office</label>
        ${empProcOn() ? `<select class="tpl-field" aria-label="Field this step fills in" title="Checking this step fills in the field on the employee's case">
          <option value="">No linked field</option>
          ${caseDef.fields.filter(f => Cases.linkable(f)).map(f => `<option value="${escapeHtml(f.key)}"${st.field === f.key ? " selected" : ""}>Fills: ${escapeHtml(f.label)}</option>`).join("")}
        </select>` : ""}
        <button type="button" class="icon-btn" data-act="up" ${i === 0 ? "disabled" : ""} aria-label="Move up">↑</button>
        <button type="button" class="icon-btn" data-act="down" ${i === tplSteps.length - 1 ? "disabled" : ""} aria-label="Move down">↓</button>
        <button type="button" class="icon-btn danger" data-act="del" aria-label="Remove step">&times;</button>
      </li>`).join("");
    ol.querySelectorAll(".tpl-field").forEach(sel => sel.addEventListener("change", () => {
      const i = Number(sel.closest(".tpl-step").dataset.i);
      if(sel.value) tplSteps[i].field = sel.value; else delete tplSteps[i].field;
    }));
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
    if(norm(before[T.caseFields]) !== norm(body[T.caseFields])) changes.push(body[T.caseFields] ? "employee-process fields" : "no longer an employee process");
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
