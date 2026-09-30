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

  function init(){
    if(!App.isAdmin) return;
    if(!wired) wire();
    ViewHooks.adminView = () => { renderTypes(); Tickets.renderRules(); renderTeam(); };
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
  }

  /* ---------- activity types ---------- */

  function renderTypes(){
    const wrap = document.getElementById("typesTable");
    const showInactive = document.getElementById("showInactiveTypes").checked;
    const types = Activity.types.filter(t => showInactive || t[T.active] !== false);
    if(!types.length){
      wrap.innerHTML = `<div class="empty-block"><p>No activities yet. Select <strong>Add activity</strong> to create the first one; it appears on everyone's Log activity tab right away.</p></div>`;
      return;
    }
    const groups = Activity.groupByCategory(types);
    wrap.innerHTML = `
      <div class="table-scroll"><table class="plain types">
        <thead><tr><th>Order</th><th>Activity</th><th>Asks for</th><th>Duplicate check</th><th>Status</th><th></th></tr></thead>
        <tbody>${groups.map(g => `
          <tr class="grp"><th colspan="6">${escapeHtml(g.name)}</th></tr>
          ${g.types.map(t => {
            const active = t[T.active] !== false;
            const mode = Activity.inputMode(t);
            const days = Activity.windowDays(t);
            return `<tr class="static${active ? "" : " inactive"}">
              <td class="num">${t[T.sortOrder] ?? ""}</td>
              <td><strong>${escapeHtml(t.Title || "")}</strong>${t[T.description] ? `<div class="muted small">${escapeHtml(t[T.description])}</div>` : ""}</td>
              <td>${mode === "click" ? "Nothing (one click)" : escapeHtml(Activity.inputLabel(t))}</td>
              <td>${mode === "click" ? "—" : days > 0 ? `Last ${plural(days, "day")}` : "Off"}</td>
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
    $("atInputType").value = t ? (INPUT_TYPES.includes(t[T.inputType]) ? t[T.inputType] : "Employee Number") : "Employee Number";
    $("atInputLabel").value = t ? t[T.inputLabel] || "" : "";
    $("atDupDays").value = t
      ? (t[T.duplicateWindowDays] ?? CONFIG.defaultDuplicateWindowDays)
      : CONFIG.defaultDuplicateWindowDays;
    $("atSortOrder").value = t ? (t[T.sortOrder] ?? "") : maxOrder + 10;
    $("atDescription").value = t ? t[T.description] || "" : "";
    $("typeSaveMsg").textContent = "";
    $("typeRenameHint").hidden = !t;

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
    document.getElementById("atInputLabelRow").hidden = click;
    document.getElementById("atDupRow").hidden = click;
    document.getElementById("atInputLabel").placeholder = mode === "Reference" ? "e.g. Pay period, Case #" : "Employee #";
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
    const dup = dupRaw === "" ? CONFIG.defaultDuplicateWindowDays : Number(dupRaw);
    if(!Number.isFinite(dup) || dup < 0){ msg.className = "save-msg err"; msg.textContent = "Duplicate check must be 0 or more days."; return; }

    const body = {
      Title: name,
      [T.category]: category || "General",
      [T.inputType]: inputType,
      [T.inputLabel]: inputType === "Click Only" ? "" : label,
      [T.duplicateWindowDays]: inputType === "Click Only" ? 0 : Math.round(dup),
      [T.sortOrder]: orderRaw === "" ? null : Number(orderRaw),
      [T.description]: $("atDescription").value.trim()
    };
    if(!editingId) body[T.active] = true;

    const btn = $("typeSaveBtn");
    btn.disabled = true;
    msg.className = "save-msg";
    msg.textContent = "Saving...";
    try{
      if(editingId) await spUpdate(CONFIG.lists.activityTypes, editingId, body);
      else await spCreate(CONFIG.lists.activityTypes, body);
      await Activity.reloadTypes();
      renderTypes();
      closeTypeModal();
      toast(editingId ? "Activity updated." : `"${name}" added to the board.`, { type: "success" });
    }catch(err){
      console.error(err);
      msg.className = "save-msg err";
      msg.textContent = `Not saved: ${err.message}`;
    }finally{
      btn.disabled = false;
    }
  }

  async function toggleType(id){
    const t = Activity.types.find(x => x.Id === id);
    if(!t) return;
    const turningOff = t[T.active] !== false;
    if(turningOff && !confirm(`Turn off "${t.Title}"? It disappears from the board. Past entries stay in reports, and you can turn it back on any time.`)) return;
    try{
      await spUpdate(CONFIG.lists.activityTypes, id, { [T.active]: !turningOff });
      await Activity.reloadTypes();
      renderTypes();
    }catch(err){
      console.error(err);
      toast(`Couldn't change "${t.Title}": ${err.message}`, { type: "error" });
    }
  }

  /* ---------- team access ---------- */

  async function reloadTeam(){
    App.team = await spGetAll(CONFIG.lists.team, `$select=Id,${M.title},${M.role}&$top=500`);
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
        <span class="role-chip${role === "Admin" ? " admin" : ""}">${escapeHtml(role)}</span>
        ${isMe ? "" : `
          <button type="button" class="link-btn" data-role="${m.Id}">${role === "Admin" ? "Make staff" : "Make admin"}</button>
          <button type="button" class="link-btn danger" data-remove="${m.Id}">Remove</button>`}
      </div>`;
    }).join("");
    wrap.querySelectorAll("[data-role]").forEach(b => b.addEventListener("click", () => toggleRole(Number(b.dataset.role))));
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
    if(!email || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)){ setStaffMsg("Enter a full email address.", "err"); return; }
    if(App.team.some(m => localPart(m[M.title]) === localPart(email))){ setStaffMsg("That person is already on the roster.", "err"); return; }
    const btn = document.getElementById("addStaffBtn");
    btn.disabled = true;
    setStaffMsg("Adding...");
    try{
      await spCreate(CONFIG.lists.team, { [M.title]: email, [M.role]: role });
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

  async function toggleRole(id){
    const m = App.team.find(x => x.Id === id);
    if(!m) return;
    const next = (m[M.role] || "Staff") === "Admin" ? "Staff" : "Admin";
    try{
      await spUpdate(CONFIG.lists.team, id, { [M.role]: next });
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
      await reloadTeam();
    }catch(err){
      console.error(err);
      toast(`Not removed: ${err.message}`, { type: "error" });
    }
  }

  return { init };
})();
