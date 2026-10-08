(window.PS_FILE_VERSIONS = window.PS_FILE_VERSIONS || {})["setup-check.js"] = "2026.10.08-1";
/* ============================================================
   setup-check.js — Admin → "Check SharePoint setup".
   Read-only. Compares every list and column the app expects
   (from config.js) with what's actually on the SharePoint site,
   and explains each problem in plain language.
   ============================================================ */

const SetupCheck = (() => {
  const TYPE_LABEL = {
    Text: "Single line of text",
    Note: "Multiple lines of text",
    Number: "Number",
    Boolean: "Yes/No",
    DateTime: "Date and time",
    Choice: "Choice"
  };

  /* Expected schema, built from the field maps in config.js so it can't drift from the code.
     Each field: [internalName, type, options]
       options.choices  — values the app writes or filters on
       options.json     — holds JSON: must be plain text and not append-only
       options.index    — should be indexed (needed once the list passes 5,000 items) */
  function schema(){
    const T = TYPE_FIELDS, F = LOG_FIELDS, R = REQ_FIELDS, A = AUDIT_FIELDS,
          AL = ALERT_FIELDS, M = TEAM_FIELDS, W = WORK_FIELDS, WL = WORKLOG_FIELDS, E = EMP_FIELDS;
    return [
      { list: CONFIG.lists.team, purpose: "who can sign in", fields: [
        [M.role, "Choice", { choices: ["Staff", "Admin"] }], [M.team, "Choice", { choices: TEAMS }]
      ]},
      { list: CONFIG.lists.activityTypes, purpose: "the activities on the Dashboard", fields: [
        [T.category, "Text"], [T.inputType, "Choice", { choices: INPUT_TYPES }], [T.inputLabel, "Text"],
        [T.duplicateWindowDays, "Number"], [T.sortOrder, "Number"], [T.description, "Text"],
        [T.active, "Boolean"], [T.checklistSteps, "Note", { json: true }],
        [T.team, "Choice", { choices: [...TEAMS, TEAM_BOTH] }], [T.caseFields, "Note", { json: true }]
      ]},
      { list: CONFIG.lists.activityLog, purpose: "every logged activity", fields: [
        [F.activityTypeId, "Number", { index: true }], [F.activityName, "Text"], [F.category, "Text"],
        [F.identifier, "Text", { index: true }], [F.staffName, "Text"], [F.staffEmail, "Text", { index: true }],
        [F.loggedAt, "DateTime", { index: true, time: true }], [F.quantity, "Number"],
        [F.voided, "Boolean"], [F.voidedBy, "Text"],
        [F.voidedOn, "DateTime", { time: true }]
      ]},
      { list: CONFIG.lists.requests, purpose: "inquiries", fields: [
        [R.email, "Text"], [R.problem, "Note", { plain: true }],
        [R.internalNotes, "Note", { plain: true }], [R.completed, "Boolean"],
        [R.completedOn, "DateTime", { index: true, time: true }], [R.completedBy, "Text"],
        [R.status, "Choice", { index: true, choices: ["New", "In Progress", "Completed", "Merged", NO_ACTION_STATUS] }],
        [R.receivedOn, "DateTime", { index: true, time: true }],
        [R.entryType, "Choice", { choices: ["Automated", "Manual"] }],
        [R.source, "Choice", { choices: ["Email", "Phone", "Walk-in", "ServiceNow"] }], [R.serviceNow, "Text"],
        [R.requesterName, "Text"], [R.phoneNumber, "Text"], [R.queue, "Text"]
      ]},
      { list: CONFIG.lists.audit, purpose: "the audit log", renamedFrom: "Request Audit Log", fields: [
        [A.area, "Text"], [A.recordId, "Number"], [A.details, "Note", { plain: true }], [A.staffEmail, "Text"],
        [A.ticketId, "Number"], [A.staffMember, "Text"], [A.logTime, "DateTime", { index: true, time: true }],
        [A.action, "Note", { plain: true }], [A.previousStatus, "Text"], [A.newStatus, "Text"],
        [A.internalNotesSnapshot, "Note", { plain: true }],
        [A.problemSnapshot, "Note", { plain: true }]
      ]},
      { list: CONFIG.lists.alertRules, purpose: "inquiry aging alerts", fields: [
        [AL.color, "Text"], [AL.thresholdValue, "Number"], [AL.thresholdUnit, "Choice", { choices: ["Hours", "Days"] }],
        [AL.excludeWeekends, "Boolean"], [AL.active, "Boolean"]
      ]},
      { list: CONFIG.lists.workItems, purpose: "tracked work (checklists)", fields: [
        [W.activityTypeId, "Number", { index: true }], [W.activityName, "Text"], [W.category, "Text"], [W.identifier, "Text"],
        [W.status, "Choice", { index: true, choices: ["Open", "Completed", "Cancelled"] }],
        [W.ownerName, "Text"], [W.ownerEmail, "Text"], [W.startedBy, "Text"],
        [W.startedOn, "DateTime", { time: true }], [W.lastActivityOn, "DateTime", { time: true }],
        [W.closedOn, "DateTime", { index: true, time: true }], [W.closedBy, "Text"],
        [W.steps, "Note", { json: true }], [W.completionLogId, "Number"],
        // 2.0 employee processes
        [W.employeeId, "Text", { index: true }], [W.employeeName, "Text"], [W.supervisor, "Text"],
        [W.receivedDate, "DateTime", { time: true }], [W.startDate, "DateTime", { time: true }], [W.endDate, "DateTime", { time: true }],
        [W.outcome, "Text"], [W.caseData, "Note", { json: true }], [W.legacyKey, "Text", { index: true }]
      ]},
      { list: CONFIG.lists.employees, purpose: "2.0: the employee roster", fields: [
        [E.employeeId, "Text", { index: true }], [E.firstName, "Text"], [E.lastName, "Text"], [E.supervisor, "Text"],
        [E.badge, "Text"], [E.rank, "Text"], [E.division, "Text"], [E.hireDate, "DateTime"],
        [E.active, "Boolean"], [E.source, "Text"],
        [E.lastRosterDate, "DateTime", { time: true }]
      ]},
      { list: CONFIG.lists.workItemLog, purpose: "tracked work history and notes", fields: [
        [WL.workItemId, "Number", { index: true }], [WL.action, "Note", { plain: true }],
        [WL.staffName, "Text"], [WL.staffEmail, "Text"], [WL.loggedAt, "DateTime", { time: true }]
      ]}
    ];
  }

  function listUrl(title, suffix){
    return `${CONFIG.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(title.replace(/'/g, "''"))}')${suffix}`;
  }

  async function listExists(title){
    try{ await spRequest(listUrl(title, "?$select=Id")); return true; }
    catch(_){ return false; }
  }

  async function getFields(title){
    const res = await spRequest(listUrl(title, `/fields?$filter=${encodeURIComponent("Hidden eq false")}`));
    return (await res.json()).value || [];
  }

  const squash = s => String(s || "").toLowerCase().replace(/[^a-z0-9]/g, "");

  async function checkList(spec){
    const out = { list: spec.list, purpose: spec.purpose, errors: [], warnings: [], ok: 0, missingList: false };
    let fields;
    try{
      fields = await getFields(spec.list);
    }catch(err){
      out.missingList = true;
      if(err.status === 404){
        if(spec.renamedFrom && await listExists(spec.renamedFrom)){
          out.errors.push(`List not found. A list named "${spec.renamedFrom}" exists: rename it to "${spec.list}" (List settings → List name, description and navigation).`);
        }else{
          out.errors.push(`List not found. Create a list named exactly "${spec.list}".`);
        }
      }else{
        out.errors.push(`Couldn't read this list: ${err.message}`);
      }
      return out;
    }

    const byInternal = new Map(fields.map(f => [f.InternalName, f]));
    for(const [name, type, opt = {}] of spec.fields){
      const f = byInternal.get(name);
      if(!f){
        // Created with a space or a different first name? Look for a near match by display name.
        const near = fields.find(x => squash(x.Title) === squash(name) || squash(x.InternalName) === squash(name));
        if(near){
          out.errors.push(`"${name}": there's a column shown as "${near.Title}", but its internal name is "${near.InternalName}". ` +
            `The app needs the internal name "${name}". Delete that column and add it again, typing ${name} (no spaces) as the name.`);
        }else{
          out.errors.push(`"${name}" is missing. Add it as ${TYPE_LABEL[type]}${opt.json || opt.plain ? " (plain text)" : ""}${opt.choices ? ` with choices: ${opt.choices.join(", ")}` : ""}.`);
        }
        continue;
      }

      let fieldOk = true;
      if(f.TypeAsString !== type){
        fieldOk = false;
        const actual = TYPE_LABEL[f.TypeAsString] || f.TypeAsString;
        if(type === "Choice" && f.TypeAsString === "MultiChoice"){
          out.errors.push(`"${name}" allows multiple selections. Edit the column and turn off "Allow multiple selections".`);
        }else{
          out.errors.push(`"${name}" is ${actual}, but should be ${TYPE_LABEL[type]}. Delete it and add it again as ${TYPE_LABEL[type]}.`);
        }
      }

      if(type === "Choice" && opt.choices && Array.isArray(f.Choices)){
        const have = new Set(f.Choices);
        const missing = opt.choices.filter(c => !have.has(c));
        if(missing.length){
          fieldOk = false;
          const close = missing.map(m => {
            const c = f.Choices.find(x => squash(x) === squash(m));
            return c ? `"${m}" (it has "${c}", which must match exactly)` : `"${m}"`;
          });
          out.errors.push(`"${name}" is missing choice${missing.length > 1 ? "s" : ""}: ${close.join(", ")}. Edit the column and add ${missing.length > 1 ? "them" : "it"}, one per line.`);
        }
      }

      if(type === "Note" && f.TypeAsString === "Note"){
        if(f.RichText === true){
          fieldOk = false;
          if(opt.json){
            out.errors.push(`"${name}" is set to rich text. It must be plain text or saved checklists will break. Edit the column → "Specify the type of text" → Plain text.`);
          }else{
            out.warnings.push(`"${name}" is rich text. Plain text is recommended so notes don't pick up HTML formatting.`);
          }
        }
        if(f.AppendOnly === true){
          fieldOk = false;
          out.errors.push(`"${name}" has "Append Changes to Existing Text" turned on. Edit the column and set it to No.`);
        }
      }

      if(type === "DateTime" && f.TypeAsString === "DateTime" && opt.time && f.DisplayFormat === 0){
        fieldOk = false;
        out.warnings.push(`"${name}" is set to date only. Edit the column → Date and Time Format → Date & Time, or times will be lost.`);
      }

      if(opt.index && f.Indexed !== true){
        fieldOk = false;
        out.warnings.push(`"${name}" isn't indexed. Add it under List settings → Indexed columns. Needed once this list passes 5,000 items.`);
      }

      if(fieldOk) out.ok++;
    }
    return out;
  }

  let lastText = "";

  async function run(){
    const btn = document.getElementById("runSetupCheckBtn");
    const wrap = document.getElementById("setupCheckResults");
    btn.disabled = true;
    btn.textContent = "Checking...";
    wrap.innerHTML = `<p class="muted">Reading ${schema().length} lists from SharePoint...</p>`;
    try{
      const results = await Promise.all(schema().map(checkList));
      const errors = results.reduce((n, r) => n + r.errors.length, 0);
      const warnings = results.reduce((n, r) => n + r.warnings.length, 0);
      render(results, errors, warnings);
      audit(AUDIT_AREAS.admin, `Ran the SharePoint setup check: ${plural(errors, "problem")}, ${plural(warnings, "warning")}`);
    }catch(err){
      console.error(err);
      wrap.innerHTML = `<p class="error-text">The check couldn't run: ${escapeHtml(err.message)}</p>`;
    }finally{
      btn.disabled = false;
      btn.textContent = "Check SharePoint setup";
    }
  }

  function render(results, errors, warnings){
    const wrap = document.getElementById("setupCheckResults");
    const when = formatDate(new Date());
    const headline = errors
      ? `${plural(errors, "problem")} to fix${warnings ? ` and ${plural(warnings, "warning")}` : ""}.`
      : warnings ? `Everything the app needs is in place. ${plural(warnings, "warning")} worth fixing.`
      : "Everything matches. All lists and columns are set up correctly.";

    wrap.innerHTML = `
      <div class="check-summary ${errors ? "bad" : warnings ? "warn" : "good"}">
        <strong>${escapeHtml(headline)}</strong>
        <button type="button" class="btn btn-ghost btn-sm" id="copySetupCheck">Copy results</button>
      </div>
      ${results.map(r => {
        const state = r.errors.length ? "bad" : r.warnings.length ? "warn" : "good";
        const icon = { bad: "✕", warn: "!", good: "✓" }[state];
        const lines = [
          ...r.errors.map(t => `<li class="bad">${escapeHtml(t)}</li>`),
          ...r.warnings.map(t => `<li class="warn">${escapeHtml(t)}</li>`)
        ].join("");
        return `<div class="check-list ${state}">
          <div class="check-head"><span class="check-icon" aria-hidden="true">${icon}</span>
            <strong>${escapeHtml(r.list)}</strong> <span class="muted small">${escapeHtml(r.purpose)}</span>
            ${!r.missingList ? `<span class="muted small check-count">${r.ok} column${r.ok === 1 ? "" : "s"} OK</span>` : ""}
          </div>
          ${lines ? `<ul>${lines}</ul>` : ""}
        </div>`;
      }).join("")}
      <p class="muted small">Checked ${escapeHtml(when)}. Columns the app doesn't use are ignored, so extra columns are fine.
        Portal version ${escapeHtml(window.PS_PAGE_VERSION || "unknown")}${outdatedFiles().length ? `; <strong class="error-text">out-of-date files: ${escapeHtml(outdatedFiles().join(", "))}</strong>` : "; all files current"}.</p>`;

    lastText = [`Personnel Support setup check, ${when}`, `Portal version ${window.PS_PAGE_VERSION || "unknown"}` +
      (outdatedFiles().length ? `; OUT-OF-DATE FILES: ${outdatedFiles().join(", ")}` : "; all files current"), headline, ""].concat(results.map(r => {
      const head = `${r.errors.length ? "PROBLEM" : r.warnings.length ? "WARNING" : "OK"}  ${r.list}`;
      return [head, ...r.errors.map(t => `   - ${t}`), ...r.warnings.map(t => `   - (warning) ${t}`)].join("\n");
    })).join("\n");

    document.getElementById("copySetupCheck").addEventListener("click", copyResults);
  }

  async function copyResults(){
    try{
      await navigator.clipboard.writeText(lastText);
      toast("Results copied. Paste them into the chat with Claude, or an email.", { type: "success" });
    }catch(_){
      // Clipboard blocked: show the text so it can be copied by hand.
      const ta = document.createElement("textarea");
      ta.value = lastText;
      ta.className = "check-copy-fallback";
      ta.readOnly = true;
      document.getElementById("setupCheckResults").appendChild(ta);
      ta.focus();
      ta.select();
      toast("Your browser blocked copying. The results are selected below. Press Ctrl+C.", { type: "info" });
    }
  }

  function init(){
    const btn = document.getElementById("runSetupCheckBtn");
    if(btn && !btn.dataset.wired){
      btn.dataset.wired = "1";
      btn.addEventListener("click", run);
    }
  }

  return { init, run, schema };
})();

/* ============================================================
   Audit coverage check (Admin).
   SharePoint itself records who last changed every item and when
   (Modified / Modified By), whatever app or tool made the change.
   This compares those records with the Audit Log and lists every
   change that has no matching audit entry.
   Limits: SharePoint keeps only the LATEST change per item in these
   fields (earlier ones are in each item's Version history), and
   deleted items can't be seen.
   ============================================================ */

const AuditCoverage = (() => {
  const A = AUDIT_FIELDS;
  const MIN = 60 * 1000;
  const NEAR = 10 * MIN;            // allows for clock differences and slow saves
  let lastRows = [];
  let wired = false;

  function specs(){
    const R = REQ_FIELDS, W = WORK_FIELDS, L = LOG_FIELDS;
    return [
      { list: CONFIG.lists.requests, label: "Inquiries", areas: [AUDIT_AREAS.inquiry],
        select: [R.entryType], what: i => `Inquiry #${Tickets.formatId(i.Id)}: ${i.Title || ""}`,
        // emails create inquiries through the intake flow, which isn't the portal: skip untouched intake items
        intake: i => i[R.entryType] !== "Manual" },
      { list: CONFIG.lists.workItems, label: "Tracked work & employee cases", areas: [AUDIT_AREAS.work, AUDIT_AREAS.inquiry],
        bulk: true, select: [W.legacyKey], what: i => `Item #${i.Id}: ${i.Title || ""}` },
      { list: CONFIG.lists.activityLog, label: "Activity entries", areas: [AUDIT_AREAS.activity, AUDIT_AREAS.work],
        select: [L.activityName, L.identifier], what: i => `Entry #${i.Id}: ${i[L.activityName] || ""} ${i[L.identifier] || ""}` },
      { list: CONFIG.lists.employees, label: "Employees", areas: [AUDIT_AREAS.employees, AUDIT_AREAS.admin],
        bulk: true, select: [], what: i => `Employee record: ${i.Title || ""}` },
      { list: CONFIG.lists.activityTypes, label: "Activities (setup)", areas: [AUDIT_AREAS.admin], select: [], what: i => `Activity: ${i.Title || ""}` },
      { list: CONFIG.lists.team, label: "Team access", areas: [AUDIT_AREAS.admin], select: [], what: i => `Roster entry: ${i.Title || ""}` },
      { list: CONFIG.lists.alertRules, label: "Aging alerts", areas: [AUDIT_AREAS.admin], select: [], what: i => `Alert rule #${i.Id}` }
    ];
  }

  const keyOf = (email, name) => (localPart(email) || String(name || "").trim().toLowerCase());
  const areaOf = e => e[A.area] || (e[A.ticketId] ? AUDIT_AREAS.inquiry : "");

  async function fetchChanged(spec, since){
    const base = ["Id", "Title", "Created", "Modified", "Author/Title", "Author/EMail", "Editor/Title", "Editor/EMail"];
    const q = sel => filterQuery(`Modified ge ${odataDate(since)}`, `$select=${sel.join(",")}&$expand=Author,Editor&$top=2000`);
    try{
      return await spGetAll(spec.list, q(base.concat(spec.select)));
    }catch(err){
      if(err.status === 400 && spec.select.length) return spGetAll(spec.list, q(base));   // a column is missing: check with basics
      throw err;
    }
  }

  function check(item, spec, audits){
    const t = new Date(item.Modified).getTime();
    const editor = item.Editor || {};
    const who = keyOf(editor.EMail, editor.Title);
    const created = new Date(item.Created).getTime();
    // An item created by the email intake flow and never changed since isn't a portal action.
    const author = item.Author || {};
    const intakeOnly = spec.intake && spec.intake(item) && Math.abs(t - created) < MIN &&
      keyOf(author.EMail, author.Title) === who;
    const mine = audits.filter(a => {
      const k = keyOf(a[A.staffEmail], a[A.staffMember]);
      return k === who || String(a[A.staffMember] || "").toLowerCase() === String(editor.Title || "").toLowerCase();
    });
    const refs = a => a[A.recordId] === item.Id || a[A.ticketId] === item.Id ||
      new RegExp(`#${item.Id}(?!\\d)`).test(String(a[A.details] || ""));
    const at = a => new Date(a[A.logTime]).getTime();
    if(mine.some(a => spec.areas.includes(areaOf(a)) && refs(a) && Math.abs(at(a) - t) <= NEAR)) return { ok: "exact" };
    // Actions recorded without an item number (e.g. a new alert rule, a step that also wrote an activity entry)
    if(mine.some(a => spec.areas.includes(areaOf(a)) && Math.abs(at(a) - t) <= 3 * MIN)) return { ok: "nearby" };
    // Bulk imports write one summary entry when they finish
    if(spec.bulk && mine.some(a => areaOf(a) === AUDIT_AREAS.imports && at(a) >= t - 5 * MIN && at(a) <= t + 120 * MIN)) return { ok: "import" };
    if(intakeOnly) return { skip: "intake" };
    return { missing: true, who: editor.Title || editor.EMail || "(unknown)",
      onRoster: App.team.some(m => localPart(m[TEAM_FIELDS.title]) === localPart(editor.EMail)) };
  }

  async function run(){
    const btn = document.getElementById("coverageBtn");
    const out = document.getElementById("coverageResults");
    const days = Number(document.getElementById("coverageDays").value) || 7;
    const since = addDays(startOfToday(), -(days - 1));
    btn.disabled = true;
    out.innerHTML = `<p class="muted">Reading changes and audit entries for the last ${plural(days, "day")}...</p>`;
    try{
      const audits = await spGetAll(CONFIG.lists.audit,
        filterQuery(`${A.logTime} ge ${odataDate(addDays(since, -1))}`, "$top=2000"));
      const summary = [];
      lastRows = [];
      for(const spec of specs()){
        let items;
        try{ items = await fetchChanged(spec, since); }
        catch(err){ summary.push({ spec, error: err.message }); continue; }
        let ok = 0, skipped = 0, missing = 0;
        items.forEach(i => {
          const r = check(i, spec, audits);
          if(r.skip){ skipped++; return; }
          if(r.ok){ ok++; return; }
          missing++;
          lastRows.push({ when: i.Modified, list: spec.label, what: spec.what(i), who: r.who, onRoster: r.onRoster, link: i.Id });
        });
        summary.push({ spec, total: items.length, ok, skipped, missing });
      }
      lastRows.sort((a, b) => new Date(b.when) - new Date(a.when));
      render(summary, days);
      const missingTotal = lastRows.length;
      audit(AUDIT_AREAS.admin, `Ran the audit coverage check (last ${plural(days, "day")}): ${plural(missingTotal, "change")} without an audit entry`);
    }catch(err){
      console.error(err);
      out.innerHTML = `<p class="error-text">The check couldn't run: ${escapeHtml(err.message)}</p>`;
    }finally{
      btn.disabled = false;
    }
  }

  function render(summary, days){
    const out = document.getElementById("coverageResults");
    const missing = lastRows.length;
    out.innerHTML = `
      <div class="check-summary ${missing ? "bad" : "good"}">
        <strong>${missing ? `${plural(missing, "change")} in the last ${plural(days, "day")} with no audit entry.` : `Every change in the last ${plural(days, "day")} has an audit entry.`}</strong>
        ${missing ? `<button type="button" class="btn btn-ghost btn-sm" id="coverageCsv">Export list</button>` : ""}
      </div>
      <table class="plain"><thead><tr><th>List</th><th class="num">Changed</th><th class="num">Audited</th><th class="num">Email intake (skipped)</th><th class="num">Missing</th></tr></thead>
      <tbody>${summary.map(s => s.error
        ? `<tr class="static"><td>${escapeHtml(s.spec.label)}</td><td colspan="4" class="error-text">${escapeHtml(s.error)}</td></tr>`
        : `<tr class="static"><td>${escapeHtml(s.spec.label)}</td><td class="num">${s.total}</td><td class="num">${s.ok}</td><td class="num">${s.skipped || ""}</td><td class="num${s.missing ? " error-text" : ""}">${s.missing || ""}</td></tr>`).join("")}</tbody></table>
      ${missing ? `<h4 class="coverage-h">Changes without an audit entry</h4>
      <div class="table-scroll"><table class="plain"><thead><tr><th>Changed</th><th>List</th><th>Item</th><th>Changed by</th></tr></thead>
      <tbody>${lastRows.slice(0, 300).map(r => `<tr class="static"><td class="nowrap">${escapeHtml(formatDate(r.when))}</td><td>${escapeHtml(r.list)}</td><td>${escapeHtml(r.what)}</td>
        <td>${escapeHtml(r.who)}${r.onRoster ? "" : ` <span class="muted small">(not on the portal roster: likely changed in SharePoint directly or by Power Automate)</span>`}</td></tr>`).join("")}</tbody></table></div>` : ""}
      <p class="muted small">How to read this: SharePoint records the last change to each item and who made it, no matter how it was changed. A change counts as audited when the same person has a matching audit entry within 10 minutes. SharePoint keeps only each item's latest change in these fields; to see every earlier change, open the item in SharePoint and use <strong>Version history</strong>. Deleted items aren't shown.</p>`;
    const csv = document.getElementById("coverageCsv");
    if(csv) csv.addEventListener("click", () => {
      downloadCsv(`audit-coverage-${toDateInput(new Date())}.csv`,
        [["Changed", "List", "Item", "Changed by", "On portal roster"], ...lastRows.map(r => [formatDate(r.when), r.list, r.what, r.who, r.onRoster ? "Yes" : "No"])]);
      audit(AUDIT_AREAS.data, `Exported the audit coverage list (${plural(lastRows.length, "row")})`);
    });
  }

  function init(){
    if(wired) return;
    const btn = document.getElementById("coverageBtn");
    if(!btn) return;
    wired = true;
    btn.addEventListener("click", run);
  }

  return { init, run, check };
})();
