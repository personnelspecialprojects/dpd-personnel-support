(window.PS_FILE_VERSIONS = window.PS_FILE_VERSIONS || {})["setup-check.js"] = "2026.10.02-2";
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
          AL = ALERT_FIELDS, M = TEAM_FIELDS, W = WORK_FIELDS, WL = WORKLOG_FIELDS;
    return [
      { list: CONFIG.lists.team, purpose: "who can sign in", fields: [
        [M.role, "Choice", { choices: ["Staff", "Admin"] }], [M.team, "Choice", { choices: TEAMS }]
      ]},
      { list: CONFIG.lists.activityTypes, purpose: "the activities on the Dashboard", fields: [
        [T.category, "Text"], [T.inputType, "Choice", { choices: INPUT_TYPES }], [T.inputLabel, "Text"],
        [T.duplicateWindowDays, "Number"], [T.sortOrder, "Number"], [T.description, "Text"],
        [T.active, "Boolean"], [T.checklistSteps, "Note", { json: true }],
        [T.team, "Choice", { choices: [...TEAMS, TEAM_BOTH] }]
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
        [R.source, "Choice", { choices: ["Email", "Phone", "Walk-in"] }],
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
        [W.activityTypeId, "Number"], [W.activityName, "Text"], [W.category, "Text"], [W.identifier, "Text"],
        [W.status, "Choice", { index: true, choices: ["Open", "Completed", "Cancelled"] }],
        [W.ownerName, "Text"], [W.ownerEmail, "Text"], [W.startedBy, "Text"],
        [W.startedOn, "DateTime", { time: true }], [W.lastActivityOn, "DateTime", { time: true }],
        [W.closedOn, "DateTime", { index: true, time: true }], [W.closedBy, "Text"],
        [W.steps, "Note", { json: true }], [W.completionLogId, "Number"]
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
