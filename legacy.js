(window.PS_FILE_VERSIONS = window.PS_FILE_VERSIONS || {})["legacy.js"] = "2026.10.05-2";
/* ============================================================
   legacy.js — 2.0 starter processes and the legacy spreadsheet import.

   STARTERS defines the employee processes built from the team's
   tracking workbook (DPD_Leaves.xlsx): their fields, checklist steps
   (linked to fields), and how each legacy tab maps onto them.

   The import reads the workbook IN THE BROWSER (SheetJS), shows a
   preview (rows, open/closed, employees found, data problems), then
   creates one case per row. Each case remembers its source row
   (LegacyKey), so running the import again skips rows already brought
   in. Spreadsheet cell colors are not read; status comes from the text.
   ============================================================ */

const Legacy = (() => {
  const W = WORK_FIELDS;
  const T = TYPE_FIELDS;

  /* ---------- value conversion ---------- */

  function serialToDay(n){
    const d = new Date(Date.UTC(1899, 11, 30) + Math.round(n * 86400000));
    const p = x => String(x).padStart(2, "0");
    return `${d.getUTCFullYear()}-${p(d.getUTCMonth() + 1)}-${p(d.getUTCDate())}`;
  }

  /* Returns { v, problem } — v is "YYYY-MM-DD" when the input is a clear date, otherwise the original text. */
  function toDay(raw){
    if(raw === null || raw === undefined || raw === "") return { v: "" };
    if(raw instanceof Date) return { v: isoDay(raw) };
    if(typeof raw === "number"){
      if(raw > 20000 && raw < 80000) return { v: serialToDay(raw) };
      return { v: String(raw), problem: "number in a date column" };
    }
    const s = String(raw).replace(/\s+/g, " ").trim();
    const all = s.match(/\d{1,2}\/\d{1,2}\/\d{2,5}/g) || [];
    if(all.length === 1){
      const [m, d, yRaw] = all[0].split("/").map(Number);
      const y = String(all[0].split("/")[2]).length === 2 ? 2000 + yRaw : yRaw;
      const dt = new Date(y, m - 1, d);
      const valid = y >= 1980 && y <= 2100 && dt.getMonth() === m - 1 && dt.getDate() === d;
      if(valid && s.replace(all[0], "").replace(/^[\s;:,-]*(yes)?[\s;:,-]*$/i, "") === "") return { v: isoDay(dt) };
      if(!valid) return { v: s, problem: `"${s}" isn't a valid date` };
    }
    return { v: s };   // words like "none", "N/A", "Waiting approval", or date ranges: kept as text
  }

  function toYes(raw){
    if(raw === null || raw === undefined || raw === "") return "";
    if(typeof raw === "number" && raw > 20000 && raw < 80000) return "Yes";
    if(raw instanceof Date) return "Yes";
    const s = String(raw).replace(/\s+/g, " ").trim();
    if(/^(yes|y|x|✓|done|true)\.?$/i.test(s)) return "Yes";
    if(/^no\.?$/i.test(s)) return "No";
    return s;   // keep the original wording ("Transferred from HRIS", "Yes; 10/7/2024", "No Data")
  }

  function toMoney(raw){
    if(raw === null || raw === undefined || raw === "") return "";
    if(typeof raw === "number") return String(raw);
    const s = String(raw).replace(/[$,\s]/g, "");
    return /^-?\d+(\.\d+)?$/.test(s) ? s : String(raw).trim();
  }

  function toText(raw){
    if(raw === null || raw === undefined) return "";
    if(raw instanceof Date) return showDay(isoDay(raw));
    return String(raw).replace(/[ \t]+\n/g, "\n").trim();
  }

  /* Choice values: match the official option ignoring case/spacing, else keep as typed. */
  function toChoice(raw, options){
    const s = toText(raw).replace(/\s+/g, " ");
    if(!s) return "";
    const hit = (options || []).find(o => o.toLowerCase().replace(/\s+/g, "") === s.toLowerCase().replace(/\s+/g, ""));
    return hit || s;
  }

  const isFilled = v => String(v ?? "").trim() !== "" && !/^no$/i.test(String(v).trim());

  /* ---------- starter processes ---------- */

  const F = (key, label, type, extra) => ({ key, label, type, ...(extra || {}) });
  const SCORES = ["6 - Intermediate", "7 - Intermediate Plus", "8 - Advanced Minus", "9 - Advanced",
    "10 - Advanced Plus", "11 - Superior Minus", "12 - Superior"];
  const LSAP_STATUS = ["Contacting EM to schedule test", "Pending Test Center to confirm testing date", "Awaiting Test Results",
    "Results Sent | Data entered into IWM and WD", "Results Sent | Did Not Meet Criteria for LSAP",
    "Missed test (1 year to reschedule)", "Missed test/Did not notify (1 year to reschedule)", "Declined LSAP"];
  const LANGS = ["Spanish", "Arabic", "Cambodian", "Chinese", "Hindi", "Korean", "Swahili", "Urdu", "Vietnamese"];

  const STARTERS = [
    {
      processKey: "asl", title: "Advanced Sick Leave", category: "Leave", order: 100,
      fields: [
        F("received", "Date request received", "date", { col: "received" }),
        F("start", "Start date", "date", { col: "start" }),
        F("end", "End date", "date", { col: "end" }),
        F("supervisor", "Supervisor", "text", { col: "supervisor" }),
        F("leaveType", "ASL / LOA type", "choice", { options: ["ASL", "ASL/Trade/Time", "Trade/Time"] }),
        F("enteredWorkday", "Entered in Workday", "yesno"),
        F("outcome", "Decision", "choice", { col: "outcome", options: ["Approved", "Denied"] }),
        F("hoursAdded", "Hours added to the absence bank", "yesno"),
        F("emailSent", "Email sent to the employee", "yesno")
      ],
      steps: [
        { text: "Entered in Workday", field: "enteredWorkday" },
        { text: "Hours added to the absence bank", field: "hoursAdded" },
        { text: "Email sent to the employee", field: "emailSent" }
      ],
      legacy: {
        sheets: ["advanced sick leave"],
        cols: { received: /^date received request/, start: /^start date/, end: /^end date/, supervisor: /^supervisor/,
          leaveType: /^asl & loa/, enteredWorkday: /^entered in workday/, outcome: /^other$/,
          hoursAdded: /^hours added/, emailSent: /^email sent to the employee/ },
        done: v => isFilled(v.outcome) || isFilled(v.emailSent)
      }
    },
    {
      processKey: "probation", title: "Extension of Probation & Training", category: "Probation & training", order: 200,
      fields: [
        F("received", "Date request received", "date", { col: "received" }),
        F("supervisor", "Supervisor", "text", { col: "supervisor" }),
        F("approvalReceived", "Date request received with approval", "date"),
        F("sentToSupervisors", "Approved and sent to supervisors", "yesno"),
        F("notes", "Notes", "longtext")
      ],
      steps: [
        { text: "Approval received", field: "approvalReceived" },
        { text: "Approved and sent to supervisors", field: "sentToSupervisors" }
      ],
      legacy: {
        sheets: ["ext of probation & training"],
        cols: { received: /^date received request$/, supervisor: /^supervisor/, approvalReceived: /with approval/,
          sentToSupervisors: /sent to supervisors/, notes: /^notes/ },
        done: v => isFilled(v.sentToSupervisors)
      }
    },
    {
      processKey: "phasedown", title: "DPD Phase Down", category: "Retirement", order: 300,
      fields: [
        F("start", "Phase down effective", "date", { col: "start" }),
        F("electedHours", "Elected hours", "number"),
        F("leaveBalances", "Leave balances", "longtext"),
        F("enteredWorkday", "Entered in Workday", "yesno"),
        F("enrollmentCompleted", "Completed phase down enrollment", "yesno"),
        F("emailDepts", "Email to all departments", "yesno"),
        F("initiateWd", "Initiate WD process", "yesno"),
        F("letterSent", "Phase down letter sent/emailed", "yesno"),
        F("estimatedDate", "Estimated phase down date", "date"),
        F("retireeStatus", "Retiree status processed in WD", "date")
      ],
      steps: [
        { text: "Entered in Workday", field: "enteredWorkday" },
        { text: "Completed phase down enrollment", field: "enrollmentCompleted" },
        { text: "Email to all departments", field: "emailDepts" },
        { text: "Initiate WD process", field: "initiateWd" },
        { text: "Phase down letter sent/emailed", field: "letterSent" },
        { text: "Process retiree status in WD", field: "retireeStatus" }
      ],
      legacy: {
        sheets: ["dpd phase down"],
        cols: { start: /^phase down effective/, electedHours: /^elected hours/, leaveBalances: /^leave balances/,
          enteredWorkday: /^entered in workday/, enrollmentCompleted: /^completed phase down/, emailDepts: /^email to all/,
          initiateWd: /^initiate wd/, letterSent: /letter sent/, estimatedDate: /^estimated phase down/, retireeStatus: /retiree status/ },
        firstName: /^employee first name/, lastName: /^employee last name/,
        done: v => isFilled(v.retireeStatus)
      }
    },
    {
      processKey: "loa", title: "Leave of Absence", category: "Leave", order: 110,
      fields: [
        F("received", "Date request received", "date", { col: "received" }),
        F("supervisor", "Employee's supervisor", "text", { col: "supervisor" }),
        F("start", "Leave start date", "date", { col: "start" }),
        F("end", "Leave end date", "date", { col: "end" }),
        F("enteredWorkday", "Date entered into Workday", "date"),
        F("emailSentBy", "Email sent by", "text"),
        F("outcome", "Approval", "choice", { col: "outcome", options: ["Approved", "Denied"] }),
        F("leaveStatus", "Leave status", "choice", { options: ["Scheduled", "On leave", "Returned"] }),
        F("placedOnLeave", "Date placed on leave (WD)", "date"),
        F("returned", "Date returned from leave (WD)", "date"),
        F("notes", "Notes", "longtext")
      ],
      steps: [
        { text: "Entered into Workday", field: "enteredWorkday" },
        { text: "Email sent", field: "emailSentBy" },
        { text: "Placed on leave in Workday", field: "placedOnLeave" },
        { text: "Returned from leave in Workday", field: "returned" }
      ],
      legacy: {
        sheets: ["leave of absence request"],
        cols: { received: /^date request received/, supervisor: /supervisor/, start: /^leave start date/, end: /^leave end date/,
          enteredWorkday: /^date entered into workday/, emailSentBy: /^e-?mail sent by/, outcome: /^approval$/,
          leaveStatus: /^status$/, placedOnLeave: /placed on leave/, returned: /returned from leave/, notes: /^notes/ },
        done: v => /returned/i.test(v.leaveStatus) || /denied/i.test(v.outcome) || isFilled(v.returned)
      }
    },
    {
      processKey: "tuition", title: "Tuition Reimbursement", category: "Tuition", order: 400,
      fields: [
        F("amount", "Amount", "money"),
        F("course", "Course title or name", "text"),
        F("classEnd", "Class end date", "date"),
        F("gradesDue", "Grades needed by", "date"),
        F("emailSent", "Email sent to employee", "yesno"),
        F("gradesReceived", "Final grades received", "yesno"),
        F("paymentSubmitted", "One-time payment submitted", "yesno"),
        F("enteredWorkday", "Date $$ entered in Workday", "date"),
        F("notifiedEmployee", "Employee notified", "yesno"),
        F("paymentReceived", "Payment received", "yesno"),
        F("outcome", "Outcome", "choice", { col: "outcome", options: ["Approved", "Dropped", "Denied"] }),
        F("notes", "Notes", "longtext")
      ],
      steps: [
        { text: "Email sent to employee", field: "emailSent" },
        { text: "Final grades received", field: "gradesReceived" },
        { text: "One-time payment submitted", field: "paymentSubmitted" },
        { text: "$$ entered in Workday", field: "enteredWorkday" },
        { text: "Employee notified", field: "notifiedEmployee" },
        { text: "Payment received", field: "paymentReceived" }
      ],
      legacy: {
        sheets: ["tuition reimbursement", "tuition reimbursement (old)"],
        cols: { amount: /^amount/, course: /^course title/, classEnd: /^class end date/, gradesDue: /^grades needed by/,
          emailSent: /^e-?mail sent/, notifiedEmployee: /^notified employee/, enteredWorkday: /^date entered \$\$ in workday/,
          notes: /^notes/, _final: /^final grades received/, _payment: /^one-time payment submitted/ },
        // Sheet-specific columns ("_final" on the old tab, "_payment" on the new one) become several fields.
        transform(v, raw, problems){
          if(raw._final !== undefined && raw._final !== null && raw._final !== ""){
            const r = raw._final;
            const s = toText(r);
            if(typeof r === "number" || r instanceof Date){
              const d = toDay(r);
              v.gradesReceived = "Yes"; v.paymentSubmitted = "Yes"; v.enteredWorkday = d.v;
            }else if(/^(yes|done)|\(done\)/i.test(s)){
              v.gradesReceived = "Yes"; v.paymentSubmitted = "Yes";
              if(!/^yes\.?$/i.test(s)) v.notes = [v.notes, s].filter(Boolean).join("\n");
            }else if(/dropped|no longer available/i.test(s)){
              v.outcome = "Dropped";
              v.notes = [v.notes, s].filter(Boolean).join("\n");
            }else{
              v.notes = [v.notes, s].filter(Boolean).join("\n");   // "Waiting on grades" stays open
            }
          }
          if(raw._payment){
            const s = toText(raw._payment);
            if(/payment received/i.test(s)){ v.paymentSubmitted = "Yes"; v.paymentReceived = "Yes"; }
            else if(/submitted/i.test(s)) v.paymentSubmitted = "Yes";
            else v.notes = [v.notes, s].filter(Boolean).join("\n");
          }
        },
        done: (v, sheet) => /dropped|denied/i.test(v.outcome) || isFilled(v.notifiedEmployee) || isFilled(v.paymentReceived) ||
          (/\(old\)/.test(sheet) && isFilled(v.gradesReceived))
      }
    },
    {
      processKey: "ppl", title: "Paid Parental Leave", category: "Leave", order: 120,
      fields: [
        F("received", "Request date", "date", { col: "received" }),
        F("supervisor", "Supervisor", "text", { col: "supervisor" }),
        F("proofSubmitted", "Proof submitted", "date"),
        F("start", "First day out (first day of leave)", "date", { col: "start" }),
        F("end", "Estimated last day out", "date", { col: "end" }),
        F("pplAdded", "PPL added in Workday", "yesno"),
        F("returned", "Returned from leave in Workday (actual last day)", "date"),
        F("audit", "Audit", "yesno"),
        F("notes", "Notes", "longtext")
      ],
      steps: [
        { text: "Proof submitted", field: "proofSubmitted" },
        { text: "PPL added in Workday", field: "pplAdded" },
        { text: "Returned from leave in Workday", field: "returned" },
        { text: "Audit completed", field: "audit" }
      ],
      legacy: {
        sheets: ["paid parental leave"],
        cols: { received: /^request date/, supervisor: /^supervisor/, proofSubmitted: /^proof submitted/,
          start: /first day out/, end: /last day out/, pplAdded: /^ppl added/, returned: /^returned from leave/,
          audit: /^audit/, notes: /^column1$|^notes/ },
        done: v => isFilled(v.returned)
      }
    },
    {
      processKey: "rtp", title: "Return to Patrol", category: "Assignments", order: 500,
      fields: [
        F("received", "Date request received", "date", { col: "received" }),
        F("start", "Start date", "date", { col: "start" }),
        F("end", "End date", "date", { col: "end" }),
        F("supervisor", "Supervisor", "text", { col: "supervisor" }),
        F("emailChain", "Email sent to chain of command", "yesno")
      ],
      steps: [{ text: "Email sent to chain of command", field: "emailChain" }],
      legacy: {
        sheets: ["return to patrol"],
        cols: { received: /^date received request/, start: /^start date/, end: /^end date/, supervisor: /^supervisor/,
          emailChain: /chain of comman/ },
        done: v => isFilled(v.emailChain)
      }
    },
    {
      processKey: "tradetime", title: "Trade Time", category: "Leave", order: 130,
      fields: [
        F("received", "Date request received", "date", { col: "received" }),
        F("supervisor", "Supervisor", "text", { col: "supervisor" }),
        F("start", "Start date", "date", { col: "start" }),
        F("end", "End date", "date", { col: "end" }),
        F("emailPio", "Email sent to PIO", "yesno"),
        F("infoCco", "Information sent to CCO", "yesno"),
        F("enteredSheet", "Entered on trade time tracking", "date")
      ],
      steps: [
        { text: "Email sent to PIO", field: "emailPio" },
        { text: "Information sent to CCO", field: "infoCco" }
      ],
      legacy: {
        sheets: ["trade time"],
        cols: { supervisor: /^supervisor/, received: /^date received request/, start: /^start date/, end: /^end date/,
          emailPio: /^email sent to pio/, infoCco: /sent to cco/, enteredSheet: /^entered on spreadsheet/ },
        done: v => isFilled(v.emailPio) && isFilled(v.infoCco)
      }
    },
    {
      processKey: "lsap", title: "LSAP (Language Skills Pay)", category: "Language pay", order: 600,
      fields: [
        F("received", "LSAP memo received", "date", { col: "received" }),
        F("language", "Language requested", "choice", { options: LANGS }),
        F("contacted", "Employee contacted to schedule test", "yesno"),
        F("testStatus", "Test status", "choice", { options: LSAP_STATUS }),
        F("testDate", "Test date", "date"),
        F("score", "LSAP score", "choice", { options: SCORES }),
        F("resultsSent", "Results sent", "yesno"),
        F("payEntered", "$$ entered in Workday & IWM", "date"),
        F("notes", "Notes", "longtext")
      ],
      steps: [
        { text: "Contacted employee to schedule test", field: "contacted" },
        { text: "Test taken", field: "testDate" },
        { text: "Results sent", field: "resultsSent" },
        { text: "$$ entered in Workday & IWM", field: "payEntered" }
      ],
      legacy: {
        sheets: ["lsap"],
        cols: { received: /^lsap memo received/, language: /^language requested/, testStatus: /^contacting em/,
          testDate: /^test date/, score: /^lsap score/, payEntered: /^entered \$\$ in workday/, notes: /^notes/ },
        transform: lsapTransform,
        done: v => isFilled(v.payEntered) || /declined|did not meet|missed/i.test(v.testStatus)
      }
    },
    {
      processKey: "lsaprecert", title: "LSAP Recertification", category: "Language pay", order: 610,
      fields: [
        F("received", "Initial notification sent", "date", { col: "received" }),
        F("language", "Language", "choice", { options: LANGS }),
        F("contacted", "Employee contacted to schedule test", "yesno"),
        F("testStatus", "Test status", "choice", { options: LSAP_STATUS }),
        F("testDate", "Test date", "date"),
        F("score", "LSAP score", "choice", { options: SCORES }),
        F("resultsSent", "Results sent", "yesno"),
        F("payEntered", "$$ entered in Workday & IWM", "date"),
        F("notes", "Notes", "longtext")
      ],
      steps: [
        { text: "Initial notification sent", field: "received" },
        { text: "Contacted employee to schedule test", field: "contacted" },
        { text: "Test taken", field: "testDate" },
        { text: "Results sent", field: "resultsSent" },
        { text: "$$ entered in Workday & IWM (or no pay change)", field: "payEntered" }
      ],
      legacy: {
        sheets: ["lsap recert"],
        cols: { received: /^date initial notif/, language: /^language requested/, testStatus: /^contacting em/,
          testDate: /^test date/, score: /^lsap score/, payEntered: /^entered \$\$ in workday/, notes: /^notes/ },
        transform: lsapTransform,
        done: v => isFilled(v.payEntered) || /declined|did not meet|missed/i.test(v.testStatus) ||
          /no change in pay|no pay change|did not respond|no re?sponse|removed|retired/i.test(v.notes)
      }
    }
  ];

  function lsapTransform(v){
    if(isFilled(v.testStatus)) v.contacted = "Yes";
    if(/results sent/i.test(v.testStatus)) v.resultsSent = "Yes";
    // A status pasted into the score column belongs in Test status.
    if(/results sent/i.test(v.score)){ v.score = ""; }
  }

  function starterDef(s){
    return JSON.stringify({ employee: true, processKey: s.processKey, fields: s.fields });
  }

  function existingProcess(key){
    return Activity.types.find(t => { const d = Cases.def(t); return d && d.processKey === key; }) || null;
  }

  async function installStarters(){
    const btn = document.getElementById("installStartersBtn");
    const out = document.getElementById("starterResult");
    btn.disabled = true;
    let added = 0;
    const skipped = [];
    try{
      for(const s of STARTERS){
        if(existingProcess(s.processKey)){ skipped.push(s.title); continue; }
        await spCreate(CONFIG.lists.activityTypes, {
          Title: s.title,
          [T.category]: s.category,
          [T.team]: TEAM_BOTH,
          [T.inputType]: "Employee Number",
          [T.inputLabel]: "Employee #",
          [T.duplicateWindowDays]: 0,
          [T.sortOrder]: s.order,
          [T.description]: "",
          [T.active]: true,
          [T.checklistSteps]: JSON.stringify(s.steps),
          [T.caseFields]: starterDef(s)
        });
        added++;
      }
      await Activity.reloadTypes();
      Cases.invalidate();
      out.innerHTML = `<p><strong>${added} process${added === 1 ? "" : "es"} added.</strong>${skipped.length ? ` Already installed: ${escapeHtml(skipped.join(", "))}.` : ""} Set each one's team (SRU/PSU) in Activities above.</p>`;
      if(added) audit(AUDIT_AREAS.admin, `Installed ${added} starter employee processes`, { details: STARTERS.filter(s => !skipped.includes(s.title)).map(s => s.title).join(", ") });
    }catch(err){
      console.error(err);
      out.innerHTML = `<p class="error-text">Stopped after ${added}: ${escapeHtml(err.message)}. Run the setup check; the Activity Types list may be missing the CaseFields column.</p>`;
    }finally{
      btn.disabled = false;
    }
  }

  /* ---------- reading the legacy workbook ---------- */

  const norm = s => String(s ?? "").replace(/\s+/g, " ").trim().toLowerCase();

  function findStarterForSheet(sheetName){
    const n = norm(sheetName);
    return STARTERS.find(s => s.legacy.sheets.includes(n)) || null;
  }

  function parseSheet(X, ws, sheetName, starter){
    const aoa = X.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    const h = aoa.findIndex((r, i) => i < 12 && r.some(c => /employee\s*id/.test(norm(c))));
    if(h < 0) return { error: "No header row with \"Employee ID\" found" };
    const headers = aoa[h].map(norm);
    const colOf = re => headers.findIndex(x => x && re.test(x));
    const L = starter.legacy;
    const idCol = colOf(/employee\s*id/);
    const lastCol = colOf(L.lastName || /^last name$/);
    const firstCol = colOf(L.firstName || /^first name$/);
    const nameCol = colOf(/^name$/);
    const map = {};
    Object.entries(L.cols).forEach(([k, re]) => { const i = colOf(re); if(i >= 0) map[k] = i; });
    const fieldByKey = Object.fromEntries(starter.fields.map(f => [f.key, f]));

    const records = [];
    for(let r = h + 1; r < aoa.length; r++){
      const row = aoa[r] || [];
      const raw = {};
      Object.entries(map).forEach(([k, i]) => { raw[k] = row[i]; });
      const id = idCol >= 0 ? Roster.normId(row[idCol]) : "";
      let first = firstCol >= 0 ? toText(row[firstCol]) : "", last = lastCol >= 0 ? toText(row[lastCol]) : "";
      if(!first && !last && nameCol >= 0){
        const full = toText(row[nameCol]);
        const parts = full.split(/\s+/);
        first = parts.length > 1 ? parts.slice(0, -1).join(" ") : full;
        last = parts.length > 1 ? parts.slice(-1)[0] : "";
      }
      const hasData = Object.values(raw).some(v => v !== null && v !== undefined && String(v).trim() !== "");
      if(!hasData && !id) continue;

      const v = {};
      const problems = [];
      Object.entries(raw).forEach(([k, val]) => {
        if(k.startsWith("_")) return;
        const f = fieldByKey[k];
        if(!f) return;
        if(f.type === "date"){
          const d = toDay(val);
          v[k] = d.v;
          if(d.problem) problems.push(`${f.label}: ${d.problem}`);
        }else if(f.type === "yesno") v[k] = toYes(val);
        else if(f.type === "money") v[k] = toMoney(val);
        else if(f.type === "number") v[k] = typeof val === "number" ? String(val) : toText(val);
        else if(f.type === "choice") v[k] = toChoice(val, f.options);
        else v[k] = toText(val);
      });
      if(L.transform) L.transform(v, raw, problems);
      Object.keys(v).forEach(k => { if(v[k] === "") delete v[k]; });

      const done = !!L.done(v, norm(sheetName));
      const days = Object.values(v).filter(x => /^\d{4}-\d{2}-\d{2}$/.test(x)).sort();
      const startDay = v.received || v.start || days[0] || "";
      const rowSig = JSON.stringify([id, first, last, v]);
      records.push({
        sheet: sheetName, row: r + 1, id, first, last, values: v, done, problems,
        startDay, lastDay: days[days.length - 1] || startDay,
        key: `${starter.processKey}|${sheetName} row ${r + 1}|${hash(rowSig)}`
      });
    }
    return { records, headerRow: h + 1, unmatched: aoa[h].filter((c, i) => c && !Object.values(map).includes(i) && ![idCol, lastCol, firstCol, nameCol].includes(i)).map(c => String(c).replace(/\s+/g, " ").trim()) };
  }

  function hash(s){
    let h = 2166136261;
    for(let i = 0; i < s.length; i++){ h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); }
    return (h >>> 0).toString(36);
  }

  /* ---------- import UI ---------- */

  let parsed = null;   // { fileName, bySheet: [{ sheet, starter, records, error, unmatched }] }

  function init(){
    const btn = document.getElementById("installStartersBtn");
    if(btn && !btn.dataset.wired){
      btn.dataset.wired = "1";
      btn.addEventListener("click", installStarters);
      document.getElementById("legacyFile").addEventListener("change", e => { if(e.target.files[0]) readWorkbook(e.target.files[0]); });
    }
    onView("adminView", renderStarterStatus);
  }

  function renderStarterStatus(){
    const el = document.getElementById("starterStatus");
    if(!el) return;
    const have = STARTERS.filter(s => existingProcess(s.processKey));
    el.textContent = have.length === STARTERS.length
      ? `All ${STARTERS.length} starter processes are installed.`
      : `${have.length} of ${STARTERS.length} installed.`;
  }

  async function readWorkbook(file){
    const out = document.getElementById("legacyArea");
    out.innerHTML = `<p class="muted">Reading ${escapeHtml(file.name)} on this computer...</p>`;
    try{
      await Roster.ready();
      const X = await loadSheetJS();
      const wb = X.read(await file.arrayBuffer(), { type: "array", cellDates: false });
      const bySheet = wb.SheetNames.map(name => {
        const starter = findStarterForSheet(name);
        if(!starter) return { sheet: name, starter: null };
        const res = parseSheet(X, wb.Sheets[name], name, starter);
        return { sheet: name, starter, ...res };
      });
      parsed = { fileName: file.name, bySheet };
      renderPreview();
    }catch(err){
      console.error(err);
      out.innerHTML = `<p class="error-text">The workbook couldn't be read: ${escapeHtml(err.message)}</p>`;
    }
  }

  function renderPreview(){
    const out = document.getElementById("legacyArea");
    const missing = [...new Set(parsed.bySheet.filter(s => s.starter && !existingProcess(s.starter.processKey)).map(s => s.starter.title))];
    let total = 0, withId = 0, inRoster = 0, open = 0, probs = 0;
    const rows = parsed.bySheet.map(s => {
      if(!s.starter) return `<tr class="static"><td>${escapeHtml(s.sheet)}</td><td colspan="6" class="muted">Not recognized; skipped</td></tr>`;
      if(s.error) return `<tr class="static"><td>${escapeHtml(s.sheet)}</td><td colspan="6" class="error-text">${escapeHtml(s.error)}</td></tr>`;
      const rec = s.records;
      const ids = rec.filter(r => r.id).length;
      const roster = rec.filter(r => r.id && Roster.find(r.id)).length;
      const op = rec.filter(r => !r.done).length;
      const pr = rec.reduce((n, r) => n + r.problems.length, 0);
      total += rec.length; withId += ids; inRoster += roster; open += op; probs += pr;
      return `<tr class="static"><td>${escapeHtml(s.sheet)}</td><td>${escapeHtml(s.starter.title)}</td>
        <td class="num">${rec.length}</td><td class="num">${rec.length - op}</td><td class="num">${op}</td>
        <td class="num">${ids}${ids ? ` <span class="muted small">(${roster} in roster)</span>` : ""}</td>
        <td class="num">${pr || ""}</td></tr>`;
    }).join("");
    const problemList = parsed.bySheet.flatMap(s => (s.records || []).flatMap(r => r.problems.map(p => `${s.sheet}, row ${r.row}: ${p}`)));
    const unmatched = parsed.bySheet.filter(s => s.unmatched && s.unmatched.length).map(s => `${s.sheet}: ${s.unmatched.join(", ")}`);
    out.innerHTML = `
      <p><strong>${escapeHtml(parsed.fileName)}</strong></p>
      ${missing.length ? `<p class="error-text">Install the starter processes first (above). Missing: ${escapeHtml(missing.join(", "))}.</p>` : ""}
      <div class="table-scroll"><table class="plain">
        <thead><tr><th>Tab</th><th>Becomes</th><th class="num">Rows</th><th class="num">Done</th><th class="num">Open</th><th class="num">With employee ID</th><th class="num">Data issues</th></tr></thead>
        <tbody>${rows}</tbody>
        <tfoot><tr><th colspan="2">Total</th><td class="num">${total}</td><td class="num">${total - open}</td><td class="num">${open}</td><td class="num">${withId}</td><td class="num">${probs || ""}</td></tr></tfoot>
      </table></div>
      ${problemList.length ? `<details class="mine"><summary>${problemList.length} data issues (kept as text; fix them on the case after import)</summary><ul class="small">${problemList.slice(0, 200).map(p => `<li>${escapeHtml(p)}</li>`).join("")}</ul></details>` : ""}
      ${unmatched.length ? `<details class="mine"><summary>Columns not imported</summary><ul class="small">${unmatched.map(u => `<li>${escapeHtml(u)}</li>`).join("")}</ul></details>` : ""}
      <div class="legacy-options">
        <label class="check"><input type="checkbox" id="legacyCloseOld" checked> Close open rows received more than
          <input type="number" id="legacyCloseDays" value="180" min="30" class="narrow-input"> days ago</label>
        <div class="hint">The spreadsheet often doesn't mark old rows finished. Without this, every unfinished old row becomes an open case in Tracked work.</div>
        <label class="check"><input type="checkbox" id="legacyAddPeople" checked> Add people missing from the roster as inactive employees (so their history has a record)</label>
      </div>
      <button type="button" class="btn btn-navy" id="legacyRunBtn" ${missing.length || !total ? "disabled" : ""}>Import ${total.toLocaleString()} rows</button>
      <div id="legacyProgress"></div>`;
    const run = document.getElementById("legacyRunBtn");
    if(run) run.addEventListener("click", runImport);
  }

  async function runImport(){
    const btn = document.getElementById("legacyRunBtn");
    btn.disabled = true;
    const prog = document.getElementById("legacyProgress");
    const closeOld = document.getElementById("legacyCloseOld").checked;
    const closeDays = Number(document.getElementById("legacyCloseDays").value) || 180;
    const addPeople = document.getElementById("legacyAddPeople").checked;
    const cutoff = isoDay(addDays(new Date(), -closeDays));
    const nowIso = new Date().toISOString();
    prog.innerHTML = `<p class="muted">Checking for rows already imported...</p>`;

    // 1. skip rows already imported (re-running is safe)
    const typeFor = {};
    const existingKeys = new Set();
    for(const s of parsed.bySheet){
      if(!s.starter || s.error) continue;
      const t = existingProcess(s.starter.processKey);
      typeFor[s.starter.processKey] = t;
      if(t && !s._keysLoaded){
        const rows = await spGetAll(CONFIG.lists.workItems,
          filterQuery(`${W.activityTypeId} eq ${t.Id}`, `$select=Id,${W.legacyKey}&$top=2000`));
        rows.forEach(r => { if(r[W.legacyKey]) existingKeys.add(r[W.legacyKey]); });
      }
    }
    const todo = parsed.bySheet.flatMap(s => (s.starter && !s.error ? s.records.map(r => ({ ...r, starter: s.starter })) : []))
      .filter(r => !existingKeys.has(r.key));
    const already = parsed.bySheet.reduce((n, s) => n + (s.records ? s.records.length : 0), 0) - todo.length;

    // 2. people not in the roster
    let peopleAdded = 0;
    if(addPeople){
      const need = new Map();
      todo.forEach(r => { if(r.id && !Roster.find(r.id) && !need.has(r.id)) need.set(r.id, r); });
      const list = [...need.values()];
      if(list.length){
        prog.innerHTML = `<p class="muted">Adding ${list.length} people from the old records...</p>`;
        const res = await runPool(list, r => Roster.ensureEmployee(r.id, r.first, r.last), 6);
        peopleAdded = res.ok;
      }
    }

    // 3. cases
    let closedByAge = 0;
    const show = (d, n) => { prog.innerHTML = `<div class="progress progress-lg"><span style="width:${Math.round(d / n * 100)}%"></span></div><p class="muted small">${d.toLocaleString()} of ${n.toLocaleString()} cases saved. Keep this tab open.</p>`; };
    const res = await runPool(todo, async r => {
      const type = typeFor[r.starter.processKey];
      const emp = r.id ? Roster.find(r.id) : null;
      const name = emp ? emp[EMP_FIELDS.title] : Roster.fullName(r.first, r.last);
      let closed = r.done;
      if(!closed && closeOld && r.startDay && r.startDay < cutoff){ closed = true; closedByAge++; }
      const steps = Activity.templateSteps(type).map(s => {
        const val = s.field ? r.values[s.field] : "";
        const done = s.field ? isFilled(val) : closed;
        return { id: Math.random().toString(36).slice(2, 10), text: s.text, external: !!s.external, ...(s.field ? { field: s.field } : {}),
          assignee: "", assigneeName: "", done, doneBy: done ? "Legacy import" : "",
          doneOn: done && /^\d{4}-\d{2}-\d{2}$/.test(val || "") ? dayToIso(val) : "" };
      });
      const startedIso = dayToIso(r.startDay) || nowIso;
      const shell = { [W.activityTypeId]: type.Id, [W.caseData]: "{}" };
      const body = {
        [W.title]: `${type.Title} — ${name || r.id || `row ${r.row}`}`.slice(0, 255),
        [W.activityTypeId]: type.Id,
        [W.activityName]: type.Title,
        [W.category]: type[T.category] || "",
        [W.identifier]: r.id || "",
        [W.status]: closed ? "Completed" : "Open",
        [W.ownerName]: App.user.name,
        [W.ownerEmail]: App.user.username,
        [W.startedBy]: "Legacy import",
        [W.startedOn]: startedIso,
        [W.lastActivityOn]: closed ? (dayToIso(r.lastDay) || startedIso) : nowIso,
        [W.steps]: JSON.stringify(steps),
        [W.employeeId]: r.id || "",
        [W.employeeName]: name || "",
        [W.legacyKey]: r.key,
        ...Cases.bodyFor(shell, r.values)
      };
      if(closed){
        body[W.closedOn] = dayToIso(r.lastDay) || startedIso;
        body[W.closedBy] = "Legacy import";
      }
      await spCreate(CONFIG.lists.workItems, body);
    }, 6, show);

    Cases.invalidate();
    await Work.refresh();
    prog.innerHTML = `<p class="${res.failed.length ? "warn-text" : ""}"><strong>Import finished.</strong> ${res.ok.toLocaleString()} cases created${already ? `, ${already} rows skipped (already imported)` : ""}${closedByAge ? `, ${closedByAge} old unfinished rows closed` : ""}${peopleAdded ? `, ${peopleAdded} people added to the roster as inactive` : ""}.${res.failed.length ? ` ${res.failed.length} failed (${escapeHtml(res.failed[0].err.message)}); run the import again to retry just those.` : ""} See them on the Employee processes tab.</p>`;
    audit(AUDIT_AREAS.imports, `Imported legacy workbook ${parsed.fileName}: ${res.ok} cases created, ${already} already imported, ${closedByAge} closed by age, ${peopleAdded} people added${res.failed.length ? `, ${res.failed.length} failed` : ""}`);
  }

  return { init, STARTERS, installStarters, toDay, toYes, parseSheet, findStarterForSheet };
})();
