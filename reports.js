(window.PS_FILE_VERSIONS = window.PS_FILE_VERSIONS || {})["reports.js"] = "2026.10.07-1";
/* ============================================================
   reports.js — activity counts by person and task for any date
   range, plus inquiries resolved, exports, and a lookup that
   answers "has anyone already done this for employee #12345?"
   ============================================================ */

const Reports = (() => {
  const F = LOG_FIELDS;
  const R = REQ_FIELDS;
  const INQUIRY_KEY = "__inquiries__";
  const SPECIAL_KEY = "__special_inquiries__";
  let wired = false;
  let loadedOnce = false;
  let entries = [];
  let range = null;      // { from: Date, to: Date (exclusive) }
  let matrix = null;

  function init(){
    if(!wired) wire();
    onView("reportsView", () => { if(!loadedOnce) run(); });
  }

  function wire(){
    wired = true;
    document.getElementById("rpPreset").addEventListener("change", () => {
      applyPreset(document.getElementById("rpPreset").value);
      run();
    });
    ["rpFrom", "rpTo"].forEach(id => document.getElementById(id).addEventListener("change", () => {
      document.getElementById("rpPreset").value = "custom";
    }));
    document.getElementById("rpRunBtn").addEventListener("click", run);
    const teamSel = document.getElementById("rpTeam");
    teamSel.innerHTML = `<option value="">All teams</option>` +
      TEAMS.map(tm => `<option value="${escapeHtml(tm)}">${escapeHtml(tm)}</option>`).join("");
    teamSel.addEventListener("change", () => { if(range) render(); });
    document.getElementById("rpExportSummaryBtn").addEventListener("click", exportSummary);
    document.getElementById("rpExportEntriesBtn").addEventListener("click", exportEntries);
    document.getElementById("lookupBtn").addEventListener("click", lookup);
    document.getElementById("lookupInput").addEventListener("keydown", e => {
      if(e.key === "Enter"){ e.preventDefault(); lookup(); }
    });
    applyPreset("thisMonth");
  }

  function presetRange(p){
    const today = startOfToday();
    const tomorrow = addDays(today, 1);
    const y = today.getFullYear(), m = today.getMonth();
    switch(p){
      case "today": return [today, tomorrow];
      case "thisWeek": return [addDays(today, -today.getDay()), tomorrow];
      case "lastWeek": { const s = addDays(today, -today.getDay() - 7); return [s, addDays(s, 7)]; }
      case "lastMonth": return [new Date(y, m - 1, 1), new Date(y, m, 1)];
      case "last30": return [addDays(today, -29), tomorrow];
      case "thisMonth":
      default: return [new Date(y, m, 1), tomorrow];
    }
  }

  function applyPreset(p){
    if(p === "custom") return;
    const [from, to] = presetRange(p);
    document.getElementById("rpFrom").value = toDateInput(from);
    document.getElementById("rpTo").value = toDateInput(addDays(to, -1));  // inputs show inclusive end
  }

  function readRange(){
    const from = fromDateInput(document.getElementById("rpFrom").value);
    const toInclusive = fromDateInput(document.getElementById("rpTo").value);
    if(!from || !toInclusive || toInclusive < from) return null;
    return { from, to: addDays(toInclusive, 1) };
  }

  async function run(){
    const r = readRange();
    const out = document.getElementById("rpMatrix");
    if(!r){ toast("Pick a start date on or before the end date.", { type: "error" }); return; }
    range = r;
    loadedOnce = true;
    out.innerHTML = `<p class="muted">Loading...</p>`;
    try{
      const expr = `${F.loggedAt} ge ${odataDate(r.from)} and ${F.loggedAt} lt ${odataDate(r.to)}`;
      const items = await spGetAll(CONFIG.lists.activityLog, filterQuery(expr, "$top=2000"));
      entries = items.filter(e => !e[F.voided]);
      render();
    }catch(err){
      console.error(err);
      out.innerHTML = `<p class="error-text">The report couldn't load: ${escapeHtml(err.message)}</p>`;
    }
  }

  function selectedTeam(){ return document.getElementById("rpTeam").value; }

  /* Entries for the chosen team: its own activities plus shared ("Both") ones.
     Activities deleted from the list count as shared. */
  function teamEntries(){
    const team = selectedTeam();
    if(!team) return entries;
    const byId = {};
    Activity.types.forEach(t => { byId[t.Id] = t; });
    return entries.filter(e => {
      const t = byId[e[F.activityTypeId]];
      const tm = t ? Activity.teamOf(t) : TEAM_BOTH;
      return tm === team || tm === TEAM_BOTH;
    });
  }

  function resolvedInRange(){
    // Inquiries are shared by the whole Personnel Support team, so they show only under "All teams".
    if(selectedTeam()) return [];
    return Tickets.all.filter(t => {
      if(Tickets.statusOf(t) !== "Completed" || !t[R.completedOn]) return false;
      const d = new Date(t[R.completedOn]);
      return d >= range.from && d < range.to;
    });
  }

  function buildMatrix(){
    const resolved = resolvedInRange();
    const staff = new Set();
    const cells = {};      // rowKey -> { staffName -> n }
    const names = {};      // typeId -> latest snapshot name/category
    const bump = (row, person, n = 1) => {
      staff.add(person);
      const r = cells[row] || (cells[row] = {});
      r[person] = (r[person] || 0) + n;
    };
    teamEntries().forEach(e => {
      const key = String(e[F.activityTypeId]);
      names[key] = names[key] || { name: e[F.activityName], category: e[F.category] };
      bump(key, e[F.staffName] || "Unknown", Activity.qty(e));   // Count activities add their number
    });
    resolved.forEach(t => bump(Tickets.queueOf(t) === "special" ? SPECIAL_KEY : INQUIRY_KEY, t[R.completedBy] || "Unassigned"));

    // Row order follows the board: current activity types first (by category/sort order),
    // then anything retired that still has entries in this range.
    const typesById = {};
    Activity.types.forEach(t => { typesById[String(t.Id)] = t; });
    const groups = Activity.groupByCategory(Activity.types.filter(t => cells[String(t.Id)]))
      .map(g => ({ name: g.name, rows: g.types.map(t => ({ key: String(t.Id), name: t.Title })) }));
    const orphanRows = Object.keys(cells)
      .filter(k => k !== INQUIRY_KEY && k !== SPECIAL_KEY && !typesById[k])
      .map(k => ({ key: k, name: `${names[k].name || "Unknown activity"} (retired)` }));
    if(orphanRows.length) groups.push({ name: "Retired activities", rows: orphanRows });
    const inqRows = [];
    if(cells[INQUIRY_KEY]) inqRows.push({ key: INQUIRY_KEY, name: "Inquiries resolved" });
    if(cells[SPECIAL_KEY]) inqRows.push({ key: SPECIAL_KEY, name: `${SPECIAL_QUEUE} resolved` });
    if(inqRows.length) groups.push({ name: "Inquiries", rows: inqRows });

    const people = [...staff].sort((a, b) => a.localeCompare(b));
    return { groups, people, cells, resolved };
  }

  function render(){
    matrix = buildMatrix();
    const { groups, people, cells, resolved } = matrix;
    const out = document.getElementById("rpMatrix");
    const toLabel = addDays(range.to, -1).toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
    const fromLabel = range.from.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });

    const isOpenInq = t => ["New", "In Progress"].includes(Tickets.statusOf(t));
    const open = Tickets.all.filter(t => isOpenInq(t) && Tickets.queueOf(t) === "main").length;
    const openSpecial = Tickets.all.filter(t => isOpenInq(t) && Tickets.queueOf(t) === "special").length;
    const resolvedSpecial = resolved.filter(t => Tickets.queueOf(t) === "special").length;
    const noAction = Tickets.all.filter(t => {
      if(Tickets.statusOf(t) !== NO_ACTION_STATUS || !t[R.completedOn]) return false;
      const d = new Date(t[R.completedOn]);
      return d >= range.from && d < range.to;
    }).length;
    const days = resolved
      .filter(t => t[R.receivedOn])
      .map(t => (new Date(t[R.completedOn]) - new Date(t[R.receivedOn])) / 86400000);
    const avg = days.length ? (days.reduce((a, b) => a + b, 0) / days.length).toFixed(1) : null;

    const shown = teamEntries();
    const totalItems = shown.reduce((a, e) => a + Activity.qty(e), 0);
    document.getElementById("rpSummary").innerHTML =
      `<strong>${totalItems.toLocaleString()}</strong> ${totalItems === 1 ? "item" : "items"} of work logged and ` +
      `<strong>${resolved.length}</strong> ${resolved.length === 1 ? "inquiry" : "inquiries"} resolved` +
      `${resolvedSpecial ? ` (${resolvedSpecial} of them ${escapeHtml(SPECIAL_QUEUE)})` : ""}, ` +
      `${fromLabel} to ${toLabel}${selectedTeam() ? `, ${escapeHtml(selectedTeam())} activities plus shared ones (inquiries show under All teams)` : ""}` +
      `${noAction ? ` (plus ${noAction} no-action ${noAction === 1 ? "email" : "emails"} closed without a response)` : ""}. ` +
      `${avg !== null ? `Inquiries took ${avg} days on average to resolve. ` : ""}` +
      `${plural(open, "inquiry", "inquiries")} open on the Dashboard` +
      `${openSpecial ? ` and ${openSpecial} in ${escapeHtml(SPECIAL_QUEUE)}` : ""} right now` +
      `${typeof Work !== "undefined" ? `, and ${plural(Work.all.filter(Work.isOpen).length, "checklist item")} in progress` : ""}.` +
      ` Checklist items count here when they're marked complete.`;

    if(!groups.length){
      out.innerHTML = `<div class="empty-block"><p>Nothing was logged in this date range.</p></div>`;
      return;
    }

    const colTotals = {};
    let grand = 0;
    const bodyRows = groups.map(g => {
      const head = `<tr class="grp"><th colspan="${people.length + 2}" scope="colgroup">${escapeHtml(g.name)}</th></tr>`;
      const rows = g.rows.map(r => {
        let total = 0;
        const tds = people.map(p => {
          const n = (cells[r.key] && cells[r.key][p]) || 0;
          total += n;
          colTotals[p] = (colTotals[p] || 0) + n;
          return `<td class="num${n ? "" : " zero"}">${n || "–"}</td>`;
        }).join("");
        grand += total;
        return `<tr><th scope="row">${escapeHtml(r.name)}</th>${tds}<td class="num total">${total}</td></tr>`;
      }).join("");
      return head + rows;
    }).join("");

    out.innerHTML = `
      <div class="table-scroll">
        <table class="matrix">
          <thead><tr><th scope="col">Activity</th>${people.map(p => `<th scope="col" class="num">${escapeHtml(p)}</th>`).join("")}<th scope="col" class="num">Total</th></tr></thead>
          <tbody>${bodyRows}</tbody>
          <tfoot><tr><th scope="row">Total</th>${people.map(p => `<td class="num">${colTotals[p] || 0}</td>`).join("")}<td class="num total">${grand}</td></tr></tfoot>
        </table>
      </div>`;
  }

  function rangeSlug(){
    const team = selectedTeam();
    return `${team ? team + "_" : ""}${toDateInput(range.from)}_to_${toDateInput(addDays(range.to, -1))}`;
  }

  function exportSummary(){
    if(!matrix){ toast("Run a report first.", { type: "error" }); return; }
    const { groups, people, cells } = matrix;
    const rows = [["Category", "Activity", ...people, "Total"]];
    groups.forEach(g => g.rows.forEach(r => {
      const counts = people.map(p => (cells[r.key] && cells[r.key][p]) || 0);
      rows.push([g.name, r.name, ...counts, counts.reduce((a, b) => a + b, 0)]);
    }));
    downloadCsv(`activity-summary_${rangeSlug()}.csv`, rows);
    audit(AUDIT_AREAS.data, `Exported activity summary, ${rangeSlug().replace("_to_", " to ")}`);
  }

  function exportEntries(){
    if(!range){ toast("Run a report first.", { type: "error" }); return; }
    const rows = [["Logged At", "Category", "Activity", "Identifier", "Quantity", "Staff Member", "Staff Email"]];
    teamEntries()
      .slice()
      .sort((a, b) => new Date(a[F.loggedAt]) - new Date(b[F.loggedAt]))
      .forEach(e => rows.push([formatDate(e[F.loggedAt]), e[F.category], e[F.activityName], e[F.identifier], Activity.qty(e), e[F.staffName], e[F.staffEmail]]));
    resolvedInRange().forEach(t => rows.push([
      formatDate(t[R.completedOn]), "Inquiries", `${Tickets.queueOf(t) === "special" ? SPECIAL_QUEUE.replace(/ies$/, "y") : "Inquiry"} resolved: ${t.Title || ""}`, t[R.email] || t[R.phoneNumber] || "", 1, t[R.completedBy], ""
    ]));
    downloadCsv(`activity-entries_${rangeSlug()}.csv`, rows);
    audit(AUDIT_AREAS.data, `Exported all activity entries (${plural(rows.length - 1, "row")}), ${rangeSlug().replace("_to_", " to ")}`);
  }

  /* ---------- lookup ---------- */

  async function lookup(){
    const value = document.getElementById("lookupInput").value.trim();
    const out = document.getElementById("lookupResults");
    if(!value){ out.innerHTML = ""; return; }
    out.innerHTML = `<p class="muted">Searching...</p>`;
    try{
      const items = await spGetAll(CONFIG.lists.activityLog,
        filterQuery(`${F.identifier} eq ${odataString(value)}`, "$top=500"));
      items.sort((a, b) => new Date(b[F.loggedAt]) - new Date(a[F.loggedAt]));
      audit(AUDIT_AREAS.data, `Looked up ${value} (${plural(items.length, "record")} found)`);
      if(!items.length){
        out.innerHTML = `<p class="muted">Nothing has been logged for <strong>${escapeHtml(value)}</strong>.</p>`;
        return;
      }
      out.innerHTML = `
        <div class="table-scroll"><table class="plain">
          <thead><tr><th>Logged</th><th>Activity</th><th>Staff member</th><th></th></tr></thead>
          <tbody>${items.map(e => `
            <tr class="static${e[F.voided] ? " voided" : ""}">
              <td>${formatDate(e[F.loggedAt])}</td>
              <td>${escapeHtml(e[F.activityName] || "")}</td>
              <td>${escapeHtml(e[F.staffName] || "")}</td>
              <td>${e[F.voided] ? `Removed by ${escapeHtml(e[F.voidedBy] || "")}` : ""}</td>
            </tr>`).join("")}</tbody>
        </table></div>`;
    }catch(err){
      console.error(err);
      out.innerHTML = `<p class="error-text">Lookup failed: ${escapeHtml(err.message)}</p>`;
    }
  }

  return { init, run };
})();
