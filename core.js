(window.PS_FILE_VERSIONS = window.PS_FILE_VERSIONS || {})["core.js"] = "2026.10.02-3";
/* ============================================================
   core.js — sign-in, SharePoint REST helpers, shared utilities,
   tab routing, polling, and boot. Loaded before the feature files.
   ============================================================ */

const App = {
  user: null,          // { name, username, key }  key = email local part
  role: "Staff",
  isAdmin: false,
  token: null,
  msal: null,
  pollId: null,
  digest: null,
  digestExpires: 0,
  team: []
};

const ViewHooks = {};  // viewId -> [functions] called when that tab is shown
function onView(viewId, fn){ (ViewHooks[viewId] = ViewHooks[viewId] || []).push(fn); }

/* ---------------- Auth ---------------- */

function getMsal(){
  if(App.msal) return App.msal;
  if(typeof msal === "undefined") return null;   // CDN fallback may still be loading
  App.msal = new msal.PublicClientApplication({
    auth: {
      clientId: CONFIG.clientId,
      authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
      redirectUri: CONFIG.redirectUri
    },
    cache: { cacheLocation: "sessionStorage" }
  });
  return App.msal;
}

function spScopes(){
  return [`${new URL(CONFIG.siteUrl).origin}/AllSites.Write`];
}

async function acquireSpToken(account, allowPopup){
  const m = getMsal();
  try{
    const r = await m.acquireTokenSilent({ scopes: spScopes(), account });
    return r.accessToken;
  }catch(err){
    if(!allowPopup) throw err;
    const r = await m.acquireTokenPopup({ scopes: spScopes(), account });
    return r.accessToken;
  }
}

async function signIn(){
  const m = getMsal();
  if(!m){
    alert("The Microsoft sign-in library didn't load. Check your connection, or ask IT whether alcdn.msauth.net and alcdn.msftauth.net are blocked on this network, then reload the page.");
    return;
  }
  try{
    // Graph token (identity) and SharePoint token (data) are always requested separately.
    const login = await m.loginPopup({ scopes: ["User.Read"] });
    m.setActiveAccount(login.account);
    App.token = await acquireSpToken(login.account, true);
    await startSession(login.account);
  }catch(err){
    console.error(err);
    showGate("signInScreen");
    toast("Sign-in didn't finish. If a pop-up blocker is on, allow pop-ups for this page and try again.", { type: "error" });
  }
}

async function tryResumeSession(){
  const m = getMsal();
  if(!m) return;
  const accounts = m.getAllAccounts();
  if(!accounts.length) return;
  const account = m.getActiveAccount() || accounts[0];
  try{
    m.setActiveAccount(account);
    showGate("loadingScreen");
    App.token = await acquireSpToken(account, false);
    await startSession(account);
  }catch(err){
    console.warn("Silent sign-in didn't work; showing the sign-in button.", err);
    showGate("signInScreen");
  }
}

function signOut(){
  stopPolling();
  const m = getMsal();
  if(m) m.logoutPopup().catch(err => console.error(err));
}

async function startSession(account){
  App.user = {
    name: account.name || account.username,
    username: (account.username || "").toLowerCase(),
    key: localPart(account.username)
  };
  showGate("loadingScreen");

  let me;
  try{
    me = await loadTeamAndFindMe();
  }catch(err){
    console.error(err);
    showGate("notAuthorizedScreen",
      `The Team Members list couldn't be read (${err.message}). If the site was just set up, confirm the list exists and has a Role column.`);
    return;
  }
  if(!me){
    // Record the attempt if SharePoint lets us; someone off the roster may not have write access.
    audit(AUDIT_AREAS.access, "Access denied: not on the team roster", { silent: true });
    showGate("notAuthorizedScreen",
      `Your account isn't on the team roster for this portal. Contact ${CONFIG.adminContact} if you should have access.`);
    return;
  }

  App.role = me[TEAM_FIELDS.role] || "Staff";
  App.isAdmin = App.role === "Admin";
  App.myTeam = TEAMS.includes(me[TEAM_FIELDS.team]) ? me[TEAM_FIELDS.team] : null;
  document.querySelectorAll("[data-admin-only]").forEach(el => { el.hidden = !App.isAdmin; });

  document.getElementById("userBox").innerHTML = `
    <div class="user-name">${escapeHtml(App.user.name)}${App.isAdmin ? ' <span class="role-tag">Admin</span>' : ""}</div>
    <div class="user-meta"><span id="syncLabel"></span><span class="ver" title="Release version">v${escapeHtml(window.PS_PAGE_VERSION || "?")}</span><button id="signOutBtn" type="button">Sign out</button></div>`;
  document.getElementById("signOutBtn").addEventListener("click", signOut);

  showGate(null);
  // Once per browser session, so page reloads don't flood the log.
  let loggedThisSession = false;
  try{ loggedThisSession = sessionStorage.getItem("ps-signin-logged") === App.user.key; }catch(_){}
  if(!loggedThisSession){
    audit(AUDIT_AREAS.access, `Signed in${App.isAdmin ? " (admin)" : ""}`);
    try{ sessionStorage.setItem("ps-signin-logged", App.user.key); }catch(_){}
  }

  // Activity board first — it's the screen people land on.
  try{
    await Activity.init();
  }catch(err){
    console.error(err);
    toast(`Activities couldn't load: ${err.message}`, { type: "error", duration: 0 });
  }
  const results = await Promise.allSettled([Work.init(), Tickets.init()]);
  results.forEach(r => { if(r.status === "rejected") console.error(r.reason); });
  // Each module starts independently: one failing piece must not lock everyone out.
  const safeInit = (name, fn) => { try{ fn(); }catch(err){ console.error(`${name} failed to start:`, err); } };
  safeInit("Reports", () => Reports.init());
  safeInit("Admin", () => Admin.init());
  safeInit("Audit", () => Audit.init());
  if(App.isAdmin) safeInit("Setup check", () => SetupCheck.init());
  markSynced();
  startPolling();
}

async function loadTeamAndFindMe(){
  // No $select: a missing optional column (e.g. Team) must never block sign-in.
  App.team = await spGetAll(CONFIG.lists.team, "$top=500");
  if(!App.user.key) return null;
  return App.team.find(m => localPart(m[TEAM_FIELDS.title]) === App.user.key) || null;
}

/* ---------------- Polling ---------------- */

async function poll(){
  if(document.hidden) return;
  const m = getMsal();
  const account = m && m.getActiveAccount();
  if(!account) return;
  try{
    App.token = await acquireSpToken(account, false);
  }catch(err){
    console.warn("Token refresh failed; will retry next cycle.", err);
    return;
  }
  const results = await Promise.allSettled([Activity.refresh(), Work.refresh(), Tickets.refresh(), Audit.refresh()]);
  results.forEach(r => { if(r.status === "rejected") console.warn("Refresh problem:", r.reason); });
  markSynced();
}

function startPolling(){
  stopPolling();
  App.pollId = setInterval(poll, CONFIG.pollSeconds * 1000);
}

function stopPolling(){
  if(App.pollId){ clearInterval(App.pollId); App.pollId = null; }
}

function markSynced(){
  const el = document.getElementById("syncLabel");
  if(el) el.textContent = `Updated ${formatTime(new Date())}`;
}

/* ---------------- SharePoint REST ---------------- */

class SpError extends Error{
  constructor(status, body, url){
    super(parseSpMessage(body) || `SharePoint returned ${status}`);
    this.status = status;
    this.body = body;
    this.url = url;
  }
}

function parseSpMessage(body){
  try{
    const j = JSON.parse(body);
    const e = j["odata.error"] || j.error;
    if(!e) return null;
    return (e.message && (e.message.value || e.message)) || null;
  }catch(_){
    return null;
  }
}

async function spRequest(url, options = {}){
  const send = () => fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${App.token}`,
      Accept: "application/json;odata=nometadata",
      ...(options.headers || {})
    }
  });
  let res = await send();
  if(res.status === 401){
    // token expired mid-session — refresh silently and retry once
    try{
      App.token = await acquireSpToken(getMsal().getActiveAccount(), false);
      res = await send();
    }catch(_){ /* fall through to error */ }
  }
  if(!res.ok){
    throw new SpError(res.status, await res.text(), url);
  }
  return res;
}

async function getDigest(force){
  if(!force && App.digest && Date.now() < App.digestExpires) return App.digest;
  const res = await spRequest(`${CONFIG.siteUrl}/_api/contextinfo`, { method: "POST" });
  const d = await res.json();
  App.digest = d.FormDigestValue;
  App.digestExpires = Date.now() + ((d.FormDigestTimeoutSeconds || 1800) - 120) * 1000;
  return App.digest;
}

async function spWrite(url, method, body){
  const attempt = async (forceDigest) => {
    const digest = await getDigest(forceDigest);
    const headers = {
      "X-RequestDigest": digest,
      "Content-Type": "application/json;odata=nometadata"
    };
    if(method === "MERGE" || method === "DELETE"){
      headers["IF-MATCH"] = "*";
      headers["X-HTTP-Method"] = method;
    }
    return spRequest(url, {
      method: "POST",
      headers,
      body: body === undefined ? undefined : JSON.stringify(body)
    });
  };
  try{
    return await attempt(false);
  }catch(err){
    if(err.status === 403) return attempt(true);   // stale digest — get a new one and retry once
    throw err;
  }
}

function listItemsUrl(listTitle, suffix = ""){
  const safe = encodeURIComponent(listTitle.replace(/'/g, "''"));
  return `${CONFIG.siteUrl}/_api/web/lists/getbytitle('${safe}')/items${suffix}`;
}

/* GET every page of results (follows odata.nextLink). */
async function spGetAll(listTitle, query = ""){
  let url = listItemsUrl(listTitle) + (query ? `?${query}` : "");
  const out = [];
  let pages = 0;
  while(url && pages++ < 100){
    const res = await spRequest(url);
    const data = await res.json();
    out.push(...(data.value || []));
    url = data["odata.nextLink"] || data["@odata.nextLink"] || null;
  }
  return out;
}

async function spGetItem(listTitle, id){
  const res = await spRequest(listItemsUrl(listTitle, `(${id})`));
  return res.json();
}

async function spCreate(listTitle, body){
  const res = await spWrite(listItemsUrl(listTitle), "POST", body);
  return res.json();
}

async function spUpdate(listTitle, id, body){
  await spWrite(listItemsUrl(listTitle, `(${id})`), "MERGE", body);
}

async function spDelete(listTitle, id){
  await spWrite(listItemsUrl(listTitle, `(${id})`), "DELETE");
}

/* OData helpers — build filter text, then encode the whole expression once. */
function odataString(value){
  return `'${String(value).replace(/'/g, "''")}'`;
}
function odataDate(date){
  return `datetime'${new Date(date).toISOString()}'`;
}
function filterQuery(expr, extra = ""){
  return `$filter=${encodeURIComponent(expr)}${extra ? "&" + extra : ""}`;
}

/* ---------------- Audit log (app-wide) ---------------- */

let auditErrorShownAt = 0;

/* Write one Audit Log row. `entry` may include any AUDIT_FIELDS columns.
   Never throws: an audit failure must not undo or block the user's actual change. */
async function writeAudit(area, entry = {}, opts = {}){
  const A = AUDIT_FIELDS;
  const body = {
    [A.area]: area,
    [A.staffMember]: App.user ? App.user.name : "",
    [A.staffEmail]: App.user ? App.user.username : "",
    [A.logTime]: new Date().toISOString(),
    ...entry
  };
  body.Title = String(entry.Title || `${area}: ${body[A.action] || ""}`).slice(0, 255);
  try{
    const row = await spCreate(CONFIG.lists.audit, body);
    if(typeof Audit !== "undefined") Audit.add(row);
    return row;
  }catch(err){
    console.error("Audit log entry failed to save:", err, body);
    if(!opts.silent && Date.now() - auditErrorShownAt > 60000){
      auditErrorShownAt = Date.now();
      toast(`Your change saved, but its audit log entry didn't (${err.message}). ` +
        `Let ${CONFIG.adminContact} know; the setup check in Admin will show what's wrong.`,
        { type: "error", duration: 0 });
    }
    return null;
  }
}

/* Shorthand: audit(AUDIT_AREAS.activity, "Logged ...", { recordId, details }) */
function audit(area, action, extra = {}){
  const A = AUDIT_FIELDS;
  const entry = { [A.action]: action };
  if(extra.recordId !== undefined && extra.recordId !== null) entry[A.recordId] = extra.recordId;
  if(extra.details) entry[A.details] = String(extra.details);
  if(extra.title) entry.Title = extra.title;
  return writeAudit(area, entry, extra);
}

/* ---------------- Utilities ---------------- */

function escapeHtml(s){
  return String(s ?? "").replace(/[&<>"']/g, c => (
    { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]
  ));
}

function localPart(email){
  return String(email || "").split("@")[0].trim().toLowerCase();
}

/* "jane.doe2@dallaspolice.gov" -> "Jane Doe". Used where only an email is on file. */
function nameFromEmail(email){
  const parts = localPart(email).replace(/[0-9]+$/, "").split(/[._-]+/).filter(Boolean);
  return parts.map(s => s.charAt(0).toUpperCase() + s.slice(1)).join(" ") || String(email || "");
}

/* Display name for a roster member: the signed-in user's real name, otherwise derived. */
function teamMemberName(email){
  if(App.user && localPart(email) === App.user.key) return App.user.name;
  return nameFromEmail(email);
}

function daysSince(d){
  if(!d) return null;
  return Math.max(0, Math.floor((Date.now() - new Date(d).getTime()) / 86400000));
}

function formatDate(d){
  if(!d) return "";
  const dt = new Date(d);
  if(isNaN(dt)) return "";
  return dt.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" }) +
    " " + formatTime(dt);
}

function formatTime(d){
  return new Date(d).toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" });
}

function formatDay(d){
  const dt = new Date(d);
  const today = startOfToday();
  const day = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate());
  const diff = Math.round((today - day) / 86400000);
  if(diff === 0) return "Today";
  if(diff === 1) return "Yesterday";
  return dt.toLocaleDateString("en-US", { weekday: "short", month: "short", day: "numeric" });
}

function startOfToday(){
  const n = new Date();
  return new Date(n.getFullYear(), n.getMonth(), n.getDate());
}

function addDays(date, days){
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

function toDateInput(date){
  const d = new Date(date);
  const p = n => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

function fromDateInput(value){
  const [y, m, d] = String(value).split("-").map(Number);
  if(!y || !m || !d) return null;
  return new Date(y, m - 1, d);
}

function plural(n, one, many){
  return `${n} ${n === 1 ? one : (many || one + "s")}`;
}

function downloadCsv(filename, rows){
  const csv = rows
    .map(row => row.map(v => `"${String(v ?? "").replace(/"/g, '""')}"`).join(","))
    .join("\n");
  const blob = new Blob([csv], { type: "text/csv" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
}

/* ---------------- Toasts ---------------- */

function toast(message, opts = {}){
  const { type = "info", actionLabel, onAction, duration = 5000 } = opts;
  const host = document.getElementById("toastHost");
  const el = document.createElement("div");
  el.className = `toast toast-${type}`;
  el.setAttribute("role", type === "error" ? "alert" : "status");
  el.innerHTML = `<span class="toast-msg">${escapeHtml(message)}</span>`;
  const close = () => { el.classList.add("leaving"); setTimeout(() => el.remove(), 200); };
  if(actionLabel && onAction){
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "toast-action";
    btn.textContent = actionLabel;
    btn.addEventListener("click", () => { btn.disabled = true; close(); onAction(); });
    el.appendChild(btn);
  }
  const x = document.createElement("button");
  x.type = "button";
  x.className = "toast-close";
  x.setAttribute("aria-label", "Dismiss");
  x.innerHTML = "&times;";
  x.addEventListener("click", close);
  el.appendChild(x);
  host.appendChild(el);
  if(duration > 0) setTimeout(close, duration);
  return el;
}

/* ---------------- Screens, tabs, overlays ---------------- */

function showGate(id, message){
  ["signInScreen", "loadingScreen", "notAuthorizedScreen"].forEach(g => {
    document.getElementById(g).hidden = g !== id;
  });
  if(id === "notAuthorizedScreen" && message){
    document.getElementById("notAuthMsg").textContent = message;
  }
  document.getElementById("mainApp").hidden = id !== null;
}

function switchTab(viewId){
  document.querySelectorAll(".tab").forEach(t => {
    const on = t.dataset.view === viewId;
    t.classList.toggle("active", on);
    t.setAttribute("aria-selected", on ? "true" : "false");
  });
  document.querySelectorAll(".view").forEach(v => v.classList.toggle("active", v.id === viewId));
  (ViewHooks[viewId] || []).forEach(fn => { try{ fn(); }catch(err){ console.error(err); } });
}

function openOverlay(id){ document.getElementById(id).classList.add("active"); }
function closeOverlay(id){ document.getElementById(id).classList.remove("active"); }

/* ---------------- Boot ---------------- */

/* ---------------- Version check ---------------- */

const PS_EXPECTED_FILES = ["config.js", "core.js", "activity.js", "work.js", "tickets.js",
  "reports.js", "admin.js", "audit.js", "setup-check.js"];

/* Files whose release stamp doesn't match the page (old cached copy, or not uploaded). */
function outdatedFiles(){
  const page = window.PS_PAGE_VERSION;
  const have = window.PS_FILE_VERSIONS || {};
  if(!page) return [];
  return PS_EXPECTED_FILES.filter(f => have[f] !== page);
}

function checkVersions(){
  const bad = outdatedFiles();
  if(!bad.length) return;
  console.warn("Out-of-date files:", bad, "page version", window.PS_PAGE_VERSION, window.PS_FILE_VERSIONS);
  const bar = document.getElementById("versionBanner");
  bar.innerHTML = `<strong>Part of the portal is out of date:</strong> ${bad.map(escapeHtml).join(", ")}. ` +
    `Press <kbd>Ctrl</kbd>+<kbd>F5</kbd> to reload. If this message stays, ` +
    `let ${escapeHtml(CONFIG.adminContact)} know: the latest ${bad.length === 1 ? "copy of that file needs" : "copies of those files need"} to be uploaded to GitHub.`;
  bar.hidden = false;
}

function boot(){
  checkVersions();
  document.getElementById("signInBtn").addEventListener("click", signIn);
  document.getElementById("notAuthSignOutBtn").addEventListener("click", signOut);
  document.querySelectorAll(".tab").forEach(t => t.addEventListener("click", () => switchTab(t.dataset.view)));
  document.addEventListener("keydown", e => {
    if(e.key !== "Escape") return;
    const open = [...document.querySelectorAll(".modal-overlay.active")].pop();
    if(open){
      const closer = open.querySelector("[data-close]");
      if(closer) closer.click(); else open.classList.remove("active");
    }
  });
  document.addEventListener("visibilitychange", () => { if(!document.hidden && App.user) poll(); });
  showGate("signInScreen");
  tryResumeSession();
}
