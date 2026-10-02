(window.PS_FILE_VERSIONS = window.PS_FILE_VERSIONS || {})["config.js"] = "2026.10.02-2";
/* ============================================================
   config.js — every environment value and SharePoint column name
   lives here. If a live field check shows a different internal
   name, fix it here once and the rest of the app follows.
   ============================================================ */

const CONFIG = {
  appName: "Personnel Support",
  clientId: "0d384d2a-9a9e-46f6-a195-ad63f9e94f06",          // PatrolBidPortal registration
  tenantId: "2935709e-c10c-4809-a302-852d369f8700",
  redirectUri: "https://personnelspecialprojects.github.io/dpd-personnel-support/",
  siteUrl: "https://dallastxgov.sharepoint.com/sites/dpdpersonnelsupport",

  lists: {
    activityTypes: "Activity Types",
    activityLog: "Activity Log",
    requests: "Requests",
    audit: "Audit Log",               // formerly "Request Audit Log" — rename the list in SharePoint
    alertRules: "Alert Rules",
    team: "Team Members",
    workItems: "Work Items",
    workItemLog: "Work Item Log"
  },

  adminContact: "Jared Nielsen",
  pollSeconds: 60,                  // background refresh interval
  recentDays: 14,                   // how far back the Log Activity tab loads
  defaultDuplicateWindowDays: 30,   // used when an activity type leaves the window blank
  maxBatch: 50,                     // most IDs accepted in one submission
  maxCount: 10000,                  // largest number accepted for a "Count" activity in one entry
  undoSeconds: 10,                  // how long the Undo button stays on the confirmation
  closedWorkDays: 30,               // how long completed/cancelled work items stay in the Tracked work tab
  auditDefaultDays: 30              // Audit log tab's default time period
};

/* Activity Types list — the task menu, managed from the Admin tab */
const TYPE_FIELDS = {
  title: "Title",                         // activity name shown on the board
  category: "Category",                   // Single line of text — groups rows on the board
  inputType: "InputType",                 // Choice: Employee Number | Reference | Click Only
  inputLabel: "InputLabel",               // Single line of text — placeholder, e.g. "Employee #"
  duplicateWindowDays: "DuplicateWindowDays", // Number — 0 turns duplicate checking off
  sortOrder: "SortOrder",                 // Number — controls row AND category order
  description: "Description",             // Single line of text (optional helper text)
  active: "Active",                       // Yes/No
  checklistSteps: "ChecklistSteps",       // Multiple lines of text (plain). Empty = simple one-step logging.
  team: "Team"                            // Choice: SRU | PSU | Both. Blank = shown to both teams
};

/* Sub-teams within Personnel Support. Activities belong to one of these, or to
   TEAM_BOTH (shared). The Dashboard's logging panel toggles between them.
   To add a team later, add it here and as a choice on both Team columns. */
const TEAMS = ["SRU", "PSU"];
const TEAM_BOTH = "Both";

const INPUT_TYPES = ["Employee Number", "Reference", "Click Only", "Count"];

/* Activity Log list — one row per completed task */
const LOG_FIELDS = {
  title: "Title",                   // identifier (or activity name for click-only)
  activityTypeId: "ActivityTypeId", // Number — Id of the Activity Types row
  activityName: "ActivityName",     // Single line — snapshot so renames don't rewrite history
  category: "ActivityCategory",     // Single line — snapshot
  identifier: "Identifier",         // Single line — employee # or reference
  staffName: "StaffName",           // Single line
  staffEmail: "StaffEmail",         // Single line
  loggedAt: "LoggedAt",             // Date and time
  quantity: "Quantity",             // Number — how many items this entry represents ("Count" activities). Blank = 1
  voided: "Voided",                 // Yes/No — Undo/Remove sets this; rows are never deleted
  voidedBy: "VoidedBy",             // Single line
  voidedOn: "VoidedOn"              // Date and time
};

/* Requests list — inquiry tickets (same model as the Secondary Employment portal) */
const REQ_FIELDS = {
  title: "Title",
  email: "Email",
  problem: "Problem",
  customerNote: "CustomerNote",           // no longer used (replies are sent from Outlook); kept for old records
  internalNotes: "InternalNotes",
  completed: "Completed",
  completedOn: "CompletedOn",
  completedBy: "CompletedBy",
  status: "Status",
  customerNotified: "CustomerNotified",   // no longer used
  receivedOn: "ReceivedOn",
  entryType: "EntryType",
  source: "Source",
  requesterName: "RequesterName",
  phoneNumber: "PhoneNumber",
  queue: "Queue"                          // Single line of text — blank = Dashboard inquiries; SPECIAL_QUEUE = Special Project Inquiries tab
};

/* Audit Log list — app-wide record of everything people do in the portal */
const AUDIT_FIELDS = {
  area: "Area",                 // Single line — Inquiry, Activity, Tracked work, Admin, Access, Data access
  recordId: "RecordId",         // Number — Id of the inquiry / activity entry / work item / activity type
  details: "Details",           // Multiple lines (plain) — extra context
  staffEmail: "StaffEmail",     // Single line
  ticketId: "TicketId",
  staffMember: "StaffMember",
  logTime: "LogTime",
  action: "Action",
  previousStatus: "PreviousStatus",
  newStatus: "NewStatus",
  internalNotesSnapshot: "InternalNotesSnapshot",
  customerNoteSnapshot: "CustomerNoteSnapshot",
  problemSnapshot: "ProblemSnapshot"
};

const ALERT_FIELDS = {
  color: "Color",
  thresholdValue: "ThresholdValue",
  thresholdUnit: "ThresholdUnit",
  excludeWeekends: "ExcludeWeekends",
  active: "Active"
};

const TEAM_FIELDS = {
  title: "Title",   // email address
  role: "Role",     // Choice: Staff | Admin
  team: "Team"      // Choice: SRU | PSU (optional) — which team's activities the Dashboard opens on
};

/* Work Items list — one row per started checklist item (e.g. one job posting) */
const WORK_FIELDS = {
  title: "Title",                     // "<activity> — <identifier>"
  activityTypeId: "ActivityTypeId",   // Number
  activityName: "ActivityName",       // Single line — snapshot
  category: "ActivityCategory",       // Single line — snapshot
  identifier: "Identifier",           // Single line — employee #, posting #, etc.
  status: "Status",                   // Choice: Open | Completed | Cancelled
  ownerName: "OwnerName",             // Single line — person currently responsible
  ownerEmail: "OwnerEmail",           // Single line
  startedBy: "StartedBy",             // Single line
  startedOn: "StartedOn",             // Date and time
  lastActivityOn: "LastActivityOn",   // Date and time — updated on every change
  closedOn: "ClosedOn",               // Date and time — completed or cancelled
  closedBy: "ClosedBy",               // Single line
  steps: "Steps",                     // Multiple lines of text (plain) — this item's own checklist (JSON)
  completionLogId: "CompletionLogId"  // Number — Activity Log entry written on completion
};

/* Work Item Log list — history and notes for each work item */
const WORKLOG_FIELDS = {
  title: "Title",
  workItemId: "WorkItemId",           // Number
  action: "Action",                   // Multiple lines of text (plain)
  staffName: "StaffName",             // Single line
  staffEmail: "StaffEmail",           // Single line
  loggedAt: "LoggedAt"                // Date and time
};

/* Status for inquiries closed without a response (FYI emails to the DL).
   Must match the choice added to the Requests list's Status column exactly.
   Senders are asked to put this phrase in the subject line. */
const NO_ACTION_STATUS = "No Action Required";

/* Inquiries set aside for a special project (e.g. a software transition). They get
   their own tab so they don't crowd the Dashboard. Stored in the Requests list's
   Queue column. */
const SPECIAL_QUEUE = "Special Project Inquiries";

const AUDIT_AREAS = {
  inquiry: "Inquiry",
  activity: "Activity",
  work: "Tracked work",
  admin: "Admin",
  access: "Access",
  data: "Data access"
};

const PRESET_COLORS = [
  { name: "Yellow", hex: "#F5C518" },
  { name: "Orange", hex: "#E67E22" },
  { name: "Red", hex: "#C0392B" },
  { name: "Purple", hex: "#8E44AD" },
  { name: "Blue", hex: "#2E86DE" },
  { name: "Gray", hex: "#7F8C8D" }
];
