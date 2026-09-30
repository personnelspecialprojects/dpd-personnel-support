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
    requestAudit: "Request Audit Log",
    alertRules: "Alert Rules",
    team: "Team Members"
  },

  adminContact: "Jared Nielsen",
  pollSeconds: 60,                  // background refresh interval
  recentDays: 14,                   // how far back the Log Activity tab loads
  defaultDuplicateWindowDays: 30,   // used when an activity type leaves the window blank
  maxBatch: 50,                     // most IDs accepted in one submission
  undoSeconds: 10                   // how long the Undo button stays on the confirmation
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
  active: "Active"                        // Yes/No
};

const INPUT_TYPES = ["Employee Number", "Reference", "Click Only"];

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
  voided: "Voided",                 // Yes/No — Undo/Remove sets this; rows are never deleted
  voidedBy: "VoidedBy",             // Single line
  voidedOn: "VoidedOn"              // Date and time
};

/* Requests list — inquiry tickets (same model as the Secondary Employment portal) */
const REQ_FIELDS = {
  title: "Title",
  email: "Email",
  problem: "Problem",
  customerNote: "CustomerNote",
  internalNotes: "InternalNotes",
  completed: "Completed",
  completedOn: "CompletedOn",
  completedBy: "CompletedBy",
  status: "Status",
  customerNotified: "CustomerNotified",
  receivedOn: "ReceivedOn",
  entryType: "EntryType",
  source: "Source",
  requesterName: "RequesterName",
  phoneNumber: "PhoneNumber"
};

const AUDIT_FIELDS = {
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
  role: "Role"      // Choice: Staff | Admin
};

const PRESET_COLORS = [
  { name: "Yellow", hex: "#F5C518" },
  { name: "Orange", hex: "#E67E22" },
  { name: "Red", hex: "#C0392B" },
  { name: "Purple", hex: "#8E44AD" },
  { name: "Blue", hex: "#2E86DE" },
  { name: "Gray", hex: "#7F8C8D" }
];
