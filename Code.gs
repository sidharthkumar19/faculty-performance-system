// =============================================================================
// VMRF-DU Institutional Monitoring System — Code.gs
// =============================================================================
// SETUP:
//  1. Open Google Sheet → Extensions → Apps Script
//  2. Paste this into Code.gs (replace all existing content)
//  3. Click + → New HTML file → name it exactly "Index" → paste Index.html
//  4. Run initializeSystem() once → approve permissions
//  5. Deploy → New Deployment → Web App
//     Execute as: Me | Access: Anyone (or your org)
//  6. Copy & share the Web App URL
// =============================================================================

var SH = {
  STAFF:       'Staff_Master',
  FACULTY:     'Faculty_Master',
  SUBMISSION:  'Weekly_Submission',
  TIMESHEET:   'Timesheet_Entries',
  SELF_ASSESS: 'Self_Assessment',
  HOD:         'HOD_Remarks',
  HOI:         'HOI_Remarks',
  IMO:         'IMO_Monitoring',
  NOTIF:       'Notifications',
  PP_WORKS:    'PendingPriorityWork',
  HOD_SUB:     'HOD_Submission',
  HOD_TS:      'HOD_Timesheet',
  HOD_SA:      'HOD_SelfAssess',
  HOD_REVIEW:  'HOD_Review',
  HOD_IMO:     'HOD_IMO_Review',
  HOI_WEEKLY:  'HOI_WeeklyMeeting',
  PROFILES:    'FacultyProfiles'
};

var SCHEMA = {
  Staff_Master:      ['StaffID','StaffName','Email','Role','Department','Campus','Institution','PasswordHash','GoogleEmail','Status','Designation','Phone','ProfileCompleted'],
  Faculty_Master:    ['FacultyName','Email','Department','Campus','Institution','Designation','PasswordHash','GoogleEmail','Status','ProfileCompleted'],
  Weekly_Submission: ['SubmissionID','FacultyEmail','AcademicYearSemester','ReportingFrom','ReportingTo','Declaration','SubmittedDateTime'],
  Timesheet_Entries: ['SubmissionID','Day','TimeSlot','ActivityType','ActivityDetails','AttachmentURL','AttachmentName'],
  Self_Assessment:   ['SubmissionID','OutcomeOfWeek','TargetPlanNextWeek'],
  HOD_Remarks:       ['SubmissionID','HOD_Remark','HOD_Status','HOD_DateTime'],
  HOI_Remarks:       ['SubmissionID','HOI_Remark','HOI_Status','HOI_DateTime'],
  IMO_Monitoring:    ['SubmissionID','IMO_Remark','IMO_Status','IMO_DateTime'],
  // NotifID | ForRole | Type | Title | Body | SubmissionID | FacultyName | IsRead | CreatedAt
  Notifications:     ['NotifID','ForRole','Type','Title','Body','SubmissionID','FacultyName','IsRead','CreatedAt'],
  PendingPriorityWork: ['PPID','FacultyEmail','WeekOf','Area','Task','ResponsiblePerson','DateOfCommencement','TargetDate','Status','CreatedAt'],
  HOD_Submission:  ['SubmissionID','HOD_ID','AcademicYearSemester','ReportingFrom','ReportingTo','Declaration','SubmittedDateTime'],
  HOD_Timesheet:   ['SubmissionID','Day','TimeSlot','ActivityType','ActivityDetails','AttachmentURL','AttachmentName'],
  HOD_SelfAssess:  ['SubmissionID','Tasks','TargetPlanNextWeek'],
  HOD_Review:      ['SubmissionID','HOI_Remark','HOI_Status','HOI_DateTime'],
  HOD_IMO_Review:  ['SubmissionID','IMO_Remark','IMO_Status','IMO_DateTime'],
  HOI_WeeklyMeeting: ['MeetingID','HOI_Email','MeetingDate','Institution','WeekStart','WeekEnd','DiscussionData','SubmittedAt'],
  FacultyProfiles:   ['Email','FacultyName','EmployeeID','Designation','Department','Institution','Campus',
                      'Qualification','Specialization','Experience','Phone','DateOfJoining',
                      'ResearchAreas','PublicationsCount','Certifications','LinkedinOrcid','Bio','LastUpdated']
};

// Activity types kept for backward compatibility with old submissions
// New submissions use 3-level structure: Category > Sub-Category > Activity
var ACTIVITY_TYPES = [
  'Academic Planning & Delivery',
  'Research & Innovation',
  'Professional Development',
  'Student Support Services',
  'Other'
];

var TIME_SLOTS = [
  '8:40 AM – 9:40 AM',
  '9:40 AM – 10:40 AM',
  '10:50 AM – 11:50 AM',
  '12:30 PM – 1:30 PM',
  '1:30 PM – 2:25 PM',
  '2:35 PM – 3:30 PM'
];

// Per-institution time slots for Chennai Campus institutions.
// Key = institution code (matches INSTITUTION_HIERARCHY codes).
// Puducherry campus institutions fall back to TIME_SLOTS (AVIT default).
var INSTITUTION_TIME_SLOTS = {
  'AVIT':     ['8:40 AM – 9:40 AM',  '9:40 AM – 10:40 AM',  '10:50 AM – 11:50 AM', '12:30 PM – 1:30 PM', '1:30 PM – 2:25 PM', '2:35 PM – 3:30 PM'],
  'AVIT-DoM': ['8:40 AM – 9:40 AM',  '9:40 AM – 10:40 AM',  '10:50 AM – 11:50 AM', '12:30 PM – 1:30 PM', '1:30 PM – 2:25 PM', '2:35 PM – 3:30 PM'],
  'SAS':      ['9:00 AM – 9:50 AM',  '9:51 AM – 10:40 AM',  '10:46 AM – 11:35 AM', '11:36 AM – 12:25 PM','12:56 PM – 1:45 PM', '1:46 PM – 2:35 PM'],
  'SAHS':     ['8:30 AM – 9:30 AM',  '9:30 AM – 10:30 AM',  '10:45 AM – 11:45 AM', '11:45 AM – 12:45 PM','1:30 PM – 2:30 PM',  '2:30 PM – 3:30 PM'],
  'VMLS':     ['8:45 AM – 9:45 AM',  '9:45 AM – 10:45 AM',  '11:00 AM – 12:00 PM', '12:00 PM – 1:00 PM', '1:45 PM – 2:45 PM',  '2:45 PM – 3:30 PM'],
  'VSEP':     ['9:00 AM – 9:50 AM',  '9:51 AM – 10:40 AM',  '10:51 AM – 11:40 AM', '11:41 AM – 12:30 PM','1:16 PM – 2:05 PM',  '2:05 PM – 2:55 PM'],
  'VSHS':     ['8:45 AM – 9:45 AM',  '9:45 AM – 10:45 AM',  '11:00 AM – 12:00 PM', '12:00 PM – 1:00 PM', '1:30 PM – 2:30 PM',  '2:30 PM – 3:30 PM']
};

// ═══════════════════════════════════════════════════════════════════════════
// INSTITUTION HIERARCHY — single source of truth
// ═══════════════════════════════════════════════════════════════════════════
// Structure: Campus (full name) → { code, institutions: { instName → { code, departments:[] } } }
// Campus codes (VMCC, VMPC) match the IPM portal.
// ⚠️  REVIEW DEPARTMENT LISTS BELOW — edit to match your actual institutional structure.
// ═══════════════════════════════════════════════════════════════════════════
var INSTITUTION_HIERARCHY = {
  "Vinayaka Mission's Chennai Campus": {
    code: 'VMCC',
    institutions: {
      'Aarupadai Veedu Institute of Technology (AVIT)': {
        code: 'AVIT',
        departments: [
          'Biotechnology',
          'Biomedical Engineering',
          'Civil Engineering',
          'Computer Science and Engineering',
          'Electronics and Communication Engineering',
          'Electrical and Electronics Engineering',
          'Mechanical Engineering',
          'Humanities & Sciences'
        ]
      },
      'AVIT - Department of Management (AVIT-DoM)': {
        code: 'AVIT-DoM',
        departments: ['Management Studies']
      },
      'School of Arts and Science (SAS)': {
        code: 'SAS',
        departments: [
          'English','Tamil','Mathematics','Physics','Chemistry',
          'Computer Science','Commerce','Business Administration',
          'Biotechnology','Visual Communication'
        ]
      },
      'School of Allied Health Sciences - Chennai Campus (SAHS)': {
        code: 'SAHS',
        noHod: true,
        departments: [
          'Medical Laboratory Technology',
          'Cardiac Care Technology',
          'Operation Theatre Technology',
          'Dialysis Technology',
          'Imaging Technology',
          'Optometry',
          'Anaesthesia Technology'
        ]
      },
      "Vinayaka Mission's Law School (VMLS)": {
        code: 'VMLS',
        noHod: true,
        departments: ['Law']
      },
      "Vinayaka Mission's School of Economics and Public Policy (VSEP)": {
        code: 'VSEP',
        noHod: true,
        departments: ['Economics','Public Policy']
      },
      "Vinayaka Mission's School of Health Systems (VSHS)": {
        code: 'VSHS',
        noHod: true,
        departments: ['Hospital Administration','Health Systems Management']
      }
    }
  },
  "Vinayaka Mission's Puducherry Campus": {
    code: 'VMPC',
    institutions: {
      'Aarupadai Veedu Medical College & Hospital (AVMC)': {
        code: 'AVMC',
        departments: [
          'Anatomy','Physiology','Biochemistry','Pharmacology',
          'Pathology','Microbiology','Community Medicine','Forensic Medicine',
          'General Medicine','General Surgery','Obstetrics & Gynaecology',
          'Paediatrics','Orthopaedics','Ophthalmology','ENT',
          'Dermatology','Psychiatry','Radiology','Anaesthesiology'
        ]
      },
      "Vinayaka Mission's College of Nursing (VMCON)": {
        code: 'VMCON',
        departments: [
          'Medical-Surgical Nursing',
          'Community Health Nursing',
          'Obstetrics & Gynaecological Nursing',
          'Paediatric Nursing',
          'Mental Health Nursing',
          'Nursing Foundation'
        ]
      },
      'School of Rehabilitation and Behavioral Sciences (SRBS)': {
        code: 'SRBS',
        departments: [
          'Pharmaceutics','Pharmacology','Pharmacognosy',
          'Pharmaceutical Chemistry','Pharmacy Practice',
          'Biotechnology','Microbiology'
        ]
      },
      'School of Physiotherapy (SPT)': {
        code: 'SPT',
        departments: [
          'Orthopaedic Physiotherapy',
          'Neurological Physiotherapy',
          'Cardiorespiratory Physiotherapy',
          'Sports Physiotherapy',
          'Community Physiotherapy'
        ]
      },
      'School of Allied Health Sciences - Puducherry Campus (SAHS-PC)': {
        code: 'SAHS-MC',
        departments: [
          'Medical Laboratory Technology',
          'Cardiac Care Technology',
          'Operation Theatre Technology',
          'Dialysis Technology',
          'Imaging Technology',
          'Anaesthesia Technology',
          'Optometry'
        ]
      }
    }
  }
};

// Derived flat arrays — keep existing code working unchanged
var CAMPUSES     = Object.keys(INSTITUTION_HIERARCHY);
var INSTITUTIONS = (function(){
  var out = [];
  CAMPUSES.forEach(function(c){
    Object.keys(INSTITUTION_HIERARCHY[c].institutions).forEach(function(i){ out.push(i); });
  });
  return out;
})();
var DEPARTMENTS  = (function(){
  var seen = {}, out = [];
  CAMPUSES.forEach(function(c){
    var insts = INSTITUTION_HIERARCHY[c].institutions;
    Object.keys(insts).forEach(function(i){
      (insts[i].departments||[]).forEach(function(d){
        if (!seen[d]) { seen[d]=1; out.push(d); }
      });
    });
  });
  return out;
})();
var UG_DEPARTMENTS = DEPARTMENTS.slice();
var PG_DEPARTMENTS = [];

// Build a client-friendly hierarchy object for getConfig()

var DESIGNATIONS   = ['Professor','Associate Professor','Assistant Professor','Lecturer','Other'];
var ACADEMIC_YEARS = ['2026–2027 (Odd Semester)','2026–2027 (Even Semester)','2025–2026 (Odd Semester)','2025–2026 (Even Semester)'];
var DAYS           = ['Day 1 (Mon)','Day 2 (Tue)','Day 3 (Wed)','Day 4 (Thu)','Day 5 (Fri)','Day 6 (Sat)'];

// ─── WEB APP ENTRY ────────────────────────────────────────────────────────────
function doGet() {
  var t = HtmlService.createTemplateFromFile('Index');
  return t.evaluate()
    .setTitle('VMRF-DU Institutional Monitoring System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Used by <?!= include('FileName') ?> in Index.html to inline other HTML files
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ─── ONE-TIME CLEANUP: Remove HR sheets ──────────────────────────────────────
// ONE-TIME UTILITY: Run from Apps Script → removeHRSheets, then delete this function
// Then delete this function after running.
function removeHRSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var toDelete = ['Leave_Requests', 'HR_Announcements'];
  var removed = [], notFound = [];
  toDelete.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (sh) { ss.deleteSheet(sh); removed.push(name); }
    else notFound.push(name);
  });
  var msg = '';
  if (removed.length)   msg += 'Deleted: ' + removed.join(', ') + '.\n';
  if (notFound.length)  msg += 'Not found (already removed): ' + notFound.join(', ') + '.';
  SpreadsheetApp.getUi().alert('HR Sheet Cleanup\n\n' + (msg || 'Nothing to remove.'));
}