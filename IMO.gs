// =============================================================================
// VMRF-DU Institutional Monitoring System — IMO Trigger & Archive Services
// =============================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('VMRF IMO')
    .addItem('Setup / Re-initialise System', 'initializeSystem')
      .addItem('Seed HOD Department Accounts', 'seedHODAccounts')
      .addItem('Reset HOD Passwords', 'resetHODPasswords')
      .addItem('Remove HOD Accounts (No-HOD Institutions)', 'removeNoHodAccounts')
      .addItem('Reset ALL Passwords (except IMO)', 'resetAllPasswordsExceptIMO')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🔒 Security Upgrade (Path A)')
      .addItem('1. Migrate Passwords to Salted Hash', 'migratePasswordsToSecureFormat')
      .addItem('2. Lock Down Existing Attachments', 'lockdownExistingAttachments'))
    .addItem('Apply Status Dropdowns', 'applyStatusDropdowns')
    .addItem('Open Web App', 'openWebApp')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Weekly KPI Archive')
      .addItem('Install Sunday 23:30 Trigger', 'menuInstallWeeklyArchiveTrigger')
      .addItem('Snapshot Last Week Now (manual)', 'menuRunWeeklyArchiveNow')
      .addSeparator()
      .addItem('Check Archive Status', 'menuCheckWeeklyArchiveStatus')
      .addItem('Remove Trigger', 'menuRemoveWeeklyArchiveTrigger'))
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Institutional Monitoring (IPM)')
      .addItem('Setup IPM Database (first time)', 'ipmSetup')
      .addItem('Sync / Re-init IPM Schema', 'ipmInitialize'))
    .addToUi();
}

/* ──────────────────────────────────────────────────────────────────────────
   WEEKLY KPI ARCHIVE — menu wrappers
   These thin wrappers exist so an admin can install / run / inspect the
   weekly archive system from the spreadsheet menu without opening the Apps
   Script editor. The underlying logic lives in archiveWeeklyKpis,
   setupWeeklyArchiveTrigger, listArchivedWeeks etc. (see "WEEKLY KPI
   ARCHIVE" section below). Each wrapper shows a clear confirmation /
   summary toast and never silently fails.
   ────────────────────────────────────────────────────────────────────── */

function menuInstallWeeklyArchiveTrigger() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    'Install Weekly KPI Archive Trigger?',
    'A time-based trigger will fire every Sunday at 23:30 to snapshot that ' +
    'week\'s KPIs into the KPI_Weekly_Archive sheet.\n\n' +
    'HoD, HoI and IMO can then browse past weeks from the "Past Weeks" tab ' +
    'in the web app, even after submissions are re-reviewed or edited.\n\n' +
    'Existing triggers pointing at archiveWeeklyKpis will be replaced ' +
    '(idempotent — safe to run multiple times).',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp !== ui.Button.OK) return;
  try {
    var r = setupWeeklyArchiveTrigger();
    ui.alert('Trigger installed',
      '✅ Weekly archive trigger installed.\n\n' +
      'Schedule: ' + r.weekday + ' at ' + r.hour + ':00 server time.\n\n' +
      'The first snapshot will be written this coming Sunday at 23:30. ' +
      'If you want a snapshot of last week right now, run ' +
      '"Snapshot Last Week Now (manual)" from the same menu.',
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Failed to install trigger', '⚠️ ' + (e && e.message ? e.message : e), ui.ButtonSet.OK);
  }
}

function menuRunWeeklyArchiveNow() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    'Snapshot last completed week now?',
    'Runs archiveWeeklyKpis() once, immediately. Captures last Mon–Sun ' +
    'across campus, institution and department dimensions and appends ' +
    'them to the KPI_Weekly_Archive sheet.\n\n' +
    'Re-running for the same week REPLACES the earlier snapshot (idempotent).\n\n' +
    'Use this to backfill a missed week, or for the first installation ' +
    'so dashboards have at least one past-week to show.',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp !== ui.Button.OK) return;
  try {
    var r = archiveWeeklyKpis();
    ui.alert('Snapshot complete',
      '✅ Snapshot written.\n\n' +
      'Week:           ' + r.weekStart + ' → ' + r.weekEnd + '\n' +
      'Dimensions:     ' + (r.dimensions || []).join(', ') + '\n' +
      'Rows archived:  ' + r.rowsArchived + '\n\n' +
      'Open the web app → Past Weeks (HoD/HoI/IMO) to verify.',
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Snapshot failed', '⚠️ ' + (e && e.message ? e.message : e), ui.ButtonSet.OK);
  }
}

function menuCheckWeeklyArchiveStatus() {
  var ui = SpreadsheetApp.getUi();
  try {
    // 1) Trigger presence
    var triggers = ScriptApp.getProjectTriggers();
    var archTriggers = triggers.filter(function(t){ return t.getHandlerFunction() === 'archiveWeeklyKpis'; });
    var triggerLine;
    if (!archTriggers.length) {
      triggerLine = '❌ NOT INSTALLED — run "Install Sunday 23:30 Trigger" from this menu.';
    } else {
      triggerLine = '✅ Installed (' + archTriggers.length + ' trigger' + (archTriggers.length === 1 ? '' : 's') + ').';
    }

    // 2) Archive sheet presence + row counts
    var sheetLine, weekLine;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('KPI_Weekly_Archive');
    if (!sh) {
      sheetLine = '❌ KPI_Weekly_Archive sheet not yet created.';
      weekLine  = '   No snapshots yet.';
    } else {
      var lastRow = sh.getLastRow();
      var dataRows = Math.max(0, lastRow - 1);
      sheetLine = '✅ KPI_Weekly_Archive exists (' + dataRows + ' archived row' + (dataRows === 1 ? '' : 's') + ').';
      var list = listArchivedWeeks({});
      var weeks = (list && list.weeks) || [];
      if (!weeks.length) {
        weekLine = '   No weeks archived yet.';
      } else {
        var newest = weeks[0], oldest = weeks[weeks.length - 1];
        weekLine = '   ' + weeks.length + ' distinct week' + (weeks.length === 1 ? '' : 's') +
                   ' archived.\n   Oldest: ' + oldest.label +
                   '\n   Newest: ' + newest.label;
      }
    }

    ui.alert('Weekly KPI Archive — Status',
      'Trigger:\n   ' + triggerLine + '\n\n' +
      'Storage:\n   ' + sheetLine + '\n' + weekLine,
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Status check failed', '⚠️ ' + (e && e.message ? e.message : e), ui.ButtonSet.OK);
  }
}

function menuRemoveWeeklyArchiveTrigger() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    'Remove the Weekly Archive trigger?',
    'This stops the Sunday 23:30 snapshot from running automatically.\n\n' +
    'Existing archived weeks in KPI_Weekly_Archive are PRESERVED — only ' +
    'the schedule is removed. You can re-install it any time from this menu.',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp !== ui.Button.OK) return;
  try {
    var triggers = ScriptApp.getProjectTriggers();
    var removed = 0;
    triggers.forEach(function(t){
      if (t.getHandlerFunction() === 'archiveWeeklyKpis') {
        ScriptApp.deleteTrigger(t); removed++;
      }
    });
    ui.alert('Trigger removed',
      removed
        ? '✅ Removed ' + removed + ' trigger' + (removed === 1 ? '' : 's') + '. Archived weeks are preserved.'
        : 'ℹ️ No archive trigger was installed. Nothing to remove.',
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Remove failed', '⚠️ ' + (e && e.message ? e.message : e), ui.ButtonSet.OK);
  }
}

function openWebApp() {
  var url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput('<script>window.open("'+url+'","_blank");google.script.host.close();<\/script>').setWidth(1).setHeight(1),
    'Opening Web App...'
  );
}

// ─── INITIALIZE ───────────────────────────────────────────────────────────────
function initializeSystem() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Also ensure HOD submission sheets exist
  ['HOD_Submission','HOD_Timesheet','HOD_SelfAssess','HOD_Review','HOD_IMO_Review','HOI_WeeklyMeeting'].forEach(function(name){
    if(!ss.getSheetByName(name)){
      var sh=ss.insertSheet(name);
      var hdrs=SCHEMA[name];
      if(hdrs){var rng=sh.getRange(1,1,1,hdrs.length);rng.setValues([hdrs]).setFontWeight('bold').setBackground('#1a3c5e').setFontColor('#ffffff');sh.setFrozenRows(1);}
    }
  });
  var names = Object.keys(SH).map(function(k){ return SH[k]; });
  names.forEach(function(n){ if(!ss.getSheetByName(n)) ss.insertSheet(n); });
  names.forEach(function(n){
    var sheet = ss.getSheetByName(n), hdrs = SCHEMA[n];
    if(sheet.getRange(1,1).getValue() !== hdrs[0]){
      // Brand new sheet — write all headers
      var r = sheet.getRange(1,1,1,hdrs.length);
      r.setValues([hdrs]);
      r.setFontWeight('bold').setBackground('#1a3c5e').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    } else {
      // Existing sheet — add any missing columns without touching existing data
      _ensureSheetColumns(sheet, hdrs);
    }
  });
  _applyValidations(ss);
  _applyFormatting(ss);
  _setColumnWidths(ss);
  _resetTriggers();
  // Seed default IMO credentials in Script Properties (change before go-live)
  var p = PropertiesService.getScriptProperties();
  if (!p.getProperty('IMO_EMAIL'))    p.setProperty('IMO_EMAIL',    'imo@vmrf.edu.in');
  if (!p.getProperty('IMO_PASSWORD')) p.setProperty('IMO_PASSWORD', 'IMO@VMRF2024');
  if (!p.getProperty('HOI_EMAIL'))    p.setProperty('HOI_EMAIL',    'hoi@vmrf.edu.in');
  if (!p.getProperty('HOI_PASSWORD')) p.setProperty('HOI_PASSWORD', 'HOI@VMRF2024');
  try { SpreadsheetApp.getUi().alert(
    '✅ VMRF System Ready!\n\n' +
    'IMO Login Credentials (set in Script Properties):\n' +
    '  IMO_EMAIL    : ' + p.getProperty('IMO_EMAIL')    + '\n' +
    '  IMO_PASSWORD : ' + p.getProperty('IMO_PASSWORD') + '\n' +
    '\nHOI Login Credentials (set in Script Properties):\n' +
    '  HOI_EMAIL    : ' + p.getProperty('HOI_EMAIL')    + '\n' +
    '  HOI_PASSWORD : ' + p.getProperty('HOI_PASSWORD') + '\n' +
    '\n' +
    'HOD: Each department has its own credentials.\n' +
    'Run "Seed HOD Department Accounts" from the VMRF IMO menu to create them.\n\n' +
    'To change IMO/HOI: Apps Script → Project Settings → Script Properties\n\n' +
    'Deploy as Web App when ready.'
  ); } catch(e) {}
}

// Run this directly from the VMRF menu if Status dropdowns are missing
function applyStatusDropdowns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var R  = 500;
  _dv(ss.getSheetByName(SH.STAFF),  2, 8, R, ['Active','Pending','Inactive']);
  _dv(ss.getSheetByName(SH.FACULTY),2, 9, R, ['Active','Pending','Inactive']);
  try { SpreadsheetApp.getUi().alert('✅ Status dropdowns applied to Staff_Master (col 8) and Faculty_Master (col 9).'); } catch(e) {}
}

function _applyValidations(ss) {
  var R = 500;
  var fm = ss.getSheetByName(SH.FACULTY);
  // Faculty_Master schema: FacultyName(1), Email(2), Department(3), Campus(4), Institution(5), Designation(6), PasswordHash(7), GoogleEmail(8), Status(9)
  _dv(fm,2,3,R,DEPARTMENTS); _dv(fm,2,4,R,CAMPUSES);
  _dv(fm,2,5,R,INSTITUTIONS); _dv(fm,2,6,R,DESIGNATIONS);
  // Clear any stale validation on PasswordHash column (col 7) from old schema offset bug
  fm.getRange(2,7,R,1).clearDataValidations();
  var ws = ss.getSheetByName(SH.SUBMISSION);
  _dv(ws,2,3,R,ACADEMIC_YEARS); _dv(ws,2,6,R,['YES','NO']);
  ws.getRange(2,4,R,2).setNumberFormat('dd-MMM-yyyy');
  // Timesheet_Entries is written programmatically — no cell validation needed
  var ts = ss.getSheetByName(SH.TIMESHEET);
  ts.getRange(2,1,R,5).clearDataValidations();
  _dv(ss.getSheetByName(SH.HOD),2,3,R,['Approved','Rejected']);
  _dv(ss.getSheetByName(SH.HOI),2,3,R,['Approved','Rejected']);
  _dv(ss.getSheetByName(SH.IMO),2,3,R,['Under Review','Finalised']);
  // Status dropdowns — col 8 in Staff_Master, col 9 in Faculty_Master
  _dv(ss.getSheetByName(SH.STAFF),  2, 8, R, ['Active','Pending','Inactive']);
  _dv(ss.getSheetByName(SH.FACULTY),2, 9, R, ['Active','Pending','Inactive']);
}

function _dv(sheet,sr,col,nr,list){
  sheet.getRange(sr,col,nr,1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(list,true).setAllowInvalid(false).build());
}

function _applyFormatting(ss) {
  [{n:SH.HOD,c:3,m:[['Approved','#b7e1cd'],['Rejected','#f4c7c3']]},
   {n:SH.HOI,c:3,m:[['Approved','#b7e1cd'],['Rejected','#f4c7c3']]},
   {n:SH.IMO,c:3,m:[['Finalised','#b7e1cd'],['Under Review','#fce8b2']]}
  ].forEach(function(cfg){
    var sheet=ss.getSheetByName(cfg.n),range=sheet.getRange(2,cfg.c,500,1);
    sheet.setConditionalFormatRules(cfg.m.map(function(r){
      return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(r[0]).setBackground(r[1]).setRanges([range]).build();
    }));
  });
}

function _setColumnWidths(ss) {
  var map={};
  map[SH.STAFF]=[120,180,220,80,180,160,200,80];
  map[SH.FACULTY]=[180,200,150,220,260,160,160,200,80];
  map[SH.SUBMISSION]=[200,120,170,120,120,80,160];
  map[SH.TIMESHEET]=[200,100,170,280,200];
  map[SH.SELF_ASSESS]=[200,360,360];
  map[SH.HOD]=map[SH.HOI]=map[SH.IMO]=[200,320,160,160];
  map[SH.NOTIF]=[180,80,120,300,500,200,160,60,160];
  Object.keys(map).forEach(function(n){
    var s=ss.getSheetByName(n);
    if(s) map[n].forEach(function(w,i){s.setColumnWidth(i+1,w);});
  });
}

function _resetTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t){ScriptApp.deleteTrigger(t);});
  ScriptApp.newTrigger('sendFridayReminders').timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(12).create();
}

// ─── IMO QUEUE ────────────────────────────────────────────────────────────────
function getIMOQueue() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var subD=ss.getSheetByName(SH.SUBMISSION).getDataRange().getValues();
  var facD=ss.getSheetByName(SH.FACULTY).getDataRange().getValues();
  var hodD=ss.getSheetByName(SH.HOD).getDataRange().getValues();
  var hoiD=ss.getSheetByName(SH.HOI).getDataRange().getValues();
  var imoD=ss.getSheetByName(SH.IMO).getDataRange().getValues();
  var saD =ss.getSheetByName(SH.SELF_ASSESS).getDataRange().getValues();
  var tsD =ss.getSheetByName(SH.TIMESHEET).getDataRange().getValues();
  var subMap=_bm(subD,subD[0]);
  var sidEmailMap=_buildSidEmailMap(subD);
  var facMap=_buildFacMap(facD, ss.getSheetByName(SH.STAFF).getDataRange().getValues());
  var hodMap=_bm(hodD,hodD[0]);
  var hoiMap=_bm(hoiD,hoiD[0]);
  var imoMap=_bm(imoD,imoD[0]);
  var saMap=_bm(saD,saD[0]);
  var tsMap=_bmMulti(tsD,tsD[0]);
  var hoiH=hoiD[0];
  var hoiStI=hoiH.indexOf('HOI_Status'), hoiSbI=hoiH.indexOf('SubmissionID');
  var out=[];
  for(var i=1;i<hoiD.length;i++){
    var sid=String(hoiD[i][hoiSbI]||'').trim();
    var hoiSt=String(hoiD[i][hoiStI]||'').trim();
    if(!sid||hoiSt!=='Approved') continue;
    var imoR=imoMap[sid]||{}, imoSt=String(imoR['IMO_Status']||'').trim();
    if(imoSt==='Finalised') continue;
    var sub=subMap[sid]||{};
    var fid=sidEmailMap[sid]||_getFidFromSub(sub);
    var fac=fid?(facMap[fid]||{}):{};
    var sa=saMap[sid]||{}, hodR=hodMap[sid]||{}, hoiR2=hoiMap[sid]||{};
    out.push(_buildItem(sid,sub,fac,sa,tsMap[sid]||[],
      {hodStatus:String(hodR['HOD_Status']||''),hodRemark:String(hodR['HOD_Remark']||'')},
      {hoiStatus:hoiSt,hoiRemark:String(hoiR2['HOI_Remark']||'')},
      {imoStatus:imoSt}));
  }
  return out;
}

function submitIMOReview(sid, remark, status) {
  // IMO role is monitoring only — this function is retained for remark logging only.
  if(!sid) throw new Error('Submission ID missing.');
  if(remark) _writeReview(SH.IMO, sid, remark, 'Finalised'); // preserve existing Finalised status
  return { ok:true };
}

// ─── ALL SUBMISSIONS (IMO view) ───────────────────────────────────────────────
function getAllSubmissions() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var subD=ss.getSheetByName(SH.SUBMISSION).getDataRange().getValues(), subH=subD[0].map(function(v){return String(v||'').trim();});
  var facD=ss.getSheetByName(SH.FACULTY).getDataRange().getValues();
  var facH=facD[0].map(function(v){return String(v||'').trim();});
  var hodD=ss.getSheetByName(SH.HOD).getDataRange().getValues();
  var hoiD=ss.getSheetByName(SH.HOI).getDataRange().getValues();
  var imoD=ss.getSheetByName(SH.IMO).getDataRange().getValues();
  var facMap=_buildFacMap(facD, ss.getSheetByName(SH.STAFF).getDataRange().getValues());
  var hodMap=_bm(hodD,hodD[0]), hoiMap=_bm(hoiD,hoiD[0]), imoMap=_bm(imoD,imoD[0]);

  // Identify faculty email column in Weekly_Submission by multiple strategies
  var _fecI = _facEmailCol(subH);
  // If header-based detection fails, scan first data row for a cell containing '@'
  if (_fecI < 0 && subD.length > 1) {
    for (var ci = 0; ci < subD[1].length; ci++) {
      if (String(subD[1][ci]||'').indexOf('@') >= 0) { _fecI = ci; break; }
    }
  }
  if (_fecI < 0) _fecI = 1; // ultimate fallback: column B

  // Pre-build a brute-force Faculty_Master lookup (scan every cell for email)
  // This catches ANY schema — old, new, renamed, shifted columns
  var _bruteMap = {};
  for (var fi = 1; fi < facD.length; fi++) {
    var frow = facD[fi];
    var fnm  = String(frow[facH.indexOf('FacultyName')>=0?facH.indexOf('FacultyName'):0]||'');
    var fdp  = String(frow[facH.indexOf('Department')>=0?facH.indexOf('Department'):2]||'');
    var fin  = String(frow[facH.indexOf('Institution')>=0?facH.indexOf('Institution'):4]||'');
    var fcm  = String(frow[facH.indexOf('Campus')>=0?facH.indexOf('Campus'):3]||'');
    var fds  = String(frow[facH.indexOf('Designation')>=0?facH.indexOf('Designation'):5]||'');
    // Index by every cell that looks like an email
    for (var fc = 0; fc < frow.length; fc++) {
      var cv = String(frow[fc]||'').trim().toLowerCase();
      if (cv.indexOf('@') >= 0 && !_bruteMap[cv]) {
        _bruteMap[cv] = {name:fnm, dept:fdp, inst:fin, campus:fcm, desig:fds};
      }
    }
  }

  var out=[];
  for(var i=1;i<subD.length;i++){
    var sid=String(subD[i][subH.indexOf('SubmissionID')>=0?subH.indexOf('SubmissionID'):0]||'');
    if(!sid) continue;
    var facID=String(subD[i][_fecI]||'').trim().toLowerCase();
    // If primary column has no email, scan entire row for an email
    if (!facID || facID.indexOf('@') < 0) {
      for (var rc = 0; rc < subD[i].length; rc++) {
        var rv = String(subD[i][rc]||'').trim().toLowerCase();
        if (rv.indexOf('@') >= 0) { facID = rv; break; }
      }
    }

    var fac=facMap[facID]||{};
    // If facMap lookup returned empty, use brute-force lookup
    var facName = String(fac['FacultyName']||'');
    var facDept = String(fac['Department']||'');
    var facInst = String(fac['Institution']||'');
    var facCamp = String(fac['Campus']||'');
    var facDesg = String(fac['Designation']||'');
    if (!facName && _bruteMap[facID]) {
      var bf = _bruteMap[facID];
      facName = bf.name; facDept = bf.dept; facInst = bf.inst;
      facCamp = bf.campus; facDesg = bf.desig;
    }
    // Final fallbacks
    if (!facName) facName = facID || 'Unknown';
    if (!facInst && INSTITUTIONS.length) facInst = INSTITUTIONS[0];
    if (!facCamp && CAMPUSES.length) facCamp = CAMPUSES[0];

    var hod=hodMap[sid]||{}, hoi=hoiMap[sid]||{}, imo=imoMap[sid]||{};
    var _isNH = _isNoHodInstitution_(facInst);
    out.push({
      submissionID: sid,
      facultyName:  facName,
      facultyID:    facID,
      department:   facDept,
      institution:  facInst,
      campus:       facCamp,
      designation:  facDesg,
      semester:     String(subD[i][subH.indexOf('AcademicYearSemester')>=0?subH.indexOf('AcademicYearSemester'):2]||''),
      from:         _fmt(subD[i][subH.indexOf('ReportingFrom')>=0?subH.indexOf('ReportingFrom'):3]),
      to:           _fmt(subD[i][subH.indexOf('ReportingTo')>=0?subH.indexOf('ReportingTo'):4]),
      submitted:    _fmtDT(subD[i][subH.indexOf('SubmittedDateTime')>=0?subH.indexOf('SubmittedDateTime'):6]),
      hodStatus:    _isNH ? 'N/A' : String(hod['HOD_Status']||'Pending'),
      hoiStatus:    hoi['HOI_Status'] ? String(hoi['HOI_Status']) : (_isNH || String(hod['HOD_Status']||'')==='Approved' ? 'Pending' : '—'),
      imoStatus:    imo['IMO_Status'] ? String(imo['IMO_Status']) : (String(hoi['HOI_Status']||'')==='Approved' ? 'Pending' : '—')
    });
  }
  return out.reverse();
}

// ─── ALL HOD SUBMISSIONS (IMO view) ──────────────────────────────────────────
function getAllHODSubmissions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var subSh = ss.getSheetByName(SH.HOD_SUB); if (!subSh) return [];
  var subD = subSh.getDataRange().getValues(), subH = subD[0];
  var revSh = ss.getSheetByName(SH.HOD_REVIEW);
  var revD = revSh ? revSh.getDataRange().getValues() : [['SubmissionID']]; var revH = revD[0];
  var imoSh = ss.getSheetByName(SH.HOD_IMO);
  var imoD = imoSh ? imoSh.getDataRange().getValues() : [['SubmissionID']]; var imoH = imoD[0];
  var staffSh = ss.getSheetByName(SH.STAFF);
  var staffMap = {};
  if (staffSh) {
    var staffD = staffSh.getDataRange().getValues(), staffH2 = staffD[0];
    var _sE = staffH2.indexOf('Email'), _sN = staffH2.indexOf('StaffName'), _sD = staffH2.indexOf('Department');
    for (var sf = 1; sf < staffD.length; sf++) {
      var se = String(staffD[sf][_sE]||'').trim().toLowerCase();
      if (se) staffMap[se] = { name: String(staffD[sf][_sN]||''), dept: String(staffD[sf][_sD]||'') };
    }
  }
  var revMap = {};
  for (var r2 = 1; r2 < revD.length; r2++) { revMap[String(revD[r2][0]||'')] = String(revD[r2][revH.indexOf('HOI_Status')>=0?revH.indexOf('HOI_Status'):2]||''); }
  var imoMap2 = {};
  var _imStI = imoH.indexOf('IMO_Status');
  for (var im = 1; im < imoD.length; im++) { imoMap2[String(imoD[im][0]||'')] = String(imoD[im][_imStI>=0?_imStI:2]||''); }

  var out = [];
  for (var i = 1; i < subD.length; i++) {
    var r = subD[i];
    var sid = String(r[0]||'').trim(); if (!sid) continue;
    var hodID = String(r[subH.indexOf('HOD_ID')>=0?subH.indexOf('HOD_ID'):1]||'').trim().toLowerCase();
    var staff = staffMap[hodID] || {};
    var hoiSt = revMap[sid] || '';
    var imoSt = imoMap2[sid] || '';
    out.push({
      submissionID: sid,
      facultyName:  staff.name || 'HOD',
      facultyID:    hodID,
      department:   staff.dept || '',
      institution:  INSTITUTIONS.length ? INSTITUTIONS[0] : '',
      campus:       CAMPUSES.length ? CAMPUSES[0] : '',
      designation:  'HOD',
      semester:     String(r[subH.indexOf('AcademicYearSemester')>=0?subH.indexOf('AcademicYearSemester'):2]||''),
      from:         _fmt(r[subH.indexOf('ReportingFrom')>=0?subH.indexOf('ReportingFrom'):3]),
      to:           _fmt(r[subH.indexOf('ReportingTo')>=0?subH.indexOf('ReportingTo'):4]),
      submitted:    _fmtDT(r[subH.indexOf('SubmittedDateTime')>=0?subH.indexOf('SubmittedDateTime'):6]),
      hodStatus:    'N/A',
      hoiStatus:    hoiSt || 'Pending',
      imoStatus:    imoSt ? imoSt : (hoiSt === 'Approved' ? 'Pending' : '—'),
      _type:        'HOD'
    });
  }
  return out.reverse();
}

// ─── FACULTY SUBMISSION HISTORY (for drill-down) ─────────────────────────────
function getFacultySubmissionHistory(email) {
  if (!email) return [];
  email = String(email).trim().toLowerCase();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var subD = ss.getSheetByName(SH.SUBMISSION).getDataRange().getValues(), subH = subD[0];
  var hodD = ss.getSheetByName(SH.HOD).getDataRange().getValues();
  var hoiD = ss.getSheetByName(SH.HOI).getDataRange().getValues();
  var imoD = ss.getSheetByName(SH.IMO).getDataRange().getValues();
  var saD  = ss.getSheetByName(SH.SELF_ASSESS).getDataRange().getValues();
  var tsD  = ss.getSheetByName(SH.TIMESHEET).getDataRange().getValues();
  var hodMap=_bm(hodD,hodD[0]), hoiMap=_bm(hoiD,hoiD[0]), imoMap=_bm(imoD,imoD[0]);
  var saMap=_bm(saD,saD[0]), tsMap=_bmMulti(tsD,tsD[0]);
  var sidEmailMap = _buildSidEmailMap(subD);

  var out = [];
  for (var sid in sidEmailMap) {
    if (sidEmailMap[sid] !== email) continue;
    var hod=hodMap[sid]||{}, hoi=hoiMap[sid]||{}, imo=imoMap[sid]||{}, sa=saMap[sid]||{};
    // Find the submission row
    var subRow = {};
    var _sidI = subH.indexOf('SubmissionID'); if(_sidI<0) _sidI=0;
    for (var r=1; r<subD.length; r++) {
      if (String(subD[r][_sidI]||'').trim() === sid) {
        subH.forEach(function(c,j){ subRow[c]=subD[r][j]; });
        break;
      }
    }
    var tsEntries = (tsMap[sid]||[]).map(function(t){
      return {
        Day:            String(t['Day']||''),
        TimeSlot:       String(t['TimeSlot']||''),
        ActivityType:   String(t['ActivityType']||''),
        Details:        String(t['ActivityDetails']||t['Details']||''),
        AttachmentURL:  String(t['AttachmentURL']||''),
        AttachmentName: String(t['AttachmentName']||'')
      };
    });
    out.push({
      submissionID: sid,
      semester:     String(subRow['AcademicYearSemester']||''),
      from:         _fmt(subRow['ReportingFrom']),
      to:           _fmt(subRow['ReportingTo']),
      submitted:    _fmtDT(subRow['SubmittedDateTime']),
      hodStatus:    String(hod['HOD_Status']||'Pending'),
      hoiStatus:    hoi['HOI_Status'] ? String(hoi['HOI_Status']) : (String(hod['HOD_Status']||'')==='Approved'?'Pending':'—'),
      imoStatus:    imo['IMO_Status'] ? String(imo['IMO_Status']) : (String(hoi['HOI_Status']||'')==='Approved'?'Pending':'—'),
      hodRemark:    String(hod['HOD_Remark']||''),
      hoiRemark:    String(hoi['HOI_Remark']||''),
      imoRemark:    String(imo['IMO_Remark']||''),
      outcome:      String(sa['OutcomeOfWeek']||''),
      target:       String(sa['TargetPlanNextWeek']||''),
      timesheet:    tsEntries
    });
  }
  return out.reverse();
}

// ─── DASHBOARD STATS ─────────────────────────────────────────────────────────

// ─── STATS DRILL-DOWN: returns faculty list for a given KPI filter ─────────────
function getStatsDrilldown(filter) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var subD = ss.getSheetByName(SH.SUBMISSION).getDataRange().getValues(), subH = subD[0];
  var facD = ss.getSheetByName(SH.FACULTY).getDataRange().getValues();
  var hodD = ss.getSheetByName(SH.HOD).getDataRange().getValues(), hodH = hodD[0];
  var hoiD = ss.getSheetByName(SH.HOI).getDataRange().getValues(), hoiH = hoiD[0];
  var imoD = ss.getSheetByName(SH.IMO).getDataRange().getValues(), imoH = imoD[0];

  var facMap = _buildFacMap(facD, ss.getSheetByName(SH.STAFF).getDataRange().getValues());
  var sidEmailMap = _buildSidEmailMap(subD);

  var _frI = subH.indexOf('ReportingFrom'); if(_frI<0) _frI=3;
  var _toI = subH.indexOf('ReportingTo');   if(_toI<0) _toI=4;
  var _sidI = subH.indexOf('SubmissionID'); if(_sidI<0) _sidI=0;
  var subInfo = {};
  for (var s = 1; s < subD.length; s++) {
    var sid2 = String(subD[s][_sidI]||'').trim();
    if (sid2) subInfo[sid2] = { from: _fmt(subD[s][_frI]), to: _fmt(subD[s][_toI]) };
  }

  var results = [];
  var role = String(filter.role || ''), status = String(filter.status || '');
  var seen = {};
  var filterDept = (filter.department||'').trim().toLowerCase();
  var filterDeptBase = filterDept.replace(/\s*\(pg\)\s*$/,'');
  var filterCampus = (filter.campus||'').trim();
  var filterInst = (filter.institution||'').trim();

  function addRow(sid, reviewStatus) {
    if (seen[sid]) return; seen[sid] = true;
    var email = sidEmailMap[sid] || '';
    var fac = email ? (facMap[email]||{}) : {};
    var facDept = String(fac['Department']||'').trim().toLowerCase();
    var facDeptBase = facDept.replace(/\s*\(pg\)\s*$/,'');
    if (filterDept && facDept !== filterDept && facDeptBase !== filterDeptBase) return;
    if (filterCampus && String(fac['Campus']||'').trim() !== filterCampus) return;
    if (filterInst && String(fac['Institution']||'').trim() !== filterInst) return;
    var info = subInfo[sid] || {};
    results.push({
      sid: sid,
      name: String(fac['FacultyName']||email||'Unknown'),
      dept: String(fac['Department']||''),
      desig: String(fac['Designation']||''),
      institution: String(fac['Institution']||''),
      campus: String(fac['Campus']||''),
      from: info.from||'', to: info.to||'',
      status: reviewStatus
    });
  }

  // Special drilldowns
  if (status === 'AllFaculty') {
    // Only iterate Faculty_Master sheet, NOT Staff_Master
    var _fH = facD[0].map(function(v){return String(v||'').trim();});
    var _fNm = _fH.indexOf('FacultyName'); if(_fNm<0) _fNm=0;
    var _fEm = _fH.indexOf('Email'); if(_fEm<0) _fEm=1;
    var _fDp = _fH.indexOf('Department'); if(_fDp<0) _fDp=2;
    var _fDs = _fH.indexOf('Designation'); if(_fDs<0) _fDs=5;
    var _fSt = _fH.indexOf('Status'); if(_fSt<0) _fSt=_fH.length-1;
    var _fCm = _fH.indexOf('Campus'); if(_fCm<0) _fCm=3;
    var _fIn = _fH.indexOf('Institution'); if(_fIn<0) _fIn=4;
    for (var af = 1; af < facD.length; af++) {
      var afSt = String(facD[af][_fSt]||'').trim();
      if (afSt && afSt !== 'Active') continue;
      var afDept = String(facD[af][_fDp]||'').trim();
      var afDeptLow = afDept.toLowerCase();
      var afDeptBase = afDeptLow.replace(/\s*\(pg\)\s*$/,'');
      if (filterDept && afDeptLow !== filterDept && afDeptBase !== filterDeptBase) continue;
      if (filter.campus && String(facD[af][_fCm]||'').trim() !== filter.campus) continue;
      if (filter.institution && String(facD[af][_fIn]||'').trim() !== filter.institution) continue;
      results.push({
        sid:'',
        name: String(facD[af][_fNm]||''),
        dept: afDept,
        desig: String(facD[af][_fDs]||''),
        institution: String(facD[af][_fIn]||''),
        campus: String(facD[af][_fCm]||''),
        from:'', to:'',
        status:'Active'
      });
    }
    return results;
  }
  if (status === 'All') {
    var hodMap3={}, hoiMap3={}, imoMap3={};
    for(var h3=1;h3<hodD.length;h3++) hodMap3[String(hodD[h3][0]||'')]=String(hodD[h3][hodH.indexOf('HOD_Status')]||'');
    for(var o3=1;o3<hoiD.length;o3++) hoiMap3[String(hoiD[o3][0]||'')]=String(hoiD[o3][hoiH.indexOf('HOI_Status')]||'');
    for(var m3=1;m3<imoD.length;m3++) imoMap3[String(imoD[m3][0]||'')]=String(imoD[m3][imoH.indexOf('IMO_Status')]||'');
    for (var sk in subInfo) {
      var hSt=hodMap3[sk]||'', oSt=hoiMap3[sk]||'', mSt=imoMap3[sk]||'';
      var disp = mSt==='Finalised'?'Finalised':mSt==='Escalated'?'Escalated':oSt==='Approved'?'Pending IMO':hSt==='Approved'?'Pending HOI':hSt?hSt:'Pending HOD';
      addRow(sk, disp);
    }
    return results;
  }

  if (role === 'HOD') {
    var hodStI = hodH.indexOf('HOD_Status'); if(hodStI<0) hodStI=2;
    var hodSidI = hodH.indexOf('SubmissionID'); if(hodSidI<0) hodSidI=0;
    for (var i = 1; i < hodD.length; i++) {
      var hs = String(hodD[i][hodStI]||'').trim();
      var sid3 = String(hodD[i][hodSidI]||'').trim();
      if (!sid3) continue;
      if (status === 'Pending' && hs === '') addRow(sid3, 'Pending');
      else if (status === hs) addRow(sid3, hs);
    }
  } else if (role === 'HOI') {
    // Build lookup maps from review sheets
    var hoiHodMap = {}, hoiHoiMap = {}, hoiImoMap2 = {};
    for (var hh=1; hh<hodD.length; hh++) { var hhi=String(hodD[hh][0]||'').trim(); if(hhi) hoiHodMap[hhi]=String(hodD[hh][hodH.indexOf('HOD_Status')]||'').trim(); }
    for (var hj=1; hj<hoiD.length; hj++) { var hji=String(hoiD[hj][0]||'').trim(); if(hji) hoiHoiMap[hji]=String(hoiD[hj][hoiH.indexOf('HOI_Status')]||'').trim(); }
    for (var hk=1; hk<imoD.length; hk++) { var hki=String(imoD[hk][0]||'').trim(); if(hki) hoiImoMap2[hki]=String(imoD[hk][imoH.indexOf('IMO_Status')]||'').trim(); }

    if (status === 'All') {
      for (var sid_hA in subInfo) {
        var hA_hod=hoiHodMap[sid_hA]||'', hA_hoi=hoiHoiMap[sid_hA]||'';
        var disp_hA = (hA_hoi==='Approved')?'Submitted to IMO':hA_hod==='Rejected'||hA_hoi==='Rejected'?'Rejected':(hA_hod==='Approved'&&hA_hoi==='')?'Forwarded to HOI':(hA_hod==='')?'Submitted to HOD':'Submitted to HOD';
        addRow(sid_hA, disp_hA);
      }
    } else if (status === 'Pending HOD') {
      // Submissions still waiting for HOD review
      for (var sid_hB in subInfo) { if((hoiHodMap[sid_hB]||'')==='') addRow(sid_hB, 'Submitted to HOD'); }
    } else if (status === 'Pending HOI' || status === 'Pending') {
      // HOD approved, HOI not yet reviewed
      for (var sid_hC in subInfo) {
        if ((hoiHodMap[sid_hC]||'')==='Approved' && (hoiHoiMap[sid_hC]||'')==='') addRow(sid_hC, 'Forwarded to HOI');
      }
    } else if (status === 'Finalised' || status === 'Pending IMO') {
      // HOI approved = Submitted to IMO (final)
      for (var sid_hD in subInfo) {
        if ((hoiHoiMap[sid_hD]||'')==='Approved' || (hoiImoMap2[sid_hD]||'')==='Finalised') addRow(sid_hD, 'Submitted to IMO');
      }
    } else if (status === 'Approved') {
      // HOI has approved (same as Finalised/Submitted to IMO)
      for (var sid_hE in subInfo) {
        if ((hoiHoiMap[sid_hE]||'')==='Approved') addRow(sid_hE, 'Submitted to IMO');
      }
    } else if (status === 'Rejected' || status === 'Needs Revision') {
      for (var sid_hF in subInfo) {
        var hF_hod=hoiHodMap[sid_hF]||'', hF_hoi=hoiHoiMap[sid_hF]||'';
        if (hF_hod==='Rejected'||hF_hoi==='Rejected'||hF_hod==='Needs Revision'||hF_hoi==='Needs Revision') addRow(sid_hF, 'Rejected');
      }
    }
  } else if (role === 'IMO') {
    // Build quick-lookup maps for HOD and HOI statuses
    var imoHodMap = {}, imoHoiMap = {}, imoImoMap = {};
    for (var hm=1; hm<hodD.length; hm++) { var hid=String(hodD[hm][0]||'').trim(); if(hid) imoHodMap[hid]=String(hodD[hm][hodH.indexOf('HOD_Status')]||'').trim(); }
    for (var om=1; om<hoiD.length; om++) { var oid=String(hoiD[om][0]||'').trim(); if(oid) imoHoiMap[oid]=String(hoiD[om][hoiH.indexOf('HOI_Status')]||'').trim(); }
    for (var mm=1; mm<imoD.length; mm++) { var mid=String(imoD[mm][0]||'').trim(); if(mid) imoImoMap[mid]=String(imoD[mm][imoH.indexOf('IMO_Status')]||'').trim(); }

    if (status === 'All') {
      // All submissions — derive display status from review chain
      for (var sid_a in subInfo) {
        var hodSt_a = imoHodMap[sid_a]||'', hoiSt_a = imoHoiMap[sid_a]||'', imoSt_a = imoImoMap[sid_a]||'';
        var disp_a = imoSt_a==='Finalised'?'Finalised':hoiSt_a==='Approved'?'Finalised':hodSt_a==='Rejected'?'Rejected by HOD':hoiSt_a==='Rejected'?'Rejected by HOI':hoiSt_a==='Pending'?'Pending HOI':hodSt_a==='Approved'?'Pending HOI':hodSt_a?hodSt_a:'Pending HOD';
        addRow(sid_a, disp_a);
      }
    } else if (status === 'Pending HOD') {
      for (var sid_b in subInfo) {
        var hs_b = imoHodMap[sid_b]||'';
        if (hs_b==='') addRow(sid_b, 'Pending HOD');
      }
    } else if (status === 'Pending HOI') {
      for (var sid_c in subInfo) {
        var hs_c = imoHodMap[sid_c]||'', ois_c = imoHoiMap[sid_c]||'';
        if (hs_c==='Approved' && ois_c==='' ) addRow(sid_c, 'Forwarded to HOI');
      }
    } else if (status === 'Finalised' || status === 'Pending IMO') {
      // Both mean Submitted to IMO in the new flow
      for (var sid_d in subInfo) {
        var ois_d = imoHoiMap[sid_d]||'', ims_d = imoImoMap[sid_d]||'';
        if (ois_d==='Approved'||ims_d==='Finalised') addRow(sid_d, 'Submitted to IMO');
      }
    } else if (status === 'Needs Revision' || status === 'Rejected') {
      for (var sid_e in subInfo) {
        var hs_e = imoHodMap[sid_e]||'', ois_e = imoHoiMap[sid_e]||'';
        if (hs_e==='Rejected'||ois_e==='Rejected'||hs_e==='Needs Revision'||ois_e==='Needs Revision') addRow(sid_e, 'Rejected');
      }
    } else {
      // Fallback: match by IMO_Status value
      for (var sid_f in imoImoMap) {
        if (imoImoMap[sid_f]===status) addRow(sid_f, status);
      }
    }
  }
  return results;
}

function getDashboardStats(hodDept, period) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var facD=ss.getSheetByName(SH.FACULTY).getDataRange().getValues(), facH=facD[0];
  var subD=ss.getSheetByName(SH.SUBMISSION).getDataRange().getValues(), subH=subD[0];
  var hodD=ss.getSheetByName(SH.HOD).getDataRange().getValues(), hodH=hodD[0];
  var hoiD=ss.getSheetByName(SH.HOI).getDataRange().getValues(), hoiH=hoiD[0];
  var imoD=ss.getSheetByName(SH.IMO).getDataRange().getValues(), imoH=imoD[0];
  var filterDept=(hodDept||'').trim().toLowerCase();

  // ── Period filter — KPIs scoped to week / month / year (except faculty) ──
  var _pb = _periodBounds_(String(period||'week'));
  var _pbS = _pb.start, _pbE = _pb.end;
  function _inPeriod(d){
    if(!_pbS && !_pbE) return true;                        // 'all' — no filter
    return (d instanceof Date) && d>=_pbS && d<=_pbE;
  }

  // Build submission → dept map AND submission → SubmittedDateTime date map
  // Using SubmittedDateTime (not ReportingFrom) matches the proven date-filter
  // approach used by getComparisonReport, which always works correctly.
  var facMap=_buildFacMap(facD, ss.getSheetByName(SH.STAFF).getDataRange().getValues());
  var _dsFi = _facEmailCol(subH); if(_dsFi<0)_dsFi=1;
  var _dsSidI = subH.indexOf('SubmissionID'); if(_dsSidI<0)_dsSidI=0;
  var _dsDtI  = subH.indexOf('SubmittedDateTime'); if(_dsDtI<0)_dsDtI=6;
  var subFidMap={}, subDateMap={};
  for(var s=1;s<subD.length;s++){
    var fid2=String(subD[s][_dsFi]||'').trim().toLowerCase();
    var sid_s=String(subD[s][_dsSidI]||'').trim();
    if(!sid_s) continue;
    subFidMap[sid_s]=fid2;
    var _dt=subD[s][_dsDtI];
    if(_dt instanceof Date) subDateMap[sid_s]=_dt;
  }
  function deptOf(sid2){
    var fid3=subFidMap[sid2]||''; var fac3=facMap[fid3]||{};
    return String(fac3['Department']||'').trim().toLowerCase();
  }
  function matchesDept(sid2){ return !filterDept||deptOf(sid2)===filterDept; }
  function matchesPeriodDept(sid2){
    var d=subDateMap[sid2];
    return _inPeriod(d) && matchesDept(sid2);
  }

  var pendHOD=0,pendHOI=0,pendIMO=0,finalised=0,escalated=0;
  var hodApproved=0,hodRevision=0,hodRejected=0;
  var hoiApproved=0,hoiRevision=0,hoiRejected=0;
  var sidI=hodH.indexOf('SubmissionID');
  for(var i=1;i<hodD.length;i++){
    var sid=String(hodD[i][sidI]||'').trim();
    if(!matchesPeriodDept(sid)) continue;
    var hs=String(hodD[i][hodH.indexOf('HOD_Status')]||'');
    if(hs==='') pendHOD++;
    if(hs==='Approved') hodApproved++;
    if(hs==='Rejected') hodRevision++;
    if(hs==='Rejected') hodRejected++;
  }
  var sidI2=hoiH.indexOf('SubmissionID');
  for(var j=1;j<hoiD.length;j++){
    var sid2=String(hoiD[j][sidI2]||'').trim();
    if(!matchesPeriodDept(sid2)) continue;
    var is=String(hoiD[j][hoiH.indexOf('HOI_Status')]||'');
    if(is==='') pendHOI++;
    if(is==='Approved') hoiApproved++;
    if(is==='Rejected') hoiRevision++;
    if(is==='Rejected') hoiRejected++;
  }
  var sidI3=imoH.indexOf('SubmissionID'); if(sidI3<0)sidI3=0;
  for(var k=1;k<imoD.length;k++){
    var sidK=String(imoD[k][sidI3]||'').trim();
    if(!_inPeriod(subDateMap[sidK])) continue;
    var ms=String(imoD[k][imoH.indexOf('IMO_Status')]||'');
    if(ms===''||ms==='Under Review') pendIMO++;
    if(ms==='Finalised') finalised++;
    if(ms==='Escalated') escalated++;
  }
  // Active Faculty — full roster headcount, never period-filtered
  var depFac=0, depSub=0;
  var facDepI=facH.indexOf('Department');
  for(var f=1;f<facD.length;f++){
    if(!filterDept||String(facD[f][facDepI]||'').trim().toLowerCase()===filterDept) depFac++;
  }
  // Total Submissions — period-scoped
  for(var ss2=1;ss2<subD.length;ss2++){
    var sid3=String(subD[ss2][_dsSidI]||'').trim();
    if(_inPeriod(subDateMap[sid3])&&matchesDept(sid3)) depSub++;
  }
  return {
    totalFaculty:depFac, totalSubmissions:depSub,
    pendingHOD:pendHOD, pendingHOI:pendHOI, pendingIMO:pendIMO,
    finalised:finalised, escalated:escalated,
    hodApproved:hodApproved, hodRevision:hodRevision, hodRejected:hodRejected,
    hoiApproved:hoiApproved, hoiRevision:hoiRevision, hoiRejected:hoiRejected
  };
}

// ─── IN-APP NOTIFICATIONS ─────────────────────────────────────────────────────
// getNotifications(role)              — newest notifications for a role
// markNotifRead(notifID)              — mark one as read
// markAllRead(role)                   — mark all as read for a role
// _pushNotif(role,type,title,body,sid,facultyName) — internal writer

// _scopeKey_ — derive the per-row "ForRole" key used by both readers and writers.
// Notifications are scoped so that:
//   • FACULTY    → 'FACULTY:<email>'             (per-faculty)
//   • HOD        → 'HOD:<department>'            (per-department)
//   • HOI        → 'HOI:<instCode>'              (per-institution; falls back to 'HOI' if code missing)
//   • IMO        → 'IMO'                         (single Chancellor's office, intentionally global)
// scopeArg is the email for FACULTY, the department for HOD, the institution code for HOI.
function _scopeKey_(role, scopeArg) {
  var r = String(role||'').toUpperCase();
  if (r === 'FACULTY' && scopeArg) return 'FACULTY:' + String(scopeArg).trim().toUpperCase();
  if (r === 'HOD'     && scopeArg) return 'HOD:'     + String(scopeArg).trim();
  if (r === 'HOI'     && scopeArg) return 'HOI:'     + String(scopeArg).trim().toUpperCase();
  return r;
}

function getNotifications(role, facultyID, scopeArg) {
  if (!role) throw new Error('Role required.');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.NOTIF);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues(), h = data[0];
  // Backward-compatible parameter shape:
  //   FACULTY: facultyID is the email
  //   HOD:     facultyID is the department (legacy positional arg)
  //   HOI:     scopeArg is the institution code (new); facultyID is unused
  //   IMO:     no scope
  var matchKey;
  if (role === 'FACULTY' && facultyID) {
    matchKey = _scopeKey_('FACULTY', facultyID);
  } else if (role === 'HOD' && facultyID) {
    matchKey = _scopeKey_('HOD', facultyID);
  } else if (role === 'HOI' && (scopeArg || facultyID)) {
    // Accept the institution code as either positional arg for back-compat
    matchKey = _scopeKey_('HOI', scopeArg || facultyID);
  } else {
    matchKey = role;
  }
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var forRole = String(row[h.indexOf('ForRole')] || '');
    // Match exact scoped key, plus fall through to legacy unscoped 'HOD' / 'HOI'
    // rows so notifications written before this scoping change still surface.
    var matches = forRole === matchKey ||
      (role === 'HOD' && forRole === 'HOD') ||
      (role === 'HOI' && forRole === 'HOI');
    if (!matches) continue;
    out.push({
      notifID:     String(row[h.indexOf('NotifID')]      || ''),
      type:        String(row[h.indexOf('Type')]         || ''),
      title:       String(row[h.indexOf('Title')]        || ''),
      body:        String(row[h.indexOf('Body')]         || ''),
      submissionID:String(row[h.indexOf('SubmissionID')] || ''),
      facultyName: String(row[h.indexOf('FacultyName')]  || ''),
      isRead:      String(row[h.indexOf('IsRead')]       || '') === 'YES',
      createdAt:   _fmtDT(row[h.indexOf('CreatedAt')])
    });
  }
  return out.reverse();
}

function markNotifRead(notifID) {
  if (!notifID) return { ok: false };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.NOTIF);
  if (!sheet) return { ok: false };
  var data = sheet.getDataRange().getValues(), h = data[0];
  var kI = h.indexOf('NotifID'), rI = h.indexOf('IsRead');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][kI]) === String(notifID)) {
      sheet.getRange(i + 1, rI + 1).setValue('YES');
      return { ok: true };
    }
  }
  return { ok: false };
}

function markAllRead(role, scopeArg) {
  if (!role) return { ok: false };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.NOTIF);
  if (!sheet) return { ok: false };
  var data = sheet.getDataRange().getValues(), h = data[0];
  var frI = h.indexOf('ForRole'), rI = h.indexOf('IsRead');
  // scopeArg = department (HOD) | instCode (HOI)
  var matchKey =
    (role === 'HOD' && scopeArg) ? _scopeKey_('HOD', scopeArg) :
    (role === 'HOI' && scopeArg) ? _scopeKey_('HOI', scopeArg) :
    role;
  for (var i = 1; i < data.length; i++) {
    var fr = String(data[i][frI]).trim();
    var matches = fr === matchKey ||
      (role === 'HOD' && fr === 'HOD') ||
      (role === 'HOI' && fr === 'HOI');
    if (matches && String(data[i][rI]) !== 'YES') {
      sheet.getRange(i + 1, rI + 1).setValue('YES');
    }
  }
  return { ok: true };
}

function clearNotifications(role, scopeArg) {
  if (!role) return { ok: false };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.NOTIF);
  if (!sheet) return { ok: false };
  var data = sheet.getDataRange().getValues(), h = data[0];
  var frI = h.indexOf('ForRole');
  var matchKey =
    (role === 'HOD' && scopeArg) ? _scopeKey_('HOD', scopeArg) :
    (role === 'HOI' && scopeArg) ? _scopeKey_('HOI', scopeArg) :
    role;
  for (var i = data.length - 1; i >= 1; i--) {
    var fr = String(data[i][frI]);
    if (fr === matchKey ||
        (role === 'HOD' && fr === 'HOD') ||
        (role === 'HOI' && fr === 'HOI')) sheet.deleteRow(i + 1);
  }
  return { ok: true };
}

function _pushNotif(forRole, type, title, body, sid, facultyName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SH.NOTIF);
    // Auto-create Notifications sheet if it doesn't exist yet
    if (!sheet) {
      sheet = ss.insertSheet(SH.NOTIF);
      var hdrs = SCHEMA[SH.NOTIF];
      var hr = sheet.getRange(1,1,1,hdrs.length);
      hr.setValues([hdrs]);
      hr.setFontWeight('bold').setBackground('#1a3c5e').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
    var nid = 'N-' + new Date().getTime() + '-' + Math.random().toString(36).slice(2,5).toUpperCase();
    sheet.appendRow([nid, forRole, type, title, body, sid || '', facultyName || '', 'NO', new Date()]);
  } catch(e) { Logger.log('_pushNotif failed: ' + e.message); }
}

// ─── FRIDAY REMINDERS ────────────────────────────────────────────────────────
function sendFridayReminders() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var facD=ss.getSheetByName(SH.FACULTY).getDataRange().getValues(),fh=facD[0];
  // The Faculty_Master schema uses 'Email' as the sole identifier. The previous
  // implementation expected a separate 'FacultyEmail' column that doesn't exist,
  // so facID was always '' and the "already-submitted-this-week" skip never matched.
  var eI=fh.indexOf('Email'), nI=fh.indexOf('FacultyName'), stI=fh.indexOf('Status');
  var subD=ss.getSheetByName(SH.SUBMISSION).getDataRange().getValues(),sh=subD[0];
  var sfI=sh.indexOf('FacultyEmail'); if (sfI < 0) sfI = 1;  // standard pos
  var sdI=sh.indexOf('ReportingFrom');
  var today=new Date(),weekAgo=new Date(today); weekAgo.setDate(today.getDate()-6);
  var done={};
  for(var i=1;i<subD.length;i++){
    if(new Date(subD[i][sdI])>=weekAgo) {
      done[String(subD[i][sfI]||'').trim().toLowerCase()] = true;
    }
  }
  var dateStr=Utilities.formatDate(today,Session.getScriptTimeZone(),'dd-MMM-yyyy');
  var sent=0;
  for(var j=1;j<facD.length;j++){
    var email = String(facD[j][eI]||'').trim();
    if (!email) continue;
    var status = (stI >= 0) ? String(facD[j][stI]||'').trim() : 'Active';
    if (status && status !== 'Active') continue;
    if (done[email.toLowerCase()]) continue;
    try{
      MailApp.sendEmail({to:email,subject:'VMRF: Weekly Performance Submission Due Today',
        body:'Dear '+facD[j][nI]+',\n\nYour Weekly Performance Submission for the week ending '+dateStr+' is due TODAY by 3:00 PM.\n\nPlease open the VMRF Institutional Monitoring System and submit your report.\n\n— Institutional Management Office, VMRF-DU'});
      sent++;
    }catch(e){Logger.log('Email failed for '+email+': '+e.message);}
  }
  return { sent:sent };
}

// ─── EMAIL HELPERS ────────────────────────────────────────────────────────────
function _notifyHOD(ss,sid,facultyEmail) {
  // Look up faculty to get their department
  var fac = _rowByKey(SH.FACULTY, String(facultyEmail||'').trim().toLowerCase(), 'Email');
  if (!fac) return;
  var dept = String(fac['Department']||'').trim();
  if (!dept) return;
  // Find the HOD for this department in Staff_Master
  var staffSh = ss.getSheetByName(SH.STAFF); if (!staffSh) return;
  var staffD = staffSh.getDataRange().getValues(), sh = staffD[0];
  var emI = sh.indexOf('Email'), rlI = sh.indexOf('Role');
  var depI = sh.indexOf('Department'), stI = sh.indexOf('Status');
  var hodEmail = null;
  for (var i = 1; i < staffD.length; i++) {
    if (String(staffD[i][rlI]||'').toUpperCase() !== 'HOD') continue;
    if (String(staffD[i][stI]||'') !== 'Active') continue;
    if (String(staffD[i][depI]||'').trim().toLowerCase() === dept.toLowerCase()) {
      hodEmail = String(staffD[i][emI]||'').trim();
      break;
    }
  }
  if (!hodEmail) return;
  // Look up submission period so email is informative
  var subRow = _rowByKey(SH.SUBMISSION, sid) || {};
  var _pFrom = subRow['ReportingFrom'] ? _fmt(subRow['ReportingFrom']) : '';
  var _pTo   = subRow['ReportingTo']   ? _fmt(subRow['ReportingTo'])   : '';
  var periodStr = (_pFrom && _pTo) ? _pFrom + ' to ' + _pTo : '';
  try {
    MailApp.sendEmail({
      to: hodEmail,
      subject: '[VMRF] New Submission Pending HOD Review — ' + String(fac['FacultyName']||'Faculty'),
      body: 'A new faculty weekly report is pending your review.\n\n' +
            'Faculty: '    + String(fac['FacultyName']||'') + '\n' +
            'Department: ' + dept + '\n' +
            (periodStr ? 'Period: ' + periodStr + '\n' : '') + '\n' +
            'Please login to the VMRF Institutional Monitoring System to review.\n\n— IMO, VMRF-DU'
    });
  } catch(e) { Logger.log('HOD email failed: ' + e.message); }
}

function _notifyHOI(ss,sid) {
  var to=_prop('HOI_DEFAULT'); if(!to) return;
  var sub=_rowByKey(SH.SUBMISSION,sid)||{};
  var fid=_getFidFromSub(sub);
  var fac=fid?(_rowByKey(SH.FACULTY,fid,'Email')||{}):{};
  var facName=String(fac['FacultyName']||fid||'A faculty member');
  MailApp.sendEmail({to:to,subject:'[VMRF] Submission Approved by HOD — HOI Review Required',
    body:'A submission by '+facName+' has been approved by the Head of Department and is now pending your review.\n\nPlease login to the VMRF Institutional Monitoring System.\n\n— IMO, VMRF-DU'});
}

function _notifyIMO(ss,sid) {
  var to=_prop('IMO_EMAIL'); if(!to) return;
  var sub=_rowByKey(SH.SUBMISSION,sid)||{};
  var fid=_getFidFromSub(sub);
  var fac=fid?(_rowByKey(SH.FACULTY,fid,'Email')||{}):{};
  var facName=String(fac['FacultyName']||fid||'A faculty member');
  MailApp.sendEmail({to:to,subject:'[VMRF] Submission Ready for IMO Monitoring',
    body:'A submission by '+facName+' has been approved by both HOD and HOI. Ready for final IMO monitoring.\n\n— Automated, VMRF-DU'});
}

function _notifyRevision(ss,sid,reviewer,remark) {
  var sub=_rowByKey(SH.SUBMISSION,sid); if(!sub) return;
  var _nrFid=_getFidFromSub(sub);
  var fac=_rowByKey(SH.FACULTY,_nrFid,'Email'); if(!fac||!fac['Email']) return;
  MailApp.sendEmail({to:String(fac['Email']),subject:'[VMRF] Your Weekly Report Requires Revision',
    body:'Dear '+String(fac['FacultyName']||'Faculty')+',\n\nYour submission has been returned for revision by your '+reviewer+'.\n\nRemark:\n'+remark+'\n\nPlease make the necessary changes and resubmit.\n\n— IMO, VMRF-DU'});
}

function _notifyFinalStatus(ss,sid,status,remark) {
  var sub=_rowByKey(SH.SUBMISSION,sid); if(!sub) return;
  var _nfFid=_getFidFromSub(sub);
  var fac=_rowByKey(SH.FACULTY,_nfFid,'Email'); if(!fac||!fac['Email']) return;
  MailApp.sendEmail({to:String(fac['Email']),subject:'[VMRF] Your Weekly Report Has Been '+status+' by IMO',
    body:'Dear '+String(fac['FacultyName']||'Faculty')+',\n\nYour submission has been marked "'+status+'" by the Institutional Management Office.\n\n'+(remark?'IMO Note:\n'+remark+'\n\n':'')+' — IMO, VMRF-DU'});
}

// ─── IMO QUEUE FOR HOD SUBMISSIONS (HOD→HOI→IMO) ─────────────────────────────
function getIMOQueueHOD() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var subSh = ss.getSheetByName(SH.HOD_SUB); if (!subSh) return [];
  var subD  = subSh.getDataRange().getValues(), subH = subD[0];
  var revSh = ss.getSheetByName(SH.HOD_REVIEW);
  var revD  = revSh ? revSh.getDataRange().getValues() : [['SubmissionID']]; var revH = revD[0];
  var imoSh = ss.getSheetByName(SH.HOD_IMO);
  var imoD  = imoSh ? imoSh.getDataRange().getValues() : [['SubmissionID']]; var imoH = imoD[0];
  var saSh  = ss.getSheetByName(SH.HOD_SA);
  var saD   = saSh ? saSh.getDataRange().getValues() : [['SubmissionID']]; var saH = saD[0];
  var tsSh  = ss.getSheetByName(SH.HOD_TS);
  var tsD   = tsSh ? tsSh.getDataRange().getValues() : [['SubmissionID']];

  // Build HOI review map
  var revMap = {};
  for (var r2 = 1; r2 < revD.length; r2++) {
    revMap[String(revD[r2][0]||'')] = { status: String(revD[r2][revH.indexOf('HOI_Status')]||''), remark: String(revD[r2][revH.indexOf('HOI_Remark')]||'') };
  }
  // Build IMO review map
  var imoMap = {};
  var _imoStI = imoH.indexOf('IMO_Status');
  for (var im = 1; im < imoD.length; im++) {
    imoMap[String(imoD[im][0]||'')] = String(imoD[im][_imoStI >= 0 ? _imoStI : 2]||'');
  }
  // Build SA map
  var saMap2 = {};
  for (var s2 = 1; s2 < saD.length; s2++) {
    saMap2[String(saD[s2][0]||'')] = { tasks: String(saD[s2][saH.indexOf('Tasks')]||''), target: String(saD[s2][saH.indexOf('TargetPlanNextWeek')]||'') };
  }
  // Build timesheet map
  var tsMap2 = {};
  for (var t2 = 1; t2 < tsD.length; t2++) {
    var k = String(tsD[t2][0]||'');
    if (!tsMap2[k]) tsMap2[k] = [];
    tsMap2[k].push({
      Day:            String(tsD[t2][1]||''),
      TimeSlot:       String(tsD[t2][2]||''),
      ActivityType:   String(tsD[t2][3]||''),
      Details:        String(tsD[t2][4]||''),
      AttachmentURL:  String(tsD[t2][5]||''),
      AttachmentName: String(tsD[t2][6]||'')
    });
  }

  // Look up HOD staff names, campus and institution
  var staffSh = ss.getSheetByName(SH.STAFF);
  var staffMap = {};
  if (staffSh) {
    var staffD = staffSh.getDataRange().getValues(), staffH = staffD[0];
    var _sEmI  = staffH.indexOf('Email'),       _sNmI = staffH.indexOf('StaffName');
    var _sDepI = staffH.indexOf('Department'),  _sCaI = staffH.indexOf('Campus');
    var _sInI  = staffH.indexOf('Institution');
    for (var sf = 1; sf < staffD.length; sf++) {
      var sEmail = String(staffD[sf][_sEmI]||'').trim().toLowerCase();
      if (sEmail) staffMap[sEmail] = {
        name:        String(staffD[sf][_sNmI]||''),
        dept:        String(staffD[sf][_sDepI]||''),
        campus:      _sCaI >= 0 ? String(staffD[sf][_sCaI]||'').trim() : '',
        institution: _sInI >= 0 ? String(staffD[sf][_sInI]||'').trim() : ''
      };
    }
  }

  var out = [];
  for (var i = 1; i < subD.length; i++) {
    var r   = subD[i];
    var sid = String(r[0]||'').trim();
    var rev = revMap[sid] || { status:'', remark:'' };
    if (rev.status !== 'Approved') continue; // only HOI-approved
    var imoSt = imoMap[sid] || '';
    if (imoSt === 'Finalised') continue; // already done
    var sa  = saMap2[sid] || {};
    var hodID = String(r[subH.indexOf('HOD_ID')]||'').trim().toLowerCase();
    var staff = staffMap[hodID] || {};
    out.push({
      submissionID: sid,
      hodID:        hodID,
      facultyName:  staff.name || 'HOD',
      department:   staff.dept || '',
      campus:       staff.campus || '',
      institution:  staff.institution || '',
      semester:     String(r[subH.indexOf('AcademicYearSemester')]||''),
      from:         _fmt(r[subH.indexOf('ReportingFrom')]),
      to:           _fmt(r[subH.indexOf('ReportingTo')]),
      submitted:    _fmtDT(r[subH.indexOf('SubmittedDateTime')]),
      outcome:      sa.tasks   || '',
      target:       sa.target  || '',
      hoiStatus:    rev.status,
      hoiRemark:    rev.remark,
      imoStatus:    imoSt || 'Pending',
      timesheet:    tsMap2[sid] || [],
      _type:        'HOD'
    });
  }
  return out.reverse();
}

function submitIMOHODReview(sid, remark, status) {
  if (!sid)    throw new Error('Submission ID missing.');
  if (!status) throw new Error('Please select a final status.');
  // Verify HOI approved
  var revRow = _rowByKey(SH.HOD_REVIEW, sid);
  if (!revRow || String(revRow['HOI_Status']||'') !== 'Approved')
    throw new Error('HOI must approve this HOD submission before IMO can finalise it.');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH.HOD_IMO);
  if (!sh) throw new Error('HOD_IMO_Review sheet not found. Please re-run initializeSystem.');
  var data = sh.getDataRange().getValues(), h = data[0];
  var idI = h.indexOf('SubmissionID'), stI = h.indexOf('IMO_Status'), rmI = h.indexOf('IMO_Remark'), dtI = h.indexOf('IMO_DateTime');
  var now = new Date();
  var found = false;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idI]) === String(sid)) {
      sh.getRange(i+1, stI+1).setValue(status);
      sh.getRange(i+1, rmI+1).setValue(remark || '');
      sh.getRange(i+1, dtI+1).setValue(now);
      found = true;
      break;
    }
  }
  if (!found) sh.appendRow([sid, remark||'', status, now]);

  var emoji = status==='Finalised'?'✅':status==='Escalated'?'⚡':'ℹ️';
  // Resolve the HOD's department so this decision only goes to the right HOD,
  // not every HOD in the system.
  var hodDeptIMO = '';
  try {
    var hodSubShI = ss.getSheetByName(SH.HOD_SUB);
    if (hodSubShI) {
      var hsdI = hodSubShI.getDataRange().getValues(), hshI = hsdI[0];
      var _hII = hshI.indexOf('HOD_ID');
      var hEmailI = '';
      for (var hxI = 1; hxI < hsdI.length; hxI++) {
        if (String(hsdI[hxI][0]||'').trim() === String(sid).trim()) {
          hEmailI = String(hsdI[hxI][_hII >= 0 ? _hII : 1]||'').trim().toLowerCase();
          break;
        }
      }
      if (hEmailI) {
        var hodFacI = _rowByKey(SH.FACULTY, hEmailI, 'Email') || {};
        hodDeptIMO = String(hodFacI['Department']||'').trim();
      }
    }
  } catch(e) { Logger.log('IMO HOD dept lookup failed: '+e.message); }
  var hodKeyIMO = hodDeptIMO ? 'HOD:'+hodDeptIMO : 'HOD';
  _pushNotif(hodKeyIMO, 'imo_hod_decision', emoji+' IMO Decision on Your Timesheet: '+status,
    'Your HOD weekly report has been marked "'+status+'" by the Institutional Management Office.'+(remark?' Note: '+remark:''),
    sid, 'IMO');
  return { ok:true };
}

// ═════════════════════════════════════════════════════════════════════════════
// COMPARISON REPORT — HOD / HOI / IMO
// ═════════════════════════════════════════════════════════════════════════════
// Aggregates submission metrics grouped by a chosen dimension so that users
// can compare performance across entities. The `dimension` controls grouping:
//
//   dimension='faculty'    → one row per faculty (email + name + dept + campus/inst)
//   dimension='department' → one row per department
//   dimension='institution'→ one row per institution
//   dimension='campus'     → one row per campus
//
// The `filter` argument narrows the universe before grouping. Role-appropriate
// narrowing is the client's responsibility but the server accepts any of:
//   { campus, institution, department }
//
// Per-row metrics returned:
//   submissions, totalFaculty, finalised (= imoFinal), pendingHOD/HOI/IMO,
//   rejected, revision,
//   complianceRate = submissions / (totalFaculty * expectedWeeks) × 100, capped 100
//   imoApprovalRate = imoFinal / submissions × 100        (fully approved)
//   hoiApprovalRate = (pendingIMO + imoFinal) / submissions × 100 (cleared HOI)
//   rejectionRate   = rejected / submissions × 100
//   onTimeRate      = onTime  / submissions × 100   (≤ Friday 15:00 of reporting week)
//   approvalRate    = hoiApprovalRate            (legacy alias; kept for UI compat)
//   avgSubsPerFac   = submissions / totalFaculty (legacy; kept for UI compat)
//   prevSubmissions = submissions in the equivalent previous period (null if 'all')
//   trendDelta      = submissions - prevSubmissions (null if 'all')
//   trendPct        = (submissions − prev) / prev × 100  (null if prev is 0 or 'all')
//   compositeScore  = weighted 0–100 score (A11). Default weights:
//                       compliance:0.45, onTime:0.25, imoApproval:0.20, rejInv:0.10
//                     Override via Script Property COMPARISON_SCORE_WEIGHTS.
//   scoreBreakdown  = {compliance, onTime, imoApproval, rejectionInverted} —
//                     each component's weighted contribution to compositeScore
//                     (null when that component had no data for this row).
//   avgHodTurnaroundHrs = mean hours from faculty submit → HOD review (D2)
//   avgHoiTurnaroundHrs = mean hours from HOD review → HOI review (or faculty
//                         submit → HOI review for no-HOD faculty)
//   avgImoTurnaroundHrs = mean hours from HOI review → IMO finalisation
//   avgEndToEndHrs      = mean hours from faculty submit → IMO Approved
//   slaCounts           = {hod, hoi, imo, end} — sample sizes behind the means
function getComparisonReport(dimension, filter, period) {
  dimension = String(dimension||'faculty').toLowerCase();
  if (['faculty','department','institution','campus'].indexOf(dimension) < 0) {
    throw new Error('Invalid dimension: ' + dimension);
  }
  filter = filter || {};
  period = String(period || 'all').toLowerCase();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Period date boundaries ──────────────────────────────────────────────
  // Tier-4 / A7: in addition to week/month/year/all, the period now accepts
  //   'semester' or 'current_semester' — the in-progress academic semester
  //   'prev_semester'                  — the immediately preceding semester
  //   'custom'                         — uses filter.from / filter.to (ISO dates)
  // Indian academic default: Odd semester Aug–Dec, Even semester Jan–May.
  // Override via Script Property SEMESTER_CONFIG (JSON), see _resolveSemester_.
  var now = new Date();
  var periodStart = null, periodEnd = null;
  if (period === 'week') {
    var dow = now.getDay(); // 0=Sun
    var daysToMon = (dow === 0) ? -6 : (1 - dow);
    periodStart = new Date(now);
    periodStart.setDate(now.getDate() + daysToMon);
    periodStart.setHours(0, 0, 0, 0);
    periodEnd = new Date(periodStart);
    periodEnd.setDate(periodStart.getDate() + 6);
    periodEnd.setHours(23, 59, 59, 999);
  } else if (period === 'month') {
    periodStart = new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0, 0);
    periodEnd   = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59, 999);
  } else if (period === 'year') {
    periodStart = new Date(now.getFullYear(), 0, 1, 0, 0, 0, 0);
    periodEnd   = new Date(now.getFullYear(), 11, 31, 23, 59, 59, 999);
  } else if (period === 'semester' || period === 'current_semester') {
    var sem = _resolveSemester_(now, /*offset=*/0);
    periodStart = sem.start; periodEnd = sem.end;
  } else if (period === 'prev_semester') {
    var psem = _resolveSemester_(now, /*offset=*/-1);
    periodStart = psem.start; periodEnd = psem.end;
  } else if (period === 'custom') {
    var cFrom = filter.from ? new Date(filter.from) : null;
    var cTo   = filter.to   ? new Date(filter.to)   : null;
    if (cFrom && !isNaN(cFrom.getTime())) {
      periodStart = new Date(cFrom); periodStart.setHours(0,0,0,0);
    }
    if (cTo && !isNaN(cTo.getTime())) {
      periodEnd = new Date(cTo); periodEnd.setHours(23,59,59,999);
    }
    // If either bound is missing, leave it null (treated as open-ended that side)
  }

  // ── Previous-period boundaries — used to compute trendPct / trendDelta ──
  // 'week'      → prior Mon-Sun (current week shifted back 7 days)
  // 'month'     → prior calendar month
  // 'year'      → prior calendar year
  // 'semester'  → previous semester (A7)
  // 'custom'    → window of equivalent length immediately before periodStart
  // 'all'       → no prev (trend is N/A)
  var prevPeriodStart = null, prevPeriodEnd = null;
  if (period === 'week' && periodStart && periodEnd) {
    prevPeriodStart = new Date(periodStart); prevPeriodStart.setDate(periodStart.getDate() - 7);
    prevPeriodEnd   = new Date(periodEnd);   prevPeriodEnd.setDate(periodEnd.getDate() - 7);
  } else if (period === 'month' && periodStart) {
    prevPeriodStart = new Date(periodStart.getFullYear(), periodStart.getMonth() - 1, 1, 0, 0, 0, 0);
    prevPeriodEnd   = new Date(periodStart.getFullYear(), periodStart.getMonth(), 0, 23, 59, 59, 999); // last day of prev month
  } else if (period === 'year' && periodStart) {
    prevPeriodStart = new Date(periodStart.getFullYear() - 1, 0, 1, 0, 0, 0, 0);
    prevPeriodEnd   = new Date(periodStart.getFullYear() - 1, 11, 31, 23, 59, 59, 999);
  } else if ((period === 'semester' || period === 'current_semester') && periodStart) {
    var prevSem = _resolveSemester_(now, /*offset=*/-1);
    prevPeriodStart = prevSem.start; prevPeriodEnd = prevSem.end;
  } else if (period === 'prev_semester' && periodStart) {
    var prevPrevSem = _resolveSemester_(now, /*offset=*/-2);
    prevPeriodStart = prevPrevSem.start; prevPeriodEnd = prevPrevSem.end;
  } else if (period === 'custom' && periodStart && periodEnd) {
    // Equivalent prior window of the same length, immediately preceding periodStart
    var lengthMs = periodEnd.getTime() - periodStart.getTime();
    prevPeriodEnd   = new Date(periodStart.getTime() - 1);
    prevPeriodStart = new Date(prevPeriodEnd.getTime() - lengthMs);
  }

  // ── Tier-5 / D4: year-ago boundaries — same window, exactly 1 year earlier
  // Used for year-on-year comparison ("Q4 this year vs Q4 last year"). Only
  // meaningful when the current window is bounded; for 'all' we leave them null.
  var yearAgoStart = null, yearAgoEnd = null;
  if (periodStart && periodEnd) {
    yearAgoStart = new Date(periodStart); yearAgoStart.setFullYear(periodStart.getFullYear() - 1);
    yearAgoEnd   = new Date(periodEnd);   yearAgoEnd.setFullYear(periodEnd.getFullYear() - 1);
  }

  var facD = ss.getSheetByName(SH.FACULTY).getDataRange().getValues();
  var facH = facD[0].map(function(v){return String(v).trim();});
  var subD = ss.getSheetByName(SH.SUBMISSION).getDataRange().getValues();
  var hodD = ss.getSheetByName(SH.HOD).getDataRange().getValues();
  var hoiD = ss.getSheetByName(SH.HOI).getDataRange().getValues();
  var imoD = ss.getSheetByName(SH.IMO).getDataRange().getValues();

  var facMap = _buildFacMap(facD, ss.getSheetByName(SH.STAFF).getDataRange().getValues());
  var sidEmailMap = _buildSidEmailMap(subD);
  var noHodSids   = _buildNoHodSidSet_(sidEmailMap, facMap);

  // Review status maps keyed by submission ID. Each value is an object
  // {status, dateTime} so we can compute both the review chain (Tier 1+2) and
  // the review-turnaround SLA (Tier 3 / D2). dateTime may be null if the row
  // hasn't been reviewed yet.
  var hodMap = {}, hoiMap = {}, imoMap = {};
  var hodStI = hodD[0].indexOf('HOD_Status'); if (hodStI < 0) hodStI = 2;
  var hoiStI = hoiD[0].indexOf('HOI_Status'); if (hoiStI < 0) hoiStI = 2;
  var imoStI = imoD[0].indexOf('IMO_Status'); if (imoStI < 0) imoStI = 2;
  var hodDtI = hodD[0].indexOf('HOD_DateTime'); if (hodDtI < 0) hodDtI = 3;
  var hoiDtI = hoiD[0].indexOf('HOI_DateTime'); if (hoiDtI < 0) hoiDtI = 3;
  var imoDtI = imoD[0].indexOf('IMO_DateTime'); if (imoDtI < 0) imoDtI = 3;
  for (var h2=1; h2<hodD.length; h2++) {
    var s2=String(hodD[h2][0]||''); if(!s2) continue;
    var _hdt = hodD[h2][hodDtI];
    hodMap[s2] = { status: String(hodD[h2][hodStI]||''), dateTime: (_hdt instanceof Date) ? _hdt : null };
  }
  for (var h3=1; h3<hoiD.length; h3++) {
    var s3=String(hoiD[h3][0]||''); if(!s3) continue;
    var _idt = hoiD[h3][hoiDtI];
    hoiMap[s3] = { status: String(hoiD[h3][hoiStI]||''), dateTime: (_idt instanceof Date) ? _idt : null };
  }
  for (var h4=1; h4<imoD.length; h4++) {
    var s4=String(imoD[h4][0]||''); if(!s4) continue;
    var _xdt = imoD[h4][imoDtI];
    imoMap[s4] = { status: String(imoD[h4][imoStI]||''), dateTime: (_xdt instanceof Date) ? _xdt : null };
  }

  // Filters — A6: each filter dimension accepts either a single string OR an
  // array of strings. The single-string form is preserved for backwards compat
  // with all existing callers; arrays unlock side-by-side comparisons like
  // "Engineering vs Pharmacy in VMCC". Empty string / empty array → no filter.
  function _normFilter_(v) {
    if (Array.isArray(v)) return v.map(function(s){return String(s||'').trim();}).filter(Boolean);
    var s = String(v||'').trim();
    return s ? [s] : [];
  }
  var fCampusArr = _normFilter_(filter.campus);
  var fInstArr   = _normFilter_(filter.institution);
  var fDeptArr   = _normFilter_(filter.department).map(function(s){return s.toLowerCase();});
  var fDeptArrB  = fDeptArr.map(function(s){return s.replace(/\s*\(pg\)\s*$/,'');});

  function _facMatchesFilter(f) {
    if (fCampusArr.length && fCampusArr.indexOf(String(f['Campus']||'').trim()) < 0) return false;
    if (fInstArr.length   && fInstArr.indexOf(String(f['Institution']||'').trim()) < 0) return false;
    if (fDeptArr.length) {
      var dept  = String(f['Department']||'').trim().toLowerCase();
      var deptB = dept.replace(/\s*\(pg\)\s*$/,'');
      if (fDeptArr.indexOf(dept) < 0 && fDeptArrB.indexOf(deptB) < 0) return false;
    }
    return true;
  }

  // Legacy single-value vars used by some downstream blocks (response shape) —
  // keep them as the FIRST element of each array so the response still looks
  // sensible to callers that pass single values.
  var fCampus = fCampusArr[0] || '';
  var fInst   = fInstArr[0]   || '';

  // Key builder per dimension — returns the group key for a faculty row, or
  // an empty string to skip this faculty from the report entirely.
  function _groupKey(f) {
    if (dimension === 'faculty')    return String(f['Email']||'').trim().toLowerCase();
    if (dimension === 'department') return String(f['Department']||'').trim();
    if (dimension === 'institution')return String(f['Institution']||'').trim();
    if (dimension === 'campus')     return String(f['Campus']||'').trim();
    return '';
  }
  function _groupMeta(f, key) {
    if (dimension === 'faculty') {
      return {
        key: key,
        label: String(f['FacultyName']||'')||key,
        subtitle: [f['Department'], f['Institution'], f['Campus']].filter(Boolean).join(' · '),
        // Expose department as a top-level field so the client picker can
        // filter faculty cards by department for cross-department comparison.
        department: String(f['Department']||'')
      };
    }
    if (dimension === 'department')  return { key:key, label:key, subtitle: [f['Institution'],f['Campus']].filter(Boolean).join(' · ') };
    if (dimension === 'institution') return { key:key, label:key, subtitle: String(f['Campus']||'') };
    if (dimension === 'campus')      return { key:key, label:key, subtitle: '' };
    return { key:key, label:key, subtitle:'' };
  }

  // Seed buckets with every in-scope faculty so that groups with zero
  // submissions still appear (important for the "no one submitted" comparison).
  var buckets = {}; // key → {key,label,subtitle,totalFaculty,submissions,finalised,pendingHOD,pendingHOI,pendingIMO,rejected,revision,prevSubmissions,onTime,late,latestSubmittedDate,hodTurnaround*,hoiTurnaround*,imoTurnaround*,endToEnd*}
  function _ensure(key, meta) {
    if (!buckets[key]) {
      buckets[key] = {
        key: meta.key, label: meta.label, subtitle: meta.subtitle,
        // (Option A) Carry the faculty's department through to the row so the
        // front-end picker can group cards by department and the comp-table
        // header can show a coloured dept badge. Empty string for non-faculty
        // dimensions where meta.department is undefined.
        department: meta.department || '',
        totalFaculty: 0, submissions: 0,
        finalised: 0, pendingHOD: 0, pendingHOI: 0, pendingIMO: 0,
        rejected: 0, revision: 0,
        // Tier-2 additions
        prevSubmissions: 0,         // submissions count in the equivalent previous period
        onTime: 0, late: 0,         // current-period submissions classified by Friday-15:00 deadline
        // Tier-3 additions (D2): SLA turnaround running sums.
        // Each *N is a sample count and *SumMs is the total elapsed time across
        // those samples. We compute means in the finalisation block.
        hodTurnaroundSumMs: 0, hodTurnaroundN: 0,
        hoiTurnaroundSumMs: 0, hoiTurnaroundN: 0,
        imoTurnaroundSumMs: 0, imoTurnaroundN: 0,
        endToEndSumMs: 0,     endToEndN: 0,
        // Tier-5 / A5: parallel HOD weekly-submission tracking. Distinct from
        // the faculty submissions above because HODs file their OWN reports.
        // Only meaningful at department/institution/campus dimensions.
        hodTotal: 0, hodSubmissions: 0,
        // Tier-5 / D4: year-ago counter (same window, 1 year earlier)
        yearAgoSubmissions: 0,
        latestSubmittedDate: null,
        categorySlots: {}   // top-level category → filled slot count
      };
    }
    return buckets[key];
  }

  // Pass 1: enumerate faculty rows to build buckets + count faculty per group
  for (var fi=1; fi<facD.length; fi++) {
    var fd = facD[fi];
    var f = {};
    for (var hi2=0; hi2<facH.length; hi2++) f[facH[hi2]] = fd[hi2];
    // Only count Active faculty — faculty Pending/Inactive shouldn't affect rates
    var st = String(f['Status']||'').trim();
    if (st && st !== 'Active') continue;
    if (!_facMatchesFilter(f)) continue;
    var key = _groupKey(f);
    if (!key) continue;
    var meta = _groupMeta(f, key);
    var b = _ensure(key, meta);
    b.totalFaculty++;
  }

  // Pass 2: walk submissions, assign each to its group via facMap
  var subDateI = subD[0].indexOf('SubmittedDateTime'); if (subDateI < 0) subDateI = 6;
  // ReportingFrom is the Monday of the reporting week — drives Friday-3PM deadline.
  var subRfI = subD[0].indexOf('ReportingFrom'); if (subRfI < 0) subRfI = 3;
  var earliestSubDate = null;       // tracked across whole dataset for 'all'-period compliance baseline
  for (var si=1; si<subD.length; si++) {
    var sid = String(subD[si][0]||'').trim();
    if (!sid) continue;
    // Track earliest submission date globally — used by _resolveExpectedWeeks_
    var _sd = subD[si][subDateI];
    if (_sd instanceof Date && (!earliestSubDate || _sd.getTime() < earliestSubDate.getTime())) {
      earliestSubDate = _sd;
    }
    var email = sidEmailMap[sid] || '';
    var fac = email ? (facMap[email]||{}) : {};
    // Promote fac to use same canonical keys as facD rows
    var f2 = {
      FacultyName: fac['FacultyName']||'',
      Email:       fac['Email']||email,
      Department:  fac['Department']||'',
      Campus:      fac['Campus']||'',
      Institution: fac['Institution']||''
    };
    if (!_facMatchesFilter(f2)) continue;
    var key2 = _groupKey(f2);
    if (!key2) continue;
    var b2 = _ensure(key2, _groupMeta(f2, key2));

    // ── Period routing — which window does this submission fall into? ─────
    var subDate2 = subD[si][subDateI];
    var inCurrent, inPrev = false, inYearAgo = false;
    if (!periodStart) {
      // 'all' period — every submission with a valid date is "current",
      // and there is no previous-period concept (trend metrics will be null).
      inCurrent = (subDate2 instanceof Date);
    } else if (subDate2 instanceof Date) {
      inCurrent = (subDate2 >= periodStart && subDate2 <= periodEnd);
      inPrev    = !!prevPeriodStart && (subDate2 >= prevPeriodStart && subDate2 <= prevPeriodEnd);
      inYearAgo = !!yearAgoStart && (subDate2 >= yearAgoStart && subDate2 <= yearAgoEnd);
    } else {
      // No date stamp — can't classify, drop entirely from period-bound counts
      inCurrent = false;
    }

    if (inCurrent) b2.submissions++;
    if (inPrev)    b2.prevSubmissions++;
    if (inYearAgo) b2.yearAgoSubmissions++;

    // Status routing, on-time, latest-date are CURRENT-PERIOD only.
    // Previous-period submissions only feed prevSubmissions for trend calc.
    if (!inCurrent) continue;

    // ── On-time vs late (A3) ─────────────────────────────────────────────
    // Deadline = Friday 15:00 of the reporting week (per VMRF policy and
    // the Friday-3PM reminder email). Submissions sent before or at the
    // deadline are 'onTime'; anything after is 'late'.
    var rfRaw = subD[si][subRfI];
    var deadline = (rfRaw instanceof Date) ? _weekDeadline_(rfRaw) : null;
    if (deadline && subDate2 <= deadline) {
      b2.onTime++;
    } else if (deadline) {
      b2.late++;
    } // (if deadline could not be derived, neither bucket is incremented —
      //  the row still counts toward submissions but is omitted from on-time%)

    // Resolve terminal status from the review chain. IMPORTANT: the review
    // chain is HOD → HOI → IMO, but each stage retains its own status row in
    // its own sheet. So a fully-approved submission has hoiS='Approved' AND
    // imoS='Approved'. We must therefore check IMO first (terminal), then
    // HOI, then HOD — otherwise IMO-approved submissions get mis-counted as
    // 'pending IMO'. (This was the cause of the `finalised` field always
    // reading 0 in earlier versions of this report.)
    // Map values are now objects {status, dateTime} (Tier 3 / D2). Status reads
    // tolerate the old string-shape so a deploy-skew here can't break review routing.
    var _hodRow = hodMap[sid] || null;
    var _hoiRow = hoiMap[sid] || null;
    var _imoRow = imoMap[sid] || null;
    var hodS = _hodRow ? (typeof _hodRow === 'string' ? _hodRow : (_hodRow.status||'')) : '';
    var hoiS = _hoiRow ? (typeof _hoiRow === 'string' ? _hoiRow : (_hoiRow.status||'')) : '';
    var imoS = _imoRow ? (typeof _imoRow === 'string' ? _imoRow : (_imoRow.status||'')) : '';
    var effectiveHodS = hodS || (noHodSids[sid] ? 'Approved' : '');
    if (imoS === 'Approved')                                                  b2.finalised++;       // terminal: fully approved
    else if (imoS === 'Escalated' || imoS === 'Rejected')                     b2.rejected++;        // IMO-rejected
    else if (hoiS === 'Approved')                                             b2.pendingIMO++;      // cleared HOI, awaiting IMO
    else if (hoiS === 'Rejected' || effectiveHodS === 'Rejected')             b2.rejected++;
    else if (hoiS === 'Needs Revision' || effectiveHodS === 'Needs Revision') b2.revision++;
    else if (effectiveHodS === 'Approved')                                    b2.pendingHOI++;
    else                                                                      b2.pendingHOD++;

    // ── D2: review-turnaround SLA ─────────────────────────────────────────
    // Captures the elapsed time between adjacent stages of the review chain.
    // Each stage contributes only when its review row has a valid timestamp
    // AND the prior anchor (faculty submit / prior reviewer) also has one,
    // so partial chains don't pollute the average.
    // For no-HOD faculty, HOI's anchor is the faculty submission (no HOD step).
    var hodDt = (_hodRow && typeof _hodRow === 'object') ? _hodRow.dateTime : null;
    var hoiDt = (_hoiRow && typeof _hoiRow === 'object') ? _hoiRow.dateTime : null;
    var imoDt = (_imoRow && typeof _imoRow === 'object') ? _imoRow.dateTime : null;
    var subDt = (subDate2 instanceof Date) ? subDate2 : null;
    if (subDt && hodDt) {
      var dHod = hodDt.getTime() - subDt.getTime();
      if (dHod >= 0) { b2.hodTurnaroundSumMs += dHod; b2.hodTurnaroundN++; }
    }
    if (hoiDt) {
      // For no-HOD faculty, anchor on subDt; otherwise on hodDt
      var anchorHoi = noHodSids[sid] ? subDt : hodDt;
      if (anchorHoi) {
        var dHoi = hoiDt.getTime() - anchorHoi.getTime();
        if (dHoi >= 0) { b2.hoiTurnaroundSumMs += dHoi; b2.hoiTurnaroundN++; }
      }
    }
    if (hoiDt && imoDt) {
      var dImo = imoDt.getTime() - hoiDt.getTime();
      if (dImo >= 0) { b2.imoTurnaroundSumMs += dImo; b2.imoTurnaroundN++; }
    }
    if (subDt && imoDt && imoS === 'Approved') {
      // End-to-end is only meaningful once IMO has actually finalised.
      var dEnd = imoDt.getTime() - subDt.getTime();
      if (dEnd >= 0) { b2.endToEndSumMs += dEnd; b2.endToEndN++; }
    }

    // Track latest submitted date for the bucket
    var d = subD[si][subDateI];
    if (d instanceof Date) {
      if (!b2.latestSubmittedDate || d.getTime() > b2.latestSubmittedDate.getTime()) {
        b2.latestSubmittedDate = d;
      }
    }
  }

  // ── Tier-5 / A5: HOD weekly submissions in the same buckets ─────────────
  // HODs file THEIR OWN weekly reports (HOD_Submission sheet). At dept /
  // institution / campus dimensions, those submissions belong to the same
  // bucket the HOD's department rolls up into. At dimension='faculty' this
  // is meaningless (faculty != HOD) so the pass is skipped — the response
  // fields stay null/zero for faculty rows.
  if (dimension !== 'faculty') {
    try {
      var staffSh = ss.getSheetByName(SH.STAFF);
      var hodSubSh = ss.getSheetByName(SH.HOD_SUB);
      if (staffSh && hodSubSh) {
        var staffD = staffSh.getDataRange().getValues();
        if (staffD.length > 1) {
          var sH = staffD[0];
          var sIdI   = sH.indexOf('StaffID');
          var sRoleI = sH.indexOf('Role');
          var sDeptI = sH.indexOf('Department');
          var sCampI = sH.indexOf('Campus');
          var sInstI = sH.indexOf('Institution');
          var sStatI = sH.indexOf('Status');
          // Build map: hodId(lowercased) → {Department, Campus, Institution}
          // for HODs that pass the active filter and are Active.
          var hodToBucket = {};
          for (var sii=1; sii<staffD.length; sii++) {
            var srow = staffD[sii];
            var role = String(srow[sRoleI]||'').trim().toUpperCase();
            if (role !== 'HOD') continue;
            var stat = sStatI >= 0 ? String(srow[sStatI]||'').trim() : '';
            if (stat && stat !== 'Active') continue;
            var hodFac = {
              Department:  String(srow[sDeptI]||'').trim(),
              Campus:      String(srow[sCampI]||'').trim(),
              Institution: String(srow[sInstI]||'').trim()
            };
            if (!_facMatchesFilter(hodFac)) continue;
            var hodKey = _groupKey(hodFac);
            if (!hodKey) continue;
            // Bucket exists from Pass 1 (faculty enumeration); ensure-create
            // for the rare case of an HOD whose dept has zero faculty.
            var hb = _ensure(hodKey, _groupMeta(hodFac, hodKey));
            hb.hodTotal++;
            var hodId = String(srow[sIdI]||'').trim().toLowerCase();
            if (hodId) hodToBucket[hodId] = hodKey;
          }

          // Walk HOD_Submission and increment hodSubmissions per bucket
          // (current period only — same window as faculty submissions).
          var hodSubD = hodSubSh.getDataRange().getValues();
          if (hodSubD.length > 1) {
            var hsH = hodSubD[0];
            var hsIdI   = hsH.indexOf('HOD_ID');     if (hsIdI < 0) hsIdI = 1;
            var hsDtI   = hsH.indexOf('SubmittedDateTime');
            for (var hsi=1; hsi<hodSubD.length; hsi++) {
              var hsRow = hodSubD[hsi];
              var hodId2 = String(hsRow[hsIdI]||'').trim().toLowerCase();
              if (!hodId2) continue;
              var hsBucketKey = hodToBucket[hodId2];
              if (!hsBucketKey || !buckets[hsBucketKey]) continue;
              // Period filter
              var hsDt = (hsDtI >= 0) ? hsRow[hsDtI] : null;
              if (periodStart) {
                if (!(hsDt instanceof Date) || hsDt < periodStart || hsDt > periodEnd) continue;
              } else if (!(hsDt instanceof Date)) {
                continue;
              }
              buckets[hsBucketKey].hodSubmissions++;
            }
          }
        }
      }
    } catch(_a5){ /* sheet missing or schema mismatch → skip A5 silently */ }
  }

  // ── Pass 3: Timesheet_Entries → aggregate activity categories per bucket.
  // For each filled slot, resolve its owning faculty via the submission ID,
  // then route the category count into the right bucket. Empty/unclassified
  // slots are skipped. This powers the "Ranked by Activity Category" view.
  try {
    var tsSh = ss.getSheetByName(SH.TIMESHEET);
    if (tsSh) {
      var tsD = tsSh.getDataRange().getValues();
      if (tsD.length > 1) {
        var tsH = tsD[0];
        var tsSidI = tsH.indexOf('SubmissionID');     if (tsSidI < 0) tsSidI = 0;
        var tsActI = tsH.indexOf('ActivityType');     if (tsActI < 0) tsActI = 3;
        // Cache per-submission → bucket-key so we don't re-resolve faculty
        // for every one of the ~30 timesheet rows per submission.
        var sidToBucketKey = {};
        for (var ti = 1; ti < tsD.length; ti++) {
          var tsSid = String(tsD[ti][tsSidI]||'').trim(); if (!tsSid) continue;
          var act   = String(tsD[ti][tsActI]||'').trim(); if (!act) continue;
          var bucketKey = sidToBucketKey[tsSid];
          if (bucketKey === undefined) {
            var tEmail = sidEmailMap[tsSid] || '';
            var tFac   = tEmail ? (facMap[tEmail]||{}) : {};
            var tf = {
              FacultyName: tFac['FacultyName']||'',
              Email:       tFac['Email']||tEmail,
              Department:  tFac['Department']||'',
              Campus:      tFac['Campus']||'',
              Institution: tFac['Institution']||''
            };
            bucketKey = _facMatchesFilter(tf) ? _groupKey(tf) : '';
            sidToBucketKey[tsSid] = bucketKey;  // cache even '' to skip next time
          }
          if (!bucketKey) continue;
          var tb = buckets[bucketKey];
          if (!tb) continue;  // in-scope faculty who somehow skipped pass 1
          var cat = act.split(' > ')[0] || act;
          if (!cat) continue;
          tb.categorySlots[cat] = (tb.categorySlots[cat] || 0) + 1;
        }
      }
    }
  } catch (tsErr) {
    // Non-fatal — worst case, the category view shows no data for this call.
  }

  // ── Resolve expectedWeeks for compliance-rate calculation ────────────────
  // Used as the denominator when computing complianceRate per bucket.
  // Capped at TODAY so future weeks in the period don't penalise compliance.
  var expectedWeeks = _resolveExpectedWeeks_(period, periodStart, periodEnd, earliestSubDate);

  // Finalise rows with derived metrics
  var rows = [];
  Object.keys(buckets).forEach(function(k){
    var b = buckets[k];
    var subs        = b.submissions || 0;
    var imoFinal    = b.finalised   || 0;
    var pendingIMO  = b.pendingIMO  || 0;
    var rejected    = b.rejected    || 0;
    var onTime      = b.onTime      || 0;
    var late        = b.late        || 0;
    var classified  = onTime + late;       // submissions whose deadline could be derived

    // ── New, semantically correct rates ───────────────────────────────────
    // complianceRate = % of expected weekly reports actually submitted by this
    //   bucket's faculty. Capped at 100 so ad-hoc duplicate submissions don't
    //   inflate compliance over the headline figure of 100%.
    var denom = (b.totalFaculty || 0) * expectedWeeks;
    var complianceRate = denom > 0
      ? Math.min(100, Math.round((subs / denom) * 100))
      : 0;
    // imoApprovalRate  = % of submissions that have been TERMINALLY approved by IMO
    var imoApprovalRate = subs > 0 ? Math.round((imoFinal / subs) * 100) : 0;
    // hoiApprovalRate  = % of submissions that have at least cleared the HOI
    //   gate (pending IMO + already finalised). This is the rate the legacy
    //   `approvalRate` field was *trying* to express.
    var hoiApprovalRate = subs > 0 ? Math.round(((pendingIMO + imoFinal) / subs) * 100) : 0;
    // rejectionRate    = % of submissions thrown back at any stage
    var rejectionRate = subs > 0 ? Math.round((rejected / subs) * 100) : 0;
    // onTimeRate       = % of submissions filed by Friday 15:00 of reporting week.
    //   Denominator is `classified` (= onTime + late), NOT `subs`, so submissions
    //   whose deadline couldn't be derived (missing ReportingFrom) don't drag
    //   the rate down. UI should show "—" when classified === 0.
    var onTimeRate = classified > 0 ? Math.round((onTime / classified) * 100) : null;

    // ── Trend vs equivalent previous period (A4) ──────────────────────────
    // Only meaningful when a previous-period window exists (week / month / year).
    // For 'all', prevSubmissions stays 0 and trend* are null.
    var prevSubs       = (period === 'all') ? null : (b.prevSubmissions || 0);
    var trendDelta     = (prevSubs === null) ? null : (subs - prevSubs);
    var trendPct       = null;
    var trendDirection = null;
    if (prevSubs !== null) {
      if (prevSubs > 0) {
        trendPct = Math.round(((subs - prevSubs) / prevSubs) * 100);
        if (trendPct > 0)        trendDirection = 'up';
        else if (trendPct < 0)   trendDirection = 'down';
        else                     trendDirection = 'flat';
      } else if (subs > 0) {
        // No baseline — can't compute %, but we can flag it as new activity
        trendDirection = 'new';
      } else {
        trendDirection = 'flat';
      }
    }

    // ── Tier-3 / A11: composite 0–100 score ───────────────────────────────
    // Combines compliance + on-time + IMO-approval + (1 − rejection) using
    // tunable weights. For buckets with zero submissions the per-component
    // rates would all be 0 (real numbers, not null) and rejectionInverted
    // would be 100, producing a misleading non-zero score. Short-circuit to
    // 0/{} when there's nothing meaningful to score.
    var _scoreWeights = _resolveScoreWeights_();
    var _scored;
    if (subs === 0 && b.totalFaculty === 0) {
      _scored = { score: 0, breakdown: {} };
    } else {
      _scored = _computeCompositeScore_({
        complianceRate:  complianceRate,
        onTimeRate:      onTimeRate,
        imoApprovalRate: imoApprovalRate,
        rejectionRate:   subs > 0 ? rejectionRate : null  // null when no submissions to judge
      }, _scoreWeights);
    }

    // ── Tier-3 / D2: review-turnaround SLA means ──────────────────────────
    // Convert the running ms sums to hours (rounded to 1 decimal) where a
    // sample exists; null otherwise. The UI should render `—` for null.
    var _hrs = function(sumMs, n) {
      if (!n) return null;
      return Math.round((sumMs / n) / 36e5 * 10) / 10;   // ms → hours, 1dp
    };
    var avgHodTurnaroundHrs = _hrs(b.hodTurnaroundSumMs, b.hodTurnaroundN);
    var avgHoiTurnaroundHrs = _hrs(b.hoiTurnaroundSumMs, b.hoiTurnaroundN);
    var avgImoTurnaroundHrs = _hrs(b.imoTurnaroundSumMs, b.imoTurnaroundN);
    var avgEndToEndHrs      = _hrs(b.endToEndSumMs,      b.endToEndN);

    // ── Legacy aliases (kept so existing UI code doesn't break) ───────────
    // `approvalRate` historically meant "reached IMO" — now equal to hoiApprovalRate.
    // `avgSubsPerFac` is unchanged.
    var approvalRate  = hoiApprovalRate;
    var avgSubsPerFac = b.totalFaculty > 0 ? Math.round((subs / b.totalFaculty) * 10) / 10 : 0;

    var latestSubStr = b.latestSubmittedDate
      ? Utilities.formatDate(b.latestSubmittedDate, Session.getScriptTimeZone(), 'dd MMM yyyy')
      : '';

    // Category-view derived fields
    var catSlots = b.categorySlots || {};
    var totalCatSlots = 0;
    Object.keys(catSlots).forEach(function(c){ totalCatSlots += (catSlots[c]||0); });
    var catArr = Object.keys(catSlots).map(function(c){ return { cat: c, slots: catSlots[c] }; });
    catArr.sort(function(a,b){ return b.slots - a.slots; });
    var topCat = catArr.length ? catArr[0].cat : '';

    // ── A9: category mix as percentages (Tier-5 fair-comparison metric) ───
    // Raw `categorySlots` favours large buckets (a 40-faculty teaching dept
    // always has more "Teaching" slots than a 4-faculty research dept).
    // The mix % normalises that — what FRACTION of this bucket's activity
    // is in each category? Apples-to-apples for cross-dept comparison.
    var categoryMixPct = {};
    if (totalCatSlots > 0) {
      Object.keys(catSlots).forEach(function(c){
        categoryMixPct[c] = Math.round((catSlots[c] / totalCatSlots) * 100);
      });
    }

    rows.push({
      key:            b.key,
      label:          b.label,
      subtitle:       b.subtitle,
      // (Option A) Department field on every row — used by the picker for
      // dept-grouped layout and by the comp-table header for the dept badge.
      // '' for non-faculty dimensions.
      department:     b.department || '',
      totalFaculty:   b.totalFaculty,
      submissions:    subs,
      finalised:      imoFinal,
      pendingHOD:     b.pendingHOD,
      pendingHOI:     b.pendingHOI,
      pendingIMO:     b.pendingIMO,
      rejected:       b.rejected,
      revision:       b.revision,
      // ── New fields ──
      expectedWeeks:    expectedWeeks,           // denominator transparency
      complianceRate:   complianceRate,          // % of expected reports actually filed
      imoApprovalRate:  imoApprovalRate,         // % terminally approved by IMO
      hoiApprovalRate:  hoiApprovalRate,         // % cleared past HOI
      rejectionRate:    rejectionRate,           // % rejected at any stage
      // Tier-2 additions
      onTime:           onTime,                  // raw count classified as on-time
      late:             late,                    // raw count classified as late
      onTimeRate:       onTimeRate,              // null when no submission could be classified
      prevSubmissions:  prevSubs,                // submissions in previous period (null for 'all')
      trendDelta:       trendDelta,              // raw count delta
      trendPct:         trendPct,                // % delta vs prev period (null when prev=0 or 'all')
      trendDirection:   trendDirection,          // 'up'|'down'|'flat'|'new'|null
      // Tier-3 additions (A11 composite score, D2 SLA turnaround)
      compositeScore:        _scored.score,      // 0–100 weighted score
      scoreBreakdown:        _scored.breakdown,  // each component's weighted contribution
      avgHodTurnaroundHrs:   avgHodTurnaroundHrs, // mean hrs faculty-submit → HOD-review
      avgHoiTurnaroundHrs:   avgHoiTurnaroundHrs, // mean hrs HOD-review → HOI-review (or submit→HOI for no-HOD)
      avgImoTurnaroundHrs:   avgImoTurnaroundHrs, // mean hrs HOI-review → IMO-finalisation
      avgEndToEndHrs:        avgEndToEndHrs,     // mean hrs faculty-submit → IMO-Approved
      slaCounts: {                               // sample sizes (for tooltip "n=12")
        hod: b.hodTurnaroundN, hoi: b.hoiTurnaroundN,
        imo: b.imoTurnaroundN, end: b.endToEndN
      },
      // Tier-5 / A5: HOD weekly submissions parallel to faculty submissions
      hodTotal:        (dimension === 'faculty') ? null : b.hodTotal,
      hodSubmissions:  (dimension === 'faculty') ? null : b.hodSubmissions,
      hodCompliance:   (dimension === 'faculty' || b.hodTotal === 0 || expectedWeeks === 0)
                         ? null
                         : Math.min(100, Math.round((b.hodSubmissions / (b.hodTotal * expectedWeeks)) * 100)),
      // Tier-5 / D4: year-on-year comparison (same window, 1 year earlier)
      yearAgoSubmissions: yearAgoStart ? b.yearAgoSubmissions : null,
      yoyDelta:           yearAgoStart ? (subs - b.yearAgoSubmissions) : null,
      yoyPct:             (yearAgoStart && b.yearAgoSubmissions > 0)
                            ? Math.round(((subs - b.yearAgoSubmissions) / b.yearAgoSubmissions) * 100)
                            : null,
      // ── Legacy aliases (do NOT remove without auditing the front-end) ──
      approvalRate:   approvalRate,
      avgSubsPerFac:  avgSubsPerFac,
      latestSubmitted:latestSubStr,
      categorySlots:  catSlots,
      categoryMixPct: categoryMixPct,         // Tier-5 / A9: fair cross-bucket comparison
      totalCategorySlots: totalCatSlots,
      topCategory:    topCat,
      categoryCount:  catArr.length
    });
  });

  // Sort: most submissions first, ties broken by label
  rows.sort(function(a,b){
    if (b.submissions !== a.submissions) return b.submissions - a.submissions;
    return String(a.label).localeCompare(String(b.label));
  });

  // ── Tier-5 / A8: anomaly flags ───────────────────────────────────────────
  // Surface attention-worthy outliers as a top-level `flags` array so the UI
  // can render an exception-based banner above the table. Detection thresholds
  // are configurable via Script Property ANOMALY_THRESHOLDS — see
  // _resolveAnomalyThresholds_. Each row also gets its own `alerts` field with
  // the same shape so the UI can render in-row indicators alongside the global
  // flags banner.
  var _anomalyThresholds = _resolveAnomalyThresholds_();
  var flags = [];
  rows.forEach(function(row) {
    var alerts = _detectAnomalies_(row, _anomalyThresholds);
    row.alerts = alerts;
    alerts.forEach(function(a) {
      flags.push({
        key:      row.key,
        label:    row.label,
        type:     a.type,
        severity: a.severity,
        detail:   a.detail
      });
    });
  });
  // Sort flags: high severity first, then medium, then low. Stable within group.
  var sevOrder = { high: 0, med: 1, low: 2 };
  flags.sort(function(a, b){ return sevOrder[a.severity] - sevOrder[b.severity]; });

  // Format prev-period bounds for the response — useful for UI tooltips like
  // "vs. Sep 2025 (124 submissions)".
  var _fmtD = function(d){ return d ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd MMM yyyy') : null; };

  return {
    dimension: dimension,
    period: period,
    filter: {
      // Legacy single-value fields — first selected value of each, for UI compat
      campus:      fCampus,
      institution: fInst,
      department:  fDeptArr[0] || '',
      // New array fields (A6) — full multi-select state
      campuses:    fCampusArr,
      institutions:fInstArr,
      departments: fDeptArr
    },
    expectedWeeks: expectedWeeks,        // # reports expected per faculty in the chosen period
    periodBounds: periodStart ? { from: _fmtD(periodStart), to: _fmtD(periodEnd) } : null,
    prevPeriod: prevPeriodStart ? { from: _fmtD(prevPeriodStart), to: _fmtD(prevPeriodEnd) } : null,
    yearAgoPeriod: yearAgoStart ? { from: _fmtD(yearAgoStart), to: _fmtD(yearAgoEnd) } : null,
    flags: flags,                        // Tier-5 / A8: anomaly flags
    rows: rows,
    generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm')
  };
}

// ═════════════════════════════════════════════════════════════════════════════
// COMPARISON REPORT — CSV EXPORT (A10)
// ═════════════════════════════════════════════════════════════════════════════
// Returns a CSV-encoded snapshot of getComparisonReport for download. The
// caller passes the same dimension/filter/period as the live report and gets
// back { filename, mimeType, content } where `content` is the raw CSV text.
// The frontend turns this into a Blob and triggers a download.
//
// Columns mirror the on-screen comparison table; rows are ordered the same
// way getComparisonReport orders them (most submissions first). Numeric
// fields are emitted as numbers (no quotes); strings are CSV-quoted with
// embedded quotes doubled per RFC 4180.
function exportComparisonReportCsv(dimension, filter, period) {
  var report = getComparisonReport(dimension, filter, period);

  // Helper: CSV-quote a field (RFC 4180). Strings get wrapped in quotes if
  // they contain a comma, newline, or quote; existing quotes are doubled.
  function _csvField(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'number') {
      return isFinite(v) ? String(v) : '';
    }
    var s = String(v);
    if (/[,"\r\n]/.test(s)) {
      return '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  }
  function _row(arr) { return arr.map(_csvField).join(','); }

  // Header row — matches the most useful comparison-report fields. Order is
  // chosen for readability when opened in Excel: identity → counts → rates
  // → trend → SLA → composite. Legacy aliases like avgSubsPerFac are omitted
  // to keep the export tidy; the new tier-1+2+3 fields are the canonical set.
  var header = [
    'Key', 'Label', 'Subtitle', 'Department',
    'Total Faculty', 'Submissions', 'Expected Weeks',
    'Compliance %', 'On-Time %', 'IMO Approval %', 'HOI Approval %', 'Rejection %',
    'Pending HOD', 'Pending HOI', 'Pending IMO', 'Rejected', 'Revision', 'Finalised',
    'On-Time Count', 'Late Count',
    'Prev Submissions', 'Trend Δ', 'Trend %', 'Trend Direction',
    'Avg HOD Hrs', 'Avg HOI Hrs', 'Avg IMO Hrs', 'Avg End-to-End Hrs',
    'Composite Score',
    'Latest Submitted'
  ];

  var lines = [_row(header)];
  report.rows.forEach(function(r){
    lines.push(_row([
      r.key, r.label, r.subtitle, r.department,
      r.totalFaculty, r.submissions, r.expectedWeeks,
      r.complianceRate, r.onTimeRate, r.imoApprovalRate, r.hoiApprovalRate, r.rejectionRate,
      r.pendingHOD, r.pendingHOI, r.pendingIMO, r.rejected, r.revision, r.finalised,
      r.onTime, r.late,
      r.prevSubmissions, r.trendDelta, r.trendPct, r.trendDirection,
      r.avgHodTurnaroundHrs, r.avgHoiTurnaroundHrs, r.avgImoTurnaroundHrs, r.avgEndToEndHrs,
      r.compositeScore,
      r.latestSubmitted
    ]));
  });

  // Filename: dimension + period + ISO timestamp. Friendly and unambiguous.
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmm');
  var fname = 'comparison_' + (report.dimension||'report') + '_' + (report.period||'all') + '_' + stamp + '.csv';

  return {
    ok:       true,
    filename: fname,
    mimeType: 'text/csv;charset=utf-8',
    content:  lines.join('\r\n'),    // RFC 4180 prefers CRLF
    rowCount: report.rows.length
  };
}

// ═════════════════════════════════════════════════════════════════════════════
// SCHEDULED COMPARISON DIGEST — D5
// ═════════════════════════════════════════════════════════════════════════════
// Sends an HTML email summary of the institution-level comparison report to a
// configurable list of recipients. Designed to be called from a time-driven
// Apps Script trigger (e.g. every Monday morning).
//
// Recipients: read from Script Property COMPARISON_DIGEST_RECIPIENTS — either
// a comma-separated email list, or a JSON array. If no property is set, the
// function logs and returns without sending.
//
// Period: read from Script Property COMPARISON_DIGEST_PERIOD (default 'week').
// Accepts any value getComparisonReport accepts.
//
// Content: top-10 institutions by composite score, with their compliance / on-
// time / approval / SLA / trend. Bottom-3 also flagged so the Chancellor's
// office sees both the leaders and the laggards on one page.
function sendComparisonDigest() {
  // 1. Resolve recipients
  var sp = PropertiesService.getScriptProperties();
  var recipRaw = sp.getProperty('COMPARISON_DIGEST_RECIPIENTS') || '';
  var recipients = [];
  if (recipRaw) {
    try {
      var asJson = JSON.parse(recipRaw);
      if (Array.isArray(asJson)) recipients = asJson;
    } catch(_) {
      recipients = recipRaw.split(/[,;\n]/);
    }
  }
  recipients = recipients.map(function(s){return String(s||'').trim();}).filter(Boolean);
  if (!recipients.length) {
    Logger.log('sendComparisonDigest: no recipients (set Script Property COMPARISON_DIGEST_RECIPIENTS)');
    return { ok: false, reason: 'no_recipients' };
  }

  // 2. Resolve period
  var period = sp.getProperty('COMPARISON_DIGEST_PERIOD') || 'week';

  // 3. Build report
  var report = getComparisonReport('institution', {}, period);
  var rows = (report.rows || []).slice(); // already sorted by submissions desc

  // Re-sort by composite score desc for the leaderboard view
  rows.sort(function(a, b){
    var sa = a.compositeScore || 0, sb = b.compositeScore || 0;
    if (sb !== sa) return sb - sa;
    return String(a.label).localeCompare(String(b.label));
  });

  var top    = rows.slice(0, 10);
  var bottom = rows.length > 10 ? rows.slice(-3).reverse() : [];

  // 4. Build HTML email
  var esc = function(s){
    return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
      .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
  };
  var fmtPct  = function(v){ return (v === null || v === undefined) ? '—' : (v + '%'); };
  var fmtNum  = function(v){ return (v === null || v === undefined) ? '—' : String(v); };
  var fmtHrs  = function(v){ return (v === null || v === undefined) ? '—' : (v + ' h'); };
  var fmtTrend = function(r){
    if (r.trendDirection === 'up')   return '<span style="color:#10b981">▲ ' + (r.trendPct || 0) + '%</span>';
    if (r.trendDirection === 'down') return '<span style="color:#ef4444">▼ ' + Math.abs(r.trendPct || 0) + '%</span>';
    if (r.trendDirection === 'flat') return '<span style="color:#6b7280">— flat</span>';
    if (r.trendDirection === 'new')  return '<span style="color:#3b82f6">new</span>';
    return '—';
  };
  var scoreColor = function(s){
    if (s >= 80) return '#10b981';
    if (s >= 60) return '#f59e0b';
    return '#ef4444';
  };

  function _renderRow(r, rank) {
    return '<tr style="border-bottom:1px solid #e5e7eb">' +
      '<td style="padding:8px 10px;font-weight:700;color:#6b7280">' + rank + '</td>' +
      '<td style="padding:8px 10px;font-weight:600;color:#1f2937">' + esc(r.label) + '</td>' +
      '<td style="padding:8px 10px;text-align:center;font-weight:800;color:' + scoreColor(r.compositeScore) + '">' +
        fmtNum(r.compositeScore) + '</td>' +
      '<td style="padding:8px 10px;text-align:center;color:#374151">' + fmtNum(r.submissions) + '</td>' +
      '<td style="padding:8px 10px;text-align:center;color:#374151">' + fmtPct(r.complianceRate) + '</td>' +
      '<td style="padding:8px 10px;text-align:center;color:#374151">' + fmtPct(r.onTimeRate) + '</td>' +
      '<td style="padding:8px 10px;text-align:center;color:#374151">' + fmtPct(r.imoApprovalRate) + '</td>' +
      '<td style="padding:8px 10px;text-align:center">' + fmtTrend(r) + '</td>' +
      '<td style="padding:8px 10px;text-align:center;color:#6b7280;font-size:12px">' + fmtHrs(r.avgEndToEndHrs) + '</td>' +
      '</tr>';
  }

  var topRows    = top.map(function(r,i){return _renderRow(r, i+1);}).join('');
  var bottomRows = bottom.length
    ? '<h3 style="margin:24px 0 8px 0;color:#991b1b;font-size:13px;text-transform:uppercase;letter-spacing:.05em">Lowest 3 by composite score</h3>' +
      '<table style="width:100%;border-collapse:collapse;font-family:Arial,sans-serif;font-size:13px">' +
      '<thead><tr style="background:#fef2f2"><th style="padding:8px 10px;text-align:left;color:#6b7280">#</th>' +
      '<th style="padding:8px 10px;text-align:left;color:#6b7280">Institution</th>' +
      '<th style="padding:8px 10px;text-align:center;color:#6b7280">Score</th>' +
      '<th style="padding:8px 10px;text-align:center;color:#6b7280">Subs</th>' +
      '<th style="padding:8px 10px;text-align:center;color:#6b7280">Compl%</th>' +
      '<th style="padding:8px 10px;text-align:center;color:#6b7280">OnTime%</th>' +
      '<th style="padding:8px 10px;text-align:center;color:#6b7280">IMO%</th>' +
      '<th style="padding:8px 10px;text-align:center;color:#6b7280">Trend</th>' +
      '<th style="padding:8px 10px;text-align:center;color:#6b7280">E2E</th></tr></thead>' +
      '<tbody>' + bottom.map(function(r,i){return _renderRow(r, rows.length - i);}).join('') + '</tbody></table>'
    : '';

  var pStr = report.periodBounds
    ? (esc(report.periodBounds.from) + ' – ' + esc(report.periodBounds.to))
    : esc(period);

  var html = ''+
    '<div style="font-family:Arial,sans-serif;max-width:760px;margin:0 auto;color:#1f2937">' +
    '<div style="background:linear-gradient(135deg,#1a2332,#28375a);padding:20px 24px;border-radius:8px 8px 0 0">' +
      '<h1 style="margin:0;color:#fff;font-size:18px;font-weight:700">VMRF Comparison Digest</h1>' +
      '<p style="margin:4px 0 0;color:rgba(255,255,255,.7);font-size:12px">Period: ' + pStr + ' · Generated ' + esc(report.generatedAt) + '</p>' +
    '</div>' +
    '<div style="border:1px solid #e5e7eb;border-top:none;padding:20px 24px;background:#fff;border-radius:0 0 8px 8px">' +
      '<h3 style="margin:0 0 8px 0;color:#1f2937;font-size:13px;text-transform:uppercase;letter-spacing:.05em">Top by composite score</h3>' +
      '<table style="width:100%;border-collapse:collapse;font-family:Arial,sans-serif;font-size:13px">' +
        '<thead><tr style="background:#f9fafb"><th style="padding:8px 10px;text-align:left;color:#6b7280">#</th>' +
          '<th style="padding:8px 10px;text-align:left;color:#6b7280">Institution</th>' +
          '<th style="padding:8px 10px;text-align:center;color:#6b7280">Score</th>' +
          '<th style="padding:8px 10px;text-align:center;color:#6b7280">Subs</th>' +
          '<th style="padding:8px 10px;text-align:center;color:#6b7280">Compl%</th>' +
          '<th style="padding:8px 10px;text-align:center;color:#6b7280">OnTime%</th>' +
          '<th style="padding:8px 10px;text-align:center;color:#6b7280">IMO%</th>' +
          '<th style="padding:8px 10px;text-align:center;color:#6b7280">Trend</th>' +
          '<th style="padding:8px 10px;text-align:center;color:#6b7280">E2E</th></tr></thead>' +
        '<tbody>' + topRows + '</tbody></table>' +
      bottomRows +
      '<p style="margin:20px 0 0;font-size:11px;color:#9ca3af;line-height:1.5">' +
        'Composite score weights: 0.45 compliance + 0.25 on-time + 0.20 IMO approval + 0.10 (1 − rejection). ' +
        'E2E = mean hours from faculty submission to IMO finalisation.' +
      '</p>' +
    '</div>' +
    '</div>';

  // 5. Send
  var subject = 'VMRF Comparison Digest — ' + (report.periodBounds ? report.periodBounds.from + ' to ' + report.periodBounds.to : period);
  try {
    MailApp.sendEmail({
      to: recipients.join(','),
      subject: subject,
      htmlBody: html,
      noReply: true
    });
    return { ok: true, sent: recipients.length, period: period };
  } catch (e) {
    Logger.log('sendComparisonDigest mail error: ' + e.message);
    return { ok: false, error: e.message };
  }
}

// One-time setup helper — installs a weekly trigger that calls sendComparisonDigest.
// Run from the Apps Script editor: setupComparisonDigestTrigger().
// Idempotent — removes any previous trigger pointing at sendComparisonDigest first.
function setupComparisonDigestTrigger(weekday, hour) {
  weekday = weekday || ScriptApp.WeekDay.MONDAY;
  hour    = (typeof hour === 'number') ? hour : 8;
  var existing = ScriptApp.getProjectTriggers();
  existing.forEach(function(t){
    if (t.getHandlerFunction() === 'sendComparisonDigest') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendComparisonDigest')
    .timeBased().onWeekDay(weekday).atHour(hour).create();
  return { ok: true, weekday: String(weekday), hour: hour };
}

// Take a snapshot of the just-ended Mon-Sun week's KPIs. Replaces any
// existing rows for that week (idempotent — safe to re-run). Runs the
// comparison report for campus, institution and department dimensions, and
// flattens every row into the archive sheet.
function archiveWeeklyKpis() {
  var sh = _kpiArchiveSheet_();
  var bounds = _lastCompletedWeekBounds_(new Date());
  var weekStart = bounds.start, weekEnd = bounds.end;
  var tz = Session.getScriptTimeZone();
  var weekStartISO = Utilities.formatDate(weekStart, tz, 'yyyy-MM-dd');

  // ── 1. Remove any earlier snapshot of the same week (idempotency) ──
  var data = sh.getDataRange().getValues();
  if (data.length > 1) {
    var rowsToDelete = [];
    for (var r = 1; r < data.length; r++) {
      var ws = data[r][0];
      if (ws instanceof Date) {
        var iso = Utilities.formatDate(ws, tz, 'yyyy-MM-dd');
        if (iso === weekStartISO) rowsToDelete.push(r + 1);
      }
    }
    // Delete from the bottom up so indices remain stable
    for (var i = rowsToDelete.length - 1; i >= 0; i--) {
      sh.deleteRow(rowsToDelete[i]);
    }
  }

  // ── 2. Build the snapshot via getComparisonReport with custom bounds ──
  var customFilter = {
    from: Utilities.formatDate(weekStart, tz, 'yyyy-MM-dd'),
    to:   Utilities.formatDate(weekEnd,   tz, 'yyyy-MM-dd')
  };
  var generatedAt = new Date();
  var dims = ['campus', 'institution', 'department'];
  var newRows = [];

  dims.forEach(function(dim){
    var report;
    try {
      report = getComparisonReport(dim, customFilter, 'custom');
    } catch (e) {
      Logger.log('archiveWeeklyKpis: ' + dim + ' failed — ' + e.message);
      return;
    }
    (report.rows || []).forEach(function(row){
      // Composite scope tag — useful for fast filtered reads downstream.
      // For department/institution we encode parent scope so an HoI can pull
      // just "departments in my institution" without scanning the JSON.
      var scopeParts = [];
      if (row.subtitle) {
        // subtitle is "Inst · Campus" for department, "Campus" for institution
        scopeParts = String(row.subtitle).split('·').map(function(s){ return s.trim(); });
      }
      var scopeKey;
      if (dim === 'campus')          scopeKey = 'all';
      else if (dim === 'institution') scopeKey = 'campus:' + (scopeParts[0] || '');
      else if (dim === 'department')  scopeKey = 'inst:' + (scopeParts[0] || '') + '|campus:' + (scopeParts[1] || '');
      else                            scopeKey = 'all';

      newRows.push([
        weekStart,
        weekEnd,
        dim,
        scopeKey,
        row.key || '',
        row.label || '',
        row.subtitle || '',
        row.department || '',
        JSON.stringify(row),
        generatedAt
      ]);
    });
  });

  if (newRows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, newRows.length, 10).setValues(newRows);
  }

  return {
    ok: true,
    weekStart: Utilities.formatDate(weekStart, tz, 'yyyy-MM-dd'),
    weekEnd:   Utilities.formatDate(weekEnd,   tz, 'yyyy-MM-dd'),
    rowsArchived: newRows.length,
    dimensions: dims
  };
}

// Read one week's archived KPIs, optionally scoped. Filter is the same shape
// the live comparison report accepts: { campus, institution, department }
// where department implies the department dimension, institution implies
// institution-or-department dimensions in that institution, and a bare
// campus implies institutions in that campus. Empty filter = system-wide.
function getArchivedWeeklyKpis(filter, weekStartISO) {
  filter = filter || {};
  var sh = _kpiArchiveSheet_();
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return { ok: true, rows: [], weekStart: weekStartISO || '', weekEnd: '' };

  var tz = Session.getScriptTimeZone();
  var fCampus = String(filter.campus      || '').trim();
  var fInst   = String(filter.institution || '').trim();
  var fDept   = String(filter.department  || '').trim();

  // Pick the dimension implied by the filter — most specific wins.
  var wantDim;
  if (fDept) wantDim = 'department';
  else if (fInst) wantDim = 'department'; // HoI sees departments in their inst
  else if (fCampus) wantDim = 'institution';
  else wantDim = 'institution'; // IMO default: institution dimension

  var rows = [];
  var weekStart = null, weekEnd = null;
  for (var r = 1; r < data.length; r++) {
    var ws = data[r][0];
    if (!(ws instanceof Date)) continue;
    var iso = Utilities.formatDate(ws, tz, 'yyyy-MM-dd');
    if (weekStartISO && iso !== weekStartISO) continue;

    var dim = String(data[r][2] || '');
    if (dim !== wantDim) continue;

    var rowKey   = String(data[r][4] || '');
    var subtitle = String(data[r][6] || '');
    // Filter by parent scope. subtitle holds the parent crumb chain.
    if (fCampus && subtitle.indexOf(fCampus) < 0) continue;
    if (fInst   && subtitle.indexOf(fInst)   < 0 && rowKey !== fInst) continue;
    if (fDept   && rowKey !== fDept) continue;

    var json = String(data[r][8] || '');
    var parsed = null;
    try { parsed = JSON.parse(json); } catch (e) { parsed = null; }
    if (!parsed) continue;
    rows.push(parsed);

    if (!weekStart) {
      weekStart = data[r][0];
      weekEnd   = data[r][1];
    }
  }

  return {
    ok: true,
    rows: rows,
    dimension: wantDim,
    weekStart: weekStart instanceof Date ? Utilities.formatDate(weekStart, tz, 'yyyy-MM-dd') : (weekStartISO || ''),
    weekEnd:   weekEnd   instanceof Date ? Utilities.formatDate(weekEnd,   tz, 'yyyy-MM-dd') : '',
    weekLabel: (weekStart && weekEnd) ? _kpiWeekLabel_(weekStart, weekEnd) : ''
  };
}

// Return the list of distinct archived weeks, newest first. Optionally
// limited to weeks where at least one row matches the given scope filter.
function listArchivedWeeks(filter) {
  filter = filter || {};
  var sh = _kpiArchiveSheet_();
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return { ok: true, weeks: [] };

  var tz = Session.getScriptTimeZone();
  var fCampus = String(filter.campus      || '').trim();
  var fInst   = String(filter.institution || '').trim();
  var fDept   = String(filter.department  || '').trim();
  var hasFilter = !!(fCampus || fInst || fDept);

  // Build a {iso → {weekStart, weekEnd, count}} map then sort desc.
  var map = {};
  for (var r = 1; r < data.length; r++) {
    var ws = data[r][0], we = data[r][1];
    if (!(ws instanceof Date)) continue;
    if (hasFilter) {
      var subtitle = String(data[r][6] || '');
      var rowKey   = String(data[r][4] || '');
      if (fCampus && subtitle.indexOf(fCampus) < 0) continue;
      if (fInst   && subtitle.indexOf(fInst)   < 0 && rowKey !== fInst) continue;
      if (fDept   && rowKey !== fDept) continue;
    }
    var iso = Utilities.formatDate(ws, tz, 'yyyy-MM-dd');
    if (!map[iso]) {
      map[iso] = {
        weekStart: iso,
        weekEnd:   we instanceof Date ? Utilities.formatDate(we, tz, 'yyyy-MM-dd') : '',
        label:     (ws instanceof Date && we instanceof Date) ? _kpiWeekLabel_(ws, we) : iso,
        count: 0
      };
    }
    map[iso].count++;
  }
  var weeks = Object.keys(map).map(function(k){ return map[k]; });
  weeks.sort(function(a, b){ return b.weekStart.localeCompare(a.weekStart); });
  return { ok: true, weeks: weeks };
}

// One-time setup helper — installs a Sunday 23:30 weekly trigger that calls
// archiveWeeklyKpis(). Run from the Apps Script editor:
//   setupWeeklyArchiveTrigger();
// Idempotent — removes any existing trigger pointing at archiveWeeklyKpis
// before creating the new one. Override hour by passing it as an argument.
function setupWeeklyArchiveTrigger(hour) {
  hour = (typeof hour === 'number') ? hour : 23;
  var existing = ScriptApp.getProjectTriggers();
  existing.forEach(function(t){
    if (t.getHandlerFunction() === 'archiveWeeklyKpis') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('archiveWeeklyKpis')
    .timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(hour).create();
  return { ok: true, weekday: 'SUNDAY', hour: hour };
}

// ═════════════════════════════════════════════════════════════════════════════
// REPORT COMPARISON — side-by-side content comparison of individual faculty
// weekly reports. Distinct from the metrics-focused Comparison Report: that
// page asks "who is more compliant"; this page asks "what did they actually
// do this week". Used by HoD (own department), HoI (own institution), and
// IMO (system-wide).
//
// Two endpoints:
//   getReportComparisonWeeks(role, scope)
//       → list of recent weeks (ISO Monday + label + submission count) that
//         have at least one submission in scope. Used to populate the week
//         dropdown so the user only sees weeks that actually contain data.
//   getReportComparison(role, scope, weekIso, facultyIds)
//       → parallel report objects for the chosen faculty for that week.
//         Faculty without a submission for the week get a column with
//         hasSubmission:false so the UI can render a placeholder rather
//         than silently dropping them.
// ═════════════════════════════════════════════════════════════════════════════

/**
 * Return scope-matching distinct weeks that have at least one submission.
 * Caps at the most-recent 52 weeks to avoid scanning all-time data on each
 * page open. Returned newest first.
 *
 * role:  'HOD' | 'HOI' | 'IMO'
 * scope: { department?, campus?, institution? }   (interpreted per role)
 * facultyEmail: optional — if provided, only return weeks where that
 *               specific faculty has a submission. Used by the Timeline
 *               comparison mode (Mode C) to populate a per-faculty week
 *               picker. Email is matched case-insensitively.
 */
function getReportComparisonWeeks(role, scope, facultyEmail) {
  try {
    role = String(role||'').toUpperCase();
    scope = scope || {};
    if (['HOD','HOI','IMO'].indexOf(role) < 0) {
      return { ok: false, error: 'Role must be HOD, HOI, or IMO', weeks: [] };
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var subD = ss.getSheetByName(SH.SUBMISSION).getDataRange().getValues();
    if (subD.length < 2) return { ok: true, weeks: [] };
    var facD = ss.getSheetByName(SH.FACULTY).getDataRange().getValues();
    var staffD = ss.getSheetByName(SH.STAFF).getDataRange().getValues();
    var facMap = _buildFacMap(facD, staffD);

    var subH = subD[0];
    var emI  = subH.indexOf('FacultyEmail');
    var rfI  = subH.indexOf('ReportingFrom');
    var sidI = subH.indexOf('SubmissionID'); if (sidI < 0) sidI = 0;
    if (emI < 0 || rfI < 0) return { ok: false, error: 'Submission sheet schema unexpected', weeks: [] };

    // Pre-normalise scope filters once
    var sDept = String(scope.department||'').trim().toLowerCase();
    var sInst = String(scope.institution||'').trim().toLowerCase();
    var sInstBare = sInst.replace(/\s*\([^)]*\)\s*$/, '').trim();
    var sCampus = String(scope.campus||'').trim().toLowerCase();
    var sFacEmail = String(facultyEmail||'').trim().toLowerCase();

    var tz = Session.getScriptTimeZone();
    var weekMap = {};  // iso → { iso, label, count }

    for (var i = 1; i < subD.length; i++) {
      var fEmail = String(subD[i][emI]||'').trim().toLowerCase();
      if (!fEmail) continue;
      // Skip rows that have no SubmissionID — they cannot be looked up later
      // and would produce "no submission" placeholders in the comparison view.
      var chkSid = String(subD[i][sidI]||'').trim();
      if (!chkSid) continue;
      // Timeline mode: hard filter to a single faculty's submissions only.
      if (sFacEmail && fEmail !== sFacEmail) continue;
      var fac = facMap[fEmail] || {};

      // Scope check
      if (role === 'HOD' && sDept) {
        var fDept = String(fac['Department']||'').trim().toLowerCase();
        if (fDept !== sDept) continue;
      }
      if (role === 'HOI' && sInst) {
        var fInst = String(fac['Institution']||'').trim().toLowerCase();
        var fInstBare = fInst.replace(/\s*\([^)]*\)\s*$/, '').trim();
        if (fInst !== sInst && fInstBare !== sInstBare && fInst !== sInstBare && fInstBare !== sInst) continue;
      }
      if (role === 'IMO' && sCampus) {
        var fCampus = String(fac['Campus']||'').trim().toLowerCase();
        if (fCampus !== sCampus) continue;
      }

      var rFrom = subD[i][rfI];
      var iso;
      if (rFrom instanceof Date) iso = Utilities.formatDate(rFrom, tz, 'yyyy-MM-dd');
      else iso = String(rFrom||'').slice(0,10);
      if (!iso || iso.length < 10) continue;

      if (!weekMap[iso]) {
        var d = new Date(iso + 'T00:00:00');
        var end = new Date(d); end.setDate(d.getDate() + 6);
        var sameMonth = d.getMonth() === end.getMonth();
        var label = sameMonth
          ? Utilities.formatDate(d, tz, 'd') + '–' + Utilities.formatDate(end, tz, 'd MMM yyyy')
          : Utilities.formatDate(d, tz, 'd MMM') + ' – ' + Utilities.formatDate(end, tz, 'd MMM yyyy');
        weekMap[iso] = { iso: iso, label: label, count: 0 };
      }
      weekMap[iso].count++;
    }

    var weeks = Object.keys(weekMap).map(function(k){ return weekMap[k]; });
    weeks.sort(function(a,b){ return b.iso.localeCompare(a.iso); });   // newest first
    if (weeks.length > 52) weeks = weeks.slice(0, 52);  // Extended to 52 weeks (~1 year) for past report access
    return { ok: true, weeks: weeks };
  } catch (e) {
    Logger.log('getReportComparisonWeeks ERROR: ' + (e && e.message) + '\n' + (e && e.stack || ''));
    return { ok: false, error: String((e && e.message) || e || 'Unknown server error'), weeks: [] };
  }
}

/**
 * Return parallel report data for `facultyIds` for `weekIso`.
 *
 * Each column carries: meta, activity-category breakdown (counts by
 * ActivityType across the timesheet — drives the side-by-side breakdown
 * chart), self-assessment text, full timesheet, attachments, and reviewer
 * remarks. Faculty without a submission for that week are still returned
 * with hasSubmission:false so the UI can show a placeholder.
 *
 * Scope is enforced server-side — even a HoD passing a faculty outside
 * their department gets an outOfScope column rather than the data.
 */
/**
 * Return parallel report data for `facultyIds` for `weekIso`.
 *
 * Each column carries: meta, activity-category breakdown (counts by
 * ActivityType across the timesheet — drives the side-by-side breakdown
 * chart), self-assessment text, full timesheet, attachments, and reviewer
 * remarks. Faculty without a submission for that week are still returned
 * with hasSubmission:false so the UI can show a placeholder.
 *
 * Scope is enforced server-side — even a HoD passing a faculty outside
 * their department gets an outOfScope column rather than the data.
 *
 * `selection` is an optional 4th argument that selects WHICH submission
 * to load per column. Backward-compatible: if omitted/null, each faculty's
 * latest submission is loaded (the original behaviour).
 *
 *   selection = null  |  undefined  |  { mode: 'latest' }
 *       → Latest submission per faculty (default, no params needed)
 *
 *   selection = { mode: 'specificWeek', weekIso: 'YYYY-MM-DD' }
 *       → All faculty load their submission for that exact week. Faculty
 *         without a submission for that week return hasSubmission:false.
 *         `facultyIds` carries the faculty to compare.
 *
 *   selection = { mode: 'perFaculty', weeks: { 'email@x': 'YYYY-MM-DD', ... } }
 *       → Each faculty loads their submission for the per-email week.
 *         If a faculty is missing from `weeks` or its iso doesn't match
 *         any of their submissions, their column comes back as
 *         hasSubmission:false. `facultyIds` carries the faculty to compare.
 *
 *   selection = { mode: 'timeline', facultyEmail: 'email@x', weeks: ['iso1','iso2',...] }
 *       → Single faculty, multiple weeks. Each entry in `weeks` becomes
 *         one column. `facultyIds` is IGNORED in this mode (the faculty
 *         is taken from selection.facultyEmail).
 */
function getReportComparison(role, scope, facultyIds, selection) {
  try {
    role = String(role||'').toUpperCase();
    scope = scope || {};
    if (['HOD','HOI','IMO'].indexOf(role) < 0) {
      return { ok: false, error: 'Role must be HOD, HOI, or IMO' };
    }
    // Normalise the selection parameter. Treat null/missing as "latest" so
    // every existing client call site continues to work unchanged.
    var sel = selection || { mode: 'latest' };
    var mode = String(sel.mode || 'latest');
    if (['latest','specificWeek','perFaculty','timeline'].indexOf(mode) < 0) {
      return { ok: false, error: 'Unknown comparison mode: ' + mode };
    }
    // Validation depends on mode. Timeline drives columns from selection.weeks;
    // the other three drive columns from facultyIds.
    if (mode === 'timeline') {
      if (!sel.facultyEmail) {
        return { ok: false, error: 'Timeline mode requires selection.facultyEmail' };
      }
      if (!Array.isArray(sel.weeks) || sel.weeks.length < 2) {
        return { ok: false, error: 'Timeline mode requires at least 2 weeks' };
      }
      if (sel.weeks.length > 4) {
        return { ok: false, error: 'Up to 4 weeks can be compared side-by-side' };
      }
    } else {
      if (!Array.isArray(facultyIds) || !facultyIds.length) {
        return { ok: false, error: 'Pick at least one faculty to compare' };
      }
      if (facultyIds.length > 4) {
        return { ok: false, error: 'Up to 4 faculty can be compared side-by-side' };
      }
      if (mode === 'specificWeek' && !sel.weekIso) {
        return { ok: false, error: 'Specific-week mode requires selection.weekIso' };
      }
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var subD = ss.getSheetByName(SH.SUBMISSION).getDataRange().getValues();
    var facD = ss.getSheetByName(SH.FACULTY).getDataRange().getValues();
    var staffD = ss.getSheetByName(SH.STAFF).getDataRange().getValues();
    var hodD = ss.getSheetByName(SH.HOD).getDataRange().getValues();
    var hoiD = ss.getSheetByName(SH.HOI).getDataRange().getValues();
    var imoD = ss.getSheetByName(SH.IMO).getDataRange().getValues();
    var saD  = ss.getSheetByName(SH.SELF_ASSESS).getDataRange().getValues();
    var tsD  = ss.getSheetByName(SH.TIMESHEET).getDataRange().getValues();

    var subMap = _bm(subD, subD[0]);
    var facMap = _buildFacMap(facD, staffD);
    var hodMap = _bm(hodD, hodD[0]);
    var hoiMap = _bm(hoiD, hoiD[0]);
    var imoMap = _bm(imoD, imoD[0]);
    var saMap  = _bm(saD, saD[0]);
    var tsMap  = _bmMulti(tsD, tsD[0]);

    var tz = Session.getScriptTimeZone();
    var subH = subD[0];
    var emI  = subH.indexOf('FacultyEmail');  if (emI  < 0) emI  = 1;
    var rfI  = subH.indexOf('ReportingFrom'); if (rfI  < 0) rfI  = 3;
    var rtI  = subH.indexOf('ReportingTo');   if (rtI  < 0) rtI  = 4;
    var sidI = subH.indexOf('SubmissionID');  if (sidI < 0) sidI = 0;

    // byFacAll: email → [{ sid, fromIso, toIso }, ...] sorted desc by fromIso.
    // Replaces the old "byFacLatest" single-record-per-faculty index — we now
    // need to resolve specific weeks too, not just the most recent one.
    var byFacAll = {};
    for (var i = 1; i < subD.length; i++) {
      var rEmail = String(subD[i][emI]||'').trim().toLowerCase();
      if (!rEmail) continue;
      var rawSid = String(subD[i][sidI]||'').trim();
      if (!rawSid) continue;  // Skip rows without a SubmissionID
      var rFrom = subD[i][rfI];
      var rFromIso = rFrom instanceof Date
        ? Utilities.formatDate(rFrom, tz, 'yyyy-MM-dd')
        : String(rFrom||'').slice(0,10);
      if (!rFromIso || rFromIso.length < 10) continue;
      var rTo = subD[i][rtI];
      var rToIso = rTo instanceof Date
        ? Utilities.formatDate(rTo, tz, 'yyyy-MM-dd')
        : String(rTo||'').slice(0,10);
      if (!byFacAll[rEmail]) byFacAll[rEmail] = [];
      byFacAll[rEmail].push({ sid: rawSid, fromIso: rFromIso, toIso: rToIso });
    }
    Object.keys(byFacAll).forEach(function(k){
      byFacAll[k].sort(function(a,b){ return b.fromIso.localeCompare(a.fromIso); });
    });

    // Pre-normalise scope filters
    var sDept = String(scope.department||'').trim().toLowerCase();
    var sInst = String(scope.institution||'').trim().toLowerCase();
    var sInstBare = sInst.replace(/\s*\([^)]*\)\s*$/, '').trim();

    // Helper: resolve which submission record to use for (email, mode).
    // Returns the record { sid, fromIso, toIso } or null if no match.
    function _pickSubmission(emailLc, weekIsoOrNull) {
      var list = byFacAll[emailLc] || [];
      if (!list.length) return null;
      if (!weekIsoOrNull) return list[0];   // latest
      for (var k = 0; k < list.length; k++) {
        if (list[k].fromIso === weekIsoOrNull) return list[k];
      }
      return null;
    }

    // Helper: apply scope rules. Returns a flag (true=allowed, false=outOfScope).
    function _inScope(fac) {
      if (role === 'HOD' && sDept) {
        var fDept = String(fac['Department']||'').trim().toLowerCase();
        if (fDept !== sDept) return false;
      }
      if (role === 'HOI' && sInst) {
        var fInst = String(fac['Institution']||'').trim().toLowerCase();
        var fInstBare = fInst.replace(/\s*\([^)]*\)\s*$/, '').trim();
        if (fInst !== sInst && fInstBare !== sInstBare && fInst !== sInstBare && fInstBare !== sInst) return false;
      }
      return true;
    }

    // Helper: hydrate a column object from a faculty record + resolved submission.
    function _buildColumn(rawFid, fac, record, weekLabelHint) {
      var col = {
        facultyId:   String(rawFid),
        facultyName: String(fac['FacultyName']||rawFid||'Unknown'),
        department:  String(fac['Department']||''),
        institution: String(fac['Institution']||''),
        campus:      String(fac['Campus']||''),
        designation: String(fac['Designation']||''),
        hasSubmission: false
      };
      if (!record) {
        // weekLabelHint lets the timeline / specific-week placeholders show
        // which week the submission was being looked up for. The UI uses
        // this to render "No submission for <week>" rather than a generic
        // empty card.
        if (weekLabelHint) col.requestedWeek = weekLabelHint;
        return col;
      }
      var sid = record.sid;
      var sub = subMap[sid]||{};
      var sa  = saMap[sid]||{};
      var ts  = tsMap[sid]||[];
      var hod = hodMap[sid]||{};
      var hoi = hoiMap[sid]||{};
      var imo = imoMap[sid]||{};

      // Activity-category breakdown — count timesheet slots per ActivityType.
      var breakdown = {};
      ts.forEach(function(t){
        var typ = String(t['ActivityType']||'').trim() || 'Other';
        breakdown[typ] = (breakdown[typ]||0) + 1;
      });
      var breakdownArr = Object.keys(breakdown).map(function(k){
        return { type: k, count: breakdown[k] };
      }).sort(function(a,b){ return b.count - a.count; });

      // Attachments — pulled out of timesheet entries that have a URL
      var atts = [];
      ts.forEach(function(t){
        var u = String(t['AttachmentURL']||'').trim();
        if (u) atts.push({ name: String(t['AttachmentName']||'Attachment'), url: u });
      });

      col.hasSubmission = true;
      col.submissionId  = sid;
      col.semester      = String(sub['AcademicYearSemester']||'');
      col.fromDate      = _fmt(sub['ReportingFrom']);
      col.toDate        = _fmt(sub['ReportingTo']);
      col.fromIso       = record.fromIso;       // raw iso for client sorting / labels
      col.toIso         = record.toIso;
      col.submitted     = _fmtDT(sub['SubmittedDateTime']);
      col.outcome       = String(sa['OutcomeOfWeek']||'');
      col.target        = String(sa['TargetPlanNextWeek']||'');
      col.breakdown     = breakdownArr;
      col.timesheet     = ts.map(function(t){
        return {
          day:     String(t['Day']||''),
          slot:    String(t['TimeSlot']||''),
          type:    String(t['ActivityType']||''),
          details: String(t['ActivityDetails']||''),
          attUrl:  String(t['AttachmentURL']||''),
          attName: String(t['AttachmentName']||'')
        };
      });
      col.attachments   = atts;
      col.hodStatus     = String(hod['HOD_Status']||'');
      col.hodRemark     = String(hod['HOD_Remark']||'');
      col.hodDateTime   = _fmtDT(hod['HOD_DateTime']);
      col.hoiStatus     = String(hoi['HOI_Status']||'');
      col.hoiRemark     = String(hoi['HOI_Remark']||'');
      col.hoiDateTime   = _fmtDT(hoi['HOI_DateTime']);
      col.imoStatus     = String(imo['IMO_Status']||'');
      return col;
    }

    var columns;

    if (mode === 'timeline') {
      // Single faculty, multiple weeks. Columns are derived from sel.weeks.
      var tEmailLc = String(sel.facultyEmail||'').trim().toLowerCase();
      var tFac = facMap[tEmailLc] || {};
      var tInScope = _inScope(tFac);
      // We use the email itself as the facultyId echo on each column so the
      // client can still group/identify by faculty.
      columns = sel.weeks.map(function(w){
        var weekIso = String(w||'').slice(0,10);
        if (!tInScope) {
          var blocked = _buildColumn(sel.facultyEmail, tFac, null, weekIso);
          blocked.outOfScope = true;
          blocked.requestedWeek = weekIso;
          return blocked;
        }
        var rec = _pickSubmission(tEmailLc, weekIso);
        var col = _buildColumn(sel.facultyEmail, tFac, rec, weekIso);
        if (!rec) col.requestedWeek = weekIso;
        return col;
      });
    } else {
      // Latest / Specific-week / Per-faculty all walk facultyIds.
      var pfWeeks = (mode === 'perFaculty' && sel.weeks && typeof sel.weeks === 'object')
        ? sel.weeks
        : null;
      var sharedWeekIso = (mode === 'specificWeek') ? String(sel.weekIso||'').slice(0,10) : null;

      columns = facultyIds.map(function(rawFid){
        var fidLc = String(rawFid||'').trim().toLowerCase();
        var fac = facMap[fidLc] || {};

        // Scope enforcement — never let a HoD see another department's data
        // even if the client somehow passes a faculty outside their scope.
        if (!_inScope(fac)) {
          var blocked = _buildColumn(rawFid, fac, null, null);
          blocked.outOfScope = true;
          return blocked;
        }

        var lookupWeek;
        if (mode === 'latest') {
          lookupWeek = null;
        } else if (mode === 'specificWeek') {
          lookupWeek = sharedWeekIso;
        } else { // perFaculty
          // pfWeeks keys can be the raw email or normalised lowercase — accept either.
          var raw = pfWeeks ? (pfWeeks[rawFid] || pfWeeks[fidLc] || '') : '';
          lookupWeek = String(raw||'').slice(0,10) || null;
        }

        var rec = _pickSubmission(fidLc, lookupWeek);
        var col = _buildColumn(rawFid, fac, rec, lookupWeek);
        if (!rec && lookupWeek) col.requestedWeek = lookupWeek;
        return col;
      });
    }

    return {
      ok: true,
      mode: mode,
      // For specific-week / timeline modes we echo back the shared week so
      // the client can label headers without recomputing from columns.
      sharedWeekIso: (mode === 'specificWeek') ? String(sel.weekIso||'').slice(0,10) : null,
      timelineFacultyEmail: (mode === 'timeline') ? String(sel.facultyEmail||'') : null,
      // Each column carries its own fromDate/toDate (and now fromIso/toIso)
      // so the client can render per-column week labels even when the
      // columns are pinned to different weeks.
      columns: columns
    };
  } catch (e) {
    Logger.log('getReportComparison ERROR: ' + (e && e.message) + '\n' + (e && e.stack || ''));
    return { ok: false, error: String((e && e.message) || e || 'Unknown server error') };
  }
}

function listSavedComparisons(ownerEmail) {
  try {
    var owner = String(ownerEmail||'').trim().toLowerCase();
    if (!owner) return { ok: false, error: 'ownerEmail required (user not identified)', views: [] };
    var sh = _savedComparisonsSheet_();
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return { ok: true, views: [] };
    var h = data[0];
    var views = [];
    for (var i=1; i<data.length; i++) {
      var r = data[i];
      var rowOwner = String(r[h.indexOf('OwnerEmail')]||'').trim().toLowerCase();
      var shared   = String(r[h.indexOf('Shared')]||'').toLowerCase() === 'true';
      if (rowOwner !== owner && !shared) continue;
      var fj = String(r[h.indexOf('FilterJSON')]||'{}');
      var filterParsed = {};
      try { filterParsed = JSON.parse(fj); } catch(_) {}
      views.push({
        viewId:    String(r[h.indexOf('ViewID')]||''),
        name:      String(r[h.indexOf('Name')]||''),
        dimension: String(r[h.indexOf('Dimension')]||''),
        filter:    filterParsed,
        period:    String(r[h.indexOf('Period')]||''),
        createdAt: r[h.indexOf('CreatedAt')] instanceof Date
                     ? Utilities.formatDate(r[h.indexOf('CreatedAt')], Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm') : '',
        updatedAt: r[h.indexOf('UpdatedAt')] instanceof Date
                     ? Utilities.formatDate(r[h.indexOf('UpdatedAt')], Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm') : '',
        shared:    shared,
        isOwner:   rowOwner === owner
      });
    }
    // Sort: own views first, then shared. Within each group, most-recent-updated first.
    views.sort(function(a,b){
      if (a.isOwner !== b.isOwner) return a.isOwner ? -1 : 1;
      return String(b.updatedAt).localeCompare(String(a.updatedAt));
    });
    return { ok: true, views: views };
  } catch (e) {
    Logger.log('listSavedComparisons ERROR: ' + (e && e.message) + '\n' + (e && e.stack || ''));
    return { ok: false, error: String((e && e.message) || e || 'Unknown server error'), views: [] };
  }
}

function saveComparisonView(ownerEmail, view) {
  // Wrap the whole body in try-catch so the client gets a useful error message
  // instead of an opaque "unknown error". Previously any throw here went to
  // withFailureHandler — fine for transport-level issues — but if the response
  // failed to serialize (e.g. unexpected object in view.filter) the client saw
  // a null response and the success handler showed "unknown error" with no
  // diagnostic. Returning {ok:false, error} always gives the user real info.
  try {
    var owner = String(ownerEmail||'').trim().toLowerCase();
    if (!owner) return { ok: false, error: 'ownerEmail required (user not identified)' };
    if (!view) return { ok: false, error: 'No view payload supplied' };
    if (!view.name)      return { ok: false, error: 'View name is required' };
    if (!view.dimension) return { ok: false, error: 'Comparison dimension is required — load a comparison first, then save' };

    var sh = _savedComparisonsSheet_();
    var data = sh.getDataRange().getValues();
    var h = data[0];
    var idCol = h.indexOf('ViewID');
    var ownerCol = h.indexOf('OwnerEmail');
    var now = new Date();

    // Defensive JSON.stringify — if a non-serializable value ended up in
    // view.filter (e.g. a Date, a function reference, or circular ref) JSON
    // would throw. Catch it cleanly so the user sees the field at fault.
    var filterJson;
    try {
      filterJson = JSON.stringify(view.filter || {});
    } catch (jerr) {
      Logger.log('saveComparisonView: filter stringify failed — ' + jerr.message);
      return { ok: false, error: 'Filter contains non-serializable values: ' + jerr.message };
    }

    // Update existing if viewId provided AND owned by caller
    if (view.viewId) {
      for (var i=1; i<data.length; i++) {
        if (String(data[i][idCol]) !== view.viewId) continue;
        if (String(data[i][ownerCol]||'').trim().toLowerCase() !== owner) {
          return { ok: false, error: 'Cannot modify a view owned by another user' };
        }
        var rng = sh.getRange(i+1, 1, 1, h.length);
        var newRow = data[i].slice();
        newRow[h.indexOf('Name')]       = view.name;
        newRow[h.indexOf('Dimension')]  = view.dimension;
        newRow[h.indexOf('FilterJSON')] = filterJson;
        newRow[h.indexOf('Period')]     = view.period || 'all';
        newRow[h.indexOf('UpdatedAt')]  = now;
        newRow[h.indexOf('Shared')]     = !!view.shared;
        rng.setValues([newRow]);
        // Return ISO strings, not Date objects — Apps Script Date serialization
        // across google.script.run has been a source of "null response" issues
        // (the client receives undefined when an unexpected Date format slips
        // through). Strings are unambiguously transportable.
        var isoNow = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm');
        return { ok: true, view: { viewId: view.viewId, name: view.name, dimension: view.dimension,
                                    filter: view.filter || {}, period: view.period || 'all',
                                    shared: !!view.shared, updatedAt: isoNow } };
      }
      // viewId given but not found → fall through to create-with-that-id
    }

    // Create
    var newId = view.viewId || ('SC_' + now.getTime().toString(36).toUpperCase() + '_' + Math.floor(Math.random()*1000));
    // Build the row by index — Array.fill is V8-only, replaced with a loop so
    // the function works under either Apps Script runtime.
    var row = [];
    for (var c = 0; c < h.length; c++) row.push('');
    row[h.indexOf('ViewID')]     = newId;
    row[h.indexOf('OwnerEmail')] = owner;
    row[h.indexOf('Name')]       = view.name;
    row[h.indexOf('Dimension')]  = view.dimension;
    row[h.indexOf('FilterJSON')] = filterJson;
    row[h.indexOf('Period')]     = view.period || 'all';
    row[h.indexOf('CreatedAt')]  = now;
    row[h.indexOf('UpdatedAt')]  = now;
    row[h.indexOf('Shared')]     = !!view.shared;
    sh.appendRow(row);
    var isoNow2 = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm');
    return { ok: true, view: { viewId: newId, name: view.name, dimension: view.dimension,
                                filter: view.filter || {}, period: view.period || 'all',
                                shared: !!view.shared, createdAt: isoNow2, updatedAt: isoNow2 } };
  } catch (e) {
    Logger.log('saveComparisonView ERROR: ' + (e && e.message) + '\n' + (e && e.stack || ''));
    return { ok: false, error: String((e && e.message) || e || 'Unknown server error') };
  }
}

function deleteComparisonView(ownerEmail, viewId) {
  try {
    var owner = String(ownerEmail||'').trim().toLowerCase();
    if (!owner)  return { ok: false, error: 'ownerEmail required (user not identified)' };
    if (!viewId) return { ok: false, error: 'viewId required' };
    var sh = _savedComparisonsSheet_();
    var data = sh.getDataRange().getValues();
    var h = data[0];
    for (var i=1; i<data.length; i++) {
      if (String(data[i][h.indexOf('ViewID')]) !== String(viewId)) continue;
      if (String(data[i][h.indexOf('OwnerEmail')]||'').trim().toLowerCase() !== owner) {
        return { ok: false, error: 'Cannot delete a view owned by another user' };
      }
      sh.deleteRow(i+1);
      return { ok: true, deleted: viewId };
    }
    return { ok: false, error: 'View not found' };
  } catch (e) {
    Logger.log('deleteComparisonView ERROR: ' + (e && e.message) + '\n' + (e && e.stack || ''));
    return { ok: false, error: String((e && e.message) || e || 'Unknown server error') };
  }
}

// ═════════════════════════════════════════════════════════════════════════════
// INSTITUTION PULSE — D1 (cross-portal unified view)
// ═════════════════════════════════════════════════════════════════════════════
// Combines the faculty-weekly side (compliance, on-time, approval, SLA, composite
// score) with the IPM institutional-reporting side (submission count, item
// completeness, status mix) into one row per institution. Intended for IMO and
// the Chancellor's office — the single page they actually want on a Monday.
//
// Authorization: this function is role-agnostic at the function level (matches
// the pattern of getComparisonReport). The frontend dispatcher must restrict
// callers to IMO.
//
// Args:
//   filter: { campus?, institution? }   (department is ignored at this level)
//   period: 'week'|'month'|'year'|'all' (default 'all')
//
// Returns:
//   {
//     period, generatedAt,
//     rows: [
//       {
//         institutionName, institutionCode, campus,
//         faculty: { totalFaculty, submissions, complianceRate, onTimeRate,
//                    imoApprovalRate, hoiApprovalRate, rejectionRate,
//                    pendingHOD, pendingHOI, pendingIMO, rejected, revision,
//                    finalised, expectedWeeks, latestSubmitted },
//         ipm:     { submissions, itemsFilled, expectedItems, completenessPct,
//                    byStatus: { 'Submitted':N, 'Reviewed':N, ... },
//                    latestSubmitted },
//         sla:     { avgHodHrs, avgHoiHrs, avgImoHrs, avgEndToEndHrs, samples },
//         composite: { score, breakdown }
//       }
//     ]
//   }
function getInstitutionPulse(filter, period) {
  filter = filter || {};
  period = String(period || 'all').toLowerCase();

  // 1. Pull faculty-side institution rows by reusing getComparisonReport.
  //    This already gives us compliance, on-time, approval, trend, composite,
  //    SLA — everything we need for the faculty.* sub-object.
  var facReport = getComparisonReport('institution', filter, period);

  // 2. Resolve period bounds locally — same algorithm as getComparisonReport
  //    so IPM filtering matches the faculty filter. Kept inline to avoid a
  //    risky refactor of getComparisonReport's period block.
  var now = new Date();
  var pStart = null, pEnd = null;
  if (period === 'week') {
    var dow = now.getDay(); var daysToMon = (dow === 0) ? -6 : (1 - dow);
    pStart = new Date(now); pStart.setDate(now.getDate() + daysToMon); pStart.setHours(0,0,0,0);
    pEnd = new Date(pStart); pEnd.setDate(pStart.getDate() + 6); pEnd.setHours(23,59,59,999);
  } else if (period === 'month') {
    pStart = new Date(now.getFullYear(), now.getMonth(), 1, 0,0,0,0);
    pEnd   = new Date(now.getFullYear(), now.getMonth()+1, 0, 23,59,59,999);
  } else if (period === 'year') {
    pStart = new Date(now.getFullYear(), 0, 1, 0,0,0,0);
    pEnd   = new Date(now.getFullYear(), 11, 31, 23,59,59,999);
  }

  // 3. Build IPM lookups — institution-name → IPM (code, campus). Tolerate
  //    a missing IPM database (e.g. ipmSetup() never ran) by returning an
  //    empty IPM section rather than throwing.
  var nameToCode = {}, nameToIpmCampus = {};
  var ipmSubsByCode = {};      // code → { count, itemsFilled, expectedItems, byStatus, latestDate }
  var ipmAvailable = false;
  try {
    var instSh = ipmSheet_('Institutions');
    if (instSh) {
      ipmAvailable = true;
      ipmRowsAsObjects_(instSh).forEach(function(i){
        var n = String(i.name||'').trim(); if (!n) return;
        nameToCode[n] = String(i.code||'').trim();
        nameToIpmCampus[n] = String(i.campus||'').trim();
      });

      // Pre-compute itemsBySid the same way ipmGetComparisonReport does.
      var itemsBySid = {};
      var itemSh = ipmSheet_('SubmissionItems');
      if (itemSh) {
        ipmRowsAsObjects_(itemSh).forEach(function(it){
          var sid = String(it.submissionId||'').trim(); if (!sid) return;
          var has = (String(it.minutes||'').trim() !== '') ||
                    (String(it.actionItems||'').trim() !== '') ||
                    (String(it.responsibility||'').trim() !== '');
          if (has) itemsBySid[sid] = (itemsBySid[sid]||0) + 1;
        });
      }

      // expectedItems per campus = number of distinct itemTitles in Activities.
      var expectedByCampus = {}, seenAct = {};
      var actSh = ipmSheet_('Activities');
      if (actSh) {
        ipmRowsAsObjects_(actSh).forEach(function(r){
          if (!r.itemTitle) return;
          var k = String(r.campus||'') + '||' + String(r.itemTitle).trim();
          if (seenAct[k]) return;
          seenAct[k] = true;
          expectedByCampus[r.campus] = (expectedByCampus[r.campus]||0) + 1;
        });
      }

      // Walk submissions, period-bound, accumulating into ipmSubsByCode.
      var subSh = ipmSheet_('Submissions');
      if (subSh) {
        ipmRowsAsObjects_(subSh).forEach(function(r){
          var d = new Date(r.timestamp);
          if (isNaN(d.getTime())) return;
          if (pStart && (d < pStart || d > pEnd)) return;          // period filter
          var code = String(r.institution||'').trim();
          if (!code) return;
          if (!ipmSubsByCode[code]) {
            ipmSubsByCode[code] = { count:0, itemsFilled:0, expectedItems:0, byStatus:{}, latestDate:null };
          }
          var c = ipmSubsByCode[code];
          c.count++;
          c.itemsFilled    += (itemsBySid[String(r.id||'')] || 0);
          c.expectedItems  += (expectedByCampus[r.campus] || 0);
          var st = String(r.status||'').trim() || 'Submitted';
          c.byStatus[st] = (c.byStatus[st]||0) + 1;
          if (!c.latestDate || d.getTime() > c.latestDate.getTime()) c.latestDate = d;
        });
      }
    }
  } catch (ipmErr) {
    // Non-fatal — pulse rows just won't have an `ipm` section populated.
    ipmAvailable = false;
  }

  // 4. Stitch together one row per faculty-side institution. If IPM has
  //    institutions the faculty side doesn't (IPM-only entities), they
  //    appear at the bottom with an empty faculty section.
  var seenInstNames = {};
  var pulseRows = facReport.rows.map(function(r){
    var instName = String(r.label||'').trim();
    seenInstNames[instName] = true;
    var ipmCode = nameToCode[instName] || '';
    var ipmCampus = nameToIpmCampus[instName] || '';
    var ipmStat = ipmCode ? (ipmSubsByCode[ipmCode] || null) : null;
    var ipmCompPct = (ipmStat && ipmStat.expectedItems > 0)
      ? Math.round((ipmStat.itemsFilled / ipmStat.expectedItems) * 100)
      : null;

    return {
      institutionName: instName,
      institutionCode: ipmCode,
      campus:          (r.subtitle && !ipmCampus) ? String(r.subtitle).trim() : ipmCampus,
      faculty: {
        totalFaculty:    r.totalFaculty,
        submissions:     r.submissions,
        finalised:       r.finalised,
        pendingHOD:      r.pendingHOD,
        pendingHOI:      r.pendingHOI,
        pendingIMO:      r.pendingIMO,
        rejected:        r.rejected,
        revision:        r.revision,
        expectedWeeks:   r.expectedWeeks,
        complianceRate:  r.complianceRate,
        onTimeRate:      r.onTimeRate,
        imoApprovalRate: r.imoApprovalRate,
        hoiApprovalRate: r.hoiApprovalRate,
        rejectionRate:   r.rejectionRate,
        latestSubmitted: r.latestSubmitted,
        trendPct:        r.trendPct,
        trendDirection:  r.trendDirection
      },
      ipm: ipmStat ? {
        submissions:     ipmStat.count,
        itemsFilled:     ipmStat.itemsFilled,
        expectedItems:   ipmStat.expectedItems,
        completenessPct: ipmCompPct,
        byStatus:        ipmStat.byStatus,
        latestSubmitted: ipmStat.latestDate
          ? Utilities.formatDate(ipmStat.latestDate, Session.getScriptTimeZone(), 'dd MMM yyyy')
          : ''
      } : null,
      sla: {
        avgHodHrs:       r.avgHodTurnaroundHrs,
        avgHoiHrs:       r.avgHoiTurnaroundHrs,
        avgImoHrs:       r.avgImoTurnaroundHrs,
        avgEndToEndHrs:  r.avgEndToEndHrs,
        samples:         r.slaCounts
      },
      composite: {
        score:     r.compositeScore,
        breakdown: r.scoreBreakdown
      }
    };
  });

  // Sort: highest composite score first; ties broken by institution name.
  pulseRows.sort(function(a, b){
    var sa = (a.composite && a.composite.score) || 0;
    var sb = (b.composite && b.composite.score) || 0;
    if (sb !== sa) return sb - sa;
    return String(a.institutionName).localeCompare(String(b.institutionName));
  });

  return {
    period: period,
    rows: pulseRows,
    ipmAvailable: ipmAvailable,
    generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm')
  };
}

// ═════════════════════════════════════════════════════════════════════════════
// NAAC COMPARISON REPORT — D3 (NAAC criterion alignment)
// ═════════════════════════════════════════════════════════════════════════════
// Aggregates institution-level activity across the seven NAAC criteria. The
// mapping from activity category → NAAC criterion is configurable via a
// `NAAC_Mapping` sheet:
//
//   Sheet: NAAC_Mapping
//   Columns: Category | Criterion | Description
//
// Where Category matches the top-level slot category found in Timesheet_Entries
// (e.g. "Teaching", "Research", "Administrative", ...) and Criterion is the
// number 1–7. If the sheet is absent or empty, a built-in best-effort default
// mapping is used (see _resolveNaacMapping_). Any category without a mapping
// falls into the special "Unmapped" bucket so nothing is silently dropped.
//
// NAAC criteria reference:
//   1. Curricular Aspects
//   2. Teaching-Learning and Evaluation
//   3. Research, Innovations and Extension
//   4. Infrastructure and Learning Resources
//   5. Student Support and Progression
//   6. Governance, Leadership and Management
//   7. Institutional Values and Best Practices
//
// Returns: { criteria: [...], rows: [{key,label,subtitle,naac:{1:N,2:N,...,7:N,Unmapped:N}, total:N}], generatedAt }
function _resolveNaacMapping_() {
  // 1. Built-in defaults — VMRF can override any mapping via the NAAC_Mapping
  //    sheet. These are chosen to be sensible starting points based on common
  //    Indian academic-monitoring activity-category names.
  var defaults = {
    'Teaching':                      2,    // Teaching-Learning
    'Curriculum':                    1,    // Curricular Aspects
    'Curriculum Development':        1,
    'Course Development':            1,
    'Examination':                   2,
    'Evaluation':                    2,
    'Research':                      3,    // Research/Innovation
    'Publication':                   3,
    'Innovation':                    3,
    'Patent':                        3,
    'Consultancy':                   3,
    'Extension':                     3,
    'Lab':                           4,    // Infrastructure
    'Laboratory':                    4,
    'Library':                       4,
    'Infrastructure':                4,
    'Student Support':               5,
    'Student Mentoring':             5,
    'Mentoring':                     5,
    'Student Activity':              5,
    'Placement':                     5,
    'Counselling':                   5,
    'Administrative':                6,    // Governance
    'Administration':                6,
    'Committee':                     6,
    'Meeting':                       6,
    'Governance':                    6,
    'Best Practices':                7,    // Institutional Values
    'Institutional Values':          7,
    'Outreach':                      7,
    'Community Service':             7
  };
  // 2. Override with NAAC_Mapping sheet (if present + valid).
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('NAAC_Mapping');
    if (sh) {
      var data = sh.getDataRange().getValues();
      if (data.length > 1) {
        var hdr = data[0].map(function(v){return String(v).trim();});
        var catI = hdr.indexOf('Category');
        var critI = hdr.indexOf('Criterion');
        if (catI >= 0 && critI >= 0) {
          for (var i=1; i<data.length; i++) {
            var cat = String(data[i][catI]||'').trim();
            var crit = parseInt(data[i][critI], 10);
            if (cat && crit >= 1 && crit <= 7) defaults[cat] = crit;
          }
        }
      }
    }
  } catch(_) { /* sheet missing or malformed → keep defaults */ }
  return defaults;
}

function getNaacComparisonReport(filter, period) {
  filter = filter || {};
  period = String(period || 'all').toLowerCase();

  // 1. Reuse getComparisonReport at institution level — gives us the buckets
  //    with categorySlots already aggregated per institution.
  var report = getComparisonReport('institution', filter, period);

  // 2. Build category → criterion mapping.
  var mapping = _resolveNaacMapping_();
  var lcMapping = {}; // case-insensitive lookup
  Object.keys(mapping).forEach(function(k){ lcMapping[String(k).toLowerCase()] = mapping[k]; });

  // NAAC criterion canonical labels for the response
  var criteriaLabels = {
    1: '1. Curricular Aspects',
    2: '2. Teaching-Learning & Evaluation',
    3: '3. Research, Innovations & Extension',
    4: '4. Infrastructure & Learning Resources',
    5: '5. Student Support & Progression',
    6: '6. Governance, Leadership & Management',
    7: '7. Institutional Values & Best Practices'
  };

  // 3. Re-bucket each row's categorySlots into NAAC criteria.
  var unmappedCategories = {};   // surfaces in response so VMRF can populate the sheet
  var pulseRows = report.rows.map(function(r){
    var naac = {1:0, 2:0, 3:0, 4:0, 5:0, 6:0, 7:0, Unmapped:0};
    var slots = r.categorySlots || {};
    Object.keys(slots).forEach(function(cat){
      var crit = lcMapping[String(cat).toLowerCase()];
      if (crit >= 1 && crit <= 7) {
        naac[crit] += slots[cat];
      } else {
        naac.Unmapped += slots[cat];
        unmappedCategories[cat] = (unmappedCategories[cat] || 0) + slots[cat];
      }
    });
    var total = naac[1]+naac[2]+naac[3]+naac[4]+naac[5]+naac[6]+naac[7]+naac.Unmapped;
    return {
      key:      r.key,
      label:    r.label,
      subtitle: r.subtitle,
      naac:     naac,
      total:    total
    };
  });

  return {
    period:               period,
    criteriaLabels:       criteriaLabels,
    rows:                 pulseRows,
    unmappedCategories:   unmappedCategories,   // tells admin which cats to add to NAAC_Mapping
    mappingSource:        Object.keys(unmappedCategories).length > 0 ? 'partial' : 'complete',
    generatedAt:          Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm')
  };
}

// ─── CATEGORY BREAKDOWN FOR DASHBOARDS ───────────────────────────────────────
// filter keys supported: { facultyEmail, department, campus, institution, weeks[] }
//   weeks[] — array of YYYY-MM-DD strings matching ReportingFrom on the
//             submission. If present, only timesheet rows whose parent
//             submission has ReportingFrom in this list are counted.
// Response now includes `availableWeeks: string[]` — all distinct
// ReportingFrom dates (faculty + HOD) passing the non-week filters, so the
// frontend can populate a week picker.
function getCategoryBreakdown(filter) {
  filter = filter || {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Faculty categories ──
  var tsD  = ss.getSheetByName(SH.TIMESHEET).getDataRange().getValues(), tsH = tsD[0];
  var subD = ss.getSheetByName(SH.SUBMISSION).getDataRange().getValues(), subH = subD[0];
  var facD = ss.getSheetByName(SH.FACULTY).getDataRange().getValues();
  var facMap = _buildFacMap(facD, ss.getSheetByName(SH.STAFF).getDataRange().getValues());
  var sidEmailMap = _buildSidEmailMap(subD);

  // Build a SubmissionID → ReportingFrom map (YYYY-MM-DD string) for week lookup
  var sidWeekMap = _buildSidWeekMap(subD);

  var tsActI = tsH.indexOf('ActivityType'); if(tsActI<0) tsActI=3;
  var tsSidI = tsH.indexOf('SubmissionID'); if(tsSidI<0) tsSidI=0;
  var tsDayI = tsH.indexOf('Day');          if(tsDayI<0) tsDayI=1;
  var filterDept = (filter.department||'').trim().toLowerCase();
  var filterFac  = (filter.facultyEmail||'').trim().toLowerCase();
  var filterCampus = (filter.campus||'').trim();
  var filterInst   = (filter.institution||'').trim();
  var weekSet = null;
  if (Array.isArray(filter.weeks) && filter.weeks.length) {
    weekSet = {};
    filter.weeks.forEach(function(w){ weekSet[String(w).trim().slice(0,10)] = true; });
  }

  var facCats = {};
  var facTotalSlots = 0, facFilledSlots = 0;
  var deptFaculty = {}; // email → name (for faculty dropdown)
  var facWeeksSet = {};
  for (var i = 1; i < tsD.length; i++) {
    var act = String(tsD[i][tsActI]||'').trim();
    var sid = String(tsD[i][tsSidI]||'').trim();
    var email = sidEmailMap[sid]||'';
    var fac = email ? (facMap[email]||{}) : {};
    var dept = String(fac['Department']||'').trim().toLowerCase();
    var camp = String(fac['Campus']||'').trim();
    var inst = String(fac['Institution']||'').trim();
    var weekKey = sidWeekMap[sid] || '';

    // Non-week filters
    if (filterDept) {
      if (dept !== filterDept && dept.replace(/\s*\(pg\)\s*$/,'') !== filterDept.replace(/\s*\(pg\)\s*$/,'')) continue;
    }
    if (filterCampus && camp !== filterCampus) continue;
    if (filterInst   && inst !== filterInst)   continue;
    // Track faculty in this department for dropdown
    if (email && fac['FacultyName'] && !deptFaculty[email]) {
      deptFaculty[email] = String(fac['FacultyName']);
    }
    // Faculty filter
    if (filterFac && email !== filterFac) continue;

    // Record week as "available" BEFORE the week filter is applied so the UI
    // knows the full set of pickable weeks even when a subset is selected.
    if (weekKey) facWeeksSet[weekKey] = true;
    // Week filter
    if (weekSet && !weekSet[weekKey]) continue;

    // Skip slots on non-working days (public holidays on any weekday, even Saturdays)
    var _fDayStr = String(tsD[i][tsDayI]||'').trim();
    var _fDayM = _fDayStr.match(/\d+/); var _fDayIdx = _fDayM ? parseInt(_fDayM[0]) - 1 : -1;
    if (_fDayIdx >= 0 && weekKey && _isNonWorkingDay_(weekKey, _fDayIdx)) continue;

    facTotalSlots++;
    if (!act) continue;
    facFilledSlots++;
    var cat = act.split(' > ')[0];
    if (cat) facCats[cat] = (facCats[cat]||0) + 1;
  }

  // Build faculty list for dropdown
  var facList = [];
  for (var fe in deptFaculty) {
    facList.push({email: fe, name: deptFaculty[fe]});
  }
  facList.sort(function(a,b){ return a.name.localeCompare(b.name); });

  // ── HOD categories ──
  var hodCats = {};
  var hodTotalSlots = 0, hodFilledSlots = 0;
  var hodWeeksSet = {};
  var hodTsSh = ss.getSheetByName(SH.HOD_TS);
  var hodSubSh = ss.getSheetByName(SH.HOD);
  if (hodTsSh) {
    var hodTsD = hodTsSh.getDataRange().getValues(), hodTsH = hodTsD[0];
    var hodTsActI = hodTsH.indexOf('ActivityType'); if(hodTsActI<0) hodTsActI=3;
    var hodTsSidI = hodTsH.indexOf('SubmissionID'); if(hodTsSidI<0) hodTsSidI=0;
    var hodTsDayI = hodTsH.indexOf('Day');          if(hodTsDayI<0) hodTsDayI=1;
    // Build HOD SubmissionID → ReportingFrom map
    var hodSidWeekMap = {};
    if (hodSubSh) {
      var hodSubD = hodSubSh.getDataRange().getValues();
      hodSidWeekMap = _buildSidWeekMap(hodSubD);
    }
    for (var j = 1; j < hodTsD.length; j++) {
      var hSid = String(hodTsD[j][hodTsSidI]||'').trim();
      var hWeekKey = hodSidWeekMap[hSid] || '';
      if (hWeekKey) hodWeeksSet[hWeekKey] = true;
      if (weekSet && !weekSet[hWeekKey]) continue;

      // Skip slots on non-working days (public holidays on any weekday, even Saturdays)
      var _hDayStr = String(hodTsD[j][hodTsDayI]||'').trim();
      var _hDayM = _hDayStr.match(/\d+/); var _hDayIdx = _hDayM ? parseInt(_hDayM[0]) - 1 : -1;
      if (_hDayIdx >= 0 && hWeekKey && _isNonWorkingDay_(hWeekKey, _hDayIdx)) continue;

      hodTotalSlots++;
      var hAct = String(hodTsD[j][hodTsActI]||'').trim();
      if (!hAct) continue;
      hodFilledSlots++;
      var hCat = hAct.split(' > ')[0];
      if (hCat) hodCats[hCat] = (hodCats[hCat]||0) + 1;
    }
  }

  // Merge faculty + HOD week sets, sort descending (newest week first)
  var allWeeks = {};
  for (var fw in facWeeksSet) allWeeks[fw] = true;
  for (var hw in hodWeeksSet) allWeeks[hw] = true;
  var availableWeeks = Object.keys(allWeeks).sort().reverse();

  return {
    facCats: facCats, hodCats: hodCats, facList: facList,
    facTotalSlots: facTotalSlots, facFilledSlots: facFilledSlots,
    hodTotalSlots: hodTotalSlots, hodFilledSlots: hodFilledSlots,
    availableWeeks: availableWeeks
  };
}

// ─── CAMPUS / INSTITUTION / DEPARTMENT HIERARCHY KPIs ───
function getCampusHierarchyStats(filter) {
  filter = filter || {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var facD = ss.getSheetByName(SH.FACULTY).getDataRange().getValues(), facH = facD[0];
  var subD = ss.getSheetByName(SH.SUBMISSION).getDataRange().getValues(), subH = subD[0];
  var facMap = _buildFacMap(facD, ss.getSheetByName(SH.STAFF).getDataRange().getValues());
  var sidEmailMap = _buildSidEmailMap(subD);

  var hodD = ss.getSheetByName(SH.HOD).getDataRange().getValues(), hodH = hodD[0];
  var hoiD = ss.getSheetByName(SH.HOI).getDataRange().getValues(), hoiH = hoiD[0];
  var imoD = ss.getSheetByName(SH.IMO).getDataRange().getValues(), imoH = imoD[0];
  var hodMap = {}, hoiMap = {}, imoMap = {};
  for (var h2=1; h2<hodD.length; h2++) hodMap[String(hodD[h2][0]||'')] = String(hodD[h2][hodH.indexOf('HOD_Status')]||'');
  for (var h3=1; h3<hoiD.length; h3++) hoiMap[String(hoiD[h3][0]||'')] = String(hoiD[h3][hoiH.indexOf('HOI_Status')]||'');
  for (var h4=1; h4<imoD.length; h4++) imoMap[String(imoD[h4][0]||'')] = String(imoD[h4][imoH.indexOf('IMO_Status')]||'');

  // node: {key, faculty:0, submissions:0, finalised:0, pending:0}
  function newNode(key) { return {key:key, faculty:0, submissions:0, finalised:0, pending:0}; }
  var campusMap = {}, instMap = {}, deptMap = {};

  // Count active faculty
  var facStI = facH.indexOf('Status');
  var facCmI = facH.indexOf('Campus');
  var facInI = facH.indexOf('Institution');
  var facDpI = facH.indexOf('Department');
  for (var fi=1; fi<facD.length; fi++) {
    var st = String(facD[fi][facStI]||'').trim();
    if (st && st !== 'Active') continue;
    var camp = String(facD[fi][facCmI]||'').trim();
    var inst = String(facD[fi][facInI]||'').trim();
    var dept = String(facD[fi][facDpI]||'').trim();
    if (camp) {
      if (!campusMap[camp]) campusMap[camp] = newNode(camp);
      campusMap[camp].faculty++;
    }
    if (camp && inst) {
      var ik = camp+'|'+inst;
      if (!instMap[ik]) { instMap[ik] = newNode(inst); instMap[ik].campus = camp; }
      instMap[ik].faculty++;
    }
    if (camp && inst && dept) {
      var dk = camp+'|'+inst+'|'+dept;
      if (!deptMap[dk]) { deptMap[dk] = newNode(dept); deptMap[dk].campus = camp; deptMap[dk].institution = inst; }
      deptMap[dk].faculty++;
    }
  }

  // Count submissions and statuses
  var sidI = subH.indexOf('SubmissionID'); if(sidI<0) sidI=0;
  for (var i=1; i<subD.length; i++) {
    var sid = String(subD[i][sidI]||'').trim();
    if (!sid) continue;
    var email = sidEmailMap[sid] || '';
    var fac = email ? (facMap[email]||{}) : {};
    var camp2 = String(fac['Campus']||'').trim();
    var inst2 = String(fac['Institution']||'').trim();
    var dept2 = String(fac['Department']||'').trim();

    var ms = imoMap[sid]||'', hs = hodMap[sid]||'', os = hoiMap[sid]||'';
    var isFinalised = ms === 'Finalised';
    var isPending = !isFinalised && ms !== 'Escalated';

    if (camp2 && campusMap[camp2]) {
      campusMap[camp2].submissions++;
      if (isFinalised) campusMap[camp2].finalised++;
      else if (isPending) campusMap[camp2].pending++;
    }
    if (camp2 && inst2) {
      var ik2 = camp2+'|'+inst2;
      if (instMap[ik2]) {
        instMap[ik2].submissions++;
        if (isFinalised) instMap[ik2].finalised++;
        else if (isPending) instMap[ik2].pending++;
      }
    }
    if (camp2 && inst2 && dept2) {
      var dk2 = camp2+'|'+inst2+'|'+dept2;
      if (deptMap[dk2]) {
        deptMap[dk2].submissions++;
        if (isFinalised) deptMap[dk2].finalised++;
        else if (isPending) deptMap[dk2].pending++;
      }
    }
  }

  // Filter by drill level
  var campuses = [], institutions = [], departments = [];
  for (var ck in campusMap) campuses.push(campusMap[ck]);
  for (var ik3 in instMap) {
    if (filter.campus && instMap[ik3].campus !== filter.campus) continue;
    institutions.push(instMap[ik3]);
  }
  for (var dk3 in deptMap) {
    if (filter.campus && deptMap[dk3].campus !== filter.campus) continue;
    if (filter.institution && deptMap[dk3].institution !== filter.institution) continue;
    departments.push(deptMap[dk3]);
  }
  campuses.sort(function(a,b){return a.key.localeCompare(b.key);});
  institutions.sort(function(a,b){return a.key.localeCompare(b.key);});
  departments.sort(function(a,b){return a.key.localeCompare(b.key);});

  return { campuses: campuses, institutions: institutions, departments: departments };
}