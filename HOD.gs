// =============================================================================
// VMRF-DU Institutional Monitoring System — HOD Reviews & Remark Sheets
// =============================================================================

// ─── HOD QUEUE ────────────────────────────────────────────────────────────────
function getHODQueue(hodDept) {
  // hodDept: department string passed from frontend after HOD login
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var subD=ss.getSheetByName(SH.SUBMISSION).getDataRange().getValues();
  var facD=ss.getSheetByName(SH.FACULTY).getDataRange().getValues();
  var hodD=ss.getSheetByName(SH.HOD).getDataRange().getValues();
  var saD =ss.getSheetByName(SH.SELF_ASSESS).getDataRange().getValues();
  var tsD =ss.getSheetByName(SH.TIMESHEET).getDataRange().getValues();
  var subMap=_bm(subD,subD[0]);
  var sidEmailMap=_buildSidEmailMap(subD);
  var facMap=_buildFacMap(facD, ss.getSheetByName(SH.STAFF).getDataRange().getValues());
  var saMap=_bm(saD,saD[0]);
  var tsMap=_bmMulti(tsD,tsD[0]);
  var hodH=hodD[0];
  var stI=hodH.indexOf('HOD_Status'), sbI=hodH.indexOf('SubmissionID');
  var filterDept = (hodDept||'').trim().toLowerCase();
  var out=[];
  for(var i=1;i<hodD.length;i++){
    var sid=String(hodD[i][sbI]||'').trim();
    var st =String(hodD[i][stI]||'').trim();
    if(!sid) continue;
    if(st!=='') continue;
    var sub=subMap[sid]||{};
    var fid=sidEmailMap[sid]||_getFidFromSub(sub);
    var fac=fid?(facMap[fid]||{}):{};
    var sa=saMap[sid]||{};
    // Skip submissions from no-HOD institutions
    var facInstCode2 = String(fac['InstCode']||fac['institutionCode']||'').trim();
    var facInstName2 = String(fac['Institution']||'').trim();
    if (_isNoHodInstitution_(facInstCode2 || facInstName2)) continue;
    // Filter by department if HOD has a department assigned
    if(filterDept){
      var facDept=String(fac['Department']||'').trim().toLowerCase();
      // Exact match first
      if(facDept!==filterDept){
        // Also try base name match (strip PG suffix) for cross-programme matching
        var baseFacDept=facDept.replace(/\s*\(pg\)\s*$/,'').trim();
        var baseHodDept=filterDept.replace(/\s*\(pg\)\s*$/,'').trim();
        if(baseFacDept!==baseHodDept) continue;
      }
    }
    out.push(_buildItem(sid,sub,fac,sa,tsMap[sid]||[],{hodStatus:st},null,null));
  }
  return out;
}

function submitHODReview(sid, remark, status) {
  if(!sid)    throw new Error('Submission ID missing.');
  if(!status) throw new Error('Please select a status.');
  if(status!=='Approved'&&!remark) throw new Error('Please enter your remarks.');
  _writeReview(SH.HOD,sid,remark||'',status);
  var ss2=SpreadsheetApp.getActiveSpreadsheet();
  var sub2=_rowByKey(SH.SUBMISSION,sid)||{};
  var _fid2=_getFidFromSub(sub2);
  var fac2=_fid2?(_rowByKey(SH.FACULTY,_fid2,'Email')||{}):{}; 
  var fn2=String(fac2['FacultyName']||'Faculty');
  var fid2=_getFidFromSub(sub2);
  // Resolve the faculty's institution code so the HOI notification only
  // reaches that institution's HOI, not every HOI in the system.
  var fac2Inst = String(fac2['Institution']||'').trim();
  var fac2InstCode = '';
  try { fac2InstCode = String(_resolveInstCode('', fac2Inst) || '').toUpperCase(); } catch(e) {}
  var hoi2Key = fac2InstCode ? 'HOI:'+fac2InstCode : 'HOI';
  if(status==='Approved'){
    _pushNotif(hoi2Key,'hod_approved','✅ HOD Approved — Review Required',
      'Submission by '+fn2+' has been approved by HOD and is awaiting your review.',sid,fn2);
    _pushNotif('FACULTY:'+fid2,'status_update','Your submission was reviewed by HOD',
      'HOD has approved your submission. It has been forwarded to the Head of Institution for review.',sid,fn2);
    try{_notifyHOI(ss2,sid);}catch(e){Logger.log(e.message);}
  } else {
    _pushNotif('FACULTY:'+fid2,'rejected','❌ Submission Rejected by HOD',
      'Your submission was rejected by HOD'+(remark?' with the remark: '+remark:'')+'. Please resubmit after making the necessary changes.',sid,fn2);
    try{_notifyRevision(ss2,sid,'Head of Department',remark);}catch(e){Logger.log(e.message);}
  }
  return { ok:true };
}

// ─── HOD SUBMISSION HISTORY (for drill-down view) ────────────────────────────
function getHODSubmissionHistory(hodEmail) {
  if (!hodEmail) return [];
  hodEmail = String(hodEmail).trim().toLowerCase();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var subSh = ss.getSheetByName(SH.HOD_SUB); if (!subSh) return [];
  var subD = subSh.getDataRange().getValues(), subH = subD[0];
  var revSh = ss.getSheetByName(SH.HOD_REVIEW);
  var revD = revSh ? revSh.getDataRange().getValues() : [['SubmissionID']];
  var imoSh = ss.getSheetByName(SH.HOD_IMO);
  var imoD = imoSh ? imoSh.getDataRange().getValues() : [['SubmissionID']];
  var saSh  = ss.getSheetByName(SH.HOD_SA);
  var saD   = saSh ? saSh.getDataRange().getValues() : [['SubmissionID']];
  var tsSh  = ss.getSheetByName(SH.HOD_TS);
  var tsD   = tsSh ? tsSh.getDataRange().getValues() : [['SubmissionID']];

  var revMap = _bm(revD, revD[0]);
  var imoMap = _bm(imoD, imoD[0]);
  var saMap  = _bm(saD, saD[0]);
  var tsMap  = _bmMulti(tsD, tsD[0]);

  var hodIdI = subH.indexOf('HOD_ID'); if(hodIdI<0) hodIdI=1;
  var sidI   = subH.indexOf('SubmissionID'); if(sidI<0) sidI=0;
  var out = [];
  for (var i = 1; i < subD.length; i++) {
    var hid = String(subD[i][hodIdI]||'').trim().toLowerCase();
    if (hid !== hodEmail) continue;
    var sid = String(subD[i][sidI]||'').trim();
    if (!sid) continue;
    var rev = revMap[sid]||{};
    var imo = imoMap[sid]||{};
    var sa  = saMap[sid]||{};
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
    var hoiSt = String(rev['HOI_Status']||'');
    var imoSt = String(imo['IMO_Status']||'');
    out.push({
      submissionID: sid,
      semester:     String(subD[i][subH.indexOf('AcademicYearSemester')>=0?subH.indexOf('AcademicYearSemester'):2]||''),
      from:         _fmt(subD[i][subH.indexOf('ReportingFrom')>=0?subH.indexOf('ReportingFrom'):3]),
      to:           _fmt(subD[i][subH.indexOf('ReportingTo')>=0?subH.indexOf('ReportingTo'):4]),
      submitted:    _fmtDT(subD[i][subH.indexOf('SubmittedDateTime')>=0?subH.indexOf('SubmittedDateTime'):6]),
      hodStatus:    'N/A',
      hoiStatus:    hoiSt || 'Pending',
      imoStatus:    imoSt ? imoSt : (hoiSt==='Approved'?'Pending':'—'),
      hoiRemark:    String(rev['HOI_Remark']||''),
      imoRemark:    String(imo['IMO_Remark']||''),
      outcome:      String(sa['Tasks']||sa['OutcomeOfWeek']||''),
      target:       String(sa['TargetPlanNextWeek']||''),
      timesheet:    tsEntries
    });
  }
  return out.reverse();
}

// ─── HOD SUBMISSION KPI DRILLDOWN ────────────────────────────────────────────
function getHODStatsDrilldown(filter) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var subSh = ss.getSheetByName(SH.HOD_SUB); if (!subSh) return [];
  var subD = subSh.getDataRange().getValues(), subH = subD[0];
  var revSh = ss.getSheetByName(SH.HOD_REVIEW);
  var revD = revSh ? revSh.getDataRange().getValues() : [['SubmissionID']];
  var imoSh = ss.getSheetByName(SH.HOD_IMO);
  var imoD = imoSh ? imoSh.getDataRange().getValues() : [['SubmissionID']];
  var staffSh = ss.getSheetByName(SH.STAFF);
  var staffMap = {};
  if (staffSh) {
    var staffD2 = staffSh.getDataRange().getValues(), sh2 = staffD2[0];
    var _sE = sh2.indexOf('Email'), _sN = sh2.indexOf('StaffName'), _sD = sh2.indexOf('Department');
    var _sCm = sh2.indexOf('Campus'), _sIn = sh2.indexOf('Institution');
    for (var sf = 1; sf < staffD2.length; sf++) {
      var se = String(staffD2[sf][_sE]||'').trim().toLowerCase();
      if (se) staffMap[se] = { name: String(staffD2[sf][_sN]||''), dept: String(staffD2[sf][_sD]||''), campus: String(staffD2[sf][_sCm>=0?_sCm:3]||''), institution: String(staffD2[sf][_sIn>=0?_sIn:4]||'') };
    }
  }
  var revMap = {}, imoMap2 = {};
  var _rvStI = revD[0].indexOf ? revD[0].indexOf('HOI_Status') : -1; if(_rvStI<0) _rvStI=2;
  for (var r2 = 1; r2 < revD.length; r2++) revMap[String(revD[r2][0]||'')] = String(revD[r2][_rvStI]||'');
  var _imStI = imoD[0].indexOf ? imoD[0].indexOf('IMO_Status') : -1; if(_imStI<0) _imStI=2;
  for (var im = 1; im < imoD.length; im++) imoMap2[String(imoD[im][0]||'')] = String(imoD[im][_imStI]||'');

  var hodIdI = subH.indexOf('HOD_ID'); if(hodIdI<0) hodIdI=1;
  var frI = subH.indexOf('ReportingFrom'); if(frI<0) frI=3;
  var toI = subH.indexOf('ReportingTo'); if(toI<0) toI=4;

  var status = String(filter.status||'');
  var _fDept = String(filter.department||'').trim().toLowerCase();
  var _fInst = String(filter.institution||'').trim();
  var _fCamp = String(filter.campus||'').trim();
  var results = [];
  for (var i = 1; i < subD.length; i++) {
    var sid = String(subD[i][0]||'').trim(); if (!sid) continue;
    var hodID = String(subD[i][hodIdI]||'').trim().toLowerCase();
    var staff = staffMap[hodID] || {};
    // Apply department / institution / campus scope if provided
    if (_fDept && String(staff.dept||'').trim().toLowerCase() !== _fDept) continue;
    if (_fInst && String(staff.institution||'').trim() !== _fInst) continue;
    if (_fCamp && String(staff.campus||'').trim() !== _fCamp) continue;
    var hoiSt = revMap[sid] || '';
    var imoSt = imoMap2[sid] || '';
    var match = false;
    if (status === 'Pending HOI' && hoiSt === '') match = true;
    else if (status === 'Approved HOI' && hoiSt === 'Approved' && imoSt !== 'Finalised') match = true;
    else if (status === 'Pending IMO' && hoiSt === 'Approved' && imoSt === '') match = true;
    else if (status === 'Finalised' && (imoSt === 'Finalised' || hoiSt === 'Approved')) match = true;
    else if (status === 'Rejected' && (hoiSt === 'Rejected')) match = true;
    else if (status === 'All') match = true;
    if (!match) continue;
    results.push({
      sid: sid, name: staff.name || 'HOD', dept: staff.dept || '', desig: 'HOD',
      institution: staff.institution || '', campus: staff.campus || '',
      from: _fmt(subD[i][frI]), to: _fmt(subD[i][toI]),
      status: imoSt || hoiSt || 'Pending'
    });
  }
  return results;
}

// =============================================================================
// ─── HOD WEEKLY TIMESHEET MODULE ─────────────────────────────────────────────
// =============================================================================

function submitHODWeeklyReport(data) {
  if (!data.hodID)            throw new Error('HOD ID is required.');
  if (!data.academicYearSem)  throw new Error('Please select the Academic Year / Semester.');
  if (!data.reportingFrom)    throw new Error('Please set the Reporting From date.');
  if (!data.reportingTo)      throw new Error('Please set the Reporting To date.');
  if (data.reportingFrom > data.reportingTo) throw new Error('Reporting From cannot be after Reporting To.');
  if (!data.tasks)            throw new Error('Please fill in the Tasks field.');
  if (data.declaration !== 'YES') throw new Error('Declaration must be YES to submit.');

  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var sid = 'HSUB-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmm') + '-' + Math.random().toString(36).slice(2,6).toUpperCase();
  var now = new Date();

  // Ensure sheets exist
  ['HOD_Submission','HOD_Timesheet','HOD_SelfAssess','HOD_Review'].forEach(function(name) {
    if (!ss.getSheetByName(name)) {
      var sh = ss.insertSheet(name);
      var hdrs = SCHEMA[name];
      if (hdrs) {
        var rng = sh.getRange(1,1,1,hdrs.length);
        rng.setValues([hdrs]).setFontWeight('bold').setBackground('#1a3c5e').setFontColor('#ffffff');
        sh.setFrozenRows(1);
      }
    }
  });

  // Main submission row
  ss.getSheetByName(SH.HOD_SUB).appendRow([
    sid, String(data.hodID||'').trim().toLowerCase(), data.academicYearSem,
    data.reportingFrom ? new Date(data.reportingFrom) : '', data.reportingTo ? new Date(data.reportingTo) : '',
    data.declaration, now
  ]);

  // Timesheet rows
  if (data.timesheet && data.timesheet.length) {
    var tsRows = data.timesheet.map(function(e) {
      return [sid, e.day, e.slot, e.activity, e.details || '', e.attachmentURL||'', e.attachmentName||''];
    });
    var tsSheet = ss.getSheetByName(SH.HOD_TS);
    _ensureSheetColumns(tsSheet, SCHEMA.HOD_Timesheet); // ensure AttachmentURL/Name cols exist
    var tsStart = tsSheet.getLastRow() + 1;
    tsSheet.getRange(tsStart, 1, tsRows.length, 7).setValues(tsRows);
  }

  // Self-assessment
  ss.getSheetByName(SH.HOD_SA).appendRow([sid, data.tasks, data.targetPlanNextWeek]);

  // Blank HOI review row
  ss.getSheetByName(SH.HOD_REVIEW).appendRow([sid, '', '', '']);

  // Resolve the HOD's institution so the HOI notification is scoped to that
  // institution only. HODs live in Staff_Master, not Faculty_Master.
  var hodFacRow = data.hodID ? (_rowByKey(SH.STAFF, data.hodID, 'Email') || {}) : {};
  var hodInst   = String(hodFacRow['Institution']||'').trim();
  var hodInstCode = '';
  try { hodInstCode = String(_resolveInstCode('', hodInst) || '').toUpperCase(); } catch(e) {}
  var hoiKeyForHOD = hodInstCode ? 'HOI:'+hodInstCode : 'HOI';

  // Notify the HOI of this institution only
  _pushNotif(hoiKeyForHOD, 'hod_submission',
    '📋 HOD Weekly Report Submitted',
    'HOD has submitted their weekly timesheet for ' + data.reportingFrom + ' to ' + data.reportingTo + '. Awaiting your review.',
    sid, 'HOD'
  );

  return { ok: true, sid: sid };
}

function getHODSubmissions(hodID) {
  if (!hodID) throw new Error('HOD ID required.');
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var subSh = ss.getSheetByName(SH.HOD_SUB); if (!subSh) return [];
  var subD = subSh.getDataRange().getValues(), subH = subD[0];
  var saSh  = ss.getSheetByName(SH.HOD_SA);
  var saD   = saSh ? saSh.getDataRange().getValues() : [['SubmissionID']]; var saH = saD[0];
  var revSh = ss.getSheetByName(SH.HOD_REVIEW);
  var revD  = revSh ? revSh.getDataRange().getValues() : [['SubmissionID']]; var revH = revD[0];
  var imoSh = ss.getSheetByName(SH.HOD_IMO);
  var imoD  = imoSh ? imoSh.getDataRange().getValues() : [['SubmissionID']]; var imoH = imoD[0];
  var tsSh  = ss.getSheetByName(SH.HOD_TS);
  var tsD   = tsSh ? tsSh.getDataRange().getValues() : [['SubmissionID']]; var tsH = tsD[0];

  function _bm2(data, h) {
    var map = {};
    for (var i = 1; i < data.length; i++) { var k = String(data[i][0]||''); map[k] = {}; h.forEach(function(col,j){ map[k][col]=data[i][j]; }); }
    return map;
  }
  var saMap  = _bm2(saD, saH);
  var revMap = _bm2(revD, revH);
  var imoMap = _bm2(imoD, imoH);

  // Build timesheet map sid -> [{day,slot,activity,details,attachmentURL,attachmentName}]
  var tsMap = {};
  for (var t = 1; t < tsD.length; t++) {
    var sid = String(tsD[t][0]||'');
    if (!tsMap[sid]) tsMap[sid] = [];
    tsMap[sid].push({
      Day:            String(tsD[t][1]||''),
      TimeSlot:       String(tsD[t][2]||''),
      ActivityType:   String(tsD[t][3]||''),
      Details:        String(tsD[t][4]||''),
      AttachmentURL:  String(tsD[t][5]||''),
      AttachmentName: String(tsD[t][6]||'')
    });
  }

  var hodIDNorm = String(hodID||'').trim().toLowerCase();
  var hodIdIdx = subH.indexOf('HOD_ID');
  if(hodIdIdx < 0) return []; // sheet not yet initialised
  var out = [];
  for (var i = 1; i < subD.length; i++) {
    var r   = subD[i];
    var hid = String(r[hodIdIdx]||'').trim().toLowerCase();
    if (hid !== hodIDNorm) continue;
    var sid2 = String(r[0]||'').trim();
    var sa   = saMap[sid2] || {};
    var rev  = revMap[sid2] || {};
    var imo  = imoMap[sid2] || {};
    var hoiSt = String(rev['HOI_Status']||'');
    var imoSt = String(imo['IMO_Status']||'');
    out.push({
      sid:          sid2,
      semester:     String(r[subH.indexOf('AcademicYearSemester')]||''),
      from:         _fmt(r[subH.indexOf('ReportingFrom')]),
      to:           _fmt(r[subH.indexOf('ReportingTo')]),
      submitted:    _fmtDT(r[subH.indexOf('SubmittedDateTime')]),
      tasks:        String(sa['Tasks']||''),
      target:       String(sa['TargetPlanNextWeek']||''),
      hoiStatus:    hoiSt || 'Pending HOI',
      hoiRemark:    String(rev['HOI_Remark']||''),
      imoStatus:    imoSt ? imoSt : (hoiSt === 'Approved' ? 'Pending' : '—'),
      imoRemark:    String(imo['IMO_Remark']||''),
      timesheet:    tsMap[sid2] || []
    });
  }
  return out.reverse();
}

// ─── PER-DEPARTMENT ACTIVITY CATEGORY BREAKDOWN (for HOI/IMO) ───
// filter keys supported: { campus, institution, department, weeks[] }
function getDeptCategoryBreakdown(filter) {
  filter = filter || {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tsD  = ss.getSheetByName(SH.TIMESHEET).getDataRange().getValues(), tsH = tsD[0];
  var subD = ss.getSheetByName(SH.SUBMISSION).getDataRange().getValues();
  var facD = ss.getSheetByName(SH.FACULTY).getDataRange().getValues();
  var facMap = _buildFacMap(facD, ss.getSheetByName(SH.STAFF).getDataRange().getValues());
  var sidEmailMap = _buildSidEmailMap(subD);
  var sidWeekMap  = _buildSidWeekMap(subD);

  var tsActI = tsH.indexOf('ActivityType'); if(tsActI<0) tsActI=3;
  var tsSidI = tsH.indexOf('SubmissionID'); if(tsSidI<0) tsSidI=0;
  var tsDayI = tsH.indexOf('Day');          if(tsDayI<0) tsDayI=1;

  var fCampus = (filter.campus||'').trim();
  var fInst   = (filter.institution||'').trim();
  var fDept   = (filter.department||'').trim().toLowerCase();
  var fDeptB  = fDept.replace(/\s*\(pg\)\s*$/,'');
  var weekSet = null;
  if (Array.isArray(filter.weeks) && filter.weeks.length) {
    weekSet = {};
    filter.weeks.forEach(function(w){ weekSet[String(w).trim().slice(0,10)] = true; });
  }

  // dept -> {cats:{}, total:0, filled:0}
  var deptStats = {};
  for (var i = 1; i < tsD.length; i++) {
    var act = String(tsD[i][tsActI]||'').trim();
    var sid = String(tsD[i][tsSidI]||'').trim();
    var email = sidEmailMap[sid]||'';
    var fac = email ? (facMap[email]||{}) : {};
    var dept = String(fac['Department']||'').trim();
    var camp = String(fac['Campus']||'').trim();
    var inst = String(fac['Institution']||'').trim();
    if (!dept) continue;
    if (fCampus && camp !== fCampus) continue;
    if (fInst && inst !== fInst) continue;
    if (fDept) {
      var dL = dept.toLowerCase(), dB = dL.replace(/\s*\(pg\)\s*$/,'');
      if (dL !== fDept && dB !== fDeptB) continue;
    }
    var wk = sidWeekMap[sid] || '';
    if (weekSet && !weekSet[wk]) continue;
    // Skip slots on non-working days (public holidays on any weekday, even Saturdays)
    var _dDayStr = String(tsD[i][tsDayI]||'').trim();
    var _dDayM = _dDayStr.match(/\d+/); var _dDayIdx = _dDayM ? parseInt(_dDayM[0]) - 1 : -1;
    if (_dDayIdx >= 0 && wk && _isNonWorkingDay_(wk, _dDayIdx)) continue;
    if (!deptStats[dept]) deptStats[dept] = {cats:{}, total:0, filled:0};
    deptStats[dept].total++;
    if (!act) continue;
    deptStats[dept].filled++;
    var cat = act.split(' > ')[0];
    if (cat) deptStats[dept].cats[cat] = (deptStats[dept].cats[cat]||0) + 1;
  }
  return { deptStats: deptStats };
}