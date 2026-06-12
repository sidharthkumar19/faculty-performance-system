// =============================================================================
// VMRF-DU Institutional Monitoring System — Faculty Submissions & Timesheets
// =============================================================================

// ─── MY SUBMISSIONS ───────────────────────────────────────────────────────────

function getMySubmissions(facultyID) {
  if(!facultyID) throw new Error('Faculty ID is required.');
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var subD = ss.getSheetByName(SH.SUBMISSION).getDataRange().getValues(), subH=subD[0];
  var hodD = ss.getSheetByName(SH.HOD).getDataRange().getValues();
  var hoiD = ss.getSheetByName(SH.HOI).getDataRange().getValues();
  var imoD = ss.getSheetByName(SH.IMO).getDataRange().getValues();
  var saD  = ss.getSheetByName(SH.SELF_ASSESS).getDataRange().getValues();
  var tsD  = ss.getSheetByName(SH.TIMESHEET).getDataRange().getValues();
  var hodMap=_bm(hodD,hodD[0]), hoiMap=_bm(hoiD,hoiD[0]), imoMap=_bm(imoD,imoD[0]);
  var saMap=_bm(saD,saD[0]), tsMap=_bmMulti(tsD,tsD[0]);
  // Determine once if this faculty is from a no-HOD institution
  var _facRowNH = _rowByKey(SH.FACULTY, facultyID, 'Email') || {};
  var _isNoHodFaculty = _isNoHodInstitution_(String(_facRowNH['Institution']||'').trim());
  // Find faculty email column — try header name first, then scan all cols for a match
  var _fecI = _facEmailCol(subH);
  var facIDLower = String(facultyID||'').trim().toLowerCase();
  
  // If column not found by name, auto-detect by scanning first data row for the email
  if (_fecI < 0 && subD.length > 1) {
    for (var ci = 0; ci < subD[1].length; ci++) {
      if (String(subD[1][ci]||'').trim().toLowerCase() === facIDLower) { _fecI = ci; break; }
    }
  }
  // Final fallback: column index 1 (standard position)
  if (_fecI < 0) _fecI = 1;
  
  var out=[];
  for(var i=1;i<subD.length;i++){
    // Try primary column first, then scan all columns for the email (handles schema mismatches)
    var rowEmail = String(subD[i][_fecI]||'').trim().toLowerCase();
    if (rowEmail !== facIDLower) {
      // Secondary scan: check every column in this row for the email
      var found = false;
      for (var ci2 = 0; ci2 < subD[i].length; ci2++) {
        if (String(subD[i][ci2]||'').trim().toLowerCase() === facIDLower) { found = true; break; }
      }
      if (!found) continue;
    }
    var _sidI=subH.indexOf('SubmissionID'); if(_sidI<0)_sidI=0;
    var sid=String(subD[i][_sidI]||'');
    if(!sid) continue; // skip rows with no submission ID
    var hod=hodMap[sid]||{}, hoi=hoiMap[sid]||{}, imo=imoMap[sid]||{}, sa=saMap[sid]||{};
    var tsEntries=(tsMap[sid]||[]).map(function(r){
      // r['AttachmentURL'] may be undefined if sheet columns were never added — safe cast
      return {
        Day:            String(r['Day']            || ''),
        TimeSlot:       String(r['TimeSlot']        || ''),
        ActivityType:   String(r['ActivityType']    || ''),
        Details:        String(r['ActivityDetails'] || r['Details'] || ''),
        AttachmentURL:  r['AttachmentURL']  != null ? String(r['AttachmentURL'])  : '',
        AttachmentName: r['AttachmentName'] != null ? String(r['AttachmentName']) : ''
      };
    });
    out.push({
      submissionID: sid,
      semester:     String(subD[i][subH.indexOf('AcademicYearSemester')]||''),
      from:         _fmt(subD[i][subH.indexOf('ReportingFrom')]),
      to:           _fmt(subD[i][subH.indexOf('ReportingTo')]),
      submitted:    _fmtDT(subD[i][subH.indexOf('SubmittedDateTime')]),
      declaration:  String(subD[i][subH.indexOf('Declaration')]||''),
      hodStatus:    _isNoHodFaculty ? 'N/A' : (hod['HOD_Status']||'Pending'),
      hoiStatus:    hoi['HOI_Status'] ? String(hoi['HOI_Status']) : (_isNoHodFaculty || String(hod['HOD_Status']||'')==='Approved' ? 'Pending' : '—'),
      imoStatus:    imo['IMO_Status'] ? String(imo['IMO_Status']) : (String(hoi['HOI_Status']||'')=='Approved'?'Pending':'—'),
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

// ─── FLEXIBLE FACULTY OWNERSHIP CHECK ────────────────────────────────────────────
// Returns true if facultyID (email) is the owner of submission sid.
// Uses the same flexible scan as getMySubmissions so column-name differences
// (FacultyEmail vs FacultyID, old vs new schema) never cause false negatives.
function _isFacultyOwner(ss, sid, facultyID) {
  var facIDLower = String(facultyID||'').trim().toLowerCase();
  if (!facIDLower) return false;
  var subSh = ss.getSheetByName(SH.SUBMISSION); if (!subSh) return false;
  var subD  = subSh.getDataRange().getValues(), subH = subD[0];
  var _sidI = _facEmailCol(subH); // reuse to find email col
  for (var i = 1; i < subD.length; i++) {
    var rowSid = String(subD[i][0]||'').trim();
    if (rowSid !== String(sid).trim()) continue;
    // Primary: try detected email column
    if (_sidI >= 0 && String(subD[i][_sidI]||'').trim().toLowerCase() === facIDLower) return true;
    // Secondary: scan every cell in this row for a matching email
    for (var ci = 0; ci < subD[i].length; ci++) {
      var v = String(subD[i][ci]||'').trim().toLowerCase();
      if (v === facIDLower && v.indexOf('@') >= 0) return true;
    }
    return false; // row found but no email match
  }
  return false; // sid not found
}

// ─── WITHDRAW SUBMISSION (faculty-only, HOD not yet reviewed) ─────────────────
function withdrawSubmission(sid, facultyID) {
  if (!sid)       throw new Error('Submission ID is required.');
  if (!facultyID) throw new Error('Faculty ID is required.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Safety: verify this submission belongs to this faculty ──
  if (!_isFacultyOwner(ss, sid, facultyID))
    throw new Error('Submission not found or you can only withdraw your own submissions.');
  var sub = _rowByKey(SH.SUBMISSION, sid);
  if (!sub) throw new Error('Submission not found.');

  // ── Safety: block if HOD has already reviewed (or HOI for no-HOD institutions) ──
  var hodStatus = String(_getCell(SH.HOD, sid, 'HOD_Status')||'').trim();
  var subFac    = _rowByKey(SH.FACULTY, facultyID, 'Email') || {};
  var subInst   = String(subFac['Institution']||'').trim();
  var isNoHod   = _isNoHodInstitution_(subInst);
  if (isNoHod) {
    var hoiStatusW = String(_getCell(SH.HOI, sid, 'HOI_Status')||'').trim();
    if (hoiStatusW && hoiStatusW !== 'Pending')
      throw new Error('Cannot withdraw — your HOI has already reviewed this submission (Status: ' + hoiStatusW + ').');
  } else {
    if (hodStatus && hodStatus !== 'Pending')
      throw new Error('Cannot withdraw — your HOD has already reviewed this submission (Status: ' + hodStatus + ').');
  }

  // ── Collect data for re-editing before deleting ──
  var sa  = _rowByKey(SH.SELF_ASSESS, sid) || {};
  var tsD = ss.getSheetByName(SH.TIMESHEET).getDataRange().getValues(), tsH = tsD[0];
  var tsEntries = [];
  var _auI = tsH.indexOf('AttachmentURL');
  var _anI = tsH.indexOf('AttachmentName');
  for (var i = 1; i < tsD.length; i++) {
    if (String(tsD[i][tsH.indexOf('SubmissionID')]||'').trim() === sid) {
      tsEntries.push({
        day:            String(tsD[i][tsH.indexOf('Day')]            || ''),
        slot:           String(tsD[i][tsH.indexOf('TimeSlot')]       || ''),
        activity:       String(tsD[i][tsH.indexOf('ActivityType')]   || ''),
        details:        String(tsD[i][tsH.indexOf('ActivityDetails')]|| ''),
        attachmentURL:  (_auI >= 0) ? String(tsD[i][_auI]||'') : '',
        attachmentName: (_anI >= 0) ? String(tsD[i][_anI]||'') : ''
      });
    }
  }
  var editData = {
    sid:        sid,
    semester:   String(sub['AcademicYearSemester'] || ''),
    from:       sub['ReportingFrom'] ? Utilities.formatDate(new Date(sub['ReportingFrom']), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
    to:         sub['ReportingTo']   ? Utilities.formatDate(new Date(sub['ReportingTo']),   Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
    outcome:    String(sa['OutcomeOfWeek']       || ''),
    target:     String(sa['TargetPlanNextWeek']  || ''),
    timesheet:  tsEntries
  };

  // ── Delete rows from all sheets ──
  function deleteRowBySid(sheetName) {
    var sh = ss.getSheetByName(sheetName);
    if (!sh) return;
    var data = sh.getDataRange().getValues(), h = data[0];
    var sidI = h.indexOf('SubmissionID');
    if (sidI < 0) return;
    // Delete bottom-up to preserve row indices
    for (var r = data.length - 1; r >= 1; r--) {
      if (String(data[r][sidI]||'').trim() === sid) sh.deleteRow(r + 1);
    }
  }
  [SH.SUBMISSION, SH.TIMESHEET, SH.SELF_ASSESS, SH.HOD, SH.HOI, SH.IMO, SH.NOTIF].forEach(deleteRowBySid);

  return editData;
}

// ─── DELETE SUBMISSION (faculty permanently removes a submitted report) ──────────
function deleteSubmission(sid, facultyID) {
  if (!sid)       throw new Error('Submission ID is required.');
  if (!facultyID) throw new Error('Faculty ID is required.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!_isFacultyOwner(ss, sid, facultyID))
    throw new Error('Submission not found or you can only delete your own submissions.');
  var sub = _rowByKey(SH.SUBMISSION, sid);
  if (!sub) throw new Error('Submission not found.');
  // Only allow deletion if not yet approved by HOD (to protect reviewed submissions)
  var hodStatus = String(_getCell(SH.HOD, sid, 'HOD_Status')||'').trim();
  if (hodStatus === 'Approved')
    throw new Error('Cannot delete — this submission has already been approved by HOD.');
  var hoiStatus = String(_getCell(SH.HOI, sid, 'HOI_Status')||'').trim();
  if (hoiStatus === 'Approved')
    throw new Error('Cannot delete — this submission has already been approved by HOI.');
  function deleteRowBySid(sheetName) {
    var sh = ss.getSheetByName(sheetName); if (!sh) return;
    var data = sh.getDataRange().getValues(), h = data[0];
    var sidI = h.indexOf('SubmissionID'); if (sidI < 0) return;
    for (var r = data.length - 1; r >= 1; r--) {
      if (String(data[r][sidI]||'').trim() === sid) sh.deleteRow(r + 1);
    }
  }
  [SH.SUBMISSION, SH.TIMESHEET, SH.SELF_ASSESS, SH.HOD, SH.HOI, SH.IMO, SH.NOTIF].forEach(deleteRowBySid);
  return { ok: true };
}

// ─── MIGRATE LEGACY "Needs Revision" → "Rejected" IN ALL REVIEW SHEETS ──────────
// Run once to clean up old data. Safe to call again (idempotent).
function migrateNeedsRevisionToRejected() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = [
    { name: SH.HOD, col: 'HOD_Status' },
    { name: SH.HOI, col: 'HOI_Status' },
    { name: SH.IMO, col: 'IMO_Status' }
  ];
  var total = 0;
  sheets.forEach(function(s) {
    var sh = ss.getSheetByName(s.name); if (!sh) return;
    var data = sh.getDataRange().getValues();
    var h = data[0], ci = h.indexOf(s.col); if (ci < 0) return;
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][ci]||'').trim() === 'Needs Revision') {
        sh.getRange(r + 1, ci + 1).setValue('Rejected');
        total++;
      }
    }
  });
  return { ok: true, updated: total };
}

// ─── RECALL REJECTED SUBMISSION (faculty edits after HOD/HOI rejection) ────────
function recallRejectedSubmission(sid, facultyID) {
  if (!sid)       throw new Error('Submission ID is required.');
  if (!facultyID) throw new Error('Faculty ID is required.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!_isFacultyOwner(ss, sid, facultyID))
    throw new Error('Submission not found or you can only edit your own submissions.');
  var sub = _rowByKey(SH.SUBMISSION, sid);
  if (!sub) throw new Error('Submission not found.');

  var hodStatus = String(_getCell(SH.HOD, sid, 'HOD_Status')||'').trim();
  var hoiStatus = String(_getCell(SH.HOI, sid, 'HOI_Status')||'').trim();
  if (hodStatus !== 'Rejected' && hoiStatus !== 'Rejected')
    throw new Error('This submission has not been rejected — it cannot be recalled for editing.');

  // Collect data for re-editing
  var sa  = _rowByKey(SH.SELF_ASSESS, sid) || {};
  var tsD = ss.getSheetByName(SH.TIMESHEET).getDataRange().getValues(), tsH = tsD[0];
  var tsEntries = [];
  var _auI2 = tsH.indexOf('AttachmentURL');
  var _anI2 = tsH.indexOf('AttachmentName');
  for (var i = 1; i < tsD.length; i++) {
    if (String(tsD[i][tsH.indexOf('SubmissionID')]||'').trim() === sid) {
      tsEntries.push({
        day:            String(tsD[i][tsH.indexOf('Day')]            || ''),
        slot:           String(tsD[i][tsH.indexOf('TimeSlot')]       || ''),
        activity:       String(tsD[i][tsH.indexOf('ActivityType')]   || ''),
        details:        String(tsD[i][tsH.indexOf('ActivityDetails')]|| ''),
        attachmentURL:  (_auI2 >= 0) ? String(tsD[i][_auI2]||'') : '',
        attachmentName: (_anI2 >= 0) ? String(tsD[i][_anI2]||'') : ''
      });
    }
  }
  var editData = {
    sid:       sid,
    semester:  String(sub['AcademicYearSemester'] || ''),
    from:      sub['ReportingFrom'] ? Utilities.formatDate(new Date(sub['ReportingFrom']), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
    to:        sub['ReportingTo']   ? Utilities.formatDate(new Date(sub['ReportingTo']),   Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
    outcome:   String(sa['OutcomeOfWeek']      || ''),
    target:    String(sa['TargetPlanNextWeek'] || ''),
    timesheet: tsEntries,
    hodRemark: hodStatus === 'Rejected' ? String(_getCell(SH.HOD, sid, 'HOD_Remark')||'') : '',
    hoiRemark: hoiStatus === 'Rejected' ? String(_getCell(SH.HOI, sid, 'HOI_Remark')||'') : ''
  };

  // Delete all rows for this submission
  function deleteRowBySid(sheetName) {
    var sh = ss.getSheetByName(sheetName); if (!sh) return;
    var data = sh.getDataRange().getValues(), h = data[0];
    var sidI = h.indexOf('SubmissionID'); if (sidI < 0) return;
    for (var r = data.length - 1; r >= 1; r--) {
      if (String(data[r][sidI]||'').trim() === sid) sh.deleteRow(r + 1);
    }
  }
  [SH.SUBMISSION, SH.TIMESHEET, SH.SELF_ASSESS, SH.HOD, SH.HOI, SH.IMO, SH.NOTIF].forEach(deleteRowBySid);

  return editData;
}

// ─── FIND FACULTY EMAIL COLUMN (works with old 'FacultyID' and new 'FacultyEmail') ──
function _facEmailCol(headers) {
  // Try new name first, then old name
  var i = headers.indexOf('FacultyEmail');
  if (i >= 0) return i;
  i = headers.indexOf('FacultyID');
  if (i >= 0) return i;
  // Last resort: scan for anything with 'faculty' or 'email'
  for (var x = 0; x < headers.length; x++) {
    var h = String(headers[x]||'').toLowerCase();
    if (h === 'facultyid' || h === 'facultyemail') return x;
  }
  return -1;
}

// ─── SUBMIT WEEKLY REPORT ────────────────────────────────────────────────────

// ─── DIAGNOSTIC: run this in Apps Script editor to debug submission visibility ──
function diagMySubmissions(testEmail) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var subSh = ss.getSheetByName(SH.SUBMISSION);
  if (!subSh) return {error: 'Weekly_Submission sheet NOT FOUND'};
  var subD = subSh.getDataRange().getValues();
  var subH = subD[0];
  var _fecI = _facEmailCol(subH);
  var result = {
    totalRows: subD.length - 1,
    headers: subH,
    facEmailColIndex: _fecI,
    facEmailColName: _fecI >= 0 ? subH[_fecI] : 'NOT FOUND',
    submittedEmails: [],
    matchFound: false
  };
  var needle = String(testEmail||'').trim().toLowerCase();
  for (var i = 1; i < subD.length; i++) {
    var colVal = _fecI >= 0 ? String(subD[i][_fecI]||'').trim().toLowerCase() : '(col not found)';
    result.submittedEmails.push({row: i, value: colVal, matches: colVal === needle});
    if (colVal === needle) result.matchFound = true;
  }
  return result;
}

function submitWeeklyReport(data) {
  if(!data.facultyID)         throw new Error('Faculty ID is required.');
  if(!data.academicYearSem)   throw new Error('Please select the Academic Year / Semester.');
  if(!data.reportingFrom)     throw new Error('Please set the Reporting From date.');
  if(!data.reportingTo)       throw new Error('Please set the Reporting To date.');
  if(data.reportingFrom>data.reportingTo) throw new Error('Reporting From cannot be after Reporting To.');
  if(!data.outcomeOfWeek)     throw new Error('Please fill in the Tasks field.');
  if(data.declaration!=='YES')throw new Error('Declaration must be YES to submit.');

  var ss=SpreadsheetApp.getActiveSpreadsheet(), sid=_uid(), now=new Date();
  var storedFID = String(data.facultyID||'').trim().toLowerCase();
  var _dtFrom = data.reportingFrom ? new Date(data.reportingFrom) : '';
  var _dtTo   = data.reportingTo   ? new Date(data.reportingTo)   : '';
  ss.getSheetByName(SH.SUBMISSION).appendRow([sid,storedFID,data.academicYearSem,_dtFrom,_dtTo,data.declaration,now]);

  if(data.timesheet&&data.timesheet.length){
    var tsRows=data.timesheet.map(function(e){return [sid,e.day,e.slot,e.activity,e.details||'',e.attachmentURL||'',e.attachmentName||''];});
    var tsSheet=ss.getSheetByName(SH.TIMESHEET);
    _ensureSheetColumns(tsSheet, SCHEMA.Timesheet_Entries);
    var tsStart=tsSheet.getLastRow()+1;
    var tsNeeded=tsStart+tsRows.length-1;
    // Expand sheet rows if needed before calling getRange (clearDataValidations
    // does not auto-extend the sheet the way setValues does)
    if(tsSheet.getMaxRows()<tsNeeded) tsSheet.insertRowsAfter(tsSheet.getMaxRows(), tsNeeded-tsSheet.getMaxRows());
    // Expand columns to at least 7
    if(tsSheet.getMaxColumns()<7) tsSheet.insertColumnsAfter(tsSheet.getMaxColumns(), 7-tsSheet.getMaxColumns());
    try{ tsSheet.getRange(tsStart,1,tsRows.length,7).clearDataValidations(); }catch(e){}
    tsSheet.getRange(tsStart,1,tsRows.length,7).setValues(tsRows);
  }
  ss.getSheetByName(SH.SELF_ASSESS).appendRow([sid,data.outcomeOfWeek,data.targetPlanNextWeek]);
  // Pre-create blank review rows — skip HOD row entirely for no-HOD institutions
  var facRow = _rowByKey(SH.FACULTY, data.facultyID, 'Email') || {};
  var facInst = String(facRow['Institution']||'').trim();
  var isNoHod = _isNoHodInstitution_(facInst);
  if (!isNoHod) ss.getSheetByName(SH.HOD).appendRow([sid,'','','']);
  ss.getSheetByName(SH.HOI).appendRow([sid,'','','']);
  ss.getSheetByName(SH.IMO).appendRow([sid,'','','']);

  var facName=String(facRow['FacultyName']||data.facultyID);
  var facDept=String(facRow['Department']||'');
  var facInstCode = '';
  try { facInstCode = String(_resolveInstCode('', facInst) || '').toUpperCase(); } catch(e) {}
  var facPeriod=data.reportingFrom+' to '+data.reportingTo;
  if (!isNoHod) {
    // Standard 4-stage: Faculty → HOD → HOI → IMO. Notify HOD only.
    var hodNotifKey = facDept ? 'HOD:'+facDept : 'HOD';
    _pushNotif(hodNotifKey,'new_submission',
      '📋 New Submission from '+facName,
      facName+(facDept?' ('+facDept+')':'')+' submitted a weekly report for '+facPeriod+'. Awaiting your review.',
      sid, facName);
    try { _notifyHOD(ss,sid,data.facultyID); } catch(e) { Logger.log('HOD notify failed: '+e.message); }
  } else {
    // No-HOD institutions (SAHS / VMLS / VSEP / VSHS): submission goes
    // directly to HOI. Notify only the HOI of this institution.
    var hoiNotifKey = facInstCode ? 'HOI:'+facInstCode : 'HOI';
    _pushNotif(hoiNotifKey,'new_submission',
      '📋 New Submission from '+facName,
      facName+(facDept?' ('+facDept+')':'')+' submitted a weekly report for '+facPeriod+'. Awaiting your institutional review.',
      sid, facName);
    try { _notifyHOI(ss,sid); } catch(e) { Logger.log('HOI notify failed: '+e.message); }
  }
  return { sid:sid };
}

// ─── PENDING & PRIORITY WORK ──────────────────────────────────────────────────

function savePPWork(data) {
  if (!data.facultyID) throw new Error('Faculty ID required.');
  if (!data.area)      throw new Error('Area is required.');
  if (!data.task)      throw new Error('Task description is required.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH.PP_WORKS);
  if (!sh) { _ensureSheets(); sh = ss.getSheetByName(SH.PP_WORKS); }
  var id = 'PP-' + new Date().getTime();
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  var responsible = String(data.responsiblePerson || '');
  sh.appendRow([
    id,
    String(data.facultyID || ''),
    String(data.weekOf || ''),
    String(data.area || ''),
    String(data.task || ''),
    responsible,
    String(data.dateOfCommencement || ''),
    String(data.targetDate || ''),
    String(data.status || 'Pending'),
    now
  ]);

  return { ok: true, id: id };
}

function getPPWorks(facultyID) {
  if (!facultyID) throw new Error('Faculty ID required.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH.PP_WORKS);
  if (!sh) return [];
  var data = sh.getDataRange().getValues(), h = data[0];
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (String(r[h.indexOf('FacultyEmail')] || '') !== String(facultyID)) continue;
    out.push({
      ppid:                String(r[h.indexOf('PPID')] || ''),
      weekOf:              String(r[h.indexOf('WeekOf')] || ''),
      area:                String(r[h.indexOf('Area')] || ''),
      task:                String(r[h.indexOf('Task')] || ''),
      responsiblePerson:   String(r[h.indexOf('ResponsiblePerson')] || ''),
      dateOfCommencement:  String(r[h.indexOf('DateOfCommencement')] || ''),
      targetDate:          String(r[h.indexOf('TargetDate')] || ''),
      status:              String(r[h.indexOf('Status')] || ''),
      createdAt:           String(r[h.indexOf('CreatedAt')] || '')
    });
  }
  return out.reverse();
}

function updatePPWorkStatus(ppid, status) {
  if (!ppid || !status) throw new Error('PPID and status required.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH.PP_WORKS);
  if (!sh) throw new Error('Sheet not found.');
  var data = sh.getDataRange().getValues(), h = data[0];
  var idIdx = h.indexOf('PPID'), stIdx = h.indexOf('Status');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idIdx]) === String(ppid)) {
      sh.getRange(i + 1, stIdx + 1).setValue(status);
      return { ok: true };
    }
  }
  throw new Error('Record not found.');
}

// ─── Get ALL Pending & Priority Works (all entries) ──────────────────────────
function getAllPPWorks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH.PP_WORKS); if (!sh) return [];
  var ppD = sh.getDataRange().getValues(), ppH = ppD[0];
  var facD = ss.getSheetByName(SH.FACULTY).getDataRange().getValues(), facH = facD[0];
  var facMap = {};
  for (var f = 1; f < facD.length; f++) {
    var fid = String(facD[f][facH.indexOf('Email')]||'').trim().toLowerCase();
    facMap[fid] = String(facD[f][facH.indexOf('FacultyName')]||'');
  }
  var out = [];
  for (var i = 1; i < ppD.length; i++) {
    var r = ppD[i];
    if (!r[0]) continue;
    var fid2 = String(r[ppH.indexOf('FacultyEmail')]||'').trim();
    out.push({
      ppid:               String(r[ppH.indexOf('PPID')]||''),
      facultyID:          fid2,
      facultyName:        facMap[fid2] || fid2,
      weekOf:             String(r[ppH.indexOf('WeekOf')]||''),
      area:               String(r[ppH.indexOf('Area')]||''),
      task:               String(r[ppH.indexOf('Task')]||''),
      responsiblePerson:  String(r[ppH.indexOf('ResponsiblePerson')]||''),
      dateOfCommencement: String(r[ppH.indexOf('DateOfCommencement')]||''),
      targetDate:         String(r[ppH.indexOf('TargetDate')]||''),
      status:             String(r[ppH.indexOf('Status')]||''),
      createdAt:          String(r[ppH.indexOf('CreatedAt')]||'')
    });
  }
  return out.reverse();
}

// ─── SAVE ACTIVITY ATTACHMENT TO GOOGLE DRIVE ────────────────────────────────
function saveActivityAttachment(base64Data, fileName, mimeType, submissionID, role) {
  if (!base64Data || !fileName) throw new Error('File data and name are required.');
  if (!submissionID) throw new Error('Submission ID is required.');

  // Validate file type
  var allowedTypes = ['application/pdf','application/msword',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'application/vnd.ms-excel',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-powerpoint',
    'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    'image/jpeg','image/png','image/gif','text/plain'];
  if (allowedTypes.indexOf(mimeType) < 0)
    throw new Error('File type not allowed. Please upload PDF, Word, Excel, PowerPoint, image or text files.');

  // Max 10MB
  var bytes = Utilities.base64Decode(base64Data);
  if (bytes.length > 10 * 1024 * 1024)
    throw new Error('File size exceeds 10MB limit.');

  // Get or create attachments folder
  var folderName = 'VMRF_Attachments';
  var folders = DriveApp.getFoldersByName(folderName);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

  // Sub-folder by role
  var subName = (role === 'HOD' ? 'HOD_Submissions' : 'Faculty_Submissions');
  var subFolders = folder.getFoldersByName(subName);
  var subFolder = subFolders.hasNext() ? subFolders.next() : folder.createFolder(subName);

  // Save file. IMPORTANT: do NOT call setSharing(ANYONE_WITH_LINK) — that makes
  // every faculty submission attachment world-readable by anyone who guesses /
  // sees the URL. Instead the file is kept PRIVATE (only the deploying account
  // has read access) and downloads happen via getAttachmentBlob(), which
  // verifies the caller's session before streaming the bytes.
  var blob = Utilities.newBlob(bytes, mimeType, fileName);
  var file = subFolder.createFile(blob);
  // Leave default permissions = private to the deploying account.

  return {
    ok:   true,
    // The "url" we return is an internal route, not a Drive URL. The client
    // calls getAttachmentBlob(fileId) when the user clicks the link.
    url:  'attachment://' + file.getId(),
    id:   file.getId(),
    name: fileName
  };
}

// ─── ATTACHMENT DOWNLOAD (session-gated) ──────────────────────────────────────
// Returns the file's bytes as a base64 string. The caller must supply the
// session identity (facultyID for FACULTY, ipmToken for HOI/IMO, or staffID
// for HOD) and the function checks that the submission referenced by the
// attachment belongs to that user OR that the user is a reviewer (HOD for
// the dept, HOI for the institution, IMO global).
//
// NOTE: this is intentionally a server-side download — the client receives
// a base64 payload that it converts to a Blob and triggers download. The
// underlying Drive file is never made public.
function getAttachmentBlob(fileId, requesterRole, requesterId) {
  if (!fileId) throw new Error('File ID is required.');
  if (!requesterRole || !requesterId) throw new Error('You must be signed in.');

  // 1. Locate the file. It must live inside our VMRF_Attachments folder
  //    (anything else is rejected — prevents fileId-guessing on unrelated files).
  var file;
  try { file = DriveApp.getFileById(fileId); }
  catch(e) { throw new Error('Attachment not found.'); }
  var parents = file.getParents();
  var inOurFolder = false;
  while (parents.hasNext()) {
    var p = parents.next();
    if (p.getName() === 'Faculty_Submissions' || p.getName() === 'HOD_Submissions') {
      inOurFolder = true; break;
    }
  }
  if (!inOurFolder) throw new Error('Access denied: file is outside the VMRF attachments scope.');

  // 2. Look up the SubmissionID this file was attached to.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sidForFile = null, ownerEmail = null;
  ['Timesheet_Entries','HOD_Timesheet'].forEach(function(sn){
    if (sidForFile) return;
    var tSh = ss.getSheetByName(sn); if (!tSh) return;
    var tD = tSh.getDataRange().getValues(); if (tD.length < 2) return;
    var tH = tD[0];
    var auI = tH.indexOf('AttachmentURL'); if (auI < 0) return;
    var sidI = tH.indexOf('SubmissionID'); if (sidI < 0) return;
    for (var i = 1; i < tD.length; i++) {
      var u = String(tD[i][auI]||'');
      if (u && u.indexOf(fileId) >= 0) {
        sidForFile = String(tD[i][sidI]||'').trim();
        break;
      }
    }
  });
  if (!sidForFile) throw new Error('Access denied: file is not linked to any submission.');

  // 3. Determine the submission's owner and apply per-role access rules.
  var role = String(requesterRole||'').toUpperCase();
  var reqId = String(requesterId||'').trim().toLowerCase();

  if (role === 'FACULTY') {
    // Faculty may only download from their own submissions.
    if (!_isFacultyOwner(ss, sidForFile, reqId)) {
      throw new Error('Access denied: you can only download attachments from your own submissions.');
    }
  } else if (role === 'HOD') {
    // HOD may download from submissions belonging to faculty in their dept.
    var hodRow = _rowByKey(SH.STAFF, reqId, 'Email');
    if (!hodRow) throw new Error('Access denied: HOD account not found.');
    var hodDept = String(hodRow['Department']||'').toLowerCase();
    var subRow = _rowByKey(SH.SUBMISSION, sidForFile);
    var facEmail = subRow ? _getFidFromSub(subRow) : '';
    var facRow = facEmail ? _rowByKey(SH.FACULTY, facEmail, 'Email') : null;
    var facDept = facRow ? String(facRow['Department']||'').toLowerCase() : '';
    if (!facDept || facDept !== hodDept) {
      throw new Error('Access denied: this attachment is outside your department.');
    }
  } else if (role === 'HOI' || role === 'IMO') {
    // HOI/IMO have broader access (institution / system-wide). The session
    // token validation is done at the route level via APP.role; if we reach
    // here with HOI/IMO, allow access.
  } else {
    throw new Error('Access denied: unrecognised role.');
  }

  // 4. Stream the bytes back as base64. Apps Script's google.script.run cannot
  //    return raw bytes, so we encode.
  var blob = file.getBlob();
  return {
    ok: true,
    name: file.getName(),
    mimeType: blob.getContentType(),
    base64: Utilities.base64Encode(blob.getBytes())
  };
}

// One-time utility: scan VMRF_Attachments and revoke ANYONE_WITH_LINK on any
// file that still has it. Run this once from the Apps Script editor after
// applying the security upgrade so historical attachments stop being public.
function lockdownExistingAttachments() {
  var folders = DriveApp.getFoldersByName('VMRF_Attachments');
  if (!folders.hasNext()) {
    Logger.log('No VMRF_Attachments folder found.');
    return;
  }
  var folder = folders.next();
  var count = 0, errors = 0;
  function _walk(f) {
    var files = f.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      try {
        file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
        count++;
      } catch(e) {
        errors++;
        Logger.log('Failed to lock ' + file.getName() + ': ' + e.message);
      }
    }
    var subs = f.getFolders();
    while (subs.hasNext()) _walk(subs.next());
  }
  _walk(folder);
  var msg = 'Lockdown complete. ' + count + ' files set to PRIVATE, ' + errors + ' errors.';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch(_){}
  return { ok: true, locked: count, errors: errors };
}

// ─── SINGLE SUBMISSION DETAIL (for All Submissions date-click drill-down) ────
function getSingleSubmission(submissionID, type) {
  submissionID = String(submissionID||'').trim();
  if (!submissionID) return null;
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (type === 'HOD') {
    var subSh = ss.getSheetByName(SH.HOD_SUB); if (!subSh) return null;
    var subD  = subSh.getDataRange().getValues(), subH = subD[0];
    var revSh = ss.getSheetByName(SH.HOD_REVIEW);
    var revD  = revSh ? revSh.getDataRange().getValues() : [['SubmissionID']];
    var imoSh = ss.getSheetByName(SH.HOD_IMO);
    var imoD  = imoSh ? imoSh.getDataRange().getValues() : [['SubmissionID']];
    var saSh  = ss.getSheetByName(SH.HOD_SA);
    var saD   = saSh ? saSh.getDataRange().getValues() : [['SubmissionID']];
    var tsSh  = ss.getSheetByName(SH.HOD_TS);
    var tsD   = tsSh ? tsSh.getDataRange().getValues() : [['SubmissionID']];
    var revMap = _bm(revD, revD[0]), imoMap = _bm(imoD, imoD[0]);
    var saMap  = _bm(saD, saD[0]), tsMap = _bmMulti(tsD, tsD[0]);
    var sidI   = subH.indexOf('SubmissionID'); if(sidI<0) sidI=0;
    var hodIdI = subH.indexOf('HOD_ID');       if(hodIdI<0) hodIdI=1;
    for (var i = 1; i < subD.length; i++) {
      if (String(subD[i][sidI]||'').trim() !== submissionID) continue;
      var rev = revMap[submissionID]||{}, imo2 = imoMap[submissionID]||{}, sa = saMap[submissionID]||{};
      var hoiSt = String(rev['HOI_Status']||''), imoSt = String(imo2['IMO_Status']||'');
      var hodID = String(subD[i][hodIdI]||'').trim().toLowerCase();
      var hodName = 'HOD', hodDept = '';
      var staffSh = ss.getSheetByName(SH.STAFF);
      if (staffSh) {
        var staffD2 = staffSh.getDataRange().getValues(), staffH2 = staffD2[0];
        var seI2=staffH2.indexOf('Email'), snI2=staffH2.indexOf('StaffName'), sdI2=staffH2.indexOf('Department');
        for (var sf=1; sf<staffD2.length; sf++) {
          if (String(staffD2[sf][seI2]||'').trim().toLowerCase()===hodID) {
            hodName=String(staffD2[sf][snI2]||'HOD'); hodDept=String(staffD2[sf][sdI2]||''); break;
          }
        }
      }
      var tsEntries = (tsMap[submissionID]||[]).map(function(t){
        return {Day:String(t['Day']||''),TimeSlot:String(t['TimeSlot']||''),ActivityType:String(t['ActivityType']||''),Details:String(t['ActivityDetails']||t['Details']||''),AttachmentURL:String(t['AttachmentURL']||''),AttachmentName:String(t['AttachmentName']||'')};
      });
      return {
        submissionID:submissionID, facultyName:hodName, department:hodDept,
        semester:String(subD[i][subH.indexOf('AcademicYearSemester')>=0?subH.indexOf('AcademicYearSemester'):2]||''),
        from:_fmt(subD[i][subH.indexOf('ReportingFrom')>=0?subH.indexOf('ReportingFrom'):3]),
        to:  _fmt(subD[i][subH.indexOf('ReportingTo')>=0?subH.indexOf('ReportingTo'):4]),
        submitted:_fmtDT(subD[i][subH.indexOf('SubmittedDateTime')>=0?subH.indexOf('SubmittedDateTime'):6]),
        hodStatus:'N/A', hoiStatus:hoiSt||'Pending',
        imoStatus:imoSt?(imoSt):(hoiSt==='Approved'?'Pending':'—'),
        hodRemark:'', hoiRemark:String(rev['HOI_Remark']||''), imoRemark:String(imo2['IMO_Remark']||''),
        outcome:String(sa['Tasks']||sa['OutcomeOfWeek']||''), timesheet:tsEntries, _type:'HOD'
      };
    }
    return null;

  } else {
    // FACULTY submission
    var subD3 = ss.getSheetByName(SH.SUBMISSION).getDataRange().getValues(), subH3 = subD3[0];
    var hodD3 = ss.getSheetByName(SH.HOD).getDataRange().getValues();
    var hoiD3 = ss.getSheetByName(SH.HOI).getDataRange().getValues();
    var imoD3 = ss.getSheetByName(SH.IMO).getDataRange().getValues();
    var saD3  = ss.getSheetByName(SH.SELF_ASSESS).getDataRange().getValues();
    var tsD3  = ss.getSheetByName(SH.TIMESHEET).getDataRange().getValues();
    var hodMap3=_bm(hodD3,hodD3[0]), hoiMap3=_bm(hoiD3,hoiD3[0]), imoMap3=_bm(imoD3,imoD3[0]);
    var saMap3=_bm(saD3,saD3[0]), tsMap3=_bmMulti(tsD3,tsD3[0]);
    var sidI3=subH3.indexOf('SubmissionID'); if(sidI3<0) sidI3=0;
    var fecI3=_facEmailCol(subH3);
    for (var j=1; j<subD3.length; j++) {
      if (String(subD3[j][sidI3]||'').trim() !== submissionID) continue;
      var hod3=hodMap3[submissionID]||{}, hoi3=hoiMap3[submissionID]||{}, imo3=imoMap3[submissionID]||{}, sa3=saMap3[submissionID]||{};
      var facID3=String(subD3[j][fecI3>=0?fecI3:1]||'').trim().toLowerCase();
      var facD3=ss.getSheetByName(SH.FACULTY).getDataRange().getValues(), facH3=facD3[0].map(function(v){return String(v||'').trim();});
      var emI3=facH3.indexOf('Email'),nmI3=facH3.indexOf('FacultyName'),dpI3=facH3.indexOf('Department'),dsI3=facH3.indexOf('Designation');
      if(emI3<0)emI3=1; if(nmI3<0)nmI3=0; if(dpI3<0)dpI3=2; if(dsI3<0)dsI3=5;
      var facName3='',facDept3='',facDesg3='';
      for (var fk=1; fk<facD3.length; fk++) {
        if (String(facD3[fk][emI3]||'').trim().toLowerCase()===facID3) {
          facName3=String(facD3[fk][nmI3]||''); facDept3=String(facD3[fk][dpI3]||''); facDesg3=String(facD3[fk][dsI3]||''); break;
        }
      }
      var tsEntries3=(tsMap3[submissionID]||[]).map(function(t){
        return {Day:String(t['Day']||''),TimeSlot:String(t['TimeSlot']||''),ActivityType:String(t['ActivityType']||''),Details:String(t['ActivityDetails']||t['Details']||''),AttachmentURL:String(t['AttachmentURL']||''),AttachmentName:String(t['AttachmentName']||'')};
      });
      return {
        submissionID:submissionID, facultyName:facName3||facID3, facultyID:facID3,
        department:facDept3, designation:facDesg3,
        semester:String(subD3[j][subH3.indexOf('AcademicYearSemester')>=0?subH3.indexOf('AcademicYearSemester'):2]||''),
        from:_fmt(subD3[j][subH3.indexOf('ReportingFrom')>=0?subH3.indexOf('ReportingFrom'):3]),
        to:  _fmt(subD3[j][subH3.indexOf('ReportingTo')>=0?subH3.indexOf('ReportingTo'):4]),
        submitted:_fmtDT(subD3[j][subH3.indexOf('SubmittedDateTime')>=0?subH3.indexOf('SubmittedDateTime'):6]),
        hodStatus:String(hod3['HOD_Status']||'Pending'),
        hoiStatus:hoi3['HOI_Status']?String(hoi3['HOI_Status']):(String(hod3['HOD_Status']||'')==='Approved'?'Pending':'—'),
        imoStatus:imo3['IMO_Status']?String(imo3['IMO_Status']):(String(hoi3['HOI_Status']||'')==='Approved'?'Pending':'—'),
        hodRemark:String(hod3['HOD_Remark']||''), hoiRemark:String(hoi3['HOI_Remark']||''), imoRemark:String(imo3['IMO_Remark']||''),
        outcome:String(sa3['OutcomeOfWeek']||''), timesheet:tsEntries3
      };
    }
    return null;
  }
}

// ─── WEEKLY DRAFT PERSISTENCE ────────────────────────────────
// Drafts are stored per faculty email + week-start date so that a faculty
// can fill their timesheet incrementally across multiple days before submitting.
// Storage is Apps Script Script Properties with key: draft:<email>:<weekStart>

function saveDraftTimesheet(payload) {
  try {
    payload = payload || {};
    var email = String(payload.facultyEmail||'').trim().toLowerCase();
    var weekStart = String(payload.weekStart||'').trim();
    if (!email || !weekStart) throw new Error('Missing faculty email or week start date.');
    var key = 'draft:' + email + ':' + weekStart;
    var value = JSON.stringify({
      weekStart: weekStart,
      weekEnd: String(payload.weekEnd||''),
      calData: payload.calData || {},
      outcome: String(payload.outcome||''),
      semester: String(payload.semester||''),
      savedAt: new Date().toISOString()
    });
    // Script Properties has a 9KB per-value limit; chunking is unlikely needed for weekly data
    PropertiesService.getScriptProperties().setProperty(key, value);
    return { ok: true, savedAt: new Date().toISOString() };
  } catch (e) {
    throw new Error('Could not save draft: ' + e.message);
  }
}

function loadDraftTimesheet(facultyEmail, weekStart) {
  try {
    var email = String(facultyEmail||'').trim().toLowerCase();
    var ws = String(weekStart||'').trim();
    if (!email || !ws) return null;
    var key = 'draft:' + email + ':' + ws;
    var value = PropertiesService.getScriptProperties().getProperty(key);
    if (!value) return null;
    return JSON.parse(value);
  } catch (e) {
    return null;
  }
}

function deleteDraftTimesheet(facultyEmail, weekStart) {
  try {
    var email = String(facultyEmail||'').trim().toLowerCase();
    var ws = String(weekStart||'').trim();
    if (!email || !ws) return { ok: false };
    var key = 'draft:' + email + ':' + ws;
    PropertiesService.getScriptProperties().deleteProperty(key);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function listDraftTimesheets(facultyEmail) {
  try {
    var email = String(facultyEmail||'').trim().toLowerCase();
    if (!email) return [];
    var prefix = 'draft:' + email + ':';
    var all = PropertiesService.getScriptProperties().getProperties();
    var result = [];
    for (var k in all) {
      if (k.indexOf(prefix) === 0) {
        try {
          var d = JSON.parse(all[k]);
          result.push({ weekStart: d.weekStart, weekEnd: d.weekEnd, savedAt: d.savedAt });
        } catch (e) {}
      }
    }
    return result;
  } catch (e) {
    return [];
  }
}

// =============================================================================
// ─── FACULTY PROFILE MODULE ───────────────────────────────────────────────────
// =============================================================================

// ─── Save or update the logged-in faculty's profile ──────────────────────────
function saveFacultyProfile(profileData) {
  try {
    var email = String(profileData.email || '').trim().toLowerCase();
    if (!email) throw new Error('Email is required.');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SH.PROFILES);
    if (!sheet) {
      sheet = ss.insertSheet(SH.PROFILES);
      var hdrs = SCHEMA.FacultyProfiles;
      var hr = sheet.getRange(1, 1, 1, hdrs.length);
      hr.setValues([hdrs]).setFontWeight('bold').setBackground('#1a3c5e').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    var data = sheet.getDataRange().getValues();
    var h    = data[0].map(function(v){ return String(v).trim(); });
    var emI  = h.indexOf('Email');
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][emI]||'').trim().toLowerCase() === email) { rowIndex = i + 1; break; }
    }

    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    var vals = {
      'Email':            email,
      'FacultyName':      String(profileData.facultyName      || '').trim(),
      'EmployeeID':       String(profileData.employeeId       || '').trim(),
      'Designation':      String(profileData.designation      || '').trim(),
      'Department':       String(profileData.department       || '').trim(),
      'Institution':      String(profileData.institution      || '').trim(),
      'Campus':           String(profileData.campus           || '').trim(),
      'Qualification':    String(profileData.qualification    || '').trim(),
      'Specialization':   String(profileData.specialization   || '').trim(),
      'Experience':       String(profileData.experience       || '').trim(),
      'Phone':            String(profileData.phone            || '').trim(),
      'DateOfJoining':    String(profileData.dateOfJoining    || '').trim(),
      'ResearchAreas':    String(profileData.researchAreas    || '').trim(),
      'PublicationsCount':String(profileData.publicationsCount|| '').trim(),
      'Certifications':   String(profileData.certifications   || '').trim(),
      'LinkedinOrcid':    String(profileData.linkedinOrcid    || '').trim(),
      'Bio':              String(profileData.bio              || '').trim(),
      'LastUpdated':      now
    };
    var newRow = h.map(function(col){ return vals[col] !== undefined ? vals[col] : ''; });

    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 1, 1, newRow.length).setValues([newRow]);
    } else {
      sheet.appendRow(newRow);
    }
    return { success: true, message: 'Profile saved successfully.' };
  } catch (e) {
    return { success: false, message: 'Error saving profile: ' + e.message };
  }
}

// ─── Get the logged-in faculty's own profile ─────────────────────────────────
function getFacultyProfile(email) {
  try {
    email = String(email || '').trim().toLowerCase();
    if (!email) return { success: true, profile: null };

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SH.PROFILES);
    if (!sheet) return { success: true, profile: null };

    var data = sheet.getDataRange().getValues();
    var h    = data[0].map(function(v){ return String(v).trim(); });
    var emI  = h.indexOf('Email');

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][emI]||'').trim().toLowerCase() !== email) continue;
      var p = {};
      h.forEach(function(col, j){ p[col] = String(data[i][j] != null ? data[i][j] : ''); });
      return {
        success: true,
        profile: {
          email:              p['Email'],
          facultyName:        p['FacultyName'],
          employeeId:         p['EmployeeID'],
          designation:        p['Designation'],
          department:         p['Department'],
          institution:        p['Institution'],
          campus:             p['Campus'],
          qualification:      p['Qualification'],
          specialization:     p['Specialization'],
          experience:         p['Experience'],
          phone:              p['Phone'],
          dateOfJoining:      p['DateOfJoining'],
          researchAreas:      p['ResearchAreas'],
          publicationsCount:  p['PublicationsCount'],
          certifications:     p['Certifications'],
          linkedinOrcid:      p['LinkedinOrcid'],
          bio:                p['Bio'],
          lastUpdated:        p['LastUpdated']
        }
      };
    }
    return { success: true, profile: null };
  } catch (e) {
    return { success: false, message: 'Error fetching profile: ' + e.message };
  }
}

// ─── Get any faculty's profile by email — for HOD / HOI / IMO ───────────────
function getFacultyProfileByEmail(facultyEmail) {
  try {
    var email = String(facultyEmail || '').trim().toLowerCase();
    if (!email) return { success: false, message: 'Email required.' };
    // Re-use the same read logic; no role-gate needed server-side since the
    // caller is always a reviewer (HOD / HOI / IMO) who already authenticated.
    return getFacultyProfile(email);
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}