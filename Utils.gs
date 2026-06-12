// =============================================================================
// VMRF-DU Institutional Monitoring System — Utilities & Helpers
// =============================================================================

function _buildHierarchyForClient_() {
  var out = {};
  Object.keys(INSTITUTION_HIERARCHY).forEach(function(campus){
    var node = INSTITUTION_HIERARCHY[campus];
    var instList = [];
    Object.keys(node.institutions).forEach(function(instName){
      var ins = node.institutions[instName];
      instList.push({ name: instName, code: ins.code, departments: ins.departments||[], noHod: !!ins.noHod });
    });
    out[campus] = { code: node.code, institutions: instList };
  });
  return out;
}

/* Returns a plain object acting as a Set: { sid: true } for every submission
   whose faculty belongs to a no-HOD institution. Used to skip the HOD layer. */
function _buildNoHodSidSet_(sidEmailMap, facMap) {
  var set = {};
  Object.keys(sidEmailMap).forEach(function(sid) {
    var email = sidEmailMap[sid];
    var fac   = (email && facMap[email]) || {};
    var instName = String(fac['Institution']||'').trim();
    if (_isNoHodInstitution_(instName)) set[sid] = true;
  });
  return set;
}

function _hashPwd(pwd) {
  var hash = 5381, s = String(pwd);
  for (var i = 0; i < s.length; i++) { hash = ((hash << 5) + hash) + s.charCodeAt(i); hash = hash & hash; }
  return 'H' + (hash >>> 0).toString(16).toUpperCase();
}

// ── Helper: generate random ID ───────────────────────────────────────────────
function _makeID(prefix) {
  var ch = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789', r = prefix + '-';
  for (var k = 0; k < 6; k++) r += ch[Math.floor(Math.random() * ch.length)];
  return r;
}

// ─── ENSURE FACULTY MASTER COLUMNS ───────────────────────────────────────────
// Adds any missing columns from SCHEMA to an existing Faculty_Master sheet.
// Called before every login so stale sheets from old initializeSystem() runs
// don't cause "PasswordHash column missing" errors.
// ─── ENSURE SHEET HAS ALL REQUIRED COLUMNS (adds missing, never removes) ────
function _ensureSheetColumns(sheet, requiredHeaders) {
  var lastCol = sheet.getLastColumn();
  var existing = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(v){ return String(v).trim(); })
    : [];
  var added = 0;
  requiredHeaders.forEach(function(hdr) {
    if (existing.indexOf(hdr) < 0) {
      var newCol = sheet.getLastColumn() + 1;
      var cell = sheet.getRange(1, newCol);
      cell.setValue(hdr);
      cell.setFontWeight('bold').setBackground('#1a3c5e').setFontColor('#ffffff');
      existing.push(hdr); // keep local list in sync
      added++;
    }
  });
  if (added > 0) sheet.setFrozenRows(1);
  return added;
}

// ─── DEBUG / REPAIR FACULTY ROW ──────────────────────────────────────────────
// Run this from Apps Script editor to see what's actually in the sheet.
// Returns the raw header row and all data rows so you can diagnose misalignment.
function debugFacultySheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.FACULTY);
  if (!sheet) return { error: 'Faculty_Master sheet not found' };
  var data = sheet.getDataRange().getValues();
  var result = { headers: data[0], rows: [] };
  for (var i = 1; i < data.length; i++) {
    var row = {};
    for (var j = 0; j < data[0].length; j++) {
      row[data[0][j] || 'col_' + j] = data[i][j];
    }
    result.rows.push(row);
  }
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

// Rebuilds Faculty_Master with correct canonical headers.
// Detects columns by POSITION if headers are missing/wrong,
// or by NAME if headers exist. Wipes and rewrites cleanly.
// Run once from the Apps Script editor after any schema change.
function repairFacultyMaster() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.FACULTY);
  if (!sheet) { Logger.log('Sheet not found'); return; }

  var canonical = SCHEMA[SH.FACULTY];
  var raw = sheet.getDataRange().getValues();

  Logger.log('=== BEFORE REPAIR ===');
  Logger.log('Row count: ' + raw.length);
  Logger.log('Headers: ' + JSON.stringify(raw[0]));
  for (var d = 1; d < raw.length; d++) Logger.log('Row ' + d + ': ' + JSON.stringify(raw[d]));

  var oldHeaders = raw[0].map(function(h){ return String(h).trim(); });
  var hasHeaders = oldHeaders.indexOf('FacultyName') >= 0;

  var objects = [];
  for (var i = 1; i < raw.length; i++) {
    if (!raw[i].some(function(v){ return v !== ''; })) continue; // skip blank rows
    var obj = {};
    if (hasHeaders) {
      // Map by column name
      for (var j = 0; j < oldHeaders.length; j++) {
        if (oldHeaders[j]) obj[oldHeaders[j]] = raw[i][j];
      }
    } else {
      // Fallback: map by old 7-column positional order (pre-schema-update)
      var pos7 = ['FacultyName','Email','Department','Campus','Institution','Designation'];
      for (var p = 0; p < pos7.length && p < raw[i].length; p++) {
        obj[pos7[p]] = raw[i][p];
      }
    }
    if (obj['FacultyName']) objects.push(obj);
  }

  // Clear and rewrite
  sheet.clearContents();
  var newRows = [canonical];
  for (var k = 0; k < objects.length; k++) {
    var row = canonical.map(function(col){ return objects[k][col] !== undefined ? objects[k][col] : ''; });
    newRows.push(row);
  }
  sheet.getRange(1, 1, newRows.length, canonical.length).setValues(newRows);

  Logger.log('=== AFTER REPAIR ===');
  Logger.log('Faculty_Master repaired. ' + objects.length + ' rows migrated.');
  Logger.log('New headers: ' + canonical.join(', '));
  SpreadsheetApp.getUi().alert('Repair complete! ' + objects.length + ' faculty rows migrated.\nNOTE: Any rows registered before this fix had no password stored — those users must re-register.');
  return { repaired: objects.length, headers: canonical };
}

// ─── PRIVATE UTILITIES ────────────────────────────────────────────────────────
function _uid(){
  return 'SUB-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd-HHmm')+'-'+Math.random().toString(36).slice(2,6).toUpperCase();
}

function _prop(k){ return PropertiesService.getScriptProperties().getProperty(k)||''; }

function _fmt(d){ try{ return d?Utilities.formatDate(new Date(d),Session.getScriptTimeZone(),'dd-MMM-yyyy'):''; }catch(e){ return ''; } }

function _fmtDT(d){ try{ return d?Utilities.formatDate(new Date(d),Session.getScriptTimeZone(),'dd-MMM-yyyy HH:mm'):''; }catch(e){ return ''; } }

// Build a lookup map from a 2D values array keyed by keyCol (default SubmissionID)
// Build a lookup map keyed by any named column — always uses header names, never positions

// ─── ROBUST FACULTY MAP BUILDER ──────────────────────────────────────────────
// Works with both old schema (has FacultyID col) and new schema (no FacultyID)
// Indexes faculty by Email (lowercase) — scans ALL columns for email-like values
// Also merges Staff_Master entries (for HOD department lookups)
function _buildFacMap(facD, staffD) {
  var map = {};
  // ── Faculty_Master ──
  if (facD && facD.length >= 2) {
    var h = facD[0].map(function(v){ return String(v||'').trim(); });
    // Find ALL columns that might contain email addresses
    var emailCols = [];
    var emI = h.indexOf('Email');
    if (emI >= 0) emailCols.push(emI);
    // Also check for any header with 'mail' in the name
    for (var x = 0; x < h.length; x++) {
      var hl = h[x].toLowerCase();
      if ((hl.indexOf('mail') >= 0 || hl.indexOf('email') >= 0) && emailCols.indexOf(x) < 0) emailCols.push(x);
    }
    // If still nothing found, auto-detect: scan first data row for cells containing '@'
    if (emailCols.length === 0 && facD.length > 1) {
      for (var a = 0; a < facD[1].length; a++) {
        if (String(facD[1][a]||'').indexOf('@') >= 0) { emailCols.push(a); }
      }
    }
    // Final fallback: column 1 (standard position after FacultyName)
    if (emailCols.length === 0) emailCols.push(1);

    // Detect name, dept, campus, institution, designation columns by header name
    var nmI  = h.indexOf('FacultyName');  if(nmI<0) nmI=0;
    var depI = h.indexOf('Department');
    var camI = h.indexOf('Campus');
    var insI = h.indexOf('Institution');
    var desI = h.indexOf('Designation');

    for (var i = 1; i < facD.length; i++) {
      var r = facD[i];
      // Build object from all headers
      var obj = {};
      h.forEach(function(col, j){ if(col) obj[col] = r[j]; });
      // Also force-set key fields by detected column index (in case headers are off)
      if (!obj['FacultyName'] && nmI >= 0)  obj['FacultyName']  = String(r[nmI]||'');
      if (!obj['Department']  && depI >= 0) obj['Department']   = String(r[depI]||'');
      if (!obj['Campus']      && camI >= 0) obj['Campus']       = String(r[camI]||'');
      if (!obj['Institution'] && insI >= 0) obj['Institution']  = String(r[insI]||'');
      if (!obj['Designation'] && desI >= 0) obj['Designation']  = String(r[desI]||'');
      // Index by ALL email columns found
      for (var ec = 0; ec < emailCols.length; ec++) {
        var email = String(r[emailCols[ec]]||'').trim().toLowerCase();
        if (email && email.indexOf('@') >= 0 && !map[email]) {
          map[email] = obj;
        }
      }
    }
  }
  // ── Staff_Master (merge HOD/HOI entries so their department is resolvable) ──
  if (staffD && staffD.length >= 2) {
    var sh2 = staffD[0].map(function(v){ return String(v||'').trim(); });
    var sEmI  = sh2.indexOf('Email');
    var sNmI  = sh2.indexOf('StaffName');
    var sDepI = sh2.indexOf('Department');
    var sRlI  = sh2.indexOf('Role');
    // Fallback for Email column
    if (sEmI < 0) {
      for (var sx = 0; sx < sh2.length; sx++) {
        if (sh2[sx].toLowerCase().indexOf('mail') >= 0) { sEmI = sx; break; }
      }
    }
    if (sEmI < 0) sEmI = 2; // standard position in Staff_Master
    if (sEmI >= 0) {
      for (var j = 1; j < staffD.length; j++) {
        var se = String(staffD[j][sEmI]||'').trim().toLowerCase();
        if (!se || map[se]) continue; // don't overwrite faculty entries
        var sobj = {
          'FacultyName': sNmI >= 0 ? String(staffD[j][sNmI]||'') : '',
          'Email':       se,
          'Department':  sDepI >= 0 ? String(staffD[j][sDepI]||'') : '',
          'Designation': sRlI >= 0 ? String(staffD[j][sRlI]||'') : '',
          'Institution': INSTITUTIONS.length ? INSTITUTIONS[0] : '',
          'Campus':      CAMPUSES.length ? CAMPUSES[0] : ''
        };
        map[se] = sobj;
      }
    }
  }
  return map;
}

function _bmByCol(data, keyCol) {
  var h=data[0], kI=h.indexOf(keyCol), map={};
  if(kI<0) return map;
  for(var i=1;i<data.length;i++){
    var k=String(data[i][kI]||'').trim();
    if(k){
      var o={}; h.forEach(function(c,j){if(c)o[String(c)]=data[i][j];});
      map[k]=o;
      // Also store lowercase key for case-insensitive lookup
      var kl=k.toLowerCase(); if(kl!==k) map[kl]=o;
    }
  }
  return map;
}

function _bm(data, headers, keyCol) {
  keyCol = keyCol || 'SubmissionID';
  var kI = headers.indexOf(keyCol), map = {};
  for(var i=1;i<data.length;i++){
    var k=String(data[i][kI]||'');
    if(k){ var o={}; headers.forEach(function(c,j){o[c]=data[i][j];}); map[k]=o; }
  }
  return map;
}

// Build a multi-row lookup map (one key → array of row objects)
function _bmMulti(data, headers) {
  var kI=headers.indexOf('SubmissionID'), map={};
  for(var i=1;i<data.length;i++){
    var k=String(data[i][kI]||'');
    if(!k) continue;
    var o={}; headers.forEach(function(c,j){o[c]=data[i][j];});
    if(!map[k]) map[k]=[];
    map[k].push(o);
  }
  return map;
}

// ─── BULLETPROOF: SubmissionID → Faculty Email map ──────────────────────────
// Scans Weekly_Submission data to extract the faculty email for each submission.
// Tries header-based detection first, then scans every cell for '@'.
function _buildSidEmailMap(subD) {
  if (!subD || subD.length < 2) return {};
  var h = subD[0].map(function(v){ return String(v||'').trim(); });
  // Find SubmissionID column
  var sidI = h.indexOf('SubmissionID'); if (sidI < 0) sidI = 0;
  // Find email column using _facEmailCol, then fallback scan
  var emI = _facEmailCol(h);
  if (emI < 0) {
    // Scan first data row for '@'
    for (var c = 0; c < subD[1].length; c++) {
      if (String(subD[1][c]||'').indexOf('@') >= 0) { emI = c; break; }
    }
  }
  if (emI < 0) emI = 1; // ultimate fallback: column B
  var map = {};
  for (var i = 1; i < subD.length; i++) {
    var sid = String(subD[i][sidI]||'').trim();
    if (!sid) continue;
    var email = String(subD[i][emI]||'').trim().toLowerCase();
    // If detected column doesn't have '@', scan entire row
    if (!email || email.indexOf('@') < 0) {
      for (var rc = 0; rc < subD[i].length; rc++) {
        var rv = String(subD[i][rc]||'').trim().toLowerCase();
        if (rv.indexOf('@') >= 0) { email = rv; break; }
      }
    }
    if (sid && email) map[sid] = email;
  }
  return map;
}

// ─── Extract faculty email from a sub object (built by _bm) ─────────────────
// Tries all known key names, then scans all values for '@'.
function _getFidFromSub(sub) {
  if (!sub) return '';
  // Try all known header names
  var fid = String(sub['FacultyEmail']||sub['FacultyID']||sub['FacultyId']||sub['Email']||sub['email']||'').trim().toLowerCase();
  if (fid && fid.indexOf('@') >= 0) return fid;
  // Brute-force: scan all values of the sub object for an email
  var keys = Object.keys(sub);
  for (var i = 0; i < keys.length; i++) {
    var v = String(sub[keys[i]]||'').trim().toLowerCase();
    if (v.indexOf('@') >= 0) return v;
  }
  return fid || '';
}

function _rowByKey(sheetName, keyVal, keyCol) {
  keyCol = keyCol || 'SubmissionID';
  var data=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getDataRange().getValues();
  var h=data[0], kI=h.indexOf(keyCol);
  if (kI < 0) return null;
  // Email columns get a case-insensitive comparison. Faculty_Master /
  // Staff_Master often store mixed-case emails while _getFidFromSub and
  // _buildSidEmailMap force lowercase — without this the lookup silently
  // returns null and the no-HOD detection downstream fails.
  var ci = String(keyCol||'').toLowerCase().indexOf('email') >= 0;
  var needle = ci ? String(keyVal||'').trim().toLowerCase() : String(keyVal);
  for(var i=1;i<data.length;i++){
    var cell = ci ? String(data[i][kI]||'').trim().toLowerCase() : String(data[i][kI]);
    if(cell===needle){ var o={}; h.forEach(function(c,j){o[c]=data[i][j];}); return o; }
  }
  return null;
}

function _getCell(sheetName, sid, col) {
  var r=_rowByKey(sheetName,sid); return r ? String(r[col]||'') : '';
}

function _writeReview(sheetName, sid, remark, status) {
  var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var data=sheet.getDataRange().getValues(), h=data[0], kI=h.indexOf('SubmissionID');
  for(var i=1;i<data.length;i++){
    if(String(data[i][kI])===String(sid)){
      sheet.getRange(i+1,2,1,3).setValues([[remark,status,new Date()]]);
      return;
    }
  }
  sheet.appendRow([sid,remark,status,new Date()]);
}

// Builds a standardized submission item object
function _buildItem(sid,sub,fac,sa,tsRows,hodInfo,hoiInfo,imoInfo) {
  return {
    submissionID: sid,
    facultyName:  String(fac['FacultyName']||'Unknown Faculty'),
    facultyID:    String(fac['Email']||fac['FacultyID']||''),
    department:   String(fac['Department']||''),
    institution:  String(fac['Institution']||''),
    campus:       String(fac['Campus']||''),
    designation:  String(fac['Designation']||''),
    email:        String(fac['Email']||''),
    semester:     String(sub['AcademicYearSemester']||''),
    from:         _fmt(sub['ReportingFrom']),
    to:           _fmt(sub['ReportingTo']),
    submitted:    _fmtDT(sub['SubmittedDateTime']),
    outcome:      String(sa['OutcomeOfWeek']||''),
    target:       String(sa['TargetPlanNextWeek']||''),
    timesheet:    tsRows,
    hodStatus:    String((hodInfo&&hodInfo.hodStatus)||''),
    hodRemark:    String((hodInfo&&hodInfo.hodRemark)||''),
    hoiStatus:    String((hoiInfo&&hoiInfo.hoiStatus)||''),
    hoiRemark:    String((hoiInfo&&hoiInfo.hoiRemark)||''),
    imoStatus:    String((imoInfo&&imoInfo.imoStatus)||'')
  };
}

// ─── Helper: resolve date boundaries for a named period ──────────────────────
// Returns { start: Date, end: Date } for 'week' / 'month' / 'year'.
// Returns { start: null, end: null } for 'all' or any unrecognised value.
// Used by getDashboardStats and getDetailedStats so both functions share
// identical period semantics.
function _periodBounds_(period) {
  var now = new Date();
  if (period === 'week') {
    var dow = now.getDay(), dtm = (dow === 0) ? -6 : (1 - dow);
    var s = new Date(now); s.setDate(now.getDate() + dtm); s.setHours(0,0,0,0);
    var e = new Date(s);   e.setDate(s.getDate() + 6);    e.setHours(23,59,59,999);
    return { start: s, end: e };
  }
  if (period === 'month') {
    var s = new Date(now.getFullYear(), now.getMonth(), 1, 0,0,0,0);
    var e = new Date(now.getFullYear(), now.getMonth()+1, 0, 23,59,59,999);
    return { start: s, end: e };
  }
  if (period === 'year') {
    var s = new Date(now.getFullYear(), 0, 1, 0,0,0,0);
    var e = new Date(now.getFullYear(), 11, 31, 23,59,59,999);
    return { start: s, end: e };
  }
  return { start: null, end: null }; // 'all' — no date filter
}

// ─── Helper: count Mondays in [start, end] inclusive ───────────────────────
// Used to derive expectedWeeks for compliance-rate calculation. A faculty is
// expected to file ONE weekly report per Monday-anchored ISO-style week.
function _countMondays_(start, end) {
  if (!(start instanceof Date) || !(end instanceof Date)) return 0;
  if (start > end) return 0;
  var d = new Date(start.getTime());
  d.setHours(0, 0, 0, 0);
  var dow = d.getDay(); // 0=Sun, 1=Mon, ..., 6=Sat
  // Advance to the first Monday on/after `start`
  var daysToMon = (dow === 1) ? 0 : ((dow === 0) ? 1 : (8 - dow));
  d.setDate(d.getDate() + daysToMon);
  var count = 0;
  while (d.getTime() <= end.getTime()) {
    count++;
    d.setDate(d.getDate() + 7);
  }
  return count;
}

// ─── Helper: resolve the expected number of reporting weeks for a period ───
// For 'week'/'month'/'year' the window is the calendar period clipped to TODAY
// (future weeks haven't happened yet, so they don't count against compliance).
// For 'all' we use Script Property `ACADEMIC_YEAR_START` (ISO date) if set;
// otherwise we fall back to 1-Aug of the current academic year (or last year
// if today is before Aug). The earliest faculty submission is honoured as a
// further floor so compliance isn't punished for pre-deployment weeks.
function _resolveExpectedWeeks_(period, periodStart, periodEnd, earliestSubDate) {
  var now = new Date();
  var s, e;
  if (period === 'week' || period === 'month' || period === 'year') {
    s = periodStart;
    e = (periodEnd && periodEnd < now) ? periodEnd : now;
  } else {
    // 'all' — derive a sensible academic-year start
    var prop = '';
    try { prop = PropertiesService.getScriptProperties().getProperty('ACADEMIC_YEAR_START') || ''; } catch(_) {}
    if (prop) {
      s = new Date(prop);
      if (isNaN(s.getTime())) s = null;
    }
    if (!s) {
      // Default: 1 August of the current academic year. Indian academic year
      // typically starts in August (odd semester). If today is before Aug,
      // we are in the even semester of last year's academic year.
      var ay = (now.getMonth() < 7) ? (now.getFullYear() - 1) : now.getFullYear();
      s = new Date(ay, 7, 1, 0, 0, 0, 0); // 1 Aug
    }
    if (earliestSubDate instanceof Date && earliestSubDate > s) {
      // If first-ever submission was after our derived start, prefer that —
      // we don't penalise institutions for weeks the system wasn't deployed.
      s = new Date(earliestSubDate.getFullYear(), earliestSubDate.getMonth(), earliestSubDate.getDate());
    }
    e = now;
  }
  return _countMondays_(s, e);
}

// ─── Helper: compute the Friday-3-PM deadline for a reporting week ────────
// VMRF policy: weekly reports for a Mon-Sun reporting period are due by 3:00 PM
// on the FRIDAY of that same reporting week. (See sendFridayReminders() —
// the email reminder fires on Friday with "due TODAY by 3:00 PM".)
// Given any date inside the reporting week, returns the Friday 15:00 of that
// week in the script timezone.
function _weekDeadline_(reportingFrom) {
  if (!(reportingFrom instanceof Date) || isNaN(reportingFrom.getTime())) return null;
  var d = new Date(reportingFrom.getTime());
  d.setHours(0, 0, 0, 0);
  var dow = d.getDay(); // Sun=0, Mon=1, ..., Fri=5, Sat=6
  // Days from current dow to Friday (5). If dow is already Fri or Sat, find this week's Fri.
  // We treat Mon-Fri as "find this week's Friday" and Sat-Sun as "find the immediately preceding Friday"
  // because Sat/Sun are *after* the deadline of the same Mon-Sun reporting week.
  var daysToFri;
  if (dow >= 1 && dow <= 5) {       // Mon..Fri
    daysToFri = 5 - dow;            // forward to Fri
  } else if (dow === 6) {           // Sat
    daysToFri = -1;                 // go back to Fri
  } else {                          // Sun (0)
    daysToFri = -2;                 // go back to Fri
  }
  d.setDate(d.getDate() + daysToFri);
  d.setHours(15, 0, 0, 0);
  return d;
}

// ─── Helper: resolve the academic semester window (A7) ───────────────────
// Indian academic default: Odd semester = 1 Aug → 31 Dec, Even semester =
// 1 Jan → 31 May. Override via Script Property SEMESTER_CONFIG, e.g.:
//   {
//     "odd":  { "startMonth": 7,  "endMonth": 11 },   // Aug=7..Dec=11 (0-indexed)
//     "even": { "startMonth": 0,  "endMonth": 4  }    // Jan=0..May=4
//   }
// Returns { start: Date, end: Date, name: 'odd'|'even', label: 'Odd 2025-26', year, offset }.
// `offset=0` returns the semester containing `ref`, `-1` the previous semester,
// `+1` the next, etc. Boundaries are clipped to start-of-day and end-of-day.
function _resolveSemester_(ref, offset) {
  ref = ref || new Date();
  offset = offset || 0;

  var cfg = { odd: { startMonth: 7, endMonth: 11 }, even: { startMonth: 0, endMonth: 4 } };
  try {
    var prop = PropertiesService.getScriptProperties().getProperty('SEMESTER_CONFIG');
    if (prop) {
      var override = JSON.parse(prop);
      if (override && override.odd  && typeof override.odd.startMonth  === 'number') cfg.odd  = override.odd;
      if (override && override.even && typeof override.even.startMonth === 'number') cfg.even = override.even;
    }
  } catch(_) { /* malformed → keep defaults */ }

  // Determine which semester `ref` is in. If month is in [odd.start, odd.end] → odd;
  // otherwise even. Then apply offset by walking semesters back/forward.
  var refMonth = ref.getMonth();
  var refYear  = ref.getFullYear();
  var inOdd = (refMonth >= cfg.odd.startMonth && refMonth <= cfg.odd.endMonth);
  // For the even semester (Jan-May), it's "academic year" of (year - 1)/(year)
  // so the academic-year label is anchored on the year the odd semester started.
  var ayStart = inOdd ? refYear : (refYear - 1);
  var current = { name: inOdd ? 'odd' : 'even', ayStart: ayStart };

  // Walk by offset
  for (var k = 0; k < Math.abs(offset); k++) {
    if (offset < 0) {
      if (current.name === 'odd') { current = { name: 'even', ayStart: current.ayStart - 1 }; }
      else                        { current = { name: 'odd',  ayStart: current.ayStart     }; }
    } else {
      if (current.name === 'odd') { current = { name: 'even', ayStart: current.ayStart     }; }
      else                        { current = { name: 'odd',  ayStart: current.ayStart + 1 }; }
    }
  }

  var conf = (current.name === 'odd') ? cfg.odd : cfg.even;
  var startYear = (current.name === 'odd') ? current.ayStart : (current.ayStart + 1);
  var endYear   = startYear; // both odd (Aug-Dec) and even (Jan-May) are within a single calendar year
  // Last day of `endMonth` — month index + 1, day 0 trick gives last day of that month
  var endDay = new Date(endYear, conf.endMonth + 1, 0).getDate();
  var start = new Date(startYear, conf.startMonth, 1, 0, 0, 0, 0);
  var end   = new Date(endYear,   conf.endMonth,   endDay, 23, 59, 59, 999);
  var label = (current.name === 'odd' ? 'Odd ' : 'Even ') +
              current.ayStart + '-' + (String(current.ayStart + 1).slice(-2));
  return { start: start, end: end, name: current.name, label: label, ayStart: current.ayStart, offset: offset };
}

// ─── Helper: resolve anomaly-detection thresholds (A8) ────────────────────
// Configurable via Script Property ANOMALY_THRESHOLDS, e.g.:
//   {
//     "rejectionPct":     40,    // % rejection above this → 'high_rejection'
//     "hodHrs":           72,    // mean HOD turnaround above this → 'slow_hod_review'
//     "hoiHrs":           168,   // mean HOI turnaround above this → 'slow_hoi_review'
//     "trendPctDrop":    -50,    // trendPct below this → 'sudden_drop'
//     "compliancePct":    50     // complianceRate below this → 'low_compliance'
//   }
// All thresholds are positive numbers (or negative for trendPctDrop). Missing
// keys keep their defaults. Bad JSON → silently falls back to defaults.
function _resolveAnomalyThresholds_() {
  var defaults = {
    rejectionPct:    40,
    hodHrs:          72,
    hoiHrs:         168,
    trendPctDrop:   -50,
    compliancePct:   50
  };
  try {
    var prop = PropertiesService.getScriptProperties().getProperty('ANOMALY_THRESHOLDS');
    if (prop) {
      var override = JSON.parse(prop);
      if (override && typeof override === 'object') {
        Object.keys(defaults).forEach(function(k){
          if (typeof override[k] === 'number') defaults[k] = override[k];
        });
      }
    }
  } catch(_) { /* malformed → keep defaults */ }
  return defaults;
}

// ─── Helper: detect anomalies in a finalised bucket row (A8) ──────────────
// Returns an array of alerts: [{type, severity, detail}]. Severity is one of
// 'high' / 'med' / 'low'. Empty array means everything looks normal.
//
// Detection rules (each can be tuned via ANOMALY_THRESHOLDS):
//   no_submissions     — totalFaculty>0 but submissions===0      → high
//   high_rejection     — rejectionRate > rejectionPct            → high
//   sudden_drop        — trendPct < trendPctDrop                 → med
//   low_compliance     — complianceRate < compliancePct          → med
//   slow_hoi_review    — avgHoiTurnaroundHrs > hoiHrs            → med
//   slow_hod_review    — avgHodTurnaroundHrs > hodHrs            → low
//   declining_compliance — trendDirection==='down' AND complianceRate<70 → low
function _detectAnomalies_(row, thresholds) {
  var alerts = [];

  // High-severity
  if (row.totalFaculty > 0 && row.submissions === 0) {
    alerts.push({ type:'no_submissions', severity:'high',
      detail: row.totalFaculty + ' faculty in scope, 0 submissions in this period' });
  }
  if (row.submissions > 0 && row.rejectionRate > thresholds.rejectionPct) {
    alerts.push({ type:'high_rejection', severity:'high',
      detail: row.rejectionRate + '% of submissions rejected (threshold ' + thresholds.rejectionPct + '%)' });
  }

  // Medium-severity
  if (typeof row.trendPct === 'number' && row.trendPct < thresholds.trendPctDrop) {
    alerts.push({ type:'sudden_drop', severity:'med',
      detail: 'Submissions ' + row.trendPct + '% vs previous period' });
  }
  if (row.totalFaculty > 0 && row.complianceRate < thresholds.compliancePct) {
    alerts.push({ type:'low_compliance', severity:'med',
      detail: row.complianceRate + '% compliance (threshold ' + thresholds.compliancePct + '%)' });
  }
  if (typeof row.avgHoiTurnaroundHrs === 'number' && row.avgHoiTurnaroundHrs > thresholds.hoiHrs) {
    alerts.push({ type:'slow_hoi_review', severity:'med',
      detail: 'Avg HOI turnaround ' + row.avgHoiTurnaroundHrs + ' hrs (threshold ' + thresholds.hoiHrs + ' hrs)' });
  }

  // Low-severity
  if (typeof row.avgHodTurnaroundHrs === 'number' && row.avgHodTurnaroundHrs > thresholds.hodHrs) {
    alerts.push({ type:'slow_hod_review', severity:'low',
      detail: 'Avg HOD turnaround ' + row.avgHodTurnaroundHrs + ' hrs (threshold ' + thresholds.hodHrs + ' hrs)' });
  }
  if (row.trendDirection === 'down' && row.complianceRate < 70 && row.complianceRate >= thresholds.compliancePct) {
    alerts.push({ type:'declining_compliance', severity:'low',
      detail: 'Compliance trending down to ' + row.complianceRate + '%' });
  }

  return alerts;
}

// ─── Helper: resolve weights for the composite 0–100 score (A11) ──────────
// Default weight allocation rationale:
//   compliance        (0.45) — did they file weekly reports at all?
//   onTime            (0.25) — did they meet the Friday-3PM deadline?
//   imoApproval       (0.20) — was the work substantive enough to clear final review?
//   rejectionInverted (0.10) — high rejection rate drags the score down
//
// Override at deploy time via Script Property COMPARISON_SCORE_WEIGHTS, e.g.:
//   { "compliance":0.5, "onTime":0.2, "imoApproval":0.2, "rejectionInverted":0.1 }
// Weights are normalized to sum to 1.0 before use, so any positive numbers work.
function _resolveScoreWeights_() {
  var defaults = {
    compliance: 0.45,
    onTime: 0.25,
    imoApproval: 0.20,
    rejectionInverted: 0.10
  };
  try {
    var prop = PropertiesService.getScriptProperties().getProperty('COMPARISON_SCORE_WEIGHTS');
    if (prop) {
      var override = JSON.parse(prop);
      if (override && typeof override === 'object') {
        Object.keys(defaults).forEach(function(k){
          if (typeof override[k] === 'number' && override[k] >= 0) defaults[k] = override[k];
        });
      }
    }
  } catch(_) { /* malformed JSON → keep defaults */ }
  // Normalize to sum=1.0 so the score stays on a 0–100 scale.
  var sum = 0;
  Object.keys(defaults).forEach(function(k){ sum += defaults[k]; });
  if (sum > 0) Object.keys(defaults).forEach(function(k){ defaults[k] = defaults[k] / sum; });
  return defaults;
}

// ─── Helper: compute composite score from component rates (A11) ───────────
// Each component is 0–100 (or null/undefined if not derivable for this row).
// Null components are dropped and their weight is redistributed proportionally
// across the remaining ones. This lets us still compute a meaningful score for
// e.g. faculty-level rows where on-time classification couldn't be done.
function _computeCompositeScore_(rates, weights) {
  // A component is treated as "present" only when its source rate is a number.
  // null / undefined / NaN → drop the component, redistribute weight elsewhere.
  // Critical: do NOT collapse `null` rejectionRate into "0 rejections" via
  // `100 - (rate || 0)` — that would silently give a perfect 100 score to a
  // bucket where rejection couldn't be measured.
  var _hasNum = function(v) { return typeof v === 'number' && !isNaN(v); };
  var components = {
    compliance:        _hasNum(rates.complianceRate)  ? rates.complianceRate  : null,
    onTime:            _hasNum(rates.onTimeRate)      ? rates.onTimeRate      : null,
    imoApproval:       _hasNum(rates.imoApprovalRate) ? rates.imoApprovalRate : null,
    rejectionInverted: _hasNum(rates.rejectionRate)   ? (100 - rates.rejectionRate) : null
  };
  var activeWeight = 0;
  Object.keys(weights).forEach(function(k){
    if (components[k] !== null && components[k] !== undefined && !isNaN(components[k])) {
      activeWeight += weights[k];
    }
  });
  if (activeWeight === 0) return { score: 0, breakdown: {} };
  var score = 0, breakdown = {};
  Object.keys(weights).forEach(function(k){
    if (components[k] === null || components[k] === undefined || isNaN(components[k])) {
      breakdown[k] = null; // signal to UI: this component skipped
      return;
    }
    var effectiveWeight = weights[k] / activeWeight;
    var contribution = effectiveWeight * components[k];
    breakdown[k] = Math.round(contribution * 10) / 10; // 1-decimal precision
    score += contribution;
  });
  return { score: Math.round(score), breakdown: breakdown };
}

// ═════════════════════════════════════════════════════════════════════════════
// WEEKLY KPI ARCHIVE — captures last week's KPIs every Sunday night so HoD,
// HoI and IMO can browse past weeks even after the underlying submissions
// change (late edits, reviewer remarks, status flips). The archive is
// READ-ONLY snapshots — it never deletes or alters live data.
// ═════════════════════════════════════════════════════════════════════════════
// Sheet: KPI_Weekly_Archive (lazy-created, mirroring the SavedComparisons
// pattern — not in the core SCHEMA so initializeSystem() doesn't seed it).
//
// Columns:
//   WeekStart    — Monday 00:00 of the archived week (Date)
//   WeekEnd      — Sunday 23:59:59 of the archived week (Date)
//   Dimension    — 'campus' | 'institution' | 'department'
//   ScopeKey     — composite scope hint, e.g. 'all', 'campus:Salem',
//                  'inst:VMKVEC|campus:Salem' (used for fast filtered reads)
//   RowKey       — the row's identity from getComparisonReport (department
//                  name, institution name, etc.)
//   RowLabel     — display label
//   Subtitle     — display subtitle
//   Department   — empty for non-faculty dimensions, kept for symmetry
//   KpiJson      — full row JSON (every metric — composite score, on-time,
//                  approval rates, SLA, etc. — exactly what the live report
//                  would show for that week, frozen in time)
//   GeneratedAt  — when the snapshot was written
//
// Faculty dimension is intentionally NOT archived — 1000 faculty × 52 weeks
// × N years grows too fast and the per-faculty view is rarely revisited at a
// past-week granularity. If needed, that data can still be reconstructed on
// demand via getComparisonReport(dim='faculty', period='custom').
//
// API:
//   archiveWeeklyKpis()                    — run this from the trigger
//   getArchivedWeeklyKpis(filter, weekStart)
//                                          — read one week's snapshot
//   listArchivedWeeks(filter)              — distinct archived weeks
//   setupWeeklyArchiveTrigger()            — one-time install of the Sunday
//                                          23:30 trigger (idempotent)
function _kpiArchiveSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('KPI_Weekly_Archive');
  if (!sh) {
    sh = ss.insertSheet('KPI_Weekly_Archive');
    var headers = ['WeekStart','WeekEnd','Dimension','ScopeKey','RowKey','RowLabel','Subtitle','Department','KpiJson','GeneratedAt'];
    sh.getRange(1,1,1,headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#1a2332').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    // Width hints — JSON column is wide, everything else compact
    sh.setColumnWidth(1, 110); sh.setColumnWidth(2, 110);
    sh.setColumnWidth(3, 110); sh.setColumnWidth(4, 200);
    sh.setColumnWidth(5, 200); sh.setColumnWidth(6, 220);
    sh.setColumnWidth(7, 220); sh.setColumnWidth(8, 160);
    sh.setColumnWidth(9, 480); sh.setColumnWidth(10, 150);
  }
  return sh;
}

// Resolve the most-recently-COMPLETED Mon-Sun week. If `now` is on Mon-Sat,
// the just-ended week is last week. If on Sun, the week ends today (Sun
// 23:59:59) — we still snapshot what's been written so far in that week.
// Override behaviour for the trigger to run BEFORE Sunday midnight rather
// than after.
function _lastCompletedWeekBounds_(now) {
  var n = now || new Date();
  var dow = n.getDay(); // 0=Sun .. 6=Sat
  // If today is Sunday, archive THIS week (Mon..today-Sun). Otherwise,
  // archive the previous Mon..Sun window so callers always get a fully-
  // completed week's data.
  var daysBackToLastSun = (dow === 0) ? 0 : (dow); // dow days back gets us to last Sunday
  var sun = new Date(n);
  sun.setDate(n.getDate() - daysBackToLastSun);
  sun.setHours(23, 59, 59, 999);
  var mon = new Date(sun);
  mon.setDate(sun.getDate() - 6);
  mon.setHours(0, 0, 0, 0);
  return { start: mon, end: sun };
}

// Format helper — keeps the "Week of: 12 May – 18 May 2026" string consistent
// across every surface that shows a week boundary.
function _kpiWeekLabel_(weekStart, weekEnd) {
  if (!(weekStart instanceof Date) || !(weekEnd instanceof Date)) return '';
  var tz = Session.getScriptTimeZone();
  var sameYear  = weekStart.getFullYear() === weekEnd.getFullYear();
  var sameMonth = sameYear && weekStart.getMonth() === weekEnd.getMonth();
  if (sameMonth) {
    return Utilities.formatDate(weekStart, tz, 'd') + '–' +
           Utilities.formatDate(weekEnd,   tz, 'd MMM yyyy');
  }
  if (sameYear) {
    return Utilities.formatDate(weekStart, tz, 'd MMM') + ' – ' +
           Utilities.formatDate(weekEnd,   tz, 'd MMM yyyy');
  }
  return Utilities.formatDate(weekStart, tz, 'd MMM yyyy') + ' – ' +
         Utilities.formatDate(weekEnd,   tz, 'd MMM yyyy');
}

// Public utility — returns the CURRENT-week label, used by dashboards.
function getCurrentWeekLabel() {
  var n = new Date();
  var dow = n.getDay();
  var dtm = (dow === 0) ? -6 : (1 - dow);
  var mon = new Date(n); mon.setDate(n.getDate() + dtm); mon.setHours(0,0,0,0);
  var sun = new Date(mon); sun.setDate(mon.getDate() + 6); sun.setHours(23,59,59,999);
  return {
    label: _kpiWeekLabel_(mon, sun),
    from:  Utilities.formatDate(mon, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    to:    Utilities.formatDate(sun, Session.getScriptTimeZone(), 'yyyy-MM-dd')
  };
}

// ═════════════════════════════════════════════════════════════════════════════
// SAVED COMPARISON VIEWS — D6
// ═════════════════════════════════════════════════════════════════════════════
// Lets users persist a comparison configuration (dimension + filter + period +
// optional name) so they can reopen it with one click. Stored in a sheet
// called SavedComparisons (auto-created on first save) with columns:
//   ViewID | OwnerEmail | Name | Dimension | FilterJSON | Period | CreatedAt | UpdatedAt | Shared
// `Shared = TRUE` makes the view visible to other users (e.g. IMO can publish a
// "Quarterly Engineering Review" view that any HOI sees).
//
// API:
//   listSavedComparisons(ownerEmail) → { ok, views: [...] }
//   saveComparisonView(ownerEmail, view) → { ok, view }   // create or update by id
//   deleteComparisonView(ownerEmail, viewId) → { ok }
function _savedComparisonsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('SavedComparisons');
  if (!sh) {
    sh = ss.insertSheet('SavedComparisons');
    var headers = ['ViewID','OwnerEmail','Name','Dimension','FilterJSON','Period','CreatedAt','UpdatedAt','Shared'];
    sh.getRange(1,1,1,headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#1a2332').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

// ── Non-working day helper ─────────────────────────────────────────────────
// Returns true when a given day of the week is non-working for a given week.
//   weekMonday : Date or YYYY-MM-DD string for the Monday of the reporting week.
//   dayIndex   : 0=Mon, 1=Tue, 2=Wed, 3=Thu, 4=Fri, 5=Sat (matches DAYS array order).
// Rules:
//   • Even Saturdays (2nd, 4th of month) are always non-working.
//   • Any day that falls on a VMRF public holiday is non-working.
var VMRF_HOLIDAYS_GS_ = {
  '2026-01-01':true,'2026-01-15':true,'2026-01-16':true,'2026-01-17':true,
  '2026-01-26':true,'2026-03-21':true,'2026-04-14':true,'2026-05-01':true,
  '2026-08-15':true,'2026-09-14':true,'2026-10-02':true,'2026-10-19':true,
  '2026-10-20':true,'2026-12-25':true
};
function _isNonWorkingDay_(weekMonday, dayIndex) {
  try {
    var mon = (weekMonday instanceof Date) ? new Date(weekMonday.getTime()) : new Date(weekMonday);
    if (isNaN(mon.getTime())) return false;
    var day = new Date(mon); day.setDate(mon.getDate() + dayIndex);
    var dow = day.getDay(); // 0=Sun, 6=Sat
    if (dow === 0) return true; // Sunday
    if (dow === 6 && Math.ceil(day.getDate() / 7) % 2 === 0) return true; // Even Saturday
    var mm = day.getMonth() + 1, dd = day.getDate();
    var key = day.getFullYear() + '-' + (mm < 10 ? '0' : '') + mm + '-' + (dd < 10 ? '0' : '') + dd;
    return !!VMRF_HOLIDAYS_GS_[key]; // Public holiday on any weekday
  } catch (e) { return false; }
}

// Helper: build SubmissionID → ReportingFrom (YYYY-MM-DD) map from a sheet's
// raw values. Works for both Weekly_Submission and HOD_Submission because
// both have SubmissionID + ReportingFrom columns by name.
function _buildSidWeekMap(sheetValues) {
  var out = {};
  if (!sheetValues || !sheetValues.length) return out;
  var H = sheetValues[0];
  var sidI = H.indexOf('SubmissionID'); if (sidI < 0) sidI = 0;
  var frI  = H.indexOf('ReportingFrom'); if (frI < 0) return out;
  for (var i = 1; i < sheetValues.length; i++) {
    var sid = String(sheetValues[i][sidI]||'').trim();
    if (!sid) continue;
    var fr = sheetValues[i][frI];
    var wk = '';
    if (fr instanceof Date) {
      wk = Utilities.formatDate(fr, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else if (fr) {
      wk = String(fr).trim().slice(0, 10);
    }
    if (wk) out[sid] = wk;
  }
  return out;
}

function ipmRowsAsObjects_(sh) {
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  const h = v.shift();
  return v.map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i]); return o; });
}

/* Append any missing header columns to an IPM sheet, non-destructively. Used
   to migrate SubmissionItems when followUp + status are added mid-deployment. */
function ipmEnsureColumns_(sh, requiredHeaders) {
  var lastCol = sh.getLastColumn();
  var existing = lastCol > 0
    ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(v){ return String(v).trim(); })
    : [];
  var added = 0;
  requiredHeaders.forEach(function(hdr){
    if (existing.indexOf(hdr) < 0) {
      var col = sh.getLastColumn() + 1;
      sh.getRange(1, col).setValue(hdr).setFontWeight('bold').setBackground('#1a2332').setFontColor('#ffffff');
      existing.push(hdr); added++;
    }
  });
  if (added > 0) sh.setFrozenRows(1);
  return added;
}

/* ---------- Security helpers ---------- */
var IPM_HASH_PREFIX = 'sha256:';
var IPM_MAX_LOGIN_ATTEMPTS = 5;
var IPM_LOGIN_LOCK_SECONDS = 600; // 10 min

function ipmHashPassword_(plain) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(plain), Utilities.Charset.UTF_8);
  return IPM_HASH_PREFIX + bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function ipmIsHashed_(v) { return typeof v === 'string' && v.indexOf(IPM_HASH_PREFIX) === 0; }

function ipmVerifyPassword_(plain, stored) {
  if (ipmIsHashed_(stored)) return ipmHashPassword_(plain) === stored;
  return String(plain) === String(stored); // legacy plaintext
}

function ipmLoginAttemptsKey_(username) { return 'ipm_attempts_' + String(username).toLowerCase(); }

function ipmGetAttempts_(username) {
  const v = CacheService.getScriptCache().get(ipmLoginAttemptsKey_(username));
  return v ? parseInt(v, 10) : 0;
}

function ipmBumpAttempts_(username) {
  const next = ipmGetAttempts_(username) + 1;
  CacheService.getScriptCache().put(ipmLoginAttemptsKey_(username), String(next), IPM_LOGIN_LOCK_SECONDS);
  return next;
}

function ipmClearAttempts_(username) { CacheService.getScriptCache().remove(ipmLoginAttemptsKey_(username)); }

// Normalize a sheet row object by converting any Date values to ISO strings.
// google.script.run silently returns null to success handlers when payloads
// contain raw Date instances in certain cases — stringify to stay safe.
function ipmNormalizeDates_(obj) {
  const out = {};
  const tz = Session.getScriptTimeZone();
  for (const k in obj) {
    const v = obj[k];
    if (v instanceof Date) {
      if (k === 'weekOf') {
        // Date-only field: format in script timezone as plain YYYY-MM-DD
        out[k] = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
      } else {
        // Timestamp and other datetime fields: standard UTC ISO string
        out[k] = v.toISOString();
      }
    } else {
      out[k] = v;
    }
  }
  return out;
}