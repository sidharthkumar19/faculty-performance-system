// =============================================================================
// VMRF-DU Institutional Monitoring System — IPM Database & API Services
// =============================================================================

/* =================================================================== */
/* ====================  I P M   S U B S Y S T E M   ================== */
/* =================================================================== */
/*  Institutional Performance Monitoring — HOI & IMO portal.           */
/*  Merged in from the standalone IPM web app.                          */
/*  All public (client-callable) functions are prefixed 'ipm'.          */
/*  All private helpers are prefixed 'ipm' to avoid colliding with the  */
/*  main app's helpers (db_, sheet_, session_, hashPassword_, etc.).    */
/*  Uses its own spreadsheet via script property 'IPM_DB_SHEET_ID' so   */
/*  the 12 institution HOI accounts, Activities, Submissions &          */
/*  SubmissionItems live in a SEPARATE sheet and do not touch the main  */
/*  Faculty / HOD / Weekly_Submission tables.                           */
/* =================================================================== */

/*****************************************************************
 * VMRF Institutional Performance Monitoring – Apps Script backend
 *
 * BEFORE DEPLOYING:
 *   1. Open script properties and run ipmSetup() once from editor.
 *      It creates the spreadsheet (ID saved to script properties),
 *      seeds Users, Institutions, Activities tabs.
 *   2. Deploy > New deployment > Web app > Execute as: Me,
 *      Access: Anyone with link (or Anyone in domain).
 *
 * Sheet tabs created:
 *   Users         | username | password | role | campus | institution | fullName | createdAt
 *   Institutions  | code     | name     | campus
 *   Activities    | campus   | sectionNo | sectionTitle | itemCode | itemTitle
 *   Submissions   | id | timestamp | username | campus | institution | semester | weekOf | tasks | status
 *   SubmissionItems | submissionId | itemCode | itemTitle | minutes | actionItems | responsibility
 *****************************************************************/

var IPM_APP_TITLE = 'VMRF Institutional Monitoring System';
var IPM_PROP_SHEET_ID = 'IPM_DB_SHEET_ID';

/* ---------- Web app entry ---------- */

/* ---------- Campus header images (CampusHeadersData.gs) ---------- */
var _IPM_CAMPUS_HEADERS_CACHE = null;
function ipmGetCampusHeaders_() {
  if (_IPM_CAMPUS_HEADERS_CACHE) return _IPM_CAMPUS_HEADERS_CACHE;
  _IPM_CAMPUS_HEADERS_CACHE = getCampusHeadersData_();   // defined in CampusHeadersData.gs
  return _IPM_CAMPUS_HEADERS_CACHE;
}

function ipmGetCampusHeaders() {
  return ipmGetCampusHeaders_();
}

/* ---------- DB helpers ---------- */
function ipmDb_() {
  const id = PropertiesService.getScriptProperties().getProperty(IPM_PROP_SHEET_ID);
  if (!id) throw new Error('Database not initialised. Run ipmSetup() from the editor.');
  return SpreadsheetApp.openById(id);
}

function ipmSheet_(name) { return ipmDb_().getSheetByName(name); }

/* ---------- Auth ---------- */
// ─── IPM SSO bridge — for embedded IPM sub-views in the main app ───
// Mints an IPM session token from main-app context so the user doesn't have to
// log in twice. Called either from staffLogin() (when Script-Properties login
// succeeds) or directly from the client when the user has picked a HOI
// institution.
//
// IMPORTANT: The main-app INSTITUTION_HIERARCHY keys campuses/institutions by
// their FULL NAMES (e.g. "Vinayaka Mission's Chennai Campus",
// "Aarupadai Veedu Institute of Technology (AVIT)"), but the IPM Users sheet
// and Activities sheet key them by CODES (VMCC, AVIT). This helper transparently
// accepts either and resolves to the IPM codes before matching.
function _ipmResolveHoiCodes_(campusArg, instArg) {
  var out = { campus: String(campusArg || ''), institution: String(instArg || '') };
  try {
    // 1. Campus: if a main-app full-name was passed, swap it for the code.
    //    Robust to missing INSTITUTION_HIERARCHY (shouldn't happen, but guard).
    if (typeof INSTITUTION_HIERARCHY !== 'undefined' && INSTITUTION_HIERARCHY) {
      if (INSTITUTION_HIERARCHY[out.campus] && INSTITUTION_HIERARCHY[out.campus].code) {
        out.campus = INSTITUTION_HIERARCHY[out.campus].code;
      }
      // 2. Institution: find the matching name anywhere in the hierarchy and
      //    use its code. If the caller already passed a code, we leave it.
      var matchedInstCode = '';
      Object.keys(INSTITUTION_HIERARCHY).forEach(function(campusName) {
        var insts = INSTITUTION_HIERARCHY[campusName].institutions || {};
        Object.keys(insts).forEach(function(instName) {
          if (instName === String(instArg || '') && insts[instName].code) {
            matchedInstCode = insts[instName].code;
          }
        });
      });
      if (matchedInstCode) out.institution = matchedInstCode;
    }
  } catch (_) {}
  return out;
}

function ipmSsoTokenForMainUser(mainRole, hoiCampus, hoiInstitution) {
  try {
    if (!mainRole) return { ok:false, error:'Role is required' };
    mainRole = String(mainRole).toUpperCase();
    if (mainRole !== 'HOI' && mainRole !== 'IMO') {
      return { ok:false, error:'SSO only supported for HOI and IMO roles' };
    }
    var users = ipmRowsAsObjects_(ipmSheet_('Users'));
    var u = null;
    if (mainRole === 'IMO') {
      u = users.find(function(x){ return String(x.role).toUpperCase() === 'IMO'; });
      if (!u) return { ok:false, error:'No IMO account found in IPM Users sheet. Run ipmSetup first.' };
    } else {
      if (!hoiCampus || !hoiInstitution) {
        return { ok:false, error:'Campus and institution are required for HOI SSO' };
      }
      // Resolve main-app names → IPM codes (VMCC, AVIT …) so the lookup works.
      var resolved = _ipmResolveHoiCodes_(hoiCampus, hoiInstitution);
      u = users.find(function(x){
        return String(x.role).toUpperCase() === 'HOI'
            && String(x.campus)      === resolved.campus
            && String(x.institution) === resolved.institution;
      });
      if (!u) {
        return {
          ok:false,
          error:'No IPM HOI account found for ' + resolved.campus + ' / ' + resolved.institution +
                '. Check the IPM Users sheet or run ipmSetup.'
        };
      }
    }
    return { ok:true, token: _ipmMintSession_(u), user: _ipmUserView_(u) };
  } catch (e) {
    return { ok:false, error: e.message };
  }
}

function _ipmMintSession_(u) {
  var token = Utilities.getUuid();
  CacheService.getScriptCache().put('ipm_sess_' + token, JSON.stringify({
    username: u.username, role: u.role, campus: u.campus,
    institution: u.institution, fullName: u.fullName
  }), 21600);
  return token;
}

function _ipmUserView_(u) {
  return {
    username: u.username, role: u.role, campus: u.campus,
    institution: u.institution, fullName: u.fullName
  };
}

function ipmLogin(payload) {
  try {
    const { role: requestedRole, campus, institution, username, password } = payload;
    if (!username || !password) return { ok: false, error: 'Username and password are required' };
    if (!requestedRole || (requestedRole !== 'HOI' && requestedRole !== 'IMO')) {
      return { ok: false, error: 'Invalid role' };
    }
    const attempts = ipmGetAttempts_(username);
    if (attempts >= IPM_MAX_LOGIN_ATTEMPTS) {
      return { ok: false, error: 'Account temporarily locked due to too many failed attempts. Try again in 10 minutes.' };
    }
    const usersSheet = ipmSheet_('Users');
    const users = ipmRowsAsObjects_(usersSheet);
    const idx = users.findIndex(x => String(x.username).toLowerCase() === String(username).toLowerCase());
    if (idx === -1 || !ipmVerifyPassword_(password, users[idx].password)) {
      const a = ipmBumpAttempts_(username);
      return { ok: false, error: 'Invalid credentials. ' + Math.max(0, IPM_MAX_LOGIN_ATTEMPTS - a) + ' attempts remaining.' };
    }
    const u = users[idx];
    // CRITICAL: requested role must match the user's actual role
    if (String(u.role) !== requestedRole) {
      ipmBumpAttempts_(username);
      return { ok: false, error: 'This account is not authorised to sign in as ' + requestedRole + '.' };
    }
    if (u.role === 'HOI') {
      if (u.campus !== campus || u.institution !== institution) {
        ipmBumpAttempts_(username);
        return { ok: false, error: 'Campus / Institution mismatch for this user' };
      }
    }
    // First-ipmLogin flag: any user whose stored password is still plaintext is on a default password
    const mustChangePassword = !ipmIsHashed_(u.password);
    // Auto-migrate plaintext to hash on successful ipmLogin
    if (mustChangePassword) {
      usersSheet.getRange(idx + 2, 2).setValue(ipmHashPassword_(password));
    }
    ipmClearAttempts_(username);
    const token = Utilities.getUuid();
    CacheService.getScriptCache().put('ipm_sess_' + token, JSON.stringify({
      username: u.username, role: u.role, campus: u.campus,
      institution: u.institution, fullName: u.fullName
    }), 21600);
    return { ok: true, token, mustChangePassword: mustChangePassword, user: { username: u.username, role: u.role, campus: u.campus, institution: u.institution, fullName: u.fullName } };
  } catch (err) { return { ok: false, error: err.message }; }
}

function ipmChangePassword(token, currentPw, newPw) {
  const s = ipmSession_(token);
  if (!newPw || newPw.length < 8) return { ok: false, error: 'New password must be at least 8 characters' };
  // Rate limit using the existing ipmLogin-attempts bucket so brute-force from a hijacked session is throttled too
  const attempts = ipmGetAttempts_(s.username);
  if (attempts >= IPM_MAX_LOGIN_ATTEMPTS) {
    return { ok: false, error: 'Too many failed attempts. Try again in 10 minutes.' };
  }
  const usersSheet = ipmSheet_('Users');
  const users = ipmRowsAsObjects_(usersSheet);
  const idx = users.findIndex(x => String(x.username).toLowerCase() === String(s.username).toLowerCase());
  if (idx === -1) return { ok: false, error: 'User not found' };
  if (!ipmVerifyPassword_(currentPw, users[idx].password)) {
    ipmBumpAttempts_(s.username);
    return { ok: false, error: 'Current password is incorrect' };
  }
  if (currentPw === newPw) return { ok: false, error: 'New password must be different from current password' };
  usersSheet.getRange(idx + 2, 2).setValue(ipmHashPassword_(newPw));
  ipmClearAttempts_(s.username);
  return { ok: true };
}

function ipmPingSession(token) {
  // Refresh session TTL on activity
  const raw = CacheService.getScriptCache().get('ipm_sess_' + token);
  if (!raw) return { ok: false };
  CacheService.getScriptCache().put('ipm_sess_' + token, raw, 21600);
  return { ok: true };
}

function ipmSession_(token) {
  const raw = CacheService.getScriptCache().get('ipm_sess_' + token);
  if (!raw) throw new Error('Session expired. Please log in again.');
  return JSON.parse(raw);
}

function ipmLogout(token) { CacheService.getScriptCache().remove('ipm_sess_' + token); return { ok: true }; }

/* ---------- Bootstrap data for ipmLogin page ---------- */
function ipmGetCampusTree() {
  const inst = ipmRowsAsObjects_(ipmSheet_('Institutions'));
  const tree = {};
  inst.forEach(i => {
    if (!tree[i.campus]) tree[i.campus] = [];
    tree[i.campus].push({ code: i.code, name: i.name });
  });
  return tree;
}

function ipmGetLogos(){
  try{return getLogosData_();}catch(e){return {};}
}

/* ---------- Activities for the logged-in HOI ---------- */
function ipmGetActivities(token) {
  const s = ipmSession_(token);
  if (s.role !== 'HOI') throw new Error('Forbidden');
  const rows = ipmRowsAsObjects_(ipmSheet_('Activities')).filter(r => r.campus === s.campus);
  const sections = [];
  const map = {};
  rows.forEach(r => {
    if (!map[r.sectionNo]) {
      map[r.sectionNo] = { sectionNo: r.sectionNo, sectionTitle: r.sectionTitle, items: [] };
      sections.push(map[r.sectionNo]);
    }
    if (r.itemTitle) map[r.sectionNo].items.push({ title: r.itemTitle });
  });
  return sections;
}

/* ---------- Submissions ---------- */
function ipmSubmitReport(token, payload) {
  const s = ipmSession_(token);
  if (s.role !== 'HOI') throw new Error('Only HOI can submit');

  // Enforce one submission per week per semester (excluding Drafts)
  var normWk = function(v){ return String(v instanceof Date?v.toISOString():(v||'')).trim().slice(0,10); };
  var weekKey = normWk(payload.weekOf);
  var sem = String(payload.semester||'').trim();
  var me  = String(s.username).trim().toLowerCase();
  var existing = ipmRowsAsObjects_(ipmSheet_('Submissions')).filter(function(r){
    return String(r.username).trim().toLowerCase() === me &&
           normWk(r.weekOf) === weekKey &&
           String(r.semester).trim() === sem &&
           String(r.status).trim() !== 'Draft';
  });
  if (existing.length > 0)
    return { ok: false, error: 'You have already submitted a report for this week (' + weekKey + ', ' + sem + ' semester). Only one submission per week is allowed.' };

  const id = Utilities.getUuid();
  const now = new Date();
  ipmSheet_('Submissions').appendRow([
    id, now, s.username, s.campus, s.institution,
    payload.semester, payload.weekOf, payload.tasks || '', 'Submitted'
  ]);
  const itemsSheet = ipmSheet_('SubmissionItems');
  ipmEnsureColumns_(itemsSheet, ['submissionId','itemTitle','minutes','actionItems','responsibility','attachments','followUp','status']);

  // Create a Drive folder for attachments if any item has files
  let attachFolder = null;
  const hasFiles = (payload.items || []).some(it => (it.files || []).length > 0);
  if (hasFiles) {
    try {
      const ROOT_NAME = 'IPM Submission Attachments';
      const roots = DriveApp.getFoldersByName(ROOT_NAME);
      const root = roots.hasNext() ? roots.next() : DriveApp.createFolder(ROOT_NAME);
      const label = s.institution + ' ' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      attachFolder = root.createFolder(id + ' (' + label + ')');
    } catch(e) {
      console.error('[IPM] Drive folder creation failed: ' + e.message);
    }
  }

  (payload.items || []).forEach(it => {
    let attachments = '';
    if (attachFolder && (it.files || []).length > 0) {
      const parts = [];
      it.files.forEach(f => {
        try {
          const blob = Utilities.newBlob(
            Utilities.base64Decode(f.data),
            f.type || 'application/octet-stream',
            f.name
          );
          const driveFile = attachFolder.createFile(blob);
          driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          parts.push(driveFile.getUrl() + '|' + f.name);
        } catch(e) {
          console.error('[IPM] File upload failed (' + f.name + '): ' + e.message);
        }
      });
      attachments = parts.join(';;');
    }
    itemsSheet.appendRow([id, it.title, it.minutes || '', it.actionItems || '', it.responsibility || '', attachments, it.followUp || '', it.status || '']);
  });
  return { ok: true, id };
}

/* ---------- ipmDeleteSubmission — HOI deletes own Submitted or Draft (not Reviewed) ---------- */
function ipmDeleteSubmission(token, id) {
  var s = ipmSession_(token);
  if (s.role !== 'HOI') throw new Error('Only HOI can delete submissions');
  var subSh   = ipmSheet_('Submissions');
  var subData = subSh.getDataRange().getValues();
  var headers = subData[0];
  var idCol   = headers.indexOf('id');
  var usrCol  = headers.indexOf('username');
  var stCol   = headers.indexOf('status');
  var rowIdx  = subData.findIndex(function(r, i) { return i > 0 && String(r[idCol]) === id; });
  if (rowIdx === -1) return { ok: false, error: 'Submission not found' };
  if (String(subData[rowIdx][usrCol]).trim().toLowerCase() !== String(s.username).trim().toLowerCase())
    return { ok: false, error: 'You can only delete your own submissions' };
  if (String(subData[rowIdx][stCol]).trim() === 'Reviewed')
    return { ok: false, error: 'Reviewed submissions cannot be deleted' };
  // Delete SubmissionItems first
  var itemSh   = ipmSheet_('SubmissionItems');
  var itemData = itemSh.getDataRange().getValues();
  var toDelete = [];
  itemData.forEach(function(r, i) { if (i > 0 && String(r[0]) === id) toDelete.push(i + 1); });
  toDelete.reverse().forEach(function(rn) { itemSh.deleteRow(rn); });
  // Delete the submission row
  subSh.deleteRow(rowIdx + 1);
  return { ok: true };
}

/* ---------- ipmGetCarryForward ----------
   Given a (semester, weekOf) the HOI is about to submit for, build the
   carry-forward list for each activity item from their PREVIOUS submission.
   Carry-forward sources, per item:
     1. Last week's NEW action items — always carried (they are new
        commitments with no status; they need follow-up next week).
     2. Last week's FOLLOW-UP action items where status === "In Progress"
        — these are ongoing tasks that aren't done yet.
   Dropped:
     • Last week's follow-up items where status === "Completed" (done; do
       not surface them again).
   Returns { hasPrevious: bool, byItem: { itemTitle: [text, ...] } }.
   If the HOI has no prior submission at all, hasPrevious=false. */
function ipmGetCarryForward(token, semester, weekOf) {
  var s = ipmSession_(token);
  if (s.role !== 'HOI') throw new Error('Forbidden');
  var normWk = function(v){ return String(v instanceof Date?v.toISOString():(v||'')).trim().slice(0,10); };
  var me = String(s.username).trim().toLowerCase();
  var sem = String(semester||'').trim();
  var wkThis = normWk(weekOf);

  // Find all non-draft submissions by this HOI in this semester,
  // sorted newest-first by timestamp. The most recent one is the "previous week".
  var subs = ipmRowsAsObjects_(ipmSheet_('Submissions')).filter(function(r){
    if (String(r.username).trim().toLowerCase() !== me) return false;
    if (String(r.status).trim() === 'Draft') return false;
    return true; // include all semesters — carry from latest regardless
  }).sort(function(a, b){
    return new Date(b.timestamp) - new Date(a.timestamp);
  });

  if (!subs.length) return { ok:true, hasPrevious:false, byItem:{} };

  var prevId = String(subs[0].id);
  var itemSh = ipmSheet_('SubmissionItems');
  ipmEnsureColumns_(itemSh, ['submissionId','itemTitle','minutes','actionItems','responsibility','attachments','followUp','status']);
  var rows = itemSh.getDataRange().getValues();
  var H = rows[0] || [];
  var cSub = H.indexOf('submissionId'), cTit = H.indexOf('itemTitle');
  var cAct = H.indexOf('actionItems'), cSt = H.indexOf('status');
  var cFu  = H.indexOf('followUp');
  var JOIN = '|||';

  // Helper — split a "|||"-joined cell value into a trimmed-string array.
  // A blank cell yields []; a single value yields [value].
  var splitJoin = function(raw){
    var v = String(raw == null ? '' : raw);
    if (!v) return [];
    return v.indexOf(JOIN) >= 0 ? v.split(JOIN) : [v];
  };

  var byItem = {};
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (String(r[cSub]) !== prevId) continue;
    var title = String(r[cTit]||'').trim();
    if (!title) continue;
    if (!byItem[title]) byItem[title] = [];
    var seen = {}; // case-insensitive de-dup, scoped per item title.

    // 1. Last week's NEW action items — always carry forward.
    var actArr = cAct >= 0 ? splitJoin(r[cAct]) : [];
    for (var k = 0; k < actArr.length; k++) {
      var actVal = String(actArr[k]||'').trim();
      if (!actVal) continue;
      var actKey = actVal.toLowerCase();
      if (seen[actKey]) continue;
      seen[actKey] = true;
      byItem[title].push(actVal);
    }

    // 2. Last week's FOLLOW-UP items where status === "In Progress".
    //    "Completed" items (and items with empty/other status) are dropped.
    //    followUp[j] is paired with status[j] (the form keeps them aligned).
    var fuArr = cFu >= 0 ? splitJoin(r[cFu]) : [];
    var stArr = cSt >= 0 ? splitJoin(r[cSt]) : [];
    for (var j = 0; j < fuArr.length; j++) {
      var fuVal = String(fuArr[j]||'').trim();
      if (!fuVal) continue;
      var stVal = String(stArr[j]||'').trim();
      if (stVal !== 'In Progress') continue; // only carry in-progress items
      var fuKey = fuVal.toLowerCase();
      if (seen[fuKey]) continue;
      seen[fuKey] = true;
      byItem[title].push(fuVal);
    }

    if (!byItem[title].length) delete byItem[title];
  }
  return { ok:true, hasPrevious:true, previousWeekOf: normWk(subs[0].weekOf), byItem: byItem };
}

function ipmSaveDraft(token, payload) {
  var s = ipmSession_(token);
  if (s.role !== 'HOI') throw new Error('Only HOI can save drafts');
  if (!payload.weekOf) return { ok: false, error: 'weekOf is required' };
  var normWk = function(v){ return String(v instanceof Date?v.toISOString():(v||'')).trim().slice(0,10); };
  var weekKey = normWk(payload.weekOf);
  var sem     = String(payload.semester||'').trim();
  var subSh   = ipmSheet_('Submissions');
  var subData = subSh.getDataRange().getValues();
  var headers = subData[0];
  var idCol=headers.indexOf('id'), usrCol=headers.indexOf('username'), stCol=headers.indexOf('status');
  var wkCol=headers.indexOf('weekOf'), smCol=headers.indexOf('semester'), tsCol=headers.indexOf('timestamp');
  // Find existing draft for this user+week+semester
  var existIdx=-1;
  for(var i=1;i<subData.length;i++){
    if(String(subData[i][usrCol]).trim().toLowerCase()===String(s.username).trim().toLowerCase()&&
       normWk(subData[i][wkCol])===weekKey&&String(subData[i][smCol]).trim()===sem&&
       String(subData[i][stCol]).trim()==='Draft'){ existIdx=i; break; }
  }
  var itemSh=ipmSheet_('SubmissionItems');
  ipmEnsureColumns_(itemSh, ['submissionId','itemTitle','minutes','actionItems','responsibility','attachments','followUp','status']);
  if(existIdx!==-1){
    var draftId=String(subData[existIdx][idCol]);
    if(tsCol!==-1) subSh.getRange(existIdx+1,tsCol+1).setValue(new Date());
    var itemData=itemSh.getDataRange().getValues();
    var toDelete=[];
    itemData.forEach(function(r,i){if(i>0&&String(r[0])===draftId)toDelete.push(i+1);});
    toDelete.reverse().forEach(function(rn){itemSh.deleteRow(rn);});
    (payload.items||[]).forEach(function(it){itemSh.appendRow([draftId,it.title,it.minutes||'',it.actionItems||'',it.responsibility||'','',it.followUp||'',it.status||'']);});
    return {ok:true,id:draftId,action:'updated'};
  }
  // Check not already submitted this week
  for(var j=1;j<subData.length;j++){
    if(String(subData[j][usrCol]).trim().toLowerCase()===String(s.username).trim().toLowerCase()&&
       normWk(subData[j][wkCol])===weekKey&&String(subData[j][smCol]).trim()===sem&&
       String(subData[j][stCol]).trim()!=='Draft') return {ok:false,error:'A submission already exists for this week'};
  }
  var id=Utilities.getUuid();
  subSh.appendRow([id,new Date(),s.username,s.campus,s.institution,sem,payload.weekOf,'','Draft']);
  (payload.items||[]).forEach(function(it){itemSh.appendRow([id,it.title,it.minutes||'',it.actionItems||'',it.responsibility||'','',it.followUp||'',it.status||'']);});
  return {ok:true,id:id,action:'created'};
}

/* ---------- ipmGetDraft ---------- */
function ipmGetDraft(token, weekOf, semester) {
  var s=ipmSession_(token);
  if(s.role!=='HOI') throw new Error('Forbidden');
  var normWk=function(v){return String(v instanceof Date?v.toISOString():(v||'')).trim().slice(0,10);};
  var me=String(s.username).trim().toLowerCase();
  var subs=ipmRowsAsObjects_(ipmSheet_('Submissions')).filter(function(r){
    return String(r.username).trim().toLowerCase()===me&&
           normWk(r.weekOf)===normWk(weekOf)&&
           String(r.semester).trim()===String(semester||'').trim()&&
           String(r.status).trim()==='Draft';
  });
  if(!subs.length) return {ok:true,draft:null};
  var draft=ipmNormalizeDates_(subs[0]);
  var itemSh = ipmSheet_('SubmissionItems');
  ipmEnsureColumns_(itemSh, ['submissionId','itemTitle','minutes','actionItems','responsibility','attachments','followUp','status']);
  var rows=itemSh.getDataRange().getValues();
  var H = rows[0] || [];
  var cIdx = {
    submissionId:   H.indexOf('submissionId'),
    itemTitle:      H.indexOf('itemTitle'),
    minutes:        H.indexOf('minutes'),
    actionItems:    H.indexOf('actionItems'),
    responsibility: H.indexOf('responsibility'),
    followUp:       H.indexOf('followUp'),
    status:         H.indexOf('status')
  };
  draft.items=rows.slice(1).filter(function(r){return r[cIdx.submissionId]===draft.id;})
    .map(function(r){return {
      itemTitle:      String(r[cIdx.itemTitle]||''),
      minutes:        String(r[cIdx.minutes]||''),
      actionItems:    String(r[cIdx.actionItems]||''),
      responsibility: String(r[cIdx.responsibility]||''),
      followUp:       cIdx.followUp>=0 ? String(r[cIdx.followUp]||'') : '',
      status:         cIdx.status>=0   ? String(r[cIdx.status]||'')   : ''
    };});
  return {ok:true,draft:draft};
}

/* ---------- ipmDeleteDraft ---------- */
function ipmDeleteDraft(token, id) {
  var s=ipmSession_(token);
  if(s.role!=='HOI') throw new Error('Forbidden');
  var subSh=ipmSheet_('Submissions'), subData=subSh.getDataRange().getValues(), headers=subData[0];
  var idCol=headers.indexOf('id'), usrCol=headers.indexOf('username'), stCol=headers.indexOf('status');
  var rowIdx=subData.findIndex(function(r,i){return i>0&&String(r[idCol])===id;});
  if(rowIdx===-1) return {ok:false,error:'Draft not found'};
  if(String(subData[rowIdx][usrCol]).trim().toLowerCase()!==String(s.username).trim().toLowerCase())
    return {ok:false,error:'Not your draft'};
  if(String(subData[rowIdx][stCol]).trim()!=='Draft') return {ok:false,error:'Not a draft'};
  var itemSh=ipmSheet_('SubmissionItems'), itemData=itemSh.getDataRange().getValues(), toDelete=[];
  itemData.forEach(function(r,i){if(i>0&&String(r[0])===id)toDelete.push(i+1);});
  toDelete.reverse().forEach(function(rn){itemSh.deleteRow(rn);});
  subSh.deleteRow(rowIdx+1);
  return {ok:true};
}

function ipmMySubmissions(token) {
  const s = ipmSession_(token);
  if (s.role !== 'HOI') throw new Error('Forbidden');
  const me = String(s.username).trim().toLowerCase();
  return ipmRowsAsObjects_(ipmSheet_('Submissions'))
    .filter(r => String(r.username).trim().toLowerCase() === me)
    .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp))
    .slice(0, 200)
    .map(ipmNormalizeDates_);
}

function ipmSubmissionDetail(token, id) {
  var s = ipmSession_(token);
  var head = ipmRowsAsObjects_(ipmSheet_('Submissions')).find(function(r){ return r.id === id; });
  if (!head) throw new Error('Not found');
  if (s.role === 'HOI' && head.username !== s.username) throw new Error('Forbidden');

  // Flat items with attachments
  var itemSh2 = ipmSheet_('SubmissionItems');
  ipmEnsureColumns_(itemSh2, ['submissionId','itemTitle','minutes','actionItems','responsibility','attachments','followUp','status']);
  var rows = itemSh2.getDataRange().getValues();
  var IH = rows[0] || [];
  var iC = {
    submissionId:   IH.indexOf('submissionId'),
    itemTitle:      IH.indexOf('itemTitle'),
    minutes:        IH.indexOf('minutes'),
    actionItems:    IH.indexOf('actionItems'),
    responsibility: IH.indexOf('responsibility'),
    attachments:    IH.indexOf('attachments'),
    followUp:       IH.indexOf('followUp'),
    status:         IH.indexOf('status')
  };
  var items = rows.slice(1)
    .filter(function(r){ return r[iC.submissionId] === id; })
    .map(function(r){
      return {
        submissionId:   r[iC.submissionId],
        itemTitle:      String(r[iC.itemTitle]||''),
        minutes:        String(r[iC.minutes]||''),
        actionItems:    String(r[iC.actionItems]||''),
        responsibility: String(r[iC.responsibility]||''),
        followUp:       iC.followUp>=0 ? String(r[iC.followUp]||'') : '',
        status:         iC.status>=0   ? String(r[iC.status]||'')   : '',
        attachmentList: iC.attachments>=0 && r[iC.attachments] ? String(r[iC.attachments]).split(';;').filter(Boolean).map(function(a){
          var sep = a.lastIndexOf('|');
          return { url: a.slice(0, sep), name: a.slice(sep + 1) };
        }) : []
      };
    });

  // Group items into sections using Activities sheet (same campus as submission)
  var campus = head.campus || s.campus || '';
  var actRows = ipmRowsAsObjects_(ipmSheet_('Activities')).filter(function(r){ return r.campus === campus; });
  var sectionOrder = [], sectionMeta = {}, itemToSec = {};
  actRows.forEach(function(r){
    var sno = String(r.sectionNo);
    if (!sectionMeta[sno]){ sectionMeta[sno] = r.sectionTitle; sectionOrder.push(sno); }
    if (r.itemTitle) itemToSec[String(r.itemTitle).trim()] = sno;
  });

  // Build sections
  var grouped = {};
  items.forEach(function(it){
    var sno = itemToSec[String(it.itemTitle).trim()] || '__other__';
    if (!grouped[sno]) grouped[sno] = [];
    grouped[sno].push(it);
  });
  var sections = sectionOrder
    .filter(function(sno){ return grouped[sno] && grouped[sno].length; })
    .map(function(sno){
      return { sectionNo: sno, sectionTitle: sectionMeta[sno] || sno, items: grouped[sno] };
    });
  if (grouped['__other__'] && grouped['__other__'].length)
    sections.push({ sectionNo: '', sectionTitle: 'Other', items: grouped['__other__'] });

  return { head: ipmNormalizeDates_(head), items: items, sections: sections };
}

/* ---------- IMO dashboard ---------- */
function ipmImoTree(token) {
  const s = ipmSession_(token);
  if (s.role !== 'IMO') throw new Error('Forbidden');
  const inst = ipmRowsAsObjects_(ipmSheet_('Institutions'));
  const subs = ipmRowsAsObjects_(ipmSheet_('Submissions'));
  const campuses = {};
  inst.forEach(i => {
    if (!campuses[i.campus]) campuses[i.campus] = { name: i.campus, total: 0, institutions: [] };
    const instSubs = subs.filter(r => r.campus === i.campus && r.institution === i.code && String(r.status).trim() !== 'Draft')
                         .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    const count = instSubs.length;
    const lastSubmission = count > 0 && instSubs[0].timestamp ? new Date(instSubs[0].timestamp).toISOString() : null;
    campuses[i.campus].total += count;
    campuses[i.campus].institutions.push({ code: i.code, name: i.name, count, lastSubmission });
  });
  return Object.values(campuses);
}

function ipmImoInstitutionSubmissions(token, campus, institutionCode) {
  const s = ipmSession_(token);
  if (s.role !== 'IMO') throw new Error('Forbidden');
  return ipmRowsAsObjects_(ipmSheet_('Submissions'))
    .filter(r => r.campus === campus && r.institution === institutionCode && String(r.status).trim() !== 'Draft')
    .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp))
    .map(ipmNormalizeDates_);
}

/* IMO marks a submission reviewed (toggle) */
function ipmSetReviewed(token, id, reviewed) {
  const s = ipmSession_(token);
  if (s.role !== 'IMO') throw new Error('Forbidden');
  const sh = ipmSheet_('Submissions');
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');
  const statusCol = headers.indexOf('status');
  if (idCol < 0 || statusCol < 0) throw new Error('Schema mismatch');
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === id) {
      sh.getRange(i + 1, statusCol + 1).setValue(reviewed ? 'Reviewed' : 'Submitted');
      return { ok: true, status: reviewed ? 'Reviewed' : 'Submitted' };
    }
  }
  throw new Error('Submission not found');
}

/* IMO exports an institution's submissions as a CSV string */
function ipmExportInstitutionCsv(token, campus, institutionCode) {
  const s = ipmSession_(token);
  if (s.role !== 'IMO') throw new Error('Forbidden');
  const subs = ipmRowsAsObjects_(ipmSheet_('Submissions'))
    .filter(r => r.campus === campus && r.institution === institutionCode)
    .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
  const items = ipmRowsAsObjects_(ipmSheet_('SubmissionItems'));
  const itemsBySub = {};
  items.forEach(it => {
    if (!itemsBySub[it.submissionId]) itemsBySub[it.submissionId] = [];
    itemsBySub[it.submissionId].push(it);
  });
  const csvEsc = v => {
    if (v == null) return '';
    const str = String(v).replace(/"/g, '""');
    return /[",\n\r]/.test(str) ? '"' + str + '"' : str;
  };
  const rows = [['Submission Date', 'Week Of', 'Submitted By', 'Semester', 'Status', 'Activity', 'Minutes', 'Action Items', 'Responsibility', 'Tasks/Notes']];
  subs.forEach(sub => {
    const subItems = itemsBySub[sub.id] || [];
    const ts = sub.timestamp ? new Date(sub.timestamp).toLocaleString() : '';
    const wk = sub.weekOf ? new Date(sub.weekOf).toLocaleDateString() : '';
    if (subItems.length === 0) {
      rows.push([ts, wk, sub.username, sub.semester, sub.status, '', '', '', '', sub.tasks || '']);
    } else {
      subItems.forEach((it, idx) => {
        rows.push([
          idx === 0 ? ts : '',
          idx === 0 ? wk : '',
          idx === 0 ? sub.username : '',
          idx === 0 ? sub.semester : '',
          idx === 0 ? sub.status : '',
          it.itemTitle || '',
          it.minutes || '',
          it.actionItems || '',
          it.responsibility || '',
          idx === 0 ? (sub.tasks || '') : ''
        ]);
      });
    }
  });
  const csv = rows.map(r => r.map(csvEsc).join(',')).join('\r\n');
  return { ok: true, csv: csv, filename: campus + '_' + institutionCode + '_submissions.csv' };
}

function ipmImoKpis(token) {
  const s = ipmSession_(token);
  if (s.role !== 'IMO') throw new Error('Forbidden');
  const subs = ipmRowsAsObjects_(ipmSheet_('Submissions')).filter(function(r){return String(r.status).trim()!=='Draft';});
  const inst = ipmRowsAsObjects_(ipmSheet_('Institutions'));
  const now = new Date();
  const byCampus = {};
  subs.forEach(r => { byCampus[r.campus] = (byCampus[r.campus] || 0) + 1; });
  const thisWeek = subs.filter(r => {
    const d = new Date(r.timestamp);
    const diff = (now - d) / 86400000;
    return diff <= 7;
  }).length;
  return {
    totalInstitutions: inst.length,
    totalSubmissions: subs.length,
    thisWeek,
    byCampus,
    reportingInstitutions: new Set(subs.map(r => r.institution)).size
  };
}

/* ======================================================
 * SETUP – run once from editor
 * ====================================================== */
function ipmSetup() {
  let id = PropertiesService.getScriptProperties().getProperty(IPM_PROP_SHEET_ID);
  let ss;
  if (id) {
    ss = SpreadsheetApp.openById(id);
    // SAFETY: if the database already has submission data, refuse to wipe.
    // ipmSetup() unconditionally clear()s every sheet (Submissions included)
    // — running it on a populated production DB would destroy every weekly
    // update. Force the operator to use ipmInitialize() instead, which is
    // additive-only.
    try {
      var existingSubsSheet = ss.getSheetByName('Submissions');
      if (existingSubsSheet && existingSubsSheet.getLastRow() > 1) {
        var msg = 'IPM database already contains submission data ('
          + (existingSubsSheet.getLastRow() - 1) + ' rows). ' +
          'ipmSetup() would WIPE every sheet — refusing to proceed. ' +
          'Run ipmInitialize() instead (additive-only) to sync schema changes.';
        try { SpreadsheetApp.getUi().alert(msg); } catch(_){}
        Logger.log(msg);
        return ss.getUrl();
      }
    } catch(_){ /* fresh install — fall through */ }
  }
  else {
    ss = SpreadsheetApp.create('VMRF IPM Database');
    PropertiesService.getScriptProperties().setProperty(IPM_PROP_SHEET_ID, ss.getId());
  }

  const schemas = {
    Users: ['username', 'password', 'role', 'campus', 'institution', 'fullName', 'createdAt'],
    Institutions: ['code', 'name', 'campus'],
    Activities: ['campus', 'sectionNo', 'sectionTitle', 'itemTitle'],
    Submissions: ['id', 'timestamp', 'username', 'campus', 'institution', 'semester', 'weekOf', 'tasks', 'status'],
    SubmissionItems: ['submissionId', 'itemTitle', 'minutes', 'actionItems', 'responsibility', 'attachments', 'followUp', 'status']
  };
  Object.keys(schemas).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    sh.clear();
    sh.getRange(1, 1, 1, schemas[name].length).setValues([schemas[name]]).setFontWeight('bold').setBackground('#1a2332').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  });
  const def = ss.getSheetByName('Sheet1'); if (def) ss.deleteSheet(def);

  // Institutions
  const institutions = [
    ['AVIT',     'Aarupadai Veedu Institute of Technology',                 'VMCC'],
    ['AVIT-DoM', 'AVIT - Department of Management',                         'VMCC'],
    ['SAS',      'School of Arts and Science',                              'VMCC'],
    ['SAHS',     'School of Allied Health Sciences - Chennai Campus',       'VMCC'],
    ['VMLS',     'Vinayaka Mission\'s Law School',                          'VMCC'],
    ['VSEP',     'Vinayaka Mission\'s School of Economics and Public Policy','VMCC'],
    ['VSHS',     'Vinayaka Mission\'s School of Health Systems',            'VMCC'],
    ['AVMC',     'Aarupadai Veedu Medical College & Hospital',              'VMPC'],
    ['VMCON',    'Vinayaka Mission\'s College of Nursing',                  'VMPC'],
    ['SRBS',     'School of Rehabilitation and Behavioral Sciences',        'VMPC'],
    ['SPT',      'School of Physiotherapy',                                 'VMPC'],
    ['SAHS-MC',  'School of Allied Health Sciences - Puducherry Campus',   'VMPC']
  ];
  ipmSheet_('Institutions').getRange(2, 1, institutions.length, 3).setValues(institutions);

  // Users – credentials (CHANGE AFTER FIRST LOGIN!)
  const users = [
    ['hoi_avit',     'psTiOfGQvIYn', 'HOI', 'VMCC', 'AVIT',     'HOI - AVIT'],
    ['hoi_avit_dom', 'jQ%83F7q5ojh', 'HOI', 'VMCC', 'AVIT-DoM', 'HOI - AVIT DoM'],
    ['hoi_sas',      'n%nupphXdm2@', 'HOI', 'VMCC', 'SAS',      'HOI - SAS'],
    ['hoi_sahs',     'sSn2YcbwI#ee', 'HOI', 'VMCC', 'SAHS',     'HOI - SAHS-CC (VMCC)'],
    ['hoi_vmls',     '6#hZvrPe$FHk', 'HOI', 'VMCC', 'VMLS',     'HOI - VMLS'],
    ['hoi_vsep',     'gFRgJ14xulAG', 'HOI', 'VMCC', 'VSEP',     'HOI - VSEP'],
    ['hoi_vshs',     'ZvBg@J9LuXnn', 'HOI', 'VMCC', 'VSHS',     'HOI - VSHS'],
    ['hoi_avmc',     'CqsWevls%uhR', 'HOI', 'VMPC', 'AVMC',     'HOI - AVMC'],
    ['hoi_vmcon',    'bEBY$7n84AE9', 'HOI', 'VMPC', 'VMCON',    'HOI - VMCON'],
    ['hoi_srbs',     'Wo42PHyBn1m@', 'HOI', 'VMPC', 'SRBS',     'HOI - SRBS (VMPC)'],
    ['hoi_spt',      'rk3RkgSBbC61', 'HOI', 'VMPC', 'SPT',      'HOI - SPT'],
    ['hoi_sahs_mc',  'NxaI2uFhoRUP', 'HOI', 'VMPC', 'SAHS-MC',  'HOI - SAHS-PC (VMPC)'],
    ['imo',          '0y%6ebW64G4E', 'IMO', '',     '',         'Institutional Monitoring Officer']
  ];
  const now = new Date();
  ipmSheet_('Users').getRange(2, 1, users.length, 7).setValues(users.map(u => [...u, now]));

  // Activities – VMCC & VMPC
  const vmccSections = [
    ['1', 'Academics', [
      ['1a', 'Academics for UG (including WP) (Attendance, PT meeting, IA, Slow Learners, Exams)'],
      ['1b', 'Academics for PG (Attendance, PT meeting, IA, Slow Learners, Exams)'],
      ['1c', 'Faculty Attendance'],
      ['1d', 'Student Affairs - Hostel, Mess, Disciplinary Actions'],
      ['1e', 'Fees Collection']]],
    ['2', 'Research & Innovation', [
      ['2a', 'Research - Publications & Patent'],
      ['2b', 'Research - Seed Money'],
      ['2c', 'Research - External Proposals'],
      ['2d', 'Research - Collaborations'],
      ['2e', 'Research - Innovation (IIC) & Consultancy']]],
    ['3', 'Industry Integrated Programmes and Skilling', [
      ['3a', 'Industry Integrated Programmes and Skilling for UG'],
      ['3b', 'Industry Integrated Programmes and Skilling for PG']]],
    ['4', 'Internship & Placement', [
      ['4a', 'Internship'],
      ['4b', 'Career Progression (Higher Studies & Job Opportunities)']]],
    ['5', 'Ranking & Accreditation', [
      ['5a', 'Ranking Update - NIRF, QS & SDG'],
      ['5b', 'Accreditation Update - IQAC (NAAC), NBA']]],
    ['6', 'Events', [
      ['6a', 'Workshops / Seminars'],
      ['6b', 'International & National Conferences']]],
    ['7', 'Branding', [
      ['7a', 'Website'],
      ['7b', 'Social Media & Newsletters']]],
    ['8', 'Physical Infrastructure', [
      ['8a', 'Maintenance'],
      ['8b', 'Small Projects']]],
    ['9', 'Collaborations', [
      ['9a', 'MOU - Follow-up Actions'],
      ['9b', 'International Collaborations']]],
    ['10', 'Clubs', [
      ['10a', 'Sports, Culturals, Associations, Professional Body'],
      ['10b', 'Alumni Relationship']]],
    ['11', 'Statutory', [
      ['11a', 'Regulatory Matters'],
      ['11b', 'Local Authority']]],
    ['12', 'New Appointments', [['12a', 'New Appointments']]],
    ['13', 'HR Updates', [['13a', 'HR Updates']]],
    ['14', 'Any Other Matters', [['14a', 'Any Other Matters']]]
  ];

  const vmpcSections = [
    ['1', 'Academics', [
      ['1a', 'Academics for UG (Attendance, PT meeting, IA, Slow Learners, Exams and MMP)'],
      ['1b', 'Academics for PG (Attendance, PT meeting, IA, Slow Learners, Exams)'],
      ['1c', 'Faculty Attendance'],
      ['1d', 'Student Affairs - Hostel, Mess, Disciplinary Actions'],
      ['1e', 'Fees Collection']]],
    ['2', 'Research & Innovation', [
      ['2a', 'Research - Publications & Patent'],
      ['2b', 'Research - Seed Money'],
      ['2c', 'Research - External Proposals'],
      ['2d', 'Research - Collaborations'],
      ['2e', 'Research - Innovation (IIC) & Consultancy']]],
    ['3', 'Skilling', [
      ['3a', 'Skilling for UG'],
      ['3b', 'Skilling for PG']]],
    ['4', 'Internship & Placement', [
      ['4a', 'CRMI'],
      ['4b', 'Career Progression (Higher Studies & Job Opportunities), NEET, USMLE etc']]],
    ['5', 'Ranking & Accreditation', [
      ['5a', 'Ranking Update - NIRF, QS & SDG'],
      ['5b', 'Accreditation Update - IQAC (NAAC), ACGME, NABH & NABL']]],
    ['6', 'Events', [
      ['6a', 'MEU, CME'],
      ['6b', 'International & National Conferences']]],
    ['7', 'Branding', [
      ['7a', 'Website'],
      ['7b', 'Social Media & Newsletters']]],
    ['8', 'Physical Infrastructure', [
      ['8a', 'Maintenance'],
      ['8b', 'Small Projects']]],
    ['9', 'Collaborations', [
      ['9a', 'MOU - Follow-up Actions'],
      ['9b', 'International Collaborations']]],
    ['10', 'Clubs', [
      ['10a', 'Sports, Culturals, Associations, Professional Body'],
      ['10b', 'Alumni Relationship']]],
    ['11', 'Statutory', [
      ['11a', 'Regulatory Matters'],
      ['11b', 'Local Authority']]],
    ['12', 'HR Updates', [['12a', 'HR Updates']]],
    ['13', 'New Appointments', [['13a', 'New Appointments']]],
    ['14', 'Any Other Matters', [['14a', 'Any Other Matters']]]
  ];

  const actRows = [];
  [['VMCC', vmccSections], ['VMPC', vmpcSections]].forEach(([campus, secs]) => {
    secs.forEach(([no, title, items]) => {
      items.forEach(([_code, itemTitle]) => {
        actRows.push([campus, no, title, itemTitle]);
      });
    });
  });
  ipmSheet_('Activities').getRange(2, 1, actRows.length, 4).setValues(actRows);

  Logger.log('Setup complete. Spreadsheet: ' + ss.getUrl());
  return ss.getUrl();
}

/* ─────────────────────────────────────────────────────────────────
 * ipmInitialize()
 *
 * Run this ONCE from the GAS editor after deploying new features:
 *   Editor → Run → ipmInitialize
 *
 * Safe to run on an existing database — it only ADDS missing
 * columns / sheets. It never clears or overwrites any data.
 * ───────────────────────────────────────────────────────────────── */
function ipmInitialize() {
  const ss = ipmDb_();
  let log = [];

  // 1. Ensure every required sheet exists (creates if missing)
  const requiredSheets = {
    Users:           ['username','password','role','campus','institution','fullName','createdAt'],
    Institutions:    ['code','name','campus'],
    Activities:      ['campus','sectionNo','sectionTitle','itemTitle'],
    Submissions:     ['id','timestamp','username','campus','institution','semester','weekOf','tasks','status'],
    SubmissionItems: ['submissionId','itemTitle','minutes','actionItems','responsibility','attachments','followUp','status']
  };
  Object.entries(requiredSheets).forEach(([name, headers]) => {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, headers.length)
        .setValues([headers])
        .setFontWeight('bold')
        .setBackground('#1a2332')
        .setFontColor('#ffffff');
      sh.setFrozenRows(1);
      log.push('Created sheet: ' + name);
    }
  });

  // 2. Add missing columns to existing sheets (never removes columns)
  Object.entries(requiredSheets).forEach(([name, headers]) => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const existing = sh.getRange(1, 1, 1, sh.getLastColumn() || 1).getValues()[0]
                       .map(h => String(h).trim());
    headers.forEach((col, idx) => {
      if (!existing.includes(col)) {
        const newCol = sh.getLastColumn() + 1;
        sh.getRange(1, newCol).setValue(col)
          .setFontWeight('bold')
          .setBackground('#1a2332')
          .setFontColor('#ffffff');
        log.push('Added column "' + col + '" to sheet "' + name + '"');
      }
    });
  });

  // 3. Done
  const msg = log.length
    ? 'ipmInitialize() completed:\n' + log.join('\n')
    : 'ipmInitialize() — nothing to do, all sheets and columns are up to date.';
  Logger.log(msg);
  console.log(msg);
}

/* ===== ipmUpdateSubmission ===== */
function ipmUpdateSubmission(token, id, payload) {
  var s=ipmSession_(token); if(s.role!=='HOI') throw new Error('Only HOI can edit submissions');
  var subSh=ipmSheet_('Submissions'),subData=subSh.getDataRange().getValues(),headers=subData[0];
  var idCol=headers.indexOf('id'),usrCol=headers.indexOf('username'),stCol=headers.indexOf('status');
  var rowIdx=subData.findIndex(function(r,i){return i>0&&String(r[idCol])===id;});
  if(rowIdx===-1) return {ok:false,error:'Submission not found'};
  if(String(subData[rowIdx][usrCol]).trim().toLowerCase()!==String(s.username).trim().toLowerCase())
    return {ok:false,error:'You can only edit your own submissions'};
  if(String(subData[rowIdx][stCol]).trim()==='Reviewed')
    return {ok:false,error:'This submission has already been reviewed and cannot be edited'};
  var itemSh=ipmSheet_('SubmissionItems'),itemData=itemSh.getDataRange().getValues(),toDelete=[];
  itemData.forEach(function(r,i){if(i>0&&String(r[0])===id)toDelete.push(i+1);});
  toDelete.reverse().forEach(function(rn){itemSh.deleteRow(rn);});
  (payload.items||[]).forEach(function(it){itemSh.appendRow([id,it.title,it.minutes||'',it.actionItems||'',it.responsibility||'','']);});
  return {ok:true};
}

/* ===== ipmExportInstitutionReport ===== */
function ipmExportInstitutionReport(token, campus, institutionCode) {
  var s=ipmSession_(token); if(s.role!=='IMO') throw new Error('Forbidden');
  var subs=ipmRowsAsObjects_(ipmSheet_('Submissions'))
    .filter(function(r){return r.campus===campus&&r.institution===institutionCode;})
    .sort(function(a,b){return new Date(a.timestamp)-new Date(b.timestamp);})
    .map(ipmNormalizeDates_);
  var allItems=ipmRowsAsObjects_(ipmSheet_('SubmissionItems')),itemsBySub={};
  allItems.forEach(function(it){if(!itemsBySub[it.submissionId])itemsBySub[it.submissionId]=[];itemsBySub[it.submissionId].push(it);});
  var actRows=ipmRowsAsObjects_(ipmSheet_('Activities')).filter(function(r){return r.campus===campus;});
  var sectionOrder=[],sectionMeta={},itemToSec={};
  actRows.forEach(function(r){var sno=String(r.sectionNo);if(!sectionMeta[sno]){sectionMeta[sno]=r.sectionTitle;sectionOrder.push(sno);}if(r.itemTitle)itemToSec[r.itemTitle]=sno;});
  var submissions=subs.map(function(sub){
    var rawItems=itemsBySub[sub.id]||[],grouped={};
    rawItems.forEach(function(it){var sno=itemToSec[it.itemTitle]||'__other__';if(!grouped[sno])grouped[sno]=[];grouped[sno].push({title:it.itemTitle,minutes:it.minutes||'',actionItems:it.actionItems||'',responsibility:it.responsibility||''}); });
    var sections=sectionOrder.filter(function(sno){return grouped[sno]&&grouped[sno].length;}).map(function(sno){return{no:sno,title:sectionMeta[sno]||sno,items:grouped[sno]};});
    if(grouped['__other__'])sections.push({no:'',title:'Other',items:grouped['__other__']});
    return{id:sub.id,timestamp:sub.timestamp,weekOf:sub.weekOf,semester:sub.semester,username:sub.username,status:sub.status,tasks:sub.tasks||'',sections:sections};
  });
  var instRow=ipmRowsAsObjects_(ipmSheet_('Institutions')).find(function(r){return r.code===institutionCode;})||{};
  return{ok:true,institution:{code:institutionCode,name:instRow.name||institutionCode,campus:campus},submissions:submissions,generatedAt:new Date().toLocaleString()};
}

/* ===== ipmGetComparisonReport ===== */
function ipmGetComparisonReport(token, campus) {
  var s=ipmSession_(token); if(s.role!=='IMO') throw new Error('Forbidden');
  var now=new Date(),allSubs=ipmRowsAsObjects_(ipmSheet_('Submissions')),allInst=ipmRowsAsObjects_(ipmSheet_('Institutions'));
  var institutions=campus?allInst.filter(function(i){return i.campus===campus;}):allInst;
  var instCodes=institutions.map(function(i){return i.code;}),instNames={};
  institutions.forEach(function(i){instNames[i.code]=i.name;});

  // ── Resolve report start date (replaces previously hard-coded Mar 2026) ──
  // Priority order:
  //   1. Script Property `IPM_REPORT_START` (ISO date e.g. '2026-03-01')
  //   2. Earliest submission timestamp in the in-scope Submissions data
  //   3. First of the current month (last-resort fallback)
  // The result is also clamped to never be in the future.
  var reportStart = null;
  try {
    var prop = PropertiesService.getScriptProperties().getProperty('IPM_REPORT_START');
    if (prop) {
      var p = new Date(prop);
      if (!isNaN(p.getTime())) reportStart = p;
    }
  } catch(_) {}
  if (!reportStart) {
    var inScopeSubs = allSubs.filter(function(r){return instCodes.indexOf(r.institution)!==-1;});
    var earliest = null;
    inScopeSubs.forEach(function(r){
      var t = new Date(r.timestamp);
      if (isNaN(t.getTime())) return;
      if (!earliest || t < earliest) earliest = t;
    });
    if (earliest) reportStart = new Date(earliest.getFullYear(), earliest.getMonth(), 1, 0, 0, 0, 0);
  }
  if (!reportStart) {
    reportStart = new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0, 0);
  }
  if (reportStart > now) reportStart = new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0, 0);

  // ── Build month axis: reportStart-month → current month ─────────────────
  var months=[];
  var cur=new Date(reportStart.getFullYear(),reportStart.getMonth(),1);
  while(cur.getFullYear()<now.getFullYear()||(cur.getFullYear()===now.getFullYear()&&cur.getMonth()<=now.getMonth())){
    months.push(Utilities.formatDate(cur,Session.getScriptTimeZone(),'MMM yyyy'));
    cur.setMonth(cur.getMonth()+1);
  }

  // ── Build week axis: first Monday on/after reportStart → this Monday ────
  var monStr_=function(dt){var t=new Date(dt);t.setHours(0,0,0,0);var dow=t.getDay();t.setDate(t.getDate()-(dow===0?6:dow-1));return t.toISOString().slice(0,10);};
  var startWk=new Date(reportStart.getTime()); startWk.setHours(0,0,0,0);
  while(startWk.getDay()!==1) startWk.setDate(startWk.getDate()+1); // advance to first Monday
  var thisMonday=new Date(now); thisMonday.setHours(0,0,0,0);
  var dow=thisMonday.getDay(); thisMonday.setDate(thisMonday.getDate()-(dow===0?6:dow-1));
  var weeks=[],wkCur=new Date(startWk);
  while(wkCur<=thisMonday){
    weeks.push(wkCur.toISOString().slice(0,10));
    wkCur.setDate(wkCur.getDate()+7);
  }

  // ── B2 / B3 prep: pre-compute itemsFilled per submission and expected items per campus ──
  // itemsBySid: SubmissionId → count of SubmissionItems rows that have ANY non-blank
  //   content across (minutes / actionItems / responsibility). This lets us reduce
  //   "completeness" to a single integer per submission for cheap summing per cell.
  // expectedByCampus: campus → number of distinct activity items that institution is
  //   expected to fill in per submission (drives the denominator for completeness).
  var allItemsRaw = ipmRowsAsObjects_(ipmSheet_('SubmissionItems'));
  var itemsBySid = {};
  allItemsRaw.forEach(function(it){
    var sid = String(it.submissionId||'').trim(); if (!sid) return;
    var hasContent = (String(it.minutes||'').trim() !== '') ||
                     (String(it.actionItems||'').trim() !== '') ||
                     (String(it.responsibility||'').trim() !== '');
    if (hasContent) itemsBySid[sid] = (itemsBySid[sid]||0) + 1;
  });
  var allActRows = ipmRowsAsObjects_(ipmSheet_('Activities'));
  var expectedByCampus = {};
  var seenAct = {};
  allActRows.forEach(function(r){
    if (!r.itemTitle) return;
    var k = String(r.campus||'') + '||' + String(r.itemTitle).trim();
    if (seenAct[k]) return;
    seenAct[k] = true;
    expectedByCampus[r.campus] = (expectedByCampus[r.campus]||0) + 1;
  });

  // ── Cell shape: every cell is an OBJECT, not just a number, to carry quality
  //   and status mix in addition to count. The frontend reads .count for the
  //   primary headline figure (backwards-compatible reading via `cell.count||0`)
  //   and uses .itemsFilled/.expectedItems for the completeness shading and
  //   .byStatus for the stack-segmented bar.
  var monthlyData={},weeklyData={};
  function _newCell(){ return { count:0, itemsFilled:0, expectedItems:0, byStatus:{} }; }
  instCodes.forEach(function(c){
    monthlyData[c]={}; months.forEach(function(m){monthlyData[c][m]=_newCell();});
    weeklyData[c]={};  weeks.forEach (function(w){weeklyData[c][w]=_newCell();});
  });
  allSubs.filter(function(r){return instCodes.indexOf(r.institution)!==-1;}).forEach(function(r){
    var d=new Date(r.timestamp);if(isNaN(d))return;
    var sid = String(r.id||'').trim();
    var status = String(r.status||'').trim() || 'Submitted';
    var filled = sid ? (itemsBySid[sid]||0) : 0;
    var expected = expectedByCampus[r.campus] || 0;

    var mk=Utilities.formatDate(d,Session.getScriptTimeZone(),'MMM yyyy');
    var mCell = monthlyData[r.institution] && monthlyData[r.institution][mk];
    if (mCell) {
      mCell.count++;
      mCell.itemsFilled    += filled;
      mCell.expectedItems  += expected;
      mCell.byStatus[status] = (mCell.byStatus[status]||0) + 1;
    }
    var wk=monStr_(d);
    var wCell = weeklyData[r.institution] && weeklyData[r.institution][wk];
    if (wCell) {
      wCell.count++;
      wCell.itemsFilled    += filled;
      wCell.expectedItems  += expected;
      wCell.byStatus[status] = (wCell.byStatus[status]||0) + 1;
    }
  });

  // ── Tier-5 / B4: participation rate per institution ─────────────────────
  // Each institution is expected to file once a week. Participation = actual /
  // elapsed-weeks-since-first-submission. Caps at 100 to avoid duplicate-
  // submission inflation. Null when no submissions found yet (can't measure).
  var firstSubByInst = {};
  allSubs.forEach(function(r){
    if (instCodes.indexOf(r.institution) === -1) return;
    var d = new Date(r.timestamp); if (isNaN(d.getTime())) return;
    if (!firstSubByInst[r.institution] || d < firstSubByInst[r.institution]) {
      firstSubByInst[r.institution] = d;
    }
  });
  var participation = {};
  instCodes.forEach(function(c){
    var first = firstSubByInst[c];
    if (!first) {
      participation[c] = { expected: 0, actual: 0, pct: null };
      return;
    }
    // Count weeks since first-submission Monday
    var startMon = new Date(first); startMon.setHours(0,0,0,0);
    var dow = startMon.getDay();
    startMon.setDate(startMon.getDate() - (dow === 0 ? 6 : dow - 1));
    var elapsedWeeks = Math.max(1, Math.floor((thisMonday.getTime() - startMon.getTime()) / (7*24*60*60*1000)) + 1);
    var actualSubs = 0;
    Object.keys(weeklyData[c] || {}).forEach(function(wk){
      actualSubs += (weeklyData[c][wk].count || 0);
    });
    participation[c] = {
      expected: elapsedWeeks,
      actual:   actualSubs,
      pct:      Math.min(100, Math.round((actualSubs / elapsedWeeks) * 100))
    };
  });

  // ── Tier-5 / B6: 4-week moving average of weekly counts per institution ─
  // Smooths out single-week noise so trends are easier to spot. For a given
  // week, average is over THAT week and the 3 preceding ones (or fewer if at
  // the start of the series). Stored as a parallel structure to weeklyData.
  var weeklyMovingAvg = {};
  instCodes.forEach(function(c){
    weeklyMovingAvg[c] = {};
    weeks.forEach(function(wk, idx){
      var sum = 0, n = 0;
      for (var k = Math.max(0, idx - 3); k <= idx; k++) {
        sum += (weeklyData[c][weeks[k]] && weeklyData[c][weeks[k]].count) || 0;
        n++;
      }
      weeklyMovingAvg[c][wk] = n > 0 ? Math.round((sum / n) * 10) / 10 : 0;
    });
  });

  return{ok:true,
    institutions:instCodes.map(function(c){return{code:c,name:instNames[c]||c};}),
    months:months,weeks:weeks,monthlyData:monthlyData,weeklyData:weeklyData,
    participation: participation,            // Tier-5 / B4
    weeklyMovingAvg: weeklyMovingAvg,         // Tier-5 / B6
    reportStart: Utilities.formatDate(reportStart, Session.getScriptTimeZone(), 'yyyy-MM-dd')};
}

/* ===== ipmGetCampusComparisonReport — Tier-5 / B5 =========================
 * Aggregates the institution-level heatmap into a campus-level view. Useful
 * for Chancellor-level comparisons (e.g. VMCC vs VMPC). Reuses the existing
 * ipmGetComparisonReport response and rolls up by campus.
 */
function ipmGetCampusComparisonReport(token) {
  var s = ipmSession_(token); if (s.role !== 'IMO') throw new Error('Forbidden');

  // Pull the un-filtered institution-level report (no campus arg)
  var instReport = ipmGetComparisonReport(token, '');
  if (!instReport.ok) return instReport;

  // Build instCode → campus map
  var allInst = ipmRowsAsObjects_(ipmSheet_('Institutions'));
  var instToCampus = {};
  allInst.forEach(function(i){
    instToCampus[String(i.code||'').trim()] = String(i.campus||'').trim();
  });

  // Collect campuses
  var campusSet = {};
  instReport.institutions.forEach(function(i){
    var c = instToCampus[i.code]; if (c) campusSet[c] = true;
  });
  var campuses = Object.keys(campusSet).sort();

  // Aggregate cells: campus × period → sum of all institutions in that campus
  function _aggCells(srcByInst, periodKeys) {
    var out = {};
    campuses.forEach(function(camp){
      out[camp] = {};
      periodKeys.forEach(function(pk){
        out[camp][pk] = { count:0, itemsFilled:0, expectedItems:0, byStatus:{} };
      });
    });
    Object.keys(srcByInst).forEach(function(instCode){
      var camp = instToCampus[instCode]; if (!camp || !out[camp]) return;
      Object.keys(srcByInst[instCode]).forEach(function(pk){
        var src = srcByInst[instCode][pk];
        if (!out[camp][pk] || !src) return;
        out[camp][pk].count          += (src.count || 0);
        out[camp][pk].itemsFilled    += (src.itemsFilled || 0);
        out[camp][pk].expectedItems  += (src.expectedItems || 0);
        var bs = src.byStatus || {};
        Object.keys(bs).forEach(function(st){
          out[camp][pk].byStatus[st] = (out[camp][pk].byStatus[st] || 0) + bs[st];
        });
      });
    });
    return out;
  }

  // Aggregate participation: sum actual + expected across institutions in each campus
  var campusParticipation = {};
  campuses.forEach(function(camp){
    var totActual = 0, totExpected = 0, anyData = false;
    Object.keys(instToCampus).forEach(function(instCode){
      if (instToCampus[instCode] !== camp) return;
      var p = instReport.participation && instReport.participation[instCode];
      if (!p) return;
      if (p.pct !== null) anyData = true;
      totActual   += (p.actual || 0);
      totExpected += (p.expected || 0);
    });
    campusParticipation[camp] = {
      actual: totActual, expected: totExpected,
      pct: anyData && totExpected > 0 ? Math.min(100, Math.round((totActual/totExpected) * 100)) : null
    };
  });

  return {
    ok: true,
    campuses: campuses.map(function(c){return{code:c,name:c};}),
    months: instReport.months,
    weeks:  instReport.weeks,
    monthlyData: _aggCells(instReport.monthlyData, instReport.months),
    weeklyData:  _aggCells(instReport.weeklyData,  instReport.weeks),
    participation: campusParticipation,
    reportStart: instReport.reportStart
  };
}

/* ===== ipmGetCellSubmissions — B7 drill-down for the comparison heatmap =====
 * Returns the actual submissions (id, week, status, items-filled summary)
 * underlying a clicked heatmap cell. UI calls this on cell click and shows a
 * side panel listing the submissions, each with a link/button to open the
 * full submission via ipmExportInstitutionReport or similar.
 *
 * Args:
 *   token, campus, institutionCode  — same auth/scope as ipmGetComparisonReport
 *   mode: 'month' | 'week'          — matches the active heatmap view
 *   bucketKey:                       — 'MMM yyyy' for month, ISO Monday for week
 */
function ipmGetCellSubmissions(token, campus, institutionCode, mode, bucketKey) {
  var s = ipmSession_(token); if (s.role !== 'IMO') throw new Error('Forbidden');
  if (!institutionCode || !bucketKey) return { ok: false, error: 'Missing inputs' };

  var subs = ipmRowsAsObjects_(ipmSheet_('Submissions'))
    .filter(function(r){
      if (campus && r.campus !== campus) return false;
      return r.institution === institutionCode;
    });

  // Pre-compute itemsFilled per submission
  var itemsBySid = {};
  ipmRowsAsObjects_(ipmSheet_('SubmissionItems')).forEach(function(it){
    var sid = String(it.submissionId||'').trim(); if (!sid) return;
    var has = (String(it.minutes||'').trim() !== '') ||
              (String(it.actionItems||'').trim() !== '') ||
              (String(it.responsibility||'').trim() !== '');
    if (has) itemsBySid[sid] = (itemsBySid[sid]||0) + 1;
  });

  // Filter by bucket
  var monStr_ = function(dt){
    var t = new Date(dt); t.setHours(0,0,0,0);
    var dow = t.getDay(); t.setDate(t.getDate() - (dow===0?6:dow-1));
    return t.toISOString().slice(0,10);
  };
  var matched = subs.filter(function(r){
    var d = new Date(r.timestamp); if (isNaN(d.getTime())) return false;
    if (mode === 'month') {
      return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMM yyyy') === bucketKey;
    }
    return monStr_(d) === bucketKey;
  });

  // Sort newest-first
  matched.sort(function(a,b){ return new Date(b.timestamp) - new Date(a.timestamp); });

  return {
    ok: true,
    bucketKey: bucketKey,
    mode: mode,
    institutionCode: institutionCode,
    submissions: matched.map(function(r){
      var sid = String(r.id||'');
      return {
        id:          sid,
        timestamp:   r.timestamp instanceof Date
                       ? Utilities.formatDate(r.timestamp, Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm')
                       : String(r.timestamp||''),
        weekOf:      r.weekOf instanceof Date
                       ? Utilities.formatDate(r.weekOf, Session.getScriptTimeZone(), 'dd MMM yyyy')
                       : String(r.weekOf||''),
        status:      String(r.status||'Submitted'),
        username:    String(r.username||''),
        itemsFilled: itemsBySid[sid] || 0
      };
    })
  };
}

/* ===== Best-practice tags — C6 ============================================
 * Lets IMO mark a specific cell in the weekly-updates comparison as a "best
 * practice" example so other institutions can learn from it. Stored on a
 * tag sheet auto-created on first use.
 *
 * Sheet: IPM_Tags
 * Columns: TagID | Type | Scope | Note | TaggedBy | TaggedAt
 *   Type  = 'best_practice' (extensible for future tag types)
 *   Scope = JSON, e.g. {"institutionCode":"MEC","weekKey":"2026-04-13","itemTitle":"Faculty Development Programme"}
 */
function _ipmTagsSheet_() {
  var sh = ipmDb_().getSheetByName('IPM_Tags');
  if (!sh) {
    sh = ipmDb_().insertSheet('IPM_Tags');
    var headers = ['TagID','Type','Scope','Note','TaggedBy','TaggedAt'];
    sh.getRange(1,1,1,headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#1a2332').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function ipmAddBestPracticeTag(token, payload) {
  var s = ipmSession_(token); if (s.role !== 'IMO') throw new Error('Forbidden');
  payload = payload || {};
  if (!payload.institutionCode || !payload.weekKey || !payload.itemTitle) {
    return { ok: false, error: 'institutionCode, weekKey, itemTitle required' };
  }
  var sh = _ipmTagsSheet_();
  var tagId = 'BP_' + Date.now().toString(36).toUpperCase() + '_' + Math.floor(Math.random()*1000);
  var scope = JSON.stringify({
    institutionCode: payload.institutionCode,
    weekKey:         payload.weekKey,
    itemTitle:       payload.itemTitle
  });
  sh.appendRow([tagId, 'best_practice', scope, String(payload.note||''), s.username||'', new Date()]);
  return { ok: true, tagId: tagId };
}

function ipmRemoveBestPracticeTag(token, tagId) {
  var s = ipmSession_(token); if (s.role !== 'IMO') throw new Error('Forbidden');
  if (!tagId) return { ok: false, error: 'tagId required' };
  var sh = _ipmTagsSheet_();
  var data = sh.getDataRange().getValues();
  for (var i=1; i<data.length; i++) {
    if (String(data[i][0]) === String(tagId)) {
      sh.deleteRow(i+1);
      return { ok: true, deleted: tagId };
    }
  }
  return { ok: false, error: 'Tag not found' };
}

function ipmListBestPracticeTags(token) {
  var s = ipmSession_(token); if (s.role !== 'IMO') throw new Error('Forbidden');
  var sh = _ipmTagsSheet_();
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return { ok: true, tags: [] };
  var tags = [];
  for (var i=1; i<data.length; i++) {
    var r = data[i];
    if (String(r[1]) !== 'best_practice') continue;
    var scope = {}; try { scope = JSON.parse(r[2]||'{}'); } catch(_){}
    tags.push({
      tagId: String(r[0]),
      scope: scope,
      note:  String(r[3]||''),
      taggedBy: String(r[4]||''),
      taggedAt: r[5] instanceof Date
                  ? Utilities.formatDate(r[5], Session.getScriptTimeZone(), 'dd MMM yyyy')
                  : ''
    });
  }
  return { ok: true, tags: tags };
}

/* ===== ipmGetWeeklyUpdatesComparison ===== */
function ipmGetWeeklyUpdatesComparison(token, payload) {
  var s=ipmSession_(token); if(s.role!=='IMO') throw new Error('Forbidden');
  var campus=payload.campus||'',instCodes=payload.institutionCodes||[],weekKeys=payload.weekKeys||[];
  if(!instCodes.length) return{ok:false,error:'Select at least one institution'};
  if(!weekKeys.length)  return{ok:false,error:'Select at least one week'};
  var normWk=function(v){return String(v instanceof Date?v.toISOString():(v||'')).trim().slice(0,10);};
  var allInst=ipmRowsAsObjects_(ipmSheet_('Institutions')),instNames={};
  allInst.forEach(function(i){instNames[i.code]=i.name;});
  var actRows=ipmRowsAsObjects_(ipmSheet_('Activities')).filter(function(r){return r.campus===campus;});
  var sectionOrder=[],sectionMeta={},itemOrder=[],itemToSec={};
  actRows.forEach(function(r){var sno=String(r.sectionNo);if(!sectionMeta[sno]){sectionMeta[sno]=r.sectionTitle;sectionOrder.push(sno);}if(r.itemTitle&&itemOrder.indexOf(r.itemTitle)===-1){itemOrder.push(r.itemTitle);itemToSec[r.itemTitle]=sno;}});
  var allSubs=ipmRowsAsObjects_(ipmSheet_('Submissions')),allItems=ipmRowsAsObjects_(ipmSheet_('SubmissionItems'));
  var subIndex={};
  allSubs.filter(function(sub){return instCodes.indexOf(sub.institution)!==-1&&weekKeys.indexOf(normWk(sub.weekOf))!==-1;})
    .forEach(function(sub){subIndex[sub.id]={instCode:sub.institution,weekKey:normWk(sub.weekOf),status:sub.status};});
  var lookup={};
  allItems.filter(function(it){return!!subIndex[it.submissionId];}).forEach(function(it){
    var meta=subIndex[it.submissionId];
    var key=meta.instCode+'||'+meta.weekKey+'||'+String(it.itemTitle).trim();
    lookup[key]={
      minutes:        it.minutes        || '',
      actionItems:    it.actionItems    || '',
      responsibility: it.responsibility || '',
      // Newly surfaced fields — already stored in SubmissionItems sheet
      // (see ipmEnsureColumns_ schema) but previously hidden from the UI.
      followUp:       it.followUp       || '',
      status:         it.status         || '',
      hasAttachments: !!(it.attachments && String(it.attachments).trim())
    };
  });
  var columns=[];
  instCodes.forEach(function(ic){weekKeys.forEach(function(wk){columns.push({instCode:ic,instName:instNames[ic]||ic,weekKey:wk});});});
  var sections=sectionOrder.map(function(sno){
    var items=itemOrder.filter(function(t){return itemToSec[t]===sno;}).map(function(title){
      var cells=columns.map(function(col){var key=col.instCode+'||'+col.weekKey+'||'+title;return lookup[key]||null;});
      return{title:title,cells:cells};
    });
    var hasData=items.some(function(it){return it.cells.some(function(c){return c!==null;});});
    return{sectionNo:sno,sectionTitle:sectionMeta[sno],items:items,hasData:hasData};
  }).filter(function(sec){return sec.hasData;});
  // Tier-5 / C7: auto-detect "cross-week" layout (1 institution × many weeks).
  // The UI uses this hint to flip column headers (week labels become primary,
  // institution name moves to a single page-level header).
  var mode = (instCodes.length === 1 && weekKeys.length > 1) ? 'cross-week' : 'cross-institution';
  var subjectInstitution = (mode === 'cross-week')
    ? { code: instCodes[0], name: instNames[instCodes[0]] || instCodes[0] }
    : null;
  return{ok:true,columns:columns,sections:sections,
    mode: mode,
    subjectInstitution: subjectInstitution};
}

/* ===== ipmGetAvailableWeeksForInstitutions ===== */
function ipmGetAvailableWeeksForInstitutions(token, campus, instCodes) {
  var s=ipmSession_(token); if(s.role!=='IMO') throw new Error('Forbidden');
  var normWk=function(v){return String(v instanceof Date?v.toISOString():(v||'')).trim().slice(0,10);};
  var seen={},weeks=[];
  ipmRowsAsObjects_(ipmSheet_('Submissions')).filter(function(sub){
    return(!campus||sub.campus===campus)&&(instCodes||[]).indexOf(sub.institution)!==-1;
  }).forEach(function(sub){var wk=normWk(sub.weekOf);if(wk&&!seen[wk]){seen[wk]=true;weeks.push(wk);}});
  weeks.sort().reverse();
  return{ok:true,weeks:weeks};
}