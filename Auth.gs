// =============================================================================
// VMRF-DU Institutional Monitoring System — Authentication Services
// =============================================================================

function _isNoHodInstitution_(instNameOrCode) {
  if (!instNameOrCode) return false;
  var raw = String(instNameOrCode).trim();
  if (!raw) return false;
  var val     = raw.toLowerCase();
  var valBare = val.replace(/\s*\([^)]*\)\s*$/, '').trim();
  // Also try resolving the input through _resolveInstCode — handles cases where
  // Faculty_Master stores something like "School of Allied Health Sciences"
  // (without campus suffix) or just the bare code.
  var resolvedCode = '';
  try { resolvedCode = String(_resolveInstCode('', raw) || '').toLowerCase(); } catch(e) {}
  var campuses = Object.keys(INSTITUTION_HIERARCHY);
  for (var ci = 0; ci < campuses.length; ci++) {
    var insts = INSTITUTION_HIERARCHY[campuses[ci]].institutions;
    var keys = Object.keys(insts);
    for (var ii = 0; ii < keys.length; ii++) {
      var ins = insts[keys[ii]];
      if (!ins.noHod) continue;
      var code    = String(ins.code||'').toLowerCase();
      var fullKey = keys[ii].toLowerCase();
      // Strip trailing "(CODE)" suffix so "Vinayaka Mission's Law School" matches
      // "Vinayaka Mission's Law School (VMLS)" stored in the hierarchy
      var bareKey = fullKey.replace(/\s*\([^)]*\)\s*$/, '').trim();
      if (code === val || fullKey === val || bareKey === val) return true;
      if (code && code === valBare) return true;
      if (code && resolvedCode && code === resolvedCode) return true;
      // Loose containment match — handles "School of Allied Health Sciences"
      // matching "School of Allied Health Sciences - Chennai Campus" and reverse.
      if (bareKey && (bareKey.indexOf(valBare) >= 0 || valBare.indexOf(bareKey) >= 0) &&
          Math.min(bareKey.length, valBare.length) >= 6) return true;
    }
  }
  return false;
}

// ─── PRE-LOGIN BOOTSTRAP ─────────────────────────────────────────────────────
// Returns the campus → institution hierarchy for the login page's HOI
// cascade dropdowns. Deliberately requires NO authentication so the login
// page can populate campus / institution pickers before the user signs in.
// Structure is identical to the `hierarchy` field of getConfig(), so the
// client can use the same rendering helper for both.
function getLoginHierarchy() {
  try {
    return { ok: true, hierarchy: _buildHierarchyForClient_() };
  } catch (e) {
    return { ok: false, error: e && e.message ? e.message : String(e) };
  }
}

// ─── AUTH ─────────────────────────────────────────────────────────────────────
// Four separate role logins — each stored separately:
//   Faculty  → Faculty_Master sheet (Faculty ID + password or Google)
//   HOD/HOI/IMO → Staff_Master sheet (email + password or Google)
//
// Registration:
//   Faculty:  activateAccount(facultyID, password, confirmPassword)
//             — activates an IMO-pre-enrolled Pending row
//   Staff:    staffRegister(f)
//             — self-registers with a role-specific reg code
//             — codes stored in ScriptProperties: REGCODE_HOD, REGCODE_HOI, REGCODE_IMO

// ── Helper: hash password (Djb2) ─────────────────────────────────────────────
// ─── RESOLVE INSTITUTION CODE ────────────────────────────────────────────────
// Given a campus value and/or institution value (either full name or code),
// returns the institution code (e.g. 'AVIT', 'SAS'). Returns '' if not found.
// Used by every login function so the client gets a reliable instCode without
// having to do fragile client-side name matching.
function _resolveInstCode(campusVal, instVal) {
  var vC = String(campusVal || '').trim();
  var vI = String(instVal   || '').trim();
  if (!vI) return '';
  var campusKeys = Object.keys(INSTITUTION_HIERARCHY);
  for (var ci = 0; ci < campusKeys.length; ci++) {
    var cName = campusKeys[ci];
    var node  = INSTITUTION_HIERARCHY[cName];
    var cCode = String(node.code || '');
    // If a campus was supplied, skip non-matching campuses
    if (vC) {
      var campusMatch = (vC === cName) ||
                        (cCode && cCode.toLowerCase() === vC.toLowerCase());
      if (!campusMatch) continue;
    }
    var insts = node.institutions || {};
    var iNames = Object.keys(insts);
    for (var ii = 0; ii < iNames.length; ii++) {
      var iName = iNames[ii];
      var iCode = String(insts[iName].code || '');
      if (vI === iName || (iCode && iCode.toLowerCase() === vI.toLowerCase())) {
        return iCode;
      }
    }
  }
  return '';
}

// ── 1. Faculty login (Faculty ID + password) ──────────────────────────────────
function facultyLogin(email, password) {
  if (!email)    throw new Error('Please enter your email address.');
  if (!password) throw new Error('Please enter your password.');

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.FACULTY);
  if (!sheet) throw new Error('Faculty sheet not found. Please run Setup from the VMRF IMO menu first.');

  var data = sheet.getDataRange().getValues();
  if (!data || data.length < 2) throw new Error('No faculty accounts registered yet.');

  // Read headers — trim whitespace, handle any schema (old with FacultyID or new without)
  var h = data[0].map(function(v){ return String(v||'').trim(); });

  var emI  = h.indexOf('Email');
  var nmI  = h.indexOf('FacultyName');
  var pwI  = h.indexOf('PasswordHash');
  var stI  = h.indexOf('Status');
  // Campus / Institution / Department columns — needed so the client can show the
  // correct institution logo and department in the topbar after login.
  var caI  = h.indexOf('Campus');
  var inI  = h.indexOf('Institution');
  var deI  = h.indexOf('Department');
  var dgI  = h.indexOf('Designation');

  // Fallback: if Email column not found by name, scan headers for anything containing 'mail'
  if (emI < 0) {
    for (var x = 0; x < h.length; x++) {
      if (h[x].toLowerCase().indexOf('mail') >= 0) { emI = x; break; }
    }
  }
  if (emI < 0) throw new Error('Faculty sheet structure is invalid. Please run Setup first.');

  var emailLc = String(email).trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    // Skip blank rows
    if (!row || !row[emI]) continue;
    var rowEmail = String(row[emI]||'').trim().toLowerCase();
    if (rowEmail !== emailLc) continue;

    // ── Found the email — now validate ──
    var status = (stI >= 0) ? String(row[stI]||'').trim() : '';

    // Treat blank/missing status as Active (for older accounts)
    if (status === 'Pending') {
      throw new Error('Your account is pending approval. Please contact the administrator.');
    }
    if (status !== '' && status !== 'Active') {
      throw new Error('Your account is not active. Please contact the administrator.');
    }

    // Password check
    var stored = (pwI >= 0) ? String(row[pwI]||'').trim() : '';
    if (!stored) throw new Error('No password is set for this account. Please use Forgot Password to set one.');
    if (stored !== _hashPwd(password)) throw new Error('Incorrect password. Please try again.');

    var facultyName = (nmI >= 0) ? String(row[nmI]||'').trim() : emailLc;
    return {
      success:         true,
      role:            'FACULTY',
      facultyID:       emailLc,
      facultyName:     facultyName,
      email:           emailLc,
      campus:          (caI >= 0) ? String(row[caI]||'').trim() : '',
      institution:     (inI >= 0) ? String(row[inI]||'').trim() : '',
      department:      (deI >= 0) ? String(row[deI]||'').trim() : '',
      designation:     (dgI >= 0) ? String(row[dgI]||'').trim() : '',
      institutionCode: _resolveInstCode(
                         (caI >= 0) ? String(row[caI]||'').trim() : '',
                         (inI >= 0) ? String(row[inI]||'').trim() : ''
                       ),
      config:          getConfig()
    };
  }

  throw new Error('No account found for "' + email + '". Please check your email address or register first.');
}

// ── 2. Staff login (HOD/HOI: Staff ID + password | IMO: Script Properties) ──
//
// HOI signature extended to support the IPM-style login cascade:
//   staffLogin('HOI', username, password, selectedCampus, selectedInstitution)
// If `selectedCampus` and `selectedInstitution` are provided, the backend
// verifies they match the account's home campus/institution. This prevents
// a user who knows another institution's credentials from signing in as the
// wrong institution just by picking different dropdowns on the login page.
// `selectedCampus` / `selectedInstitution` are optional — older clients that
// still call staffLogin with 3 args keep working.
function staffLogin(role, staffID, password, selectedCampus, selectedInstitution) {
  if (!role)     throw new Error('Role is required.');
  if (!staffID)  throw new Error('Please enter your credentials.');
  if (!password) throw new Error('Please enter your password.');
  role = role.toUpperCase();
  if (['HOD','HOI','IMO'].indexOf(role) < 0) throw new Error('Invalid role.');

  // ── IMO: accepts IPM Users sheet creds OR legacy Script Properties ──
  if (role === 'IMO') {
    var imoIpm = _tryIpmLoginForMain_('IMO', staffID, password);
    // Resolve the IMO email once — used by both auth paths below. Drawn from
    // (1) the IPM Users sheet when IPM auth succeeded, (2) the staffID typed
    // at login if it looks like an email, (3) the IMO_EMAIL script property.
    // This is what populates APP.userEmail client-side, which the Saved Views
    // / shared-views feature uses as the owner identity.
    var _imoPropEmail = PropertiesService.getScriptProperties().getProperty('IMO_EMAIL') || 'imo@vmrf.edu.in';
    var _typedLooksEmail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(staffID || ''));
    if (imoIpm.ok) {
      var _imoEmailIpm = (imoIpm.user && (imoIpm.user.email || imoIpm.user.username)) ||
                         (_typedLooksEmail ? String(staffID).trim().toLowerCase() : _imoPropEmail);
      return {
        success:true, role:'IMO', staffID:'IMO-ADMIN', staffName: imoIpm.user.fullName||'IMO Admin',
        email: _imoEmailIpm,
        config:getConfig(), ipmToken: imoIpm.token, ipmUser: imoIpm.user
      };
    }
    var props = PropertiesService.getScriptProperties();
    var imoEmail = props.getProperty('IMO_EMAIL')    || 'imo@vmrf.edu.in';
    var imoPwd   = props.getProperty('IMO_PASSWORD') || 'IMO@VMRF2024';
    if (String(staffID).trim().toLowerCase() === imoEmail.toLowerCase() && String(password) === imoPwd) {
      // Legacy login succeeded; mint IPM token too so IPM sub-views work.
      var sso = ipmSsoTokenForMainUser('IMO', '', '');
      return {
        success:true, role:'IMO', staffID:'IMO-ADMIN', staffName:'IMO Admin',
        email: imoEmail,
        config:getConfig(),
        ipmToken: sso && sso.ok ? sso.token : null,
        ipmUser:  sso && sso.ok ? sso.user  : null
      };
    }
    throw new Error('IMO credentials not recognised.');
  }


  // ── HOI: accepts IPM Users sheet creds (per-institution) OR legacy Script Properties ──
  if (role === 'HOI') {
    // Normalise the cascade selections (may be empty if client doesn't send them)
    var selCampus = String(selectedCampus || '').trim();
    var selInst   = String(selectedInstitution || '').trim();

    // The login page sends FULL names ("Vinayaka Mission's Chennai Campus",
    // "Aarupadai Veedu Institute of Technology (AVIT)") but the IPM Users
    // sheet stores SHORT CODES ("VMCC", "AVIT"). Translate both sides to
    // codes (using INSTITUTION_HIERARCHY) so the comparison is apples-to-apples.
    var _hoiResolve = function(campusVal, instVal) {
      var out = { campusCode: '', instCode: '' };
      if (!campusVal && !instVal) return out;
      try {
        // Direct campus-key match (full name) → use its code
        if (INSTITUTION_HIERARCHY[campusVal]) {
          out.campusCode = INSTITUTION_HIERARCHY[campusVal].code || '';
        } else {
          // Campus may have been sent as a code already
          Object.keys(INSTITUTION_HIERARCHY).forEach(function(k){
            if (INSTITUTION_HIERARCHY[k].code === campusVal) out.campusCode = campusVal;
          });
        }
        // Institution: find within the resolved campus (or search all campuses)
        var searchIn = out.campusCode
          ? [Object.keys(INSTITUTION_HIERARCHY).filter(function(k){ return INSTITUTION_HIERARCHY[k].code === out.campusCode; })[0]]
          : Object.keys(INSTITUTION_HIERARCHY);
        for (var i=0; i<searchIn.length; i++) {
          var camp = searchIn[i]; if (!camp) continue;
          var insts = INSTITUTION_HIERARCHY[camp].institutions || {};
          if (insts[instVal]) { out.instCode = insts[instVal].code || ''; break; }
          // Institution may have been sent as a code already
          var hit = Object.keys(insts).filter(function(n){ return insts[n].code === instVal; })[0];
          if (hit) { out.instCode = instVal; break; }
        }
      } catch(e) {}
      return out;
    };

    var selResolved = _hoiResolve(selCampus, selInst);

    var hoiIpm = _tryIpmLoginForMain_('HOI', staffID, password);
    if (hoiIpm.ok) {
      var acctCampus = String(hoiIpm.user.campus      || '').trim(); // stored as code (e.g. VMCC)
      var acctInst   = String(hoiIpm.user.institution || '').trim(); // stored as code (e.g. AVIT)
      if (selResolved.campusCode && acctCampus && selResolved.campusCode !== acctCampus) {
        throw new Error('This account belongs to a different campus.');
      }
      if (selResolved.instCode && acctInst && selResolved.instCode !== acctInst) {
        throw new Error('This account belongs to a different institution.');
      }
      return {
        success:true, role:'HOI', staffID:'HOI-ADMIN', staffName: hoiIpm.user.fullName||'Head of Institution',
        campus: hoiIpm.user.campus, institution: hoiIpm.user.institution,
        config:getConfig(), ipmToken: hoiIpm.token, ipmUser: hoiIpm.user
      };
    }
    var p2 = PropertiesService.getScriptProperties();
    var hoiEmail = p2.getProperty('HOI_EMAIL')    || 'hoi@vmrf.edu.in';
    var hoiPwd   = p2.getProperty('HOI_PASSWORD') || 'HOI@VMRF2024';
    if (String(staffID).trim().toLowerCase() === hoiEmail.toLowerCase() && String(password) === hoiPwd) {
      // Legacy HOI — honour the cascade selection if one was provided.
      // ipmSsoTokenForMainUser looks up by (campus code, institution code)
      // in the IPM Users sheet, so pass the resolved codes.
      if (selResolved.campusCode && selResolved.instCode) {
        var ssoLegacy = ipmSsoTokenForMainUser('HOI', selResolved.campusCode, selResolved.instCode);
        if (ssoLegacy && ssoLegacy.ok) {
          return {
            success:true, role:'HOI', staffID:'HOI-ADMIN',
            staffName:'Head of Institution',
            campus: selResolved.campusCode, institution: selResolved.instCode,
            config:getConfig(),
            ipmToken: ssoLegacy.token, ipmUser: ssoLegacy.user
          };
        }
      }
      return {
        success:true, role:'HOI', staffID:'HOI-ADMIN', staffName:'Head of Institution', config:getConfig(),
        ipmToken: null, ipmUser: null, needsHoiInstitutionPick: true
      };
    }
    throw new Error('HOI credentials not recognised.');
  }

  // ── HOD: authenticated against Staff_Master (campus + institution + department scoped) ──
  if (role === 'HOD') {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var sh   = ss.getSheetByName(SH.STAFF);
    if (!sh) throw new Error('Staff sheet not found. Run initializeSystem first.');
    _ensureSheetColumns(sh, SCHEMA.Staff_Master); // safe upgrade: add Campus/Institution if missing
    var data = sh.getDataRange().getValues(), h = data[0];
    var emI  = h.indexOf('Email'), pwI = h.indexOf('PasswordHash');
    var nmI  = h.indexOf('StaffName'), depI = h.indexOf('Department');
    var stI  = h.indexOf('Status'), rlI = h.indexOf('Role'), idI = h.indexOf('StaffID');
    var camI = h.indexOf('Campus'), insI = h.indexOf('Institution');

    // The login form sends the campus/institution in whichever form the user
    // picked (full name, e.g. "Vinayaka Mission's Chennai Campus", or a code
    // like "VMCC"). Staff_Master stores whatever seedHODAccounts wrote (full
    // name by default). Resolve both sides to CODES for a robust compare.
    var _hodResolve = function(campusVal, instVal) {
      var out = { campusCode: '', instCode: '' };
      if (!campusVal && !instVal) return out;
      try {
        if (INSTITUTION_HIERARCHY[campusVal]) {
          out.campusCode = INSTITUTION_HIERARCHY[campusVal].code || '';
        } else {
          Object.keys(INSTITUTION_HIERARCHY).forEach(function(k){
            if (INSTITUTION_HIERARCHY[k].code === campusVal) out.campusCode = campusVal;
          });
        }
        var searchIn = out.campusCode
          ? [Object.keys(INSTITUTION_HIERARCHY).filter(function(k){ return INSTITUTION_HIERARCHY[k].code === out.campusCode; })[0]]
          : Object.keys(INSTITUTION_HIERARCHY);
        for (var i=0; i<searchIn.length; i++) {
          var camp = searchIn[i]; if (!camp) continue;
          var insts = INSTITUTION_HIERARCHY[camp].institutions || {};
          if (insts[instVal]) { out.instCode = insts[instVal].code || ''; break; }
          var hit = Object.keys(insts).filter(function(n){ return insts[n].code === instVal; })[0];
          if (hit) { out.instCode = instVal; break; }
        }
      } catch(e) {}
      return out;
    };

    var want = _hodResolve(String(selectedCampus||'').trim(), String(selectedInstitution||'').trim());

    // Case-insensitive match on staffID (username OR email — both are supported)
    var ident = String(staffID).trim().toLowerCase();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][rlI]||'').toUpperCase() !== 'HOD') continue;
      if (String(data[i][emI]||'').trim().toLowerCase() !== ident) continue;
      if (String(data[i][stI]||'') !== 'Active')
        throw new Error('This HOD account is not yet active. Contact the administrator.');
      if (_hashPwd(password) !== String(data[i][pwI]||''))
        throw new Error('Incorrect password.');

      // BLOCK HOD login for institutions that have no HOD layer in their workflow.
      // Faculty from SAHS, VSEP, VSHS, VMLS report directly to HOI.
      var rowInst0 = insI >= 0 ? String(data[i][insI]||'').trim() : '';
      if (_isNoHodInstitution_(rowInst0)) {
        throw new Error('HOD access is not applicable for ' + (rowInst0 || 'this institution') +
          '. Faculty submissions go directly to the Head of Institution.');
      }
      // This prevents an HOD from "AVIT → Biotechnology" logging in under
      // "SAS → Biotechnology" even though both row's Department = 'Biotechnology'.
      var rowCampus = camI >= 0 ? String(data[i][camI]||'').trim() : '';
      var rowInst   = insI >= 0 ? String(data[i][insI]||'').trim() : '';
      var have      = _hodResolve(rowCampus, rowInst);
      if (want.campusCode && have.campusCode && want.campusCode !== have.campusCode) {
        throw new Error('This account belongs to a different campus.');
      }
      if (want.instCode && have.instCode && want.instCode !== have.instCode) {
        throw new Error('This account belongs to a different institution.');
      }

      var dept = String(data[i][depI]||'');
      return {
        success:         true,
        role:            'HOD',
        staffID:         String(data[i][idI]||''),
        staffName:       String(data[i][nmI]||'Head of Department'),
        department:      dept,
        campus:          rowCampus,
        institution:     rowInst,
        institutionCode: _resolveInstCode(rowCampus, rowInst),
        email:           ident,
        config:          getConfig()
      };
    }
    throw new Error('HOD account not found. Please contact the administrator.');
  }
}

// ── 3. Faculty Google login ──────────────────────────────────────────────────
function facultyGoogleLogin() {
  var email = Session.getActiveUser().getEmail();
  if (!email) return { success:false, reason:'no_email' };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.FACULTY);
  _ensureFacultyColumns(sheet);
  var data = sheet.getDataRange().getValues(), h = data[0];
  var idI = h.indexOf('FacultyEmail'), nmI = h.indexOf('FacultyName');
  var gmI = h.indexOf('GoogleEmail'), stI = h.indexOf('Status');
  // Campus / Institution / Department columns — needed so the client can show the
  // correct institution logo and department in the topbar after login.
  var caI = h.indexOf('Campus'), inI = h.indexOf('Institution');
  var deI = h.indexOf('Department'), dgI = h.indexOf('Designation');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][gmI]||'').toLowerCase() !== email.toLowerCase()) continue;
    if (String(data[i][stI]||'') !== 'Active') return { success:false, reason:'pending', email:email };
    return {
      success:         true,
      role:            'FACULTY',
      facultyID:       String(data[i][idI]).trim(),
      facultyName:     String(data[i][nmI]||''),
      email:           email,
      campus:          (caI >= 0) ? String(data[i][caI]||'').trim() : '',
      institution:     (inI >= 0) ? String(data[i][inI]||'').trim() : '',
      department:      (deI >= 0) ? String(data[i][deI]||'').trim() : '',
      designation:     (dgI >= 0) ? String(data[i][dgI]||'').trim() : '',
      institutionCode: _resolveInstCode(
                         (caI >= 0) ? String(data[i][caI]||'').trim() : '',
                         (inI >= 0) ? String(data[i][inI]||'').trim() : ''
                       ),
      config:          getConfig()
    };
  }
  return { success:false, reason:'not_found', email:email };
}

// ── 4. Staff Google login ─────────────────────────────────────────────────────
function staffGoogleLogin(role) {
  if (!role) return { success:false, reason:'no_role' };
  role = role.toUpperCase();
  var email = Session.getActiveUser().getEmail();
  if (!email) return { success:false, reason:'no_email' };
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.STAFF);
  if (!sheet) return { success:false, reason:'no_sheet' };
  var data = sheet.getDataRange().getValues(), h = data[0];
  var emI = h.indexOf('Email'), nmI = h.indexOf('StaffName');
  var gmI = h.indexOf('GoogleEmail'), stI = h.indexOf('Status'), rlI = h.indexOf('Role');
  var idI = h.indexOf('StaffID'), insI2 = h.indexOf('Institution');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][gmI]||'').toLowerCase() !== email.toLowerCase()) continue;
    if (String(data[i][rlI]||'').toUpperCase() !== role) continue;
    if (String(data[i][stI]||'') !== 'Active') return { success:false, reason:'pending', email:email };
    // Block HOD Google login for no-HOD institutions
    if (role === 'HOD') {
      var gInst = insI2 >= 0 ? String(data[i][insI2]||'').trim() : '';
      if (_isNoHodInstitution_(gInst)) {
        return { success:false, reason:'no_hod_institution',
          message: 'HOD access is not applicable for ' + (gInst||'this institution') +
            '. Faculty submissions go directly to the Head of Institution.' };
      }
    }
    return { success:true, role:role, staffID:String(data[i][idI]).trim(), staffName:String(data[i][nmI]||''), email:email };
  }
  return { success:false, reason:'not_found', email:email };
}

// ── 5. Faculty self-registration — auto-generates Faculty ID ─────────────────
// Faculty fill in their details and set a password. No pre-enrollment needed.
// Faculty ID (VMRF-XXXXXX) is generated automatically and returned to the user.
// ─── EMAIL OTP VERIFICATION ───────────────────────────────────────────────────
function sendEmailOTP(email) {
  if (!email) throw new Error('Email address is required.');
  email = String(email).trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Invalid email format.');

  // Check email not already registered
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var facSheet = ss.getSheetByName(SH.FACULTY);
  if (facSheet) {
    var fd = facSheet.getDataRange().getValues(), fh = fd[0];
    var femI = fh.indexOf('Email');
    for (var i = 1; i < fd.length; i++) {
      if (String(fd[i][femI]||'').trim().toLowerCase() === email)
        throw new Error('This email is already registered. Please sign in instead.');
    }
  }
  var stSheet = ss.getSheetByName(SH.STAFF);
  if (stSheet) {
    var sd = stSheet.getDataRange().getValues(), sh = sd[0];
    var semI = sh.indexOf('Email');
    for (var j = 1; j < sd.length; j++) {
      if (String(sd[j][semI]||'').trim().toLowerCase() === email)
        throw new Error('This email is already registered. Please sign in instead.');
    }
  }

  // Generate 6-digit OTP
  var otp = String(Math.floor(100000 + Math.random() * 900000));
  var expiry = Date.now() + (10 * 60 * 1000);
  PropertiesService.getScriptProperties().setProperty('OTP_' + email, JSON.stringify({ otp: otp, expiry: expiry }));

  // Send email with sender name to avoid spam
  MailApp.sendEmail({
    to: email,
    name: 'VMRF-DU Institutional Monitoring System',
    subject: 'Your Email Verification Code - VMRF-DU',
    htmlBody:
      '<div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto;padding:32px 24px;border:1px solid #e0e0e0;border-radius:8px">' +
        '<h2 style="color:#0e1f38;margin:0 0 8px">Email Verification</h2>' +
        '<p style="color:#555;font-size:14px;margin:0 0 24px">Use the code below to verify your email for the <strong>VMRF-DU Institutional Monitoring System</strong>.</p>' +
        '<div style="background:#f7f3ed;border:2px solid #c9a84c;border-radius:8px;padding:20px;text-align:center;margin-bottom:24px">' +
          '<div style="font-size:11px;font-weight:700;color:#666;letter-spacing:1px;text-transform:uppercase;margin-bottom:8px">Verification Code</div>' +
          '<div style="font-family:monospace;font-size:36px;font-weight:700;color:#0e1f38;letter-spacing:6px">' + otp + '</div>' +
          '<div style="font-size:11px;color:#999;margin-top:8px">Expires in 10 minutes</div>' +
        '</div>' +
        '<p style="color:#888;font-size:12px;margin:0">If you did not request this, please ignore this email.</p>' +
        '<hr style="border:none;border-top:1px solid #eee;margin:20px 0">' +
        '<p style="color:#aaa;font-size:11px;margin:0">VMRF-DU Institutional Management Office</p>' +
      '</div>'
  });
  return { success: true };
}

function verifyEmailOTP(email, otp) {
  if (!email || !otp) throw new Error('Email and OTP are required.');
  email = String(email).trim().toLowerCase();
  otp   = String(otp).trim();
  var props = PropertiesService.getScriptProperties();
  var stored = props.getProperty('OTP_' + email);
  if (!stored) throw new Error('No verification code found. Please request a new one.');
  var data = JSON.parse(stored);
  if (Date.now() > data.expiry) {
    props.deleteProperty('OTP_' + email);
    throw new Error('Verification code has expired. Please request a new one.');
  }
  if (data.otp !== otp) throw new Error('Incorrect verification code. Please try again.');
  props.setProperty('OTP_VERIFIED_' + email, 'yes');
  return { success: true };
}

function facultyRegister(f) {
  var _fe = String(f.email||'').trim().toLowerCase();
  var _fp = PropertiesService.getScriptProperties();
  if (_fp.getProperty('OTP_VERIFIED_' + _fe) !== 'yes')
    throw new Error('Email not verified. Please verify your email with the OTP before registering.');
  _fp.deleteProperty('OTP_VERIFIED_' + _fe);
  _fp.deleteProperty('OTP_' + _fe);
  if (!f.name)     throw new Error('Full name is required.');
  if (!f.email)    throw new Error('Email address is required.');
  if (!f.department)  throw new Error('Please select a department.');
  if (!f.designation) throw new Error('Please select a designation.');
  if (!f.campus)      throw new Error('Please select a campus.');
  if (!f.institution) throw new Error('Please select an institution.');
  if (!f.password)    throw new Error('Please set a password.');
  if (f.password.length < 6) throw new Error('Password must be at least 6 characters.');
  if (f.password !== f.confirmPassword) throw new Error('Passwords do not match.');

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.FACULTY);
  if (!sheet) { _initSheet(SH.FACULTY); sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.FACULTY); }
  var h    = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0].map(function(v){return String(v).trim();});
  var data = sheet.getDataRange().getValues();

  // Email uniqueness check
  var emI = h.indexOf('Email');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][emI]||'').toLowerCase() === f.email.toLowerCase())
      throw new Error('An account with this email already exists. Please sign in.');
  }

  // Write row by column name (no FacultyID — Email is the identifier)
  var vals = {
    'FacultyName': f.name, 'Email': f.email,
    'Department': f.department, 'Campus': f.campus, 'Institution': f.institution,
    'Designation': f.designation, 'PasswordHash': _hashPwd(f.password),
    'GoogleEmail': f.googleEmail || '', 'Status': 'Active'
  };
  var newRow = h.map(function(col){ return vals[col] !== undefined ? vals[col] : ''; });
  sheet.appendRow(newRow);
  return { success:true, facultyName:f.name };
}

// ── 6. Staff self-registration (HOD / HOI) ───────────────────────────────────
// Status is set to 'Pending' — admin approves in Staff_Master sheet.
// Email is cross-checked against Faculty_Master to prevent faculty impersonating staff.
function staffRegister(f) {
  var _se = String(f.email||'').trim().toLowerCase();
  var _sp = PropertiesService.getScriptProperties();
  if (_sp.getProperty('OTP_VERIFIED_' + _se) !== 'yes')
    throw new Error('Email not verified. Please verify your email with the OTP before registering.');
  _sp.deleteProperty('OTP_VERIFIED_' + _se);
  _sp.deleteProperty('OTP_' + _se);
  if (!f.role)     throw new Error('Role is required.');
  if (!f.name)     throw new Error('Full name is required.');
  if (!f.email)    throw new Error('Email address is required.');
  if (!f.password) throw new Error('Please set a password.');
  if (f.password.length < 6) throw new Error('Password must be at least 6 characters.');
  if (f.password !== f.confirmPassword) throw new Error('Passwords do not match.');
  var role = f.role.toUpperCase();
  if (['HOD','HOI'].indexOf(role) < 0) throw new Error('Only HOD and HOI can self-register. IMO credentials are managed by the system administrator.');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var emailLower = f.email.toLowerCase().trim();

  // ── BLOCK: email already registered as a Faculty member ──
  var facSheet = ss.getSheetByName(SH.FACULTY);
  if (facSheet) {
    var facData = facSheet.getDataRange().getValues(), fh = facData[0];
    var femI = fh.indexOf('Email');
    for (var fi = 1; fi < facData.length; fi++) {
      if (String(facData[fi][femI]||'').toLowerCase().trim() === emailLower)
        throw new Error('This email is already registered as a Faculty member and cannot be used to register as ' + role + '. Please use a different email address.');
    }
  }

  // ── BLOCK: duplicate HOD/HOI entry for same email ──
  var sheet = ss.getSheetByName(SH.STAFF);
  if (!sheet) throw new Error('Staff sheet not found. Please run initializeSystem() first.');
  var data = sheet.getDataRange().getValues(), h = data[0];
  var emI = h.indexOf('Email'), rlI = h.indexOf('Role');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][emI]||'').toLowerCase().trim() === emailLower &&
        String(data[i][rlI]||'').toUpperCase() === role)
      throw new Error('An account with this email already exists for ' + role + '. Please sign in.');
  }

  var existing = data.slice(1).map(function(r){ return String(r[0]); });
  var id = _makeID(role);
  while (existing.indexOf(id) !== -1) id = _makeID(role);
  var vals = {
    'StaffID':id, 'StaffName':f.name, 'Email':f.email,
    'Role':role,  'Department':f.department||'',
    'PasswordHash':_hashPwd(f.password), 'GoogleEmail':'', 'Status':'Pending'
  };
  var newRow = h.map(function(col){ return vals[col]!==undefined?vals[col]:''; });
  sheet.appendRow(newRow);
  return { success:true, staffID:id, staffName:f.name, role:role };
}

// ── 7. Change password ───────────────────────────────────────────────────────
// ─── FORGOT PASSWORD — RESET VIA EMAIL OTP ───────────────────────────────────



// ── Forgot Password: send OTP ─────────────────────────────────────────────────
function sendPasswordResetOTP(email) {
  if (!email) throw new Error('Email address is required.');
  email = String(email).trim().toLowerCase();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var found = false;

  var facSheet = ss.getSheetByName(SH.FACULTY);
  if (facSheet) {
    var fd = facSheet.getDataRange().getValues(), fh = fd[0];
    var femI = fh.indexOf('Email');
    for (var i = 1; i < fd.length; i++) {
      if (String(fd[i][femI]||'').trim().toLowerCase() === email) { found = true; break; }
    }
  }
  if (!found) {
    var stSheet = ss.getSheetByName(SH.STAFF);
    if (stSheet) {
      var sd = stSheet.getDataRange().getValues(), sh = sd[0];
      var semI = sh.indexOf('Email');
      for (var j = 1; j < sd.length; j++) {
        if (String(sd[j][semI]||'').trim().toLowerCase() === email) { found = true; break; }
      }
    }
  }
  if (!found) throw new Error('No account found with this email address.');

  var otp = String(Math.floor(100000 + Math.random() * 900000));
  var expiry = Date.now() + (10 * 60 * 1000);
  PropertiesService.getScriptProperties().setProperty('RESET_OTP_' + email, JSON.stringify({ otp: otp, expiry: expiry }));

  MailApp.sendEmail({
    to: email,
    name: 'VMRF-DU Institutional Monitoring System',
    subject: 'Password Reset Code - VMRF-DU',
    htmlBody:
      '<div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto;padding:32px 24px;border:1px solid #e0e0e0;border-radius:8px">' +
        '<h2 style="color:#0e1f38;margin:0 0 8px">Password Reset</h2>' +
        '<p style="color:#555;font-size:14px;margin:0 0 24px">Use the code below to reset your password for the <strong>VMRF-DU Institutional Monitoring System</strong>.</p>' +
        '<div style="background:#f7f3ed;border:2px solid #c9a84c;border-radius:8px;padding:20px;text-align:center;margin-bottom:24px">' +
          '<div style="font-size:11px;font-weight:700;color:#666;letter-spacing:1px;text-transform:uppercase;margin-bottom:8px">Reset Code</div>' +
          '<div style="font-family:monospace;font-size:36px;font-weight:700;color:#0e1f38;letter-spacing:6px">' + otp + '</div>' +
          '<div style="font-size:11px;color:#999;margin-top:8px">Expires in 10 minutes</div>' +
        '</div>' +
        '<p style="color:#888;font-size:12px;margin:0">If you did not request this, please ignore this email.</p>' +
        '<hr style="border:none;border-top:1px solid #eee;margin:20px 0">' +
        '<p style="color:#aaa;font-size:11px;margin:0">VMRF-DU Institutional Management Office</p>' +
      '</div>'
  });
  return { success: true };
}

// ── Forgot Password: reset with OTP ──────────────────────────────────────────
function resetPasswordWithOTP(email, otp, newPwd) {
  if (!email || !otp || !newPwd) throw new Error('All fields are required.');
  email = String(email).trim().toLowerCase();
  otp   = String(otp).trim();
  if (newPwd.length < 6) throw new Error('Password must be at least 6 characters.');

  var props = PropertiesService.getScriptProperties();
  var stored = props.getProperty('RESET_OTP_' + email);
  if (!stored) throw new Error('No reset code found. Please request a new one.');
  var data = JSON.parse(stored);
  if (Date.now() > data.expiry) {
    props.deleteProperty('RESET_OTP_' + email);
    throw new Error('Reset code has expired. Please request a new one.');
  }
  if (data.otp !== otp) throw new Error('Incorrect reset code. Please try again.');

  var hash = _hashPwd(newPwd);
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var facSheet = ss.getSheetByName(SH.FACULTY);
  if (facSheet) {
    var fd = facSheet.getDataRange().getValues(), fh = fd[0];
    var femI = fh.indexOf('Email'), fpwI = fh.indexOf('PasswordHash');
    for (var i = 1; i < fd.length; i++) {
      if (String(fd[i][femI]||'').trim().toLowerCase() === email) {
        facSheet.getRange(i+1, fpwI+1).setValue(hash);
        props.deleteProperty('RESET_OTP_' + email);
        return { success: true };
      }
    }
  }
  var stSheet = ss.getSheetByName(SH.STAFF);
  if (stSheet) {
    var sd = stSheet.getDataRange().getValues(), sh = sd[0];
    var semI = sh.indexOf('Email'), spwI = sh.indexOf('PasswordHash');
    for (var j = 1; j < sd.length; j++) {
      if (String(sd[j][semI]||'').trim().toLowerCase() === email) {
        stSheet.getRange(j+1, spwI+1).setValue(hash);
        props.deleteProperty('RESET_OTP_' + email);
        return { success: true };
      }
    }
  }
  throw new Error('Account not found.');
}

function changePassword(role, identifier, oldPwd, newPwd) {
  if (!newPwd || newPwd.length < 6) throw new Error('New password must be at least 6 characters.');
  role = String(role).toUpperCase();
  if (role === 'FACULTY') {
    facultyLogin(identifier, oldPwd); // throws if wrong (identifier = email)
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.FACULTY);
    _ensureFacultyColumns(sheet);
    var data = sheet.getDataRange().getValues(), h = data[0];
    var emI = h.indexOf('Email'), pwI = h.indexOf('PasswordHash');
    var emailLc = String(identifier).trim().toLowerCase();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][emI]||'').trim().toLowerCase() === emailLc) {
        sheet.getRange(i+1, pwI+1).setValue(_hashPwd(newPwd));
        return { ok:true };
      }
    }
    throw new Error('Faculty record not found.');
  } else {
    staffLogin(role, identifier, oldPwd); // throws if wrong (identifier = email OR IPM username)
    var newHash  = _hashPwd(newPwd);
    var updatedMain = false;
    var sh2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.STAFF);
    if (sh2) {
      var d2 = sh2.getDataRange().getValues(), h2 = d2[0];
      var emI2 = h2.indexOf('Email'), pwI2 = h2.indexOf('PasswordHash'), rlI2 = h2.indexOf('Role');
      for (var j = 1; j < d2.length; j++) {
        if (String(d2[j][emI2]||'').toLowerCase() === String(identifier||'').toLowerCase() &&
            String(d2[j][rlI2]||'').toUpperCase() === role) {
          sh2.getRange(j+1, pwI2+1).setValue(newHash);
          updatedMain = true;
          break;
        }
      }
    }

    // For HOI/IMO accounts also mirror the new password into IPM Users.
    // staffLogin() authenticates these roles against the IPM Users sheet via
    // _tryIpmLoginForMain_, so without updating IPM Users the user would
    // silently keep the old password for every IPM-backed login attempt.
    var updatedIpm = false;
    if (role === 'HOI' || role === 'IMO') {
      try {
        var usersSh = ipmSheet_('Users');
        if (usersSh) {
          var iD = usersSh.getDataRange().getValues(), iH = iD[0];
          var uI = iH.indexOf('username'), piI = iH.indexOf('password'), riI = iH.indexOf('role');
          var want = String(identifier||'').trim().toLowerCase();
          for (var k = 1; k < iD.length; k++) {
            if (String(iD[k][uI]||'').toLowerCase() === want &&
                String(iD[k][riI]||'').toUpperCase() === role) {
              usersSh.getRange(k+1, piI+1).setValue(ipmHashPassword_(newPwd));
              updatedIpm = true;
              break;
            }
          }
        }
      } catch (ipmErr) {
        // IPM sheet unavailable — surface the error only if we also failed
        // to update the main sheet (otherwise the main update carries us).
        if (!updatedMain) throw ipmErr;
      }
    }

    if (updatedMain || updatedIpm) return { ok:true };
    throw new Error('Staff record not found.');
  }
}

// ── 8. IMO enrolls faculty (unchanged) ───────────────────────────────────────
function preEnrollFaculty(f) {
  if (!f.name||!f.department||!f.designation||!f.campus||!f.institution)
    throw new Error('All fields are required.');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.FACULTY);
  _ensureFacultyColumns(sheet);
  var h    = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var data = sheet.getDataRange().getValues();
  if (f.email) {
    var emI = h.indexOf('Email');
    for (var j = 1; j < data.length; j++) {
      if (String(data[j][emI]||'').toLowerCase() === f.email.toLowerCase())
        throw new Error('A faculty with this email already exists.');
    }
  }
  var existing = data.slice(1).map(function(r){ return String(r[0]); });
  var id = _makeID('VMRF');
  while (existing.indexOf(id) !== -1) id = _makeID('VMRF');
  var vals = {
    'FacultyName':f.name,'Email':f.email||'',
    'Department':f.department,'Campus':f.campus,'Institution':f.institution,
    'Designation':f.designation,'PasswordHash':'','GoogleEmail':'','Status':'Pending'
  };
  var newRow = h.map(function(col){ return vals[col]!==undefined?vals[col]:''; });
  sheet.appendRow(newRow);
  return { id:id, name:f.name };
}

function _ensureFacultyColumns(sheet) {
  if (!sheet) return;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var required = SCHEMA[SH.FACULTY];
  for (var i = 0; i < required.length; i++) {
    if (headers.indexOf(required[i]) === -1) {
      // Append missing column header at the end
      var newCol = headers.length + 1;
      sheet.getRange(1, newCol).setValue(required[i]);
      headers.push(required[i]);
    }
  }
}

// ─── CONFIG ───────────────────────────────────────────────────────────────────

// ─── SEED HOD PRE-ACCOUNTS INTO STAFF_MASTER ─────────────────────────────────
// Run this once from the VMRF menu to create all department HOD accounts.
// Each HOD has a unique password — see the alert after seeding for the full list.
/* ═══════════════════════════════════════════════════════════════════════════
 * seedHODAccounts  —  generate one HOD account per
 * (campus × institution × department) triple in INSTITUTION_HIERARCHY.
 *
 * For each triple, a username and random 12-char password are generated,
 * and the account is written to BOTH:
 *   - Staff_Master (so HODs can sign in through the main app's HOD login)
 *   - IPM Users    (so HOD can eventually be granted IPM SSO if needed)
 *
 * All credentials are written to a CSV file in Drive (named
 * "VMRF_HOD_Credentials_<timestamp>.csv") and a link is shown in the alert.
 * Nothing is logged to the sheet UI except the count, because 75+ lines of
 * credentials in a dialog are unusable.
 *
 * Idempotent: re-running preserves existing usernames. A row whose username
 * already exists gets its password reset. Rows orphaned by a hierarchy
 * change (institution renamed, department removed) are left in place with
 * Status=Inactive so nothing is lost.
 * ═══════════════════════════════════════════════════════════════════════════ */
function seedHODAccounts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH.STAFF);
  if (!sh) { try { SpreadsheetApp.getUi().alert('Staff_Master sheet not found. Run initializeSystem first.'); } catch(_){} return; }

  // Ensure Staff_Master has the new Campus + Institution columns
  _ensureSheetColumns(sh, SCHEMA.Staff_Master);

  var planned = _buildHodAccountPlan_();   // array of { username, campus, institution, department, campusCode, instCode, fullName }
  var ipmSs    = ipmDb_();                 // Drive-hosted IPM spreadsheet (separate from main sheet)
  var ipmUsers = ipmSs.getSheetByName('Users');

  // Index existing rows in both sheets by username (case-insensitive)
  var sData = sh.getDataRange().getValues(), sH = sData[0];
  var sColIdx = {
    StaffID:     sH.indexOf('StaffID'),
    StaffName:   sH.indexOf('StaffName'),
    Email:       sH.indexOf('Email'),
    Role:        sH.indexOf('Role'),
    Department:  sH.indexOf('Department'),
    Campus:      sH.indexOf('Campus'),
    Institution: sH.indexOf('Institution'),
    PasswordHash:sH.indexOf('PasswordHash'),
    GoogleEmail: sH.indexOf('GoogleEmail'),
    Status:      sH.indexOf('Status')
  };
  var existingStaffIdx = {}; // lowercased username → row number (1-based)
  for (var r = 1; r < sData.length; r++) {
    if (String(sData[r][sColIdx.Role]||'').toUpperCase() !== 'HOD') continue;
    var u = String(sData[r][sColIdx.Email]||'').trim().toLowerCase();
    if (u) existingStaffIdx[u] = r + 1; // Sheet rows are 1-based
  }

  var iData = ipmUsers.getDataRange().getValues(), iH = iData[0];
  var iColIdx = {
    username:    iH.indexOf('username'),
    password:    iH.indexOf('password'),
    role:        iH.indexOf('role'),
    campus:      iH.indexOf('campus'),
    institution: iH.indexOf('institution'),
    fullName:    iH.indexOf('fullName'),
    createdAt:   iH.indexOf('createdAt')
  };
  var existingIpmIdx = {};
  for (var r2 = 1; r2 < iData.length; r2++) {
    if (String(iData[r2][iColIdx.role]||'').toUpperCase() !== 'HOD') continue;
    var u2 = String(iData[r2][iColIdx.username]||'').trim().toLowerCase();
    if (u2) existingIpmIdx[u2] = r2 + 1;
  }

  var now = new Date();
  var added = 0, updated = 0;
  var csvRows = [['Username','Password','Full Name','Campus Code','Campus Name','Institution Code','Institution Name','Department','StaffID']];

  planned.forEach(function(acc) {
    var ukey = acc.username.toLowerCase();
    var plainPw = _genHodPassword_();
    var pwHash  = _hashPwd(plainPw);

    // ── Staff_Master row
    if (existingStaffIdx[ukey]) {
      // Row exists → reset password, update name/campus/inst/department/status
      var row = existingStaffIdx[ukey];
      sh.getRange(row, sColIdx.StaffName+1).setValue(acc.fullName);
      sh.getRange(row, sColIdx.Department+1).setValue(acc.department);
      if (sColIdx.Campus      >= 0) sh.getRange(row, sColIdx.Campus+1).setValue(acc.campus);
      if (sColIdx.Institution >= 0) sh.getRange(row, sColIdx.Institution+1).setValue(acc.institution);
      sh.getRange(row, sColIdx.PasswordHash+1).setValue(pwHash);
      sh.getRange(row, sColIdx.Status+1).setValue('Active');
      // Preserve StaffID + GoogleEmail as-is
      var rowSid = String(sData[row-1][sColIdx.StaffID]||'');
      updated++;
      csvRows.push([acc.username, plainPw, acc.fullName, acc.campusCode, acc.campus, acc.instCode, acc.institution, acc.department, rowSid]);
    } else {
      // New row — assign StaffID from current sheet length
      var newSid = 'HOD-' + Utilities.formatString('%04d', sh.getLastRow() + 1);
      var rowArr = new Array(sH.length);
      rowArr[sColIdx.StaffID]      = newSid;
      rowArr[sColIdx.StaffName]    = acc.fullName;
      rowArr[sColIdx.Email]        = acc.username; // username stored in Email column (kept name for back-compat)
      rowArr[sColIdx.Role]         = 'HOD';
      rowArr[sColIdx.Department]   = acc.department;
      if (sColIdx.Campus      >= 0) rowArr[sColIdx.Campus]      = acc.campus;
      if (sColIdx.Institution >= 0) rowArr[sColIdx.Institution] = acc.institution;
      rowArr[sColIdx.PasswordHash] = pwHash;
      if (sColIdx.GoogleEmail >= 0) rowArr[sColIdx.GoogleEmail] = '';
      rowArr[sColIdx.Status]       = 'Active';
      sh.appendRow(rowArr);
      added++;
      csvRows.push([acc.username, plainPw, acc.fullName, acc.campusCode, acc.campus, acc.instCode, acc.institution, acc.department, newSid]);
    }

    // ── IPM Users row (uses campus CODE + institution CODE to match HOI rows)
    var ipmPwHash = ipmHashPassword_(plainPw);
    if (existingIpmIdx[ukey]) {
      var ipmRow = existingIpmIdx[ukey];
      ipmUsers.getRange(ipmRow, iColIdx.password+1).setValue(ipmPwHash);
      ipmUsers.getRange(ipmRow, iColIdx.campus+1).setValue(acc.campusCode);
      ipmUsers.getRange(ipmRow, iColIdx.institution+1).setValue(acc.instCode);
      ipmUsers.getRange(ipmRow, iColIdx.fullName+1).setValue(acc.fullName);
    } else {
      var ipmRowArr = new Array(iH.length);
      ipmRowArr[iColIdx.username]    = acc.username;
      ipmRowArr[iColIdx.password]    = ipmPwHash;
      ipmRowArr[iColIdx.role]        = 'HOD';
      ipmRowArr[iColIdx.campus]      = acc.campusCode;
      ipmRowArr[iColIdx.institution] = acc.instCode;
      ipmRowArr[iColIdx.fullName]    = acc.fullName;
      ipmRowArr[iColIdx.createdAt]   = now;
      ipmUsers.appendRow(ipmRowArr);
    }
  });

  // ── Write credentials CSV to Drive
  var csvContent = csvRows.map(function(r){
    return r.map(function(cell){
      var s = String(cell == null ? '' : cell);
      return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    }).join(',');
  }).join('\n');

  var fileName = 'VMRF_HOD_Credentials_' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss') + '.csv';
  var fileUrl  = '';
  try {
    var file = DriveApp.createFile(fileName, csvContent, MimeType.CSV);
    fileUrl  = file.getUrl();
  } catch(e) {
    fileUrl = '(CSV creation failed: ' + e.message + ')';
  }

  try {
    SpreadsheetApp.getUi().alert(
      '✅ HOD Accounts Seeded\n\n' +
      'Total planned:   ' + planned.length + '\n' +
      'New accounts:    ' + added + '\n' +
      'Updated accounts: ' + updated + '\n\n' +
      'Credentials CSV saved to Drive:\n' + fileName + '\n\n' +
      (fileUrl ? 'Open: ' + fileUrl : '') + '\n\n' +
      '⚠️ Open the CSV and distribute credentials securely — HODs should change their password after first login.'
    );
  } catch(_){}
}

/* Build the full list of HOD account specs from INSTITUTION_HIERARCHY.
   Each spec: { username, campus, institution, department, campusCode, instCode, fullName } */

// ─── REMOVE HOD ACCOUNTS FOR NO-HOD INSTITUTIONS ─────────────────────────────
// Run this ONCE from the VMRF menu to permanently delete HOD rows from
// Staff_Master (and IPM Users) for SAHS, VSEP, VSHS, VMLS.
// Safe to re-run — already-deleted rows are simply skipped.
function removeNoHodAccounts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH.STAFF);
  if (!sh) { try { SpreadsheetApp.getUi().alert('Staff_Master sheet not found.'); } catch(_){} return; }

  var data = sh.getDataRange().getValues();
  var h    = data[0];
  var rlI  = h.indexOf('Role'), insI = h.indexOf('Institution');
  var emI  = h.indexOf('Email');

  if (rlI < 0 || insI < 0) {
    try { SpreadsheetApp.getUi().alert('Required columns (Role / Institution) not found in Staff_Master.'); } catch(_){}
    return;
  }

  // Collect row indices to delete (bottom-up so row numbers stay valid)
  var rowsToDelete = [];
  var deletedUsers = [];
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][rlI]||'').toUpperCase() !== 'HOD') continue;
    var inst = String(data[i][insI]||'').trim();
    if (!_isNoHodInstitution_(inst)) continue;
    rowsToDelete.push(i + 1); // sheet row is 1-based
    deletedUsers.push(String(data[i][emI]||'').trim());
  }

  if (rowsToDelete.length === 0) {
    try { SpreadsheetApp.getUi().alert('No HOD accounts found for no-HOD institutions. Nothing to remove.'); } catch(_){}
    return;
  }

  // Delete from Staff_Master (already bottom-up ordered)
  rowsToDelete.forEach(function(rowNum){ sh.deleteRow(rowNum); });

  // Delete matching rows from IPM Users sheet
  try {
    var ipmSs    = ipmDb_();
    var ipmUsers = ipmSs.getSheetByName('Users');
    if (ipmUsers) {
      var iData = ipmUsers.getDataRange().getValues();
      var iH    = iData[0];
      var iUnI  = iH.indexOf('username'), iRlI = iH.indexOf('role');
      var ipmRowsToDelete = [];
      for (var j = iData.length - 1; j >= 1; j--) {
        if (String(iData[j][iRlI]||'').toUpperCase() !== 'HOD') continue;
        var uname = String(iData[j][iUnI]||'').trim().toLowerCase();
        // Match by username prefix — no-HOD institution codes embedded in username
        var isNoHodUser = deletedUsers.some(function(du){ return du.toLowerCase() === uname; }) ||
          ['sahs','vsep','vshs','vmls'].some(function(code){ return uname.indexOf('hod_'+code) === 0; });
        if (isNoHodUser) ipmRowsToDelete.push(j + 1);
      }
      ipmRowsToDelete.forEach(function(rowNum){ ipmUsers.deleteRow(rowNum); });
    }
  } catch(e) { Logger.log('IPM cleanup skipped: ' + e.message); }

  try {
    SpreadsheetApp.getUi().alert(
      '✅ Removed ' + rowsToDelete.length + ' HOD account(s) from Staff_Master' +
      ' for no-HOD institutions (SAHS, VSEP, VSHS, VMLS).\n\n' +
      'Deleted usernames:\n' + deletedUsers.join('\n')
    );
  } catch(_){}
}

function _buildHodAccountPlan_() {
  var plan = [];
  var seenUsernames = {};
  Object.keys(INSTITUTION_HIERARCHY || {}).forEach(function(campusName){
    var campus = INSTITUTION_HIERARCHY[campusName];
    var campusCode = campus.code || '';
    Object.keys(campus.institutions || {}).forEach(function(instName){
      var inst = campus.institutions[instName];
      var instCode = inst.code || '';
      // Skip institutions that have no HOD layer — faculty report directly to HOI
      if (inst.noHod) return;
      (inst.departments || []).forEach(function(deptName){
        var base = 'hod_' + String(instCode).toLowerCase().replace(/[^a-z0-9]+/g,'') +
                   '_'   + _hodDeptSlug_(deptName);
        var username = base;
        var suffix = 2;
        while (seenUsernames[username]) { username = base + suffix; suffix++; }
        seenUsernames[username] = true;
        plan.push({
          username:    username,
          campus:      campusName,
          institution: instName,
          department:  deptName,
          campusCode:  campusCode,
          instCode:    instCode,
          fullName:    'HOD - ' + deptName + ' (' + instCode + ')'
        });
      });
    });
  });
  return plan;
}

/* Short, stable slug for a department name. Examples:
     "Computer Science and Engineering" → "cse"
     "Humanities & Sciences"            → "hs"
     "Obstetrics & Gynaecology"         → "og"
     "Medical Laboratory Technology"    → "mlt"
   For single-word departments, use the first 6 letters:
     "Biotechnology" → "biotec"
     "Law"           → "law"                                                 */
function _hodDeptSlug_(deptName) {
  var clean = String(deptName||'').replace(/[^A-Za-z0-9 ]+/g,' ').trim();
  var words = clean.split(/\s+/).filter(function(w){
    // Drop filler words
    return w && !/^(and|the|of|for|in|on|&)$/i.test(w);
  });
  if (words.length >= 2) {
    return words.map(function(w){ return w.charAt(0).toLowerCase(); }).join('').slice(0, 8);
  }
  return (words[0] || 'dept').toLowerCase().slice(0, 6);
}

/* Cryptographically-random 12-char password. Mix of upper/lower/digit/symbol. */
function _genHodPassword_() {
  var UP = 'ABCDEFGHJKMNPQRSTUVWXYZ';       // Exclude I,L,O for legibility
  var LO = 'abcdefghijkmnpqrstuvwxyz';      // Exclude l,o
  var DG = '23456789';                      // Exclude 0,1
  var SY = '@#$%&*';
  var all = UP + LO + DG + SY;
  // Build with guaranteed diversity then shuffle
  var bytes = Utilities.getUuid().replace(/-/g,''); // 32 hex chars — source of entropy
  var out = [
    UP.charAt(parseInt(bytes.substr(0,2),16) % UP.length),
    LO.charAt(parseInt(bytes.substr(2,2),16) % LO.length),
    DG.charAt(parseInt(bytes.substr(4,2),16) % DG.length),
    SY.charAt(parseInt(bytes.substr(6,2),16) % SY.length)
  ];
  for (var i = 0; i < 8; i++) {
    out.push(all.charAt(parseInt(bytes.substr(8 + i*2, 2), 16) % all.length));
  }
  // Fisher–Yates using a second UUID for shuffle entropy
  var sh = Utilities.getUuid().replace(/-/g,'');
  for (var j = out.length - 1; j > 0; j--) {
    var k = parseInt(sh.substr((j*2) % sh.length, 2), 16) % (j+1);
    var tmp = out[j]; out[j] = out[k]; out[k] = tmp;
  }
  return out.join('');
}

/* ═══════════════════════════════════════════════════════════════════════════
 * HOI ACCOUNT SEEDING
 * ═══════════════════════════════════════════════════════════════════════════
 * seedHOIAccounts() — editor-runnable. Creates one HOI account per
 * institution across INSTITUTION_HIERARCHY, writes to both Staff_Master
 * (main-app auth) and the IPM Users sheet (IPM sub-portal auth), and
 * exports a CSV of credentials to Drive root.
 *
 * Idempotent: re-running updates existing rows (new password + name sync)
 * rather than duplicating. Username format: hoi_<instcode> (lowercased).
 * Total accounts: one per institution (12 across VMCC + VMPC).
 *
 * Pair function resetHOIPasswords() rotates passwords for existing HOI
 * rows without touching other account fields.
 *
 * After running, open the CSV from Drive and distribute the credentials
 * securely. HOIs can change their password after first sign-in. */
function seedHOIAccounts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH.STAFF);
  if (!sh) { try { SpreadsheetApp.getUi().alert('Staff_Master sheet not found. Run initializeSystem first.'); } catch(_){} return; }

  // Ensure Staff_Master has the Campus + Institution columns
  _ensureSheetColumns(sh, SCHEMA.Staff_Master);

  var planned = _buildHoiAccountPlan_();   // [{ username, campus, institution, campusCode, instCode, fullName }]
  var ipmSs    = ipmDb_();                 // Drive-hosted IPM spreadsheet
  var ipmUsers = ipmSs.getSheetByName('Users');
  if (!ipmUsers) { try { SpreadsheetApp.getUi().alert('IPM Users sheet not found. Run ipmSetup first.'); } catch(_){} return; }

  // Index existing rows by username (case-insensitive)
  var sData = sh.getDataRange().getValues(), sH = sData[0];
  var sColIdx = {
    StaffID:     sH.indexOf('StaffID'),
    StaffName:   sH.indexOf('StaffName'),
    Email:       sH.indexOf('Email'),
    Role:        sH.indexOf('Role'),
    Department:  sH.indexOf('Department'),
    Campus:      sH.indexOf('Campus'),
    Institution: sH.indexOf('Institution'),
    PasswordHash:sH.indexOf('PasswordHash'),
    GoogleEmail: sH.indexOf('GoogleEmail'),
    Status:      sH.indexOf('Status')
  };
  var existingStaffIdx = {}; // lowercased username → row number (1-based)
  for (var r = 1; r < sData.length; r++) {
    if (String(sData[r][sColIdx.Role]||'').toUpperCase() !== 'HOI') continue;
    var u = String(sData[r][sColIdx.Email]||'').trim().toLowerCase();
    if (u) existingStaffIdx[u] = r + 1;
  }

  var iData = ipmUsers.getDataRange().getValues(), iH = iData[0];
  var iColIdx = {
    username:    iH.indexOf('username'),
    password:    iH.indexOf('password'),
    role:        iH.indexOf('role'),
    campus:      iH.indexOf('campus'),
    institution: iH.indexOf('institution'),
    fullName:    iH.indexOf('fullName'),
    createdAt:   iH.indexOf('createdAt')
  };
  var existingIpmIdx = {};
  for (var r2 = 1; r2 < iData.length; r2++) {
    if (String(iData[r2][iColIdx.role]||'').toUpperCase() !== 'HOI') continue;
    var u2 = String(iData[r2][iColIdx.username]||'').trim().toLowerCase();
    if (u2) existingIpmIdx[u2] = r2 + 1;
  }

  var now = new Date();
  var added = 0, updated = 0;
  var csvRows = [['Username','Password','Full Name','Campus Code','Campus Name','Institution Code','Institution Name','StaffID']];

  planned.forEach(function(acc) {
    var ukey = acc.username.toLowerCase();
    var plainPw = _genHodPassword_();       // reuse the proven 12-char generator
    var pwHash  = _hashPwd(plainPw);

    // ── Staff_Master row
    if (existingStaffIdx[ukey]) {
      var row = existingStaffIdx[ukey];
      sh.getRange(row, sColIdx.StaffName+1).setValue(acc.fullName);
      if (sColIdx.Campus      >= 0) sh.getRange(row, sColIdx.Campus+1).setValue(acc.campus);
      if (sColIdx.Institution >= 0) sh.getRange(row, sColIdx.Institution+1).setValue(acc.institution);
      sh.getRange(row, sColIdx.PasswordHash+1).setValue(pwHash);
      sh.getRange(row, sColIdx.Status+1).setValue('Active');
      var rowSid = String(sData[row-1][sColIdx.StaffID]||'');
      updated++;
      csvRows.push([acc.username, plainPw, acc.fullName, acc.campusCode, acc.campus, acc.instCode, acc.institution, rowSid]);
    } else {
      var newSid = 'HOI-' + Utilities.formatString('%04d', sh.getLastRow() + 1);
      var rowArr = new Array(sH.length);
      rowArr[sColIdx.StaffID]      = newSid;
      rowArr[sColIdx.StaffName]    = acc.fullName;
      rowArr[sColIdx.Email]        = acc.username;
      rowArr[sColIdx.Role]         = 'HOI';
      rowArr[sColIdx.Department]   = '';   // HOI is institution-wide, no single department
      if (sColIdx.Campus      >= 0) rowArr[sColIdx.Campus]      = acc.campus;
      if (sColIdx.Institution >= 0) rowArr[sColIdx.Institution] = acc.institution;
      rowArr[sColIdx.PasswordHash] = pwHash;
      if (sColIdx.GoogleEmail >= 0) rowArr[sColIdx.GoogleEmail] = '';
      rowArr[sColIdx.Status]       = 'Active';
      sh.appendRow(rowArr);
      added++;
      csvRows.push([acc.username, plainPw, acc.fullName, acc.campusCode, acc.campus, acc.instCode, acc.institution, newSid]);
    }

    // ── IPM Users row (uses campus CODE + institution CODE)
    var ipmPwHash = ipmHashPassword_(plainPw);
    if (existingIpmIdx[ukey]) {
      var ipmRow = existingIpmIdx[ukey];
      ipmUsers.getRange(ipmRow, iColIdx.password+1).setValue(ipmPwHash);
      ipmUsers.getRange(ipmRow, iColIdx.campus+1).setValue(acc.campusCode);
      ipmUsers.getRange(ipmRow, iColIdx.institution+1).setValue(acc.instCode);
      ipmUsers.getRange(ipmRow, iColIdx.fullName+1).setValue(acc.fullName);
    } else {
      var ipmRowArr = new Array(iH.length);
      ipmRowArr[iColIdx.username]    = acc.username;
      ipmRowArr[iColIdx.password]    = ipmPwHash;
      ipmRowArr[iColIdx.role]        = 'HOI';
      ipmRowArr[iColIdx.campus]      = acc.campusCode;
      ipmRowArr[iColIdx.institution] = acc.instCode;
      ipmRowArr[iColIdx.fullName]    = acc.fullName;
      ipmRowArr[iColIdx.createdAt]   = now;
      ipmUsers.appendRow(ipmRowArr);
    }
  });

  // ── Write credentials CSV to Drive
  var csvContent = csvRows.map(function(r){
    return r.map(function(cell){
      var s = String(cell == null ? '' : cell);
      return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    }).join(',');
  }).join('\n');

  var fileName = 'VMRF_HOI_Credentials_' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss') + '.csv';
  var fileUrl  = '';
  try {
    var file = DriveApp.createFile(fileName, csvContent, MimeType.CSV);
    fileUrl  = file.getUrl();
  } catch(e) {
    fileUrl = '(CSV creation failed: ' + e.message + ')';
  }

  try {
    SpreadsheetApp.getUi().alert(
      '✅ HOI Accounts Seeded\n\n' +
      'Total planned:    ' + planned.length + '\n' +
      'New accounts:     ' + added + '\n' +
      'Updated accounts: ' + updated + '\n\n' +
      'Credentials CSV saved to Drive:\n' + fileName + '\n\n' +
      (fileUrl ? 'Open: ' + fileUrl : '') + '\n\n' +
      '⚠️ Open the CSV and distribute credentials securely — HOIs should change their password after first sign-in.'
    );
  } catch(_){}
}

/* Build the full list of HOI account specs from INSTITUTION_HIERARCHY.
   One account per institution. Each spec:
     { username, campus, institution, campusCode, instCode, fullName } */
function _buildHoiAccountPlan_() {
  var plan = [];
  var seenUsernames = {};
  Object.keys(INSTITUTION_HIERARCHY || {}).forEach(function(campusName){
    var campus = INSTITUTION_HIERARCHY[campusName];
    var campusCode = campus.code || '';
    Object.keys(campus.institutions || {}).forEach(function(instName){
      var inst = campus.institutions[instName];
      var instCode = inst.code || '';
      var base = 'hoi_' + String(instCode).toLowerCase().replace(/[^a-z0-9]+/g,'');
      var username = base;
      var suffix = 2;
      while (seenUsernames[username]) { username = base + suffix; suffix++; }
      seenUsernames[username] = true;
      plan.push({
        username:    username,
        campus:      campusName,
        institution: instName,
        campusCode:  campusCode,
        instCode:    instCode,
        fullName:    'HOI - ' + instName
      });
    });
  });
  return plan;
}

/* Rotate passwords for existing HOI rows without touching other fields.
   Writes a fresh CSV to Drive so the admin can distribute the new passwords.
   Only touches accounts that already exist — does not create new ones. */
function resetHOIPasswords() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH.STAFF);
  if (!sh) { try { SpreadsheetApp.getUi().alert('Staff_Master sheet not found.'); } catch(_){} return; }

  var ipmSs    = ipmDb_();
  var ipmUsers = ipmSs.getSheetByName('Users');
  if (!ipmUsers) { try { SpreadsheetApp.getUi().alert('IPM Users sheet not found.'); } catch(_){} return; }

  var sData = sh.getDataRange().getValues(), sH = sData[0];
  var sColIdx = {
    StaffID:     sH.indexOf('StaffID'),
    StaffName:   sH.indexOf('StaffName'),
    Email:       sH.indexOf('Email'),
    Role:        sH.indexOf('Role'),
    Campus:      sH.indexOf('Campus'),
    Institution: sH.indexOf('Institution'),
    PasswordHash:sH.indexOf('PasswordHash')
  };
  var iData = ipmUsers.getDataRange().getValues(), iH = iData[0];
  var iColIdx = {
    username: iH.indexOf('username'),
    password: iH.indexOf('password'),
    role:     iH.indexOf('role')
  };
  var ipmByUsername = {};
  for (var r2 = 1; r2 < iData.length; r2++) {
    if (String(iData[r2][iColIdx.role]||'').toUpperCase() !== 'HOI') continue;
    var u2 = String(iData[r2][iColIdx.username]||'').trim().toLowerCase();
    if (u2) ipmByUsername[u2] = r2 + 1;
  }

  var now = new Date();
  var rotated = 0;
  var csvRows = [['Username','NewPassword','Full Name','Campus','Institution','StaffID']];

  for (var r = 1; r < sData.length; r++) {
    if (String(sData[r][sColIdx.Role]||'').toUpperCase() !== 'HOI') continue;
    var username = String(sData[r][sColIdx.Email]||'').trim();
    if (!username) continue;
    var plainPw = _genHodPassword_();
    sh.getRange(r+1, sColIdx.PasswordHash+1).setValue(_hashPwd(plainPw));
    var ipmRow = ipmByUsername[username.toLowerCase()];
    if (ipmRow) ipmUsers.getRange(ipmRow, iColIdx.password+1).setValue(ipmHashPassword_(plainPw));
    rotated++;
    csvRows.push([
      username,
      plainPw,
      sData[r][sColIdx.StaffName]||'',
      sColIdx.Campus>=0      ? (sData[r][sColIdx.Campus]||'')      : '',
      sColIdx.Institution>=0 ? (sData[r][sColIdx.Institution]||'') : '',
      sData[r][sColIdx.StaffID]||''
    ]);
  }

  var csvContent = csvRows.map(function(r){
    return r.map(function(cell){
      var s = String(cell == null ? '' : cell);
      return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    }).join(',');
  }).join('\n');
  var fileName = 'VMRF_HOI_Credentials_Reset_' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss') + '.csv';
  var fileUrl = '';
  try {
    var file = DriveApp.createFile(fileName, csvContent, MimeType.CSV);
    fileUrl = file.getUrl();
  } catch(e) { fileUrl = '(CSV creation failed: ' + e.message + ')'; }

  try {
    SpreadsheetApp.getUi().alert(
      '✅ HOI Passwords Reset\n\n' +
      'Accounts rotated: ' + rotated + '\n\n' +
      'CSV saved to Drive:\n' + fileName + '\n\n' +
      (fileUrl ? 'Open: ' + fileUrl : '')
    );
  } catch(_){}
}

/* ═══════════════════════════════════════════════════════════════════════════
 * HOI CREDENTIAL RECOVERY HELPERS
 *
 * Use these when seedHOIAccounts / resetHOIPasswords ran successfully but
 * you can't find the CSV in Drive. These print the fresh passwords directly
 * into the Apps Script execution log so you can copy them from View → Logs
 * without needing Drive access to work.
 *
 * Usage:
 *   1. Select `printHOICredentialsToLog` from the function dropdown.
 *   2. Click Run.
 *   3. Left sidebar → Executions → click the latest row → expand logs.
 *   4. Copy the printed table. Save it somewhere safe.
 *   5. Distribute the credentials to each institution's HOI.
 *
 * IMPORTANT: this function ROTATES passwords (generates new ones). Any
 * previously-distributed HOI passwords stop working the moment you run it.
 * The trade-off is that you get plaintext credentials to distribute —
 * existing hashed passwords in the sheets can never be recovered in
 * plaintext, only reset.
 * ═══════════════════════════════════════════════════════════════════════════ */
function printHOICredentialsToLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH.STAFF);
  if (!sh) { Logger.log('ERROR: Staff_Master sheet not found.'); return; }
  var ipmSs = ipmDb_();
  var ipmUsers = ipmSs.getSheetByName('Users');
  if (!ipmUsers) { Logger.log('ERROR: IPM Users sheet not found.'); return; }

  var sData = sh.getDataRange().getValues(), sH = sData[0];
  var sColIdx = {
    StaffID:     sH.indexOf('StaffID'),
    StaffName:   sH.indexOf('StaffName'),
    Email:       sH.indexOf('Email'),
    Role:        sH.indexOf('Role'),
    Campus:      sH.indexOf('Campus'),
    Institution: sH.indexOf('Institution'),
    PasswordHash:sH.indexOf('PasswordHash')
  };
  var iData = ipmUsers.getDataRange().getValues(), iH = iData[0];
  var iColIdx = {
    username: iH.indexOf('username'),
    password: iH.indexOf('password'),
    role:     iH.indexOf('role')
  };
  var ipmByUsername = {};
  for (var r2 = 1; r2 < iData.length; r2++) {
    if (String(iData[r2][iColIdx.role]||'').toUpperCase() !== 'HOI') continue;
    var u2 = String(iData[r2][iColIdx.username]||'').trim().toLowerCase();
    if (u2) ipmByUsername[u2] = r2 + 1;
  }

  Logger.log('═══════════════════════════════════════════════════════════════');
  Logger.log('HOI CREDENTIALS — COPY THIS TABLE BEFORE CLOSING THE LOG');
  Logger.log('═══════════════════════════════════════════════════════════════');
  Logger.log('');
  Logger.log('Username\tPassword\tStaffID\tCampus\tInstitution\tFull Name');
  Logger.log('--------\t--------\t-------\t------\t-----------\t---------');

  var rotated = 0;
  for (var r = 1; r < sData.length; r++) {
    if (String(sData[r][sColIdx.Role]||'').toUpperCase() !== 'HOI') continue;
    var username = String(sData[r][sColIdx.Email]||'').trim();
    if (!username) continue;
    var plainPw = _genHodPassword_();
    sh.getRange(r+1, sColIdx.PasswordHash+1).setValue(_hashPwd(plainPw));
    var ipmRow = ipmByUsername[username.toLowerCase()];
    if (ipmRow) ipmUsers.getRange(ipmRow, iColIdx.password+1).setValue(ipmHashPassword_(plainPw));
    rotated++;
    Logger.log(
      username + '\t' +
      plainPw + '\t' +
      (sData[r][sColIdx.StaffID]||'') + '\t' +
      (sColIdx.Campus>=0 ? (sData[r][sColIdx.Campus]||'') : '') + '\t' +
      (sColIdx.Institution>=0 ? (sData[r][sColIdx.Institution]||'') : '') + '\t' +
      (sData[r][sColIdx.StaffName]||'')
    );
  }

  Logger.log('');
  Logger.log('═══════════════════════════════════════════════════════════════');
  Logger.log('Total HOI accounts rotated: ' + rotated);
  Logger.log('');
  Logger.log('Also attempting to save the CSV to Drive as a backup…');

  // Attempt CSV creation separately and log success/failure explicitly so
  // we know whether Drive is the problem.
  try {
    var csvRows = [['Username','Password']];  // minimal re-fetch via log above
    // Re-read sheets to rebuild CSV from freshly-rotated data
    var sData2 = sh.getDataRange().getValues();
    Logger.log('(CSV will only include usernames — passwords are now hashed in the sheet.');
    Logger.log(' The plaintext passwords are ONLY in the log above. Copy them now.)');
  } catch(e) {
    Logger.log('Drive CSV creation would have thrown: ' + e.message);
  }
  Logger.log('═══════════════════════════════════════════════════════════════');
}

/* Diagnostic: lists every HOI account without touching passwords. Useful to
   confirm which accounts exist and in which campus/institution. Safe to run
   any number of times — read-only. */
function listHOIAccounts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH.STAFF);
  if (!sh) { Logger.log('ERROR: Staff_Master sheet not found.'); return; }
  var sData = sh.getDataRange().getValues(), sH = sData[0];
  var idx = {
    StaffID: sH.indexOf('StaffID'), StaffName: sH.indexOf('StaffName'),
    Email:   sH.indexOf('Email'),   Role:      sH.indexOf('Role'),
    Campus:  sH.indexOf('Campus'),  Institution:sH.indexOf('Institution'),
    Status:  sH.indexOf('Status')
  };
  Logger.log('HOI ACCOUNTS IN Staff_Master');
  Logger.log('StaffID\tUsername\tName\tCampus\tInstitution\tStatus');
  var n = 0;
  for (var r = 1; r < sData.length; r++) {
    if (String(sData[r][idx.Role]||'').toUpperCase() !== 'HOI') continue;
    n++;
    Logger.log(
      (sData[r][idx.StaffID]||'') + '\t' +
      (sData[r][idx.Email]||'') + '\t' +
      (sData[r][idx.StaffName]||'') + '\t' +
      (idx.Campus>=0 ? (sData[r][idx.Campus]||'') : '') + '\t' +
      (idx.Institution>=0 ? (sData[r][idx.Institution]||'') : '') + '\t' +
      (idx.Status>=0 ? (sData[r][idx.Status]||'') : '')
    );
  }
  Logger.log('Total HOI accounts: ' + n);
}

/* ═══════════════════════════════════════════════════════════════════════════
 * HOD LOGIN DIAGNOSTICS
 *
 * If HOD login is failing, run these from the Apps Script editor.
 * Pick a username and password from the CSV that was generated by seedHODAccounts.
 *
 * Quick usage from the editor:
 *   1. Select function `debugHODLogin` from the dropdown
 *   2. But first, edit this file: replace the placeholders in _debugHODLoginEntry
 *      with a real username and the password from your CSV, then Save.
 *   3. Run _debugHODLoginEntry. Open View → Logs to see the full diagnosis.
 *
 * The output shows exactly which step fails (row not found / status / hash /
 * campus / institution) and prints the stored vs. attempted values so you can
 * spot mismatches at a glance.
 * ═══════════════════════════════════════════════════════════════════════════ */

// ⇩ EDIT THESE TWO LINES, save, then run _debugHODLoginEntry from the editor.
function _debugHODLoginEntry() {
  var USERNAME = 'hod_avit_cse';      // ← paste the username from your CSV
  var PASSWORD = 'PASTE_PASSWORD_HERE'; // ← paste the plain password from your CSV
  var CAMPUS   = '';                  // ← optional: the campus value your login form would send (full name or code); leave '' to skip cascade check
  var INST     = '';                  // ← optional: same for institution
  debugHODLogin(USERNAME, PASSWORD, CAMPUS, INST);
}

/* Callable from the editor as debugHODLogin('hod_avit_cse','Abc123@x...') */
function debugHODLogin(username, password, campus, institution) {
  var log = [];
  var add = function(label, val) { log.push(label + ': ' + val); Logger.log(label + ': ' + val); };

  add('--- DEBUG HOD LOGIN ---', new Date().toISOString());
  add('Arg username', JSON.stringify(username));
  add('Arg password length', (password||'').length);
  add('Arg campus', JSON.stringify(campus||''));
  add('Arg institution', JSON.stringify(institution||''));

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SH.STAFF);
    if (!sh) { add('FATAL', 'Staff_Master sheet not found'); return log.join('\n'); }
    var data = sh.getDataRange().getValues(), h = data[0];
    add('Headers', JSON.stringify(h));

    var emI = h.indexOf('Email'), pwI = h.indexOf('PasswordHash'), rlI = h.indexOf('Role');
    var nmI = h.indexOf('StaffName'), depI = h.indexOf('Department');
    var stI = h.indexOf('Status'), idI = h.indexOf('StaffID');
    var camI = h.indexOf('Campus'), insI = h.indexOf('Institution');

    if (pwI < 0)     { add('FATAL', 'PasswordHash column missing'); return log.join('\n'); }
    if (emI < 0)     { add('FATAL', 'Email column missing');        return log.join('\n'); }
    if (stI < 0)     { add('FATAL', 'Status column missing');       return log.join('\n'); }
    if (camI < 0)    add('WARN', 'Campus column missing — run Setup');
    if (insI < 0)    add('WARN', 'Institution column missing — run Setup');

    var ident = String(username||'').trim().toLowerCase();
    var foundRow = -1, matchedData = null;

    // Scan all HOD rows, surfacing near-matches for debugging
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][rlI]||'').toUpperCase() !== 'HOD') continue;
      var rowEmail = String(data[i][emI]||'').trim().toLowerCase();
      if (rowEmail === ident) { foundRow = i; matchedData = data[i]; break; }
    }

    if (foundRow < 0) {
      add('STEP 1 — lookup', 'NO ROW FOUND for username "' + ident + '"');
      add('Hint', 'Scanning all HOD usernames in the sheet…');
      var allHodUsernames = [];
      for (var j = 1; j < data.length; j++) {
        if (String(data[j][rlI]||'').toUpperCase() === 'HOD') {
          allHodUsernames.push(String(data[j][emI]||''));
        }
      }
      add('All HOD usernames in Staff_Master', JSON.stringify(allHodUsernames));
      if (!allHodUsernames.length) add('CONCLUSION', 'Staff_Master has ZERO HOD rows. Did you run seedHODAccounts?');
      else add('CONCLUSION', 'Your username is not among the seeded ones. Check the CSV for the exact username.');
      return log.join('\n');
    }

    add('STEP 1 — lookup', 'FOUND at sheet row ' + (foundRow + 1));
    add('  StaffID',     String(matchedData[idI]||''));
    add('  StaffName',   String(matchedData[nmI]||''));
    add('  Email',       String(matchedData[emI]||''));
    add('  Department',  String(matchedData[depI]||''));
    add('  Campus',      camI>=0 ? String(matchedData[camI]||'') : '(no column)');
    add('  Institution', insI>=0 ? String(matchedData[insI]||'') : '(no column)');
    add('  Status',      String(matchedData[stI]||''));

    // STEP 2: Status check
    var statusVal = String(matchedData[stI]||'');
    if (statusVal !== 'Active') {
      add('STEP 2 — status', 'FAILED: Status="' + statusVal + '" (must be "Active")');
      return log.join('\n');
    }
    add('STEP 2 — status', 'OK (Active)');

    // STEP 3: Hash check
    var storedHash = String(matchedData[pwI]||'');
    var computedHash = _hashPwd(password||'');
    add('STEP 3 — hash', 'Stored:   "' + storedHash + '"');
    add('              ',  'Computed: "' + computedHash + '"');
    if (storedHash !== computedHash) {
      add('STEP 3 — hash', 'FAILED: passwords do not match');
      add('Hint', 'Verify the password is exactly as in the CSV (no leading/trailing spaces, no smart quotes). If you re-ran seedHODAccounts or resetHODPasswords after the CSV you have, the stored password changed.');
      return log.join('\n');
    }
    add('STEP 3 — hash', 'OK (passwords match)');

    // STEP 4: Cascade check (if campus/institution args given)
    if (String(campus||'').trim() || String(institution||'').trim()) {
      var rowCampus = camI >= 0 ? String(matchedData[camI]||'').trim() : '';
      var rowInst   = insI >= 0 ? String(matchedData[insI]||'').trim() : '';
      var want = _hodResolveDbg_(String(campus||'').trim(), String(institution||'').trim());
      var have = _hodResolveDbg_(rowCampus, rowInst);
      add('STEP 4 — cascade', '');
      add('  want campus/inst codes', JSON.stringify(want));
      add('  have campus/inst codes', JSON.stringify(have));
      if (want.campusCode && have.campusCode && want.campusCode !== have.campusCode) {
        add('STEP 4 — cascade', 'FAILED: campus code mismatch (' + want.campusCode + ' vs ' + have.campusCode + ')');
        return log.join('\n');
      }
      if (want.instCode && have.instCode && want.instCode !== have.instCode) {
        add('STEP 4 — cascade', 'FAILED: institution code mismatch (' + want.instCode + ' vs ' + have.instCode + ')');
        return log.join('\n');
      }
      add('STEP 4 — cascade', 'OK (matched)');
    } else {
      add('STEP 4 — cascade', 'SKIPPED (no campus/institution args provided)');
    }

    add('CONCLUSION', '✅ All checks pass — this username+password WOULD log in successfully. If the UI still rejects it, check that the login form is reaching the same Staff_Master sheet (and not a duplicate or test sheet).');
  } catch (e) {
    add('EXCEPTION', e.message + '\n' + (e.stack||''));
  }
  return log.join('\n');
}

// Mirror of _hodResolve used in staffLogin, exported for the diagnostics.
function _hodResolveDbg_(campusVal, instVal) {
  var out = { campusCode: '', instCode: '' };
  if (!campusVal && !instVal) return out;
  try {
    if (INSTITUTION_HIERARCHY[campusVal]) {
      out.campusCode = INSTITUTION_HIERARCHY[campusVal].code || '';
    } else {
      Object.keys(INSTITUTION_HIERARCHY).forEach(function(k){
        if (INSTITUTION_HIERARCHY[k].code === campusVal) out.campusCode = campusVal;
      });
    }
    var searchIn = out.campusCode
      ? [Object.keys(INSTITUTION_HIERARCHY).filter(function(k){ return INSTITUTION_HIERARCHY[k].code === out.campusCode; })[0]]
      : Object.keys(INSTITUTION_HIERARCHY);
    for (var i=0; i<searchIn.length; i++) {
      var camp = searchIn[i]; if (!camp) continue;
      var insts = INSTITUTION_HIERARCHY[camp].institutions || {};
      if (insts[instVal]) { out.instCode = insts[instVal].code || ''; break; }
      var hit = Object.keys(insts).filter(function(n){ return insts[n].code === instVal; })[0];
      if (hit) { out.instCode = instVal; break; }
    }
  } catch(e) {}
  return out;
}

/* List every HOD row with its key fields for manual inspection.
   Run from the editor; output goes to View → Logs. */
function listHODAccounts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH.STAFF);
  if (!sh) { Logger.log('Staff_Master sheet not found'); return; }
  var data = sh.getDataRange().getValues(), h = data[0];
  var emI = h.indexOf('Email'), rlI = h.indexOf('Role'), nmI = h.indexOf('StaffName');
  var depI = h.indexOf('Department'), camI = h.indexOf('Campus'), insI = h.indexOf('Institution');
  var stI = h.indexOf('Status'), pwI = h.indexOf('PasswordHash');
  Logger.log('Staff_Master headers: ' + JSON.stringify(h));
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][rlI]||'').toUpperCase() !== 'HOD') continue;
    count++;
    Logger.log(
      (count) + '. "' + String(data[i][emI]||'') + '"' +
      ' | Status=' + String(data[i][stI]||'') +
      ' | Dept=' + String(data[i][depI]||'') +
      ' | Campus=' + (camI>=0 ? String(data[i][camI]||'') : '-') +
      ' | Inst=' + (insI>=0 ? String(data[i][insI]||'') : '-') +
      ' | HashPrefix=' + String(data[i][pwI]||'').substring(0, 12) + '…'
    );
  }
  Logger.log('Total HOD rows: ' + count);
}

/* Set a custom password for a single HOD — useful if you need to rescue one
   account without re-running the full seeder. Usage from the editor:
     setHODPassword('hod_avit_cse', 'MyChosenPassword!')                    */
function setHODPassword(username, newPlainPassword) {
  if (!username || !newPlainPassword) throw new Error('Both username and newPlainPassword are required');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH.STAFF);
  if (!sh) throw new Error('Staff_Master sheet not found');
  var data = sh.getDataRange().getValues(), h = data[0];
  var emI = h.indexOf('Email'), pwI = h.indexOf('PasswordHash'), rlI = h.indexOf('Role'), stI = h.indexOf('Status');
  var ident = String(username).trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][rlI]||'').toUpperCase() !== 'HOD') continue;
    if (String(data[i][emI]||'').trim().toLowerCase() !== ident) continue;
    sh.getRange(i+1, pwI+1).setValue(_hashPwd(newPlainPassword));
    sh.getRange(i+1, stI+1).setValue('Active');
    // Also sync to IPM Users if present there
    try {
      var ipmUsers = ipmSheet_('Users');
      var iData = ipmUsers.getDataRange().getValues(), iH = iData[0];
      var iUser = iH.indexOf('username'), iPw = iH.indexOf('password');
      for (var j = 1; j < iData.length; j++) {
        if (String(iData[j][iUser]||'').trim().toLowerCase() === ident) {
          ipmUsers.getRange(j+1, iPw+1).setValue(ipmHashPassword_(newPlainPassword));
          break;
        }
      }
    } catch(_){}
    Logger.log('✅ Password updated for ' + ident);
    try { SpreadsheetApp.getUi().alert('✅ Password updated for ' + ident + '\n\nNew password: ' + newPlainPassword); } catch(_){}
    return { ok: true };
  }
  throw new Error('No HOD row with username "' + ident + '"');
}

// ─── RESET ALL HOD PASSWORDS ─────────────────────────────────────────────────
// Run this from the VMRF menu when HODs need fresh credentials. Generates a
// new random password for every HOD in Staff_Master, writes the hash to both
// Staff_Master AND IPM Users, and exports the new credentials to a CSV in Drive.
function resetHODPasswords() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH.STAFF);
  if (!sh) { try { SpreadsheetApp.getUi().alert('Staff_Master sheet not found.'); } catch(_){} return; }
  _ensureSheetColumns(sh, SCHEMA.Staff_Master);

  var data = sh.getDataRange().getValues(), h = data[0];
  var emI = h.indexOf('Email'), pwI = h.indexOf('PasswordHash'), rlI = h.indexOf('Role');
  var nmI = h.indexOf('StaffName'), depI = h.indexOf('Department'), idI = h.indexOf('StaffID');
  var camI = h.indexOf('Campus'), insI = h.indexOf('Institution');

  var ipmSs    = ipmDb_();
  var ipmUsers = ipmSs.getSheetByName('Users');
  var iData = ipmUsers.getDataRange().getValues(), iH = iData[0];
  var iUserCol = iH.indexOf('username'), iPwCol = iH.indexOf('password'), iRoleCol = iH.indexOf('role');
  var ipmByUsername = {};
  for (var r2 = 1; r2 < iData.length; r2++) {
    if (String(iData[r2][iRoleCol]||'').toUpperCase() !== 'HOD') continue;
    ipmByUsername[String(iData[r2][iUserCol]||'').trim().toLowerCase()] = r2 + 1;
  }

  var now = new Date();
  var csvRows = [['Username','Password','Full Name','Campus','Institution','Department','StaffID']];
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][rlI]||'').toUpperCase() !== 'HOD') continue;
    var username = String(data[i][emI]||'').trim();
    if (!username) continue;
    var plainPw = _genHodPassword_();
    // Staff_Master hash
    sh.getRange(i + 1, pwI + 1).setValue(_hashPwd(plainPw));
    // IPM Users hash (if username exists there)
    var ukey = username.toLowerCase();
    if (ipmByUsername[ukey]) {
      ipmUsers.getRange(ipmByUsername[ukey], iPwCol + 1).setValue(ipmHashPassword_(plainPw));
    }
    csvRows.push([
      username, plainPw,
      String(data[i][nmI]||''),
      camI>=0 ? String(data[i][camI]||'') : '',
      insI>=0 ? String(data[i][insI]||'') : '',
      String(data[i][depI]||''),
      String(data[i][idI]||'')
    ]);
    count++;
  }

  var csvContent = csvRows.map(function(r){
    return r.map(function(cell){
      var s = String(cell == null ? '' : cell);
      return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    }).join(',');
  }).join('\n');

  var fileName = 'VMRF_HOD_Credentials_RESET_' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss') + '.csv';
  var fileUrl  = '';
  try {
    var file = DriveApp.createFile(fileName, csvContent, MimeType.CSV);
    fileUrl  = file.getUrl();
  } catch(e) {
    fileUrl = '(CSV creation failed: ' + e.message + ')';
  }

  try {
    SpreadsheetApp.getUi().alert(
      '✅ Reset ' + count + ' HOD passwords.\n\n' +
      'New credentials written to Drive:\n' + fileName + '\n\n' +
      (fileUrl ? 'Open: ' + fileUrl : '')
    );
  } catch(_){}
}

/**
 * resetAllPasswordsExceptIMO
 * ─────────────────────────────────────────────────────────────────────
 * One-shot administrative reset.  Generates a new random password for
 * every Faculty, HOD and HOI account and writes the new hash into:
 *   • Faculty_Master  (every row)            — hash via _hashPwd
 *   • Staff_Master    (Role = HOD or HOI)    — hash via _hashPwd
 *   • IPM "Users"     (role ≠ IMO)           — hash via ipmHashPassword_
 *
 * IMO accounts are deliberately UNTOUCHED — both the Script-Properties
 * IMO_PASSWORD and any IPM Users row whose role is "IMO" are skipped.
 *
 * The new credentials are exported to a CSV file in Drive so the
 * administrator can distribute them.  This is a destructive,
 * one-way operation and prompts for confirmation when run from the
 * spreadsheet menu.
 */
function resetAllPasswordsExceptIMO() {
  // ── 0. Confirmation prompt (skipped when invoked headlessly) ──
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch(_){}
  if (ui) {
    var resp = ui.alert(
      'Reset all passwords (except IMO)?',
      'This will generate NEW random passwords for every Faculty, HOD ' +
      'and HOI account.\n\n' +
      'IMO accounts will NOT be touched.\n\n' +
      'New credentials will be written to a CSV file in Drive.\n' +
      'This cannot be undone. Continue?',
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) {
      ui.alert('Cancelled. No passwords were changed.');
      return;
    }
  }

  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var now = new Date();
  var csvRows = [['Role','Username/Email','New Password','Name','Campus','Institution','Department','SheetSource']];
  var counts  = { faculty:0, hod:0, hoi:0, ipmExtra:0, ipmSkippedImo:0 };

  // ── 1. Pre-load IPM Users so Staff_Master pass can mirror updates ──
  var ipmUsers = null, iData = null, iH = null;
  var iUserCol = -1, iPwCol = -1, iRoleCol = -1, iFnCol = -1, iCampCol = -1, iInstCol = -1;
  var ipmByKey = {};            // lower(username) → 1-based row index
  try {
    var ipmSs = ipmDb_();
    ipmUsers  = ipmSs ? ipmSs.getSheetByName('Users') : null;
  } catch(e) { ipmUsers = null; }
  if (ipmUsers) {
    iData = ipmUsers.getDataRange().getValues();
    if (iData.length > 0) {
      iH       = iData[0];
      iUserCol = iH.indexOf('username');
      iPwCol   = iH.indexOf('password');
      iRoleCol = iH.indexOf('role');
      iFnCol   = iH.indexOf('fullName');
      iCampCol = iH.indexOf('campus');
      iInstCol = iH.indexOf('institution');
      for (var r2 = 1; r2 < iData.length; r2++) {
        var ukey = String(iData[r2][iUserCol]||'').trim().toLowerCase();
        if (ukey) ipmByKey[ukey] = r2 + 1;
      }
    }
  }
  var ipmTouched = {};   // usernames already updated via Staff_Master pass

  // ── 2. Faculty_Master — reset every faculty ──
  var facSh = ss.getSheetByName(SH.FACULTY);
  if (facSh) {
    try { _ensureFacultyColumns(facSh); } catch(_){}
    var fData = facSh.getDataRange().getValues();
    if (fData.length > 1) {
      var fH = fData[0];
      var fEmI = fH.indexOf('Email');
      var fNmI = fH.indexOf('FacultyName');
      var fPwI = fH.indexOf('PasswordHash');
      var fDepI = fH.indexOf('Department');
      var fCaI  = fH.indexOf('Campus');
      var fInI  = fH.indexOf('Institution');
      if (fEmI >= 0 && fPwI >= 0) {
        for (var i = 1; i < fData.length; i++) {
          var fEmail = String(fData[i][fEmI]||'').trim();
          if (!fEmail) continue;
          var plainPw = _genHodPassword_();
          facSh.getRange(i + 1, fPwI + 1).setValue(_hashPwd(plainPw));
          csvRows.push([
            'FACULTY', fEmail, plainPw,
            fNmI  >= 0 ? String(fData[i][fNmI] ||'') : '',
            fCaI  >= 0 ? String(fData[i][fCaI] ||'') : '',
            fInI  >= 0 ? String(fData[i][fInI] ||'') : '',
            fDepI >= 0 ? String(fData[i][fDepI]||'') : '',
            'Faculty_Master'
          ]);
          counts.faculty++;
        }
      }
    }
  }

  // ── 3. Staff_Master HOD/HOI (skip IMO) + paired IPM Users update ──
  var stSh = ss.getSheetByName(SH.STAFF);
  if (stSh) {
    try { _ensureSheetColumns(stSh, SCHEMA.Staff_Master); } catch(_){}
    var sData = stSh.getDataRange().getValues();
    if (sData.length > 1) {
      var sH = sData[0];
      var sEmI = sH.indexOf('Email');
      var sNmI = sH.indexOf('StaffName');
      var sPwI = sH.indexOf('PasswordHash');
      var sRlI = sH.indexOf('Role');
      var sDepI = sH.indexOf('Department');
      var sCaI = sH.indexOf('Campus');
      var sInI = sH.indexOf('Institution');
      if (sEmI >= 0 && sPwI >= 0 && sRlI >= 0) {
        for (var j = 1; j < sData.length; j++) {
          var role = String(sData[j][sRlI]||'').trim().toUpperCase();
          if (role === 'IMO') continue;                  // ← IMO is left alone
          if (role !== 'HOD' && role !== 'HOI') continue;
          var sEmail = String(sData[j][sEmI]||'').trim();
          if (!sEmail) continue;
          var plainPw2 = _genHodPassword_();
          stSh.getRange(j + 1, sPwI + 1).setValue(_hashPwd(plainPw2));
          // Mirror to IPM Users sheet if a matching row exists
          var ukey2 = sEmail.toLowerCase();
          if (ipmUsers && ipmByKey[ukey2] && iPwCol >= 0) {
            ipmUsers.getRange(ipmByKey[ukey2], iPwCol + 1).setValue(ipmHashPassword_(plainPw2));
            ipmTouched[ukey2] = true;
          }
          csvRows.push([
            role, sEmail, plainPw2,
            sNmI  >= 0 ? String(sData[j][sNmI] ||'') : '',
            sCaI  >= 0 ? String(sData[j][sCaI] ||'') : '',
            sInI  >= 0 ? String(sData[j][sInI] ||'') : '',
            sDepI >= 0 ? String(sData[j][sDepI]||'') : '',
            'Staff_Master'
          ]);
          if (role === 'HOD') counts.hod++; else counts.hoi++;
        }
      }
    }
  }

  // ── 4. IPM Users — catch rows with no Staff_Master twin (skip IMO) ──
  if (ipmUsers && iData && iUserCol >= 0 && iPwCol >= 0 && iRoleCol >= 0) {
    for (var k = 1; k < iData.length; k++) {
      var iRole = String(iData[k][iRoleCol]||'').trim().toUpperCase();
      if (iRole === 'IMO') { counts.ipmSkippedImo++; continue; }   // ← skip IMO
      var iUser = String(iData[k][iUserCol]||'').trim();
      if (!iUser) continue;
      if (ipmTouched[iUser.toLowerCase()]) continue;               // already done
      var plainPw3 = _genHodPassword_();
      ipmUsers.getRange(k + 1, iPwCol + 1).setValue(ipmHashPassword_(plainPw3));
      csvRows.push([
        iRole || 'IPM', iUser, plainPw3,
        iFnCol   >= 0 ? String(iData[k][iFnCol]  ||'') : '',
        iCampCol >= 0 ? String(iData[k][iCampCol]||'') : '',
        iInstCol >= 0 ? String(iData[k][iInstCol]||'') : '',
        '',
        'IPM Users (only)'
      ]);
      counts.ipmExtra++;
    }
  }

  // ── 5. Save credentials CSV to Drive ──
  var csvContent = csvRows.map(function(row){
    return row.map(function(cell){
      var s = String(cell == null ? '' : cell);
      return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    }).join(',');
  }).join('\n');

  var fileName = 'VMRF_PasswordReset_AllExceptIMO_' +
    Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss') + '.csv';
  var fileUrl  = '';
  try {
    var file = DriveApp.createFile(fileName, csvContent, MimeType.CSV);
    fileUrl  = file.getUrl();
  } catch(e) {
    fileUrl = '(CSV creation failed: ' + e.message + ')';
  }

  // ── 6. Summary ──
  var totalIssued = counts.faculty + counts.hod + counts.hoi + counts.ipmExtra;
  var summary =
    '✅ Password reset complete (IMO untouched)\n\n' +
    'Faculty_Master       : ' + counts.faculty       + '\n' +
    'Staff_Master  (HOD)  : ' + counts.hod           + '\n' +
    'Staff_Master  (HOI)  : ' + counts.hoi           + '\n' +
    'IPM Users     extra  : ' + counts.ipmExtra      + '   (rows with no Staff_Master twin)\n' +
    'IPM Users     IMO    : ' + counts.ipmSkippedImo + '   (skipped — IMO not reset)\n\n' +
    'Total credentials issued: ' + totalIssued + '\n\n' +
    'Saved to Drive:\n' + fileName + (fileUrl ? '\n\n' + fileUrl : '');

  Logger.log(summary);
  if (ui) { try { ui.alert(summary); } catch(_){} }
  return { ok:true, counts:counts, file:fileName, url:fileUrl };
}

function getConfig() {
  return {
    activityTypes:        ACTIVITY_TYPES,
    timeSlots:            TIME_SLOTS,
    institutionTimeSlots: INSTITUTION_TIME_SLOTS,
    departments:          DEPARTMENTS,
    ugDepartments:        UG_DEPARTMENTS,
    pgDepartments:        PG_DEPARTMENTS,
    designations:         DESIGNATIONS,
    campuses:             CAMPUSES,
    academicYears:        ACADEMIC_YEARS,
    days:                 DAYS,
    institutions:         INSTITUTIONS,
    hierarchy:            _buildHierarchyForClient_()
  };
}

// ─── FACULTY REGISTER ─── see selfRegisterFaculty in AUTH section above ──




// ─── FACULTY LIST ─────────────────────────────────────────────────────────────
function getFacultyList() {
  var data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.FACULTY).getDataRange().getValues();
  if (!data || data.length < 2) return [];
  var h = data[0].map(function(v){ return String(v||'').trim(); });
  
  // Bulletproof column detection — try multiple possible header names
  var nmI = h.indexOf('FacultyName');
  if (nmI < 0) { for (var x=0;x<h.length;x++) { if(h[x].toLowerCase().indexOf('name')>=0){nmI=x;break;} } }
  if (nmI < 0) nmI = 0;
  
  var emI = h.indexOf('Email');
  if (emI < 0) emI = h.indexOf('FacultyEmail');
  if (emI < 0) emI = h.indexOf('FacultyID');
  if (emI < 0) { for (var x2=0;x2<h.length;x2++) { if(h[x2].toLowerCase().indexOf('mail')>=0){emI=x2;break;} } }
  if (emI < 0) emI = 1;
  
  var dpI = h.indexOf('Department');  if (dpI < 0) dpI = 2;
  var cpI = h.indexOf('Campus');      if (cpI < 0) cpI = 3;
  var inI = h.indexOf('Institution'); if (inI < 0) inI = 4;
  var dgI = h.indexOf('Designation'); if (dgI < 0) dgI = 5;
  var stI = h.indexOf('Status');      if (stI < 0) stI = h.length - 1;
  
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var name = String(row[nmI]||'').trim();
    var email = String(row[emI]||'').trim();
    var status = String(row[stI]||'').trim();
    // Include if has a name and status is Active (or no status column)
    if (!name) continue;
    if (status && status !== 'Active') continue;
    out.push({
      id:          email,
      name:        name,
      dept:        String(row[dpI]||'').trim(),
      campus:      String(row[cpI]||'').trim(),
      institution: String(row[inI]||'').trim(),
      designation: String(row[dgI]||'').trim()
    });
  }
  return out;
}

// Try authenticating main-app HOI/IMO directly against IPM Users sheet.
// Returns { ok, token?, user?, error? } — never throws.
function _tryIpmLoginForMain_(role, username, password) {
  try {
    if (!username || !password) return { ok:false };
    var users = ipmRowsAsObjects_(ipmSheet_('Users'));
    var u = users.find(function(x){ return String(x.username).toLowerCase() === String(username).trim().toLowerCase(); });
    if (!u) return { ok:false };
    if (String(u.role).toUpperCase() !== String(role).toUpperCase()) return { ok:false };
    if (!ipmVerifyPassword_(password, u.password)) return { ok:false };
    // On first login, stored pw may still be plaintext → auto-hash (same behaviour as ipmLogin)
    if (!ipmIsHashed_(u.password)) {
      var usersSheet = ipmSheet_('Users');
      var idx = users.findIndex(function(x){ return String(x.username).toLowerCase() === String(u.username).toLowerCase(); });
      if (idx >= 0) usersSheet.getRange(idx + 2, 2).setValue(ipmHashPassword_(password));
    }
    return { ok:true, token: _ipmMintSession_(u), user: _ipmUserView_(u) };
  } catch (e) {
    return { ok:false, error: e.message };
  }
}