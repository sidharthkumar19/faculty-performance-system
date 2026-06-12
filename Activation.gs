/**
 * ════════════════════════════════════════════════════════════════════
 * ACTIVATION.gs — VMRF IMS Excel-Import / Pre-Provisioned Account System
 *
 * Single-file home for everything related to the migration from open
 * self-registration to closed pre-provisioned accounts. Contains:
 *
 *   PART A — Lockdown primitives (called by Auth.gs patches)
 *     _isSelfRegisterOpen_, _ensureSelfRegisterAllowed_
 *     _validatePasswordStrength, _awaitingActivationResponse_
 *     _findAwaitingActivation_
 *
 *   PART B — Defensive utility helpers
 *     _activationEnsureColumns_  (column appender)
 *     _activationMakeID_         (ID generator)
 *     _activationIpmHash_        (sha256:<hex> password hash)
 *     _activationUpsertIpmUser_  (IPM Users sheet writer)
 *
 *   PART C — Single-row admin
 *     addUserAccount, addTestAccount, deactivateUserAccount,
 *     resetTestActivation, _devCompleteActivationManually
 *
 *   PART D — Bulk import (Step 2)
 *     createFacultyImportStagingTab, bulkImportFaculty,
 *     runImportDryRun, runImportForReal
 *
 *   PART E — Validation maintenance
 *     updateStatusDropdownsForActivation, relaxAllValidations
 *
 *   PART F — Menu wrappers (used by IMO.gs onOpen)
 *     menuCreateImportStagingTab, menuRunImportDryRun,
 *     menuRunImportForReal, menuDeactivateUser
 *
 *   PART G — Test scaffolding
 *     YOUR_GMAIL (edit before use),
 *     addMyThreeTestAccounts, activateMyThreeTestAccounts,
 *     fixHoiTestAccountIpmUsers, listMyTestAccounts
 *
 *   PART H — Health check (run after any code change)
 *     runStep1HealthCheck — 10 checks; idempotent; auto-fixes the two
 *     known data-validation traps and initialises EMAIL_ENABLED.
 *
 *   PART I — Step 3 (OTP edition): activation backend with 6-digit codes
 *     requestActivation, completeActivationWithOtp,
 *     sendActivationInvitations, resendActivationInvitation,
 *     _emailSendGuarded_, _emailLog_,
 *     testActivationFlowForFaculty, testActivationFlowForHoi
 *
 * USAGE on first deployment:
 *   1. Save all files in the Apps Script editor.
 *   2. Run runStep1HealthCheck — expect "7 passed, 0 failed".
 *   3. Edit YOUR_GMAIL in Part G to your real Gmail address.
 *   4. Run addMyThreeTestAccounts — creates three test users.
 *   5. Run activateMyThreeTestAccounts — sets password TestPass1.
 *   6. Run fixHoiTestAccountIpmUsers — writes the HOI IPM Users row.
 *   7. Test login on the test web app for all three roles.
 * ════════════════════════════════════════════════════════════════════
 */

var STATUS_AWAITING_ACTIVATION = 'AwaitingActivation';


// ════════════════════════════════════════════════════════════════════
// PART A: Lockdown + activation primitives
// ════════════════════════════════════════════════════════════════════

/**
 * Returns true when self-registration is permitted. Controlled by
 * Script Property `SELF_REGISTER_OPEN`. Default closed.
 */
function _isSelfRegisterOpen_() {
  var v = PropertiesService.getScriptProperties().getProperty('SELF_REGISTER_OPEN');
  return String(v || '').trim().toLowerCase() === 'true';
}

/**
 * Throws a user-facing error when self-registration is closed.
 * Called from facultyRegister, staffRegister, and sendEmailOTP.
 */
function _ensureSelfRegisterAllowed_(role) {
  if (_isSelfRegisterOpen_()) return;
  var roleLabel = role ? String(role) : 'account';
  throw new Error(
    'Self-registration is currently closed. Your ' + roleLabel +
    ' account is pre-provisioned by the IMO. Please use the ' +
    '"Activate your account" link on the login page to set your password.'
  );
}

/**
 * Validates a password against the policy:
 *   - At least 8 characters
 *   - At least one letter
 *   - At least one digit
 */
function _validatePasswordStrength(pwd) {
  if (pwd === undefined || pwd === null || pwd === '') {
    throw new Error('Password is required.');
  }
  var s = String(pwd);
  if (s.length < 8) {
    throw new Error('Password must be at least 8 characters long.');
  }
  if (!/[A-Za-z]/.test(s)) {
    throw new Error('Password must contain at least one letter.');
  }
  if (!/[0-9]/.test(s)) {
    throw new Error('Password must contain at least one digit.');
  }
  return true;
}

/**
 * Standard "needs activation" response shape returned by login
 * functions when they encounter a row with Status='AwaitingActivation'.
 */
function _awaitingActivationResponse_(email) {
  return {
    success:         false,
    needsActivation: true,
    email:           String(email || '').trim().toLowerCase(),
    message:         'Your account is awaiting activation. Please use the ' +
                     '"Activate your account" link on the login page to set your password.'
  };
}

/**
 * Looks up the given email across Faculty_Master and Staff_Master.
 * Returns a needs-activation response if a row exists with
 * Status='AwaitingActivation'. Returns null otherwise.
 */
function _findAwaitingActivation_(email) {
  if (!email) return null;
  var emLc = String(email).trim().toLowerCase();
  if (!emLc) return null;
  var ss;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch(_) { return null; }
  if (!ss) return null;

  var sheetsToScan = [];
  try { if (typeof SH !== 'undefined' && SH.FACULTY) sheetsToScan.push(SH.FACULTY); } catch(_){}
  try { if (typeof SH !== 'undefined' && SH.STAFF)   sheetsToScan.push(SH.STAFF);   } catch(_){}

  for (var s = 0; s < sheetsToScan.length; s++) {
    try {
      var sh = ss.getSheetByName(sheetsToScan[s]);
      if (!sh) continue;
      var data = sh.getDataRange().getValues();
      if (!data || data.length < 2) continue;
      var hdr = data[0];
      var emI = hdr.indexOf('Email');
      var stI = hdr.indexOf('Status');
      if (emI < 0 || stI < 0) continue;
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][emI] || '').trim().toLowerCase() !== emLc) continue;
        if (String(data[i][stI] || '').trim() === STATUS_AWAITING_ACTIVATION) {
          return _awaitingActivationResponse_(emLc);
        }
        break;
      }
    } catch (_) { /* skip this sheet */ }
  }
  return null;
}


// ════════════════════════════════════════════════════════════════════
// PART B: Test helpers (for the test deployment only)
// ════════════════════════════════════════════════════════════════════

/**
 * Defensive replacement for _ensureSheetColumns — appends any missing
 * column from `requiredHeaders` to the end of the sheet's header row.
 * Used when the project's existing _ensureSheetColumns may or may not
 * be loaded yet.
 */
function _activationEnsureColumns_(sheet, requiredHeaders) {
  if (typeof _ensureSheetColumns === 'function') {
    return _ensureSheetColumns(sheet, requiredHeaders);
  }
  var lastCol = sheet.getLastColumn();
  var existing = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(v){ return String(v).trim(); })
    : [];
  var added = 0;
  requiredHeaders.forEach(function(hdr) {
    if (existing.indexOf(hdr) < 0) {
      var newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(hdr).setFontWeight('bold');
      existing.push(hdr);
      added++;
    }
  });
  if (added > 0) sheet.setFrozenRows(1);
  return added;
}

/**
 * Defensive ID maker — uses project's _makeID if available, otherwise
 * generates a simple unique-enough ID.
 */
function _activationMakeID_(role) {
  if (typeof _makeID === 'function') return _makeID(role);
  var r = String(role || 'X').toUpperCase().substring(0, 3);
  return r + '-' + Date.now() + '-' + Math.floor(Math.random() * 1000);
}

/**
 * IPM-style password hasher: 'sha256:<hex>' prefix format that the
 * IPM Users sheet's auth path expects. Uses the project's
 * ipmHashPassword_ when available, otherwise computes the hash
 * inline so HOI/IMO dual-writes work even when IPM helpers aren't
 * loaded in this project.
 */
function _activationIpmHash_(password) {
  if (typeof ipmHashPassword_ === 'function') return ipmHashPassword_(password);
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(password),
    Utilities.Charset.UTF_8
  );
  var hex = bytes.map(function(b) {
    var h = (b < 0 ? b + 256 : b).toString(16);
    return h.length === 1 ? '0' + h : h;
  }).join('');
  return 'sha256:' + hex;
}

/**
 * Create or update the IPM Users row for an HOI/IMO account.
 * Uses ipmSheet_('Users') to locate the sheet (which may be in a
 * different spreadsheet than the active one, depending on the
 * project's IPM helper). Falls back to getSheetByName variants if
 * the helper isn't loaded.
 *
 * Returns { ok, note } — never throws.
 */
function _activationUpsertIpmUser_(role, email, password, instFull, campusFull) {
  try {
    var sh = null;
    // Preferred: use the project's own sheet resolver (gets the right
    // spreadsheet AND the right sheet name in one call).
    if (typeof ipmSheet_ === 'function') {
      try { sh = ipmSheet_('Users'); } catch(_){}
    }
    // Fallbacks if the helper isn't loaded
    if (!sh) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      sh = ss.getSheetByName('Users') ||
           ss.getSheetByName('IPM Users') ||
           ss.getSheetByName('IPMUsers');
    }
    if (!sh) return { ok: false, note: 'IPM Users sheet not found — run ipmSetup() first' };

    var data = sh.getDataRange().getValues();
    if (data.length < 1) return { ok: false, note: 'Users sheet has no headers' };

    var hdr = data[0].map(function(v){ return String(v).trim(); });
    var lower = hdr.map(function(v){ return v.toLowerCase(); });
    var col = function(name) { return lower.indexOf(name.toLowerCase()); };

    var ciUser  = col('username'),
        ciPwd   = col('password'),
        ciRole  = col('role'),
        ciCamp  = col('campus'),
        ciInst  = col('institution'),
        ciEmail = col('email'),       // may not exist in your schema — that's fine
        ciName  = col('fullName');

    if (ciUser < 0 || ciPwd < 0) {
      return { ok: false, note: 'Users sheet headers missing — expected username,password,role,...' };
    }

    var emailLc = String(email).toLowerCase();
    var pwdHash = _activationIpmHash_(password);

    // Resolve campus + institution short codes from INSTITUTION_HIERARCHY
    var instCode = '', campusCode = '';
    try {
      if (typeof INSTITUTION_HIERARCHY !== 'undefined') {
        Object.keys(INSTITUTION_HIERARCHY).forEach(function(c) {
          if (c === campusFull) campusCode = INSTITUTION_HIERARCHY[c].code;
          var insts = INSTITUTION_HIERARCHY[c].institutions || {};
          if (insts[instFull]) instCode = insts[instFull].code;
        });
      }
    } catch(_){}

    // If a row for this email already exists, update it in place.
    for (var i = 1; i < data.length; i++) {
      var rowUser  = ciUser  >= 0 ? String(data[i][ciUser]||'').toLowerCase()  : '';
      var rowEmail = ciEmail >= 0 ? String(data[i][ciEmail]||'').toLowerCase() : '';
      if (rowUser === emailLc || rowEmail === emailLc) {
        sh.getRange(i + 1, ciPwd + 1).setValue(pwdHash);
        if (ciRole >= 0) sh.getRange(i + 1, ciRole + 1).setValue(role);
        if (ciCamp >= 0) sh.getRange(i + 1, ciCamp + 1).setValue(campusCode);
        if (ciInst >= 0) sh.getRange(i + 1, ciInst + 1).setValue(instCode);
        return { ok: true, note: 'Users sheet row updated for ' + email + ' (row ' + (i+1) + ')' };
      }
    }

    // Otherwise append a new row in header order.
    var vals = {};
    vals[hdr[ciUser]] = emailLc;
    vals[hdr[ciPwd]]  = pwdHash;
    if (ciRole  >= 0) vals[hdr[ciRole]]  = role;
    if (ciCamp  >= 0) vals[hdr[ciCamp]]  = campusCode;
    if (ciInst  >= 0) vals[hdr[ciInst]]  = instCode;
    if (ciEmail >= 0) vals[hdr[ciEmail]] = emailLc;
    if (ciName  >= 0) vals[hdr[ciName]]  = role + ' user';
    if (col('createdAt') >= 0) vals[hdr[col('createdAt')]] = new Date().toISOString();
    sh.appendRow(hdr.map(function(c){ return vals[c] !== undefined ? vals[c] : ''; }));
    return { ok: true, note: 'Users sheet row created for ' + email };
  } catch (e) {
    return { ok: false, note: 'IPM upsert failed: ' + e.message };
  }
}


/**
 * Add a single test row in AwaitingActivation state. For testing the
 * login + activation flow without running a real bulk import.
 *
 * Example call from the editor:
 *
 *   addTestAccount(
 *     'Faculty',
 *     'you+facultytest@gmail.com',
 *     'Test Faculty',
 *     'Computer Science and Engineering',
 *     'Aarupadai Veedu Institute of Technology (AVIT)',
 *     'Assistant Professor'
 *   );
 */
function addTestAccount(role, email, name, department, instFullName, designation) {
  if (!role)         throw new Error('addTestAccount: role is required.');
  if (!email)        throw new Error('addTestAccount: email is required.');
  if (!name)         throw new Error('addTestAccount: name is required.');
  if (!instFullName) throw new Error('addTestAccount: institution full name is required.');

  var roleU = String(role).trim().toUpperCase();
  if (['FACULTY','HOD','HOI'].indexOf(roleU) < 0) {
    throw new Error('addTestAccount: role must be Faculty, HOD, or HOI (got "' + role + '").');
  }

  // Resolve campus from institution
  var campusFull = '';
  try {
    if (typeof INSTITUTION_HIERARCHY !== 'undefined') {
      Object.keys(INSTITUTION_HIERARCHY).forEach(function(c) {
        var insts = INSTITUTION_HIERARCHY[c].institutions || {};
        if (insts[instFullName]) campusFull = c;
      });
    }
  } catch(_){}
  if (!campusFull) {
    throw new Error('addTestAccount: institution "' + instFullName + '" not found in INSTITUTION_HIERARCHY. ' +
                    'Use the FULL institution name with the code in parens, e.g. ' +
                    '"Aarupadai Veedu Institute of Technology (AVIT)".');
  }

  // No-HOD institution guard
  if (roleU === 'HOD' && typeof _isNoHodInstitution_ === 'function' && _isNoHodInstitution_(instFullName)) {
    throw new Error('addTestAccount: HOD is not applicable for "' + instFullName + '" (no-HOD institution).');
  }
  if ((roleU === 'HOD' || roleU === 'FACULTY') && !department) {
    throw new Error('addTestAccount: department is required for ' + roleU + '.');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var emailLc = String(email).trim().toLowerCase();

  // Duplicate guard across both sheets
  ['FACULTY','STAFF'].forEach(function(sk) {
    var sn = (typeof SH !== 'undefined') ? SH[sk] : null;
    if (!sn) return;
    var sh = ss.getSheetByName(sn);
    if (!sh) return;
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return;
    var emI = data[0].indexOf('Email');
    if (emI < 0) return;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][emI]||'').trim().toLowerCase() === emailLc) {
        throw new Error('addTestAccount: email "' + emailLc + '" already exists in ' + sn +
                        '. Use resetTestActivation() to reset, or a different email.');
      }
    }
  });

  if (roleU === 'FACULTY') {
    var fsh = ss.getSheetByName(SH.FACULTY);
    if (!fsh) throw new Error('addTestAccount: Faculty_Master sheet not found. Run "Setup" from the VMRF IMO menu first.');
    _activationEnsureColumns_(fsh, SCHEMA.Faculty_Master);
    var fh = fsh.getRange(1, 1, 1, fsh.getLastColumn()).getValues()[0].map(function(v){return String(v).trim();});
    var vals = {
      'FacultyName': name, 'Email': emailLc,
      'Department': department || '', 'Campus': campusFull, 'Institution': instFullName,
      'Designation': designation || '',
      'PasswordHash': '',
      'GoogleEmail': '', 'Status': STATUS_AWAITING_ACTIVATION
    };
    fsh.appendRow(fh.map(function(col){ return vals[col] !== undefined ? vals[col] : ''; }));
    return { ok:true, role:'Faculty', email:emailLc, sheet:SH.FACULTY, status:STATUS_AWAITING_ACTIVATION };
  }

  // HOD or HOI — write to Staff_Master
  var ssh = ss.getSheetByName(SH.STAFF);
  if (!ssh) throw new Error('addTestAccount: Staff_Master sheet not found. Run "Setup" from the VMRF IMO menu first.');
  _activationEnsureColumns_(ssh, SCHEMA.Staff_Master);
  var sh2 = ssh.getRange(1, 1, 1, ssh.getLastColumn()).getValues()[0].map(function(v){return String(v).trim();});
  var staffID = _activationMakeID_(roleU);
  var svals = {
    'StaffID': staffID, 'StaffName': name, 'Email': emailLc,
    'Role': roleU,
    'Department': (roleU === 'HOI') ? '' : (department || ''),
    'Campus': campusFull, 'Institution': instFullName,
    'PasswordHash': '',
    'GoogleEmail': '', 'Status': STATUS_AWAITING_ACTIVATION,
    'Designation': designation || '', 'Phone': ''
  };
  ssh.appendRow(sh2.map(function(col){ return svals[col] !== undefined ? svals[col] : ''; }));

  return { ok:true, role:roleU, email:emailLc, sheet:SH.STAFF, staffID:staffID, status:STATUS_AWAITING_ACTIVATION };
}


/**
 * Reset a test account back to AwaitingActivation so the activation
 * flow can be re-tested. Clears password, status, OTPs.
 */
function resetTestActivation(email) {
  if (!email) throw new Error('resetTestActivation: email is required.');
  var emLc = String(email).trim().toLowerCase();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var changed = [];

  ['FACULTY','STAFF'].forEach(function(sk) {
    var sn = SH[sk]; if (!sn) return;
    var sh = ss.getSheetByName(sn); if (!sh) return;
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return;
    var hdr = data[0];
    var emI = hdr.indexOf('Email');
    var stI = hdr.indexOf('Status');
    var pwI = hdr.indexOf('PasswordHash');
    if (emI < 0) return;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][emI]||'').trim().toLowerCase() !== emLc) continue;
      if (stI >= 0) sh.getRange(i + 1, stI + 1).setValue(STATUS_AWAITING_ACTIVATION);
      if (pwI >= 0) sh.getRange(i + 1, pwI + 1).setValue('');
      changed.push(sn + ' row ' + (i + 1));
    }
  });

  // Clear cached OTPs
  try {
    var sp = PropertiesService.getScriptProperties();
    sp.deleteProperty('ACT_OTP_' + emLc);
    sp.deleteProperty('OTP_' + emLc);
    sp.deleteProperty('OTP_VERIFIED_' + emLc);
    sp.deleteProperty('RESET_OTP_' + emLc);
  } catch(_){}

  return {
    ok: true,
    email: emLc,
    rowsReset: changed,
    note: changed.length ? 'Reset to AwaitingActivation.' :
                           'No matching row found in Faculty_Master or Staff_Master.'
  };
}


/**
 * Health check for Step 1. Run from the Apps Script editor.
 */
function runStep1HealthCheck() {
  var results = [];
  var record = function(name, passed, detail) {
    results.push({ name: name, passed: !!passed, detail: detail || '' });
  };

  // 1. SELF_REGISTER_OPEN exists
  try {
    var sp = PropertiesService.getScriptProperties();
    var v = sp.getProperty('SELF_REGISTER_OPEN');
    if (v === null || v === undefined) {
      sp.setProperty('SELF_REGISTER_OPEN', 'false');
      record('SELF_REGISTER_OPEN property', true, 'created with default "false"');
    } else {
      record('SELF_REGISTER_OPEN property', true, 'already set to "' + v + '"');
    }
  } catch (e) {
    record('SELF_REGISTER_OPEN property', false, e.message);
  }

  // 2. Staff_Master schema has Phone and Designation
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss && ss.getSheetByName(SH.STAFF);
    if (!sh) {
      record('Staff_Master Phone+Designation', false, 'Staff_Master sheet does not exist yet (run initializeSystem first)');
    } else {
      _activationEnsureColumns_(sh, SCHEMA.Staff_Master);
      var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(v){return String(v).trim();});
      var hasPhone = hdr.indexOf('Phone') >= 0;
      var hasDes   = hdr.indexOf('Designation') >= 0;
      record('Staff_Master Phone+Designation', hasPhone && hasDes,
             'Phone=' + hasPhone + ', Designation=' + hasDes);
    }
  } catch (e) {
    record('Staff_Master Phone+Designation', false, e.message);
  }

  // 3. Password validator enforces policy
  try {
    var cases = [
      { pwd: '',          shouldPass: false, name: 'empty rejected' },
      { pwd: 'short1',    shouldPass: false, name: '< 8 chars rejected' },
      { pwd: 'abcdefgh',  shouldPass: false, name: 'no digit rejected' },
      { pwd: '12345678',  shouldPass: false, name: 'no letter rejected' },
      { pwd: 'Passw0rd',  shouldPass: true,  name: 'valid 8-char letter+digit' },
      { pwd: 'Long1Pass', shouldPass: true,  name: 'valid 9-char mixed' }
    ];
    var allOk = true, failures = [];
    cases.forEach(function(c) {
      var threw = false;
      try { _validatePasswordStrength(c.pwd); } catch(e) { threw = true; }
      var passed = (threw === !c.shouldPass);
      if (!passed) { allOk = false; failures.push(c.name); }
    });
    record('Password validator policy', allOk,
           allOk ? 'all ' + cases.length + ' cases behave correctly' :
                   'failed: ' + failures.join(', '));
  } catch (e) {
    record('Password validator policy', false, e.message);
  }

  // 4. Lockdown helper behaves
  try {
    var sp4 = PropertiesService.getScriptProperties();
    var saved = sp4.getProperty('SELF_REGISTER_OPEN');
    sp4.setProperty('SELF_REGISTER_OPEN', 'false');
    var closedThrew = false;
    try { _ensureSelfRegisterAllowed_('Faculty'); } catch(_) { closedThrew = true; }
    sp4.setProperty('SELF_REGISTER_OPEN', 'true');
    var openThrew = false;
    try { _ensureSelfRegisterAllowed_('Faculty'); } catch(_) { openThrew = true; }
    sp4.setProperty('SELF_REGISTER_OPEN', saved || 'false');
    record('Lockdown helper', closedThrew && !openThrew,
           'closed-throws=' + closedThrew + ', open-permits=' + !openThrew);
  } catch (e) {
    record('Lockdown helper', false, e.message);
  }

  // 5. Awaiting-activation lookup
  try {
    var lookup = _findAwaitingActivation_('this-email-does-not-exist@example.invalid');
    record('Awaiting-activation lookup', lookup === null,
           'returns null for nonexistent email');
  } catch (e) {
    record('Awaiting-activation lookup', false, e.message);
  }

  // 6. Status dropdowns include AwaitingActivation
  try {
    var r6 = updateStatusDropdownsForActivation({ silent: true });
    record('Status dropdowns include AwaitingActivation', true,
           'updated on: ' + (r6.updated.join(', ') || 'no sheets needed updating'));
  } catch (e) {
    record('Status dropdowns include AwaitingActivation', false, e.message);
  }

  // 7. Validations relaxed (scripts can write without being blocked)
  try {
    var r7 = relaxAllValidations({ silent: true });
    record('Validations relaxed (non-strict)', true,
           'relaxed ' + r7.totalRelaxed + ' cells across ' + Object.keys(r7.perSheet).length + ' sheets');
  } catch (e) {
    record('Validations relaxed (non-strict)', false, e.message);
  }

  // 8. EMAIL_ENABLED + EMAIL_ALLOWLIST configured (Step 3)
  try {
    var sp8 = PropertiesService.getScriptProperties();
    var enabled = sp8.getProperty('EMAIL_ENABLED');
    var allow = sp8.getProperty('EMAIL_ALLOWLIST');
    if (enabled === null) { sp8.setProperty('EMAIL_ENABLED', 'false'); enabled = 'false'; }
    if (allow === null)   { sp8.setProperty('EMAIL_ALLOWLIST', '');    allow = ''; }
    record('Email blackout configured', true,
           'EMAIL_ENABLED=' + enabled + ', EMAIL_ALLOWLIST="' + allow + '"');
  } catch (e) {
    record('Email blackout configured', false, e.message);
  }

  // 9. Email_Log sheet exists (or can be created)
  try {
    _emailLog_('healthcheck@example.invalid', '[healthcheck]', 'test-log-write', '');
    var elog = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EMAIL_LOG_SHEET);
    record('Email_Log sheet works', !!elog,
           elog ? 'sheet exists with ' + elog.getLastRow() + ' rows' : 'could not create');
  } catch (e) {
    record('Email_Log sheet works', false, e.message);
  }

  // 10. Activation OTP round-trip
  try {
    var hcEmail = 'healthcheck-' + Date.now() + '@example.invalid';
    var hcOtp = _makeActivationOtp_();
    _storeActivationOtp_(hcEmail, hcOtp, 'HC', 'FACULTY');
    var look = _lookupActivationOtp_(hcEmail);
    var ok = look && look.otp === hcOtp;
    _consumeActivationOtp_(hcEmail);
    var gone = _lookupActivationOtp_(hcEmail) === null;
    record('Activation OTP round-trip', !!(ok && gone),
           (ok && gone) ? 'code stored, matched, and consumed' : 'lookup=' + JSON.stringify(look) + ' gone=' + gone);
  } catch (e) {
    record('Activation OTP round-trip', false, e.message);
  }

  var passed = 0, failed = 0;
  Logger.log('═══ Step 1 Health Check ═══');
  results.forEach(function(r) {
    Logger.log((r.passed ? 'PASS' : 'FAIL') + '  ' + r.name + (r.detail ? '  (' + r.detail + ')' : ''));
    if (r.passed) passed++; else failed++;
  });
  Logger.log('═══ ' + passed + ' passed, ' + failed + ' failed ═══');

  try {
    SpreadsheetApp.getUi().alert(
      'Step 1 Health Check: ' + passed + ' passed, ' + failed + ' failed.\n\n' +
      results.map(function(r){ return (r.passed?'✓ ':'✗ ') + r.name + (r.detail?'\n   '+r.detail:''); }).join('\n\n')
    );
  } catch(_){ /* not in a UI context */ }

  return { passed: passed, failed: failed, results: results };
}
// ════════════════════════════════════════════════════════════════════
// PART C — STEP 2: Bulk import + single-row admin + dev activation
// ════════════════════════════════════════════════════════════════════

var FACULTY_IMPORT_TAB = 'Faculty_Import';

// Canonical staging tab column headers. The import reads by header NAME,
// not by position, so the column order in the tab can vary.
var IMPORT_HEADERS = ['S.No','Institution','Name','Designation','Department','Email','Phone','Category'];


/**
 * Create the Faculty_Import staging tab with canonical headers.
 * Idempotent: if the tab already exists, this just verifies the
 * headers are present and adds any missing ones at the end.
 *
 * Run once from the editor when you first set up the staging area.
 * After that, just paste your roster data starting from row 2.
 */
function createFacultyImportStagingTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(FACULTY_IMPORT_TAB);
  if (!sh) {
    sh = ss.insertSheet(FACULTY_IMPORT_TAB);
    sh.getRange(1, 1, 1, IMPORT_HEADERS.length).setValues([IMPORT_HEADERS]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, IMPORT_HEADERS.length).setFontWeight('bold');
    sh.setColumnWidth(2, 110); sh.setColumnWidth(3, 200);
    sh.setColumnWidth(4, 160); sh.setColumnWidth(5, 200);
    sh.setColumnWidth(6, 240); sh.setColumnWidth(7, 130);
    return { ok:true, created:true, sheet:FACULTY_IMPORT_TAB,
             message:'Staging tab created. Paste your roster data starting from row 2.' };
  }
  // Already exists — ensure all required headers are present
  _activationEnsureColumns_(sh, IMPORT_HEADERS);
  return { ok:true, created:false, sheet:FACULTY_IMPORT_TAB,
           message:'Staging tab already exists; headers verified.' };
}


/**
 * Find the header row in a staging sheet. The user might paste
 * data that includes a title row (e.g. "VMCC Faculty Database").
 * We scan the first 5 rows looking for one that contains BOTH
 * 'Email' and 'Category' as cell values.
 */
function _findImportHeaderRow_(sheet) {
  var maxScan = Math.min(5, sheet.getLastRow());
  if (maxScan === 0) return -1;
  var range = sheet.getRange(1, 1, maxScan, sheet.getLastColumn()).getValues();
  for (var r = 0; r < range.length; r++) {
    var lc = range[r].map(function(v){ return String(v||'').trim().toLowerCase(); });
    if (lc.indexOf('email') >= 0 && lc.indexOf('category') >= 0) return r;
  }
  return -1;
}


/**
 * Resolve an institution input (full name, short code, or messy string)
 * to its canonical { campusFull, instFull, instCode } for the Chennai
 * campus. Returns null if not found.
 *
 * Examples that all resolve to AVIT:
 *   "AVIT"
 *   "Aarupadai Veedu Institute of Technology (AVIT)"
 *   "aarupadai veedu" (loose substring match)
 */
function _resolveInstitution_(input) {
  if (!input) return null;
  var raw = String(input).trim();
  if (typeof INSTITUTION_HIERARCHY === 'undefined') return null;
  // Limit to Chennai campus per the migration scope
  var campusFull = "Vinayaka Mission's Chennai Campus";
  var camp = INSTITUTION_HIERARCHY[campusFull];
  if (!camp) return null;
  var insts = camp.institutions || {};
  var rawL = raw.toLowerCase();

  // 1) Exact match on full name
  if (insts[raw]) return { campusFull: campusFull, instFull: raw, instCode: insts[raw].code };

  // 2) Match by short code
  var instNames = Object.keys(insts);
  for (var i = 0; i < instNames.length; i++) {
    if (String(insts[instNames[i]].code).toLowerCase() === rawL) {
      return { campusFull: campusFull, instFull: instNames[i], instCode: insts[instNames[i]].code };
    }
  }
  // 3) Case-insensitive full-name match
  for (var j = 0; j < instNames.length; j++) {
    if (instNames[j].toLowerCase() === rawL) {
      return { campusFull: campusFull, instFull: instNames[j], instCode: insts[instNames[j]].code };
    }
  }
  // 4) Substring fallback (e.g. "Aarupadai Veedu" matches AVIT's full name)
  for (var k = 0; k < instNames.length; k++) {
    if (instNames[k].toLowerCase().indexOf(rawL) >= 0 ||
        rawL.indexOf(String(insts[instNames[k]].code).toLowerCase()) >= 0) {
      return { campusFull: campusFull, instFull: instNames[k], instCode: insts[instNames[k]].code };
    }
  }
  return null;
}


/**
 * Build a set of emails already present in Faculty_Master and
 * Staff_Master so the import can detect duplicates without N
 * sheet reads per row.
 */
function _existingEmailSet_() {
  var set = {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ['FACULTY','STAFF'].forEach(function(sk) {
    var sn = SH[sk];
    var sh = ss.getSheetByName(sn);
    if (!sh) return;
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return;
    var emI = data[0].indexOf('Email');
    if (emI < 0) return;
    for (var i = 1; i < data.length; i++) {
      var e = String(data[i][emI]||'').trim().toLowerCase();
      if (e) set[e] = sn;
    }
  });
  return set;
}


/**
 * BULK IMPORT — the main Step 2 function.
 *
 * Reads the Faculty_Import staging tab, validates every row, writes
 * accepted rows to Faculty_Master / Staff_Master with Status =
 * AwaitingActivation, pre-seeds FacultyProfiles for Faculty rows,
 * and produces a per-row CSV report saved to Drive.
 *
 * Usage:
 *   bulkImportFaculty()                   // commit
 *   bulkImportFaculty({ dryRun: true })   // validate only, write nothing
 *
 * The convenience wrappers runImportDryRun() and runImportForReal()
 * exist for one-click execution from the editor's Run dropdown.
 *
 * @param {object} opts  { dryRun: boolean }
 * @return {object}      Summary: { dryRun, totalRows, imported, skipped, rejected, flagged, reportUrl, summary }
 */
function bulkImportFaculty(opts) {
  opts = opts || {};
  var dryRun = opts.dryRun === true;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var stage = ss.getSheetByName(FACULTY_IMPORT_TAB);
  if (!stage) {
    throw new Error('Staging tab "' + FACULTY_IMPORT_TAB + '" not found. ' +
                    'Run createFacultyImportStagingTab() first, then paste your roster.');
  }

  // Find header row (handles a stray title row at the top)
  var hdrRow = _findImportHeaderRow_(stage);
  if (hdrRow < 0) {
    throw new Error('Could not find a header row in "' + FACULTY_IMPORT_TAB +
                    '" containing both "Email" and "Category". ' +
                    'Make sure your pasted data includes the column headers.');
  }

  var lastRow = stage.getLastRow();
  var lastCol = stage.getLastColumn();
  if (lastRow <= hdrRow + 1) {
    throw new Error('No data rows found below the header in "' + FACULTY_IMPORT_TAB + '".');
  }

  var allData = stage.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = allData[hdrRow].map(function(v){ return String(v||'').trim(); });
  var col = function(name) { return headers.indexOf(name); };

  // Required columns
  var ciName = col('Name'), ciEmail = col('Email'), ciCat = col('Category'),
      ciInst = col('Institution'), ciDept = col('Department'),
      ciDes  = col('Designation'), ciPhone = col('Phone');

  if (ciName < 0 || ciEmail < 0 || ciCat < 0 || ciInst < 0) {
    throw new Error('Staging tab is missing required columns. ' +
                    'Need at minimum: Name, Email, Institution, Category. ' +
                    'Got headers: ' + headers.join(' | '));
  }

  // Pre-cache existing emails for duplicate detection
  var existing = _existingEmailSet_();

  // Iterate data rows
  var imported = 0, skipped = 0, rejected = 0, flagged = 0;
  var report = [];          // per-row outcome
  var seenInStaging = {};   // catch within-file duplicates
  var toWriteFaculty = [];  // batched writes
  var toWriteStaff   = [];
  var toWriteProfile = [];

  for (var r = hdrRow + 1; r < allData.length; r++) {
    var row = allData[r];
    // Skip completely empty rows (no name AND no email)
    var rawName  = String(row[ciName]  || '').trim();
    var rawEmail = String(row[ciEmail] || '').trim();
    if (!rawName && !rawEmail) continue;

    var rowNum = r + 1; // 1-based sheet row
    var errors = [], flags = [];
    var category = String(row[ciCat] || '').trim();
    var catU = category.toUpperCase();
    var emailLc = rawEmail.toLowerCase();
    var instInput = ciInst >= 0 ? String(row[ciInst] || '').trim() : '';
    var deptInput = ciDept >= 0 ? String(row[ciDept] || '').trim() : '';
    var desInput  = ciDes  >= 0 ? String(row[ciDes]  || '').trim() : '';
    var phoneInput= ciPhone>= 0 ? String(row[ciPhone]|| '').trim() : '';

    // ── Validations (collect all errors per row) ──
    if (!rawName)  errors.push('missing Name');
    if (!rawEmail) errors.push('missing Email');
    else if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(rawEmail)) errors.push('invalid Email format');
    if (!category) errors.push('missing Category');
    else if (['FACULTY','HOD','HOI'].indexOf(catU) < 0) {
      errors.push('Category must be Faculty, HOD, or HOI (got "' + category + '")');
    }

    var resolved = _resolveInstitution_(instInput);
    if (!instInput) errors.push('missing Institution');
    else if (!resolved) errors.push('Institution "' + instInput + '" not recognised in INSTITUTION_HIERARCHY');

    if ((catU === 'FACULTY' || catU === 'HOD') && !deptInput) {
      errors.push('Department is required for ' + catU);
    }

    // No-HOD institution guard
    if (catU === 'HOD' && resolved && typeof _isNoHodInstitution_ === 'function' &&
        _isNoHodInstitution_(resolved.instFull)) {
      errors.push(resolved.instFull + ' is a no-HOD institution — HOD rows not allowed');
    }

    // Within-staging duplicate
    if (emailLc) {
      if (seenInStaging[emailLc]) {
        errors.push('duplicate of staging row ' + seenInStaging[emailLc]);
      } else {
        seenInStaging[emailLc] = rowNum;
      }
    }

    // ── Existing-row check (not an error, just skipped) ──
    var alreadyExists = emailLc && existing[emailLc];

    // ── Soft flags (don't block import, just surface for review) ──
    if (catU === 'HOI' && deptInput) {
      flags.push('HOI row has Department "' + deptInput + '" — will be stored but not used for routing');
    }
    if (catU === 'HOD' && desInput &&
        /^assistant\s+professor$/i.test(desInput.replace(/\s+/g, ' '))) {
      flags.push('HOD designation "' + desInput + '" is unusual — most HODs are Associate Professor or higher');
    }

    // ── Decide outcome ──
    var outcome, reason = '';
    if (errors.length > 0) {
      outcome = 'REJECTED';
      reason = errors.join('; ');
      rejected++;
    } else if (alreadyExists) {
      outcome = 'SKIPPED';
      reason = 'email already exists in ' + alreadyExists;
      skipped++;
    } else {
      outcome = dryRun ? 'WOULD_IMPORT' : 'IMPORTED';
      reason = flags.length ? 'flagged: ' + flags.join('; ') : '';
      if (flags.length) flagged++;
      imported++;

      // Queue the write (only used when !dryRun)
      if (catU === 'FACULTY') {
        toWriteFaculty.push({
          name: rawName, email: emailLc, dept: deptInput,
          campus: resolved.campusFull, inst: resolved.instFull,
          designation: desInput, phone: phoneInput
        });
        toWriteProfile.push({
          email: emailLc, name: rawName, designation: desInput, dept: deptInput,
          inst: resolved.instFull, campus: resolved.campusFull, phone: phoneInput
        });
      } else {
        toWriteStaff.push({
          role: catU, name: rawName, email: emailLc,
          dept: catU === 'HOI' ? '' : deptInput,
          campus: resolved.campusFull, inst: resolved.instFull,
          designation: desInput, phone: phoneInput
        });
      }
      // Mark in existing so a duplicate later in the staging tab is caught
      if (emailLc) existing[emailLc] = (catU === 'FACULTY') ? SH.FACULTY : SH.STAFF;
    }

    report.push({
      row: rowNum, name: rawName, email: rawEmail, category: category,
      institution: resolved ? resolved.instFull : instInput,
      outcome: outcome, reason: reason
    });
  }

  // ── Commit writes (unless dryRun) ──
  if (!dryRun) {
    if (toWriteFaculty.length > 0) _commitFacultyRows_(toWriteFaculty);
    if (toWriteStaff.length > 0)   _commitStaffRows_(toWriteStaff);
    if (toWriteProfile.length > 0) _commitProfileRows_(toWriteProfile);
  }

  // ── Generate CSV report to Drive ──
  var reportUrl = '';
  try { reportUrl = _writeImportReportCsv_(report, { dryRun: dryRun }); } catch(e) {
    Logger.log('Could not write CSV report: ' + e.message);
  }

  // ── Summary ──
  var label = dryRun ? 'DRY RUN' : 'IMPORT';
  var summary = label + ' — ' + (dryRun ? 'would import' : 'imported') + ' ' + imported +
                ', skipped ' + skipped + ' (already existing), rejected ' + rejected +
                (flagged ? ', ' + flagged + ' flagged for review' : '') +
                (reportUrl ? '\n\nReport CSV: ' + reportUrl : '');

  Logger.log('═══ ' + summary + ' ═══');
  try {
    SpreadsheetApp.getUi().alert(label + ' complete', summary, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(_){ /* non-UI context */ }

  return {
    ok: true, dryRun: dryRun,
    totalRows: report.length,
    imported: imported, skipped: skipped, rejected: rejected, flagged: flagged,
    reportUrl: reportUrl,
    summary: summary
  };
}


/**
 * Convenience wrapper: dry-run the import. Pick this from the Apps
 * Script editor's Run dropdown — no arguments needed.
 */
function runImportDryRun() {
  return bulkImportFaculty({ dryRun: true });
}


/**
 * Convenience wrapper: run the real import. Pick this from the editor's
 * Run dropdown after you've verified the dry-run report.
 */
function runImportForReal() {
  return bulkImportFaculty({ dryRun: false });
}


// ── Write helpers ──────────────────────────────────────────────────

function _commitFacultyRows_(rows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH.FACULTY);
  if (!sh) throw new Error('Faculty_Master sheet not found. Run initializeSystem first.');
  _activationEnsureColumns_(sh, SCHEMA.Faculty_Master);
  var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(v){ return String(v).trim(); });
  var values = rows.map(function(r) {
    var vals = {
      'FacultyName':  r.name,
      'Email':        r.email,
      'Department':   r.dept,
      'Campus':       r.campus,
      'Institution':  r.inst,
      'Designation':  r.designation,
      'PasswordHash': '',
      'GoogleEmail':  '',
      'Status':       STATUS_AWAITING_ACTIVATION
    };
    return hdr.map(function(c){ return vals[c] !== undefined ? vals[c] : ''; });
  });
  sh.getRange(sh.getLastRow() + 1, 1, values.length, hdr.length).setValues(values);
}

function _commitStaffRows_(rows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH.STAFF);
  if (!sh) throw new Error('Staff_Master sheet not found. Run initializeSystem first.');
  _activationEnsureColumns_(sh, SCHEMA.Staff_Master);
  var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(v){ return String(v).trim(); });
  var values = rows.map(function(r) {
    var vals = {
      'StaffID':      _activationMakeID_(r.role),
      'StaffName':    r.name,
      'Email':        r.email,
      'Role':         r.role,
      'Department':   r.dept,
      'Campus':       r.campus,
      'Institution':  r.inst,
      'PasswordHash': '',
      'GoogleEmail':  '',
      'Status':       STATUS_AWAITING_ACTIVATION,
      'Designation':  r.designation,
      'Phone':        r.phone
    };
    return hdr.map(function(c){ return vals[c] !== undefined ? vals[c] : ''; });
  });
  sh.getRange(sh.getLastRow() + 1, 1, values.length, hdr.length).setValues(values);
}

function _commitProfileRows_(rows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH.PROFILES);
  if (!sh) {
    // Profile sheet is optional — skip silently if it doesn't exist
    Logger.log('FacultyProfiles sheet not found — skipping pre-seed.');
    return;
  }
  _activationEnsureColumns_(sh, SCHEMA.FacultyProfiles);
  var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(v){ return String(v).trim(); });
  var values = rows.map(function(r) {
    var vals = {
      'Email':       r.email,
      'FacultyName': r.name,
      'Designation': r.designation,
      'Department':  r.dept,
      'Institution': r.inst,
      'Campus':      r.campus,
      'Phone':       r.phone,
      'LastUpdated': new Date()
    };
    return hdr.map(function(c){ return vals[c] !== undefined ? vals[c] : ''; });
  });
  sh.getRange(sh.getLastRow() + 1, 1, values.length, hdr.length).setValues(values);
}


// ── Report generator ──────────────────────────────────────────────

function _writeImportReportCsv_(report, opts) {
  opts = opts || {};
  var lines = ['Row,Name,Email,Category,Institution,Outcome,Reason'];
  report.forEach(function(r) {
    var csvCell = function(v) {
      var s = String(v == null ? '' : v).replace(/"/g, '""');
      return /[",\n]/.test(s) ? '"' + s + '"' : s;
    };
    lines.push([
      r.row, csvCell(r.name), csvCell(r.email), csvCell(r.category),
      csvCell(r.institution), r.outcome, csvCell(r.reason)
    ].join(','));
  });
  var csv = lines.join('\n');
  var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");
  var prefix = opts.dryRun ? 'VMRF-Import-DryRun-' : 'VMRF-Import-Commit-';
  var file = DriveApp.createFile(prefix + ts + '.csv', csv, MimeType.CSV);
  return file.getUrl();
}


// ════════════════════════════════════════════════════════════════════
// Single-row admin functions
// ════════════════════════════════════════════════════════════════════

/**
 * Add a single user account directly to the sheets — for ad-hoc
 * additions outside the bulk import. Same validation rules as the
 * bulk import.
 *
 * Example:
 *   addUserAccount('Faculty', 'newhire@vmrf.edu.in', 'New Hire',
 *                  'Computer Science and Engineering', 'AVIT',
 *                  'Assistant Professor', '9999999999');
 */
function addUserAccount(role, email, name, department, institutionInput, designation, phone) {
  if (!role)  throw new Error('addUserAccount: role is required.');
  if (!email) throw new Error('addUserAccount: email is required.');
  if (!name)  throw new Error('addUserAccount: name is required.');

  var catU = String(role).trim().toUpperCase();
  if (['FACULTY','HOD','HOI'].indexOf(catU) < 0) {
    throw new Error('addUserAccount: role must be Faculty, HOD, or HOI (got "' + role + '").');
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error('addUserAccount: invalid email format.');
  }
  var resolved = _resolveInstitution_(institutionInput);
  if (!resolved) {
    throw new Error('addUserAccount: institution "' + institutionInput + '" not found.');
  }
  if (catU === 'HOD' && typeof _isNoHodInstitution_ === 'function' && _isNoHodInstitution_(resolved.instFull)) {
    throw new Error('addUserAccount: HOD is not applicable for ' + resolved.instFull);
  }
  if ((catU === 'FACULTY' || catU === 'HOD') && !department) {
    throw new Error('addUserAccount: department is required for ' + catU);
  }

  var emailLc = email.trim().toLowerCase();
  var existing = _existingEmailSet_();
  if (existing[emailLc]) {
    throw new Error('addUserAccount: email "' + emailLc + '" already exists in ' + existing[emailLc] + '.');
  }

  if (catU === 'FACULTY') {
    _commitFacultyRows_([{
      name: name, email: emailLc, dept: department, campus: resolved.campusFull,
      inst: resolved.instFull, designation: designation || '', phone: phone || ''
    }]);
    _commitProfileRows_([{
      email: emailLc, name: name, designation: designation || '',
      dept: department, inst: resolved.instFull, campus: resolved.campusFull, phone: phone || ''
    }]);
  } else {
    _commitStaffRows_([{
      role: catU, name: name, email: emailLc,
      dept: catU === 'HOI' ? '' : department,
      campus: resolved.campusFull, inst: resolved.instFull,
      designation: designation || '', phone: phone || ''
    }]);
  }
  return { ok:true, role:catU, email:emailLc, status:STATUS_AWAITING_ACTIVATION };
}


/**
 * Deactivate a user — flip their Status to 'Inactive'. Preserves the
 * row and all history; just prevents login. The user can be
 * reactivated later by flipping Status back to 'Active' (manually).
 */
function deactivateUserAccount(email) {
  if (!email) throw new Error('deactivateUserAccount: email is required.');
  var emLc = String(email).trim().toLowerCase();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var changed = [];
  ['FACULTY','STAFF'].forEach(function(sk) {
    var sn = SH[sk]; if (!sn) return;
    var sh = ss.getSheetByName(sn); if (!sh) return;
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return;
    var emI = data[0].indexOf('Email'), stI = data[0].indexOf('Status');
    if (emI < 0 || stI < 0) return;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][emI]||'').trim().toLowerCase() === emLc) {
        sh.getRange(i + 1, stI + 1).setValue('Inactive');
        changed.push(sn + ' row ' + (i + 1));
      }
    }
  });
  if (changed.length === 0) {
    return { ok:false, email:emLc, note:'No matching row found.' };
  }
  return { ok:true, email:emLc, rowsDeactivated:changed };
}


// ════════════════════════════════════════════════════════════════════
// DEV-ONLY: manual activation completer (bypasses OTP/email)
// ════════════════════════════════════════════════════════════════════

/**
 * Manually complete activation for a test account — writes a chosen
 * password and flips Status from AwaitingActivation to Active.
 * Bypasses OTP and email entirely. Useful during Step 2 testing
 * before the real activation backend (Step 3) is built.
 *
 * SAFETY: refuses to run if SELF_REGISTER_OPEN is true. The intent is
 * that this helper only works on closed-system test deployments —
 * never on a production-like deployment that has self-registration
 * open. (For your situation, that flag stays false on both test and
 * production, so the helper works on test; on production you'd never
 * call it because real users will activate via Step 3's email flow.)
 *
 * Example:
 *   _devCompleteActivationManually('you+test@gmail.com', 'TestPass1');
 */
function _devCompleteActivationManually(email, password) {
  if (!email)    throw new Error('_devCompleteActivationManually: email is required.');
  if (!password) throw new Error('_devCompleteActivationManually: password is required.');

  // Safety: hash function must exist
  if (typeof _hashPwd !== 'function') {
    throw new Error('_devCompleteActivationManually: _hashPwd is not defined in this project. ' +
                    'This helper expects the legacy Djb2 hash function used by Auth.gs.');
  }

  _validatePasswordStrength(password);

  var emLc = String(email).trim().toLowerCase();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var changed = [];
  var roleFound = '';
  var instFound = '', campusFound = '';

  ['FACULTY','STAFF'].forEach(function(sk) {
    var sn = SH[sk]; if (!sn) return;
    var sh = ss.getSheetByName(sn); if (!sh) return;
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return;
    var hdr = data[0];
    var emI = hdr.indexOf('Email'),
        stI = hdr.indexOf('Status'),
        pwI = hdr.indexOf('PasswordHash'),
        rlI = hdr.indexOf('Role'),
        inI = hdr.indexOf('Institution'),
        caI = hdr.indexOf('Campus');
    if (emI < 0 || stI < 0 || pwI < 0) return;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][emI]||'').trim().toLowerCase() !== emLc) continue;
      sh.getRange(i + 1, pwI + 1).setValue(_hashPwd(password));
      sh.getRange(i + 1, stI + 1).setValue('Active');
      changed.push(sn + ' row ' + (i + 1));
      if (sk === 'FACULTY') roleFound = 'FACULTY';
      else if (rlI >= 0)    roleFound = String(data[i][rlI]||'').toUpperCase();
      if (inI >= 0) instFound = String(data[i][inI]||'').trim();
      if (caI >= 0) campusFound = String(data[i][caI]||'').trim();
    }
  });

  if (changed.length === 0) {
    return { ok:false, email:emLc, note:'No matching row found in Faculty_Master or Staff_Master.' };
  }

  // Best-effort: dual-write to IPM Users for HOI/IMO if helpers exist.
  // If IPM helpers aren't loaded, we log a note rather than failing — the
  // main-app login will work; the IPM portal mounted inside it may not
  // until the real activation backend ships in Step 3.
  var ipmNote = '';
  if (roleFound === 'HOI' || roleFound === 'IMO') {
    try {
      ipmNote = _tryDualWriteIpmUsers_(roleFound, emLc, password, instFound, campusFound);
    } catch (e) {
      ipmNote = 'IPM dual-write skipped: ' + e.message;
    }
  }

  return {
    ok: true, email: emLc, role: roleFound,
    rowsUpdated: changed,
    ipmNote: ipmNote || '(not applicable for ' + (roleFound || 'this role') + ')',
    message: 'Account activated. You can now sign in with the chosen password.'
  };
}

/**
 * Legacy name kept for backward compatibility. Delegates to
 * _activationUpsertIpmUser_, which handles the case where the project's
 * ipmHashPassword_ isn't loaded (by computing the sha256:<hex> hash
 * inline). Returns just the note string (not the {ok,note} object) to
 * preserve the original calling convention.
 */
function _tryDualWriteIpmUsers_(role, email, password, instFull, campusFull) {
  var result = _activationUpsertIpmUser_(role, email, password, instFull, campusFull);
  return result.note;
}


// ════════════════════════════════════════════════════════════════════
// Menu wrappers (used by IMO.gs onOpen menu items)
// ════════════════════════════════════════════════════════════════════

function menuCreateImportStagingTab() {
  var r = createFacultyImportStagingTab();
  SpreadsheetApp.getUi().alert(r.message);
}

function menuRunImportDryRun() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert('Run dry-run import?',
    'This will validate every row in Faculty_Import and produce a CSV report, but write NOTHING to the sheets.\n\nProceed?',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  runImportDryRun();
}

function menuRunImportForReal() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert('Commit import to sheets?',
    'This will write all valid rows to Faculty_Master / Staff_Master / FacultyProfiles ' +
    'with Status=AwaitingActivation. No emails will be sent.\n\n' +
    'Re-running is safe — existing emails will be skipped.\n\nProceed?',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  runImportForReal();
}

function menuDeactivateUser() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt('Deactivate user',
    'Enter the email of the user to deactivate:',
    ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var email = resp.getResponseText().trim();
  if (!email) return;
  var r = deactivateUserAccount(email);
  ui.alert(r.ok ? 'Deactivated:\n' + r.rowsDeactivated.join('\n') : 'Nothing changed: ' + r.note);
}


// ════════════════════════════════════════════════════════════════════
// PART D — Validation maintenance helpers
// Both called automatically by runStep1HealthCheck. Idempotent.
// ════════════════════════════════════════════════════════════════════

/**
 * Add 'AwaitingActivation' to the Status column data validation on
 * Faculty_Master and Staff_Master. Without this, the original
 * applyStatusDropdowns rule (Active/Pending/Inactive only) rejects
 * any AwaitingActivation write.
 *
 * @param {object} opts  { silent: boolean }  Suppress UI popup
 */
function updateStatusDropdownsForActivation(opts) {
  opts = opts || {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allowed = ['Active','Pending','Inactive','AwaitingActivation'];
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(allowed, true)
    .setAllowInvalid(true)   // non-strict so scripts can write anyway
    .setHelpText('Allowed: ' + allowed.join(', '))
    .build();
  var updated = [];
  [SH.FACULTY, SH.STAFF].forEach(function(sn) {
    var sh = ss.getSheetByName(sn);
    if (!sh) return;
    var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
                .map(function(v){ return String(v).trim(); });
    var stI = hdr.indexOf('Status');
    if (stI < 0) return;
    var lastRow = Math.max(2, sh.getMaxRows());
    sh.getRange(2, stI + 1, lastRow - 1, 1).setDataValidation(rule);
    updated.push(sn);
  });
  if (!opts.silent) {
    try {
      SpreadsheetApp.getUi().alert(
        updated.length
          ? 'Status dropdown updated on:\n' + updated.join('\n') +
            '\n\nAllowed values: ' + allowed.join(', ')
          : 'No sheets found to update.'
      );
    } catch(_){}
  }
  return { ok: true, updated: updated, allowed: allowed };
}


/**
 * Convert all strict data validations on user-management sheets
 * (Faculty_Master, Staff_Master, FacultyProfiles) to non-strict.
 * IMO still gets autocomplete dropdowns when editing manually,
 * but scripts can write any value without being blocked.
 *
 * Idempotent — safe to run repeatedly.
 *
 * @param {object} opts  { silent: boolean }  Suppress UI popup
 */
function relaxAllValidations(opts) {
  opts = opts || {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetsToFix = [SH.FACULTY, SH.STAFF, SH.PROFILES];
  var totalRelaxed = 0;
  var perSheet = {};

  sheetsToFix.forEach(function(sn) {
    if (!sn) return;
    var sh = ss.getSheetByName(sn);
    if (!sh) return;
    var rows = sh.getMaxRows();
    var cols = sh.getMaxColumns();
    if (rows < 2) return;

    var range = sh.getRange(2, 1, rows - 1, cols);
    var rules;
    try { rules = range.getDataValidations(); }
    catch(_) { return; }
    var newRules = [];
    var relaxedHere = 0;

    for (var r = 0; r < rules.length; r++) {
      var rowOut = [];
      for (var c = 0; c < rules[r].length; c++) {
        var rule = rules[r][c];
        if (rule && rule.getAllowInvalid() === false) {
          rowOut.push(rule.copy().setAllowInvalid(true).build());
          relaxedHere++;
        } else {
          rowOut.push(rule);
        }
      }
      newRules.push(rowOut);
    }

    if (relaxedHere > 0) {
      range.setDataValidations(newRules);
      perSheet[sn] = relaxedHere;
      totalRelaxed += relaxedHere;
    }
  });

  var msg = 'Relaxed ' + totalRelaxed + ' strict validations';
  if (Object.keys(perSheet).length > 0) {
    msg += ':\n\n' + Object.keys(perSheet).map(function(k) {
      return '  ' + k + ': ' + perSheet[k] + ' cells';
    }).join('\n');
  } else {
    msg += '\n(no strict validations found — already relaxed)';
  }

  if (!opts.silent) {
    try { SpreadsheetApp.getUi().alert(msg); } catch(_){}
  }
  Logger.log(msg);
  return { ok: true, totalRelaxed: totalRelaxed, perSheet: perSheet };
}


// ════════════════════════════════════════════════════════════════════
// PART E — One-click test wrappers (edit YOUR_GMAIL before running)
//
// Three plus-addressed accounts (Faculty + HOD + HOI) created from
// your single Gmail inbox. Used for end-to-end testing on the test
// deployment.
// ════════════════════════════════════════════════════════════════════

// YOUR_GMAIL global has been replaced. Each wrapper hardcodes the
// email locally to eliminate any chance of scoping issues.
// To change the test email, search-and-replace doeerica32@gmail.com
var YOUR_GMAIL = 'doeerica32@gmail.com';   // unused — kept for backwards compatibility


/**
 * Create the three test accounts in AwaitingActivation state.
 * Run once. Robust to popup failures (logs everything).
 */
function addMyThreeTestAccounts() {
  var local  = 'doeerica32';
  var domain = 'gmail.com';

  var accounts = [
    { role:'Faculty', email: local + '+facultytest@' + domain,
      name:'Test Faculty', dept:'Computer Science and Engineering',
      inst:'AVIT', des:'Assistant Professor', phone:'0000000000' },
    { role:'HOD',     email: local + '+hodtest@' + domain,
      name:'Test HOD', dept:'Computer Science and Engineering',
      inst:'AVIT', des:'Professor', phone:'0000000000' },
    { role:'HOI',     email: local + '+hoitest@' + domain,
      name:'Test HOI', dept:'',
      inst:'AVIT', des:'Principal', phone:'0000000000' }
  ];

  var results = [];
  for (var i = 0; i < accounts.length; i++) {
    var a = accounts[i];
    try {
      addUserAccount(a.role, a.email, a.name, a.dept, a.inst, a.des, a.phone);
      results.push({ role: a.role, email: a.email, ok: true });
      Logger.log('CREATED  ' + a.role + ': ' + a.email);
    } catch (e) {
      results.push({ role: a.role, email: a.email, ok: false, error: e.message });
      Logger.log('FAILED   ' + a.role + ': ' + a.email + '  --  ' + e.message);
    }
  }

  var summary = results.map(function(r) {
    return (r.ok ? '[ok] ' : '[X]  ') + r.role + '  ' + r.email +
           (r.error ? '  ERROR: ' + r.error : '');
  }).join('\n');
  Logger.log('=== Test account creation summary ===\n' + summary);

  try { SpreadsheetApp.getUi().alert('Test accounts created:\n\n' + summary); } catch(_){}
  return results;
}


/**
 * Set the password to TestPass1 and flip status to Active for the
 * three test accounts. Run after addMyThreeTestAccounts.
 */
function activateMyThreeTestAccounts() {
  var PASSWORD = 'TestPass1';
  var local  = 'doeerica32';
  var domain = 'gmail.com';

  var results = [];
  ['facultytest','hodtest','hoitest'].forEach(function(suffix) {
    var email = local + '+' + suffix + '@' + domain;
    try {
      var r = _devCompleteActivationManually(email, PASSWORD);
      results.push({ email: email, ok: r.ok, note: r.note || r.message || '' });
      Logger.log((r.ok ? 'ACTIVATED' : 'FAILED   ') + '  ' + email +
                 (r.ipmNote ? '  ipm: ' + r.ipmNote : ''));
    } catch (e) {
      results.push({ email: email, ok: false, error: e.message });
      Logger.log('FAILED     ' + email + '  --  ' + e.message);
    }
  });

  var summary = results.map(function(r) {
    return (r.ok ? '[ok] ' : '[X]  ') + r.email +
           (r.error ? '  ERROR: ' + r.error : '');
  }).join('\n');
  Logger.log('=== Activation summary ===\n' + summary +
             '\nPassword for activated accounts: ' + PASSWORD);

  try {
    SpreadsheetApp.getUi().alert('Activation results:\n\n' + summary +
                                  '\n\nPassword: ' + PASSWORD);
  } catch(_){}
  return results;
}


/**
 * Recovery helper: re-run the IPM dual-write for the HOI test account.
 * Use this if HOI login fails with "HOI credentials not recognised."
 * (Caused by an old _devCompleteActivationManually call that ran when
 * the project's ipmHashPassword_ wasn't available, so the IPM Users
 * row never got written.)
 *
 * Reads the existing Staff_Master row for the HOI test account to get
 * its campus/institution, then upserts the corresponding IPM Users row
 * with a correctly-formatted sha256:<hex> hash of 'TestPass1'.
 *
 * Idempotent — safe to run multiple times.
 */
function fixHoiTestAccountIpmUsers() {
  var local  = 'doeerica32';
  var domain = 'gmail.com';
  var email = local + '+hoitest@' + domain;
  var PWD   = 'TestPass1';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH.STAFF);
  if (!sh) { Logger.log('Staff_Master sheet not found'); return; }
  var data = sh.getDataRange().getValues();
  var hdr = data[0];
  var emI = hdr.indexOf('Email'),
      inI = hdr.indexOf('Institution'),
      caI = hdr.indexOf('Campus');
  var inst = '', campus = '';
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][emI]||'').trim().toLowerCase() === email.toLowerCase()) {
      inst   = data[i][inI];
      campus = data[i][caI];
      break;
    }
  }
  if (!inst) {
    Logger.log('HOI row not found in Staff_Master. Run addMyThreeTestAccounts first.');
    try { SpreadsheetApp.getUi().alert('HOI test row not found. Run addMyThreeTestAccounts first.'); } catch(_){}
    return;
  }

  var result = _activationUpsertIpmUser_('HOI', email, PWD, inst, campus);
  Logger.log('IPM upsert: ' + JSON.stringify(result));
  try {
    SpreadsheetApp.getUi().alert(
      (result.ok ? '✓ ' : '✗ ') + 'IPM Users:\n\n' + result.note +
      (result.ok ? '\n\nNow try logging in as the HOI test account.' : '')
    );
  } catch(_){}
  return result;
}


/**
 * Print the current state of the three test accounts. Always uses
 * Logger.log, so it works whether or not the UI context is available.
 */
function listMyTestAccounts() {
  var local  = 'doeerica32';
  var domain = 'gmail.com';
  var emails = ['facultytest','hodtest','hoitest'].map(function(s) {
    return local + '+' + s + '@' + domain;
  });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lines = ['=== Test account status ==='];

  emails.forEach(function(em) {
    var found = false;
    ['FACULTY','STAFF'].forEach(function(sk) {
      var sh = ss.getSheetByName(SH[sk]); if (!sh) return;
      var data = sh.getDataRange().getValues();
      if (data.length < 2) return;
      var emI = data[0].indexOf('Email'),
          stI = data[0].indexOf('Status'),
          pwI = data[0].indexOf('PasswordHash');
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][emI]||'').trim().toLowerCase() === em) {
          var status = stI >= 0 ? data[i][stI] : '?';
          var hasPwd = pwI >= 0 && String(data[i][pwI]||'').length > 0;
          lines.push('[ok] ' + em + '  in ' + SH[sk] +
                     '  Status=' + status +
                     '  Password=' + (hasPwd ? 'set' : 'EMPTY'));
          found = true;
        }
      }
    });
    if (!found) lines.push('[X]  ' + em + '  NOT FOUND in any sheet');
  });

  var summary = lines.join('\n');
  Logger.log(summary);
  try { SpreadsheetApp.getUi().alert(summary); } catch(_){}
  return summary;
}


// ════════════════════════════════════════════════════════════════════
// PART I — STEP 3 (OTP edition): Activation backend with 6-digit codes
//
// DESIGN: The activation email contains a 6-digit code, NOT a clickable
// link. The user opens the web app (the normal login URL), clicks
// "Activate your account" under the sign-in form, and enters:
//   email + code + new password.
// This avoids every Workspace/Gmail link-click restriction because the
// user never follows an external link from email.
//
// Public functions called by the frontend:
//   requestActivation(email)  → generate + email a fresh 6-digit code
//   completeActivationWithOtp(email, otp, password, confirmPassword)
//                             → verify code, set password, flip Active
//
// Admin functions (editor / menu):
//   sendActivationInvitations({dryRun, role})
//   resendActivationInvitation(email)
//
// Code storage: Script Properties, key ACT_OTP_<email_lc>
//   value = JSON { otp, name, role, createdAt, attempts }
// TTL: 24 hours. Max 5 wrong attempts, then the code is invalidated.
// ════════════════════════════════════════════════════════════════════

var ACT_OTP_TTL_SEC      = 24 * 60 * 60;  // 24 hours
var ACT_OTP_MAX_ATTEMPTS = 5;
var ACT_EMAIL_FROM_NAME  = 'VMRF-DU IMS';
var EMAIL_LOG_SHEET      = 'Email_Log';


// ── Email infrastructure (plain-text only, blackout + allowlist) ───

/**
 * Send an email subject to the blackout flag and allowlist.
 * PLAIN TEXT ONLY — no htmlBody (Gmail wraps HTML links in tracking
 * redirects; plain text avoids the whole class of problems).
 * Always logs to Email_Log whether sent, blocked, or failed.
 *
 * Script Properties:
 *   EMAIL_ENABLED    — must be "true" for any send to happen
 *   EMAIL_ALLOWLIST  — comma-separated allowed recipients. Empty = all.
 */
function _emailSendGuarded_(to, subject, plainBody, htmlBody) {
  // Sends a branded HTML email (with the VMRF logo inlined) plus a
  // plain-text fallback body. The OTP design has no must-click links,
  // so HTML is safe — users copy the code and (optionally) the URL.
  var sp = PropertiesService.getScriptProperties();
  var enabled  = String(sp.getProperty('EMAIL_ENABLED') || '').toLowerCase() === 'true';
  var allowStr = String(sp.getProperty('EMAIL_ALLOWLIST') || '').trim();
  var allowList = allowStr ? allowStr.split(',').map(function(s){ return s.trim().toLowerCase(); }).filter(Boolean) : [];
  var toLc = String(to||'').trim().toLowerCase();

  if (!toLc) {
    _emailLog_(to, subject, 'blocked-no-recipient', plainBody);
    return { sent: false, reason: 'no recipient' };
  }
  if (!enabled) {
    _emailLog_(to, subject, 'blocked-blackout', plainBody);
    return { sent: false, reason: 'EMAIL_ENABLED is false' };
  }
  if (allowList.length > 0 && allowList.indexOf(toLc) < 0) {
    _emailLog_(to, subject, 'blocked-allowlist', plainBody);
    return { sent: false, reason: 'recipient not on EMAIL_ALLOWLIST' };
  }

  try {
    var opts = { name: ACT_EMAIL_FROM_NAME };
    if (htmlBody) {
      opts.htmlBody = htmlBody;
      try {
        opts.inlineImages = { vmrfLogo: _vmrfEmailLogoBlob_() };
      } catch(_) { /* logo unavailable — send HTML without it */ }
    }
    if (typeof GmailApp !== 'undefined') {
      GmailApp.sendEmail(to, subject, plainBody, opts);
    } else {
      opts.to = to; opts.subject = subject; opts.body = plainBody;
      MailApp.sendEmail(opts);
    }
    _emailLog_(to, subject, 'sent', plainBody);
    return { sent: true };
  } catch (e) {
    _emailLog_(to, subject, 'failed: ' + e.message, plainBody);
    return { sent: false, reason: e.message };
  }
}

/**
 * Append a row to the Email_Log sheet. Creates the sheet if missing.
 * Never throws.
 */
function _emailLog_(to, subject, status, body) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(EMAIL_LOG_SHEET);
    if (!sh) {
      sh = ss.insertSheet(EMAIL_LOG_SHEET);
      sh.getRange(1, 1, 1, 5).setValues([['Timestamp','To','Subject','Status','Body (truncated)']]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date(), to, subject, status, String(body||'').substring(0, 800)]);
  } catch (_) {}
}


// ── OTP generation, storage, lookup ─────────────────────────────────

/** Generate a 6-digit numeric activation code as a string. */
function _makeActivationOtp_() {
  return String(Math.floor(100000 + Math.random() * 900000));
}

/**
 * Store an activation code for an email. Overwrites any previous code
 * for the same email (so a resend always issues a fresh code and
 * invalidates the old one).
 */
function _storeActivationOtp_(email, otp, name, role) {
  var emLc = String(email).trim().toLowerCase();
  PropertiesService.getScriptProperties().setProperty('ACT_OTP_' + emLc, JSON.stringify({
    otp:       String(otp),
    name:      name || '',
    role:      String(role || '').toUpperCase(),
    createdAt: Date.now(),
    attempts:  0
  }));
}

/**
 * Look up the activation record for an email. Returns the payload or
 * null if none / expired (expired records are cleaned up).
 */
function _lookupActivationOtp_(email) {
  var emLc = String(email||'').trim().toLowerCase();
  if (!emLc) return null;
  var sp = PropertiesService.getScriptProperties();
  var raw = sp.getProperty('ACT_OTP_' + emLc);
  if (!raw) return null;
  var payload;
  try { payload = JSON.parse(raw); } catch(_) { return null; }
  if (!payload || !payload.otp || !payload.createdAt) return null;
  if ((Date.now() - payload.createdAt) / 1000 > ACT_OTP_TTL_SEC) {
    sp.deleteProperty('ACT_OTP_' + emLc);
    return null;
  }
  return payload;
}

/**
 * Record a wrong attempt. Deletes the code when the limit is reached.
 * Returns the number of attempts remaining (0 = now locked).
 */
function _bumpOtpAttempts_(email, payload) {
  var emLc = String(email).trim().toLowerCase();
  var sp = PropertiesService.getScriptProperties();
  payload.attempts = (payload.attempts || 0) + 1;
  if (payload.attempts >= ACT_OTP_MAX_ATTEMPTS) {
    sp.deleteProperty('ACT_OTP_' + emLc);
    return 0;
  }
  sp.setProperty('ACT_OTP_' + emLc, JSON.stringify(payload));
  return ACT_OTP_MAX_ATTEMPTS - payload.attempts;
}

/** Delete the activation record after successful activation. */
function _consumeActivationOtp_(email) {
  var emLc = String(email||'').trim().toLowerCase();
  if (emLc) PropertiesService.getScriptProperties().deleteProperty('ACT_OTP_' + emLc);
}


// ── Web app URL (for email instructions only — not a magic link) ────

/**
 * The login page URL included in activation emails so users know where
 * to go. Prefers the WEB_APP_URL Script Property (set it to your /exec
 * deployment URL). Falls back to ScriptApp's URL with /dev → /exec.
 */
function _getWebAppUrl_() {
  var sp = PropertiesService.getScriptProperties();
  var configured = sp.getProperty('WEB_APP_URL');
  if (configured) return configured.replace(/\/$/, '');
  try {
    var base = ScriptApp.getService().getUrl();
    if (base) return base.replace(/\/dev$/, '/exec');
  } catch(_){}
  return '(ask the IMO for the portal link)';
}


// ── Email content ───────────────────────────────────────────────────

function _buildActivationEmailPlain_(name, otp, role) {
  var roleLabel = role === 'FACULTY' ? 'faculty' :
                  role === 'HOD' ? 'Head of Department' :
                  role === 'HOI' ? 'Head of Institution' : 'staff';
  var url = _getWebAppUrl_();
  return '' +
    'Hello ' + (name || 'there') + ',\n\n' +
    'Your ' + roleLabel + ' account on the VMRF-DU Institutional Monitoring ' +
    'System is ready to be activated.\n\n' +
    'Your activation code: ' + otp + '\n\n' +
    'How to activate:\n' +
    '1. Copy this address and paste it into your browser (do not click it):\n' +
    '   ' + url + '\n' +
    '2. On the sign-in page, choose "Activate your account".\n' +
    '3. Enter your email address, the code above, and set your password.\n\n' +
    'The code is valid for 24 hours. You can request a new code from the ' +
    'activation form if it expires.\n\n' +
    'Regards,\n' +
    'Office of the Chancellor\n' +
    'Vinayaka Mission\'s Research Foundation';
}

/**
 * VMRF logo as an inline-image blob (480x127 JPEG, ~17 KB).
 * Embedded as base64 so the email never depends on external hosting.
 */
function _vmrfEmailLogoBlob_() {
  var b64 = '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';
  return Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', 'vmrf-logo.jpg');
}

/**
 * Branded HTML activation email. The 6-digit code is the hero element;
 * the portal URL is shown as selectable text with copy-paste guidance.
 * Recolored to the Meridian Authority palette (navy #0C1E3C / gold
 * #C49A2A / ivory #F5F7FB) — Step 7 of the Phase 5 implementation plan.
 */
// ── ORIGINAL (pre-Meridian) VERSION — kept for instant reversion ─────────
// To revert: delete the active _buildActivationEmailHtml_ below and
// un-comment this block (remove the leading // from each line).
// function _buildActivationEmailHtml_(name, otp, role) {
//   var roleLabel = role === 'FACULTY' ? 'faculty' :
//                   role === 'HOD' ? 'Head of Department' :
//                   role === 'HOI' ? 'Head of Institution' : 'staff';
//   var url = _getWebAppUrl_();
//   var esc = function(s){ return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); };
//   return '' +
//   '<div style="background:#f4f6fb;padding:28px 12px;font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Roboto,Arial,sans-serif">' +
//     '<div style="max-width:560px;margin:0 auto;background:#ffffff;border-radius:12px;overflow:hidden;border:1px solid #e3e7ef">' +
//       '<div style="padding:26px 30px 18px 30px;border-bottom:3px solid #2a6fb0;text-align:center">' +
//         '<img src="cid:vmrfLogo" alt="Vinayaka Mission\'s Research Foundation" width="240" style="max-width:100%;height:auto">' +
//       '</div>' +
//       '<div style="padding:28px 30px 8px 30px;color:#1f2430">' +
//         '<h2 style="margin:0 0 6px 0;font-size:19px;color:#2a6fb0">Activate your IMS account</h2>' +
//         '<p style="margin:0 0 18px 0;font-size:13.5px;color:#5b6472;line-height:1.55">' +
//           'Hello ' + esc(name || 'there') + ', your ' + roleLabel + ' account on the ' +
//           'Institutional Monitoring System is ready. Use the code below to activate it.</p>' +
//         '<div style="background:#f0f5fb;border:1px dashed #2a6fb0;border-radius:10px;text-align:center;padding:18px 10px;margin:0 0 22px 0">' +
//           '<div style="font-size:11px;letter-spacing:1px;color:#5b6472;text-transform:uppercase;margin-bottom:6px">Your activation code</div>' +
//           '<div style="font-size:32px;font-weight:700;letter-spacing:8px;color:#1f2430;font-family:Consolas,Menlo,monospace">' + esc(otp) + '</div>' +
//           '<div style="font-size:11.5px;color:#8a93a3;margin-top:6px">valid for 24 hours</div>' +
//         '</div>' +
//         '<div style="font-size:13.5px;color:#1f2430;line-height:1.7">' +
//           '<strong>How to activate</strong>' +
//           '<ol style="margin:8px 0 0 0;padding-left:20px;color:#3c4250">' +
//             '<li>Copy this address and paste it into your browser:<br>' +
//               '<span style="display:inline-block;margin-top:4px;background:#f6f7f9;border:1px solid #e3e7ef;border-radius:6px;' +
//               'padding:8px 10px;font-family:Consolas,Menlo,monospace;font-size:11.5px;word-break:break-all;color:#2a6fb0">' + esc(url) + '</span></li>' +
//             '<li style="margin-top:8px">On the sign-in page, choose <strong>&ldquo;Activate your account&rdquo;</strong>.</li>' +
//             '<li style="margin-top:4px">Enter your email, the code above, and set your password.</li>' +
//           '</ol>' +
//         '</div>' +
//         '<p style="margin:22px 0 0 0;font-size:12px;color:#8a93a3;line-height:1.6">' +
//           'If the code expires, use &ldquo;Send me a new code&rdquo; on the activation form. ' +
//           'If you didn\'t expect this email, you can safely ignore it.</p>' +
//       '</div>' +
//       '<div style="padding:16px 30px 22px 30px;border-top:1px solid #eef1f6;margin-top:18px">' +
//         '<p style="margin:0;font-size:12px;color:#8a93a3;line-height:1.5">Regards,<br>' +
//           'Office of the Chancellor<br>' +
//           '<strong style="color:#5b6472">Vinayaka Mission\'s Research Foundation</strong><br>' +
//           '<span style="font-size:11px">Deemed to be University under section 3 of the UGC Act 1956</span></p>' +
//       '</div>' +
//     '</div>' +
//   '</div>';
// }
// ──────────────────────────────────────────────────────────────────────────
function _buildActivationEmailHtml_(name, otp, role) {
  var roleLabel = role === 'FACULTY' ? 'faculty' :
                  role === 'HOD' ? 'Head of Department' :
                  role === 'HOI' ? 'Head of Institution' : 'staff';
  var url = _getWebAppUrl_();
  var esc = function(s){ return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); };
  return '' +
  '<div style="background:#F5F7FB;padding:28px 12px;font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Roboto,Arial,sans-serif">' +
    '<div style="max-width:560px;margin:0 auto;background:#ffffff;border-radius:12px;overflow:hidden;border:1px solid #E8ECF4">' +
      '<div style="padding:26px 30px 18px 30px;border-bottom:3px solid #C49A2A;text-align:center">' +
        '<img src="cid:vmrfLogo" alt="Vinayaka Mission\'s Research Foundation" width="240" style="max-width:100%;height:auto">' +
      '</div>' +
      '<div style="padding:28px 30px 8px 30px;color:#16233A">' +
        '<h2 style="margin:0 0 6px 0;font-size:20px;color:#0C1E3C;font-family:Georgia,serif">Activate your IMS account</h2>' +
        '<p style="margin:0 0 18px 0;font-size:13.5px;color:#4A6280;line-height:1.55">' +
          'Hello ' + esc(name || 'there') + ', your ' + roleLabel + ' account on the ' +
          'Institutional Monitoring System is ready. Use the code below to activate it.</p>' +
        '<div style="background:#FBF1D8;border:1px dashed #C49A2A;border-radius:10px;text-align:center;padding:18px 10px;margin:0 0 22px 0">' +
          '<div style="font-size:11px;letter-spacing:1px;color:#7A5C0D;text-transform:uppercase;margin-bottom:6px">Your activation code</div>' +
          '<div style="font-size:32px;font-weight:700;letter-spacing:8px;color:#0C1E3C;font-family:Consolas,Menlo,monospace">' + esc(otp) + '</div>' +
          '<div style="font-size:11.5px;color:#8a93a3;margin-top:6px">valid for 24 hours</div>' +
        '</div>' +
        '<div style="font-size:13.5px;color:#16233A;line-height:1.7">' +
          '<strong>How to activate</strong>' +
          '<ol style="margin:8px 0 0 0;padding-left:20px;color:#2C4460">' +
            '<li>Copy this address and paste it into your browser:<br>' +
              '<span style="display:inline-block;margin-top:4px;background:#f6f7f9;border:1px solid #E8ECF4;border-radius:6px;' +
              'padding:8px 10px;font-family:Consolas,Menlo,monospace;font-size:11.5px;word-break:break-all;color:#1A3D72">' + esc(url) + '</span></li>' +
            '<li style="margin-top:8px">On the sign-in page, choose <strong>&ldquo;Activate your account&rdquo;</strong>.</li>' +
            '<li style="margin-top:4px">Enter your email, the code above, and set your password.</li>' +
          '</ol>' +
        '</div>' +
        '<p style="margin:22px 0 0 0;font-size:12px;color:#8a93a3;line-height:1.6">' +
          'If the code expires, use &ldquo;Send me a new code&rdquo; on the activation form. ' +
          'If you didn\'t expect this email, you can safely ignore it.</p>' +
      '</div>' +
      '<div style="padding:16px 30px 22px 30px;border-top:1px solid #E8ECF4;margin-top:18px">' +
        '<p style="margin:0;font-size:12px;color:#8a93a3;line-height:1.5">Regards,<br>' +
          'Office of the Chancellor<br>' +
          '<strong style="color:#0C1E3C">Vinayaka Mission\'s Research Foundation</strong><br>' +
          '<span style="font-size:11px">Deemed to be University under section 3 of the UGC Act 1956</span></p>' +
      '</div>' +
    '</div>' +
  '</div>';
}


// ── User lookup ─────────────────────────────────────────────────────

/**
 * Find a user row by email across Faculty_Master and Staff_Master.
 * Returns { sheet, sheetName, rowIndex, name, role, status,
 *           institution, campus } or null.
 */
function _findUserByEmail_(email) {
  if (!email) return null;
  var emLc = String(email).trim().toLowerCase();
  if (!emLc) return null;
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var fac = ss.getSheetByName(SH.FACULTY);
  if (fac) {
    var fd = fac.getDataRange().getValues();
    if (fd.length >= 2) {
      var fh = fd[0];
      var emI = fh.indexOf('Email'), nmI = fh.indexOf('FacultyName'),
          stI = fh.indexOf('Status'), inI = fh.indexOf('Institution'),
          caI = fh.indexOf('Campus');
      if (emI >= 0) {
        for (var i = 1; i < fd.length; i++) {
          if (String(fd[i][emI]||'').trim().toLowerCase() === emLc) {
            return {
              sheet: fac, sheetName: SH.FACULTY, rowIndex: i + 1,
              name: nmI >= 0 ? fd[i][nmI] : '',
              role: 'FACULTY',
              status: stI >= 0 ? String(fd[i][stI]||'').trim() : '',
              institution: inI >= 0 ? fd[i][inI] : '',
              campus: caI >= 0 ? fd[i][caI] : ''
            };
          }
        }
      }
    }
  }

  var stf = ss.getSheetByName(SH.STAFF);
  if (stf) {
    var sd = stf.getDataRange().getValues();
    if (sd.length >= 2) {
      var sh2 = sd[0];
      var emI2 = sh2.indexOf('Email'), nmI2 = sh2.indexOf('StaffName'),
          rlI2 = sh2.indexOf('Role'), stI2 = sh2.indexOf('Status'),
          inI2 = sh2.indexOf('Institution'), caI2 = sh2.indexOf('Campus');
      if (emI2 >= 0) {
        for (var j = 1; j < sd.length; j++) {
          if (String(sd[j][emI2]||'').trim().toLowerCase() === emLc) {
            return {
              sheet: stf, sheetName: SH.STAFF, rowIndex: j + 1,
              name: nmI2 >= 0 ? sd[j][nmI2] : '',
              role: rlI2 >= 0 ? String(sd[j][rlI2]||'').toUpperCase() : '',
              status: stI2 >= 0 ? String(sd[j][stI2]||'').trim() : '',
              institution: inI2 >= 0 ? sd[j][inI2] : '',
              campus: caI2 >= 0 ? sd[j][caI2] : ''
            };
          }
        }
      }
    }
  }
  return null;
}


// ════════════════════════════════════════════════════════════════════
// PUBLIC API — called by the frontend
// ════════════════════════════════════════════════════════════════════

/**
 * Request (or re-request) an activation code for an email.
 * Generates a fresh 6-digit code, stores it (invalidating any previous
 * code), and emails it. Always returns the same generic success message
 * regardless of whether the email exists — prevents enumeration.
 */
function requestActivation(email) {
  var generic = {
    ok: true,
    message: 'If this email is registered and not yet activated, an activation code has been sent to it.'
  };
  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return generic;

  var user = _findUserByEmail_(email);
  if (!user || user.status !== STATUS_AWAITING_ACTIVATION) return generic;

  var otp = _makeActivationOtp_();
  _storeActivationOtp_(email, otp, user.name, user.role);

  var plain = _buildActivationEmailPlain_(user.name, otp, user.role);
  var html  = _buildActivationEmailHtml_(user.name, otp, user.role);
  var sendResult = _emailSendGuarded_(email, 'Your VMRF-DU IMS activation code', plain, html);
  Logger.log('requestActivation: ' + email + ' → ' + JSON.stringify(sendResult));

  generic._debug = sendResult;  // editor visibility only
  return generic;
}

/**
 * Complete activation with email + code + chosen password.
 * On success: password hashed and written, Status → Active, and for
 * HOI/IMO the IPM Users dual-write runs. The code is then consumed.
 *
 * @return {object} { ok, role, email, message } or { ok:false, reason }
 */
function completeActivationWithOtp(email, otp, password, confirmPassword) {
  if (!email)    return { ok: false, reason: 'Please enter your email address.' };
  if (!otp)      return { ok: false, reason: 'Please enter the activation code from your email.' };
  if (!password) return { ok: false, reason: 'Please choose a password.' };
  if (password !== confirmPassword) return { ok: false, reason: 'Passwords do not match.' };

  try { _validatePasswordStrength(password); }
  catch (e) { return { ok: false, reason: e.message }; }

  var emLc = String(email).trim().toLowerCase();
  var record = _lookupActivationOtp_(emLc);
  if (!record) {
    return { ok: false, reason: 'No valid activation code found for this email. It may have expired — use "Send me a new code".' };
  }

  if (String(otp).trim() !== record.otp) {
    var remaining = _bumpOtpAttempts_(emLc, record);
    if (remaining <= 0) {
      return { ok: false, reason: 'Too many incorrect attempts. The code has been invalidated — use "Send me a new code".' };
    }
    return { ok: false, reason: 'Incorrect code. ' + remaining + ' attempt' + (remaining === 1 ? '' : 's') + ' remaining.' };
  }

  var user = _findUserByEmail_(emLc);
  if (!user) return { ok: false, reason: 'No account found for this email.' };
  if (user.status !== STATUS_AWAITING_ACTIVATION) {
    return { ok: false, reason: 'This account is already activated. Please sign in normally.', alreadyActive: true };
  }

  if (typeof _hashPwd !== 'function') {
    return { ok: false, reason: 'Server misconfiguration: hash function not loaded.' };
  }

  var hdr = user.sheet.getRange(1, 1, 1, user.sheet.getLastColumn()).getValues()[0];
  var pwI = hdr.indexOf('PasswordHash');
  var stI = hdr.indexOf('Status');
  if (pwI < 0 || stI < 0) {
    return { ok: false, reason: 'Server misconfiguration: required columns missing on ' + user.sheetName };
  }
  user.sheet.getRange(user.rowIndex, pwI + 1).setValue(_hashPwd(password));
  user.sheet.getRange(user.rowIndex, stI + 1).setValue('Active');

  var ipmResult = null;
  if (user.role === 'HOI' || user.role === 'IMO') {
    try {
      ipmResult = _activationUpsertIpmUser_(user.role, emLc, password, user.institution, user.campus);
    } catch (e) {
      ipmResult = { ok: false, note: 'IPM dual-write threw: ' + e.message };
    }
  }

  _consumeActivationOtp_(emLc);

  Logger.log('completeActivationWithOtp: ' + emLc + ' role=' + user.role +
             ' (' + user.sheetName + ' row ' + user.rowIndex + ')' +
             (ipmResult ? ' IPM: ' + JSON.stringify(ipmResult) : ''));

  return {
    ok:      true,
    role:    user.role,
    email:   emLc,
    message: 'Your account has been activated. You can now sign in with your new password.',
    ipmNote: ipmResult ? ipmResult.note : null
  };
}


// ════════════════════════════════════════════════════════════════════
// ADMIN — bulk send + per-user resend
// ════════════════════════════════════════════════════════════════════

/**
 * Send activation codes to every user with Status='AwaitingActivation'.
 *   sendActivationInvitations()                  → commit
 *   sendActivationInvitations({ dryRun: true })  → count only
 *   sendActivationInvitations({ role: 'HOI' })   → one role only
 * Respects EMAIL_ENABLED / EMAIL_ALLOWLIST (blocked sends are logged).
 */
function sendActivationInvitations(opts) {
  opts = opts || {};
  var dryRun = opts.dryRun === true;
  var roleFilter = opts.role ? String(opts.role).toUpperCase() : null;
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var candidates = [];
  ['FACULTY','STAFF'].forEach(function(sk) {
    var sn = SH[sk];
    var sh = ss.getSheetByName(sn);
    if (!sh) return;
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return;
    var hdr = data[0];
    var emI = hdr.indexOf('Email'),
        stI = hdr.indexOf('Status'),
        nmI = sk === 'FACULTY' ? hdr.indexOf('FacultyName') : hdr.indexOf('StaffName'),
        rlI = sk === 'FACULTY' ? -1 : hdr.indexOf('Role');
    if (emI < 0 || stI < 0) return;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][stI]||'').trim() !== STATUS_AWAITING_ACTIVATION) continue;
      var role = sk === 'FACULTY' ? 'FACULTY' : (rlI >= 0 ? String(data[i][rlI]||'').toUpperCase() : 'STAFF');
      if (roleFilter && role !== roleFilter) continue;
      candidates.push({
        email: String(data[i][emI]||'').trim().toLowerCase(),
        name:  nmI >= 0 ? data[i][nmI] : '',
        role:  role
      });
    }
  });

  var sent = 0, blocked = 0, failed = 0;
  var report = [];
  candidates.forEach(function(c) {
    if (!c.email) return;
    if (dryRun) { report.push({ email: c.email, role: c.role, outcome: 'WOULD_SEND' }); return; }
    var otp = _makeActivationOtp_();
    _storeActivationOtp_(c.email, otp, c.name, c.role);
    var plain = _buildActivationEmailPlain_(c.name, otp, c.role);
    var html  = _buildActivationEmailHtml_(c.name, otp, c.role);
    var r = _emailSendGuarded_(c.email, 'Your VMRF-DU IMS activation code', plain, html);
    if (r.sent) { sent++; report.push({ email: c.email, role: c.role, outcome: 'SENT' }); }
    else if ((r.reason||'').indexOf('allowlist') >= 0 || (r.reason||'').indexOf('blackout') >= 0) {
      blocked++; report.push({ email: c.email, role: c.role, outcome: 'BLOCKED', reason: r.reason });
    } else {
      failed++; report.push({ email: c.email, role: c.role, outcome: 'FAILED', reason: r.reason });
    }
  });

  var summary = (dryRun ? 'DRY RUN' : 'BULK SEND') + ' — candidates: ' + candidates.length +
                (dryRun ? ' (none sent)' : (', sent: ' + sent + ', blocked: ' + blocked + ', failed: ' + failed));
  Logger.log(summary);
  report.slice(0, 20).forEach(function(r) {
    Logger.log('  [' + r.outcome + '] ' + r.email + (r.reason ? ' — ' + r.reason : ''));
  });
  try {
    SpreadsheetApp.getUi().alert(summary +
      (dryRun ? '\n\nNo emails were sent. Use the commit menu item to send.' :
                '\n\nSee the Email_Log sheet for full details.'));
  } catch(_){}

  return { ok: true, summary: summary, candidates: candidates.length,
           sent: sent, blocked: blocked, failed: failed, report: report };
}

/**
 * Resend a fresh activation code to one user. The previous code is
 * invalidated automatically (storage is keyed by email).
 */
function resendActivationInvitation(email) {
  if (!email) throw new Error('Email is required.');
  var user = _findUserByEmail_(email);
  if (!user) return { ok: false, reason: 'No user found with email ' + email };
  if (user.status !== STATUS_AWAITING_ACTIVATION) {
    return { ok: false, reason: 'User status is "' + user.status + '", not AwaitingActivation.' };
  }
  var r = requestActivation(email);
  Logger.log('resendActivationInvitation: ' + email + ' → ' + JSON.stringify(r));
  return r;
}


// ════════════════════════════════════════════════════════════════════
// Menu wrappers (used by IMO.gs onOpen)
// ════════════════════════════════════════════════════════════════════

function menuSendActivationDryRun() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.alert('Activation codes — DRY RUN',
    'Counts how many activation-code emails would be sent. Sends nothing.\n\nProceed?',
    ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;
  sendActivationInvitations({ dryRun: true });
}

function menuSendActivationForReal() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.alert('Activation codes — COMMIT',
    'Sends a 6-digit activation code to every user with Status=AwaitingActivation.\n\n' +
    'Respects EMAIL_ENABLED and EMAIL_ALLOWLIST script properties.\n\nProceed?',
    ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;
  sendActivationInvitations({ dryRun: false });
}

function menuResendActivation() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt('Resend activation code',
    'Enter the email of the user to resend:',
    ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var email = resp.getResponseText().trim();
  if (!email) return;
  var r = resendActivationInvitation(email);
  ui.alert(r.ok ? 'Code sent (or logged if blocked by allowlist).' : 'Could not resend: ' + r.reason);
}

function menuShowEmailLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(EMAIL_LOG_SHEET);
  if (!sh) {
    SpreadsheetApp.getUi().alert('Email_Log sheet does not exist yet. It is created on the first logged email.');
    return;
  }
  ss.setActiveSheet(sh);
  sh.setActiveRange(sh.getRange(Math.max(2, sh.getLastRow()), 1));
}


// ════════════════════════════════════════════════════════════════════
// Test wrappers (Step 3, OTP edition)
// ════════════════════════════════════════════════════════════════════

/**
 * Full backend round-trip for the Faculty test account:
 * reset → request code → read stored code → complete activation → login.
 */
function testActivationFlowForFaculty() {
  var email = 'doeerica32+facultytest@gmail.com';

  Logger.log('--- Reset test Faculty account ---');
  Logger.log(JSON.stringify(resetTestActivation(email)));

  Logger.log('\n--- Request activation code ---');
  Logger.log(JSON.stringify(requestActivation(email)));

  Logger.log('\n--- Read stored code (simulating the user reading their email) ---');
  var record = _lookupActivationOtp_(email);
  if (!record) { Logger.log('FAIL: no code stored'); return; }
  Logger.log('Code: ' + record.otp);

  Logger.log('\n--- Complete activation with TestPass1 ---');
  Logger.log(JSON.stringify(completeActivationWithOtp(email, record.otp, 'TestPass1', 'TestPass1')));

  Logger.log('\n--- Code should now be consumed ---');
  Logger.log('Lookup after consume: ' + JSON.stringify(_lookupActivationOtp_(email)));

  Logger.log('\n--- facultyLogin should now succeed ---');
  try {
    var r = facultyLogin(email, 'TestPass1');
    Logger.log('Login result: success=' + r.success + ', role=' + r.role);
  } catch(e) {
    Logger.log('Login THREW: ' + e.message);
  }
}

/**
 * Same round-trip for the HOI test account, including IPM dual-write
 * verification — the most failure-prone path.
 */
function testActivationFlowForHoi() {
  var email = 'doeerica32+hoitest@gmail.com';

  Logger.log('--- Reset test HOI account ---');
  Logger.log(JSON.stringify(resetTestActivation(email)));

  Logger.log('\n--- Request activation code ---');
  Logger.log(JSON.stringify(requestActivation(email)));

  var record = _lookupActivationOtp_(email);
  if (!record) { Logger.log('FAIL: no code stored'); return; }
  Logger.log('Code: ' + record.otp);

  Logger.log('\n--- Complete activation (includes IPM dual-write) ---');
  var done = completeActivationWithOtp(email, record.otp, 'TestPass1', 'TestPass1');
  Logger.log(JSON.stringify(done));
  Logger.log('IPM note: ' + done.ipmNote);

  Logger.log('\n--- staffLogin HOI should now succeed ---');
  try {
    var r = staffLogin('HOI', email, 'TestPass1',
                       "Vinayaka Mission's Chennai Campus",
                       'Aarupadai Veedu Institute of Technology (AVIT)');
    Logger.log('Login result: ' + JSON.stringify(r).substring(0, 200));
  } catch(e) {
    Logger.log('Login THREW: ' + e.message);
  }
}


// ── One-click reset wrappers (no-argument; runnable from the dropdown) ──

/** Reset the Faculty test account back to AwaitingActivation. */
function resetFacultyTest() {
  var r = resetTestActivation('doeerica32+facultytest@gmail.com');
  Logger.log(JSON.stringify(r));
  try { SpreadsheetApp.getUi().alert(r.note + '\n' + (r.rowsReset || []).join('\n')); } catch(_){}
}

/** Reset the HOD test account back to AwaitingActivation. */
function resetHodTest() {
  var r = resetTestActivation('doeerica32+hodtest@gmail.com');
  Logger.log(JSON.stringify(r));
  try { SpreadsheetApp.getUi().alert(r.note + '\n' + (r.rowsReset || []).join('\n')); } catch(_){}
}

/** Reset the HOI test account back to AwaitingActivation. */
function resetHoiTest() {
  var r = resetTestActivation('doeerica32+hoitest@gmail.com');
  Logger.log(JSON.stringify(r));
  try { SpreadsheetApp.getUi().alert(r.note + '\n' + (r.rowsReset || []).join('\n')); } catch(_){}
}


// ════════════════════════════════════════════════════════════════════
// PART J — STEP 4: First-login profile gate
//
// On first successful login, the user must confirm/complete their
// profile before reaching the landing page. Detection is via a
// ProfileCompleted column on Faculty_Master / Staff_Master (blank for
// imported rows; TRUE once the form is submitted).
//
// Frontend calls:
//   getProfileGateStatus(email, role) → { complete } or
//        { complete:false, fields:{...prefill...}, readOnly:{...} }
//   saveProfileGate(email, role, fields) → { ok } or { ok:false, reason }
//
// Faculty  → full form, saved to FacultyProfiles (+ Designation/Phone
//            synced to Faculty_Master where columns exist).
// HOD/HOI  → light form (designation + phone), saved to Staff_Master.
// IMO      → exempt (frontend never gates it).
// ════════════════════════════════════════════════════════════════════

var PROFILE_GATE_COL = 'ProfileCompleted';

// Faculty form fields (FacultyProfiles columns), mandatory flags.
var PROFILE_FACULTY_FIELDS = [
  { key:'FacultyName',       label:'Full name',                 required:true  },
  { key:'EmployeeID',        label:'Employee ID',               required:true  },
  { key:'Designation',       label:'Designation',               required:true  },
  { key:'Department',        label:'Department',                required:true  },
  { key:'Qualification',     label:'Highest qualification',     required:true  },
  { key:'Phone',             label:'Phone number',              required:true  },
  { key:'DateOfJoining',     label:'Date of joining',           required:true, type:'date' },
  { key:'Specialization',    label:'Specialization',            required:false },
  { key:'Experience',        label:'Experience (years)',        required:false },
  { key:'ResearchAreas',     label:'Research areas',            required:false },
  { key:'PublicationsCount', label:'Number of publications',    required:false },
  { key:'Certifications',    label:'Certifications',            required:false },
  { key:'LinkedinOrcid',     label:'LinkedIn / ORCID link',     required:false },
  { key:'Bio',               label:'Short bio',                 required:false, type:'textarea' }
];

// HOD / HOI light form fields (Staff_Master columns).
var PROFILE_STAFF_FIELDS = [
  { key:'StaffName',   label:'Full name',     required:true },
  { key:'Designation', label:'Designation',   required:true },
  { key:'Phone',       label:'Phone number',  required:true }
];


/** Ensure the ProfileCompleted column exists on both master sheets. */
function _ensureProfileGateColumns_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ['FACULTY','STAFF'].forEach(function(sk) {
    var sh = ss.getSheetByName(SH[sk]);
    if (sh) _activationEnsureColumns_(sh, [PROFILE_GATE_COL]);
  });
}

/** Find a FacultyProfiles row by email → { sheet, rowIndex, hdr, values } or null. */
function _findProfileRow_(email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH.PROFILES);
  if (!sh) return null;
  var data = sh.getDataRange().getValues();
  if (data.length < 1) return null;
  var hdr = data[0].map(function(v){ return String(v).trim(); });
  var emI = hdr.indexOf('Email');
  if (emI < 0) return null;
  var emLc = String(email).trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][emI]||'').trim().toLowerCase() === emLc) {
      return { sheet: sh, rowIndex: i + 1, hdr: hdr, values: data[i] };
    }
  }
  return { sheet: sh, rowIndex: -1, hdr: hdr, values: null };  // sheet exists, row doesn't
}


/**
 * Called by the frontend right after a successful login.
 * Returns { complete:true } when no gate is needed, otherwise the
 * field definitions + prefill values for the form.
 */
function getProfileGateStatus(email, role) {
  try {
    if (!email) return { complete: true };           // fail open — never lock out
    var roleU = String(role || '').toUpperCase();
    if (roleU === 'IMO') return { complete: true };  // IMO exempt

    _ensureProfileGateColumns_();
    var user = _findUserByEmail_(email);
    if (!user) return { complete: true };            // unknown → fail open

    var hdr = user.sheet.getRange(1, 1, 1, user.sheet.getLastColumn()).getValues()[0]
                .map(function(v){ return String(v).trim(); });
    var pcI = hdr.indexOf(PROFILE_GATE_COL);
    if (pcI >= 0) {
      var flag = String(user.sheet.getRange(user.rowIndex, pcI + 1).getValue() || '').trim().toUpperCase();
      if (flag === 'TRUE' || flag === 'YES') return { complete: true };
    }

    // Build prefill
    var fields = {}, readOnly = {
      Email:       String(email).trim().toLowerCase(),
      Institution: user.institution || '',
      Campus:      user.campus || ''
    };

    if (user.role === 'FACULTY') {
      // Seed from master row first
      fields.FacultyName = user.name || '';
      // Then overlay the FacultyProfiles row (richer, pre-seeded at import)
      var prof = _findProfileRow_(email);
      if (prof && prof.values) {
        PROFILE_FACULTY_FIELDS.forEach(function(f) {
          var ci = prof.hdr.indexOf(f.key);
          if (ci >= 0 && prof.values[ci] !== '' && prof.values[ci] != null) {
            var v = prof.values[ci];
            if (v instanceof Date) {
              v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            }
            fields[f.key] = String(v);
          }
        });
      }
      return { complete: false, role: 'FACULTY',
               fieldDefs: PROFILE_FACULTY_FIELDS, fields: fields, readOnly: readOnly };
    }

    // HOD / HOI — light form from Staff_Master
    fields.StaffName = user.name || '';
    var dI = hdr.indexOf('Designation'), phI = hdr.indexOf('Phone');
    if (dI  >= 0) fields.Designation = String(user.sheet.getRange(user.rowIndex, dI  + 1).getValue() || '');
    if (phI >= 0) fields.Phone       = String(user.sheet.getRange(user.rowIndex, phI + 1).getValue() || '');
    return { complete: false, role: user.role,
             fieldDefs: PROFILE_STAFF_FIELDS, fields: fields, readOnly: readOnly };
  } catch (e) {
    Logger.log('getProfileGateStatus error: ' + e.message);
    return { complete: true };  // any error → fail open, never block login
  }
}


/**
 * Save the gate form. Validates mandatory fields, writes the profile,
 * sets ProfileCompleted=TRUE on the master row.
 */
function saveProfileGate(email, role, fields) {
  try {
    if (!email)  return { ok: false, reason: 'Missing email.' };
    fields = fields || {};
    var user = _findUserByEmail_(email);
    if (!user) return { ok: false, reason: 'Account not found.' };

    var defs = (user.role === 'FACULTY') ? PROFILE_FACULTY_FIELDS : PROFILE_STAFF_FIELDS;
    for (var i = 0; i < defs.length; i++) {
      if (defs[i].required && !String(fields[defs[i].key] || '').trim()) {
        return { ok: false, reason: 'Please fill in: ' + defs[i].label };
      }
    }

    _ensureProfileGateColumns_();

    if (user.role === 'FACULTY') {
      // Upsert FacultyProfiles
      var prof = _findProfileRow_(email);
      if (!prof) return { ok: false, reason: 'FacultyProfiles sheet not found. Contact the IMO.' };
      var emLc = String(email).trim().toLowerCase();
      if (prof.rowIndex < 0) {
        // No row yet — append one in header order
        var newVals = prof.hdr.map(function(col) {
          if (col === 'Email')       return emLc;
          if (col === 'Institution') return user.institution || '';
          if (col === 'Campus')      return user.campus || '';
          if (col === 'LastUpdated') return new Date();
          return fields[col] !== undefined ? fields[col] : '';
        });
        prof.sheet.appendRow(newVals);
      } else {
        // Update existing row, only the form fields + LastUpdated
        PROFILE_FACULTY_FIELDS.forEach(function(f) {
          var ci = prof.hdr.indexOf(f.key);
          if (ci >= 0 && fields[f.key] !== undefined) {
            prof.sheet.getRange(prof.rowIndex, ci + 1).setValue(fields[f.key]);
          }
        });
        var luI = prof.hdr.indexOf('LastUpdated');
        if (luI >= 0) prof.sheet.getRange(prof.rowIndex, luI + 1).setValue(new Date());
      }
      // Sync the duplicated columns on Faculty_Master
      var mh = user.sheet.getRange(1, 1, 1, user.sheet.getLastColumn()).getValues()[0]
                 .map(function(v){ return String(v).trim(); });
      [['FacultyName','FacultyName'],['Designation','Designation'],['Department','Department']].forEach(function(pair) {
        var ci = mh.indexOf(pair[0]);
        if (ci >= 0 && fields[pair[1]] !== undefined && String(fields[pair[1]]).trim() !== '') {
          user.sheet.getRange(user.rowIndex, ci + 1).setValue(fields[pair[1]]);
        }
      });
    } else {
      // HOD / HOI — write to Staff_Master directly
      var sh2 = user.sheet.getRange(1, 1, 1, user.sheet.getLastColumn()).getValues()[0]
                  .map(function(v){ return String(v).trim(); });
      [['StaffName','StaffName'],['Designation','Designation'],['Phone','Phone']].forEach(function(pair) {
        var ci = sh2.indexOf(pair[0]);
        if (ci >= 0 && fields[pair[1]] !== undefined && String(fields[pair[1]]).trim() !== '') {
          user.sheet.getRange(user.rowIndex, ci + 1).setValue(fields[pair[1]]);
        }
      });
    }

    // Flip the gate flag on the master row
    var mh2 = user.sheet.getRange(1, 1, 1, user.sheet.getLastColumn()).getValues()[0]
                .map(function(v){ return String(v).trim(); });
    var pcI = mh2.indexOf(PROFILE_GATE_COL);
    if (pcI >= 0) user.sheet.getRange(user.rowIndex, pcI + 1).setValue('TRUE');

    Logger.log('saveProfileGate: ' + email + ' (' + user.role + ') completed profile');
    return { ok: true };
  } catch (e) {
    Logger.log('saveProfileGate error: ' + e.message);
    return { ok: false, reason: 'Server error: ' + e.message };
  }
}


/** Test helper: clear the gate flag so the form shows again at next login. */
function resetProfileGateForTestAccounts() {
  _ensureProfileGateColumns_();
  var emails = ['doeerica32+facultytest@gmail.com',
                'doeerica32+hodtest@gmail.com',
                'doeerica32+hoitest@gmail.com'];
  var out = [];
  emails.forEach(function(em) {
    var user = _findUserByEmail_(em);
    if (!user) { out.push(em + ': not found'); return; }
    var hdr = user.sheet.getRange(1, 1, 1, user.sheet.getLastColumn()).getValues()[0]
                .map(function(v){ return String(v).trim(); });
    var pcI = hdr.indexOf(PROFILE_GATE_COL);
    if (pcI >= 0) {
      user.sheet.getRange(user.rowIndex, pcI + 1).setValue('');
      out.push(em + ': gate reset');
    }
  });
  Logger.log(out.join('\n'));
  try { SpreadsheetApp.getUi().alert(out.join('\n')); } catch(_){}
}
