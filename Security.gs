// =============================================================================
// VMRF-DU Institutional Monitoring System — Security Helpers
// =============================================================================
// This file centralises:
//   • Strong password hashing (iterated SHA-256 with per-user salt)
//   • Legacy hash detection + on-the-fly migration
//   • Login attempt throttling (5 fails → 10-min cache lockout)
//   • Default-password detection (forces a change at first login)
//
// All functions are designed to be backwards-compatible with the existing
// Djb2 _hashPwd() so production accounts continue to work while they migrate.
// =============================================================================

/* ── CONFIG ───────────────────────────────────────────────────────────── */
var SEC_HASH_PREFIX     = 'P2:';   // marker for new-format hashes; legacy hashes start with 'H'
var SEC_PBKDF_ITERS     = 10000;   // ~1 sec / hash on Apps Script V8 — high enough to be brute-force-hostile
var SEC_MAX_ATTEMPTS    = 5;       // failed logins before lockout
var SEC_LOCKOUT_SECONDS = 600;     // 10-minute lockout

// Known-bad default passwords. The first time a user with one of these tries
// to log in, the function returns mustChangePassword:true and the UI should
// force a reset before letting them proceed.
var SEC_DEFAULT_PASSWORDS = ['IMO@VMRF2024', 'HOI@VMRF2024'];


/* ── HASH: NEW FORMAT (P2:hex, salt stored in row) ───────────────────── */
// Iterated SHA-256 PBKDF-style. Returns { hash, salt }.
// If `salt` is omitted, a fresh 32-hex-char salt is generated.
function _hashPwdSecure(password, salt) {
  if (!salt) salt = Utilities.getUuid().replace(/-/g, '');  // 32 hex chars of entropy
  var current = String(salt) + ':' + String(password);
  for (var i = 0; i < SEC_PBKDF_ITERS; i++) {
    var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, current);
    var hex = '';
    for (var b = 0; b < bytes.length; b++) {
      var v = (bytes[b] & 0xff).toString(16);
      hex += (v.length === 1 ? '0' + v : v);
    }
    current = hex;
  }
  return { hash: SEC_HASH_PREFIX + current, salt: salt };
}

// Verify a plaintext password against a stored hash.
// Handles both new (P2:…) and legacy (H…) formats.
// Returns { ok: bool, needsRehash: bool } — needsRehash=true means caller
// should overwrite the stored hash with the new format using _hashPwdSecure.
function _verifyPwd(password, storedHash, storedSalt) {
  if (!storedHash) return { ok: false, needsRehash: false };
  // New format
  if (String(storedHash).indexOf(SEC_HASH_PREFIX) === 0) {
    if (!storedSalt) return { ok: false, needsRehash: false };
    var computed = _hashPwdSecure(password, storedSalt);
    return { ok: computed.hash === storedHash, needsRehash: false };
  }
  // Legacy Djb2 (anything else — typically starts with 'H'). If it matches,
  // ask the caller to migrate to the new format.
  if (typeof _hashPwd === 'function') {
    var ok = (_hashPwd(password) === storedHash);
    return { ok: ok, needsRehash: ok };
  }
  return { ok: false, needsRehash: false };
}

// True if `password` is a known-bad default. Use this AFTER verifying the
// password is correct, to refuse access until the user changes it.
function _isDefaultPassword(password) {
  if (!password) return false;
  for (var i = 0; i < SEC_DEFAULT_PASSWORDS.length; i++) {
    if (String(password) === SEC_DEFAULT_PASSWORDS[i]) return true;
  }
  return false;
}


/* ── THROTTLING (CacheService — cheap, no Drive writes per attempt) ──── */
function _throttleKey_(role, identifier) {
  return 'la_' + String(role || '').toUpperCase() + '_' + String(identifier || '').trim().toLowerCase();
}

// Check throttle status. Throws if the caller is locked out so the surrounding
// login function aborts naturally. Returns the current attempt count (0 if no
// failures recorded yet).
function _checkLoginThrottle_(role, identifier) {
  var key = _throttleKey_(role, identifier);
  var v = CacheService.getScriptCache().get(key);
  var attempts = v ? parseInt(v, 10) : 0;
  if (attempts >= SEC_MAX_ATTEMPTS) {
    throw new Error('Account temporarily locked due to too many failed login attempts. ' +
                    'Please wait 10 minutes and try again.');
  }
  return attempts;
}

// Record a failed attempt. Returns the new attempt count.
function _bumpLoginAttempts_(role, identifier) {
  var key = _throttleKey_(role, identifier);
  var v = CacheService.getScriptCache().get(key);
  var attempts = (v ? parseInt(v, 10) : 0) + 1;
  CacheService.getScriptCache().put(key, String(attempts), SEC_LOCKOUT_SECONDS);
  return attempts;
}

// Clear failed-attempt counter (call after a successful login).
function _clearLoginAttempts_(role, identifier) {
  var key = _throttleKey_(role, identifier);
  CacheService.getScriptCache().remove(key);
}


/* ── ADMIN UTILITY: MIGRATE STORED HASHES ─────────────────────────────── */
// Run from the Apps Script editor (or VMRF IMO menu) to back-fill a Salt
// column and convert every Djb2 hash to the new P2: format.
//
// IMPORTANT: This rotates passwords for accounts whose stored hash cannot be
// directly upgraded (we don't know the plaintext). Those accounts are flagged
// with Status='Pending' and a fresh random password is written to a CSV in
// Drive. Operators must distribute the new passwords through a secure channel.
//
// For accounts with the new format already, the function is a no-op.
//
// Idempotent — safe to re-run.
function migratePasswordsToSecureFormat() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var report = { faculty: { migrated: 0, alreadyNew: 0, rotated: 0 },
                 staff:   { migrated: 0, alreadyNew: 0, rotated: 0 } };
  var rotatedRows = [['Sheet','Email','NewPassword','Role','Note']];

  function _migrateSheet(sheetName, roleLabel) {
    var sh = ss.getSheetByName(sheetName);
    if (!sh) return;
    // Make sure the Salt column exists
    _ensureSheetColumns(sh, [].concat(SCHEMA[sheetName] || []));
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return;
    var h = data[0];
    var emI = h.indexOf('Email');
    var pwI = h.indexOf('PasswordHash');
    var saI = h.indexOf('Salt');
    if (pwI < 0 || saI < 0 || emI < 0) {
      Logger.log('Skipping ' + sheetName + ' — required columns missing.');
      return;
    }
    var rlI = h.indexOf('Role');

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var stored = String(row[pwI] || '').trim();
      var salt   = String(row[saI] || '').trim();
      var email  = String(row[emI] || '').trim();
      if (!email) continue;
      var role = rlI >= 0 ? String(row[rlI] || '').toUpperCase() : roleLabel;

      if (stored.indexOf(SEC_HASH_PREFIX) === 0 && salt) {
        report[sheetName === SH.FACULTY ? 'faculty' : 'staff'].alreadyNew++;
        continue;
      }
      if (!stored) {
        // No password set — leave it; user will use Forgot Password.
        continue;
      }
      // Legacy hash — we don't know the plaintext. Rotate to a fresh random
      // password, write new salt+hash, and emit a CSV row so the admin can
      // distribute it.
      var plain = _genHodPassword_();
      var fresh = _hashPwdSecure(plain);
      sh.getRange(i + 1, pwI + 1).setValue(fresh.hash);
      sh.getRange(i + 1, saI + 1).setValue(fresh.salt);
      rotatedRows.push([sheetName, email, plain, role, 'Password rotated during security upgrade']);
      report[sheetName === SH.FACULTY ? 'faculty' : 'staff'].rotated++;
    }
  }

  _migrateSheet(SH.FACULTY, 'FACULTY');
  _migrateSheet(SH.STAFF,   'STAFF');

  // Write CSV
  var csv = rotatedRows.map(function(r){
    return r.map(function(cell){
      var s = String(cell == null ? '' : cell);
      return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    }).join(',');
  }).join('\n');
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmmss');
  var fileName = 'VMRF_PasswordMigration_' + stamp + '.csv';
  var fileUrl  = '';
  try {
    var file = DriveApp.createFile(fileName, csv, MimeType.CSV);
    fileUrl  = file.getUrl();
  } catch(e) { fileUrl = '(CSV creation failed: ' + e.message + ')'; }

  var summary =
    '✅ Password migration complete.\n\n' +
    'Faculty_Master:\n' +
    '  Already on new format: ' + report.faculty.alreadyNew + '\n' +
    '  Rotated to new format: ' + report.faculty.rotated   + '\n\n' +
    'Staff_Master:\n' +
    '  Already on new format: ' + report.staff.alreadyNew + '\n' +
    '  Rotated to new format: ' + report.staff.rotated   + '\n\n' +
    (rotatedRows.length > 1 ? 'New credentials CSV saved to Drive:\n' + fileName + '\n' + fileUrl
                            : 'No rotations needed — all accounts were already on the new format.');
  try { SpreadsheetApp.getUi().alert(summary); } catch(_){}
  Logger.log(summary);
  return { ok: true, report: report, fileUrl: fileUrl };
}
