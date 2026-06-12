// =============================================================================
// VMRF-DU Institutional Monitoring System — HOI Meetings & KPI Services
// =============================================================================

// ─── HOI QUEUE ────────────────────────────────────────────────────────────────
function getHOIQueue(hoiInstitution, hoiInstCode) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var subD=ss.getSheetByName(SH.SUBMISSION).getDataRange().getValues();
  var facD=ss.getSheetByName(SH.FACULTY).getDataRange().getValues();
  var hodD=ss.getSheetByName(SH.HOD).getDataRange().getValues();
  var hoiD=ss.getSheetByName(SH.HOI).getDataRange().getValues();
  var saD =ss.getSheetByName(SH.SELF_ASSESS).getDataRange().getValues();
  var tsD =ss.getSheetByName(SH.TIMESHEET).getDataRange().getValues();
  var subMap=_bm(subD,subD[0]);
  var sidEmailMap=_buildSidEmailMap(subD);
  var facMap=_buildFacMap(facD, ss.getSheetByName(SH.STAFF).getDataRange().getValues());
  var hodMap=_bm(hodD,hodD[0]);
  var hoiMap=_bm(hoiD,hoiD[0]);
  var saMap=_bm(saD,saD[0]);
  var tsMap=_bmMulti(tsD,tsD[0]);
  var noHodSids = _buildNoHodSidSet_(sidEmailMap, facMap);
  // Normalize the HOI's institution for scoping Pass 2
  var hoiInstNorm = String(hoiInstitution||'').trim().toLowerCase();
  var hoiCodeNorm = String(hoiInstCode||'').trim().toLowerCase();
  var out=[];
  var seen={};

  // Pass 1: HOD-approved submissions (normal flow)
  var hodH=hodD[0];
  var hodStI=hodH.indexOf('HOD_Status'), hodSbI=hodH.indexOf('SubmissionID');
  for(var i=1;i<hodD.length;i++){
    var sid=String(hodD[i][hodSbI]||'').trim();
    var hst=String(hodD[i][hodStI]||'').trim();
    if(!sid||hst!=='Approved') continue;
    var hoiR=hoiMap[sid]||{}, hoiSt=String(hoiR['HOI_Status']||'').trim();
    if(hoiSt!=='') continue;
    var sub=subMap[sid]||{};
    var fid=sidEmailMap[sid]||_getFidFromSub(sub);
    var fac=fid?(facMap[fid]||{}):{};
    // Scope to HOI's institution
    if (hoiInstNorm || hoiCodeNorm) {
      var p1FacInstN = String(fac['Institution']||'').trim().toLowerCase();
      var p1FacCodeN = _resolveInstCode('', String(fac['Institution']||'').trim()).toLowerCase();
      var p1HoiBare  = hoiInstNorm.replace(/\s*\([^)]*\)\s*$/, '').trim();
      var p1FacBare  = p1FacInstN.replace(/\s*\([^)]*\)\s*$/, '').trim();
      var p1MatchCode = hoiCodeNorm && (p1FacCodeN === hoiCodeNorm || p1FacInstN === hoiCodeNorm);
      var p1MatchName = hoiInstNorm && (p1FacInstN === hoiInstNorm || p1FacBare === p1HoiBare || p1FacBare === hoiInstNorm || p1FacInstN === p1HoiBare);
      if (!p1MatchCode && !p1MatchName) continue;
    }
    var sa=saMap[sid]||{}, hodR=hodMap[sid]||{};
    out.push(_buildItem(sid,sub,fac,sa,tsMap[sid]||[],
      {hodStatus:hst,hodRemark:String(hodR['HOD_Remark']||'')},
      {hoiStatus:hoiSt},null));
    seen[sid]=true;
  }

  // Pass 2: No-HOD institution submissions — check HOI sheet directly
  var hoiH=hoiD[0];
  var hoiStI2=hoiH.indexOf('HOI_Status'), hoiSbI2=hoiH.indexOf('SubmissionID');
  for(var j=1;j<hoiD.length;j++){
    var sid2=String(hoiD[j][hoiSbI2]||'').trim();
    if(!sid2||seen[sid2]) continue;
    if(!noHodSids[sid2]) continue; // only no-HOD sids
    var hoiSt2=String(hoiD[j][hoiStI2]||'').trim();
    if(hoiSt2!=='') continue; // already reviewed
    var sub2=subMap[sid2]||{};
    var fid2=sidEmailMap[sid2]||_getFidFromSub(sub2);
    var fac2=fid2?(facMap[fid2]||{}):{};
    // Scope to this HOI's institution if provided
    if (hoiInstNorm || hoiCodeNorm) {
      var facInstN = String(fac2['Institution']||'').trim().toLowerCase();
      var facCodeN = _resolveInstCode('', String(fac2['Institution']||'').trim()).toLowerCase();
      // Strip (CODE) suffix from both sides so "Vinayaka Mission's Law School (VMLS)"
      // matches "Vinayaka Mission's Law School" stored in Faculty_Master
      var hoiBare  = hoiInstNorm.replace(/\s*\([^)]*\)\s*$/, '').trim();
      var facBare  = facInstN.replace(/\s*\([^)]*\)\s*$/, '').trim();
      var matchByCode = hoiCodeNorm && (facCodeN === hoiCodeNorm || facInstN === hoiCodeNorm);
      var matchByName = hoiInstNorm && (facInstN === hoiInstNorm || facBare === hoiBare || facBare === hoiInstNorm || facInstN === hoiBare);
      if (!matchByCode && !matchByName) continue;
    }
    var sa2=saMap[sid2]||{};
    out.push(_buildItem(sid2,sub2,fac2,sa2,tsMap[sid2]||[],
      {hodStatus:'N/A',hodRemark:''},
      {hoiStatus:''},null));
  }
  return out;
}

function submitHOIReview(sid, remark, status) {
  if(!sid)    throw new Error('Submission ID missing.');
  if(!status) throw new Error('Please select a status.');
  if(status!=='Approved'&&!remark) throw new Error('Please enter your remarks.');
  var ss3=SpreadsheetApp.getActiveSpreadsheet();

  // ─── Determine no-HOD status with TWO redundant strategies ─────────────────
  // The previous implementation relied solely on looking up the faculty's
  // Institution string via _rowByKey + Faculty_Master, which was fragile to
  // mixed-case emails and to institution-name format drift. We now combine:
  //
  //   1) HOD sheet row presence — submitFacultyReport intentionally skips
  //      appending an HOD_Remarks row for no-HOD institutions, so absence of
  //      a row for this SID is conclusive evidence that this is a no-HOD case.
  //
  //   2) Faculty institution string — kept as a belt-and-suspenders check using
  //      _buildFacMap (case-insensitive) so old-data submissions made before
  //      the no-HOD flag was added still resolve correctly.
  //
  // If EITHER strategy says no-HOD, we accept it. Only when both checks fail
  // and HOD_Status is not 'Approved' do we throw.

  var hodRow = _rowByKey(SH.HOD, sid);
  var noHodByRow = !hodRow;

  var noHodByInst = false;
  try {
    var subD   = ss3.getSheetByName(SH.SUBMISSION).getDataRange().getValues();
    var facD   = ss3.getSheetByName(SH.FACULTY).getDataRange().getValues();
    var staffD = ss3.getSheetByName(SH.STAFF).getDataRange().getValues();
    var sidEmailMap4 = _buildSidEmailMap(subD);
    var facMap4      = _buildFacMap(facD, staffD);
    var email4       = sidEmailMap4[String(sid).trim()] || '';
    var fac4         = (email4 && facMap4[email4]) || {};
    var instName4    = String(fac4['Institution']||'').trim();
    if (instName4) {
      noHodByInst = _isNoHodInstitution_(instName4);
      if (!noHodByInst) {
        try { noHodByInst = _isNoHodInstitution_(_resolveInstCode('', instName4)); } catch(e2) {}
      }
    }
  } catch(e) { Logger.log('submitHOIReview noHodByInst check failed: '+e.message); }

  var _isNH4 = noHodByRow || noHodByInst;
  var hodSt  = hodRow ? String(hodRow['HOD_Status']||'').trim() : '';

  if(!_isNH4 && hodSt!=='Approved')
    throw new Error('HOD must approve this submission before HOI can review it.');

  _writeReview(SH.HOI,sid,remark||'',status);
  var sub3=_rowByKey(SH.SUBMISSION,sid)||{};
  var _fid3=_getFidFromSub(sub3);
  // Prefer case-insensitive facMap lookup; fall back to _rowByKey
  var fac3 = {};
  if (_fid3) {
    try {
      var _facMap3 = _buildFacMap(
        ss3.getSheetByName(SH.FACULTY).getDataRange().getValues(),
        ss3.getSheetByName(SH.STAFF).getDataRange().getValues()
      );
      fac3 = _facMap3[String(_fid3).toLowerCase()] || {};
    } catch(e) { fac3 = {}; }
    if (!fac3 || !Object.keys(fac3).length) {
      fac3 = _rowByKey(SH.FACULTY, _fid3, 'Email') || {};
    }
  }
  var fn3=String(fac3['FacultyName']||'Faculty');
  var fid3=_fid3;
  if(status==='Approved'){
    // HOI approval is final — auto-write Finalised to IMO monitoring sheet
    _writeReview(SH.IMO, sid, '', 'Finalised');
    _pushNotif('FACULTY:'+fid3,'status_update','✅ Submission Finalised',
      'Your submission has been approved by HOI and is now finalised.',sid,fn3);
    try{_notifyIMO(ss3,sid);}catch(e){Logger.log(e.message);}
  } else {
    _pushNotif('FACULTY:'+fid3,'rejected','❌ Submission Rejected by HOI',
      'Your submission was rejected by HOI'+(remark?' with the remark: '+remark:'')+'. Please resubmit after making the necessary changes.',sid,fn3);
    try{_notifyRevision(ss3,sid,'Head of Institution',remark);}catch(e){Logger.log(e.message);}
  }
  return { ok:true };
}

function getHOIStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Faculty Submission Sheets ──
  var subD = ss.getSheetByName(SH.SUBMISSION).getDataRange().getValues(), subH = subD[0];
  var facD = ss.getSheetByName(SH.FACULTY).getDataRange().getValues(), facH = facD[0];
  var hodD = ss.getSheetByName(SH.HOD).getDataRange().getValues(), hodH = hodD[0];
  var hoiD = ss.getSheetByName(SH.HOI).getDataRange().getValues(), hoiH = hoiD[0];

  // ── HOD Submission Sheets ──
  var hodSubSh = ss.getSheetByName(SH.HOD_SUB);
  var hodSubD  = hodSubSh ? hodSubSh.getDataRange().getValues() : [['SubmissionID']];
  var hodSubH  = hodSubD[0];
  var hodRevSh = ss.getSheetByName(SH.HOD_REVIEW);
  var hodRevD  = hodRevSh ? hodRevSh.getDataRange().getValues() : [['SubmissionID']];
  var hodRevH  = hodRevD[0];

  // ── Build lookup maps ──
  var facMap   = _buildFacMap(facD, ss.getSheetByName(SH.STAFF).getDataRange().getValues());
  var _fecI    = _facEmailCol(subH); if(_fecI<0)_fecI=1;
  var _sidIdx  = subH.indexOf('SubmissionID'); if(_sidIdx<0)_sidIdx=0;
  var _depIdx  = facH.indexOf('Department');

  // Map submissionID → { fid, dept }
  var subInfoMap = {};
  for(var s=1;s<subD.length;s++){
    var sid_s = String(subD[s][_sidIdx]||'').trim();
    if(!sid_s) continue;
    var fid_s = String(subD[s][_fecI]||'').trim().toLowerCase();
    var fac_s = facMap[fid_s]||{};
    subInfoMap[sid_s] = {
      fid:  fid_s,
      dept: String(fac_s['Department']||'').trim(),
      name: String(fac_s['FacultyName']||fid_s)
    };
  }

  // ── Faculty submissions by dept ──
  // dept → { pending, awaitingHOI, approved, revision, rejected, total }
  var deptStats = {};
  function getDeptStat(dept) {
    if(!deptStats[dept]) deptStats[dept] = {
      dept:dept, facPending:0, facAwaitingHOI:0, facApproved:0,
      facRevision:0, facRejected:0, facTotal:0
    };
    return deptStats[dept];
  }

  // HOD remarks map
  var hodMap = _bm(hodD, hodH);
  // HOI remarks map
  var hoiMap = _bm(hoiD, hoiH);

  var _hodSidI = hodH.indexOf('SubmissionID'); if(_hodSidI<0)_hodSidI=0;
  for(var i=1;i<hodD.length;i++){
    var sid = String(hodD[i][_hodSidI]||'').trim(); if(!sid) continue;
    var info = subInfoMap[sid]||{};
    var dept = info.dept||'Unassigned';
    var ds   = getDeptStat(dept);
    ds.facTotal++;
    var hs = String(hodD[i][hodH.indexOf('HOD_Status')]||'').trim();
    var hoiRow = hoiMap[sid]||{};
    var is = String(hoiRow['HOI_Status']||'').trim();
    if(hs===''){ ds.facPending++; }
    else if(hs==='Approved'){
      if(is===''||is==='Pending')  ds.facAwaitingHOI++;
      else if(is==='Approved')     ds.facApproved++;
      else if(is==='Rejected')     ds.facRejected++;
      else                         ds.facAwaitingHOI++;
    } else if(hs==='Rejected'||hs==='Needs Revision'){ ds.facRejected++; }
  }

  // ── HOD submissions (weekly reports from HODs) ──
  var hodRevMap = {};
  var _hrSidI = hodRevH.indexOf('SubmissionID'); if(_hrSidI<0)_hrSidI=0;
  for(var r=1;r<hodRevD.length;r++){
    var k=String(hodRevD[r][_hrSidI]||'').trim();
    if(k){ var o={}; hodRevH.forEach(function(c,j){if(c)o[c]=hodRevD[r][j];}); hodRevMap[k]=o; }
  }

  var hodSubStats = { pending:0, approved:0, revision:0, rejected:0, total:0 };
  var _hsSidI = hodSubH.indexOf('SubmissionID'); if(_hsSidI<0)_hsSidI=0;
  for(var h=1;h<hodSubD.length;h++){
    var hsid=String(hodSubD[h][_hsSidI]||'').trim(); if(!hsid) continue;
    hodSubStats.total++;
    var rev = hodRevMap[hsid]||{};
    var rst = String(rev['HOI_Status']||'').trim();
    if(rst==='') hodSubStats.pending++;
    else if(rst==='Approved')                          hodSubStats.approved++;
    else if(rst==='Rejected'||rst==='Needs Revision')  hodSubStats.rejected++;
    else                                               hodSubStats.pending++;
  }

  // ── Overall faculty totals ──
  var overallFac = { pending:0, awaitingHOI:0, approved:0, revision:0, rejected:0, total:0 };
  Object.keys(deptStats).forEach(function(d){
    var ds=deptStats[d];
    overallFac.pending     += ds.facPending;
    overallFac.awaitingHOI += ds.facAwaitingHOI;
    overallFac.approved    += ds.facApproved;
    overallFac.revision    += ds.facRevision;
    overallFac.rejected    += ds.facRejected;
    overallFac.total       += ds.facTotal;
  });

  // ── Total registered faculty ──
  var totalFaculty = Math.max(0, facD.length-1);

  return {
    totalFaculty:  totalFaculty,
    overallFac:    overallFac,
    hodSubStats:   hodSubStats,
    deptStats:     Object.keys(deptStats).sort().map(function(d){ return deptStats[d]; })
  };
}

function getHOIQueueHOD(hoiInstitution, hoiCampus) {
  // Normalise scope filters sent by the HOI client.
  // Both are optional (legacy callers pass nothing and see everything,
  // but every current HOI login sends its institution + campus).
  var fInst   = String(hoiInstitution || '').trim();
  var fCampus = String(hoiCampus      || '').trim();

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var subSh = ss.getSheetByName(SH.HOD_SUB); if (!subSh) return [];
  var subD  = subSh.getDataRange().getValues(), subH = subD[0];
  var revSh = ss.getSheetByName(SH.HOD_REVIEW);
  var revD  = revSh ? revSh.getDataRange().getValues() : [['SubmissionID']]; var revH = revD[0];
  var tsSh  = ss.getSheetByName(SH.HOD_TS);
  var tsD   = tsSh ? tsSh.getDataRange().getValues() : [['SubmissionID']];
  var saSh  = ss.getSheetByName(SH.HOD_SA);
  var saD   = saSh ? saSh.getDataRange().getValues() : [['SubmissionID']]; var saH = saD[0];

  // Look up HOD staff details from Staff_Master — now includes Campus & Institution
  // so we can filter to only the HOI's own institution.
  var staffMap = {};
  var staffSh = ss.getSheetByName(SH.STAFF);
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

  // Build review map
  var revMap = {};
  for (var r2 = 1; r2 < revD.length; r2++) {
    revMap[String(revD[r2][0]||'')] = { status: String(revD[r2][revH.indexOf('HOI_Status')]||''), remark: String(revD[r2][revH.indexOf('HOI_Remark')]||'') };
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

  var out = [];
  for (var i = 1; i < subD.length; i++) {
    var r   = subD[i];
    var sid = String(r[0]||'').trim();
    var rev = revMap[sid] || { status:'', remark:'' };
    if (rev.status === 'Approved') continue; // already reviewed
    var sa  = saMap2[sid] || {};
    var hodID = String(r[subH.indexOf('HOD_ID')]||'').trim().toLowerCase();
    var staff = staffMap[hodID] || {};

    // ── Scope guard: only show submissions whose HOD belongs to this HOI's
    //    institution (and campus if provided). Submissions from HODs in other
    //    institutions are never returned to this HOI.
    if (fInst && String(staff.institution||'').trim() !== fInst) continue;
    if (fCampus && String(staff.campus||'').trim()      !== fCampus) continue;

    out.push({
      sid:        sid,
      hodID:      hodID,
      name:       staff.name || 'HOD',
      department: staff.dept || '',
      semester:   String(r[subH.indexOf('AcademicYearSemester')]||''),
      from:       _fmt(r[subH.indexOf('ReportingFrom')]),
      to:         _fmt(r[subH.indexOf('ReportingTo')]),
      submitted:  _fmtDT(r[subH.indexOf('SubmittedDateTime')]),
      tasks:      sa.tasks || '',
      target:     sa.target || '',
      hoiStatus:  rev.status || 'Pending HOI',
      hoiRemark:  rev.remark || '',
      timesheet:  tsMap2[sid] || []
    });
  }
  return out.reverse();
}

function submitHOIHODReview(sid, remark, status, hoiInstitution) {
  if (!sid || !status) throw new Error('Submission ID and status are required.');
  // Institution guard — verify this submission belongs to the calling HOI's institution
  if (hoiInstitution) {
    var fInst = String(hoiInstitution).trim();
    var staffSh2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.STAFF);
    if (staffSh2) {
      var sd2 = staffSh2.getDataRange().getValues(), sh2 = sd2[0];
      var _eIdx = sh2.indexOf('Email'), _iIdx = sh2.indexOf('Institution');
      // Find the HOD_ID for this submission
      var hodSubSh2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH.HOD_SUB);
      if (hodSubSh2) {
        var hsd = hodSubSh2.getDataRange().getValues(), hsh = hsd[0];
        var _hIdx = hsh.indexOf('HOD_ID');
        var hodEmail = '';
        for (var hx = 1; hx < hsd.length; hx++) {
          if (String(hsd[hx][0]||'').trim() === String(sid).trim()) {
            hodEmail = String(hsd[hx][_hIdx >= 0 ? _hIdx : 1]||'').trim().toLowerCase();
            break;
          }
        }
        if (hodEmail && _iIdx >= 0) {
          var hodInst = '';
          for (var sx = 1; sx < sd2.length; sx++) {
            if (String(sd2[sx][_eIdx]||'').trim().toLowerCase() === hodEmail) {
              hodInst = String(sd2[sx][_iIdx]||'').trim();
              break;
            }
          }
          if (hodInst && hodInst !== fInst)
            throw new Error('Access denied: this submission belongs to a different institution.');
        }
      }
    }
  }
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var sh  = ss.getSheetByName(SH.HOD_REVIEW);
  if (!sh) throw new Error('HOD_Review sheet not found.');
  var data = sh.getDataRange().getValues(), h = data[0];
  var idI = h.indexOf('SubmissionID'), stI = h.indexOf('HOI_Status'), rmI = h.indexOf('HOI_Remark'), dtI = h.indexOf('HOI_DateTime');
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idI]) === String(sid)) {
      sh.getRange(i+1, stI+1).setValue(status);
      sh.getRange(i+1, rmI+1).setValue(remark || '');
      sh.getRange(i+1, dtI+1).setValue(now);
      if (status === 'Approved') {
        // HOI approval is final — auto-write Finalised to IMO monitoring sheet
        var imoSh = ss.getSheetByName(SH.HOD_IMO);
        if (!imoSh) {
          imoSh = ss.insertSheet(SH.HOD_IMO);
          var hdrs = SCHEMA[SH.HOD_IMO];
          if (hdrs) { imoSh.getRange(1,1,1,hdrs.length).setValues([hdrs]).setFontWeight('bold').setBackground('#1a3c5e').setFontColor('#ffffff'); imoSh.setFrozenRows(1); }
        }
        imoSh.appendRow([sid, '', 'Finalised', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')]);
      }
      // Resolve the HOD's department so the notification only reaches the
      // correct HOD, not every HOD in the system. Look up via HOD_Submission
      // → HOD_ID → Faculty_Master.Department.
      var hodDeptForNotif = '';
      try {
        var hodSubSh3 = ss.getSheetByName(SH.HOD_SUB);
        if (hodSubSh3) {
          var hsd3 = hodSubSh3.getDataRange().getValues(), hsh3 = hsd3[0];
          var _hI3 = hsh3.indexOf('HOD_ID');
          var hEmail3 = '';
          for (var hx3 = 1; hx3 < hsd3.length; hx3++) {
            if (String(hsd3[hx3][0]||'').trim() === String(sid).trim()) {
              hEmail3 = String(hsd3[hx3][_hI3 >= 0 ? _hI3 : 1]||'').trim().toLowerCase();
              break;
            }
          }
          if (hEmail3) {
            var hodFac3 = _rowByKey(SH.FACULTY, hEmail3, 'Email') || {};
            hodDeptForNotif = String(hodFac3['Department']||'').trim();
          }
        }
      } catch(e) { Logger.log('HOD dept lookup failed: '+e.message); }
      var hodKey3 = hodDeptForNotif ? 'HOD:'+hodDeptForNotif : 'HOD';
      _pushNotif(hodKey3, 'hod_review_done',
        status === 'Approved' ? '✅ HOI Approved Your Timesheet' : '⚠️ HOI Reviewed Your Timesheet',
        'Your weekly timesheet has been ' + status.toLowerCase() + ' by HOI.' + (remark ? ' Remark: ' + remark : ''),
        sid, 'HOI'
      );
      return { ok: true };
    }
  }
  throw new Error('Submission not found.');
}

// ─── Faculty list for HOD dropdowns (used in Pending & Priority Work form) ──
function getHRFacultyList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH.FACULTY); if (!sh) return [];
  var data = sh.getDataRange().getValues(), h = data[0].map(function(v){return String(v).trim();});
  var nmI = h.indexOf('FacultyName'), dsI = h.indexOf('Designation'), stI = h.indexOf('Status');
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var status = stI >= 0 ? String(data[i][stI]||'') : 'Active';
    if (status !== 'Active') continue;
    var name = nmI >= 0 ? String(data[i][nmI]||'').trim() : '';
    if (!name) continue;
    out.push({ name: name, designation: dsI >= 0 ? String(data[i][dsI]||'') : '' });
  }
  return out;
}

// ─── LIGHTWEIGHT INSTITUTION KPIs FOR MODAL (fast — reads minimal sheets) ────
function getInstKPIs(filter) {
  filter = filter || {};
  var ss  = SpreadsheetApp.getActiveSpreadsheet();

  // Read only 5 sheets (no STAFF, no TIMESHEET, no SELF_ASSESS, no zeroSubmit)
  var facD = ss.getSheetByName(SH.FACULTY).getDataRange().getValues(), facH = facD[0];
  var subD = ss.getSheetByName(SH.SUBMISSION).getDataRange().getValues(), subH = subD[0];
  var hodD = ss.getSheetByName(SH.HOD).getDataRange().getValues(), hodH = hodD[0];
  var hoiD = ss.getSheetByName(SH.HOI).getDataRange().getValues(), hoiH = hoiD[0];
  var imoD = ss.getSheetByName(SH.IMO).getDataRange().getValues(), imoH = imoD[0];

  // Build maps inline (no _buildFacMap which reads STAFF sheet)
  var facEm = facH.indexOf('Email'), facNm = facH.indexOf('FacultyName');
  var facDp = facH.indexOf('Department'), facCm = facH.indexOf('Campus');
  var facIn = facH.indexOf('Institution'), facSt = facH.indexOf('Status');

  var facMap = {};
  for (var fi = 1; fi < facD.length; fi++) {
    var em = String(facD[fi][facEm]||'').trim().toLowerCase();
    if (!em) continue;
    facMap[em] = {
      dept: String(facD[fi][facDp]||'').trim(),
      campus: String(facD[fi][facCm]||'').trim(),
      institution: String(facD[fi][facIn]||'').trim(),
      status: String(facD[fi][facSt]||'Active'),
      name: String(facD[fi][facNm]||'')
    };
  }

  // Build submission → email map
  var subH0 = subH.map(function(v){return String(v).trim();});
  var sidEmailI = subH0.indexOf('FacultyEmail'); if(sidEmailI<0) sidEmailI=1;
  var sidMap = {};
  for (var si = 1; si < subD.length; si++) {
    var sid = String(subD[si][0]||'').trim();
    if (sid) sidMap[sid] = String(subD[si][sidEmailI]||'').trim().toLowerCase();
  }

  // Build review status maps
  var hodStI = hodH.map(function(v){return String(v).trim();}).indexOf('HOD_Status');
  var hoiStI = hoiH.map(function(v){return String(v).trim();}).indexOf('HOI_Status');
  var imoStI = imoH.map(function(v){return String(v).trim();}).indexOf('IMO_Status');
  var hodMap = {}, hoiMap2 = {}, imoMap2 = {};
  for (var hi=1; hi<hodD.length; hi++) { var k=String(hodD[hi][0]||'').trim(); if(k) hodMap[k]=String(hodD[hi][hodStI>=0?hodStI:2]||'').trim(); }
  for (var oi=1; oi<hoiD.length; oi++) { var k=String(hoiD[oi][0]||'').trim(); if(k) hoiMap2[k]=String(hoiD[oi][hoiStI>=0?hoiStI:2]||'').trim(); }
  for (var mi=1; mi<imoD.length; mi++) { var k=String(imoD[mi][0]||'').trim(); if(k) imoMap2[k]=String(imoD[mi][imoStI>=0?imoStI:2]||'').trim(); }

  var fInst = (filter.institution||'').trim(), fCamp = (filter.campus||'').trim();
  var fDeptLow = (filter.department||'').trim().toLowerCase();
  var fDeptBase = fDeptLow.replace(/\s*\(pg\)\s*$/,'');

  function matchesFac(fac) {
    if (fInst && fac.institution !== fInst) return false;
    if (fCamp && fac.campus     !== fCamp) return false;
    if (fDeptLow) {
      var dl = fac.dept.toLowerCase(), db = dl.replace(/\s*\(pg\)\s*$/,'');
      if (dl !== fDeptLow && db !== fDeptBase) return false;
    }
    return true;
  }

  // Count active faculty (filtered)
  var totalFaculty = 0;
  for (var fi2=1; fi2<facD.length; fi2++) {
    var fac2 = facMap[String(facD[fi2][facEm]||'').trim().toLowerCase()];
    if (!fac2) continue;
    if (fac2.status && fac2.status !== 'Active') continue;
    if (matchesFac(fac2)) totalFaculty++;
  }

  // Count submissions (filtered)
  var totalSubs=0, finalised=0, pendingHOI=0, pendingHOD=0, rejected=0;
  var deptStats = {};
  for (var xi=1; xi<subD.length; xi++) {
    var xid = String(subD[xi][0]||'').trim(); if(!xid) continue;
    var xem = sidMap[xid] || '';
    var xfac = facMap[xem] || {};
    if (!matchesFac(xfac)) continue;
    var xdept = xfac.dept || 'Unassigned';
    var hs = hodMap[xid]||'', his = hoiMap2[xid]||'', ms = imoMap2[xid]||'';
    var st = (ms==='Finalised'||his==='Approved')?'Finalised':hs==='Approved'?'Pending HOI':(hs==='Rejected'||his==='Rejected')?'Rejected':'Pending HOD';
    totalSubs++;
    if(st==='Finalised') finalised++;
    else if(st==='Pending HOI') pendingHOI++;
    else if(st==='Rejected') rejected++;
    else pendingHOD++;
    if(!deptStats[xdept]) deptStats[xdept]={dept:xdept,facTotal:0,facApproved:0,facPending:0,facRejected:0};
    deptStats[xdept].facTotal++;
    if(st==='Finalised') deptStats[xdept].facApproved++;
    else if(st==='Rejected') deptStats[xdept].facRejected++;
    else deptStats[xdept].facPending++;
  }

  // HOD submissions (lightweight)
  var hodSubSh = ss.getSheetByName(SH.HOD_SUB);
  var hodSubD2 = hodSubSh ? hodSubSh.getDataRange().getValues() : [['SubmissionID']];
  var hodRevSh = ss.getSheetByName(SH.HOD_REVIEW);
  var hodRevD2 = hodRevSh ? hodRevSh.getDataRange().getValues() : [['SubmissionID','','HOI_Status']];
  var _hrStI2 = hodRevD2[0].indexOf ? hodRevD2[0].indexOf('HOI_Status') : 2;
  var hodRevMap3 = {};
  for (var hr=1; hr<hodRevD2.length; hr++) { hodRevMap3[String(hodRevD2[hr][0]||'')]=String(hodRevD2[hr][_hrStI2>=0?_hrStI2:2]||''); }
  var hodTotal=0, hodPending=0, hodApproved=0, hodRejected=0;
  for (var hs2=1; hs2<hodSubD2.length; hs2++) {
    var hsid=String(hodSubD2[hs2][0]||'').trim(); if(!hsid) continue;
    hodTotal++; var rst=hodRevMap3[hsid]||'';
    if(rst==='Approved') hodApproved++;
    else if(rst==='Rejected') hodRejected++;
    else hodPending++;
  }

  return {
    totalFaculty:   totalFaculty,
    totalSubmissions: totalSubs,
    finalised:      finalised,
    pendingHOI:     pendingHOI,
    pendingHOD:     pendingHOD,
    pendingIMO:     0,
    rejected:       rejected,
    hodSubStats:    { total:hodTotal, pending:hodPending, approved:hodApproved, rejected:hodRejected },
    deptStats:      Object.keys(deptStats).sort().map(function(d){return deptStats[d];})
  };
}

// ─── IMO DETAILED STATS WITH FILTERS ─────────────────────────────────────────
function getDetailedStats(filter) {
  filter = filter || {};
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var facD = ss.getSheetByName(SH.FACULTY).getDataRange().getValues(), facH = facD[0].map(function(v){return String(v).trim();});
  var subD = ss.getSheetByName(SH.SUBMISSION).getDataRange().getValues(), subH = subD[0].map(function(v){return String(v).trim();});
  var hodD = ss.getSheetByName(SH.HOD).getDataRange().getValues(), hodH = hodD[0].map(function(v){return String(v).trim();});
  var hoiD = ss.getSheetByName(SH.HOI).getDataRange().getValues(), hoiH = hoiD[0].map(function(v){return String(v).trim();});
  var imoD = ss.getSheetByName(SH.IMO).getDataRange().getValues(), imoH = imoD[0].map(function(v){return String(v).trim();});

  // ── Period filter — driven by filter.period ('week'/'month'/'year'/'all') ─
  var _pb2 = _periodBounds_(String(filter.period||'week'));
  var _pb2S = _pb2.start, _pb2E = _pb2.end;
  function _gdInPeriod(d){
    if(!_pb2S && !_pb2E) return true;
    return (d instanceof Date) && d>=_pb2S && d<=_pb2E;
  }
  var _gdDtI = subH.indexOf('SubmittedDateTime'); if(_gdDtI<0)_gdDtI=6;

  // Build faculty map keyed by email
  var facMap = _buildFacMap(facD, ss.getSheetByName(SH.STAFF).getDataRange().getValues());
  var sidEmailMap = _buildSidEmailMap(subD);

  // Build review maps keyed by submission ID
  var hodMap = {}, hoiMap = {}, imoMap = {};
  for (var h2=1; h2<hodD.length; h2++) { var s2=String(hodD[h2][0]||''); if(s2) hodMap[s2]=String(hodD[h2][hodH.indexOf('HOD_Status')]||''); }
  for (var h3=1; h3<hoiD.length; h3++) { var s3=String(hoiD[h3][0]||''); if(s3) hoiMap[s3]=String(hoiD[h3][hoiH.indexOf('HOI_Status')]||''); }
  for (var h4=1; h4<imoD.length; h4++) { var s4=String(imoD[h4][0]||''); if(s4) imoMap[s4]=String(imoD[h4][imoH.indexOf('IMO_Status')]||''); }

  // Collect unique filter values
  var allDepts = {}, allCampuses = {}, allInsts = {};
  for (var fi=1; fi<facD.length; fi++) {
    var fd = facD[fi];
    var dep = String(fd[facH.indexOf('Department')]||'').trim();
    var cam = String(fd[facH.indexOf('Campus')]||'').trim();
    var ins = String(fd[facH.indexOf('Institution')]||'').trim();
    if(dep) allDepts[dep]=1;
    if(cam) allCampuses[cam]=1;
    if(ins) allInsts[ins]=1;
  }

  // Process submissions — scoped to current week + any dimension filter
  var deptStats = {};
  var totalSubs=0, pendingHOD=0, pendingHOI=0, pendingIMO=0, finalised=0, escalated=0, rejected=0, revision=0;
  var facultySubCount = {};

  for (var i=1; i<subD.length; i++) {
    var sid  = String(subD[i][0]||'').trim();
    if (!sid) continue;
    // ── Period filter — skip if SubmittedDateTime is outside the selected period ─
    var _gdDt = subD[i][_gdDtI];
    if (!_gdInPeriod(_gdDt)) continue;
    var fid = sidEmailMap[sid] || '';
    var fac = fid ? (facMap[fid]||{}) : {};
    var dept  = String(fac['Department']||'').trim();
    var camp  = String(fac['Campus']||'').trim();
    var inst  = String(fac['Institution']||'').trim();

    // Apply filters (with PG suffix handling for department)
    var filterDeptLow = (filter.department||'').trim().toLowerCase();
    var filterDeptBase2 = filterDeptLow.replace(/\s*\(pg\)\s*$/,'');
    var deptLow = dept.toLowerCase();
    var deptBase = deptLow.replace(/\s*\(pg\)\s*$/,'');
    if (filterDeptLow && deptLow !== filterDeptLow && deptBase !== filterDeptBase2) continue;
    if (filter.campus     && filter.campus     !== camp) continue;
    if (filter.institution&& filter.institution!== inst) continue;

    var sid  = String(subD[i][0]||'').trim();
    var hs   = hodMap[sid]||'';
    var his  = hoiMap[sid]||'';
    var ms   = imoMap[sid]||'';

    // Overall status
    // HOI approval = Submitted to IMO (final state)
    var overallStatus = (ms==='Finalised'||his==='Approved')?'Finalised':ms==='Escalated'?'Escalated':hs==='Approved'?'Pending HOI':hs==='Rejected'||his==='Rejected'||hs==='Needs Revision'||his==='Needs Revision'?'Rejected':'Pending HOD';

    totalSubs++;
    if(overallStatus==='Finalised')        finalised++;
    else if(overallStatus==='Escalated')   escalated++;
    else if(overallStatus==='Rejected')    rejected++;
    else if(overallStatus==='Pending HOI') pendingHOI++;
    else pendingHOD++;

    // Department breakdown
    if (!dept) dept = 'Unassigned';
    if (!deptStats[dept]) deptStats[dept] = {total:0,finalised:0,escalated:0,revision:0,rejected:0,pending:0};
    deptStats[dept].total++;
    if(overallStatus==='Finalised')        deptStats[dept].finalised++;
    else if(overallStatus==='Escalated')   deptStats[dept].escalated++;
    else if(overallStatus==='Rejected')    deptStats[dept].rejected++;
    else if(overallStatus==='Pending HOI') deptStats[dept].pending++;
    else                                   deptStats[dept].pending++;

    // Faculty submission count
    var facName = String(fac['FacultyName']||fid);
    if (!facultySubCount[fid]) facultySubCount[fid] = {name:facName, dept:dept, count:0};
    facultySubCount[fid].count++;
  }

  // Faculty with zero submissions (filtered)
  var zeroSubmit = [];
  var _filterDL4 = (filter.department||'').trim().toLowerCase();
  var _filterDB4 = _filterDL4.replace(/\s*\(pg\)\s*$/,'');
  for (var fi2=1; fi2<facD.length; fi2++) {
    var fd2 = facD[fi2];
    var fid2 = String(fd2[facH.indexOf('Email')]||'').trim().toLowerCase();
    var dep2  = String(fd2[facH.indexOf('Department')]||'').trim();
    var dep2L = dep2.toLowerCase(), dep2B = dep2L.replace(/\s*\(pg\)\s*$/,'');
    var cam2  = String(fd2[facH.indexOf('Campus')]||'').trim();
    var ins2  = String(fd2[facH.indexOf('Institution')]||'').trim();
    var st2   = String(fd2[facH.indexOf('Status')]||'');
    if (st2 && st2 !== 'Active') continue;
    if (_filterDL4 && dep2L !== _filterDL4 && dep2B !== _filterDB4) continue;
    if (filter.campus     && filter.campus     !== cam2) continue;
    if (filter.institution&& filter.institution!== ins2) continue;
    if (!facultySubCount[fid2]) {
      zeroSubmit.push({name:String(fd2[facH.indexOf('FacultyName')]||fid2), dept:dep2, campus:cam2});
    }
  }

  // Total active faculty (filtered)
  var totalFac = 0;
  var _filterDL3 = (filter.department||'').trim().toLowerCase();
  var _filterDB3 = _filterDL3.replace(/\s*\(pg\)\s*$/,'');
  for (var fi3=1; fi3<facD.length; fi3++) {
    var st3 = String(facD[fi3][facH.indexOf('Status')]||'');
    if (st3 && st3 !== 'Active') continue;
    var fd3 = String(facD[fi3][facH.indexOf('Department')]||'').trim();
    var fd3L = fd3.toLowerCase(), fd3B = fd3L.replace(/\s*\(pg\)\s*$/,'');
    if (_filterDL3 && fd3L !== _filterDL3 && fd3B !== _filterDB3) continue;
    if (filter.campus     && filter.campus     !== String(facD[fi3][facH.indexOf('Campus')]||'').trim()) continue;
    if (filter.institution&& filter.institution!== String(facD[fi3][facH.indexOf('Institution')]||'').trim()) continue;
    totalFac++;
  }

  // ── HOD Submission Stats — current week only ─────────────────────────────
  var hodSubSh = ss.getSheetByName(SH.HOD_SUB);
  var hodSubD  = hodSubSh ? hodSubSh.getDataRange().getValues() : [['SubmissionID']];
  var hodRevSh = ss.getSheetByName(SH.HOD_REVIEW);
  var hodRevD  = hodRevSh ? hodRevSh.getDataRange().getValues() : [['SubmissionID']];
  var hodImoSh = ss.getSheetByName(SH.HOD_IMO);
  var hodImoD  = hodImoSh ? hodImoSh.getDataRange().getValues() : [['SubmissionID']];

  var hodSubTotal=0, hodPendHOI=0, hodApprHOI=0, hodPendIMO=0, hodFinalised=0, hodEscalated=0;
  var hodRevMap2={}, hodImoMap2={};
  var _hrStI = hodRevD[0].indexOf ? hodRevD[0].indexOf('HOI_Status') : -1;
  for (var hr=1; hr<hodRevD.length; hr++) { hodRevMap2[String(hodRevD[hr][0]||'')]=String(hodRevD[hr][_hrStI>=0?_hrStI:2]||''); }
  var _hiStI = hodImoD[0].indexOf ? hodImoD[0].indexOf('IMO_Status') : -1;
  for (var hi=1; hi<hodImoD.length; hi++) { hodImoMap2[String(hodImoD[hi][0]||'')]=String(hodImoD[hi][_hiStI>=0?_hiStI:2]||''); }
  // SubmittedDateTime is at index 6 in HOD_Submission schema
  var _hsStI = hodSubD[0].indexOf ? hodSubD[0].indexOf('SubmittedDateTime') : 6; if(_hsStI<0)_hsStI=6;
  for (var hs=1; hs<hodSubD.length; hs++) {
    var hsid=String(hodSubD[hs][0]||'').trim(); if(!hsid) continue;
    // Period filter — only count HOD submissions within the selected period
    if(!_gdInPeriod(hodSubD[hs][_hsStI])) continue;
    hodSubTotal++;
    var hoiSt2=hodRevMap2[hsid]||'';
    var imoSt2=hodImoMap2[hsid]||'';
    if(imoSt2==='Finalised') hodFinalised++;
    else if(imoSt2==='Escalated') hodEscalated++;
    else if(hoiSt2==='Approved') hodPendIMO++;
    else if(hoiSt2==='') hodPendHOI++;
    else hodApprHOI++;
  }

  return {
    totalFaculty:    totalFac,
    totalSubmissions:totalSubs,
    finalised:   finalised,
    escalated:   escalated,
    rejected:    rejected,
    revision:    revision,
    pendingHOD:  pendingHOD,
    pendingHOI:  pendingHOI,
    pendingIMO:  pendingIMO,
    pending:     pendingHOD + pendingHOI + pendingIMO,
    deptStats:   deptStats,
    zeroSubmit:  zeroSubmit,
    hodSubStats: {
      total:     hodSubTotal,
      pendingHOI:hodPendHOI,
      approvedHOI:hodApprHOI,
      pendingIMO:hodPendIMO,
      finalised: hodFinalised,
      escalated: hodEscalated
    },
    filterOptions: {
      departments: Object.keys(allDepts).sort(),
      campuses:    Object.keys(allCampuses).sort(),
      institutions:Object.keys(allInsts).sort()
    }
  };
}

// ─── HOI WEEKLY MEETING ──────────────────────────────────────────────────────
function submitHOIWeeklyMeeting(data) {
  if (!data.hoiEmail)    throw new Error('HOI email required.');
  if (!data.meetingDate) throw new Error('Meeting date is required.');
  if (!data.weekStart)   throw new Error('Week start is required.');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH.HOI_WEEKLY);
  if (!sh) {
    sh = ss.insertSheet(SH.HOI_WEEKLY);
    var hdrs = SCHEMA.HOI_WeeklyMeeting;
    var r = sh.getRange(1,1,1,hdrs.length);
    r.setValues([hdrs]).setFontWeight('bold').setBackground('#1a3c5e').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  var mid = 'WM-' + new Date().getTime();
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  sh.appendRow([
    mid,
    String(data.hoiEmail || '').trim().toLowerCase(),
    String(data.meetingDate || ''),
    String(data.institution || ''),
    String(data.weekStart || ''),
    String(data.weekEnd || ''),
    JSON.stringify(data.discussionData || []),
    now
  ]);
  return { ok: true, id: mid };
}

function getHOIWeeklyMeetings(hoiEmail) {
  if (!hoiEmail) return [];
  hoiEmail = String(hoiEmail).trim().toLowerCase();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SH.HOI_WEEKLY);
  if (!sh) return [];
  var data = sh.getDataRange().getValues(), h = data[0];
  var emI = h.indexOf('HOI_Email'), midI = h.indexOf('MeetingID');
  var dtI = h.indexOf('MeetingDate'), instI = h.indexOf('Institution');
  var wsI = h.indexOf('WeekStart'), weI = h.indexOf('WeekEnd');
  var ddI = h.indexOf('DiscussionData'), saI = h.indexOf('SubmittedAt');
  var out = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][emI]||'').trim().toLowerCase() !== hoiEmail) continue;
    var dd = [];
    try { dd = JSON.parse(String(data[i][ddI]||'[]')); } catch(e) {}
    out.push({
      meetingID:      String(data[i][midI]||''),
      meetingDate:    String(data[i][dtI]||''),
      institution:    String(data[i][instI]||''),
      weekStart:      String(data[i][wsI]||''),
      weekEnd:        String(data[i][weI]||''),
      discussionData: dd,
      submittedAt:    String(data[i][saI]||'')
    });
  }
  return out.reverse();
}

/* ---------- HOI dashboard KPIs ---------- */
function ipmHoiKpis(token) {
  const s = ipmSession_(token);
  if (s.role !== 'HOI') throw new Error('Forbidden');
  const me = String(s.username).trim().toLowerCase();
  // Exclude Drafts — they must not affect any KPI count or status
  const allSubs = ipmRowsAsObjects_(ipmSheet_('Submissions'))
    .filter(r => String(r.username).trim().toLowerCase() === me &&
                 String(r.status).trim() !== 'Draft')
    .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));

  const now = new Date();
  const yr = now.getFullYear(), mo = now.getMonth();

  // Normalise a date/string to YYYY-MM-DD Monday
  var toMonday_ = function(dt) {
    var t = new Date(dt); t.setHours(0,0,0,0);
    var dow = t.getDay(); t.setDate(t.getDate() - (dow === 0 ? 6 : dow - 1));
    return t.toISOString().slice(0, 10);
  };
  var normWk_ = function(v) {
    return String(v instanceof Date ? v.toISOString() : (v || '')).trim().slice(0, 10);
  };

  // All Mondays in current month up to today
  var getMondays_ = function(y, m) {
    var mons = [], d = new Date(y, m, 1);
    while (d.getDay() !== 1) d.setDate(d.getDate() + 1);
    while (d.getMonth() === m) { mons.push(d.toISOString().slice(0, 10)); d.setDate(d.getDate() + 7); }
    return mons;
  };

  // Map of submitted week keys → status
  var submittedKeys = {};
  allSubs.forEach(function(r) { submittedKeys[normWk_(r.weekOf)] = r.status; });

  // This week
  var thisWeekMon  = toMonday_(now);
  var thisWeekDone = !!submittedKeys[thisWeekMon];

  // Pending (overdue unsubmitted weeks this month)
  var monthMondays = getMondays_(yr, mo);
  var pendingWeeks = monthMondays.filter(function(mon) {
    return mon <= thisWeekMon && !submittedKeys[mon];
  });

  // This month count
  var thisMonth = allSubs.filter(function(r) {
    var d = new Date(r.timestamp);
    return d.getMonth() === mo && d.getFullYear() === yr;
  });

  // Streak: consecutive submitted weeks going back from last submitted
  var streak = 0;
  if (allSubs.length > 0) {
    var check = toMonday_(new Date(allSubs[0].timestamp));
    for (var i = 0; i < 26; i++) {
      if (submittedKeys[check]) {
        streak++;
        var prev = new Date(check + 'T00:00:00'); prev.setDate(prev.getDate() - 7);
        check = prev.toISOString().slice(0, 10);
      } else { break; }
    }
  }

  // Last 5 submissions for recent-activity list
  var recent = allSubs.slice(0, 5).map(function(r) {
    return {
      weekOf:    normWk_(r.weekOf),
      semester:  r.semester,
      status:    r.status,
      timestamp: r.timestamp instanceof Date ? r.timestamp.toISOString() : String(r.timestamp)
    };
  });

  // This year count
  var thisYear = allSubs.filter(function(r) {
    return new Date(r.timestamp).getFullYear() === yr;
  }).length;

  // Last 6 months submission counts
  var monthly = {};
  for (var j = 5; j >= 0; j--) {
    var d2 = new Date(yr, mo - j, 1);
    var key = Utilities.formatDate(d2, Session.getScriptTimeZone(), 'MMM yy');
    monthly[key] = 0;
  }
  allSubs.forEach(function(r) {
    var d3 = new Date(r.timestamp);
    var key = Utilities.formatDate(d3, Session.getScriptTimeZone(), 'MMM yy');
    if (monthly[key] !== undefined) monthly[key]++;
  });

  return {
    total:        allSubs.length,
    thisMonth:    thisMonth.length,
    thisYear:     thisYear,
    thisWeekDone: thisWeekDone,
    thisWeekMon:  thisWeekMon,
    pendingWeeks: pendingWeeks,
    streak:       streak,
    recent:       recent,
    monthly:      monthly,
    lastSubmitted: allSubs[0] && allSubs[0].timestamp
      ? (allSubs[0].timestamp instanceof Date
          ? allSubs[0].timestamp.toISOString()
          : String(allSubs[0].timestamp))
      : null
  };
}