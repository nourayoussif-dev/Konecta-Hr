/**
 * KONECTA EGYPT — EMPLOYEE / HR / IT WEB APP  (v2, role-aware)
 * One deployment. The app decides what each person sees from their Google login.
 *   - email in HR_ADMINS  -> HR console (pending queue, offer form, verify & issue ID)
 *   - email in IT_USERS   -> IT panel (set konecta_email on identity-verified records)
 *   - otherwise           -> that person's own self-service record
 *
 * THREE GATES
 *   Gate 1  identity  : national_id, full_name_en, full_name_ar, personal_email, mobile
 *                       -> HR matches the physical ID card -> issue employee ID + signal IT
 *   Gate 2  offer     : hire_date, job_title, function, contract_type, direct_manager, basic_salary
 *   Gate 3  payment   : insurance_number, bank_name, account_number, iban, bank_verified
 *
 * Deploy: Execute as = ME (the deploying user).  Access = Konecta domain.
 *
 *   "Execute as: Me" is load-bearing, not a preference. The script reaches the
 *   spreadsheet on its own authority, so employees never need access to a
 *   Sheet holding every salary, national ID and bank account in the company.
 *   Switching to "User accessing" would require granting all of them exactly
 *   that, and they could then open the Sheet directly and bypass every check
 *   in this file. See appsscript.json.
 */

// ================== CONFIG ==================
const SHEET_ID  = '1DrcqCT1XYdN4fiP-NYhPGR74RcEGcRhICNSo-sUZpq0';
const HR_ADMINS = ['eghr@konecta.com', 'egpersonnel@konecta.com', 'malak.tarek@konecta.com'];   // + Aliaa, Emad, Nada
const IT_USERS  = ['egit@konecta.com'];             // + IT team emails

const TAB = { EMP:'EMPLOYEES', LOG:'CHANGE LOG', LISTS:'LISTS' };
const ID_PREFIX='EG', ID_PAD=4;

const GATE1 = ['national_id','full_name_en','full_name_ar','personal_email','mobile'];
const GATE2 = ['hire_date','job_title','function','contract_type','direct_manager','basic_salary'];
const GATE3 = ['insurance_number','bank_name','account_number','iban','bank_verified'];

// Employee-editable (self-service + new joiner). Bank handled separately.
const EMPLOYEE_EDITABLE = [
  'full_name_en','full_name_ar','personal_email','passport_number','date_of_birth','nationality','religion','education_level','has_disability',
  'marital_status','dependants','dependant1_name','dependant1_dob','dependant1_relation','dependant1_national_id',
  'dependant2_name','dependant2_dob','dependant2_relation','dependant2_national_id','dependant3_name','dependant3_dob',
  'dependant3_relation','dependant3_national_id','mobile','address','city','governorate',
  'emergency_contact_name','emergency_contact_phone','emergency_contact_relation','insurance_number'
];
const BANK_FIELDS = ['bank_name','account_number','iban'];

// Visible to employee, read-only
// NOTE: 'gcm' is deliberately NOT here — employees must never see their GCM level.
const READ_ONLY_VISIBLE = [
  'employee_id','national_id','national_id_verified','konecta_email','record_status','hire_date',
  'job_title','grade','function','subfunction','contract_type','direct_manager','work_location',
  'basic_salary','kpi_target','kpi_frequency','incentive','transportation','bank_verified','payment_status'
];

// HR offer fields shown on the approval form
const HR_OFFER_FIELDS = [
  'hire_date','job_title','grade','function','subfunction','contract_type','direct_manager',
  'work_location','offer_ref','basic_salary','kpi_target','kpi_frequency','incentive','transportation',
  'training_contract','company_type','cost_centre','work_modality','employee_classification',
  'scope','digital_flag','bench_flag','variable_plan','hiring_reason','contract_time',
  'corporation_code','agreement_hours_cba','contract_hours'
];

// ================== ENTRY ==================
function doGet(){
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Konecta Egypt — My Details')
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}
function include_(f){ return HtmlService.createHtmlOutputFromFile(f).getContent(); }

// ================== HELPERS ==================
//
// PER-EXECUTION MEMO
// Apps Script charges a round trip for every call into the Spreadsheet
// service. ss_() and sheet_() were called 149 times across this file and
// headers_() 43 times, each one re-opening the file or re-reading the header
// row. Worse, headers_() called sheet_() twice on one line, so every header
// read cost two openById plus two getSheetByName.
//
// This memo lives for one execution only. A new request, or a trigger firing,
// starts with an empty one, so nothing is cached across users or across runs.
//
// Only truthy sheets are memoised: several functions call sheet_() on a tab
// that does not exist yet and then create it (TAB_CONTRACTS, TAB_APPOINTMENTS),
// and caching the miss would leave them permanently invisible.
var _MEMO = { ss:null, sheets:{}, headers:{} };

function ss_(){
  if(!_MEMO.ss) _MEMO.ss = SpreadsheetApp.openById(SHEET_ID);
  return _MEMO.ss;
}
function sheet_(n){
  if(_MEMO.sheets[n]) return _MEMO.sheets[n];
  var sh = ss_().getSheetByName(n);
  if(sh) _MEMO.sheets[n] = sh;
  return sh;
}
function headers_(n){
  if(_MEMO.headers[n]) return _MEMO.headers[n];
  var sh = sheet_(n);
  if(!sh) return [];
  var lastCol = sh.getLastColumn();
  if(lastCol < 1) return [];
  var h = sh.getRange(1,1,1,lastCol).getValues()[0];
  _MEMO.headers[n] = h;
  return h;
}
// Call if a column is added mid-execution. Data writes do not need this —
// only the header row is memoised, not row content.
function clearHeaderMemo_(n){
  if(n) delete _MEMO.headers[n]; else _MEMO.headers = {};
}
function currentUser_(){
  // Under "Execute as: Me" within the same Workspace domain, getActiveUser() still returns the visitor.
  var e = Session.getActiveUser().getEmail();
  if(!e){
    // Fail SAFE: if we cannot identify the visitor, refuse rather than guess.
    // A blank identity must never fall through to showing someone else's data.
    throw new Error('We could not verify your Konecta account. Please make sure you are signed in with your @konecta.com account and try again. If this persists, contact HR.');
  }
  return e.toLowerCase();
}
function inList_(list){ const me=currentUser_(); return list.some(function(a){return String(a).toLowerCase().trim()===me;}); }

// A function that is only ever meant to run on a time-driven trigger, from the
// Apps Script editor, or as an internal call — never straight from a browser
// via google.script.run. Under "execute as Me", a web-app visitor's ACTIVE
// user differs from the EFFECTIVE (owner) user; in a trigger or editor run the
// two are equal (both the owner). A blank active user means no web session, so
// we let it through — currentUser_ already fails closed on the real web paths.
function assertNotDirectCall_(){
  var a='', e='';
  try{ a=Session.getActiveUser().getEmail(); }catch(_){}
  try{ e=Session.getEffectiveUser().getEmail(); }catch(_){}
  if(a && e && a.toLowerCase().trim()!==e.toLowerCase().trim())
    throw new Error('This runs on a schedule and cannot be called directly.');
}
function isHR_(){ return inList_(HR_ADMINS); }
function isIT_(){ return inList_(IT_USERS); }

// ================================================================
// VISIBILITY — leavers disappear from working views.
//   Visible : Active, Serving Notice, Final Month, On Hold,
//             and anyone who left within the last 30 days.
//   Hidden  : everyone else who has left.
// The rehire check and attrition report deliberately bypass this
// and read the full history.
// ================================================================
const VISIBLE_STATUSES = ['Active','Serving Notice','Final Month','On Hold','Identity Verified','Pending'];
const LEAVER_GRACE_DAYS = 30;

function isVisibleEmployee_(status, exitDate){
  const st=String(status||'').trim();
  if(VISIBLE_STATUSES.indexOf(st)!==-1) return true;
  // recently departed: still visible while final pay and clearance are live
  if(exitDate){
    const d=new Date(exitDate);
    if(!isNaN(d)){
      const days=Math.floor((new Date()-d)/86400000);
      if(days<=LEAVER_GRACE_DAYS) return true;
    }
  }
  return false;
}

// One cached read of the employee sheet, used by every operational function.
// Cuts repeated full-sheet reads, which is most of the slowness.
function empData_(includeAll){
  const key='empdata_'+(includeAll?'all':'active');
  if(!includeAll){
    const cache=CacheService.getScriptCache();
    const hit=cache.get(key);
    if(hit){ try{ return JSON.parse(hit); }catch(e){} }
  }
  const sh=sheet_(TAB.EMP), hdr=headers_(TAB.EMP), data=sh.getDataRange().getValues();
  const si=hdr.indexOf('record_status'), xi=hdr.indexOf('exit_date'), ei=hdr.indexOf('employee_id');
  const rows=[];
  for(let r=1;r<data.length;r++){
    if(!String(data[r][ei]).trim()) continue;
    if(!includeAll && !isVisibleEmployee_(data[r][si], data[r][xi])) continue;
    rows.push({row:r+1, values:data[r]});
  }
  const out={hdr:hdr, rows:rows};
  if(!includeAll){
    try{ CacheService.getScriptCache().put(key, JSON.stringify(out), 300); }catch(e){}  // 5 minutes
  }
  return out;
}

// Call after any write, so the next read is fresh.
function clearEmpCache_(){
  try{ CacheService.getScriptCache().removeAll(['empdata_active','empdata_all']); }catch(e){}
}


function fmt_(v){ if(v instanceof Date) return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd'); return v==null?'':String(v); }

function firstFreeRow_(sh,hdr){
  // Key off employee_id, NOT national_id.
  //
  // An earlier version looked for a blank national_id and treated that row as free.
  // That was wrong and it destroyed data: a leaver who never gave a national ID, or an
  // active employee who has not filled in their details, both have real records —
  // name, employee ID, salary, history — with an empty national_id. Those rows were
  // overwritten by new joiners.
  //
  // employee_id exists on every real record from the moment it is created, so a row
  // with no employee_id is genuinely empty. Belt and braces: also check full_name_en,
  // in case a row was part-written before an error.
  const cId=hdr.indexOf('employee_id');
  const cNm=hdr.indexOf('full_name_en');
  const cNid=hdr.indexOf('national_id');
  const n=Math.max(sh.getMaxRows()-1,1);
  const width=Math.max(cId,cNm,cNid)+1;
  const block=sh.getRange(2,1,n,width).getValues();
  for(let i=0;i<block.length;i++){
    const hasId  = cId  !==-1 && String(block[i][cId] ||'').trim()!=='';
    const hasName= cNm  !==-1 && String(block[i][cNm] ||'').trim()!=='';
    const hasNid = cNid !==-1 && String(block[i][cNid]||'').trim()!=='';
    if(!hasId && !hasName && !hasNid) return i+2;      // genuinely empty
  }
  return sh.getMaxRows()+1;
}
// ================================================================
// SCHEMA GUARD
//
// Every sheet read resolves columns by header name at runtime. A renamed
// or deleted header does not throw: indexOf returns -1, row[-1] is
// undefined, and the value flows onward as if the cell were blank. The
// failure is silent and shows up later as wrong data — a payslip with no
// salary, a leave request with no approver.
//
// This guard makes it loud instead. It checks that every tab the code
// depends on exists and still carries the columns the code reads.
//
//   hrSchemaCheck()     on demand from the HR console
//   schemaDailyCheck_() at the start of leaveDailyRun — emails HR when
//                       something is broken, every day until it is fixed
//
// The EMPLOYEES list is assembled from the field-group constants at the
// top of this file, so extending a gate keeps the guard in step for free.
// The workflow tabs are listed conservatively: only columns the code
// verifiably reads, because a wrong entry here is a false alarm.
// ================================================================
function requiredSchema_(){
  const employees = {};
  [].concat(GATE1, GATE2, GATE3, EMPLOYEE_EDITABLE, BANK_FIELDS,
            READ_ONLY_VISIBLE,
            ['exit_date','exit_type','dotted_manager','project','company_type',
             'leave_entitlement','updated_at','updated_by','created_at','created_by'])
    .forEach(function(f){ employees[f]=true; });

  return {
    'EMPLOYEES': Object.keys(employees),
    'LEAVE': ['request_id','employee_id','employee_name','leave_type','track',
              'start_date','end_date','days_requested','days_approved',
              'final_status','direct_status','dotted_status',
              'direct_manager','dotted_manager','weekend_pattern','submitted_at'],
    'LEAVE_ADJUSTMENTS': ['employee_id','days','reason','adjustment_date','added_by'],
    'DELEGATES': ['manager_id','delegate_email','active','from_date','to_date'],
    'RESIGNATIONS': ['resignation_id','employee_id','final_status','proposed_last_day',
                     'withdraw_status','direct_manager','dotted_manager',
                     'submitted_at','reminder_count','last_reminder_at'],
    'CLEARANCE': ['clearance_id','employee_id','final_status',
                  'it_status','fac_status','hr_status'],
    'NO_SHOW': ['noshow_id','employee_id','hr_status','absent_since'],
    'TERMINATIONS': ['termination_id','employee_id','hr_status','final_status',
                     'direct_manager'],
    'DEPENDANTS': ['employee_id','name','relation','status','requested_at','notes'],
    'SIGNING_APPOINTMENTS': ['appointment_id','employee_id','status','appointment_date'],
    // layout-only checks: these tabs must exist, but their columns are either
    // positional (CHANGE LOG appends 13 cells), a single date column
    // (HOLIDAYS), or live on a non-standard header row (INTAKE, MANAGERS).
    'CHANGE LOG': [],
    'HOLIDAYS': [],
    'LEAVE_TYPES': []
  };
}

function schemaProblems_(){
  const want=requiredSchema_();
  const problems=[];
  Object.keys(want).forEach(function(tab){
    const sh=sheet_(tab);
    if(!sh){ problems.push({tab:tab, missing_tab:true}); return; }
    if(!want[tab].length) return;
    const hdr=headers_(tab);
    const missing=want[tab].filter(function(f){ return hdr.indexOf(f)===-1; });
    if(missing.length) problems.push({tab:tab, missing_columns:missing});
  });
  return problems;
}

// HR console: the full report.
function hrSchemaCheck(){
  if(!isHR_()) throw new Error('HR only.');
  const problems=schemaProblems_();
  return {ok:!problems.length, problems:problems,
          msg: problems.length? 'Schema problems found — see the list. Reads on these columns are coming back BLANK.'
                              : 'Every tab and column the code depends on is present.'};
}

// Called by leaveDailyRun. Emails HR while anything is broken.
function schemaDailyCheck_(){
  try{
    const problems=schemaProblems_();
    if(!problems.length) return true;
    const lines=problems.map(function(p){
      return p.missing_tab
        ? 'TAB MISSING: '+p.tab
        : p.tab+' is missing column(s): '+p.missing_columns.join(', ');
    });
    notifyHR_('SCHEMA BROKEN — the app is reading blanks',
      'A tab or column the code depends on has been renamed or deleted.\n\n'+
      lines.join('\n')+
      '\n\nUntil this is restored, every read of those columns silently returns '+
      'nothing — leave balances, approvals and payroll inputs may all be wrong. '+
      'Restore the original header names (see docs/SCHEMA.md in the repository).');
    return false;
  }catch(e){ console.error('schemaDailyCheck_ failed: '+e); return false; }
}

// ================================================================
// ROW IDENTITY GUARD
//
// The row number a browser sends can be stale: HR sorts or filters the
// sheet while an editor is open, rows shift, and a write aimed at row N
// lands on whoever occupies row N now. For salary, bank and status writes
// that is silent corruption of the wrong person's record.
//
// So no client-facing function trusts a bare row any more. The client
// sends back the identity it was given when it fetched the list (an
// employee_id, request_id, clearance_id, ...) and the write goes through
// here first:
//
//   still in place -> the same row, one cheap read to confirm.
//   moved          -> the ONE row that matches every key, found by a
//                     narrow column scan. The caller gets the new row.
//   zero or many   -> refuse with a message a human can act on. Nothing
//                     is written.
//
// Comparison uses fmt_ on the cell, because fmt_ is what built the value
// the client is holding — so dates compare as yyyy-MM-dd, not by whatever
// Date.toString happens to produce.
// ================================================================
function guardRow_(sh, hdr, row, keys, firstDataRow){
  firstDataRow = firstDataRow || 2;
  const fields = Object.keys(keys || {});
  if(!fields.length) throw new Error('guardRow_: no identity given.');
  const want = {};
  for (var k = 0; k < fields.length; k++){
    const f = fields[k];
    const v = String(keys[f] == null ? '' : keys[f]).trim();
    // A blank key is a page from before this guard existed, or a renderer
    // that failed to pass one. Refusing beats writing to an unverified row.
    if(!v) throw new Error('This page is out of date. Refresh and try again.');
    if(hdr.indexOf(f) === -1)
      throw new Error('Column "' + f + '" is missing from ' + sh.getName() +
                      ' — it was renamed or deleted. Restore it before saving.');
    want[f] = v;
  }
  const last = sh.getLastRow();
  const matchesAt = function(r){
    return fields.every(function(f){
      return fmt_(sh.getRange(r, hdr.indexOf(f) + 1).getValue()).trim() === want[f];
    });
  };
  row = parseInt(row, 10);
  if(row >= firstDataRow && row <= last && matchesAt(row)) return row;

  // The sheet changed under the client. Re-find by identity: one narrow
  // column read per key field, then intersect the matching rows.
  var hits = null;
  if(last >= firstDataRow){
    for (var k2 = 0; k2 < fields.length; k2++){
      const f2 = fields[k2];
      const col = sh.getRange(firstDataRow, hdr.indexOf(f2) + 1,
                              last - firstDataRow + 1, 1).getValues();
      const these = {};
      for (var r2 = 0; r2 < col.length; r2++){
        if(fmt_(col[r2][0]).trim() === want[f2]) these[r2 + firstDataRow] = true;
      }
      if(hits === null) hits = these;
      else Object.keys(hits).forEach(function(r3){ if(!these[r3]) delete hits[r3]; });
    }
  }
  const found = Object.keys(hits || {});
  if(found.length === 1) return parseInt(found[0], 10);

  const what = fields.map(function(f){ return f + ' ' + want[f]; }).join(', ');
  if(found.length > 1)
    throw new Error('More than one row matches ' + what +
                    ' — there is a duplicate in ' + sh.getName() + '. Fix it before saving.');
  throw new Error('Could not find the record (' + what + '). It may have been ' +
                  'changed or removed — refresh and try again.');
}

// EMPLOYEES convenience: before the ID is issued a record is identified by
// its national ID, afterwards by the employee ID. EG-prefix tells them apart.
function guardEmpRow_(sh, hdr, row, key){
  const k = String(key == null ? '' : key).trim();
  const keys = {};
  keys[k.indexOf(ID_PREFIX) === 0 ? 'employee_id' : 'national_id'] = k;
  return guardRow_(sh, hdr, row, keys);
}

function getBootstrap(){
  const role=getRole();
  const out={role:role};
  try{ out.lists=getLists(); }catch(e){ out.lists={}; }
  try{ out.projmap=getProjectMap(); }catch(e){ out.projmap={}; }
  try{ out.managers=getManagerOptions(); }catch(e){ out.managers={managers:[],globals:[],others:[]}; }
  if(!role.hr && !role.it){
    try{ out.record=getMyRecord(); }catch(e){ out.record={found:false}; }
    try{ out.team=getMyTeam(); }catch(e){ out.team={isManager:false,team:[]}; }
  }
  return out;
}

function getRole(){
  return { email: currentUser_(), hr: isHR_(), it: isIT_(), fac: isFacilities_() };
}

// ================== EMPLOYEE: read own record ==================
function getMyRecord(){
  const email=currentUser_(), sh=sheet_(TAB.EMP), hdr=headers_(TAB.EMP), data=sh.getDataRange().getValues();
  const kCol=hdr.indexOf('konecta_email');
  for(let r=1;r<data.length;r++){
    if(String(data[r][kCol]).toLowerCase().trim()===email){
      const rec={},editable={},readonly={},bank={};
      hdr.forEach(function(h,c){ const v=fmt_(data[r][c]); rec[h]=v;
        if(EMPLOYEE_EDITABLE.indexOf(h)!==-1) editable[h]=v;
        else if(BANK_FIELDS.indexOf(h)!==-1) bank[h]=v;
        else if(READ_ONLY_VISIBLE.indexOf(h)!==-1) readonly[h]=v; });

      // Where their medical cover stands, so they are not left guessing —
      // particularly if they have added dependants and are waiting.
      try{
        const med=medicalRecordFor_(rec.employee_id);
        const deps=medicalPayloadFor_(rec.employee_id).enrolled.length;
        if(med && med.status==='Enrolled'){
          readonly['medical_status'] = deps
            ? ('Enrolled — you and your '+deps+' dependant'+(deps>1?'s':'')+' are covered')
            : 'Enrolled';
        } else if(med && med.status==='Removed'){
          readonly['medical_status']='Cover ended '+(med.last_working_day||med.removed_at);
        } else {
          const chk=medicalEnrolmentReady_(rec.employee_id);
          readonly['medical_status'] = !chk.signed
            ? 'Pending contract signing'
            : 'In progress';
        }
      }catch(e){ /* medical module may not be loaded yet */ }

      // Show the bonus in the period the employee is actually on, so it matches their offer.
      // Stored value is the MONTHLY amount; annual plans see it x12.
      delete readonly.kpi_target;
      delete readonly.kpi_frequency;
      const plan = String(rec.kpi_frequency||'');
      const monthly = parseFloat(String(rec.kpi_target||'').replace(/,/g,''));
      if(plan) readonly['bonus_plan'] = plan;
      if(!isNaN(monthly) && monthly>0){
        const isAnnual = /annual|semi/i.test(plan) && !/monthly/i.test(plan);
        const amount = isAnnual ? monthly*12 : monthly;
        const label  = isAnnual ? 'bonus_annual' : 'bonus_monthly';
        readonly[label] = Math.round(amount).toLocaleString('en-US');
      }
      return {found:true,row:r+1,editable:editable,readonly:readonly,bank:bank,
        completeness:rec['completeness_%'],blocking:rec['blocking_gaps'],chase:rec['chase_gaps'],email:email};
    }
  }
  return {found:false,email:email};
}


// Project -> cost centre map, for auto-fill in the editor.

// Manager options for the direct_manager dropdown:
// every employee (EID - Name) PLUS the global managers from the MANAGERS tab.
function getManagerOptions(){
  const E=empData_(false), h=E.hdr;
  const col=function(f){return h.indexOf(f);};
  const cEid=col('employee_id'), cNm=col('full_name_en'), cDm=col('direct_manager');
  // who already manages someone? those are the real managers
  const isManager={};
  E.rows.forEach(function(rec){
    const m=String(rec.values[cDm]).trim();
    if(m) isManager[m]=true;
  });
  const managers=[], others=[];
  E.rows.forEach(function(rec){
    const id=String(rec.values[cEid]).trim(); if(!id) return;
    const item={id:id, label:id+' — '+String(rec.values[cNm]).trim()};
    (isManager[id]? managers : others).push(item);
  });
  const globals=[];
  const ms=sheet_('MANAGERS');
  if(ms && ms.getLastRow()>4){
    ms.getRange(5,1,ms.getLastRow()-4,3).getValues().forEach(function(row){
      if(!row[0]) return;
      globals.push({id:String(row[0]).trim(), label:String(row[1]).trim()+' (global)'});
    });
  }
  const byLabel=function(a,b){return a.label.localeCompare(b.label);};
  managers.sort(byLabel); others.sort(byLabel); globals.sort(byLabel);
  return {managers:managers, globals:globals, others:others};
}

function getManagerIdentity_(){
  const me=currentUser_();
  // internal: find the employee whose konecta_email == me, use their employee_id
  const sh=sheet_(TAB.EMP), hdr=headers_(TAB.EMP), data=sh.getDataRange().getValues();
  const ke=hdr.indexOf('konecta_email'), ei=hdr.indexOf('employee_id');
  for(let r=1;r<data.length;r++){
    if(String(data[r][ke]).toLowerCase().trim()===me){
      return {id:String(data[r][ei]).trim(), source:'employee'};
    }
  }
  // global: MANAGERS tab has manager_id, name, can_view, email(optional col 4)
  const ms=sheet_('MANAGERS');
  if(ms){
    const md=ms.getRange(5,1,Math.max(ms.getLastRow()-4,1),4).getValues();
    for(let i=0;i<md.length;i++){
      const email=String(md[i][3]||'').toLowerCase().trim();
      if(email && email===me && String(md[i][2]).toLowerCase().trim()==='yes'){
        return {id:String(md[i][0]).trim(), source:'global'};
      }
    }
  }
  return null;
}

function getProjectMap(){
  const sh=sheet_('PROJECT_MAP'); if(!sh) return {};
  const data=sh.getRange(5,1,Math.max(sh.getLastRow()-4,1),2).getValues();
  const m={};
  data.forEach(function(row){ if(row[0]) m[String(row[0]).trim()]=String(row[1]).trim(); });
  return m;
}

function getLists(){
  const sh=sheet_(TAB.LISTS), data=sh.getDataRange().getValues(), out={};
  data[0].forEach(function(name,c){ if(!name) return; const vals=[];
    for(let r=1;r<data.length;r++) if(data[r][c]!=='') vals.push(data[r][c]); out[name]=vals; });
  return out;
}

// ================== EMPLOYEE: writes ==================
function submitPersonalUpdate(payload){
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const me=getMyRecord(); if(!me.found) throw new Error('No record found for your account.');
    const sh=sheet_(TAB.EMP), hdr=headers_(TAB.EMP), changes=[], missingCols=[];
    Object.keys(payload).forEach(function(field){
      if(EMPLOYEE_EDITABLE.indexOf(field)===-1) return;
      const ci=hdr.indexOf(field);
      if(ci===-1){ missingCols.push(field); return; }   // column not on the sheet — skip, do not crash
      const c=ci+1, oldV=fmt_(sh.getRange(me.row,c).getValue()), newV=String(payload[field]).trim();
      if(oldV===newV) return;
      sh.getRange(me.row,c).setValue(newV); changes.push([field,oldV,newV]);
    });
    stampUpdate_(sh,hdr,me.row);
    changes.forEach(function(ch){ logChange_(me.readonly.employee_id,me.readonly.national_id,ch[0],ch[1],ch[2],'Web app','Applied','Employee self-service'); });

    // A dependant cannot be enrolled on the medical scheme without a national ID
    // and a certificate. Ask for the document as soon as one is declared.
    const newDeps=[];
    for(let i=1;i<=3;i++){
      const nameChanged = changes.some(function(c){ return c[0]==='dependant'+i+'_name' && c[2]; });
      if(!nameChanged) continue;
      const c=hdr.indexOf('dependant'+i+'_relation');
      const rel = c===-1? '' : fmt_(sh.getRange(me.row,c+1).getValue());
      const nc=hdr.indexOf('dependant'+i+'_national_id');
      const nid = nc===-1? '' : fmt_(sh.getRange(me.row,nc+1).getValue());
      const nameC=hdr.indexOf('dependant'+i+'_name');
      newDeps.push({name: nameC===-1?'':fmt_(sh.getRange(me.row,nameC+1).getValue()),
                    relation:rel, national_id:nid});
    }
    if(newDeps.length) requestDependantDocs_(me, newDeps);

    if(missingCols.length) notifyHR_('Missing column(s) on EMPLOYEES',
      'These fields are in EMPLOYEE_EDITABLE but have no column on the sheet, so they were skipped:\n\n  '+
      missingCols.join('\n  ')+'\n\nAdd the columns, or remove them from EMPLOYEE_EDITABLE.');
    return {ok:true,count:changes.length};
  } finally { lock.releaseLock(); }
}


// A dependant needs a national ID and a certificate before they can be enrolled.
// Same pattern as sick notes: reply to this email with the document attached,
// and the reference in the subject lets HR match it.
function requestDependantDocs_(me, deps){
  const to=me.email;
  if(!to) return;
  const eid=me.readonly.employee_id;
  let rows='';
  deps.forEach(function(d){
    const cert = String(d.relation).toLowerCase()==='spouse' ? 'marriage certificate' : 'birth certificate';
    rows+='<li><strong>'+escapeHtml_(d.name)+'</strong> ('+escapeHtml_(d.relation)+') — '+cert+
          (d.national_id? '' : ', and their national ID is still missing')+'</li>';
  });
  try{
    MailApp.sendEmail({
      to: to,
      cc: [MEDICAL_CONTACT].concat(HR_ADMINS).join(','),
      replyTo: HR_ADMINS.join(','),
      subject:'Action needed: dependant documents — '+eid,
      htmlBody:
        '<div style="font-family:Arial,sans-serif;max-width:520px">'+
        '<p>Hello '+escapeHtml_(me.editable.full_name_en||'')+',</p>'+
        '<p>You have added dependants to your record. Before they can be added to the medical '+
        'insurance scheme we need a document for each.</p>'+
        '<ul>'+rows+'</ul>'+
        '<div style="background:#FFF9D6;border-left:4px solid #FFE100;padding:12px 16px;margin:16px 0">'+
        '<strong>Reply to this email with the documents attached.</strong><br>'+
        'Keep the subject line as it is — it carries your employee number so we can match them.</div>'+
        '<p style="font-size:13px;color:#6b6b80">Your dependants are not covered until these are received.</p>'+
        '<p style="font-size:13px;color:#6b6b80">Konecta Egypt — People team</p></div>',
      name:'Konecta Egypt — People Team'
    });
  }catch(e){}
}

function submitBankChange(payload){
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const me=getMyRecord(); if(!me.found) throw new Error('No record found.');
    const iban=String(payload.iban||'').replace(/\s/g,'').toUpperCase();
    if(!/^EG\d{27}$/.test(iban)) return {ok:false,msg:'IBAN must be EG followed by 27 digits.'};
    const sh=sheet_(TAB.EMP),hdr=headers_(TAB.EMP);
    BANK_FIELDS.forEach(function(f){
      const v=f==='iban'?iban:String(payload[f]||'').trim();
      // write the value to the record but leave it UNVERIFIED — payroll won't pay until HR verifies
      const c=hdr.indexOf(f); if(c!==-1) sh.getRange(me.row,c+1).setValue(v);
      logChange_(me.readonly.employee_id,me.readonly.national_id,f,me.bank[f],v,'Web app','Pending approval',
        'Bank change submitted. HR must verify against the ID before it is used for payment.');
    });
    const bv=hdr.indexOf('bank_verified')+1;
    if(bv>0) sh.getRange(me.row,bv).setValue('Pending verification');
    stampUpdate_(sh,hdr,me.row);
    notifyHR_('Bank change — verify required', me.readonly.employee_id+' ('+me.email+') submitted new bank details.\nVerify against the ID, then set bank_verified to Verified.\n\nBank: '+payload.bank_name+'\nAccount: '+payload.account_number+'\nIBAN: '+iban);
    return {ok:true,msg:'Submitted. HR will verify against your ID before this takes effect.'};
  } finally { lock.releaseLock(); }
}

function reportIssue(field,comment){
  const me=getMyRecord(); if(!me.found) throw new Error('No record found.');
  logChange_(me.readonly.employee_id,me.readonly.national_id,field,me.readonly[field]||me.editable[field]||'','(reported incorrect)','Web app','Pending approval','EMPLOYEE REPORT: '+comment);
  notifyHR_('Data reported incorrect — '+me.readonly.employee_id, me.email+' reports "'+field+'" wrong.\nComment: '+comment);
  return {ok:true};
}

// ================== NEW JOINER ==================
function submitNewEmployee(payload){
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const nid=String(payload.national_id||'').trim();
    const v=validateNationalId(nid); if(!v.valid) return {ok:false,msg:v.msg};
    const dupe=findByNationalId_(nid);
    if(dupe){
      if(String(dupe.record_status).toLowerCase().indexOf('closed')!==-1){
        appendRecord_(payload,nid,'Blocked - possible rehire',dupe.employee_id);
        notifyHR_('BLOCKED — possible rehire','National ID '+nid+' matches Closed record '+dupe.employee_id+'. No ID issued.');
        return {ok:false,blocked:true,msg:'Your record is held for HR review because our records show previous employment. HR will contact you.'};
      }
      return {ok:false,msg:'This national ID already exists on an active record. Please speak to HR.'};
    }
    appendRecord_(payload,nid,'Pending','');
    notifyHR_('New joiner submitted — verify ID & complete offer','A new joiner completed identity details.\nNational ID: '+nid+'\nName: '+payload.full_name_en+'\nPersonal email: '+payload.personal_email+'\n\nMatch the physical ID card, then complete the offer and issue the ID.');
    return {ok:true,msg:'Thank you. Your details are with HR. You will receive your employee ID by email once approved.'};
  } finally { lock.releaseLock(); }
}

function appendRecord_(payload,nid,status,prevId){
  const sh=sheet_(TAB.EMP),hdr=headers_(TAB.EMP),row=firstFreeRow_(sh,hdr);
  const write=function(f,val){ const c=hdr.indexOf(f); if(c!==-1) sh.getRange(row,c+1).setValue(val); };
  write('national_id',nid); write('record_status',status); write('previous_employee_id',prevId);
  write('national_id_verified','Pending');
  EMPLOYEE_EDITABLE.concat(BANK_FIELDS).forEach(function(f){ if(payload[f]!==undefined) write(f,String(payload[f]).trim()); });
  write('bank_verified','Pending verification');
  write('created_at',new Date()); write('created_by',currentUser_());
  logChange_('',nid,'record_status','',status,'Web app','Applied','New joiner submission');
  return row;
}

// ================== HR CONSOLE ==================
function hrGetPending(){
  if(!isHR_()) throw new Error('HR only.');
  const sh=sheet_(TAB.EMP),hdr=headers_(TAB.EMP),data=sh.getDataRange().getValues();
  const idx=function(f){return hdr.indexOf(f);};
  const out=[];
  for(let r=1;r<data.length;r++){
    const nid=data[r][idx('national_id')]; if(!nid) continue;
    const st=String(data[r][idx('record_status')]);
    if(['Pending','Identity Verified','Blocked - possible rehire'].indexOf(st)===-1) continue;
    const rec={row:r+1};
    hdr.forEach(function(h,c){ rec[h]=fmt_(data[r][c]); });
    out.push(rec);
  }
  return out;
}

// HR confirms the physical ID card matches, and issues the employee ID
function hrVerifyAndIssue(row, key){
  if(!isHR_()) throw new Error('HR only.');
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const sh=sheet_(TAB.EMP),hdr=headers_(TAB.EMP);
    row=guardRow_(sh,hdr,row,{national_id:key});
    const get=function(f){ return sh.getRange(row,hdr.indexOf(f)+1).getValue(); };
    const set=function(f,v){ const c=hdr.indexOf(f); if(c!==-1) sh.getRange(row,c+1).setValue(v); };

    if(get('employee_id')) throw new Error('This record already has an employee ID.');
    // Gate 1 must be complete
    const missing=GATE1.filter(function(f){ return !String(get(f)).trim(); });
    if(missing.length) return {ok:false,msg:'Cannot issue — missing: '+missing.join(', ')};

    // generate ID under lock
    const idCol=hdr.indexOf('employee_id')+1;
    const existing=sh.getRange(2,idCol,Math.max(sh.getMaxRows()-1,1),1).getValues()
      .map(function(r){return String(r[0]);}).filter(function(s){return s.indexOf(ID_PREFIX)===0;})
      .map(function(s){return parseInt(s.replace(ID_PREFIX,''),10);}).filter(function(n){return !isNaN(n);});
    const next=(existing.length?Math.max.apply(null,existing):0)+1;
    const id=ID_PREFIX+String(next).padStart(ID_PAD,'0');

    set('employee_id',id);
    set('national_id_verified','Verified');
    set('record_status','Identity Verified');
    stampUpdate_(sh,hdr,row);
    logChange_(id,get('national_id'),'employee_id','',id,'HR console','Applied','ID card matched by HR. ID issued.');

    // signal IT
    notifyIT_('Create Konecta email — '+id, 'Identity verified for '+get('full_name_en')+' ('+id+').\nPlease create the Konecta email and enter it in the IT panel.');
    // email joiner
    const pe=get('personal_email');
    if(pe) MailApp.sendEmail(pe,'Welcome to Konecta Egypt — your employee ID','Your identity has been verified.\n\nEmployee ID: '+id+'\n\nYour Konecta email and system access will follow shortly.\n\nKonecta Egypt — People team');
    return {ok:true,id:id};
  } finally { lock.releaseLock(); }
}

// HR saves offer fields
function hrSaveOffer(row,payload,key){
  if(!isHR_()) throw new Error('HR only.');
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const sh=sheet_(TAB.EMP),hdr=headers_(TAB.EMP),changes=[];
    row=guardEmpRow_(sh,hdr,row,key);
    HR_OFFER_FIELDS.forEach(function(f){
      if(payload[f]===undefined) return;
      const c=hdr.indexOf(f)+1; if(c===0) return;
      const oldV=fmt_(sh.getRange(row,c).getValue()), newV=String(payload[f]).trim();
      if(oldV===newV) return;
      sh.getRange(row,c).setValue(newV); changes.push([f,oldV,newV]);
    });
    // activate if all gate-2 present and ID already issued
    const get=function(f){ return String(sh.getRange(row,hdr.indexOf(f)+1).getValue()).trim(); };
    if(get('employee_id') && GATE2.every(function(f){return get(f);}) && get('record_status')==='Identity Verified'){
      sh.getRange(row,hdr.indexOf('record_status')+1).setValue('Active');
    }
    stampUpdate_(sh,hdr,row);
    const eid=get('employee_id'), nid=get('national_id');
    changes.forEach(function(ch){ logChange_(eid,nid,ch[0],ch[1],ch[2],'HR console','Applied','Offer detail set by HR'); });
    return {ok:true,count:changes.length,active:get('record_status')==='Active'};
  } finally { lock.releaseLock(); }
}

// HR verifies bank

// Employees who have submitted bank details but are not yet verified.
function hrGetBankPending(){
  if(!isHR_()) throw new Error('HR only.');
  const E=empData_(false), h=E.hdr;
  const col=function(f){return h.indexOf(f);};
  const cEid=col('employee_id'), cNm=col('full_name_en'), cNid=col('national_id'),
        cBv=col('bank_verified'), cBn=col('bank_name'), cAc=col('account_number'), cIb=col('iban');
  const out=[];
  E.rows.forEach(function(rec){
    const v=rec.values;
    const bv=String(v[cBv]);
    if(bv==='Verified'||bv==='') return;
    if(!String(v[cBn]).trim() && !String(v[cAc]).trim() && !String(v[cIb]).trim()) return;
    out.push({row:rec.row, employee_id:fmt_(v[cEid]), full_name_en:fmt_(v[cNm]),
              national_id:fmt_(v[cNid]), bank_name:fmt_(v[cBn]),
              account_number:fmt_(v[cAc]), iban:fmt_(v[cIb]), bank_verified:bv});
  });
  return out;
}

function hrVerifyBank(row, decision, key){
  if(!isHR_()) throw new Error('HR only.');
  const sh=sheet_(TAB.EMP),hdr=headers_(TAB.EMP);
  row=guardRow_(sh,hdr,row,{employee_id:key});
  const val = decision==='reject' ? 'Rejected' : 'Verified';
  const old = fmt_(sh.getRange(row,hdr.indexOf('bank_verified')+1).getValue());
  sh.getRange(row,hdr.indexOf('bank_verified')+1).setValue(val);
  stampUpdate_(sh,hdr,row);
  const eid=sh.getRange(row,hdr.indexOf('employee_id')+1).getValue();
  const nid=sh.getRange(row,hdr.indexOf('national_id')+1).getValue();
  logChange_(eid,nid,'bank_verified',old,val,'HR console','Applied',
    val==='Verified'?'Bank details verified against ID by HR':'Bank details rejected by HR — employee must resubmit');
  return {ok:true, status:val};
}


// ================================================================
// HR FULL-RECORD EDITOR — open any record, edit any field, all logged.
// Nobody touches the sheet.
// ================================================================

// Search by employee_id (exact) or name (partial, case-insensitive)

// Records needing HR attention: incomplete, awaiting offer, bank unverified, or reporting stale.
// Fields HR is responsible for. Employee-owned gaps are deliberately NOT listed here —
// those get chased by reminder to the employee, not parked on HR's desk.
const HR_OWNED_FIELDS = [
  'hire_date','job_title','grade','function','subfunction','contract_type','contract_time',
  'direct_manager','project','cost_centre','work_location','basic_salary','kpi_frequency',
  'company_type','work_modality','employee_classification','scope','corporation_code',
  'agreement_hours_cba','contract_hours','n_level','gcm','weekend_pattern'
];

function hrGetTaskList(){
  if(!isHR_()) throw new Error('HR only.');
  const E=empData_(false), h=E.hdr;
  const col=function(f){return h.indexOf(f);};
  const cEid=col('employee_id'), cNm=col('full_name_en'), cNid=col('national_id'),
        cSt=col('record_status'), cBv=col('bank_verified'), cIb=col('iban'),
        cRv=col('reporting_validated'), cKe=col('konecta_email');
  // pre-resolve the HR field columns once
  const hrCols=HR_OWNED_FIELDS.map(function(f){return {f:f, c:col(f)};})
                              .filter(function(x){return x.c!==-1;});
  const out=[];
  E.rows.forEach(function(rec){
    const v=rec.values;
    const status=String(v[cSt]);
    const reasons=[], missing=[];

    // identity is HR's to resolve — they match the card
    if(!String(v[cNid]).trim()) reasons.push('NATIONAL ID MISSING');
    if(status==='Identity Verified') reasons.push('Offer details needed');
    if(!String(v[cKe]).trim() && status!=='Pending') reasons.push('No Konecta email');

    // which HR-owned fields are blank?
    hrCols.forEach(function(x){
      if(!String(v[x.c]).trim()) missing.push(x.f);
    });
    if(missing.length) reasons.push(missing.length+' field(s) for you to fill');

    // actions that sit with HR
    if(String(v[cIb]).trim() && String(v[cBv])!=='Verified' && String(v[cBv])!=='Rejected')
      reasons.push('Bank to verify');
    if(String(v[cRv])==='Needs review') reasons.push('Reporting to re-check');

    if(!reasons.length) return;
    out.push({row:rec.row, employee_id:fmt_(v[cEid]), full_name_en:fmt_(v[cNm]),
              national_id:fmt_(v[cNid]),
              record_status:status, missingCount:missing.length,
              missing:missing, reasons:reasons});
  });
  // most outstanding HR work first
  out.sort(function(a,b){ return b.missingCount-a.missingCount; });
  return out;
}

function hrSearchEmployees(term){
  if(!isHR_()) throw new Error('HR only.');
  term=String(term||'').trim().toLowerCase();
  if(!term) return [];
  const E=empData_(false), h=E.hdr;
  const col=function(f){return h.indexOf(f);};
  const cEid=col('employee_id'), cNm=col('full_name_en'), cNid=col('national_id'),
        cSt=col('record_status'), cComp=col('completeness_%');
  const out=[];
  for(let i=0;i<E.rows.length && out.length<25;i++){
    const v=E.rows[i].values;
    const hay=(String(v[cEid])+' '+String(v[cNm])+' '+String(v[cNid])).toLowerCase();
    if(hay.indexOf(term)===-1) continue;
    out.push({row:E.rows[i].row, employee_id:fmt_(v[cEid]), full_name_en:fmt_(v[cNm]),
              national_id:fmt_(v[cNid]),
              record_status:fmt_(v[cSt]), completeness:fmt_(v[cComp])});
  }
  return out;
}

function hrGetRecord(row, key){
  if(!isHR_()) throw new Error('HR only.');
  const sh=sheet_(TAB.EMP), hdr=headers_(TAB.EMP);
  row=guardEmpRow_(sh,hdr,row,key);
  const vals=sh.getRange(row,1,1,hdr.length).getValues()[0];
  const rec={}; hdr.forEach(function(h,c){ rec[h]=fmt_(vals[c]); });
  rec._row=row;
  return rec;
}

// Fields HR may NOT edit (system-managed / derived). Everything else is editable.
var HR_LOCKED = ['employee_id','completeness_%','blocking_gaps','chase_gaps','report_name',
  'report_surname','citizenship_code','gender_code','has_disability_code','training_contract_code',
  'company_type_code','contract_type_code','contract_time_code','exit_type_code',
  'created_at','created_by','updated_at','updated_by','insurance_wage'];

function hrSaveRecord(row, payload, key){
  if(!isHR_()) throw new Error('HR only.');
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const sh=sheet_(TAB.EMP), hdr=headers_(TAB.EMP);
    row=guardEmpRow_(sh,hdr,row,key);
    const eid=sh.getRange(row,hdr.indexOf('employee_id')+1).getValue();
    const nid=sh.getRange(row,hdr.indexOf('national_id')+1).getValue();
    const changes=[];
    Object.keys(payload).forEach(function(field){
      const c=hdr.indexOf(field);
      if(c===-1) return;
      if(HR_LOCKED.indexOf(field)!==-1) return;             // never write locked fields
      const oldV=fmt_(sh.getRange(row,c+1).getValue());
      const newV=String(payload[field]).trim();
      if(oldV===newV) return;
      sh.getRange(row,c+1).setValue(newV);
      changes.push([field,oldV,newV]);
    });
    stampUpdate_(sh,hdr,row);
    changes.forEach(function(ch){ logChange_(eid,nid,ch[0],ch[1],ch[2],'HR editor','Applied','Edited by HR'); });
    return {ok:true,count:changes.length};
  } finally { lock.releaseLock(); }
}

// ================== IT PANEL ==================
function itGetQueue(){
  if(!isIT_()&&!isHR_()) throw new Error('IT only.');
  const E=empData_(false), h=E.hdr;
  const col=function(f){return h.indexOf(f);};
  const cEid=col('employee_id'), cNm=col('full_name_en'), cPe=col('personal_email'),
        cNv=col('national_id_verified'), cKe=col('konecta_email');
  const out=[];
  E.rows.forEach(function(rec){
    const v=rec.values;
    if(String(v[cNv])!=='Verified') return;
    if(String(v[cKe]).trim()) return;
    out.push({row:rec.row, employee_id:fmt_(v[cEid]),
              full_name_en:fmt_(v[cNm]), personal_email:fmt_(v[cPe])});
  });
  return out;
}

function itSetEmail(row,email,key){
  if(!isIT_()&&!isHR_()) throw new Error('IT only.');
  email=String(email).trim().toLowerCase();
  if(!/^[^@]+@konecta\.com$/.test(email)) return {ok:false,msg:'Must be a @konecta.com address.'};
  const sh=sheet_(TAB.EMP),hdr=headers_(TAB.EMP);
  row=guardRow_(sh,hdr,row,{employee_id:key});
  sh.getRange(row,hdr.indexOf('konecta_email')+1).setValue(email);
  stampUpdate_(sh,hdr,row);
  const eid=sh.getRange(row,hdr.indexOf('employee_id')+1).getValue();
  const nid=sh.getRange(row,hdr.indexOf('national_id')+1).getValue();
  logChange_(eid,nid,'konecta_email','',email,'IT panel','Applied','Konecta email created by IT');
  // activate + email the new hire to finish their details in the web app
  const stCol=hdr.indexOf('record_status')+1;
  if(String(sh.getRange(row,stCol).getValue())==='Identity Verified') sh.getRange(row,stCol).setValue('Active');
  const pe=sh.getRange(row,hdr.indexOf('personal_email')+1).getValue();
  const nm=sh.getRange(row,hdr.indexOf('full_name_en')+1).getValue();
  const url=ScriptApp.getService().getUrl();
  if(pe){
    MailApp.sendEmail({
      to: pe,
      subject: 'Welcome to Konecta Egypt — your account is ready',
      htmlBody: welcomeEmailHtml_(nm, email, url),
      name: 'Konecta Egypt — People Team'
    });
  }
  return {ok:true};
}



// ================================================================
// STAGE 1 — PUBLIC FORM INTAKE
// The public Google Form writes to the INTAKE tab. This trigger validates
// each submission automatically. HR then approves, which creates the real
// record and Employee ID. Nothing reaches EMPLOYEES without HR approval.
// ================================================================

const TAB_INTAKE = 'INTAKE';

// Map the form's question titles to intake columns.
// EDIT the left-hand titles to match your form questions EXACTLY.
const FORM_MAP = {
  'National ID'          : 'national_id',
  'Full name (English)'  : 'full_name_en',
  'Full name (Arabic)'   : 'full_name_ar',
  'Personal email'       : 'personal_email',
  'Mobile number'        : 'mobile',
  'Job title'            : 'intake_job_title',
  'Account / team'       : 'intake_account'
};

// Attach this to the FORM (not the sheet): Triggers -> add trigger ->
// onFormSubmit -> From form -> On form submit.
function onFormSubmit(e){
  // A real Forms submission carries namedValues. A browser google.script.run
  // call passes nothing — refuse it, so intake rows can only come from the form.
  if(!e || !e.namedValues) throw new Error('This is triggered by the intake form, not callable directly.');
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const sh=sheet_(TAB_INTAKE), hdr=intakeHeaders_();
    const vals={};
    // e.namedValues: { 'Question title': ['answer'], ... }
    Object.keys(e.namedValues||{}).forEach(function(q){
      const col=FORM_MAP[q.trim()];
      if(col) vals[col]=String(e.namedValues[q][0]||'').trim();
    });

    const nid=vals.national_id||'';
    const v=validateNationalId(nid);
    const dupe=v.valid?findByNationalId_(nid):null;
    let status='Validated - ready', note=v.valid?'':v.msg, dedup='', prev='';
    if(!v.valid){ status='REVIEW - invalid ID'; }
    else if(dupe){
      prev=dupe.employee_id||'';
      if(String(dupe.record_status).toLowerCase().indexOf('closed')!==-1){ status='REVIEW - possible rehire'; dedup='Matches CLOSED record'; }
      else { status='REVIEW - possible rehire'; dedup='Matches ACTIVE record'; }
    }

    const row=firstFreeRow_(sh,hdr);
    const set=function(f,x){ const c=hdr.indexOf(f); if(c!==-1) sh.getRange(row,c+1).setValue(x); };
    set('intake_id','INT-'+String(row).padStart(5,'0'));
    set('timestamp',new Date());
    Object.keys(vals).forEach(function(f){ set(f,vals[f]); });
    set('id_valid', v.valid?'Yes':'No');
    set('id_check_note', note || (v.valid?('Born '+v.dob+', '+v.governorate):''));
    set('dedup_flag', dedup);
    set('previous_employee_id', prev);
    set('review_status', status);

    notifyHR_('New form submission — '+(status.indexOf('REVIEW')===0?'NEEDS REVIEW':'ready to approve'),
      (vals.full_name_en||'(no name)')+'  '+(vals.intake_job_title||'')+' / '+(vals.intake_account||'')+
      '\nNational ID: '+nid+'\nStatus: '+status+(dedup?('\n'+dedup+' — previous ID '+prev):'')+
      '\n\nOpen the HR console to review and approve.');
  } finally { lock.releaseLock(); }
}

// HR console: list intake rows awaiting action
// INTAKE tab has its header on ROW 4 (rows 1-3 are title/notes). Data starts row 5.
function intakeHeaders_(){
  const sh=sheet_(TAB_INTAKE);
  return sh.getRange(4,1,1,sh.getLastColumn()).getValues()[0];
}
function hrGetIntake(){
  if(!isHR_()) throw new Error('HR only.');
  const sh=sheet_(TAB_INTAKE), hdr=intakeHeaders_(), data=sh.getDataRange().getValues();
  const out=[];
  for(let r=4;r<data.length;r++){            // data rows start at index 4 (sheet row 5)
    const rec={}; hdr.forEach(function(h,c){ rec[h]=fmt_(data[r][c]); });
    if(!rec.national_id) continue;
    if(['New','Validated - ready','REVIEW - possible rehire','REVIEW - invalid ID'].indexOf(rec.review_status)===-1) continue;
    rec._row=r+1; out.push(rec);
  }
  return out;
}

// HR approves an intake row -> creates the EMPLOYEES record + issues the Employee ID
function hrApproveIntake(intakeRow, key){
  if(!isHR_()) throw new Error('HR only.');
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const ish=sheet_(TAB_INTAKE), ihdr=intakeHeaders_();
    intakeRow=guardRow_(ish,ihdr,intakeRow,{national_id:key},5);
    const iget=function(f){ return fmt_(ish.getRange(intakeRow,ihdr.indexOf(f)+1).getValue()); };
    const nid=iget('national_id');
    if(!nid) throw new Error('No national ID on this intake row.');
    if(iget('created_employee_id')) throw new Error('This submission was already approved.');

    // final dedup guard at the moment of creation
    const dupe=findByNationalId_(nid);
    if(dupe) return {ok:false,msg:'A record with this national ID already exists ('+ (dupe.employee_id||'no ID') +'). Not created.'};

    // create the EMPLOYEES record
    const esh=sheet_(TAB.EMP), ehdr=headers_(TAB.EMP), erow=firstFreeRow_(esh,ehdr);
    const eset=function(f,x){ const c=ehdr.indexOf(f); if(c!==-1) esh.getRange(erow,c+1).setValue(x); };
    eset('national_id',nid);
    eset('full_name_en',iget('full_name_en'));
    eset('full_name_ar',iget('full_name_ar'));
    eset('personal_email',iget('personal_email'));
    eset('mobile',iget('mobile'));
    eset('national_id_verified','Verified');          // HR approval IS the verification
    eset('record_status','Identity Verified');
    eset('bank_verified','Pending verification');
    eset('created_at',new Date()); eset('created_by',currentUser_());

    // issue Employee ID under the same lock
    const idCol=ehdr.indexOf('employee_id')+1;
    const existing=esh.getRange(2,idCol,Math.max(esh.getMaxRows()-1,1),1).getValues()
      .map(function(r){return String(r[0]);}).filter(function(x){return x.indexOf(ID_PREFIX)===0;})
      .map(function(x){return parseInt(x.replace(ID_PREFIX,''),10);}).filter(function(n){return !isNaN(n);});
    const id=ID_PREFIX+String((existing.length?Math.max.apply(null,existing):0)+1).padStart(ID_PAD,'0');
    eset('employee_id',id);

    logChange_(id,nid,'employee_id','',id,'HR intake approval','Applied','Approved from public form. ID card validated by HR.');

    // mark intake row done
    const iset=function(f,x){ const c=ihdr.indexOf(f); if(c!==-1) ish.getRange(intakeRow,c+1).setValue(x); };
    iset('review_status','Approved'); iset('approved_by',currentUser_());
    iset('approved_at',new Date()); iset('created_employee_id',id);

    // signal IT to create the Konecta email
    notifyIT_('Create Konecta email — '+id,
      'Approved new hire:\n\nName: '+iget('full_name_en')+'\nEmployee ID: '+id+
      '\nJob/Account: '+iget('intake_job_title')+' / '+iget('intake_account')+
      '\n\nCreate the Konecta email and enter it in the IT panel. The new hire is emailed automatically once you do.');

    return {ok:true,id:id};
  } finally { lock.releaseLock(); }
}

function hrRejectIntake(intakeRow,reason,key){
  if(!isHR_()) throw new Error('HR only.');
  const ish=sheet_(TAB_INTAKE), ihdr=intakeHeaders_();
  intakeRow=guardRow_(ish,ihdr,intakeRow,{national_id:key},5);
  const set=function(f,x){ const c=ihdr.indexOf(f); if(c!==-1) ish.getRange(intakeRow,c+1).setValue(x); };
  set('review_status','Rejected'); set('approved_by',currentUser_()); set('approved_at',new Date());
  set('notes',reason||'');
  return {ok:true};
}


// ================================================================
// MANAGER VIEW — who reports to me, and a read-only limited card.
// ================================================================

// Does the logged-in user manage anyone? Return their team (limited fields).
function getMyTeam(){
  const identity=getManagerIdentity_();
  if(!identity) return {isManager:false, team:[]};
  const E=empData_(false), h=E.hdr;
  const col=function(f){return h.indexOf(f);};
  const cEid=col('employee_id'), cDm=col('direct_manager'),
        cNm=col('full_name_en'), cJt=col('job_title'), cPr=col('project');
  const team=[];
  E.rows.forEach(function(rec){
    const v=rec.values;
    if(String(v[cDm]).trim()!==identity.id) return;      // direct reports only
    team.push({row:rec.row, employee_id:fmt_(v[cEid]), full_name_en:fmt_(v[cNm]),
               job_title:fmt_(v[cJt]), project:fmt_(v[cPr])});
  });
  return {isManager:true, myId:identity.id, team:team};
}

function getTeamMemberCard(row, key){
  const identity=getManagerIdentity_();
  if(!identity) throw new Error('Not authorised.');
  const sh=sheet_(TAB.EMP), hdr=headers_(TAB.EMP);
  row=guardRow_(sh,hdr,row,{employee_id:key});
  const vals=sh.getRange(row,1,1,hdr.length).getValues()[0];
  const get=function(f){ const c=hdr.indexOf(f); return c===-1?'':fmt_(vals[c]); };
  // authorisation: this person must report to the caller
  if(String(get('direct_manager')).trim()!==identity.id) throw new Error('Not your report.');
  // age from DOB
  var age='';
  const dob=get('date_of_birth');
  if(dob){ const d=new Date(dob); if(!isNaN(d)){ const t=new Date(); age=t.getFullYear()-d.getFullYear()-((t.getMonth()<d.getMonth()||(t.getMonth()===d.getMonth()&&t.getDate()<d.getDate()))?1:0); } }
  return {
    full_name_en:get('full_name_en'),
    konecta_email:get('konecta_email'),
    mobile:get('mobile'),
    job_title:get('job_title'),
    hire_date:get('hire_date'),
    date_of_birth:dob,
    age:age,
    project:get('project'),
    basic_salary:get('basic_salary'),
    grade:get('grade'),
    gcm:get('gcm'),
    has_disability:get('has_disability'),
    direct_manager:get('direct_manager'),
    dotted_manager:get('dotted_manager')
  };
}

// ================================================================
// N-LEVEL — reporting depth from the top of the chain.
// Walks direct_manager upward and counts hops. Run after manager changes.
// ================================================================
const NLEVEL_TOP = 'GLOBAL-NOUREDIN';   // top of the reporting tree

function recalcNLevels(){
  if(!isHR_()) throw new Error('HR only.');
  const sh=sheet_(TAB.EMP), hdr=headers_(TAB.EMP), data=sh.getDataRange().getValues();
  const ei=hdr.indexOf('employee_id'), dm=hdr.indexOf('direct_manager'), nl=hdr.indexOf('n_level');
  if(nl===-1) throw new Error('Column n_level not found. Add it to the sheet first.');

  const mgrOf={}, rowOf={};
  for(let r=1;r<data.length;r++){
    const id=String(data[r][ei]).trim(); if(!id) continue;
    mgrOf[id]=String(data[r][dm]).trim();
    rowOf[id]=r+1;
  }
  const cache={};
  function depth(id, seen){
    if(cache[id]!==undefined) return cache[id];
    seen=seen||{};
    if(seen[id]) return null;             // cycle guard
    seen[id]=true;
    const m=mgrOf[id];
    if(!m) return (cache[id]=null);
    if(m===NLEVEL_TOP) return (cache[id]=1);
    const up=depth(m,seen);
    return (cache[id] = (up===null?null:up+1));
  }
  let written=0, broken=0;
  Object.keys(mgrOf).forEach(function(id){
    const d=depth(id);
    const val = d===null ? '' : ('N-'+d);
    if(d===null) broken++;
    sh.getRange(rowOf[id], nl+1).setValue(val);
    written++;
  });
  return {ok:true, written:written, unresolved:broken};
}

// Fields a DIRECT MANAGER may edit on their own reports. Nothing else.
const MANAGER_EDITABLE = ['grade','gcm'];

function managerSaveTeamMember(row, payload, key){
  const identity=getManagerIdentity_();
  if(!identity) throw new Error('Not authorised.');
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const sh=sheet_(TAB.EMP), hdr=headers_(TAB.EMP);
    row=guardRow_(sh,hdr,row,{employee_id:key});
    const get=function(f){ const c=hdr.indexOf(f); return c===-1?'':fmt_(sh.getRange(row,c+1).getValue()); };
    // must be their own direct report
    if(String(get('direct_manager')).trim()!==identity.id) throw new Error('Not your report.');

    const eid=get('employee_id'), nid=get('national_id');
    const changes=[];
    MANAGER_EDITABLE.forEach(function(f){
      if(payload[f]===undefined) return;
      const c=hdr.indexOf(f); if(c===-1) return;
      const oldV=fmt_(sh.getRange(row,c+1).getValue());
      const newV=String(payload[f]).trim();
      if(oldV===newV) return;
      sh.getRange(row,c+1).setValue(newV);
      changes.push([f,oldV,newV]);
    });
    stampUpdate_(sh,hdr,row);
    changes.forEach(function(ch){
      logChange_(eid,nid,ch[0],ch[1],ch[2],'Manager view','Applied','Set by direct manager '+identity.id);
    });
    return {ok:true,count:changes.length};
  } finally { lock.releaseLock(); }
}
// ================== NATIONAL ID ==================
const GOVS={'01':'Cairo','02':'Alexandria','03':'Port Said','04':'Suez','11':'Damietta','12':'Dakahlia','13':'Sharqia','14':'Qalyubia','15':'Kafr El Sheikh','16':'Gharbia','17':'Monufia','18':'Beheira','19':'Ismailia','21':'Giza','22':'Beni Suef','23':'Fayoum','24':'Minya','25':'Asyut','26':'Sohag','27':'Qena','28':'Aswan','29':'Luxor','31':'Red Sea','32':'New Valley','33':'Matrouh','34':'North Sinai','35':'South Sinai','88':'Born abroad'};
function validateNationalId(nid){
  nid=String(nid).trim();
  if(!/^\d{14}$/.test(nid)) return {valid:false,msg:'National ID must be exactly 14 digits.'};
  const c=nid[0]; if(c!=='2'&&c!=='3') return {valid:false,msg:'Invalid century digit.'};
  const yy=nid.substr(1,2),mm=nid.substr(3,2),dd=nid.substr(5,2);
  const year=(c==='2'?1900:2000)+parseInt(yy,10),m=parseInt(mm,10),d=parseInt(dd,10);
  if(m<1||m>12||d<1||d>31) return {valid:false,msg:'Invalid date of birth in the ID.'};
  const dob=new Date(year,m-1,d);
  if(dob.getMonth()!==m-1||dob.getDate()!==d) return {valid:false,msg:'Date of birth in the ID is not real.'};
  const gov=nid.substr(7,2); if(!GOVS[gov]) return {valid:false,msg:'Unrecognised governorate code.'};
  return {valid:true,dob:Utilities.formatDate(dob,Session.getScriptTimeZone(),'yyyy-MM-dd'),
    governorate:GOVS[gov],gender:(parseInt(nid[12],10)%2===1)?'Male':'Female'};
}
function findByNationalId_(nid){
  const sh=sheet_(TAB.EMP),hdr=headers_(TAB.EMP),data=sh.getDataRange().getValues(),c=hdr.indexOf('national_id');
  for(let r=1;r<data.length;r++){ if(String(data[r][c]).trim()===String(nid).trim()){
    const o={}; hdr.forEach(function(h,i){o[h]=data[r][i];}); o._row=r+1; return o; } }
  return null;
}

// ================== LOG / UTIL ==================
function logChange_(empId,nid,field,oldV,newV,source,status,notes){
  const sh=sheet_(TAB.LOG),row=sh.getLastRow()+1,id='LOG-'+String(row).padStart(6,'0');
  sh.getRange(row,1,1,13).setValues([[id,new Date(),empId,nid,field,oldV,newV,currentUser_(),source,'','',status,notes]]);
}
function stampUpdate_(sh,hdr,row){
  clearEmpCache_();          // any write invalidates the cached read
  const ua=hdr.indexOf('updated_at'),ub=hdr.indexOf('updated_by');
  if(ua!==-1) sh.getRange(row,ua+1).setValue(new Date());
  if(ub!==-1) sh.getRange(row,ub+1).setValue(currentUser_());
}

// Branded welcome email (Konecta indigo #2800C8, lilac #EEEDFE, yellow #FFE100).
function welcomeEmailHtml_(name, konectaEmail, url){
  return '' +
'<div style="margin:0;padding:0;background:#EEEDFE;font-family:Arial,Helvetica,sans-serif">' +
  '<div style="max-width:560px;margin:0 auto;padding:24px 16px">' +
    '<div style="background:#2800C8;border-radius:12px 12px 0 0;padding:28px 28px 22px">' +
      '<div style="color:#FFFFFF;font-size:22px;font-weight:bold;letter-spacing:.3px">Konecta Egypt</div>' +
      '<div style="color:#FFE100;font-size:13px;font-weight:bold;margin-top:4px">Welcome onboard</div>' +
    '</div>' +
    '<div style="background:#FFFFFF;border-radius:0 0 12px 12px;padding:28px">' +
      '<p style="font-size:15px;color:#1a1a2e;margin:0 0 14px">Hello ' + escapeHtml_(name) + ',</p>' +
      '<p style="font-size:14px;color:#3a3a50;line-height:1.6;margin:0 0 20px">' +
        'Your Konecta account has been created and is now active. Welcome to the team.</p>' +
      '<div style="background:#EEEDFE;border-radius:10px;padding:18px 20px;margin:0 0 22px">' +
        '<div style="font-size:11px;color:#6b6b80;text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px">Your Konecta email</div>' +
        '<div style="font-size:17px;color:#2800C8;font-weight:bold;word-break:break-all">' + escapeHtml_(konectaEmail) + '</div>' +
      '</div>' +
      '<p style="font-size:14px;color:#3a3a50;line-height:1.6;margin:0 0 22px">' +
        'Please sign in with your new Konecta email, then open your profile below to complete the rest of your details — ' +
        'address, emergency contact, dependants, and bank information.</p>' +
      '<div style="text-align:center;margin:0 0 24px">' +
        '<a href="' + url + '" style="display:inline-block;background:#2800C8;color:#FFFFFF;text-decoration:none;' +
        'font-size:15px;font-weight:bold;padding:14px 34px;border-radius:8px">Complete my details</a>' +
      '</div>' +
      '<div style="background:#FFF9D6;border-left:4px solid #FFE100;border-radius:6px;padding:12px 16px;margin:0 0 8px">' +
        '<div style="font-size:13px;color:#7a6a00;line-height:1.5">' +
          'You must be signed in as <strong>' + escapeHtml_(konectaEmail) + '</strong> to open the link.</div>' +
      '</div>' +
    '</div>' +
    '<div style="text-align:center;color:#9a9ab0;font-size:11px;padding:16px 0 0">' +
      'Konecta Egypt &middot; People Team</div>' +
  '</div>' +
'</div>';
}
function escapeHtml_(s){ return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

function notifyHR_(s,b){ try{ MailApp.sendEmail(HR_ADMINS.join(','),'[Konecta HR] '+s,b);}catch(e){console.error(e);} }
function notifyIT_(s,b){ try{ MailApp.sendEmail(IT_USERS.join(','),'[Konecta IT] '+s,b);}catch(e){console.error(e);} }


// ================================================================
// ATTRITION — the ONE place leavers are visible. Reads full history.
// ================================================================
function hrAttritionReport(year){
  if(!isHR_()) throw new Error('HR only.');
  const E=empData_(true), h=E.hdr;                 // full history, no visibility filter
  const col=function(f){return h.indexOf(f);};
  const cEid=col('employee_id'), cNm=col('full_name_en'), cSt=col('record_status'),
        cHire=col('hire_date'), cExit=col('exit_date'), cType=col('exit_type'),
        cFn=col('function'), cProj=col('project'), cMgr=col('direct_manager'),
        cJt=col('job_title');
  const y=year? parseInt(year):null;
  const leavers=[]; let headcount=0;
  E.rows.forEach(function(rec){
    const v=rec.values;
    const st=String(v[cSt]).trim();
    const exit=fmt_(v[cExit]);
    if(VISIBLE_STATUSES.indexOf(st)!==-1 && !exit){ headcount++; return; }
    if(!exit) return;
    if(y && new Date(exit).getFullYear()!==y) return;
    // tenure in months
    let tenure='';
    const hire=fmt_(v[cHire]);
    if(hire){
      const hd=new Date(hire), xd=new Date(exit);
      if(!isNaN(hd)&&!isNaN(xd)) tenure=Math.max(Math.round((xd-hd)/2629800000),0);
    }
    leavers.push({employee_id:fmt_(v[cEid]), name:fmt_(v[cNm]),
      hire_date:hire, exit_date:exit, exit_type:fmt_(v[cType])||'(not recorded)',
      tenure_months:tenure, job_title:fmt_(v[cJt]), func:fmt_(v[cFn]),
      project:fmt_(v[cProj]), manager:fmt_(v[cMgr])});
  });
  leavers.sort(function(a,b){ return (b.exit_date||'').localeCompare(a.exit_date||''); });

  // summaries
  const byType={}, byMonth={}, byProject={};
  leavers.forEach(function(l){
    byType[l.exit_type]=(byType[l.exit_type]||0)+1;
    const m=(l.exit_date||'').slice(0,7);
    if(m) byMonth[m]=(byMonth[m]||0)+1;
    const p=l.project||'(none)';
    byProject[p]=(byProject[p]||0)+1;
  });
  const avgTenure = leavers.filter(function(l){return l.tenure_months!=='';})
                           .reduce(function(s,l,i,a){return s+l.tenure_months/a.length;},0);
  return {year:y, headcount:headcount, leavers:leavers, total:leavers.length,
          byType:byType, byMonth:byMonth, byProject:byProject,
          avgTenureMonths: Math.round(avgTenure*10)/10,
          rate: headcount? Math.round(leavers.length/(headcount+leavers.length)*1000)/10 : 0};
}


// ================================================================
// LEAVE MODULE — Stage 1: submit a request, count days, show balances
// Tabs required: LEAVE, HOLIDAYS, LEAVE_TYPES  (headers on row 1, data row 2)
// ================================================================

const TAB_LEAVE   = 'LEAVE';
const TAB_HOLIDAY = 'HOLIDAYS';
const TAB_LTYPES  = 'LEAVE_TYPES';

// Annual entitlement: 15 in the first year, 21 from 1 Jan after the first year.


function getHolidays_(){
  const sh=sheet_(TAB_HOLIDAY); if(!sh) return {};
  const data=sh.getRange(2,1,Math.max(sh.getLastRow()-1,1),1).getValues();
  const out={};
  data.forEach(function(r){
    if(!r[0]) return;
    const d = (r[0] instanceof Date) ? Utilities.formatDate(r[0],Session.getScriptTimeZone(),'yyyy-MM-dd') : String(r[0]).trim();
    out[d]=true;
  });
  return out;
}

// Working days between two dates, honouring the employee's weekend pattern.
//   'Fri & Sat'  -> skip Fri, Sat
//   'Sat & Sun'  -> skip Sat, Sun
//   'Rotational' -> no fixed weekend; the employee ticks the days they were scheduled,
//                   passed in as workedDays (array of yyyy-MM-dd). Public holidays always excluded.
const WEEKEND_PATTERNS = {
  'Fri & Sat': [5,6],     // JS getDay(): 0 Sun .. 5 Fri, 6 Sat
  'Sat & Sun': [6,0],
  'Rotational': null
};

function countLeaveDays(startStr, endStr, pattern, workedDays){
  const hol=getHolidays_();
  const off=WEEKEND_PATTERNS[pattern];
  const worked = workedDays ? workedDays.reduce(function(a,d){a[d]=true;return a;},{}) : null;
  const s=new Date(startStr), e=new Date(endStr);
  if(isNaN(s)||isNaN(e)||e<s) return {days:0, error:'Check the dates — the end date must not be before the start date.'};
  let n=0; const cur=new Date(s); const skipped=[]; const all=[];
  while(cur<=e){
    const key=Utilities.formatDate(cur,Session.getScriptTimeZone(),'yyyy-MM-dd');
    const dow=cur.getDay();
    let reason=null;
    if(hol[key]) reason='public holiday';
    else if(off && off.indexOf(dow)!==-1) reason='weekend';
    else if(!off && worked && !worked[key]) reason='not scheduled';
    all.push({date:key, dow:dow, counted:!reason, reason:reason});
    if(reason) skipped.push(key+' ('+reason+')'); else n++;
    cur.setDate(cur.getDate()+1);
  }
  return {days:n, skipped:skipped, detail:all, rotational:(off===null)};
}

// The day list a rotational employee ticks. Public holidays are pre-excluded.
function getDayPicker(startStr, endStr){
  const hol=getHolidays_();
  const s=new Date(startStr), e=new Date(endStr);
  if(isNaN(s)||isNaN(e)||e<s) return [];
  const out=[]; const cur=new Date(s);
  const names=['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  while(cur<=e){
    const key=Utilities.formatDate(cur,Session.getScriptTimeZone(),'yyyy-MM-dd');
    out.push({date:key, label:names[cur.getDay()]+' '+Utilities.formatDate(cur,Session.getScriptTimeZone(),'d MMM'),
              holiday:!!hol[key]});
    cur.setDate(cur.getDate()+1);
  }
  return out;
}

function getLeaveTypes(){
  const sh=sheet_(TAB_LTYPES); if(!sh) return [];
  const data=sh.getRange(2,1,Math.max(sh.getLastRow()-1,1),5).getValues();
  return data.filter(function(r){return r[0];}).map(function(r){
    return {type:String(r[0]).trim(), track:String(r[1]).trim(), entitlement:String(r[2]).trim(),
            document:String(r[3]).trim(), notice:parseInt(r[4])||0};
  });
}

// Everything the Leave tab needs when it opens: my record, entitlement, taken, remaining, my requests.
function getMyLeaveInfo(){
  const me=getMyRecord();
  if(!me.found) return {found:false};
  const sh=sheet_(TAB.EMP), hdr=headers_(TAB.EMP);
  const get=function(f){ const c=hdr.indexOf(f); return c===-1?'':fmt_(sh.getRange(me.row,c+1).getValue()); };
  const eid=get('employee_id'), hire=get('hire_date');

  // full entitlement engine: HR override > disability > age 50+ > tenure
  const recForEnt={
    leave_entitlement:get('leave_entitlement'),
    has_disability:get('has_disability'),
    date_of_birth:get('date_of_birth'),
    hire_date:hire,
    exit_date:get('exit_date')
  };
  const ent=entitlementFor_(recForEnt, new Date().getFullYear());
  const entitlement=ent.days;
  const year=new Date().getFullYear();
  const adj=adjustmentsFor_(eid, year);

  // my requests this year
  const ls=sheet_(TAB_LEAVE);
  const mine=[]; let annualTaken=0;
  if(ls && ls.getLastRow()>1){
    const lh=ls.getRange(1,1,1,ls.getLastColumn()).getValues()[0];
    const li=function(f){return lh.indexOf(f);};
    const rows=ls.getRange(2,1,ls.getLastRow()-1,ls.getLastColumn()).getValues();
    rows.forEach(function(r){
      if(String(r[li('employee_id')]).trim()!==eid) return;
      const start=fmt_(r[li('start_date')]);
      if(start && new Date(start).getFullYear()!==year) return;
      const rec={};
      lh.forEach(function(h,c){ rec[h]=fmt_(r[c]); });
      mine.push(rec);
      const fs=String(rec.final_status||'');
      if(rec.leave_type==='Annual' && (fs.indexOf('Approved')===0 || fs==='Auto-approved')){
        annualTaken += parseFloat(rec.days_approved||rec.days_requested||0)||0;
      }
    });
  }
  mine.reverse();
  return {found:true, employee_id:eid, name:get('full_name_en'), hire_date:hire,
          weekend_pattern:get('weekend_pattern')||'Fri & Sat',
          entitlement_basis:ent.basis,
          adjustments:adj.total, adjustment_items:adj.items,
          konecta_email:get('konecta_email'),
          direct_manager:get('direct_manager'), dotted_manager:get('dotted_manager'),
          annual_entitlement:entitlement,
          annual_taken:annualTaken,
          annual_remaining:Math.max(entitlement + adj.total - annualTaken, 0),
          prorated:ent.prorated, annual_rate:ent.rate,
          types:getLeaveTypes(), requests:mine};
}

// Submit a request. Stage 1 = it lands as Pending; approvals come in stage 2.
function submitLeaveRequest(p){
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const info=getMyLeaveInfo();
    if(!info.found) throw new Error('No employee record linked to your account.');

    const types=getLeaveTypes();
    const t=types.filter(function(x){return x.type===p.leave_type;})[0];
    if(!t) return {ok:false,msg:'Please choose a leave type.'};

    const c=countLeaveDays(p.start_date,p.end_date,info.weekend_pattern,p.worked_days);
    if(c.error) return {ok:false,msg:c.error};
    if(c.days<=0) return {ok:false,msg:'That range contains no working days — it falls on weekends or public holidays.'};

    // how far back may this be dated?
    const BACKDATE_LIMIT_DAYS = 30;
    const start=new Date(p.start_date);
    const today0=new Date(); today0.setHours(0,0,0,0);
    if(start < today0){
      const daysBack=Math.round((today0-start)/86400000);
      if(t.notice>0){
        return {ok:false,msg:t.type+' leave cannot be requested for dates in the past. If this is a correction, please contact HR.'};
      }
      if(daysBack > BACKDATE_LIMIT_DAYS){
        return {ok:false,msg:'This starts '+daysBack+' days ago. Requests can only be backdated up to '+BACKDATE_LIMIT_DAYS+' days — please contact HR to record anything older.'};
      }
    }

    // notice period (annual = 5 business days)
    if(t.notice>0){
      const today=new Date(); today.setHours(0,0,0,0);
      const notice=countLeaveDays(Utilities.formatDate(today,Session.getScriptTimeZone(),'yyyy-MM-dd'),
                                  Utilities.formatDate(new Date(start.getTime()-86400000),Session.getScriptTimeZone(),'yyyy-MM-dd'),
                                  info.weekend_pattern);
      if(notice.days < t.notice){
        return {ok:false,msg:t.type+' leave needs at least '+t.notice+' working days notice. You have given '+notice.days+'.'};
      }
    }

    // annual balance check
    if(p.leave_type==='Annual' && c.days > info.annual_remaining){
      return {ok:false,msg:'You have '+info.annual_remaining+' day(s) of annual leave left, but requested '+c.days+'.'};
    }

    const sh=sheet_(TAB_LEAVE), hdr=headers_(TAB_LEAVE);
    const row=sh.getLastRow()+1;
    const id='LV-'+String(row-1).padStart(6,'0');
    const set=function(f,v){ const i=hdr.indexOf(f); if(i!==-1) sh.getRange(row,i+1).setValue(v); };

    const dm=String(p.direct_manager||info.direct_manager||'').trim();
    const dt=String(p.dotted_manager||info.dotted_manager||'').trim();
    const corrected = (dm!==String(info.direct_manager||'').trim()) || (dt!==String(info.dotted_manager||'').trim());

    set('request_id',id); set('submitted_at',new Date());
    set('employee_id',info.employee_id); set('employee_name',info.name);
    set('konecta_email',info.konecta_email);
    set('leave_type',t.type); set('track',t.track);
    set('start_date',p.start_date); set('end_date',p.end_date);
    set('days_requested',c.days); set('reason',String(p.reason||'').trim());
    set('direct_manager',dm); set('dotted_manager',dt);
    if(corrected){ set('direct_manager_stated',dm); set('dotted_manager_stated',dt); set('manager_correction','Yes'); }
    if(t.track==='Discretionary'){ set('direct_status','Pending'); if(dt) set('dotted_status','Pending'); }
    else { set('hr_status','Pending'); set('document_received','No'); }
    set('final_status','Pending');
    set('reminder_count',0);

    // Entitled leave needs a document. Send the employee a reply-to-this email so the
    // attachment comes back with the request ID in the subject, and HR can match it.
    if(t.track==='Entitled' && t.document){
      const pe=info.konecta_email;
      if(pe){
        // the direct manager is copied so they know about the absence,
        // and both HR mailboxes so either can validate the document
        const mgrEmail=emailForApprover_(resolveApprover_(info.direct_manager));
        const ccList=[mgrEmail].concat(HR_ADMINS).filter(String).join(',');
        try{
          MailApp.sendEmail({
            to: pe,
            cc: ccList,
            subject: 'Action needed: attach your document — '+id,
            htmlBody:
              '<div style="font-family:Arial,sans-serif;max-width:520px">'+
              '<p>Hello '+escapeHtml_(info.name)+',</p>'+
              '<p>Your <strong>'+escapeHtml_(t.type)+'</strong> request <strong>'+id+'</strong> '+
              '('+p.start_date+' to '+p.end_date+', '+c.days+' day(s)) has been received.</p>'+
              '<div style="background:#FFF9D6;border-left:4px solid #FFE100;padding:12px 16px;margin:16px 0">'+
              '<strong>One thing left to do.</strong><br>Reply to this email with your '+
              escapeHtml_(t.document.toLowerCase())+' attached. Keep the subject line as it is — '+
              'it carries your request number so HR can match the document to your request.</div>'+
              '<p style="font-size:13px;color:#6b6b80">Your leave is not confirmed until HR has seen the document. '+
              'Your manager and the People team are copied on this so they know you are away.</p>'+
              '<p style="font-size:13px;color:#6b6b80">Konecta Egypt — People team</p></div>',
            replyTo: HR_ADMINS.join(','),
            name: 'Konecta Egypt — People Team'
          });
        }catch(e){}
      }
    }

    const who = (t.track==='Entitled') ? 'HR (document validation)' : 'your manager';
    notifyHR_('Leave request '+id+' — '+t.type,
      info.name+' ('+info.employee_id+') requested '+c.days+' day(s) of '+t.type+
      '\n'+p.start_date+' to '+p.end_date+
      '\nReason: '+(p.reason||'-')+
      (t.document? ('\n\nDOCUMENT REQUIRED: '+t.document+' — the employee must email it to HR.') : '')+
      (corrected? '\n\nNOTE: the employee corrected their reporting line. Confirm and update the record.' : ''));

    return {ok:true, id:id, days:c.days, skipped:c.skipped,
            msg:'Request '+id+' submitted for '+c.days+' working day(s). It is now with '+who+'.'+
                (t.document? ' Please email your '+t.document.toLowerCase()+' to HR.' : '')};
  } finally { lock.releaseLock(); }
}




// ================================================================
// LEAVE ENTITLEMENT & BALANCE
//   1. HR override on the record   -> use it
//   2. has_disability = Yes        -> 45
//   3. Age 50+                     -> 30
//   4. otherwise                   -> 15 first year, 21 from 1 Jan after
//   + adjustments (holiday worked, corrections)  - approved leave  = balance
// ================================================================

const TAB_ADJUST = 'LEAVE_ADJUSTMENTS';

function ageOn_(dobStr, ref){
  if(!dobStr) return null;
  const d=new Date(dobStr); if(isNaN(d)) return null;
  const t=ref||new Date();
  let a=t.getFullYear()-d.getFullYear();
  if(t.getMonth()<d.getMonth() || (t.getMonth()===d.getMonth() && t.getDate()<d.getDate())) a--;
  return a;
}

// ================================================================
// LEAVE ENTITLEMENT
//   Category comes from leave_entitlement on the record where HR has set it;
//   otherwise it is derived. The JOINING year is 15 days for everyone,
//   whatever their category, prorated from the hire date. From the second
//   year onward they get their full category rate with no proration. The
//   exit year is prorated to the last working day.
//
//   Adjustments — a public holiday worked, a correction — are whole days
//   added on top of the prorated figure. They are NOT prorated, and they are
//   NOT part of the entitlement: an employee should see "21 days, plus 2 for
//   holidays worked", not "23 days".
// ================================================================

const LEAVE_FIRST_YEAR_DAYS = 15;

// The annual rate before any proration.
function leaveCategoryFor_(rec){
  const set=String(rec.leave_entitlement||'').trim();
  if(set && !isNaN(parseFloat(set))) return {days:parseFloat(set), basis:'Set by HR'};
  if(String(rec.has_disability||'').toLowerCase()==='yes')
    return {days:45, basis:'Entitlement for employees with a disability'};
  const age=ageOn_(rec.date_of_birth);
  if(age!==null && age>=50) return {days:30, basis:'Age 50 or over'};
  return {days:21, basis:'Standard entitlement'};
}

// Days in a year, so proration handles leap years without a special case.
function daysInYear_(y){ return ((y%4===0 && y%100!==0)||y%400===0)? 366 : 365; }

// What they are entitled to for one calendar year, prorated at each end.
function entitlementFor_(rec, year){
  year = year || new Date().getFullYear();
  const cat=leaveCategoryFor_(rec);

  const hire = rec.hire_date? new Date(rec.hire_date) : null;
  const exit = rec.exit_date? new Date(rec.exit_date) : null;
  const yStart=new Date(year,0,1), yEnd=new Date(year,11,31);

  // no hire date: fall back to the flat rate rather than guessing
  if(!hire || isNaN(hire)) return {days:cat.days, basis:cat.basis, prorated:false, rate:cat.days};

  if(hire>yEnd) return {days:0, basis:'Not yet employed in '+year, prorated:false, rate:cat.days};
  if(exit && !isNaN(exit) && exit<yStart)
    return {days:0, basis:'Left before '+year, prorated:false, rate:cat.days};

  const isJoiningYear = hire.getFullYear()===year;
  // everyone starts on 15 in their first year, whatever category they are in
  const rate = isJoiningYear? LEAVE_FIRST_YEAR_DAYS : cat.days;

  const from = hire>yStart? hire : yStart;
  const to   = (exit && !isNaN(exit) && exit<yEnd)? exit : yEnd;
  const spanDays = Math.round((to-from)/86400000)+1;
  const fullYear = daysInYear_(year);
  const partial  = spanDays < fullYear;

  if(!partial) return {days:rate, basis: isJoiningYear? 'First year of service':cat.basis,
                       prorated:false, rate:rate};

  // half days are real — round to the nearest half rather than swallowing them
  const days = Math.round((rate*spanDays/fullYear)*2)/2;
  const why = isJoiningYear
    ? 'First year of service — '+rate+' days, prorated from '+fmt_(rec.hire_date)
    : cat.basis+' — '+rate+' days, prorated to '+fmt_(rec.exit_date);
  return {days:days, basis:why, prorated:true, rate:rate,
          span_days:spanDays, year_days:fullYear};
}

// Adjustments: +1 for working a public holiday, corrections, opening balances.
function adjustmentsFor_(eid, year){
  const sh=sheet_(TAB_ADJUST);
  if(!sh || sh.getLastRow()<2) return {total:0, items:[]};
  const hdr=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const i=function(f){return hdr.indexOf(f);};
  const data=sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
  let total=0; const items=[];
  data.forEach(function(r){
    if(String(r[i('employee_id')]).trim()!==eid) return;
    const d=fmt_(r[i('adjustment_date')]);
    if(d && new Date(d).getFullYear()!==year) return;
    const days=parseFloat(r[i('days')])||0;
    total+=days;
    items.push({date:d, days:days, reason:fmt_(r[i('reason')]), by:fmt_(r[i('added_by')])});
  });
  return {total:total, items:items};
}

// HR: add an adjustment (e.g. worked a public holiday -> +1 day)
function hrAddLeaveAdjustment(eid, days, reason, dateStr){
  if(!isHR_()) throw new Error('HR only.');
  const n=parseFloat(days);
  if(isNaN(n) || n===0) return {ok:false,msg:'Enter the number of days to add or subtract.'};
  if(!String(reason||'').trim()) return {ok:false,msg:'A reason is required — this is an audited change.'};
  const sh=sheet_(TAB_ADJUST);
  if(!sh) throw new Error('Tab '+TAB_ADJUST+' not found.');
  const hdr=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const row=sh.getLastRow()+1;
  const set=function(f,v){ const c=hdr.indexOf(f); if(c!==-1) sh.getRange(row,c+1).setValue(v); };
  set('adjustment_id','ADJ-'+String(row-1).padStart(5,'0'));
  set('employee_id',eid);
  set('adjustment_date', dateStr || Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd'));
  set('days',n);
  set('reason',String(reason).trim());
  set('added_by',currentUser_());
  set('added_at',new Date());
  logChange_(eid,'','leave_adjustment','',n+' day(s): '+reason,'HR console','Applied','Leave balance adjustment');
  return {ok:true, msg:(n>0?'+':'')+n+' day(s) recorded.'};
}

// Employee: flag a balance they believe is wrong. Opens a ticket, changes nothing.
function reportBalanceIssue(comment){
  const me=getMyLeaveInfo();
  if(!me.found) throw new Error('No record found.');
  if(!String(comment||'').trim()) return {ok:false,msg:'Please describe what looks wrong.'};
  logChange_(me.employee_id,'','leave_balance','(reported as wrong)',String(comment).trim(),
             'Web app','Pending approval','EMPLOYEE REPORT: leave balance queried');
  notifyHR_('Leave balance queried — '+me.employee_id,
    me.name+' ('+me.employee_id+') believes their leave balance is wrong.\n\n'+
    'System shows: entitlement '+me.annual_entitlement+', taken '+me.annual_taken+', remaining '+me.annual_remaining+
    '\nBasis: '+me.entitlement_basis+
    '\n\nTheir comment: '+comment);
  return {ok:true,msg:'Sent to HR. Your balance has not been changed — HR will review and come back to you.'};
}

// HR: the full leave picture for one employee (for the editor panel)
function hrGetLeaveBalance(row, key){
  if(!isHR_()) throw new Error('HR only.');
  const sh=sheet_(TAB.EMP), hdr=headers_(TAB.EMP);
  row=guardRow_(sh,hdr,row,{employee_id:key});
  const get=function(f){ const c=hdr.indexOf(f); return c===-1?'':fmt_(sh.getRange(row,c+1).getValue()); };
  const eid=get('employee_id');
  const rec={leave_entitlement:get('leave_entitlement'), has_disability:get('has_disability'),
             date_of_birth:get('date_of_birth'), hire_date:get('hire_date'),
             exit_date:get('exit_date')};
  const ent=entitlementFor_(rec, new Date().getFullYear());
  const year=new Date().getFullYear();
  const adj=adjustmentsFor_(eid,year);

  // approved annual leave this year
  let taken=0;
  const ls=sheet_(TAB_LEAVE);
  if(ls && ls.getLastRow()>1){
    const lh=ls.getRange(1,1,1,ls.getLastColumn()).getValues()[0];
    const li=function(f){return lh.indexOf(f);};
    ls.getRange(2,1,ls.getLastRow()-1,ls.getLastColumn()).getValues().forEach(function(r){
      if(String(r[li('employee_id')]).trim()!==eid) return;
      if(String(r[li('leave_type')])!=='Annual') return;
      const st=String(r[li('final_status')]||'');
      if(st.indexOf('Approved')!==0 && st!=='Auto-approved') return;
      const d=fmt_(r[li('start_date')]);
      if(d && new Date(d).getFullYear()!==year) return;
      taken += parseFloat(r[li('days_approved')]||r[li('days_requested')]||0)||0;
    });
  }
  return {employee_id:eid, override:get('leave_entitlement'),
          entitlement:ent.days, basis:ent.basis,
          adjustments:adj.total, adjustment_items:adj.items,
          taken:taken, remaining:Math.max(ent.days+adj.total-taken,0)};
}
// ================================================================
// LEAVE — STAGE 2: APPROVALS
//   Discretionary : direct manager -> dotted manager (each may REDUCE days)
//   Entitled      : HR validates the document, no manager involvement
//   Delegation    : a manager may nominate someone to act on their behalf
// ================================================================

const TAB_DELEGATES = 'DELEGATES';

// Who may this person act for? Returns the manager IDs they can approve as.
function actingFor_(){
  const me=currentUser_();
  const ids=[];
  const identity=getManagerIdentity_();
  if(identity) ids.push(identity.id);          // themselves
  const sh=sheet_(TAB_DELEGATES);
  if(sh && sh.getLastRow()>1){
    const hdr=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    const i=function(f){return hdr.indexOf(f);};
    sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues().forEach(function(r){
      if(String(r[i('delegate_email')]).toLowerCase().trim()!==me) return;
      const act=String(r[i('active')]||'').toLowerCase();
      if(act && act!=='yes') return;
      // optional window: blank dates mean open-ended
      const today=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');
      const from=fmt_(r[i('from_date')]), to=fmt_(r[i('to_date')]);
      if(from && today<from) return;
      if(to && today>to) return;
      const mid=String(r[i('manager_id')]).trim();
      if(mid && ids.indexOf(mid)===-1) ids.push(mid);
    });
  }
  return ids;
}


// When a request's manager is a GLOBAL manager who cannot log in, route the approval
// to whoever the MANAGERS tab says acts for them (approves_via). Most point to Ahmad
// (EG0001); a few point elsewhere, e.g. EG0146.
var _approverCache={};
function resolveApprover_(managerId){
  const id=String(managerId||'').trim();
  if(!id) return '';
  if(_approverCache[id]!==undefined) return _approverCache[id];
  const keep=function(v){ _approverCache[id]=v; return v; };
  if(id.indexOf('GLOBAL-')!==0) return keep(id);
  const sh=sheet_('MANAGERS');
  if(!sh || sh.getLastRow()<5) return keep(id);
  const hdr=sh.getRange(4,1,1,sh.getLastColumn()).getValues()[0];
  const iId=hdr.indexOf('manager_id'), iVia=hdr.indexOf('approves_via'), iView=hdr.indexOf('can_view');
  const rows=sh.getRange(5,1,sh.getLastRow()-4,sh.getLastColumn()).getValues();
  for(var i=0;i<rows.length;i++){
    if(String(rows[i][iId]).trim()!==id) continue;
    const canView=String(rows[i][iView]||'').toLowerCase().trim()==='yes';
    if(canView) return keep(id);
    const via = iVia===-1? '' : String(rows[i][iVia]||'').trim();
    return keep(via || id);
  }
  return keep(id);
}


// The specific working dates a request covers. If the approver has already trimmed
// days, approved_dates holds what survived; otherwise derive from the range.
function leaveDatesOf_(rec){
  if(rec.approved_dates){
    return String(rec.approved_dates).split(',').map(function(d){return d.trim();}).filter(String);
  }
  const worked = rec.worked_days ? String(rec.worked_days).split(',').map(function(d){return d.trim();}) : null;
  const c=countLeaveDays(rec.start_date, rec.end_date, rec.weekend_pattern||'Fri & Sat', worked);
  if(!c.detail) return [];
  return c.detail.filter(function(d){return d.counted;}).map(function(d){return d.date;});
}

function leaveHdr_(){ const sh=sheet_(TAB_LEAVE); return sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]; }

// Requests waiting on ME (as manager or delegate)
function getLeaveApprovals(){
  const acting=actingFor_();
  if(!acting.length) return {isApprover:false, items:[]};
  const sh=sheet_(TAB_LEAVE);
  if(!sh || sh.getLastRow()<2) return {isApprover:true, acting:acting, items:[]};
  const hdr=leaveHdr_(); const i=function(f){return hdr.indexOf(f);};
  const data=sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
  const out=[];
  data.forEach(function(r,idx){
    if(String(r[i('track')])!=='Discretionary') return;      // entitled leave never reaches managers
    const fs=String(r[i('final_status')]||'');
    if(fs && fs!=='Pending') return;
    const dm=resolveApprover_(r[i('direct_manager')]);
    const dt=resolveApprover_(r[i('dotted_manager')]);
    const ds=String(r[i('direct_status')]||'');
    const ts=String(r[i('dotted_status')]||'');

    let myRole=null;
    if(acting.indexOf(dm)!==-1 && ds==='Pending') myRole='direct';
    else if(acting.indexOf(dt)!==-1 && ds==='Approved' && ts==='Pending') myRole='dotted';
    if(!myRole) return;

    const rec={row:idx+2, role:myRole};
    hdr.forEach(function(h,c){ rec[h]=fmt_(r[c]); });
    // the actual dates this request covers, so the approver ticks specific days
    rec.days_list = leaveDatesOf_(rec);
    // the dotted approver works on what the direct manager allowed
    rec.max_days = myRole==='dotted' ? (parseFloat(rec.direct_days)||parseFloat(rec.days_requested))
                                     : parseFloat(rec.days_requested);
    rec.acting_as = myRole==='direct' ? dm : dt;
    rec.on_behalf = (rec.acting_as !== (getManagerIdentity_()||{}).id);
    out.push(rec);
  });
  return {isApprover:true, acting:acting, items:out};
}

// Manager decision. days = how many they allow (may be fewer, never more).
function decideLeave(row, decision, days, comment, keptDates, key){
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const acting=actingFor_();
    if(!acting.length) throw new Error('You are not an approver.');
    const sh=sheet_(TAB_LEAVE), hdr=leaveHdr_();
    row=guardRow_(sh,hdr,row,{request_id:key});
    const i=function(f){return hdr.indexOf(f);};
    const get=function(f){ const c=i(f); return c===-1?'':fmt_(sh.getRange(row,c+1).getValue()); };
    const set=function(f,v){ const c=i(f); if(c!==-1) sh.getRange(row,c+1).setValue(v); };

    if(String(get('track'))!=='Discretionary') throw new Error('This leave type is validated by HR, not by managers.');
    const dm=resolveApprover_(get('direct_manager')), dt=resolveApprover_(get('dotted_manager'));
    const ds=String(get('direct_status')||''), ts=String(get('dotted_status')||'');

    let role=null;
    if(acting.indexOf(dm)!==-1 && ds==='Pending') role='direct';
    else if(acting.indexOf(dt)!==-1 && ds==='Approved' && ts==='Pending') role='dotted';
    if(!role) throw new Error('This request is not waiting on you.');

    const requested=parseFloat(get('days_requested'))||0;
    // the dates currently on the table: what the previous approver left, or the full range
    const onTable = String(get('approved_dates')||'').split(',').map(function(d){return d.trim();}).filter(String);
    const ceilingDates = onTable.length ? onTable : leaveDatesOf_({
        start_date:get('start_date'), end_date:get('end_date'),
        weekend_pattern:get('weekend_pattern'), worked_days:get('worked_days')});
    let kept = [];
    if(decision!=='reject'){
      kept = (keptDates && keptDates.length) ? keptDates.filter(function(d){ return ceilingDates.indexOf(d)!==-1; })
                                             : ceilingDates.slice();
    }
    let allowed = kept.length;

    const stamp=new Date(), who=currentUser_();
    set('approved_dates', kept.join(','));
    if(role==='direct'){
      set('direct_status', allowed>0 ? 'Approved':'Rejected');
      set('direct_days', allowed); set('direct_by', who); set('direct_at', stamp);
    } else {
      set('dotted_status', allowed>0 ? 'Approved':'Rejected');
      set('dotted_days', allowed); set('dotted_by', who); set('dotted_at', stamp);
    }
    if(comment) set('notes', (get('notes')? get('notes')+' | ':'') + role+': '+comment);

    // work out the final position
    const needDotted = !!dt;
    let final='Pending', approvedDays=null;
    if(allowed===0){ final='Rejected'; approvedDays=0; }
    else if(role==='direct' && !needDotted){ final = allowed<requested?'Partially approved':'Approved'; approvedDays=allowed; }
    else if(role==='dotted'){ final = allowed<requested?'Partially approved':'Approved'; approvedDays=allowed; }

    if(approvedDays!==null){
      set('days_approved', approvedDays);
      set('days_rejected', requested-approvedDays);   // LIABILITY under labour law
      set('final_status', final);
      notifyLeaveOutcome_(row, final, approvedDays, requested);
      if(approvedDays>0){
        try{ createLeaveCalendarBlock_(row, hdr); }catch(e){}
        try{ flagUnpaidForPayroll_(row, hdr); }catch(e){}
      }
    }
    return {ok:true, role:role, allowed:allowed, final:final};
  } finally { lock.releaseLock(); }
}

// HR validates entitled leave once the document is in hand
function hrValidateLeave(row, decision, days, comment, key){
  if(!isHR_()) throw new Error('HR only.');
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const sh=sheet_(TAB_LEAVE), hdr=leaveHdr_();
    row=guardRow_(sh,hdr,row,{request_id:key});
    const i=function(f){return hdr.indexOf(f);};
    const get=function(f){ const c=i(f); return c===-1?'':fmt_(sh.getRange(row,c+1).getValue()); };
    const set=function(f,v){ const c=i(f); if(c!==-1) sh.getRange(row,c+1).setValue(v); };
    if(String(get('track'))!=='Entitled') throw new Error('This request goes to managers, not HR.');

    const requested=parseFloat(get('days_requested'))||0;
    const allowed = decision==='reject' ? 0 : Math.min(parseFloat(days)||requested, requested);
    set('hr_status', allowed>0?'Validated':'Rejected');
    set('document_received', decision==='reject' ? 'No':'Yes');
    set('hr_by', currentUser_()); set('hr_at', new Date());
    if(comment) set('notes',(get('notes')? get('notes')+' | ':'')+'HR: '+comment);
    set('days_approved', allowed);
    set('days_rejected', requested-allowed);
    const final = allowed===0?'Rejected':(allowed<requested?'Partially approved':'Approved');
    set('final_status', final);
    notifyLeaveOutcome_(row, final, allowed, requested);
    if(allowed>0){ try{ createLeaveCalendarBlock_(row, hdr); }catch(e){} }
    return {ok:true, final:final, allowed:allowed};
  } finally { lock.releaseLock(); }
}

// HR queue for entitled leave awaiting document validation
function hrGetLeaveToValidate(){
  if(!isHR_()) throw new Error('HR only.');
  const sh=sheet_(TAB_LEAVE);
  if(!sh || sh.getLastRow()<2) return [];
  const hdr=leaveHdr_(); const i=function(f){return hdr.indexOf(f);};
  const out=[];
  sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues().forEach(function(r,idx){
    if(String(r[i('track')])!=='Entitled') return;
    const fs=String(r[i('final_status')]||'');
    if(fs && fs!=='Pending') return;
    const rec={row:idx+2}; hdr.forEach(function(h,c){ rec[h]=fmt_(r[c]); });
    // what document does this type need?
    const t=getLeaveTypes().filter(function(x){return x.type===rec.leave_type;})[0];
    rec.document_needed = t? t.document : '';
    out.push(rec);
  });
  return out;
}

function notifyLeaveOutcome_(row, final, approved, requested){
  const sh=sheet_(TAB_LEAVE), hdr=leaveHdr_();
  const i=function(f){return hdr.indexOf(f);};
  const g=function(f){ const c=i(f); return c===-1?'':fmt_(sh.getRange(row,c+1).getValue()); };
  const to=g('konecta_email'); if(!to) return;

  // the managers who were told about the absence are told the outcome,
  // and HR is copied so the trail sits in the mailbox as well as the sheet
  const dmEmail = emailForApprover_(resolveApprover_(g('direct_manager')));
  const dtEmail = emailForApprover_(resolveApprover_(g('dotted_manager')));
  const ccList = [dmEmail, dtEmail].concat(HR_ADMINS)
                   .filter(String)
                   .filter(function(e,idx,arr){ return e!==to && arr.indexOf(e)===idx; })
                   .join(',');

  let body='Your '+g('leave_type')+' request '+g('request_id')+' ('+g('start_date')+' to '+g('end_date')+') has been '+final.toLowerCase()+'.';
  if(final==='Partially approved'){
    body+='\n\nApproved: '+approved+' of '+requested+' day(s).';
    if(g('approved_dates')) body+='\n\nDays approved: '+g('approved_dates').split(',').join(', ');
    body+='\n\nThe '+(requested-approved)+' day(s) not approved remain available to you — you may request them again later.';
  } else if(final==='Rejected'){
    body+='\n\nNo days were approved. Your balance is unchanged and you may request again.';
  } else {
    body+='\n\nApproved: '+approved+' day(s).';
    if(g('approved_dates')) body+='\nDays: '+g('approved_dates').split(',').join(', ');
  }
  if(g('notes')) body+='\n\nNotes: '+g('notes');

  try{
    MailApp.sendEmail({
      to: to,
      cc: ccList,
      subject: 'Leave '+final.toLowerCase()+' — '+g('request_id'),
      body: body+'\n\nKonecta Egypt — People team'
    });
  }catch(e){}
}


// ---------- delegation, managed by the manager themselves ----------

// Who this manager could delegate to: their own team first, then everyone else.
function getDelegateOptions(){
  const identity=getManagerIdentity_();
  if(!identity) return {canDelegate:false};
  // Only return this manager's OWN team. The old version shipped every active
  // employee's name AND Konecta email to any manager — the whole company
  // directory — as the "others" list. A delegate outside the team is chosen by
  // server-side typeahead (delegateSearch) instead, so no bulk email list ever
  // crosses to the client.
  const E=empData_(false), h=E.hdr;
  const ei=h.indexOf('employee_id'), ni=h.indexOf('full_name_en'),
        ke=h.indexOf('konecta_email'), dm=h.indexOf('direct_manager');
  const team=[];
  E.rows.forEach(function(rec){
    const v=rec.values;
    if(String(v[dm]).trim()!==identity.id) return;
    const email=String(v[ke]).trim(); if(!email) return;   // must be able to log in
    team.push({id:String(v[ei]).trim(), email:email,
               label:String(v[ni]).trim()+' ('+String(v[ei]).trim()+')'});
  });
  team.sort(function(a,b){return a.label.localeCompare(b.label);});
  return {canDelegate:true, managerId:identity.id, team:team, others:[], typeahead:true,
          current:myDelegations_()};
}

// Server-side typeahead for delegating outside your team: needs 2+ characters,
// returns at most 10 matches, and only to a manager. Replaces shipping every
// Konecta email to the client.
function delegateSearch(term){
  const identity=getManagerIdentity_();
  if(!identity) throw new Error('Only a manager can delegate.');
  const q=String(term||'').trim().toLowerCase();
  if(q.length<2) return [];
  const E=empData_(false), h=E.hdr;
  const ei=h.indexOf('employee_id'), ni=h.indexOf('full_name_en'), ke=h.indexOf('konecta_email');
  const out=[];
  for(let i=0;i<E.rows.length && out.length<10;i++){
    const v=E.rows[i].values;
    const email=String(v[ke]).trim(); if(!email) continue;
    const name=String(v[ni]).trim();
    if((name+' '+String(v[ei])+' '+email).toLowerCase().indexOf(q)===-1) continue;
    out.push({id:String(v[ei]).trim(), email:email, label:name+' ('+String(v[ei]).trim()+')'});
  }
  return out;
}

function myDelegations_(){
  const identity=getManagerIdentity_();
  if(!identity) return [];
  const sh=sheet_(TAB_DELEGATES);
  if(!sh || sh.getLastRow()<2) return [];
  const hdr=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const i=function(f){return hdr.indexOf(f);};
  const out=[];
  sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues().forEach(function(r,idx){
    if(String(r[i('manager_id')]).trim()!==identity.id) return;
    out.push({row:idx+2, delegate_email:fmt_(r[i('delegate_email')]),
              active:fmt_(r[i('active')]), from_date:fmt_(r[i('from_date')]),
              to_date:fmt_(r[i('to_date')]), note:fmt_(r[i('note')])});
  });
  return out;
}

// Manager sets up a delegation. Dates optional — blank means until they turn it off.
function setDelegate(delegateEmail, fromDate, toDate, note){
  const identity=getManagerIdentity_();
  if(!identity) throw new Error('Only a manager can delegate approvals.');
  const email=String(delegateEmail||'').trim().toLowerCase();
  if(!/^[^@]+@konecta\.com$/.test(email)) return {ok:false,msg:'Choose a colleague with a Konecta email.'};
  const sh=sheet_(TAB_DELEGATES);
  if(!sh) throw new Error('Tab '+TAB_DELEGATES+' not found.');
  const hdr=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const row=sh.getLastRow()+1;
  const set=function(f,v){ const c=hdr.indexOf(f); if(c!==-1) sh.getRange(row,c+1).setValue(v); };
  set('manager_id',identity.id); set('delegate_email',email); set('active','Yes');
  set('from_date',fromDate||''); set('to_date',toDate||'');
  set('note',note||''); 
  const c=hdr.indexOf('created_by'); if(c!==-1) sh.getRange(row,c+1).setValue(currentUser_());
  const c2=hdr.indexOf('created_at'); if(c2!==-1) sh.getRange(row,c2+1).setValue(new Date());

  const window = (fromDate||toDate) ? (' from '+(fromDate||'now')+' to '+(toDate||'further notice')) : ' until you turn it off';
  try{
    MailApp.sendEmail(email,'You can now approve leave on behalf of '+identity.id,
      'You have been given delegated approval rights'+window+'.\n\n'+
      'Leave requests waiting on '+identity.id+' will appear in your My Team tab, marked as acting on their behalf.\n\n'+
      'Konecta Egypt — People team');
  }catch(e){}
  return {ok:true, msg:'Delegation set. '+email+' has been notified.'};
}

function endDelegate(row, key){
  const identity=getManagerIdentity_();
  if(!identity) throw new Error('Not authorised.');
  const sh=sheet_(TAB_DELEGATES);
  const hdr=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  row=guardRow_(sh,hdr,row,{manager_id:identity.id, delegate_email:key, active:'Yes'});
  const mi=hdr.indexOf('manager_id');
  if(String(sh.getRange(row,mi+1).getValue()).trim()!==identity.id) throw new Error('That is not your delegation.');
  const ai=hdr.indexOf('active');
  sh.getRange(row,ai+1).setValue('No');
  return {ok:true};
}
// ================================================================
// LEAVE — STAGE 3: reminders, auto-approval, calendar blocking
//
// Auto-approval applies to ANNUAL LEAVE ONLY. Silence becomes a yes there,
// so the reminders matter: day 2 to the approver, day 4 to the approver AND HR.
// Sick and every entitled type are excluded — HR must see the document.
// Unpaid and casual are excluded too — those are deliberate decisions.
//
// Set ONE time-driven trigger: leaveDailyRun, daily, early morning.
// ================================================================

const AUTO_APPROVE_TYPES = ['Annual'];
const AUTO_APPROVE_AFTER_DAYS = 5;      // working days with no decision
const REMIND_ON_DAYS = [2, 4];          // day 4 copies HR

function leaveDailyRun(){
  assertNotDirectCall_();
  schemaDailyCheck_();
  const sh=sheet_(TAB_LEAVE);
  if(!sh || sh.getLastRow()<2) return;
  const hdr=leaveHdr_(); const i=function(f){return hdr.indexOf(f);};
  const rows=sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
  const today=new Date(); today.setHours(0,0,0,0);
  let reminded=0, autoApproved=0;

  rows.forEach(function(r,idx){
    const row=idx+2;
    const fs=String(r[i('final_status')]||'');
    if(fs && fs!=='Pending') return;                       // already decided
    const submitted=r[i('submitted_at')];
    if(!submitted) return;
    const sub=new Date(submitted); sub.setHours(0,0,0,0);

    // working days elapsed since submission, using the employee's own pattern
    const pattern=String(r[i('weekend_pattern')]||'Fri & Sat');
    const elapsed=countLeaveDays(
      Utilities.formatDate(sub,Session.getScriptTimeZone(),'yyyy-MM-dd'),
      Utilities.formatDate(today,Session.getScriptTimeZone(),'yyyy-MM-dd'),
      pattern, null).days - 1;                             // day of submission is day 0
    if(elapsed<1) return;

    const track=String(r[i('track')]);
    const type=String(r[i('leave_type')]);
    const eligibleForAuto = (track==='Discretionary' && AUTO_APPROVE_TYPES.indexOf(type)!==-1);

    // --- auto-approve ---
    if(eligibleForAuto && elapsed>=AUTO_APPROVE_AFTER_DAYS){
      autoApproveLeave_(row, hdr, r, elapsed);
      autoApproved++;
      return;
    }

    // --- reminders ---
    if(track!=='Discretionary') return;   // entitled leave sits with HR, not the manager
    const already=parseInt(r[i('reminder_count')])||0;
    REMIND_ON_DAYS.forEach(function(d, n){
      if(elapsed===d && already<=n){
        sendLeaveReminder_(row, hdr, r, elapsed, d===Math.max.apply(null,REMIND_ON_DAYS), eligibleForAuto);
        sh.getRange(row, i('reminder_count')+1).setValue(n+1);
        sh.getRange(row, i('last_reminder_at')+1).setValue(new Date());
        reminded++;
      }
    });
  });
  return {reminded:reminded, autoApproved:autoApproved};
}

function autoApproveLeave_(row, hdr, r, elapsed){
  const sh=sheet_(TAB_LEAVE); const i=function(f){return hdr.indexOf(f);};
  const requested=parseFloat(r[i('days_requested')])||0;
  // auto-approval grants what is still on the table
  const onTable=String(r[i('approved_dates')]||'').split(',').map(function(d){return d.trim();}).filter(String);
  const dates = onTable.length ? onTable : leaveDatesOf_({
      start_date:fmt_(r[i('start_date')]), end_date:fmt_(r[i('end_date')]),
      weekend_pattern:fmt_(r[i('weekend_pattern')]), worked_days:fmt_(r[i('worked_days')])});
  const days=dates.length || requested;

  sh.getRange(row,i('approved_dates')+1).setValue(dates.join(','));
  sh.getRange(row,i('days_approved')+1).setValue(days);
  sh.getRange(row,i('days_rejected')+1).setValue(Math.max(requested-days,0));
  sh.getRange(row,i('final_status')+1).setValue('Auto-approved');
  sh.getRange(row,i('auto_approved')+1).setValue('Yes');
  sh.getRange(row,i('notes')+1).setValue(
    (fmt_(r[i('notes')])? fmt_(r[i('notes')])+' | ':'')+
    'Auto-approved after '+elapsed+' working days with no decision.');

  const emp=fmt_(r[i('konecta_email')]);
  const id=fmt_(r[i('request_id')]);
  if(emp){
    try{ MailApp.sendEmail(emp,'Leave approved — '+id,
      'Your '+fmt_(r[i('leave_type')])+' request '+id+' ('+fmt_(r[i('start_date')])+' to '+fmt_(r[i('end_date')])+') '+
      'has been approved.\n\nDays: '+dates.join(', ')+
      '\n\nNo decision was recorded within '+AUTO_APPROVE_AFTER_DAYS+' working days, so it was approved automatically under company policy.'+
      '\n\nKonecta Egypt — People team'); }catch(e){}
  }
  notifyHR_('Auto-approved leave — '+id,
    fmt_(r[i('employee_name')])+' ('+fmt_(r[i('employee_id')])+') — '+days+' day(s) of '+fmt_(r[i('leave_type')])+
    '\n'+fmt_(r[i('start_date')])+' to '+fmt_(r[i('end_date')])+
    '\n\nNo manager acted within '+AUTO_APPROVE_AFTER_DAYS+' working days. Approved automatically.'+
    '\nDirect manager: '+fmt_(r[i('direct_manager')])+'   Dotted: '+fmt_(r[i('dotted_manager')])+
    '\n\nThis is flagged so you can follow up with the manager.');
  createLeaveCalendarBlock_(row, hdr);
}

function sendLeaveReminder_(row, hdr, r, elapsed, isFinal, willAutoApprove){
  const i=function(f){return hdr.indexOf(f);};
  // who is it waiting on?
  const ds=String(r[i('direct_status')]||'');
  const waitingOn = (ds==='Approved') ? resolveApprover_(fmt_(r[i('dotted_manager')]))
                                      : resolveApprover_(fmt_(r[i('direct_manager')]));
  const to=emailForApprover_(waitingOn);
  const id=fmt_(r[i('request_id')]);
  const left=Math.max(AUTO_APPROVE_AFTER_DAYS-elapsed,0);
  let body=fmt_(r[i('employee_name')])+' ('+fmt_(r[i('employee_id')])+') is waiting on your decision.\n\n'+
    fmt_(r[i('leave_type')])+' — '+fmt_(r[i('days_requested')])+' day(s), '+
    fmt_(r[i('start_date')])+' to '+fmt_(r[i('end_date')])+'\nRequest '+id+
    '\n\nSubmitted '+elapsed+' working day(s) ago.';
  if(willAutoApprove){
    body+='\n\nIf no decision is recorded within '+left+' more working day(s), this request will be '+
          'APPROVED AUTOMATICALLY under company policy.';
  }
  body+='\n\nOpen the app and go to My Team to approve or decline.\n\nKonecta Egypt — People team';

  const subject=(isFinal?'FINAL REMINDER':'Reminder')+' — leave awaiting your approval ('+id+')';
  try{
    if(to) MailApp.sendEmail(to, subject, body);
    if(isFinal) notifyHR_(subject, 'No decision after '+elapsed+' working days.\n\n'+body);
  }catch(e){}
}

// find the login email for an approver id (employee or global manager)
function emailForApprover_(id){
  if(!id) return '';
  const sh=sheet_(TAB.EMP), hdr=headers_(TAB.EMP), data=sh.getDataRange().getValues();
  const ei=hdr.indexOf('employee_id'), ke=hdr.indexOf('konecta_email');
  for(let r=1;r<data.length;r++){
    if(String(data[r][ei]).trim()===id) return String(data[r][ke]).trim();
  }
  const ms=sheet_('MANAGERS');
  if(ms && ms.getLastRow()>4){
    const mh=ms.getRange(4,1,1,ms.getLastColumn()).getValues()[0];
    const mi=mh.indexOf('manager_id'), me=mh.indexOf('email');
    const rows=ms.getRange(5,1,ms.getLastRow()-4,ms.getLastColumn()).getValues();
    for(let i=0;i<rows.length;i++){
      if(String(rows[i][mi]).trim()===id) return String(rows[i][me]||'').trim();
    }
  }
  return '';
}

// ---------- calendar OOO block on approval ----------
function createLeaveCalendarBlock_(row, hdr){
  const sh=sheet_(TAB_LEAVE); const i=function(f){return hdr.indexOf(f);};
  const g=function(f){ const c=i(f); return c===-1?'':fmt_(sh.getRange(row,c+1).getValue()); };
  if(g('calendar_event_id')) return;                  // already blocked
  const email=g('konecta_email'); if(!email) return;
  const dates=String(g('approved_dates')||'').split(',').map(function(d){return d.trim();}).filter(String);
  if(!dates.length) return;
  try{
    const cal=CalendarApp.getCalendarById(email);
    if(!cal) return;
    const ids=[];
    dates.forEach(function(d){
      const day=new Date(d);
      const ev=cal.createAllDayEvent('Out of office — '+g('leave_type'), day);
      ev.setDescription('Approved leave '+g('request_id')+'. Recorded in the Konecta HR system.');
      try{ ev.removeAllReminders(); }catch(e){}
      ids.push(ev.getId());
    });
    sh.getRange(row,i('calendar_event_id')+1).setValue(ids.join(','));
  }catch(e){
    console.error('Calendar block failed for '+g('request_id')+': '+e);
  }
}

// ---------- payroll flag: unpaid leave must reach payroll ----------
function flagUnpaidForPayroll_(row, hdr){
  const sh=sheet_(TAB_LEAVE); const i=function(f){return hdr.indexOf(f);};
  const g=function(f){ const c=i(f); return c===-1?'':fmt_(sh.getRange(row,c+1).getValue()); };
  if(g('leave_type')!=='Unpaid') return;
  const days=parseFloat(g('days_approved'))||0;
  if(days<=0) return;
  sh.getRange(row,i('payroll_flag')+1).setValue('Deduct '+days+' day(s)');
  notifyHR_('Unpaid leave approved — payroll deduction due',
    g('employee_name')+' ('+g('employee_id')+') — '+days+' unpaid day(s)\n'+
    g('start_date')+' to '+g('end_date')+'  ('+g('request_id')+')\n\n'+
    'Deduct these days in the next payroll run.');
}
// ================================================================
// HR LEAVE PANEL — balances, bulk holiday-worked, backlog entry,
// and the year-end liability report.
// ================================================================

// ---------- bulk: who worked a public holiday ----------
// HR pastes employee IDs (one per line, or comma separated) and picks the holiday.
// Each person gets +1 day, recorded with the reason and the date.
function hrBulkHolidayWorked(idText, holidayDate, note){
  if(!isHR_()) throw new Error('HR only.');
  const ids=String(idText||'').split(/[\s,;]+/).map(function(x){return x.trim().toUpperCase();}).filter(String);
  if(!ids.length) return {ok:false,msg:'Paste at least one employee ID.'};
  if(!holidayDate) return {ok:false,msg:'Choose which public holiday they worked.'};

  // check the date really is a holiday — guards against a mistyped date
  const hol=getHolidays_();
  if(!hol[holidayDate]) return {ok:false,msg:holidayDate+' is not in the HOLIDAYS tab. Add it first, or check the date.'};

  const valid=validEmployeeIds_();
  const done=[], unknown=[];
  ids.forEach(function(id){
    if(!valid[id]){ unknown.push(id); return; }
    hrAddLeaveAdjustment(id, 1, 'Worked the public holiday on '+holidayDate+(note?(' — '+note):''), holidayDate);
    done.push(id);
  });
  return {ok:true, added:done.length, unknown:unknown,
          msg:done.length+' employee(s) credited with 1 day.'+
              (unknown.length? ' '+unknown.length+' ID(s) not recognised: '+unknown.join(', ') : '')};
}

function validEmployeeIds_(){
  const sh=sheet_(TAB.EMP), hdr=headers_(TAB.EMP), data=sh.getDataRange().getValues();
  const ei=hdr.indexOf('employee_id');
  const out={};
  for(let r=1;r<data.length;r++){
    const id=String(data[r][ei]).trim().toUpperCase();
    if(id) out[id]=true;
  }
  return out;
}

// ---------- bulk: historic leave backlog (March onward) ----------
// One line per record:  EG0123, 2026-03-10, 2026-03-12, Annual
// Loaded straight as approved history — no approval flow, because it already happened.
function hrBulkBacklog(text){
  if(!isHR_()) throw new Error('HR only.');
  const lines=String(text||'').split('\n').map(function(l){return l.trim();}).filter(String);
  if(!lines.length) return {ok:false,msg:'Paste at least one line.'};
  const valid=validEmployeeIds_();
  const sh=sheet_(TAB_LEAVE), hdr=leaveHdr_();
  const types=getLeaveTypes();
  const ok=[], bad=[];

  lines.forEach(function(line, n){
    const p=line.split(',').map(function(x){return x.trim();});
    if(p.length<4){ bad.push('line '+(n+1)+': expected 4 values'); return; }
    const id=p[0].toUpperCase(), start=p[1], end=p[2], type=p[3];
    if(!valid[id]){ bad.push('line '+(n+1)+': unknown employee '+id); return; }
    const t=types.filter(function(x){return x.type.toLowerCase()===type.toLowerCase();})[0];
    if(!t){ bad.push('line '+(n+1)+': unknown leave type "'+type+'"'); return; }

    const emp=employeeSnapshot_(id);
    const c=countLeaveDays(start,end,emp.weekend_pattern||'Fri & Sat',null);
    if(c.error || c.days<=0){ bad.push('line '+(n+1)+': '+(c.error||'no working days in that range')); return; }

    const row=sh.getLastRow()+1;
    const set=function(f,v){ const i=hdr.indexOf(f); if(i!==-1) sh.getRange(row,i+1).setValue(v); };
    const dates=(c.detail||[]).filter(function(d){return d.counted;}).map(function(d){return d.date;});
    set('request_id','LV-'+String(row-1).padStart(6,'0'));
    set('submitted_at',new Date()); set('employee_id',id);
    set('employee_name',emp.name); set('konecta_email',emp.email);
    set('leave_type',t.type); set('track',t.track);
    set('start_date',start); set('end_date',end);
    set('days_requested',c.days); set('days_approved',c.days); set('days_rejected',0);
    set('approved_dates',dates.join(','));
    set('final_status','Approved');
    set('notes','Historic record loaded by HR — this leave was taken before the system went live.');
    set('hr_by',currentUser_()); set('hr_at',new Date());
    ok.push(id+' '+start+'→'+end+' ('+c.days+'d)');
  });
  return {ok:true, loaded:ok.length, errors:bad,
          msg:ok.length+' record(s) loaded.'+(bad.length? ' '+bad.length+' line(s) skipped.' : '')};
}

function employeeSnapshot_(id){
  const sh=sheet_(TAB.EMP), hdr=headers_(TAB.EMP), data=sh.getDataRange().getValues();
  const ei=hdr.indexOf('employee_id');
  for(let r=1;r<data.length;r++){
    if(String(data[r][ei]).trim().toUpperCase()!==id) continue;
    const g=function(f){ const c=hdr.indexOf(f); return c===-1?'':fmt_(data[r][c]); };
    return {name:g('full_name_en'), email:g('konecta_email'),
            weekend_pattern:g('weekend_pattern'), hire_date:g('hire_date')};
  }
  return {};
}

// ---------- HR: set an entitlement override ----------
function hrSetEntitlement(row, days, reason, key){
  if(!isHR_()) throw new Error('HR only.');
  const sh=sheet_(TAB.EMP), hdr=headers_(TAB.EMP);
  row=guardRow_(sh,hdr,row,{employee_id:key});
  const c=hdr.indexOf('leave_entitlement');
  if(c===-1) throw new Error('Column leave_entitlement not found.');
  const old=fmt_(sh.getRange(row,c+1).getValue());
  const val=String(days||'').trim();
  sh.getRange(row,c+1).setValue(val);
  const eid=fmt_(sh.getRange(row,hdr.indexOf('employee_id')+1).getValue());
  logChange_(eid,'','leave_entitlement',old,val,'HR console','Applied',reason||'Entitlement set by HR');
  return {ok:true, msg: val? ('Entitlement set to '+val+' days.') : 'Override cleared — the automatic rules now apply.'};
}

// ---------- year-end: rejected days carried as a liability ----------
// Under Egyptian labour law, leave the employee REQUESTED and the company REFUSED must be
// paid the following year if unconsumed. Leave never requested is a different question.
function hrRejectedDaysReport(year){
  if(!isHR_()) throw new Error('HR only.');
  const y=parseInt(year)||new Date().getFullYear();
  const sh=sheet_(TAB_LEAVE);
  if(!sh || sh.getLastRow()<2) return {year:y, rows:[], totalDays:0};
  const hdr=leaveHdr_(); const i=function(f){return hdr.indexOf(f);};
  const byEmp={};
  sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues().forEach(function(r){
    if(String(r[i('leave_type')])!=='Annual') return;         // only annual is encashable
    const start=fmt_(r[i('start_date')]);
    if(!start || new Date(start).getFullYear()!==y) return;
    const rej=parseFloat(r[i('days_rejected')])||0;
    if(rej<=0) return;
    const id=fmt_(r[i('employee_id')]);
    if(!byEmp[id]) byEmp[id]={employee_id:id, name:fmt_(r[i('employee_name')]), days:0, requests:[]};
    byEmp[id].days+=rej;
    byEmp[id].requests.push({id:fmt_(r[i('request_id')]), dates:start+' to '+fmt_(r[i('end_date')]),
      rejected:rej, by:fmt_(r[i('dotted_by')])||fmt_(r[i('direct_by')]),
      when:fmt_(r[i('dotted_at')])||fmt_(r[i('direct_at')])});
  });
  const rows=Object.keys(byEmp).map(function(k){return byEmp[k];})
                   .sort(function(a,b){return b.days-a.days;});
  return {year:y, rows:rows, totalDays:rows.reduce(function(s,r){return s+r.days;},0)};
}

// Holiday list for the HR dropdown
function getHolidayList(){
  const sh=sheet_(TAB_HOLIDAY);
  if(!sh || sh.getLastRow()<2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,2).getValues()
    .filter(function(r){return r[0];})
    .map(function(r){
      const d=(r[0] instanceof Date)? Utilities.formatDate(r[0],Session.getScriptTimeZone(),'yyyy-MM-dd') : String(r[0]).trim();
      return {date:d, name:String(r[1]||'').trim()};
    });
}
// ================================================================
// RESIGNATION
//   The employee proposes their last working day. Both managers must land on
//   the SAME date — neither can overrule the other. If they have not agreed
//   within 10 days, the EMPLOYEE'S date stands. That is deliberate: the cost
//   of inaction falls on the managers, not the person leaving.
//
//   Withdrawal is immediate within 10 days of submitting. After that the
//   direct manager must accept it, and the dotted manager is told.
// ================================================================

const TAB_RESIGN = 'RESIGNATIONS';
const RESIGN_AUTO_DAYS = 10;          // calendar days before the employee's date stands
const RESIGN_REMIND_EVERY = 2;        // reminder cadence, in days
const WITHDRAW_FREE_DAYS = 10;        // withdraw without approval inside this window

function resignHdr_(){ const sh=sheet_(TAB_RESIGN); return sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]; }

function getMyResignation(){
  const me=getMyRecord();
  if(!me.found) return {found:false};
  const sh=sheet_(TAB_RESIGN);
  const eid=me.readonly.employee_id;
  const base={found:true, employee_id:eid, name:me.editable.full_name_en||'',
              notice_period:me.readonly.notice_period||'',
              direct_manager:me.readonly.direct_manager||'', existing:null};
  if(!sh || sh.getLastRow()<2) return base;
  const hdr=resignHdr_(), i=function(f){return hdr.indexOf(f);};
  const rows=sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
  for(let r=rows.length-1;r>=0;r--){                    // most recent first
    if(String(rows[r][i('employee_id')]).trim()!==eid) continue;
    const st=String(rows[r][i('final_status')]||'');
    if(st==='Withdrawn') continue;                       // a withdrawn one does not block a new one
    const rec={row:r+2};
    hdr.forEach(function(h,c){ rec[h]=fmt_(rows[r][c]); });
    const sub=new Date(rec.submitted_at);
    rec.days_since = isNaN(sub)? 0 : Math.floor((new Date()-sub)/86400000);
    rec.can_withdraw_freely = rec.days_since <= WITHDRAW_FREE_DAYS;
    base.existing=rec;
    break;
  }
  return base;
}

function submitResignation(p){
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const info=getMyResignation();
    if(!info.found) throw new Error('No employee record linked to your account.');
    if(info.existing && ['Pending','Accepted','Auto-approved'].indexOf(info.existing.final_status)!==-1){
      return {ok:false,msg:'You already have a resignation in progress ('+info.existing.resignation_id+'). Withdraw it first if you need to change anything.'};
    }
    const last=String(p.proposed_last_day||'').trim();
    if(!last) return {ok:false,msg:'Please choose your proposed last working day.'};
    const d=new Date(last), today=new Date(); today.setHours(0,0,0,0);
    if(isNaN(d)) return {ok:false,msg:'That date is not valid.'};
    if(d<today) return {ok:false,msg:'Your last working day cannot be in the past.'};

    const emp=getMyRecord();
    const sh=sheet_(TAB_RESIGN), hdr=resignHdr_();
    const row=sh.getLastRow()+1;
    const id='RS-'+String(row-1).padStart(5,'0');
    const set=function(f,v){ const c=hdr.indexOf(f); if(c!==-1) sh.getRange(row,c+1).setValue(v); };

    const dm=resolveApprover_(emp.readonly.direct_manager);
    const dt=resolveApprover_(emp.readonly.dotted_manager);

    set('resignation_id',id); set('submitted_at',new Date());
    set('employee_id',info.employee_id); set('employee_name',info.name);
    set('konecta_email',emp.email);
    set('proposed_last_day',last); set('reason',String(p.reason||'').trim());
    set('direct_manager',dm); set('direct_status','Pending');
    if(dt){ set('dotted_manager',dt); set('dotted_status','Pending'); }
    set('final_status','Pending'); set('reminder_count',0);

    notifyResignApprovers_(row, hdr, 'submitted');
    notifyHR_('Resignation submitted — '+id,
      info.name+' ('+info.employee_id+') has resigned.\n'+
      'Proposed last working day: '+last+'\n'+
      (p.reason? ('Reason: '+p.reason+'\n') : '')+
      '\nWith '+dm+(dt?(' and '+dt):'')+' to agree the date. '+
      'If they have not agreed within '+RESIGN_AUTO_DAYS+' days, the employee\'s date stands.');

    return {ok:true, id:id,
      msg:'Your resignation has been submitted ('+id+'). Your managers will confirm the last working day. '+
          'You can withdraw it yourself within '+WITHDRAW_FREE_DAYS+' days; after that your manager must agree.'};
  } finally { lock.releaseLock(); }
}

// A manager either accepts the date on the table, or proposes a different one.
function decideResignation(row, decision, proposedDate, comment, key){
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const acting=actingFor_();
    if(!acting.length) throw new Error('You are not an approver.');
    const sh=sheet_(TAB_RESIGN), hdr=resignHdr_();
    row=guardRow_(sh,hdr,row,{resignation_id:key});
    const i=function(f){return hdr.indexOf(f);};
    const get=function(f){ const c=i(f); return c===-1?'':fmt_(sh.getRange(row,c+1).getValue()); };
    const set=function(f,v){ const c=i(f); if(c!==-1) sh.getRange(row,c+1).setValue(v); };

    if(get('final_status')!=='Pending') throw new Error('This resignation is no longer open.');
    const dm=get('direct_manager'), dt=get('dotted_manager');
    let role=null;
    if(acting.indexOf(dm)!==-1) role='direct';
    else if(dt && acting.indexOf(dt)!==-1) role='dotted';
    if(!role) throw new Error('This resignation is not waiting on you.');

    const myDate = decision==='accept' ? get('proposed_last_day') : String(proposedDate||'').trim();
    if(!myDate) return {ok:false,msg:'Choose a date, or accept the one proposed.'};
    if(isNaN(new Date(myDate))) return {ok:false,msg:'That date is not valid.'};

    set(role+'_status', decision==='accept' ? 'Accepted' : 'Proposed different date');
    set(role+'_date', myDate);
    set(role+'_by', currentUser_());
    set(role+'_at', new Date());
    if(comment) set('notes',(get('notes')? get('notes')+' | ':'')+role+': '+comment);

    // both landed on the same date?
    const a=get('direct_date'), b=dt? get('dotted_date') : a;
    const bothIn = a && (!dt || b);
    if(bothIn && a===b){
      set('agreed_last_day', a);
      set('final_status','Accepted');
      finaliseResignation_(row, hdr, a, false);
      return {ok:true, agreed:true, date:a};
    }
    if(bothIn && a!==b){
      // deliberate: no one overrules the other. They must converge.
      notifyResignApprovers_(row, hdr, 'disagreement');
      return {ok:true, agreed:false,
              msg:'Recorded. You and the other manager have proposed different dates ('+a+' and '+b+'). '+
                  'You need to agree one between you — neither date can override the other. '+
                  'If no agreement is reached within '+RESIGN_AUTO_DAYS+' days of submission, the employee\'s date will stand.'};
    }
    return {ok:true, agreed:false, msg:'Recorded. Waiting on the other manager.'};
  } finally { lock.releaseLock(); }
}

function finaliseResignation_(row, hdr, lastDay, auto){
  const sh=sheet_(TAB_RESIGN); const i=function(f){return hdr.indexOf(f);};
  const g=function(f){ const c=i(f); return c===-1?'':fmt_(sh.getRange(row,c+1).getValue()); };
  const set=function(f,v){ const c=i(f); if(c!==-1) sh.getRange(row,c+1).setValue(v); };
  set('agreed_last_day', lastDay);
  if(auto){ set('final_status','Auto-approved'); set('auto_approved','Yes'); }

  // settlement lands the month AFTER departure
  const d=new Date(lastDay);
  const settle=new Date(d.getFullYear(), d.getMonth()+1, 1);
  set('settlement_month', Utilities.formatDate(settle,Session.getScriptTimeZone(),'yyyy-MM'));

  // move the employee record to Serving Notice and record the exit
  const eid=g('employee_id');
  const esh=sheet_(TAB.EMP), ehdr=headers_(TAB.EMP), data=esh.getDataRange().getValues();
  const ei=ehdr.indexOf('employee_id');
  for(let r=1;r<data.length;r++){
    if(String(data[r][ei]).trim()!==eid) continue;
    const put=function(f,v){ const c=ehdr.indexOf(f); if(c!==-1) esh.getRange(r+1,c+1).setValue(v); };
    put('record_status','Serving Notice');
    put('last_working_day', lastDay);
    put('exit_date', lastDay);
    put('exit_type','Resignation');
    // Monthly salary continues through the notice period — they are still working.
    // It is the FINAL settlement that waits for clearance.
    put('payment_status','Release');
    put('hold_reason','Final settlement held until clearance is complete');
    break;
  }
  clearEmpCache_();

  const to=g('konecta_email');
  if(to){
    try{ MailApp.sendEmail(to,'Your resignation — last working day confirmed ('+g('resignation_id')+')',
      'Your last working day is confirmed as '+lastDay+'.'+
      (auto? '\n\nYour managers did not agree a different date within '+RESIGN_AUTO_DAYS+' days, so the date you proposed stands.' : '')+
      '\n\nYour remaining leave balance will be settled in the month after you leave.'+
      '\n\nHR will be in touch about clearance and handover.'+
      '\n\nKonecta Egypt — People team'); }catch(e){}
  }
  // open the clearance record — handover starts with the managers
  try{ openClearance_(eid, lastDay); }catch(e){ console.error('clearance open failed: '+e); }

  notifyHR_('Resignation confirmed — '+g('resignation_id'),
    g('employee_name')+' ('+eid+') — last working day '+lastDay+
    (auto? '\n\nAUTO-APPROVED: the managers did not agree within '+RESIGN_AUTO_DAYS+' days, so the employee\'s proposed date stands.' : '')+
    '\n\nRecord moved to Serving Notice. Leave settlement due '+g('settlement_month')+'.');
}

// ---------- withdrawal ----------
function withdrawResignation(comment){
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const info=getMyResignation();
    if(!info.found || !info.existing) return {ok:false,msg:'You have no resignation to withdraw.'};
    const rec=info.existing;
    if(['Withdrawn','Rejected'].indexOf(rec.final_status)!==-1) return {ok:false,msg:'This resignation is already closed.'};

    const sh=sheet_(TAB_RESIGN), hdr=resignHdr_();
    const set=function(f,v){ const c=hdr.indexOf(f); if(c!==-1) sh.getRange(rec.row,c+1).setValue(v); };
    set('withdraw_reason', String(comment||'').trim());
    set('withdrawn_at', new Date());

    if(rec.can_withdraw_freely){
      set('final_status','Withdrawn'); set('withdraw_status','Withdrawn');
      set('withdraw_by', currentUser_());
      restoreEmployeeFromNotice_(rec.employee_id);
      notifyResignApprovers_(rec.row, hdr, 'withdrawn');
      notifyHR_('Resignation withdrawn — '+rec.resignation_id,
        rec.employee_name+' ('+rec.employee_id+') has withdrawn their resignation within the '+
        WITHDRAW_FREE_DAYS+'-day window. No approval was needed.');
      return {ok:true, msg:'Your resignation has been withdrawn. Nothing further is needed.'};
    }
    // outside the window: the direct manager must agree
    set('withdraw_status','Pending manager');
    notifyResignApprovers_(rec.row, hdr, 'withdraw_requested');
    return {ok:true, msg:'Your withdrawal request has gone to your manager. It is more than '+
      WITHDRAW_FREE_DAYS+' days since you resigned, so they need to agree to it. You will be told either way.'};
  } finally { lock.releaseLock(); }
}

function decideWithdrawal(row, accept, comment, key){
  const acting=actingFor_();
  if(!acting.length) throw new Error('Not authorised.');
  const sh=sheet_(TAB_RESIGN), hdr=resignHdr_();
  row=guardRow_(sh,hdr,row,{resignation_id:key});
  const i=function(f){return hdr.indexOf(f);};
  const g=function(f){ const c=i(f); return c===-1?'':fmt_(sh.getRange(row,c+1).getValue()); };
  const set=function(f,v){ const c=i(f); if(c!==-1) sh.getRange(row,c+1).setValue(v); };
  if(acting.indexOf(g('direct_manager'))===-1) throw new Error('Only the direct manager can decide a withdrawal.');
  if(g('withdraw_status')!=='Pending manager') throw new Error('There is no withdrawal request open.');

  set('withdraw_by', currentUser_());
  if(comment) set('notes',(g('notes')? g('notes')+' | ':'')+'withdrawal: '+comment);
  const to=g('konecta_email');
  if(accept){
    set('withdraw_status','Approved'); set('final_status','Withdrawn');
    restoreEmployeeFromNotice_(g('employee_id'));
    if(to){ try{ MailApp.sendEmail(to,'Your resignation has been withdrawn ('+g('resignation_id')+')',
      'Your manager has agreed to withdraw your resignation. Your employment continues as normal.'+
      '\n\nKonecta Egypt — People team'); }catch(e){} }
    notifyHR_('Withdrawal approved — '+g('resignation_id'),
      g('employee_name')+' ('+g('employee_id')+') is staying. Record returned to Active.');
  } else {
    set('withdraw_status','Declined');
    if(to){ try{ MailApp.sendEmail(to,'Withdrawal not agreed ('+g('resignation_id')+')',
      'Your manager has not agreed to withdraw your resignation. Your last working day remains '+
      (g('agreed_last_day')||g('proposed_last_day'))+'.'+
      '\n\nPlease speak to HR if you need to discuss this.'+
      '\n\nKonecta Egypt — People team'); }catch(e){} }
  }
  return {ok:true};
}

function restoreEmployeeFromNotice_(eid){
  const esh=sheet_(TAB.EMP), ehdr=headers_(TAB.EMP), data=esh.getDataRange().getValues();
  const ei=ehdr.indexOf('employee_id');
  for(let r=1;r<data.length;r++){
    if(String(data[r][ei]).trim()!==eid) continue;
    const put=function(f,v){ const c=ehdr.indexOf(f); if(c!==-1) esh.getRange(r+1,c+1).setValue(v); };
    put('record_status','Active'); put('last_working_day',''); put('exit_date',''); put('exit_type','');
    break;
  }
  clearEmpCache_();
}

// ---------- notifications ----------
function notifyResignApprovers_(row, hdr, kind){
  const sh=sheet_(TAB_RESIGN); const i=function(f){return hdr.indexOf(f);};
  const g=function(f){ const c=i(f); return c===-1?'':fmt_(sh.getRange(row,c+1).getValue()); };
  const id=g('resignation_id'), who=g('employee_name')+' ('+g('employee_id')+')';
  const recips=[emailForApprover_(g('direct_manager')), emailForApprover_(g('dotted_manager'))]
                 .filter(String).join(',');
  if(!recips) return;
  let subject='', body='';
  const proposed=g('proposed_last_day');

  if(kind==='submitted'){
    subject='Resignation to confirm — '+who;
    body=who+' has resigned.\n\nProposed last working day: '+proposed+
      (g('reason')? ('\nReason: '+g('reason')) : '')+
      '\n\nYou and the other manager need to AGREE one last working day between you. '+
      'Neither of you can override the other — if you propose different dates, you must converge.'+
      '\n\nIMPORTANT: if no agreement is recorded within '+RESIGN_AUTO_DAYS+' days, '+
      'the date the employee proposed ('+proposed+') will stand automatically.'+
      '\n\nOpen the app and go to My Team.';
  } else if(kind==='disagreement'){
    subject='Different dates proposed — '+who;
    body='You and the other manager have proposed different last working days for '+who+':'+
      '\n\n  Direct manager: '+(g('direct_date')||'not yet')+
      '\n  Dotted manager: '+(g('dotted_date')||'not yet')+
      '\n\nNeither date overrides the other. Please agree one between you.'+
      '\n\nIf nothing is agreed within '+RESIGN_AUTO_DAYS+' days of the resignation, '+
      'the employee\'s proposed date ('+proposed+') will stand.';
  } else if(kind==='withdrawn'){
    subject='Resignation withdrawn — '+who;
    body=who+' has withdrawn their resignation within the '+WITHDRAW_FREE_DAYS+
      '-day window, so no approval was needed. They are staying.';
  } else if(kind==='withdraw_requested'){
    subject='Withdrawal request — '+who;
    body=who+' wants to withdraw their resignation.'+
      '\n\nIt is more than '+WITHDRAW_FREE_DAYS+' days since they resigned, so this needs your agreement.'+
      (g('withdraw_reason')? ('\n\nTheir reason: '+g('withdraw_reason')) : '')+
      '\n\nOpen the app and go to My Team to accept or decline.';
  }
  try{ MailApp.sendEmail(recips,'[Konecta] '+subject, body+'\n\nKonecta Egypt — People team'); }catch(e){}
}

// ---------- daily run: reminders every 2 days, auto-approve at 10 ----------
// Add ONE time-driven daily trigger for this.
function resignationDailyRun(){
  assertNotDirectCall_();
  const sh=sheet_(TAB_RESIGN);
  if(!sh || sh.getLastRow()<2) return;
  const hdr=resignHdr_(); const i=function(f){return hdr.indexOf(f);};
  const rows=sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
  let reminded=0, auto=0;

  rows.forEach(function(r,idx){
    const row=idx+2;
    if(String(r[i('final_status')])!=='Pending') return;
    const sub=r[i('submitted_at')]; if(!sub) return;
    const days=Math.floor((new Date()-new Date(sub))/86400000);

    if(days>=RESIGN_AUTO_DAYS){
      // the employee's date stands — the managers had their chance
      finaliseResignation_(row, hdr, fmt_(r[i('proposed_last_day')]), true);
      auto++;
      return;
    }
    if(days>0 && days%RESIGN_REMIND_EVERY===0){
      const already=parseInt(r[i('reminder_count')])||0;
      const due=Math.floor(days/RESIGN_REMIND_EVERY);
      if(already<due){
        sendResignReminder_(row, hdr, days);
        sh.getRange(row,i('reminder_count')+1).setValue(due);
        sh.getRange(row,i('last_reminder_at')+1).setValue(new Date());
        reminded++;
      }
    }
  });
  return {reminded:reminded, autoApproved:auto};
}

function sendResignReminder_(row, hdr, daysElapsed){
  const sh=sheet_(TAB_RESIGN); const i=function(f){return hdr.indexOf(f);};
  const g=function(f){ const c=i(f); return c===-1?'':fmt_(sh.getRange(row,c+1).getValue()); };
  const recips=[emailForApprover_(g('direct_manager')), emailForApprover_(g('dotted_manager'))]
                 .filter(String).join(',');
  if(!recips) return;
  const left=RESIGN_AUTO_DAYS-daysElapsed;
  const a=g('direct_date'), b=g('dotted_date');
  let state;
  if(!a && !b) state='Neither of you has responded yet.';
  else if(a && b && a!==b) state='You have proposed different dates ('+a+' and '+b+'). You need to converge on one.';
  else if(a && !b) state='The direct manager proposed '+a+'. The dotted manager has not responded.';
  else state='The dotted manager proposed '+b+'. The direct manager has not responded.';

  const body=g('employee_name')+' ('+g('employee_id')+') resigned '+daysElapsed+' day(s) ago.'+
    '\n\nProposed last working day: '+g('proposed_last_day')+
    '\n\n'+state+
    '\n\nWARNING: if no agreed date is recorded within '+left+' more day(s), this resignation will be '+
    'APPROVED AUTOMATICALLY at the date the employee proposed ('+g('proposed_last_day')+'). '+
    'Not responding does not delay the departure — it simply removes your say in the date.'+
    '\n\nOpen the app and go to My Team.';
  try{ MailApp.sendEmail(recips,'[Konecta] REMINDER: resignation awaiting your agreement — '+g('employee_name'), 
    body+'\n\nKonecta Egypt — People team'); }catch(e){}
}

// ---------- manager queue ----------
function getResignationApprovals(){
  const acting=actingFor_();
  if(!acting.length) return {isApprover:false, items:[]};
  const sh=sheet_(TAB_RESIGN);
  if(!sh || sh.getLastRow()<2) return {isApprover:true, items:[]};
  const hdr=resignHdr_(); const i=function(f){return hdr.indexOf(f);};
  const out=[];
  sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues().forEach(function(r,idx){
    const dm=String(r[i('direct_manager')]).trim(), dt=String(r[i('dotted_manager')]).trim();
    const mine = acting.indexOf(dm)!==-1 || (dt && acting.indexOf(dt)!==-1);
    if(!mine) return;
    const fs=String(r[i('final_status')]||'');
    const ws=String(r[i('withdraw_status')]||'');
    const needsMe = (fs==='Pending') || (ws==='Pending manager' && acting.indexOf(dm)!==-1);
    if(!needsMe) return;
    const rec={row:idx+2, role: acting.indexOf(dm)!==-1?'direct':'dotted',
               isWithdrawal: ws==='Pending manager'};
    hdr.forEach(function(h,c){ rec[h]=fmt_(r[c]); });
    const sub=new Date(rec.submitted_at);
    rec.days_since=isNaN(sub)?0:Math.floor((new Date()-sub)/86400000);
    rec.days_left=Math.max(RESIGN_AUTO_DAYS-rec.days_since,0);
    out.push(rec);
  });
  return {isApprover:true, items:out};
}


function submitResignationFor(p){
  if(!isHR_()) throw new Error('HR only.');
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const eid=String(p.employee_id||'').trim().toUpperCase();
    if(!eid) return {ok:false,msg:'Enter the employee ID.'};
    const last=String(p.proposed_last_day||'').trim();
    if(!last) return {ok:false,msg:'Enter the proposed last working day.'};

    const f=employeeFieldsOf_(eid,['employee_id','full_name_en','konecta_email',
                                   'record_status','direct_manager','dotted_manager']);
    if(!f.employee_id) return {ok:false,msg:'No employee found with ID '+eid+'.'};
    if(['Closed','Cleared'].indexOf(String(f.record_status))!==-1)
      return {ok:false,msg:'That record is already closed.'};

    // already one in progress?
    const sh=sheet_(TAB_RESIGN), hdr=resignHdr_();
    if(sh.getLastRow()>1){
      const i=function(x){return hdr.indexOf(x);};
      const rows=sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
      for(let r=0;r<rows.length;r++){
        if(String(rows[r][i('employee_id')]).trim()!==eid) continue;
        const st=String(rows[r][i('final_status')]||'');
        if(['Pending','Accepted','Auto-approved'].indexOf(st)!==-1){
          return {ok:false,msg:'There is already a resignation in progress for '+eid+'.'};
        }
      }
    }

    const row=sh.getLastRow()+1;
    const id='RS-'+String(row-1).padStart(5,'0');
    const set=function(x,v){ const c=hdr.indexOf(x); if(c!==-1) sh.getRange(row,c+1).setValue(v); };
    const dm=resolveApprover_(f.direct_manager), dt=resolveApprover_(f.dotted_manager);

    set('resignation_id',id); set('submitted_at',new Date());
    set('employee_id',eid); set('employee_name',f.full_name_en);
    set('konecta_email',f.konecta_email);
    set('proposed_last_day',last); set('reason',String(p.reason||'').trim());
    set('direct_manager',dm); set('direct_status','Pending');
    if(dt){ set('dotted_manager',dt); set('dotted_status','Pending'); }
    set('final_status','Pending'); set('reminder_count',0);
    set('notes','Recorded by '+currentUser_()+' on behalf of the employee'+
                (p.on_behalf_note? ' — '+p.on_behalf_note : ''));

    // The employee must be told. A resignation recorded in someone's name without
    // their knowledge is dangerous even when it is genuine.
    if(f.konecta_email){
      try{ MailApp.sendEmail(f.konecta_email,
        'Your resignation has been recorded — '+id,
        'Hello '+f.full_name_en+',\n\n'+
        'HR has recorded your resignation, with a proposed last working day of '+last+'.\n\n'+
        'IF THIS IS WRONG, contact HR straight away.\n\n'+
        'Your managers will now agree the final date between them. You can see the status, '+
        'and withdraw it yourself within 10 days, in the Resignation tab of the app.\n\n'+
        'Konecta Egypt — People team'); }catch(e){}
    }

    notifyResignApprovers_(row, hdr, 'submitted');
    notifyHR_('Resignation recorded on behalf — '+id,
      f.full_name_en+' ('+eid+')\nProposed last working day: '+last+
      '\nRecorded by: '+currentUser_()+
      (p.reason? ('\nReason: '+p.reason) : '')+
      '\n\nThe employee has been emailed to confirm. The managers still agree the date as normal.');

    return {ok:true, id:id,
      msg:'Recorded as '+id+'. '+f.full_name_en+' has been emailed to confirm, and their managers have been asked to agree the date.'};
  } finally { lock.releaseLock(); }
}


// ================================================================
// CLEARANCE
//   1. HANDOVER   direct + dotted manager confirm handover is complete
//                 (blocks everything else)
//   2. IT + FACILITIES in parallel
//   3. HR/PERSONNEL final gate — exit formalities, documents, Comp & Ben
//   -> settlement releases
//
//   Unreturned items and outstanding loans accumulate as deductions
//   against the final payout. The employee is told what they owe.
// ================================================================

const TAB_CLEAR  = 'CLEARANCE';
const FACILITIES_USERS = ['lobna.adel@konecta.com'];

const IT_ITEMS  = [
  {f:'it_laptop',            label:'Laptop'},
  {f:'it_headset',           label:'Headset'},
  {f:'it_charger',           label:'Laptop charger'},
  {f:'it_bag',               label:'Laptop bag'},
  {f:'it_email_deactivated', label:'Email deactivated', noAsset:true}
];
const FAC_ITEMS = [
  {f:'fac_access_card', label:'Access card'},
  {f:'fac_mifi',        label:'MiFi'},
  {f:'fac_mobile_line', label:'Mobile line'},
  {f:'fac_locker',      label:'Locker'}
];
const HR_ITEMS  = [
  {f:'hr_exit_formalities', label:'Exit formalities completed'},
  {f:'hr_documents_handed', label:'Hiring documents handed over'},
  {f:'hr_compben_notified', label:'Comp & Ben notified — medical + social insurance deactivation'}
];

function isFacilities_(){ return inList_(FACILITIES_USERS); }
function clearHdr_(){ const sh=sheet_(TAB_CLEAR); return sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]; }

// Open a clearance record. Called when a resignation is confirmed, or by HR directly.
function openClearance_(employeeId, lastWorkingDay){
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const eid=String(employeeId).trim();
    const sh=sheet_(TAB_CLEAR), hdr=clearHdr_();
    // already open?
    if(sh.getLastRow()>1){
      const rows=sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
      const ci=hdr.indexOf('employee_id'), fi=hdr.indexOf('final_status');
      for(let r=0;r<rows.length;r++){
        if(String(rows[r][ci]).trim()===eid && String(rows[r][fi])!=='Cleared'){
          return {ok:false,msg:'A clearance is already open for '+eid+'.'};
        }
      }
    }
    const emp=employeeSnapshot_(eid);
    if(!emp.name) return {ok:false,msg:'No employee record for '+eid+'.'};

    const row=sh.getLastRow()+1;
    const id='CL-'+String(row-1).padStart(5,'0');
    const set=function(f,v){ const c=hdr.indexOf(f); if(c!==-1) sh.getRange(row,c+1).setValue(v); };
    set('clearance_id',id); set('employee_id',eid); set('employee_name',emp.name);
    set('konecta_email',emp.email); set('last_working_day',lastWorkingDay||'');
    set('opened_at',new Date());
    set('handover_direct_status','Pending');
    set('handover_dotted_status','Pending');
    set('it_status','Blocked'); set('fac_status','Blocked'); set('hr_status','Blocked');
    set('final_status','In progress');

    // tell the managers it starts with them
    const dm=emailForApprover_(resolveApprover_(directManagerOf_(eid)));
    const dt=emailForApprover_(resolveApprover_(dottedManagerOf_(eid)));
    const to=[dm,dt].filter(String).join(',');
    if(to){
      try{ MailApp.sendEmail(to,'[Konecta] Handover confirmation needed — '+emp.name,
        emp.name+' ('+eid+') is leaving on '+(lastWorkingDay||'a date to be confirmed')+'.\n\n'+
        'Clearance starts with you. Please confirm all handover activities are complete — '+
        'IT and Facilities cannot begin collecting equipment until you do.\n\n'+
        'Open the app and go to My Team.\n\nKonecta Egypt — People team'); }catch(e){}
    }
    notifyHR_('Clearance opened — '+id, emp.name+' ('+eid+'), last working day '+(lastWorkingDay||'TBC')+
      '.\n\nWaiting on the managers to confirm handover.');
    return {ok:true, id:id};
  } finally { lock.releaseLock(); }
}

function directManagerOf_(eid){ const s=employeeFieldsOf_(eid,['direct_manager']); return s.direct_manager||''; }
function dottedManagerOf_(eid){ const s=employeeFieldsOf_(eid,['dotted_manager']); return s.dotted_manager||''; }
/**
 * A few fields for one employee.
 *
 * This used to call sh.getDataRange().getValues() — the ENTIRE employee table,
 * every column of every row — to read two or three cells for one person. With
 * 26 call sites, four of them inside loops (hrInviteToSign, hrPendingDependants,
 * hrMovementReport, hrMarkSigning), that was the single most expensive pattern
 * in the file: inviting 100 people to sign meant 100 full-table reads.
 *
 * Now it reads the employee_id column to locate the row, then that one row.
 * Cells transferred go from rows x columns to rows + columns — at 2,000
 * employees and ~80 columns, from ~160,000 to ~2,080 per call.
 *
 * Deliberately still reads live rather than memoising, so a caller that writes
 * to EMPLOYEES and reads back in the same execution cannot see stale data.
 */
function employeeFieldsOf_(eid, fields){
  const want=String(eid||'').trim();
  if(!want) return {};
  const sh=sheet_(TAB.EMP), hdr=headers_(TAB.EMP);
  const ei=hdr.indexOf('employee_id');
  if(ei===-1) return {};
  const lastRow=sh.getLastRow();
  if(lastRow<2) return {};

  // One narrow read of the ID column to find the row.
  const ids=sh.getRange(2,ei+1,lastRow-1,1).getValues();
  let found=-1;
  for(let r=0;r<ids.length;r++){
    if(String(ids[r][0]).trim()===want){ found=r+2; break; }
  }
  const out={};
  if(found===-1) return out;

  // One read of just that employee's row.
  const row=sh.getRange(found,1,1,hdr.length).getValues()[0];
  fields.forEach(function(f){ const c=hdr.indexOf(f); out[f]= c===-1?'':fmt_(row[c]); });
  return out;
}

// ---------- stage 1: handover ----------
function confirmHandover(row, note, key){
  const acting=actingFor_();
  if(!acting.length) throw new Error('Only a manager can confirm handover.');
  const sh=sheet_(TAB_CLEAR), hdr=clearHdr_();
  row=guardRow_(sh,hdr,row,{clearance_id:key});
  const g=function(f){ const c=hdr.indexOf(f); return c===-1?'':fmt_(sh.getRange(row,c+1).getValue()); };
  const set=function(f,v){ const c=hdr.indexOf(f); if(c!==-1) sh.getRange(row,c+1).setValue(v); };
  const eid=g('employee_id');
  const dm=resolveApprover_(directManagerOf_(eid)), dt=resolveApprover_(dottedManagerOf_(eid));

  let role=null;
  if(acting.indexOf(dm)!==-1) role='direct';
  else if(dt && acting.indexOf(dt)!==-1) role='dotted';
  if(!role) throw new Error('This handover is not waiting on you.');

  set('handover_'+role+'_status','Confirmed');
  set('handover_'+role+'_by',currentUser_());
  set('handover_'+role+'_at',new Date());
  if(note) set('handover_'+role+'_note',note);

  const needDotted=!!dt;
  const bothIn = g('handover_direct_status')==='Confirmed' &&
                 (!needDotted || g('handover_dotted_status')==='Confirmed');
  if(bothIn){
    set('handover_complete_at',new Date());
    set('it_status','Pending'); set('fac_status','Pending');
    const to=[FACILITIES_USERS.join(','), IT_USERS.join(',')].filter(String).join(',');
    try{ MailApp.sendEmail(to,'[Konecta] Clearance — please collect equipment for '+g('employee_name'),
      g('employee_name')+' ('+eid+') leaves on '+g('last_working_day')+'.\n\n'+
      'Handover is confirmed, so equipment collection can start. Open the app to record what comes back.\n\n'+
      'Anything not returned is charged against the final payout, so please mark items honestly.\n\n'+
      'Konecta Egypt — People team'); }catch(e){}
    notifyHR_('Handover confirmed — '+g('clearance_id'),
      g('employee_name')+' — IT and Facilities can now collect equipment.');
  }
  return {ok:true, bothConfirmed:bothIn};
}

// ---------- stage 2: IT and Facilities ----------
function submitClearanceItems(row, dept, items, amounts, note, loans, key){
  const sh=sheet_(TAB_CLEAR), hdr=clearHdr_();
  row=guardRow_(sh,hdr,row,{clearance_id:key});
  const g=function(f){ const c=hdr.indexOf(f); return c===-1?'':fmt_(sh.getRange(row,c+1).getValue()); };
  const set=function(f,v){ const c=hdr.indexOf(f); if(c!==-1) sh.getRange(row,c+1).setValue(v); };

  if(dept==='it'  && !isIT_()  && !isHR_()) throw new Error('IT only.');
  if(dept==='fac' && !isFacilities_() && !isHR_()) throw new Error('Facilities only.');
  if(g(dept+'_status')==='Blocked') throw new Error('Handover has not been confirmed yet.');

  const list = dept==='it'? IT_ITEMS : FAC_ITEMS;
  let deduction=0; const missing=[]; const detail=[];
  list.forEach(function(it){
    const v=String(items[it.f]||'').trim();
    set(it.f, v);
    if(v==='Not returned' && !it.noAsset){
      // the amount is judged NOW — current value, condition, depreciation
      const amt=parseFloat((amounts||{})[it.f])||0;
      deduction+=amt;
      missing.push(it.label+(amt? (' (EGP '+amt+')') : ' (no amount entered)'));
      detail.push(it.label+': EGP '+amt);
    }
  });
  set(dept+'_deductions', deduction);
  if(detail.length) set(dept+'_note', (note? note+' | ':'')+detail.join('; '));
  else if(note) set(dept+'_note', note);
  set(dept+'_status','Cleared');
  set(dept+'_by',currentUser_()); set(dept+'_at',new Date());
  if(dept==='fac' && loans!==undefined){
    set('outstanding_loans', loans.has? 'Yes':'No');
    set('loan_amount', loans.amount||'');
    set('loan_note', loans.note||'');
  }
  recalcClearance_(row, hdr);
  return {ok:true, missing:missing, deduction:deduction};
}

// ---------- stage 3: HR final gate ----------
function hrCompleteClearance(row, items, note, key){
  if(!isHR_()) throw new Error('HR only.');
  const sh=sheet_(TAB_CLEAR), hdr=clearHdr_();
  row=guardRow_(sh,hdr,row,{clearance_id:key});
  const g=function(f){ const c=hdr.indexOf(f); return c===-1?'':fmt_(sh.getRange(row,c+1).getValue()); };
  const set=function(f,v){ const c=hdr.indexOf(f); if(c!==-1) sh.getRange(row,c+1).setValue(v); };
  if(g('it_status')!=='Cleared' || g('fac_status')!=='Cleared'){
    return {ok:false,msg:'IT and Facilities must clear first.'};
  }
  HR_ITEMS.forEach(function(it){ set(it.f, String(items[it.f]||'').trim()); });
  const allDone=HR_ITEMS.every(function(it){ return String(items[it.f]||'')==='Done'; });
  set('hr_status', allDone? 'Cleared':'In progress');
  set('hr_by',currentUser_()); set('hr_at',new Date());
  if(note) set('hr_note',note);
  if(allDone){
    set('final_status','Cleared'); set('cleared_at',new Date());
    set('settlement_released','Ready');
    // employee record -> Cleared
    const eid=g('employee_id');
    const esh=sheet_(TAB.EMP), ehdr=headers_(TAB.EMP), data=esh.getDataRange().getValues();
    const ei=ehdr.indexOf('employee_id');
    for(let r=1;r<data.length;r++){
      if(String(data[r][ei]).trim()!==eid) continue;
      const put=function(f,v){ const c=ehdr.indexOf(f); if(c!==-1) esh.getRange(r+1,c+1).setValue(v); };
      put('record_status','Cleared'); put('cleared_date',new Date());
      break;
    }
    clearEmpCache_();
    notifyClearanceOutcome_(row, hdr);
  }
  recalcClearance_(row, hdr);
  return {ok:true, cleared:allDone};
}

function recalcClearance_(row, hdr){
  const sh=sheet_(TAB_CLEAR);
  const g=function(f){ const c=hdr.indexOf(f); return c===-1?'':fmt_(sh.getRange(row,c+1).getValue()); };
  const set=function(f,v){ const c=hdr.indexOf(f); if(c!==-1) sh.getRange(row,c+1).setValue(v); };
  const total=(parseFloat(g('it_deductions'))||0)+(parseFloat(g('fac_deductions'))||0)+
              (parseFloat(g('loan_amount'))||0);
  set('total_deductions', total);
  if(g('it_status')==='Cleared' && g('fac_status')==='Cleared' && g('hr_status')==='Blocked'){
    set('hr_status','Pending');
    notifyHR_('Clearance ready for you — '+g('clearance_id'),
      g('employee_name')+' ('+g('employee_id')+')\n\n'+
      'IT and Facilities have cleared. Yours is the last step: exit formalities, hiring documents, '+
      'and notifying Comp & Ben to deactivate medical and social insurance.'+
      (total>0? ('\n\nDeductions to carry to payroll: EGP '+total) : ''));
  }
}

function notifyClearanceOutcome_(row, hdr){
  const sh=sheet_(TAB_CLEAR);
  const g=function(f){ const c=hdr.indexOf(f); return c===-1?'':fmt_(sh.getRange(row,c+1).getValue()); };
  const total=parseFloat(g('total_deductions'))||0;
  const to=g('konecta_email');
  if(to && total>0){
    try{ MailApp.sendEmail(to,'Your clearance — outstanding amount to be deducted',
      'Hello '+g('employee_name')+',\n\nYour clearance is complete.\n\n'+
      'The following will be deducted from your final payment:\n'+
      (parseFloat(g('it_deductions'))>0? ('  IT equipment not returned: EGP '+g('it_deductions')+'\n'):'')+
      (parseFloat(g('fac_deductions'))>0? ('  Facilities items not returned: EGP '+g('fac_deductions')+'\n'):'')+
      (parseFloat(g('loan_amount'))>0? ('  Outstanding loan or advance: EGP '+g('loan_amount')+'\n'):'')+
      '\n  Total: EGP '+total+
      '\n\nIf you believe any of this is wrong, contact HR before your final payment is processed.'+
      '\n\nKonecta Egypt — People team'); }catch(e){}
  }
  notifyHR_('CLEARED — '+g('clearance_id')+' — settlement can be released',
    g('employee_name')+' ('+g('employee_id')+') is fully cleared.\n\n'+
    'Last working day: '+g('last_working_day')+
    (total>0? ('\n\nDEDUCT FROM FINAL PAYOUT: EGP '+total) : '\n\nNo deductions.')+
    '\n\nRelease the final settlement, including the prorated leave balance.');
}

// ---------- queues ----------
function getMyClearanceTasks(){
  const acting=actingFor_();
  const it=isIT_(), fac=isFacilities_(), hr=isHR_();
  const sh=sheet_(TAB_CLEAR);
  if(!sh || sh.getLastRow()<2) return {items:[]};
  const hdr=clearHdr_(); const i=function(f){return hdr.indexOf(f);};

  // Read the employee sheet ONCE and index the two manager columns.
  // This used to call employeeFieldsOf_ twice per open clearance, and each of
  // those read all 835 rows by 103 columns — so twenty clearances meant forty
  // full-sheet reads before anything rendered. That is what made every panel
  // that shows clearance slow, for everyone.
  const mgrOf={};
  try{
    const E=empData_(true), eh=E.hdr;
    const cEid=eh.indexOf('employee_id'), cDm=eh.indexOf('direct_manager'), cDt=eh.indexOf('dotted_manager');
    E.rows.forEach(function(rec){
      const id=String(rec.values[cEid]).trim();
      if(id) mgrOf[id]={dm:fmt_(rec.values[cDm]), dt:fmt_(rec.values[cDt])};
    });
  }catch(e){}

  const out=[];
  sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues().forEach(function(r,idx){
    if(String(r[i('final_status')])==='Cleared') return;
    const rec={row:idx+2}; hdr.forEach(function(h,c){ rec[h]=fmt_(r[c]); });
    const eid=String(rec.employee_id).trim();
    const m=mgrOf[eid]||{dm:'',dt:''};
    const stages=[];
    if(acting.length){
      const dm=resolveApprover_(m.dm), dt=resolveApprover_(m.dt);
      if(dm && acting.indexOf(dm)!==-1 && rec.handover_direct_status==='Pending') stages.push('handover-direct');
      if(dt && acting.indexOf(dt)!==-1 && rec.handover_dotted_status==='Pending') stages.push('handover-dotted');
    }
    if((it||hr) && rec.it_status==='Pending') stages.push('it');
    if((fac||hr) && rec.fac_status==='Pending') stages.push('fac');
    if(hr && rec.hr_status==='Pending') stages.push('hr');
    if(!stages.length) return;
    rec.stages=stages;
    rec.itItems=IT_ITEMS; rec.facItems=FAC_ITEMS; rec.hrItems=HR_ITEMS;
    out.push(rec);
  });
  return {items:out, roles:{it:it,fac:fac,hr:hr,manager:acting.length>0}};
}
// ================================================================
// NO SHOW
//   Anyone can report anyone. Trainers, team leaders, project managers —
//   the reporter does not need to be the person's manager, because a trainee
//   who stops turning up may not have a manager assigned yet.
//
//   On report: payment goes on HOLD immediately, the record moves to On Hold,
//   and clearance opens. Everyone who needs to know is emailed.
//
//   The absence is NOT deducted from the leave balance — it is unpaid absence,
//   and it goes to payroll as a deduction. Disciplinary consequences run
//   separately, under the existing consequence framework.
// ================================================================

const TAB_NOSHOW = 'NO_SHOW';

// Two very different situations that share the same mechanics:
//   No show  — they have vanished. Nobody knows why. Urgent.
//   Drop out — they told us they are leaving, usually during probation. Known.
// Holding payment and opening clearance is the same. The message is not.
const DEPARTURE_TYPES = ['No show', 'Drop out'];

function noshowHdr_(){ const sh=sheet_(TAB_NOSHOW); return sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]; }

// Look up one employee by ID — deliberately NOT a dropdown of everyone,
// so reporting does not expose the whole employee list.
function lookupEmployeeForNoShow(employeeId){
  const eid=String(employeeId||'').trim().toUpperCase();
  if(!eid) return {found:false};
  const f=employeeFieldsOf_(eid,['employee_id','full_name_en','konecta_email','job_title',
                                 'record_status','direct_manager','dotted_manager','project']);
  if(!f.employee_id) return {found:false, msg:'No employee found with ID '+eid+'. Check the ID and try again.'};
  return {found:true, employee_id:f.employee_id, name:f.full_name_en, job_title:f.job_title,
          project:f.project, status:f.record_status};
}

function reportNoShow(p){
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const eid=String(p.employee_id||'').trim().toUpperCase();
    if(!eid) return {ok:false,msg:'Enter the employee ID of the person who has not shown up.'};
    const since=String(p.absent_since||'').trim();
    if(!since) return {ok:false,msg:'Enter the date they were last expected.'};

    const f=employeeFieldsOf_(eid,['employee_id','full_name_en','konecta_email',
                                   'record_status','direct_manager','dotted_manager']);
    if(!f.employee_id) return {ok:false,msg:'No employee found with ID '+eid+'.'};
    if(String(f.record_status)==='Closed') return {ok:false,msg:'That record is already closed.'};

    // who is reporting
    const me=currentUser_();
    const reporter=employeeByEmail_(me);

    const sh=sheet_(TAB_NOSHOW), hdr=noshowHdr_();
    const row=sh.getLastRow()+1;
    const id='NS-'+String(row-1).padStart(5,'0');
    const set=function(fl,v){ const c=hdr.indexOf(fl); if(c!==-1) sh.getRange(row,c+1).setValue(v); };

    // everyone who should know: HR always, the absent person's managers,
    // the reporter's own manager, and anyone the reporter added by hand
    const cc=String(p.cc_list||'').split(/[,;\s]+/).map(function(x){return x.trim();})
              .filter(function(x){ return /^[^@]+@[^@]+$/.test(x); });
    const recips={};
    HR_ADMINS.forEach(function(a){ recips[a]=true; });
    [f.direct_manager, f.dotted_manager].forEach(function(m){
      const e=emailForApprover_(resolveApprover_(m)); if(e) recips[e]=true;
    });
    if(reporter.direct_manager){
      const e=emailForApprover_(resolveApprover_(reporter.direct_manager)); if(e) recips[e]=true;
    }
    cc.forEach(function(e){ recips[e]=true; });
    const to=Object.keys(recips).join(',');

    set('noshow_id',id); set('reported_at',new Date());
    set('reported_by',me); set('reporter_employee_id',reporter.employee_id||'');
    set('reporter_name',reporter.full_name_en||'');
    set('employee_id',eid); set('employee_name',f.full_name_en);
    set('konecta_email',f.konecta_email);
    set('absent_since',since); set('last_seen_note',String(p.note||'').trim());
    set('event_type', String(p.event_type||'No show').trim());
    set('leaving_reason', String(p.leaving_reason||'').trim());
    set('reason_detail', String(p.reason_detail||'').trim());
    set('cc_list',cc.join(', ')); set('notified_to',to);
    set('hr_status','Open');

    // hold the money and park the record — before anyone investigates
    holdPaymentFor_(eid, String(p.event_type||'No show')+' '+id+' — from '+since);
    set('payment_held_at', new Date());

    // open clearance so equipment recovery starts
    let clid='';
    try{ const c=openClearance_(eid, ''); if(c && c.ok) clid=c.id; }catch(e){}
    set('clearance_id', clid);

    const isDrop = String(p.event_type||'No show')==='Drop out';
    const common =
      '\nWhat has happened automatically:\n'+
      '  - Payment is on hold pending clearance.\n'+
      '  - The record has been moved to On Hold.\n'+
      '  - Clearance has been opened so equipment can be recovered.\n';

    let subject, body;
    if(isDrop){
      subject='[Konecta] Drop out — '+f.full_name_en+' ('+eid+')';
      body =
        f.full_name_en+' ('+eid+') has left voluntarily.\n\n'+
        'Last day: '+since+'\n'+
        'Primary reason: '+(p.leaving_reason||'not recorded')+'\n'+
        (p.reason_detail? ('Detail: '+p.reason_detail+'\n') : '')+
        'Reported by: '+(reporter.full_name_en||me)+'\n'+
        (p.note? ('Note: '+p.note+'\n') : '')+
        common+
        '\nAny unpaid days are handled in payroll. This is not deducted from their leave balance.\n\n'+
        'HR will process the exit. Reference: '+id;
    } else {
      subject='[Konecta] NO SHOW — '+f.full_name_en+' ('+eid+')';
      body =
        f.full_name_en+' ('+eid+') has not shown up for work and we have had no contact.\n\n'+
        'Absent since: '+since+'\n'+
        (p.leaving_reason? ('Suspected reason: '+p.leaving_reason+'\n') : '')+
        (p.reason_detail? ('Detail: '+p.reason_detail+'\n') : '')+
        'Reported by: '+(reporter.full_name_en||me)+'\n'+
        (p.note? ('Note: '+p.note+'\n') : '')+
        common+
        '\nThis absence is NOT deducted from their leave balance. It is unpaid absence and '+
        'goes to payroll as a deduction.\n\n'+
        'If you know where this person is, or this has been reported in error, please reply as soon '+
        'as possible — the hold affects their pay.\n\nReference: '+id;
    }
    try{ MailApp.sendEmail(to, subject, body+'\n\nKonecta Egypt — People team'); }catch(e){}

    return {ok:true, id:id,
      msg:'Recorded as '+id+'. Payment is on hold, clearance has been opened, and everyone listed has been notified.'};
  } finally { lock.releaseLock(); }
}

function employeeByEmail_(email){
  const sh=sheet_(TAB.EMP), hdr=headers_(TAB.EMP), data=sh.getDataRange().getValues();
  const ke=hdr.indexOf('konecta_email');
  const out={};
  for(let r=1;r<data.length;r++){
    if(String(data[r][ke]).toLowerCase().trim()!==String(email).toLowerCase().trim()) continue;
    ['employee_id','full_name_en','direct_manager'].forEach(function(f){
      const c=hdr.indexOf(f); out[f]= c===-1?'':fmt_(data[r][c]);
    });
    break;
  }
  return out;
}

// The single place payment gets held, whatever the reason.
function holdPaymentFor_(eid, reason){
  const sh=sheet_(TAB.EMP), hdr=headers_(TAB.EMP), data=sh.getDataRange().getValues();
  const ei=hdr.indexOf('employee_id');
  for(let r=1;r<data.length;r++){
    if(String(data[r][ei]).trim()!==String(eid).trim()) continue;
    const put=function(f,v){ const c=hdr.indexOf(f); if(c!==-1) sh.getRange(r+1,c+1).setValue(v); };
    put('payment_status','Hold');
    put('hold_reason', reason);
    put('record_status','On Hold');
    logChange_(eid,'','payment_status','','Hold','System','Applied',reason);
    break;
  }
  clearEmpCache_();
}

// ---------- HR view ----------
function hrGetNoShows(){
  if(!isHR_()) throw new Error('HR only.');
  const sh=sheet_(TAB_NOSHOW);
  if(!sh || sh.getLastRow()<2) return [];
  const hdr=noshowHdr_(); const out=[];
  sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues().forEach(function(r,idx){
    const rec={row:idx+2}; hdr.forEach(function(h,c){ rec[h]=fmt_(r[c]); });
    if(rec.hr_status==='Closed') return;
    const since=new Date(rec.absent_since);
    rec.days_absent = isNaN(since)? '' : Math.floor((new Date()-since)/86400000);
    out.push(rec);
  });
  out.sort(function(a,b){ return (b.days_absent||0)-(a.days_absent||0); });
  return out;
}

// HR resolves it: the person came back, or it becomes a termination.
function hrResolveNoShow(row, outcome, note, key){
  if(!isHR_()) throw new Error('HR only.');
  const sh=sheet_(TAB_NOSHOW), hdr=noshowHdr_();
  row=guardRow_(sh,hdr,row,{noshow_id:key});
  const g=function(f){ const c=hdr.indexOf(f); return c===-1?'':fmt_(sh.getRange(row,c+1).getValue()); };
  const set=function(f,v){ const c=hdr.indexOf(f); if(c!==-1) sh.getRange(row,c+1).setValue(v); };
  const eid=g('employee_id');

  set('outcome',outcome); set('outcome_date',new Date());
  set('hr_status','Closed'); set('hr_by',currentUser_()); set('hr_at',new Date());
  if(note) set('notes',(g('notes')? g('notes')+' | ':'')+note);

  if(outcome==='Returned to work'){
    // release the hold and put them back
    const esh=sheet_(TAB.EMP), ehdr=headers_(TAB.EMP), data=esh.getDataRange().getValues();
    const ei=ehdr.indexOf('employee_id');
    for(let r=1;r<data.length;r++){
      if(String(data[r][ei]).trim()!==eid) continue;
      const put=function(f,v){ const c=ehdr.indexOf(f); if(c!==-1) esh.getRange(r+1,c+1).setValue(v); };
      put('payment_status','Release'); put('hold_reason',''); put('record_status','Active');
      break;
    }
    clearEmpCache_();
    logChange_(eid,'','payment_status','Hold','Release','HR console','Applied','No show resolved — returned to work');
    notifyHR_('No show resolved — '+g('noshow_id'),
      g('employee_name')+' has returned. Payment hold released, record back to Active.\n\n'+
      'The unpaid days still need deducting in payroll.');
  } else {
    notifyHR_('No show — '+outcome+' — '+g('noshow_id'),
      g('employee_name')+' ('+eid+')\n\nOutcome: '+outcome+
      '\n\nPayment remains on hold. Complete clearance before releasing anything.');
  }
  return {ok:true};
}


// ================================================================
// CONTRACT EXPIRY WARNING
//   60 days before a fixed-period contract ends, the direct manager gets a
//   note listing THEIR people, and HR gets one consolidated list.
//
//   This matters more than it looks: contract end dates cluster on
//   31 December and 30 June, so expiry arrives in batches. Without a prompt
//   people are renewed by default or lapse unnoticed — both bad.
//
//   Add ONE daily time-driven trigger: contractExpiryRun.
// ================================================================

const EXPIRY_WARN_DAYS = 60;
const EXPIRY_REMINDER_TAB = 'EXPIRY_LOG';   // so nobody is warned twice for the same contract

function contractExpiryRun(){
  assertNotDirectCall_();
  const E=empData_(false), h=E.hdr;
  const col=function(f){return h.indexOf(f);};
  const cEid=col('employee_id'), cNm=col('full_name_en'), cEnd=col('contract_end_date'),
        cType=col('contract_type'), cDm=col('direct_manager'), cJt=col('job_title'),
        cProj=col('project'), cHire=col('hire_date'), cSt=col('record_status');
  if(cEnd===-1) return {error:'contract_end_date column not found'};

  const today=new Date(); today.setHours(0,0,0,0);
  const already=expiryAlreadyWarned_();
  const byManager={}; const all=[];

  E.rows.forEach(function(rec){
    const v=rec.values;
    const st=String(v[cSt]).trim();
    if(['Active','Serving Notice'].indexOf(st)===-1) return;      // already leaving, or not started
    const endRaw=fmt_(v[cEnd]);
    if(!endRaw) return;
    const end=new Date(endRaw); if(isNaN(end)) return;
    const days=Math.round((end-today)/86400000);
    if(days<0 || days>EXPIRY_WARN_DAYS) return;                    // only inside the window

    const eid=String(v[cEid]).trim();
    const key=eid+'|'+endRaw;
    if(already[key]) return;                                       // warned already for this contract

    const item={employee_id:eid, name:fmt_(v[cNm]), job_title:fmt_(v[cJt]),
                project:fmt_(v[cProj]), hire_date:fmt_(v[cHire]),
                contract_end:endRaw, days_left:days,
                contract_type:fmt_(v[cType]), manager:String(v[cDm]).trim()};
    all.push(item);
    const m=resolveApprover_(item.manager) || '(no manager)';
    (byManager[m]=byManager[m]||[]).push(item);
  });

  if(!all.length) return {warned:0};

  // one email per manager, listing only their people
  Object.keys(byManager).forEach(function(mgr){
    const to=emailForApprover_(mgr);
    if(!to) return;
    const list=byManager[mgr];
    let body='The following contracts in your team end within the next '+EXPIRY_WARN_DAYS+' days.\\n\\n';
    list.forEach(function(x){
      body+='  '+x.name+' ('+x.employee_id+')\\n'+
            '    '+(x.job_title||'')+(x.project? (' — '+x.project):'')+'\\n'+
            '    Contract ends '+x.contract_end+'  ('+x.days_left+' days)\\n\\n';
    });
    body+='Please confirm for each whether the contract is being renewed.\\n\\n'+
          'If nothing is done, the contract simply lapses on its end date — which is rarely what anyone intended. '+
          'Speak to HR if you want to renew, extend, or let it end.\\n';
    try{ MailApp.sendEmail(to,'[Konecta] Contracts ending soon — '+list.length+' in your team',
      body+'\\nKonecta Egypt — People team'); }catch(e){}
  });

  // one consolidated list to HR
  all.sort(function(a,b){ return a.contract_end.localeCompare(b.contract_end) || a.name.localeCompare(b.name); });
  let hrBody='These contracts end within '+EXPIRY_WARN_DAYS+' days. Each manager has been emailed their own people.\\n\\n';
  let currentDate='';
  all.forEach(function(x){
    if(x.contract_end!==currentDate){
      currentDate=x.contract_end;
      hrBody+='\\n--- ending '+currentDate+' ---\\n';
    }
    hrBody+='  '+x.employee_id+'  '+x.name+'  ('+(x.job_title||'')+')  manager: '+(x.manager||'none')+'\\n';
  });
  hrBody+='\\nTotal: '+all.length+' contract(s).\\n';
  notifyHR_('Contracts ending within '+EXPIRY_WARN_DAYS+' days — '+all.length+' people', hrBody);

  logExpiryWarned_(all);
  return {warned:all.length, managers:Object.keys(byManager).length};
}

function expiryAlreadyWarned_(){
  const sh=sheet_(EXPIRY_REMINDER_TAB);
  if(!sh || sh.getLastRow()<2) return {};
  const out={};
  sh.getRange(2,1,sh.getLastRow()-1,2).getValues().forEach(function(r){
    if(r[0]) out[String(r[0]).trim()+'|'+fmt_(r[1])]=true;
  });
  return out;
}

function logExpiryWarned_(items){
  const sh=sheet_(EXPIRY_REMINDER_TAB);
  if(!sh) return;
  const rows=items.map(function(x){ return [x.employee_id, x.contract_end, new Date(), x.name, x.manager]; });
  sh.getRange(sh.getLastRow()+1,1,rows.length,5).setValues(rows);
}

// HR view: everything expiring, whether or not a warning has gone out
function hrContractsExpiring(days){
  if(!isHR_()) throw new Error('HR only.');
  const window=parseInt(days)||90;
  const E=empData_(false), h=E.hdr;
  const col=function(f){return h.indexOf(f);};
  const cEid=col('employee_id'), cNm=col('full_name_en'), cEnd=col('contract_end_date'),
        cJt=col('job_title'), cProj=col('project'), cDm=col('direct_manager'),
        cType=col('contract_type'), cSt=col('record_status');
  const today=new Date(); today.setHours(0,0,0,0);
  const out=[];
  E.rows.forEach(function(rec){
    const v=rec.values;
    if(['Active','Serving Notice'].indexOf(String(v[cSt]).trim())===-1) return;
    const endRaw=fmt_(v[cEnd]); if(!endRaw) return;
    const end=new Date(endRaw); if(isNaN(end)) return;
    const d=Math.round((end-today)/86400000);
    if(d<0 || d>window) return;
    out.push({row:rec.row, employee_id:fmt_(v[cEid]), name:fmt_(v[cNm]),
              job_title:fmt_(v[cJt]), project:fmt_(v[cProj]),
              contract_end:endRaw, days_left:d,
              contract_type:fmt_(v[cType]), manager:String(v[cDm]).trim()});
  });
  out.sort(function(a,b){ return a.days_left-b.days_left; });
  return out;
}


// ================================================================
// TERMINATION — company-initiated exits
//
//   One form, four reasons. They share the same shape: someone initiates,
//   the record is created, clearance opens, people are notified. What differs
//   is what the reason requires and whether HR must APPROVE or is simply told.
//
//   Violation           HR APPROVES. A manager must give evidence; HR reviews
//                       it before anyone is dismissed. HR can also initiate
//                       directly — a harassment or drugs case may not come
//                       through the manager, and may concern them.
//   Failed probation    HR notified. BLOCKED after the legal window closes.
//   WFR                 HR notified. Already approved in the plan.
//   Contract non-renewal HR notified. The decision is not to act.
//
//   The skip manager (one level above the direct manager) is always notified,
//   because a violation may involve the direct manager themselves.
// ================================================================

const TAB_TERM = 'TERMINATIONS';
const PROBATION_DAYS = 90;                 // Egyptian labour law: three months

const TERM_REASONS = {
  'Violation':                  {hrApproves:true,  needsEvidence:true,  code:'D'},
  'Failed probation':           {hrApproves:false, needsEvidence:false, code:'N'},
  'Workforce reduction (WFR)':  {hrApproves:false, needsPlan:true,      code:'D'},
  'Contract ended — no renewal':{hrApproves:false, needsEvidence:false, code:'F'}
};

function termHdr_(){ const sh=sheet_(TAB_TERM); return sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]; }

// One level above the direct manager. Derived, so it never goes stale.
function skipManagerOf_(eid){
  const dm=directManagerOf_(eid);
  if(!dm) return '';
  const resolved=resolveApprover_(dm);
  if(resolved.indexOf('GLOBAL-')===0) return resolved;   // already at the top
  return directManagerOf_(resolved) || '';
}

function getTerminationContext(employeeId){
  const eid=String(employeeId||'').trim().toUpperCase();
  if(!eid) return {found:false};
  const f=employeeFieldsOf_(eid,['employee_id','full_name_en','konecta_email','job_title',
          'hire_date','record_status','direct_manager','dotted_manager','contract_end_date',
          'contract_type','probation_end_date','project']);
  if(!f.employee_id) return {found:false,msg:'No employee found with ID '+eid+'.'};

  // Same rule as initiateTermination: HR, or the person's direct manager
  // (delegates included via actingFor_). This function was the one ungated
  // door in the app — without this check any employee could pull any
  // colleague's hire date, contract end, managers and probation status from
  // the browser console.
  if(!isHR_()){
    const acting=actingFor_();
    if(acting.indexOf(resolveApprover_(f.direct_manager))===-1)
      throw new Error('Only the direct manager or HR can view this.');
  }

  // is the probation window still open?
  let probation={applies:false};
  if(f.hire_date){
    const hire=new Date(f.hire_date);
    if(!isNaN(hire)){
      const days=Math.floor((new Date()-hire)/86400000);
      probation={applies:true, daysSinceHire:days, windowDays:PROBATION_DAYS,
                 open: days<=PROBATION_DAYS, daysLeft: PROBATION_DAYS-days,
                 endDate: f.probation_end_date||''};
    }
  }
  // konecta_email is used by initiateTermination when it writes the
  // termination row — it was fetched above but never returned, so every
  // termination stored a blank email and the employee notifications that
  // read that column silently went nowhere.
  return {found:true, employee_id:f.employee_id, name:f.full_name_en, job_title:f.job_title,
          project:f.project, status:f.record_status, hire_date:f.hire_date,
          konecta_email:f.konecta_email,
          contract_end:f.contract_end_date, contract_type:f.contract_type,
          direct_manager:f.direct_manager, dotted_manager:f.dotted_manager,
          skip_manager:skipManagerOf_(eid), probation:probation,
          reasons:Object.keys(TERM_REASONS)};
}

function initiateTermination(p){
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const eid=String(p.employee_id||'').trim().toUpperCase();
    const reason=String(p.reason||'').trim();
    const rule=TERM_REASONS[reason];
    if(!rule) return {ok:false,msg:'Choose a termination reason.'};

    const ctx=getTerminationContext(eid);
    if(!ctx.found) return {ok:false,msg:ctx.msg||'Employee not found.'};
    if(['Closed','Cleared'].indexOf(String(ctx.status))!==-1) return {ok:false,msg:'That record is already closed.'};

    const me=currentUser_();
    const isHR=isHR_();
    const acting=actingFor_();
    const isManager = acting.indexOf(resolveApprover_(ctx.direct_manager))!==-1;
    if(!isHR && !isManager) return {ok:false,msg:'Only the direct manager or HR can start this.'};

    // Probation terminations are blocked once the legal window has closed.
    // Past it the employee is confirmed automatically and dismissal on those
    // grounds would not stand.
    if(reason==='Failed probation' && ctx.probation.applies && !ctx.probation.open){
      return {ok:false, blocked:true,
        msg:'The probation window closed '+Math.abs(ctx.probation.daysLeft)+' day(s) ago — '+
            ctx.probation.daysSinceHire+' days since they were hired, and the window is '+PROBATION_DAYS+' days. '+
            'They are confirmed in role, so this cannot go through as a probation failure. Speak to HR about the options.'};
    }
    // A manager asking for a dismissal must say why, with something behind it.
    if(rule.needsEvidence && !isHR && !String(p.evidence_note||'').trim()){
      return {ok:false,msg:'Describe the evidence. HR cannot approve a dismissal on assertion alone — '+
                           'set out what happened, when, and any warnings already issued.'};
    }
    if(rule.needsPlan && !String(p.plan_reference||'').trim()){
      return {ok:false,msg:'Give the approved plan reference this reduction sits under.'};
    }

    const sh=sheet_(TAB_TERM), hdr=termHdr_();
    const row=sh.getLastRow()+1;
    const id='TM-'+String(row-1).padStart(5,'0');
    const set=function(f,v){ const c=hdr.indexOf(f); if(c!==-1) sh.getRange(row,c+1).setValue(v); };

    // if the case concerns the direct manager, HR can keep them out of it
    const excludeManager = isHR && !!p.exclude_manager;

    set('termination_id',id); set('initiated_at',new Date());
    set('initiated_by',me); set('initiator_role', isHR? 'HR':'Manager');
    set('reason',reason);
    set('employee_id',eid); set('employee_name',ctx.name); set('konecta_email',ctx.konecta_email||'');
    set('direct_manager',ctx.direct_manager); set('dotted_manager',ctx.dotted_manager);
    set('skip_manager',ctx.skip_manager);
    set('proposed_last_day',String(p.proposed_last_day||'').trim());
    set('details',String(p.details||'').trim());
    set('evidence_note',String(p.evidence_note||'').trim());
    set('plan_reference',String(p.plan_reference||'').trim());
    set('exclude_manager', excludeManager? 'Yes':'No');
    set('exclude_reason', excludeManager? String(p.exclude_reason||'').trim() : '');
    set('exit_type_code', rule.code);

    if(rule.hrApproves && !isHR){
      set('hr_status','Pending'); set('final_status','Awaiting HR approval');
    } else {
      set('hr_status', isHR? 'Initiated by HR':'Notified');
      set('final_status','Approved');
    }

    const recips=terminationRecipients_(ctx, excludeManager);
    set('notified_to', recips.join(','));

    // Only proceed to the exit if it does not need HR's approval first.
    let clid='';
    if(!(rule.hrApproves && !isHR)){
      clid=applyTermination_(row, hdr, ctx, reason, rule, p.proposed_last_day);
    }
    set('clearance_id', clid);

    notifyTermination_(row, hdr, ctx, reason, rule, recips, excludeManager, isHR);

    return {ok:true, id:id,
      msg: (rule.hrApproves && !isHR)
        ? 'Submitted as '+id+'. HR will review the evidence before any decision is taken. Nothing has changed for the employee yet.'
        : 'Recorded as '+id+'. Payment is on hold, clearance has been opened, and everyone who needs to know has been notified.'};
  } finally { lock.releaseLock(); }
}

// HR approves or rejects a violation raised by a manager
function hrDecideTermination(row, approve, note, key){
  if(!isHR_()) throw new Error('HR only.');
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const sh=sheet_(TAB_TERM), hdr=termHdr_();
    row=guardRow_(sh,hdr,row,{termination_id:key});
    const g=function(f){ const c=hdr.indexOf(f); return c===-1?'':fmt_(sh.getRange(row,c+1).getValue()); };
    const set=function(f,v){ const c=hdr.indexOf(f); if(c!==-1) sh.getRange(row,c+1).setValue(v); };
    if(g('hr_status')!=='Pending') throw new Error('This is not awaiting a decision.');

    set('hr_by',currentUser_()); set('hr_at',new Date());
    if(note) set('hr_note',note);

    if(!approve){
      set('hr_status','Rejected'); set('final_status','Rejected');
      const to=[emailForApprover_(resolveApprover_(g('direct_manager')))].filter(String).join(',');
      if(to){ try{ MailApp.sendEmail(to,'[Konecta] Termination not approved — '+g('employee_name'),
        'The termination you raised for '+g('employee_name')+' ('+g('employee_id')+') has not been approved.'+
        (note? ('\\n\\nHR note: '+note) : '')+
        '\\n\\nNothing has changed for the employee. Speak to HR if you want to discuss it.'+
        '\\n\\nReference: '+g('termination_id')+'\\n\\nKonecta Egypt — People team'); }catch(e){} }
      return {ok:true, approved:false};
    }

    set('hr_status','Approved'); set('final_status','Approved');
    const ctx=getTerminationContext(g('employee_id'));
    const rule=TERM_REASONS[g('reason')]||{code:'D'};
    const clid=applyTermination_(row, hdr, ctx, g('reason'), rule, g('proposed_last_day'));
    set('clearance_id', clid);
    const recips=String(g('notified_to')||'').split(',').filter(String);
    notifyTermination_(row, hdr, ctx, g('reason'), rule, recips, g('exclude_manager')==='Yes', true, true);
    return {ok:true, approved:true};
  } finally { lock.releaseLock(); }
}

// The actual exit: hold payment, set the exit fields, open clearance.
function applyTermination_(row, hdr, ctx, reason, rule, lastDay){
  const eid=ctx.employee_id;
  const when = String(lastDay||'').trim() ||
               Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');
  const esh=sheet_(TAB.EMP), ehdr=headers_(TAB.EMP), data=esh.getDataRange().getValues();
  const ei=ehdr.indexOf('employee_id');
  for(let r=1;r<data.length;r++){
    if(String(data[r][ei]).trim()!==eid) continue;
    const put=function(f,v){ const c=ehdr.indexOf(f); if(c!==-1) esh.getRange(r+1,c+1).setValue(v); };
    put('record_status','On Hold');
    put('exit_date',when); put('last_working_day',when);
    put('exit_type', reason==='Failed probation' ? 'Probation not passed'
                    : reason==='Contract ended — no renewal' ? 'Contract ended' : 'Dismissal');
    put('payment_status','Hold');
    put('hold_reason','Termination — '+reason+' — pending clearance');
    break;
  }
  clearEmpCache_();
  logChange_(eid,'','record_status','','On Hold','Termination','Applied',reason);

  let clid='';
  try{ const c=openClearance_(eid, when); if(c && c.ok) clid=c.id; }catch(e){}
  return clid;
}

function terminationRecipients_(ctx, excludeManager){
  const out={};
  HR_ADMINS.forEach(function(a){ out[a]=true; });
  if(!excludeManager){
    const dm=emailForApprover_(resolveApprover_(ctx.direct_manager)); if(dm) out[dm]=true;
  }
  const dt=emailForApprover_(resolveApprover_(ctx.dotted_manager)); if(dt) out[dt]=true;   // always
  const sk=emailForApprover_(resolveApprover_(ctx.skip_manager));   if(sk) out[sk]=true;   // always
  return Object.keys(out);
}

function notifyTermination_(row, hdr, ctx, reason, rule, recips, excludeManager, byHR, isApproval){
  const sh=sheet_(TAB_TERM);
  const g=function(f){ const c=hdr.indexOf(f); return c===-1?'':fmt_(sh.getRange(row,c+1).getValue()); };
  const id=g('termination_id');
  const to=(recips||[]).filter(String).join(',');
  if(!to) return;

  const pending = rule.hrApproves && !byHR && !isApproval;
  let body=ctx.name+' ('+ctx.employee_id+')\\n\\n'+
    'Reason: '+reason+'\\n'+
    (g('proposed_last_day')? ('Last working day: '+g('proposed_last_day')+'\\n') : '')+
    'Raised by: '+g('initiated_by')+' ('+g('initiator_role')+')\\n'+
    (g('details')? ('\\nDetails: '+g('details')+'\\n') : '')+
    (g('plan_reference')? ('Plan reference: '+g('plan_reference')+'\\n') : '');

  if(pending){
    body+='\\nThis is WITH HR FOR REVIEW. Nothing has changed for the employee yet — no hold, no exit date. '+
          'HR will review the evidence and decide.\\n';
    if(rule.needsEvidence){
      body+='\\nIf there are documents supporting this, email them to HR with '+id+' in the subject line.\\n';
    }
  } else {
    body+='\\nWhat has happened:\\n'+
          '  - Payment is on hold.\\n'+
          '  - The record is On Hold and the exit date is set.\\n'+
          '  - Clearance has been opened so equipment can be recovered.\\n';
  }
  if(excludeManager){
    body+='\\nThe direct manager has deliberately not been copied on this.\\n';
  }
  body+='\\nReference: '+id;

  const subject='[Konecta] '+(pending? 'Termination raised — HR review' : 'Termination — '+reason)+
                ' — '+ctx.name;
  try{ MailApp.sendEmail(to, subject, body+'\\n\\nKonecta Egypt — People team'); }catch(e){}
}

// HR queue: violations awaiting a decision
function hrGetTerminations(){
  if(!isHR_()) throw new Error('HR only.');
  const sh=sheet_(TAB_TERM);
  if(!sh || sh.getLastRow()<2) return [];
  const hdr=termHdr_(); const out=[];
  sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues().forEach(function(r,idx){
    const rec={row:idx+2}; hdr.forEach(function(h,c){ rec[h]=fmt_(r[c]); });
    if(rec.hr_status!=='Pending') return;
    out.push(rec);
  });
  return out;
}

// ================================================================
// MEDICAL INSURANCE
//   Enrolment is triggered by TWO conditions together: the contract is signed
//   AND the employee record is complete. Either alone is not enough — a signed
//   contract with half a record means we cannot send the insurer usable details.
//
//   Removal is deliberately NOT gated behind clearance. Cover has to stop on the
//   last working day whether or not the laptop is back. Waiting for clearance
//   means paying premiums for someone who has left, and leaving them able to
//   claim on a policy they are no longer entitled to.
// ================================================================

const MEDICAL_CONTACT = 'hagar.mostafa@konecta.com';
const TAB_MEDICAL = 'MEDICAL_INSURANCE';

function medHdr_(){ const sh=sheet_(TAB_MEDICAL); return sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]; }

function medicalRecordFor_(eid){
  const sh=sheet_(TAB_MEDICAL);
  if(!sh || sh.getLastRow()<2) return null;
  const hdr=medHdr_(), i=function(f){return hdr.indexOf(f);};
  const rows=sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
  for(let r=0;r<rows.length;r++){
    if(String(rows[r][i('employee_id')]).trim()!==eid) continue;
    const o={row:r+2}; hdr.forEach(function(h,c){ o[h]=fmt_(rows[r][c]); });
    return o;
  }
  return null;
}

// Everything the insurer needs, gathered from the record
function medicalPayloadFor_(eid){
  eid=String(eid||'').trim().toUpperCase();
  const f=employeeFieldsOf_(eid,['employee_id','full_name_en','full_name_ar','national_id',
    'date_of_birth','gender','hire_date','job_title','konecta_email','mobile',
    'bank_name','account_number','iban','company_type']);
  const all=dependantsFor_(eid);
  const live=all.filter(function(d){
    const s=String(d.status||'').trim();
    return s!=='Removed' && s!=='Not eligible';
  });
  const deps=live.map(function(d){
    return {name:d.name, dob:d.date_of_birth, relation:d.relation,
            national_id:d.national_id||'', funding:d.funding||'Company',
            status:d.status||'', insurance_id:d.insurance_id||'',
            premium_amount:d.premium_amount||'', row:d.row};
  });
  return {emp:f, dependants:deps,
          pending: deps.filter(function(d){return d.status==='Pending enrolment';}),
          enrolled: deps.filter(function(d){return d.status==='Enrolled';}),
          paid: deps.filter(function(d){return d.funding==='Employee-paid';})};
}

// Ready to enrol? Contract signed AND record complete.
function medicalEnrolmentReady_(eid){
  const f=employeeFieldsOf_(eid,['completeness_%','record_status']);
  const complete = (parseInt(f['completeness_%'])||0) >= 100;
  const fs=sheet_(TAB_FILE_STATUS);
  let signed=false;
  if(fs && fs.getLastRow()>1){
    const hdr=fileStatusHdr_(), i=function(x){return hdr.indexOf(x);};
    const rows=fs.getRange(2,1,fs.getLastRow()-1,fs.getLastColumn()).getValues();
    for(let r=0;r<rows.length;r++){
      if(String(rows[r][i('employee_id')]).trim()!==eid) continue;
      signed = String(rows[r][i('contract_signed')])==='Yes';
      break;
    }
  }
  return {ready: complete && signed, complete:complete, signed:signed};
}

// HR sends the enrolment
function medicalEnrol(eid, note){
  if(!isHR_()) throw new Error('HR only.');
  eid=String(eid).trim().toUpperCase();
  const check=medicalEnrolmentReady_(eid);
  if(!check.ready){
    const why=[];
    if(!check.signed) why.push('the contract is not recorded as signed');
    if(!check.complete) why.push('the employee record is not yet complete');
    return {ok:false, msg:'Cannot enrol yet — '+why.join(' and ')+'.'};
  }
  const existing=medicalRecordFor_(eid);
  if(existing && existing.status==='Enrolled'){
    return {ok:false, msg:'Already enrolled on '+existing.enrolled_at+'.'};
  }
  const p=medicalPayloadFor_(eid);

  let body='Please add the following to the medical insurance scheme.\n\n'+
    'EMPLOYEE\n'+
    '  Name (EN): '+p.emp.full_name_en+'\n'+
    '  Name (AR): '+(p.emp.full_name_ar||'')+'\n'+
    '  Employee ID: '+eid+'\n'+
    '  National ID: '+p.emp.national_id+'\n'+
    '  Date of birth: '+p.emp.date_of_birth+'\n'+
    '  Gender: '+p.emp.gender+'\n'+
    '  Hire date: '+p.emp.hire_date+'\n'+
    '  Job title: '+p.emp.job_title+'\n'+
    '  Email: '+p.emp.konecta_email+'\n'+
    '  Mobile: '+p.emp.mobile+'\n\n'+
    'BANK — for claim reimbursement\n'+
    '  Bank: '+(p.emp.bank_name||'not on file')+'\n'+
    '  Account: '+(p.emp.account_number||'not on file')+'\n'+
    '  IBAN: '+(p.emp.iban||'not on file')+'\n\n';
  if(p.dependants.length){
    body+='DEPENDANTS ('+p.dependants.length+')\n';
    p.dependants.forEach(function(d,i){
      body+='  '+(i+1)+'. '+d.name+'\n'+
            '     Relationship: '+d.relation+'\n'+
            '     Date of birth: '+d.dob+'\n'+
            '     National ID: '+(d.national_id||'MISSING — chase before enrolling')+'\n';
    });
  } else {
    body+='DEPENDANTS: none declared\n';
  }
  if(note) body+='\nNote: '+note+'\n';

  try{ MailApp.sendEmail({to:MEDICAL_CONTACT, cc:HR_ADMINS.join(','),
    subject:'[Konecta] Medical insurance — please enrol '+p.emp.full_name_en+' ('+eid+')',
    body:body+'\nKonecta Egypt — People team'}); }catch(e){}

  const sh=sheet_(TAB_MEDICAL), hdr=medHdr_();
  const row=existing? existing.row : sh.getLastRow()+1;
  const set=function(f,v){ const c=hdr.indexOf(f); if(c!==-1) sh.getRange(row,c+1).setValue(v); };
  set('employee_id',eid); set('employee_name',p.emp.full_name_en);
  set('status','Enrolled'); set('enrolled_at',new Date()); set('enrolled_by',currentUser_());
  set('dependants_count',p.dependants.length);
  if(note) set('notes',note);
  return {ok:true, msg:'Enrolment sent to '+MEDICAL_CONTACT+'.'+
    (p.dependants.filter(function(d){return !d.national_id;}).length
      ? ' Note: some dependants have no national ID — chase those.' : '')};
}

// HR stops cover. Available as soon as an exit is known — not gated behind clearance.
function medicalRemove(eid, lastWorkingDay, note){
  if(!isHR_()) throw new Error('HR only.');
  eid=String(eid).trim().toUpperCase();
  const p=medicalPayloadFor_(eid);
  if(!p.emp.employee_id) return {ok:false,msg:'No employee found.'};
  const existing=medicalRecordFor_(eid);
  if(existing && existing.status==='Removed'){
    return {ok:false,msg:'Cover was already stopped on '+existing.removed_at+'.'};
  }
  const last=String(lastWorkingDay||'').trim() ||
             employeeFieldsOf_(eid,['last_working_day']).last_working_day || '';

  let body='Please remove the following from the medical insurance scheme.\n\n'+
    '  Name: '+p.emp.full_name_en+'\n'+
    '  Employee ID: '+eid+'\n'+
    '  National ID: '+p.emp.national_id+'\n'+
    '  Insurance ID: '+((existing&&existing.insurance_id)||'not on file')+'\n'+
    '  LAST WORKING DAY: '+(last||'not recorded — please confirm with HR')+'\n\n'+
    'Cover should stop for the employee AND all dependants listed below.\n\n';
  if(p.dependants.length){
    body+='DEPENDANTS TO REMOVE ('+p.dependants.length+')\n';
    p.dependants.forEach(function(d,i){
      body+='  '+(i+1)+'. '+d.name+' — '+d.relation+' — '+(d.national_id||'no national ID on file')+'\n';
    });
  } else {
    body+='DEPENDANTS: none on record\n';
  }
  if(note) body+='\nNote: '+note+'\n';

  try{ MailApp.sendEmail({to:MEDICAL_CONTACT, cc:HR_ADMINS.join(','),
    subject:'[Konecta] Medical insurance — STOP cover for '+p.emp.full_name_en+' ('+eid+')',
    body:body+'\nKonecta Egypt — People team'}); }catch(e){}

  const sh=sheet_(TAB_MEDICAL), hdr=medHdr_();
  const row=existing? existing.row : sh.getLastRow()+1;
  const set=function(f,v){ const c=hdr.indexOf(f); if(c!==-1) sh.getRange(row,c+1).setValue(v); };
  set('employee_id',eid); set('employee_name',p.emp.full_name_en);
  set('status','Removed'); set('last_working_day',last);
  set('removed_at',new Date()); set('removed_by',currentUser_());
  if(note) set('notes',note);
  return {ok:true, msg:'Removal sent to '+MEDICAL_CONTACT+'. Cover stops '+(last||'— confirm the date')+'.'};
}

// Anyone who has left but whose cover has not been stopped.
// This is the kind of thing that costs money quietly, so it goes on the task list.
function medicalOutstanding(){
  if(!isHR_()) throw new Error('HR only.');
  const E=empData_(true), h=E.hdr;                // include leavers
  const col=function(f){return h.indexOf(f);};
  const cEid=col('employee_id'), cNm=col('full_name_en'),
        cSt=col('record_status'), cLwd=col('last_working_day');
  const out=[];
  E.rows.forEach(function(rec){
    const v=rec.values;
    const st=String(v[cSt]).trim();
    if(['Serving Notice','Final Month','On Hold','Cleared','Closed','Tail'].indexOf(st)===-1) return;
    const lwd=fmt_(v[cLwd]); if(!lwd) return;
    const eid=String(v[cEid]).trim();
    const m=medicalRecordFor_(eid);
    if(m && m.status==='Removed') return;
    const days=Math.floor((new Date()-new Date(lwd))/86400000);
    out.push({employee_id:eid, name:fmt_(v[cNm]), status:st,
              last_working_day:lwd, days_since:isNaN(days)?'':days,
              ever_enrolled: !!(m && m.status==='Enrolled')});
  });
  out.sort(function(a,b){ return (b.days_since||0)-(a.days_since||0); });
  return out;
}
// ================================================================
// EMPLOYEE FILE REVIEW
// The physical file remains the record of truth. This exists so you can
// see who is missing what without opening every folder.
//   DOC_CHECKLIST : the master list of documents  (doc_no, doc_name_en,
//                   doc_name_ar, applies_to, notes)
//   EMPLOYEE_DOCS : one row per employee per document, with who and when
//   FILE_STATUS   : the roll-up, one row per employee
// ================================================================

const TAB_DOC_CHECKLIST  = 'DOC_CHECKLIST';
const TAB_FILE_STATUS    = 'FILE_STATUS';
const TAB_EMPLOYEE_DOCS  = 'EMPLOYEE_DOCS';

// Employment relationships where Konecta is NOT the employer of record.
// Their vendor holds the employment paperwork; we hold the NDA and consent only.
const NON_EMPLOYEE_MARKERS = ['outsourc','vendor','subcontract','agency','third party',
                              'freelance','consultant','intern','apprentice','external'];

// What each applies_to value means. Anything not listed here is treated as
// "always required", so a new value can never silently hide a document.
const APPLIES_RULES = {
  'all':              {show:function(){return true;},                 optional:false, note:''},
  'males':            {show:function(e){return e.gender==='male';},   optional:false, note:''},
  'females':          {show:function(e){return e.gender==='female';}, optional:false, note:''},
  'egyptian':         {show:function(e){return e.egyptian;},          optional:false, note:''},
  'non_egyptian':     {show:function(e){return !e.egyptian;},         optional:false, note:''},
  'if_any':           {show:function(){return true;},                 optional:true,  note:'Only if they have any'},
  'if_worked_before': {show:function(){return true;},                 optional:true,  note:'Only if previously employed'}
};

function fileChecklist_(who){
  const sh=sheet_(TAB_DOC_CHECKLIST);
  if(!sh || sh.getLastRow()<2) return [];
  const hdr=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const i=function(f){return hdr.indexOf(f);};
  const rows=sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
  const out=[];
  rows.forEach(function(r){
    const no=String(r[i('doc_no')]||'').trim();
    const en=String(r[i('doc_name_en')]||'').trim();
    if(!no || !en) return;
    const key=String(r[i('applies_to')]||'all').trim().toLowerCase();
    const rule=APPLIES_RULES[key] || APPLIES_RULES['all'];
    if(who && !rule.show(who)) return;
    const ar = i('doc_name_ar')===-1? '' : String(r[i('doc_name_ar')]||'').trim();
    const own= i('notes')===-1? '' : String(r[i('notes')]||'').trim();
    out.push({no:no, name: en + (ar? ' — '+ar : ''),
              optional: rule.optional, note: own || rule.note,
              mandatory: !rule.optional});
  });
  return out;
}

function fileStatusHdr_(){
  const sh=sheet_(TAB_FILE_STATUS);
  if(!sh) throw new Error('FILE_STATUS tab not found.');
  return sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
}
function docsHdr_(){
  const sh=sheet_(TAB_EMPLOYEE_DOCS);
  if(!sh) throw new Error('EMPLOYEE_DOCS tab not found.');
  return sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
}

function fileStatusFor_(eid){
  const sh=sheet_(TAB_FILE_STATUS);
  if(!sh || sh.getLastRow()<2) return null;
  const hdr=fileStatusHdr_(), i=function(f){return hdr.indexOf(f);};
  const rows=sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
  for(let r=0;r<rows.length;r++){
    if(String(rows[r][i('employee_id')]).trim().toUpperCase()!==eid) continue;
    const o={row:r+2}; hdr.forEach(function(h,c){ o[h]=fmt_(rows[r][c]); });
    return o;
  }
  return null;
}

// Every document row already recorded for this employee, keyed by doc_no.
function docsFor_(eid){
  const sh=sheet_(TAB_EMPLOYEE_DOCS);
  const out={};
  if(!sh || sh.getLastRow()<2) return out;
  const hdr=docsHdr_(), i=function(f){return hdr.indexOf(f);};
  const rows=sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
  rows.forEach(function(r,n){
    if(String(r[i('employee_id')]).trim().toUpperCase()!==eid) return;
    out[String(r[i('doc_no')]).trim()] = {row:n+2,
      status:String(r[i('status')]||'').trim(),
      received_at:fmt_(r[i('received_at')]),
      notes:String(r[i('notes')]||'').trim()};
  });
  return out;
}

function isNonEmployee_(companyType){
  const t=String(companyType||'').toLowerCase();
  if(!t) return false;
  return NON_EMPLOYEE_MARKERS.some(function(m){ return t.indexOf(m)!==-1; });
}

function monthsOfService_(hireDate){
  const h=String(hireDate||'').trim(); if(!h) return 0;
  const d=new Date(h); if(isNaN(d.getTime())) return 0;
  const now=new Date();
  return Math.max(0, (now.getFullYear()-d.getFullYear())*12 + (now.getMonth()-d.getMonth()));
}

function whoFor_(f){
  return {gender: String(f.gender||'').trim().toLowerCase(),
          egyptian: String(f.nationality||'').trim().toLowerCase().indexOf('egypt')!==-1};
}

function getFileReview(eid){
  if(!isHR_()) throw new Error('HR only.');
  eid=String(eid||'').trim().toUpperCase();
  if(!eid) return {found:false, msg:'Enter an employee ID.'};

  const f=employeeFieldsOf_(eid,['employee_id','full_name_en','hire_date','job_title',
    'company_type','contract_type','gender','nationality','record_status']);
  if(!f.employee_id) return {found:false, msg:'No employee found with ID '+eid+'.'};

  const nonEmp=isNonEmployee_(f.company_type);
  const months=monthsOfService_(f.hire_date);
  const rec=fileStatusFor_(eid);
  const held=docsFor_(eid);

  const docs=fileChecklist_(whoFor_(f)).map(function(d){
    const h=held[d.no];
    return {no:d.no, name:d.name, optional:d.optional, note:d.note,
            status: (h && h.status==='Received') ? 'Received' : ''};
  });

  // the screen reads conversion_note and the signed flags; map from your columns
  const review = rec ? {
    contract_signed:      rec.contract_signed,
    contract_signed_date: rec.contract_signed_date,
    conversion_note:      rec.conversion_note,
    renewal_signed:       rec.renewal_signed,
    renewal_signed_date:  rec.renewal_signed_date,
    nda_signed:           rec.nda_signed,
    nda_signed_date:      rec.nda_signed_date,
    vendor_confirmed:     rec.vendor_confirmed,
    notes:                rec.notes
  } : {};

  const lists=getLists();
  return {found:true, employee_id:f.employee_id, name:f.full_name_en,
    hire_date:f.hire_date, months_service:months, job_title:f.job_title,
    company_type:f.company_type, contract_type:f.contract_type,
    non_employee:nonEmp, contract_types: lists['contract_type'] || [],
    renewal_applicable: !nonEmp && months>=12,
    documents:docs, review:review};
}

function saveFileReview(p){
  if(!isHR_()) throw new Error('HR only.');
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const eid=String(p && p.employee_id || '').trim().toUpperCase();
    if(!eid) return {ok:false, msg:'No employee ID.'};

    const f=employeeFieldsOf_(eid,['employee_id','full_name_en','national_id','hire_date',
      'company_type','contract_type','gender','nationality']);
    if(!f.employee_id) return {ok:false, msg:'No employee found.'};

    const checklist=fileChecklist_(whoFor_(f));
    const ticked=p.docs||{};
    const now=new Date(), me=currentUser_();

    // ---- EMPLOYEE_DOCS: one row per document, updated in place or appended
    const dsh=sheet_(TAB_EMPLOYEE_DOCS);
    const dhdr=docsHdr_(), di=function(x){return dhdr.indexOf(x);};
    const held=docsFor_(eid);
    const appends=[];
    let receivedCount=0;

    checklist.forEach(function(d){
      const isIn = !!ticked[d.no];
      if(isIn) receivedCount++;
      const status = isIn? 'Received' : 'Missing';
      const h=held[d.no];
      if(h){
        if(h.status===status) return;                 // unchanged
        dsh.getRange(h.row, di('status')+1).setValue(status);
        dsh.getRange(h.row, di('received_at')+1).setValue(isIn? now : '');
        dsh.getRange(h.row, di('received_by')+1).setValue(isIn? me : '');
      } else {
        const row=new Array(dhdr.length).fill('');
        row[di('employee_id')]=eid;
        row[di('employee_name')]=f.full_name_en;
        row[di('doc_no')]=d.no;
        row[di('doc_name')]=d.name;
        row[di('status')]=status;
        if(isIn){ row[di('received_at')]=now; row[di('received_by')]=me; }
        appends.push(row);
      }
    });
    if(appends.length){
      dsh.getRange(dsh.getLastRow()+1,1,appends.length,dhdr.length).setValues(appends);
    }

    const outstanding=checklist.filter(function(d){ return d.mandatory && !ticked[d.no]; }).length;

    // ---- FILE_STATUS: the roll-up row
    const sh=sheet_(TAB_FILE_STATUS);
    const hdr=fileStatusHdr_();
    const existing=fileStatusFor_(eid);
    const row=existing? existing.row : sh.getLastRow()+1;
    const set=function(field,v){ const c=hdr.indexOf(field); if(c!==-1) sh.getRange(row,c+1).setValue(v); };

    const months=monthsOfService_(f.hire_date);
    const nonEmp=isNonEmployee_(f.company_type);
    const newType=String(p.contract_type||'').trim();
    const oldType=String(f.contract_type||'').trim();

    set('employee_id',eid);
    set('employee_name',f.full_name_en);
    set('hire_date',f.hire_date);
    set('months_service',months);
    set('company_type',f.company_type);
    set('contract_signed', p.contract_signed? 'Yes':'No');
    set('contract_signed_date', p.contract_signed_date||'');
    set('contract_type_at_signing', newType||oldType);
    set('nda_signed', p.nda_signed? 'Yes':'No');
    set('nda_signed_date', p.nda_signed_date||'');
    set('vendor_confirmed', p.vendor_confirmed? 'Yes':'No');
    if(p.vendor_confirmed) set('vendor_confirmed_at', now);
    set('renewal_applicable', (!nonEmp && months>=12)? 'Yes':'No');
    set('renewal_signed', p.renewal_signed||'');
    set('renewal_signed_date', p.renewal_signed_date||'');
    set('docs_complete', receivedCount);
    set('docs_outstanding', outstanding);
    set('last_reviewed_at', now);
    set('last_reviewed_by', me);
    set('notes', p.notes||'');

    // ---- conversion: recorded here and written back to the master record
    let converted='';
    if(newType && newType!==oldType){
      set('converted_from', oldType);
      set('converted_to', newType);
      set('converted_date', now);
      set('conversion_note', p.conversion_note||'');
      const esh=sheet_(TAB.EMP), ehdr=headers_(TAB.EMP);
      const data=esh.getDataRange().getValues();
      const ei=ehdr.indexOf('employee_id'), ci=ehdr.indexOf('contract_type');
      for(let r=1;r<data.length;r++){
        if(String(data[r][ei]).trim().toUpperCase()!==eid) continue;
        esh.getRange(r+1,ci+1).setValue(newType);
        stampUpdate_(esh,ehdr,r+1);
        logChange_(eid, f.national_id, 'contract_type', oldType, newType,
                   'File review','Applied', p.conversion_note||'Changed during file review');
        converted=' Contract type updated to '+newType+'.';
        break;
      }
      clearEmpCache_();
    } else if(p.conversion_note){
      set('conversion_note', p.conversion_note);
    }

    return {ok:true, msg:'Saved — '+receivedCount+' document(s) on file'+
      (outstanding? (', '+outstanding+' still outstanding.') : ', file complete.')+converted};
  } finally { lock.releaseLock(); }
}

function fileReviewProgress(){
  if(!isHR_()) throw new Error('HR only.');
  const E=empData_(false);
  const total=E.rows.length;
  const sh=sheet_(TAB_FILE_STATUS);
  if(!sh || sh.getLastRow()<2) return {reviewed:0, total:total, signed:0, docsComplete:0, unsigned:0};

  const hdr=fileStatusHdr_(), i=function(f){return hdr.indexOf(f);};
  const rows=sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
  let reviewed=0, signed=0, unsigned=0, docsComplete=0;
  rows.forEach(function(r){
    if(!String(r[i('employee_id')]).trim()) return;
    reviewed++;
    const cs=String(r[i('contract_signed')]).trim();
    if(cs==='Yes') signed++; else if(cs==='No') unsigned++;
    // complete means nothing mandatory outstanding FOR THAT PERSON
    if(Number(r[i('docs_outstanding')])===0) docsComplete++;
  });
  return {reviewed:reviewed, total:total, signed:signed, docsComplete:docsComplete, unsigned:unsigned};
}

// ================================================================
// MEDICAL — status readout for a single employee
// ================================================================
function medicalStatusOf(eid){
  if(!isHR_()) throw new Error('HR only.');
  eid=String(eid||'').trim().toUpperCase();
  if(!eid) return {found:false, msg:'Enter an employee ID.'};
  const p=medicalPayloadFor_(eid);
  if(!p.emp.employee_id) return {found:false, msg:'No employee found with ID '+eid+'.'};
  const m=medicalRecordFor_(eid);
  return {found:true, employee_id:p.emp.employee_id, name:p.emp.full_name_en,
    status: m? m.status : 'Not enrolled',
    enrolled_at: m? m.enrolled_at : '',
    removed_at: m? m.removed_at : '',
    insurance_id: m? (m.insurance_id||'') : '',
    dependants: p.dependants,
    missing_ids: p.dependants.filter(function(d){ return !d.national_id; }).length};
}

// ================================================================
// LEAVE ON BEHALF
//   A manager (or HR) files a request for someone who cannot file it
//   themselves — most often sick leave, where the person is genuinely ill
//   and not logging in.
//
//   SEPARATION OF DUTIES
//   Entitled leave goes to HR for document validation, so a manager
//   submitting it approves nothing. No conflict.
//   Discretionary leave goes to managers, so submitting it IS the
//   submitter's approval — recorded as such, in their name, and passed
//   to the dotted manager if there is one. Nobody silently approves
//   their own submission.
//
//   The employee is emailed immediately either way. A leave record in
//   someone's name that they never saw is the thing to avoid.
// ================================================================

// Everything getMyLeaveInfo returns, but for any employee.
function leaveInfoForEmployee_(eid){
  eid=String(eid||'').trim().toUpperCase();
  const sh=sheet_(TAB.EMP), hdr=headers_(TAB.EMP), data=sh.getDataRange().getValues();
  const ei=hdr.indexOf('employee_id');
  let row=-1;
  for(let r=1;r<data.length;r++){
    if(String(data[r][ei]).trim().toUpperCase()===eid){ row=r; break; }
  }
  if(row===-1) return {found:false};
  const get=function(f){ const c=hdr.indexOf(f); return c===-1?'':fmt_(data[row][c]); };
  const hire=get('hire_date');
  const ent=entitlementFor_({leave_entitlement:get('leave_entitlement'),
    has_disability:get('has_disability'), date_of_birth:get('date_of_birth'), hire_date:hire});
  const year=new Date().getFullYear();
  const adj=adjustmentsFor_(eid, year);

  let annualTaken=0;
  const ls=sheet_(TAB_LEAVE);
  if(ls && ls.getLastRow()>1){
    const lh=ls.getRange(1,1,1,ls.getLastColumn()).getValues()[0];
    const li=function(f){return lh.indexOf(f);};
    ls.getRange(2,1,ls.getLastRow()-1,ls.getLastColumn()).getValues().forEach(function(r){
      if(String(r[li('employee_id')]).trim().toUpperCase()!==eid) return;
      if(String(r[li('leave_type')])!=='Annual') return;
      const fs=String(r[li('final_status')]||'');
      if(fs.indexOf('Approved')!==0 && fs!=='Auto-approved') return;
      const d=fmt_(r[li('start_date')]);
      if(d && new Date(d).getFullYear()!==year) return;
      annualTaken += parseFloat(r[li('days_approved')]||r[li('days_requested')]||0)||0;
    });
  }
  return {found:true, employee_id:get('employee_id'), name:get('full_name_en'),
    hire_date:hire, record_status:get('record_status'),
    weekend_pattern:get('weekend_pattern')||'Fri & Sat',
    konecta_email:get('konecta_email'),
    direct_manager:get('direct_manager'), dotted_manager:get('dotted_manager'),
    annual_entitlement:ent.days + (adj.total>0?adj.total:0),
    annual_remaining:Math.max(ent.days + adj.total - annualTaken, 0)};
}

// Look up one person before filing for them, so the manager sees who they picked.
function lookupEmployeeForLeave(employeeId){
  const info=leaveInfoForEmployee_(employeeId);
  if(!info.found) return {found:false, msg:'No employee found with ID '+String(employeeId).trim().toUpperCase()+'.'};
  const acting=actingFor_();
  const dm=resolveApprover_(info.direct_manager), dt=resolveApprover_(info.dotted_manager);
  const mine = acting.indexOf(dm)!==-1 || (dt && acting.indexOf(dt)!==-1);
  if(!mine && !isHR_()) return {found:false, msg:'You can only file leave for someone who reports to you.'};
  return {found:true, employee_id:info.employee_id, name:info.name,
          weekend_pattern:info.weekend_pattern, status:info.record_status,
          annual_remaining:info.annual_remaining, types:getLeaveTypes()};
}

function submitLeaveRequestFor(p){
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const eid=String(p.employee_id||'').trim().toUpperCase();
    if(!eid) return {ok:false,msg:'Enter the employee ID.'};
    const info=leaveInfoForEmployee_(eid);
    if(!info.found) return {ok:false,msg:'No employee found with ID '+eid+'.'};
    if(['Closed','Cleared'].indexOf(String(info.record_status))!==-1)
      return {ok:false,msg:'That record is closed.'};

    // --- who is filing, and are they allowed to?
    const me=currentUser_();
    const hr=isHR_();
    const acting=actingFor_();
    const dm=resolveApprover_(info.direct_manager);
    const dt=resolveApprover_(info.dotted_manager);
    const isDirect = acting.indexOf(dm)!==-1;
    const isDotted = !!dt && acting.indexOf(dt)!==-1;
    if(!hr && !isDirect && !isDotted)
      return {ok:false,msg:'You can only file leave for someone who reports to you.'};

    const t=getLeaveTypes().filter(function(x){return x.type===p.leave_type;})[0];
    if(!t) return {ok:false,msg:'Please choose a leave type.'};

    const c=countLeaveDays(p.start_date,p.end_date,info.weekend_pattern,p.worked_days);
    if(c.error) return {ok:false,msg:c.error};
    if(c.days<=0) return {ok:false,msg:'That range contains no working days — it falls on weekends or public holidays.'};

    // --- dating. Filing on behalf exists precisely because the absence has
    // already happened, so entitled leave may be backdated freely. Annual
    // still may not: nobody takes annual leave without telling anyone.
    const start=new Date(p.start_date);
    const today0=new Date(); today0.setHours(0,0,0,0);
    if(start<today0 && t.notice>0){
      return {ok:false,msg:t.type+' leave cannot be backdated. If this needs correcting, contact HR.'};
    }

    if(p.leave_type==='Annual' && c.days>info.annual_remaining){
      return {ok:false,msg:info.name+' has '+info.annual_remaining+' day(s) of annual leave left, but you have requested '+c.days+'.'};
    }

    const sh=sheet_(TAB_LEAVE), hdr=headers_(TAB_LEAVE);
    const row=sh.getLastRow()+1;
    const id='LV-'+String(row-1).padStart(6,'0');
    const set=function(f,v){ const i=hdr.indexOf(f); if(i!==-1) sh.getRange(row,i+1).setValue(v); };
    const dates=(c.detail||[]).filter(function(d){return d.counted;}).map(function(d){return d.date;});
    const stamp=new Date();

    set('request_id',id); set('submitted_at',stamp);
    set('employee_id',info.employee_id); set('employee_name',info.name);
    set('konecta_email',info.konecta_email);
    set('leave_type',t.type); set('track',t.track);
    set('start_date',p.start_date); set('end_date',p.end_date);
    set('days_requested',c.days); set('reason',String(p.reason||'').trim());
    set('direct_manager',info.direct_manager); set('dotted_manager',info.dotted_manager);
    set('submitted_on_behalf','Yes'); set('submitted_by',me);
    set('reminder_count',0);

    let routing='';
    if(t.track==='Entitled'){
      // straight to HR. The submitter approves nothing, so no conflict arises.
      set('hr_status','Pending'); set('document_received','No');
      set('final_status','Pending');
      routing='It is with HR, who will validate it once they have the '+
              (t.document? t.document.toLowerCase() : 'supporting document')+'.';
    } else if(hr && !isDirect && !isDotted){
      // HR filing for someone who is not their report: managers still decide
      set('direct_status','Pending'); if(dt) set('dotted_status','Pending');
      set('final_status','Pending');
      routing='It has gone to their manager for approval.';
    } else {
      // A manager filed it. Submitting IS their approval — recorded in their
      // name so the audit trail shows who allowed it, not a blank.
      const role = isDirect ? 'direct' : 'dotted';
      set(role+'_status','Approved');
      set(role+'_days', c.days);
      set(role+'_by', me);
      set(role+'_at', stamp);
      set('approved_dates', dates.join(','));
      set('notes','Filed by '+me+' on behalf of the employee. Submitting it counts as the '+
                  role+' manager\'s approval.');
      if(role==='direct' && dt){
        set('dotted_status','Pending'); set('final_status','Pending');
        routing='Your approval is recorded. It has gone to the dotted manager.';
      } else {
        set('days_approved', c.days); set('days_rejected', 0);
        set('final_status','Approved');
        routing='Your approval is recorded with it, so it is approved.';
        try{ createLeaveCalendarBlock_(row, headers_(TAB_LEAVE)); }catch(e){}
        try{ flagUnpaidForPayroll_(row, headers_(TAB_LEAVE)); }catch(e){}
      }
    }

    // --- the employee is told. A record in their name they never saw is the risk.
    if(info.konecta_email){
      try{
        MailApp.sendEmail({
          to: info.konecta_email,
          cc: HR_ADMINS.join(','),
          subject:'Leave recorded for you — '+id,
          htmlBody:
            '<div style="font-family:Arial,sans-serif;max-width:520px">'+
            '<p>Hello '+escapeHtml_(info.name)+',</p>'+
            '<p><strong>'+escapeHtml_(me)+'</strong> has recorded a leave request for you.</p>'+
            '<div style="background:#EEEDFE;border-radius:8px;padding:14px 18px;margin:14px 0">'+
            escapeHtml_(t.type)+' — '+c.days+' day(s)<br>'+
            escapeHtml_(p.start_date)+' to '+escapeHtml_(p.end_date)+'<br>'+
            'Reference '+id+'</div>'+
            '<div style="background:#FFF9D6;border-left:4px solid #FFE100;padding:12px 16px;margin:14px 0">'+
            '<strong>If this is wrong, contact HR straight away.</strong></div>'+
            (t.track==='Entitled' && t.document
              ? '<p>Please reply to this email with your '+escapeHtml_(t.document.toLowerCase())+
                ' attached. Your leave is not confirmed until HR has seen it.</p>'
              : '')+
            '<p style="font-size:13px;color:#6b6b80">You can see this in the Leave tab of the app.</p>'+
            '<p style="font-size:13px;color:#6b6b80">Konecta Egypt — People team</p></div>',
          replyTo: HR_ADMINS.join(','),
          name:'Konecta Egypt — People Team'
        });
      }catch(e){}
    }

    notifyHR_('Leave filed on behalf — '+id,
      info.name+' ('+info.employee_id+') — '+c.days+' day(s) of '+t.type+'\n'+
      p.start_date+' to '+p.end_date+'\n'+
      'Filed by: '+me+(hr?' (HR)':' (manager)')+'\n'+
      (p.reason? ('Reason: '+p.reason+'\n') : '')+
      '\nThe employee has been emailed to confirm.'+
      (t.document? ('\n\nDOCUMENT REQUIRED: '+t.document) : ''));

    return {ok:true, id:id, days:c.days,
      msg:'Recorded as '+id+' — '+c.days+' day(s) for '+info.name+'. '+routing+
          ' They have been emailed to confirm.'};
  } finally { lock.releaseLock(); }
}

// ================================================================
// BULK FIELD UPDATE
//   Paste "employee ID, new value" one per line. PREVIEW first —
//   current value against new value, with anything unrecognised flagged —
//   then apply. Nothing is written until you have seen what will change.
//
//   Only whitelisted fields can be touched. Every change is logged
//   individually to CHANGE LOG, exactly as if it had been edited by hand.
// ================================================================

// field -> how to validate it, and what to call it on screen
const BULK_FIELDS = {
  'direct_manager':    {label:'Direct manager',    kind:'manager'},
  'dotted_manager':    {label:'Dotted manager',    kind:'manager'},
  'weekend_pattern':   {label:'Weekend pattern',   kind:'weekend'},
  'project':           {label:'Project',           kind:'project'},
  'cost_centre':       {label:'Cost centre',       kind:'text'},
  'work_location':     {label:'Work location',     kind:'text'},
  'contract_end_date': {label:'Contract end date', kind:'date'},
  'basic_salary':      {label:'Basic salary (EGP)',kind:'number'},
  'kpi_target':        {label:'KPI target (monthly EGP)',kind:'number'},
  'kpi_frequency':     {label:'Bonus plan',       kind:'text'},
  'hire_date':         {label:'Hire date',        kind:'date'},
  'job_title':         {label:'Job title',        kind:'text'},
  'department':        {label:'Department',       kind:'department'},
  'bank_name':         {label:'Bank name',        kind:'text'},
  'account_number':    {label:'Account number',   kind:'text'},
  'iban':              {label:'IBAN',             kind:'iban'}
};

function bulkFieldOptions(){
  if(!isHR_()) throw new Error('HR only.');
  return Object.keys(BULK_FIELDS).map(function(f){
    return {field:f, label:BULK_FIELDS[f].label};
  });
}

// Parse the pasted block into {eid, value} pairs. Accepts comma or tab.
function parseBulkLines_(text){
  return String(text||'').split('\n')
    .map(function(l){return l.trim();})
    .filter(String)
    .map(function(l,n){
      const parts=l.split(/[,\t]/).map(function(x){return x.trim();});
      return {line:n+1, eid:String(parts[0]||'').toUpperCase(),
              value: parts.slice(1).join(', ').trim()};
    });
}

// Work out what would happen, without writing anything.
function hrBulkUpdatePreview(field, text){
  if(!isHR_()) throw new Error('HR only.');
  const spec=BULK_FIELDS[field];
  if(!spec) return {ok:false, msg:'That field cannot be bulk updated.'};

  const rows=parseBulkLines_(text);
  if(!rows.length) return {ok:false, msg:'Paste at least one line.'};

  const E=empData_(false), h=E.hdr;
  const cEid=h.indexOf('employee_id'), cNm=h.indexOf('full_name_en'), cF=h.indexOf(field);
  if(cF===-1) return {ok:false, msg:'Column '+field+' not found on the EMPLOYEES tab.'};

  // index the sheet once
  const byId={};
  E.rows.forEach(function(rec){
    byId[String(rec.values[cEid]).trim().toUpperCase()]=
      {row:rec.row, name:fmt_(rec.values[cNm]), current:fmt_(rec.values[cF])};
  });

  // things we validate against
  const validIds = spec.kind==='manager' ? validEmployeeIds_() : null;
  const globals={};
  if(spec.kind==='manager'){
    const ms=sheet_('MANAGERS');
    if(ms && ms.getLastRow()>4){
      ms.getRange(5,1,ms.getLastRow()-4,1).getValues().forEach(function(r){
        if(r[0]) globals[String(r[0]).trim().toUpperCase()]=true;
      });
    }
  }
  const projects = spec.kind==='project' ? getProjectMap() : null;

  const out=[]; let okCount=0;
  rows.forEach(function(r){
    const emp=byId[r.eid];
    let status='ok', note='';
    if(!emp){ status='error'; note='No active employee with this ID'; }
    else if(!r.value){ status='error'; note='No value given on this line'; }
    else {
      const v=r.value;
      if(spec.kind==='manager'){
        const up=v.toUpperCase();
        if(!validIds[up] && !globals[up]){ status='error'; note='Not a known employee or global manager'; }
        else if(up===r.eid){ status='error'; note='Cannot report to themselves'; }
      } else if(spec.kind==='weekend'){
        if(!WEEKEND_PATTERNS.hasOwnProperty(v)){
          status='error'; note='Must be one of: '+Object.keys(WEEKEND_PATTERNS).join(', ');
        }
      } else if(spec.kind==='project'){
        if(!projects.hasOwnProperty(v)){ status='error'; note='Not in the PROJECT_MAP tab'; }
      } else if(spec.kind==='date'){
        if(isNaN(new Date(v))){ status='error'; note='Not a valid date — use yyyy-MM-dd'; }
            } else if(spec.kind==='department'){
        if(!departmentNames_().hasOwnProperty(v)){
          status='error'; note='Not in the DEPARTMENTS tab';
        }
      } else if(spec.kind==='iban'){
        const ibanV=String(v).replace(/\s/g,'').toUpperCase();
        if(!/^EG\d{27}$/.test(ibanV)){ status='error'; note='Must be EG followed by 27 digits'; }
      } else if(spec.kind==='number'){
        if(isNaN(parseFloat(String(v).replace(/,/g,'')))){ status='error'; note='Not a number'; }
      }
      if(status==='ok' && String(emp.current).trim()===String(v).trim()){
        status='same'; note='Already set to this';
      }
    }
    if(status==='ok') okCount++;
    out.push({line:r.line, employee_id:r.eid, name:emp? emp.name : '',
              current: emp? emp.current : '', proposed:r.value,
              status:status, note:note});
  });

  return {ok:true, field:field, label:spec.label, rows:out,
          willChange:okCount, total:rows.length,
          errors: out.filter(function(x){return x.status==='error';}).length,
          unchanged: out.filter(function(x){return x.status==='same';}).length};
}

// Apply only the lines the preview marked ok. Re-validated here, so a stale
// preview cannot push through something that has since become invalid.
function hrBulkUpdateApply(field, text){
  if(!isHR_()) throw new Error('HR only.');
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const pre=hrBulkUpdatePreview(field, text);
    if(!pre.ok) return pre;
    const good=pre.rows.filter(function(r){return r.status==='ok';});
    if(!good.length) return {ok:false, msg:'Nothing to apply — every line is either invalid or already set.'};

    const sh=sheet_(TAB.EMP), hdr=headers_(TAB.EMP);
    const data=sh.getDataRange().getValues();
    const ei=hdr.indexOf('employee_id'), cF=hdr.indexOf(field), cNid=hdr.indexOf('national_id');
    const rowOf={};
    for(let r=1;r<data.length;r++){
      const id=String(data[r][ei]).trim().toUpperCase();
      if(id) rowOf[id]={row:r+1, nid:fmt_(data[r][cNid])};
    }

    const projects = BULK_FIELDS[field].kind==='project' ? getProjectMap() : null;
    const cCost = hdr.indexOf('cost_centre');
    let applied=0, ccFilled=0;

    good.forEach(function(r){
      const t=rowOf[r.employee_id]; if(!t) return;
      sh.getRange(t.row, cF+1).setValue(r.proposed);
      // moving someone to a project moves their cost centre with them,
      // otherwise the two drift apart silently
      if(projects && cCost!==-1 && projects[r.proposed]!==undefined){
        const oldCc=fmt_(sh.getRange(t.row,cCost+1).getValue());
        if(oldCc!==projects[r.proposed]){
          sh.getRange(t.row,cCost+1).setValue(projects[r.proposed]);
          logChange_(r.employee_id, t.nid, 'cost_centre', oldCc, projects[r.proposed],
                     'Bulk update','Applied','Followed the project change');
          ccFilled++;
        }
      }
      stampUpdate_(sh, hdr, t.row);
      logChange_(r.employee_id, t.nid, field, r.current, r.proposed,
                 'Bulk update','Applied','Bulk field update by '+currentUser_());
      applied++;
    });
    clearEmpCache_();

    // reporting depth depends on the manager chain, so it has to be rebuilt
    let nlevel='';
    if(field==='direct_manager'){
      try{ const n=recalcNLevels(); nlevel=' N-levels recalculated ('+n.written+' records, '+n.unresolved+' unresolved).'; }
      catch(e){ nlevel=' N-level recalculation failed: '+e.message; }
    }

    notifyHR_('Bulk update applied — '+BULK_FIELDS[field].label,
      applied+' record(s) updated by '+currentUser_()+'.\n\n'+
      'Field: '+field+'\n'+
      'Skipped: '+pre.errors+' invalid, '+pre.unchanged+' already set.\n\n'+
      'Every change is in the CHANGE LOG.');

    return {ok:true, applied:applied,
      msg:applied+' record(s) updated.'+
          (ccFilled? ' '+ccFilled+' cost centre(s) followed the project.' : '')+
          (pre.errors? ' '+pre.errors+' line(s) skipped as invalid.' : '')+
          (pre.unchanged? ' '+pre.unchanged+' already had this value.' : '')+
          nlevel};
  } finally { lock.releaseLock(); }
}

// ================================================================
// CAREER HISTORY
//   One person's movement, pulled out of the change log. Only the fields
//   that constitute a move — not address changes and phone numbers.
// ================================================================

const CAREER_FIELDS = {
  'job_title':'Job title', 'grade':'Job level', 'gcm':'GCM level',
  'basic_salary':'Basic salary', 'direct_manager':'Direct manager',
  'dotted_manager':'Dotted manager', 'contract_type':'Contract type',
  'project':'Project', 'function':'Function', 'record_status':'Status',
  'n_level':'N-level', 'work_location':'Work location'
};

function careerHistory(eid){
  if(!isHR_()) throw new Error('HR only.');
  eid=String(eid||'').trim().toUpperCase();
  if(!eid) return {found:false, msg:'Enter an employee ID.'};

  const f=employeeFieldsOf_(eid,['employee_id','full_name_en','job_title','grade',
    'basic_salary','hire_date','direct_manager']);
  if(!f.employee_id) return {found:false, msg:'No employee found with ID '+eid+'.'};

  const sh=sheet_(TAB.LOG);
  const items=[];
  if(sh && sh.getLastRow()>1){
    // written by position, so read by position
    const rows=sh.getRange(2,1,sh.getLastRow()-1,13).getValues();
    rows.forEach(function(r){
      if(String(r[2]).trim().toUpperCase()!==eid) return;
      const field=String(r[4]).trim();
      if(!CAREER_FIELDS.hasOwnProperty(field)) return;
      const oldV=fmt_(r[5]), newV=fmt_(r[6]);
      if(oldV===newV) return;
      items.push({when:fmt_(r[1]), field:field, label:CAREER_FIELDS[field],
                  from:oldV||'(blank)', to:newV||'(blank)',
                  by:String(r[7]||''), source:String(r[8]||''),
                  notes:String(r[12]||'')});
    });
  }
  items.sort(function(a,b){ return String(b.when).localeCompare(String(a.when)); });

  // changes on the same day are one move, not several
  const grouped=[]; let cur=null;
  items.forEach(function(it){
    const day=String(it.when).slice(0,10);
    if(!cur || cur.date!==day){ cur={date:day, changes:[]}; grouped.push(cur); }
    cur.changes.push(it);
  });

  // a promotion is a job title or grade move; salary alone is a review
  grouped.forEach(function(g){
    const fields=g.changes.map(function(c){return c.field;});
    g.is_promotion = fields.indexOf('job_title')!==-1 || fields.indexOf('grade')!==-1;
    g.is_pay_only  = !g.is_promotion && fields.indexOf('basic_salary')!==-1;
  });

  const lastMove = grouped.filter(function(g){return g.is_promotion;})[0];
  let monthsSinceMove='';
  if(lastMove){
    const d=new Date(lastMove.date);
    if(!isNaN(d)) monthsSinceMove=monthsOfService_(lastMove.date);
  }

  return {found:true, employee_id:f.employee_id, name:f.full_name_en,
    hire_date:f.hire_date, months_service:monthsOfService_(f.hire_date),
    current:{job_title:f.job_title, grade:f.grade, basic_salary:f.basic_salary,
             direct_manager:f.direct_manager},
    last_promotion: lastMove? lastMove.date : '',
    months_since_promotion: monthsSinceMove,
    events:grouped};
}

// Everyone who moved in a date range — the same data, across the company.
function hrMovementReport(fromDate, toDate){
  if(!isHR_()) throw new Error('HR only.');
  const from=String(fromDate||'').trim(), to=String(toDate||'').trim();
  const sh=sheet_(TAB.LOG);
  if(!sh || sh.getLastRow()<2) return {rows:[], total:0};
  const rows=sh.getRange(2,1,sh.getLastRow()-1,13).getValues();
  const byPerson={};
  rows.forEach(function(r){
    const field=String(r[4]).trim();
    if(['job_title','grade','basic_salary'].indexOf(field)===-1) return;
    const when=fmt_(r[1]).slice(0,10);
    if(from && when<from) return;
    if(to && when>to) return;
    const oldV=fmt_(r[5]), newV=fmt_(r[6]);
    if(oldV===newV) return;
    const eid=String(r[2]).trim().toUpperCase();
    if(!eid) return;
    if(!byPerson[eid]) byPerson[eid]={employee_id:eid, changes:[]};
    byPerson[eid].changes.push({when:when, field:field, from:oldV||'(blank)', to:newV||'(blank)'});
  });
  const out=Object.keys(byPerson).map(function(k){
    const p=byPerson[k];
    const f=employeeFieldsOf_(k,['full_name_en','job_title']);
    p.name=f.full_name_en; p.job_title=f.job_title;
    const fields=p.changes.map(function(c){return c.field;});
    p.promotion = fields.indexOf('job_title')!==-1 || fields.indexOf('grade')!==-1;
    p.changes.sort(function(a,b){ return a.when.localeCompare(b.when); });
    return p;
  });
  out.sort(function(a,b){ return (b.promotion?1:0)-(a.promotion?1:0); });
  return {rows:out, total:out.length,
          promotions: out.filter(function(x){return x.promotion;}).length,
          from:from, to:to};
}

// ================================================================
// DEPENDANTS — read and write, replacing the three slots on EMPLOYEES
//
//   No cap. Beyond the company-funded allowance a dependant is marked
//   Employee-paid and held for HR to confirm the premium, rather than
//   silently refused or silently covered.
//
//   An employee may add and edit their own, but only while a dependant is
//   still pending. Once the insurer has them, editing would put the record
//   and the policy out of step, so changes go to HR.
// ================================================================

const COMPANY_FUNDED_LIMIT = 3;     // beyond this, the employee pays
const DEP_STATUSES = ['Pending enrolment','Enrolled','Removed','Not eligible','Removal requested'];

// What the employee sees on their own record.
function getMyDependants(){
  const me=getMyRecord();
  if(!me.found) return {found:false};
  const eid=me.readonly.employee_id;
  const list=dependantsFor_(eid).filter(function(d){
    return String(d.status).trim()!=='Removed' && String(d.status).trim()!=='Not eligible';
  });
  return {found:true, employee_id:eid,
    limit: COMPANY_FUNDED_LIMIT,
    dependants: list.map(function(d){
      const st=String(d.status||'').trim();
      return {row:d.row, name:d.name, date_of_birth:d.date_of_birth,
              relation:d.relation, national_id:d.national_id,
              funding:d.funding, status:st,
              premium_amount:d.premium_amount,
              locked: st==='Enrolled' || st==='Removal requested',
              needs: certificateFor_(d.relation)};
    })};
}

// The employee saves their own. Pending rows are replaced wholesale;
// enrolled rows are left exactly as they are.
function saveMyDependants(list){
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const me=getMyRecord();
    if(!me.found) throw new Error('No record found for your account.');
    const eid=me.readonly.employee_id;
    const name=me.editable.full_name_en||'';

    const sh=sheet_(TAB_DEPENDANTS), hdr=depHdr_();
    const set=function(row,f,v){ const c=hdr.indexOf(f); if(c!==-1) sh.getRange(row,c+1).setValue(v); };

    const existing=dependantsFor_(eid);
    const locked=existing.filter(function(d){
      const s=String(d.status).trim();
      return s==='Enrolled' || s==='Removed' || s==='Not eligible' || s==='Removal requested';
    });
    const editable=existing.filter(function(d){ return String(d.status).trim()==='Pending enrolment'; });

    const incoming=(list||[]).filter(function(d){ return String(d.name||'').trim(); });

    // an employee cannot quietly drop someone the insurer is covering
    const lockedCount=locked.filter(function(d){
      return String(d.status).trim()==='Enrolled';
    }).length;

    // rewrite the pending rows in place, appending or blanking as needed
    let n=lockedCount;
    const added=[], changed=[];
    incoming.forEach(function(d,idx){
      n++;
      const funding = n<=COMPANY_FUNDED_LIMIT ? 'Company' : 'Employee-paid';
      const target = editable[idx];
      if(target){
        const before=target.name;
        set(target.row,'name',String(d.name).trim());
        set(target.row,'date_of_birth',String(d.date_of_birth||'').trim());
        set(target.row,'relation',String(d.relation||'').trim());
        set(target.row,'national_id',String(d.national_id||'').trim());
        set(target.row,'funding',funding);
        set(target.row,'dependant_no',n);
        if(before!==String(d.name).trim()) changed.push(d.name);
      } else {
        const row=sh.getLastRow()+1;
        set(row,'employee_id',eid); set(row,'employee_name',name);
        set(row,'dependant_no',n);
        set(row,'name',String(d.name).trim());
        set(row,'date_of_birth',String(d.date_of_birth||'').trim());
        set(row,'relation',String(d.relation||'').trim());
        set(row,'national_id',String(d.national_id||'').trim());
        set(row,'funding',funding);
        set(row,'status','Pending enrolment');
        set(row,'requested_at','');
        set(row,'notes','Added by the employee');
        added.push({name:String(d.name).trim(), relation:String(d.relation||'').trim(),
                    national_id:String(d.national_id||'').trim(),
                    date_of_birth:String(d.date_of_birth||'').trim()});
      }
    });

    // any pending row the employee removed from the list is cleared
    for(let i=incoming.length;i<editable.length;i++){
      set(editable[i].row,'status','Removed');
      set(editable[i].row,'notes','Removed by the employee before enrolment');
    }

    // the count on the employee record follows, so the two never disagree
    try{
      const esh=sheet_(TAB.EMP), ehdr=headers_(TAB.EMP);
      const c=ehdr.indexOf('dependants');
      if(c!==-1) esh.getRange(me.row,c+1).setValue(lockedCount+incoming.length);
      stampUpdate_(esh,ehdr,me.row);
    }catch(e){}

    if(added.length){
      logChange_(eid, me.readonly.national_id, 'dependants','',
                 added.map(function(a){return a.name;}).join('; '),
                 'Web app','Applied','Dependant(s) added by the employee');
      chaseDependantDocs_(eid, added.concat([]));
    }

    const paid=Math.max(lockedCount+incoming.length-COMPANY_FUNDED_LIMIT,0);
    return {ok:true, count:incoming.length, added:added.length,
      msg:'Saved.'+
        (added.length? ' We have emailed you asking for the documents we need.' : '')+
        (paid? ' '+paid+' dependant(s) are beyond the '+COMPANY_FUNDED_LIMIT+
               ' the company covers — HR will confirm the cost before they are added.' : '')};
  } finally { lock.releaseLock(); }
}

// An employee asking to remove someone the insurer already covers.
// Cover has to be stopped with the insurer, so this is a request, not an act.
function requestDependantRemoval(row, reason, key){
  const me=getMyRecord();
  if(!me.found) throw new Error('No record found for your account.');
  const sh=sheet_(TAB_DEPENDANTS), hdr=depHdr_();
  row=guardRow_(sh,hdr,row,{employee_id:me.readonly.employee_id, name:key});
  const g=function(f){ const c=hdr.indexOf(f); return c===-1?'':fmt_(sh.getRange(row,c+1).getValue()); };
  if(String(g('employee_id')).trim().toUpperCase()!==me.readonly.employee_id)
    throw new Error('That is not your dependant.');
  const set=function(f,v){ const c=hdr.indexOf(f); if(c!==-1) sh.getRange(row,c+1).setValue(v); };
  set('status','Removal requested');
  set('notes',(g('notes')? g('notes')+' | ':'')+'Removal requested by the employee: '+String(reason||''));
  notifyHR_('Dependant removal requested — '+me.readonly.employee_id,
    me.editable.full_name_en+' has asked to remove '+g('name')+' ('+g('relation')+') from the medical scheme.'+
    (reason? ('\n\nReason: '+reason) : '')+
    '\n\nStop cover with the insurer, then set the status to Removed.');
  return {ok:true, msg:'Sent to HR. Cover continues until they confirm it with the insurer.'};
}

// ---------- HR side ----------

function hrDependantsFor(eid){
  if(!isHR_()) throw new Error('HR only.');
  eid=String(eid||'').trim().toUpperCase();
  if(!eid) return {found:false, msg:'Enter an employee ID.'};
  const f=employeeFieldsOf_(eid,['employee_id','full_name_en','company_type','record_status']);
  if(!f.employee_id) return {found:false, msg:'No employee found with ID '+eid+'.'};
  const list=dependantsFor_(eid);
  const med=medicalRecordFor_(eid);
  return {found:true, employee_id:f.employee_id, name:f.full_name_en,
    company_type:f.company_type||'(not recorded)',
    employee_cover: med? med.status : 'Not enrolled',
    limit:COMPANY_FUNDED_LIMIT,
    statuses:DEP_STATUSES,
    dependants:list.map(function(d){
      return {row:d.row, dependant_no:d.dependant_no, name:d.name,
              date_of_birth:d.date_of_birth, relation:d.relation,
              national_id:d.national_id, funding:d.funding,
              premium_amount:d.premium_amount, status:d.status,
              insurance_id:d.insurance_id, requested_at:d.requested_at,
              notes:d.notes, needs:certificateFor_(d.relation)};
    })};
}

function hrSaveDependant(row, data, key){
  if(!isHR_()) throw new Error('HR only.');
  const sh=sheet_(TAB_DEPENDANTS), hdr=depHdr_();
  row=guardRow_(sh,hdr,row,{employee_id:(key||{}).employee_id, name:(key||{}).name});
  const g=function(f){ const c=hdr.indexOf(f); return c===-1?'':fmt_(sh.getRange(row,c+1).getValue()); };
  const set=function(f,v){ const c=hdr.indexOf(f); if(c!==-1) sh.getRange(row,c+1).setValue(v); };
  const eid=g('employee_id');
  if(!eid) return {ok:false, msg:'That row has no employee on it.'};

  const fields=['name','date_of_birth','relation','national_id','funding',
                'premium_amount','status','insurance_id','payroll_deduction','notes'];
  const changes=[];
  fields.forEach(function(f){
    if(data[f]===undefined) return;
    const oldV=g(f), newV=String(data[f]).trim();
    if(oldV===newV) return;
    set(f,newV); changes.push([f,oldV,newV]);
  });
  if(data.status==='Enrolled' && g('approved_by')==='') {
    set('approved_by',currentUser_()); set('approved_at',new Date());
  }
  changes.forEach(function(c){
    logChange_(eid,'','dependant_'+c[0],c[1],c[2],'HR console','Applied',
               'Dependant '+g('name')+' updated by HR');
  });
  return {ok:true, count:changes.length};
}

function hrAddDependant(eid, data){
  if(!isHR_()) throw new Error('HR only.');
  eid=String(eid||'').trim().toUpperCase();
  const f=employeeFieldsOf_(eid,['employee_id','full_name_en']);
  if(!f.employee_id) return {ok:false, msg:'No employee found with ID '+eid+'.'};
  if(!String(data.name||'').trim()) return {ok:false, msg:'Enter the dependant name.'};

  const existing=dependantsFor_(eid).filter(function(d){
    return ['Removed','Not eligible'].indexOf(String(d.status).trim())===-1;
  });
  const n=existing.length+1;
  const sh=sheet_(TAB_DEPENDANTS), hdr=depHdr_();
  const row=sh.getLastRow()+1;
  const set=function(fl,v){ const c=hdr.indexOf(fl); if(c!==-1) sh.getRange(row,c+1).setValue(v); };
  set('employee_id',eid); set('employee_name',f.full_name_en);
  set('dependant_no',n);
  set('name',String(data.name).trim());
  set('date_of_birth',String(data.date_of_birth||'').trim());
  set('relation',String(data.relation||'').trim());
  set('national_id',String(data.national_id||'').trim());
  set('funding', data.funding || (n<=COMPANY_FUNDED_LIMIT? 'Company':'Employee-paid'));
  set('premium_amount',String(data.premium_amount||'').trim());
  set('status', data.status || 'Pending enrolment');
  set('approved_by',currentUser_()); set('approved_at',new Date());
  set('notes',String(data.notes||'Added by HR').trim());
  logChange_(eid,'','dependants','',String(data.name).trim(),'HR console','Applied','Dependant added by HR');
  return {ok:true, msg:'Added as dependant '+n+
    (n>COMPANY_FUNDED_LIMIT? ' — beyond the '+COMPANY_FUNDED_LIMIT+' the company covers, so marked Employee-paid.' : '.')};
}
// ================================================================
// DEPENDANTS TAB — reader and document chase
//   One row per dependant, so there is no three-per-person cap and
//   each one carries its own status and funding.
//
//   The chase email is the same one that fires when an employee adds a
//   dependant themselves. This lets HR send it for dependants already on
//   record — the ones declared before the tab existed, who were never
//   asked for anything.
// ================================================================

const TAB_DEPENDANTS = 'DEPENDANTS';

function depHdr_(){
  const sh=sheet_(TAB_DEPENDANTS);
  if(!sh) throw new Error('DEPENDANTS tab not found.');
  return sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
}

// All dependants for one employee, optionally filtered by status.
function dependantsFor_(eid, status){
  const sh=sheet_(TAB_DEPENDANTS);
  if(!sh || sh.getLastRow()<2) return [];
  eid=String(eid||'').trim().toUpperCase();
  const hdr=depHdr_(), i=function(f){return hdr.indexOf(f);};
  const rows=sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
  const out=[];
  rows.forEach(function(r,n){
    if(String(r[i('employee_id')]).trim().toUpperCase()!==eid) return;
    if(status && String(r[i('status')]).trim()!==status) return;
    const o={row:n+2};
    hdr.forEach(function(h,c){ o[h]=fmt_(r[c]); });
    out.push(o);
  });
  return out;
}

// Everyone who has a dependant at a given status.
function dependantsByStatus_(status){
  const sh=sheet_(TAB_DEPENDANTS);
  if(!sh || sh.getLastRow()<2) return {};
  const hdr=depHdr_(), i=function(f){return hdr.indexOf(f);};
  const rows=sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
  const byEmp={};
  rows.forEach(function(r,n){
    if(String(r[i('status')]).trim()!==status) return;
    const eid=String(r[i('employee_id')]).trim().toUpperCase();
    if(!eid) return;
    const o={row:n+2};
    hdr.forEach(function(h,c){ o[h]=fmt_(r[c]); });
    (byEmp[eid]=byEmp[eid]||[]).push(o);
  });
  return byEmp;
}

// Which certificate a relation needs. Anything unrecognised asks for both,
// rather than guessing and sending the employee down the wrong path.
function certificateFor_(relation){
  const r=String(relation||'').trim().toLowerCase();
  if(r==='spouse' || r==='wife' || r==='husband') return 'marriage certificate';
  if(r==='child' || r==='son' || r==='daughter') return 'birth certificate';
  return 'marriage or birth certificate, whichever applies';
}

// The chase email for one employee's outstanding dependants.
function chaseDependantDocs_(eid, deps){
  eid=String(eid||'').trim().toUpperCase();
  if(!eid) return {ok:false, msg:'No employee ID — refusing to send.'};
  const f=employeeFieldsOf_(eid,['employee_id','full_name_en','konecta_email']);
  if(!f.konecta_email) return {ok:false, msg:'No Konecta email on record for '+eid+'.'};

  let rows='';
  deps.forEach(function(d){
    const cert=certificateFor_(d.relation);
    const needs=[cert];
    if(!String(d.national_id||'').trim()) needs.push('their national ID');
    rows+='<li><strong>'+escapeHtml_(d.name)+'</strong>'+
          (d.relation? ' ('+escapeHtml_(d.relation)+')' : '')+
          (d.date_of_birth? ' — born '+escapeHtml_(d.date_of_birth) : '')+
          '<br><span style="color:#6b6b80">We need: '+escapeHtml_(needs.join(', '))+'</span></li>';
  });

  try{
    MailApp.sendEmail({
      to: f.konecta_email,
      cc: [MEDICAL_CONTACT].concat(HR_ADMINS).join(','),
      replyTo: HR_ADMINS.join(','),
      subject:'Action needed: dependant documents — '+eid,
      htmlBody:
        '<div style="font-family:Arial,sans-serif;max-width:520px">'+
        '<p>Hello '+escapeHtml_(f.full_name_en||'')+',</p>'+
        '<p>You have dependants on your record who are <strong>not yet covered</strong> by the '+
        'medical insurance scheme. Before we can add them, we need a document for each.</p>'+
        '<ul>'+rows+'</ul>'+
        '<div style="background:#FFF9D6;border-left:4px solid #FFE100;padding:12px 16px;margin:16px 0">'+
        '<strong>Reply to this email with the documents attached.</strong><br>'+
        'Keep the subject line as it is — it carries your employee number so we can match them.</div>'+
        '<p style="font-size:13px;color:#6b6b80">Until these are received your dependants are not covered, '+
        'so please send them as soon as you can.</p>'+
        '<p style="font-size:13px;color:#6b6b80">Konecta Egypt — People team</p></div>',
      name:'Konecta Egypt — People Team'
    });
  }catch(e){ return {ok:false, msg:e.message}; }
  return {ok:true, count:deps.length};
}

// HR: chase one employee.
function hrChaseDependantDocs(eid){
  if(!isHR_()) throw new Error('HR only.');
  eid=String(eid||'').trim().toUpperCase();
  if(!eid) return {ok:false, msg:'Enter an employee ID.'};
  const deps=dependantsFor_(eid,'Pending enrolment');
  if(!deps.length) return {ok:false, msg:'No dependants awaiting enrolment for '+eid+'.'};
  const res=chaseDependantDocs_(eid, deps);
  if(!res.ok) return res;
  stampChased_(deps);
  return {ok:true, msg:'Asked for '+deps.length+' document set(s). '+MEDICAL_CONTACT+' and HR are copied.'};
}

// HR: chase everyone at once. Used after a migration, or periodically.
function hrChaseAllPendingDependants(){
  if(!isHR_()) throw new Error('HR only.');
  const byEmp=dependantsByStatus_('Pending enrolment');
  const ids=Object.keys(byEmp);
  if(!ids.length) return {ok:false, msg:'Nobody has dependants awaiting enrolment.'};
  const sent=[], failed=[];
  ids.forEach(function(eid){
    const res=chaseDependantDocs_(eid, byEmp[eid]);
    if(res.ok){ sent.push(eid); stampChased_(byEmp[eid]); }
    else failed.push(eid+': '+res.msg);
  });
  notifyHR_('Dependant documents requested — '+sent.length+' employee(s)',
    'Chase emails sent to:\n  '+sent.join('\n  ')+
    (failed.length? ('\n\nCould not send:\n  '+failed.join('\n  ')) : '')+
    '\n\nEach was asked for a national ID and the right certificate for every dependant '+
    'still awaiting enrolment. '+MEDICAL_CONTACT+' is copied so the documents land in one place.');
  return {ok:true, sent:sent.length, failed:failed,
    msg:'Asked '+sent.length+' employee(s) for their dependant documents.'+
        (failed.length? ' '+failed.length+' could not be emailed — see the summary sent to HR.' : '')};
}

// Record that we asked, and when, so nobody is chased twice in a week
// and so a stale request is visible.
function stampChased_(deps){
  const sh=sheet_(TAB_DEPENDANTS), hdr=depHdr_();
  const c=hdr.indexOf('requested_at');
  if(c===-1) return;
  const now=new Date();
  deps.forEach(function(d){ sh.getRange(d.row, c+1).setValue(now); });
}

// HR console: who is still outstanding, and how long since we asked.
function hrPendingDependants(){
  if(!isHR_()) throw new Error('HR only.');
  const byEmp=dependantsByStatus_('Pending enrolment');
  const out=[];
  Object.keys(byEmp).forEach(function(eid){
    const f=employeeFieldsOf_(eid,['full_name_en','company_type','record_status']);
    const deps=byEmp[eid];
    let asked='', days='';
    deps.forEach(function(d){ if(d.requested_at && d.requested_at>asked) asked=d.requested_at; });
    if(asked){
      const dt=new Date(asked);
      if(!isNaN(dt)) days=Math.floor((new Date()-dt)/86400000);
    }
    out.push({employee_id:eid, name:f.full_name_en,
      company_type:f.company_type||'(not recorded)',
      record_status:f.record_status,
      count:deps.length,
      missing_ids:deps.filter(function(d){return !String(d.national_id||'').trim();}).length,
      last_asked:asked, days_since_asked:days,
      dependants:deps.map(function(d){
        return {name:d.name, relation:d.relation, dob:d.date_of_birth,
                national_id:d.national_id, needs:certificateFor_(d.relation)};
      })});
  });
  out.sort(function(a,b){ return (b.days_since_asked||9999)-(a.days_since_asked||9999); });
  return out;
}

// ================================================================
// HISTORIC LEAVERS — bulk close records for people who left before
// the system existed.
//
//   Deliberately NOT the same as a resignation or termination. Those run
//   a process: clearance opens, equipment is recovered, medical cover is
//   stopped, people are emailed. None of that applies to someone who left
//   eighteen months ago — the laptop is long gone and the emails would be
//   sent to strangers.
//
//   So this writes the record and nothing else. It sets exit date, exit
//   type and status TOGETHER, because a date without a status leaves
//   someone showing as Active with a leaving date on them, which is worse
//   than either alone.
// ================================================================

const HISTORIC_EXIT_TYPES = ['Resignation','Dismissal','Probation not passed',
                             'Contract ended','No show','Drop out','Redundancy','Other'];

function historicExitTypes(){
  if(!isHR_()) throw new Error('HR only.');
  return HISTORIC_EXIT_TYPES;
}

// One line each: employee ID, last working day, exit type
function parseLeaverLines_(text){
  return String(text||'').split('\n')
    .map(function(l){return l.trim();})
    .filter(String)
    .map(function(l,n){
      const p=l.split(/[,\t]/).map(function(x){return x.trim();});
      return {line:n+1, eid:String(p[0]||'').toUpperCase(),
              date:String(p[1]||''), type:String(p[2]||'')};
    });
}

function hrBulkLeaversPreview(text){
  if(!isHR_()) throw new Error('HR only.');
  const rows=parseLeaverLines_(text);
  if(!rows.length) return {ok:false, msg:'Paste at least one line.'};

  const E=empData_(true), h=E.hdr;              // include everyone, even already-closed
  const cEid=h.indexOf('employee_id'), cNm=h.indexOf('full_name_en'),
        cSt=h.indexOf('record_status'), cEx=h.indexOf('exit_date'),
        cHire=h.indexOf('hire_date');
  const byId={};
  E.rows.forEach(function(rec){
    byId[String(rec.values[cEid]).trim().toUpperCase()]={
      row:rec.row, name:fmt_(rec.values[cNm]), status:fmt_(rec.values[cSt]),
      exit:fmt_(rec.values[cEx]), hire:fmt_(rec.values[cHire])};
  });

  const today=new Date(); today.setHours(0,0,0,0);
  const out=[]; let okCount=0;
  rows.forEach(function(r){
    const emp=byId[r.eid];
    let status='ok', note='';
    if(!emp){ status='error'; note='No employee with this ID'; }
    else if(!r.date){ status='error'; note='No last working day on this line'; }
    else if(isNaN(new Date(r.date))){ status='error'; note='Date not valid — use yyyy-MM-dd'; }
    else if(!r.type){ status='error'; note='No exit type on this line'; }
    else if(HISTORIC_EXIT_TYPES.indexOf(r.type)===-1){
      status='error'; note='Exit type must be one of: '+HISTORIC_EXIT_TYPES.join(', ');
    }
    else if(emp.exit){ status='same'; note='Already has an exit date of '+emp.exit; }
    else {
      const d=new Date(r.date);
      if(d>today){ status='error'; note='That date is in the future — this is for historic leavers only'; }
      else if(emp.hire && new Date(r.date)<new Date(emp.hire)){
        status='error'; note='Before their hire date of '+emp.hire;
      }
    }
    if(status==='ok') okCount++;
    out.push({line:r.line, employee_id:r.eid, name:emp? emp.name:'',
              current_status: emp? emp.status:'', hire_date: emp? emp.hire:'',
              exit_date:r.date, exit_type:r.type, status:status, note:note});
  });

  return {ok:true, rows:out, willChange:okCount, total:rows.length,
    errors: out.filter(function(x){return x.status==='error';}).length,
    unchanged: out.filter(function(x){return x.status==='same';}).length};
}

function hrBulkLeaversApply(text){
  if(!isHR_()) throw new Error('HR only.');
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const pre=hrBulkLeaversPreview(text);
    if(!pre.ok) return pre;
    const good=pre.rows.filter(function(r){return r.status==='ok';});
    if(!good.length) return {ok:false, msg:'Nothing to apply — every line is invalid or already closed.'};

    const sh=sheet_(TAB.EMP), hdr=headers_(TAB.EMP);
    const data=sh.getDataRange().getValues();
    const ei=hdr.indexOf('employee_id'), cNid=hdr.indexOf('national_id');
    const rowOf={};
    for(let r=1;r<data.length;r++){
      const id=String(data[r][ei]).trim().toUpperCase();
      if(id) rowOf[id]={row:r+1, nid:fmt_(data[r][cNid])};
    }

    let applied=0; const stillCovered=[];
    good.forEach(function(r){
      const t=rowOf[r.employee_id]; if(!t) return;
      const put=function(f,v){ const c=hdr.indexOf(f); if(c!==-1) sh.getRange(t.row,c+1).setValue(v); };
      put('exit_date', r.exit_date);
      put('last_working_day', r.exit_date);
      put('exit_type', r.exit_type);
      put('record_status','Closed');
      stampUpdate_(sh,hdr,t.row);
      logChange_(r.employee_id, t.nid, 'record_status', r.current_status, 'Closed',
                 'Historic leaver load','Applied',
                 'Left '+r.exit_date+' — '+r.exit_type+'. Loaded in bulk; no clearance was run.');
      applied++;

      // the thing that costs money quietly: cover still running for someone long gone
      try{
        const m=medicalRecordFor_(r.employee_id);
        if(m && m.status==='Enrolled') stillCovered.push(r.employee_id+'  '+r.name);
      }catch(e){}
    });
    clearEmpCache_();

    notifyHR_('Historic leavers loaded — '+applied+' record(s)',
      applied+' record(s) closed by '+currentUser_()+'.\n\n'+
      'Exit date, exit type and status were set. No clearance was opened and nobody was emailed, '+
      'because these people left before the system existed.\n\n'+
      (stillCovered.length
        ? ('STILL ON MEDICAL COVER — stop these with the insurer:\n  '+stillCovered.join('\n  ')+'\n\n')
        : 'None of them are still showing as enrolled on medical cover.\n\n')+
      'Every change is in the CHANGE LOG.');

    return {ok:true, applied:applied, stillCovered:stillCovered,
      msg:applied+' record(s) closed.'+
        (pre.errors? ' '+pre.errors+' line(s) skipped as invalid.' : '')+
        (pre.unchanged? ' '+pre.unchanged+' already had an exit date.' : '')+
        (stillCovered.length? ' '+stillCovered.length+' are STILL ON MEDICAL COVER — see the email.' : '')};
  } finally { lock.releaseLock(); }
}

// ================================================================
// SIK MONTHLY REPORT
//   59 columns, one row per person paid in the month.
//   Reads three tabs: PAYROLL for cost, EMPLOYEES for attributes,
//   LEAVE for absence. Writes a new tab you can download.
//
//   Who is in it: everyone in PAYROLL for that period who was employed
//   during the month. Someone who left the month before is out; someone
//   who joins the month after is out; someone who left DURING the month
//   is in, with their exit date and discharge code.
// ================================================================

const TAB_PAYROLL = 'PAYROLL';
const SIK_COUNTRY = 45;
const SIK_HOURS_PER_DAY = 8;

// Leave type -> which absence column it lands in. Matched on lowercase
// substring, so 'Sick leave' and 'Sick' both work. Anything not listed
// is ignored — annual leave is paid time off, not absence for this purpose.
const SIK_ABSENCE_MAP = [
  {col:'temporary_disability', match:['sick']},
  {col:'maternity',            match:['maternity','paternity']},
  {col:'others_paid',          match:['casual','emergency','compassion']},
  {col:'others_unpaid_just',   match:['unpaid']},
  {col:'others_unpaid_unjust', match:['no show','no-show','absence without']}
];

function sikAbsenceColumn_(leaveType){
  const t=String(leaveType||'').toLowerCase();
  for(let i=0;i<SIK_ABSENCE_MAP.length;i++){
    const m=SIK_ABSENCE_MAP[i];
    for(let j=0;j<m.match.length;j++){
      if(t.indexOf(m.match[j])!==-1) return m.col;
    }
  }
  return null;
}

function sikDate_(v){
  if(!v) return '';
  const d=(v instanceof Date)? v : new Date(v);
  if(isNaN(d)) return '';
  return Utilities.formatDate(d,Session.getScriptTimeZone(),'yyyyMMdd');
}
function sikNum_(v){
  if(v===''||v===null||v===undefined) return 0;
  const n=parseFloat(String(v).replace(/,/g,''));
  return isNaN(n)? 0 : n;
}

// Absence hours per employee for the month, split by SIK column.
function sikAbsence_(year, month){
  const out={};
  const sh=sheet_(TAB_LEAVE);
  if(!sh || sh.getLastRow()<2) return out;
  const hdr=leaveHdr_(), i=function(f){return hdr.indexOf(f);};
  const rows=sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
  const mStart=new Date(year, month-1, 1);
  const mEnd=new Date(year, month, 0);

  rows.forEach(function(r){
    const eid=String(r[i('employee_id')]).trim().toUpperCase();
    if(!eid) return;
    const st=String(r[i('final_status')]||'');
    if(st.indexOf('Approved')!==0 && st!=='Auto-approved' && st!=='Partially approved') return;
    const col=sikAbsenceColumn_(r[i('leave_type')]);
    if(!col) return;

    // count only the days that fall inside this month
    let days=0;
    const approved=String(r[i('approved_dates')]||'').split(',')
                     .map(function(x){return x.trim();}).filter(String);
    if(approved.length){
      approved.forEach(function(d){
        const dt=new Date(d);
        if(!isNaN(dt) && dt>=mStart && dt<=mEnd) days++;
      });
    } else {
      // no day list: fall back to the range, clipped to the month
      const s=new Date(fmt_(r[i('start_date')])), e=new Date(fmt_(r[i('end_date')]));
      if(isNaN(s)||isNaN(e)) return;
      const from=s>mStart? s : mStart, to=e<mEnd? e : mEnd;
      if(to<from) return;
      const span=Math.round((to-from)/86400000)+1;
      const total=Math.round((e-s)/86400000)+1;
      const approvedDays=sikNum_(r[i('days_approved')])||sikNum_(r[i('days_requested')]);
      days = total>0 ? (approvedDays*span/total) : 0;
    }
    if(days<=0) return;
    if(!out[eid]) out[eid]={temporary_disability:0,maternity:0,others_paid:0,
                            others_unpaid_just:0,others_unpaid_unjust:0};
    out[eid][col] += days*SIK_HOURS_PER_DAY;
  });
  return out;
}

const SIK_HEADERS = ['COUNTRY','YEAR','MONTH','EMPLOYEE_ID','BIRTH DATE','GENDER','CITIZENSHIP',
'DISABILITY','TRAINING','COMPANY','CORPORATION CODE','CONTRACT TYPE','CONTRACT TIME',
'AGREEMENT HOURS (CBA)','CONTRACT /EMPLOYEE HOURS','HIRING DATE','END DATE','DISCHARGE CODE',
'Temporary disability','Maternity / Paternity leave','Labor union hours -','Others Paid Abs',
'Suspensions -','Others Unpaid (Justified reasons)','Others Unpaid (Unjustified reasons)',
'WORK CENTRE','COST CENTRE','CATEGORY','TOTAL COST_INVOICE (TOTAL LABOR COST)','GROSS SALARY',
'VARIABLE / INCENTIVE REMUNERATION(Other local variable plans)','BONUS /SPORADIC PRIZES(MBO)',
'OVERTIME PAY','SOCIAL BENEFITS','DISMISSAL PAY','WAGE ARREARS','SOCIAL CONTRIBUTION',
'OTHERS COMPANY COSTS -','PROVISION','OTHERS COMPANY COSTS 1. -','OTHERS COMPANY COSTS 2. -',
'OTHERS COMPANY COSTS 3. -','OTHERS COMPANY COSTS 4. -','OTHERS COMPANY COSTS 5. -',
'LIQUID SALARY -','FUNCTION','SUBFUNCTION','WORK MODALITY','EMPLOYEE CLASSIFICATION','SCOPE',
'NAME','SURNAME','ANNUAL GROSS SALARY','ON TARGET ANNUAL VARIABLE','DIGITAL','BENCH','PROJECT',
'VARIABLE PLAN','HIRING REASON'];

function hrSikReport(year, month){
  if(!isHR_()) throw new Error('HR only.');
  year=parseInt(year); month=parseInt(month);
  if(!year||!month||month<1||month>12) return {ok:false,msg:'Choose a year and a month.'};
  const period=year+'-'+String(month).padStart(2,'0');

  // ---- payroll, if there is any for this month. Cost only; it does NOT
  // decide who is in the report. Subcontractors are paid on a vendor
  // invoice and never reach payroll, but they are still headcount.
  const pay={};
  const psh=sheet_(TAB_PAYROLL);
  let pi=null, phdr=null;
  if(psh && psh.getLastRow()>1){
    phdr=psh.getRange(1,1,1,psh.getLastColumn()).getValues()[0];
    pi=function(f){return phdr.indexOf(f);};
    if(pi('period')!==-1){
      psh.getRange(2,1,psh.getLastRow()-1,psh.getLastColumn()).getValues().forEach(function(r){
        if(String(r[pi('period')]).trim()!==period) return;
        const c=String(r[pi('Code')]).trim().toUpperCase();
        if(c) pay[c]=r;
      });
    }
  }

  const E=empData_(true), eh=E.hdr;
  const ec=function(f){return eh.indexOf(f);};
  const mStart=new Date(year, month-1, 1);
  const mEnd=new Date(year, month, 0);
  const absence=sikAbsence_(year, month);

  const out=[], noPayroll=[];
  E.rows.forEach(function(rec){
    const v=rec.values;
    const eid=String(v[ec('employee_id')]).trim().toUpperCase();
    if(!eid) return;
    const g=function(f){ const c=ec(f); return c===-1? '' : fmt_(v[c]); };

    // employed at any point during the month?
    const hire=g('hire_date')? new Date(g('hire_date')) : null;
    const exit=g('exit_date')? new Date(g('exit_date')) : null;
    if(!hire || isNaN(hire)) return;                         // never started
    if(hire>mEnd) return;                                    // joins later
    if(exit && !isNaN(exit) && exit<mStart) return;          // left earlier
    const leftThisMonth = exit && !isNaN(exit) && exit>=mStart && exit<=mEnd;

    const p=pay[eid]||null;
    const pv=function(f){ return (p && pi && pi(f)!==-1)? sikNum_(p[pi(f)]) : 0; };
    if(!p) noPayroll.push({employee_id:eid, name:g('full_name_en'),
                           company_type:g('company_type')||'(not recorded)'});

    const a=absence[eid]||{temporary_disability:0,maternity:0,others_paid:0,
                           others_unpaid_just:0,others_unpaid_unjust:0};
    const basic = p? pv('Worth Basic salary (EGP)') : sikNum_(g('basic_salary'));
    const kpiMonthly=sikNum_(g('kpi_target'));

    out.push([
      SIK_COUNTRY, year, month, eid,
      sikDate_(g('date_of_birth')), g('gender_code'), g('citizenship_code'),
      g('has_disability_code'), g('training_contract_code'), g('company_type_code'),
      g('corporation_code'), g('contract_type_code'), g('contract_time_code'),
      g('agreement_hours_cba'), g('contract_hours'),
      sikDate_(g('hire_date')),
      leftThisMonth? sikDate_(g('exit_date')) : '',
      leftThisMonth? g('exit_type_code') : '',
      a.temporary_disability, a.maternity, 0, a.others_paid, 0,
      a.others_unpaid_just, a.others_unpaid_unjust,
      g('work_location'), g('cost_centre'), g('job_title'),
      pv('Employee Total Cost'),
      p? basic : 0,
      pv('Incentive')+pv('KPI'),
      pv('Bonus (Annual)'),
      pv('Overtime Amount (total)'),
      pv('Employee Share'),
      pv('Lieu of Notice'),
      pv('Retro / Back-pay'),
      pv('Employer Share'),
      pv('ER Fund'),
      pv('End of Service'),
      '','','','','',
      pv('Net Salary'),
      g('function'), g('subfunction'), g('work_modality'),
      g('employee_classification'), g('scope'),
      g('report_name'), g('report_surname'),
      basic*12, kpiMonthly*12,
      g('digital_flag'), g('bench_flag'), g('project'),
      g('variable_plan'), g('hiring_reason')
    ]);
  });

  if(!out.length) return {ok:false,msg:'Nobody was employed during '+period+'.'};

  const name='SIK_'+year+'_'+String(month).padStart(2,'0');
  const ss=ss_();
  let sh=ss.getSheetByName(name);
  if(sh) ss.deleteSheet(sh);
  sh=ss.insertSheet(name);
  sh.getRange(1,1,1,SIK_HEADERS.length).setValues([SIK_HEADERS]);
  sh.getRange(1,1,1,SIK_HEADERS.length).setFontWeight('bold')
    .setBackground('#1F3864').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  sh.getRange(2,1,out.length,SIK_HEADERS.length).setValues(out);

  const totalCost=out.reduce(function(s,r){return s+(parseFloat(r[28])||0);},0);
  const leavers=out.filter(function(r){return r[16];}).length;

  return {ok:true, tab:name, rows:out.length, leavers:leavers,
    noPayroll:noPayroll, totalCost:Math.round(totalCost*100)/100,
    msg:out.length+' row(s) written to the '+name+' tab. '+
        leavers+' left during the month. Total labour cost '+
        Math.round(totalCost).toLocaleString('en-US')+' EGP.'+
        (noPayroll.length? ' '+noPayroll.length+' have no payroll row — their cost columns are zero and need the invoice value adding.' : '')};
}

// ================================================================
// SIK PREFLIGHT
//   Runs the same population as the report, but writes nothing.
//   Tells you what is missing, by field and by person, so the gaps get
//   fixed on the record rather than patched in the output file.
// ================================================================

// Every attribute the group file expects. required=true means the cell
// must not be blank; the rest are reported but will not stop a submission.
const SIK_REQUIRED = [
  {f:'date_of_birth',           label:'Birth date',              required:true},
  {f:'gender_code',             label:'Gender',                  required:true},
  {f:'citizenship_code',        label:'Citizenship',             required:true},
  {f:'has_disability_code',     label:'Disability',              required:true},
  {f:'training_contract_code',  label:'Training',                required:true},
  {f:'company_type_code',       label:'Company',                 required:true},
  {f:'corporation_code',        label:'Corporation code',        required:true},
  {f:'contract_type_code',      label:'Contract type',           required:true},
  {f:'contract_time_code',      label:'Contract time',           required:true},
  {f:'agreement_hours_cba',     label:'Agreement hours',         required:true},
  {f:'contract_hours',          label:'Contract hours',          required:true},
  {f:'hire_date',               label:'Hiring date',             required:true},
  {f:'work_location',           label:'Work centre',             required:true},
  {f:'cost_centre',             label:'Cost centre',             required:true},
  {f:'job_title',               label:'Category (job title)',    required:true},
  {f:'function',                label:'Function',                required:true},
  {f:'subfunction',             label:'Subfunction',             required:true},
  {f:'work_modality',           label:'Work modality',           required:true},
  {f:'employee_classification', label:'Employee classification', required:true},
  {f:'scope',                   label:'Scope',                   required:true},
  {f:'report_name',             label:'Name',                    required:true},
  {f:'report_surname',          label:'Surname',                 required:true},
  {f:'bench_flag',              label:'Bench',                   required:false},
  {f:'project',                 label:'Project',                 required:true},
  {f:'kpi_target',              label:'KPI target',              required:false},
  {f:'variable_plan',           label:'Variable plan',           required:false},
  {f:'digital_flag',            label:'Digital',                 required:false},
  {f:'hiring_reason',           label:'Hiring reason',           required:false}
];

function hrSikPreflight(year, month){
  if(!isHR_()) throw new Error('HR only.');
  year=parseInt(year); month=parseInt(month);
  if(!year||!month) return {ok:false,msg:'Choose a year and a month.'};
  const period=year+'-'+String(month).padStart(2,'0');

  const psh=sheet_(TAB_PAYROLL);
  if(!psh || psh.getLastRow()<2) return {ok:false,msg:'The PAYROLL tab is empty.'};
  const phdr=psh.getRange(1,1,1,psh.getLastColumn()).getValues()[0];
  const pi=function(f){return phdr.indexOf(f);};
  if(pi('period')===-1) return {ok:false,msg:'The PAYROLL tab has no period column.'};
  const prows=psh.getRange(2,1,psh.getLastRow()-1,psh.getLastColumn()).getValues()
                 .filter(function(r){ return String(r[pi('period')]).trim()===period; });
  if(!prows.length) return {ok:false,msg:'No payroll rows for '+period+'. Paste the month in first.'};

  const E=empData_(true), eh=E.hdr;
  const ec=function(f){return eh.indexOf(f);};
  const emp={};
  E.rows.forEach(function(rec){
    emp[String(rec.values[ec('employee_id')]).trim().toUpperCase()]=rec.values;
  });

  const mStart=new Date(year, month-1, 1);
  const mEnd=new Date(year, month, 0);

  const noRecord=[], people=[], byField={};
  let inScope=0, clean=0;

  prows.forEach(function(p){
    const eid=String(p[pi('Code')]).trim().toUpperCase();
    if(!eid) return;
    const e=emp[eid];
    if(!e){
      noRecord.push({employee_id:eid, name:String(p[pi('Employee Name')]||'')});
      return;
    }
    const g=function(f){ const c=ec(f); return c===-1? '' : String(fmt_(e[c])).trim(); };

    const hire=g('hire_date')? new Date(g('hire_date')) : null;
    const exit=g('exit_date')? new Date(g('exit_date')) : null;
    if(hire && !isNaN(hire) && hire>mEnd) return;
    if(exit && !isNaN(exit) && exit<mStart) return;
    inScope++;
    const leftThisMonth = exit && !isNaN(exit) && exit>=mStart && exit<=mEnd;

    const missing=[], optional=[];
    SIK_REQUIRED.forEach(function(x){
      const c=ec(x.f);
      if(c===-1){ return; }                       // column not on the sheet at all
      if(g(x.f)) return;
      if(x.required){ missing.push(x.label); byField[x.label]=(byField[x.label]||0)+1; }
      else optional.push(x.label);
    });
    // a leaver with no discharge code cannot be reported properly
    if(leftThisMonth && !g('exit_type_code')){
      missing.push('Discharge code'); byField['Discharge code']=(byField['Discharge code']||0)+1;
    }

    if(!missing.length){ clean++; return; }
    people.push({employee_id:eid, name:g('full_name_en')||String(p[pi('Employee Name')]||''),
                 job_title:g('job_title'), missing:missing, optional:optional,
                 leaver:!!leftThisMonth});
  });

  // worst first, so the biggest gaps get fixed first
  people.sort(function(a,b){ return b.missing.length-a.missing.length; });
  const fields=Object.keys(byField).map(function(k){return {field:k, count:byField[k]};})
                     .sort(function(a,b){return b.count-a.count;});

  // columns the sheet does not have at all — a different kind of problem
  const absentCols=SIK_REQUIRED.filter(function(x){return ec(x.f)===-1;})
                               .map(function(x){return x.label+' ('+x.f+')';});

  return {ok:true, period:period,
    inScope:inScope, clean:clean, incomplete:people.length,
    noRecord:noRecord, fields:fields, people:people, absentCols:absentCols,
    ready: people.length===0 && noRecord.length===0,
    msg: people.length===0 && noRecord.length===0
      ? inScope+' employee(s) in scope, all complete. Ready to build.'
      : inScope+' in scope — '+clean+' complete, '+people.length+' with something missing'+
        (noRecord.length? ', and '+noRecord.length+' with no employee record at all' : '')+'.'};
}

// ================================================================
// CONTRACT GENERATION
//   Copies the Google Doc template, fills it from the record, and puts
//   the result in a Drive folder. Word output, so HR can complete the
//   few fields the database cannot supply before printing.
//
//   Nothing is signed here. Generating a contract does not mark it
//   signed — that stays with the file review, which is what the medical
//   enrolment gate reads.
// ================================================================

const CONTRACT_TEMPLATES = {
  'Fixed period': '1MF9EEk67zMBl1UHlz5rS-gQlq2slurZACcMNvMCb8A8',
  'Task-Based':   '1C65BzlJatIKztntX_KTUxNVIyBBwZN-UscH-p2ksSaQ'
};

// Where generated contracts land. Left blank, they go to the root of My Drive;
// put a folder ID here to keep them together.
const CONTRACT_FOLDER_ID = '1XmrgkIRqWvDWWRsCQuSP63bYIdwdaqm5';

const TAB_CONTRACTS = 'CONTRACTS';

function contractHdr_(){
  const sh=sheet_(TAB_CONTRACTS);
  if(!sh) return null;
  return sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
}

function contractTypesAvailable(){
  if(!isHR_()) throw new Error('HR only.');
  return Object.keys(CONTRACT_TEMPLATES);
}

// Everything the template asks for. Anything the record does not hold is
// left as the placeholder, so it is obvious in the document rather than
// silently blank.
function contractFieldsFor_(eid, contractType){
  const f=employeeFieldsOf_(eid,['employee_id','full_name_en','full_name_ar','national_id',
    'date_of_birth','nationality','address','city','governorate','mobile','personal_email',
    'insurance_number','bank_name','iban','emergency_contact_name','emergency_contact_phone',
    'emergency_contact_relation','job_title','function','project','work_location',
    'direct_manager','basic_salary','hire_date','contract_end_date','contract_type']);
  if(!f.employee_id) return null;

  const addr=[f.address, f.city, f.governorate].filter(String).join(', ');
  const emerg=[f.emergency_contact_name,
               f.emergency_contact_relation? '('+f.emergency_contact_relation+')' : '',
               f.emergency_contact_phone].filter(String).join(' ');
  const bank=[f.bank_name, f.iban].filter(String).join(' / ');

  // the manager is stored as an ID; the contract needs a person
  let mgr=f.direct_manager;
  if(mgr){
    const m=employeeFieldsOf_(resolveApprover_(mgr),['full_name_en']);
    if(m.full_name_en) mgr=m.full_name_en;
  }

  const salary=f.basic_salary? Number(String(f.basic_salary).replace(/,/g,'')).toLocaleString('en-US') : '';

  return {
    '{{FULL_NAME}}':      f.full_name_en,
    '{{FULL_NAME_AR}}':   f.full_name_ar,
    '{{SIG_NAME}}':       f.full_name_en,
    '{{SIG_NAME_AR}}':    f.full_name_ar,
    '{{NATIONAL_ID}}':    f.national_id,
    '{{DOB}}':            f.date_of_birth,
    '{{NATIONALITY}}':    f.nationality,
    '{{ADDRESS}}':        addr,
    '{{MOBILE}}':         f.mobile,
    '{{EMAIL}}':          f.personal_email,
    '{{INSURANCE_NO}}':   f.insurance_number,
    '{{BANK_IBAN}}':      bank,
    '{{EMERGENCY}}':      emerg,
    '{{JOB_TITLE}}':      f.job_title,
    '{{DEPARTMENT}}':     f['function'],
    '{{PROJECT}}':        f.project,
    '{{WORK_LOCATION}}':  f.work_location,
    '{{DIRECT_MANAGER}}': mgr,
    '{{BASIC_SALARY}}':   salary,
    '{{START_DATE}}':     f.hire_date,
    '{{END_DATE}}':       f.contract_end_date,
    '{{CONTRACT_DATE}}':  f.hire_date,
    '{{CONTRACT_REF}}':   'KE-'+(f.hire_date? String(f.hire_date).slice(0,4) : new Date().getFullYear())+'-'+f.employee_id,
    // left for HR: job description is Annex 2, and the task fields are per assignment
    '{{JOB_DESCRIPTION}}':'',
    '{{TASK_NAME}}':      '',
    '{{TASK_DESCRIPTION}}':'',
    '{{DELIVERABLES}}':   '',
    '{{DURATION}}':       '',
    '_emp': f
  };
}

// What is missing before a contract is worth printing.
function contractPreflight(eid, contractType){
  if(!isHR_()) throw new Error('HR only.');
  eid=String(eid||'').trim().toUpperCase();
  if(!eid) return {found:false, msg:'Enter an employee ID.'};
  const vals=contractFieldsFor_(eid, contractType);
  if(!vals) return {found:false, msg:'No employee found with ID '+eid+'.'};
  const f=vals._emp;

  const need=[
    ['{{FULL_NAME}}','Full name (English)'],['{{FULL_NAME_AR}}','Full name (Arabic)'],
    ['{{NATIONAL_ID}}','National ID'],['{{DOB}}','Date of birth'],
    ['{{NATIONALITY}}','Nationality'],['{{ADDRESS}}','Address'],
    ['{{MOBILE}}','Mobile'],['{{EMAIL}}','Personal email'],
    ['{{INSURANCE_NO}}','Social insurance number'],['{{BANK_IBAN}}','Bank / IBAN'],
    ['{{EMERGENCY}}','Emergency contact'],['{{JOB_TITLE}}','Job title'],
    ['{{DEPARTMENT}}','Department'],['{{WORK_LOCATION}}','Work location'],
    ['{{DIRECT_MANAGER}}','Direct manager'],['{{BASIC_SALARY}}','Basic salary'],
    ['{{START_DATE}}','Start date']
  ];
  if(String(contractType)==='Fixed period') need.push(['{{END_DATE}}','Contract end date']);

  const missing=need.filter(function(x){ return !String(vals[x[0]]||'').trim(); })
                    .map(function(x){ return x[1]; });

  return {found:true, employee_id:f.employee_id, name:f.full_name_en,
    contract_type_on_record:f.contract_type,
    hire_date:f.hire_date, missing:missing,
    ready: missing.length===0,
    msg: missing.length? missing.length+' field(s) missing — the contract will print with gaps.'
                       : 'Everything the contract needs is on the record.'};
}

function generateContract(eid, contractType, extras){
  if(!isHR_()) throw new Error('HR only.');
  eid=String(eid||'').trim().toUpperCase();
  const tplId=CONTRACT_TEMPLATES[contractType];
  if(!tplId) return {ok:false, msg:'No template for "'+contractType+'".'};

  const vals=contractFieldsFor_(eid, contractType);
  if(!vals) return {ok:false, msg:'No employee found with ID '+eid+'.'};
  const f=vals._emp;

  // anything HR typed on the form wins over the blank default
  if(extras){
    ['JOB_DESCRIPTION','TASK_NAME','TASK_DESCRIPTION','DELIVERABLES','DURATION'].forEach(function(k){
      if(extras[k]) vals['{{'+k+'}}']=String(extras[k]).trim();
    });
  }

  const stamp=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd-HHmm');
  const title=eid+' — '+f.full_name_en+' — '+contractType+' — '+stamp;

  const copy=DriveApp.getFileById(tplId).makeCopy(title);
  const doc=DocumentApp.openById(copy.getId());
  const body=doc.getBody();

  Object.keys(vals).forEach(function(k){
    if(k==='_emp') return;
    // a placeholder with nothing behind it stays visible, so the gap is obvious
    body.replaceText(k.replace(/[{}]/g,'\\$&'), String(vals[k]||k));
  });
  // headers and footers carry the ref too
  ['getHeader','getFooter'].forEach(function(fn){
    try{
      const s=doc[fn]();
      if(s) Object.keys(vals).forEach(function(k){
        if(k==='_emp') return;
        s.replaceText(k.replace(/[{}]/g,'\\$&'), String(vals[k]||k));
      });
    }catch(e){}
  });
  doc.saveAndClose();

  if(CONTRACT_FOLDER_ID){
    try{ DriveApp.getFolderById(CONTRACT_FOLDER_ID).addFile(copy);
         DriveApp.getRootFolder().removeFile(copy); }catch(e){}
  }

  // Word, so HR can finish it before printing
  const url='https://docs.google.com/document/d/'+copy.getId()+'/export?format=docx';

  logContract_(eid, f, contractType, copy.getId(), vals['{{CONTRACT_REF}}']);

  return {ok:true, id:copy.getId(), url:url,
    editUrl:'https://docs.google.com/document/d/'+copy.getId()+'/edit',
    ref:vals['{{CONTRACT_REF}}'],
    msg:'Generated for '+f.full_name_en+'. Download it, complete anything left blank, then print for signature.'};
}

// A record of what was produced, so a reprint is traceable.
function logContract_(eid, f, contractType, fileId, ref){
  let sh=sheet_(TAB_CONTRACTS);
  if(!sh){
    sh=ss_().insertSheet(TAB_CONTRACTS);
    const h=['contract_ref','employee_id','employee_name','contract_type','generated_at',
             'generated_by','file_id','file_url','signed','signed_date','notes'];
    sh.getRange(1,1,1,h.length).setValues([h]);
    sh.getRange(1,1,1,h.length).setFontWeight('bold')
      .setBackground('#1F3864').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  }
  const hdr=contractHdr_();
  const row=sh.getLastRow()+1;
  const set=function(k,v){ const c=hdr.indexOf(k); if(c!==-1) sh.getRange(row,c+1).setValue(v); };
  set('contract_ref',ref); set('employee_id',eid); set('employee_name',f.full_name_en);
  set('contract_type',contractType); set('generated_at',new Date());
  set('generated_by',currentUser_()); set('file_id',fileId);
  set('file_url','https://docs.google.com/document/d/'+fileId+'/edit');
  set('signed','No');
  logChange_(eid, f.national_id, 'contract_generated','', contractType+' — '+ref,
             'HR console','Applied','Contract generated');
}

// Every contract produced for one person.
function contractHistory(eid){
  if(!isHR_()) throw new Error('HR only.');
  eid=String(eid||'').trim().toUpperCase();
  const sh=sheet_(TAB_CONTRACTS);
  if(!sh || sh.getLastRow()<2) return [];
  const hdr=contractHdr_(), i=function(f){return hdr.indexOf(f);};
  const out=[];
  sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues().forEach(function(r,n){
    if(String(r[i('employee_id')]).trim().toUpperCase()!==eid) return;
    const o={row:n+2}; hdr.forEach(function(h,c){ o[h]=fmt_(r[c]); });
    out.push(o);
  });
  out.reverse();
  return out;
}

// ================================================================
// CONTRACT SIGNING APPOINTMENTS
//   The employee books a day, not a time. Ten a day, Sunday to Thursday.
//   Once a day is full the next request rolls to the following working day,
//   so nobody is choosing from a calendar and nobody double-books.
//
//   A missed appointment is NOT given back to that day — the slot stays
//   spent. Rebooking lands wherever the queue has got to, which is usually
//   a week or more out. That is deliberate: the cost of not turning up
//   should fall on the person who did not turn up.
//
//   Signing is open to everyone. Receiving a COPY is what requires a
//   complete document file.
// ================================================================

const TAB_APPOINTMENTS = 'SIGNING_APPOINTMENTS';
const APPTS_PER_DAY = 10;
const APPT_WORKING_DAYS = [0,1,2,3,4];      // Sun=0 .. Thu=4
const APPT_SEARCH_LIMIT = 30;               // working days to look ahead

function apptHdr_(){
  let sh=sheet_(TAB_APPOINTMENTS);
  if(!sh){
    sh=ss_().insertSheet(TAB_APPOINTMENTS);
    const h=['appointment_id','employee_id','employee_name','konecta_email','contract_type',
             'appointment_date','requested_at','status','outcome_by','outcome_at','notes'];
    sh.getRange(1,1,1,h.length).setValues([h]);
    sh.getRange(1,1,1,h.length).setFontWeight('bold')
      .setBackground('#1F3864').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  }
  return sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
}

function apptRows_(){
  const hdr=apptHdr_();
  const sh=sheet_(TAB_APPOINTMENTS);
  if(sh.getLastRow()<2) return [];
  const i=function(f){return hdr.indexOf(f);};
  return sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues()
    .map(function(r,n){
      const o={row:n+2};
      hdr.forEach(function(h,c){ o[h]=fmt_(r[c]); });
      return o;
    });
}

function isApptWorkingDay_(d){ return APPT_WORKING_DAYS.indexOf(d.getDay())!==-1; }

// The next day with room. Counts every booking made for that day, whether or
// not the person turned up — a spent slot stays spent.
function nextAvailableApptDate_(){
  const taken={};
  apptRows_().forEach(function(a){
    const d=String(a.appointment_date||'').trim();
    if(!d) return;
    if(String(a.status).trim()==='Cancelled') return;   // only a cancellation frees a slot
    taken[d]=(taken[d]||0)+1;
  });
  const cur=new Date(); cur.setHours(0,0,0,0);
  cur.setDate(cur.getDate()+1);                          // never today — HR needs notice
  for(let n=0;n<APPT_SEARCH_LIMIT*2;n++){
    if(isApptWorkingDay_(cur)){
      const key=Utilities.formatDate(cur,Session.getScriptTimeZone(),'yyyy-MM-dd');
      if((taken[key]||0) < APPTS_PER_DAY) return {date:key, booked:taken[key]||0};
    }
    cur.setDate(cur.getDate()+1);
  }
  return null;
}

function myOpenAppointment_(eid){
  const open=apptRows_().filter(function(a){
    return String(a.employee_id).trim().toUpperCase()===eid &&
           String(a.status).trim()==='Booked';
  });
  return open.length? open[open.length-1] : null;
}

// ---------- what the employee sees on My Contract ----------
function getMyContract(){
  const me=getMyRecord();
  if(!me.found) return {found:false};
  const eid=me.readonly.employee_id;

  const f=employeeFieldsOf_(eid,['employee_id','full_name_en','job_title','hire_date',
    'contract_type','contract_end_date','company_type']);

  // documents: what is in, what is not, judged against their own list
  let docs=[], outstanding=0, received=0, signed='No';
  try{
    const held=docsFor_(eid);
    const who=whoFor_(employeeFieldsOf_(eid,['gender','nationality']));
    docs=fileChecklist_(who).map(function(d){
      const h=held[d.no];
      const isIn=!!(h && h.status==='Received');
      if(isIn) received++; else if(d.mandatory) outstanding++;
      return {no:d.no, name:d.name, optional:d.optional,
              status:isIn? 'Received':'Missing'};
    });
    const fs=fileStatusFor_(eid);
    if(fs){
      signed=String(fs.contract_signed||'No');
      if(String(fs.docs_outstanding||'')!=='') outstanding=Number(fs.docs_outstanding)||0;
    }
  }catch(e){}
// Konecta is not the employer of record here — no contract, no appointment
  if(isNonEmployee_(f.company_type)){
    return {found:true, non_employee:true, employee_id:eid, name:f.full_name_en,
            company_type:f.company_type};
  }
  const appt=myOpenAppointment_(eid);
  const next=appt? null : nextAvailableApptDate_();

  // the contract on file, for reading only
  let contract=null;
  try{
    const hist=contractHistoryFor_(eid);
    if(hist.length) contract={ref:hist[0].contract_ref, type:hist[0].contract_type,
                              generated_at:hist[0].generated_at, id:hist[0].file_id};
  }catch(e){}

  return {found:true, employee_id:eid, name:f.full_name_en,
    job_title:f.job_title, hire_date:f.hire_date,
    contract_type:f.contract_type, contract_end_date:f.contract_end_date,
    contract_signed:signed,
    documents:docs, docs_received:received, docs_outstanding:outstanding,
    can_get_copy: outstanding===0,
    contract:contract,
    appointment:appt? {id:appt.appointment_id, date:appt.appointment_date,
                       requested_at:appt.requested_at} : null,
    next_available: next? next.date : null,
    per_day: APPTS_PER_DAY};
}

// contractHistory is HR-guarded; the employee needs the same read on themselves
function contractHistoryFor_(eid){
  eid=String(eid||'').trim().toUpperCase();
  const sh=sheet_(TAB_CONTRACTS);
  if(!sh || sh.getLastRow()<2) return [];
  const hdr=contractHdr_(), i=function(f){return hdr.indexOf(f);};
  const out=[];
  sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues().forEach(function(r,n){
    if(String(r[i('employee_id')]).trim().toUpperCase()!==eid) return;
    const o={row:n+2}; hdr.forEach(function(h,c){ o[h]=fmt_(r[c]); });
    out.push(o);
  });
  out.reverse();
  return out;
}

// ---------- the employee books ----------
function requestSigningAppointment(){
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try{
    const me=getMyRecord();
    if(!me.found) throw new Error('No record found for your account.');
    const eid=me.readonly.employee_id;
    if(!eid) throw new Error('Your record has no employee ID. Please contact HR.');
const co=employeeFieldsOf_(eid,['company_type']).company_type;
    if(isNonEmployee_(co)) return {ok:false,
      msg:'Your contract is held by your own employer, not by Konecta. Please speak to them directly.'};
    const existing=myOpenAppointment_(eid);
    if(existing) return {ok:false,
      msg:'You already have an appointment on '+existing.appointment_date+'. Cancel it first if you need a different day.'};

    const slot=nextAvailableApptDate_();
    if(!slot) return {ok:false,
      msg:'No appointment days are available in the next few weeks. Please contact HR directly.'};

    const f=employeeFieldsOf_(eid,['full_name_en','konecta_email','contract_type','job_title']);
    const sh=sheet_(TAB_APPOINTMENTS), hdr=apptHdr_();
    const row=sh.getLastRow()+1;
    const id='SG-'+String(row-1).padStart(5,'0');
    const set=function(k,v){ const c=hdr.indexOf(k); if(c!==-1) sh.getRange(row,c+1).setValue(v); };
    set('appointment_id',id); set('employee_id',eid);
    set('employee_name',f.full_name_en); set('konecta_email',f.konecta_email);
    set('contract_type',f.contract_type);
    set('appointment_date',slot.date); set('requested_at',new Date());
    set('status','Booked');

    const pretty=Utilities.formatDate(new Date(slot.date),Session.getScriptTimeZone(),'EEEE d MMMM');

    // HR needs to have the contract printed before they arrive
    notifyHR_('Contract signing — '+f.full_name_en+' on '+slot.date,
      f.full_name_en+' ('+eid+') has booked a signing appointment.\n\n'+
      '  Date: '+pretty+'\n'+
      '  Job title: '+(f.job_title||'')+'\n'+
      '  Contract type: '+(f.contract_type||'not recorded')+'\n'+
      '  Reference: '+id+'\n\n'+
      'Generate and print their contract before the day. Mark the outcome in the HR console '+
      'once they have been — that is what records who signed and who did not turn up.');

    if(f.konecta_email){
      try{
        MailApp.sendEmail({
          to: f.konecta_email,
          subject:'Your contract signing appointment — '+pretty,
          htmlBody:
            '<div style="font-family:Arial,sans-serif;max-width:520px">'+
            '<p>Hello '+escapeHtml_(f.full_name_en||'')+',</p>'+
            '<p>Your appointment to sign your employment contract is booked for:</p>'+
            '<div style="background:#EEEDFE;border-radius:8px;padding:16px 20px;margin:14px 0">'+
            '<div style="font-size:18px;font-weight:bold;color:#2800C8">'+escapeHtml_(pretty)+'</div>'+
            '<div style="font-size:13px;color:#6b6b80;margin-top:4px">Reference '+id+'</div></div>'+
            '<p>Come to the People team during the working day. Bring your national ID.</p>'+
            '<div style="background:#FFF9D6;border-left:4px solid #FFE100;padding:12px 16px;margin:14px 0">'+
            'If you cannot make it, cancel it in the app so someone else can take the slot. '+
            'If you simply do not come, the appointment is used up and your next available date '+
            'will be considerably later.</div>'+
            '<p style="font-size:13px;color:#6b6b80">Konecta Egypt — People team</p></div>',
          name:'Konecta Egypt — People Team'
        });
      }catch(e){}
    }

    return {ok:true, id:id, date:slot.date,
      msg:'Booked for '+pretty+'. We have emailed you the details.'};
  } finally { lock.releaseLock(); }
}

function cancelSigningAppointment(){
  const me=getMyRecord();
  if(!me.found) throw new Error('No record found for your account.');
  const eid=me.readonly.employee_id;
  const appt=myOpenAppointment_(eid);
  if(!appt) return {ok:false,msg:'You have no appointment to cancel.'};
  const sh=sheet_(TAB_APPOINTMENTS), hdr=apptHdr_();
  const set=function(k,v){ const c=hdr.indexOf(k); if(c!==-1) sh.getRange(appt.row,c+1).setValue(v); };
  set('status','Cancelled'); set('outcome_by',currentUser_()); set('outcome_at',new Date());
  set('notes','Cancelled by the employee');
  notifyHR_('Signing appointment cancelled — '+appt.employee_name,
    appt.employee_name+' ('+eid+') has cancelled their appointment on '+appt.appointment_date+'.\n\n'+
    'The slot is free again.');
  return {ok:true, msg:'Cancelled. You can book again whenever you are ready.'};
}

// ---------- HR ----------
function hrSigningAppointments(){
  if(!isHR_()) throw new Error('HR only.');
  const today=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');
  const rows=apptRows_().filter(function(a){ return String(a.status).trim()==='Booked'; });

  const byDate={};
  rows.forEach(function(a){
    const d=a.appointment_date;
    if(!byDate[d]) byDate[d]={date:d, isToday:d===today, isPast:d<today, people:[]};
    let outstanding='';
    try{
      const fs=fileStatusFor_(String(a.employee_id).trim().toUpperCase());
      outstanding = fs? (fs.docs_outstanding||'0') : '';
    }catch(e){}
    byDate[d].people.push({row:a.row, appointment_id:a.appointment_id,
      employee_id:a.employee_id, name:a.employee_name,
      contract_type:a.contract_type, requested_at:a.requested_at,
      docs_outstanding:outstanding});
  });
  const out=Object.keys(byDate).sort().map(function(k){return byDate[k];});
  return {days:out, per_day:APPTS_PER_DAY,
          total:rows.length,
          overdue: out.filter(function(d){return d.isPast;})
                      .reduce(function(s,d){return s+d.people.length;},0)};
}

// Signed, or did not turn up. Signed writes through to FILE_STATUS, which is
// what the medical enrolment gate reads — so signing unblocks their insurance.
function hrMarkSigning(row, outcome, note, key){
  if(!isHR_()) throw new Error('HR only.');
  const sh=sheet_(TAB_APPOINTMENTS), hdr=apptHdr_();
  row=guardRow_(sh,hdr,row,{appointment_id:key});
  const g=function(k){ const c=hdr.indexOf(k); return c===-1?'':fmt_(sh.getRange(row,c+1).getValue()); };
  const set=function(k,v){ const c=hdr.indexOf(k); if(c!==-1) sh.getRange(row,c+1).setValue(v); };
  const eid=String(g('employee_id')).trim().toUpperCase();
  if(!eid) return {ok:false,msg:'That row has no employee on it.'};
  if(String(g('status')).trim()!=='Booked') return {ok:false,msg:'That appointment is already closed.'};

  set('status', outcome==='signed'? 'Signed':'Did not attend');
  set('outcome_by',currentUser_()); set('outcome_at',new Date());
  if(note) set('notes',note);

  if(outcome!=='signed'){
    if(g('konecta_email')){
      try{ MailApp.sendEmail(g('konecta_email'),
        'You missed your contract signing appointment',
        'Hello '+g('employee_name')+',\n\n'+
        'You did not attend your appointment on '+g('appointment_date')+'.\n\n'+
        'Your contract still needs signing. Book again in the app — but note the slot you '+
        'missed has been used, so the next available date may be some way out.\n\n'+
        'Konecta Egypt — People team'); }catch(e){}
    }
    return {ok:true, msg:'Recorded as a no-show. They have been emailed and can book again.'};
  }

  // write the signature through to the file status
  let carried='';
  try{
    const fsh=sheet_(TAB_FILE_STATUS), fhdr=fileStatusHdr_();
    const existing=fileStatusFor_(eid);
    const frow=existing? existing.row : fsh.getLastRow()+1;
    const fset=function(k,v){ const c=fhdr.indexOf(k); if(c!==-1) fsh.getRange(frow,c+1).setValue(v); };
    if(!existing){
      const f=employeeFieldsOf_(eid,['full_name_en','hire_date','company_type']);
      fset('employee_id',eid); fset('employee_name',f.full_name_en);
      fset('hire_date',f.hire_date); fset('company_type',f.company_type);
    }
    fset('contract_signed','Yes');
    fset('contract_signed_date', g('appointment_date'));
    fset('contract_type_at_signing', g('contract_type'));
    fset('last_reviewed_at',new Date()); fset('last_reviewed_by',currentUser_());
    carried=' Recorded as signed on their file.';
  }catch(e){ carried=' Could not update FILE_STATUS: '+e.message; }

  logChange_(eid,'','contract_signed','No','Yes','Signing appointment','Applied',
             'Signed on '+g('appointment_date'));

  // signing is the gate the medical module waits on
  let med='';
  try{
    const already=medicalRecordFor_(eid);
    if(!already || already.status!=='Enrolled'){
      const r=medicalEnrol(eid,'Auto-enrolled — contract signed at appointment '+g('appointment_id'));
      med = r.ok? ' Medical enrolment sent.' : ' Medical enrolment not sent: '+r.msg;
    }
  }catch(e){ med=' Medical enrolment failed: '+e.message; }

  return {ok:true, msg:'Recorded as signed.'+carried+med};
}

// ================================================================
// WHO HAS NOT SIGNED — and inviting them in
//   The other half of the appointment queue. Rather than waiting for
//   people to come forward, HR works through a list and books them in.
//   Same ten-a-day cap, so inviting fifteen people spreads them over two
//   days automatically. Nobody is managing a calendar.
// ================================================================

// Everyone who ought to have a signed contract and does not.
function hrUnsignedContracts(){
  if(!isHR_()) throw new Error('HR only.');
  const E=empData_(false), h=E.hdr;
  const c=function(f){return h.indexOf(f);};
  const cEid=c('employee_id'), cNm=c('full_name_en'), cHire=c('hire_date'),
        cType=c('contract_type'), cCo=c('company_type'), cJt=c('job_title'),
        cProj=c('project'), cSt=c('record_status');

  // one read each, rather than per person
  const signed={}, outstanding={};
  try{
    const fsh=sheet_(TAB_FILE_STATUS);
    if(fsh && fsh.getLastRow()>1){
      const fh=fileStatusHdr_(), fi=function(f){return fh.indexOf(f);};
      fsh.getRange(2,1,fsh.getLastRow()-1,fsh.getLastColumn()).getValues().forEach(function(r){
        const id=String(r[fi('employee_id')]).trim().toUpperCase();
        if(!id) return;
        signed[id]=String(r[fi('contract_signed')]).trim();
        outstanding[id]=r[fi('docs_outstanding')];
      });
    }
  }catch(e){}

  const booked={};
  apptRows_().forEach(function(a){
    if(String(a.status).trim()!=='Booked') return;
    booked[String(a.employee_id).trim().toUpperCase()]=a.appointment_date;
  });

  const out=[];
  E.rows.forEach(function(rec){
    const v=rec.values;
    const eid=String(v[cEid]).trim().toUpperCase();
    if(!eid) return;
    // a subcontractor signs with their vendor, not with us
    if(isNonEmployee_(fmt_(v[cCo]))) return;
    if(signed[eid]==='Yes') return;

    const hire=fmt_(v[cHire]);
    let days='';
    if(hire){ const d=new Date(hire); if(!isNaN(d)) days=Math.floor((new Date()-d)/86400000); }

    out.push({row:rec.row, employee_id:eid, name:fmt_(v[cNm]),
      job_title:fmt_(v[cJt]), project:fmt_(v[cProj]),
      hire_date:hire, days_since_hire:days,
      contract_type:fmt_(v[cType])||'(not recorded)',
      record_status:fmt_(v[cSt]),
      docs_outstanding: outstanding[eid]===undefined? '' : outstanding[eid],
      file_status: signed[eid]||'(never reviewed)',
      booked: booked[eid]||''});
  });

  // longest-serving unsigned first — those are the ones that have been let slide
  out.sort(function(a,b){ return (b.days_since_hire||0)-(a.days_since_hire||0); });

  return {rows:out, total:out.length,
    unbooked: out.filter(function(x){return !x.booked;}).length,
    booked: out.filter(function(x){return x.booked;}).length};
}

// HR books people in. Pass a list of employee IDs; they fill the queue in
// order, rolling to the next working day as each one fills.
function hrInviteToSign(employeeIds, note){
  if(!isHR_()) throw new Error('HR only.');
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const ids=(employeeIds||[]).map(function(x){return String(x).trim().toUpperCase();})
                               .filter(String);
    if(!ids.length) return {ok:false,msg:'Choose at least one person.'};

    const sh=sheet_(TAB_APPOINTMENTS), hdr=apptHdr_();
    const booked=[], skipped=[];

    ids.forEach(function(eid){
      if(myOpenAppointment_(eid)){
        skipped.push(eid+' — already has an appointment');
        return;
      }
      const slot=nextAvailableApptDate_();
      if(!slot){ skipped.push(eid+' — no day available'); return; }

      const f=employeeFieldsOf_(eid,['full_name_en','konecta_email','contract_type','job_title']);
      if(!f.full_name_en){ skipped.push(eid+' — no employee record'); return; }

      const row=sh.getLastRow()+1;
      const id='SG-'+String(row-1).padStart(5,'0');
      const set=function(k,v){ const c=hdr.indexOf(k); if(c!==-1) sh.getRange(row,c+1).setValue(v); };
      set('appointment_id',id); set('employee_id',eid);
      set('employee_name',f.full_name_en); set('konecta_email',f.konecta_email);
      set('contract_type',f.contract_type);
      set('appointment_date',slot.date); set('requested_at',new Date());
      set('status','Booked');
      set('notes','Booked by HR'+(note? ' — '+note : ''));

      const pretty=Utilities.formatDate(new Date(slot.date),Session.getScriptTimeZone(),'EEEE d MMMM');
      if(f.konecta_email){
        try{
          MailApp.sendEmail({
            to: f.konecta_email,
            subject:'Please come and sign your employment contract — '+pretty,
            htmlBody:
              '<div style="font-family:Arial,sans-serif;max-width:520px">'+
              '<p>Hello '+escapeHtml_(f.full_name_en)+',</p>'+
              '<p>Your employment contract is ready to sign. We have booked you in for:</p>'+
              '<div style="background:#EEEDFE;border-radius:8px;padding:16px 20px;margin:14px 0">'+
              '<div style="font-size:18px;font-weight:bold;color:#2800C8">'+escapeHtml_(pretty)+'</div>'+
              '<div style="font-size:13px;color:#6b6b80;margin-top:4px">Reference '+id+'</div></div>'+
              '<p>Come to the People team during the working day and bring your national ID. '+
              'You can read your contract in the app before you come — open the <strong>My Contract</strong> tab.</p>'+
              (note? '<p>'+escapeHtml_(note)+'</p>' : '')+
              '<div style="background:#FFF9D6;border-left:4px solid #FFE100;padding:12px 16px;margin:14px 0">'+
              'If that day does not work, cancel it in the app and book another. '+
              'If you simply do not come, the slot is used up and the next available date will be later.</div>'+
              '<p style="font-size:13px;color:#6b6b80">Konecta Egypt — People team</p></div>',
            name:'Konecta Egypt — People Team'
          });
        }catch(e){}
      }
      booked.push({employee_id:eid, name:f.full_name_en, date:slot.date});
    });

    // one summary rather than an email per person
    if(booked.length){
      const byDate={};
      booked.forEach(function(b){ (byDate[b.date]=byDate[b.date]||[]).push(b); });
      let body=booked.length+' person(s) booked in by '+currentUser_()+'.\n\n';
      Object.keys(byDate).sort().forEach(function(d){
        body+='--- '+d+' ('+byDate[d].length+') ---\n';
        byDate[d].forEach(function(b){ body+='  '+b.employee_id+'  '+b.name+'\n'; });
        body+='\n';
      });
      body+='Each has been emailed. Print their contracts before the day.';
      notifyHR_('Signing appointments booked — '+booked.length+' person(s)', body);
    }

    const days=Object.keys(booked.reduce(function(a,b){a[b.date]=1;return a;},{})).sort();
    return {ok:true, booked:booked.length, skipped:skipped, days:days,
      msg: booked.length+' booked'+
        (days.length? ' across '+days.length+' day(s): '+days.join(', ') : '')+'.'+
        (skipped.length? ' '+skipped.length+' skipped.' : '')+
        ' Everyone booked has been emailed.'};
  } finally { lock.releaseLock(); }
}

// ================================================================
// DEPARTMENTS
//   The organisational cut, separate from `function` which is the SIK
//   reporting classification. Used to validate a bulk load, and to work
//   out who a document chase escalates to.
// ================================================================
function departmentNames_(){
  const sh=sheet_('DEPARTMENTS');
  const out={};
  if(!sh || sh.getLastRow()<2) return out;
  const hdr=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const i=function(f){return hdr.indexOf(f);};
  sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues().forEach(function(r){
    const d=String(r[i('department')]||'').trim();
    if(!d) return;
    out[d]={head_id:String(r[i('head_employee_id')]||'').trim(),
            head_name:String(r[i('head_name')]||'').trim()};
  });
  return out;
}

function departmentList(){
  if(!isHR_()) throw new Error('HR only.');
  return Object.keys(departmentNames_());
}
