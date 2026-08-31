# Server-side authorization — the entry-point sweep

Every function the browser can reach via `google.script.run` must enforce its
own authorization **server-side**. The client's role routing (`routeByRole`)
only decides what to *render*; any employee can open the browser console and
call any server function directly, so a function without its own gate is
public to the whole company.

This file is the record of the sweep of all 88 entry points (26/08/2026),
re-checkable with the classifier described at the bottom.

## The gates

| Gate | Meaning |
|---|---|
| `isHR_()` | caller's email is in `HR_ADMINS` |
| `isIT_()` | caller's email is in `IT_USERS` |
| `isFacilities_()` | caller's email is in `FACILITIES_USERS` |
| `getManagerIdentity_()` | caller appears as someone's manager |
| `actingFor_()` | caller is a manager or an active delegate |
| self-scoped | operates only on the record matching `currentUser_()`'s email — the caller cannot name a target |
| row guard | `guardRow_` / `guardEmpRow_` verifies the target row's identity (see the row-identity commit) |

## Sweep result

**HR console (34 functions)** — all gated `isHR_()`:
`bulkFieldOptions` `careerHistory` `contractPreflight` `contractTypesAvailable`
`fileReviewProgress` `generateContract` `getFileReview` `historicExitTypes`
`hrApproveIntake` `hrAttritionReport` `hrBulkBacklog` `hrBulkHolidayWorked`
`hrBulkLeaversApply` `hrBulkLeaversPreview` `hrBulkUpdateApply`
`hrBulkUpdatePreview` `hrCompleteClearance` `hrDecideTermination`
`hrGetBankPending` `hrGetIntake` `hrGetLeaveToValidate` `hrGetNoShows`
`hrGetPending` `hrGetRecord` `hrGetTaskList` `hrGetTerminations`
`hrInviteToSign` `hrMarkSigning` `hrRejectIntake` `hrRejectedDaysReport`
`hrResolveNoShow` `hrSaveOffer` `hrSaveRecord` `hrSearchEmployees`
`hrSigningAppointments` `hrSikPreflight` `hrSikReport` `hrUnsignedContracts`
`hrValidateLeave` `hrVerifyAndIssue` `hrVerifyBank` `medicalOutstanding`
`medicalRemove` `saveFileReview` `submitResignationFor`

**IT** — `itGetQueue`, `itSetEmail`: gated `isIT_() || isHR_()`.

**Manager / approver** — gated on manager identity, with delegation honoured;
row-guarded where they write:
`getLeaveApprovals` `decideLeave` `getResignationApprovals` `decideResignation`
`decideWithdrawal` `getTeamMemberCard` `managerSaveTeamMember`
`getDelegateOptions` `setDelegate` `endDelegate` `confirmHandover`
`getMyClearanceTasks` `submitClearanceItems` (per-department: IT/Facilities/HR)
`submitLeaveRequestFor` (HR, or the target's direct/dotted manager — checked)
`initiateTermination` (HR, or the target's direct manager — checked)
`getTerminationContext` (fixed 26/08/2026, mirrors `initiateTermination`'s rule);
`getDelegateOptions` (now returns only the caller's own team; outside-team
delegates via the `delegateSearch` server-side typeahead, so the full company
directory no longer crosses to the client)

**Self-scoped (employee self-service)** — the caller cannot name a target;
the record is resolved from their own login:
`getBootstrap` `getRole` `getMyRecord` `getMyLeaveInfo` `submitLeaveRequest`
`getMyResignation` `submitResignation` `withdrawResignation` `getMyContract`
`getMyDependants` `saveMyDependants` `submitPersonalUpdate` `submitBankChange`
`reportIssue` `reportBalanceIssue` `requestSigningAppointment`
`cancelSigningAppointment` `getMyPayslip` `makePayslipPdf`
(both payslip functions honour a target employee ID **only after `isHR_()`**)

**Deliberately open to any signed-in Konecta employee** — reviewed and
accepted, with the reason:

| Function | Why it is open |
|---|---|
| `countLeaveDays`, `getDayPicker` | pure calendar arithmetic on the caller's own inputs |
| `getHolidayList` | the public-holiday list |
| `getLists`, `getProjectMap`, `getManagerOptions` | dropdown option lists |
| `listPayslipPeriods` | period names only, no amounts |
| `lookupEmployeeForNoShow`, `reportNoShow` | the NO_SHOW module is deliberately reportable by anyone — a trainee who stops showing up may have no manager assigned yet. The lookup returns name, job title, project and status only. The report puts payment on hold, so HR reviews every one. |
| `submitNewEmployee`, `validateNationalId` | the joiner intake flow — new joiners have no record yet |

## Residual risks, accepted for now

- The role lists (`HR_ADMINS`, `IT_USERS`, `FACILITIES_USERS`) are hardcoded
  constants. Changing HR staff requires a deploy; a leaver in one of those
  lists keeps access until someone edits code. Planned fix: a ROLES tab
  (Phase 4), at which point this file should be updated.
- `reportNoShow` lets any employee place any employee's payment on hold.
  This is the module's documented intent, HR-reviewed — but it is a large
  lever, and worth revisiting if it is ever abused.

## The reachable set (corrected 26/08/2026)

The first version of this sweep intersected `function fn(` with the functions
`Index.html` actually calls. **That was wrong and it hid a critical hole.**
Apps Script exposes *every* top-level function whose name does not end in `_`
to `google.script.run`, whether or not the client references it. So
`buildPayslip(employeeId, period)` — the ungated engine sitting behind the
gated `getMyPayslip` — was never reviewed, and any employee could read any
colleague's full payslip with one console call. Same for
`publishPayrollMonth`, `runPayrollChecks`, `include`, `openClearance`, and the
daily-trigger functions.

The reachable set is therefore: **every top-level `function name(` across all
`.gs` files whose name does not end in `_`.** Each one must be:

1. **gated** — `isHR_`, `isIT_`, `isFacilities_`, a manager/approver check, or
   self-scoped from `currentUser_`; or
2. **private** — renamed with a trailing `_` so it is unreachable; or
3. **deliberately open** — listed in `tests/reachable.test.js` (`ACCEPTED_OPEN`)
   with a one-line reason.

`tests/reachable.test.js` enforces this on every `node tests/run.js`: it
enumerates the reachable set and fails on any newcomer that is none of the
three. Adding a function without gating, privatising, or accepting it breaks
the build — which is what should have happened to `buildPayslip`.

### Functions that run only on a schedule / from the editor

`leaveDailyRun`, `resignationDailyRun`, `contractExpiryRun`,
`runChecksThisMonth` and `publishPayrollMonth` are reachable globals (installable
triggers bind by name, so they cannot be given a trailing `_` without breaking
the trigger wiring). They are guarded by `assertNotDirectCall_()`, which throws
when the active user differs from the effective (owner) user — the signature of
a web-app call under "execute as Me". A trigger or editor run has active ==
effective, so it passes. `onFormSubmit` refuses any call lacking `e.namedValues`,
which only a real Forms submission carries.

> If you rename any of these functions, update the matching installable trigger
> in the Apps Script editor (Triggers panel) — triggers bind by function name.
