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
`getTerminationContext` (**was the one ungated door** — fixed 26/08/2026, now
mirrors `initiateTermination`'s rule)

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

## Re-running the sweep

Extract the entry points (`google.script.run .fn(` in `Index.html`,
intersected with `function fn(` across the `.gs` files) and check the first
lines of each for a gate. Any new entry point added without one of the gates
above should fail review.
