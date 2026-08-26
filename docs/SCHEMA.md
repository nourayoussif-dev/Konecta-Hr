# Schema — the columns this code depends on

Every sheet read in this app resolves columns **by header name at runtime**:

```js
const hdr = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
const i   = hdr.indexOf('employee_id');
```

There is no schema validation anywhere. That has three consequences worth
being explicit about:

1. **Renaming a column in the Sheet breaks code, and reads break _silently_.**
   `indexOf` returns `-1`. On a read, `row[-1]` is `undefined` — no error, the
   value just arrives blank and flows onward as if the cell were empty. On a
   write, `getRange(r, -1 + 1)` is `getRange(r, 0)`, which throws. So a renamed
   column shows up as mysteriously blank data long before anything errors.
2. **Reordering columns is safe.** Lookup is by name, not position. This is a
   genuine strength of the current design and should be preserved.
3. **Adding columns is safe**, as long as the header row stays row 1.

This file is the inventory of every column name the code requires, grouped by
the domain that uses it. Treat it as the contract between the Sheet and the
code. If you rename a header here, grep for it first.

> Generated from `Code.gs` by scanning header lookups
> (`indexOf('…')`, `i('…')`, `c('…')`). Regenerate after structural changes.

---

## Core: EMPLOYEES / CHANGE LOG / LISTS / INTAKE
_Code.gs lines 1-1123 — 23 columns referenced_

`account_number`, `bank_name`, `bank_verified`, `closed`, `direct_manager`, `employee_id`, `exit_date`, `exit_type`, `full_name_en`, `function`, `hire_date`, `iban`, `job_title`, `konecta_email`, `n_level`, `national_id`, `national_id_verified`, `personal_email`, `project`, `record_status`, `reporting_validated`, `updated_at`, `updated_by`

## Leave: LEAVE, HOLIDAYS, LEAVE_TYPES, LEAVE_ADJUSTMENTS, DELEGATES
_Code.gs lines 1124-2237 — 48 columns referenced_

`active`, `added_by`, `adjustment_date`, `approved_dates`, `approves_via`, `auto_approved`, `calendar_event_id`, `can_view`, `created_at`, `created_by`, `days`, `days_approved`, `days_rejected`, `days_requested`, `delegate_email`, `direct_at`, `direct_by`, `direct_manager`, `direct_status`, `dotted_at`, `dotted_by`, `dotted_manager`, `dotted_status`, `email`, `employee_id`, `employee_name`, `end_date`, `final_status`, `from_date`, `full_name_en`, `konecta_email`, `last_reminder_at`, `leave_entitlement`, `leave_type`, `manager_id`, `note`, `notes`, `payroll_flag`, `reason`, `record_status`, `reminder_count`, `request_id`, `start_date`, `submitted_at`, `to_date`, `track`, `weekend_pattern`, `worked_days`

## Resignations: RESIGNATIONS
_Code.gs lines 2238-2701 — 9 columns referenced_

`direct_manager`, `dotted_manager`, `employee_id`, `final_status`, `last_reminder_at`, `proposed_last_day`, `reminder_count`, `submitted_at`, `withdraw_status`

## Clearance: CLEARANCE
_Code.gs lines 2702-2995 — 4 columns referenced_

`direct_manager`, `dotted_manager`, `employee_id`, `final_status`

## No-show: NO_SHOW
_Code.gs lines 2996-3212 — 2 columns referenced_

`employee_id`, `konecta_email`

## Contract expiry: EXPIRY_LOG
_Code.gs lines 3213-3352 — 9 columns referenced_

`contract_end_date`, `contract_type`, `direct_manager`, `employee_id`, `full_name_en`, `hire_date`, `job_title`, `project`, `record_status`

## Terminations: TERMINATIONS
_Code.gs lines 3353-3619 — 1 columns referenced_

`employee_id`

## Medical: MEDICAL_INSURANCE
_Code.gs lines 3620-3819 — 5 columns referenced_

`contract_signed`, `employee_id`, `full_name_en`, `last_working_day`, `record_status`

## Documents: DOC_CHECKLIST, FILE_STATUS, EMPLOYEE_DOCS
_Code.gs lines 3820-4334 — 12 columns referenced_

`applies_to`, `contract_signed`, `contract_type`, `doc_name_ar`, `doc_name_en`, `doc_no`, `docs_outstanding`, `egypt`, `employee_id`, `notes`, `received_at`, `status`

## Bulk update + career history (EMPLOYEES)
_Code.gs lines 4335-4639 — 7 columns referenced_

`basic_salary`, `cost_centre`, `employee_id`, `full_name_en`, `grade`, `job_title`, `national_id`

## Dependants: DEPENDANTS
_Code.gs lines 4640-5050 — 4 columns referenced_

`dependants`, `employee_id`, `requested_at`, `status`

## Bulk leavers (EMPLOYEES)
_Code.gs lines 5051-5186 — 6 columns referenced_

`employee_id`, `exit_date`, `full_name_en`, `hire_date`, `national_id`, `record_status`

## Payroll + statutory: PAYROLL, SIK report
_Code.gs lines 5187-5531 — 8 columns referenced_

`approved_dates`, `days_approved`, `days_requested`, `employee_id`, `end_date`, `final_status`, `leave_type`, `start_date`

## Contracts: CONTRACTS
_Code.gs lines 5532-5759 — 1 columns referenced_

`employee_id`

## Signing appointments: SIGNING_APPOINTMENTS
_Code.gs lines 5760-6249 — 11 columns referenced_

`company_type`, `contract_type`, `department`, `employee_id`, `full_name_en`, `head_employee_id`, `head_name`, `hire_date`, `job_title`, `project`, `record_status`

---

## `EMPLOYEES` — the master record

The employee table is the spine of the app; most other tabs reference it by
`employee_id`. These groups are declared explicitly in `Code.gs` lines 25–55.

**Gate 1 — identity** (`GATE1`)
`national_id`, `full_name_en`, `full_name_ar`, `personal_email`, `mobile`

**Gate 2 — offer** (`GATE2`)
`hire_date`, `job_title`, `function`, `contract_type`, `direct_manager`, `basic_salary`

**Gate 3 — payment** (`GATE3`)
`insurance_number`, `bank_name`, `account_number`, `iban`, `bank_verified`

**Employee-editable** (`EMPLOYEE_EDITABLE`) — what self-service may write
`full_name_en`, `full_name_ar`, `personal_email`, `passport_number`,
`date_of_birth`, `nationality`, `religion`, `education_level`,
`has_disability`, `marital_status`, `dependants`, `dependant1_name`,
`dependant1_dob`, `dependant1_relation`, `dependant1_national_id`,
`dependant2_name`, `dependant2_dob`, `dependant2_relation`,
`dependant2_national_id`, `dependant3_name`, `dependant3_dob`,
`dependant3_relation`, `dependant3_national_id`, `mobile`, `address`, `city`,
`governorate`, `emergency_contact_name`, `emergency_contact_phone`,
`emergency_contact_relation`, `insurance_number`

**Bank** (`BANK_FIELDS`) — separate flow, requires HR verification
`bank_name`, `account_number`, `iban`

**Read-only to the employee** (`READ_ONLY_VISIBLE`)
`employee_id`, `national_id`, `national_id_verified`, `konecta_email`,
`record_status`, `hire_date`, `job_title`, `grade`, `function`, `subfunction`,
`contract_type`, `direct_manager`, `work_location`, `basic_salary`,
`kpi_target`, `kpi_frequency`, `incentive`, `transportation`, `bank_verified`,
`payment_status`

> `gcm` is deliberately excluded — employees must never see their GCM level.
> Any change to the field-visibility lists should re-check that.

**HR-locked** (`HR_LOCKED`) — computed, never hand-edited
`employee_id`, `completeness_%`, `blocking_gaps`, `chase_gaps`, `report_name`

---

## `record_status` — the employee state machine

Values treated as *visible* in working views (`VISIBLE_STATUSES`):
`Active`, `Serving Notice`, `Final Month`, `On Hold`, `Identity Verified`,
`Pending`

Leavers stay visible for `LEAVER_GRACE_DAYS` (30) after `exit_date`, so final
pay and clearance can complete. The rehire check and the attrition report
deliberately bypass this and read full history.

---

## Business rules that live in code, not in the Sheet

These are hardcoded constants. Changing any of them today requires a code
change and a redeploy — they are strong candidates for a `CONFIG` tab.

| Constant | File | Value | Meaning |
|---|---|---:|---|
| `LEAVE_FIRST_YEAR_DAYS` | `Code.gs` | 15 | Annual leave in the joining year |
| standard entitlement | `Code.gs` | 21 | Days from the year after joining |
| age-50 entitlement | `Code.gs` | 30 | Days once aged 50+ |
| disability entitlement | `Code.gs` | 45 | Days where `has_disability = yes` |
| `AUTO_APPROVE_AFTER_DAYS` | `Code.gs` | 5 | Working days before annual leave self-approves |
| `REMIND_ON_DAYS` | `Code.gs` | 2, 4 | Reminder cadence; day 4 copies HR |
| `RESIGN_AUTO_DAYS` | `Code.gs` | 10 | Calendar days before the employee's date stands |
| `WITHDRAW_FREE_DAYS` | `Code.gs` | 10 | Window to withdraw without approval |
| `PROBATION_DAYS` | `Code.gs` | 90 | Egyptian labour law probation |
| `EXPIRY_WARN_DAYS` | `Code.gs` | 60 | Contract expiry warning lead time |
| `LEAVER_GRACE_DAYS` | `Code.gs` | 30 | How long leavers stay visible |
| `COMPANY_FUNDED_LIMIT` | `Code.gs` | 3 | Dependants the company funds |
| `APPTS_PER_DAY` | `Code.gs` | 10 | Contract-signing slots per day |
| `ID_PREFIX` / `ID_PAD` | `Code.gs` | `EG` / 4 | Employee ID format, e.g. `EG0001` |
| `PERSONAL_EXEMPTION_ANNUAL` | `payrollengine.gs` | 20,000 | Annual personal tax exemption |
| `SI_EMPLOYEE` / `SI_EMPLOYER` | `payrollengine.gs` | 11% / 18.75% | Social insurance shares |
| `INS_WAGE_MIN` / `INS_WAGE_MAX` | `payrollengine.gs` | 2,300 / 16,700 | Insurable wage floor and ceiling |
| `BRACKET_LOWER` / `RATE_DELTA` | `payrollengine.gs` | — | Income tax brackets |
| `CANCEL` | `payrollengine.gs` | — | Bracket cancellation thresholds |

Identities are hardcoded too, and go stale when people change role:
`HR_ADMINS`, `IT_USERS`, `FACILITIES_USERS`, `MEDICAL_CONTACT`,
`NLEVEL_TOP`, `SI_EXEMPT`, `SHEET_ID`, `CONTRACT_FOLDER_ID`.
