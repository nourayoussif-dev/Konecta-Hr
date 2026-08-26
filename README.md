# Konecta Egypt — HR Platform

A Google Apps Script web app bound to a Google Sheet. It runs the employee
lifecycle for Konecta Egypt end to end: onboarding, leave, resignation,
clearance, termination, medical insurance, dependants, contracts, statutory
reporting and payroll.

Access is restricted to `@konecta.com` accounts. What each person sees is
decided server-side from their Google identity.

---

## Roles

| Role | How it is decided | What they get |
|---|---|---|
| **Employee** | anyone signed in who is not below | Their own record, leave, payslip, contract, dependants |
| **Manager** | appears as `direct_manager` / `dotted_manager` on someone's record | The above, plus My Team and approvals |
| **HR admin** | email in `HR_ADMINS` (`Code.gs`) | Full console — every queue, every record, bulk tools, reports |
| **IT** | email in `IT_USERS` (`Code.gs`) | Provisioning queue: set `konecta_email` on verified records |
| **Facilities** | email in `FACILITIES_USERS` (`Code.gs`) | Clearance items owned by Facilities |

## The three onboarding gates

1. **Identity** — `national_id`, `full_name_en`, `full_name_ar`, `personal_email`, `mobile`
   → HR matches the physical ID card → issues the employee ID → signals IT
2. **Offer** — `hire_date`, `job_title`, `function`, `contract_type`, `direct_manager`, `basic_salary`
3. **Payment** — `insurance_number`, `bank_name`, `account_number`, `iban`, `bank_verified`

---

## Files

| File | Lines | What it holds |
|---|---:|---|
| `Code.gs` | 6,249 | The monolith — ~15 business domains (see the map below) |
| `Index.html` | 2,852 | The entire frontend: CSS, markup and client JS |
| `payrollengine.gs` | 247 | Egyptian income tax, social insurance, overtime, gross-up, proration |
| `payslip.gs` | 167 | Payslip assembly |
| `payslippdf.gs` | 133 | Payslip PDF rendering to Drive |
| `payrollchecks.gs` | 182 | Pre-payroll validation |
| `payrollarchive.gs` | 75 | Payroll month archive / publish |

> `Index.html` must keep its capital **I** — `doGet` calls
> `HtmlService.createTemplateFromFile('Index')`, and that lookup is
> case-sensitive.

### Domain map for `Code.gs`

| Lines | Domain | Primary tabs |
|---|---|---|
| 1–1123 | Core, onboarding, offers, bank, IT provisioning, intake | `EMPLOYEES`, `CHANGE LOG`, `LISTS`, `INTAKE` |
| 1124–2237 | Leave: entitlement, requests, approvals, delegation, auto-approve | `LEAVE`, `HOLIDAYS`, `LEAVE_TYPES`, `LEAVE_ADJUSTMENTS`, `DELEGATES` |
| 2238–2701 | Resignations, withdrawal, finalisation | `RESIGNATIONS` |
| 2702–2995 | Clearance and handover | `CLEARANCE` |
| 2996–3212 | No-show / drop-out | `NO_SHOW` |
| 3213–3352 | Contract expiry warnings | `EXPIRY_LOG` |
| 3353–3619 | Terminations and probation | `TERMINATIONS` |
| 3620–3819 | Medical insurance enrolment and removal | `MEDICAL_INSURANCE` |
| 3820–4334 | Document checklist and file review | `DOC_CHECKLIST`, `FILE_STATUS`, `EMPLOYEE_DOCS` |
| 4335–4639 | Bulk field updates, career history, movement report | `EMPLOYEES` |
| 4640–5050 | Dependants and medical eligibility | `DEPENDANTS` |
| 5051–5186 | Bulk historic leavers | `EMPLOYEES` |
| 5187–5531 | Payroll wiring and the SIK statutory report | `PAYROLL` |
| 5532–5759 | Contract generation from templates | `CONTRACTS` |
| 5760–6249 | Signing appointments, departments | `SIGNING_APPOINTMENTS` |

See [`docs/SCHEMA.md`](docs/SCHEMA.md) for the columns each domain depends on.

---

## Deployment

The repository is the source of truth. Code flows **git → Apps Script**.

### Recommended: `clasp` (no more copy-paste)

```bash
npm install -g @google/clasp
clasp login

cp .clasp.json.example .clasp.json
# put your real scriptId in it — find it in the Apps Script editor under
# Project Settings → IDs → Script ID
```

Then, from the repo root:

```bash
clasp push          # git working tree  ->  Apps Script project
clasp open          # open the project in the browser
clasp deployments   # list deployments
```

`.clasp.json` is git-ignored because the scriptId is environment-specific.

### Manual fallback

Copy each file's contents into the matching Apps Script file. Filenames must
match exactly, including `Index` (capital I).

### Deployment settings

`appsscript.json` pins:

- **Timezone** `Africa/Cairo` — every date the app formats depends on this
- **Runtime** V8
- **Web app access** `DOMAIN` — Konecta accounts only
- **Execute as** `USER_DEPLOYING`

> **`executeAs` matters more than anything else in this file.** With
> `USER_DEPLOYING` ("Me"), the script reaches the Sheet on its own authority
> and employees never need access to the spreadsheet itself. With
> `USER_ACCESSING`, every employee would need read access to a Sheet holding
> **every salary, national ID and bank account in the company** — and could
> simply open it directly, bypassing every check in this code.
> Confirm the live deployment is set to "Me" before pushing anything else.

If a push produces an authorization error, delete the `oauthScopes` block and
let Apps Script re-detect scopes on the next run.

---

## Scheduled triggers

Set these up under **Triggers** in the Apps Script editor. Nothing in the repo
creates them, so they must be checked after any project rebuild.

| Function | Cadence | What it does |
|---|---|---|
| `leaveDailyRun` | Daily, early morning | Reminders, escalation, auto-approval of annual leave |
| `resignationDailyRun` | Daily | Resignation reminders, auto-standing of the employee's date |
| `contractExpiryRun` | Daily | Warns on contracts expiring within 60 days |
| `runChecksThisMonth` | Monthly, before payroll | Pre-payroll validation |
| `onFormSubmit` | On form submit | Ingests the intake form into `INTAKE` |

---

## Spreadsheet tabs

`EMPLOYEES` · `CHANGE LOG` · `LISTS` · `INTAKE` · `LEAVE` · `HOLIDAYS` ·
`LEAVE_TYPES` · `LEAVE_ADJUSTMENTS` · `DELEGATES` · `RESIGNATIONS` ·
`CLEARANCE` · `NO_SHOW` · `EXPIRY_LOG` · `TERMINATIONS` ·
`MEDICAL_INSURANCE` · `DOC_CHECKLIST` · `FILE_STATUS` · `EMPLOYEE_DOCS` ·
`DEPENDANTS` · `PAYROLL` · `CONTRACTS` · `SIGNING_APPOINTMENTS`

---

## Payroll

`payrollengine.gs` was validated on 20/08/2026 against the July 2026 payroll
produced by the outsourced provider:

| Component | Match |
|---|---|
| Basic, overtime, martyrs fund | 100% |
| Income tax | 99.33% |
| Social insurance | 99.0% |

Tax brackets, insurance ceilings and allowance rates live in `CFG`,
`BRACKET_LOWER`, `RATE_DELTA` and `CANCEL` at the top of `payrollengine.gs`.
**These are law-dependent and are not versioned by effective date** — when the
law changes, historic recalculation will silently use the new figures.

---

## Tests

The payroll engine handles money and had no regression protection. `tests/`
runs its logic offline — no Google account, no network, no dependencies:

```bash
node tests/run.js
```

`tests/harness.js` loads `.gs` files into a sandbox that reproduces Apps
Script's single shared global scope and stubs the Google services. Anything
that genuinely needs a spreadsheet is out of scope by design; the point is to
pin down the arithmetic.

The runner reports two kinds of result:

- **`✓` / `✗`** — assertions. A failure is a regression and exits non-zero.
- **`⚠` flagged for review** — behaviour that looks wrong but has *not* been
  changed, because altering payroll maths needs a human decision. These are
  printed in full at the end of every run and never fail the build.

Nothing is flagged today. Three items were, and two turned out to be real
bugs, now fixed and covered by regression tests: someone joining and leaving in
the same month was paid from the 1st rather than from their hire date, and a
hire dated after the payroll month was paid a full 30 days. The third — a hire
on the 31st being paid zero days — is *correct*: Egypt runs a 30-day month
convention, so nobody is paid for a 31st.

The engine also now refuses a record it cannot price. A blank or non-numeric
`basic_salary` used to propagate `NaN` through the whole payslip while the run
reported success; it now stops with the employee's ID in the message.

---

## Conventions

- A trailing underscore means private/server-internal: `empData_`, `logChange_`.
- `hr*` functions are HR-console entry points and must check `isHR_()` server-side.
- Sheet columns are `snake_case` and are looked up **by header name at runtime**,
  so renaming or reordering a column in the Sheet changes behaviour without any
  code change. See [`docs/SCHEMA.md`](docs/SCHEMA.md).
