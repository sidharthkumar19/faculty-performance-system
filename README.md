# Institutional Monitoring System (IMS)

> A multi-institution, role-based platform for weekly faculty activity reporting, multi-tier review, and institutional performance monitoring across a university's campuses — built entirely on **Google Apps Script**.

---

## Overview

The IMS is the institutional reporting and oversight platform for a multi-campus university, spanning **12 affiliated institutions across 2 campuses**. Faculty submit structured weekly timesheets and self-assessments through a guided, institution-aware interface — including institution-specific time slots and a 3-level activity classification (Category → Sub-Category → Activity). Each submission flows through a configurable multi-tier review chain (**Faculty → HOD → HOI → IMO**), with the HOD tier skipped for the institutions that have no HOD role.

HODs additionally file their own weekly reports, reviewed independently by HOI and IMO. A companion **Institutional Performance Monitoring (IPM)** module gives HOIs and the IMO campus- and institution-level comparison views, best-practice tagging, and exportable reports — including NAAC-aligned summaries.

---

## Stats at a Glance

| 17 | 19 | 12 | 2 |
|:---:|:---:|:---:|:---:|
| Data Sheets | Source Files | Institutions | Campuses |

---

## Tech Stack

| Layer | Details |
|---|---|
| **Backend** | Google Apps Script — 13 server-side modules, SpreadsheetApp API |
| **Frontend** | 6 HTML files — main portal SPA + a separate IPM portal, vanilla JS/CSS |
| **Database** | Google Sheets — 17 named sheets |
| **Auth** | Email + password login (salted hash, throttled), Google SSO, OTP-based activation |
| **IPM Session** | Independent token-based session layer, SSO-bridged from the main HOI login |
| **Email** | MailApp — OTP, activation invites, Friday reminders, comparison digests |
| **Scheduling** | Time-based triggers — Friday reminders, weekly KPI archive, comparison digests |
| **Config** | PropertiesService — IMO/HOI credentials, web app URL, feature toggles |
| **Hosting** | Apps Script Web App deployment |

---

## How It Works

```
FACULTY WEEKLY REPORTS
  Faculty → HOD → HOI → IMO
  (institutions without a HOD tier skip directly: Faculty → HOI → IMO)

HOD's OWN WEEKLY REPORTS
  HOD → HOI → IMO

IPM MODULE (parallel, institution-level)
  Submissions → HOI / IMO comparison & KPI views
```

Each stage fires an in-app notification (and email, where configured) to the next reviewer. Faculty can **withdraw** a pending submission or **recall** a rejected one for resubmission.

---

## User Roles

| Role | Sheet | Responsibilities |
|---|---|---|
| **Faculty** | `Faculty_Master` | Submit weekly timesheet & self-assessment · manage drafts, attachments, and profile · track Pending Priority Works · view history & comparisons |
| **HOD** | `Staff_Master` | Review faculty submissions (Approve/Reject) for their department · submit their own weekly report (reviewed by HOI → IMO) · department KPI drilldowns |
| **HOI** | `Staff_Master` | Review HOD-approved faculty submissions and the HOD's own submissions · log weekly HOI meetings · institution KPIs, stats & report comparisons · IPM access |
| **IMO** | `Staff_Master` | Final monitoring (Under Review/Finalised) for both review chains · system-wide dashboards, NAAC-aligned & comparison reports, weekly KPI archive · notifications, reminders & account provisioning |

---

## Institutions & Campuses

The institution hierarchy (`INSTITUTION_HIERARCHY` in `Code.gs`) models 2 campuses with 12 affiliated institutions in total. Each institution has its own department list, an optional HOD tier, and (for Campus A) its own daily time-slot schedule.

### Campus A

| Institution | Departments | HOD Tier |
|---|:---:|:---:|
| Institution A1 | 8 | ✓ |
| Institution A2 | 1 | ✓ |
| Institution A3 | 10 | ✓ |
| Institution A4 | 7 | no HOD |
| Institution A5 | 1 | no HOD |
| Institution A6 | 2 | no HOD |
| Institution A7 | 2 | no HOD |

### Campus B

| Institution | Departments | HOD Tier |
|---|:---:|:---:|
| Institution B1 | 19 | ✓ |
| Institution B2 | 6 | ✓ |
| Institution B3 | 7 | ✓ |
| Institution B4 | 5 | ✓ |
| Institution B5 | 7 | ✓ |

> Each Campus A institution has its own 6-slot daily timetable (`INSTITUTION_TIME_SLOTS` in `Code.gs`); Campus B institutions use a shared default schedule. The weekly grid covers **Mon–Sat**, with odd/even-Saturday and holiday logic (`_isNonWorkingDay_`) determining which Saturdays count as working days.

---

## Key Features

### Weekly Submission & Timesheets
- 6-day grid (Mon–Sat) with non-working-day logic for alternating Saturdays and institutional holidays
- Institution-specific time slots, with Campus B institutions defaulting to a shared schedule
- Activities organized in a 3-level hierarchy (Category → Sub-Category → Activity); the legacy flat `ACTIVITY_TYPES` list is retained for backward compatibility with pre-migration submissions
- Per-week draft saving/loading and file attachments per activity

### Multi-Tier Review
- Faculty path: Faculty → HOD → HOI → IMO, with HOD skipped for the institutions configured without a HOD tier
- HOD path: HODs submit their own weekly reports, reviewed by HOI (`HOD_Review`) then IMO (`HOD_IMO_Review`)
- Standardized statuses — HOD/HOI use **Approved / Rejected**; IMO uses **Under Review / Finalised**. The earlier "Needs Revision" and "Escalated" states were retired (`migrateNeedsRevisionToRejected()` migrates historical data)
- Faculty can withdraw a pending submission or recall a rejected one for resubmission

### Pending Priority Works
- Faculty log follow-up action items (area, task, responsible person, target date) that carry forward week to week until marked complete

### Faculty Profiles
- Extended CV-style profile — qualification, specialization, research areas, publications, certifications, ORCID/LinkedIn — stored in `FacultyProfiles`, with profile-completion gating at activation

### Report Comparison & Analytics
- Side-by-side comparison across faculty, department, and institution scopes, with saved comparison views per IMO user
- CSV export of comparison reports and scheduled email digests
- Composite scoring and anomaly detection feed weekly KPI snapshots, archived for historical trend views
- Institution "pulse" dashboards, category breakdowns, campus-hierarchy stats, and NAAC-aligned comparison reports for accreditation reporting

### Institutional Performance Monitoring (IPM)
- A companion portal with its own token-based session layer, SSO-bridged from the main HOI login
- Campus/institution submission tracking, IMO review flags, carry-forward items, and best-practice tagging
- Its own comparison reports (institution-vs-institution, campus-vs-campus) with CSV/report export

### Account Activation & Provisioning
- Accounts are pre-provisioned — individually or via bulk Excel import with a dry-run preview — rather than freely self-registered; open self-registration can still be toggled per role
- Users activate via an emailed OTP and set their own password; HOD/HOI credentials can also be seeded per department/institution
- In-app notifications plus Friday reminder emails for faculty who haven't submitted

### Security
- Salted password hashing with a one-time migration path from the legacy hash
- Login throttling after repeated failures, and detection of accounts still on default passwords

### Branding & UI
- Institution logos and campus header images served from `Logos.gs` / `LogosData.gs` and `CampusHeadersData.gs`
- "Meridian / Sovereign Academic" visual theme (deep navy, antique gold, ivory; Playfair Display + DM Sans/Mono), with phased polish layers — e.g. `MeridianPhase5.html` adds a login animation, a refined report-comparison layout, and a WCAG AA contrast pass

---

## Data Schema

| Sheet | Purpose |
|---|---|
| `Staff_Master` | HOD / HOI / IMO accounts — login, role, status |
| `Faculty_Master` | Faculty accounts — login by email, department/campus/institution, status |
| `Weekly_Submission` | Faculty weekly report headers |
| `Timesheet_Entries` | Per-slot activity entries for each submission |
| `Self_Assessment` | Weekly outcome & next-week target |
| `HOD_Remarks` | HOD review of faculty submissions (Approved/Rejected) |
| `HOI_Remarks` | HOI review of faculty submissions (Approved/Rejected) |
| `IMO_Monitoring` | IMO final status for faculty submissions (Under Review/Finalised) |
| `Notifications` | In-app notification feed, scoped by role/institution |
| `PendingPriorityWork` | Carry-forward action items / follow-ups |
| `HOD_Submission` | HOD's own weekly report headers |
| `HOD_Timesheet` | HOD's own timesheet entries |
| `HOD_SelfAssess` | HOD's own self-assessment |
| `HOD_Review` | HOI's review of the HOD's submission |
| `HOD_IMO_Review` | IMO's final status on the HOD's submission |
| `HOI_WeeklyMeeting` | HOI weekly meeting log |
| `FacultyProfiles` | Extended faculty profile/CV data |

> `SubmissionID` links `Weekly_Submission` to its timesheet, self-assessment, and HOD/HOI/IMO review rows; the parallel `HOD_*` sheets follow the same pattern for the HOD's own submissions.

---

## Project Structure

### Backend (`.gs`)

| File | Role |
|---|---|
| `Code.gs` | Config & constants — sheet names/schema, institution hierarchy, time slots, academic years; web app entry point (`doGet`, `include`) |
| `Auth.gs` | Email-based login for all roles, Google SSO, password reset, `getConfig`/`getFacultyList`; HOD/HOI account-seeding & credential tools |
| `Security.gs` | Salted password hashing/verification, login throttling, default-password detection, hash migration |
| `Activation.gs` | OTP account activation, bulk Excel import (staging → dry run → commit), profile-completion gating, activation emails |
| `HOD.gs` | HOD review queue, HOD's own weekly submission, department stats |
| `HOI.gs` | HOI review queues (faculty + HOD), institution KPIs/stats, weekly meeting log |
| `IMO.gs` | System setup, IMO queues & dashboards, notifications/reminders, weekly KPI archive, report comparison & saved views, NAAC reporting |
| `Submissions.gs` | Submission CRUD (submit/withdraw/delete/recall), PP-Works, attachments, draft timesheets, faculty profiles |
| `Utils.gs` | Shared lookups & formatting, period/non-working-day logic, composite scoring, anomaly detection, IPM helpers |
| `IPM_Backend.gs` | IPM portal backend — token sessions, submissions/drafts, IMO views, comparison reports, best-practice tags |
| `Logos.gs` / `LogosData.gs` | Institution logo assets (base64) |
| `CampusHeadersData.gs` | Campus header/banner images (base64) |

### Frontend (`.html`)

| File | Role |
|---|---|
| `Index.html` | App shell — login/registration wizard + role-based portal layout, served by `doGet()` |
| `IndexStyles.html` | Design system & "Meridian / Sovereign Academic" theme |
| `IndexScripts.html` | Main client-side logic — state, navigation, all portal views, `google.script.run` calls |
| `MeridianPhase5.html` | Phase 5 UI polish layer (append-only addition to `IndexStyles.html`) |
| `IPMStyles.html` | IPM portal styling |
| `IPMScripts.html` | IPM portal client-side logic — token session, comparison views, KPIs |

---

## Deployment

```bash
# 1. Open the Google Sheet → Extensions → Apps Script
# 2. Create each file above with the EXACT names shown (Code.gs, Auth.gs, ... Index.html, etc.)
# 3. Run initializeSystem() (in IMO.gs) once
#    → creates/repairs all 17 sheets, headers, validations & triggers
#    → seeds IMO_EMAIL / IMO_PASSWORD / HOI_EMAIL / HOI_PASSWORD in Script Properties
# 4. Run seedHODAccounts() (in Auth.gs) to generate per-department HOD logins
# 5. Deploy → New Deployment → Web App
#    Execute as: Me  |  Access: Anyone (or your org)
# 6. Update default credentials in Script Properties before go-live
# 7. Share the Web App URL
```

### Default Credentials (Script Properties)

| Property | Default Value |
|---|---|
| `IMO_EMAIL` | `imo@yourdomain.edu` |
| `IMO_PASSWORD` | `(set in Script Properties)` |
| `HOI_EMAIL` | `hoi@yourdomain.edu` |
| `HOI_PASSWORD` | `(set in Script Properties)` |

> ⚠️ Set real values in **Extensions → Apps Script → Project Settings → Script Properties** before go-live. HOD credentials are generated per department by `seedHODAccounts()`.

---

## Architecture

```
Browser
  ├─ Main Portal  (Index.html / IndexScripts.html / IndexStyles.html)
  │     │  google.script.run — email+password / Google SSO
  └─ IPM Portal   (IPMScripts.html / IPMStyles.html)
        │  ipmLogin() — token session, SSO-bridged from HOI login
        ▼
Apps Script Server
  Code.gs · Auth.gs · Security.gs · Activation.gs
  HOD.gs · HOI.gs · IMO.gs · Submissions.gs · Utils.gs
  IPM_Backend.gs · Logos.gs / LogosData.gs · CampusHeadersData.gs
        │  SpreadsheetApp API
        ▼
Google Sheets (17 named sheets)
        │  MailApp / Time-based Triggers
        ▼
Email — OTP activation · Friday reminders · comparison digests
```

---

## Status

This system is under active development. The schema, review-status vocabulary, and account-provisioning model have all evolved substantially from the original single-portal prototype — `migratePasswordsToSecureFormat()` and `migrateNeedsRevisionToRejected()` exist specifically to carry existing data forward through these changes.

---

*Built with Google Apps Script · 2026*
