# Smart Reminder Hub & Audit-Grade PDF Engine Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Upgrade the IQC Nexus Reminder and PDF Export systems with a 0ms Client-Side Smart Reminder Hub (1-Click Copy by Part, <0.2s SheetJS Excel export, 1-Click Email dispatch) and a Batch-Accelerated Audit-Grade PDF Engine (`UrlFetchApp.fetchAll` parallel signature loading + Vector Print-to-PDF & Direct Server PDF).

**Architecture:**
1. In `Code.js`, replace sequential `DriveApp.getFileById` calls with `resolveSignaturesBatch(signVals)` using `CacheService` + `UrlFetchApp.fetchAll()` to fetch 200+ signatures in parallel (<1.5s).
2. Add `getReminderData(courseId, adminToken)` and `triggerManualReminderEmail(courseId, adminToken)` in `Code.js`, and upgrade `sendIncompleteReport()` with a Part-level summary table.
3. Add `buildAuditReportHtml(course, allCourseAssigns, signMap)` in `Code.js` and expose both `getCourseAuditReportHtml(courseID, adminToken)` (for 1s Vector Print-to-PDF) and `exportCourseReport(courseID, adminToken)` (for direct `.pdf` file download).
4. In `Index.html`, add `#modalReminderHub` and `#modalPdfOptions` (or dual buttons/actions), implement `exportRemindSingle(courseId)` and `exportConsoleReminder()` to open the Smart Reminder Hub, and implement 1-Click Copy, SheetJS 2-sheet `.xlsx` export, and Vector Print-to-PDF.

**Tech Stack:** Google Apps Script (`UrlFetchApp.fetchAll`, `CacheService`, `MailApp`), HTML5/CSS3 (`@media print` A4 Audit Layout), SheetJS (`XLSX.utils.book_new`, `XLSX.writeFile`), Bootstrap 5 Modal.

---

### Task 1: Backend Parallel Signature Resolver & Audit PDF Builder (`Code.js`)

**Files:**
- Modify: `d:\Vibe_coding\xacnhan\Code.js:214-226`
- Modify: `d:\Vibe_coding\xacnhan\Code.js:984-1436`

### Task 2: Frontend Smart Reminder Hub & Dual-Mode Audit PDF (`Index.html`)

**Files:**
- Modify: `d:\Vibe_coding\xacnhan\Index.html:1050-1058`
- Modify: `d:\Vibe_coding\xacnhan\Index.html:1450-1580`
- Modify: `d:\Vibe_coding\xacnhan\Index.html:1905-1915`
- Modify: `d:\Vibe_coding\xacnhan\Index.html:2340-2380`

### Task 3: Commit, Push & Deploy to Google Apps Script Production
- Commit changes to local Git repository.
- Push via `clasp push` and deploy to primary & secondary production deployments.
